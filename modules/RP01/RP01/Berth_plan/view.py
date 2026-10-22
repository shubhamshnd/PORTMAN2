"""
RP02 — Berth Plan (page + daily-report data), living on the shared

"""

from datetime import datetime, timedelta
from functools import wraps

from flask import request, jsonify, render_template, session, redirect, url_for

from .. import bp          # shared RP01 blueprint
from database import get_db, get_cursor, get_user_permissions
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from flask import send_file
# from modules.LUEU01.model import get_started_parcels

MODULE_CODE = 'RP01'
# ---- adjust if your real column name differs ----------
SAIL_COLUMN = 'cast_off_datetime'   # ldud_header column: actual sail time
LDUD_TERMINAL_COL = 'terminal_name'        # ldud_parcel_ops: terminal name column
LUEU_EQUIPMENT_COL = 'equipment_name'      # lueu_parcel_log: equipment name column
# BERTHS constant removed — berths are now fetched live from port_berth_master
# -------------------------------------------------------------------------
# ---- adjust if your real column name differs ----------
SAIL_COLUMN = 'cast_off_datetime'   # ldud_header column: actual sail time
# BERTHS constant removed — berths are now fetched live from port_berth_master
# -------------------------------------------------------------------------


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), MODULE_CODE)


def get_berths(cur=None):
    """Live list of berth names from port_berth_master (no hardcoding)."""
    own_conn = cur is None
    if own_conn:
        conn = get_db()
        cur = get_cursor(conn)
    cur.execute('SELECT berth_name FROM port_berth_master ORDER BY berth_name')
    berths = [r['berth_name'] for r in cur.fetchall()]
    if own_conn:
        conn.close()
    return berths

def get_expected_waiting_vessels(window_start, window_end):
    conn = get_db()
    cur = get_cursor(conn)

    try:
        cur.execute(f"""
            SELECT
                vh.id,
                vh.vessel_name,
                vh.via_number,
                vh.loa,
                vh.draft,
                vh.vessel_agent_name AS agents,
                vh.berth_name,
                vh.doc_date,
                vh.cargo_type AS cargo_name,

                parcels.terminal_name AS declared_terminal_name,
                parcels.total_quantity AS cargo_quantity,
                parcels.equipment_names,
                parcels.consignee_codes,
                vh.load_port,
                ldud.alongside_datetime,
                ldud.anchored_datetime,
                ldud.nor_accepted,

                ops.terminal_name AS ops_terminal_name,
                ops.has_equipment

            FROM vcn_header vh

            LEFT JOIN LATERAL (
                SELECT
                    STRING_AGG(DISTINCT NULLIF(TRIM(p.unload_terminal), ''), ', ') AS terminal_name,
                    STRING_AGG(DISTINCT NULLIF(TRIM(p.equipment_names), ''), ', ') AS equipment_names,
                    STRING_AGG(DISTINCT NULLIF(TRIM(COALESCE(NULLIF(TRIM(vc.customer_code), ''), NULLIF(TRIM(p.consigner_name), ''))), ''), ', ') AS consignee_codes,
                    (
                        SELECT SUM(NULLIF(q.quantity, '')::numeric)
                        FROM (
                            SELECT quantity FROM vcn_consigners WHERE vcn_id = vh.id
                            UNION ALL
                            SELECT quantity FROM vcn_export_cargo_declaration WHERE vcn_id = vh.id
                        ) q
                    ) AS total_quantity
                FROM (
                    SELECT unload_terminal, equipment_names, quantity, consigner_name
                    FROM vcn_consigners
                    WHERE vcn_id = vh.id

                    UNION ALL

                    SELECT unload_terminal, equipment_names, quantity, consigner_name
                    FROM vcn_export_cargo_declaration
                    WHERE vcn_id = vh.id
                ) p
                LEFT JOIN vessel_customers vc ON (
                    UPPER(TRIM(vc.name)) = UPPER(TRIM(p.consigner_name))
                    OR UPPER(TRIM(vc.customer_code)) = UPPER(TRIM(p.consigner_name))
                )
            ) parcels ON TRUE

            LEFT JOIN LATERAL (
                SELECT
                    alongside_datetime,
                    anchored_datetime,
                    nor_accepted
                FROM ldud_header
                WHERE vcn_id = vh.id
                ORDER BY id DESC
                LIMIT 1
            ) ldud ON TRUE

            LEFT JOIN LATERAL (
                SELECT
                    STRING_AGG(DISTINCT NULLIF(TRIM(po.{LDUD_TERMINAL_COL}), ''), ', ') AS terminal_name,
                    CASE WHEN BOOL_OR(eq.has_eq) THEN 'Y' ELSE 'N' END AS has_equipment
                FROM ldud_header lh
                JOIN ldud_parcel_ops po ON po.ldud_id = lh.id
                LEFT JOIN LATERAL (
                    SELECT EXISTS (
                        SELECT 1
                        FROM lueu_parcel_log l
                        WHERE l.parcel_op_id = po.id
                          AND l.is_deleted IS NOT TRUE
                          AND NULLIF(TRIM(l.{LUEU_EQUIPMENT_COL}), '') IS NOT NULL
                    ) AS has_eq
                ) eq ON TRUE
                WHERE lh.vcn_id = vh.id
            ) ops ON TRUE

            WHERE
                ldud.alongside_datetime IS NULL
                OR NULLIF(TRIM(ldud.alongside_datetime::text), '') IS NULL

            ORDER BY vh.doc_date ASC NULLS LAST
        """)

        rows = [dict(r) for r in cur.fetchall()]

    finally:
        conn.close()

    out = []

    def _fmt_q(q_val):
        try:
            f = float(q_val)
            return str(int(f)) if f.is_integer() else str(f)
        except (ValueError, TypeError):
            return str(q_val)

    for r in rows:
        tank_terminal = "JJLTPL"
        load_port =  r.get("declared_terminal_name") or ""

        # CONS shows consignee customer code(s) assigned to the vessel
        cons = (r.get("consignee_codes") or '').strip()

        tot = r.get("cargo_quantity")

        out.append({
            "terminal":     tank_terminal,
            "vessel_name":  r.get("vessel_name"),
            "via_no":       r.get("via_number"),
            "loa":          r.get("loa"),
            "dft":          r.get("draft"),
            "agt_tnk_cons": f"{r.get('agents') or ''} / {load_port or ''} / {cons}",
            "cargo":        r.get("cargo_name"),
            "mla":          r.get("equipment_names"),
            "quantity":     _fmt_q(tot) if tot is not None else "",
            "eta":          _fmt_dt(r.get("doc_date"), fmt='%d-%m-%Y'),
            "ata":          _fmt_dt(r.get("anchored_datetime")),
            "nor":          _fmt_dt(r.get("nor_accepted")),
            "berth":        r.get("berth_name"),
        })

    return out

# ══════════════════════════════════════════════════════════════════
#  Page route
# ══════════════════════════════════════════════════════════════════

@bp.route('/module/RP01/berth-plan/')
@login_required
def berth_plan_page():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('berth_plan.html', permissions=perms, berths=get_berths())


def get_report_window(plan_date_str):
    """
    Report Window

    Selected Date = 23-07-2026

    Window:
        22-07-2026 07:00
            ↓
        23-07-2026 07:00

    Today:
        Window end = current time (if current time is before 07:00,
        use current time)
    """

    plan_date = datetime.strptime(plan_date_str, "%Y-%m-%d").date()
    today = datetime.now().date()

    window_end = datetime.combine(
        plan_date,
        datetime.min.time()
    ) + timedelta(hours=7)

    window_start = window_end - timedelta(days=1)

    if plan_date == today:
        now = datetime.now()
        if now < window_end:
            window_end = now

    return window_start, window_end


def _fmt_dt(v, fmt='%d-%m-%Y %H:%M'):
    """Format a value that may be a real datetime OR a raw text timestamp
    (e.g. po.start_dt/end_dt stored as 'YYYY-MM-DDTHH:MM' strings)."""
    if not v:
        return ''
    if isinstance(v, str):
        try:
            return datetime.strptime(v.replace('T', ' ')[:16], '%Y-%m-%d %H:%M').strftime(fmt)
        except ValueError:
            return v
    return v.strftime(fmt)


def _cumulative_as_of(cur, ldud_id, window_end):
    """Everything logged strictly BEFORE window_end (i.e. as of the report's
    cut-off date/time) — so 'Till now', 'Balance', and 'Present Flow Rate' for
    a past-dated report reflect only the data that existed up to that day,
    not entries operators logged later for subsequent days.

    Returns (logged_qty, total_hours) — real (non-shortclose) rows only, so
    the flow-rate math matches get_started_parcels' avg_rate definition.
    Shortclosed quantity is intentionally excluded here too, matching the
    'real_qty' used for avg_rate in LUEU01; balance still nets against target
    the same way callers already expect (target - logged), consistent with
    how a report for an in-progress vessel should read as of that date.
    """
    cur.execute('''
        SELECT po.id AS parcel_op_id
        FROM ldud_parcel_ops po WHERE po.ldud_id = %s
    ''', [ldud_id])
    pop_ids = [r['parcel_op_id'] for r in cur.fetchall()]
    if not pop_ids:
        return 0.0, 0.0, 0.0

    cur.execute('''
        SELECT parcel_op_id,
               COALESCE(SUM(quantity), 0) AS qty,
               COALESCE(SUM(
                 EXTRACT(EPOCH FROM (
                   COALESCE(NULLIF(to_time,'')::time, '00:00'::time)
                   - COALESCE(NULLIF(from_time,'')::time, '00:00'::time)
                   + CASE WHEN NULLIF(to_time,'')::time < NULLIF(from_time,'')::time
                          THEN INTERVAL '24 hours' ELSE INTERVAL '0' END
                 )) / 3600.0
               ), 0) AS hrs
        FROM lueu_parcel_log
        WHERE parcel_op_id = ANY(%s)
          AND is_deleted IS NOT TRUE
          AND is_shortclose IS NOT TRUE
          AND COALESCE(UPPER(TRIM(remarks)), '') != 'SHORT CLOSE'
          AND (remarks IS NULL OR remarks NOT ILIKE '%%short%%close%%')
          AND (NULLIF(entry_date, '')::timestamp
               + COALESCE(NULLIF(from_time, '')::time, '00:00'::time)) < %s
        GROUP BY parcel_op_id
    ''', [pop_ids, window_end])
    rows = cur.fetchall()

    total_qty = 0.0
    max_hrs = 0.0
    total_rate = 0.0
    for r in rows:
        q = float(r['qty'] or 0)
        h = float(r['hrs'] or 0)
        total_qty += q
        if h > max_hrs:
            max_hrs = h
        if h > 0:
            total_rate = total_qty / max_hrs if max_hrs > 0 else 0.0

    return total_qty, max_hrs, round(total_rate, 2)


MIN_HOURS_FOR_ACTUAL = 4      # must match MIN_HOURS_FOR_ACTUAL in lueu01.html
MAX_REASONABLE_DAYS = 90      # if the actual-rate ETC projects beyond this, treat it as unreliable
def _per_parcel_completion(cur, ldud_id, window_end):
    """ETC per parcel_op (= per discharge line running in parallel), computed
    the same way LUEU01 shows it: LUEU01.current_etc (last log end + balance /
    current flow rate), using ALL logged entries (no window cutoff, no
    minimum-hours gate) so it matches LUEU01's live ETC exactly. Falls back to
    Expected Start/Rate only when there's no actual logged data at all yet."""
    from modules.LUEU01.model import current_etc

    cur.execute('''
        SELECT id, quantity, start_dt, end_dt, expected_start, expected_flow_rate
        FROM ldud_parcel_ops
        WHERE ldud_id = %s
    ''', [ldud_id])
    pos = cur.fetchall()

    results = []
    for po in pos:
        po_id = po['id']
        target = float(po['quantity'] or 0)

        # If parcel already has an end_dt, it has completed and has no ongoing ETC
        if po.get('end_dt') and str(po['end_dt']).strip():
            results.append((None, False))
            continue

        cur.execute('''
            SELECT entry_date, from_time, to_time, COALESCE(quantity, 0) AS quantity
            FROM lueu_parcel_log
            WHERE parcel_op_id = %s
              AND is_deleted IS NOT TRUE
              AND is_shortclose IS NOT TRUE
              AND COALESCE(UPPER(TRIM(remarks)), '') != 'SHORT CLOSE'
              AND (remarks IS NULL OR remarks NOT ILIKE '%%short%%close%%')
            ORDER BY entry_date, from_time NULLS LAST, id
        ''', [po_id])
        rows = [dict(r) for r in cur.fetchall()]
        logged_qty = sum(float(r['quantity'] or 0) for r in rows)

        # Count shortclose quantity towards target completion (matching LUEU01)
        cur.execute('''
            SELECT COALESCE(SUM(quantity), 0) AS sc_qty
            FROM lueu_parcel_log
            WHERE parcel_op_id = %s
              AND is_deleted IS NOT TRUE
              AND (is_shortclose IS TRUE
                   OR UPPER(TRIM(remarks)) = 'SHORT CLOSE'
                   OR remarks ILIKE '%%short%%close%%')
        ''', [po_id])
        sc_qty = float(cur.fetchone()['sc_qty'] or 0)
        balance = target - (logged_qty + sc_qty)

        etc, planned = None, False
        if balance > 0:
            _, etc = current_etc(rows, balance)
            if etc is None:
                exp_rate = float(po.get('expected_flow_rate') or 0)
                exp_start = po.get('expected_start')
                if exp_rate and exp_start and target:
                    try:
                        start_dt = datetime.strptime(
                            str(exp_start).replace('T', ' ')[:16], '%Y-%m-%d %H:%M')
                        etc = start_dt + timedelta(hours=target / exp_rate)
                        planned = True
                    except ValueError:
                        pass
        results.append((etc, planned))
    return results

def _enrich_vessel(cur, vcn_id, ldud_id, window_start, window_end):
    from modules.LUEU01.model import get_started_parcels

    cur.execute('SELECT operation_type FROM vcn_header WHERE id=%s', [vcn_id])
    row = cur.fetchone()
    is_export = (row or {}).get('operation_type') == 'Export'
    tbl = 'vcn_export_cargo_declaration' if is_export else 'vcn_consigners'

    cur.execute(f'''
        SELECT DISTINCT
            COALESCE(NULLIF(TRIM(vc.customer_code), ''), TRIM(t.consigner_name)) AS code
        FROM {tbl} t
        LEFT JOIN vessel_customers vc ON (
            UPPER(TRIM(vc.name)) = UPPER(TRIM(t.consigner_name))
            OR UPPER(TRIM(vc.customer_code)) = UPPER(TRIM(t.consigner_name))
        )
        WHERE t.vcn_id = %s AND NULLIF(TRIM(t.consigner_name), '') IS NOT NULL
    ''', [vcn_id])
    consigner = ', '.join(r['code'] for r in cur.fetchall() if r['code'])

    cur.execute(f'''
        SELECT DISTINCT
            COALESCE(
                NULLIF(TRIM(vc.cargo_code), ''),
                NULLIF(TRIM(REGEXP_REPLACE(t.cargo_name, '\\s*\\[.*\\]', '')), ''),
                TRIM(t.cargo_name)
            ) AS code
        FROM {tbl} t
        LEFT JOIN vessel_cargo vc ON (
            UPPER(TRIM(vc.cargo_name)) = UPPER(TRIM(t.cargo_name))
            OR UPPER(TRIM(vc.cargo_code)) = UPPER(TRIM(t.cargo_name))
            OR UPPER(TRIM(vc.cargo_name)) = UPPER(TRIM(REGEXP_REPLACE(t.cargo_name, '\\s*\\[.*\\]', '')))
            OR UPPER(TRIM(vc.cargo_code)) = UPPER(TRIM(REGEXP_REPLACE(t.cargo_name, '\\s*\\[.*\\]', '')))
        )
        WHERE t.vcn_id = %s AND NULLIF(TRIM(t.cargo_name), '') IS NOT NULL
    ''', [vcn_id])
    cargos = [r['code'] for r in cur.fetchall() if r['code']]
    cargo_code_str = ', '.join(cargos) if cargos else None

    cur.execute('''SELECT MIN(po.start_dt::timestamp) AS started
                   FROM ldud_parcel_ops po WHERE po.ldud_id=%s''', [ldud_id])
                    
    ops_commenced = cur.fetchone()['started']
    cur.execute('''
        SELECT MAX(po.end_dt::timestamp) AS completed
        FROM ldud_parcel_ops po
        WHERE po.ldud_id = %s
        AND po.end_dt IS NOT NULL
        AND NULLIF(TRIM(po.end_dt::text), '') IS NOT NULL
    ''', [ldud_id])

    cargo_completed = cur.fetchone()['completed']
    
    # ── Remarks: last remark of every IN-PROGRESS parcel (balance > 0).
    #    If every parcel is complete, fall back to a shortclose remark if
    #    one exists, else the last parcel's last remark. ──
    cur.execute('''
        SELECT id, quantity
        FROM ldud_parcel_ops
        WHERE ldud_id = %s
        ORDER BY id ASC
    ''', [ldud_id])
    remark_pos = cur.fetchall()

    parcel_remarks = []       # last remark text for each in-progress parcel
    last_parcel_remark = ""   # last remark of the last parcel (any status)
    shortclose_remark = ""    # most recent shortclose remark, if any

    for idx, po in enumerate(remark_pos, start=1):
        po_id = po['id']
        target = float(po['quantity'] or 0)

        cur.execute('''
            SELECT COALESCE(SUM(quantity), 0) AS qty
            FROM lueu_parcel_log
            WHERE parcel_op_id = %s
              AND is_deleted IS NOT TRUE
              AND is_shortclose IS NOT TRUE
              AND COALESCE(UPPER(TRIM(remarks)), '') != 'SHORT CLOSE'
              AND (remarks IS NULL OR remarks NOT ILIKE '%%short%%close%%')
        ''', [po_id])
        logged_qty = float(cur.fetchone()['qty'] or 0)
        in_progress = (target - logged_qty) > 0

        cur.execute('''
            SELECT remarks
            FROM lueu_parcel_log
            WHERE parcel_op_id = %s
              AND is_deleted IS NOT TRUE
            ORDER BY
                NULLIF(entry_date,'')::date DESC,
                NULLIF(to_time,'')::time DESC,
                id DESC
            LIMIT 1
        ''', [po_id])
        last_row = cur.fetchone()
        last_remark = (last_row['remarks'] if last_row else '') or ''

        if in_progress and last_remark:
            parcel_remarks.append(last_remark)
        if idx == len(remark_pos):
            last_parcel_remark = last_remark

        cur.execute('''
            SELECT remarks
            FROM lueu_parcel_log
            WHERE parcel_op_id = %s
              AND is_deleted IS NOT TRUE
              AND is_shortclose IS TRUE
            ORDER BY
                NULLIF(entry_date,'')::date DESC,
                NULLIF(to_time,'')::time DESC,
                id DESC
            LIMIT 1
        ''', [po_id])
        sc_row = cur.fetchone()
        if sc_row and sc_row['remarks']:
            shortclose_remark = sc_row['remarks']

    if parcel_remarks:
        remarks = ', '.join(parcel_remarks)
    elif shortclose_remark:
        remarks = shortclose_remark
    else:
        remarks = last_parcel_remark

    # target_qty comes from the live VCN parcel quantities (via get_started_parcels) —
    # this is a "current truth" number and is NOT date-dependent.
    parcels = get_started_parcels(vcn_id)
    target_qty = float(sum(p['target_qty'] for p in parcels))

    # ── logged_qty / total_hours are date-dependent: they must reflect ONLY
    #    what had actually been logged as of this report's date, not the
    #    all-time total (which may already include days after the report date
    #    if operators have logged ahead). This makes 'Till now', 'Balance',
    #    and 'Present Flow Rate' change day-by-day correctly for a vessel that
    #    spans multiple days, instead of always showing the final/latest state. ──
    logged_qty, total_hours, present_flow_rate = _cumulative_as_of(cur, ldud_id, window_end)
    balance_qty = round(target_qty - logged_qty, 3)

    # Last-24-hours figure — already correctly window-bound (unchanged).
    cur.execute('''
        SELECT COALESCE(SUM(l.quantity), 0) AS q
        FROM lueu_parcel_log l
        JOIN ldud_parcel_ops po ON po.id = l.parcel_op_id
        WHERE po.ldud_id = %s
        AND l.is_deleted IS NOT TRUE
        AND l.is_shortclose IS NOT TRUE
        AND COALESCE(UPPER(TRIM(l.remarks)), '') != 'SHORT CLOSE'
        AND (l.remarks IS NULL OR l.remarks NOT ILIKE '%%short%%close%%')
        AND (
                NULLIF(l.entry_date, '')::timestamp
                + COALESCE(NULLIF(l.from_time, '')::time, '00:00'::time)
            ) >= %s
        AND (
                NULLIF(l.entry_date, '')::timestamp
                + COALESCE(NULLIF(l.from_time, '')::time, '00:00'::time)
            ) < %s
    ''', [ldud_id, window_start, window_end])
    last_24hr_qty = round(float(cur.fetchone()['q'] or 0), 3)

    expected_completion = '-'
    is_planned = False
    display_rate = present_flow_rate   # default: show the real actual rate (as-of this date)

    def _planned_etc():
        """Latest projected ETC + its rate, across all parcels (Expected Start/Rate)."""
        best_etc, best_rate = None, 0
        for p in parcels:
            rate = float(p.get('expected_flow_rate') or 0)
            start = p.get('expected_start')
            tgt = float(p.get('target_qty') or 0)
            if not (rate and start and tgt):
                continue
            try:
                start_dt = datetime.strptime(str(start).replace('T', ' ')[:16], '%Y-%m-%d %H:%M')
            except ValueError:
                continue
            etc_dt = start_dt + timedelta(hours=tgt / rate)
            if best_etc is None or etc_dt > best_etc:
                best_etc, best_rate = etc_dt, rate
        return best_etc, best_rate

    # whether the actual logged rate (as of this report's date) is trustworthy
    # enough to use as-is
    has_actual = present_flow_rate > 0 and total_hours >= MIN_HOURS_FOR_ACTUAL

    def _parse_completion_dt(v):
        if not v:
            return None
        if isinstance(v, datetime):
            return v
        s = str(v).replace('T', ' ').strip()
        for fmt in ('%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M'):
            try:
                return datetime.strptime(s[:16], fmt)
            except ValueError:
                pass
        try:
            return datetime.fromisoformat(s)
        except Exception:
            return None

    active_lines = []
    for p in parcels:
        if p.get('is_removed'):
            continue

        # 1. Parcel has a live ETC
        if p.get('etc'):
            dt = _parse_completion_dt(p['etc'])
            if dt:
                active_lines.append((dt, False))
                continue

        # 2. Parcel has an actual completion end time
        if p.get('end_dt'):
            dt = _parse_completion_dt(p['end_dt'])
            if dt:
                active_lines.append((dt, False))
                continue

        # 3. Parcel has planned start & expected flow rate (not started yet)
        rate = float(p.get('expected_flow_rate') or 0)
        start = p.get('expected_start')
        tgt = float(p.get('target_qty') or 0)
        if rate and start and tgt:
            start_dt = _parse_completion_dt(start)
            if start_dt:
                active_lines.append((start_dt + timedelta(hours=tgt / rate), True))

    completion_is_planned = False
    if active_lines:
        # Pick the maximum (latest) completion time among all parcels (ETCs and end times)
        etc, completion_is_planned = max(active_lines, key=lambda x: x[0])
        expected_completion = etc.strftime('%d-%m-%Y %H:%M')
    else:
        expected_completion = '-'

    rate_is_planned = False
    if not has_actual:
        _, best_rate = _planned_etc()
        if best_rate:
            display_rate = best_rate
            rate_is_planned = True

    res = {
        'consigner': consigner,
        'quantity': target_qty,
        'ops_commenced': _fmt_dt(ops_commenced),          # First parcel start
        'cargo_completion': _fmt_dt(cargo_completed),     # Last parcel end
        'last_24hr_qty': last_24hr_qty,
        'till_now_qty': round(logged_qty, 3),
        'balance': balance_qty,
        'expected_completion': expected_completion,
        'present_flow_rate': display_rate,
        'is_planned': rate_is_planned,
        'etc_is_planned': completion_is_planned,
        'remarks': remarks,
    }
    if cargo_code_str:
        res['cargo'] = cargo_code_str
    return res


def _base_row(cur, h):
    raw_cargo = (h.get('cargo_type') or '').strip()
    cargo_val = raw_cargo
    if raw_cargo and cur:
        cur.execute('''
            SELECT COALESCE(
                NULLIF(TRIM(cargo_code), ''),
                NULLIF(TRIM(REGEXP_REPLACE(%s, '\\s*\\[.*\\]', '')), ''),
                TRIM(cargo_name)
            ) AS code
            FROM vessel_cargo
            WHERE UPPER(TRIM(cargo_name)) = UPPER(TRIM(%s))
               OR UPPER(TRIM(cargo_code)) = UPPER(TRIM(%s))
               OR UPPER(TRIM(cargo_name)) = UPPER(TRIM(REGEXP_REPLACE(%s, '\\s*\\[.*\\]', '')))
               OR UPPER(TRIM(cargo_code)) = UPPER(TRIM(REGEXP_REPLACE(%s, '\\s*\\[.*\\]', '')))
            LIMIT 1
        ''', [raw_cargo, raw_cargo, raw_cargo, raw_cargo, raw_cargo])
        r = cur.fetchone()
        if r and r['code']:
            cargo_val = r['code']
        else:
            import re
            cargo_val = re.sub(r'\\s*\\[.*\\]', '', raw_cargo).strip() or raw_cargo

    return {
        # carried through so callers can reach the vessel's own parcels
        # (BPL01 renders them read-only); ignored by this module's own output
        'vcn_id': h.get('vcn_id'),
        'via_no': h.get('via_number') or '',
        'vessel_name': h.get('vessel_name') or '',
        'loa': h.get('loa'),
        'draft': h.get('draft'),
        'agent': h.get('vessel_agent_name') or '',
        'cargo': cargo_val,
        'berth_name': h.get('berth_name') or '',
        'imo_num': h.get('imo_num') or '',
        'nationality': h.get('nationality') or '',
        'vessel_run_type': h.get('vessel_run_type') or '',
        'remarks': '',
    }

def get_berthed_vessels(window_start, window_end, berths):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f'''
        SELECT DISTINCT h.id AS vcn_id, l.id AS ldud_id, h.via_number, h.vessel_name,
               h.loa, h.draft, h.vessel_agent_name, h.cargo_type, h.berth_name,
               h.operation_type,
               v.imo_num, v.nationality,
               l.alongside_datetime,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(ec.unload_terminal), ''), ', ')
                  FROM vcn_export_cargo_declaration ec
                 WHERE ec.vcn_id = h.id) AS exp_terminal,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(ec.pipeline_name), ''), ', ')
                  FROM vcn_export_cargo_declaration ec
                 WHERE ec.vcn_id = h.id) AS exp_pipeline,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(cn.unload_terminal), ''), ', ')
                  FROM vcn_consigners cn
                 WHERE cn.vcn_id = h.id) AS imp_terminal,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(cn.pipeline_name), ''), ', ')
                  FROM vcn_consigners cn
                 WHERE cn.vcn_id = h.id) AS imp_pipeline
        FROM ldud_header l
        JOIN vcn_header h ON h.id = l.vcn_id
        LEFT JOIN vessels v
            ON UPPER(REPLACE(TRIM(v.vessel_name), 'MT ', ''))
            =
            UPPER(REPLACE(TRIM(h.vessel_name), 'MT ', ''))
        WHERE h.berth_name = ANY(%s)
          AND l.alongside_datetime IS NOT NULL
          AND (NULLIF(l.alongside_datetime::text, ''))::timestamp < %s
            AND (
                l.{SAIL_COLUMN} IS NULL
                OR NULLIF(TRIM(l.{SAIL_COLUMN}::text), '') IS NULL
                OR l.{SAIL_COLUMN}::timestamp >= %s
                )
                    ORDER BY h.berth_name, l.alongside_datetime DESC
    ''', [berths, window_end, window_end])
    headers = [dict(r) for r in cur.fetchall()]

    # ===== Keep only the LATEST vessel per berth (a berth can physically =====
    # ===== hold just one vessel — dedupe any stale/overlapping records) =====
    latest_per_berth = {}
    for h in headers:
        b = h['berth_name']
        if b not in latest_per_berth:
            latest_per_berth[b] = h  # first row per berth = latest, thanks to ORDER BY above
    headers = list(latest_per_berth.values())

    out = []
    for h in headers:
        row = _base_row(cur, h)
        row['alongside'] = _fmt_dt(h['alongside_datetime'])
        row['vessel_agent'] = h['vessel_agent_name']
        row['terminal'] = h['exp_terminal'] if h['operation_type'] == 'Export' else h['imp_terminal']
        row['pipeline'] = h['exp_pipeline'] if h['operation_type'] == 'Export' else h['imp_pipeline']
        row.update(_enrich_vessel(cur, h['vcn_id'], h['ldud_id'], window_start, window_end))
        # Check whether discharge was completed as of the selected report date
        sail_dt = h.get("cast_off_datetime")

        if sail_dt and sail_dt < window_end:
            continue

        out.append(row)
    conn.close()
    return out



def get_sailed_vessels(window_start, window_end, berths):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(f'''
        SELECT DISTINCT h.id AS vcn_id, l.id AS ldud_id, h.via_number, h.vessel_name,
               h.loa, h.draft, h.vessel_agent_name, h.cargo_type, h.berth_name,
               COALESCE(NULLIF(TRIM(l.operation_type), ''), NULLIF(TRIM(h.operation_type), ''), '') AS operation_type,
               v.imo_num, v.nationality,
               l.alongside_datetime, NULLIF(TRIM(l.{SAIL_COLUMN}::text), '')::timestamp AS sail_dt,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(ec.unload_terminal), ''), ', ')
                  FROM vcn_export_cargo_declaration ec
                 WHERE ec.vcn_id = h.id) AS exp_terminal,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(ec.pipeline_name), ''), ', ')
                  FROM vcn_export_cargo_declaration ec
                 WHERE ec.vcn_id = h.id) AS exp_pipeline,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(cn.unload_terminal), ''), ', ')
                  FROM vcn_consigners cn
                 WHERE cn.vcn_id = h.id) AS imp_terminal,
               (SELECT STRING_AGG(DISTINCT NULLIF(TRIM(cn.pipeline_name), ''), ', ')
                  FROM vcn_consigners cn
                 WHERE cn.vcn_id = h.id) AS imp_pipeline,
               (SELECT MAX(po.end_dt::timestamp)
                  FROM ldud_parcel_ops po
                 WHERE po.ldud_id = l.id) AS cargo_completion_dt
        FROM ldud_header l
        JOIN vcn_header h ON h.id = l.vcn_id
        LEFT JOIN vessels v
            ON UPPER(REPLACE(TRIM(v.vessel_name), 'MT ', ''))
            =
            UPPER(REPLACE(TRIM(h.vessel_name), 'MT ', ''))
        ORDER BY h.berth_name
    ''')
    headers = [dict(r) for r in cur.fetchall()]

    def _in_window(dt):
        # window_start (inclusive) -> window_end (exclusive), so a vessel
        # landing exactly on the boundary belongs to ONE day only, never both
        return dt is not None and window_start <= dt < window_end

    out = []
    for h in headers:
        sail_dt = h['sail_dt']
        completion_dt = h['cargo_completion_dt']

        # Show only vessels that have actually sailed
        if sail_dt is None:
            continue

        # Show only if sailed during the selected report window
        if not (window_start <= sail_dt < window_end):
            continue

        row = _base_row(cur, h)
        row['alongside'] = _fmt_dt(h['alongside_datetime'])
        row['cast_off'] = _fmt_dt(sail_dt)
        row['cargo_completion'] = _fmt_dt(completion_dt)
        row['vessel_agent'] = h['vessel_agent_name']
        is_export = str(h.get('operation_type') or '').strip().lower() == 'export'
        row['terminal'] = (
            h['exp_terminal']
            if is_export
            else h['imp_terminal']
        )
        row['pipeline'] = (
            h['exp_pipeline']
            if is_export
            else h['imp_pipeline']
        )

        row.update(
            _enrich_vessel(
                cur,
                h['vcn_id'],
                h['ldud_id'],
                window_start,
                window_end
            )
        )

        # Section B QUANTITY: Show Net BL Quantity matching LDUD01
        # (Declared VCN BL + additional - short close)
        cur.execute('''
            SELECT COALESCE(SUM(po.additional_qty), 0) AS a
            FROM ldud_parcel_ops po
            WHERE po.ldud_id = %s
        ''', [h['ldud_id']])
        add_qty = float(cur.fetchone()['a'] or 0)

        cur.execute('''
            SELECT COALESCE(SUM(lg.quantity), 0) AS sc
            FROM lueu_parcel_log lg
            JOIN ldud_parcel_ops po ON po.id = lg.parcel_op_id
            WHERE po.ldud_id = %s
              AND lg.is_shortclose IS TRUE
              AND lg.is_deleted IS NOT TRUE
        ''', [h['ldud_id']])
        sc_qty = float(cur.fetchone()['sc'] or 0)

        decl_target = float(row.get('quantity') or 0)
        net_bl_qty = decl_target + add_qty - sc_qty
        row['declared_quantity'] = decl_target
        row['quantity'] = int(net_bl_qty) if net_bl_qty.is_integer() else round(net_bl_qty, 3)

        if row.get("ops_commenced") and row.get("cargo_completion") and row.get("quantity"):
            try:
                ops = datetime.strptime(row["ops_commenced"], "%d-%m-%Y %H:%M")
                completion = datetime.strptime(row["cargo_completion"], "%d-%m-%Y %H:%M")

                hours = (completion - ops).total_seconds() / 3600

                if hours > 0:
                    row["present_flow_rate"] = round(
                        float(row["quantity"]) / hours,
                        2
                    )
            except (ValueError, TypeError):
                pass

        out.append(row)
    conn.close()
    return out

def get_daily_report(plan_date_str):
    window_start, window_end = get_report_window(plan_date_str)
    conn = get_db()
    cur = get_cursor(conn)
    berths = get_berths(cur)
    conn.close()

    return {
        'as_on_date': window_end.strftime('%d-%m-%Y'),   # <-- window_end, not window_start
        'as_on_time': window_end.strftime('%H:%M'),        # <-- window_end, not window_start
        'window_start': window_start.isoformat(),
        'window_end': window_end.isoformat(),
        'berthed': get_berthed_vessels(window_start, window_end, berths),
        'sailed': get_sailed_vessels(window_start, window_end, berths),
        'berths': berths,
        'expected': get_expected_waiting_vessels(window_start, window_end),
    }

@bp.route('/api/module/RP02/berthplan/data')
@login_required
def rp02_data():
    plan_date = request.args.get('date') or datetime.now().strftime('%Y-%m-%d')
    try:
        report = get_daily_report(plan_date)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    return jsonify(report)

FIELDS_A = [
    ('VIA/F/OVERSEAS/IMO', '__via_f_ovs_imo'), ('VESSEL NAME', 'vessel_name'),
    ('LOA/BHC/DFT', '__loa_dft'), ('AGT / TF/PL', '__agt_tf_pl'), ('Consigner', 'consigner'),
    ('CARGO', 'cargo'), ('QUANTITY', 'quantity'), ('Alongside (Date/Time)', 'alongside'),
    ('OPS COMMENECED', 'ops_commenced'), ('Last 24 hrs Load/Discharge', 'last_24hr_qty'),
    ('Till now discharged / Load', 'till_now_qty'), ('BALANCE', 'balance'),
    ('Expected completion (Date/Time)', 'expected_completion'),
    ('Present Flow Rate(MT/hr)', 'present_flow_rate'), ('Remarks', 'remarks'),
]
FIELDS_B = [
    ('VIA/F/OVERSEAS/IMO', '__via_f_ovs_imo'), ('VESSEL NAME', 'vessel_name'),
    ('LOA/BHC/DFT', '__loa_dft'), ('AGT / TF/PL', '__agt_tf_pl'), ('Consigner', 'consigner'),
    ('CARGO', 'cargo'), ('QUANTITY', 'quantity'), ('Alongside (Date/Time)', 'alongside'),
    ('OPS COMMENECED', 'ops_commenced'), ('Cargo Completion Time', 'cargo_completion'),
    ('Sail Cast off time', 'cast_off'), ('Flow Rate (MT/hr)', 'present_flow_rate'),
]

def _xl_field_value(row, field):
    if row is None:
        return ''
    if field == '__loa_dft':
        loa, dft = row.get('loa'), row.get('draft')
        return f"{loa or ''} / {dft or ''}" if (loa or dft) else ''
    if field == '__via_f_ovs_imo':
        via = row.get('via_no') or ''
        run_type = (row.get('vessel_run_type') or '').strip().lower()

        if run_type == 'foreign':
            ovs = 'Overseas'
        elif run_type == 'coastal':
            ovs = 'Coastal'
        else:
            ovs = ''
        imo = row.get('imo_num') or ''
        return f"{via} / F / {ovs} / {imo}"
    if field == '__agt_tf_pl':
        agt = row.get('vessel_agent') or row.get('agent') or ''
        tf = row.get('terminal') or ''
        pl = row.get('pipeline') or ''
        return f"{agt} / {tf} / {pl}"
    val = row.get(field)
    if val is None:
        return ''
    if field == 'expected_completion' and row.get('etc_is_planned') and val and val != '-':
        return f"{val} (Exp)"
    return val


@bp.route('/api/module/RP02/berthplan/export/excel')
@login_required
def rp02_export_excel():
    plan_date = request.args.get('date') or datetime.now().strftime('%Y-%m-%d')
    report = get_daily_report(plan_date)
    berths = report['berths']

    wb = Workbook()
    ws = wb.active
    ws.title = 'Berth Plan'

    bold = Font(bold=True)
    bold_underline = Font(bold=True, underline='single')   # title lines 1 & 2
    section_font = Font(bold=True, underline='single', color='0000C8')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill('solid', fgColor='DCE6F1')
    label_fill = PatternFill('solid', fgColor='F4F6FA')
    thin = Side(style='thin', color='444444')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # value-cell fonts, matching the webpage's CSS color rules
    FONT_VALUE   = Font(color='0000C8', bold=True)   # normal logged value      -> blue
    FONT_PLANNED = Font(color='B8860B', bold=True)   # planned/no data yet      -> amber
    FONT_EMPTY   = Font(color='CBD5E1', bold=False)  # no vessel at this berth  -> light grey

    r = 1
    title_rows = []
    title_rows.append((r, 'JSW JNPT LIQUID TERMINAL PRIVATE LIMITED', bold_underline)); r += 1
    title_rows.append((r, 'JAWAHARLAL NEHRU PORT AUTHORITY (BULK TERMINAL)', bold_underline)); r += 1
    title_rows.append((r, 'DAILY PERFORMANCE REPORT', bold)); r += 2
    
    

    ws.cell(r, 1, 'As on').font = bold
    date_cell = ws.cell(r, 2, report['as_on_date'])
    date_cell.font = Font(bold=True, color='C00000')
    date_cell.alignment = Alignment(horizontal='center')
    ws.cell(r, 3, '@').font = bold
    time_cell = ws.cell(r, 4, report['as_on_time'])
    time_cell.font = Font(bold=True, color='0000C8')
    time_cell.alignment = Alignment(horizontal='center')
    r += 2
    
    c_headers = ['TERMINAL', 'VESSEL NAME', 'VIA NO.', 'LOA', 'DFT', 'AGT/TNK/CONS', 'CARGO',
                 'MLA', 'QTY.', 'ETA', 'ATA', 'NOR', 'BERTH']



    # ------------------------------------------------------------------
    # Column-width matching logic
    # ------------------------------------------------------------------
    # Section C always has len(c_headers) = 15 physical columns (cols 1-15).
    # Section A/B use column 1 as the "DETAILS" label, so their data needs
    # to fill columns 2..15 (14 columns) for the two tables' right edges
    # to line up exactly - whether those columns hold real vessel data or
    # are empty, they get merged/stretched to fill the same total width
    # as Section C (same as how the webpage stretches columns to 100%).
    #
    # If there are genuinely more concurrent vessels than 14 slots, we
    # expand beyond 14 (never hide real data) - in that rare case Section
    # C will end up narrower than A/B, since C's column count is fixed.
    # ------------------------------------------------------------------

    def group_by_berth(vessels):
        by_berth = {b: [] for b in berths}
        for v in vessels:
            b = v.get('berth_name') or '(no berth)'
            by_berth.setdefault(b, []).append(v)
        return [(b, vs if vs else [None]) for b, vs in by_berth.items()]

    cols_A = group_by_berth(report['berthed'])
    cols_B = group_by_berth(report['sailed'])

    # per-berth minimum columns needed = max vessels at that berth across A & B,
    # so berth columns line up vertically between Section A and Section B too
    mins = [max(len(cols_A[i][1]), len(cols_B[i][1])) for i in range(len(cols_A))]
    total_min = sum(mins) if mins else 0

    target_data_cols = max(len(c_headers) - 1, total_min)   # normally 14
    extra = target_data_cols - total_min
    spans = mins[:]
    if spans:
        i = 0
        while extra > 0:                       # spread leftover width evenly, round-robin
            spans[i % len(spans)] += 1
            extra -= 1
            i += 1

    max_col = max(1 + target_data_cols, len(c_headers))

    def distribute(total, parts):
        """Split `total` columns into `parts` groups, as evenly as possible."""
        if parts <= 0:
            return []
        base, rem = divmod(total, parts)
        return [base + (1 if i < rem else 0) for i in range(parts)]

    fields_map = {
        'A] VESSEL OPERATION :- BERTHED VESSELS': FIELDS_A,
        'B] VESSEL OPERATION :- SAILED VESSELS': FIELDS_B,
    }

    def write_vertical_section(title, cols):
        nonlocal r
        ws.cell(r, 1, title).font = section_font
        r += 1

        ws.cell(r, 1, 'DETAILS').font = bold
        ws.cell(r, 1).fill = header_fill
        ws.cell(r, 1).border = border

        c = 2
        vessel_ranges = []   # (start_col, span) per vessel, in cols order
        for (b, vs), span in zip(cols, spans):
            start_c = c
            for cc in range(start_c, start_c + span):
                cell = ws.cell(r, cc)
                cell.fill = header_fill
                cell.border = border
            ws.cell(r, start_c, b).font = bold
            ws.cell(r, start_c).alignment = center
            if span > 1:
                ws.merge_cells(start_row=r, start_column=start_c,
                                end_row=r, end_column=start_c + span - 1)
            sub_spans = distribute(span, len(vs))
            cc = start_c
            for vs_span in sub_spans:
                vessel_ranges.append((cc, vs_span))
                cc += vs_span
            c = start_c + span
        r += 1

        for label, field in fields_map[title]:
            ws.cell(r, 1, label).font = bold
            ws.cell(r, 1).fill = label_fill
            ws.cell(r, 1).border = border
            idx = 0
            for b, vs in cols:
                for v in vs:
                    start_c, span = vessel_ranges[idx]
                    idx += 1
                    val = _xl_field_value(v, field)
                    cell = ws.cell(r, start_c, val)
                    for cc in range(start_c, start_c + span):
                        ws.cell(r, cc).border = border
                    cell.alignment = center
                    if span > 1:
                        ws.merge_cells(start_row=r, start_column=start_c,
                                        end_row=r, end_column=start_c + span - 1)
                    if v is None:
                        cell.font = FONT_EMPTY
                    elif v.get('is_planned') and field in ('present_flow_rate', 'expected_completion'):
                        cell.font = FONT_PLANNED
                    else:
                        cell.font = FONT_VALUE
            r += 1
        r += 1

    write_vertical_section('A] VESSEL OPERATION :- BERTHED VESSELS', cols_A)
    write_vertical_section('B] VESSEL OPERATION :- SAILED VESSELS', cols_B)

    ws.cell(r, 1, 'C] Expected /Waiting Tank Vessels at JJLTPL Berth').font = section_font
    r += 1
    for col_i, h in enumerate(c_headers, start=1):
            cell = ws.cell(r, col_i, h)
            cell.font = bold; cell.fill = header_fill; cell.alignment = center; cell.border = border
    r += 1

    for row in report['expected']:
        vals = [row.get('terminal'), row.get('vessel_name'), row.get('via_no'), row.get('loa'),
                row.get('dft'), row.get('agt_tnk_cons'), row.get('cargo'), row.get('mla'),
                row.get('quantity'), row.get('eta'), row.get('ata'), 
                row.get('nor'), row.get('berth')]
        for col_i, v in enumerate(vals, start=1):
            cell = ws.cell(r, col_i, v if v is not None else '')
            cell.border = border
            cell.alignment = center
        r += 1

    # merge + center the three title rows now that max_col is final
    for row_idx, text, font in title_rows:
        ws.cell(row_idx, 1, text).font = font
        ws.cell(row_idx, 1).alignment = Alignment(horizontal='center')
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=max_col)

    ws.column_dimensions['A'].width = 26
    for i in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16

    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"Berth_Plan_{plan_date}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )