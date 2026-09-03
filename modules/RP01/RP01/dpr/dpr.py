from .. import bp

from flask import (
    render_template,
    session,
    redirect,
    url_for,
    request,
    jsonify,
    send_file,
)
from functools import wraps
from datetime import datetime, time, timedelta
from io import BytesIO
from database import get_db, get_cursor, get_user_permissions

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MODULE_CODE = 'RP01'


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


def _parse_date(date_str):
    if date_str:
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y'):
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                pass
    return datetime.now().date()


def _get_window(selected_date):
    """
    7 AM to 7 AM daily window logic:
    If selected_date is 28-08-2026, window is:
    window_start = 27-08-2026 07:00:00
    window_end   = 28-08-2026 07:00:00
    """
    window_end = datetime.combine(selected_date, time(7, 0, 0))
    window_start = window_end - timedelta(days=1)
    return window_start, window_end


def _get_month_window(selected_date, window_end):
    month_start = datetime.combine(selected_date.replace(day=1), time(7, 0, 0))
    return month_start, window_end


def _get_fy_info(selected_date):
    if selected_date.month >= 4:
        fy_start_year = selected_date.year
    else:
        fy_start_year = selected_date.year - 1
    fy_end_short = (fy_start_year + 1) % 100
    fy_label = f"{fy_start_year}-{fy_end_short:02d}"
    fy_start_dt = datetime(fy_start_year, 4, 1, 7, 0, 0)
    return fy_label, fy_start_dt


def _fmt_dt_compact(dt_val):
    """Format datetime as DDMMYYYY HHMM e.g. 24082026 1800"""
    if not dt_val:
        return ''
    if isinstance(dt_val, str):
        s = dt_val.strip()
        for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M'):
            try:
                dt_val = datetime.strptime(s[:16], fmt[:len(s)])
                break
            except Exception:
                pass
    if isinstance(dt_val, datetime):
        return dt_val.strftime('%d-%m-%Y %H:%M')
    return str(dt_val)



def _fmt_eta(dt_val):
    """Format ETA as 20-Aug-26 14.0"""
    if not dt_val:
        return ''
    if isinstance(dt_val, str):
        s = dt_val.strip()
        for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%d-%m-%Y %H:%M'):
            try:
                dt_val = datetime.strptime(s[:16], fmt[:len(s)])
                break
            except Exception:
                pass
    if isinstance(dt_val, datetime):
        return dt_val.strftime('%d-%b-%y %H.%M')
    return str(dt_val)


def _parse_entry_dt(entry_date, from_time=None):
    """Safely parse entry_date (text or date) + optional from_time into a datetime object without SQL casting errors."""
    if not entry_date:
        return None

    base_dt = None
    if isinstance(entry_date, datetime):
        base_dt = entry_date
    elif hasattr(entry_date, 'year') and hasattr(entry_date, 'month') and not hasattr(entry_date, 'hour'):
        base_dt = datetime.combine(entry_date, time(0, 0))
    else:
        s = str(entry_date).strip()
        for fmt in (
            '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
            '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M',
            '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d', '%d.%m.%Y'
        ):
            try:
                base_dt = datetime.strptime(s, fmt)
                break
            except Exception:
                continue
        if not base_dt:
            try:
                base_dt = datetime.fromisoformat(s)
            except Exception:
                try:
                    d = datetime.strptime(s[:10], '%Y-%m-%d').date()
                    base_dt = datetime.combine(d, time(0, 0))
                except Exception:
                    return None

    if from_time:
        try:
            parts = str(from_time).strip().split(':')
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                return datetime.combine(base_dt.date(), time(int(parts[0]), int(parts[1])))
        except Exception:
            pass

    return base_dt


def _load_vessel_cargo_map(cur):
    """Load cargo master mapping dynamically from database category columns across all 3 tables."""
    master_map = {}

    def resolve_category(cat_val, fallback_name=""):
        s = str(cat_val or '').strip().upper()
        if not s:
            s = str(fallback_name or '').strip().upper()
        if 'EDIBLE' in s:
            return 'Edible Oil'
        if 'POL' in s:
            return 'POL'
        if 'CHEM' in s:
            return 'Chemical'
        return 'Other liquid'

    try:
        # 1. vessel_cargo (master table) -> check cargo_sub_category_2 FIRST, then cargo_category, cargo_type
        cur.execute("""
            SELECT LOWER(TRIM(cargo_name)) AS norm_name, cargo_sub_category_2, cargo_category, cargo_type
            FROM vessel_cargo
            WHERE cargo_name IS NOT NULL AND TRIM(cargo_name) != ''
        """)
        for r in cur.fetchall():
            primary = r.get('cargo_sub_category_2') or r.get('cargo_category') or r.get('cargo_type')
            master_map[r['norm_name']] = resolve_category(primary, r['norm_name'])

        # 2. mis_vessel_master -> use new_cat, category1, category
        cur.execute("""
            SELECT LOWER(TRIM(cargo)) AS norm_name, new_cat, category1, category
            FROM mis_vessel_master
            WHERE cargo IS NOT NULL AND TRIM(cargo) != ''
        """)
        for r in cur.fetchall():
            cn = r['norm_name']
            if cn not in master_map:
                primary = r.get('new_cat') or r.get('category1') or r.get('category')
                master_map[cn] = resolve_category(primary, cn)

        # 3. mis_history -> use cargo_sub_category_2, cargo_category, cargo_type
        cur.execute("""
            SELECT LOWER(TRIM(cargo_name)) AS norm_name, cargo_sub_category_2, cargo_category, cargo_type
            FROM mis_history
            WHERE cargo_name IS NOT NULL AND TRIM(cargo_name) != ''
        """)
        for r in cur.fetchall():
            cn = r['norm_name']
            if cn not in master_map:
                primary = r.get('cargo_sub_category_2') or r.get('cargo_category') or r.get('cargo_type')
                master_map[cn] = resolve_category(primary, cn)

        return master_map
    except Exception:
        return master_map


def _categorize_cargo(cargo_name, cargo_sub2=None, cargo_cat=None, cargo_sub=None, vc_map=None):
    """
    Map cargo to Section B categories: Edible Oil, Other liquid, Chemical, POL.
    In mis_history and master tables, cargo_sub_category_2 has the exact specific bucket
    (CHEMICAL, EDIBLE OIL, FARM LIQUIDS, OTHER LIQUID, POL).
    FARM LIQUIDS and OTHER LIQUID combine into 'Other liquid'.
    """
    # 1. First priority: explicit category fields (cargo_sub2 checked first, e.g. cargo_sub_category_2 or new_cat)
    for val in (cargo_sub2, cargo_cat, cargo_sub):
        if not val:
            continue
        s = str(val).strip().upper()
        if 'EDIBLE' in s:
            return 'Edible Oil'
        if 'CHEM' in s:
            return 'Chemical'
        if 'POL' in s:
            return 'POL'
        if 'FARM' in s or 'OTHER' in s:
            return 'Other liquid'

    # 2. Second priority: master mapping by cargo name
    c_norm = str(cargo_name or '').strip().lower()
    if vc_map and c_norm in vc_map:
        mapped = vc_map[c_norm]
        if mapped in ('Edible Oil', 'Other liquid', 'Chemical', 'POL'):
            return mapped

    # 3. Third priority: cargo name keywords
    s_name = str(cargo_name or '').strip().upper()
    if 'EDIBLE' in s_name or 'PALM' in s_name or 'SOYA' in s_name or 'SUNFLOWER' in s_name or 'CPO' in s_name or 'CDSBO' in s_name:
        return 'Edible Oil'
    if 'POL' in s_name or 'MOTOR SPIRIT' in s_name or 'HSD' in s_name or 'NAPHTHA' in s_name or 'DIESEL' in s_name or 'PETROL' in s_name:
        return 'POL'
    if 'CHEM' in s_name or 'BENZENE' in s_name or 'HEXANE' in s_name or 'METHANOL' in s_name or 'ETHANOL' in s_name or 'ACETONE' in s_name:
        return 'Chemical'

    return 'Other liquid'


def _get_section_a(cur, window_start, window_end):
    """
    Section A: Vessel handling at JJLTPL
    Shows vessels alongside or active at berth slots during 7 AM to 7 AM window.
    A vessel stays on berth if alongside <= window_end AND (cast_off IS NULL OR cast_off > window_end).
    If cast_off <= window_end, the vessel has cast off and is removed from Section A.

    Status logic:
      - If ANY parcel of the vessel is working / in progress: Status = "Operations"
      - If ALL parcels are completed (and no cast_off yet): Status = "Completed"
    """
    cur.execute("""
        SELECT
            vh.id AS vh_id,
            vh.berth_name AS berth,
            vh.vessel_name,
            vh.via_number AS via,
            po.id AS parcel_op_id,
            po.parcel_ids,
            COALESCE(po.cargo_name, vh.cargo_type, 'Liquid Bulk') AS cargo,
            po.quantity AS op_qty,
            po.start_dt,
            po.end_dt,
            po.expected_start,
            COALESCE(po.expected_flow_rate, 0) AS expected_flow_rate,
            lh.alongside_datetime,
            lh.cast_off_datetime
        FROM vcn_header vh
        JOIN ldud_header lh ON lh.vcn_id = vh.id
        LEFT JOIN ldud_parcel_ops po ON po.ldud_id = lh.id
        WHERE (vh.berth_name LIKE '%%LB%%' OR vh.berth_name LIKE '%%03%%' OR vh.berth_name LIKE '%%04%%')
          AND NULLIF(lh.alongside_datetime,'') IS NOT NULL
          AND NULLIF(lh.alongside_datetime,'')::timestamp <= %s
          AND (
              NULLIF(lh.cast_off_datetime,'') IS NULL
              OR NULLIF(lh.cast_off_datetime,'')::timestamp > %s
          )
        ORDER BY vh.berth_name, vh.id, po.id
    """, (window_end, window_end))

    raw_rows = cur.fetchall()

    grouped = {}
    vessel_order = []

    for r in raw_rows:
        vh_id = r['vh_id']
        berth_raw = (r.get('berth') or 'LB03').replace('-', '').strip()
        v_name = (r.get('vessel_name') or '').strip()
        key = (vh_id, berth_raw, v_name)

        if key not in grouped:
            grouped[key] = {
                'vh_id': vh_id,
                'berth': berth_raw,
                'vessel_name': v_name,
                'cargo_list': [],
                'total_bl_qty': 0.0,
                'commenced_times': [],
                'completed_times': [],
                'parcel_statuses': [],
            }
            vessel_order.append(key)

        g = grouped[key]

        c_name = (r.get('cargo') or '').strip()
        if c_name and c_name not in g['cargo_list']:
            g['cargo_list'].append(c_name)

        p_qty = 0.0
        if r.get('op_qty'):
            try:
                p_qty = float(str(r.get('op_qty')).replace(',', '').strip())
            except Exception:
                p_qty = 0.0
        elif r.get('parcel_ids'):
            try:
                cur.execute("""
                    SELECT SUM(q.qty) AS tot FROM (
                        SELECT NULLIF(regexp_replace(COALESCE(quantity, '0'), '[^0-9.]', '', 'g'), '')::numeric AS qty FROM vcn_consigners WHERE vcn_id = %s
                        UNION ALL
                        SELECT NULLIF(regexp_replace(COALESCE(quantity, '0'), '[^0-9.]', '', 'g'), '')::numeric AS qty FROM vcn_export_cargo_declaration WHERE vcn_id = %s
                    ) q
                """, (vh_id, vh_id))
                q_row = cur.fetchone()
                if q_row and q_row.get('tot'):
                    p_qty = float(q_row['tot'])
            except Exception:
                p_qty = 0.0

        g['total_bl_qty'] += p_qty

        start_time = _parse_entry_dt(r.get('start_dt'))
        if not start_time and r.get('alongside_datetime'):
            start_time = _parse_entry_dt(r.get('alongside_datetime'))
        if start_time:
            g['commenced_times'].append(start_time)

        comp_time = _parse_entry_dt(r.get('end_dt'))
        exp_start = _parse_entry_dt(r.get('expected_start'))
        pid = r.get('parcel_op_id')
        real_qty = 0.0
        hours = 0.0

        if pid:
            try:
                cur.execute("""
                    SELECT from_time, to_time, COALESCE(quantity, 0) AS q, is_shortclose, remarks
                    FROM lueu_parcel_log
                    WHERE parcel_op_id = %s AND is_deleted IS NOT TRUE
                    ORDER BY entry_date, from_time NULLS LAST, id
                """, (pid,))
                log_rows = cur.fetchall()
                for lr in log_rows:
                    if lr.get('is_shortclose') or 'short' in str(lr.get('remarks') or '').lower():
                        continue
                    real_qty += float(lr.get('q') or 0)
                    try:
                        fh, fm = (int(x) for x in str(lr['from_time']).split(':')[:2])
                        th, tm = (int(x) for x in str(lr['to_time']).split(':')[:2])
                        m_diff = (th * 60 + tm) - (fh * 60 + fm)
                        if m_diff < 0:
                            m_diff += 1440
                        hours += m_diff / 60.0
                    except Exception:
                        pass
            except Exception:
                pass

        remaining = max(p_qty - real_qty, 0.0) if p_qty > 0 else 0.0

        if comp_time or (p_qty > 0 and remaining <= 1e-6):
            g['parcel_statuses'].append('Completed')
            if comp_time:
                g['completed_times'].append(comp_time)
        else:
            g['parcel_statuses'].append('Operations')
            if start_time and hours > 0 and real_qty > 0 and p_qty > 0:
                avg_rate = real_qty / hours
                if avg_rate > 0:
                    comp_time = start_time + timedelta(hours=(p_qty / avg_rate))
            elif exp_start and r.get('expected_flow_rate') and float(r['expected_flow_rate']) > 0 and p_qty > 0:
                comp_time = exp_start + timedelta(hours=(p_qty / float(r['expected_flow_rate'])))
            if comp_time:
                g['completed_times'].append(comp_time)

    vessels = []
    for key in vessel_order:
        g = grouped[key]

        min_commenced = min(g['commenced_times']) if g['commenced_times'] else None
        max_completed = max(g['completed_times']) if g['completed_times'] else None

        if any(st == 'Operations' for st in g['parcel_statuses']):
            v_status = 'Operations'
        else:
            v_status = 'Completed'

        vessels.append({
            'berth': g['berth'],
            'vessel_name': g['vessel_name'],
            'cargo': " / ".join(g['cargo_list']) if g['cargo_list'] else 'Liquid Bulk',
            'bl_quantity': round(g['total_bl_qty'], 3),
            'discharge_commenced': _fmt_dt_compact(min_commenced),
            'discharge_completed': _fmt_dt_compact(max_completed),
            'status': v_status,
        })


    # Ensure 4 columns as in template (LB03, LB04, LB03, LB04)
    result = []
    default_berths = ['LB03', 'LB04', 'LB03', 'LB04']
    for idx in range(4):
        if idx < len(vessels):
            result.append(vessels[idx])
        else:
            result.append({
                'berth': default_berths[idx],
                'vessel_name': '',
                'cargo': '',
                'bl_quantity': 0.0,
                'discharge_commenced': '',
                'discharge_completed': '',
                'status': ''
            })

    return result




def _get_section_b(cur, selected_date, window_start, window_end):
    """
    Section B: Cargo Handled
    Row categories: Edible Oil, Phosphoric, Lube Oil, Chemicals, POL
    Rows:
    1. Balance on Board to be unloaded/Loaded
    2. Cargo discharged during Day
    3. Cargo discharge in this Month
    4. Cum handled FY 2026-27
    5. Cum handled FY 2025-26
    6. Cum handled FY 2024-25
    7. Cum Loading/unloading till date
    """
    vc_map = _load_vessel_cargo_map(cur)
    categories = ['Edible Oil', 'Other liquid', 'Chemical', 'POL']
    fy_label, _ = _get_fy_info(selected_date)
    month_start, _ = _get_month_window(selected_date, window_end)
    month_name = selected_date.strftime('%b %Y')

    grid = {
        'balance_on_board': {c: 0.0 for c in categories},
        'day_discharged': {c: 0.0 for c in categories},
        'month_discharged': {c: 0.0 for c in categories},
        'cum_fy_curr': {c: 0.0 for c in categories},
        'cum_fy_2025_26': {c: 0.0 for c in categories},
        'cum_fy_2024_25': {c: 0.0 for c in categories},
        'cum_till_date': {c: 0.0 for c in categories},
    }

    # 1. Balance on Board — ONLY vessels currently on berth (LB-03/LB-04, alongside but not cast off)
    cur.execute("""
        SELECT
            po.id AS parcel_op_id,
            COALESCE(po.cargo_name, vh.cargo_type, '') AS cargo_name,
            COALESCE(po.quantity, 0) AS op_qty,
            vc.cargo_category,
            vc.cargo_sub_category,
            vc.cargo_sub_category_2
        FROM ldud_parcel_ops po
        JOIN ldud_header lh ON lh.id = po.ldud_id
        JOIN vcn_header vh ON vh.id = lh.vcn_id
        LEFT JOIN vessel_cargo vc ON LOWER(TRIM(COALESCE(po.cargo_name, ''))) = LOWER(TRIM(vc.cargo_name))
        WHERE COALESCE(lh.is_deleted, FALSE) = FALSE
          AND (vh.berth_name LIKE '%%LB%%' OR vh.berth_name LIKE '%%03%%' OR vh.berth_name LIKE '%%04%%')
          AND NULLIF(lh.alongside_datetime, '') IS NOT NULL
          AND NULLIF(lh.alongside_datetime, '')::timestamp <= %s
          AND (
              NULLIF(lh.cast_off_datetime, '') IS NULL
              OR NULLIF(lh.cast_off_datetime, '')::timestamp > %s
          )
    """, (window_end, window_end))
    active_rows = cur.fetchall()
    for r in active_rows:
        pid = r.get('parcel_op_id')
        c_name = r.get('cargo_name')
        p_qty = 0.0
        if r.get('op_qty'):
            try:
                p_qty = float(str(r.get('op_qty')).replace(',', '').strip())
            except Exception:
                p_qty = 0.0

        real_qty = 0.0
        if pid:
            try:
                cur.execute("""
                    SELECT SUM(COALESCE(quantity, 0)) AS q
                    FROM lueu_parcel_log
                    WHERE parcel_op_id = %s
                      AND COALESCE(is_deleted, FALSE) = FALSE
                      AND COALESCE(is_shortclose, FALSE) = FALSE
                      AND LOWER(COALESCE(remarks, '')) NOT LIKE '%%short%%close%%'
                """, (pid,))
                lr = cur.fetchone()
                if lr and lr.get('q'):
                    real_qty = float(lr['q'])
            except Exception:
                pass

        rem = max(p_qty - real_qty, 0.0)
        cat = _categorize_cargo(c_name, r.get('cargo_category'), r.get('cargo_sub_category'), r.get('cargo_sub_category_2'), vc_map=vc_map)
        if cat not in categories:
            cat = 'Other liquid'
        grid['balance_on_board'][cat] += rem

    # 2 & 3. Cargo discharged during Day and Month from lueu_parcel_log (safe Python datetime parsing)
    cur.execute("""
        SELECT
            po.cargo_name,
            COALESCE(log.quantity, 0) AS qty,
            log.entry_date,
            log.from_time,
            log.to_time,
            log.remarks,
            log.is_shortclose
        FROM lueu_parcel_log log
        JOIN ldud_parcel_ops po ON po.id = log.parcel_op_id
        JOIN ldud_header lh ON lh.id = po.ldud_id
        WHERE COALESCE(log.is_deleted, FALSE) = FALSE
          AND COALESCE(log.is_shortclose, FALSE) = FALSE
          AND COALESCE(lh.is_deleted, FALSE) = FALSE
          AND NULLIF(TRIM(log.entry_date), '') IS NOT NULL
    """)
    for r in cur.fetchall():
        if r.get('is_shortclose') or 'short' in str(r.get('remarks') or '').lower():
            continue
        entry_dt = _parse_entry_dt(r.get('entry_date'), r.get('from_time'))
        if not entry_dt:
            continue
        c_name = r.get('cargo_name')
        q = float(r.get('qty') or 0)
        cat = _categorize_cargo(c_name, vc_map=vc_map)
        if cat not in categories:
            cat = 'Other liquid'

        if window_start <= entry_dt < window_end:
            grid['day_discharged'][cat] += q

        if month_start <= entry_dt < window_end:
            grid['month_discharged'][cat] += q

    # 4. Cum handled FY 2026-27
    # Part 4A: Reconciled MIS data from mis_vessel_master (covers Apr to Jun 2026)
    cur.execute("""
        SELECT
            mvm.cargo,
            mvm.new_cat,
            mvm.category1,
            mvm.category,
            SUM(COALESCE(mvm.quantity, 0)) AS qty
        FROM mis_vessel_master mvm
        WHERE mvm.fin_year = '2026-27'
        GROUP BY mvm.cargo, mvm.new_cat, mvm.category1, mvm.category
    """)
    for r in cur.fetchall():
        q = float(r.get('qty') or 0)
        cat = _categorize_cargo(r.get('cargo'), cargo_sub2=r.get('new_cat'), cargo_cat=r.get('category1'), cargo_sub=r.get('category'), vc_map=vc_map)
        if cat not in categories:
            cat = 'Other liquid'
        grid['cum_fy_curr'][cat] += q

    # Part 4B: Live operational data (from ldud_header, ldud_parcel_ops, lueu_parcel_log)
    # Includes vessels with cast_off_datetime after June (from Jul 1 onwards) up to window_end.
    # Uses vessel_cargo.cargo_sub_category_2 to classify the cargo.
    if selected_date.month < 4:
        live_start = datetime(selected_date.year - 1, 7, 1, 7, 0, 0)
    else:
        live_start = datetime(selected_date.year, 7, 1, 7, 0, 0)

    cur.execute("""
        SELECT
            lh.id AS ldud_id,
            lh.vessel_name,
            lh.cast_off_datetime,
            po.cargo_name,
            vc.cargo_sub_category_2,
            vc.cargo_category,
            vc.cargo_type,
            SUM(COALESCE(log.quantity, 0)) AS qty
        FROM ldud_header lh
        JOIN ldud_parcel_ops po ON po.ldud_id = lh.id
        JOIN lueu_parcel_log log ON log.parcel_op_id = po.id
        LEFT JOIN vessel_cargo vc ON LOWER(TRIM(vc.cargo_name)) = LOWER(TRIM(po.cargo_name))
        WHERE NULLIF(TRIM(lh.cast_off_datetime::text), '') IS NOT NULL
          AND COALESCE(log.is_deleted, FALSE) = FALSE
          AND COALESCE(log.is_shortclose, FALSE) = FALSE
          AND LOWER(COALESCE(log.remarks, '')) NOT LIKE '%%short%%'
          AND COALESCE(lh.is_deleted, FALSE) = FALSE
        GROUP BY
            lh.id,
            lh.vessel_name,
            lh.cast_off_datetime,
            po.cargo_name,
            vc.cargo_sub_category_2,
            vc.cargo_category,
            vc.cargo_type
    """)
    for r in cur.fetchall():
        cast_off_dt = _parse_entry_dt(r.get('cast_off_datetime'))
        if not cast_off_dt:
            continue

        if live_start <= cast_off_dt <= window_end:
            q = float(r.get('qty') or 0)
            if q <= 0:
                continue
            cat = _categorize_cargo(
                r.get('cargo_name'),
                cargo_sub2=r.get('cargo_sub_category_2'),
                cargo_cat=r.get('cargo_category'),
                cargo_sub=r.get('cargo_type'),
                vc_map=vc_map
            )
            if cat not in categories:
                cat = 'Other liquid'
            grid['cum_fy_curr'][cat] += q

    # 5. Cum handled FY 2025-26 (from mis_history)
    cur.execute("""
        SELECT
            mh.cargo_name,
            mh.cargo_sub_category_2,
            mh.cargo_category,
            mh.cargo_type,
            SUM(COALESCE(mh.quantity, 0)) AS qty
        FROM mis_history mh
        WHERE mh.fin_year = '2025-26'
        GROUP BY mh.cargo_name, mh.cargo_sub_category_2, mh.cargo_category, mh.cargo_type
    """)
    for r in cur.fetchall():
        q = float(r.get('qty') or 0)
        cat = _categorize_cargo(r.get('cargo_name'), cargo_sub2=r.get('cargo_sub_category_2'), cargo_cat=r.get('cargo_category'), cargo_sub=r.get('cargo_type'), vc_map=vc_map)
        if cat not in categories:
            cat = 'Other liquid'
        grid['cum_fy_2025_26'][cat] += q

    # 6. Cum handled FY 2024-25 (from mis_history)
    cur.execute("""
        SELECT
            mh.cargo_name,
            mh.cargo_sub_category_2,
            mh.cargo_category,
            mh.cargo_type,
            SUM(COALESCE(mh.quantity, 0)) AS qty
        FROM mis_history mh
        WHERE mh.fin_year = '2024-25'
        GROUP BY mh.cargo_name, mh.cargo_sub_category_2, mh.cargo_category, mh.cargo_type
    """)
    for r in cur.fetchall():
        q = float(r.get('qty') or 0)
        cat = _categorize_cargo(r.get('cargo_name'), cargo_sub2=r.get('cargo_sub_category_2'), cargo_cat=r.get('cargo_category'), cargo_sub=r.get('cargo_type'), vc_map=vc_map)
        if cat not in categories:
            cat = 'Other liquid'
        grid['cum_fy_2024_25'][cat] += q

    # 7. Cum Loading/unloading till date = FY 2026-27 + FY 2025-26 + FY 2024-25
    for c in categories:
        grid['cum_till_date'][c] = (
            grid['cum_fy_curr'][c] +
            grid['cum_fy_2025_26'][c] +
            grid['cum_fy_2024_25'][c]
        )

    # Compute result rows formatted for Section B
    result_rows = []
    row_specs = [
        ('Balance on Board to be unloaded/Loaded', 'balance_on_board'),
        ('Cargo discharged during Day', 'day_discharged'),
        (f'Cargo discharge in this Month ({month_name})', 'month_discharged'),
        (f'Cum handled FY {fy_label}', 'cum_fy_curr'),
        ('Cum handled FY 2025-26', 'cum_fy_2025_26'),
        ('Cum handled FY 2024-25', 'cum_fy_2024_25'),
        ('Cum Loading/unloading till date', 'cum_till_date'),
    ]

    for label, key in row_specs:
        row_dict = {'particulars': label}
        row_total = 0.0
        for cat in categories:
            val = round(grid[key].get(cat, 0.0), 3)
            row_dict[cat] = val
            row_total += val
        row_dict['total'] = round(row_total, 3)
        result_rows.append(row_dict)

    return result_rows


def _get_section_c(cur, selected_date):
    """
    Section C: Vessels Expected
    Dynamic fetch EXCLUSIVELY from EV01 module (expected_vessels table).
    Shows all active / pending expected vessels from EV01.
    Columns: Vessel, Cargo, Berth, B/L Qty, EXIM, Agent, ETA, Remarks
    """
    cur.execute("""
        SELECT
            vessel_name,
            COALESCE(cargo_name, '') AS cargo,
            COALESCE(berth_name, 'LB03/04') AS berth,
            quantity AS bl_qty_str,
            'Import' AS exim,
            COALESCE(agents, '') AS agent,
            eta,
            remarks
        FROM expected_vessels
        WHERE (doc_status IS NULL OR LOWER(doc_status) NOT LIKE '%%closed%%')
        ORDER BY eta ASC NULLS LAST, id
    """)

    ev_rows = cur.fetchall()

    vessels = []
    seen = set()

    for r in ev_rows:
        v_name = (r.get('vessel_name') or '').strip()
        if not v_name or v_name.upper() in seen:
            continue
        seen.add(v_name.upper())

        qty = 0.0
        if r.get('bl_qty_str'):
            try:
                raw_str = str(r.get('bl_qty_str')).replace(',', ' ').strip()
                parts = [float(p) for p in raw_str.split() if p.replace('.', '', 1).isdigit()]
                qty = sum(parts) if parts else 0.0
            except Exception:
                qty = 0.0

        vessels.append({
            'vessel': v_name,
            'cargo': r.get('cargo') or '',
            'berth': r.get('berth') or 'LB03/04',
            'bl_qty': round(qty, 3),
            'exim': r.get('exim') or 'Import',
            'agent': r.get('agent') or '',
            'eta': _fmt_eta(r.get('eta')),
            'remarks': (r.get('remarks') or '').strip(),
        })

    return vessels


def _get_section_d(cur, selected_date):
    """
    Section D: Vessels Handled / operations / declared in [Selected Month Year]
    Dynamic query combining:
    1. Vessels currently running (in operation / alongside) on the selected date window
    2. Vessels completed in the selected month
    Columns: Sr No, Vessel, Cargo, Customer, Quantity, Berth, Agent, Remarks
    """
    window_start, window_end = _get_window(selected_date)
    month_start = datetime.combine(selected_date.replace(day=1), time(7, 0, 0))
    if selected_date.month == 12:
        next_month_start = datetime(selected_date.year + 1, 1, 1, 7, 0, 0)
    else:
        next_month_start = datetime(selected_date.year, selected_date.month + 1, 1, 7, 0, 0)

    month_str = selected_date.strftime('%b-%y')      # e.g. Aug-26
    month_str_long = selected_date.strftime('%b-%Y') # e.g. Aug-2026
    month_name = selected_date.strftime('%B %Y')     # e.g. August 2026

    vessels = []
    seen = set()
    sr = 1

    # 1. RUNNING VESSELS on selected date window
    cur.execute("""
        SELECT
            vh.id AS vh_id,
            vh.vessel_name,
            COALESCE(po.cargo_name, vh.cargo_type, 'Liquid Bulk') AS cargo,
            COALESCE(vh.vessel_agent_name, '') AS agent,
            COALESCE(vh.berth_name, 'LB-03') AS berth,
            lh.alongside_datetime,
            lh.cast_off_datetime
        FROM vcn_header vh
        JOIN ldud_header lh ON lh.vcn_id = vh.id
        LEFT JOIN ldud_parcel_ops po ON po.ldud_id = lh.id
        WHERE NULLIF(lh.alongside_datetime,'') IS NOT NULL
          AND NULLIF(lh.alongside_datetime,'')::timestamp <= %s
          AND (NULLIF(lh.cast_off_datetime,'') IS NULL OR NULLIF(lh.cast_off_datetime,'')::timestamp > %s)
        ORDER BY vh.berth_name, vh.id
    """, (window_end, window_start))

    running_rows = cur.fetchall()
    for r in running_rows:
        v_name = (r.get('vessel_name') or '').strip()
        if not v_name or v_name.upper() in seen:
            continue
        seen.add(v_name.upper())

        vh_id = r.get('vh_id')
        customer_names = []
        qty = 0.0

        if vh_id:
            try:
                cur.execute("""
                    SELECT STRING_AGG(DISTINCT NULLIF(TRIM(c.consigner_name), ''), ', ') AS custs,
                           SUM(NULLIF(regexp_replace(COALESCE(c.quantity, '0'), '[^0-9.]', '', 'g'), '')::numeric) AS tot
                    FROM (
                        SELECT consigner_name, quantity FROM vcn_consigners WHERE vcn_id = %s
                        UNION ALL
                        SELECT consigner_name, quantity FROM vcn_export_cargo_declaration WHERE vcn_id = %s
                    ) c
                """, (vh_id, vh_id))
                c_row = cur.fetchone()
                if c_row:
                    if c_row.get('custs'):
                        customer_names.append(c_row['custs'])
                    if c_row.get('tot'):
                        qty = float(c_row['tot'])
            except Exception:
                pass

        try:
            cur.execute("""
                SELECT SUM(COALESCE(log.quantity, 0)) AS log_qty
                FROM lueu_parcel_log log
                JOIN ldud_parcel_ops po ON po.id = log.parcel_op_id
                JOIN ldud_header lh ON lh.id = po.ldud_id
                WHERE lh.vcn_id = %s AND log.is_deleted IS NOT TRUE
            """, (vh_id,))
            log_r = cur.fetchone()
            if log_r and log_r.get('log_qty') and float(log_r['log_qty']) > 0:
                qty = float(log_r['log_qty'])
        except Exception:
            pass

        has_cast_off = bool(r.get('cast_off_datetime') and str(r.get('cast_off_datetime')).strip())
        has_alongside = bool(r.get('alongside_datetime') and str(r.get('alongside_datetime')).strip())
        display_name = f"{v_name} (In progress)" if (has_alongside and not has_cast_off) else v_name

        vessels.append({
            'sr_no': sr,
            'vessel': display_name,
            'cargo': r.get('cargo') or '',
            'customer': ', '.join(customer_names) if customer_names else 'Multiple Customers',
            'quantity': round(qty, 3),
            'berth': r.get('berth') or 'LB-03',
            'agent': r.get('agent') or '',
            'remarks': '',
            'is_running': not has_cast_off,
        })
        sr += 1

    # 2. LIVE COMPLETED VESSELS in selected month
    cur.execute("""
        SELECT
            vh.id AS vh_id,
            vh.vessel_name,
            COALESCE(po.cargo_name, vh.cargo_type, 'Liquid Bulk') AS cargo,
            COALESCE(vh.vessel_agent_name, '') AS agent,
            COALESCE(vh.berth_name, 'LB-03') AS berth
        FROM vcn_header vh
        JOIN ldud_header lh ON lh.vcn_id = vh.id
        LEFT JOIN ldud_parcel_ops po ON po.ldud_id = lh.id
        WHERE NULLIF(lh.cast_off_datetime,'') IS NOT NULL
          AND NULLIF(lh.cast_off_datetime,'')::timestamp >= %s
          AND NULLIF(lh.cast_off_datetime,'')::timestamp < %s
        ORDER BY lh.cast_off_datetime ASC
    """, (month_start, next_month_start))

    completed_rows = cur.fetchall()
    for r in completed_rows:
        v_name = (r.get('vessel_name') or '').strip()
        if not v_name or v_name.upper() in seen:
            continue
        seen.add(v_name.upper())

        vh_id = r.get('vh_id')
        customer_names = []
        qty = 0.0

        if vh_id:
            try:
                cur.execute("""
                    SELECT STRING_AGG(DISTINCT NULLIF(TRIM(c.consigner_name), ''), ', ') AS custs,
                           SUM(NULLIF(regexp_replace(COALESCE(c.quantity, '0'), '[^0-9.]', '', 'g'), '')::numeric) AS tot
                    FROM (
                        SELECT consigner_name, quantity FROM vcn_consigners WHERE vcn_id = %s
                        UNION ALL
                        SELECT consigner_name, quantity FROM vcn_export_cargo_declaration WHERE vcn_id = %s
                    ) c
                """, (vh_id, vh_id))
                c_row = cur.fetchone()
                if c_row:
                    if c_row.get('custs'):
                        customer_names.append(c_row['custs'])
                    if c_row.get('tot'):
                        qty = float(c_row['tot'])
            except Exception:
                pass

        rem_val = (r.get('remarks') or '').strip()
        if rem_val.lower() == 'completed':
            rem_val = ''

        vessels.append({
            'sr_no': sr,
            'vessel': v_name,
            'cargo': r.get('cargo') or '',
            'customer': ', '.join(customer_names) if customer_names else 'Multiple Customers',
            'quantity': round(qty, 3),
            'berth': r.get('berth') or 'LB-03',
            'agent': r.get('agent') or '',
            'remarks': rem_val,
        })
        sr += 1

    # 3. HISTORICAL COMPLETED VESSELS from mis_vessel_master in selected month
    cur.execute("""
        SELECT
            m.vessel_name,
            m.cargo,
            m.consigner AS customer,
            m.quantity,
            m.berth_no AS berth,
            m.agent,
            m.remarks
        FROM mis_vessel_master m
        WHERE m.month = %s OR m.month = %s
        ORDER BY m.sr_no NULLS LAST, m.id
    """, (month_str, month_str_long))

    hist_rows = cur.fetchall()
    for r in hist_rows:
        v_name = (r.get('vessel_name') or '').strip()
        if not v_name or v_name.upper() in seen:
            continue
        seen.add(v_name.upper())

        h_rem = (r.get('remarks') or '').strip()
        if h_rem.lower() == 'completed':
            h_rem = ''

        vessels.append({
            'sr_no': sr,
            'vessel': v_name,
            'cargo': r.get('cargo') or '',
            'customer': r.get('customer') or '',
            'quantity': float(r.get('quantity') or 0.0),
            'berth': r.get('berth') or 'LB-03',
            'agent': r.get('agent') or '',
            'remarks': h_rem,
        })
        sr += 1

    total_qty = sum(v['quantity'] for v in vessels)

    return {
        'month_name': month_name,
        'vessels': vessels,
        'total_quantity': round(total_qty, 3)
    }



def _get_dpr_payload(selected_date):
    window_start, window_end = _get_window(selected_date)
    reporting_date = selected_date - timedelta(days=1)

    conn = get_db()
    try:
        cur = get_cursor(conn)
        section_a = _get_section_a(cur, window_start, window_end)
        section_b = _get_section_b(cur, selected_date, window_start, window_end)
        section_c = _get_section_c(cur, selected_date)
        section_d = _get_section_d(cur, selected_date)
    finally:
        conn.close()

    return {
        'report_date': selected_date.strftime('%Y-%m-%d'),
        'report_date_display': selected_date.strftime('%d %B %Y'),
        'reporting_date_display': reporting_date.strftime('%d %B %Y'),
        'window_start_display': window_start.strftime('%d-%m-%Y %H:%M'),
        'window_end_display': window_end.strftime('%d-%m-%Y %H:%M'),
        'section_a': section_a,
        'section_b': section_b,
        'section_c': section_c,
        'section_d': section_d,
    }


# ══════════════════════════════════════════════════════════════════
# Routes
# ══════════════════════════════════════════════════════════════════

@bp.route('/module/RP01/dpr/')
@login_required
def dpr_page():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('dpr/dpr.html', permissions=perms)


@bp.route('/api/module/RP01/dpr/data')
@login_required
def dpr_data():
    selected_date = _parse_date(request.args.get('date'))
    try:
        return jsonify(_get_dpr_payload(selected_date))
    except Exception as e:
        import traceback, sys
        tb = traceback.format_exc()
        print(f"[DPR ERROR] {e}\n{tb}", file=sys.stderr, flush=True)
        return jsonify({'error': str(e), 'traceback': tb}), 500


@bp.route('/api/module/RP01/dpr/export')
@login_required
def dpr_export():
    selected_date = _parse_date(request.args.get('date'))
    payload = _get_dpr_payload(selected_date)

    wb = Workbook()
    ws = wb.active
    ws.title = "DPR"

    FONT_NAME = "Arial"
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    fill_header = PatternFill("solid", fgColor="BCD6EE")
    fill_section = PatternFill("solid", fgColor="DDEBF7")
    fill_month = PatternFill("solid", fgColor="FFF2A8")
    fill_total = PatternFill("solid", fgColor="FCE0CD")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    fill_title = PatternFill("solid", fgColor="1F4E78")

    font_title = Font(name=FONT_NAME, size=12, bold=True, color="FFFFFF")
    font_header = Font(name=FONT_NAME, bold=True, size=10)
    font_section = Font(name=FONT_NAME, bold=True, size=10)
    font_normal = Font(name=FONT_NAME, size=10)
    font_value = Font(name=FONT_NAME, size=10, color="1F4E78")
    font_total = Font(name=FONT_NAME, bold=True, size=10)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    NUM_FMT = '#,##0.000;(#,##0.000);"-"'

    def set_cell(coord, value, font=font_normal, fill=None, align=center, fmt=None):
        c = ws[coord]
        c.value = value
        c.font = font
        c.border = border
        c.alignment = align
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt
        return c

    def merge(rng, value=None, font=font_normal, fill=None, align=center, fmt=None):
        ws.merge_cells(rng)
        top_left = rng.split(":")[0]
        set_cell(top_left, value, font=font, fill=fill, align=align, fmt=fmt)
        for row in ws[rng]:
            for cell in row:
                cell.border = border
                if fill:
                    cell.fill = fill

    # Column widths (A to H)
    widths = {
        "A": 36, "B": 22, "C": 18, "D": 18,
        "E": 18, "F": 18, "G": 22, "H": 26
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Title Bar (A1:H1)
    merge("A1:H1", "DAILY REPORT - JSW JNPT LIQUID TERMINAL PRIVATE LIMITED (JJLTPL)",
          font=font_title, fill=fill_title, align=center)

    # Meta dates strip (A3:H3)
    set_cell("A3", "Report Date", font=font_header, fill=fill_header, align=left)
    set_cell("B3", payload['report_date_display'], font=font_value, fill=fill_white, align=center)
    merge("C3:E3", "", fill=fill_white)
    set_cell("F3", "Reporting Date", font=font_header, fill=fill_header, align=left)
    merge("G3:H3", payload['reporting_date_display'], font=font_value, fill=fill_white, align=center)

    r = 5
    # Section A
    merge(f"A{r}:H{r}", "Section A: Vessel handling at JJLTPL", font=font_section, fill=fill_section, align=left)

    r += 1
    sec_a_fields = [
        ("Berth", [v['berth'] for v in payload['section_a']]),
        ("Vessel Name", [v['vessel_name'] for v in payload['section_a']]),
        ("Cargo", [v['cargo'] for v in payload['section_a']]),
        ("BL Quantity", [v['bl_quantity'] for v in payload['section_a']]),
        ("Discharge Commenced", [v['discharge_commenced'] for v in payload['section_a']]),
        ("Discharge Completed (ETC)", [v['discharge_completed'] for v in payload['section_a']]),
        ("Status", [v['status'] for v in payload['section_a']]),
    ]

    for label, vals in sec_a_fields:
        set_cell(f"A{r}", label, font=font_header, fill=fill_header, align=left)
        for idx in range(7):  # Columns B through H
            col_let = get_column_letter(2 + idx)
            val = vals[idx] if idx < len(vals) else None
            is_bl = (label == "BL Quantity") and isinstance(val, (int, float))
            set_cell(f"{col_let}{r}", val, font=font_value if val is not None else font_normal,
                     fill=fill_white, align=right if is_bl else center,
                     fmt=NUM_FMT if is_bl else None)
        r += 1

    # Section B
    r += 1
    merge(f"A{r}:H{r}", "Section B: Cargo Handled", font=font_section, fill=fill_section, align=left)

    r += 1
    b_headers = ["Particulars", "Edible Oil", "Other liquid", "Chemical", "POL", "Total"]
    for i, h in enumerate(b_headers):
        col_let = get_column_letter(1 + i)
        set_cell(f"{col_let}{r}", h, font=font_header, fill=fill_header, align=left if i == 0 else right)
    set_cell(f"G{r}", None, fill=fill_white)
    set_cell(f"H{r}", None, fill=fill_white)

    for row_data in payload['section_b']:
        r += 1
        is_total = "till date" in row_data['particulars'].lower()
        is_month = "month" in row_data['particulars'].lower()
        row_fill = fill_total if is_total else (fill_month if is_month else fill_white)
        row_font = font_total if is_total else font_normal

        set_cell(f"A{r}", row_data['particulars'], font=row_font, fill=row_fill, align=left)

        cats = ["Edible Oil", "Other liquid", "Chemical", "POL", "total"]
        for idx, cat in enumerate(cats):
            col_let = get_column_letter(2 + idx)
            val = row_data.get(cat, 0.0)
            cell_font = font_total if (cat == "total" or is_total) else font_value
            cell_fill = fill_total if (cat == "total" or is_total) else row_fill
            set_cell(f"{col_let}{r}", val, font=cell_font, fill=cell_fill, align=right, fmt=NUM_FMT)

        set_cell(f"G{r}", None, fill=fill_white)
        set_cell(f"H{r}", None, fill=fill_white)

    # Section C
    r += 2
    merge(f"A{r}:H{r}", "Section C: Vessels Expected", font=font_section, fill=fill_section, align=left)

    r += 1
    c_headers = ["Vessel", "Cargo", "Berth", "B/L Qty", "EXIM", "Agent", "ETA", "Remarks"]
    for i, h in enumerate(c_headers):
        col_let = get_column_letter(1 + i)
        set_cell(f"{col_let}{r}", h, font=font_header, fill=fill_header, align=right if i == 3 else (left if i in (0, 1, 5, 7) else center))

    if payload['section_c']:
        for v in payload['section_c']:
            r += 1
            vals = [v['vessel'], v['cargo'], v['berth'], v['bl_qty'], v['exim'], v['agent'], v['eta'], v['remarks']]
            for i, val in enumerate(vals):
                col_let = get_column_letter(1 + i)
                is_num = (i == 3 and isinstance(val, (int, float)))
                set_cell(f"{col_let}{r}", val, font=font_value, fill=fill_white,
                         align=right if is_num else (left if i in (0, 1, 5, 7) else center),
                         fmt=NUM_FMT if is_num else None)
    else:
        r += 1
        merge(f"A{r}:H{r}", "No expected vessels for selected month", font=font_normal, fill=fill_white, align=center)

    # Section D
    r += 2
    sec_d = payload['section_d']
    merge(f"A{r}:H{r}", f"Section D: Vessels Handled / operations / declared in {sec_d['month_name']}",
          font=font_section, fill=fill_section, align=left)

    r += 1
    d_headers = ["Vessel", "Cargo", "Customer", "Quantity", "Berth", "Agent", "Remarks"]
    for i, h in enumerate(d_headers):
        col_let = get_column_letter(1 + i)
        set_cell(f"{col_let}{r}", h, font=font_header, fill=fill_header, align=right if i == 3 else (left if i in (0, 1, 2, 5, 6) else center))
    set_cell(f"H{r}", None, fill=fill_white)

    if sec_d['vessels']:
        for v in sec_d['vessels']:
            r += 1
            vals = [v['vessel'], v['cargo'], v['customer'], v['quantity'], v['berth'], v['agent'], v['remarks']]
            for i, val in enumerate(vals):
                col_let = get_column_letter(1 + i)
                is_num = (i == 3 and isinstance(val, (int, float)))
                set_cell(f"{col_let}{r}", val, font=font_value, fill=fill_white,
                         align=right if is_num else (left if i in (0, 1, 2, 5, 6) else center),
                         fmt=NUM_FMT if is_num else None)
            set_cell(f"H{r}", None, fill=fill_white)
    else:
        r += 1
        merge(f"A{r}:G{r}", "No vessels handled in selected month", font=font_normal, fill=fill_white, align=center)
        set_cell(f"H{r}", None, fill=fill_white)

    # Total Row for Section D
    r += 1
    merge(f"A{r}:C{r}", "Total", font=font_total, fill=fill_total, align=left)
    set_cell(f"D{r}", sec_d['total_quantity'], font=font_total, fill=fill_total, align=right, fmt=NUM_FMT)
    merge(f"E{r}:H{r}", "", fill=fill_total)

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = True

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"Daily_Performance_Report_{payload['report_date']}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
