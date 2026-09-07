from database import get_db, get_cursor
from datetime import datetime
# Target resolution lives in LDUD01 (owner of ldud_parcel_ops) so every
# screen computes the BL/target figure the same way.
from modules.LDUD01.model import (parcel_source_table, source_quantities,
                                  effective_target)

# parcel_ids on ldud_parcel_ops point at the VCN's parcel source table,
# chosen by the linked VCN's operation_type (whitelisted — safe to interpolate).
def _parse_ids(csv):
    return [int(x) for x in str(csv or '').split(',') if str(x).strip().isdigit()]


def _num(v):
    if v is None or (isinstance(v, str) and v.strip() == ''):
        return None
    return v


def _hours(f, t):
    """Duration in hours between two 'HH:MM' strings (wraps past midnight)."""
    try:
        fh, fm = (int(x) for x in str(f).split(':')[:2])
        th, tm = (int(x) for x in str(t).split(':')[:2])
    except (ValueError, AttributeError):
        return 0.0
    mins = (th * 60 + tm) - (fh * 60 + fm)
    if mins < 0:
        mins += 1440
    return mins / 60.0


def get_vessels_with_started_parcels():
    """Vessels that have any LDUD parcel-ops rows. Parcel start/end is entered
    here in LUEU01, so vessels appear as soon as parcels exist (not gated on
    start_dt)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT h.id AS vcn_id, h.vcn_doc_num, h.vessel_name, h.berth_name,
               COUNT(po.id) AS parcel_count
        FROM ldud_parcel_ops po
        JOIN ldud_header l ON l.id = po.ldud_id
        JOIN vcn_header h ON h.id = l.vcn_id
        GROUP BY h.id, h.vcn_doc_num, h.vessel_name, h.berth_name
        ORDER BY h.vcn_doc_num DESC
    ''')
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_started_parcels(vcn_id):
    """Each parcel-ops row (parcel + terminal) for the vessel. The per-row target
    is ldud_parcel_ops.quantity; remaining = target - logged. start/end are
    entered in LUEU01."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT po.id AS parcel_op_id, po.parcel_ids, po.cargo_name, po.terminal_name,
               po.quantity AS op_qty, po.start_dt, po.end_dt, po.expected_start,
               po.expected_flow_rate, po.additional_qty, po.additional_reason,
               l.alongside_datetime, l.doc_status AS ldud_status
        FROM ldud_parcel_ops po
        JOIN ldud_header l ON l.id = po.ldud_id
        WHERE l.vcn_id = %s
        ORDER BY po.id
    ''', [vcn_id])
    parcels = [dict(r) for r in cur.fetchall()]

    # resolve parcel_no label + CURRENT quantity from the operation-type source
    # table, so the validation target tracks VCN updates (falls back to the
    # parcel-op's own quantity when the parcels can't be resolved).
    cur.execute('SELECT operation_type FROM vcn_header WHERE id=%s', [vcn_id])
    row = cur.fetchone()
    is_export = (row or {}).get('operation_type') == 'Export'
    tbl = parcel_source_table((row or {}).get('operation_type') if row else None)
    # export parcels mirror import since jnpa35 — same columns on both tables
    all_ids = sorted({pid for p in parcels for pid in _parse_ids(p['parcel_ids'])})
    labels, src_qty, src_equip, src_pipe, src_term = {}, {}, {}, {}, {}
    src_removed, src_rm_reason = {}, {}
    if all_ids:
        cur.execute(f'''SELECT id, parcel_no, quantity AS q, equipment_names AS equip,
                               pipeline_name AS pipe, unload_terminal AS term,
                               COALESCE(is_removed, FALSE) AS is_removed,
                               removed_reason
                        FROM {tbl} WHERE id = ANY(%s)''', [all_ids])
        for r in cur.fetchall():
            src_removed[r['id']] = bool(r['is_removed'])
            src_rm_reason[r['id']] = r['removed_reason'] or ''
            labels[r['id']] = r['parcel_no'] or f"#{r['id']}"
            src_equip[r['id']] = r['equip'] or ''
            src_pipe[r['id']] = r['pipe'] or ''
            src_term[r['id']] = r['term'] or ''
            # A removed parcel contributes zero — mirrors LDUD01.source_quantities,
            # which keeps the id present (0.0) rather than omitting it, so the
            # op-snapshot fallback cannot resurrect the removed quantity.
            try:
                src_qty[r['id']] = (0.0 if r['is_removed'] else
                                    (float(str(r['q']).replace(',', '')) if r['q'] is not None else 0.0))
            except (ValueError, TypeError):
                src_qty[r['id']] = 0.0

    # per-parcel target — shared rule, see LDUD01.effective_target
    targets = {p['parcel_op_id']: effective_target(src_qty, _parse_ids(p['parcel_ids']),
                                                   p['op_qty'], p['additional_qty'])
               for p in parcels}

    # logged qty + operating hours per parcel (non-deleted), for total & avg flow rate.
    # ponytail: hardcoded completion cap — once cumulative qty reaches the target,
    # later log rows (top-ups, idle entries) are dropped from Run hours so they
    # can't drag the actual ETC out. Rows must be ordered for the cap to apply.
    pop_ids = [p['parcel_op_id'] for p in parcels]
    agg = {}  # parcel_op_id -> [logged_qty, hours]
    if pop_ids:
        cur.execute('''SELECT parcel_op_id, from_time, to_time, COALESCE(quantity,0) AS q, is_shortclose
                       FROM lueu_parcel_log
                       WHERE parcel_op_id = ANY(%s) AND is_deleted IS NOT TRUE
                       ORDER BY parcel_op_id, entry_date, from_time NULLS LAST, id''', [pop_ids])
        for r in cur.fetchall():
            a = agg.setdefault(r['parcel_op_id'], [0.0, 0.0, 0.0])  # [real_qty, hours, shortclose_qty]
            tgt = targets.get(r['parcel_op_id'], 0)
            if tgt > 0 and (a[0] + a[2]) >= tgt - 1e-6:
                continue  # parcel already complete — ignore this row
            if r['is_shortclose']:
                a[2] += float(r['q'] or 0)  # counts toward completion, NOT toward avg rate
            else:
                a[0] += float(r['q'] or 0)
                a[1] += _hours(r['from_time'], r['to_time'])
    conn.close()

    out = []
    for p in parcels:
        ids = _parse_ids(p['parcel_ids'])
        target = targets[p['parcel_op_id']]
        logged_real, hours, shortclosed = agg.get(p['parcel_op_id'], [0.0, 0.0, 0.0])
        logged = logged_real + shortclosed  # total toward target (Remaining)
        # distinct equipment / pipelines across the VCN parcel(s)
        def _distinct(src):
            vals = []
            for i in ids:
                for x in str(src.get(i, '')).split(','):
                    if x.strip() and x.strip() not in vals:
                        vals.append(x.strip())
            return vals
        equip = _distinct(src_equip)
        # terminal from the live VCN consigner (unload_terminal); fall back to the
        # op snapshot (covers export parcels + legacy rows without a source terminal)
        terminals = _distinct(src_term)
        out.append({
            'parcel_op_id': p['parcel_op_id'],
            'parcel_no': ', '.join(labels.get(i, f"#{i}") for i in ids) or '—',
            'cargo_name': p['cargo_name'] or '',
            'terminal_name': ', '.join(terminals) or (p['terminal_name'] or ''),
            'target_qty': round(target, 3),
            'logged_qty': round(logged, 3),
            'remaining_qty': round(target - logged, 3),
            'op_hours': round(hours, 2),
            'avg_rate': round(logged_real / hours, 2) if hours > 0 else 0,
            'is_shortclosed': shortclosed > 1e-6,
            'additional_qty': round(_num(p['additional_qty']) or 0, 3),
            'additional_reason': p['additional_reason'] or '',
            # Removed parcels stay on screen, greyed and restorable.
            'is_removed': all(src_removed.get(i, False) for i in ids) if ids else False,
            'removed_reason': next((src_rm_reason.get(i) for i in ids
                                    if src_removed.get(i)), ''),
            'parcel_ids': ids,
            'uom': 'MT',
            'equipment_names': ', '.join(equip),
            'pipeline_name': ', '.join(_distinct(src_pipe)),
            'expected_start': p['expected_start'],
            'expected_flow_rate': _num(p['expected_flow_rate']),
            'alongside_datetime': p['alongside_datetime'],
            'ldud_status': p['ldud_status'],
            'start_dt': p['start_dt'],
            'end_dt': p['end_dt'],
            'status': 'Completed' if p['end_dt'] else 'In Progress',
        })
    return out


# Master export layout: (Excel header, field). 'dt:' fields split into Date + Time
# columns — see excel_export.
EXPORT_COLS = [
    ('Vessel', 'vessel_name'), ('VCN', 'vcn_doc_num'), ('Berth', 'berth_name'),
    ('Parcel No', 'parcel_no'), ('Cargo', 'cargo_name'), ('Terminal', 'terminal_name'),
    ('Equipment', 'equipment_names'), ('Pipeline', 'pipeline_name'),
    ('Target Qty', 'target_qty'), ('Logged Qty', 'logged_qty'), ('Remaining Qty', 'remaining_qty'),
    ('UOM', 'uom'), ('Expected Start', 'dt:expected_start'), ('Expected Rate (MT/Hr)', 'expected_flow_rate'),
    ('Start', 'dt:start_dt'), ('End', 'dt:end_dt'), ('Run Hours', 'op_hours'),
    ('Avg Rate (MT/Hr)', 'avg_rate'), ('Status', 'status'), ('Short-closed', 'is_shortclosed'),
    ('Additional Qty', 'additional_qty'), ('Additional Reason', 'additional_reason'),
    ('Removed', 'is_removed'), ('Removed Reason', 'removed_reason'),
    ('LDUD Status', 'ldud_status'),
]


def export_all_parcels():
    """Every parcel-op across all vessels, flattened for the master export.
    ponytail: loops get_started_parcels per vessel rather than one hand-tuned
    query — same numbers as the screen, no second source of truth. Fold into a
    single query if the vessel list ever grows past a few hundred."""
    rows = []
    for v in get_vessels_with_started_parcels():
        for p in get_started_parcels(v['vcn_id']):
            rows.append({**p, 'vessel_name': v['vessel_name'],
                         'vcn_doc_num': v['vcn_doc_num'], 'berth_name': v['berth_name']})
    return rows


def ops_locked(parcel_op_ids):
    """True if any parcel op belongs to a fully Closed LDUD — LUEU01 entry is
    then locked. (Partial Close is a billing cut-off; ops may continue.)"""
    ids = [int(i) for i in parcel_op_ids if i]
    if not ids:
        return False
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT 1 FROM ldud_parcel_ops po
                   JOIN ldud_header l ON l.id = po.ldud_id
                   WHERE po.id = ANY(%s) AND l.doc_status = 'Closed' LIMIT 1''', [ids])
    locked = cur.fetchone() is not None
    conn.close()
    return locked


def logs_locked(log_ids):
    """True if any logbook row belongs to a fully Closed LDUD."""
    ids = [int(i) for i in log_ids if i]
    if not ids:
        return False
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT 1 FROM lueu_parcel_log lg
                   JOIN ldud_parcel_ops po ON po.id = lg.parcel_op_id
                   JOIN ldud_header l ON l.id = po.ldud_id
                   WHERE lg.id = ANY(%s) AND l.doc_status = 'Closed' LIMIT 1''', [ids])
    locked = cur.fetchone() is not None
    conn.close()
    return locked


def set_expected_start(parcel_op_id, expected_start, expected_flow_rate=None):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('UPDATE ldud_parcel_ops SET expected_start=%s, expected_flow_rate=%s WHERE id=%s',
                [expected_start or None, _num(expected_flow_rate), parcel_op_id])
    conn.commit()
    conn.close()


def set_parcel_times(parcel_op_id, start_dt, end_dt):
    """Operators enter the parcel start/end here; persisted on ldud_parcel_ops."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('UPDATE ldud_parcel_ops SET start_dt=%s, end_dt=%s WHERE id=%s',
                [start_dt or None, end_dt or None, parcel_op_id])
    conn.commit()
    conn.close()


_LOG_COLS = ['parcel_op_id', 'entry_date', 'from_time', 'to_time', 'quantity',
             'pressure', 'quantity_uom', 'medium', 'equipment_name', 'delay_name',
             'shift', 'operator_name', 'shift_incharge', 'berth_name', 'remarks']


def get_log(parcel_op_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT * FROM lueu_parcel_log
                   WHERE parcel_op_id=%s AND is_deleted IS NOT TRUE
                   ORDER BY entry_date, from_time, id''', [parcel_op_id])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def save_log(data):
    # Direct Pipe carries no equipment
    if data.get('medium') == 'Direct Pipe':
        data['equipment_name'] = None
    data['quantity'] = _num(data.get('quantity'))
    data['pressure'] = _num(data.get('pressure'))
    conn = get_db()
    cur = get_cursor(conn)
    if data.get('id'):
        sets = ', '.join(f'{c}=%s' for c in _LOG_COLS)
        cur.execute(f'UPDATE lueu_parcel_log SET {sets} WHERE id=%s',
                    [data.get(c) for c in _LOG_COLS] + [data['id']])
        row_id = data['id']
    else:
        cols = _LOG_COLS + ['created_by', 'created_date']
        vals = [data.get(c) for c in _LOG_COLS] + [data.get('created_by'),
                                                   datetime.now().strftime('%Y-%m-%d')]
        ph = ', '.join(['%s'] * len(cols))
        cur.execute(f'INSERT INTO lueu_parcel_log ({", ".join(cols)}) VALUES ({ph}) RETURNING id', vals)
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def soft_delete_log(ids, username):
    conn = get_db()
    cur = get_cursor(conn)
    today = datetime.now().strftime('%Y-%m-%d')
    for log_id in ids:
        cur.execute('''UPDATE lueu_parcel_log
                       SET is_deleted=TRUE, deleted_by=%s, deleted_date=%s
                       WHERE id=%s AND is_deleted IS NOT TRUE''', [username, today, log_id])
    conn.commit()
    conn.close()


def _single_parcel_target(cur, parcel_op_id):
    """Target qty for one parcel-op — same shared rule as get_started_parcels."""
    cur.execute('''SELECT po.parcel_ids, po.quantity AS op_qty, po.additional_qty,
                          h.operation_type
                   FROM ldud_parcel_ops po
                   JOIN ldud_header l ON l.id = po.ldud_id
                   LEFT JOIN vcn_header h ON h.id = l.vcn_id
                   WHERE po.id=%s''', [parcel_op_id])
    row = cur.fetchone()
    if not row:
        return 0.0
    ids = _parse_ids(row['parcel_ids'])
    src_qty = source_quantities(cur, parcel_source_table(row['operation_type']), ids)
    return effective_target(src_qty, ids, row['op_qty'], row['additional_qty'])


def set_additional_qty(parcel_op_id, qty, reason, username):
    """Amend the BL upward for one parcel-op.

    The extra tonnage is already in the logbook — the operator logged it hour by
    hour and the completion cap discarded it. This records the amendment so the
    target rises and those real rows start counting; nothing is invented, and
    the VCN's declared quantity is left alone.

    Reversible via clear_additional_qty."""
    qty = _num(qty)
    if not qty or qty <= 0:
        raise ValueError('Additional quantity must be greater than zero')
    reason = (reason or '').strip()
    if not reason:
        raise ValueError('A reason is required for an additional quantity')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE ldud_parcel_ops
                   SET additional_qty=%s, additional_reason=%s,
                       additional_by=%s, additional_date=%s
                   WHERE id=%s""",
                [qty, reason, username, datetime.now().strftime('%Y-%m-%d'), parcel_op_id])
    n = cur.rowcount
    conn.commit()
    conn.close()
    if not n:
        raise ValueError('Parcel operation not found')
    return qty


def clear_additional_qty(parcel_op_id, username):
    """Undo an additional-quantity amendment; the target drops back to the
    declared BL and the over-logged rows go back to being ignored."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""UPDATE ldud_parcel_ops
                   SET additional_qty=0, additional_reason=NULL,
                       additional_by=NULL, additional_date=NULL
                   WHERE id=%s AND COALESCE(additional_qty, 0) <> 0""", [parcel_op_id])
    n = cur.rowcount
    conn.commit()
    conn.close()
    if not n:
        raise ValueError('No additional quantity to clear')
    return n


def shortclose_parcel(parcel_op_id, username):
    """Close a parcel's leftover quantity: insert one flagged, timeless log row
    carrying the remaining qty so Remaining -> 0. Raises ValueError if nothing
    is left to close. Reversible via revert_shortclose."""
    conn = get_db()
    cur = get_cursor(conn)
    target = _single_parcel_target(cur, parcel_op_id)
    cur.execute('''SELECT COALESCE(SUM(quantity), 0) AS q FROM lueu_parcel_log
                   WHERE parcel_op_id=%s AND is_deleted IS NOT TRUE''', [parcel_op_id])
    logged = float(cur.fetchone()['q'] or 0)
    remaining = round(target - logged, 3)
    if remaining <= 1e-6:
        conn.close()
        raise ValueError('Nothing to short-close — no remaining quantity')
    today = datetime.now().strftime('%Y-%m-%d')
    cur.execute('''INSERT INTO lueu_parcel_log
                   (parcel_op_id, entry_date, quantity, quantity_uom, is_shortclose,
                    remarks, created_by, created_date)
                   VALUES (%s, %s, %s, 'MT', TRUE, 'Short close', %s, %s) RETURNING id''',
                [parcel_op_id, today, remaining, username, today])
    row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def revert_shortclose(parcel_op_id, username):
    """Undo a short-close: soft-delete the parcel's short-close row(s), restoring
    the previous Remaining and avg rate. Raises ValueError if there is none."""
    conn = get_db()
    cur = get_cursor(conn)
    today = datetime.now().strftime('%Y-%m-%d')
    cur.execute('''UPDATE lueu_parcel_log
                   SET is_deleted=TRUE, deleted_by=%s, deleted_date=%s
                   WHERE parcel_op_id=%s AND is_shortclose IS TRUE AND is_deleted IS NOT TRUE''',
                [username, today, parcel_op_id])
    n = cur.rowcount
    conn.commit()
    conn.close()
    if not n:
        raise ValueError('No short-close to revert')
    return n
