from database import get_db, get_cursor

def _clean_empty(data):
    """Convert empty strings to None so timestamp/date columns get NULL."""
    for k in data:
        if data[k] == '':
            data[k] = None
    return data

def get_next_doc_num():
    import datetime
    conn = get_db()
    cur = get_cursor(conn)
    # Use financial year suffix: FY starting April, so Mar→prev year pair
    now = datetime.datetime.now()
    fy_start = now.year if now.month >= 4 else now.year - 1
    fy_suffix = f"{str(fy_start)[2:]}{str(fy_start + 1)[2:]}"  # e.g. "2526"
    prefix = f"LDUD-{fy_suffix}-"
    cur.execute(
        "SELECT MAX(CAST(SPLIT_PART(doc_num, '-', 3) AS INTEGER)) FROM ldud_header WHERE doc_num LIKE %s",
        (prefix + '%',)
    )
    result = cur.fetchone()['max']
    conn.close()
    next_num = (result or 0) + 1
    return f"{prefix}{next_num:03d}"

def _build_vcn_list(rows):
    result = []
    for r in rows:
        display = f"{r['vcn_doc_num']} / {r['vessel_name']}"
        result.append({
            'value': display,
            'vcn_id': r['id'],
            'vcn_doc_num': r['vcn_doc_num'],
            'vessel_name': r['vessel_name'],
            'anchored_datetime': r.get('anchorage_arrival'),
            'doc_date': r.get('doc_date') or '',
            'operation_type': r.get('operation_type') or ''
        })
    return result

def get_vcn_list():
    """Get all approved VCN entries with doc date and operation type for dropdown"""
    conn = get_db()
    cur = get_cursor(conn)
    # One LDUD per VCN — drop VCNs already linked to an LDUD from the picker.
    cur.execute('''
        SELECT h.id, h.vcn_doc_num, h.vessel_name, h.doc_date, h.operation_type, a.anchorage_arrival
        FROM vcn_header h
        LEFT JOIN vcn_anchorage a ON a.vcn_id = h.id
        WHERE h.doc_status = 'Approved'
          AND h.id NOT IN (SELECT vcn_id FROM ldud_header WHERE vcn_id IS NOT NULL)
        ORDER BY h.vcn_doc_num DESC
    ''')
    rows = cur.fetchall()
    conn.close()
    return _build_vcn_list(rows)

def get_data(page=1, size=20, filters=None):
    conn = get_db()
    cur = get_cursor(conn)

    allowed = {'doc_num','vessel_name','doc_status','doc_date','vcn_doc_num',
               'operation_type','cargo_type'}
    # VCN doc numbers are editable in VCN01, so ldud_header.vcn_doc_num is only a
    # snapshot taken at link time. Display and filtering both use the live value.
    live_vcn_doc = ("COALESCE((SELECT h.vcn_doc_num FROM vcn_header h WHERE h.id = ldud_header.vcn_id),"
                    " ldud_header.vcn_doc_num)")
    col_sql = {'vcn_doc_num': live_vcn_doc}
    # soft-deleted LDUDs (VCN sent back to Expected) stay hidden
    where_clauses, params = ['is_deleted IS NOT TRUE'], []
    for f in (filters or []):
        field = f.get('field', '')
        if field not in allowed:
            continue
        field = col_sql.get(field, field)
        ftype = f.get('type')
        if ftype == 'contains' and f.get('value'):
            where_clauses.append(f"{field} ILIKE %s")
            params.append(f"%{f['value']}%")
        elif ftype == 'multi' and f.get('values'):
            ph = ','.join(['%s'] * len(f['values']))
            where_clauses.append(f"{field} IN ({ph})")
            params.extend(f['values'])
        elif ftype == 'range':
            if f.get('from'):
                where_clauses.append(f"{field} >= %s")
                params.append(f['from'])
            if f.get('to'):
                where_clauses.append(f"{field} <= %s")
                params.append(f['to'])

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""
    try:
        cur.execute(f'SELECT COUNT(*) FROM ldud_header {where_sql}', params)
        total = cur.fetchone()['count']
        cur.execute(f'SELECT * FROM ldud_header {where_sql} ORDER BY id DESC LIMIT %s OFFSET %s',
                    params + [size, (page - 1) * size])
        rows = [dict(r) for r in cur.fetchall()]

        # Collect vcn_ids to batch-fetch computed fields
        vcn_ids = list(set(r['vcn_id'] for r in rows if r.get('vcn_id')))

        # Short-closed (written-off) quantity per LDUD — deducted from the BL total shown
        shortclose = {}
        if rows:
            cur.execute('''SELECT po.ldud_id, COALESCE(SUM(lg.quantity), 0) AS sc
                           FROM lueu_parcel_log lg
                           JOIN ldud_parcel_ops po ON po.id = lg.parcel_op_id
                           WHERE po.ldud_id = ANY(%s) AND lg.is_shortclose IS TRUE
                             AND lg.is_deleted IS NOT TRUE
                           GROUP BY po.ldud_id''', ([r['id'] for r in rows],))
            shortclose = {s['ldud_id']: float(s['sc'] or 0) for s in cur.fetchall()}

        vcn_cargo = {}   # vcn_id -> {cargo_names, bl_quantities}
        vcn_agents = {}  # vcn_id -> {agent_name, stevedore_name}
        vcn_meta = {}    # vcn_id -> {doc_date}

        if vcn_ids:
            # Fetch doc_date for display
            cur.execute('''SELECT id, doc_date, doc_status, vcn_doc_num, operation_type
                           FROM vcn_header WHERE id = ANY(%s)''', (vcn_ids,))
            for v in cur.fetchall():
                vcn_meta[v['id']] = {'doc_date': v['doc_date'] or '', 'doc_status': v['doc_status'] or '',
                                     'vcn_doc_num': v['vcn_doc_num'],
                                     'operation_type': v['operation_type'] or ''}

            # Cargo names, BL quantities and UOM from VCN cargo declarations (Import + Export)
            cur.execute('''SELECT vcn_id, cargo_name, bl_quantity, quantity_uom FROM vcn_cargo_declaration
                           WHERE vcn_id = ANY(%s) AND cargo_name IS NOT NULL''', (vcn_ids,))
            import_cargo = cur.fetchall()
            # Export cargo now mirrors the import consigner shape: quantity (TEXT), no UOM
            cur.execute('''SELECT vcn_id, cargo_name, quantity FROM vcn_export_cargo_declaration
                           WHERE vcn_id = ANY(%s) AND cargo_name IS NOT NULL''', (vcn_ids,))
            export_cargo = []
            for c in cur.fetchall():
                try:
                    qty = float(str(c['quantity']).replace(',', '')) if c['quantity'] else 0.0
                except ValueError:
                    qty = 0.0
                export_cargo.append({'vcn_id': c['vcn_id'], 'cargo_name': c['cargo_name'],
                                     'bl_quantity': qty, 'quantity_uom': 'MT'})
            # Import cargo now lives in the VCN consigner (IGM line) table
            cur.execute('''SELECT vcn_id, cargo_name, quantity FROM vcn_consigners
                           WHERE vcn_id = ANY(%s) AND cargo_name IS NOT NULL''', (vcn_ids,))
            consigner_cargo = []
            for c in cur.fetchall():
                try:
                    qty = float(str(c['quantity']).replace(',', '')) if c['quantity'] else 0.0
                except ValueError:
                    qty = 0.0
                consigner_cargo.append({'vcn_id': c['vcn_id'], 'cargo_name': c['cargo_name'],
                                        'bl_quantity': qty, 'quantity_uom': 'MT'})
            # Legacy vcn_cargo_declaration is historic data only — once a VCN has
            # consigner (IGM line) parcels those are the truth, so drop the legacy
            # rows for it rather than counting both.
            _with_consigners = {c['vcn_id'] for c in consigner_cargo}
            import_cargo = [c for c in import_cargo if c['vcn_id'] not in _with_consigners]
            # One entry per parcel/cargo line — do NOT sum same-named cargo
            # (e.g. two EDIBLE OIL parcels stay as two separate lines).
            # Import and export parcels live in different tables and a VCN owns
            # exactly one side; flipping operation_type leaves the old side's rows
            # behind, so summing both double-counts. operation_type decides.
            for row_list, side in ((import_cargo, 'Import'), (consigner_cargo, 'Import'),
                                   (export_cargo, 'Export')):
                for c in row_list:
                    vid = c['vcn_id']
                    own_side = 'Export' if vcn_meta.get(vid, {}).get('operation_type') == 'Export' else 'Import'
                    if side != own_side:
                        continue
                    if vid not in vcn_cargo:
                        vcn_cargo[vid] = {'names': [], 'quantities': [], 'uoms': []}
                    vcn_cargo[vid]['names'].append(c['cargo_name'])
                    vcn_cargo[vid]['quantities'].append(float(c['bl_quantity'] or 0))
                    vcn_cargo[vid]['uoms'].append(c['quantity_uom'] or '')

            # Agent, Stevedore and meta from VCN header
            cur.execute('''SELECT id, vessel_agent_name, importer_exporter_name
                           FROM vcn_header WHERE id = ANY(%s)''', (vcn_ids,))
            for v in cur.fetchall():
                vcn_agents[v['id']] = {
                    'agent_name': v['vessel_agent_name'],
                    'stevedore_name': v['importer_exporter_name']
                }

        # Enrich rows
        for r in rows:
            vid = r.get('vcn_id')

            # Cargo info from VCN
            ci = vcn_cargo.get(vid, {'names': [], 'quantities': [], 'uoms': []})
            uoms = ci.get('uoms', [])
            r['cargo_names_display'] = ', '.join(ci['names']) if ci['names'] else ''
            # ponytail: single total, first non-empty UOM wins (mixed UOMs are not a real case here)
            if ci['quantities']:
                uom = next((u for u in uoms if u), '')
                total = sum(ci['quantities']) - shortclose.get(r['id'], 0.0)
                r['bl_quantities_display'] = f"{total:.3f} {uom}".strip()
            else:
                r['bl_quantities_display'] = ''

            # VCN doc date for display
            vm = vcn_meta.get(vid, {})
            r['vcn_doc_date'] = vm.get('doc_date', '')
            r['vcn_doc_status'] = vm.get('doc_status', '')
            # renamed in VCN01 after linking → show the current number, don't rewrite the row
            if vm.get('vcn_doc_num'):
                r['vcn_doc_num'] = vm['vcn_doc_num']

            # Agent and Stevedore
            ai = vcn_agents.get(vid, {})
            r['agent_name'] = ai.get('agent_name', '')
            r['stevedore_name'] = ai.get('stevedore_name', '')

        return rows, total
    finally:
        conn.close()

def save_header(data):
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')

    # Convert empty strings to None so timestamp/date columns get NULL
    for k in data:
        if data[k] == '':
            data[k] = None

    # One-to-one: a VCN may be linked to at most one LDUD (DB-enforced too).
    if data.get('vcn_id'):
        cur.execute('SELECT id FROM ldud_header WHERE vcn_id=%s AND id IS DISTINCT FROM %s',
                    [data['vcn_id'], row_id])
        if cur.fetchone():
            conn.close()
            raise ValueError('This VCN is already linked to another LDUD (one LDUD per VCN).')

    if row_id:
        _computed = {'id', 'doc_num', 'vcn_display', 'vcn_doc_date', 'vcn_doc_status', 'cargo_names_display', 'bl_quantities_display',
                     'balance_display', 'agent_name', 'stevedore_name', 'ops_started', 'ops_completed'}
        cols = [k for k in data if k not in _computed]
        cur.execute(f"UPDATE ldud_header SET {', '.join([f'{c}=%s' for c in cols])} WHERE id=%s",
                   [data[c] for c in cols] + [row_id])
    else:
        data['doc_num'] = get_next_doc_num()
        _computed = {'id', 'vcn_display', 'cargo_names_display', 'bl_quantities_display',
                     'balance_display', 'agent_name', 'stevedore_name', 'ops_started', 'ops_completed'}
        cols = [k for k in data if k not in _computed]
        cur.execute(f"INSERT INTO ldud_header ({', '.join(cols)}) VALUES ({', '.join(['%s']*len(cols))}) RETURNING id",
                   [data[c] for c in cols])
        row_id = cur.fetchone()['id']

    conn.commit()
    conn.close()
    return row_id, data.get('doc_num')

def delete_header(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM ldud_header WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


def parcels_completion(ldud_id):
    """Whether every parcel-op for this LDUD has an actual End time recorded in
    LUEU01 — gates the post-departure SOF times (cast off, pilot board, etc.)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT COUNT(*) AS total,
                          COUNT(*) FILTER (WHERE end_dt IS NULL OR end_dt = '') AS incomplete
                   FROM ldud_parcel_ops WHERE ldud_id=%s''', [ldud_id])
    r = cur.fetchone()
    conn.close()
    total, incomplete = r['total'], r['incomplete']
    return {'total': total, 'incomplete': incomplete, 'all_done': total > 0 and incomplete == 0}


# Parcel Operations sub-table — each row covers one VCN parcel, or several
# parcels MERGED together (only allowed when they share the same cargo name).
# parcel_ids is a CSV of parcel ids: vcn_consigners.id for Import LDUDs,
# vcn_export_cargo_declaration.id for Export LDUDs (resolved via the linked
# VCN's operation_type — one LDUD is one VCN of one operation type).
def _parse_ids(csv):
    return [int(x) for x in str(csv or '').split(',') if str(x).strip().isdigit()]


def _parcel_table_for_ldud(cur, ldud_id):
    """Return the VCN parcel source table for this LDUD based on operation_type."""
    cur.execute('''SELECT h.operation_type
                   FROM ldud_header l JOIN vcn_header h ON h.id = l.vcn_id
                   WHERE l.id=%s''', [ldud_id])
    row = cur.fetchone()
    op = (row or {}).get('operation_type') if row else None
    return 'vcn_export_cargo_declaration' if op == 'Export' else 'vcn_consigners'


def get_parcel_ops(ldud_id):
    conn = get_db()
    cur = get_cursor(conn)
    tbl = _parcel_table_for_ldud(cur, ldud_id)
    cur.execute('SELECT * FROM ldud_parcel_ops WHERE ldud_id=%s ORDER BY id', (ldud_id,))
    rows = [dict(r) for r in cur.fetchall()]
    # Resolve parcel labels for display from the correct source table
    all_ids = sorted({pid for r in rows for pid in _parse_ids(r['parcel_ids'])})
    labels = {}
    if all_ids:
        cur.execute(f'SELECT id, parcel_no FROM {tbl} WHERE id = ANY(%s)', (all_ids,))
        labels = {r['id']: (r['parcel_no'] or f"#{r['id']}") for r in cur.fetchall()}
    conn.close()
    for r in rows:
        ids = _parse_ids(r['parcel_ids'])
        r['parcel_nos_display'] = ', '.join(labels.get(i, f"#{i}") for i in ids)
    return rows


def save_parcel_op(data):
    # One row = one parcel + terminal + quantity. The same parcel may appear on
    # multiple rows (e.g. one per terminal), so no merge/uniqueness guard.
    # start_dt/end_dt are NOT written here — they are owned by LUEU01 (operators
    # enter parcel start/end in the logbook). An LDUD edit must not wipe them.
    _clean_empty(data)
    ids = _parse_ids(data.get('parcel_ids'))
    parcel_ids = ','.join(map(str, ids)) if ids else None
    quantity = data.get('quantity')
    if isinstance(quantity, str) and quantity.strip() == '':
        quantity = None
    conn = get_db()
    cur = get_cursor(conn)

    # Guard: quantity can't exceed the VCN parcel quantity. Across rows that share
    # a parcel (e.g. one parcel split over terminals) the total can't exceed it either.
    if ids and quantity is not None:
        tbl = _parcel_table_for_ldud(cur, data['ldud_id'])
        # both source tables now use a TEXT 'quantity' column (export mirrors import)
        cur.execute(f'SELECT quantity AS q FROM {tbl} WHERE id = ANY(%s)', (ids,))
        cap = 0.0
        for r in cur.fetchall():
            try:
                cap += float(str(r['q']).replace(',', '')) if r['q'] is not None else 0.0
            except (ValueError, TypeError):
                pass
        cur.execute('SELECT id, parcel_ids, quantity FROM ldud_parcel_ops WHERE ldud_id=%s', [data['ldud_id']])
        used = 0.0
        for r in cur.fetchall():
            if data.get('id') and r['id'] == data['id']:
                continue
            if set(_parse_ids(r['parcel_ids'])) & set(ids):
                used += float(r['quantity'] or 0)
        if cap > 0 and float(quantity) + used > cap + 1e-6:
            conn.close()
            raise ValueError(f"Quantity {round(float(quantity), 3)} exceeds the available VCN parcel "
                             f"quantity ({round(cap - used, 3)} MT).")

    cols = ['parcel_ids', 'cargo_name', 'terminal_name', 'quantity']
    vals = [parcel_ids, data.get('cargo_name'), data.get('terminal_name'), quantity]
    if data.get('id'):
        cur.execute(f"UPDATE ldud_parcel_ops SET {', '.join(f'{c}=%s' for c in cols)} WHERE id=%s",
                    vals + [data['id']])
        row_id = data['id']
    else:
        cur.execute(f'''INSERT INTO ldud_parcel_ops (ldud_id, {', '.join(cols)})
                       VALUES ({', '.join(['%s'] * (len(cols) + 1))}) RETURNING id''',
                    [data['ldud_id']] + vals)
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id


def delete_parcel_op(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM ldud_parcel_ops WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


# Closure functions
def get_doc_status(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT doc_status FROM ldud_header WHERE id=%s', (record_id,))
    row = cur.fetchone()
    conn.close()
    return row['doc_status'] if row else None


def get_closure_eligibility(ldud_id):
    """Closure gate: vessel name + NOR tendered (≥1 proof doc is enforced in
    the view), plus the quantity picture shown before closing:
      ops_total        — ACTUAL LUEU01 logged quantity (short-close EXCLUDED)
      bl_total         — parcel target total (live VCN parcel quantities,
                         falling back to the op snapshot)
      shortclose_total — written-off leftover, reported separately
    Full close only when the actual quantity matches the BL total."""
    conn = get_db()
    cur = get_cursor(conn)
    missing = []

    cur.execute('SELECT vessel_name, nor_tendered, vcn_id FROM ldud_header WHERE id=%s', (ldud_id,))
    header = cur.fetchone()
    if not header:
        conn.close()
        return {'eligible': False, 'missing': ['Record not found'], 'ops_total': 0,
                'bl_total': 0, 'shortclose_total': 0, 'can_full_close': False}

    if not header['vessel_name']:
        missing.append('Vessel Name (select a VCN to populate)')
    if not header['nor_tendered']:
        missing.append('NOR Tendered (header field)')

    # BL total: per-op target = live source-parcel quantities, fallback op snapshot
    cur.execute('SELECT parcel_ids, quantity AS op_qty FROM ldud_parcel_ops WHERE ldud_id=%s', [ldud_id])
    ops = [dict(r) for r in cur.fetchall()]
    src_qty = {}
    all_ids = sorted({pid for o in ops for pid in _parse_ids(o['parcel_ids'])})
    if all_ids:
        tbl = _parcel_table_for_ldud(cur, ldud_id)
        cur.execute(f'SELECT id, quantity FROM {tbl} WHERE id = ANY(%s)', [all_ids])
        for r in cur.fetchall():
            try:
                src_qty[r['id']] = float(str(r['quantity']).replace(',', '')) if r['quantity'] is not None else 0.0
            except (ValueError, TypeError):
                src_qty[r['id']] = 0.0
    bl_total = 0.0
    for o in ops:
        ids = _parse_ids(o['parcel_ids'])
        try:
            fallback = float(str(o['op_qty']).replace(',', '')) if o['op_qty'] is not None else 0.0
        except (ValueError, TypeError):
            fallback = 0.0
        bl_total += sum(src_qty.get(i, 0.0) for i in ids) or fallback

    # Actual handled (LUEU01) — short-close is a write-off, NOT actual quantity
    cur.execute('''SELECT
            COALESCE(SUM(CASE WHEN COALESCE(lg.is_shortclose, FALSE) THEN 0 ELSE lg.quantity END), 0) AS actual,
            COALESCE(SUM(CASE WHEN COALESCE(lg.is_shortclose, FALSE) THEN lg.quantity ELSE 0 END), 0) AS sc
        FROM lueu_parcel_log lg
        JOIN ldud_parcel_ops po ON po.id = lg.parcel_op_id
        WHERE po.ldud_id = %s AND lg.is_deleted IS NOT TRUE''', [ldud_id])
    q = cur.fetchone()
    conn.close()

    ops_total = round(float(q['actual'] or 0), 3)
    shortclose_total = round(float(q['sc'] or 0), 3)
    bl_total = round(bl_total, 3)
    eligible = len(missing) == 0
    # Displayed actual excludes short-close, but the write-off still counts
    # toward completion — otherwise a short-closed vessel could never Full Close.
    return {
        'eligible': eligible,
        'missing': missing,
        'ops_total': ops_total,
        'bl_total': bl_total,
        'shortclose_total': shortclose_total,
        'can_full_close': eligible and abs((ops_total + shortclose_total) - bl_total) < 0.005,
    }


def close_record(record_id, close_type, username):
    """close_type: 'Closed' or 'Partial Close'"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('UPDATE ldud_header SET doc_status=%s WHERE id=%s', (close_type, record_id))
    cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                   VALUES ('LDUD01', %s, %s, NULL, %s)""", (record_id, close_type, username))
    conn.commit()
    conn.close()


def get_closure_log(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    # explicit columns only: proof_bytes is BYTEA and must never reach JSON
    cur.execute("""SELECT id, action, comment, actioned_by,
                          to_char(actioned_at, 'DD-MM-YYYY HH24:MI') AS actioned_at,
                          (proof_bytes IS NOT NULL) AS has_proof
                   FROM approval_log WHERE module_code='LDUD01' AND record_id=%s
                   ORDER BY actioned_at DESC""", (record_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]
