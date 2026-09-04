from database import get_db, get_cursor

def _clean_empty(data):
    """Convert empty strings to None so timestamp/date columns get NULL."""
    for k in data:
        if data[k] == '':
            data[k] = None
    return data

def default_discharge_port():
    """Name of the port flagged default in VPM01 (auto-fills VCN Discharge Port)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT name FROM port_master WHERE is_default_discharge LIMIT 1")
    row = cur.fetchone()
    conn.close()
    return row['name'] if row else None

def get_next_doc_num():
    import datetime
    conn = get_db()
    cur = get_cursor(conn)
    now = datetime.datetime.now()
    fy_start = now.year if now.month >= 4 else now.year - 1
    fy_suffix = f"{str(fy_start)[2:]}{str(fy_start + 1)[2:]}"  # e.g. "2526"
    prefix = f"VCN-{fy_suffix}-"
    cur.execute(
        "SELECT MAX(CAST(SPLIT_PART(vcn_doc_num, '-', 3) AS INTEGER)) FROM vcn_header WHERE vcn_doc_num LIKE %s",
        (prefix + '%',)
    )
    result = cur.fetchone()['max']
    conn.close()
    next_num = (result or 0) + 1
    return f"{prefix}{next_num:03d}"

def doc_num_taken(doc_num, exclude_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT 1 FROM vcn_header WHERE vcn_doc_num=%s AND id<>%s LIMIT 1',
                [doc_num, exclude_id])
    taken = cur.fetchone() is not None
    conn.close()
    return taken

def get_vessels():
    """Get vessels from VC01 for dropdown"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT doc_num, vessel_name, pbl FROM vessels ORDER BY doc_num')
    rows = cur.fetchall()
    conn.close()
    return [{'value': f"{r['doc_num']}/{r['vessel_name']}", 'doc_num': r['doc_num'],
             'vessel_name': r['vessel_name'], 'pbl': r['pbl']} for r in rows]

def get_data(page=1, size=20, filters=None):
    conn = get_db()
    cur = get_cursor(conn)

    allowed = {'operation_type','vcn_doc_num','vessel_name','vessel_agent_name',
               'cargo_type','doc_status','doc_date',
               'customer_name','load_port','discharge_port'}
    where_clauses, params = [], []
    for f in (filters or []):
        field = f.get('field', '')
        if field not in allowed:
            continue
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
        cur.execute(f'SELECT COUNT(*) FROM vcn_header {where_sql}', params)
        total = cur.fetchone()['count']
        # total_delay_mins: sum of the Pre Berthing/Anchoring Delays sub-table, so the
        # grid shows it without opening each row. Times are TEXT 'YYYY-MM-DDTHH:MM' —
        # the regex skips junk, and the >= compares fine lexicographically in that format.
        cur.execute(f'''SELECT h.*, (
                            SELECT COALESCE(SUM(EXTRACT(EPOCH FROM
                                       (d.delay_end::timestamp - d.delay_start::timestamp)) / 60), 0)
                            FROM vcn_delays d
                            WHERE d.vcn_id = h.id
                              AND d.delay_start ~ '^[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}T[0-9]{{2}}:[0-9]{{2}}'
                              AND d.delay_end   ~ '^[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}T[0-9]{{2}}:[0-9]{{2}}'
                              AND d.delay_end >= d.delay_start
                        ) AS total_delay_mins
                        FROM vcn_header h {where_sql} ORDER BY h.id DESC LIMIT %s OFFSET %s''',
                    params + [size, (page - 1) * size])
        rows = []
        for r in cur.fetchall():
            r = dict(r)
            r.pop('igm_document', None)   # BYTEA — not JSON-serializable
            r['has_igm_doc'] = bool(r.get('igm_document_name'))
            r['total_delay_mins'] = int(r.get('total_delay_mins') or 0)   # Decimal -> JSON
            rows.append(r)
        return rows, total
    finally:
        conn.close()

def save_header(data):
    _clean_empty(data)
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')

    # computed / blob fields never come through the JSON save path
    for k in ('has_igm_doc', 'igm_document', 'igm_document_name', 'total_delay_mins'):
        data.pop(k, None)

    if row_id:
        cur.execute('SELECT vcn_doc_num FROM vcn_header WHERE id=%s', [row_id])
        old_doc = (cur.fetchone() or {}).get('vcn_doc_num')
        cols = [k for k in data if k != 'id']
        cur.execute(f"UPDATE vcn_header SET {', '.join([f'{c}=%s' for c in cols])} WHERE id=%s",
                   [data[c] for c in cols] + [row_id])
        # parcel labels embed the doc number ('<doc>/P<seq>') — renumber them too
        new_doc = data.get('vcn_doc_num')
        if new_doc and new_doc != old_doc:
            for t in ('vcn_consigners', 'vcn_export_cargo_declaration'):
                cur.execute(f"UPDATE {t} SET parcel_no = %s || '/P' || parcel_seq "
                            f"WHERE vcn_id=%s AND parcel_seq IS NOT NULL", [new_doc, row_id])
    else:
        data['vcn_doc_num'] = get_next_doc_num()
        cols = [k for k in data if k != 'id']
        cur.execute(f"INSERT INTO vcn_header ({', '.join(cols)}) VALUES ({', '.join(['%s']*len(cols))}) RETURNING id",
                   [data[c] for c in cols])
        row_id = cur.fetchone()['id']

    # PBL entered on the VCN flows back to the VC01 vessel master (one-way; only
    # when provided, so a blank never clears the master). vessel_master_doc is
    # 'DOCNUM/NAME' — match the vessel by its doc_num.
    pbl = data.get('pbl')
    vmd = data.get('vessel_master_doc')
    if pbl not in (None, '') and vmd:
        cur.execute('UPDATE vessels SET pbl=%s WHERE doc_num=%s', [pbl, str(vmd).split('/', 1)[0]])

    conn.commit()
    conn.close()
    return row_id, data.get('vcn_doc_num')

def delete_header(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM vcn_header WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


def send_back_to_expected(vcn_id, username):
    """Undo an accidental EV01→VCN move: recreate the expected_vessels row
    from the VCN header (reverse of EV01 move_to_vcn) and delete the VCN.
    Any LDUD for the VCN is soft-deleted (hidden but recoverable if the
    vessel comes back later). Caller checks billed/Approved."""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('SELECT * FROM vcn_header WHERE id=%s', [vcn_id])
        h = cur.fetchone()
        if not h:
            raise ValueError('Record not found')
        h = dict(h)
        # Soft-delete the LDUD (keeps its parcel ops / LUEU logbook rows).
        # vcn_id is detached: this vcn_header row is deleted below and every
        # LDUD/LUEU listing joins through vcn_id, so hidden rows drop out
        # everywhere. vcn_doc_num/vessel_name text stays for tracing/restore.
        cur.execute('''UPDATE ldud_header
                       SET is_deleted=TRUE, deleted_by=%s, deleted_date=now(), vcn_id=NULL
                       WHERE vcn_id=%s''', [username, vcn_id])
        cur.execute('''SELECT DISTINCT consigner_name FROM vcn_consigners
                       WHERE vcn_id=%s AND consigner_name IS NOT NULL ORDER BY consigner_name''', [vcn_id])
        consignees = ', '.join(r['consigner_name'] for r in cur.fetchall()) or None
        cur.execute('SELECT cargo_name, total_qty FROM vcn_cargo_quota WHERE vcn_id=%s ORDER BY cargo_name', [vcn_id])
        quotas = [dict(r) for r in cur.fetchall()]
        cargo = ', '.join(q['cargo_name'] for q in quotas) or h.get('cargo_type')
        qty = ', '.join(str(q['total_qty']) for q in quotas if q['total_qty'] is not None) or None
        cur.execute('''INSERT INTO expected_vessels
                       (vessel_name, via_number, loa, draft, agents, consignees, cargo_name,
                        quantity, nor, eta, berth_name, doc_status, created_by)
                       VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'Pending',%s) RETURNING id''',
                    [h.get('vessel_name'), h.get('via_number'), h.get('loa'), h.get('draft'),
                     h.get('vessel_agent_name'), consignees, cargo, qty,
                     h.get('nor_tendered'), h.get('doc_date'), h.get('berth_name'), username])
        ev_id = cur.fetchone()['id']
        cur.execute('DELETE FROM vcn_header WHERE id=%s', [vcn_id])
        conn.commit()
        return ev_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

# Consigner (customer details) sub-table — each row is one PARCEL (one IGM/FORM III
# line: product + receiver + BL). vessel agent is captured on the header.
_CONSIGNER_COLS = ['igm_line_no', 'bl_no', 'bl_date', 'cargo_name', 'quantity',
                   'consigner_name', 'importer_name',
                   'pipeline_name', 'unload_terminal',
                   'toll_applicable', 'toll_reason', 'equipment_names']

# Export parcels mirror import parcels minus the BL fields (see spec
# 2026-07-01-vcn01-export-parcels). Same list-driven CRUD, different table.
_EXPORT_PARCEL_COLS = [c for c in _CONSIGNER_COLS if c not in ('bl_no', 'bl_date')]


def get_equipment():
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT name FROM equipment ORDER BY name')
    rows = [r['name'] for r in cur.fetchall()]
    conn.close()
    return rows


def _parcel_no(cur, vcn_id, seq):
    """Build the stored parcel label '<vcn_doc_num>/P<seq>' (or 'P<seq>' if the
    parent VCN has no doc number yet — e.g. a brand-new draft)."""
    cur.execute('SELECT vcn_doc_num FROM vcn_header WHERE id=%s', [vcn_id])
    row = cur.fetchone()
    doc = (row or {}).get('vcn_doc_num') if row else None
    return f"{doc}/P{seq}" if doc else f"P{seq}"


def _sync_header_cargo(cur, vcn_id):
    """Keep vcn_header.cargo_type in sync with the parcels' cargo names.
    Recomputes it from the distinct cargo names across import consigners and
    export declarations, so editing a parcel's cargo reflects on the header."""
    if not vcn_id:
        return
    cur.execute('''
        SELECT cargo_name FROM vcn_consigners WHERE vcn_id=%s AND cargo_name IS NOT NULL
        UNION
        SELECT cargo_name FROM vcn_export_cargo_declaration WHERE vcn_id=%s AND cargo_name IS NOT NULL
    ''', (vcn_id, vcn_id))
    names = []
    for r in cur.fetchall():
        for name in (r['cargo_name'] or '').split(','):   # consigner rows may be comma-separated
            name = name.strip()
            if name and name not in names:
                names.append(name)
    cur.execute('UPDATE vcn_header SET cargo_type=%s WHERE id=%s',
                [', '.join(sorted(names)), vcn_id])


def get_header_cargo_type(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT cargo_type FROM vcn_header WHERE id=%s', [vcn_id])
    row = cur.fetchone()
    conn.close()
    return (row or {}).get('cargo_type') if row else None


def get_parcels(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM vcn_consigners WHERE vcn_id=%s ORDER BY parcel_seq NULLS LAST, id',
                (vcn_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_parcel(row_id):
    """Single parcel row — used to return the generated parcel_no after a save."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, parcel_no, parcel_seq FROM vcn_consigners WHERE id=%s', [row_id])
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None

# back-compat alias — existing callers/endpoints use get_consigners
get_consigners = get_parcels


def _operation_type(cur, vcn_id):
    cur.execute('SELECT operation_type FROM vcn_header WHERE id=%s', [vcn_id])
    row = cur.fetchone()
    return (row or {}).get('operation_type') if row else None


def get_picker_parcels(vcn_id):
    """Operation-type-aware parcel list for cross-module pickers (LDUD).
    Import → consigner rows; Export → export cargo declaration rows.
    Returns: id, parcel_no, cargo_name, consigner_name, quantity, terminals (list).
    Terminals come from the consigner/export row's unload_terminal (multi-value,
    comma separated)."""
    conn = get_db()
    cur = get_cursor(conn)
    is_export = _operation_type(cur, vcn_id) == 'Export'
    if is_export:
        cur.execute('''SELECT id, parcel_no, cargo_name, consigner_name, quantity, unload_terminal
                       FROM vcn_export_cargo_declaration WHERE vcn_id=%s
                       ORDER BY parcel_seq NULLS LAST, id''', (vcn_id,))
    else:
        cur.execute('''SELECT id, parcel_no, cargo_name, consigner_name, quantity, unload_terminal
                       FROM vcn_consigners WHERE vcn_id=%s
                       ORDER BY parcel_seq NULLS LAST, id''', (vcn_id,))
    rows = []
    for r in cur.fetchall():
        d = dict(r)
        d['terminals'] = [t.strip() for t in str(d.pop('unload_terminal', '') or '').split(',') if t.strip()]
        rows.append(d)
    conn.close()
    return rows


# Master parcel export — mirrors the on-screen parcels sub-table, prefixed with
# the vessel/VCN identity. 'dt:' fields split into Date + Time columns (excel_export).
EXPORT_PARCEL_COLS = [
    ('Vessel Name', 'vessel_name'), ('VCN Doc', 'vcn_doc_num'), ('Operation Type', 'operation_type'),
    ('Berth', 'berth_name'), ('Vessel Agent', 'vessel_agent_name'), ('VCN Status', 'doc_status'),
    ('Doc Date', 'doc_date'), ('NOR Tendered', 'dt:nor_tendered'),
    ('Parcel No', 'parcel_no'), ('Ln', 'igm_line_no'), ('BL No', 'bl_no'), ('BL Date', 'bl_date'),
    ('Cargo Name', 'cargo_name'), ('Qty (MT)', 'quantity'), ('Consignee', 'consigner_name'),
    ('Payment will be made by', 'importer_name'), ('Pipeline', 'pipeline_name'),
    ('Terminal', 'unload_terminal'), ('Toll', 'toll_applicable'), ('Toll Reason', 'toll_reason'),
    ('Equipment', 'equipment_names'),
]


def export_all_parcels():
    """Every parcel across every VCN — import consigners and export declarations
    in one list, same columns the parcels sub-table shows. Export rows carry no
    BL, so those cells come back empty."""
    conn = get_db()
    cur = get_cursor(conn)
    rows = []
    for tbl, bl in (('vcn_consigners', 'p.bl_no, p.bl_date'),
                    ('vcn_export_cargo_declaration', "NULL AS bl_no, NULL AS bl_date")):
        cur.execute(f'''
            SELECT h.vessel_name, h.vcn_doc_num, h.operation_type, h.berth_name,
                   h.vessel_agent_name, h.doc_status, h.doc_date, h.nor_tendered,
                   p.parcel_no, p.parcel_seq, p.igm_line_no, {bl}, p.cargo_name, p.quantity,
                   p.consigner_name, p.importer_name, p.pipeline_name, p.unload_terminal,
                   p.toll_applicable, p.toll_reason, p.equipment_names
            FROM {tbl} p JOIN vcn_header h ON h.id = p.vcn_id
        ''')
        rows += [dict(r) for r in cur.fetchall()]
    conn.close()
    rows.sort(key=lambda r: (r['vcn_doc_num'] or '', r['parcel_seq'] or 0))
    return rows


def get_delay_window(vcn_id):
    """LDUD01's Anchorage / Pilot Pickup times — the window pre-berthing delays
    must fall inside. Either may be unset; the UI then enforces only what it has."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT anchored_datetime, pilot_pickup_time FROM ldud_header
                   WHERE vcn_id=%s AND is_deleted IS NOT TRUE ORDER BY id DESC LIMIT 1''', [vcn_id])
    row = cur.fetchone() or {}
    conn.close()
    return {'anchored': row.get('anchored_datetime') or '',
            'pilot_pickup': row.get('pilot_pickup_time') or ''}


def get_export_parcel(row_id):
    """Single export-cargo parcel — returns generated parcel_no after a save."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id, parcel_no, parcel_seq FROM vcn_export_cargo_declaration WHERE id=%s', [row_id])
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_consigner_vcn_id(row_id):
    """vcn_id owning an import consigner parcel (for billed-lock checks before delete)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT vcn_id FROM vcn_consigners WHERE id=%s', [row_id])
    row = cur.fetchone()
    conn.close()
    return row['vcn_id'] if row else None


def get_export_parcel_vcn_id(row_id):
    """vcn_id owning an export cargo parcel (for billed-lock checks before delete)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT vcn_id FROM vcn_export_cargo_declaration WHERE id=%s', [row_id])
    row = cur.fetchone()
    conn.close()
    return row['vcn_id'] if row else None


def _parse_qty(v):
    try:
        return float(str(v).replace(',', '')) if v not in (None, '') else 0.0
    except (TypeError, ValueError):
        return 0.0


def save_consigner(data):
    # The per-cargo quota (from EV01) is informational only — the "Available per
    # cargo" panel shows allocated vs total and flags over-allocation, but parcels
    # are NOT blocked from exceeding it.
    _clean_empty(data)
    conn = get_db()
    cur = get_cursor(conn)
    if data.get('id'):
        cur.execute(f"UPDATE vcn_consigners SET {', '.join(f'{c}=%s' for c in _CONSIGNER_COLS)} WHERE id=%s",
                   [data.get(c) for c in _CONSIGNER_COLS] + [data['id']])
        row_id = data['id']
        # backfill parcel_no if it was created on a draft before the VCN had a doc number
        cur.execute('SELECT parcel_seq, parcel_no FROM vcn_consigners WHERE id=%s', [row_id])
        cur_row = cur.fetchone()
        if cur_row and cur_row['parcel_seq'] and not cur_row['parcel_no']:
            cur.execute('UPDATE vcn_consigners SET parcel_no=%s WHERE id=%s',
                        [_parcel_no(cur, data['vcn_id'], cur_row['parcel_seq']), row_id])
    else:
        cur.execute('SELECT COALESCE(MAX(parcel_seq), 0) + 1 AS nxt FROM vcn_consigners WHERE vcn_id=%s',
                    [data['vcn_id']])
        seq = cur.fetchone()['nxt']
        parcel_no = _parcel_no(cur, data['vcn_id'], seq)
        cols = _CONSIGNER_COLS + ['parcel_seq', 'parcel_no']
        vals = [data.get(c) for c in _CONSIGNER_COLS] + [seq, parcel_no]
        cur.execute(f'''INSERT INTO vcn_consigners (vcn_id, {', '.join(cols)})
                       VALUES ({', '.join(['%s'] * (len(cols) + 1))}) RETURNING id''',
                   [data['vcn_id']] + vals)
        row_id = cur.fetchone()['id']
    _sync_header_cargo(cur, data.get('vcn_id'))
    conn.commit()
    conn.close()
    return row_id

# back-compat alias
save_parcel = save_consigner


def save_cargo_quotas(vcn_id, quotas):
    """Upsert per-cargo totals {cargo_name: qty} for a VCN (captured on EV01 move)."""
    conn = get_db()
    cur = get_cursor(conn)
    for cargo, qty in (quotas or {}).items():
        cur.execute('''INSERT INTO vcn_cargo_quota (vcn_id, cargo_name, total_qty)
                       VALUES (%s, %s, %s)
                       ON CONFLICT (vcn_id, cargo_name) DO UPDATE SET total_qty = EXCLUDED.total_qty''',
                    [vcn_id, cargo, qty])
    conn.commit()
    conn.close()


def get_cargo_quotas(vcn_id):
    """Per-cargo available/allocated/remaining for a VCN's parcel allocation."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT cargo_name, total_qty FROM vcn_cargo_quota WHERE vcn_id=%s ORDER BY cargo_name', [vcn_id])
    quotas = [dict(r) for r in cur.fetchall()]
    cur.execute('SELECT cargo_name, quantity FROM vcn_consigners WHERE vcn_id=%s', [vcn_id])
    alloc = {}
    for r in cur.fetchall():
        alloc[r['cargo_name']] = alloc.get(r['cargo_name'], 0.0) + _parse_qty(r['quantity'])
    conn.close()
    out = []
    for q in quotas:
        total = float(q['total_qty'] or 0)
        a = round(alloc.get(q['cargo_name'], 0.0), 3)
        out.append({'cargo_name': q['cargo_name'], 'total_qty': round(total, 3),
                    'allocated': a, 'remaining': round(total - a, 3)})
    return out

def save_igm_document(vcn_id, filename, file_bytes):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('UPDATE vcn_header SET igm_document=%s, igm_document_name=%s WHERE id=%s',
                [file_bytes, filename, vcn_id])
    conn.commit()
    conn.close()

def get_igm_document(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT igm_document, igm_document_name FROM vcn_header WHERE id=%s', (vcn_id,))
    row = cur.fetchone()
    conn.close()
    if not row or not row['igm_document']:
        return None, None
    return bytes(row['igm_document']), row['igm_document_name']

def delete_consigner(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT vcn_id FROM vcn_consigners WHERE id=%s', (row_id,))
    r = cur.fetchone()
    vcn_id = r['vcn_id'] if r else None
    cur.execute('DELETE FROM vcn_consigners WHERE id=%s', (row_id,))
    _sync_header_cargo(cur, vcn_id)
    conn.commit()
    conn.close()
    return vcn_id

# Delays sub-table operations
def get_delays(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    # chronological — the grid reads as a log and "+ Add" appends to the bottom
    cur.execute('SELECT * FROM vcn_delays WHERE vcn_id=%s ORDER BY id', (vcn_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def save_delay(data):
    _clean_empty(data)
    conn = get_db()
    cur = get_cursor(conn)
    if data.get('id'):
        cur.execute('UPDATE vcn_delays SET delay_name=%s, delay_start=%s, delay_end=%s WHERE id=%s',
                   [data.get('delay_name'), data.get('delay_start'), data.get('delay_end'), data['id']])
        row_id = data['id']
    else:
        cur.execute('INSERT INTO vcn_delays (vcn_id, delay_name, delay_start, delay_end) VALUES (%s, %s, %s, %s) RETURNING id',
                   [data['vcn_id'], data.get('delay_name'), data.get('delay_start'), data.get('delay_end')])
        row_id = cur.fetchone()['id']
    conn.commit()
    conn.close()
    return row_id

def delete_delay(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM vcn_delays WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()

# Import cargo is declared in the consigner table now; vcn_cargo_declaration
# remains read-only for historic data (billing/LDUD still query it).
def get_all_cargo_names_for_vcn(vcn_id):
    """All cargo names for a VCN — consigners plus historic declarations."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT DISTINCT cargo_name FROM (
            SELECT cargo_name FROM vcn_consigners WHERE vcn_id=%s AND cargo_name IS NOT NULL
            UNION
            SELECT cargo_name FROM vcn_cargo_declaration WHERE vcn_id=%s AND cargo_name IS NOT NULL
            UNION
            SELECT cargo_name FROM vcn_export_cargo_declaration WHERE vcn_id=%s AND cargo_name IS NOT NULL
        ) combined ORDER BY cargo_name
    ''', (vcn_id, vcn_id, vcn_id))
    rows = cur.fetchall()
    conn.close()
    names = []
    for r in rows:
        if not r['cargo_name']:
            continue
        # consigner rows may hold comma-separated cargo lists
        for name in r['cargo_name'].split(','):
            name = name.strip()
            if name and name not in names:
                names.append(name)
    return sorted(names)

# Export Cargo Declaration sub-table operations
def get_export_cargo_declarations(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM vcn_export_cargo_declaration WHERE vcn_id=%s ORDER BY parcel_seq NULLS LAST, id', (vcn_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def save_export_cargo_declaration(data):
    _clean_empty(data)
    conn = get_db()
    cur = get_cursor(conn)
    if data.get('id'):
        cur.execute(f"UPDATE vcn_export_cargo_declaration SET {', '.join(f'{c}=%s' for c in _EXPORT_PARCEL_COLS)} WHERE id=%s",
                    [data.get(c) for c in _EXPORT_PARCEL_COLS] + [data['id']])
        row_id = data['id']
        cur.execute('SELECT parcel_seq, parcel_no FROM vcn_export_cargo_declaration WHERE id=%s', [row_id])
        cur_row = cur.fetchone()
        if cur_row and cur_row['parcel_seq'] and not cur_row['parcel_no']:
            cur.execute('UPDATE vcn_export_cargo_declaration SET parcel_no=%s WHERE id=%s',
                        [_parcel_no(cur, data['vcn_id'], cur_row['parcel_seq']), row_id])
    else:
        cur.execute('SELECT COALESCE(MAX(parcel_seq), 0) + 1 AS nxt FROM vcn_export_cargo_declaration WHERE vcn_id=%s',
                    [data['vcn_id']])
        seq = cur.fetchone()['nxt']
        parcel_no = _parcel_no(cur, data['vcn_id'], seq)
        cols = _EXPORT_PARCEL_COLS + ['parcel_seq', 'parcel_no']
        vals = [data.get(c) for c in _EXPORT_PARCEL_COLS] + [seq, parcel_no]
        cur.execute(f'''INSERT INTO vcn_export_cargo_declaration (vcn_id, {', '.join(cols)})
                       VALUES ({', '.join(['%s'] * (len(cols) + 1))}) RETURNING id''',
                    [data['vcn_id']] + vals)
        row_id = cur.fetchone()['id']
    _sync_header_cargo(cur, data.get('vcn_id'))
    conn.commit()
    conn.close()
    return row_id

def delete_export_cargo_declaration(row_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT vcn_id FROM vcn_export_cargo_declaration WHERE id=%s', (row_id,))
    r = cur.fetchone()
    vcn_id = r['vcn_id'] if r else None
    cur.execute('DELETE FROM vcn_export_cargo_declaration WHERE id=%s', (row_id,))
    _sync_header_cargo(cur, vcn_id)
    conn.commit()
    conn.close()
    return vcn_id


def get_export_cargo_names_for_vcn(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT DISTINCT cargo_name FROM vcn_export_cargo_declaration WHERE vcn_id=%s AND cargo_name IS NOT NULL', (vcn_id,))
    rows = cur.fetchall()
    conn.close()
    return [r['cargo_name'] for r in rows if r['cargo_name']]

def get_export_cargo_total_quantity(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT COALESCE(SUM(NULLIF(quantity, '')::numeric), 0) AS s FROM vcn_export_cargo_declaration WHERE vcn_id=%s", (vcn_id,))
    result = cur.fetchone()['s']
    conn.close()
    return float(result or 0)

def get_export_loading_totals(vcn_id):
    """Loading totals per cargo. ponytail: ldud_vessel_operations was dropped
    when LDUD01 moved to parcel-based ops; returns empty until re-sourced."""
    return {}


def get_hold_completion_by_vcn(vcn_id):
    """Hold completion across LDUDs for a VCN. ponytail: ldud_hold_completion
    was dropped; returns empty until the parcel-based ops flow re-feeds it."""
    return []


# Approval functions
def get_doc_status(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT doc_status FROM vcn_header WHERE id=%s', (record_id,))
    row = cur.fetchone()
    conn.close()
    return row['doc_status'] if row else None


def get_approval_eligibility(vcn_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT operation_type, vessel_name, vessel_agent_name,
                          cargo_type, discharge_port, via_number, berth_name, vessel_run_type
                   FROM vcn_header WHERE id=%s''', (vcn_id,))
    header = cur.fetchone()
    if not header:
        conn.close()
        return {'eligible': False, 'missing': ['Record not found']}

    missing = []
    if not header['operation_type']:
        missing.append('Operation Type')
    if not header['vessel_name']:
        missing.append('Vessel Name')
    if not header['vessel_agent_name']:
        missing.append('Agent Name')
    if not header['via_number']:
        missing.append('VIA No.')
    if not header['berth_name']:
        missing.append('Berth')
    if not header['vessel_run_type']:
        missing.append('Vessel Run Type')
    if not header['discharge_port']:
        missing.append('Discharge Port')

    # Every parcel must be complete for closure. Export and import parcels now
    # share the same shape, so validate the fields billing/ops need on each row.
    op_type = header['operation_type']
    tbl = 'vcn_export_cargo_declaration' if op_type == 'Export' else 'vcn_consigners'
    cur.execute(f'''SELECT parcel_no, cargo_name, quantity, consigner_name, importer_name,
                           pipeline_name, unload_terminal
                    FROM {tbl} WHERE vcn_id=%s ORDER BY parcel_seq NULLS LAST, id''', (vcn_id,))
    parcels = cur.fetchall()
    conn.close()

    if not parcels:
        missing.append('At least one parcel')
    else:
        required = [('cargo_name', 'Cargo'), ('quantity', 'Qty'),
                    ('consigner_name', 'Consignee'), ('importer_name', 'Payment by'),
                    ('pipeline_name', 'Pipeline'), ('unload_terminal', 'Terminal')]
        bad = []
        for p in parcels:
            gaps = [label for field, label in required
                    if not (str(p[field]).strip() if p[field] is not None else '')]
            if gaps:
                bad.append(f"{p['parcel_no'] or '#?'} ({', '.join(gaps)})")
        if bad:
            missing.append('Incomplete parcels — ' + '; '.join(bad))

    return {'eligible': len(missing) == 0, 'missing': missing}


def approve_record(record_id, username):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("UPDATE vcn_header SET doc_status='Approved' WHERE id=%s", (record_id,))
    cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                   VALUES ('VCN01', %s, 'Approved', NULL, %s)""", (record_id, username))
    conn.commit()
    conn.close()


def send_back_to_draft(record_id, comment, username):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("UPDATE vcn_header SET doc_status='Draft' WHERE id=%s", (record_id,))
    cur.execute("""INSERT INTO approval_log (module_code, record_id, action, comment, actioned_by)
                   VALUES ('VCN01', %s, 'Back to Draft', %s, %s)""", (record_id, comment, username))
    conn.commit()
    conn.close()


def get_approval_log(record_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""SELECT action, comment, actioned_by,
                          to_char(actioned_at, 'DD-MM-YYYY HH24:MI') AS actioned_at
                   FROM approval_log WHERE module_code='VCN01' AND record_id=%s
                   ORDER BY actioned_at DESC""", (record_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]
