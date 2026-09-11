from database import get_db, get_cursor
from datetime import datetime
from modules.FCAM01 import model as fcam_model


# ===== CARGO BILLING HELPERS =====

def _mark_cargo_source_billed(cur, cargo_source_type, cargo_source_id, bill_qty, bill_id):
    """Deprecated no-op. Billed-status now lives in the parcel_charge_billed ledger
    (see record_parcel_charge/billed_qty); the legacy declaration columns are no
    longer maintained (the export table's were dropped in jnpa35)."""
    return


def _unmark_cargo_source_billed(cur, cargo_source_type, cargo_source_id, bill_qty):
    """Deprecated no-op — see _mark_cargo_source_billed."""
    return


# ===== BILL FUNCTIONS =====

# ===== GO-LIVE CUTOVER: NUMBERING SEEDS =====
# A seed is a FLOOR, never an assignment: the next number is whichever is
# higher, the natural increment or the seed. Once real documents pass the seed
# it stops mattering, so a stale seed can never collide with a live document.

def lookup_seed(seed_type, doc_series='', financial_year=''):
    """Configured start sequence for one numbering series, or None.

    Tolerates the table not existing yet (pre-migration): a failed statement
    poisons the transaction, so roll back before returning."""
    conn = get_db()
    cur = get_cursor(conn)
    try:
        cur.execute('''SELECT start_seq FROM cutover_seed
                       WHERE seed_type=%s AND doc_series=%s AND financial_year=%s''',
                    [seed_type, doc_series or '', financial_year or ''])
        row = cur.fetchone()
    except Exception:
        conn.rollback()
        conn.close()
        return None
    conn.close()
    return int(row['start_seq']) if row else None


def next_from_seed(current_max, seed):
    """Next sequence number: the natural increment, floored at the seed."""
    nxt = int(current_max or 0) + 1
    return max(nxt, int(seed)) if seed else nxt


def next_invoice_seq(cur, doc_series, financial_year):
    """Next invoice doc_series_seq for one series + FY, seed floor applied."""
    cur.execute('SELECT MAX(doc_series_seq) AS max FROM invoice_header WHERE doc_series=%s AND financial_year=%s',
                [doc_series, financial_year])
    row = cur.fetchone()
    return next_from_seed(row['max'] if row else 0,
                          lookup_seed('invoice', doc_series, financial_year))


def would_overbill(already_billed, quantity, cap, tol=1e-6):
    """True when billing `quantity` on top of `already_billed` would push a
    parcel+service past its billable `cap`. Legitimate partial billing stays
    allowed; a stale page or double-submit replaying the same lines does not."""
    return (float(already_billed) + float(quantity)) - float(cap) > tol


def get_next_bill_number(cur=None):
    """Generate next bill number. Pass a cursor to run inside a transaction."""
    conn = None if cur is not None else get_db()
    if conn is not None:
        cur = get_cursor(conn)
    cur.execute(
        "SELECT MAX(CAST(SUBSTR(bill_number, 5) AS INTEGER)) FROM bill_header WHERE bill_number LIKE 'BILL%%'"
    )
    result = cur.fetchone()['max']
    if conn is not None:
        conn.close()
    return f"BILL{next_from_seed(result, lookup_seed('bill', 'BILL')):04d}"


def get_bill_data(page=1, size=20, status_filter=None):
    """Get paginated bills"""
    conn = get_db()
    cur = get_cursor(conn)

    where_clause = ""
    params = []
    if status_filter:
        where_clause = "WHERE b.bill_status = %s"
        params.append(status_filter)

    cur.execute(f'SELECT COUNT(*) FROM bill_header b {where_clause}', params)
    total = cur.fetchone()['count']
    cur.execute(f'''
        SELECT
            b.*,
            ca.agreement_code,
            ca.agreement_name,
            -- TCS lives on the lines, not the header; roll it up for the list.
            COALESCE((SELECT SUM(bl.tcs_amount) FROM bill_lines bl WHERE bl.bill_id = b.id), 0) AS tcs_amount,
            NULLIF(
                TRIM(
                    COALESCE(ca.agreement_code, '') ||
                    CASE
                        WHEN COALESCE(ca.agreement_name, '') <> '' THEN ' - ' || ca.agreement_name
                        ELSE ''
                    END
                ),
                ''
            ) AS agreement_display
        FROM bill_header b
        LEFT JOIN customer_agreements ca ON b.agreement_id = ca.id
        {where_clause}
        ORDER BY b.id DESC
        LIMIT %s OFFSET %s
    ''', params + [size, (page-1)*size])
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


def save_bill_header(data):
    """Save bill header"""
    conn = get_db()
    cur = get_cursor(conn)
    row_id = data.get('id')

    if row_id:
        cols = [k for k in data if k not in ['id', 'bill_number']]
        cur.execute(f'''UPDATE bill_header
            SET {', '.join([f'{c}=%s' for c in cols])}
            WHERE id=%s''',
            [data[c] for c in cols] + [row_id])
    else:
        data['bill_number'] = get_next_bill_number()
        cols = [k for k in data if k != 'id']
        cur.execute(f'''INSERT INTO bill_header
            ({', '.join(cols)})
            VALUES ({', '.join(['%s']*len(cols))})
            RETURNING id''',
            [data[c] for c in cols])
        row_id = cur.fetchone()['id']

    conn.commit()
    conn.close()
    return row_id, data.get('bill_number')


def get_bill_by_id(bill_id):
    """Get bill header by ID"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT
            b.*,
            ca.agreement_code,
            ca.agreement_name,
            NULLIF(
                TRIM(
                    COALESCE(ca.agreement_code, '') ||
                    CASE
                        WHEN COALESCE(ca.agreement_name, '') <> '' THEN ' - ' || ca.agreement_name
                        ELSE ''
                    END
                ),
                ''
            ) AS agreement_display
        FROM bill_header b
        LEFT JOIN customer_agreements ca ON b.agreement_id = ca.id
        WHERE b.id = %s
    ''', (bill_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def quantity_totals(lines):
    """Per-UOM quantity totals for a bill's lines — a single 'Total' is
    meaningless when MT and HRS sit in the same bill.
    Returns [{'uom': 'MT', 'quantity': 1234.5}, ...] ordered by UOM."""
    totals = {}
    for l in lines:
        uom = (l.get('uom') or '').strip() or '—'
        totals[uom] = totals.get(uom, 0.0) + float(l.get('quantity') or 0)
    return [{'uom': u, 'quantity': round(q, 3)} for u, q in sorted(totals.items())]


def get_bill_lines(bill_id):
    """Get all lines for a bill"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM bill_lines WHERE bill_id = %s ORDER BY id', (bill_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_bill_line(data):
    """Save bill line (supports both EU lines and service records)"""
    conn = get_db()
    cur = get_cursor(conn)
    existing_line = None

    # Look up TDS/TCS config from service master (FSTM01)
    tds_applicable = int(data.get('tds_applicable') or 0)
    tds_percent = float(data.get('tds_percent') or 0)
    tds_amount = float(data.get('tds_amount') or 0)
    tcs_applicable = int(data.get('tcs_applicable') or 0)
    tcs_percent = float(data.get('tcs_percent') or 0)
    tcs_amount = float(data.get('tcs_amount') or 0)
    service_code = data.get('service_code') or ''

    svc_id = data.get('service_type_id')
    if svc_id:
        cur.execute(
            'SELECT service_code, is_tds, tds_percent, is_tcs, tcs_percent, gst_rate_id, sac_code FROM finance_service_types WHERE id = %s',
            [svc_id]
        )
        svc = cur.fetchone()
        if svc:
            service_code = service_code or (svc.get('service_code') or '')
            if not data.get('sac_code'):
                data['sac_code'] = svc.get('sac_code') or ''
            # TDS — calculated on basic amount only
            if not data.get('tds_applicable') and svc.get('is_tds'):
                tds_applicable = 1
                tds_percent = float(svc.get('tds_percent') or 0)
                line_amount = float(data.get('line_amount') or 0)
                tds_amount = round(line_amount * tds_percent / 100, 2)
            # TCS — calculated on basic + GST (set after GST computation below)
            if not data.get('tcs_applicable') and svc.get('is_tcs'):
                tcs_applicable = 1
                tcs_percent = float(svc.get('tcs_percent') or 0)
            # GST — auto-compute if not already provided
            gst_rate_id = svc.get('gst_rate_id')
            if gst_rate_id and not data.get('cgst_amount') and not data.get('igst_amount'):
                cur.execute('SELECT cgst_rate, sgst_rate, igst_rate FROM gst_rates WHERE id = %s', [gst_rate_id])
                gst = cur.fetchone()
                if gst:
                    line_amount = float(data.get('line_amount') or 0)
                    data['gst_rate_id'] = gst_rate_id
                    # Determine CGST+SGST vs IGST
                    customer_gstin = data.get('customer_gstin') or ''
                    customer_state = data.get('customer_state_code') or ''
                    # Get port state code from FIN01 module config (seller_gstin / port_gst_state_code)
                    from database import get_module_config
                    fin_cfg = get_module_config('FIN01')
                    port_state_code = str(fin_cfg.get('port_gst_state_code') or '').strip()
                    seller_gstin = str(fin_cfg.get('seller_gstin') or '').strip()
                    # Derive port state from explicit config first, then GSTIN prefix
                    if not port_state_code and seller_gstin:
                        port_state_code = seller_gstin[:2]
                    # Compare state codes
                    if customer_state and port_state_code:
                        same_state = customer_state.strip() == port_state_code
                    elif customer_gstin and port_state_code:
                        same_state = customer_gstin[:2] == port_state_code
                    else:
                        # Cannot determine — default to intra-state (safer: no IGST surprise)
                        same_state = True
                    if same_state:
                        # Intra-state: CGST + SGST
                        data['cgst_rate'] = float(gst['cgst_rate'] or 0)
                        data['sgst_rate'] = float(gst['sgst_rate'] or 0)
                        data['igst_rate'] = 0
                        data['cgst_amount'] = round(line_amount * data['cgst_rate'] / 100, 2)
                        data['sgst_amount'] = round(line_amount * data['sgst_rate'] / 100, 2)
                        data['igst_amount'] = 0
                    else:
                        # Inter-state: IGST
                        data['cgst_rate'] = 0
                        data['sgst_rate'] = 0
                        data['igst_rate'] = float(gst['igst_rate'] or 0)
                        data['cgst_amount'] = 0
                        data['sgst_amount'] = 0
                        data['igst_amount'] = round(line_amount * data['igst_rate'] / 100, 2)

    # Compute line_total = line_amount + GST
    la = float(data.get('line_amount') or 0)
    ca = float(data.get('cgst_amount') or 0)
    sa = float(data.get('sgst_amount') or 0)
    ia = float(data.get('igst_amount') or 0)
    data['line_total'] = round(la + ca + sa + ia, 2)

    # TCS — calculated on basic + GST
    if tcs_applicable and tcs_percent > 0:
        tcs_amount = round((la + ca + sa + ia) * tcs_percent / 100, 2)

    if data.get('id'):
        cur.execute(
            'SELECT cargo_source_type, cargo_source_id, quantity FROM bill_lines WHERE id=%s',
            [data['id']]
        )
        existing_line = cur.fetchone()
        if existing_line:
            _unmark_cargo_source_billed(
                cur,
                existing_line.get('cargo_source_type'),
                existing_line.get('cargo_source_id'),
                float(existing_line.get('quantity') or 0)
            )
        cur.execute('''UPDATE bill_lines
            SET cargo_source_type=%s, cargo_source_id=%s, service_record_id=%s, service_type_id=%s, service_name=%s,
                service_description=%s, quantity=%s, uom=%s, rate=%s, line_amount=%s,
                gst_rate_id=%s, cgst_rate=%s, sgst_rate=%s, igst_rate=%s,
                cgst_amount=%s, sgst_amount=%s, igst_amount=%s,
                line_total=%s, gl_code=%s, sac_code=%s, remarks=%s,
                service_code=%s, tds_applicable=%s, tds_percent=%s, tds_amount=%s,
                tcs_applicable=%s, tcs_percent=%s, tcs_amount=%s
            WHERE id=%s''',
            [data.get('cargo_source_type'), data.get('cargo_source_id'), data.get('service_record_id'),
             data.get('service_type_id'), data.get('service_name'),
             data.get('service_description'), data.get('quantity'), data.get('uom'),
             data.get('rate'), data.get('line_amount'), data.get('gst_rate_id'),
             data.get('cgst_rate'), data.get('sgst_rate'), data.get('igst_rate'),
             data.get('cgst_amount'), data.get('sgst_amount'), data.get('igst_amount'),
             data.get('line_total'), data.get('gl_code'), data.get('sac_code'),
             data.get('remarks'), service_code, tds_applicable, tds_percent, tds_amount,
             tcs_applicable, tcs_percent, tcs_amount,
             data['id']])
        row_id = data['id']
    else:
        cur.execute('''INSERT INTO bill_lines
            (bill_id, cargo_source_type, cargo_source_id, service_record_id, service_type_id, service_name,
             service_description, quantity, uom, rate, line_amount, gst_rate_id,
             cgst_rate, sgst_rate, igst_rate, cgst_amount, sgst_amount, igst_amount,
             line_total, gl_code, sac_code, remarks,
             service_code, tds_applicable, tds_percent, tds_amount,
             tcs_applicable, tcs_percent, tcs_amount)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id''',
            [data['bill_id'], data.get('cargo_source_type'), data.get('cargo_source_id'), data.get('service_record_id'),
             data.get('service_type_id'), data.get('service_name'),
             data.get('service_description'),
             data.get('quantity'), data.get('uom'), data.get('rate'), data.get('line_amount'),
             data.get('gst_rate_id'), data.get('cgst_rate'), data.get('sgst_rate'),
             data.get('igst_rate'), data.get('cgst_amount'), data.get('sgst_amount'),
             data.get('igst_amount'), data.get('line_total'), data.get('gl_code'),
             data.get('sac_code'), data.get('remarks'),
             service_code, tds_applicable, tds_percent, tds_amount,
             tcs_applicable, tcs_percent, tcs_amount])
        row_id = cur.fetchone()['id']

    # Mark cargo declaration source as billed
    _mark_cargo_source_billed(
        cur,
        data.get('cargo_source_type'),
        data.get('cargo_source_id'),
        float(data.get('quantity') or 0),
        data.get('bill_id')
    )

    # Mark the service record as billed if service_record_id is provided
    if data.get('service_record_id'):
        cur.execute('UPDATE service_records SET is_billed = 1, bill_id = %s WHERE id = %s',
                     [data.get('bill_id'), data.get('service_record_id')])

    conn.commit()
    conn.close()
    return row_id


def delete_bill_line(row_id):
    """Delete bill line and reverse billed tracking on cargo source."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(
        'SELECT cargo_source_type, cargo_source_id, quantity, service_record_id FROM bill_lines WHERE id=%s',
        (row_id,)
    )
    bl = cur.fetchone()
    if bl:
        _unmark_cargo_source_billed(
            cur,
            bl['cargo_source_type'],
            bl['cargo_source_id'],
            float(bl['quantity'] or 0)
        )
        if bl.get('service_record_id'):
            cur.execute(
                'UPDATE service_records SET is_billed=0, bill_id=NULL WHERE id=%s',
                [bl['service_record_id']]
            )
    cur.execute('DELETE FROM bill_lines WHERE id=%s', (row_id,))
    conn.commit()
    conn.close()


def delete_bill(bill_id):
    """Delete bill header and all lines"""
    conn = get_db()
    cur = get_cursor(conn)
    # Reverse billed tracking on cargo declaration tables
    cur.execute('''
        SELECT cargo_source_type, cargo_source_id, quantity
        FROM bill_lines
        WHERE bill_id = %s AND cargo_source_type IS NOT NULL AND cargo_source_id IS NOT NULL
    ''', (bill_id,))
    for row in cur.fetchall():
        _unmark_cargo_source_billed(
            cur,
            row['cargo_source_type'],
            row['cargo_source_id'],
            float(row['quantity'] or 0)
        )
    # Unmark service records as billed
    cur.execute('''UPDATE service_records SET is_billed=0, bill_id=NULL
        WHERE bill_id IN (SELECT id FROM bill_header WHERE id=%s)''', (bill_id,))
    # Delete bill (cascades to lines)
    cur.execute('DELETE FROM bill_header WHERE id=%s', (bill_id,))
    conn.commit()
    conn.close()


# ===== INVOICE FUNCTIONS =====

def get_next_invoice_number(series='INV'):
    """Generate next invoice number"""
    year = datetime.now().year
    prefix = f"{series}{year}-"

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute(
        "SELECT MAX(CAST(SUBSTR(invoice_number, LENGTH(%s) + 1) AS INTEGER)) FROM invoice_header WHERE invoice_number LIKE %s",
        [prefix, f"{prefix}%"]
    )
    result = cur.fetchone()['max']
    conn.close()
    next_num = (result or 0) + 1
    return f"{prefix}{next_num:04d}"


def get_financial_year(date_str):
    """Get financial year from date (FY runs Apr-Mar)"""
    dt = datetime.strptime(date_str, '%Y-%m-%d')
    if dt.month >= 4:
        return f"{dt.year}-{str(dt.year + 1)[2:]}"
    else:
        return f"{dt.year - 1}-{str(dt.year)[2:]}"


def get_invoice_data(page=1, size=20, status_filter=None):
    """Get paginated invoices"""
    conn = get_db()
    cur = get_cursor(conn)

    where_clause = ""
    params = []
    if status_filter:
        where_clause = "WHERE invoice_status = %s"
        params.append(status_filter)

    cur.execute(f'SELECT COUNT(*) FROM invoice_header {where_clause}', params)
    total = cur.fetchone()['count']
    cur.execute(f'''
        SELECT * FROM invoice_header {where_clause}
        ORDER BY id DESC
        LIMIT %s OFFSET %s
    ''', params + [size, (page-1)*size])
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows], total


def create_invoice_from_bills(bill_ids, invoice_data):
    """Create invoice from approved bills"""
    conn = get_db()
    cur = get_cursor(conn)

    # Get customer details from first bill (all bills should be for same customer)
    cur.execute('SELECT * FROM bill_header WHERE id=%s', (bill_ids[0],))
    first_bill = dict(cur.fetchone())

    # Add customer details from bill to invoice_data
    invoice_data['customer_id'] = first_bill['customer_id']
    invoice_data['customer_type'] = first_bill['customer_type']
    invoice_data['customer_name'] = first_bill['customer_name']
    invoice_data['customer_gstin'] = first_bill['customer_gstin']
    invoice_data['customer_gst_state_code'] = first_bill['customer_gst_state_code']
    invoice_data['customer_gl_code'] = first_bill['customer_gl_code']

    # Generate invoice number and FY
    if invoice_data.get('_invoice_number_override'):
        invoice_number = invoice_data.pop('_invoice_number_override')
    else:
        invoice_number = get_next_invoice_number(invoice_data.get('invoice_series', 'INV'))
    financial_year = get_financial_year(invoice_data['invoice_date'])

    invoice_data['invoice_number'] = invoice_number
    invoice_data['financial_year'] = financial_year

    # Insert invoice header
    cols = [k for k in invoice_data if k not in ('id', '_invoice_number_override')]
    cur.execute(f'''INSERT INTO invoice_header
        ({', '.join(cols)})
        VALUES ({', '.join(['%s']*len(cols))})
        RETURNING id''',
        [invoice_data[c] for c in cols])
    invoice_id = cur.fetchone()['id']

    # Copy bill lines to invoice lines
    line_number = 1
    for bill_id in bill_ids:
        # Get bill details
        cur.execute('SELECT * FROM bill_header WHERE id=%s', (bill_id,))
        bill = dict(cur.fetchone())

        # Create mapping entry
        cur.execute('''INSERT INTO invoice_bill_mapping
            (invoice_id, bill_id, bill_number, bill_amount)
            VALUES (%s, %s, %s, %s)''',
            [invoice_id, bill_id, bill['bill_number'], bill['total_amount']])

        # Copy bill lines to invoice lines
        cur.execute('SELECT * FROM bill_lines WHERE bill_id=%s', (bill_id,))
        bill_lines = cur.fetchall()
        for bl in bill_lines:
            bl = dict(bl)
            cur.execute('''INSERT INTO invoice_lines
                (invoice_id, bill_id, bill_number, line_number, service_name, service_description,
                 quantity, uom, rate, line_amount, cgst_rate, sgst_rate, igst_rate,
                 cgst_amount, sgst_amount, igst_amount, line_total, gl_code, sac_code,
                 profit_center, cost_center,
                 service_code, tds_applicable, tds_percent, tds_amount,
                 tcs_applicable, tcs_percent, tcs_amount)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)''',
                [invoice_id, bill_id, bill['bill_number'], line_number, bl['service_name'],
                 bl['service_description'], bl['quantity'], bl['uom'], bl['rate'],
                 bl['line_amount'], bl['cgst_rate'], bl['sgst_rate'], bl['igst_rate'],
                 bl['cgst_amount'], bl['sgst_amount'], bl['igst_amount'], bl['line_total'],
                 bl['gl_code'], bl['sac_code'], invoice_data.get('profit_center'),
                 invoice_data.get('cost_center'),
                 bl.get('service_code'), bl.get('tds_applicable', 0),
                 bl.get('tds_percent', 0), bl.get('tds_amount', 0),
                 bl.get('tcs_applicable', 0), bl.get('tcs_percent', 0),
                 bl.get('tcs_amount', 0)])
            line_number += 1

        # Mark bill as invoiced
        cur.execute("UPDATE bill_header SET bill_status='Invoiced' WHERE id=%s", (bill_id,))

    # Auto-calculate invoice header tds_amount and tcs_amount from line totals
    cur.execute(
        'SELECT COALESCE(SUM(tds_amount), 0) AS total_tds, COALESCE(SUM(tcs_amount), 0) AS total_tcs FROM invoice_lines WHERE invoice_id = %s',
        [invoice_id]
    )
    row = cur.fetchone()
    total_tds = row['total_tds']
    total_tcs = row['total_tcs']
    if total_tds or total_tcs:
        cur.execute(
            'UPDATE invoice_header SET tds_amount = %s, tcs_amount = %s WHERE id = %s',
            [total_tds, total_tcs, invoice_id]
        )

    conn.commit()
    conn.close()
    return invoice_id, invoice_number


def get_invoice_lines(invoice_id):
    """Get all lines for an invoice"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM invoice_lines WHERE invoice_id = %s ORDER BY line_number', (invoice_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_invoice_bills(invoice_id):
    """Get all bills included in an invoice"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT m.*, b.bill_date, b.customer_name
        FROM invoice_bill_mapping m
        JOIN bill_header b ON m.bill_id = b.id
        WHERE m.invoice_id = %s
        ORDER BY b.bill_date
    ''', (invoice_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_invoice_by_id(invoice_id):
    """Get invoice header by ID"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT * FROM invoice_header WHERE id = %s', (invoice_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_invoice_sac_summary(invoice_id):
    """Get SAC-wise summary for invoice (grouped by SAC code)"""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT
            il.sac_code,
            SUM(il.line_amount) as taxable_value,
            SUM(il.cgst_amount) as cgst,
            SUM(il.sgst_amount) as sgst,
            SUM(il.igst_amount) as igst
        FROM invoice_lines il
        WHERE il.invoice_id = %s
        GROUP BY il.sac_code
        ORDER BY il.sac_code
    ''', (invoice_id,))
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ===== PARCEL CHARGE BILLED LEDGER (billed-status store for the AR engine) =====


def record_parcel_charge(cur, cargo_source_type, cargo_source_id, service_type_id,
                         service_code, bill_id, billed_quantity, created_by):
    """Ledger a billed parcel charge. Called inside the bill-generation transaction
    (takes the caller's cursor; does NOT commit)."""
    cur.execute('''INSERT INTO parcel_charge_billed
        (cargo_source_type, cargo_source_id, service_type_id, service_code,
         bill_id, billed_quantity, billed_date, created_by)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)''',
        [cargo_source_type, cargo_source_id, service_type_id, service_code,
         bill_id, billed_quantity, datetime.now().strftime('%Y-%m-%d'), created_by])


def void_bill_charges(cur, bill_id):
    """Remove all ledger rows for a bill (bill cancellation/reversal). Takes the
    caller's cursor; does NOT commit."""
    cur.execute('DELETE FROM parcel_charge_billed WHERE bill_id=%s', [bill_id])


def billed_qty(cargo_source_type, cargo_source_id, service_type_id):
    """Total quantity already billed for one parcel + service."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT COALESCE(SUM(billed_quantity), 0) AS q FROM parcel_charge_billed
                   WHERE cargo_source_type=%s AND cargo_source_id=%s AND service_type_id=%s''',
                [cargo_source_type, cargo_source_id, service_type_id])
    q = float(cur.fetchone()['q'] or 0)
    conn.close()
    return q


def is_vcn_billed(vcn_id):
    """True if any parcel of this VCN (import consigner or export cargo) has been
    billed. Used to lock a billed vessel against edits / revert-to-draft."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT EXISTS (
        SELECT 1 FROM parcel_charge_billed pcb
        WHERE (pcb.cargo_source_type = 'VCN_IMPORT'
               AND pcb.cargo_source_id IN (SELECT id FROM vcn_consigners WHERE vcn_id=%s))
           OR (pcb.cargo_source_type = 'VCN_EXPORT'
               AND pcb.cargo_source_id IN (SELECT id FROM vcn_export_cargo_declaration WHERE vcn_id=%s))
    ) AS billed''', [vcn_id, vcn_id])
    billed = bool(cur.fetchone()['billed'])
    conn.close()
    return billed


def vcn_billing_blockers(vcn_id):
    """What holds this VCN's billed-lock: the distinct bill numbers behind its
    ledger rows, and whether any row is a cutover flag (bill_id NULL — billed
    in the legacy system, no PORTMAN bill behind it)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''SELECT DISTINCT bh.bill_number, (pcb.bill_id IS NULL) AS cutover
        FROM parcel_charge_billed pcb
        LEFT JOIN bill_header bh ON bh.id = pcb.bill_id
        WHERE (pcb.cargo_source_type = 'VCN_IMPORT'
               AND pcb.cargo_source_id IN (SELECT id FROM vcn_consigners WHERE vcn_id=%s))
           OR (pcb.cargo_source_type = 'VCN_EXPORT'
               AND pcb.cargo_source_id IN (SELECT id FROM vcn_export_cargo_declaration WHERE vcn_id=%s))''',
        [vcn_id, vcn_id])
    rows = cur.fetchall()
    conn.close()
    return {
        'bill_numbers': sorted({r['bill_number'] for r in rows if r['bill_number']}),
        'cutover': any(r['cutover'] for r in rows),
    }


def generate_bill(data, created_by, bill_status, approved_by=None):
    """Create ONE bill across the selected vessels. Reuses save_bill_header
    (numbering) and save_bill_line (GST/TDS calc, mutates the line dict with the
    computed amounts), then records totals, bill_vessels, and the parcel ledger.
    Returns (bill_id, bill_number). No MBC — only VCN_IMPORT/VCN_EXPORT lines.
    Pass approved_by when the bill is born Approved (password-confirmed)."""
    lines = data.get('lines') or []
    if not lines:
        raise ValueError('No lines to bill')
    for l in lines:
        if not l.get('cargo_source_type') or not l.get('cargo_source_id') or not l.get('vcn_id'):
            raise ValueError('Each bill line needs cargo_source_type, cargo_source_id and vcn_id')
    vcn_ids = sorted({l['vcn_id'] for l in lines if l.get('vcn_id')})

    conn = get_db()
    cur = get_cursor(conn)
    docs = []
    if vcn_ids:
        cur.execute("SELECT vcn_doc_num FROM vcn_header WHERE id = ANY(%s) ORDER BY vcn_doc_num", [vcn_ids])
        docs = [r['vcn_doc_num'] for r in cur.fetchall() if r['vcn_doc_num']]

    # Overbill guard: a stale page or a double-submit must not bill the same
    # parcel twice. Declared quantity is the cap — partial billing stays legal.
    caps = {}
    for src, table in (('VCN_IMPORT', 'vcn_consigners'),
                       ('VCN_EXPORT', 'vcn_export_cargo_declaration')):
        ids = sorted({l['cargo_source_id'] for l in lines if l.get('cargo_source_type') == src})
        if not ids:
            continue
        cur.execute(f'SELECT id, quantity, parcel_no FROM {table} WHERE id = ANY(%s)', [ids])
        for r in cur.fetchall():
            caps[(src, r['id'])] = (_to_float(r['quantity']), r['parcel_no'])
    cur.execute('''SELECT cargo_source_type, cargo_source_id, service_type_id,
                          COALESCE(SUM(billed_quantity), 0) AS q
                   FROM parcel_charge_billed GROUP BY 1, 2, 3''')
    already = {(r['cargo_source_type'], r['cargo_source_id'], r['service_type_id']): float(r['q'] or 0)
               for r in cur.fetchall()}
    conn.close()

    for l in lines:
        key = (l['cargo_source_type'], l['cargo_source_id'])
        cap, parcel_no = caps.get(key, (None, None))
        if cap is None:
            continue
        prior = already.get(key + (l.get('service_type_id'),), 0.0)
        if would_overbill(prior, l.get('quantity') or 0, cap):
            raise ValueError(
                f"Parcel {parcel_no or l['cargo_source_id']} / {l.get('service_code') or 'service'} "
                f"is already billed for {prior:g} of {cap:g} — reload the page and bill "
                f"only the remaining quantity.")

    header = {
        'source_type': 'MULTI', 'source_id': None,
        'source_display': ', '.join(docs),
        'customer_type': data.get('customer_type'), 'customer_id': data.get('customer_id'),
        'customer_name': data.get('customer_name'), 'customer_gstin': data.get('customer_gstin'),
        'customer_gst_state_code': data.get('customer_gst_state_code'),
        'customer_gl_code': data.get('customer_gl_code'),
        'currency_code': data.get('currency_code') or 'INR',
        'agreement_id': data.get('agreement_id') or None,
        'bill_status': bill_status,
        'bill_date': data.get('bill_date') or datetime.now().strftime('%Y-%m-%d'),
        'created_by': created_by,
        'created_date': datetime.now().strftime('%Y-%m-%d'),
    }
    if approved_by:
        header['approved_by'] = approved_by
        header['approved_date'] = datetime.now().strftime('%Y-%m-%d')
    bill_id, bill_number = save_bill_header(header)

    subtotal = cgst = sgst = igst = 0.0
    for l in lines:
        line_amount = round(float(l.get('quantity') or 0) * float(l.get('rate') or 0), 2)
        ld = {
            'bill_id': bill_id, 'service_type_id': l.get('service_type_id'),
            'service_code': l.get('service_code'), 'service_name': l.get('service_name'),
            'service_description': l.get('service_name'),
            'quantity': l.get('quantity'), 'uom': l.get('uom'), 'rate': l.get('rate'),
            'line_amount': line_amount, 'gst_rate_id': l.get('gst_rate_id'),
            'sac_code': l.get('sac_code'), 'gl_code': l.get('gl_code'),
            'tds_applicable': l.get('tds_applicable'), 'tds_percent': l.get('tds_percent'),
            'cargo_source_type': l.get('cargo_source_type'), 'cargo_source_id': l.get('cargo_source_id'),
            'customer_gstin': data.get('customer_gstin'),
            'customer_state_code': data.get('customer_gst_state_code'),
        }
        save_bill_line(ld)  # computes + stores cgst/sgst/igst/tds/line_total on ld and the row
        subtotal += line_amount
        cgst += float(ld.get('cgst_amount') or 0)
        sgst += float(ld.get('sgst_amount') or 0)
        igst += float(ld.get('igst_amount') or 0)

    total = round(subtotal + cgst + sgst + igst, 2)
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE bill_header
        SET subtotal=%s, cgst_amount=%s, sgst_amount=%s, igst_amount=%s, total_amount=%s
        WHERE id=%s''', [subtotal, cgst, sgst, igst, total, bill_id])
    for vid in vcn_ids:
        cur.execute('INSERT INTO bill_vessels (bill_id, vcn_id) VALUES (%s, %s)', [bill_id, vid])
    for l in lines:
        record_parcel_charge(cur, l.get('cargo_source_type'), l.get('cargo_source_id'),
                             l.get('service_type_id'), l.get('service_code'), bill_id,
                             float(l.get('quantity') or 0), created_by)
    conn.commit()
    conn.close()
    return bill_id, bill_number


# ===== BILLABLES ENGINE (parcels -> 4 charges, grouped by vessel) =====

_CARGO_GATE = ('Closed', 'Partial Close')

# Order the billables screen and the pro-forma list services in: cargo handling,
# then infrastructure, then the extras. Rows for one service stay together.
_SERVICE_DISPLAY_ORDER = ['CHGU01', 'CHGL01', 'INFM01', 'MLAC01', 'TOLL01']


def parcel_charge_codes(src, equipment_names, toll_applicable):
    """Service codes one parcel yields: cargo handling (by direction) plus
    infrastructure, then equipment and toll when applicable. Single source of
    truth for the billables engine, the picker counts and Admin Cutover."""
    codes = ['CHGU01' if src == 'VCN_IMPORT' else 'CHGL01', 'INFM01']
    if (equipment_names or '').strip():
        codes.append('MLAC01')
    if toll_applicable:
        codes.append('TOLL01')
    return codes


def _to_float(v):
    try:
        return float(str(v).replace(',', '')) if v not in (None, '') else 0.0
    except (ValueError, TypeError):
        return 0.0


def _actual_qty_map(cur, ldud_id, declared_by_parcel):
    """Actual handled quantity per parcel id for one LDUD, from the LUEU01
    logbook (short-close rows excluded — that quantity was never handled).
    Merged ops (one op covering several same-cargo parcels) are apportioned
    pro-rata by declared parcel quantity. Parcels not covered by any op are
    absent from the map (caller falls back to the declared quantity)."""
    cur.execute('''SELECT po.parcel_ids, COALESCE(SUM(lg.quantity), 0) AS q
                   FROM ldud_parcel_ops po
                   LEFT JOIN lueu_parcel_log lg
                          ON lg.parcel_op_id = po.id
                         AND COALESCE(lg.is_deleted, FALSE) = FALSE
                         AND COALESCE(lg.is_shortclose, FALSE) = FALSE
                   WHERE po.ldud_id = %s
                   GROUP BY po.id, po.parcel_ids''', [ldud_id])
    actual = {}
    for r in cur.fetchall():
        ids = [int(x) for x in str(r['parcel_ids'] or '').split(',') if str(x).strip().isdigit()]
        if not ids:
            continue
        q = float(r['q'] or 0)
        weights = [max(_to_float(declared_by_parcel.get(pid)), 0.0) for pid in ids]
        wsum = sum(weights)
        for pid, w in zip(ids, weights):
            # ponytail: pro-rata by declared qty; equal split when nothing declared
            share = q * (w / wsum) if wsum > 0 else q / len(ids)
            actual[pid] = actual.get(pid, 0.0) + share
    return actual


def unclosed_vcn_docs(vcn_ids):
    """VCN doc nums whose latest LDUD is NOT Closed/Partial Close (incl. no
    LDUD at all) — these may only get a pro forma, never an actual bill."""
    if not vcn_ids:
        return []
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT h.vcn_doc_num
        FROM vcn_header h
        LEFT JOIN (SELECT DISTINCT ON (vcn_id) vcn_id, doc_status
                   FROM ldud_header ORDER BY vcn_id, id DESC) ll ON ll.vcn_id = h.id
        WHERE h.id = ANY(%s)
          AND (ll.doc_status IS NULL OR NOT (ll.doc_status = ANY(%s)))
    ''', [list(vcn_ids), list(_CARGO_GATE)])
    docs = [r['vcn_doc_num'] for r in cur.fetchall()]
    conn.close()
    return docs


def verify_user_password(user_id, password):
    """Re-authenticate the logged-in user (same plaintext check as login)."""
    if not password:
        return False
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT 1 AS ok FROM users WHERE id=%s AND password=%s', [user_id, password])
    ok = cur.fetchone() is not None
    conn.close()
    return ok


def get_customer_billables(customer_type, customer_id):
    """Billable charges for a customer's parcels, grouped by vessel. Read-only.
    Bills the payer (importer_name). Two stages per vessel:
      'proforma' — VCN Approved but LDUD not yet Closed: declared quantities,
                   pro forma invoice only (no bill generation);
      'actual'   — latest LDUD Closed/Partial Close: quantities come from the
                   LUEU01 logbook (fallback to declared when no logbook data).
    Remaining per charge comes from the parcel_charge_billed ledger."""
    conn = get_db()
    cur = get_cursor(conn)

    if customer_type == 'Customer':
        cur.execute("SELECT name FROM vessel_customers WHERE id=%s", [customer_id])
    else:
        cur.execute("SELECT name FROM vessel_agents WHERE id=%s", [customer_id])
    row = cur.fetchone()
    customer_name = row['name'] if row else ''

    # GST rates come along so the billables screen can total the tax live,
    # instead of the user finding out what GST is only after generating.
    cur.execute("""SELECT s.id, s.service_code, s.service_name, s.sac_code, s.uom,
                          s.gst_rate_id, s.is_tds, s.tds_percent, s.is_tcs, s.tcs_percent,
                          g.cgst_rate, g.sgst_rate, g.igst_rate
                   FROM finance_service_types s
                   LEFT JOIN gst_rates g ON g.id = s.gst_rate_id
                   WHERE s.service_code IN ('CHGU01','CHGL01','INFM01','MLAC01','TOLL01')""")
    svc = {r['service_code']: dict(r) for r in cur.fetchall()}

    cur.execute("""
        WITH ldud_latest AS (
            SELECT DISTINCT ON (vcn_id) vcn_id, id AS ldud_id, doc_status
            FROM ldud_header ORDER BY vcn_id, id DESC
        )
        SELECT 'VCN_IMPORT' AS src, c.id, c.parcel_no, c.cargo_name, c.quantity,
               c.equipment_names, c.toll_applicable,
               h.id AS vcn_id, h.vcn_doc_num, h.vessel_name,
               ll.ldud_id, ll.doc_status AS ldud_status
        FROM vcn_consigners c
        JOIN vcn_header h ON h.id = c.vcn_id
        LEFT JOIN ldud_latest ll ON ll.vcn_id = h.id
        WHERE c.importer_name = %s
          AND COALESCE(c.is_removed, FALSE) = FALSE
          AND (h.doc_status = 'Approved' OR ll.doc_status = ANY(%s))
        UNION ALL
        SELECT 'VCN_EXPORT' AS src, e.id, e.parcel_no, e.cargo_name, e.quantity,
               e.equipment_names, e.toll_applicable,
               h.id AS vcn_id, h.vcn_doc_num, h.vessel_name,
               ll.ldud_id, ll.doc_status AS ldud_status
        FROM vcn_export_cargo_declaration e
        JOIN vcn_header h ON h.id = e.vcn_id
        LEFT JOIN ldud_latest ll ON ll.vcn_id = h.id
        WHERE e.importer_name = %s
          AND COALESCE(e.is_removed, FALSE) = FALSE
          AND (h.doc_status = 'Approved' OR ll.doc_status = ANY(%s))
        ORDER BY vcn_doc_num, parcel_no
    """, [customer_name, list(_CARGO_GATE), customer_name, list(_CARGO_GATE)])
    parcels = [dict(r) for r in cur.fetchall()]

    # Actual handled qty (LUEU01) per parcel, for closed-LDUD vessels.
    # ponytail: keyed by bare parcel id — one VCN's LDUD ops reference only one
    # source table (its operation_type), so cross-table id collisions don't occur.
    by_vcn = {}
    for p in parcels:
        by_vcn.setdefault(p['vcn_id'], []).append(p)
    actual_by_vcn = {}
    for vcn_id, plist in by_vcn.items():
        if plist[0]['ldud_status'] in _CARGO_GATE and plist[0]['ldud_id']:
            declared = {p['id']: p['quantity'] for p in plist}
            actual_by_vcn[vcn_id] = _actual_qty_map(cur, plist[0]['ldud_id'], declared)

    # Ledger + rates in two batch queries on this connection. Doing them per
    # charge cost a fresh DB connection each (~180ms) and dominated the whole
    # call — pricing N parcels was O(N) connections, not O(N) queries.
    billed = {}
    if parcels:
        cur.execute('''SELECT cargo_source_type, cargo_source_id, service_type_id,
                              COALESCE(SUM(billed_quantity), 0) AS q
                       FROM parcel_charge_billed
                       WHERE (cargo_source_type, cargo_source_id) IN %s
                       GROUP BY 1, 2, 3''',
                    [tuple((p['src'], p['id']) for p in parcels)])
        billed = {(r['cargo_source_type'], r['cargo_source_id'], r['service_type_id']): float(r['q'] or 0)
                  for r in cur.fetchall()}
    rates_by_cargo, rates_by_service = fcam_model.get_customer_rates_map(
        customer_type, customer_id, cur=cur)
    conn.close()

    vessels = {}
    for p in parcels:
        src = p['src']
        stage = 'actual' if p['ldud_status'] in _CARGO_GATE else 'proforma'
        declared = _to_float(p['quantity'])
        actual = None
        if stage == 'actual':
            aq = actual_by_vcn.get(p['vcn_id'], {}).get(p['id'])
            actual = round(aq, 3) if aq is not None else None
        qty = actual if actual is not None else declared

        # (service_code, cargo_name_for_rate) — cargo_name only for cargo-priced services
        charges = [(code, p['cargo_name'] if code in ('CHGU01', 'CHGL01', 'INFM01') else None)
                   for code in parcel_charge_codes(src, p['equipment_names'], p['toll_applicable'])]

        v = vessels.setdefault(p['vcn_id'], {
            'vcn_id': p['vcn_id'], 'vcn_doc_num': p['vcn_doc_num'],
            'vessel_name': p['vessel_name'], 'ldud_status': p['ldud_status'],
            'stage': stage,
            'lines': [], 'total_amount': 0.0,
        })
        for code, cargo_for_rate in charges:
            st = svc.get(code)
            if not st:
                continue
            remaining = round(qty - billed.get((src, p['id'], st['id']), 0.0), 3)
            if remaining <= 1e-6:
                continue
            # Cargo-specific rate first, then any rate for the service — the
            # same precedence get_customer_rate applies (it skips the
            # cargo-specific lookup entirely when there is no cargo name).
            rate_info = ((rates_by_cargo.get((st['id'], cargo_for_rate)) if cargo_for_rate else None)
                         or rates_by_service.get(st['id']))
            rate = float(rate_info['rate']) if rate_info and rate_info.get('rate') is not None else 0.0
            amount = round(remaining * rate, 2)
            v['lines'].append({
                'cargo_source_type': src, 'cargo_source_id': p['id'],
                'parcel_no': p['parcel_no'], 'service_type_id': st['id'],
                'service_code': code, 'service_name': st['service_name'],
                'cargo_name': p['cargo_name'] or '', 'qty': remaining,
                'declared_qty': declared, 'actual_qty': actual,
                'uom': st['uom'] or 'MT', 'rate': rate, 'amount': amount,
                'sac_code': st['sac_code'] or '', 'gst_rate_id': st['gst_rate_id'],
                # None (not 0) when the service has no usable GST rate — the
                # screen must say "not configured", never quietly show 0.00
                'cgst_rate': _to_float(st['cgst_rate']) if st['cgst_rate'] is not None else None,
                'sgst_rate': _to_float(st['sgst_rate']) if st['sgst_rate'] is not None else None,
                'igst_rate': _to_float(st['igst_rate']) if st['igst_rate'] is not None else None,
                'is_tds': st['is_tds'], 'tds_percent': float(st['tds_percent'] or 0),
                'is_tcs': st['is_tcs'], 'tcs_percent': float(st['tcs_percent'] or 0),
            })
            v['total_amount'] = round(v['total_amount'] + amount, 2)

    # Club each vessel's lines by service instead of the per-parcel hopscotch
    # the loop above emits (P1/handling, P1/infra, P2/handling…). sort is
    # stable, so parcels keep their order inside each service block.
    for v in vessels.values():
        v['lines'].sort(key=lambda l: _SERVICE_DISPLAY_ORDER.index(l['service_code'])
                        if l['service_code'] in _SERVICE_DISPLAY_ORDER else 99)

    return {'vessels': list(vessels.values())}


def customers_with_billables():
    """{payer_name: {'actual_count': n, 'proforma_count': n}} — parcels that
    still have quantity to bill, counted per payer (importer_name).

    Same rules as get_customer_billables, reusing _actual_qty_map so the two
    can't drift: declared quantity while the LDUD is open ('proforma'), LUEU01
    actual once it is Closed/Partial Close ('actual'), minus what the
    parcel_charge_billed ledger already records. Counts parcels, not charges —
    the number matches what the vessel list shows.
    """
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("""
        WITH ldud_latest AS (
            SELECT DISTINCT ON (vcn_id) vcn_id, id AS ldud_id, doc_status
            FROM ldud_header ORDER BY vcn_id, id DESC
        ),
        billed AS (
            SELECT cargo_source_type, cargo_source_id,
                   COALESCE(SUM(billed_quantity), 0) AS q
            FROM parcel_charge_billed
            GROUP BY cargo_source_type, cargo_source_id
        )
        SELECT 'VCN_IMPORT' AS src, c.id, c.importer_name AS payer, c.quantity,
               c.equipment_names, c.toll_applicable,
               ll.ldud_id, ll.doc_status AS ldud_status, COALESCE(b.q, 0) AS billed
        FROM vcn_consigners c
        JOIN vcn_header h ON h.id = c.vcn_id
        LEFT JOIN ldud_latest ll ON ll.vcn_id = h.id
        LEFT JOIN billed b ON b.cargo_source_type = 'VCN_IMPORT' AND b.cargo_source_id = c.id
        WHERE COALESCE(c.importer_name, '') <> ''
          AND (h.doc_status = 'Approved' OR ll.doc_status = ANY(%s))
        UNION ALL
        SELECT 'VCN_EXPORT' AS src, e.id, e.importer_name AS payer, e.quantity,
               e.equipment_names, e.toll_applicable,
               ll.ldud_id, ll.doc_status AS ldud_status, COALESCE(b.q, 0) AS billed
        FROM vcn_export_cargo_declaration e
        JOIN vcn_header h ON h.id = e.vcn_id
        LEFT JOIN ldud_latest ll ON ll.vcn_id = h.id
        LEFT JOIN billed b ON b.cargo_source_type = 'VCN_EXPORT' AND b.cargo_source_id = e.id
        WHERE COALESCE(e.importer_name, '') <> ''
          AND (h.doc_status = 'Approved' OR ll.doc_status = ANY(%s))
    """, [list(_CARGO_GATE), list(_CARGO_GATE)])
    parcels = [dict(r) for r in cur.fetchall()]

    declared_by_ldud = {}
    for p in parcels:
        if p['ldud_status'] in _CARGO_GATE and p['ldud_id']:
            declared_by_ldud.setdefault(p['ldud_id'], {})[p['id']] = p['quantity']
    actual_by_ldud = {lid: _actual_qty_map(cur, lid, decl)
                      for lid, decl in declared_by_ldud.items()}
    conn.close()

    counts = {}
    for p in parcels:
        actual = p['ldud_status'] in _CARGO_GATE
        qty = _to_float(p['quantity'])
        if actual:
            aq = actual_by_ldud.get(p['ldud_id'], {}).get(p['id'])
            if aq is not None:
                qty = round(aq, 3)
        charges = len(parcel_charge_codes(p['src'], p['equipment_names'], p['toll_applicable']))
        if qty * charges - float(p['billed'] or 0) <= 1e-6:
            continue  # every applicable charge is fully billed
        c = counts.setdefault(p['payer'], {'actual_count': 0, 'proforma_count': 0})
        c['actual_count' if actual else 'proforma_count'] += 1
    return counts


def unbill_invoice_sources(cur, invoice_id):
    """Unbill everything behind a cancelled invoice so the cargo can be re-billed.

    For each bill linked via invoice_bill_mapping: reverse the parcel_charge_billed
    ledger (this releases the VCN billed-lock — the ledger is the authoritative
    billed-status store; the legacy declaration-column helpers are no-ops now),
    unmark any service records, and set the bill to 'Cancelled' (kept for audit).
    Runs inside the caller's transaction — takes a cursor and does NOT commit.
    Returns the list of affected bill numbers for the cancellation remark."""
    cur.execute('''SELECT ibm.bill_id, ibm.bill_number
        FROM invoice_bill_mapping ibm WHERE ibm.invoice_id = %s''', [invoice_id])
    bills = [dict(r) for r in cur.fetchall()]
    for bill in bills:
        bill_id = bill['bill_id']
        void_bill_charges(cur, bill_id)  # reverse the parcel ledger (releases billed-lock)
        cur.execute('UPDATE service_records SET is_billed=0, bill_id=NULL WHERE bill_id=%s', [bill_id])
        cur.execute("UPDATE bill_header SET bill_status='Cancelled' WHERE id=%s", [bill_id])
    return [b['bill_number'] for b in bills]
