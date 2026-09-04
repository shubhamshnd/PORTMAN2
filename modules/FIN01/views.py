from flask import render_template, request, redirect, url_for, session, jsonify
from . import bp
from . import model
from database import get_user_permissions, get_db, get_cursor, get_module_config
from mail_service import notify_module_approver, get_module_approver_info, build_approval_mail_html


def _queue_bill_approval_request(bill_id, bill_number, customer_name, total_amount):
    info = get_module_approver_info('FIN01')
    if not info.get('approval_add'):
        return
    bill_url = request.host_url.rstrip('/') + url_for('FIN01.view_bill', bill_id=bill_id)
    notify_module_approver(
        module_code='FIN01',
        ref_id=bill_id,
        subject=f"[Portbird DPPL] Bill {bill_number} — Pending Approval",
        body_html=build_approval_mail_html(
            approver_name=info.get('username'),
            action_label='Pending Approval',
            subtitle='Billing — Approval Required',
            details=[
                ('Bill No',       bill_number or '—'),
                ('Customer',      customer_name or '—'),
                ('Total Amount',  f'₹ {float(total_amount or 0):,.2f}'),
            ],
            action_url=bill_url,
            action_btn_label='Review &amp; Approve Bill',
            submitted_by=session.get('username'),
            badge_color='#d97706',
        ),
    )

@bp.route('/module/FIN01/')
def index():
    """Main FIN01 index - redirect to bills"""
    return redirect(url_for('FIN01.bills'))


@bp.route('/module/FIN01/invoices')
def legacy_invoices():
    """Legacy invoice list route; moved to FINV01"""
    return redirect(url_for('FINV01.invoices'))


@bp.route('/module/FIN01/invoice/generate')
def legacy_generate_invoice():
    """Legacy invoice generation route; moved to FINV01"""
    return redirect(url_for('FINV01.generate_invoice'))


@bp.route('/module/FIN01/invoice/print/<int:invoice_id>')
def legacy_print_invoice(invoice_id):
    """Legacy invoice print route; moved to FINV01"""
    return redirect(url_for('FINV01.print_invoice', invoice_id=invoice_id))


@bp.route('/module/FIN01/bills')
def bills():
    """List all bills"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_user_permissions(session['user_id'], 'FIN01')
    page = int(request.args.get('page', 1))
    status_filter = request.args.get('status')
    data, total = model.get_bill_data(page, status_filter=status_filter)

    config = get_module_config('FIN01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or bool(session.get('is_admin'))

    return render_template('bills.html',
                         data=data,
                         page=page,
                         last_page=(total + 19) // 20,
                         status_filter=status_filter,
                         perms=perms,
                         is_approver=is_approver,
                         username=session.get('username'))


@bp.route('/module/FIN01/bill/<int:bill_id>')
def view_bill(bill_id):
    """View bill details"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_user_permissions(session['user_id'], 'FIN01')

    # Get bill header
    bill = model.get_bill_by_id(bill_id)
    if not bill:
        return "Bill not found", 404

    # Get bill lines
    bill_lines = model.get_bill_lines(bill_id)

    config = get_module_config('FIN01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id) or bool(session.get('is_admin'))

    return render_template('bill_view.html',
                         bill=bill,
                         bill_lines=bill_lines,
                         qty_totals=model.quantity_totals(bill_lines),
                         perms=perms,
                         is_approver=is_approver,
                         username=session.get('username'))


@bp.route('/module/FIN01/bill/generate')
def generate_bill():
    """Generate bill - customer-centric"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    perms = get_user_permissions(session['user_id'], 'FIN01')

    from datetime import datetime
    current_date = datetime.now().strftime('%Y-%m-%d')

    return render_template('generate_bill.html',
                         current_date=current_date,
                         perms=perms,
                         username=session.get('username'))


def _proof_doc_payload(row, module_code, source_id):
    return {
        'id': row['id'],
        'original_filename': row['original_filename'],
        'uploaded_by': row['uploaded_by'],
        'uploaded_at': str(row['uploaded_at'])[:16],
        'source_module': module_code,
        'source_id': source_id,
        'file_url': f'/api/module/{module_code}/proof_docs/file/{row["id"]}',
    }


def _fetch_source_proof_docs(cur, module_code, source_id):
    if module_code == 'LDUD01':
        cur.execute('''
            SELECT id, original_filename, uploaded_by, uploaded_at
            FROM ldud_proof_documents
            WHERE ldud_id = %s
            ORDER BY uploaded_at
        ''', [source_id])
    else:
        return []
    return [_proof_doc_payload(r, module_code, source_id) for r in cur.fetchall()]


@bp.route('/api/module/FIN01/proof_docs/by_source/<source_module>/<int:source_id>')
def proof_docs_by_source(source_module, source_id):
    """Return proof documents for one LDUD source."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    module_code = {'LDUD': 'LDUD01', 'LDUD01': 'LDUD01'}.get(source_module.upper())
    if not module_code:
        return jsonify({'error': 'Invalid proof document source'}), 400

    conn = get_db()
    cur = get_cursor(conn)
    docs = _fetch_source_proof_docs(cur, module_code, source_id)
    conn.close()
    return jsonify({'docs': docs, 'source_module': module_code, 'source_id': source_id})


@bp.route('/api/module/FIN01/proof_docs/by_bill/<int:bill_id>')
def proof_docs_by_bill(bill_id):
    """Return LDUD proof documents attached to cargo lines on a bill."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT DISTINCT cargo_source_type, cargo_source_id
        FROM bill_lines
        WHERE bill_id = %s
          AND cargo_source_type IN ('VCN_IMPORT', 'VCN_EXPORT')
          AND cargo_source_id IS NOT NULL
    ''', [bill_id])
    sources = cur.fetchall()

    docs = []
    seen_sources = set()
    seen_docs = set()

    for src in sources:
        module_code = None
        source_id = None

        if src['cargo_source_type'] in ('VCN_IMPORT', 'VCN_EXPORT'):
            table = 'vcn_cargo_declaration' if src['cargo_source_type'] == 'VCN_IMPORT' else 'vcn_export_cargo_declaration'
            cur.execute(f'SELECT vcn_id FROM {table} WHERE id = %s', [src['cargo_source_id']])
            decl = cur.fetchone()
            if not decl:
                continue
            cur.execute('SELECT id FROM ldud_header WHERE vcn_id = %s ORDER BY id DESC LIMIT 1', [decl['vcn_id']])
            source = cur.fetchone()
            if source:
                module_code = 'LDUD01'
                source_id = source['id']

        source_key = (module_code, source_id)
        if not module_code or not source_id or source_key in seen_sources:
            continue
        seen_sources.add(source_key)

        for doc in _fetch_source_proof_docs(cur, module_code, source_id):
            doc_key = (doc['source_module'], doc['id'])
            if doc_key in seen_docs:
                continue
            seen_docs.add(doc_key)
            docs.append(doc)

    conn.close()
    return jsonify({'docs': docs})


@bp.route('/api/module/FIN01/bill/save', methods=['POST'])
def save_bill():
    """Save bill header"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    perms = get_user_permissions(session['user_id'], 'FIN01')
    if not perms['can_add'] and not perms['can_edit']:
        return jsonify({'success': False, 'error': 'No permission'})

    data = request.json

    # Approved/Invoiced bills are final — no edits, no revert to Draft
    if data.get('id'):
        conn = get_db()
        cur = get_cursor(conn)
        cur.execute('SELECT bill_status FROM bill_header WHERE id=%s', [data['id']])
        row = cur.fetchone()
        conn.close()
        if row and row['bill_status'] in ('Approved', 'Invoiced'):
            return jsonify({'success': False,
                            'error': f"Bill is {row['bill_status']} — it is locked and cannot be modified"})

    # Extract lines from data before saving header (lines belong to bill_lines table, not bill_header)
    lines = data.pop('lines', [])

    data['created_by'] = session.get('username')
    data['created_date'] = __import__('datetime').datetime.now().strftime('%Y-%m-%d')

    # Set bill status based on approval config
    config = get_module_config('FIN01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id)
    is_admin = session.get('is_admin')

    if config.get('approval_add'):
        data['bill_status'] = 'Pending Approval'
    else:
        data['bill_status'] = 'Draft'

    # Get source display name if not provided
    if not data.get('source_display') and data.get('source_type') and data.get('source_id'):
        conn = get_db()
        cur = get_cursor(conn)
        if data['source_type'] == 'VCN':
            cur.execute('SELECT vcn_doc_num FROM vcn_header WHERE id=%s', (data['source_id'],))
            row = cur.fetchone()
            data['source_display'] = row['vcn_doc_num'] if row else ''
        conn.close()

    # Extract fields not in bill_header table before saving
    customer_state_code = data.pop('customer_state_code', '') or ''

    row_id, bill_number = model.save_bill_header(data)

    # Save bill lines and calculate totals
    subtotal = 0
    cgst_total = 0
    sgst_total = 0
    igst_total = 0

    customer_gstin = data.get('customer_gstin') or ''

    for line in lines:
        line['bill_id'] = row_id
        line['customer_gstin'] = customer_gstin
        line['customer_state_code'] = customer_state_code
        # Map frontend field names to model field names
        if not line.get('service_name') and line.get('description'):
            line['service_name'] = line['description']
        if not line.get('service_description'):
            line['service_description'] = line.get('description', '')
        model.save_bill_line(line)
        subtotal += float(line.get('line_amount') or 0)
        cgst_total += float(line.get('cgst_amount') or 0)
        sgst_total += float(line.get('sgst_amount') or 0)
        igst_total += float(line.get('igst_amount') or 0)

    # Update bill header with calculated totals + mark source records as billed
    total_amount = subtotal + cgst_total + sgst_total + igst_total
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE bill_header
        SET subtotal=%s, cgst_amount=%s, sgst_amount=%s, igst_amount=%s, total_amount=%s
        WHERE id=%s''',
        [subtotal, cgst_total, sgst_total, igst_total, total_amount, row_id])

    # Mark service records as billed
    for line in lines:
        if line.get('line_type') == 'service_record' and line.get('service_record_id'):
            cur.execute('UPDATE service_records SET is_billed=1, bill_id=%s WHERE id=%s',
                        [row_id, line['service_record_id']])

    conn.commit()
    conn.close()

    if data.get('bill_status') == 'Pending Approval':
        _queue_bill_approval_request(row_id, bill_number, data.get('customer_name'), total_amount)

    return jsonify({'success': True, 'id': row_id, 'bill_number': bill_number})


@bp.route('/api/module/FIN01/bill/generate', methods=['POST'])
def bill_generate():
    """Generate the ACTUAL bill across selected vessels from picked billable
    lines. Password-confirmed: the user re-enters their password and the bill
    is born Approved (no Draft stage, not deletable). Only vessels whose
    latest LDUD is Closed/Partial Close can be billed — earlier the customer
    only gets a pro forma invoice."""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})
    perms = get_user_permissions(session['user_id'], 'FIN01')
    if not perms['can_add']:
        return jsonify({'success': False, 'error': 'No add permission'})

    data = request.json or {}
    lines = data.get('lines') or []
    if not lines:
        return jsonify({'success': False, 'error': 'No lines selected'})
    if any(float(l.get('rate') or 0) <= 0 for l in lines):
        return jsonify({'success': False, 'error': 'Every selected line needs a rate greater than 0'})

    if not model.verify_user_password(session['user_id'], data.get('password')):
        return jsonify({'success': False, 'error': 'Incorrect password — bill not generated'})

    vcn_ids = sorted({l.get('vcn_id') for l in lines if l.get('vcn_id')})
    unclosed = model.unclosed_vcn_docs(vcn_ids)
    if unclosed:
        return jsonify({'success': False,
                        'error': 'LDUD not closed for: ' + ', '.join(unclosed) +
                                 ' — only a pro forma invoice is available until closure'})

    try:
        bill_id, bill_number = model.generate_bill(
            data, session.get('username'), 'Approved', approved_by=session.get('username'))
    except Exception as e:
        return jsonify({'success': False, 'error': 'Generate failed: ' + str(e)})
    return jsonify({'success': True, 'id': bill_id, 'bill_number': bill_number})


def _inr(n):
    """Indian digit grouping: 305843 -> '3,05,843' (paise only when nonzero)."""
    n = round(float(n or 0), 2)
    neg = n < 0
    n = abs(n)
    r = int(n)
    p = int(round((n - r) * 100))
    s = str(r)
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ','.join(parts + [tail])
    return ('-' if neg else '') + s + (f'.{p:02d}' if p else '')


def _amount_in_words(amount):
    """Indian-system words: 305843 -> 'Rupees Three Lakh Five Thousand ... Only'."""
    ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
            'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen',
            'Seventeen', 'Eighteen', 'Nineteen']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

    def words(n):
        if n < 20:
            return ones[n]
        if n < 100:
            return (tens[n // 10] + (' ' + ones[n % 10] if n % 10 else '')).strip()
        if n < 1000:
            return (ones[n // 100] + ' Hundred' + (' ' + words(n % 100) if n % 100 else '')).strip()
        if n < 100000:
            return (words(n // 1000) + ' Thousand' + (' ' + words(n % 1000) if n % 1000 else '')).strip()
        if n < 10000000:
            return (words(n // 100000) + ' Lakh' + (' ' + words(n % 100000) if n % 100000 else '')).strip()
        return (words(n // 10000000) + ' Crore' + (' ' + words(n % 10000000) if n % 10000000 else '')).strip()

    total = int(round(float(amount or 0)))
    return 'Rupees ' + (words(total) or 'Zero') + ' Only.'


# JJLTPL stationery constants — overridable via FIN01 module config keys of the
# same name if these ever change.
_PI_GSTIN = '27AAGCJ3665D1ZK'
_PI_PAN = 'AAGCJ3665D'
_PI_PAYMENT_NOTE = ('Note : Payment to be made through DD / Bankers Cheque/RTGS drawn in favour of '
                    'JSW JNPT LIQUID TERMINAL PRIVATE LIMITED, (Axis Bank Ltd- Kalina Branch, '
                    'Mumbai – 400098, Escrow Account- 924020046923953, IFS CODE- UTIB0000776)')


@bp.route('/module/FIN01/proforma/<customer_type>/<int:customer_id>/<int:vcn_id>')
def proforma_invoice(customer_type, customer_id, vcn_id):
    """Printable PRO-FORMA invoice for one vessel (JJLTPL letterhead format) —
    declared/BL quantities at agreement rates, no GST (matches the manual
    pro forma issued today). Display document only: nothing persisted."""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    billables = model.get_customer_billables(customer_type, customer_id)
    vessel = next((v for v in (billables.get('vessels') or []) if v['vcn_id'] == vcn_id), None)
    if not vessel or not vessel['lines']:
        return "No billable lines found for this vessel/customer.", 404

    # ?l=SRC:cargo_id:service_type_id,... — print only the lines ticked on the
    # billables screen. Absent (a bookmarked/older link) means every line.
    picked = request.args.get('l')
    if picked:
        want = set(picked.split(','))
        vessel = {**vessel, 'lines': [
            l for l in vessel['lines']
            if f"{l['cargo_source_type']}:{l['cargo_source_id']}:{l['service_type_id']}" in want]}
        if not vessel['lines']:
            return "No lines selected for this vessel/customer.", 404

    conn = get_db()
    cur = get_cursor(conn)
    tbl = 'vessel_customers' if customer_type == 'Customer' else 'vessel_agents'
    cur.execute(f'''SELECT name, gstin, billing_address, city, pincode
                    FROM {tbl} WHERE id=%s''', [customer_id])
    cust = dict(cur.fetchone() or {})
    conn.close()

    lines, subtotal = [], 0.0
    for l in vessel['lines']:
        amount = round(float(l['qty']) * float(l['rate'] or 0), 2)
        lines.append({**l, 'amount': amount, 'amount_fmt': _inr(amount),
                      'qty_fmt': f"{float(l['qty']):.3f}".rstrip('0').rstrip('.'),
                      'rate_fmt': f"{float(l['rate'] or 0):.2f}"})
        subtotal += amount
    subtotal = round(subtotal, 2)
    sac_codes = sorted({l['sac_code'] for l in vessel['lines'] if l.get('sac_code')})

    from datetime import datetime
    now = datetime.now()
    fy = (f'{now.year % 100}-{(now.year + 1) % 100:02d}' if now.month >= 4
          else f'{(now.year - 1) % 100}-{now.year % 100:02d}')
    # ponytail: ref derives from the VCN doc num (stateless); add a numbered
    # pro forma register if finance wants sequential PI numbers
    ref_no = f"JJLTPL/PI/{fy}/{vessel['vcn_doc_num']}"

    config = get_module_config('FIN01')
    return render_template('proforma_print.html',
                           vessel=vessel, lines=lines, customer=cust,
                           ref_no=ref_no, date_str=now.strftime('%d.%m.%Y'),
                           sac_codes=', '.join(sac_codes),
                           subtotal_fmt=_inr(subtotal),
                           amount_words=_amount_in_words(subtotal),
                           seller_gstin=config.get('seller_gstin') or _PI_GSTIN,
                           seller_pan=config.get('seller_pan') or _PI_PAN,
                           payment_note=config.get('payment_note') or _PI_PAYMENT_NOTE)


@bp.route('/api/module/FIN01/bill/approve', methods=['POST'])
def approve_bill():
    """Approve a bill - only approver or admin"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    config = get_module_config('FIN01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id)
    is_admin = session.get('is_admin')

    if not is_approver and not is_admin:
        return jsonify({'success': False, 'error': 'Only approver or admin can approve bills'})

    bill_id = request.json.get('id')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE bill_header
        SET bill_status='Approved', approved_by=%s, approved_date=%s
        WHERE id=%s''',
        [session.get('username'), __import__('datetime').datetime.now().strftime('%Y-%m-%d'), bill_id])
    conn.commit()
    conn.close()

    return jsonify({'success': True})


@bp.route('/api/module/FIN01/bill/submit', methods=['POST'])
def submit_bill():
    """Submit bill for approval"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    bill_id = request.json.get('id')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE bill_header
        SET bill_status='Pending Approval'
        WHERE id=%s''', [bill_id])
    cur.execute('SELECT bill_number, customer_name, total_amount FROM bill_header WHERE id=%s', [bill_id])
    bill = cur.fetchone()
    conn.commit()
    conn.close()

    if bill:
        _queue_bill_approval_request(
            bill_id,
            bill.get('bill_number'),
            bill.get('customer_name'),
            bill.get('total_amount'),
        )

    return jsonify({'success': True})


@bp.route('/api/module/FIN01/bill/reject', methods=['POST'])
def reject_bill():
    """Reject a bill - only approver or admin"""
    if 'user_id' not in session:
        return jsonify({'success': False, 'error': 'Not logged in'})

    config = get_module_config('FIN01')
    user_id = session.get('user_id')
    is_approver = str(config.get('approver_id', '')) == str(user_id)
    is_admin = session.get('is_admin')

    if not is_approver and not is_admin:
        return jsonify({'success': False, 'error': 'Only approver or admin can reject bills'})

    bill_id = request.json.get('id')
    reason = request.json.get('reason', '')
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''UPDATE bill_header
        SET bill_status='Rejected', rejection_reason=%s
        WHERE id=%s''', [reason, bill_id])
    conn.commit()
    conn.close()

    return jsonify({'success': True})


@bp.route('/api/module/FIN01/bill-lines/<int:bill_id>')
def get_bill_lines_api(bill_id):
    """Get bill lines for a specific bill (used in invoice generation page)"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    lines = model.get_bill_lines(bill_id)
    return jsonify({'lines': lines})



@bp.route('/api/module/FIN01/service-types')
def get_service_types():
    """Get all active service types with GST rate details"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT s.id, s.service_name, s.service_code, s.sac_code, s.uom, s.gl_code,
               s.gst_rate_id,
               COALESCE(g.cgst_rate, 0) as cgst_rate,
               COALESCE(g.sgst_rate, 0) as sgst_rate,
               COALESCE(g.igst_rate, 0) as igst_rate,
               g.rate_name as gst_rate_name
        FROM finance_service_types s
        LEFT JOIN gst_rates g ON s.gst_rate_id = g.id
        WHERE s.is_active = 1
        ORDER BY s.service_name
    ''')
    rows = cur.fetchall()
    conn.close()

    return jsonify({'data': [dict(r) for r in rows]})


@bp.route('/api/module/FIN01/port-config')
def get_port_config():
    """Get port GST config (state code, GSTIN) from FIN01 module config"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    config = get_module_config('FIN01')
    return jsonify({
        'port_gst_state_code': config.get('port_gst_state_code', ''),
        'port_gstin': config.get('port_gstin', ''),
        'seller_gstin': config.get('seller_gstin', ''),
        'seller_legal_name': config.get('seller_legal_name', ''),
        'seller_address': config.get('seller_address', ''),
        'seller_location': config.get('seller_location', ''),
        'seller_pincode': config.get('seller_pincode', ''),
        'seller_phone': config.get('seller_phone', ''),
        'seller_email': config.get('seller_email', '')
    })


@bp.route('/api/module/FIN01/customer-agreements/<int:customer_id>')
def get_customer_agreements(customer_id):
    """Get all valid active approved agreements for a customer"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    from datetime import datetime
    today = datetime.now().strftime('%Y-%m-%d')

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT id, agreement_code, agreement_name, currency_code, valid_from, valid_to
        FROM customer_agreements
        WHERE customer_id = %s
        AND is_active = 1
        AND agreement_status = 'Approved'
        AND valid_from <= %s
        AND (valid_to IS NULL OR valid_to >= %s)
        ORDER BY valid_from DESC
    ''', [customer_id, today, today])
    rows = cur.fetchall()
    conn.close()

    return jsonify({'data': [dict(r) for r in rows]})


@bp.route('/api/module/FIN01/agreement-rate/<customer_type>/<int:customer_id>/<int:service_type_id>')
def get_agreement_rate(customer_type, customer_id, service_type_id):
    """Get rate from active customer/agent agreement. Optionally filter by agreement_id and cargo_name."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    from datetime import datetime

    conn = get_db()
    cur = get_cursor(conn)
    today = datetime.now().strftime('%Y-%m-%d')
    agreement_id = request.args.get('agreement_id')
    cargo_name = request.args.get('cargo_name')

    if agreement_id:
        # Try cargo-specific rate first
        if cargo_name:
            cur.execute('''
                SELECT cal.rate, cal.uom, cal.currency_code,
                       ca.agreement_code, ca.agreement_name, cal.cargo_name
                FROM customer_agreement_lines cal
                INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
                WHERE ca.id = %s AND cal.service_type_id = %s AND cal.cargo_name = %s
            ''', [agreement_id, service_type_id, cargo_name])
            row = cur.fetchone()
            if row:
                conn.close()
                return jsonify({'success': True, 'data': dict(row)})

        # Fallback to generic (no cargo) rate
        cur.execute('''
            SELECT cal.rate, cal.uom, cal.currency_code,
                   ca.agreement_code, ca.agreement_name, cal.cargo_name
            FROM customer_agreement_lines cal
            INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
            WHERE ca.id = %s AND cal.service_type_id = %s
              AND (cal.cargo_id IS NULL OR cal.cargo_name IS NULL)
        ''', [agreement_id, service_type_id])
    else:
        # Try cargo-specific rate first
        if cargo_name:
            cur.execute('''
                SELECT cal.rate, cal.uom, cal.currency_code,
                       ca.agreement_code, ca.agreement_name, cal.cargo_name
                FROM customer_agreement_lines cal
                INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
                WHERE ca.customer_type = %s
                AND ca.customer_id = %s
                AND cal.service_type_id = %s
                AND cal.cargo_name = %s
                AND ca.is_active = 1
                AND ca.agreement_status = 'Approved'
                AND ca.valid_from <= %s
                AND (ca.valid_to IS NULL OR ca.valid_to >= %s)
                ORDER BY ca.valid_from DESC
                LIMIT 1
            ''', [customer_type, customer_id, service_type_id, cargo_name, today, today])
            row = cur.fetchone()
            if row:
                conn.close()
                return jsonify({'success': True, 'data': dict(row)})

        # Fallback to generic rate
        cur.execute('''
            SELECT cal.rate, cal.uom, cal.currency_code,
                   ca.agreement_code, ca.agreement_name, cal.cargo_name
            FROM customer_agreement_lines cal
            INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
            WHERE ca.customer_type = %s
            AND ca.customer_id = %s
            AND cal.service_type_id = %s
            AND ca.is_active = 1
            AND ca.agreement_status = 'Approved'
            AND ca.valid_from <= %s
            AND (ca.valid_to IS NULL OR ca.valid_to >= %s)
            ORDER BY ca.valid_from DESC
            LIMIT 1
        ''', [customer_type, customer_id, service_type_id, today, today])
    row = cur.fetchone()
    conn.close()

    if row:
        return jsonify({'success': True, 'data': dict(row)})
    else:
        return jsonify({'success': False, 'error': 'No valid agreement found'})


@bp.route('/api/module/FIN01/cargo-rates/<customer_type>/<int:customer_id>/<int:service_type_id>')
def get_cargo_rates(customer_type, customer_id, service_type_id):
    """Get all cargo-specific rates for a service type from the agreement."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    from datetime import datetime
    conn = get_db()
    cur = get_cursor(conn)
    today = datetime.now().strftime('%Y-%m-%d')
    agreement_id = request.args.get('agreement_id')

    if agreement_id:
        cur.execute('''
            SELECT cal.rate, cal.uom, cal.currency_code, cal.cargo_id, cal.cargo_name
            FROM customer_agreement_lines cal
            INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
            WHERE ca.id = %s AND cal.service_type_id = %s AND cal.cargo_name IS NOT NULL
        ''', [agreement_id, service_type_id])
    else:
        cur.execute('''
            SELECT cal.rate, cal.uom, cal.currency_code, cal.cargo_id, cal.cargo_name
            FROM customer_agreement_lines cal
            INNER JOIN customer_agreements ca ON cal.agreement_id = ca.id
            WHERE ca.customer_type = %s
            AND ca.customer_id = %s
            AND cal.service_type_id = %s
            AND cal.cargo_name IS NOT NULL
            AND ca.is_active = 1
            AND ca.agreement_status = 'Approved'
            AND ca.valid_from <= %s
            AND (ca.valid_to IS NULL OR ca.valid_to >= %s)
            ORDER BY ca.valid_from DESC
        ''', [customer_type, customer_id, service_type_id, today, today])
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    # Build a map: cargo_name -> rate
    rate_map = {}
    for r in rows:
        if r['cargo_name'] and r['cargo_name'] not in rate_map:
            rate_map[r['cargo_name']] = r['rate']

    return jsonify({'success': True, 'rates': rate_map})


@bp.route('/api/module/FIN01/customer-billables/<customer_type>/<int:customer_id>')
def get_customer_billables(customer_type, customer_id):
    """Billables for a customer, grouped by vessel (see FIN01/model)."""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401
    return jsonify(model.get_customer_billables(customer_type, customer_id))


@bp.route('/api/module/FIN01/service-records/<customer_type>/<int:customer_id>')
def get_service_records(customer_type, customer_id):
    """Get approved, unbilled service records for a customer/agent"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    from modules.SRV01 import model as srv_model
    records = srv_model.get_unbilled_records_for_customer(customer_type, customer_id)

    conn = get_db()
    cur = get_cursor(conn)
    for rec in records:
        cur.execute('''
            SELECT sfd.field_label, srv.field_value
            FROM service_record_values srv
            JOIN service_field_definitions sfd ON srv.field_definition_id = sfd.id
            WHERE srv.service_record_id = %s
            ORDER BY sfd.display_order, sfd.id
        ''', [rec['id']])
        rec['field_values'] = [dict(r) for r in cur.fetchall()]
    conn.close()

    return jsonify({'data': records})


@bp.route('/api/module/FIN01/customers/<path:customer_type>')
def get_customers_for_billing(customer_type):
    """Get customers or agents with billing details"""
    if 'user_id' not in session:
        return jsonify({'error': 'Not logged in'}), 401

    conn = get_db()
    cur = get_cursor(conn)
    if customer_type == 'Customer':
        cur.execute('''
            SELECT id, name, gstin, gst_state_code,
                   billing_address, city, pincode, contact_phone, contact_email
            FROM vessel_customers ORDER BY name
        ''')
    elif customer_type == 'Agent':
        cur.execute('''
            SELECT id, name, gstin, gst_state_code,
                   billing_address, city, pincode, contact_phone, contact_email
            FROM vessel_agents WHERE is_active = 1 ORDER BY name
        ''')
    else:
        conn.close()
        return jsonify({'error': 'Invalid customer type'}), 400
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    # ?with_billables=1 narrows the list to parties that actually have something
    # to bill, with a per-stage count. The unfiltered list stays the default —
    # Admin Cutover needs the full master, it works on already-billed cargo.
    if request.args.get('with_billables') == '1':
        counts = model.customers_with_billables()
        rows = [dict(r, **counts[r['name']]) for r in rows if counts.get(r['name'])]
        rows.sort(key=lambda r: (-r['actual_count'], r['name'] or ''))

    return jsonify({'data': rows})
