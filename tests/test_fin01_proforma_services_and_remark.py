"""Section B gets the same pro-forma as section A, and both carry a remark
that defaults to the VCN's VIA No.

Dev DB with a throwaway customer + VCN + service record, cleaned up.
"""
import re
import zlib
from pathlib import Path

from database import get_db, get_cursor
from modules.FIN01 import model, proforma_pdf, views

GEN_BILL = Path('modules/FIN01/generate_bill.html').read_text(encoding='utf-8')
NAME = 'ZZ PF SERVICES CO'
VIA = 'ZZVIA9'
RATE = 1200.0


def _setup(cur, with_vcn_ref=True):
    cur.execute('SELECT id FROM finance_service_types WHERE is_active=1 ORDER BY id LIMIT 1')
    svc = cur.fetchone()['id']
    cur.execute('INSERT INTO vessel_customers (name) VALUES (%s) RETURNING id', [NAME])
    cid = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_header (operation_type, vcn_doc_num, vessel_name, via_number)
                   VALUES ('Import','VCN-PF-1','PFVESSEL',%s) RETURNING id""", [VIA])
    vcn = cur.fetchone()['id']
    ref = ('VCN', vcn, 'VCN-PF-1 / PFVESSEL') if with_vcn_ref else (None, None, None)
    cur.execute("""INSERT INTO service_records
        (module_code, record_number, service_type_id, source_type, source_id, source_display,
         ref_source_type, ref_source_id, ref_source_display,
         record_date, billable_quantity, billable_uom, doc_status, is_billed)
        VALUES ('SRV02','ZZPFS1',%s,'Customer',%s,%s,%s,%s,%s,'2026-09-18',2,'OTH','Approved',0)
        RETURNING id""", [svc, cid, NAME, ref[0], ref[1], ref[2]])
    rec = cur.fetchone()['id']
    cur.execute("""INSERT INTO customer_agreements (customer_type, customer_id, agreement_code,
        agreement_name, agreement_status, is_active, valid_from, valid_to)
        VALUES ('Customer', %s, 'ZZPFSAG', 'pf', 'Approved', 1, '2000-01-01', NULL)
        RETURNING id""", [cid])
    ag = cur.fetchone()['id']
    cur.execute("""INSERT INTO customer_agreement_lines
        (agreement_id, service_type_id, cargo_name, rate, uom, currency_code)
        VALUES (%s, %s, NULL, %s, 'OTH', 'INR')""", [ag, svc, RATE])
    return cid, vcn, rec, ag


def _teardown(cid, vcn, rec, ag):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('DELETE FROM service_records WHERE id=%s', [rec])
    cur.execute('DELETE FROM customer_agreement_lines WHERE agreement_id=%s', [ag])
    cur.execute('DELETE FROM customer_agreements WHERE id=%s', [ag])
    cur.execute('DELETE FROM vcn_header WHERE id=%s', [vcn])
    cur.execute('DELETE FROM vessel_customers WHERE id=%s', [cid])
    conn.commit(); conn.close()


def _drawn(pdf_bytes):
    raw = ''
    for m in re.finditer(rb'stream\r?\n(.*?)\r?\nendstream', pdf_bytes, re.S):
        try:
            raw += zlib.decompress(m.group(1)).decode('latin-1')
        except Exception:
            pass
    return re.findall(r'\((.*?)\)\s*Tj', raw)


# ── Section B pro-forma ─────────────────────────────────────────────────────

def test_services_get_the_same_document_as_cargo():
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    conn.commit(); conn.close()
    try:
        ctx, err, _ = views._services_ctx('Customer', cid, str(rec),
                                          'JJLTPL/PI/', '0485', 'VIA No: ' + VIA)
        assert err is None, err
        # every key the PDF renderer reads, same as section A
        for k in ('vessel_name', 'customer', 'rows', 'ref_no', 'subtotal',
                  'tax_rows', 'total', 'amount_words', 'remark', 'seller_gstin'):
            assert k in ctx, k
        assert ctx['ref_no'] == 'JJLTPL/PI/0485'
        assert ctx['subtotal'] == 2400.0            # 2 @ 1200
        assert ctx['total'] == round(2400.0 * 1.18, 2)
        assert proforma_pdf.render(ctx).startswith(b'%PDF')
    finally:
        _teardown(cid, vcn, rec, ag)


def test_only_the_ticked_records_print():
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    conn.commit(); conn.close()
    try:
        _ctx, err, status = views._services_ctx('Customer', cid, str(rec + 99999))
        assert err and status == 404
    finally:
        _teardown(cid, vcn, rec, ag)


def test_one_vcn_is_named_once_in_the_heading_not_on_every_line():
    """A service has no cargo to break down, so it is one flat row carrying its
    own figures — repeating the reference under a heading that already says it
    adds nothing.

    The heading is the VESSEL, matching section A. It used to be the VCN
    number, because it was parsed out of ref_source_display ("VCN / vessel /
    anchored") by splitting on '/' and keeping the first field.
    """
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    conn.commit(); conn.close()
    try:
        ctx, _, _ = views._services_ctx('Customer', cid, str(rec))
        assert ctx['vessel_name'] == 'PFVESSEL'
        assert len(ctx['rows']) == 1
        row = ctx['rows'][0]
        assert 'VCN-PF-1' not in row['label'], row['label']
        assert 'PFVESSEL' not in row['label'], row['label']
        assert row['indent'] is False
        # the figures are on the row itself, not on a detail row underneath
        assert row['qty'] == 2.0 and row['rate'] == RATE and row['amount'] == 2400.0
    finally:
        _teardown(cid, vcn, rec, ag)


def test_several_vcns_keep_their_reference_on_the_line():
    """With more than one call on the document the heading cannot name them,
    so the reference is the only thing telling the lines apart."""
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    cur.execute("""INSERT INTO vcn_header (operation_type, vcn_doc_num, vessel_name, via_number)
                   VALUES ('Import','VCN-PF-2','PFVESSEL2','ZZVIA8') RETURNING id""")
    vcn2 = cur.fetchone()['id']
    cur.execute('SELECT service_type_id FROM service_records WHERE id=%s', [rec])
    svc = cur.fetchone()['service_type_id']
    cur.execute("""INSERT INTO service_records
        (module_code, record_number, service_type_id, source_type, source_id, source_display,
         ref_source_type, ref_source_id, ref_source_display,
         record_date, billable_quantity, billable_uom, doc_status, is_billed)
        VALUES ('SRV02','ZZPFS2',%s,'Customer',%s,%s,'VCN',%s,'VCN-PF-2 / PFVESSEL2',
                '2026-09-18',1,'OTH','Approved',0) RETURNING id""", [svc, cid, NAME, vcn2])
    rec2 = cur.fetchone()['id']
    conn.commit(); conn.close()
    try:
        ctx, _, _ = views._services_ctx('Customer', cid, f'{rec},{rec2}')
        assert ctx['vessel_name'] == 'Other Services'
        labels = ' '.join(r['label'] for r in ctx['rows'])
        assert 'VCN-PF-1' in labels and 'VCN-PF-2' in labels, labels
    finally:
        conn = get_db(); cur = get_cursor(conn)
        cur.execute('DELETE FROM service_records WHERE id=%s', [rec2])
        cur.execute('DELETE FROM vcn_header WHERE id=%s', [vcn2])
        conn.commit(); conn.close()
        _teardown(cid, vcn, rec, ag)


def test_flat_lines_club_by_service_and_rate():
    """Same rule group_lines applies: one service at two rates stays two rows,
    and the quantities add up rather than being averaged."""
    from modules.FIN01.proforma_pdf import flat_lines
    rows = flat_lines([
        {'service_name': 'Gangway', 'qty': 2, 'rate': 100, 'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
        {'service_name': 'Gangway', 'qty': 3, 'rate': 100, 'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
        {'service_name': 'Gangway', 'qty': 1, 'rate': 250, 'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
    ])
    assert [(r['label'], r['qty'], r['rate'], r['amount']) for r in rows] == [
        ('Gangway', 5.0, 100.0, 500.0),
        ('Gangway', 1.0, 250.0, 250.0),
    ]
    # and they sum to the subtotal without a heading double counting them
    assert sum(r['amount'] for r in rows) == 750.0
    assert all(r['indent'] is False for r in rows)


def test_records_with_no_vcn_reference_still_print():
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur, with_vcn_ref=False)
    conn.commit(); conn.close()
    try:
        ctx, err, _ = views._services_ctx('Customer', cid, str(rec))
        assert err is None
        assert ctx['vessel_name'] == 'Other Services'
        assert proforma_pdf.render(ctx).startswith(b'%PDF')
    finally:
        _teardown(cid, vcn, rec, ag)


# ── Remark ──────────────────────────────────────────────────────────────────

def test_the_remark_prints_under_the_heading():
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    conn.commit(); conn.close()
    try:
        ctx, _, _ = views._services_ctx('Customer', cid, str(rec), remark='VIA No: ' + VIA)
        drawn = _drawn(proforma_pdf.render(ctx))
        i = drawn.index('PFVESSEL')      # heading is the vessel, not the VCN
        assert drawn[i + 1] == 'VIA No: ' + VIA
    finally:
        _teardown(cid, vcn, rec, ag)


def test_an_empty_remark_prints_no_row():
    """The operator can clear the pre-filled VIA No — that must leave a blank
    line behind, not an empty row in the particulars column."""
    for remark in (None, '', '   '):
        ctx = dict(_MIN_CTX, remark=remark)
        drawn = _drawn(proforma_pdf.render(ctx))
        i = drawn.index('HEADERVESSEL')
        assert drawn[i + 1] != '', drawn[i:i + 3]
        assert not any('VIA' in t for t in drawn)


def test_the_via_number_reaches_the_screen():
    """The dialog pre-fills from data already on the page, so both sections
    need the VIA No on their rows."""
    conn = get_db(); cur = get_cursor(conn)
    cid, vcn, rec, ag = _setup(cur)
    conn.commit(); conn.close()
    try:
        s = model.get_unbilled_services('Customer', cid)[0]
        assert s['via_number'] == VIA
    finally:
        _teardown(cid, vcn, rec, ag)


_MIN_CTX = {
    'vessel_name': 'HEADERVESSEL', 'ref_no': 'ZZ/1', 'date_str': '18.09.2026',
    'customer': {'name': 'C'}, 'ac_names': [],
    'rows': [{'label': 'Svc', 'indent': False, 'qty': None, 'rate': None,
              'amount': None, 'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18}],
    'sac_codes': '996719', 'subtotal': 0.0, 'tax_rows': [], 'total': 0.0,
    'amount_words': 'Nil.', 'escalation_note': '', 'seller_gstin': 'g',
    'seller_pan': 'p', 'payment_note': 'n',
}


# ── Screen wiring ───────────────────────────────────────────────────────────

def test_the_dialog_asks_for_a_remark_prefilled_with_the_via_no():
    assert 'id="pfRemark"' in GEN_BILL
    assert 'VIA No: ${viaNumber}' in GEN_BILL
    assert 'rm=' in GEN_BILL


def test_section_b_has_its_own_pro_forma_button():
    assert 'openServiceProforma' in GEN_BILL
    # both sections drive the one dialog, so they cannot drift apart
    assert GEN_BILL.count('showPfDialog(') == 3   # definition + two callers


def test_both_sections_reach_the_same_resolver():
    """Preview, PDF and mail send must all pick the document the same way."""
    src = Path('modules/FIN01/views.py').read_text(encoding='utf-8')
    for fn in ('def proforma_invoice(', 'def proforma_invoice_pdf(', 'def send_proforma('):
        body = src[src.index(fn):]
        body = body[:body.index('\n@bp.route') if '\n@bp.route' in body else len(body)]
        assert '_ctx_for_request(' in body, fn
    # and the query string carries every selector through to the PDF and mail
    qs = src[src.index('def _proforma_qs('):]
    qs = qs[:qs.index('\n@bp.route')]
    for key in ("'l'", "'r'", "'s'", "'n'", "'rm'"):
        assert key in qs, key
