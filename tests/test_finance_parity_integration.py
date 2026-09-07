"""Integration checks for the finance-parity work, against the dev DB with
throwaway rows that are always cleaned up:

  * LDUD01 billed lock (409 on every write; reopen lives in ADMIN and is
    refused outright for a billed vessel)
  * billing customer picker counts (?with_billables=1)
  * Admin Cutover flagging, unmarking and the lock
  * seeded bill numbering
  * FSAP01 console endpoints

No SAP or IRP network calls.
"""
from flask import Flask, session

from database import get_db, get_cursor
from modules.ADMIN import cutover
from modules.FIN01 import model as fin
from modules.FSAP01 import views as fsap
from modules.LDUD01 import views as ldud
from modules.ADMIN import views as admin_views

_app = Flask(__name__)
_app.secret_key = 'pytest-secret'


# ── fixtures-by-hand ─────────────────────────────────────────────────────────

def _mk_vessel(customer_name, ldud_status=None, vcn_status='Approved', qty='100'):
    """VCN + one import parcel payable by `customer_name`, optionally with an
    LDUD in `ldud_status`. Returns (vcn_id, parcel_id, ldud_id)."""
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""INSERT INTO vcn_header (operation_type, vessel_name, doc_status, vcn_doc_num)
                   VALUES ('Import','PARITY TEST',%s,'VCN-PARITY') RETURNING id""", [vcn_status])
    vcn_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_consigners
                   (vcn_id, cargo_name, quantity, consigner_name, importer_name,
                    pipeline_name, unload_terminal, parcel_seq, parcel_no)
                   VALUES (%s,'OIL',%s,'CNS',%s,'PL1','T1',1,'P1') RETURNING id""",
                [vcn_id, qty, customer_name])
    parcel_id = cur.fetchone()['id']
    ldud_id = None
    if ldud_status:
        cur.execute("""INSERT INTO ldud_header (vcn_id, vessel_name, doc_status, operation_type)
                       VALUES (%s,'PARITY TEST',%s,'Import') RETURNING id""", [vcn_id, ldud_status])
        ldud_id = cur.fetchone()['id']
    conn.commit(); conn.close()
    return vcn_id, parcel_id, ldud_id


def _drop_vessel(vcn_id, parcel_id):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("DELETE FROM parcel_charge_billed WHERE cargo_source_type='VCN_IMPORT' AND cargo_source_id=%s",
                [parcel_id])
    cur.execute("DELETE FROM approval_log WHERE module_code='LDUD01' AND record_id IN "
                "(SELECT id FROM ldud_header WHERE vcn_id=%s)", [vcn_id])
    cur.execute("DELETE FROM ldud_parcel_ops WHERE ldud_id IN (SELECT id FROM ldud_header WHERE vcn_id=%s)",
                [vcn_id])
    cur.execute("DELETE FROM ldud_header WHERE vcn_id=%s", [vcn_id])
    cur.execute("DELETE FROM vcn_header WHERE id=%s", [vcn_id])  # cascades consigners
    conn.commit(); conn.close()


def _service_id(code='CHGU01'):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT id FROM finance_service_types WHERE service_code=%s LIMIT 1', [code])
    row = cur.fetchone(); conn.close()
    return row['id'] if row else None


def _bill_parcel(parcel_id, qty=100, bill_id=None, service_code='CHGU01'):
    conn = get_db(); cur = get_cursor(conn)
    fin.record_parcel_charge(cur, 'VCN_IMPORT', parcel_id, _service_id(service_code),
                             service_code, bill_id, qty, 'pytest')
    conn.commit(); conn.close()


def _mk_customer(name):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('INSERT INTO vessel_customers (name) VALUES (%s) RETURNING id', [name])
    cid = cur.fetchone()['id']
    conn.commit(); conn.close()
    return cid


def _drop_customer(cid):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('DELETE FROM vessel_customers WHERE id=%s', [cid])
    conn.commit(); conn.close()


# ── D. LDUD01 billed lock ────────────────────────────────────────────────────

def _post(view, payload, admin=False, user_id=1):
    with _app.test_request_context('/', json=payload):
        session['user_id'] = user_id
        session['username'] = 'pytest'
        if admin:
            session['is_admin'] = True
        return view()


# 1x1 PNG - the reopen endpoint requires a real image, not just a filename.
_PNG = (b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
        b'\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01'
        b'\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82')


def _post_reopen(module, record_id, comment, admin=True, with_proof=True, user_id=1):
    """Reopen through ADMIN - multipart, admin-only, proof image mandatory."""
    import io as _io
    data = {'module': module, 'id': str(record_id), 'comment': comment}
    if with_proof:
        data['file'] = (_io.BytesIO(_PNG), 'proof.png', 'image/png')
    with _app.test_request_context('/', method='POST', data=data,
                                   content_type='multipart/form-data'):
        session['user_id'] = user_id
        session['username'] = 'pytest'
        if admin:
            session['is_admin'] = True
        res = admin_views.reopen_record()
        return res if isinstance(res, tuple) else (res, 200)


def test_billed_vessel_locks_every_ldud_write(monkeypatch):
    vcn_id, parcel_id, ldud_id = _mk_vessel('PARITY LOCK CO', ldud_status='Closed')
    try:
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""INSERT INTO bill_header (bill_number, bill_date, source_type, customer_type, customer_id, bill_status)
                       VALUES ('BILL-PARITY-1','2026-09-01','MULTI','Customer',0,'Approved') RETURNING id""")
        bill_id = cur.fetchone()['id']
        cur.execute("""INSERT INTO ldud_parcel_ops (ldud_id, parcel_ids, quantity)
                       VALUES (%s, %s, 50) RETURNING id""", [ldud_id, str(parcel_id)])
        op_id = cur.fetchone()['id']
        conn.commit(); conn.close()
        _bill_parcel(parcel_id, bill_id=bill_id)

        # Reopen no longer exists in LDUD01 - it moved to ADMIN, admin-only.
        assert not hasattr(ldud, 'reopen')

        # Admin does not get a bypass on the other four write paths.
        for view, payload in (
            (ldud.save, {'id': ldud_id, 'vcn_id': vcn_id}),
            (ldud.save_parcel_op, {'ldud_id': ldud_id, 'parcel_ids': str(parcel_id), 'quantity': 10}),
            (ldud.delete_parcel_op, {'id': op_id}),
            (ldud.delete, {'id': ldud_id}),
        ):
            body, status = _post(view, payload, admin=True)
            assert status == 409, (view.__name__, status)

        # A reopen without a proof image is refused outright.
        body, status = _post_reopen('LDUD01', ldud_id, 'no proof', with_proof=False)
        assert status == 400 and 'proof image is required' in body.get_json()['error']

        # A billed vessel can NEVER be reopened — there is no admin override.
        body, status = _post_reopen('LDUD01', ldud_id, 'legacy correction')
        assert status == 409
        assert 'never be reopened' in body.get_json()['error']

        # Refused means refused: status untouched, nothing written to the log,
        # and the proof documents are still there (the delete sits after the
        # billed check, so a refusal must not destroy them).
        conn = get_db(); cur = get_cursor(conn)
        cur.execute('SELECT doc_status FROM ldud_header WHERE id=%s', [ldud_id])
        assert cur.fetchone()['doc_status'] == 'Closed'
        conn.close()
        assert 'Back to Draft' not in [r['action'] for r in ldud.model.get_closure_log(ldud_id)]
        assert fin.is_vcn_billed(vcn_id) is True
    finally:
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("DELETE FROM bill_header WHERE bill_number='BILL-PARITY-1'")
        conn.commit(); conn.close()
        _drop_vessel(vcn_id, parcel_id)


def test_unbilled_vessel_is_not_locked():
    vcn_id, parcel_id, ldud_id = _mk_vessel('PARITY OPEN CO', ldud_status='Closed')
    try:
        assert ldud._billed_locked(ldud_id) is None
    finally:
        _drop_vessel(vcn_id, parcel_id)


# ── E. Billing customer picker ───────────────────────────────────────────────

def test_picker_counts_by_stage_and_drops_fully_billed():
    name = 'PARITY PICKER CO'
    cid = _mk_customer(name)
    vcn_id, parcel_id, _ = _mk_vessel(name, ldud_status=None)
    try:
        counts = fin.customers_with_billables()
        assert counts[name]['proforma_count'] == 1
        assert counts[name]['actual_count'] == 0

        # Close the LDUD: same parcel, now an actual-stage billable.
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""INSERT INTO ldud_header (vcn_id, vessel_name, doc_status, operation_type)
                       VALUES (%s,'PARITY TEST','Closed','Import')""", [vcn_id])
        conn.commit(); conn.close()

        counts = fin.customers_with_billables()
        assert counts[name]['actual_count'] == 1
        assert counts[name]['proforma_count'] == 0

        # Bill every applicable charge in full — the party drops out.
        for code in fin.parcel_charge_codes('VCN_IMPORT', None, None):
            _bill_parcel(parcel_id, qty=100, bill_id=None, service_code=code)
        assert name not in fin.customers_with_billables()
    finally:
        _drop_vessel(vcn_id, parcel_id)
        _drop_customer(cid)


def test_picker_ignores_parties_with_no_parcels():
    name = 'PARITY EMPTY CO'
    cid = _mk_customer(name)
    try:
        assert name not in fin.customers_with_billables()
    finally:
        _drop_customer(cid)


# ── F. Admin Cutover ─────────────────────────────────────────────────────────

def test_cutover_flag_locks_the_vessel_and_unmark_spares_real_bills():
    vcn_id, parcel_id, _ = _mk_vessel('PARITY CUTOVER CO')
    try:
        cutover.mark_items_billed(
            [{'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': parcel_id}], 'pytest')

        # A cutover flag is a ledger row with no bill behind it — and it locks.
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""SELECT COUNT(*) AS n FROM parcel_charge_billed
                       WHERE cargo_source_id=%s AND bill_id IS NULL""", [parcel_id])
        assert cur.fetchone()['n'] > 0
        conn.close()
        assert fin.is_vcn_billed(vcn_id) is True

        # Add a genuine bill's row, then unmark: only the NULL rows go.
        _bill_parcel(parcel_id, qty=10, bill_id=424242, service_code='INFM01')
        cutover.unmark_items_billed(
            [{'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': parcel_id}], 'pytest')

        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""SELECT bill_id FROM parcel_charge_billed WHERE cargo_source_id=%s""", [parcel_id])
        remaining = [r['bill_id'] for r in cur.fetchall()]
        conn.close()
        assert remaining == [424242]
    finally:
        _drop_vessel(vcn_id, parcel_id)


def test_cutover_lock_refuses_every_write():
    vcn_id, parcel_id, _ = _mk_vessel('PARITY LOCKED CO')
    cutover.set_lock(True, 'pytest')
    try:
        for fn, args in (
            (cutover.mark_items_billed,
             ([{'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': parcel_id}], 'pytest')),
            (cutover.unmark_items_billed,
             ([{'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': parcel_id}], 'pytest')),
            (cutover.set_bill_seed, (999999, 'pytest')),
        ):
            try:
                fn(*args)
                assert False, f'{fn.__name__} should refuse while locked'
            except cutover.CutoverLocked:
                pass
    finally:
        cutover.set_lock(False, 'pytest')
        _drop_vessel(vcn_id, parcel_id)
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("DELETE FROM cutover_audit WHERE performed_by='pytest'")
        conn.commit(); conn.close()


def test_bill_seed_is_a_floor_then_normal_incrementing_resumes():
    start = cutover.current_max('bill') + 500
    try:
        cutover.set_bill_seed(start, 'pytest')
        assert fin.get_next_bill_number() == f'BILL{start:04d}'

        # Once a real document passes the seed, incrementing takes over.
        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""INSERT INTO bill_header (bill_number, bill_date, source_type, customer_type, customer_id, bill_status)
                       VALUES (%s,'2026-09-01','MULTI','Customer',0,'Draft')""", [f'BILL{start + 3:04d}'])
        conn.commit(); conn.close()
        assert fin.get_next_bill_number() == f'BILL{start + 4:04d}'
    finally:
        conn = get_db(); cur = get_cursor(conn)
        cur.execute('DELETE FROM bill_header WHERE bill_number=%s', [f'BILL{start + 3:04d}'])
        cur.execute("DELETE FROM cutover_seed WHERE seed_type='bill'")
        cur.execute("DELETE FROM cutover_audit WHERE performed_by='pytest'")
        conn.commit(); conn.close()


def test_seed_must_be_above_the_current_maximum():
    try:
        cutover.set_bill_seed(cutover.current_max('bill'), 'pytest')
        assert False, 'a seed at the current maximum must be rejected'
    except ValueError as e:
        assert 'current maximum' in str(e)


# ── B. FSAP01 console endpoints ──────────────────────────────────────────────

def _get(view, query='', user_id=1, admin=True, **kwargs):
    with _app.test_request_context('/?' + query):
        session['user_id'] = user_id
        if admin:
            session['is_admin'] = True
        return view(**kwargs)


def test_fsap01_console_endpoints_return_a_paginated_envelope():
    for view in (fsap.callback_logs, fsap.outbound_logs, fsap.sap_queue_list):
        body = _get(view, 'page=1&size=5').get_json()
        assert set(body) == {'data', 'last_page', 'total'}, view.__name__
        assert isinstance(body['data'], list)
        assert len(body['data']) <= 5


def test_fsap01_manual_send_refuses_without_can_edit():
    body, status = _post(fsap.sap_queue_manual_send, {'queue_id': 1}, admin=False, user_id=-1)
    assert status == 403
    assert body.get_json()['ok'] is False


# ── Batched rate lookup must match the per-call one ──────────────────────────

def test_rates_map_matches_get_customer_rate():
    """get_customer_billables prices every line from get_customer_rates_map
    instead of one get_customer_rate call (= one DB connection) per charge.
    The two must resolve identically, cargo-specific line and fallback alike."""
    from modules.FCAM01 import model as fcam
    name = 'PARITY RATE CO'
    cid = _mk_customer(name)
    svc_a, svc_b = _service_id('CHGU01'), _service_id('INFM01')
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""INSERT INTO customer_agreements
                   (agreement_code, customer_type, customer_id, customer_name,
                    valid_from, is_active, agreement_status)
                   VALUES ('AGR-PARITY','Customer',%s,%s,'2000-01-01',1,'Approved')
                   RETURNING id""", [cid, name])
    agr = cur.fetchone()['id']
    cur.execute("""INSERT INTO customer_agreement_lines (agreement_id, service_type_id, rate, cargo_name)
                   VALUES (%s,%s,42.5,'OIL'), (%s,%s,10.0,NULL), (%s,%s,7.25,'COAL')""",
                [agr, svc_a, agr, svc_a, agr, svc_b])
    conn.commit(); conn.close()
    try:
        by_cargo, by_service = fcam.get_customer_rates_map('Customer', cid)
        for service_type_id, cargo in ((svc_a, 'OIL'), (svc_a, 'COAL'), (svc_a, None),
                                       (svc_b, 'COAL'), (svc_b, 'OIL'), (svc_b, None)):
            old = fcam.get_customer_rate('Customer', cid, service_type_id, cargo_name=cargo)
            new = ((by_cargo.get((service_type_id, cargo)) if cargo else None)
                   or by_service.get(service_type_id))
            new = {k: new[k] for k in old} if (new and old) else new
            assert new == old, (service_type_id, cargo, new, old)
    finally:
        conn = get_db(); cur = get_cursor(conn)
        cur.execute('DELETE FROM customer_agreement_lines WHERE agreement_id=%s', [agr])
        cur.execute('DELETE FROM customer_agreements WHERE id=%s', [agr])
        conn.commit(); conn.close()
        _drop_customer(cid)
