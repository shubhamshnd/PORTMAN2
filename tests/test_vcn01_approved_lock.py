"""Approved is shut; Draft is open; billed can never be reopened.

The screen already hid these controls (isDraftRow); the server did not enforce
it — parcels on an Approved VCN were editable through the API. These assert the
backend now holds the line.
"""
import io

import pytest

from app import app as flask_app
from database import get_db, get_cursor
from modules.FIN01 import model as fin
from modules.VCN01 import model as vcn_model

_PNG = (b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
        b'\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01'
        b'\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82')


def _client(is_admin=True):
    c = flask_app.test_client()
    with c.session_transaction() as sess:
        sess['user_id'] = 1
        sess['username'] = 'pytest'
        sess['is_admin'] = is_admin
    return c


@pytest.fixture
def vcn():
    """VCN with one import parcel. Status is set per-test."""
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""INSERT INTO vcn_header (operation_type, vessel_name, doc_status, vcn_doc_num)
                   VALUES ('Import','MT LOCK TEST','Draft','VCN-TEST-LOCK') RETURNING id""")
    vcn_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_consigners (vcn_id, cargo_name, quantity, parcel_seq, parcel_no)
                   VALUES (%s,'OIL','100',1,'P1') RETURNING id""", [vcn_id])
    parcel_id = cur.fetchone()['id']
    conn.commit(); conn.close()
    yield vcn_id, parcel_id
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("DELETE FROM parcel_charge_billed WHERE cargo_source_type='VCN_IMPORT' AND cargo_source_id=%s", [parcel_id])
    cur.execute("DELETE FROM approval_log WHERE module_code='VCN01' AND record_id=%s", [vcn_id])
    cur.execute('DELETE FROM vcn_header WHERE id=%s', [vcn_id])
    conn.commit(); conn.close()


def _set_status(vcn_id, status):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('UPDATE vcn_header SET doc_status=%s WHERE id=%s', [status, vcn_id])
    conn.commit(); conn.close()


# ── Draft: do whatever you want ─────────────────────────────────────────────

def test_draft_parcels_are_editable(vcn):
    vcn_id, parcel_id = vcn
    res = _client().post('/api/module/VCN01/consigners/save',
                         json={'id': parcel_id, 'vcn_id': vcn_id,
                               'cargo_name': 'OIL', 'quantity': '150'})
    assert res.status_code == 200
    assert vcn_model.get_consigners(vcn_id)[0]['quantity'] == '150'


def test_draft_header_is_editable(vcn):
    vcn_id, _ = vcn
    res = _client().post('/api/module/VCN01/save',
                         json={'id': vcn_id, 'vessel_name': 'MT LOCK TEST 2'})
    assert res.status_code == 200


# ── Approved: shut to everyone, admin included ──────────────────────────────

def test_approved_header_is_locked_even_for_admin(vcn):
    vcn_id, _ = vcn
    _set_status(vcn_id, 'Approved')
    res = _client(is_admin=True).post('/api/module/VCN01/save',
                                      json={'id': vcn_id, 'vessel_name': 'HACKED'})
    assert res.status_code == 403
    assert 'Reopen to Draft' in res.get_json()['error']
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT vessel_name FROM vcn_header WHERE id=%s', [vcn_id])
    assert cur.fetchone()['vessel_name'] == 'MT LOCK TEST'
    conn.close()


def test_approved_parcel_save_is_locked(vcn):
    """The gap this closes: the API used to allow this while the UI hid it."""
    vcn_id, parcel_id = vcn
    _set_status(vcn_id, 'Approved')
    res = _client().post('/api/module/VCN01/consigners/save',
                         json={'id': parcel_id, 'vcn_id': vcn_id,
                               'cargo_name': 'OIL', 'quantity': '999'})
    assert res.status_code == 403
    assert vcn_model.get_consigners(vcn_id)[0]['quantity'] == '100'


def test_approved_parcel_delete_is_locked(vcn):
    vcn_id, parcel_id = vcn
    _set_status(vcn_id, 'Approved')
    res = _client().post('/api/module/VCN01/consigners/delete', json={'id': parcel_id})
    assert res.status_code == 403
    assert len(vcn_model.get_consigners(vcn_id)) == 1


def test_delays_stay_editable_after_approval(vcn):
    """Deliberate exception — delays are discovered after approval."""
    vcn_id, _ = vcn
    _set_status(vcn_id, 'Approved')
    res = _client().post('/api/module/VCN01/delays/save',
                         json={'vcn_id': vcn_id, 'delay_start': '2026-09-01T10:00',
                               'delay_end': '2026-09-01T12:00'})
    assert res.status_code == 200


# ── Billed: never reopened ──────────────────────────────────────────────────

def test_billed_vessel_cannot_be_reopened(vcn):
    vcn_id, parcel_id = vcn
    _set_status(vcn_id, 'Approved')
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT id FROM finance_service_types LIMIT 1')
    svc = cur.fetchone()
    fin.record_parcel_charge(cur, 'VCN_IMPORT', parcel_id, svc['id'], 'CHGU01',
                             None, 100, 'pytest')
    conn.commit(); conn.close()
    assert fin.is_vcn_billed(vcn_id) is True

    res = _client().post('/admin/api/reopen',
                         data={'module': 'VCN01', 'id': str(vcn_id), 'comment': 'try it',
                               'file': (io.BytesIO(_PNG), 'proof.png', 'image/png')},
                         content_type='multipart/form-data')
    assert res.status_code == 409
    assert 'never be reopened' in res.get_json()['error']
    assert vcn_model.get_doc_status(vcn_id) == 'Approved'
    assert vcn_model.get_approval_log(vcn_id) == []      # nothing written


def test_unbilled_approved_vessel_can_be_reopened(vcn):
    vcn_id, _ = vcn
    _set_status(vcn_id, 'Approved')
    res = _client().post('/admin/api/reopen',
                         data={'module': 'VCN01', 'id': str(vcn_id), 'comment': 'amendment',
                               'file': (io.BytesIO(_PNG), 'proof.png', 'image/png')},
                         content_type='multipart/form-data')
    assert res.status_code == 200
    assert vcn_model.get_doc_status(vcn_id) == 'Draft'
    # and now it is editable again
    assert _client().post('/api/module/VCN01/save',
                          json={'id': vcn_id, 'vessel_name': 'MT LOCK TEST'}).status_code == 200
