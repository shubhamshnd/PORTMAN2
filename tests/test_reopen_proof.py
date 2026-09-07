"""Reopen to Draft is admin-only and needs a photo as proof.

VCN01/LDUD01 no longer expose send-back at all (MT Hodaka Galaxy, 2026-09-05);
the single path back to Draft is Admin > Reopen to Draft, which stores the
proof image on the approval_log row it writes.
"""
import io

import pytest

from app import app as flask_app
from database import get_db, get_cursor
from modules.VCN01 import model as vcn_model

# 1x1 PNG — the endpoint reads the bytes, so a bare filename will not do.
_PNG = (b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01'
        b'\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01'
        b'\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82')


@pytest.fixture
def approved_vcn():
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""INSERT INTO vcn_header (operation_type, vessel_name, doc_status, vcn_doc_num, created_by)
                   VALUES ('Import','MT REOPEN TEST','Approved','VCN-TEST-REOPEN','pytest')
                   RETURNING id""")
    vcn_id = cur.fetchone()['id']
    conn.commit(); conn.close()
    yield vcn_id
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("DELETE FROM approval_log WHERE module_code='VCN01' AND record_id=%s", [vcn_id])
    cur.execute('DELETE FROM vcn_header WHERE id=%s', [vcn_id])
    conn.commit(); conn.close()


def _client(is_admin=True):
    c = flask_app.test_client()
    with c.session_transaction() as sess:
        sess['user_id'] = 1
        sess['username'] = 'pytest'
        sess['is_admin'] = is_admin
    return c


def _payload(vcn_id, filename='proof.png', mime='image/png', blob=_PNG, comment='amendment'):
    data = {'module': 'VCN01', 'id': str(vcn_id), 'comment': comment}
    if filename:
        data['file'] = (io.BytesIO(blob), filename, mime)
    return data


def _status(vcn_id):
    return vcn_model.get_doc_status(vcn_id)


def test_non_admin_is_refused(approved_vcn):
    res = _client(is_admin=False).post('/admin/api/reopen', data=_payload(approved_vcn),
                                       content_type='multipart/form-data')
    assert res.status_code in (302, 403)          # admin_required redirects
    assert _status(approved_vcn) == 'Approved'    # nothing changed


def test_missing_proof_is_refused(approved_vcn):
    res = _client().post('/admin/api/reopen', data=_payload(approved_vcn, filename=None),
                         content_type='multipart/form-data')
    assert res.status_code == 400
    assert 'proof image is required' in res.get_json()['error']
    assert _status(approved_vcn) == 'Approved'


def test_non_image_is_refused(approved_vcn):
    res = _client().post('/admin/api/reopen',
                         data=_payload(approved_vcn, 'scan.pdf', 'application/pdf', b'%PDF-1.4'),
                         content_type='multipart/form-data')
    assert res.status_code == 400
    assert 'PNG or JPG' in res.get_json()['error']
    assert _status(approved_vcn) == 'Approved'


def test_missing_reason_is_refused(approved_vcn):
    res = _client().post('/admin/api/reopen', data=_payload(approved_vcn, comment='  '),
                         content_type='multipart/form-data')
    assert res.status_code == 400
    assert _status(approved_vcn) == 'Approved'


def test_admin_with_proof_reopens_and_stores_the_image(approved_vcn):
    res = _client().post('/admin/api/reopen', data=_payload(approved_vcn),
                         content_type='multipart/form-data')
    assert res.status_code == 200 and res.get_json()['success'] is True
    assert _status(approved_vcn) == 'Draft'

    log = vcn_model.get_approval_log(approved_vcn)
    entry = next(e for e in log if e['action'] == 'Back to Draft')
    assert entry['has_proof'] is True
    assert entry['comment'] == 'amendment'

    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT proof_bytes, proof_filename, proof_mime FROM approval_log WHERE id=%s',
                [entry['id']])
    row = cur.fetchone(); conn.close()
    assert bytes(row['proof_bytes']) == _PNG
    assert row['proof_filename'] == 'proof.png'
    assert row['proof_mime'] == 'image/png'

    # The stored image is served back to any logged-in user, admin or not.
    res = _client(is_admin=False).get(f"/admin/api/approval-proof/{entry['id']}")
    assert res.status_code == 200
    assert res.data == _PNG


def test_approval_log_stays_json_serialisable(approved_vcn):
    """proof_bytes is BYTEA — it must never leak into an API response."""
    _client().post('/admin/api/reopen', data=_payload(approved_vcn),
                   content_type='multipart/form-data')
    res = _client().get(f'/api/module/VCN01/approval-log/{approved_vcn}')
    assert res.status_code == 200
    assert all('proof_bytes' not in e for e in res.get_json())


def test_draft_record_cannot_be_reopened(approved_vcn):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("UPDATE vcn_header SET doc_status='Draft' WHERE id=%s", [approved_vcn])
    conn.commit(); conn.close()
    res = _client().post('/admin/api/reopen', data=_payload(approved_vcn),
                         content_type='multipart/form-data')
    assert res.status_code == 400
    assert 'Only Approved' in res.get_json()['error']
