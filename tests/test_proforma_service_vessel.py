"""A service record raised against a VCN shows that vessel on the pro forma.

The heading came from `ref_source_display.split('/')[0]` — but that string is
the SRV01 dropdown label, "VCN-2627-001 / SC GARNET / 2026-01-15 10:30", so
splitting it kept the VCN number and threw the vessel name away. Section A
(cargo) has always shown the vessel, so a services pro forma looked different
for no reason.

The vessel now comes off vcn_header via the join get_unbilled_services already
had — it was selecting only via_number from it.

Hits the dev DB; every row it writes is removed again.
"""
import pytest

from database import get_db, get_cursor
from modules.FIN01 import model, views


@pytest.fixture
def fixtures():
    """A customer, a service type, two VCNs, and three service records:
    one against vessel A, one against vessel B, one against nothing."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id FROM finance_service_types ORDER BY id LIMIT 1')
    svc = cur.fetchone()
    cur.execute('SELECT id FROM vessel_customers ORDER BY id LIMIT 1')
    cust = cur.fetchone()
    if not svc or not cust:
        conn.close()
        pytest.skip('no service type / customer configured')

    made_vcns, made_records = [], []
    for num, name in (('PYTEST-VCN-A', 'PYTEST VESSEL A'), ('PYTEST-VCN-B', 'PYTEST VESSEL B')):
        cur.execute("""INSERT INTO vcn_header (vcn_doc_num, vessel_name, doc_status)
                       VALUES (%s, %s, 'Approved') RETURNING id""", [num, name])
        made_vcns.append({'id': cur.fetchone()['id'], 'vcn_doc_num': num, 'vessel_name': name})

    def add(record_number, vcn):
        cur.execute("""INSERT INTO service_records
            (record_number, module_code, service_type_id, source_type, source_id,
             source_display, ref_source_type, ref_source_id, ref_source_display,
             record_date, billable_quantity, billable_uom, doc_status, is_billed)
            VALUES (%s, 'SRV02', %s, 'Customer', %s, 'pytest', %s, %s, %s,
                    CURRENT_DATE::text, 5, 'DAY', 'Approved', 0) RETURNING id""",
            [record_number, svc['id'], cust['id'],
             'VCN' if vcn else None, vcn['id'] if vcn else None,
             f"{vcn['vcn_doc_num']} / {vcn['vessel_name']} / 2026-01-01 00:00" if vcn else None])
        made_records.append(cur.fetchone()['id'])

    add('PYTEST-SRV-A', made_vcns[0])
    add('PYTEST-SRV-B', made_vcns[1])
    add('PYTEST-SRV-N', None)
    conn.commit()
    conn.close()

    yield cust['id'], made_records, made_vcns

    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('DELETE FROM service_records WHERE id = ANY(%s)', [made_records])
    cur.execute('DELETE FROM vcn_header WHERE id = ANY(%s)', [[v['id'] for v in made_vcns]])
    conn.commit()
    conn.close()


def _ctx(customer_id, picked):
    ctx, err, _ = views._services_ctx('Customer', customer_id,
                                      ','.join(str(x) for x in picked))
    assert not err, err
    return ctx


def test_the_line_carries_the_vessel_off_vcn_header(fixtures):
    """get_unbilled_services joins vcn_header but used to select only
    via_number, so the vessel was never available to the document."""
    customer_id, records, vcns = fixtures
    rows = model.get_unbilled_services('Customer', customer_id)
    row = next(r for r in rows if r['service_record_id'] == records[0])
    assert row['vessel_name'] == vcns[0]['vessel_name']
    assert row['vcn_doc_num'] == vcns[0]['vcn_doc_num']
    assert row['ref_source_id'] == vcns[0]['id']


def test_one_vessel_becomes_the_heading(fixtures):
    customer_id, records, vcns = fixtures
    assert _ctx(customer_id, [records[0]])['vessel_name'] == vcns[0]['vessel_name']


def test_two_vessels_put_each_on_its_own_line(fixtures):
    """No single vessel can head the document, so the vessel moves onto the
    lines — it is the only thing telling them apart."""
    customer_id, records, vcns = fixtures
    ctx = _ctx(customer_id, records[:2])
    assert ctx['vessel_name'] == 'Other Services'
    labels = ' | '.join(r['label'] for r in ctx['rows'])
    assert vcns[0]['vessel_name'] in labels
    assert vcns[1]['vessel_name'] in labels


def test_a_record_with_no_vessel_is_not_printed_under_one(fixtures):
    """The regression this guards: an unreferenced record mixed in with a
    referenced one was filtered out of the VCN set, so the document headed
    itself with a vessel that record was never raised against."""
    customer_id, records, vcns = fixtures
    ctx = _ctx(customer_id, [records[0], records[2]])
    assert ctx['vessel_name'] == 'Other Services', \
        'a service with no VCN was attributed to a vessel'
    labels = ' | '.join(r['label'] for r in ctx['rows'])
    assert 'PYTEST-SRV-N' in labels, 'the unreferenced record must identify itself'


def test_no_vessel_anywhere_does_not_repeat_the_heading(fixtures):
    """With nothing to break down, the document stays flat — a detail row
    under the heading would just say the same thing twice."""
    customer_id, records, _ = fixtures
    ctx = _ctx(customer_id, [records[2]])
    assert ctx['vessel_name'] == 'Other Services'
    assert len(ctx['rows']) == 1, [r['label'] for r in ctx['rows']]
