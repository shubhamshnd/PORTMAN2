"""A removed parcel stops counting everywhere, but is never deleted.

The flag lives on the VCN parcel row, not ldud_parcel_ops, because FIN01
selects billable parcels straight from the VCN tables — flagging the op would
have left the parcel billable at its full DECLARED quantity.
"""
import pytest

from database import get_db, get_cursor
from modules.FIN01 import model as fin
from modules.LDUD01 import model as ldud_model
from modules.LUEU01 import model as lueu_model
from modules.VCN01 import model as vcn_model


@pytest.fixture
def vessel():
    """VCN with two 100 MT import parcels, an LDUD and one op per parcel."""
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("INSERT INTO vessel_customers (name) VALUES ('RM CUSTOMER') RETURNING id")
    cust_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_header (operation_type, vessel_name, vcn_doc_num,
                                           doc_status, nor_tendered, via_number,
                                           berth_name, vessel_run_type, discharge_port,
                                           vessel_agent_name)
                   VALUES ('Import','MT REMOVE TEST','VCN-TEST-RM','Draft',
                           '2026-09-01T08:00','V1','B1','Foreign','JNPA','AG')
                   RETURNING id""")
    vcn_id = cur.fetchone()['id']
    pids = []
    for seq, cargo in ((1, 'OIL'), (2, 'ACID')):
        cur.execute("""INSERT INTO vcn_consigners
                       (vcn_id, cargo_name, quantity, consigner_name, importer_name,
                        pipeline_name, unload_terminal, parcel_seq, parcel_no)
                       VALUES (%s,%s,'100','CNS','RM CUSTOMER','PL1','T1',%s,%s)
                       RETURNING id""", [vcn_id, cargo, seq, f'P{seq}'])
        pids.append(cur.fetchone()['id'])
    cur.execute("""INSERT INTO ldud_header (vcn_id, vessel_name, nor_tendered, doc_status)
                   VALUES (%s,'MT REMOVE TEST','2026-09-01T08:00','Closed') RETURNING id""",
                [vcn_id])
    ldud_id = cur.fetchone()['id']
    ops = []
    for pid in pids:
        cur.execute("""INSERT INTO ldud_parcel_ops (ldud_id, parcel_ids, cargo_name, quantity)
                       VALUES (%s,%s,'X','100') RETURNING id""", [ldud_id, str(pid)])
        op_id = cur.fetchone()['id']
        ops.append(op_id)
        # LDUD is Closed, so billing uses the ACTUAL logged quantity
        cur.execute("""INSERT INTO lueu_parcel_log
                       (parcel_op_id, entry_date, from_time, to_time, quantity, quantity_uom)
                       VALUES (%s,'2026-09-01','08:00','18:00',100,'MT')""", [op_id])
    conn.commit(); conn.close()
    yield vcn_id, pids, ldud_id, ops, cust_id
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("DELETE FROM parcel_charge_billed WHERE cargo_source_type='VCN_IMPORT' AND cargo_source_id = ANY(%s)", [pids])
    cur.execute("""DELETE FROM lueu_parcel_log WHERE parcel_op_id IN
                   (SELECT po.id FROM ldud_parcel_ops po JOIN ldud_header l ON l.id=po.ldud_id
                    WHERE l.vcn_id=%s)""", [vcn_id])
    cur.execute("DELETE FROM ldud_header WHERE vcn_id=%s", [vcn_id])
    cur.execute("DELETE FROM vcn_header WHERE id=%s", [vcn_id])
    cur.execute("DELETE FROM vessel_customers WHERE id=%s", [cust_id])
    conn.commit(); conn.close()


def test_removed_parcel_is_flagged_not_deleted(vessel):
    vcn_id, pids, _, _, cust_id = vessel
    vcn_model.set_parcel_removed(pids[0], True, 'Cargo cancelled by shipper', 'pytest')
    rows = vcn_model.get_consigners(vcn_id)
    assert len(rows) == 2                       # still listed, for display
    removed = next(r for r in rows if r['id'] == pids[0])
    assert removed['is_removed'] is True
    assert removed['removed_reason'] == 'Cargo cancelled by shipper'
    assert removed['quantity'] == '100'          # declaration untouched


def test_removed_parcel_drops_out_of_billing(vessel):
    """The trap: flagging the OP would have billed it at 100 MT instead."""
    vcn_id, pids, _, _, cust_id = vessel
    before = fin.get_customer_billables('Customer', cust_id)
    v = next((x for x in before.get('vessels', []) if x['vcn_id'] == vcn_id), None)
    assert v is not None and len(v['lines']) > 0
    parcels_before = {l.get('parcel_no') for l in v['lines']}
    assert 'P1' in parcels_before

    vcn_model.set_parcel_removed(pids[0], True, 'cancelled', 'pytest')

    after = fin.get_customer_billables('Customer', cust_id)
    v2 = next((x for x in after.get('vessels', []) if x['vcn_id'] == vcn_id), None)
    parcels_after = {l.get('parcel_no') for l in (v2['lines'] if v2 else [])}
    assert 'P1' not in parcels_after
    assert 'P2' in parcels_after                 # the other parcel is unaffected


def test_removed_parcel_drops_out_of_closure_and_targets(vessel):
    vcn_id, pids, ldud_id, ops, cust_id = vessel
    assert ldud_model.get_closure_eligibility(ldud_id)['bl_total'] == 200.0

    vcn_model.set_parcel_removed(pids[0], True, 'cancelled', 'pytest')

    # the op-snapshot fallback must NOT resurrect the removed quantity
    assert ldud_model.get_closure_eligibility(ldud_id)['bl_total'] == 100.0
    p = next(x for x in lueu_model.get_started_parcels(vcn_id)
             if x['parcel_op_id'] == ops[0])
    assert p['target_qty'] == 0.0
    assert p['is_removed'] is True


def test_removed_parcel_is_not_pickable_and_not_in_header_cargo(vessel):
    vcn_id, pids, _, _, cust_id = vessel
    vcn_model.set_parcel_removed(pids[0], True, 'cancelled', 'pytest')
    assert [p['parcel_no'] for p in vcn_model.get_picker_parcels(vcn_id)] == ['P2']
    assert vcn_model.get_header_cargo_type(vcn_id) == 'ACID'


def test_restore_brings_it_all_back(vessel):
    vcn_id, pids, ldud_id, _, cust_id = vessel
    vcn_model.set_parcel_removed(pids[0], True, 'cancelled', 'pytest')
    vcn_model.set_parcel_removed(pids[0], False, None, 'pytest')

    assert ldud_model.get_closure_eligibility(ldud_id)['bl_total'] == 200.0
    assert len(vcn_model.get_picker_parcels(vcn_id)) == 2
    row = next(r for r in vcn_model.get_consigners(vcn_id) if r['id'] == pids[0])
    assert row['is_removed'] is False
    assert row['removed_reason'] is None


def test_billed_parcel_cannot_be_removed(vessel):
    vcn_id, pids, _, _, cust_id = vessel
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('SELECT id FROM finance_service_types LIMIT 1')
    fin.record_parcel_charge(cur, 'VCN_IMPORT', pids[0], cur.fetchone()['id'],
                             'CHGU01', None, 100, 'pytest')
    conn.commit(); conn.close()
    with pytest.raises(ValueError, match='billed'):
        vcn_model.set_parcel_removed(pids[0], True, 'too late', 'pytest')
    assert vcn_model.get_consigners(vcn_id)[0]['is_removed'] is False


def test_reason_is_required_to_remove(vessel):
    _, pids, _, _, _ = vessel
    for reason in (None, '', '   '):
        with pytest.raises(ValueError, match='reason'):
            vcn_model.set_parcel_removed(pids[0], True, reason, 'pytest')
