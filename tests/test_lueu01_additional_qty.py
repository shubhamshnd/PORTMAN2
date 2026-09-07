"""Over-discharge is reconciled in LUEU01 without touching the VCN.

Before this, quantity logged above the declared BL was discarded by the
completion cap in get_started_parcels — yet FIN01 still billed it, because
billing sums the raw log rows. Recording an additional quantity raises the
target so the real rows count everywhere at once.
"""
from database import get_db, get_cursor
from modules.LDUD01 import model as ldud_model
from modules.LUEU01 import model as lueu_model


def _seed(declared='100', logged=(60, 60)):
    """VCN + parcel + LDUD + op, with `logged` MT of log rows against it."""
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""INSERT INTO vcn_header (operation_type, vessel_name, vcn_doc_num, nor_tendered)
                   VALUES ('Import','MT ADDL TEST','VCN-TEST-ADDL','2026-09-01T08:00')
                   RETURNING id""")
    vcn_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_consigners (vcn_id, cargo_name, quantity, parcel_seq, parcel_no)
                   VALUES (%s,'OIL',%s,1,'P1') RETURNING id""", [vcn_id, declared])
    parcel_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO ldud_header (vcn_id, vessel_name, nor_tendered)
                   VALUES (%s,'MT ADDL TEST','2026-09-01T08:00') RETURNING id""", [vcn_id])
    ldud_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO ldud_parcel_ops (ldud_id, parcel_ids, cargo_name, quantity)
                   VALUES (%s,%s,'OIL',%s) RETURNING id""", [ldud_id, str(parcel_id), declared])
    op_id = cur.fetchone()['id']
    for i, q in enumerate(logged):
        cur.execute("""INSERT INTO lueu_parcel_log
                       (parcel_op_id, entry_date, from_time, to_time, quantity, quantity_uom)
                       VALUES (%s,'2026-09-01',%s,%s,%s,'MT')""",
                    [op_id, f'{8+i:02d}:00', f'{9+i:02d}:00', q])
    conn.commit(); conn.close()
    return vcn_id, ldud_id, op_id


def _drop(vcn_id):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("""DELETE FROM lueu_parcel_log WHERE parcel_op_id IN
                   (SELECT po.id FROM ldud_parcel_ops po JOIN ldud_header l ON l.id=po.ldud_id
                    WHERE l.vcn_id=%s)""", [vcn_id])
    cur.execute("DELETE FROM ldud_header WHERE vcn_id=%s", [vcn_id])   # ops cascade
    cur.execute("DELETE FROM vcn_header WHERE id=%s", [vcn_id])
    conn.commit(); conn.close()


def _parcel(vcn_id, op_id):
    return next(p for p in lueu_model.get_started_parcels(vcn_id)
                if p['parcel_op_id'] == op_id)


def test_rows_past_the_target_are_dropped_while_billing_still_counts_them():
    """The divergence this feature closes.

    180 MT logged against a 100 MT BL. The cap admits rows until the target is
    crossed, then discards the rest — so LUEU01 shows 120. FIN01 sums the raw
    log rows and sees all 180. Same parcel, two different quantities.
    """
    vcn_id, ldud_id, op_id = _seed(declared='100', logged=(60, 60, 60))
    try:
        p = _parcel(vcn_id, op_id)
        assert p['target_qty'] == 100.0
        assert p['logged_qty'] == 120.0       # third row silently dropped

        conn = get_db(); cur = get_cursor(conn)
        cur.execute("""SELECT COALESCE(SUM(quantity),0) AS q FROM lueu_parcel_log
                       WHERE parcel_op_id=%s AND is_deleted IS NOT TRUE""", [op_id])
        billed_view = float(cur.fetchone()['q'])
        conn.close()
        assert billed_view == 180.0           # what billing charges
        assert billed_view != p['logged_qty']  # ... and what the screen shows
    finally:
        _drop(vcn_id)


def test_additional_quantity_raises_the_target_and_reveals_the_excess():
    vcn_id, ldud_id, op_id = _seed(declared='100', logged=(60, 60, 60))
    try:
        lueu_model.set_additional_qty(op_id, 80, 'Discharged in excess of BL', 'pytest')

        p = _parcel(vcn_id, op_id)
        assert p['target_qty'] == 180.0                # 100 declared + 80 amended
        assert p['logged_qty'] == 180.0                # all three rows now count
        assert p['remaining_qty'] == 0.0
        assert p['additional_qty'] == 80.0
        assert p['additional_reason'] == 'Discharged in excess of BL'

        # closure agrees, so the vessel can still Full Close
        elig = ldud_model.get_closure_eligibility(ldud_id)
        assert elig['bl_total'] == 180.0
        assert elig['ops_total'] == 180.0
        assert elig['can_full_close'] is True
    finally:
        _drop(vcn_id)


def test_the_vcn_declared_quantity_is_never_touched():
    vcn_id, ldud_id, op_id = _seed(declared='100', logged=(60, 60))
    try:
        lueu_model.set_additional_qty(op_id, 80, 'excess', 'pytest')
        conn = get_db(); cur = get_cursor(conn)
        cur.execute('SELECT quantity FROM vcn_consigners WHERE vcn_id=%s', [vcn_id])
        assert cur.fetchone()['quantity'] == '100'
        conn.close()
    finally:
        _drop(vcn_id)


def test_clearing_restores_the_declared_target():
    vcn_id, ldud_id, op_id = _seed(declared='100', logged=(60, 60, 60))
    try:
        lueu_model.set_additional_qty(op_id, 80, 'excess', 'pytest')
        lueu_model.clear_additional_qty(op_id, 'pytest')
        p = _parcel(vcn_id, op_id)
        assert p['target_qty'] == 100.0
        assert p['logged_qty'] == 120.0     # back to the capped view
        assert ldud_model.get_closure_eligibility(ldud_id)['bl_total'] == 100.0
    finally:
        _drop(vcn_id)


def test_grid_bl_total_includes_the_amendment():
    vcn_id, ldud_id, op_id = _seed(declared='100', logged=(60, 60))
    try:
        lueu_model.set_additional_qty(op_id, 80, 'excess', 'pytest')
        rows, _ = ldud_model.get_data(1, 50, [])
        row = next(r for r in rows if r['id'] == ldud_id)
        assert row['bl_quantities_display'] == '180.000 MT'
    finally:
        _drop(vcn_id)


def test_validation():
    vcn_id, ldud_id, op_id = _seed()
    try:
        for qty, reason in ((0, 'x'), (-5, 'x'), (10, ''), (10, '   ')):
            try:
                lueu_model.set_additional_qty(op_id, qty, reason, 'pytest')
                assert False, f'accepted {qty!r}/{reason!r}'
            except ValueError:
                pass
        try:
            lueu_model.clear_additional_qty(op_id, 'pytest')
            assert False, 'cleared when nothing was set'
        except ValueError:
            pass
    finally:
        _drop(vcn_id)
