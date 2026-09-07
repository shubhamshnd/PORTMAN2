"""The LDUD01 grid must show only the cargo side its VCN actually operates.

Import and export parcels live in different tables. Flipping vcn_header.
operation_type leaves the old side's rows behind, and summing both made the
grid report roughly double the real quantity (MT Hodaka Galaxy, 2026-09-05).
"""
from database import get_db, get_cursor
from modules.LDUD01 import model as ldud_model
from modules.VCN01 import model as vcn_model


def _seed(cur, op_type):
    """VCN carrying parcels on BOTH sides — the state a flipped op type leaves."""
    cur.execute("INSERT INTO vcn_header (operation_type, vcn_doc_num) VALUES (%s, %s) RETURNING id",
                (op_type, 'VCN-TEST-OPTYPE'))
    vcn_id = cur.fetchone()['id']
    cur.execute("""INSERT INTO vcn_consigners (vcn_id, cargo_name, quantity)
                   VALUES (%s, 'IMPORT OIL', '100')""", (vcn_id,))
    cur.execute("""INSERT INTO vcn_export_cargo_declaration (vcn_id, cargo_name, quantity)
                   VALUES (%s, 'EXPORT OIL', '250')""", (vcn_id,))
    cur.execute("INSERT INTO ldud_header (vcn_id, doc_num) VALUES (%s, %s) RETURNING id",
                (vcn_id, 'LDUD-TEST-OPTYPE'))
    return vcn_id, cur.fetchone()['id']


def _grid_row(ldud_id):
    rows, _ = ldud_model.get_data(1, 50, [{'field': 'doc_num', 'type': 'contains',
                                           'value': 'LDUD-TEST-OPTYPE'}])
    return next(r for r in rows if r['id'] == ldud_id)


def _cleanup(vcn_id):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute('DELETE FROM ldud_header WHERE vcn_id=%s', (vcn_id,))
    cur.execute('DELETE FROM vcn_header WHERE id=%s', (vcn_id,))   # parcels cascade
    conn.commit(); conn.close()


def test_export_vcn_ignores_stale_import_parcels():
    conn = get_db(); cur = get_cursor(conn)
    vcn_id, ldud_id = _seed(cur, 'Export'); conn.commit(); conn.close()
    try:
        row = _grid_row(ldud_id)
        # 250 only — NOT 350 (the old bug summed both sides)
        assert row['bl_quantities_display'] == '250.000 MT'
        assert row['cargo_names_display'] == 'EXPORT OIL'
    finally:
        _cleanup(vcn_id)


def test_import_vcn_ignores_stale_export_parcels():
    conn = get_db(); cur = get_cursor(conn)
    vcn_id, ldud_id = _seed(cur, 'Import'); conn.commit(); conn.close()
    try:
        row = _grid_row(ldud_id)
        assert row['bl_quantities_display'] == '100.000 MT'
        assert row['cargo_names_display'] == 'IMPORT OIL'
    finally:
        _cleanup(vcn_id)


def test_header_cargo_type_follows_operation_type():
    """_sync_header_cargo must not UNION both parcel tables."""
    conn = get_db(); cur = get_cursor(conn)
    vcn_id, _ = _seed(cur, 'Export'); conn.commit(); conn.close()
    try:
        conn = get_db(); cur = get_cursor(conn)
        vcn_model._sync_header_cargo(cur, vcn_id)
        conn.commit(); conn.close()
        assert vcn_model.get_header_cargo_type(vcn_id) == 'EXPORT OIL'
    finally:
        _cleanup(vcn_id)
