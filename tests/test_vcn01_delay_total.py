"""get_data's total_delay_mins sums the Pre Berthing/Anchoring Delays sub-table.
Uses the dev DB directly, like the other VCN01 tests; the throwaway header is
deleted at the end (ON DELETE CASCADE clears the delays)."""
from database import get_db, get_cursor
from modules.VCN01 import model


def _find(rows, vcn_id):
    return next(r for r in rows if r['id'] == vcn_id)


def test_total_delay_mins_sums_valid_rows_only():
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("INSERT INTO vcn_header (operation_type) VALUES ('Import') RETURNING id")
    vcn_id = cur.fetchone()['id']
    cur.executemany(
        "INSERT INTO vcn_delays (vcn_id, delay_name, delay_start, delay_end) VALUES (%s,%s,%s,%s)",
        [
            (vcn_id, 'Weather',    '2026-01-01T10:00', '2026-01-01T14:30'),   # 270 min
            (vcn_id, 'Tide',       '2026-01-02T00:00', '2026-01-03T02:00'),   # 1560 min
            (vcn_id, 'Open',       '2026-01-04T08:00', ''),                   # ignored
            (vcn_id, 'Reversed',   '2026-01-05T09:00', '2026-01-05T08:00'),   # ignored
            (vcn_id, 'Junk',       'not-a-date',       'also-junk'),          # ignored
        ])
    conn.commit(); conn.close()
    try:
        rows, _ = model.get_data(page=1, size=200)
        assert _find(rows, vcn_id)['total_delay_mins'] == 270 + 1560
    finally:
        model.delete_header(vcn_id)


def test_total_delay_mins_zero_without_delays():
    conn = get_db(); cur = get_cursor(conn)
    cur.execute("INSERT INTO vcn_header (operation_type) VALUES ('Import') RETURNING id")
    vcn_id = cur.fetchone()['id']
    conn.commit(); conn.close()
    try:
        rows, _ = model.get_data(page=1, size=200)
        assert _find(rows, vcn_id)['total_delay_mins'] == 0
    finally:
        model.delete_header(vcn_id)
