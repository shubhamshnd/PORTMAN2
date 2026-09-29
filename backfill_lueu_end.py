"""One-off fix for existing LUEU01 parcels: end_dt was being set to the last
logbook row's date/time instead of the row where quantity actually completed
(see lueu01.html maybeAutoEnd fix). Recomputes end_dt for every parcel that
already has one, using the same "row where running qty first reaches target"
rule, and updates any that drifted onto a later delay row or kept the
previous day's date on a row that crosses midnight.

Usage: python backfill_lueu_end.py [--apply]
Without --apply it only prints what would change (dry run).
"""
import datetime
import sys
from database import get_db, get_cursor
from modules.LUEU01.model import _single_parcel_target


def _tmin(hhmm):
    h, _, m = str(hhmm).partition(':')
    return int(h or 0) * 60 + int(m or 0)


def correct_end(cur, parcel_op_id, target):
    cur.execute('''SELECT entry_date, from_time, to_time, COALESCE(quantity,0) AS q
                   FROM lueu_parcel_log
                   WHERE parcel_op_id=%s AND is_deleted IS NOT TRUE
                   ORDER BY entry_date, from_time NULLS LAST, id''', [parcel_op_id])
    running = 0.0
    for r in cur.fetchall():
        running += float(r['q'] or 0)
        if running >= target - 1e-6:
            if r['entry_date'] and r['to_time']:
                hhmm = r['to_time'] if len(str(r['to_time'])) == 5 else str(r['to_time']).zfill(5)
                date = r['entry_date']
                # row crossing midnight (To < From, e.g. 23:00 -> 00:30) ends the next day
                if r['from_time'] and _tmin(hhmm) < _tmin(r['from_time']):
                    date = (datetime.date.fromisoformat(date) + datetime.timedelta(days=1)).isoformat()
                return f"{date}T{hhmm}"
            return None  # completion row has no to_time yet — leave end untouched
    return None  # target not actually reached — leave untouched


def main():
    apply = '--apply' in sys.argv
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT id, end_dt FROM ldud_parcel_ops WHERE end_dt IS NOT NULL AND end_dt != '' ORDER BY id")
    parcels = cur.fetchall()

    changed = 0
    for p in parcels:
        target = _single_parcel_target(cur, p['id'])
        if target <= 0:
            continue
        new_end = correct_end(cur, p['id'], target)
        if new_end and new_end != p['end_dt']:
            changed += 1
            print(f"parcel_op_id={p['id']}: {p['end_dt']} -> {new_end}")
            if apply:
                cur.execute('UPDATE ldud_parcel_ops SET end_dt=%s WHERE id=%s', [new_end, p['id']])

    if apply:
        conn.commit()
        print(f"\nUpdated {changed} parcel(s).")
    else:
        print(f"\n{changed} parcel(s) would change. Re-run with --apply to write.")
    conn.close()


if __name__ == '__main__':
    main()
