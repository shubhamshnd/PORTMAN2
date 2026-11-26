import sys
sys.path.insert(0, '.')
from database import get_db, get_cursor

conn = get_db()
cur = get_cursor(conn)

print("--- All VCNs in vcn_header ---")
cur.execute("SELECT id, vcn_doc_num, via_number, vessel_name, customer_name, vessel_agent_name, importer_exporter_name FROM vcn_header ORDER BY id;")
for r in cur.fetchall():
    print(dict(r))

print("\n--- All LDUD headers ---")
cur.execute("SELECT id, vcn_id, doc_status, cast_off_datetime, discharge_commenced, discharge_completed FROM ldud_header ORDER BY id;")
for r in cur.fetchall():
    print(dict(r))

print("\n--- All LDUD Parcel Ops ---")
cur.execute("SELECT id, ldud_id, cargo_name, quantity, start_dt, end_dt FROM ldud_parcel_ops ORDER BY id;")
for r in cur.fetchall():
    print(dict(r))
