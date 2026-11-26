import sys
sys.path.insert(0, '.')
from database import get_db, get_cursor

conn = get_db()
cur = get_cursor(conn)

print("--- 1. Searching for SC JASPER or JASPER ---")
cur.execute("SELECT * FROM mis_history WHERE LOWER(vessel_name) LIKE '%jasper%';")
rows = cur.fetchall()
print(f"mis_history rows: {len(rows)}")
for r in rows:
    print(dict(r))

cur.execute("""
    SELECT v.*, l.id as ldud_id, l.cast_off_datetime, l.discharge_commenced, l.discharge_completed
    FROM vcn_header v
    LEFT JOIN ldud_header l ON l.vcn_id = v.id
    WHERE LOWER(v.vessel_name) LIKE '%jasper%';
""")
v_rows = cur.fetchall()
print(f"\nvcn_header rows: {len(v_rows)}")
for r in v_rows:
    print({k: v for k, v in dict(r).items() if v is not None})

print("\n--- 2. Checking vcn_header columns ---")
cur.execute("""
    SELECT column_name, data_type 
    FROM information_schema.columns 
    WHERE table_name = 'vcn_header'
    ORDER BY ordinal_position;
""")
cols = cur.fetchall()
for c in cols:
    print(c['column_name'], c['data_type'])
