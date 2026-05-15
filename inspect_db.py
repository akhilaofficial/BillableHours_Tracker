import sqlite3

conn = sqlite3.connect("hours.db")
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

tables = cursor.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
print("TABLES:", [t[0] for t in tables])
for t in tables:
    count = cursor.execute("SELECT COUNT(*) FROM " + t[0]).fetchone()[0]
    cols = [c[1] for c in cursor.execute("PRAGMA table_info(" + t[0] + ")").fetchall()]
    print(f"\n  [{t[0]}]  {count} rows")
    print(f"  columns: {', '.join(cols)}")

conn.close()
