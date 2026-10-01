import os
import sys

print("=" * 70)
print("REEDOY v49 - ONLINE WORKERS SCHEMA CHECK")
print("=" * 70)

url = os.environ.get("DATABASE_URL", "").strip()

if not url:
    print("[ERROR] DATABASE_URL is not set on this PC.")
    print()
    print("এটা স্বাভাবিক হতে পারে। Render-এর DATABASE_URL আপনার PC-তে")
    print("থাকার কথা নয়।")
    sys.exit(1)

try:
    import psycopg2
except ImportError:
    print("[ERROR] psycopg2 is not installed.")
    print("Run: pip install psycopg2-binary")
    sys.exit(1)

try:
    conn = psycopg2.connect(
        url,
        sslmode="require",
        connect_timeout=10
    )

    cur = conn.cursor()

    cur.execute("""
        SELECT column_name, data_type
        FROM information_schema.columns
        WHERE table_schema='public'
          AND table_name='workers'
        ORDER BY ordinal_position
    """)

    columns = cur.fetchall()

    cur.execute("SELECT COUNT(*) FROM workers")
    count = cur.fetchone()[0]

    print("[OK] Connected to PostgreSQL")
    print()
    print("ONLINE WORKERS")
    print("-" * 70)
    print("Worker count:", count)
    print()
    print("Columns:")
    for name, dtype in columns:
        print(f"  {name}: {dtype}")

    cur.close()
    conn.close()

    print()
    print("=" * 70)
    print("CHECK COMPLETE - NO DATA WAS MODIFIED")
    print("=" * 70)

except Exception as e:
    print("[ERROR] Could not connect to Online PostgreSQL")
    print(type(e).__name__ + ":", e)
