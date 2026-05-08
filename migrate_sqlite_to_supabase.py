from pathlib import Path
import sqlite3
import os
import psycopg
from psycopg.rows import dict_row

PROJECT_ROOT = Path(__file__).resolve().parent

SQLITE_DB = PROJECT_ROOT / "data" / "app.db"

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL ontbreekt")


def pg_connect():
    return psycopg.connect(
        DATABASE_URL,
        row_factory=dict_row,
        prepare_threshold=None,
    )


def sqlite_connect():
    conn = sqlite3.connect(SQLITE_DB)
    conn.row_factory = sqlite3.Row
    return conn


def table_columns_sqlite(conn, table):
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return [r["name"] for r in rows]


def table_columns_pg(conn, table):
    rows = conn.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema='public'
          AND table_name=%s
        ORDER BY ordinal_position
        """,
        (table,),
    ).fetchall()

    return [r["column_name"] for r in rows]


def migrate_table(table_name, conflict_column=None):
    print(f"\n--- Migreren: {table_name}")

    with sqlite_connect() as sq_conn, pg_connect() as pg_conn:

        sqlite_cols = table_columns_sqlite(sq_conn, table_name)
        pg_cols = table_columns_pg(pg_conn, table_name)

        common_cols = [c for c in sqlite_cols if c in pg_cols]

        if not common_cols:
            print("Geen overeenkomende kolommen")
            return

        rows = sq_conn.execute(
            f"SELECT {', '.join(common_cols)} FROM {table_name}"
        ).fetchall()

        if not rows:
            print("Geen records")
            return

        placeholders = ", ".join(["%s"] * len(common_cols))
        columns_sql = ", ".join(common_cols)

        if conflict_column:
            sql = f"""
                INSERT INTO {table_name} ({columns_sql})
                VALUES ({placeholders})
                ON CONFLICT ({conflict_column}) DO NOTHING
            """
        else:
            sql = f"""
                INSERT INTO {table_name} ({columns_sql})
                VALUES ({placeholders})
            """

        inserted = 0

        for row in rows:
            values = [row[c] for c in common_cols]

            try:
                pg_conn.execute(sql, values)
                inserted += 1
            except Exception as e:
                print(f"FOUT row: {e}")

        pg_conn.commit()

        print(f"{inserted} records gemigreerd")


def main():
    print("=== SQLITE -> SUPABASE MIGRATIE ===")

    migrate_table("offers", "offer_no")
    migrate_table("offer_counters", "month_key")
    migrate_table("no_plate_vehicles", "id")
    migrate_table("users", "username")

    print("\nKLAAR")


if __name__ == "__main__":
    main()