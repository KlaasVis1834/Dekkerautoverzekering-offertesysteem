from __future__ import annotations

import sqlite3
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent
DB_PATH = PROJECT_ROOT / "data" / "app.db"


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """
    (Re)initialiseert de database met alle tabellen.
    Wordt gebruikt door de UI reset-knop.
    """
    with connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS offers (
                offer_no TEXT PRIMARY KEY,
                created_at TEXT DEFAULT (datetime('now')),
                month_key TEXT,
                batch_id TEXT,

                klantnaam TEXT,
                adres TEXT,
                postcode TEXT,
                plaats TEXT,
                telefoon TEXT,
                email TEXT,

                kenteken TEXT,
                merk TEXT,
                model TEXT,
                type_model TEXT,

                klant_type TEXT,       -- particulier/prospect/zakelijk
                voertuig_type TEXT,    -- personenauto/bestelauto

                bouwjaar INTEGER,
                regio INTEGER,
                dekking TEXT,
                benodigde_svj INTEGER,

                delivery_method TEXT,  -- email/post
                delivery_status TEXT,  -- nieuw/email_klaar/post_klaar/...
                offer_pdf_path TEXT,
                eml_path TEXT,
                post_letter_path TEXT,

                is_blocked INTEGER DEFAULT 0,
                block_reason TEXT,
                block_note TEXT,

                follow_up_due_at TEXT,
                last_call_at TEXT,
                call_status TEXT DEFAULT 'open',          -- open/te_bellen/gebeld/niet_bellen
                decision_status TEXT DEFAULT 'onbekend',   -- onbekend/akkoord/niet_akkoord
                call_notes TEXT
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS offer_counters (
                month_key TEXT PRIMARY KEY,
                last_seq INTEGER NOT NULL
            )
        """)

        conn.commit()


def ensure_offer_counter(conn: sqlite3.Connection) -> None:
    """
    Zorgt dat de offer_counters tabel bestaat.
    Handig als je scripts direct draait zonder eerst init_db.
    """
    conn.execute("""
        CREATE TABLE IF NOT EXISTS offer_counters (
            month_key TEXT PRIMARY KEY,
            last_seq INTEGER NOT NULL
        )
    """)
    conn.commit()


def next_offer_no(conn: sqlite3.Connection, month_key: str) -> str:
    """
    Genereert offertenummer:
      YYMM-XXX
    Waarbij XXX per maand oploopt vanaf 001.
    """
    ensure_offer_counter(conn)

    row = conn.execute(
        "SELECT last_seq FROM offer_counters WHERE month_key = ?",
        (month_key,),
    ).fetchone()

    if row is None:
        last_seq = 0
        conn.execute(
            "INSERT INTO offer_counters (month_key, last_seq) VALUES (?, ?)",
            (month_key, last_seq),
        )
        conn.commit()
    else:
        last_seq = int(row["last_seq"])

    new_seq = last_seq + 1
    conn.execute(
        "UPDATE offer_counters SET last_seq = ? WHERE month_key = ?",
        (new_seq, month_key),
    )
    conn.commit()

    # month_key = "YYYY-MM"
    yy = month_key[2:4]
    mm = month_key[5:7]
    return f"{yy}{mm}-{new_seq:03d}"
