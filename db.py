from __future__ import annotations

import os
from pathlib import Path

import psycopg
from psycopg.rows import dict_row

PROJECT_ROOT = Path(__file__).resolve().parent
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


def connect():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL ontbreekt. Zet deze in Render Environment Variables.")

    return psycopg.connect(
        DATABASE_URL,
        row_factory=dict_row,
        connect_timeout=15,
        prepare_threshold=None,
    )


def init_db() -> None:
    """
    Initialiseert de Supabase/PostgreSQL database met de basistabellen.
    """
    with connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS offers (
                offer_no TEXT PRIMARY KEY,
                created_at TEXT,
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

                klant_type TEXT,
                voertuig_type TEXT,

                bouwjaar INTEGER,
                regio INTEGER,
                dekking TEXT,
                benodigde_svj INTEGER,

                delivery_method TEXT,
                delivery_status TEXT,
                offer_pdf_path TEXT,
                eml_path TEXT,
                post_letter_path TEXT,

                is_blocked INTEGER DEFAULT 0,
                block_reason TEXT,
                block_note TEXT,

                follow_up_due_at TEXT,
                last_call_at TEXT,
                call_status TEXT DEFAULT 'open',
                decision_status TEXT DEFAULT 'onbekend',
                call_notes TEXT,

                maandpremie DOUBLE PRECISION,
                dienstverlening_bedrag DOUBLE PRECISION,
                svj_override INTEGER,
                is_bestaande_klant INTEGER DEFAULT 0,

                revision_of TEXT,
                revision_no INTEGER DEFAULT 0,

                dekking_override TEXT,
                extra_svi INTEGER DEFAULT 0,
                extra_rb INTEGER DEFAULT 0,

                created_by TEXT,
                updated_by TEXT,
                updated_at TEXT,

                no_plate_vehicle_id INTEGER,
                np_gewicht TEXT,
                np_maandpremie DOUBLE PRECISION,
                np_cataloguswaarde TEXT,
                np_cataloguswaarde_part TEXT,
                np_cataloguswaarde_zak TEXT,

                mail_template_type TEXT DEFAULT 'auto'
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS offer_counters (
                month_key TEXT PRIMARY KEY,
                last_seq INTEGER NOT NULL
            )
            """
        )

        conn.commit()


def ensure_offer_counter(conn) -> None:
    """
    Zorgt dat de offer_counters tabel bestaat.
    """
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS offer_counters (
            month_key TEXT PRIMARY KEY,
            last_seq INTEGER NOT NULL
        )
        """
    )
    conn.commit()


def next_offer_no(conn, month_key: str) -> str:
    """
    Genereert offertenummer:
      YYMM-XXX
    Waarbij XXX per maand oploopt vanaf 001.
    """
    ensure_offer_counter(conn)

    row = conn.execute(
        "SELECT last_seq FROM offer_counters WHERE month_key = %s",
        (month_key,),
    ).fetchone()

    if row is None:
        last_seq = 0
        conn.execute(
            "INSERT INTO offer_counters (month_key, last_seq) VALUES (%s, %s)",
            (month_key, last_seq),
        )
        conn.commit()
    else:
        last_seq = int(row["last_seq"])

    new_seq = last_seq + 1

    conn.execute(
        "UPDATE offer_counters SET last_seq = %s WHERE month_key = %s",
        (new_seq, month_key),
    )
    conn.commit()

    yy = month_key[2:4]
    mm = month_key[5:7]
    return f"{yy}{mm}-{new_seq:03d}"
