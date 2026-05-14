# ui/server.py
import sys
import time
import re
import os
import secrets
import requests
import msal
import base64
import json
from functools import wraps
from pathlib import Path
from datetime import datetime
from pypdf import PdfReader, PdfWriter

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from psycopg.rows import dict_row
from psycopg import OperationalError
from psycopg_pool import ConnectionPool
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, jsonify, session,
)
from werkzeug.security import generate_password_hash, check_password_hash

from importers import import_excel, get_last_batch_id
from outlook_msg import write_msg_outlook
from voertuigdata import get_vehicle_info
from rules import bepaal_dekking
from rolls_kiwa import get_meldcode_en_type
from rdw_estimator import estimate_vehicle_data_from_rdw
from pdfgen import generate_offer_pdf
from postgen import generate_post_letter_pdf
from mailgen import load_template, render_template as render_mail_template, guess_aanhef_en_achternaam

AANVRAAG_LINK = "https://www.klaasvis.nl/aanvraagformulier/"
AANVRAAG_API_SECRET = os.environ.get("AANVRAAG_API_SECRET", "").strip()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dekker-offertesysteem-local")

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
LOCAL_SQLITE_DB_PATH = PROJECT_ROOT / "data" / "app.db"

MICROSOFT_CLIENT_ID = os.environ.get("MICROSOFT_CLIENT_ID", "").strip()
MICROSOFT_TENANT_ID = os.environ.get("MICROSOFT_TENANT_ID", "").strip()
MICROSOFT_CLIENT_SECRET = os.environ.get("MICROSOFT_CLIENT_SECRET", "").strip()
MICROSOFT_REDIRECT_URI = os.environ.get("MICROSOFT_REDIRECT_URI", "").strip()
MICROSOFT_AUTHORITY = f"https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}" if MICROSOFT_TENANT_ID else ""
MICROSOFT_SCOPES = ["User.Read", "Mail.ReadWrite"]

DB_READY = False
DB_POOL = None


@app.after_request
def add_no_cache_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

DEFAULT_USERS = [
    ("randy", "Randy", "admin"),
    ("tim", "Tim", "medewerker"),
    ("dirk", "Dirk", "medewerker"),
    ("marcel", "Marcel", "medewerker"),
    ("petra", "Petra", "medewerker"),
    ("marjolein", "Marjolein", "medewerker"),
]


def current_user_display():
    return session.get("display_name") or session.get("username") or "Onbekend"


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("Alleen de beheerder heeft toegang tot gebruikersbeheer.", "error")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return wrapper


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


def connect():
    global DB_POOL

    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL ontbreekt. Zet deze in Render Environment Variables.")

    if DB_POOL is None:
        DB_POOL = ConnectionPool(
            conninfo=DATABASE_URL,
            min_size=1,
            max_size=5,
            kwargs={
                "row_factory": dict_row,
                "connect_timeout": 15,
                "prepare_threshold": None,
            },
        )

    return DB_POOL.connection()


def _table_columns(conn, table_name: str) -> set[str]:
    rows = conn.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = 'public'
          AND table_name = %s
        """,
        (table_name,),
    ).fetchall()
    return {r["column_name"] for r in rows}


def _ensure_column(conn, table: str, col: str, ddl_type: str):
    cols = _table_columns(conn, table)
    if col not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl_type}")


def _execute_retry(conn, sql: str, params=(), retries: int = 8, sleep_s: float = 0.25):
    last_err = None
    for _ in range(retries):
        try:
            return conn.execute(sql, params)
        except OperationalError as e:
            last_err = e
            time.sleep(sleep_s)
    raise last_err


def seed_default_users(conn):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for username, display_name, role in DEFAULT_USERS:
        existing = conn.execute(
            "SELECT id FROM users WHERE username = %s",
            (username,),
        ).fetchone()

        if existing:
            continue

        temp_password = f"{display_name}2026!"
        conn.execute(
            """
            INSERT INTO users (username, display_name, password_hash, role, created_at, active)
            VALUES (%s, %s, %s, %s, %s, 1)
            """,
            (username, display_name, generate_password_hash(temp_password), role, now),
        )


def ensure_db():
    global DB_READY

    if DB_READY:
        return

    with connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS offers (
                offer_no TEXT PRIMARY KEY,
                created_at TEXT,
                month_key TEXT,
                batch_id TEXT,
                klantnaam TEXT,
                klant_type TEXT,
                adres TEXT,
                postcode TEXT,
                plaats TEXT,
                telefoon TEXT,
                email TEXT,
                kenteken TEXT,
                merk TEXT,
                model TEXT,
                type_model TEXT,
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
                call_status TEXT DEFAULT 'open',
                decision_status TEXT DEFAULT 'open',
                call_notes TEXT,
                last_call_at TEXT
            )
            """
        )

        for col, ddl in [
            ("maandpremie", "DOUBLE PRECISION"),
            ("dienstverlening_bedrag", "DOUBLE PRECISION"),
            ("svj_override", "INTEGER"),
            ("is_bestaande_klant", "INTEGER DEFAULT 0"),
            ("revision_of", "TEXT"),
            ("revision_no", "INTEGER DEFAULT 0"),
            ("dekking_override", "TEXT"),
            ("extra_svi", "INTEGER DEFAULT 0"),
            ("extra_rb", "INTEGER DEFAULT 0"),
            ("created_by", "TEXT"),
            ("updated_by", "TEXT"),
            ("updated_at", "TEXT"),
            ("mail_template_type", "TEXT DEFAULT 'auto'"),
            ("no_plate_vehicle_id", "INTEGER"),
            ("np_gewicht", "TEXT"),
            ("np_maandpremie", "DOUBLE PRECISION"),
            ("np_cataloguswaarde", "TEXT"),
            ("np_cataloguswaarde_part", "TEXT"),
            ("np_cataloguswaarde_zak", "TEXT"),
        ]:
            _ensure_column(conn, "offers", col, ddl)

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS no_plate_vehicles (
                id SERIAL PRIMARY KEY,
                merk TEXT,
                model TEXT,
                type_model TEXT,
                voertuig_type TEXT,
                bouwjaar INTEGER,
                brandstof TEXT,
                cataloguswaarde TEXT,
                gewicht TEXT,
                cataloguswaarde_part TEXT,
                cataloguswaarde_zak TEXT,
                premie_part_r1 DOUBLE PRECISION,
                premie_part_r2 DOUBLE PRECISION,
                premie_part_r3 DOUBLE PRECISION,
                premie_part_r4 DOUBLE PRECISION,
                premie_zak_r1 DOUBLE PRECISION,
                premie_zak_r2 DOUBLE PRECISION,
                premie_zak_r3 DOUBLE PRECISION,
                premie_zak_r4 DOUBLE PRECISION,
                created_at TEXT
            )
            """
        )

        _ensure_column(conn, "no_plate_vehicles", "brandstof", "TEXT")
        _ensure_column(conn, "no_plate_vehicles", "cataloguswaarde_part", "TEXT")
        _ensure_column(conn, "no_plate_vehicles", "cataloguswaarde_zak", "TEXT")

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                display_name TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'medewerker',
                created_at TEXT
            )
            """
        )

        for col, ddl in [
            ("active", "INTEGER DEFAULT 1"),
            ("last_login_at", "TEXT"),
            ("ms_graph_email", "TEXT"),
            ("ms_graph_access_token", "TEXT"),
            ("ms_graph_refresh_token", "TEXT"),
            ("ms_graph_token_expires_at", "TEXT"),
            ("ms_graph_connected_at", "TEXT"),
        ]:
            _ensure_column(conn, "users", col, ddl)

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS applications (
                id SERIAL PRIMARY KEY,
                offer_no TEXT,
                naam TEXT,
                email TEXT,
                telefoon TEXT,
                status TEXT DEFAULT 'nieuw',
                source TEXT DEFAULT 'aanvraagformulier',
                raw_payload TEXT,
                pdf_path TEXT,
                json_path TEXT,
                created_at TEXT,
                updated_at TEXT
            )
            """
        )

        for col, ddl in [
            ("aanvraag_ontvangen_at", "TEXT"),
            ("aanvraag_naam", "TEXT"),
            ("aanvraag_email", "TEXT"),
            ("aanvraag_status", "TEXT DEFAULT 'geen'"),
        ]:
            _ensure_column(conn, "offers", col, ddl)

        seed_default_users(conn)
        conn.commit()

    DB_READY = True


@app.route("/login", methods=["GET", "POST"])
def login():
    session.pop("_flashes", None)
    session.modified = True
    ensure_db()

    if session.get("logged_in"):
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = (request.form.get("username") or "").strip().lower()
        password = request.form.get("password") or ""

        with connect() as conn:
            user = conn.execute(
                """
                SELECT id, username, display_name, password_hash, role, active
                FROM users
                WHERE username = %s
                """,
                (username,),
            ).fetchone()

        if user and int(user["active"] or 1) != 1:
            flash("Dit account is uitgeschakeld.", "error")
            return render_template("login.html")

        if user and check_password_hash(user["password_hash"], password):
            session["logged_in"] = True
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["display_name"] = user["display_name"]
            session["role"] = user["role"]

            with connect() as conn:
                conn.execute(
                    "UPDATE users SET last_login_at = %s WHERE id = %s",
                    (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user["id"]),
                )
                conn.commit()

            return redirect(url_for("dashboard"))

        flash("Ongeldige gebruikersnaam of wachtwoord.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/clear-flashes", methods=["GET", "POST"])
@login_required
def clear_flashes():
    session.pop("_flashes", None)
    session.modified = True

    if request.method == "GET":
        return redirect(url_for("dashboard"))

    return ("", 204)


@app.route("/account", methods=["GET", "POST"])
@login_required
def account():
    ensure_db()

    if request.method == "POST":
        current_password = request.form.get("current_password") or ""
        new_password = request.form.get("new_password") or ""
        confirm_password = request.form.get("confirm_password") or ""

        if len(new_password) < 8:
            flash("Nieuw wachtwoord moet minimaal 8 tekens bevatten.", "error")
            return redirect(url_for("account"))

        if new_password != confirm_password:
            flash("De nieuwe wachtwoorden komen niet overeen.", "error")
            return redirect(url_for("account"))

        username = session.get("username")

        with connect() as conn:
            user = conn.execute(
                "SELECT id, password_hash FROM users WHERE username = %s",
                (username,),
            ).fetchone()

            if not user or not check_password_hash(user["password_hash"], current_password):
                flash("Huidig wachtwoord klopt niet.", "error")
                return redirect(url_for("account"))

            conn.execute(
                "UPDATE users SET password_hash = %s WHERE id = %s",
                (generate_password_hash(new_password), user["id"]),
            )
            conn.commit()

        flash("Wachtwoord succesvol gewijzigd.", "ok")
        return redirect(url_for("account"))

    with connect() as conn:
        user = conn.execute(
            """
            SELECT ms_graph_email, ms_graph_connected_at
            FROM users
            WHERE id = %s
            """,
            (session.get("user_id"),),
        ).fetchone()

    return render_template(
        "account.html",
        ms_graph_email=user["ms_graph_email"] if user else None,
        ms_graph_connected_at=user["ms_graph_connected_at"] if user else None,
    )


@app.route("/account/microsoft/connect")
@login_required
def microsoft_connect():
    ensure_db()

    if not MICROSOFT_CLIENT_ID or not MICROSOFT_CLIENT_SECRET or not MICROSOFT_TENANT_ID or not MICROSOFT_REDIRECT_URI:
        flash("Microsoft Graph is nog niet volledig ingesteld in Render.", "error")
        return redirect(url_for("account"))

    app_msal = msal.ConfidentialClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=MICROSOFT_AUTHORITY,
        client_credential=MICROSOFT_CLIENT_SECRET,
    )

    flow = app_msal.initiate_auth_code_flow(
        scopes=MICROSOFT_SCOPES,
        redirect_uri=MICROSOFT_REDIRECT_URI,
        prompt="select_account",
    )

    session["ms_auth_flow"] = flow
    return redirect(flow["auth_uri"])


@app.route("/auth/microsoft/callback")
@login_required
def microsoft_callback():
    ensure_db()

    flow = session.get("ms_auth_flow")
    if not flow:
        flash("Microsoft koppeling mislukt: sessie verlopen. Probeer opnieuw.", "error")
        return redirect(url_for("account"))

    app_msal = msal.ConfidentialClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=MICROSOFT_AUTHORITY,
        client_credential=MICROSOFT_CLIENT_SECRET,
    )

    result = app_msal.acquire_token_by_auth_code_flow(flow, dict(request.args))

    if "access_token" not in result:
        msg = result.get("error_description") or result.get("error") or "geen toegangstoken ontvangen"
        flash(f"Microsoft koppeling mislukt: {msg}", "error")
        return redirect(url_for("account"))

    access_token = result.get("access_token")
    refresh_token = result.get("refresh_token")
    expires_in = int(result.get("expires_in", 3600))
    expires_at = datetime.fromtimestamp(datetime.now().timestamp() + expires_in).strftime("%Y-%m-%d %H:%M:%S")

    profile_res = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=20,
    )

    if profile_res.status_code >= 400:
        flash("Microsoft account gekoppeld, maar profiel ophalen is mislukt.", "error")
        return redirect(url_for("account"))

    profile = profile_res.json()
    graph_email = profile.get("mail") or profile.get("userPrincipalName") or ""

    with connect() as conn:
        conn.execute(
            """
            UPDATE users
            SET ms_graph_email = %s,
                ms_graph_access_token = %s,
                ms_graph_refresh_token = %s,
                ms_graph_token_expires_at = %s,
                ms_graph_connected_at = %s
            WHERE id = %s
            """,
            (
                graph_email,
                access_token,
                refresh_token,
                expires_at,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                session.get("user_id"),
            ),
        )
        conn.commit()

    session.pop("ms_auth_flow", None)

    flash(f"Microsoft Outlook gekoppeld: {graph_email}", "ok")
    return redirect(url_for("account"))


@app.post("/account/microsoft/disconnect")
@login_required
def microsoft_disconnect():
    ensure_db()

    with connect() as conn:
        conn.execute(
            """
            UPDATE users
            SET ms_graph_email = NULL,
                ms_graph_access_token = NULL,
                ms_graph_refresh_token = NULL,
                ms_graph_token_expires_at = NULL,
                ms_graph_connected_at = NULL
            WHERE id = %s
            """,
            (session.get("user_id"),),
        )
        conn.commit()

    session.pop("ms_auth_flow", None)

    flash("Microsoft Outlook koppeling verwijderd.", "ok")
    return redirect(url_for("account"))


@app.route("/admin/users")
@admin_required
def admin_users():
    ensure_db()

    with connect() as conn:
        users = conn.execute(
            """
            SELECT id, username, display_name, role, active, created_at, last_login_at
            FROM users
            ORDER BY 
                CASE WHEN role = 'admin' THEN 0 ELSE 1 END,
                display_name ASC
            """
        ).fetchall()

    return render_template("admin_users.html", users=users)


@app.post("/admin/users/<int:user_id>/reset-password")
@admin_required
def admin_reset_password(user_id: int):
    ensure_db()

    with connect() as conn:
        user = conn.execute(
            "SELECT id, username, display_name FROM users WHERE id = %s",
            (user_id,),
        ).fetchone()

        if not user:
            flash("Gebruiker niet gevonden.", "error")
            return redirect(url_for("admin_users"))

        new_password = f"{user['display_name']}2026!"

        conn.execute(
            "UPDATE users SET password_hash = %s WHERE id = %s",
            (generate_password_hash(new_password), user_id),
        )
        conn.commit()

    flash(f"Wachtwoord gereset voor {user['display_name']}. Tijdelijk wachtwoord: {new_password}", "ok")
    return redirect(url_for("admin_users"))


@app.post("/admin/users/<int:user_id>/toggle-active")
@admin_required
def admin_toggle_user_active(user_id: int):
    ensure_db()

    if user_id == session.get("user_id"):
        flash("Je kunt je eigen account niet uitschakelen.", "error")
        return redirect(url_for("admin_users"))

    with connect() as conn:
        user = conn.execute(
            "SELECT id, display_name, active FROM users WHERE id = %s",
            (user_id,),
        ).fetchone()

        if not user:
            flash("Gebruiker niet gevonden.", "error")
            return redirect(url_for("admin_users"))

        new_active = 0 if int(user["active"] or 1) == 1 else 1

        conn.execute(
            "UPDATE users SET active = %s WHERE id = %s",
            (new_active, user_id),
        )
        conn.commit()

    status = "ingeschakeld" if new_active == 1 else "uitgeschakeld"
    flash(f"{user['display_name']} is {status}.", "ok")
    return redirect(url_for("admin_users"))


def denylist_exists() -> bool:
    return (PROJECT_ROOT / "data" / "denylist.docx").exists()


def safe_relpath(p: Path) -> str:
    try:
        return str(p.relative_to(PROJECT_ROOT)).replace("\\", "/")
    except Exception:
        return str(p).replace("\\", "/")
        
def combine_post_package_pdf(post_letter_path: str, offer_pdf_path: str, offer_no: str, klantnaam: str) -> str:
    post_abs = (PROJECT_ROOT / post_letter_path).resolve()
    offer_abs = (PROJECT_ROOT / offer_pdf_path).resolve()

    if not post_abs.exists():
        raise RuntimeError("Postbrief niet gevonden voor postpakket.")

    if not offer_abs.exists():
        raise RuntimeError("Offerte-PDF niet gevonden voor postpakket.")

    out_dir = PROJECT_ROOT / "data" / "post"
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = _safe_filename(f"Postpakket_{_short_offer_no_for_filename(offer_no)}_{klantnaam or 'klant'}.pdf")
    out_abs = out_dir / filename

    writer = PdfWriter()

    for pdf_abs in [post_abs, offer_abs]:
        reader = PdfReader(str(pdf_abs))
        for page in reader.pages:
            writer.add_page(page)

    with out_abs.open("wb") as f:
        writer.write(f)

    return safe_relpath(out_abs)
    
def _parse_float(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("€", "").strip()
    s = s.replace(".", "").replace(",", ".") if ("," in s and "." in s) else s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _parse_int(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def _load_template_safe(rel_path: str):
    try:
        return load_template(rel_path)
    except Exception:
        return None


def _pick_np_premie(np_row, klant_type: str, regio: int):
    kt = (klant_type or "").strip().lower()
    try:
        r = int(regio)
    except Exception:
        r = 0

    if r not in (1, 2, 3, 4):
        return None

    col = f"premie_zak_r{r}" if kt == "zakelijk" else f"premie_part_r{r}"
    v = np_row[col]
    try:
        return float(v) if v is not None and str(v).strip() != "" else None
    except Exception:
        return None


def _pick_np_catalogus(np_row, klant_type: str, offer_row):
    kt = (klant_type or "").strip().lower()

    if kt == "zakelijk":
        for key in ("np_cataloguswaarde_zak", "np_cataloguswaarde"):
            if key in offer_row.keys():
                v = (offer_row[key] or "").strip()
                if v:
                    return v
        for key in ("cataloguswaarde_zak", "cataloguswaarde"):
            if key in np_row.keys():
                v = (np_row[key] or "").strip()
                if v:
                    return v
        return ""

    for key in ("np_cataloguswaarde_part", "np_cataloguswaarde"):
        if key in offer_row.keys():
            v = (offer_row[key] or "").strip()
            if v:
                return v
    for key in ("cataloguswaarde_part", "cataloguswaarde"):
        if key in np_row.keys():
            v = (np_row[key] or "").strip()
            if v:
                return v
    return ""


def _kenteken_lookup_value(kenteken: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", (kenteken or "")).upper().strip()


def _format_nl_kenteken(kenteken: str) -> str:
    s = _kenteken_lookup_value(kenteken)
    if not s:
        return ""

    if len(s) != 6:
        return s

    patterns = [
        (r"^[A-Z]{2}\d{4}$", (2, 4)),        # XX-99-99
        (r"^\d{4}[A-Z]{2}$", (2, 4)),        # 99-99-XX
        (r"^\d{2}[A-Z]{2}\d{2}$", (2, 4)),  # 99-XX-99
        (r"^[A-Z]{2}\d{2}[A-Z]{2}$", (2, 4)), # XX-99-XX
        (r"^[A-Z]{4}\d{2}$", (2, 4)),        # XX-XX-99
        (r"^\d{2}[A-Z]{4}$", (2, 4)),        # 99-XX-XX
        (r"^\d{2}[A-Z]{3}\d$", (2, 5)),      # 99-XXX-9
        (r"^\d[A-Z]{3}\d{2}$", (1, 4)),      # 9-XXX-99
        (r"^[A-Z]{2}\d{3}[A-Z]$", (2, 5)),  # XX-999-X
        (r"^[A-Z]\d{3}[A-Z]{2}$", (1, 4)),  # X-999-XX
        (r"^[A-Z]{3}\d{2}[A-Z]$", (3, 5)),  # XXX-99-X
        (r"^[A-Z]\d{2}[A-Z]{3}$", (1, 3)),  # X-99-XXX
        (r"^\d[A-Z]{2}\d{3}$", (1, 3)),      # 9-XX-999
        (r"^\d{3}[A-Z]{2}\d$", (3, 5)),      # 999-XX-9
    ]

    for pattern, cuts in patterns:
        if re.match(pattern, s):
            a, b = cuts
            return f"{s[:a]}-{s[a:b]}-{s[b:]}"

    return f"{s[:2]}-{s[2:4]}-{s[4:]}"


def _safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[<>:\"/\\\\|?*]", " ", s)
    s = re.sub(r"[\x00-\x1f]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _strip_known_titles(name: str) -> str:
    s = (name or "").strip()
    s = re.sub(r"^(de\s+heer|dhr\.?|heer)\s+", "", s, flags=re.I)
    s = re.sub(r"^(mevrouw|mevr\.?|mw\.?)\s+", "", s, flags=re.I)
    s = re.sub(r"^(mr\.?|dr\.?|ing\.?|ir\.?)\s+", "", s, flags=re.I)
    return s.strip()


def _initials_from_name(full_name: str, achternaam: str) -> str:
    full = _strip_known_titles(full_name)
    a = (achternaam or "").strip()

    base = full[: len(full) - len(a)].strip() if a and full.lower().endswith(a.lower()) else full
    if not base:
        return ""

    tokens = [t for t in re.split(r"\s+", base) if t.strip()]
    existing = [t for t in tokens if "." in t]
    if existing:
        return " ".join(existing).strip()

    initials = []
    for t in tokens:
        for p in re.split(r"[-/]", t):
            p = re.sub(r"[^A-Za-zÀ-ÿ]", "", p)
            if p:
                initials.append(p[0].upper() + ".")
    return " ".join(initials).strip()


def _klant_display_for_filename(klantnaam: str, klant_type: str) -> str:
    kt = (klant_type or "").strip().lower()
    kn = (klantnaam or "").strip()

    if kt == "zakelijk":
        return kn

    aanhef, achternaam = guess_aanhef_en_achternaam(kn)
    initials = _initials_from_name(kn, achternaam)
    parts = [p for p in [aanhef, initials, achternaam] if p]
    return " ".join(parts).strip() or kn


def _short_offer_no_for_filename(offer_no: str) -> str:
    s = (offer_no or "").strip()
    parts = s.split("-")
    if parts and parts[-1].isdigit():
        parts[-1] = parts[-1].lstrip("0") or "0"
        return "-".join(parts)
    return s


def _offer_pdf_filename_base(klantnaam: str, klant_type: str, offer_no: str) -> str:
    return _safe_filename(
        f"Verzekeringsvoorstel voor {_klant_display_for_filename(klantnaam, klant_type)} - {_short_offer_no_for_filename(offer_no)}"
    )


def _normalize_dekking_override(v: str) -> str:
    s = (v or "").strip().lower()
    if not s:
        return ""
    if s == "wa":
        return "WA"
    if s in ("wa beperkt casco", "wa / beperkt casco", "beperkt casco"):
        return "WA / Beperkt Casco"
    if s in (
        "allrisk",
        "wa casco compleet",
        "wa / casco compleet",
        "wa / casco compleet (allrisk)",
        "casco compleet",
        "casco compleet (allrisk)",
    ):
        return "WA / Casco Compleet (Allrisk)"
    return (v or "").strip()


def _compose_dekking(auto_dekking: str, dekking_override: str, extra_svi, extra_rb) -> str:
    dekking = _normalize_dekking_override(dekking_override) or (auto_dekking or "").strip()
    parts = [p.strip() for p in str(dekking).split("/") if p.strip()]
    seen = {p.lower() for p in parts}

    if int(extra_svi or 0) == 1:
        svi = "Schadeverzekering voor Inzittenden"
        if svi.lower() not in seen and "schadeverzekering inzittenden" not in seen:
            parts.append(svi)
            seen.add(svi.lower())

    if int(extra_rb or 0) == 1:
        rb = "Rechtsbijstandsverzekering Verkeer"
        if rb.lower() not in seen:
            parts.append(rb)
            seen.add(rb.lower())

    return " / ".join(parts)


@app.context_processor
def inject_application_counts():
    try:
        ensure_db()
        with connect() as conn:
            pending = conn.execute(
                """
                SELECT COUNT(*) AS c
                FROM applications
                WHERE COALESCE(status, 'nieuw') = 'nieuw'
                """
            ).fetchone()["c"]
        return {"pending_applications_count": pending}
    except Exception:
        return {"pending_applications_count": 0}


@app.route("/")
@login_required
def dashboard():
    ensure_db()
    last_batch_id = get_last_batch_id()

    with connect() as conn:
        total = conn.execute("SELECT COUNT(*) AS c FROM offers").fetchone()["c"]
        open_deliveries = conn.execute(
            """
            SELECT COUNT(*) AS c FROM offers
            WHERE is_blocked = 0 AND delivery_status IN ('email_klaar','post_klaar','outlook_concept_klaar')
            """
        ).fetchone()["c"]
        blocked = conn.execute("SELECT COUNT(*) AS c FROM offers WHERE is_blocked = 1").fetchone()["c"]

    return render_template(
        "dashboard.html",
        total=total,
        open_deliveries=open_deliveries,
        blocked=blocked,
        last_batch_id=last_batch_id,
        denylist_exists=denylist_exists(),
        db_path="Supabase PostgreSQL",
    )


@app.route("/followups")
@login_required
def followups():
    flash("Bel-lijst is uitgeschakeld. Je zit nu op Offertes.", "ok")
    return redirect(url_for("offers"))


@app.route("/import", methods=["GET", "POST"])
@login_required
def import_page():
    ensure_db()

    inbox_dir = PROJECT_ROOT / "data" / "inbox"
    inbox_dir.mkdir(parents=True, exist_ok=True)
    fixed_deny = PROJECT_ROOT / "data" / "denylist.docx"

    if request.method == "POST":
        # Oude meldingen wissen voordat een nieuwe import start.
        session.pop("_flashes", None)
        session.modified = True

        excel_file = (
            request.files.get("excel")
            or request.files.get("file")
            or request.files.get("upload")
            or request.files.get("bestand")
        )
        deny_file = request.files.get("denylist")

        if not excel_file or not (excel_file.filename or "").strip():
            flash("Kies een Excel bestand.", "error")
            return redirect(url_for("import_page"))

        original_name = _safe_filename(excel_file.filename or "import.xlsx")
        if not original_name.lower().endswith((".xlsx", ".xls")):
            flash("Import fout: upload een Excel-bestand (.xlsx of .xls).", "error")
            return redirect(url_for("import_page"))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = inbox_dir / f"{timestamp}_{original_name}"
        excel_file.save(excel_path)

        if deny_file and (deny_file.filename or "").strip():
            fixed_deny.parent.mkdir(parents=True, exist_ok=True)
            deny_file.save(fixed_deny)

        deny_path = str(fixed_deny) if fixed_deny.exists() else None

        try:
            n = import_excel(str(excel_path), deny_path)
            batch_id = get_last_batch_id()

            if batch_id:
                with connect() as conn:
                    conn.execute(
                        """
                        UPDATE offers
                        SET created_by = COALESCE(NULLIF(created_by, ''), %s),
                            updated_by = %s,
                            updated_at = %s
                        WHERE batch_id = %s
                        """,
                        (
                            current_user_display(),
                            current_user_display(),
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            batch_id,
                        ),
                    )
                    conn.commit()

            flash(f"Import gelukt: {n} offertes toegevoegd.", "ok")
            return redirect(url_for("offers", delivery="all"))

        except Exception as e:
            print("IMPORT FOUT:", repr(e))
            flash(f"Import fout: {type(e).__name__}: {e}", "error")
            return redirect(url_for("import_page"))

    excel_candidates = sorted(
        inbox_dir.glob("*.xls*"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )[:10]

    return render_template(
        "import.html",
        excel_candidates=[safe_relpath(p) for p in excel_candidates],
        denylist_present=fixed_deny.exists(),
        denylist_path=safe_relpath(fixed_deny) if fixed_deny.exists() else None,
    )

@app.route("/offers")
@login_required
def offers():
    session.pop("_flashes", None)
    session.modified = True
    ensure_db()

    q = request.args.get("q", "").strip()
    month = request.args.get("month", "").strip()
    delivery = request.args.get("delivery", "").strip()

    try:
        page = int(request.args.get("page", "1"))
    except Exception:
        page = 1

    if page < 1:
        page = 1

    per_page = 50
    offset = (page - 1) * per_page

    sql = """
    SELECT offer_no, created_at, month_key, batch_id,
           klantnaam, klant_type, email, telefoon,
           kenteken, merk, model, type_model, voertuig_type, bouwjaar,
           regio, dekking,
           delivery_method, delivery_status, offer_pdf_path, eml_path, post_letter_path,
           is_blocked, block_reason, block_note,
           follow_up_due_at, call_status, decision_status,
           maandpremie, dienstverlening_bedrag,
           svj_override, is_bestaande_klant,
           revision_of, revision_no,
           no_plate_vehicle_id, np_gewicht, np_maandpremie,
           np_cataloguswaarde, np_cataloguswaarde_part, np_cataloguswaarde_zak,
           created_by, updated_by, updated_at, mail_template_type
    FROM offers
    WHERE 1=1
    """
    params = []

    count_sql = "SELECT COUNT(*) AS c FROM offers WHERE 1=1"
    count_params = []

    if month:
        sql += " AND month_key = %s"
        count_sql += " AND month_key = %s"
        params.append(month)
        count_params.append(month)

    if q:
        sql += " AND (klantnaam ILIKE %s OR kenteken ILIKE %s OR email ILIKE %s OR offer_no ILIKE %s)"
        count_sql += " AND (klantnaam ILIKE %s OR kenteken ILIKE %s OR email ILIKE %s OR offer_no ILIKE %s)"
        like = f"%{q}%"
        params.extend([like, like, like, like])
        count_params.extend([like, like, like, like])

    if delivery and delivery != "all":
        if delivery == "geblokkeerd":
            sql += " AND is_blocked = 1"
            count_sql += " AND is_blocked = 1"
        else:
            sql += " AND is_blocked = 0 AND delivery_status = %s"
            count_sql += " AND is_blocked = 0 AND delivery_status = %s"
            params.append(delivery)
            count_params.append(delivery)

    sql += """
    ORDER BY
        CASE
            WHEN COALESCE(offer_pdf_path,'')='' AND COALESCE(eml_path,'')='' AND COALESCE(post_letter_path,'')=''
            THEN 0 ELSE 1
        END ASC,
        CASE WHEN lower(COALESCE(klant_type,''))='zakelijk' THEN 1 ELSE 0 END ASC,
        COALESCE(regio, 999) ASC,
        created_at DESC,
        offer_no DESC
    LIMIT %s OFFSET %s
    """
    params.extend([per_page, offset])

    with connect() as conn:
        rows = conn.execute(sql, tuple(params)).fetchall()
        total_rows = conn.execute(count_sql, tuple(count_params)).fetchone()["c"]
        months = conn.execute(
            """
            SELECT DISTINCT month_key FROM offers
            WHERE month_key IS NOT NULL AND month_key != ''
            ORDER BY month_key DESC
            LIMIT 24
            """
        ).fetchall()

    total_pages = max(1, (total_rows + per_page - 1) // per_page)

    return render_template(
        "offers.html",
        rows=rows,
        q=q,
        month=month,
        delivery=delivery or "all",
        months=[r["month_key"] for r in months],
        page=page,
        per_page=per_page,
        total_rows=total_rows,
        total_pages=total_pages,
        has_prev=page > 1,
        has_next=page < total_pages,
    )


@app.route("/blocked")
@login_required
def blocked():
    ensure_db()

    with connect() as conn:
        rows = conn.execute(
            """
            SELECT offer_no, klantnaam, kenteken, email, block_reason, block_note, created_at
            FROM offers
            WHERE is_blocked = 1
            ORDER BY created_at DESC
            LIMIT 500
            """
        ).fetchall()

    return render_template("blocked.html", rows=rows)


@app.route("/aanvragen", endpoint="aanvragen")
@app.route("/applications", endpoint="applications")
@login_required
def aanvragen():
    session.pop("_flashes", None)
    session.modified = True
    ensure_db()

    try:
        with connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    a.id,
                    a.offer_no,
                    a.naam,
                    a.email,
                    a.telefoon,
                    a.status,
                    a.created_at,
                    a.updated_at,
                    a.pdf_path,
                    a.json_path,
                    o.klantnaam
                FROM applications a
                LEFT JOIN offers o ON TRIM(o.offer_no) = TRIM(a.offer_no)
                ORDER BY
                    CASE WHEN COALESCE(a.status, 'nieuw') = 'nieuw' THEN 0 ELSE 1 END,
                    a.created_at DESC,
                    a.id DESC
                LIMIT 500
                """
            ).fetchall()
    except Exception as e:
        print("Aanvragen ophalen mislukt:", repr(e))
        rows = []

    return render_template("aanvragen.html", rows=rows)

@app.post("/applications/<int:application_id>/complete")
@login_required
def complete_application(application_id: int):
    ensure_db()

    try:
        with connect() as conn:
            conn.execute(
                """
                UPDATE applications
                SET status = 'afgehandeld',
                    updated_at = %s
                WHERE id = %s
                """,
                (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    application_id,
                ),
            )
            conn.commit()
    except Exception as e:
        print("APPLICATION COMPLETE FOUT:", repr(e))
        flash(f"Aanvraag afhandelen mislukt: {type(e).__name__}: {e}", "error")

    return redirect(url_for("applications"))
    
@app.route("/api/aanvraag", methods=["POST", "OPTIONS"])
def api_aanvraag_ontvangen():
    if request.method == "OPTIONS":
        response = jsonify({"ok": True})
        response.headers["Access-Control-Allow-Origin"] = "https://www.klaasvis.nl"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Aanvraag-Secret"
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        return response
    ensure_db()

    if AANVRAAG_API_SECRET:
        received_secret = (
            request.headers.get("X-Aanvraag-Secret")
            or request.form.get("secret")
            or ""
        ).strip()

        if received_secret != AANVRAAG_API_SECRET:
            return jsonify({"ok": False, "error": "unauthorized"}), 401

    data = request.get_json(silent=True) or request.form.to_dict() or {}

    offer_no = (
        data.get("offer_no")
        or data.get("offerte_nummer")
        or data.get("offerte")
        or ""
    ).strip()

    naam = (
        data.get("naam")
        or data.get("name")
        or data.get("customer_name")
        or data.get("klantnaam")
        or ""
    ).strip()

    email = (
        data.get("email")
        or data.get("from_email")
        or ""
    ).strip()

    telefoon = (
        data.get("telefoon")
        or data.get("phone")
        or data.get("from_phone")
        or ""
    ).strip()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    applications_dir = PROJECT_ROOT / "data" / "applications"
    applications_dir.mkdir(parents=True, exist_ok=True)

    safe_offer = _safe_filename(offer_no or "zonder_offertenummer")
    json_filename = _safe_filename(
        f"aanvraag_{safe_offer}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    )
    json_abs = applications_dir / json_filename
    json_abs.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    json_path = safe_relpath(json_abs)

    with connect() as conn:
        conn.execute(
            """
            INSERT INTO applications (
                offer_no,
                naam,
                email,
                telefoon,
                status,
                source,
                raw_payload,
                json_path,
                created_at,
                updated_at
            )
            VALUES (%s, %s, %s, %s, 'nieuw', 'aanvraagformulier', %s, %s, %s, %s)
            """,
            (
                offer_no,
                naam,
                email,
                telefoon,
                json.dumps(data, ensure_ascii=False),
                json_path,
                now,
                now,
            ),
        )

        if offer_no:
            conn.execute(
                """
                UPDATE offers
                SET aanvraag_ontvangen_at = %s,
                    aanvraag_naam = %s,
                    aanvraag_email = %s,
                    aanvraag_status = 'nieuw',
                    decision_status = 'aanvraag_ontvangen',
                    updated_by = 'Aanvraagformulier',
                    updated_at = %s
                WHERE TRIM(offer_no) = TRIM(%s)
                """,
                (
                    now,
                    naam,
                    email,
                    now,
                    offer_no,
                ),
            )

        conn.commit()

    response = jsonify({"ok": True, "offer_no": offer_no, "json_path": json_path})
    response.headers["Access-Control-Allow-Origin"] = "https://www.klaasvis.nl"
    return response


@app.post("/offer/<offer_no>/update-meta")
@login_required
def update_offer_meta(offer_no: str):
    ensure_db()
    next_url = request.form.get("next") or url_for("offers")

    maandpremie = _parse_float((request.form.get("maandpremie") or "").strip())
    np_maandpremie = _parse_float((request.form.get("np_maandpremie") or "").strip())
    svj_override = _parse_int((request.form.get("svj_override") or "").strip())

    dekking_override = _normalize_dekking_override(request.form.get("dekking_override") or "") or None
    extra_svi = 1 if request.form.get("extra_svi") in ("1", "on", "true", "True") else 0
    extra_rb = 1 if request.form.get("extra_rb") in ("1", "on", "true", "True") else 0
    is_bestaande_klant = 1 if request.form.get("is_bestaande_klant") in ("1", "on", "true", "True") else 0

    revision_of = (request.form.get("revision_of") or "").strip() or None
    revision_no = _parse_int((request.form.get("revision_no") or "").strip())
    if revision_of is None:
        revision_no = 0

    mail_template_type = (request.form.get("mail_template_type") or "auto").strip()
    if mail_template_type not in (
        "auto",
        "definitief",
        "prospect",
        "bestaand_particulier",
        "bestaand_zakelijk",
        "aangepast",
    ):
        mail_template_type = "auto"

    dienstverlening = round(maandpremie * 0.18, 2) if maandpremie is not None else None

    with connect() as conn:
        _execute_retry(
            conn,
            """
            UPDATE offers
            SET maandpremie = %s,
                dienstverlening_bedrag = %s,
                np_maandpremie = %s,
                svj_override = %s,
                dekking_override = %s,
                extra_svi = %s,
                extra_rb = %s,
                is_bestaande_klant = %s,
                revision_of = %s,
                revision_no = %s,
                mail_template_type = %s,
                updated_by = %s,
                updated_at = %s
            WHERE TRIM(offer_no) = TRIM(%s)
            """,
            (
                maandpremie,
                dienstverlening,
                np_maandpremie,
                svj_override,
                dekking_override,
                extra_svi,
                extra_rb,
                is_bestaande_klant,
                revision_of,
                revision_no,
                mail_template_type,
                current_user_display(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                offer_no,
            ),
        )
        conn.commit()

    flash(f"Gegevens opgeslagen voor {offer_no}.", "ok")
    return redirect(next_url)


@app.post("/offer/<offer_no>/decision")
@login_required
def set_decision(offer_no: str):
    ensure_db()

    decision = (request.form.get("decision") or "").strip()
    next_url = request.form.get("next") or url_for("offers")

    if decision not in ("akkoord", "niet_akkoord", "open", "aanvraag_ontvangen"):
        flash("Ongeldige keuze.", "error")
        return redirect(next_url)

    with connect() as conn:
        new_call_status = "afgehandeld" if decision in ("akkoord", "niet_akkoord") else "open"
        _execute_retry(
            conn,
            """
            UPDATE offers
            SET decision_status = %s,
                call_status = %s,
                last_call_at = CASE WHEN %s != 'open' THEN %s ELSE last_call_at END,
                updated_by = %s,
                updated_at = %s
            WHERE TRIM(offer_no) = TRIM(%s)
            """,
            (
                decision,
                new_call_status,
                decision,
                datetime.now().strftime("%Y-%m-%d"),
                current_user_display(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                offer_no,
            ),
        )
        conn.commit()

    flash(f"Beslissing opgeslagen: {offer_no} → {decision}", "ok")
    return redirect(next_url)



@app.post("/offer/<offer_no>/delete")
@login_required
def delete_offer(offer_no: str):
    ensure_db()

    next_url = request.form.get("next") or url_for("offers")

    # Verwijder alle bestaande flash-meldingen direct.
    session.pop("_flashes", None)
    session.modified = True

    try:
        with connect() as conn:
            row = conn.execute(
                """
                SELECT offer_no, offer_pdf_path, eml_path, post_letter_path
                FROM offers
                WHERE TRIM(offer_no) = TRIM(%s)
                """,
                (offer_no,),
            ).fetchone()

            # Eventuele bestanden verwijderen
            if row:
                for rel in (
                    row["offer_pdf_path"],
                    row["eml_path"],
                    row["post_letter_path"],
                ):
                    if not rel:
                        continue

                    try:
                        p = (PROJECT_ROOT / rel).resolve()
                        if p.exists() and PROJECT_ROOT in p.parents:
                            p.unlink()
                    except Exception as file_error:
                        print("Bestand verwijderen overgeslagen:", repr(file_error))

            # Gekoppelde aanvragen verwijderen
            conn.execute(
                """
                DELETE FROM applications
                WHERE TRIM(COALESCE(offer_no, '')) = TRIM(%s)
                """,
                (offer_no,),
            )

            # Offerte verwijderen
            deleted = conn.execute(
                """
                DELETE FROM offers
                WHERE TRIM(offer_no) = TRIM(%s)
                RETURNING offer_no
                """,
                (offer_no,),
            ).fetchone()

            conn.commit()

        # Alleen foutmelding tonen als er niets is verwijderd
        if not deleted:
            flash(f"Offerte niet gevonden: {offer_no}", "error")

    except Exception as e:
        print("DELETE FOUT:", repr(e))
        flash(f"Verwijderen mislukt: {type(e).__name__}: {e}", "error")

    return redirect(next_url)
    
@app.get("/offer/<offer_no>/download-postbrief")
@login_required
def download_postbrief(offer_no: str):
    ensure_db()
    now = datetime.now()

    with connect() as conn:
        r = conn.execute(
            "SELECT * FROM offers WHERE TRIM(offer_no) = TRIM(%s)",
            (offer_no,),
        ).fetchone()

        if not r:
            flash("Offerte niet gevonden.", "error")
            return redirect(url_for("offers"))

        email = (r["email"] or "").strip()
        if email:
            flash("Deze klant heeft een e-mailadres; postbrief is niet nodig.", "error")
            return redirect(url_for("offers"))

        vinfo = get_vehicle_info(
            kenteken=r["kenteken"] or "",
            merk=r["merk"] or "",
            model=r["model"] or "",
            db_path=LOCAL_SQLITE_DB_PATH,
        )

        auto_str = " ".join([x for x in [vinfo.merk, vinfo.model] if x]).strip()

        if r["post_letter_path"]:
            existing = (PROJECT_ROOT / r["post_letter_path"]).resolve()
            if existing.exists() and PROJECT_ROOT in existing.parents:
                return send_file(existing, as_attachment=True, download_name=f"Postbrief_{offer_no}.pdf")

        post_letter_path = generate_post_letter_pdf(
            out_base_dir="data/post",
            dt=now,
            offer_no=offer_no,
            klantnaam=r["klantnaam"] or "",
            adres=r["adres"] or "",
            postcode=r["postcode"] or "",
            plaats=r["plaats"] or "",
            auto=auto_str or "auto",
            behandeld_door="Dirk Slootweg",
        )

        _execute_retry(
            conn,
            """
            UPDATE offers
            SET post_letter_path = %s,
                delivery_method = 'post',
                delivery_status = 'post_klaar',
                updated_by = %s,
                updated_at = %s
            WHERE TRIM(offer_no) = TRIM(%s)
            """,
            (
                post_letter_path,
                current_user_display(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                offer_no,
            ),
        )
        conn.commit()

    abs_path = (PROJECT_ROOT / post_letter_path).resolve()
    return send_file(abs_path, as_attachment=True, download_name=f"Postbrief_{offer_no}.pdf")


@app.route("/no-plate", methods=["GET", "POST"])
@login_required
def no_plate():
    session.pop("_flashes", None)
    session.modified = True
    ensure_db()

    if request.method == "POST":
        vid = (request.form.get("id") or "").strip()

        merk = (request.form.get("merk") or "").strip()
        model = (request.form.get("model") or "").strip()
        type_model = (request.form.get("type_model") or "").strip()

        voertuig_type = (request.form.get("voertuig_type") or "personenauto").strip().lower()
        if voertuig_type not in ("personenauto", "bestelauto"):
            voertuig_type = "personenauto"

        bouwjaar = _parse_int(request.form.get("bouwjaar"))
        brandstof = (request.form.get("brandstof") or "").strip()
        gewicht = (request.form.get("gewicht") or "").strip()

        catalogus_part = (request.form.get("cataloguswaarde_part") or "").strip()
        catalogus_zak = (request.form.get("cataloguswaarde_zak") or "").strip()
        catalogus_legacy = (request.form.get("cataloguswaarde") or "").strip()

                # ---------------------------------------------------------
        # RDW schatting voor voertuigen zonder kenteken
        # Alleen lege velden worden aangevuld.
        # Handmatig ingevulde waarden blijven leidend.
        # ---------------------------------------------------------
        rdw_data = {}

        if merk and model:
            try:
                rdw_data = estimate_vehicle_data_from_rdw(
                    merk=merk,
                    model=model,
                    type_model=type_model,
                    bouwjaar=bouwjaar,
                )

                if rdw_data is None:
                    rdw_data = {}

                if not isinstance(rdw_data, dict):
                    print("RDW gaf geen dictionary terug:", rdw_data)
                    rdw_data = {}

            except Exception as e:
                print("RDW-schatting volledig overgeslagen door fout:", repr(e))
                rdw_data = {}

        if rdw_data:
            if not brandstof:
                brandstof = rdw_data.get("brandstof") or ""

            if not gewicht:
                gewicht = rdw_data.get("gewicht") or ""

            rdw_catalogus = rdw_data.get("cataloguswaarde") or ""

            if not catalogus_part:
                catalogus_part = rdw_catalogus

            if not catalogus_zak:
                catalogus_zak = rdw_catalogus

            if not catalogus_legacy:
                catalogus_legacy = rdw_catalogus

            rdw_voertuig_type = rdw_data.get("voertuig_type") or ""
            if rdw_voertuig_type in ("personenauto", "bestelauto"):
                voertuig_type = rdw_voertuig_type

        print("RDW resultaat veilig verwerkt:", rdw_data)

        def f(name):
            return _parse_float((request.form.get(name) or "").strip())

        data = (
            merk,
            model,
            type_model,
            voertuig_type,
            bouwjaar,
            brandstof,
            catalogus_legacy,
            gewicht,
            catalogus_part,
            catalogus_zak,
            f("premie_part_r1"),
            f("premie_part_r2"),
            f("premie_part_r3"),
            f("premie_part_r4"),
            f("premie_zak_r1"),
            f("premie_zak_r2"),
            f("premie_zak_r3"),
            f("premie_zak_r4"),
        )

        with connect() as conn:
            if vid:
                _execute_retry(
                    conn,
                    """
                    UPDATE no_plate_vehicles
                    SET merk=%s, model=%s, type_model=%s,
                        voertuig_type=%s, bouwjaar=%s,
                        brandstof=%s,
                        cataloguswaarde=%s,
                        gewicht=%s,
                        cataloguswaarde_part=%s,
                        cataloguswaarde_zak=%s,
                        premie_part_r1=%s, premie_part_r2=%s, premie_part_r3=%s, premie_part_r4=%s,
                        premie_zak_r1=%s, premie_zak_r2=%s, premie_zak_r3=%s, premie_zak_r4=%s
                    WHERE id=%s
                    """,
                    data + (int(vid),),
                )
                flash("No-plate voertuig bijgewerkt. RDW-schatting is toegepast waar velden leeg waren.", "ok")
            else:
                created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                _execute_retry(
                    conn,
                    """
                    INSERT INTO no_plate_vehicles (
                        merk, model, type_model,
                        voertuig_type, bouwjaar,
                        brandstof,
                        cataloguswaarde,
                        gewicht,
                        cataloguswaarde_part,
                        cataloguswaarde_zak,
                        premie_part_r1, premie_part_r2, premie_part_r3, premie_part_r4,
                        premie_zak_r1, premie_zak_r2, premie_zak_r3, premie_zak_r4,
                        created_at
                    ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """,
                    data + (created_at,),
                )
                flash("No-plate voertuig toegevoegd. RDW-schatting is toegepast waar velden leeg waren.", "ok")

            conn.commit()

        return redirect(url_for("no_plate"))

    q = (request.args.get("q") or "").strip()

    with connect() as conn:
        if q:
            like = f"%{q}%"
            rows = conn.execute(
                """
                SELECT * FROM no_plate_vehicles
                WHERE merk ILIKE %s OR model ILIKE %s OR type_model ILIKE %s
                ORDER BY created_at DESC, id DESC
                LIMIT 500
                """,
                (like, like, like),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM no_plate_vehicles ORDER BY created_at DESC, id DESC LIMIT 500"
            ).fetchall()

    return render_template("no_plate.html", rows=rows, q=q)


@app.post("/no-plate/<int:vid>/delete")
@login_required
def no_plate_delete(vid: int):
    session.pop("_flashes", None)
    session.modified = True
    ensure_db()

    session.pop("_flashes", None)
    session.modified = True

    try:
        with connect() as conn:
            _execute_retry(conn, "DELETE FROM no_plate_vehicles WHERE id = %s", (vid,))
            conn.commit()
    except Exception as e:
        print("NO-PLATE DELETE FOUT:", repr(e))
        flash(f"No-plate voertuig verwijderen mislukt: {type(e).__name__}: {e}", "error")

    return redirect(url_for("no_plate"))

@app.get("/no-plate/search")
@login_required
def no_plate_search():
    ensure_db()
    q = (request.args.get("q") or "").strip()
    limit = 25

    with connect() as conn:
        if q:
            like = f"%{q}%"
            rows = conn.execute(
                """
                SELECT id, merk, model, type_model, voertuig_type
                FROM no_plate_vehicles
                WHERE merk ILIKE %s OR model ILIKE %s OR type_model ILIKE %s
                ORDER BY created_at DESC, id DESC
                LIMIT %s
                """,
                (like, like, like, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT id, merk, model, type_model, voertuig_type
                FROM no_plate_vehicles
                ORDER BY created_at DESC, id DESC
                LIMIT %s
                """,
                (limit,),
            ).fetchall()

    items = []
    for r in rows:
        label = " ".join(
            [
                x for x in [
                    (r["merk"] or "").strip(),
                    (r["model"] or "").strip(),
                    (r["type_model"] or "").strip(),
                ] if x
            ]
        ).strip()

        items.append(
            {
                "id": r["id"],
                "label": label or f"#{r['id']}",
                "voertuig_type": (r["voertuig_type"] or "").strip(),
            }
        )

    return jsonify({"items": items})


@app.post("/offer/<offer_no>/set-no-plate")
@login_required
def set_no_plate_for_offer(offer_no: str):
    ensure_db()
    next_url = request.form.get("next") or url_for("offers")
    vid = (request.form.get("no_plate_vehicle_id") or "").strip()

    if not vid.isdigit():
        flash("Kies eerst een no-plate voertuig.", "error")
        return redirect(next_url)

    with connect() as conn:
        np_row = conn.execute(
            "SELECT * FROM no_plate_vehicles WHERE id = %s",
            (int(vid),),
        ).fetchone()

        if not np_row:
            flash("No-plate voertuig niet gevonden.", "error")
            return redirect(next_url)

        _execute_retry(
            conn,
            """
            UPDATE offers
            SET no_plate_vehicle_id = %s,
                kenteken = '',
                merk = %s,
                model = %s,
                type_model = %s,
                voertuig_type = COALESCE(NULLIF(%s,''), voertuig_type),
                bouwjaar = COALESCE(%s, bouwjaar),
                updated_by = %s,
                updated_at = %s
            WHERE TRIM(offer_no) = TRIM(%s)
            """,
            (
                int(vid),
                (np_row["merk"] or "").strip(),
                (np_row["model"] or "").strip(),
                (np_row["type_model"] or "").strip(),
                (np_row["voertuig_type"] or "").strip().lower(),
                np_row["bouwjaar"],
                current_user_display(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                offer_no,
            ),
        )
        conn.commit()

    flash(f"No-plate voertuig gekoppeld aan {offer_no}.", "ok")
    return redirect(next_url)


def _get_current_user_graph_tokens(conn):
    user_id = session.get("user_id")
    if not user_id:
        return None

    return conn.execute(
        """
        SELECT id, ms_graph_email, ms_graph_access_token,
               ms_graph_refresh_token, ms_graph_token_expires_at
        FROM users
        WHERE id = %s
        """,
        (user_id,),
    ).fetchone()


def _refresh_graph_token_if_needed(conn, user):
    if not user:
        return None

    refresh_token = user.get("ms_graph_refresh_token")
    access_token = user.get("ms_graph_access_token")
    expires_raw = user.get("ms_graph_token_expires_at")

    token_is_valid = False
    if access_token and expires_raw:
        try:
            expires_at = datetime.strptime(expires_raw, "%Y-%m-%d %H:%M:%S")
            token_is_valid = expires_at.timestamp() > (datetime.now().timestamp() + 300)
        except Exception:
            token_is_valid = False

    if access_token and token_is_valid:
        return access_token

    if not refresh_token:
        return access_token if access_token and not expires_raw else None

    app_msal = msal.ConfidentialClientApplication(
        MICROSOFT_CLIENT_ID,
        authority=MICROSOFT_AUTHORITY,
        client_credential=MICROSOFT_CLIENT_SECRET,
    )

    result = app_msal.acquire_token_by_refresh_token(
        refresh_token,
        scopes=MICROSOFT_SCOPES,
    )

    if "access_token" not in result:
        return None

    new_access_token = result.get("access_token")
    new_refresh_token = result.get("refresh_token") or refresh_token
    expires_in = int(result.get("expires_in", 3600))
    expires_at = datetime.fromtimestamp(datetime.now().timestamp() + expires_in).strftime("%Y-%m-%d %H:%M:%S")

    conn.execute(
        """
        UPDATE users
        SET ms_graph_access_token = %s,
            ms_graph_refresh_token = %s,
            ms_graph_token_expires_at = %s
        WHERE id = %s
        """,
        (new_access_token, new_refresh_token, expires_at, user["id"]),
    )

    return new_access_token


def create_outlook_draft_with_attachment(conn, to_addr, subject, body_html, pdf_path):
    user = _get_current_user_graph_tokens(conn)
    access_token = _refresh_graph_token_if_needed(conn, user)

    if not access_token:
        return None

    message_payload = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body_html,
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": to_addr,
                }
            }
        ],
    }

    create_res = requests.post(
        "https://graph.microsoft.com/v1.0/me/messages",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        json=message_payload,
        timeout=30,
    )

    if create_res.status_code >= 400:
        raise RuntimeError(f"Graph concept aanmaken mislukt: {create_res.text}")

    message = create_res.json()
    message_id = message.get("id")
    
    abs_pdf = (PROJECT_ROOT / pdf_path).resolve()
    if not abs_pdf.exists():
        raise RuntimeError("PDF-bestand niet gevonden voor Outlook-concept.")

    pdf_bytes = abs_pdf.read_bytes()
    attachment_payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": abs_pdf.name,
        "contentType": "application/pdf",
        "contentBytes": base64.b64encode(pdf_bytes).decode("ascii"),
    }

    attach_res = requests.post(
        f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        json=attachment_payload,
        timeout=30,
    )

    if attach_res.status_code >= 400:
        raise RuntimeError(f"Graph bijlage toevoegen mislukt: {attach_res.text}")

    return {
        "message_id": message_id,
        "graph_email": user["ms_graph_email"] if user else "",
    }


def _choose_mail_template(
    klant_type: str,
    is_bestaande_klant: bool,
    revision_no: int,
    mail_template_type: str,
    tpl_def,
    tpl_pro,
    tpl_bestaand_part,
    tpl_bestaand_zak,
    tpl_aangepast,
):
    selected = (mail_template_type or "auto").strip()

    if selected == "prospect":
        return tpl_pro
    if selected == "bestaand_particulier" and tpl_bestaand_part is not None:
        return tpl_bestaand_part
    if selected == "bestaand_zakelijk" and tpl_bestaand_zak is not None:
        return tpl_bestaand_zak
    if selected == "aangepast" and tpl_aangepast is not None:
        return tpl_aangepast
    if selected == "definitief":
        return tpl_def

    if revision_no > 0 and tpl_aangepast is not None:
        return tpl_aangepast

    if is_bestaande_klant:
        if klant_type == "zakelijk" and tpl_bestaand_zak is not None:
            return tpl_bestaand_zak
        if klant_type != "zakelijk" and tpl_bestaand_part is not None:
            return tpl_bestaand_part

    return tpl_pro if klant_type == "prospect" else tpl_def


def _subject_for_offer(klant_type: str, revision_no: int, offer_no: str) -> str:
    if revision_no > 0:
        return f"Aangepast verzekeringsvoorstel Dekkerautoverzekering {offer_no}"
    if klant_type == "prospect":
        return f"Indicatief verzekeringsvoorstel Dekkerautoverzekering {offer_no}"
    return f"Verzekeringsvoorstel Dekkerautoverzekering {offer_no}"


def _build_pdf_and_delivery(conn, r, now: datetime):
    offer_no = r["offer_no"]
    klant_type = (r["klant_type"] or "particulier").strip().lower()
    email = (r["email"] or "").strip() or None

    is_bestaande_klant = int(r["is_bestaande_klant"] or 0) == 1
    revision_no = int(r["revision_no"] or 0)
    mail_template_type = r.get("mail_template_type") or "auto"

    tpl_def = load_template("templates/mail_definitief.html")
    tpl_pro = load_template("templates/mail_prospect.html")
    tpl_bestaand_part = _load_template_safe("templates/mail_bestaand_particulier.html")
    tpl_bestaand_zak = _load_template_safe("templates/mail_bestaand_zakelijk.html")
    tpl_aangepast = _load_template_safe("templates/mail_aangepast.html")

    kenteken_db = (r["kenteken"] or "").strip()
    np_row = None

    if not kenteken_db and r["no_plate_vehicle_id"] is not None:
        try:
            np_row = conn.execute(
                "SELECT * FROM no_plate_vehicles WHERE id = %s",
                (int(r["no_plate_vehicle_id"]),),
            ).fetchone()
        except Exception:
            np_row = None

    if np_row:
        np_merk = (np_row["merk"] or "").strip()
        np_model = (np_row["model"] or "").strip()
        voertuig_type = (np_row["voertuig_type"] or "personenauto").strip().lower()
        if voertuig_type not in ("personenauto", "bestelauto"):
            voertuig_type = "personenauto"

        bouwjaar_final = str(np_row["bouwjaar"]) if np_row["bouwjaar"] is not None else str(r["bouwjaar"] or "").strip()
        brandstof_final = (np_row["brandstof"] or "").strip()
        gewicht_final = (r["np_gewicht"] or np_row["gewicht"] or "").strip()
        catalogus_final = _pick_np_catalogus(np_row, klant_type, r)
        auto_str = " ".join([x for x in [np_merk, np_model] if x]).strip()

        bj_int = _parse_int(bouwjaar_final)
        dekking = bepaal_dekking(bj_int)

        np_maandpremie = r["np_maandpremie"] if "np_maandpremie" in r.keys() else None
        maandpremie = r["maandpremie"] if "maandpremie" in r.keys() else None
        premie_final = np_maandpremie if np_maandpremie is not None else maandpremie

        if premie_final is None:
            premie_final = _pick_np_premie(np_row, klant_type, int(r["regio"] or 0))

        dekking_final = _compose_dekking(
            dekking,
            r["dekking_override"] if "dekking_override" in r.keys() else None,
            r["extra_svi"] if "extra_svi" in r.keys() else 0,
            r["extra_rb"] if "extra_rb" in r.keys() else 0,
        )

        offer_pdf_path = generate_offer_pdf(
            out_base_dir="data/offers",
            dt=now,
            offer_no=offer_no,
            klant={
                "naam": r["klantnaam"] or "",
                "adres": r["adres"] or "",
                "postcode": r["postcode"] or "",
                "plaats": r["plaats"] or "",
                "telefoon": r["telefoon"] or "",
                "email": email or "",
                "klant_type": klant_type,
            },
            voertuig={
                "auto": auto_str,
                "kenteken": "",
                "brandstof": brandstof_final,
                "voertuig_type": voertuig_type,
                "merk": np_merk,
                "model": np_model,
            },
            offer={
                "regio": r["regio"] if r["regio"] is not None else "",
                "dekking": dekking_final,
                "dekking_override": r["dekking_override"] or "",
                "extra_svi": r["extra_svi"],
                "extra_rb": r["extra_rb"],
                "gewicht": gewicht_final,
                "bouwjaar": bouwjaar_final,
                "cataloguswaarde": catalogus_final,
                "dagwaarde": "",
                "bpm": "",
                "meldcode": "-",
                "premie_maand": premie_final,
                "svj_override": r["svj_override"],
                "waarde_al_in_context": "1",
            },
            filename_base=_offer_pdf_filename_base(r["klantnaam"] or "", klant_type, offer_no),
        )

    else:
        kenteken_lookup = _kenteken_lookup_value(kenteken_db)
        kenteken_display = _format_nl_kenteken(kenteken_db)

        vinfo = get_vehicle_info(
            kenteken=kenteken_lookup,
            merk=r["merk"] or "",
            model=r["model"] or "",
            db_path=LOCAL_SQLITE_DB_PATH,
        )

        db_voertuig_type = str(r["voertuig_type"] or "").strip()
        vinfo_voertuig_type = getattr(vinfo, "voertuig_type", "") or ""
        voertuig_type = (vinfo_voertuig_type or db_voertuig_type or "personenauto").strip().lower()

        bj_int = _parse_int(getattr(vinfo, "bouwjaar", "") or "")
        dekking = bepaal_dekking(bj_int)

        meldcode_rolls = ""
        voertuig_type_rolls = ""
        if kenteken_lookup:
            try:
                meldcode_rolls, voertuig_type_rolls = get_meldcode_en_type(kenteken_lookup)
            except Exception:
                meldcode_rolls, voertuig_type_rolls = "", ""

        if voertuig_type_rolls and (not vinfo_voertuig_type) and (not db_voertuig_type):
            voertuig_type = voertuig_type_rolls.strip().lower()

        meldcode_final = (meldcode_rolls or getattr(vinfo, "meldcode", "") or "—").strip()

        dekking_final = _compose_dekking(
            dekking,
            r["dekking_override"] if "dekking_override" in r.keys() else None,
            r["extra_svi"] if "extra_svi" in r.keys() else 0,
            r["extra_rb"] if "extra_rb" in r.keys() else 0,
        )

        auto_str = " ".join([x for x in [vinfo.merk, vinfo.model] if x]).strip()

        offer_pdf_path = generate_offer_pdf(
            out_base_dir="data/offers",
            dt=now,
            offer_no=offer_no,
            klant={
                "naam": r["klantnaam"] or "",
                "adres": r["adres"] or "",
                "postcode": r["postcode"] or "",
                "plaats": r["plaats"] or "",
                "telefoon": r["telefoon"] or "",
                "email": email or "",
                "klant_type": klant_type,
            },
            voertuig={
                "auto": auto_str,
                "kenteken": kenteken_display,
                "brandstof": getattr(vinfo, "brandstof", "") or "",
                "voertuig_type": voertuig_type,
                "merk": getattr(vinfo, "merk", "") or (r["merk"] or ""),
                "model": getattr(vinfo, "model", "") or (r["model"] or ""),
            },
            offer={
                "regio": r["regio"] if r["regio"] is not None else "",
                "dekking": dekking_final,
                "dekking_override": r["dekking_override"] or "",
                "extra_svi": r["extra_svi"],
                "extra_rb": r["extra_rb"],
                "gewicht": getattr(vinfo, "ledig_gewicht", "") or "",
                "bouwjaar": getattr(vinfo, "bouwjaar", "") or "",
                "cataloguswaarde": getattr(vinfo, "cataloguswaarde", "") or "",
                "dagwaarde": getattr(vinfo, "dagwaarde", "") or "",
                "bpm": getattr(vinfo, "bpm", "") or "",
                "meldcode": meldcode_final,
                "is_schatting": "1" if getattr(vinfo, "is_schatting", False) else "",
                "schatting_toelichting": getattr(vinfo, "schatting_toelichting", "") or "",
                "premie_maand": r["maandpremie"],
                "svj_override": r["svj_override"],
            },
            filename_base=_offer_pdf_filename_base(r["klantnaam"] or "", klant_type, offer_no),
        )

    if email:
        template = _choose_mail_template(
            klant_type=klant_type,
            is_bestaande_klant=is_bestaande_klant,
            revision_no=revision_no,
            mail_template_type=mail_template_type,
            tpl_def=tpl_def,
            tpl_pro=tpl_pro,
            tpl_bestaand_part=tpl_bestaand_part,
            tpl_bestaand_zak=tpl_bestaand_zak,
            tpl_aangepast=tpl_aangepast,
        )

        if klant_type == "zakelijk":
            aanhef = "heer/mevrouw"
            achternaam = ""
            aanhefregel = "Geachte heer/mevrouw,"
        else:
            aanhef, achternaam = guess_aanhef_en_achternaam(r["klantnaam"] or "")
            aanhefregel = f"Geachte {aanhef} {achternaam}," if achternaam else f"Geachte {aanhef},"

        auto_show = " ".join(
            [x for x in [(r["merk"] or "").strip(), (r["model"] or "").strip()] if x]
        ).strip() or "auto"

        base_url = MICROSOFT_REDIRECT_URI.replace("/auth/microsoft/callback", "").rstrip("/")
        logo_url = f"{base_url}/static/logo_klaasvis.png"

        body = render_mail_template(
            template,
            {
                "aanhefregel": aanhefregel,
                "aanhef": aanhef,
                "achternaam": achternaam,
                "auto": auto_show,
                "offerte_nummer": offer_no,
                "aanvraag_link": f"{AANVRAAG_LINK}?offerte={offer_no}",
                "revision_no": revision_no,
                "revision_of": r["revision_of"] or "",
                "svj_override": r["svj_override"] if r["svj_override"] is not None else "",
                "heeft_svj_override": r["svj_override"] is not None,
                "logo_url": logo_url,
            },
        )

        subject = _subject_for_offer(klant_type, revision_no, offer_no)


        graph_info = None
        graph_error = ""

        try:
            graph_info = create_outlook_draft_with_attachment(
                conn=conn,
                to_addr=email,
                subject=subject,
                body_html=body,
                pdf_path=offer_pdf_path,
            )
        except Exception as e:
            graph_error = str(e)

        if graph_info:
            _execute_retry(
                conn,
                """
                UPDATE offers
                SET offer_pdf_path = %s,
                    eml_path = NULL,
                    post_letter_path = NULL,
                    delivery_method = 'email',
                    delivery_status = 'outlook_concept_klaar',
                    updated_by = %s,
                    updated_at = %s
                WHERE TRIM(offer_no) = TRIM(%s)
                """,
                (
                    offer_pdf_path,
                    current_user_display(),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    offer_no,
                ),
            )

            return {
                "kind": "email",
                "pdf": offer_pdf_path,
                "msg": None,
                "graph": graph_info,
            }

        msg_path = write_msg_outlook(
            out_base_dir="data/outbox",
            dt=now,
            offer_no=offer_no,
            to_addr=email,
            subject=subject,
            body_text=body,
            pdf_path=offer_pdf_path,
        )

        _execute_retry(
            conn,
            """
            UPDATE offers
            SET offer_pdf_path = %s,
                eml_path = %s,
                post_letter_path = NULL,
                delivery_method = 'email',
                delivery_status = 'email_klaar',
                updated_by = %s,
                updated_at = %s
            WHERE TRIM(offer_no) = TRIM(%s)
            """,
            (
                offer_pdf_path,
                msg_path,
                current_user_display(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                offer_no,
            ),
        )

        if graph_error:
            print("Outlook concept maken mislukt, MSG fallback gebruikt:", graph_error)

        return {
            "kind": "email",
            "pdf": offer_pdf_path,
            "msg": msg_path,
        }

    auto_show = " ".join(
        [x for x in [(r["merk"] or "").strip(), (r["model"] or "").strip()] if x]
    ).strip() or "auto"

    post_letter_path = generate_post_letter_pdf(
        out_base_dir="data/post",
        dt=now,
        offer_no=offer_no,
        klantnaam=r["klantnaam"] or "",
        adres=r["adres"] or "",
        postcode=r["postcode"] or "",
        plaats=r["plaats"] or "",
        auto=auto_show,
        behandeld_door="Dirk Slootweg",
    )

    post_package_path = combine_post_package_pdf(
        post_letter_path=post_letter_path,
        offer_pdf_path=offer_pdf_path,
        offer_no=offer_no,
        klantnaam=r["klantnaam"] or "",
    )

    _execute_retry(
        conn,
        """
        UPDATE offers
        SET offer_pdf_path = %s,
            post_letter_path = %s,
            eml_path = NULL,
            delivery_method = 'post',
            delivery_status = 'post_klaar',
            updated_by = %s,
            updated_at = %s
        WHERE TRIM(offer_no) = TRIM(%s)
        """,
        (
            offer_pdf_path,
            post_package_path,
            current_user_display(),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            offer_no,
        ),
    )

    return {
        "kind": "post",
        "pdf": offer_pdf_path,
        "post": post_package_path,
    }



@app.post("/export-last-batch")
@login_required
def export_last_batch():
    ensure_db()
    now = datetime.now()
    batch_id = get_last_batch_id()

    session.pop("_flashes", None)
    session.modified = True

    if not batch_id:
        flash("Geen batch gevonden om te exporteren.", "error")
        return redirect(url_for("dashboard"))

    processed = 0
    mails = 0
    posts = 0
    errors = 0
    msg_rows = []
    post_rows = []

    with connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM offers
            WHERE is_blocked = 0
              AND batch_id = %s
            ORDER BY
                CASE
                    WHEN COALESCE(offer_pdf_path,'')='' AND COALESCE(eml_path,'')='' AND COALESCE(post_letter_path,'')=''
                    THEN 0 ELSE 1
                END ASC,
                CASE WHEN lower(COALESCE(klant_type,''))='zakelijk' THEN 1 ELSE 0 END ASC,
                COALESCE(regio, 999) ASC,
                created_at DESC,
                offer_no DESC
            """,
            (batch_id,),
        ).fetchall()

        for r in rows:
            try:
                info = _build_pdf_and_delivery(conn, r, now)
                processed += 1

                if info["kind"] == "email":
                    mails += 1
                    msg_rows.append(
                        {
                            "offer_no": r["offer_no"],
                            "klantnaam": r["klantnaam"] or "",
                            "email": (r["email"] or "").strip(),
                            "msg_path": info.get("msg"),
                            "pdf_path": info["pdf"],
                        }
                    )
                else:
                    posts += 1
                    post_rows.append(
                        {
                            "offer_no": r["offer_no"],
                            "klantnaam": r["klantnaam"] or "",
                            "post_letter_path": info["post"],
                        }
                    )
            except Exception as e:
                errors += 1
                print(f"Export fout bij {r['offer_no']}:", repr(e))
                continue

        conn.commit()

    if errors:
        flash(f"Export deels gelukt: {processed} verwerkt, {errors} fout(en). Zie Render logs.", "error")
    else:
        flash(f"Export gelukt: {processed} verwerkt.", "ok")

    return render_template(
        "batch_done_msg.html",
        batch_id=batch_id,
        processed=processed,
        mails=mails,
        posts=posts,
        msgs=msg_rows,
        post_rows=post_rows,
    )

@app.post("/offer/<offer_no>/export-one")
@login_required
def export_one_offer(offer_no: str):
    ensure_db()
    now = datetime.now()
    next_url = request.form.get("next") or url_for("offers")

    session.pop("_flashes", None)
    session.modified = True

    try:
        with connect() as conn:
            r = conn.execute(
                "SELECT * FROM offers WHERE TRIM(offer_no) = TRIM(%s)",
                (offer_no,),
            ).fetchone()

            if not r:
                flash("Offerte niet gevonden.", "error")
                return redirect(next_url)

            if int(r["is_blocked"] or 0) == 1:
                flash("Deze offerte is geblokkeerd en kan niet geëxporteerd worden.", "error")
                return redirect(next_url)

            info = _build_pdf_and_delivery(conn, r, now)
            conn.commit()

        flash(f"Offerte geëxporteerd ({info['kind']}): {offer_no}", "ok")

    except Exception as e:
        print("EXPORT FOUT:", repr(e))
        flash(f"Export mislukt voor {offer_no}: {type(e).__name__}: {e}", "error")

    return redirect(next_url)

@app.get("/offer/<offer_no>/preview-pdf")
@login_required
def preview_offer_pdf(offer_no: str):
    ensure_db()
    now = datetime.now()

    with connect() as conn:
        r = conn.execute(
            "SELECT * FROM offers WHERE TRIM(offer_no) = TRIM(%s)",
            (offer_no,),
        ).fetchone()

        if not r:
            flash("Offerte niet gevonden.", "error")
            return redirect(url_for("offers"))

        info = _build_pdf_and_delivery(conn, r, now)
        conn.commit()

    abs_path = (PROJECT_ROOT / info["pdf"]).resolve()
    if not abs_path.exists():
        flash("PDF kon niet worden gevonden.", "error")
        return redirect(url_for("offers"))

    return send_file(abs_path, as_attachment=False)


@app.get("/file")
@login_required
def get_file():
    rel = request.args.get("path", "").strip()

    if not rel:
        flash("Geen bestand opgegeven.", "error")
        return redirect(url_for("offers"))

    abs_path = (PROJECT_ROOT / rel).resolve()

    if PROJECT_ROOT not in abs_path.parents and abs_path != PROJECT_ROOT:
        flash("Onveilig pad geweigerd.", "error")
        return redirect(url_for("offers"))

    if not abs_path.exists():
        flash("Bestand bestaat niet.", "error")
        return redirect(url_for("offers"))

    return send_file(abs_path, as_attachment=False)


@app.route("/browse/<kind>")
@login_required
def browse(kind: str):
    ensure_db()

    mapping = {
        "offers": PROJECT_ROOT / "data" / "offers",
        "outbox": PROJECT_ROOT / "data" / "outbox",
        "post": PROJECT_ROOT / "data" / "post",
        "inbox": PROJECT_ROOT / "data" / "inbox",
    }

    base = mapping.get(kind)
    if not base:
        flash("Onbekende map.", "error")
        return redirect(url_for("dashboard"))

    base.mkdir(parents=True, exist_ok=True)

    items = []
    for p in sorted(base.rglob("*"), key=lambda x: str(x).lower()):
        if p.is_dir():
            continue

        st = p.stat()
        items.append(
            {
                "name": p.name,
                "rel": safe_relpath(p),
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
            }
        )

    return render_template("browse.html", kind=kind, base=str(base), items=items)


if __name__ == "__main__":
    ensure_db()
    app.run(debug=True)
