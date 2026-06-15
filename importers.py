from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import pandas as pd

from db import connect, ensure_offer_counter, next_offer_no
from rules import bepaal_regio, bepaal_dekking, benodigde_svj
from denylist import load_denylist_docx


DEFAULT_DENYLIST = Path("data/denylist/Lijstje GEEN leads.docx")


def _safe_str(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s.lower() in {"nan", "none"}:
        return ""
    return s


def _month_key(dt: date) -> str:
    return dt.strftime("%Y-%m")


def _batch_id(dt: datetime) -> str:
    return dt.strftime("%Y%m%d-%H%M%S")


def _followup_due(dt: date) -> str:
    return (dt + timedelta(days=7)).isoformat()


def _normalize_klant_type(raw: str) -> str:
    s = _safe_str(raw).lower()
    if not s:
        return "particulier"

    code = s[:2]

    if "zakel" in s or code == "14":
        return "zakelijk"

    if "prospect" in s:
        return "prospect"

    if "lead" in s:
        return "particulier"

    if "particulier" in s or "berijder" in s or code == "10":
        return "particulier"

    if code == "00":
        return "particulier"

    return "particulier"


def _normalize_relatie_geslacht(raw: str) -> str:
    s = _safe_str(raw).strip().lower()
    if not s:
        return ""

    if s in {"org", "organisatie", "bedrijf", "zakelijk"}:
        return "zakelijk"

    if s in {"o", "onbekend", "unknown"}:
        return ""

    if s in {"m", "v", "man", "vrouw", "particulier"}:
        return "particulier"

    return ""


def _looks_business_name(name: str) -> bool:
    s = f" {_safe_str(name).lower()} "
    markers = (
        " b.v", " bv ", " v.o.f", " vof ", " n.v", " nv ",
        " stichting ", " vereniging ", " holding ", " beheer ",
        " bedrijf ", " bedrijfswagens ", " installatiebedrijf ",
        " loodgieter", " loonbedrijf ", " aannemingsbedrijf ",
        " bouw ", " bouwservice ", " montage ", " techniek ",
        " elektrotechniek ", " schoonmaak ", " carwash ",
        " totaal bouw ", " geoservices ", " adviesgroep ",
    )
    return any(marker in s for marker in markers)


def _infer_klant_type(klantnaam: str, raw_klant_type: str, relatie_geslacht: str) -> str:
    raw = _safe_str(raw_klant_type)
    raw_lower = raw.lower()

    if raw:
        normalized = _normalize_klant_type(raw)
        if (
            normalized != "particulier"
            or "particulier" in raw_lower
            or "berijder" in raw_lower
            or raw_lower[:2] in {"00", "10"}
        ):
            return normalized

    from_gender = _normalize_relatie_geslacht(relatie_geslacht)
    if from_gender:
        return from_gender

    if _looks_business_name(klantnaam):
        return "zakelijk"

    return "particulier"


def _klant_type_sort_key(klant_type: str) -> int:
    kt = _safe_str(klant_type).lower()
    if kt == "particulier":
        return 0
    if kt == "zakelijk":
        return 1
    if kt == "prospect":
        return 2
    return 9


def _regio_sort_key(regio) -> int:
    try:
        return int(str(regio).strip())
    except Exception:
        return 999


def _norm_kenteken(k: str) -> str:
    return "".join(ch for ch in _safe_str(k).upper() if ch.isalnum())


def _norm_chassisnummer(chassis: str) -> str:
    return "".join(ch for ch in _safe_str(chassis).upper() if ch.isalnum())


def _meldcode_from_chassis(chassis: str) -> str:
    normalized = _norm_chassisnummer(chassis)
    return normalized[-4:] if len(normalized) >= 4 else ""


def _is_valid_kenteken(k: str) -> bool:
    kk = _norm_kenteken(k)
    if not kk:
        return False
    if len(kk) != 6:
        return False
    if not kk.isalnum():
        return False
    if kk == "000000":
        return False
    if len(set(kk)) == 1:
        return False
    return True


def import_excel(excel_path: str, denylist_path: str | None = None) -> int:
    now = datetime.now()
    today = now.date()
    batch = _batch_id(now)
    created_at = now.strftime("%Y-%m-%d %H:%M:%S")

    deny_path = Path(denylist_path) if denylist_path else DEFAULT_DENYLIST
    deny_entries = load_denylist_docx(str(deny_path)) if deny_path.exists() else []

    df = pd.read_excel(excel_path)

    def col(*names):
        for n in names:
            if n in df.columns:
                return n
        return None

    c_klantnaam = col("Klantnaam", "Naam", "Relatienaam")
    c_adres = col("Adres", "Straat", "Straatnaam")
    c_postcode = col("Postcode", "Post code")
    c_plaats = col("Woonplaats", "Plaats")
    c_tel = col("Telefoonnummer", "Telefoon", "Mobiel")
    c_email = col("E-mail", "Email", "Mail")
    c_kenteken = col("Kenteken", "LicensePlate")
    c_chassis = col(
        "Chassisnummer",
        "Chassis nummer",
        "Chassis",
        "VIN",
        "VIN nummer",
        "Voertuigidentificatienummer",
    )
    c_merk = col("Merk auto", "Merk")
    c_model = col("Model auto", "Model")
    c_type = col("Type model", "Type")

    c_klanttype = col("Klanttype", "Type klant", "Categorie")
    c_relatie_geslacht = col("Relatie geslacht", "Relatiegeslacht", "Geslacht")
    c_voertuigtype = col("Voertuigtype", "Soort voertuig")
    c_bouwjaar = col("Bouwjaar", "Jaar")

    rows: list[dict] = []

    for _, row in df.iterrows():
        klantnaam = _safe_str(row.get(c_klantnaam)) if c_klantnaam else ""
        adres = _safe_str(row.get(c_adres)) if c_adres else ""
        postcode = _safe_str(row.get(c_postcode)) if c_postcode else ""
        plaats = _safe_str(row.get(c_plaats)) if c_plaats else ""
        telefoon = _safe_str(row.get(c_tel)) if c_tel else ""
        email = _safe_str(row.get(c_email)) if c_email else ""

        kenteken_raw = _safe_str(row.get(c_kenteken)) if c_kenteken else ""
        kenteken_norm = _norm_kenteken(kenteken_raw)
        kenteken = kenteken_norm if _is_valid_kenteken(kenteken_norm) else ""
        chassisnummer = _norm_chassisnummer(row.get(c_chassis)) if c_chassis else ""
        meldcode = _meldcode_from_chassis(chassisnummer)

        merk = _safe_str(row.get(c_merk)) if c_merk else ""
        model = _safe_str(row.get(c_model)) if c_model else ""
        type_model = _safe_str(row.get(c_type)) if c_type else ""

        raw_kt = ""
        if c_klanttype:
            raw_kt = _safe_str(row.get(c_klanttype))
        else:
            try:
                if len(df.columns) > 5 and not c_relatie_geslacht:
                    raw_kt = _safe_str(row.iloc[5])
            except Exception:
                raw_kt = ""

        relatie_geslacht = _safe_str(row.get(c_relatie_geslacht)) if c_relatie_geslacht else ""
        klant_type = _infer_klant_type(klantnaam, raw_kt, relatie_geslacht)

        voertuig_type = (_safe_str(row.get(c_voertuigtype)) if c_voertuigtype else "personenauto").lower().strip()
        if voertuig_type not in {"personenauto", "bestelauto"}:
            voertuig_type = "bestelauto" if "bestel" in voertuig_type else "personenauto"

        bouwjaar_val = None
        if c_bouwjaar:
            bj = row.get(c_bouwjaar)
            try:
                bouwjaar_val = int(bj) if str(bj).strip() != "" else None
            except Exception:
                bouwjaar_val = None

        regio = bepaal_regio(postcode)

        rows.append(
            {
                "klantnaam": klantnaam,
                "adres": adres,
                "postcode": postcode,
                "plaats": plaats,
                "telefoon": telefoon,
                "email": email,
                "kenteken": kenteken,
                "chassisnummer": chassisnummer,
                "meldcode": meldcode,
                "merk": merk,
                "model": model,
                "type_model": type_model,
                "klant_type": klant_type,
                "voertuig_type": voertuig_type,
                "bouwjaar": bouwjaar_val,
                "regio": regio,
            }
        )

    rows.sort(
        key=lambda r: (
            _klant_type_sort_key(r.get("klant_type", "")),
            _regio_sort_key(r.get("regio")),
            _safe_str(r.get("klantnaam", "")).lower(),
            _safe_str(r.get("postcode", "")).lower(),
        )
    )

    processed = 0

    with connect() as conn:
        ensure_offer_counter(conn)

        for r in rows:
            processed += 1

            klantnaam = r["klantnaam"]
            adres = r["adres"]
            postcode = r["postcode"]
            plaats = r["plaats"]
            telefoon = r["telefoon"]
            email = r["email"]
            kenteken = r["kenteken"]
            chassisnummer = r["chassisnummer"]
            meldcode = r["meldcode"]
            merk = r["merk"]
            model = r["model"]
            type_model = r["type_model"]
            klant_type = r["klant_type"]
            voertuig_type = r["voertuig_type"]
            bouwjaar_val = r["bouwjaar"]
            regio = r["regio"]

            is_blocked = 0
            block_reason = None
            block_note = None

            for d in deny_entries:
                if d and klantnaam and d.lower() in klantnaam.lower():
                    is_blocked = 1
                    block_reason = "denylist"
                    block_note = f"Matched: {d}"
                    break

            dekking = bepaal_dekking(bouwjaar_val)
            ben_svj = benodigde_svj(klant_type, voertuig_type, regio)

            mk = _month_key(today)
            offer_no = next_offer_no(conn, mk)

            delivery_method = "email" if email else "post"
            delivery_status = "nieuw"

            conn.execute(
                """
                INSERT INTO offers (
                    offer_no, created_at, month_key, batch_id,
                    klantnaam, adres, postcode, plaats, telefoon, email,
                    kenteken, chassisnummer, meldcode, merk, model, type_model,
                    klant_type, voertuig_type,
                    bouwjaar, regio, dekking, benodigde_svj,
                    delivery_method, delivery_status,
                    is_blocked, block_reason, block_note,
                    call_status, decision_status, follow_up_due_at,
                    mail_template_type
                ) VALUES (
                    %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s,
                    %s, %s,
                    %s, %s, %s, %s,
                    %s, %s,
                    %s, %s, %s,
                    'open', 'open', %s,
                    'auto'
                )
                """,
                (
                    offer_no,
                    created_at,
                    mk,
                    batch,
                    klantnaam,
                    adres,
                    postcode,
                    plaats,
                    telefoon,
                    email,
                    kenteken,
                    chassisnummer,
                    meldcode,
                    merk,
                    model,
                    type_model,
                    klant_type,
                    voertuig_type,
                    bouwjaar_val,
                    regio,
                    dekking,
                    ben_svj,
                    delivery_method,
                    delivery_status,
                    is_blocked,
                    block_reason,
                    block_note,
                    _followup_due(today),
                ),
            )

        conn.commit()

    return processed


def get_last_batch_id():
    with connect() as conn:
        row = conn.execute(
            """
            SELECT batch_id
            FROM offers
            WHERE batch_id IS NOT NULL
              AND batch_id != ''
            ORDER BY created_at DESC, offer_no DESC
            LIMIT 1
            """
        ).fetchone()

    return row["batch_id"] if row else None

