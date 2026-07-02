from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import re
import pandas as pd

from db import connect, ensure_offer_counter, next_offer_no
from rules import bepaal_regio, bepaal_dekking, benodigde_svj
from denylist import load_denylist_docx
from mailgen import normalize_person_name


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


def _first_nonempty(*values) -> str:
    for value in values:
        s = _safe_str(value)
        if s:
            return s
    return ""


def _compose_address(street: str, house_no: str, house_addition: str) -> str:
    parts = [_safe_str(street), clean_house_number(house_no), _safe_str(house_addition)]
    return " ".join(part for part in parts if part).strip()


def clean_house_number(value) -> str:
    s = _safe_str(value)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    try:
        number = float(s)
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return s


def clean_phone(value) -> str:
    s = _safe_str(value)
    if not s:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    s = re.sub(r"[\s\-\(\)]", "", s)
    if s.startswith("+31"):
        s = "0" + s[3:]
    elif s.startswith("0031"):
        s = "0" + s[4:]

    digits = re.sub(r"\D", "", s)
    if len(digits) == 9 and not digits.startswith("0"):
        digits = "0" + digits
    return digits


def _row_has_import_data(values: list[str]) -> bool:
    return any(_safe_str(value) for value in values)


def _normalize_header(header: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", _safe_str(header).lower())


NEW_EXCEL_REQUIRED_HEADERS = {"relatie", "datumverkoop"}
CONVERTED_ORDERBOEK_REQUIRED_HEADERS = {
    "relatie",
    "kenteken",
    "merk",
    "autoomschrijving",
    "chassisnummer",
    "relatiegeslacht",
}
OLD_EXCEL_REQUIRED_HEADERS = {"klantnaam", "orderdatum"}


def _normalized_header_set(values) -> set[str]:
    return {_normalize_header(v) for v in values if _safe_str(v)}


def _strip_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_safe_str(c) for c in df.columns]
    return df


def _read_excel_with_detected_header(excel_path: str) -> pd.DataFrame:
    preview = pd.read_excel(excel_path, header=None, nrows=12)
    for idx, row in preview.iterrows():
        normalized = _normalized_header_set(row.tolist())
        if (
            NEW_EXCEL_REQUIRED_HEADERS.issubset(normalized)
            or CONVERTED_ORDERBOEK_REQUIRED_HEADERS.issubset(normalized)
            or OLD_EXCEL_REQUIRED_HEADERS.issubset(normalized)
        ):
            return _strip_dataframe_columns(pd.read_excel(excel_path, header=int(idx)))
    return _strip_dataframe_columns(pd.read_excel(excel_path))


def _detect_import_format(df: pd.DataFrame) -> str:
    normalized = _normalized_header_set(df.columns)
    if NEW_EXCEL_REQUIRED_HEADERS.issubset(normalized):
        return "nieuw"
    if CONVERTED_ORDERBOEK_REQUIRED_HEADERS.issubset(normalized):
        return "nieuw"
    if OLD_EXCEL_REQUIRED_HEADERS.issubset(normalized):
        return "oud"
    raise ValueError("Onbekend Excel-formaat. Controleer kolomnamen.")


def _normalize_relatie_geslacht_code(raw: str) -> str:
    s = _safe_str(raw).strip().upper()
    if s in {"M", "MAN", "HEER", "H", "DHR", "DE HEER"}:
        return "M"
    if s in {"V", "VROUW", "MEVROUW", "MEVR", "MW"}:
        return "V"
    if s in {"Z", "ZAKELIJK", "BEDRIJF", "ORG", "ORGANISATIE"}:
        return "Z"
    if s in {"O", "ONBEKEND", "UNKNOWN"}:
        return "O"
    return "O"


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
    code = _normalize_relatie_geslacht_code(raw)
    if code == "Z":
        return "zakelijk"
    if code in {"M", "V", "O"}:
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


def _format_created_at(value, fallback: str) -> str:
    if value is None or _safe_str(value) == "":
        return fallback
    try:
        ts = pd.to_datetime(value)
        if pd.isna(ts):
            return fallback
        return ts.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return fallback


def _normalize_occasion(raw: str) -> str:
    s = _safe_str(raw).strip().lower()
    if not s:
        return ""
    if "gebruik" in s or s in {"ja", "j", "occasion", "used"}:
        return "gebruikt"
    if "nieuw" in s or s in {"nee", "n", "new"}:
        return "nieuw"
    return _safe_str(raw)


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


def import_excel(
    excel_path: str,
    denylist_path: str | None = None,
    return_batch: bool = False,
) -> int | tuple[int, str]:
    now = datetime.now()
    today = now.date()
    batch = _batch_id(now)
    created_at = now.strftime("%Y-%m-%d %H:%M:%S")

    deny_path = Path(denylist_path) if denylist_path else DEFAULT_DENYLIST
    deny_entries = load_denylist_docx(str(deny_path)) if deny_path.exists() else []

    df = _read_excel_with_detected_header(excel_path)
    import_format = _detect_import_format(df)
    normalized_columns = {_normalize_header(c): c for c in df.columns}

    def col(*names):
        for n in names:
            if n in df.columns:
                return n

            normalized = _normalize_header(n)
            if normalized in normalized_columns:
                return normalized_columns[normalized]

        return None

    c_orderdatum = col("Orderdatum", "Datum verkoop")
    c_verkoper = col("Verkoper")
    c_klantnaam = col("Relatie") if import_format == "nieuw" else col("Klantnaam", "Naam", "Relatienaam")
    c_adres = col("Adres", "Straat", "Straatnaam")
    c_straat = col("Relatie straat", "Straat", "Straatnaam")
    c_huisnr = col("Relatie huisnr.", "Relatie huisnummer", "Huisnummer", "Huisnr.")
    c_huisnr_toev = col("Relatie huisnr. toev.", "Relatie huisnr toev", "Toevoeging")
    c_postcode = col("Postcode", "Post code", "Relatie postcode")
    c_plaats = col("Woonplaats", "Plaats", "Relatie plaats")
    c_tel = col("Telefoonnummer", "Telefoon", "Mobiel", "Relatie tel prive")
    c_tel_prive = col("Relatie tel prive", "Telefoon prive")
    c_tel_mobiel = col("Relatie mobiel", "Mobiel")
    c_email = col("E-mail", "Email", "Mail", "Relatie email")
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
    c_model = col("Afleveringmodel") if import_format == "nieuw" else col("Model auto", "Model")
    c_model_fallback = col("Autoomschrijving")
    c_type = col("Autoomschrijving") if import_format == "nieuw" else col("Type model", "Type")
    c_occasion = col("Occasion", "Nieuw Gebruikt")

    c_klanttype = col("Klanttype", "Type klant", "Categorie")
    c_relatie_geslacht = col("Relatie geslacht", "Relatiegeslacht", "Geslacht")
    c_voertuigtype = col("Voertuigtype", "Soort voertuig")
    c_bouwjaar = col("Bouwjaar", "Jaar")

    missing_required = []
    for label, value in [
        ("Relatie/klantnaam", c_klantnaam),
        ("Relatie postcode/postcode", c_postcode),
        ("Relatie plaats/plaats", c_plaats),
        ("Kenteken", c_kenteken),
        ("Merk", c_merk),
        ("Autoomschrijving/model", c_model or c_model_fallback),
    ]:
        if not value:
            missing_required.append(label)

    if missing_required:
        available = ", ".join(_safe_str(c) for c in df.columns)
        raise ValueError(
            "Importbestand wordt niet herkend. Ontbrekende kolommen: "
            + ", ".join(missing_required)
            + f". Gevonden kolommen: {available}"
        )

    rows: list[dict] = []

    invalid_rows: list[str] = []

    for idx, row in df.iterrows():
        klantnaam = _safe_str(row.get(c_klantnaam)) if c_klantnaam else ""
        row_created_at = _format_created_at(row.get(c_orderdatum) if c_orderdatum else None, created_at)
        verkoper = _safe_str(row.get(c_verkoper)) if c_verkoper else ""
        straat = _safe_str(row.get(c_straat)) if c_straat else ""
        huisnummer = clean_house_number(row.get(c_huisnr)) if c_huisnr else ""
        huisnummer_toevoeging = _safe_str(row.get(c_huisnr_toev)) if c_huisnr_toev else ""
        adres = _safe_str(row.get(c_adres)) if c_adres else ""
        if not adres:
            adres = _compose_address(straat, huisnummer, huisnummer_toevoeging)
        postcode = _safe_str(row.get(c_postcode)) if c_postcode else ""
        plaats = _safe_str(row.get(c_plaats)) if c_plaats else ""
        telefoon = _first_nonempty(
            clean_phone(row.get(c_tel_mobiel)) if c_tel_mobiel else "",
            clean_phone(row.get(c_tel_prive)) if c_tel_prive else "",
            clean_phone(row.get(c_tel)) if c_tel else "",
        )
        email = _safe_str(row.get(c_email)) if c_email else ""

        kenteken_raw = _safe_str(row.get(c_kenteken)) if c_kenteken else ""
        kenteken_norm = _norm_kenteken(kenteken_raw)
        kenteken = kenteken_norm if _is_valid_kenteken(kenteken_norm) else ""
        chassisnummer = _norm_chassisnummer(row.get(c_chassis)) if c_chassis else ""
        meldcode = _meldcode_from_chassis(chassisnummer)

        merk = _safe_str(row.get(c_merk)) if c_merk else ""
        model = _safe_str(row.get(c_model)) if c_model else ""
        if not model and c_model_fallback:
            model = _safe_str(row.get(c_model_fallback))
        type_model = _safe_str(row.get(c_type)) if c_type else ""
        occasion = _normalize_occasion(row.get(c_occasion) if c_occasion else "")

        if not _row_has_import_data([klantnaam, adres, postcode, plaats, email, telefoon, kenteken, merk, model, chassisnummer]):
            continue

        raw_kt = ""
        if c_klanttype:
            raw_kt = _safe_str(row.get(c_klanttype))
        else:
            try:
                if len(df.columns) > 5 and not c_relatie_geslacht:
                    raw_kt = _safe_str(row.iloc[5])
            except Exception:
                raw_kt = ""

        relatie_geslacht = _normalize_relatie_geslacht_code(row.get(c_relatie_geslacht) if c_relatie_geslacht else "")
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

        missing_values = []
        if not klantnaam:
            missing_values.append("relatie/klantnaam")
        if not adres:
            missing_values.append("adres")
        if not postcode:
            missing_values.append("postcode")
        if not plaats:
            missing_values.append("plaats")
        if regio is None:
            missing_values.append("regio uit postcode")
        if not merk:
            missing_values.append("merk")
        if not model:
            missing_values.append("autoomschrijving/model")

        if missing_values:
            excel_row_no = idx + 2
            invalid_rows.append(
                f"rij {excel_row_no}: {', '.join(missing_values)}"
            )
            continue

        rows.append(
            {
                "klantnaam": klantnaam,
                "created_at": row_created_at,
                "verkoper": verkoper,
                "relatie_geslacht": relatie_geslacht,
                "straat": straat,
                "huisnummer": huisnummer,
                "huisnummer_toevoeging": huisnummer_toevoeging,
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
                "occasion": occasion,
                "klant_type": klant_type,
                "voertuig_type": voertuig_type,
                "bouwjaar": bouwjaar_val,
                "regio": regio,
            }
        )

    if invalid_rows:
        raise ValueError(
            "Import afgebroken: niet alle verplichte gegevens konden uit het Excelbestand worden gelezen. "
            + "; ".join(invalid_rows[:10])
        )

    if not rows:
        raise ValueError("Import afgebroken: er zijn geen bruikbare leadregels gevonden in het Excelbestand.")

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
            offer_created_at = r["created_at"]
            verkoper = r["verkoper"]
            relatie_geslacht = r["relatie_geslacht"]
            straat = r["straat"]
            huisnummer = r["huisnummer"]
            huisnummer_toevoeging = r["huisnummer_toevoeging"]
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
            occasion = r["occasion"]
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
                    klantnaam, relatie_geslacht, straat, huisnummer, huisnummer_toevoeging,
                    adres, postcode, plaats, telefoon, email,
                    kenteken, chassisnummer, meldcode, merk, model, type_model,
                    occasion, klant_type, voertuig_type,
                    bouwjaar, regio, dekking, benodigde_svj,
                    delivery_method, delivery_status,
                    is_blocked, block_reason, block_note,
                    call_status, decision_status, follow_up_due_at,
                    mail_template_type
                ) VALUES (
                    %s, %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s,
                    %s, %s, %s,
                    %s, %s, %s, %s,
                    %s, %s,
                    %s, %s, %s,
                    'open', 'open', %s,
                    'auto'
                )
                """,
                (
                    offer_no,
                    offer_created_at,
                    mk,
                    batch,
                    klantnaam,
                    relatie_geslacht,
                    straat,
                    huisnummer,
                    huisnummer_toevoeging,
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
                    occasion,
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
            normalized = normalize_person_name(klantnaam)
            print(
                "IMPORT LEAD:",
                {
                    "offer_no": offer_no,
                    "klantnaam_origineel": klantnaam,
                    "klantnaam_genormaliseerd": normalized.get("display", klantnaam),
                    "relatie_geslacht": relatie_geslacht,
                    "chassisnummer": chassisnummer,
                    "formaat": import_format,
                },
            )

        conn.commit()

    if return_batch:
        return processed, batch

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

