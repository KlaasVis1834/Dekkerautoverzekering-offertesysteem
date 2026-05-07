from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Tuple, Dict
import sqlite3
import re

import requests


RDW_BASE = "https://opendata.rdw.nl/resource"
DATASET_VOERTUIG = "m9d7-ebf2"   # gekentekende voertuigen
DATASET_BRANDSTOF = "8ys7-d773"  # brandstof
DATASET_MASSA = "mdqe-txpd"      # massa / gewichten


@dataclass
class VoertuigInfo:
    kenteken: str = ""
    merk: str = ""
    model: str = ""
    brandstof: str = ""
    bouwjaar: str = ""
    ledig_gewicht: str = ""
    cataloguswaarde: str = ""
    dagwaarde: str = ""
    meldcode: str = "—"
    is_schatting: bool = False
    schatting_toelichting: str = ""

    # ✅ nieuw: herkennen personenauto/bestelauto
    voertuig_type: str = ""   # "personenauto" of "bestelauto"
    soort_raw: str = ""       # ruwe RDW tekst (optioneel)

    # ✅ nieuw: bpm (als RDW beschikbaar)
    bpm: str = ""


def _safe(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _norm_kenteken(k: str) -> str:
    k = _safe(k).upper()
    return re.sub(r"[^A-Z0-9]", "", k)


def _rdw_get(dataset: str, params: dict, timeout: int = 12) -> list[dict]:
    url = f"{RDW_BASE}/{dataset}.json"
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    data = r.json()
    return data if isinstance(data, list) else []


def _pick_first(row: dict, keys: list[str]) -> str:
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


def _bouwjaar_from_datestr(s: str) -> str:
    t = _safe(s)
    if not t:
        return ""
    t = t.replace("-", "")
    if len(t) >= 4 and t[:4].isdigit():
        return t[:4]
    return ""


def _title_case_merk_model(s: str) -> str:
    """
    Maakt "FORD" -> "Ford", "MERCEDES-BENZ" -> "Mercedes-Benz"
    """
    t = _safe(s)
    if not t:
        return ""
    t = t.lower()

    parts = []
    for chunk in re.split(r"\s+", t):
        if not chunk:
            continue
        subparts = chunk.split("-")
        subparts = [sp[:1].upper() + sp[1:] if sp else "" for sp in subparts]
        parts.append("-".join(subparts))
    return " ".join(parts).strip()


def _digits_only_number(s: str) -> str:
    """
    Haalt rommel weg en laat alleen een 'integer/float' string over.
    Voorbeeld: "62.463" -> "62463"
              "€ 62.463,00" -> "62463.00"
    """
    t = _safe(s)
    if not t:
        return ""

    t = t.replace("€", "").strip()

    if "." in t and "," in t:
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")

    t = re.sub(r"[^0-9.]", "", t)

    if t.count(".") > 1:
        parts = t.split(".")
        t = "".join(parts[:-1]) + "." + parts[-1]

    return t.strip()


def _to_float(v: str) -> Optional[float]:
    t = _digits_only_number(v)
    if not t:
        return None
    try:
        return float(t)
    except Exception:
        return None


def _format_kg_with_thousands(s: str) -> str:
    """
    Zorgt dat '1250' -> '1.250' (als >= 1000)
    """
    t = _safe(s)
    if not t:
        return ""
    digits = re.sub(r"[^0-9]", "", t)
    if not digits:
        return ""
    try:
        n = int(digits)
    except Exception:
        return t
    if n >= 1000:
        return f"{n:,}".replace(",", ".")
    return str(n)


def _normalize_brandstof_from_rdw(rows_b: list[dict]) -> str:
    """
    - Benzine/Elektriciteit => "Hybride (Benzine/Elektrisch)"
    - Diesel/Elektriciteit => "Hybride (Diesel/Elektrisch)"
    - Alleen Elektriciteit => "Elektrisch"
    - Spelfout 'Elecktriciteit' fixen
    """
    fuels: list[str] = []
    for rb in rows_b:
        b = _pick_first(rb, ["brandstof_omschrijving", "brandstofomschrijving", "brandstof"])
        b = _safe(b)
        if not b:
            continue
        if b.lower() == "elecktriciteit":
            b = "Elektriciteit"
        fuels.append(b)

    if not fuels:
        return ""

    seen = set()
    uniq = []
    for f in fuels:
        key = f.strip().lower()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(f.strip())

    lowset = {u.lower() for u in uniq}

    if "elektriciteit" in lowset:
        if "benzine" in lowset:
            return "Hybride (Benzine/Elektrisch)"
        if "diesel" in lowset:
            return "Hybride (Diesel/Elektrisch)"
        if "lpg" in lowset:
            return "Hybride (LPG/Elektrisch)"
        return "Elektrisch"

    first = uniq[0]
    if first.lower() == "elecktriciteit":
        first = "Elektriciteit"
    return first


def _guess_voertuig_type_from_rdw(row: dict) -> str:
    """
    Probeert te bepalen of het een personenauto of bestelauto/bedrijfsauto is.
    """
    blob = " ".join(
        [
            _safe(row.get("voertuigsoort")),
            _safe(row.get("inrichting")),
            _safe(row.get("voertuigcategorie")),
            _safe(row.get("europese_voertuigcategorie")),
        ]
    ).lower()

    if any(k in blob for k in ["bedrijfsauto", "bestelauto", "lichte bedrijfsauto", "n1"]):
        return "bestelauto"

    return "personenauto"


def _estimate_dagwaarde_from_catalogus(cataloguswaarde: str, bouwjaar: str) -> str:
    """
    Optie A: eenvoudige dagwaarde-schatting op basis van cataloguswaarde + leeftijd.
    Geeft een numerieke string terug (geschikt voor pdfgen _fmt_euro).
    """
    cv = _to_float(cataloguswaarde)
    if cv is None:
        return ""

    try:
        bj = int(str(bouwjaar).strip())
    except Exception:
        return ""

    leeftijd = datetime.now().year - bj

    if leeftijd <= 0:
        factor = 0.90
    elif leeftijd <= 1:
        factor = 0.80
    elif leeftijd <= 3:
        factor = 0.70
    elif leeftijd <= 6:
        factor = 0.55
    elif leeftijd <= 10:
        factor = 0.35
    else:
        factor = 0.20

    return str(round(cv * factor, 2))


def fetch_rdw_by_plate(kenteken: str) -> VoertuigInfo:
    plate = _norm_kenteken(kenteken)
    info = VoertuigInfo(kenteken=plate)

    if not plate:
        return info

    # 1) basis
    rows = []
    try:
        rows = _rdw_get(DATASET_VOERTUIG, {"kenteken": plate, "$limit": 1})
    except Exception:
        rows = []

    if rows:
        row = rows[0]

        info.voertuig_type = _guess_voertuig_type_from_rdw(row)
        info.soort_raw = _safe(row.get("voertuigsoort")) or _safe(row.get("inrichting")) or ""

        info.merk = _pick_first(row, ["merk"])
        info.model = _pick_first(row, ["handelsbenaming", "typeaanduiding"])
        info.bouwjaar = _bouwjaar_from_datestr(_pick_first(row, ["datum_eerste_toelating"]))

        info.cataloguswaarde = _pick_first(row, ["catalogusprijs", "cataloguswaarde"])
        info.ledig_gewicht = _pick_first(row, ["massa_ledig_voertuig", "ledig_gewicht", "massa_rijklaar"])

        # ✅ BPM (veldnaam verschilt soms; daarom meerdere opties)
        info.bpm = _pick_first(
            row,
            [
                "bpm",
                "bruto_bpm",
                "bpm_bedrag",
                "bedrag_bpm",
                "bedrag_bruto_bpm",
            ],
        )

    # 2) brandstof
    try:
        rows_b = _rdw_get(DATASET_BRANDSTOF, {"kenteken": plate, "$limit": 5})
    except Exception:
        rows_b = []

    if rows_b:
        info.brandstof = _normalize_brandstof_from_rdw(rows_b)

    # 3) massa fallback
    if not info.ledig_gewicht:
        try:
            rows_m = _rdw_get(DATASET_MASSA, {"kenteken": plate, "$limit": 1})
        except Exception:
            rows_m = []
        if rows_m:
            rm = rows_m[0]
            info.ledig_gewicht = _pick_first(rm, ["massa_ledig_voertuig", "ledig_gewicht", "massa_rijklaar"])

    # ---- normalisaties ----
    info.merk = _title_case_merk_model(info.merk)
    info.model = _title_case_merk_model(info.model)

    info.ledig_gewicht = _format_kg_with_thousands(info.ledig_gewicht)

    # geldvelden schoonmaken voor pdfgen
    info.cataloguswaarde = _digits_only_number(info.cataloguswaarde)
    info.dagwaarde = _digits_only_number(info.dagwaarde)
    info.bpm = _digits_only_number(info.bpm)

    # brandstof spellingsfix (voor het geval)
    if info.brandstof.lower() == "elecktriciteit":
        info.brandstof = "Elektriciteit"

    # ✅ dagwaarde schatten als die leeg is, maar we wél cataloguswaarde+bouwjaar hebben
    if not info.dagwaarde and info.cataloguswaarde and info.bouwjaar:
        info.dagwaarde = _estimate_dagwaarde_from_catalogus(info.cataloguswaarde, info.bouwjaar)

    return info


def estimate_from_db(
    db_path: Path,
    merk: str = "",
    model: str = "",
) -> Tuple[Dict[str, str], str]:
    merk_u = _safe(merk).upper()
    model_u = _safe(model).upper()

    if not db_path.exists():
        return {}, "Geen database gevonden voor schatting."

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT merk, model
            FROM offers
            WHERE merk IS NOT NULL AND merk != ''
              AND model IS NOT NULL AND model != ''
              AND UPPER(merk) = ?
              AND UPPER(model) = ?
            ORDER BY created_at DESC
            LIMIT 25
            """,
            (merk_u, model_u),
        ).fetchall()

        if not rows:
            rows = conn.execute(
                """
                SELECT merk, model
                FROM offers
                WHERE merk IS NOT NULL AND merk != ''
                  AND UPPER(merk) = ?
                ORDER BY created_at DESC
                LIMIT 25
                """,
                (merk_u,),
            ).fetchall()

        if not rows:
            return {}, "Geen vergelijkbare voertuigen gevonden in eerdere offertes."

        chosen = rows[0]
        result = {"merk": _safe(chosen["merk"]), "model": _safe(chosen["model"])}

        result["merk"] = _title_case_merk_model(result["merk"])
        result["model"] = _title_case_merk_model(result["model"])

        return result, "Schatting gebaseerd op eerder verwerkte, vergelijkbare voertuigen in ons systeem."
    finally:
        conn.close()


def get_vehicle_info(
    *,
    kenteken: str = "",
    merk: str = "",
    model: str = "",
    db_path: Optional[Path] = None,
) -> VoertuigInfo:
    plate = _norm_kenteken(kenteken)

    if plate:
        info = fetch_rdw_by_plate(plate)
        if not info.merk:
            info.merk = _title_case_merk_model(_safe(merk))
        if not info.model:
            info.model = _title_case_merk_model(_safe(model))
        return info

    info = VoertuigInfo(
        kenteken="",
        merk=_title_case_merk_model(_safe(merk)),
        model=_title_case_merk_model(_safe(model)),
        is_schatting=True,
    )

    if db_path:
        est, why = estimate_from_db(db_path, merk=merk, model=model)
        info.merk = est.get("merk", info.merk)
        info.model = est.get("model", info.model)
        info.schatting_toelichting = why
    else:
        info.schatting_toelichting = "Kenteken ontbreekt; gegevens zijn een schatting."

    return info