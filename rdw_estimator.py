# rdw_estimator.py
"""
Geavanceerde RDW schattingsmodule voor voertuigen zonder kenteken.

Zoekt op:
- merk
- model
- type_model (uitvoering)
- bouwjaar

Geeft terug:
- cataloguswaarde
- gewicht
- brandstof
- voertuig_type
- sample_size
- match_level
- is_schatting

Match levels:
- exact_type     -> model + uitvoering gevonden
- partial_type   -> model + gedeeltelijke uitvoering gevonden
- model_only     -> alleen model match
- manual_required -> onvoldoende data
"""

from collections import Counter
from statistics import median
import re
import requests

RDW_API_URL = "https://opendata.rdw.nl/resource/m9d7-ebf2.json"


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def _to_float(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None


def _to_int(value):
    val = _to_float(value)
    if val is None:
        return None
    return int(round(val))


def _normalize(text):
    text = (text or "").strip().lower()
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _tokens(text):
    return [t for t in _normalize(text).split() if len(t) >= 2]


def _median(values):
    nums = [v for v in values if v is not None]
    if not nums:
        return None
    return int(round(median(nums)))


def _most_common(values):
    vals = [v for v in values if v]
    if not vals:
        return ""
    return Counter(vals).most_common(1)[0][0]


def _parse_year(row):
    datum = row.get("datum_eerste_toelating") or ""
    if len(datum) >= 4 and datum[:4].isdigit():
        return int(datum[:4])
    return None


def _vehicle_type_from_rdw(voertuigsoort):
    s = (voertuigsoort or "").lower()
    if "bestel" in s:
        return "bestelauto"
    return "personenauto"


def _contains_all_tokens(text, token_list):
    normalized = _normalize(text)
    return all(token in normalized for token in token_list)


# ---------------------------------------------------------
# Filter logic
# ---------------------------------------------------------
def _filter_rows(rows, model, type_model=None, bouwjaar=None):
    model_tokens = _tokens(model)
    type_tokens = _tokens(type_model)

    if not model_tokens:
        return [], "manual_required"

    # 1. Exact model + alle type tokens
    exact_rows = []
    if type_tokens:
        for row in rows:
            handelsbenaming = row.get("handelsbenaming") or ""
            if (
                _contains_all_tokens(handelsbenaming, model_tokens)
                and _contains_all_tokens(handelsbenaming, type_tokens)
            ):
                exact_rows.append(row)

        if bouwjaar:
            exact_rows = [
                r for r in exact_rows
                if (_parse_year(r) is None or abs(_parse_year(r) - int(bouwjaar)) <= 1)
            ]

        if len(exact_rows) >= 3:
            return exact_rows, "exact_type"

    # 2. Gedeeltelijke type match
    partial_rows = []
    if type_tokens:
        for row in rows:
            handelsbenaming = row.get("handelsbenaming") or ""
            if not _contains_all_tokens(handelsbenaming, model_tokens):
                continue

            matched = sum(
                1 for token in type_tokens
                if token in _normalize(handelsbenaming)
            )

            if matched >= max(1, len(type_tokens) // 2):
                partial_rows.append(row)

        if bouwjaar:
            partial_rows = [
                r for r in partial_rows
                if (_parse_year(r) is None or abs(_parse_year(r) - int(bouwjaar)) <= 1)
            ]

        if len(partial_rows) >= 3:
            return partial_rows, "partial_type"

    # 3. Alleen model
    model_rows = []
    for row in rows:
        handelsbenaming = row.get("handelsbenaming") or ""
        if _contains_all_tokens(handelsbenaming, model_tokens):
            model_rows.append(row)

    if bouwjaar:
        model_rows = [
            r for r in model_rows
            if (_parse_year(r) is None or abs(_parse_year(r) - int(bouwjaar)) <= 1)
        ]

    if len(model_rows) >= 3:
        return model_rows, "model_only"

    return [], "manual_required"


# ---------------------------------------------------------
# Main function
# ---------------------------------------------------------
def estimate_vehicle_data_from_rdw(
    merk,
    model,
    type_model="",
    bouwjaar=None,
    timeout=20,
):
    merk = (merk or "").strip()
    model = (model or "").strip()
    type_model = (type_model or "").strip()

    if not merk or not model:
        return None

    # RDW query
    params = {
        "$limit": 500,
        "merk": merk.upper(),
    }

    try:
        response = requests.get(
            RDW_API_URL,
            params=params,
            timeout=timeout,
        )
        response.raise_for_status()
        rows = response.json()
    except Exception:
        return None

    if not rows:
        return None

    filtered_rows, match_level = _filter_rows(
        rows=rows,
        model=model,
        type_model=type_model,
        bouwjaar=bouwjaar,
    )

    if not filtered_rows:
        return {
            "cataloguswaarde": "",
            "gewicht": "",
            "brandstof": "",
            "voertuig_type": "personenauto",
            "sample_size": 0,
            "match_level": "manual_required",
            "is_schatting": True,
        }

    cataloguswaarden = []
    gewichten = []
    brandstoffen = []
    voertuigsoorten = []

    for row in filtered_rows:
        cataloguswaarden.append(
            _to_int(row.get("catalogusprijs"))
        )
        gewichten.append(
            _to_int(row.get("massa_ledig_voertuig"))
        )

        brandstof = row.get("brandstof_omschrijving")
        if brandstof:
            brandstoffen.append(brandstof)

        voertuigsoort = row.get("voertuigsoort")
        if voertuigsoort:
            voertuigsoorten.append(voertuigsoort)

    cataloguswaarde = _median(cataloguswaarden)
    gewicht = _median(gewichten)
    brandstof = _most_common(brandstoffen)
    voertuigsoort = _most_common(voertuigsoorten)

    return {
        "cataloguswaarde": str(cataloguswaarde) if cataloguswaarde else "",
        "gewicht": str(gewicht) if gewicht else "",
        "brandstof": brandstof,
        "voertuig_type": _vehicle_type_from_rdw(voertuigsoort),
        "sample_size": len(filtered_rows),
        "match_level": match_level,
        "is_schatting": True,
    }
