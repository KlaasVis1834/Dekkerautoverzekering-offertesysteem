from __future__ import annotations

from statistics import median
import re

import requests


RDW_VEHICLES_URL = "https://opendata.rdw.nl/resource/m9d7-ebf2.json"


def normalize_vehicle_query(merk, model, type_model="", bouwjaar=None, brandstof=""):
    return {
        "merk": _clean_text(merk).upper(),
        "model": _clean_text(model),
        "type_model": _clean_text(type_model),
        "bouwjaar": _to_int(bouwjaar),
        "brandstof": _normalize_fuel(brandstof),
    }


def estimate_no_plate_vehicle_data(merk, model, type_model="", bouwjaar=None, brandstof="", timeout=8):
    query = normalize_vehicle_query(merk, model, type_model, bouwjaar, brandstof)
    print(
        "NO-PLATE SCHATTING INPUT:",
        {
            "merk": query["merk"],
            "model": query["model"],
            "type_model": query["type_model"],
            "bouwjaar": query["bouwjaar"],
            "brandstof": query["brandstof"],
        },
    )

    if not query["merk"] or not query["model"]:
        print("NO-PLATE SCHATTING: geen merk/model, overslaan")
        return None

    rows = find_rdw_vehicle_matches(query, timeout=timeout)
    if not rows:
        print("NO-PLATE SCHATTING: geen RDW-resultaten")
        return None

    profile = safe_pick_vehicle_profile(rows, query)
    if not profile:
        print("NO-PLATE SCHATTING: geen veilig profiel gevonden")
        return None

    print(
        "NO-PLATE SCHATTING PROFIEL:",
        {
            "matches": profile.get("sample_size"),
            "confidence": profile.get("confidence"),
            "brandstof": profile.get("brandstof"),
            "gewicht": profile.get("gewicht"),
            "cataloguswaarde": profile.get("cataloguswaarde"),
            "voertuig_type": profile.get("voertuig_type"),
            "bouwjaar": profile.get("bouwjaar"),
            "reden": profile.get("match_reason"),
        },
    )
    return profile


def find_rdw_vehicle_matches(query, timeout=8):
    try:
        response = requests.get(
            RDW_VEHICLES_URL,
            params={
                "$limit": 500,
                "merk": query["merk"],
            },
            timeout=timeout,
        )
        response.raise_for_status()
        rows = response.json()
    except Exception as e:
        print("NO-PLATE SCHATTING RDW FOUT:", repr(e))
        return []

    print("NO-PLATE SCHATTING RDW MATCHES:", len(rows))
    return rows


def safe_pick_vehicle_profile(rows, query):
    match_rows, match_reason = safe_pick_best_match(
        rows,
        query["model"],
        query["type_model"],
        query["bouwjaar"],
        query["brandstof"],
    )

    if not match_rows:
        print("NO-PLATE SCHATTING: niets gevonden", match_reason)
        return None

    result = parse_rdw_vehicle_result(match_rows)
    result["match_reason"] = match_reason
    result["is_schatting"] = True
    result["source"] = "RDW"
    result["confidence"] = _confidence(match_reason, len(match_rows), bool(query["brandstof"]))
    return result


def safe_pick_best_match(rows, model, type_model="", bouwjaar=None, brandstof=""):
    model_tokens = _tokens(model)
    type_tokens = _tokens(type_model)
    year = _to_int(bouwjaar)
    fuel = _normalize_fuel(brandstof)

    if not model_tokens:
        return [], "geen modeltokens"

    candidates = [
        row for row in rows
        if _contains_all_tokens(row.get("handelsbenaming") or "", model_tokens)
    ]
    candidates = _filter_by_year(candidates, year)
    fuel_candidates = _filter_by_fuel(candidates, fuel)
    if fuel and fuel_candidates:
        candidates = fuel_candidates

    if not candidates:
        return [], "geen modelmatch"

    if type_tokens:
        exact = [
            row for row in candidates
            if _contains_all_tokens(row.get("handelsbenaming") or "", type_tokens)
        ]
        if exact:
            return exact, "model en type_model match"

        partial = []
        for row in candidates:
            name = _normalize(row.get("handelsbenaming") or "")
            matched = sum(1 for token in type_tokens if token in name)
            if matched >= max(1, len(type_tokens) // 2):
                partial.append(row)
        if partial:
            return partial, "model en gedeeltelijke type_model match"

    return candidates, "model match"


def parse_rdw_vehicle_result(rows):
    cataloguswaarden = []
    gewichten = []
    bouwjaren = []
    brandstoffen = []
    voertuigtypes = []

    for row in rows:
        cataloguswaarden.append(_to_int(row.get("catalogusprijs")))
        gewichten.append(_to_int(row.get("massa_ledig_voertuig")))
        bouwjaren.append(_year_from_row(row))

        brandstof = row.get("brandstof_omschrijving") or row.get("brandstof")
        if brandstof:
            brandstoffen.append(str(brandstof).strip())

        voertuigtypes.append(_vehicle_type_from_rdw(row.get("voertuigsoort")))

    cataloguswaarde = _median(cataloguswaarden)
    gewicht = _median(gewichten)
    bouwjaar = _median(bouwjaren)

    return {
        "brandstof": _most_common(brandstoffen),
        "gewicht": str(gewicht) if gewicht else "",
        "cataloguswaarde": str(cataloguswaarde) if cataloguswaarde else "",
        "voertuig_type": _most_common(voertuigtypes) or "personenauto",
        "bouwjaar": bouwjaar,
        "sample_size": len(rows),
    }


def parse_rdw_result(rows):
    return parse_rdw_vehicle_result(rows)


def _clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def _normalize(value):
    text = _clean_text(value).lower()
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _tokens(value):
    return [token for token in _normalize(value).split() if len(token) >= 2]


def _contains_all_tokens(value, tokens):
    normalized = _normalize(value)
    return all(token in normalized for token in tokens)


def _to_int(value):
    if value in (None, ""):
        return None
    try:
        return int(round(float(str(value).replace(",", "."))))
    except Exception:
        return None


def _year_from_row(row):
    value = row.get("datum_eerste_toelating") or ""
    return int(value[:4]) if len(value) >= 4 and value[:4].isdigit() else None


def _filter_by_year(rows, bouwjaar):
    if not bouwjaar:
        return rows
    filtered = [
        row for row in rows
        if _year_from_row(row) is None or abs(_year_from_row(row) - bouwjaar) <= 1
    ]
    return filtered or rows


def _median(values):
    nums = [value for value in values if value is not None]
    return int(round(median(nums))) if nums else None


def _most_common(values):
    counts = {}
    for value in values:
        if not value:
            continue
        counts[value] = counts.get(value, 0) + 1
    if not counts:
        return ""
    return sorted(counts.items(), key=lambda item: item[1], reverse=True)[0][0]


def _vehicle_type_from_rdw(value):
    text = _normalize(value)
    if "bestel" in text:
        return "bestelauto"
    return "personenauto"


def _normalize_fuel(value):
    text = _normalize(value)
    if not text:
        return ""
    if "benzine" in text:
        return "benzine"
    if "diesel" in text:
        return "diesel"
    if "elektr" in text:
        return "elektrisch"
    if "hybride" in text or "hybrid" in text:
        return "hybride"
    if "lpg" in text:
        return "lpg"
    return text


def _row_fuel(row):
    return _normalize_fuel(row.get("brandstof_omschrijving") or row.get("brandstof") or "")


def _filter_by_fuel(rows, fuel):
    if not fuel:
        return rows
    return [row for row in rows if _row_fuel(row) == fuel]


def _confidence(match_reason, sample_size, fuel_used):
    if sample_size >= 3 and fuel_used and "model" in match_reason:
        return "hoog"
    if sample_size >= 3 and "model" in match_reason:
        return "middel"
    return "laag"
