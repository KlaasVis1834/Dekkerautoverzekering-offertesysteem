# rolls_kiwa.py
import os
import json
import re
import requests
from typing import Any, Dict, Optional, Union, Tuple

ROLLS_BASE = "https://vergelijken.rolls.nl"
ROLLS_VEHICLE_ENDPOINT = "/beheer/data/kiwa/{kenteken}"

ROLLS_TOKEN_ENV = "ROLLS_TOKEN_URL"
ROLLS_SESSION_COOKIE_ENV = "ROLLS_PHPSESSID"
DEBUG_ENV = "ROLLS_DEBUG"

JsonVal = Union[Dict[str, Any], list]


def _safe_str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _dbg(*args):
    if os.getenv(DEBUG_ENV):
        print("ROLLS_KIWA_DEBUG:", *args)


def _norm_plate(k: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", _safe_str(k).upper())


def _maybe_load_json(s: Any) -> Optional[JsonVal]:
    if isinstance(s, (dict, list)):
        return s
    if isinstance(s, (bytes, bytearray)):
        try:
            s = s.decode("utf-8", errors="ignore")
        except Exception:
            return None
    if isinstance(s, str):
        s = s.strip()
        if not s:
            return None
        try:
            return json.loads(s)
        except Exception:
            return None
    return None


def _looks_like_html(t: str) -> bool:
    t = (t or "").lower()
    return "<html" in t or "<!doctype" in t


def _map_voertuig_type(raw: str) -> str:
    raw = _safe_str(raw).upper()
    if raw == "PA":
        return "personenauto"
    if raw == "BA":
        return "bestelauto"

    raw_low = _safe_str(raw).lower()
    if "personen" in raw_low:
        return "personenauto"
    if "bestel" in raw_low or "bedrijfs" in raw_low:
        return "bestelauto"

    return ""


def _iter_nodes(obj: Any):
    yield obj
    if isinstance(obj, dict):
        for v in obj.values():
            yield from _iter_nodes(v)
    elif isinstance(obj, list):
        for item in obj:
            yield from _iter_nodes(item)


def _find_first_value(obj: Any, keys: tuple[str, ...]) -> str:
    wanted = {k.lower() for k in keys}

    for node in _iter_nodes(obj):
        if isinstance(node, dict):
            for k, v in node.items():
                if str(k).lower() in wanted:
                    s = _safe_str(v)
                    if s:
                        return s
    return ""


def _find_best_ob_dict(obj: Any) -> Dict[str, Any]:
    candidates: list[Dict[str, Any]] = []

    for node in _iter_nodes(obj):
        if isinstance(node, dict):
            score = 0
            keys_lower = {str(k).lower() for k in node.keys()}

            if "ob_mldcode" in keys_lower:
                score += 5
            if "ob_objsrt" in keys_lower:
                score += 4
            if "ob_merk" in keys_lower:
                score += 3
            if "ob_model" in keys_lower:
                score += 3
            if "ob_brandstof" in keys_lower:
                score += 2
            if "ob_bouwjaar" in keys_lower:
                score += 2

            if score > 0:
                tmp = dict(node)
                tmp["_score_internal_rolls"] = score
                candidates.append(tmp)

    if not candidates:
        return {}

    candidates.sort(key=lambda d: d.get("_score_internal_rolls", 0), reverse=True)
    best = dict(candidates[0])
    best.pop("_score_internal_rolls", None)
    return best


def _extract_payload(text: str) -> Dict[str, Any]:
    payload = _maybe_load_json(text)
    if not payload:
        return {}

    if isinstance(payload, list):
        payload = payload[0] if payload and isinstance(payload[0], dict) else {}

    if not isinstance(payload, dict):
        return {}

    if isinstance(payload.get("text"), str):
        inner = _maybe_load_json(payload["text"])
        if isinstance(inner, dict):
            payload = inner

    if isinstance(payload.get("data"), dict):
        payload = payload["data"]

    for key in ("result", "response", "payload", "vehicle"):
        if isinstance(payload.get(key), str):
            inner = _maybe_load_json(payload[key])
            if isinstance(inner, dict):
                payload[key] = inner

    return payload if isinstance(payload, dict) else {}


def parse_rolls_kiwa_response_text(text: str) -> Dict[str, Any]:
    payload = _extract_payload(text)
    if not payload:
        return {}

    ob = _find_best_ob_dict(payload)

    meldcode = (
        _safe_str(ob.get("OB_MLDCODE"))
        or _find_first_value(payload, ("OB_MLDCODE", "MeldCode", "meldcode", "ml_decode"))
    )

    voertuig_type_raw = (
        _safe_str(ob.get("OB_OBJSRT"))
        or _find_first_value(payload, ("OB_OBJSRT", "voertuig_type", "voertuigsoort", "objectsoort"))
    )
    voertuig_type = _map_voertuig_type(voertuig_type_raw)

    merk = (
        _safe_str(ob.get("OB_MERK"))
        or _find_first_value(payload, ("OB_MERK", "Merk", "merk"))
    )

    model = (
        _safe_str(ob.get("OB_MODEL"))
        or _find_first_value(payload, ("OB_MODEL", "Model", "model"))
    )

    type_model = (
        _safe_str(ob.get("OB_TYPEMODEL"))
        or _safe_str(ob.get("OB_TYPE"))
        or _find_first_value(payload, ("OB_TYPEMODEL", "OB_TYPE", "type_model", "Type", "type"))
    )

    brandstof = (
        _safe_str(ob.get("OB_BRANDSTOF"))
        or _find_first_value(payload, ("OB_BRANDSTOF", "Brandstof", "brandstof"))
    )

    bouwjaar = (
        _safe_str(ob.get("OB_BOUWJAAR"))
        or _find_first_value(payload, ("OB_BOUWJAAR", "Bouwjaar", "bouwjaar"))
    )

    gewicht = (
        _safe_str(ob.get("OB_GEWICHT"))
        or _find_first_value(payload, ("OB_GEWICHT", "Gewicht", "gewicht"))
    )

    return {
        "meldcode": meldcode,
        "voertuig_type": voertuig_type,
        "voertuig_type_raw": voertuig_type_raw,
        "merk": merk,
        "model": model,
        "type_model": type_model,
        "brandstof": brandstof,
        "bouwjaar": bouwjaar,
        "gewicht": gewicht,
        "raw_data": payload,
    }


def fetch_rolls_vehicle_text(kenteken: str) -> str:
    plate = _norm_plate(kenteken)
    if not plate:
        return ""

    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36 Edg/147.0.0.0",
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "*/*",
        "Referer": ROLLS_BASE + "/",
    })

    php_sessid = _safe_str(os.getenv(ROLLS_SESSION_COOKIE_ENV))
    if php_sessid:
        s.cookies.set("PHPSESSID", php_sessid, domain="vergelijken.rolls.nl")
        _dbg("PHPSESSID handmatig gezet")

    token_url = _safe_str(os.getenv(ROLLS_TOKEN_ENV))
    if token_url:
        try:
            warm = s.get(token_url, timeout=15)
            _dbg("Token URL bezocht:", warm.status_code)
            _dbg("Cookies na warmup:", s.cookies.get_dict())
        except Exception as e:
            _dbg("Token URL fout:", e)

    url = f"{ROLLS_BASE}{ROLLS_VEHICLE_ENDPOINT.format(kenteken=plate)}"
    r = s.get(url, timeout=15, allow_redirects=True)

    _dbg("final url:", r.url)
    _dbg("history:", [resp.status_code for resp in r.history])
    _dbg("status:", r.status_code)
    _dbg("content-type:", r.headers.get("content-type"))
    _dbg("cookies na request:", s.cookies.get_dict())
    _dbg("body start:", r.text[:1000])

    if r.status_code != 200 or _looks_like_html(r.text):
        return ""

    return r.text


def get_vehicle_data(kenteken: str) -> Dict[str, Any]:
    text = fetch_rolls_vehicle_text(kenteken)
    data = parse_rolls_kiwa_response_text(text) if text else {}

    if os.getenv(DEBUG_ENV):
        _dbg("parsed vehicle data:", data)

    return data


def get_meldcode_en_type(kenteken: str) -> Tuple[str, str]:
    data = get_vehicle_data(kenteken)
    return data.get("meldcode", ""), data.get("voertuig_type", "")