from __future__ import annotations

from pathlib import Path
import re
from typing import Any


def _project_root() -> Path:
    return Path(__file__).resolve().parent


def _resolve_in_templates(name: str) -> Path:
    """
    Zoek een template in:
    - exact pad (als het bestaat)
    - projectroot/templates/<name>
    - als <name> geen extensie heeft: prefer .html, fallback .txt
    """
    p = Path(name)

    # 1) Exact pad zoals doorgegeven
    if p.exists():
        return p

    templates_dir = _project_root() / "templates"

    # 2) Als alleen bestandsnaam of relatief pad: in templates map zoeken
    candidate = templates_dir / p.name
    if candidate.exists():
        return candidate

    # 3) Geen extensie? -> prefer html, fallback txt
    if p.suffix == "":
        html = templates_dir / f"{p.name}.html"
        if html.exists():
            return html
        txt = templates_dir / f"{p.name}.txt"
        if txt.exists():
            return txt

    # 4) Als er wel extensie is maar file bestaat niet: nog 1 keer in templates
    if p.suffix:
        candidate2 = templates_dir / p.name
        if candidate2.exists():
            return candidate2

    raise FileNotFoundError(f"Template niet gevonden: {name}")


def load_template(path_like: Any) -> str:
    """
    Laad template tekst (html of txt).
    Je mag meegeven:
    - "mail_definitief"  -> pakt mail_definitief.html als die bestaat, anders .txt
    - "mail_definitief.html" / ".txt"
    - "templates/mail_definitief.html"
    """
    p = _resolve_in_templates(str(path_like))
    return p.read_text(encoding="utf-8-sig")


def normalize_vehicle_name(s: str) -> str:
    """
    Normaliseert voertuignaam:
    - Alleen als de input (bijna) volledig uppercase is -> nette titelvorm.
    - Behoudt veelvoorkomende afkortingen/codes (GTI, ST, RS, TDI, PHEV, etc.).
    """
    s = (s or "").strip()
    if not s:
        return ""

    letters = [ch for ch in s if ch.isalpha()]
    if letters:
        upper_ratio = sum(ch.isupper() for ch in letters) / len(letters)
    else:
        upper_ratio = 0

    # Alleen aanpassen als het praktisch ALL CAPS is
    if upper_ratio < 0.9:
        return s

    parts = re.split(r"(\s+)", s)

    def fix_word(w: str) -> str:
        if not w or w.isspace():
            return w

        # Afkortingen (2-5 letters) behouden
        if re.fullmatch(r"[A-Z]{2,5}", w):
            return w

        # Codes zoals i4 / e208 (maak eerste letter klein)
        if re.fullmatch(r"[A-Za-z]\d+[A-Za-z]?", w):
            return w.lower()[0] + w[1:]

        # Woorden met '-' per segment titelvorm (maar afkortingen behouden)
        if "-" in w:
            segs = w.split("-")
            new = []
            for seg in segs:
                if re.fullmatch(r"[A-Z]{2,5}", seg):
                    new.append(seg)
                else:
                    new.append(seg[:1].upper() + seg[1:].lower())
            return "-".join(new)

        # Standaard titelvorm
        return w[:1].upper() + w[1:].lower()

    return "".join(fix_word(p) for p in parts)


def render_template(template_text: str, values: dict[str, Any]) -> str:
    """
    Vervangt {{key}} placeholders.
    - Ondersteunt ook spaties: {{ key }}
    - None wordt ""
    - Onbekende placeholders blijven staan (handig voor debug)
    """

    def repl(match: re.Match) -> str:
        key = match.group(1).strip()
        if key in values:
            v = values.get(key)
            if v is None:
                return ""

            # ✅ Alleen voor voertuignaam: maak ALL CAPS netjes (Ford Kuga)
            if key == "auto":
                return normalize_vehicle_name(str(v))

            return str(v)
        return match.group(0)

    return re.sub(r"\{\{\s*([^}]+?)\s*\}\}", repl, template_text)


# Tussenvoegsels (NL + wat vaak voorkomende varianten)
_TUSSENVOEGSELS = {
    "de", "den", "der",
    "van", "von",
    "ten", "ter", "te",
    "op", "aan", "in",
    "la", "le", "du", "da", "di", "del", "della",
    "v.d.", "vd", "v/d",
}


def guess_aanhef_en_achternaam(naam: str) -> tuple[str, str]:
    """
    Probeert aanhef (heer/mevrouw) en achternaam (incl. tussenvoegsels) uit een naamregel te halen.
    Voorbeelden:
    - "Dhr A. de Boer" -> ("heer", "de Boer")
    - "Mevr. J van der Meer" -> ("mevrouw", "van der Meer")
    - "A. Jansen" -> ("heer/mevrouw", "Jansen")
    """
    n = (naam or "").strip()
    if not n:
        return "heer/mevrouw", ""

    low = n.lower()

    # Aanhef detectie (incl. puntjes)
    aanhef = "heer/mevrouw"
    if re.search(r"\b(mevr|mw|mevrouw)\b\.?", low):
        aanhef = "mevrouw"
    elif re.search(r"\b(dhr|de\s+heer|heer)\b\.?", low):
        aanhef = "heer"

    # Verwijder aanhefwoorden uit de originele string (voor achternaam parsing)
    cleaned = re.sub(r"\b(dhr|de\s+heer|heer|mevr|mw|mevrouw)\b\.?", "", n, flags=re.I).strip()
    cleaned = re.sub(r"\s{2,}", " ", cleaned)

    if "," in cleaned:
        surname, _rest = cleaned.split(",", 1)
        surname = re.sub(r"\s{2,}", " ", surname).strip()
        if surname:
            return aanhef, surname

    parts = cleaned.split()

    if not parts:
        return aanhef, ""

    # Achternaam van achteren opbouwen, incl. tussenvoegsels
    surname_parts = [parts[-1]]
    i = len(parts) - 2
    while i >= 0:
        token = parts[i]
        token_low = token.lower().strip(".")
        if token_low in _TUSSENVOEGSELS:
            surname_parts.insert(0, token)
            i -= 1
            continue
        break

    achternaam = " ".join(surname_parts)
    return aanhef, achternaam


def build_aanhefregel(is_zakelijk: bool, naam: str | None = None) -> str:
    """
    Maakt exact de aanhefregel die in de template komt.
    - Zakelijk: altijd "Geachte heer/mevrouw,"
    - Particulier: "Geachte <heer/mevrouw> <achternaam>,"
    """
    if is_zakelijk:
        return "Geachte heer/mevrouw,"

    aanhef, achternaam = guess_aanhef_en_achternaam(naam or "")
    if achternaam:
        return f"Geachte {aanhef} {achternaam},"
    return f"Geachte {aanhef},"
