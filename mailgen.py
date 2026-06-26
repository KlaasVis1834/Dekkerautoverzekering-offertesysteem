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


def _clean_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def _strip_titles(name: str) -> str:
    return re.sub(
        r"\b(dhr|de\s+heer|heer|mevr|mw|mevrouw)\b\.?",
        "",
        (name or "").strip(),
        flags=re.I,
    ).strip()


def _format_initials(value: str) -> str:
    raw = re.sub(r"[^A-Za-z.]+", " ", value or "").strip()
    if not raw:
        return ""

    chunks = []
    for part in raw.split():
        letters = re.findall(r"[A-Za-z]", part)
        if not letters:
            continue
        if "." in part:
            chunks.extend(letters)
        elif len(part) <= 3 and part.isupper():
            chunks.extend(letters)
        else:
            chunks.append(letters[0])

    return " ".join(f"{ch.upper()}." for ch in chunks)


def _format_surname_word(value: str) -> str:
    value = (value or "").strip(" ,")
    if not value:
        return ""
    if value.isupper() or value.islower():
        return value[:1].upper() + value[1:].lower()
    return value[:1].upper() + value[1:]


def _split_prefix_and_surname(tokens: list[str]) -> tuple[str, str]:
    clean_tokens = [t.strip(" ,") for t in tokens if t.strip(" ,")]
    if not clean_tokens:
        return "", ""

    prefix = []
    while len(clean_tokens) > 1 and clean_tokens[0].lower().strip(".") in _TUSSENVOEGSELS:
        prefix.append(clean_tokens.pop(0).lower().strip("."))

    surname = " ".join(_format_surname_word(t) for t in clean_tokens if t)
    return " ".join(prefix), surname


def _split_initials_and_prefix(tokens: list[str]) -> tuple[str, str]:
    initials = []
    prefix = []

    for token in tokens:
        clean = token.strip(" ,")
        if not clean:
            continue

        low = clean.lower().strip(".")
        letters = re.findall(r"[A-Za-z]", clean)

        if low in _TUSSENVOEGSELS:
            prefix.append(low)
            continue

        if letters and (
            "." in clean
            or len("".join(letters)) == 1
            or (len("".join(letters)) <= 3 and clean.isupper())
        ):
            initials.append(clean)

    return _format_initials(" ".join(initials)), " ".join(prefix)


def normalize_person_name(raw_name: str) -> dict[str, str]:
    """
    Normaliseert particuliere namen voor PDF, mail en bestandsnamen.
    Ondersteunt o.a. "Baltaci, O.", "Vries, A. de", "De Vries, A." en "A. de Vries".
    """
    raw = _clean_spaces(_strip_titles(raw_name))
    if not raw:
        return {
            "initials": "",
            "tussenvoegsel": "",
            "achternaam": "",
            "display": "",
            "aanhef_naam": "",
        }

    initials = ""
    tussenvoegsel = ""
    achternaam = ""

    if "," in raw:
        surname_part, rest = raw.split(",", 1)
        surname_prefix, surname = _split_prefix_and_surname(surname_part.split())
        rest_initials, rest_prefix = _split_initials_and_prefix(rest.split())

        initials = rest_initials
        tussenvoegsel = rest_prefix or surname_prefix
        achternaam = surname
    else:
        tokens = raw.split()
        initial_tokens = []
        while tokens:
            clean = tokens[0].strip(" ,")
            letters = re.findall(r"[A-Za-z]", clean)
            if letters and ("." in clean or len("".join(letters)) == 1):
                initial_tokens.append(tokens.pop(0))
                continue
            break

        initials = _format_initials(" ".join(initial_tokens))
        tussenvoegsel, achternaam = _split_prefix_and_surname(tokens)

        if not achternaam and tokens:
            achternaam = _format_surname_word(tokens[-1])

    aanhef_naam = " ".join([p for p in [tussenvoegsel, achternaam] if p]).strip()
    display = " ".join([p for p in [initials, tussenvoegsel, achternaam] if p]).strip()

    return {
        "initials": initials,
        "tussenvoegsel": tussenvoegsel,
        "achternaam": achternaam,
        "display": display or raw,
        "aanhef_naam": aanhef_naam,
    }


def normalize_customer_name(raw_name: str) -> dict[str, str]:
    return normalize_person_name(raw_name)


def salutation_from_relatie_geslacht(relatie_geslacht: str | None, raw_name: str | None = None) -> str:
    s = _clean_spaces(relatie_geslacht or "").lower()
    if s in {"v", "vrouw", "mevrouw", "mevr", "mw"}:
        return "mevrouw"
    if s in {"m", "h", "man", "heer", "dhr", "de heer"}:
        return "heer"
    if s in {"o", "z", "onbekend", "unknown", "bedrijf", "zakelijk", "org", "organisatie"}:
        return "heer/mevrouw"

    low_name = (raw_name or "").lower()
    if re.search(r"\b(mevr|mw|mevrouw)\b\.?", low_name):
        return "mevrouw"
    if re.search(r"\b(dhr|de\s+heer|heer)\b\.?", low_name):
        return "heer"
    return "heer/mevrouw"


def filename_salutation_from_relatie_geslacht(relatie_geslacht: str | None, raw_name: str | None = None) -> str:
    aanhef = salutation_from_relatie_geslacht(relatie_geslacht, raw_name)
    if aanhef == "heer":
        return "de heer"
    if aanhef == "mevrouw":
        return "mevrouw"
    return "de heer/mevrouw"


def split_dealer_customer_name(naam: str) -> tuple[str, str]:
    """
    Splitst dealerformaten zoals "De Jong, D." in achternaam en voorletters.
    Geeft ("", "") terug als er geen bruikbare naam staat.
    """
    normalized = normalize_person_name(naam)
    if not normalized["achternaam"]:
        return "", ""

    return normalized["aanhef_naam"], normalized["initials"]


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

    aanhef = salutation_from_relatie_geslacht(None, n)
    normalized = normalize_person_name(n)
    return aanhef, normalized["aanhef_naam"]


def build_aanhefregel(
    is_zakelijk: bool,
    naam: str | None = None,
    relatie_geslacht: str | None = None,
) -> str:
    """
    Maakt exact de aanhefregel die in de template komt.
    - Zakelijk: altijd "Geachte heer/mevrouw,"
    - Particulier: "Geachte <heer/mevrouw> <achternaam>,"
    """
    aanhef = salutation_from_relatie_geslacht(relatie_geslacht, naam)
    if is_zakelijk or aanhef == "heer/mevrouw":
        return "Geachte heer/mevrouw,"

    normalized = normalize_person_name(naam or "")
    if normalized["aanhef_naam"]:
        return f"Geachte {aanhef} {normalized['aanhef_naam']},"
    return "Geachte heer/mevrouw,"
