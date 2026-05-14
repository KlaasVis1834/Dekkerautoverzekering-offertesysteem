# pdfgen.py
from __future__ import annotations

from pathlib import Path
from datetime import datetime
from typing import Optional

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

POLICY_LINK = "https://www.klaasvis.nl/dekkerautoverzekering/#policies"
AANVRAAG_LINK = "https://www.klaasvis.nl/aanvraagformulier/"
CONTACT_LINK = "https://www.klaasvis.nl/contact"
DISCLAIMER_TEKST = (
    "De aanbieding is geldig behoudens tussentijdse wijzigingen in van toepassing zijnde tarieven en/of voorwaarden."
)

# ============================================================
# ✅ BONUSKORTING (met 67,5% en 72,5%)
# ============================================================
# Particulier Personenauto Regio 1-3 (max korting vanaf 8 SVJ)
BONUS_TABEL_PARTICULIER_REGIO_1_3: dict[int, float] = {
    0: 45,
    1: 50,
    2: 55,
    3: 60,
    4: 65,
    5: 67.5,
    6: 70,
    7: 72.5,
    8: 75,
}

# Regio 4 / Zakelijke personenauto (max korting vanaf 9 SVJ)
BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN: dict[int, float] = {
    0: 40,
    1: 45,
    2: 50,
    3: 55,
    4: 60,
    5: 65,
    6: 67.5,
    7: 70,
    8: 72.5,
    9: 75,
}

# Bestelauto (particulier/zakelijk) (max korting vanaf 11 SVJ)
BONUS_TABEL_BESTELAUTO: dict[int, float] = {
    0: 25,
    1: 35,
    2: 40,
    3: 45,
    4: 50,
    5: 55,
    6: 60,
    7: 65,
    8: 67.5,
    9: 70,
    10: 72.5,
    11: 75,
}


def _week_folder(dt: datetime) -> str:
    iso_year, iso_week, _ = dt.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def _ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _assets_path() -> Path:
    return Path(__file__).resolve().parent / "assets"


def _safe(v) -> str:
    return ("" if v is None else str(v)).strip()


def _parse_money_to_float(v) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("€", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _fmt_euro(v) -> str:
    """
    € 10.000,00 (duizendtallen punt, decimalen komma)
    """
    if v is None or str(v).strip() == "":
        return "€ —"
    if isinstance(v, (int, float)):
        f = float(v)
    else:
        f = _parse_money_to_float(v)
        if f is None:
            s = str(v).strip()
            return s if s.startswith("€") else f"€ {s}"
    t = f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"€ {t}"


def _fmt_pct(pct) -> str:
    """
    Toon percent netjes:
    - 67.5 -> '67,5%'
    - 72.5 -> '72,5%'
    - 70   -> '70%'
    """
    if pct is None:
        return "—"
    try:
        f = float(pct)
    except Exception:
        return f"{pct}%"
    if abs(f - round(f)) < 1e-9:
        return f"{int(round(f))}%"
    s = f"{f:.1f}".replace(".", ",")
    return f"{s}%"


def _draw_image_fit(
    c: canvas.Canvas, img_path: Path, x: float, y: float, max_w: float, max_h: float
) -> None:
    if not img_path.exists():
        return
    img = ImageReader(str(img_path))
    iw, ih = img.getSize()
    if iw <= 0 or ih <= 0:
        return
    scale = min(max_w / iw, max_h / ih)
    w = iw * scale
    h = ih * scale
    c.drawImage(str(img_path), x, y, width=w, height=h, mask="auto")


def _section_title(c: canvas.Canvas, x: float, y: float, title: str, right_edge: float) -> float:
    c.setFillColor(colors.HexColor("#1F2A44"))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, y, title)
    c.setStrokeColor(colors.HexColor("#D9DEE7"))
    c.setLineWidth(1)
    c.line(x, y - 2.5 * mm, right_edge, y - 2.5 * mm)
    return y - 8 * mm


def _kv(c: canvas.Canvas, x: float, y: float, label: str, value: str, col_w=32 * mm) -> float:
    c.setFillColor(colors.HexColor("#2B2B2B"))
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(x, y, label)
    c.setFont("Helvetica", 9.5)
    c.setFillColor(colors.HexColor("#111111"))
    c.drawString(x + col_w, y, value if value else "—")
    return y - 5.9 * mm


def _auto_zonder_uitvoering(voertuig: dict) -> str:
    merk = _safe(voertuig.get("merk", ""))
    model = _safe(voertuig.get("model", ""))
    if merk or model:
        return f"{merk} {model}".strip()

    auto = _safe(voertuig.get("auto", ""))
    if not auto:
        return ""

    parts = auto.split()
    vt = _safe(voertuig.get("voertuig_type")).lower()
    n = 3 if vt == "bestelauto" else 2
    return " ".join(parts[:n]) if len(parts) >= n else auto


def _link(c: canvas.Canvas, x: float, y: float, text: str, url: str, size=9.5) -> None:
    c.setFillColor(colors.HexColor("#0B5ED7"))
    c.setFont("Helvetica", size)
    c.drawString(x, y, text)
    width = c.stringWidth(text, "Helvetica", size)
    c.linkURL(url, (x, y - 1.5 * mm, x + width, y + 2.5 * mm), relative=0)


def _benodigde_svj_text(offer: dict, voertuig: dict, klant: dict) -> int:
    """
    Logica:
    - Bestelauto (particulier of zakelijk): 11
    - Zakelijke personenauto: 9
    - Personenauto regio 4: 9
    - Particuliere personenauto regio 1-3: 8
    """
    klant_type = _safe(klant.get("klant_type")).lower()
    voertuig_type = _safe(voertuig.get("voertuig_type")).lower()
    soort = _safe(voertuig.get("soort")).lower()
    regio = str(_safe(offer.get("regio"))).strip()

    if voertuig_type == "bestelauto":
        return 11

    bestel_hint = any(x in (voertuig_type + " " + soort) for x in ["bestel", "bestelauto", "van", "bedrijfsauto"])
    if bestel_hint or "bestel" in str(_safe(offer.get("voertuigsoort"))).lower():
        return 11

    if klant_type == "zakelijk":
        return 9

    if regio == "4":
        return 9

    return 8


def _mini_block(c: canvas.Canvas, x: float, y: float, w: float, h: float) -> None:
    c.setFillColor(colors.HexColor("#F6F8FB"))
    c.setStrokeColor(colors.HexColor("#D9DEE7"))
    c.roundRect(x, y - h, w, h, 6, stroke=1, fill=1)


def _wrap_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_w: float,
    font="Helvetica",
    size=9.4,
    leading=4.6 * mm,
) -> float:
    c.setFont(font, size)
    words = (text or "").split()
    if not words:
        return y
    line = ""
    for w_ in words:
        test = (line + " " + w_).strip()
        if c.stringWidth(test, font, size) <= max_w:
            line = test
        else:
            c.drawString(x, y, line)
            y -= leading
            line = w_
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def _waarde_suffix(kenteken: str) -> str:
    return "(RDW)" if _safe(kenteken) else "(ingeschat)"


def _waarde_context(klant: dict, voertuig: dict, offer: dict) -> str:
    """
    Jouw regels:
    - Particulier personenauto: incl. btw
    - Particulier bestelauto: incl. btw
    - Zakelijke personenauto: excl. btw / incl. bpm
    - Zakelijke bestelauto: excl. btw / excl. bpm
    """
    kt = _safe(klant.get("klant_type")).lower()
    vt = _safe(voertuig.get("voertuig_type")).lower()

    if kt == "zakelijk":
        if vt == "bestelauto":
            return "excl. btw / excl. bpm"
        return "excl. btw / incl. bpm"

    return "incl. btw"


def _waarde_naar_context_float(
    waarde_raw,
    klant: dict,
    voertuig: dict,
    offer: dict,
) -> tuple[Optional[float], str]:
    """
    RDW catalogus/dagwaarde is doorgaans INCL btw (en catalogus ook incl bpm).

    Regels:
    - Particulier personenauto: incl. btw -> geen omrekening
    - Particulier bestelauto: incl. btw -> geen omrekening
    - Zakelijke personenauto: excl. btw / incl. bpm -> waarde / 1.21
    - Zakelijke bestelauto: excl. btw / excl. bpm -> (waarde / 1.21) - bpm  (als bpm bekend)

    ✅ No-plate: als offer['waarde_al_in_context'] == '1' -> NIET omrekenen.
    """
    if str(offer.get("waarde_al_in_context") or "").strip() in ("1", "true", "True", "yes", "Y"):
        base = _parse_money_to_float(waarde_raw)
        return (base, "") if base is not None else (None, "")

    vt = _safe(voertuig.get("voertuig_type")).lower()
    kt = _safe(klant.get("klant_type")).lower()

    base = _parse_money_to_float(waarde_raw)
    if base is None:
        return None, ""

    # ✅ particulier altijd incl btw, ook bestelauto
    if kt != "zakelijk":
        return base, ""

    # Zakelijke personenauto: excl btw / incl bpm
    if vt != "bestelauto" and kt == "zakelijk":
        return (base / 1.21), ""

    # Zakelijke bestelauto: excl btw / excl bpm
    if vt == "bestelauto" and kt == "zakelijk":
        excl_btw = base / 1.21
        bpm = _parse_money_to_float(offer.get("bpm"))
        if bpm is None:
            return excl_btw, " (bpm onbekend)"
        return (excl_btw - bpm), ""

    return base, ""


def _dekking_parts(dekking: str) -> list[str]:
    d = _safe(dekking)
    if not d:
        return []
    return [p.strip() for p in d.split("/") if p.strip()]


def _base_dekking(dekking: str) -> str:
    parts = _dekking_parts(dekking)
    if not parts:
        return ""
    normalized = [p.lower() for p in parts]
    if any("casco compleet" in p or "allrisk" in p for p in normalized):
        return "WA / Casco Compleet (Allrisk)"
    if any("beperkt casco" in p for p in normalized):
        return "WA / Beperkt Casco"
    return "WA"


def _is_wa_only(dekking: str) -> bool:
    return _base_dekking(dekking).strip().lower() == "wa"


def _is_beperkt_casco(dekking: str) -> bool:
    return _base_dekking(dekking).strip().lower() == "wa / beperkt casco"


def _dekking_lines(dekking: str) -> list[str]:
    """
    Zet dekking onder elkaar en voegt (Allrisk) toe achter Casco Compleet.
    Input verwacht zoals: "WA / Casco Compleet / Schadeverzekering Inzittenden"
    """
    d = _safe(dekking)
    if not d:
        return ["—"]

    parts = [p.strip() for p in d.split("/") if p.strip()]
    out: list[str] = []
    for p in parts:
        if "casco compleet" in p.lower() and "(allrisk)" not in p.lower():
            out.append(f"{p} (Allrisk)")
        else:
            out.append(p)
    return out


def _parse_int(v) -> Optional[int]:
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def _bonus_pct(klant: dict, voertuig: dict, offer: dict, svj_used: int) -> float:
    """
    Bonustabel selecteren + waarde bepalen.
    - Bestelauto -> BONUS_TABEL_BESTELAUTO
    - Regio 4 OF klant_type zakelijk (personenauto) -> BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN
    - Anders -> BONUS_TABEL_PARTICULIER_REGIO_1_3
    """
    vt = _safe(voertuig.get("voertuig_type")).lower()
    regio = str(_safe(offer.get("regio"))).strip()
    kt = _safe(klant.get("klant_type")).lower()

    if svj_used < 0:
        svj_used = 0

    if vt == "bestelauto":
        table = BONUS_TABEL_BESTELAUTO
    elif kt == "zakelijk" or regio == "4":
        table = BONUS_TABEL_REGIO4_OF_ZAKELIJK_PERSONEN
    else:
        table = BONUS_TABEL_PARTICULIER_REGIO_1_3

    if not table:
        return 75.0

    if svj_used in table:
        return float(table[svj_used])

    m = max(table.keys())
    if svj_used > m:
        return float(table[m])

    mn = min(table.keys())
    if svj_used < mn:
        return float(table[mn])

    lower_keys = [k for k in table.keys() if k <= svj_used]
    return float(table[max(lower_keys)]) if lower_keys else float(table[mn])


def generate_offer_pdf(
    out_base_dir: str,
    dt: datetime,
    offer_no: str,
    klant: dict,
    voertuig: dict,
    offer: dict,
    filename_base: str | None = None,
) -> str:
    base = Path(out_base_dir)
    week_dir = base / _week_folder(dt)
    _ensure_dir(week_dir)

    # ✅ FIX: indentatie + nette pdf-extensie
    base_name = (filename_base or offer_no).strip() or offer_no
    if not base_name.lower().endswith(".pdf"):
        base_name += ".pdf"
    out_path = week_dir / base_name

    assets = _assets_path()
    logo_dekker = assets / "logo_dekker.png"
    logo_klaasvis = assets / "logo_klaasvis.png"

    klantnaam = _safe(klant.get("naam"))
    adres = _safe(klant.get("adres"))
    postcode = _safe(klant.get("postcode"))
    plaats = _safe(klant.get("plaats"))
    telefoon = _safe(klant.get("telefoon"))
    email = _safe(klant.get("email"))

    auto_basic = _auto_zonder_uitvoering(voertuig)
    kenteken = _safe(voertuig.get("kenteken"))

    # ✅ FIX: geen kenteken -> toon '-' bij Kenteken en Meldcode
    kenteken_display = kenteken if kenteken else "-"
    meldcode_raw = _safe(offer.get("meldcode"))
    meldcode_display = meldcode_raw if kenteken else "-"

    gewicht = _safe(offer.get("gewicht"))
    bouwjaar = _safe(offer.get("bouwjaar"))
    brandstof = _safe(voertuig.get("brandstof"))
    dekking = _safe(offer.get("dekking"))

    cataloguswaarde = _safe(offer.get("cataloguswaarde"))
    dagwaarde = _safe(offer.get("dagwaarde"))

    premie_raw = offer.get("premie_maand", offer.get("premie"))
    premie_float = _parse_money_to_float(premie_raw)
    premie_txt = _fmt_euro(premie_float if premie_float is not None else premie_raw)

    # ✅ DIENSTVERLENING 18%
    beloning_float = (premie_float * 0.18) if premie_float is not None else None
    beloning_txt = _fmt_euro(beloning_float) if beloning_float is not None else "€ —"

    required_svj = _benodigde_svj_text(offer, voertuig, klant)
    svj_override = _parse_int(offer.get("svj_override"))

    is_gezinsauto_regeling = (svj_override == -10)
    if is_gezinsauto_regeling:
        used_svj = required_svj
        bonus_pct = 60.0
    else:
        used_svj = svj_override if svj_override is not None else required_svj
        bonus_pct = _bonus_pct(klant, voertuig, offer, used_svj)

    if not dekking:
        dekking = "WA / Casco Compleet (Allrisk) / Schadeverzekering Inzittenden"

    is_beperkt = _is_beperkt_casco(dekking)
    is_wa_only = _is_wa_only(dekking)
    waarde_label = "Dagwaarde" if is_beperkt else "Cataloguswaarde"
    waarde_raw = dagwaarde if is_beperkt else cataloguswaarde

    suffix = _waarde_suffix(kenteken)
    waarde_ctx = _waarde_context(klant, voertuig, offer)

    waarde_txt = ""
    if (not is_wa_only) and waarde_raw:
        waarde_conv, note = _waarde_naar_context_float(waarde_raw, klant, voertuig, offer)
        if waarde_conv is not None:
            waarde_txt = f"{_fmt_euro(waarde_conv)} {waarde_ctx}{note} {suffix}".strip()
        else:
            waarde_txt = f"{_fmt_euro(waarde_raw)} {waarde_ctx} {suffix}".strip()

    c = canvas.Canvas(str(out_path), pagesize=A4)
    w, h = A4
    margin_x = 15 * mm

    c.setFillColor(colors.white)
    c.rect(0, h - 32 * mm, w, 32 * mm, fill=1, stroke=0)

    _draw_image_fit(c, logo_dekker, margin_x, h - 28 * mm, max_w=95 * mm, max_h=16 * mm)
    _draw_image_fit(c, logo_klaasvis, w - margin_x - 40 * mm, h - 30 * mm, max_w=40 * mm, max_h=20 * mm)

    c.setFillColor(colors.HexColor("#1F2A44"))
    c.setFont("Helvetica-Bold", 14)
    c.drawString(margin_x, h - 40 * mm, "Verzekeringsvoorstel Dekkerautoverzekering")
    c.setFont("Helvetica-Bold", 10.5)
    c.drawRightString(w - margin_x, h - 40 * mm, f"Offertenummer: {offer_no}")

    y = h - 52 * mm

    col_gap = 12 * mm
    col_w = (w - 2 * margin_x - col_gap) / 2
    left_x = margin_x
    right_x = margin_x + col_w + col_gap

    y_title = y
    left_end = left_x + col_w
    right_end = right_x + col_w

    y_left = _section_title(c, left_x, y_title, "Klantgegevens", left_end)
    y_right = _section_title(c, right_x, y_title, "Uw adviseur", right_end)

    yl = y_left
    yl = _kv(c, left_x, yl, "Naam", klantnaam, col_w=28 * mm)
    yl = _kv(c, left_x, yl, "Adres", adres, col_w=28 * mm)
    yl = _kv(c, left_x, yl, "Postcode/plaats", f"{postcode} {plaats}".strip(), col_w=28 * mm)
    yl = _kv(c, left_x, yl, "Telefoon", telefoon, col_w=28 * mm)
    yl = _kv(c, left_x, yl, "E-mail", email if email else "— (verzending per post)", col_w=28 * mm)

    yr = y_right
    yr = _kv(c, right_x, yr, "Bedrijf", "B.V. Assurantiekantoor Klaas Vis ao. 1834", col_w=26 * mm)
    yr = _kv(c, right_x, yr, "Adres", "Zuiderweg 7", col_w=26 * mm)
    yr = _kv(c, right_x, yr, "Postcode/plaats", "1456 NC WIJDEWORMER", col_w=26 * mm)
    yr = _kv(c, right_x, yr, "Telefoon", "075 – 631 42 61", col_w=26 * mm)
    yr = _kv(c, right_x, yr, "E-mail", "info@klaasvis.nl", col_w=26 * mm)

    y = min(yl, yr) - 6 * mm

    voertuig_x = margin_x
    svj_x = margin_x + col_w + col_gap
    voertuig_end = voertuig_x + col_w
    svj_end = svj_x + col_w

    y_voertuig = _section_title(c, voertuig_x, y, "Voertuiggegevens", voertuig_end)
    y_svj = _section_title(c, svj_x, y, "Dekking & schadevrije jaren", svj_end)

    yv = y_voertuig
    yv = _kv(c, voertuig_x, yv, "Auto", auto_basic, col_w=28 * mm)
    yv = _kv(c, voertuig_x, yv, "Kenteken", kenteken_display, col_w=28 * mm)
    yv = _kv(c, voertuig_x, yv, "Bouwjaar", bouwjaar, col_w=28 * mm)
    yv = _kv(c, voertuig_x, yv, "Brandstof", brandstof, col_w=28 * mm)
    yv = _kv(c, voertuig_x, yv, "Meldcode", meldcode_display, col_w=28 * mm)
    yv = _kv(c, voertuig_x, yv, "Gewicht", (gewicht + " kg").strip() if gewicht else "", col_w=28 * mm)
    if not is_wa_only:
        yv = _kv(c, voertuig_x, yv, waarde_label, waarde_txt, col_w=28 * mm)

    ys = y_svj
    box2_y_top = ys + 2 * mm
    box2_h = 44 * mm
    _mini_block(c, svj_x, box2_y_top, col_w, box2_h)

    pad = 6 * mm
    tx = svj_x + pad
    ty = box2_y_top - 8 * mm
    max_w2 = col_w - 2 * pad

    c.setFillColor(colors.HexColor("#333333"))
    c.setFont("Helvetica-Bold", 9.6)
    c.drawString(tx, ty, "Dekking:")
    ty -= 5.2 * mm

    c.setFont("Helvetica", 9.4)
    for line in _dekking_lines(dekking):
        ty = _wrap_text(c, line, tx, ty, max_w2, size=9.4, leading=4.5 * mm)

    ty -= 1.0 * mm

    if is_gezinsauto_regeling:
        svj_line = "Op basis van de 2de gezinsautoregeling (+3 extra treden)"
    elif svj_override is not None:
        svj_line = f"Uitgegaan van {used_svj} schadevrije jaren in Roydata."
    else:
        svj_line = f"Uitgegaan van minimaal {required_svj} schadevrije jaren in Roydata voor de maximale korting."

    c.setFillColor(colors.HexColor("#333333"))
    c.setFont("Helvetica", 9.2)
    ty = _wrap_text(c, svj_line, tx, ty, max_w2, size=9.2, leading=4.4 * mm)

    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica", 9.0)
    _wrap_text(c, "Bij minder schadevrije jaren: neem contact met ons op.", tx, ty, max_w2, size=9.0, leading=4.3 * mm)

    y = min(yv, (box2_y_top - box2_h)) - 6 * mm

    y = _section_title(c, margin_x, y, "Belangrijke kenmerken", w - margin_x)

    kenmerken = [
        "Een hoge bonuskorting ook zonder schadevrije jaren;",
        "Bij schadeherstel door Dekkerautogroep geen eigen risico *;",
        "Gratis vervangend vervoer bij schadeherstel door Dekkerautogroep *;",
        "Geen terugval in bonuskorting wanneer de schade verzekerde niet kan worden aangerekend (vandalisme, onbekende dader, etc.);",
        "Uitstekende nieuwwaarderegeling voor personenauto’s van 36 maanden waarbij de eerste 12 maanden 100%.",
    ]
    foot = (
        "* Uiteraard alleen indien er sprake is van dekking conform de polisvoorwaarden Casco Compleet en/of Beperkt. "
        "Met uitzondering bij vervanging van een ruit en/of panorama dak, hiervoor geldt een verlaagd eigen risico van € 75,00."
    )

    block_x = margin_x
    block_w = w - 2 * margin_x
    block_h = 46 * mm
    _mini_block(c, block_x, y + 2 * mm, block_w, block_h)

    pad = 6 * mm
    tx2 = block_x + pad
    ty2 = (y + 2 * mm) - 8 * mm
    max_text_w = block_w - 2 * pad

    for b in kenmerken:
        c.setFillColor(colors.HexColor("#1F2A44"))
        c.setFont("Helvetica-Bold", 9.6)
        c.drawString(tx2, ty2, "•")
        c.setFillColor(colors.HexColor("#333333"))
        c.setFont("Helvetica", 9.4)
        ty2 = _wrap_text(c, b, tx2 + 4 * mm, ty2, max_text_w - 4 * mm, size=9.4, leading=4.6 * mm)

    ty2 -= 0.5 * mm
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica", 8.4)
    _wrap_text(c, foot, tx2, ty2, max_text_w, size=8.4, leading=4.1 * mm)

    y = (y + 2 * mm) - block_h - 8 * mm

    y = _section_title(c, margin_x, y, "Maandpremie", w - margin_x)

    box_x = margin_x
    box_y = y - 16 * mm
    box_w = 78 * mm
    box_h = 16 * mm

    c.setFillColor(colors.HexColor("#E9F7EF"))
    c.setStrokeColor(colors.HexColor("#B7E3C6"))
    c.roundRect(box_x, box_y, box_w, box_h, 6, stroke=1, fill=1)

    c.setFillColor(colors.HexColor("#0F5132"))
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(box_x + box_w / 2, box_y + 5.0 * mm, premie_txt)

    text_x = box_x + box_w + 10 * mm
    text_y = box_y + 10.5 * mm

    c.setFillColor(colors.HexColor("#111111"))
    c.setFont("Helvetica-Bold", 10)
    c.drawString(text_x, text_y, "Uitgangspunten")
    c.setFont("Helvetica", 9.5)

    text_y -= 5.8 * mm
    c.drawString(text_x, text_y, f"• Bonuskorting: {_fmt_pct(bonus_pct)}")
    text_y -= 5.2 * mm
    c.drawString(text_x, text_y, "• Premie incl. 21% assurantiebelasting")
    text_y -= 5.2 * mm
    c.drawString(text_x, text_y, f"• Kosten dienstverlening: {beloning_txt}")

    y = box_y - 10 * mm

    c.setFillColor(colors.HexColor("#1F2A44"))
    c.setFont("Helvetica-Bold", 9.8)
    c.drawString(margin_x, y, "Akkoord met dit voorstel?")
    y -= 5.6 * mm

    c.setFillColor(colors.HexColor("#444444"))
    c.setFont("Helvetica", 9.5)
    c.drawString(margin_x, y, "U kunt de aanvraag snel en eenvoudig online indienen via onderstaand aanvraagformulier.")
    y -= 5.8 * mm

    # ============================================================
    # ✅ AANGEPAST: alleen "contact" klikbaar + blauw
    # ============================================================
    prefix = "Komt u er niet uit of heeft u vragen? Neem gerust "
    link_word = "contact"
    suffix = " met ons op."

    c.setFillColor(colors.HexColor("#444444"))
    c.setFont("Helvetica", 9.5)
    c.drawString(margin_x, y, prefix)

    prefix_w = c.stringWidth(prefix, "Helvetica", 9.5)
    link_w = c.stringWidth(link_word, "Helvetica", 9.5)

    c.setFillColor(colors.HexColor("#0B5ED7"))
    c.drawString(margin_x + prefix_w, y, link_word)
    c.linkURL(
        CONTACT_LINK,
        (
            margin_x + prefix_w,
            y - 1.5 * mm,
            margin_x + prefix_w + link_w,
            y + 2.5 * mm,
        ),
        relative=0,
    )

    c.setFillColor(colors.HexColor("#444444"))
    c.drawString(margin_x + prefix_w + link_w, y, suffix)

    y -= 6.2 * mm

    # Persoonlijke aanvraaglink met offertenummer
    aanvraag_url = f"{AANVRAAG_LINK}?offerte={offer_no}"

    c.setFillColor(colors.HexColor("#1F2A44"))
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(margin_x, y, "Aanvraagformulier:")
    _link(c, margin_x + 34 * mm, y, aanvraag_url, aanvraag_url, size=9.5)
    y -= 10 * mm

    c.setStrokeColor(colors.HexColor("#D9DEE7"))
    c.line(margin_x, 30 * mm, w - margin_x, 30 * mm)

    c.setFillColor(colors.HexColor("#1F2A44"))
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(margin_x, 24.5 * mm, "Polisvoorwaarden & verzekeringskaarten:")
    _link(c, margin_x, 19.8 * mm, POLICY_LINK, POLICY_LINK, size=9.5)

    c.setFillColor(colors.HexColor("#444444"))
    c.setFont("Helvetica", 8.8)
    c.drawString(margin_x, 13.2 * mm, DISCLAIMER_TEKST)

    c.showPage()
    c.save()

    return out_path.as_posix()
