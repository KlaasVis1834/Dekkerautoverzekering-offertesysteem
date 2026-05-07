from __future__ import annotations

from pathlib import Path
from datetime import datetime

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

from mailgen import guess_aanhef_en_achternaam  # ✅ toegevoegd


# ----------------- helpers -----------------

def _week_folder(dt: datetime) -> str:
    iso_year, iso_week, _ = dt.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def _ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _assets_path() -> Path:
    return Path(__file__).resolve().parent / "assets"


def _safe(v) -> str:
    return ("" if v is None else str(v)).strip()


def _draw_image_fit(c, img_path: Path, x, y, max_w, max_h):
    if not img_path.exists():
        return
    img = ImageReader(str(img_path))
    iw, ih = img.getSize()
    scale = min(max_w / iw, max_h / ih)
    c.drawImage(
        str(img_path),
        x,
        y,
        width=iw * scale,
        height=ih * scale,
        mask="auto",
    )


def _wrap_text(c, text, x, y, max_w, font="Helvetica", size=10.5, leading=5.6 * mm):
    c.setFont(font, size)
    words = (text or "").split()
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, font, size) <= max_w:
            line = test
        else:
            c.drawString(x, y, line)
            y -= leading
            line = w
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def _auto_merk_model(auto: str) -> str:
    parts = _safe(auto).split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[1]}"
    return _safe(auto)


# ----------------- main -----------------

def generate_post_letter_pdf(
    out_base_dir: str,
    dt: datetime,
    offer_no: str,
    klantnaam: str,
    adres: str,
    postcode: str,
    plaats: str,
    auto: str,
    behandeld_door: str = "",
) -> str:
    base = Path(out_base_dir)
    week_dir = base / _week_folder(dt)
    _ensure_dir(week_dir)

    out_path = week_dir / f"{offer_no}_postbrief.pdf"

    assets = _assets_path()
    logo_klaasvis = assets / "logo_klaasvis.png"

    c = canvas.Canvas(str(out_path), pagesize=A4)
    w, h = A4

    margin_x = 18 * mm

    # ---------------- Header ----------------
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#1F2A44"))
    c.drawString(margin_x, h - 22 * mm, "Verzekeringsvoorstel Dekkerautoverzekering")

    c.setFont("Helvetica", 9.5)
    c.setFillColor(colors.HexColor("#555555"))
    c.drawString(margin_x, h - 28 * mm, f"Offertenummer: {offer_no}")

    # ---------------- Adresblok (lager + witruimte) ----------------
    addr_y = h - 62 * mm
    c.setFont("Helvetica", 10.5)
    c.setFillColor(colors.black)

    c.drawString(margin_x, addr_y, _safe(klantnaam))
    c.drawString(margin_x, addr_y - 5.8 * mm, _safe(adres))
    c.drawString(
        margin_x,
        addr_y - 11.6 * mm,
        f"{_safe(postcode)}  {_safe(plaats)}".strip(),
    )

    # Datum rechts
    c.setFont("Helvetica", 9.5)
    c.setFillColor(colors.HexColor("#555555"))
    c.drawRightString(w - margin_x, addr_y, dt.strftime("%d-%m-%Y"))

    # Extra witruimte voor envelopvenster
    y = addr_y - 34 * mm

    # ---------------- Briefinhoud (zoals e-mail) ----------------
    auto_basic = _auto_merk_model(auto)

    # ✅ toegevoegd: aanhef + achternaam bepalen
    aanhef, achternaam = guess_aanhef_en_achternaam(klantnaam)

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10.5)

    y = _wrap_text(
        c,
        f"Geachte {aanhef} {achternaam},",
        margin_x,
        y,
        w - 2 * margin_x,
    )
    y -= 2 * mm

    y = _wrap_text(
        c,
        f"Van harte gefeliciteerd met de aankoop van uw {auto_basic} bij Dekkerautogroep B.V.",
        margin_x,
        y,
        w - 2 * margin_x,
    )
    y -= 2 * mm

    y = _wrap_text(
        c,
        "Om haar klanten optimale service te verlenen, heeft Dekkerautogroep B.V. in samenwerking met "
        "Assurantiekantoor Klaas Vis de Dekkerautoverzekering ontwikkeld. "
        "Bijgaand doen wij u een aanbieding toekomen van deze speciaal ontwikkelde en hoogwaardige verzekering.",
        margin_x,
        y,
        w - 2 * margin_x,
    )
    y -= 2 * mm

    y = _wrap_text(
        c,
        "In de bijgevoegde offerte vindt u een overzicht van de dekking en de bijbehorende premie.",
        margin_x,
        y,
        w - 2 * margin_x,
    )
    y -= 5 * mm

    y = _wrap_text(
        c,
        "Wij zullen u, naar aanleiding van dit verzekeringsvoorstel, telefonisch benaderen om te informeren "
        "of alles duidelijk is en om eventuele vragen met u door te nemen. "
        "Uiteraard kunt u ook altijd zelf contact met ons opnemen.",
        margin_x,
        y,
        w - 2 * margin_x,
    )

    y -= 10 * mm
    c.drawString(margin_x, y, "Met vriendelijke groet,")
    y -= 18 * mm

    # ---------------- Logo + bedrijfsgegevens (groter logo) ----------------
    logo_w = 65 * mm
    logo_h = 35 * mm

    _draw_image_fit(
        c,
        logo_klaasvis,
        margin_x,
        y - logo_h + 3 * mm,
        max_w=logo_w,
        max_h=logo_h,
    )

    text_x = margin_x
    ty = y - logo_h - 4 * mm

    c.setFont("Helvetica-Bold", 10.2)
    c.setFillColor(colors.black)
    c.drawString(text_x, ty, "Assurantiekantoor Klaas Vis")
    ty -= 5.6 * mm

    c.setFont("Helvetica", 9.8)
    c.setFillColor(colors.HexColor("#222222"))
    c.drawString(text_x, ty, "Zuiderweg 7")
    ty -= 5.2 * mm
    c.drawString(text_x, ty, "1456 NC Wijdewormer")
    ty -= 5.2 * mm
    c.drawString(text_x, ty, "Telefoon: (075) 631 42 61")
    ty -= 5.2 * mm
    c.drawString(text_x, ty, "E-mail: info@klaasvis.nl")
    ty -= 5.2 * mm
    c.drawString(text_x, ty, "Website: www.klaasvis.nl")

    # ---------------- Footer ----------------
    c.setStrokeColor(colors.HexColor("#D9DEE7"))
    c.setLineWidth(1)
    c.line(margin_x, 22 * mm, w - margin_x, 22 * mm)

    c.setFont("Helvetica", 8.8)
    c.setFillColor(colors.HexColor("#666666"))
    c.drawString(margin_x, 16 * mm, "Bijlage: verzekeringsvoorstel")

    c.showPage()
    c.save()

    return out_path.as_posix()
