from __future__ import annotations

from pathlib import Path
from datetime import datetime
import re

import pythoncom
import win32com.client


def _week_folder(dt: datetime) -> str:
    iso_year, iso_week, _ = dt.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def _looks_like_html(s: str) -> bool:
    if not s:
        return False
    t = s.strip().lower()

    # duidelijke HTML signalen
    if "<html" in t or "<body" in t or "<div" in t or "<p" in t or "<br" in t:
        return True

    # als er meerdere tags voorkomen, beschouw als html
    tag_count = len(re.findall(r"</?[a-z][a-z0-9]*\b[^>]*>", t))
    return tag_count >= 3


def _attach_inline_logo(mail, logo_path: Path, cid: str = "klaasvis_logo") -> None:
    """
    Voegt logo toe als inline afbeelding voor Outlook HTMLBody:
    <img src="cid:klaasvis_logo">
    """
    if not logo_path.exists():
        return

    att = mail.Attachments.Add(str(logo_path.resolve()))
    pa = att.PropertyAccessor

    # PR_ATTACH_CONTENT_ID (string)
    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
    # PR_ATTACH_CONTENT_LOCATION (string) - helpt bij rendering
    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3713001F", cid)
    # PR_ATTACHMENT_HIDDEN (bool) - voorkomt losse bijlageweergave
    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)


def write_msg_outlook(
    out_base_dir: str,
    dt: datetime,
    offer_no: str,
    to_addr: str,
    subject: str,
    body_text: str,
    pdf_path: str | None = None,
) -> str:
    """
    Maakt een .msg bestand via Outlook COM en koppelt (optioneel) de PDF als bijlage.
    Ondersteunt HTMLBody + inline logo via CID.
    Returned: relative pad (posix) zodat UI hem kan openen via /file?path=...
    """
    pythoncom.CoInitialize()
    try:
        base = Path(out_base_dir)
        week_dir = base / _week_folder(dt)
        week_dir.mkdir(parents=True, exist_ok=True)

        msg_file = week_dir / f"{offer_no}.msg"

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem

        mail.To = to_addr
        mail.Subject = subject or ""

        is_html = _looks_like_html(body_text)

        if is_html:
            # Belangrijk: zet BodyFormat op HTML
            mail.BodyFormat = 2  # 2 = olFormatHTML

            # Inline logo attachen (zorg dat je HTML: <img src="cid:klaasvis_logo"> gebruikt)
            project_root = Path(__file__).resolve().parent
            logo_path = project_root / "assets" / "logo_klaasvis.png"
            _attach_inline_logo(mail, logo_path, cid="klaasvis_logo")

            # Outlook verwacht complete HTML
            t = (body_text or "").strip()
            if "<html" not in t.lower():
                t = f"<html><body>{t}</body></html>"

            mail.HTMLBody = t
        else:
            mail.Body = body_text or ""

        # PDF bijlage
        if pdf_path:
            p = Path(pdf_path)
            if p.exists():
                mail.Attachments.Add(str(p.resolve()))

        # Save as .msg (3 = olMSG)
        mail.SaveAs(str(msg_file.resolve()), 3)

        return msg_file.as_posix()

    finally:
        pythoncom.CoUninitialize()
