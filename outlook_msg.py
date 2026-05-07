from pathlib import Path

def write_msg_outlook(
    out_base_dir,
    dt,
    offer_no,
    to_addr,
    subject,
    body_text,
    pdf_path,
):
    """
    Tijdelijke Render/Linux fallback.
    Maakt dummy .txt bestand i.p.v. Outlook .msg
    """

    out_dir = Path(out_base_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = f"{offer_no}.txt"
    out_path = out_dir / filename

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"TO: {to_addr}\n")
        f.write(f"SUBJECT: {subject}\n\n")
        f.write(body_text)

    return str(out_path)
