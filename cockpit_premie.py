from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Optional

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


@dataclass
class PremieResultaat:
    premie: Optional[float] = None                 # totaal incl assurantiebelasting (maand)
    assurantiebelasting: Optional[float] = None     # maand
    eindpremie_excl: Optional[float] = None         # maand excl assurantiebelasting
    raw_fields: Dict[str, Any] = field(default_factory=dict)
    landed_url: str = ""


def _extract_premievelden_from_json(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Zoekt in applicationForm.applicationFormQuestionList naar velden met premie.
    Geeft dict terug: {questionRef.externalIdentifier: value}
    """
    out: Dict[str, Any] = {}

    app_form = payload.get("applicationForm") or {}
    qlist = app_form.get("applicationFormQuestionList") or []
    if not isinstance(qlist, list):
        return out

    for item in qlist:
        if not isinstance(item, dict):
            continue
        qref = (item.get("questionRef") or {})
        key = qref.get("externalIdentifier")
        if not key:
            continue
        out[str(key)] = item.get("value")

    return out


def _pick_best(fields: Dict[str, Any]) -> PremieResultaat:
    """
    Probeert de beste maandpremie-set te kiezen.
    Jij zag o.a.:
      - finTotaalpremieBd
      - finAssurantiebelastingBd
      - finEindpremieBd
    (Zakelijk) en soms finZakelijkTotaalpremieBd etc.
    """
    def f(key: str) -> Optional[float]:
        v = fields.get(key)
        try:
            return None if v is None else float(v)
        except Exception:
            return None

    # voorkeur: “normale” keys op S3
    premie = f("finTotaalpremieBd") or f("finZakelijkTotaalpremieBd")
    assur = f("finAssurantiebelastingBd") or f("finZakelijkAssurantiebelastingBd")
    eind_excl = f("finEindpremieBd") or f("finZakelijkEindpremieBd")

    return PremieResultaat(
        premie=premie,
        assurantiebelasting=assur,
        eindpremie_excl=eind_excl,
        raw_fields=fields,
    )


def haal_premie_uit_cockpit(
    url: str,
    storage_state_path: str = "state.json",
    headless: bool = True,
    timeout_ms: int = 45000,
) -> PremieResultaat:
    """
    Navigeert naar de QIS/Axon S3 url en onderschept de process-action JSON response.
    Werkt het meest stabiel als state.json geldig is.
    """
    result = PremieResultaat()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=storage_state_path)
        page = context.new_page()

        captured: Dict[str, Any] = {}

        def on_response(resp):
            nonlocal captured
            try:
                u = resp.url or ""
                if "policy-request-forms/process-action" in u and resp.status == 200:
                    ct = (resp.headers.get("content-type") or "").lower()
                    if "application/json" in ct:
                        data = resp.json()
                        if isinstance(data, dict):
                            # haal velden eruit
                            fields = _extract_premievelden_from_json(data)
                            # alleen opslaan als we echt premie-achtige keys zien
                            if any(k.startswith("fin") or "Premie" in k or "premie" in k for k in fields.keys()):
                                captured = fields
            except Exception:
                pass

        page.on("response", on_response)

        # extra debug: waar landen we?
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        result.landed_url = page.url

        # Wacht even op XHRs / process-action
        # (Soms moet je scrollen of iets triggert pas na “render”. Daarom: korte wait-loop)
        try:
            page.wait_for_timeout(1500)
            # Als er nog niets is, wachten we met networkidle (kan soms nooit komen) -> daarom try/except
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except PlaywrightTimeoutError:
                pass

            # nog even extra tijd om responses binnen te krijgen
            page.wait_for_timeout(2500)
        finally:
            pass

        if captured:
            pr = _pick_best(captured)
            pr.landed_url = result.landed_url
            browser.close()
            return pr

        # Niets gevangen
        browser.close()
        return result
