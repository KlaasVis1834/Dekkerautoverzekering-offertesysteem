import json
from typing import Any, Dict, Optional
from playwright.sync_api import sync_playwright

START_URL = "https://www.asrcockpit.nl/"
STATE_FILE = "cockpit_state.json"

# Zet hier een “diepe link” naar jouw offerte-scherm (S2 of S3) zodra je die hebt.
# Voor nu laten we hem op home starten. Later zetten we DEEPLINK naar jouw exacte offerte-url.
DEEPLINK: Optional[str] = None

PROCESS_ACTION_SUBSTR = "/Axon/services/rest/authenticated/policy-request-forms/process-action"


def find_premies(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Zoekt in applicationFormQuestionList naar bekende premievelden.
    Returnt dict met gevonden waarden.
    """
    out: Dict[str, Any] = {}
    app_form = payload.get("applicationForm") or {}
    qlist = app_form.get("applicationFormQuestionList") or []
    if not isinstance(qlist, list):
        return out

    # Key = questionRef.externalIdentifier, Value = value
    for q in qlist:
        if not isinstance(q, dict):
            continue
        qref = (q.get("questionRef") or {}).get("externalIdentifier")
        val = q.get("value")
        if isinstance(qref, str) and qref:
            # Pak alleen de relevante premievelden (breid gerust uit)
            if "premie" in qref.lower() or "assurantie" in qref.lower():
                out[qref] = val

    return out


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=STATE_FILE)
        page = context.new_page()

        # 1) Naar cockpit
        page.goto(START_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        if DEEPLINK:
            page.goto(DEEPLINK, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

        print("Landed on:", page.url)

        found: Dict[str, Any] = {}

        def on_response(resp):
            nonlocal found
            try:
                url = resp.url
                if PROCESS_ACTION_SUBSTR not in url:
                    return
                if resp.request.method != "POST":
                    return
                # Soms is content-type json, soms niet — toch proberen.
                txt = resp.text()
                if not txt or txt.strip()[0] not in "{[":
                    return
                payload = json.loads(txt)
                premies = find_premies(payload)
                if premies:
                    found = premies
                    print("\n✅ Premievelden gevonden uit process-action:")
                    for k, v in premies.items():
                        print(f" - {k}: {v}")
            except Exception:
                return

        page.on("response", on_response)

        print("\n➡️ Actie nodig:")
        print("Ga in de open browser naar het offerte-scherm en klik 1x op 'Volgende', 'Bereken' of iets dat de premie herberekent.")
        print("Zodra die POST wordt gedaan, print dit script de premievelden.\n")

        # Wachten tot je klaar bent (of premie gevonden)
        for _ in range(600):  # ~10 min bij 1s
            if found:
                break
            page.wait_for_timeout(1000)

        if not found:
            print("\n❌ Geen premievelden gezien.")
            print("Waarschijnlijk is de process-action call niet getriggerd tijdens deze run, of je zit op een scherm dat geen herberekening doet.")
            print("Tip: klik in Cockpit op een knop die van scherm wisselt (S2->S3) of wijzig 1 veld en klik 'Volgende'.")

        browser.close()


if __name__ == "__main__":
    main()
