import json
import re
from playwright.sync_api import sync_playwright


S3_URL = (
    "https://qis.asrcockpit.nl/Axon/add-quotation/form/fill-form/screen/S3"
    "?formExtId=autoZakelijk"
    "&formDialogueExtId=offerte"
    "&formBehaviourExtId=%257B%2522prefix%2522%253A%2522A%2522%252C%2522agreementExtId%2522%253A%25221109%2522%257D"
    "&dateFormDefinition=2026-02-04"
    "&originExtId=Extranet"
)

# Welke velden willen we er sowieso uit vissen?
FIELD_WHITELIST = {
    # Zakelijk totaal/ass belasting
    "finZakelijkEindpremieBd",
    "finZakelijkAssurantiebelastingBd",
    "finZakelijkTotaalpremieBd",

    # Termijn-premies (vaak deze gebruiken voor maand)
    "finEindpremieBd",
    "finAssurantiebelastingBd",
    "finTotaalpremieBd",

    # Soms zit de “premie op scherm”
    "dlgAutoZakelijkWACascoAllriskPremieBd",
}


def extract_premies_from_process_action(payload: dict) -> dict:
    """
    payload = JSON response van /process-action
    zoekt in applicationFormQuestionList naar whitelisted fields
    """
    out = {}
    app_form = payload.get("applicationForm") or {}
    qlist = app_form.get("applicationFormQuestionList") or []
    if not isinstance(qlist, list):
        return out

    for q in qlist:
        if not isinstance(q, dict):
            continue
        qref = (q.get("questionRef") or {}).get("externalIdentifier")
        if not qref or qref not in FIELD_WHITELIST:
            continue
        out[qref] = q.get("value")
    return out


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="state.json")

        premies_found = {}

        def on_response(resp):
            nonlocal premies_found
            url = resp.url

            # We pakken specifiek de process-action call
            if "policy-request-forms/process-action" not in url:
                return

            try:
                if "application/json" not in (resp.headers.get("content-type", "")):
                    return
                data = resp.json()
                premies = extract_premies_from_process_action(data)
                if premies:
                    premies_found.update(premies)
            except Exception:
                pass

        context.on("response", on_response)

        page = context.new_page()
        page.goto(S3_URL, wait_until="networkidle")

        # Wacht even zodat eventuele XHR/process-action binnenkomt
        page.wait_for_timeout(4000)

        print("\n✅ Premievelden gevonden uit process-action:")
        if not premies_found:
            print(" - (niets gevonden) -> waarschijnlijk wordt process-action nog niet getriggerd op dit moment.")
            print("   Tip: klik 1x op 'Volgende' of verander een veld, dan triggert ie vaak.")
        else:
            for k in sorted(premies_found.keys()):
                print(f" - {k}: {premies_found[k]}")

        input("\nDruk ENTER om te sluiten...")
        browser.close()


if __name__ == "__main__":
    main()
