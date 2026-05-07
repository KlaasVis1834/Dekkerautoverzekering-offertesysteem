from playwright.sync_api import sync_playwright

LOGIN_URL = "https://login.asrcockpit.nl"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.goto(LOGIN_URL)
    print("➡️ Log nu handmatig in (zonder dat je hier wachtwoord deelt).")
    print("✅ Als je volledig bent ingelogd en je Cockpit omgeving ziet, druk dan hier Enter...")
    input()

    context.storage_state(path="cockpit_state.json")
    browser.close()
    print("✅ opgeslagen: cockpit_state.json")
