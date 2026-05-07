from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)

    # ✅ storage_state hoort hier
    context = browser.new_context(storage_state="state.json")

    page = context.new_page()
    page.goto("https://www.asrcockpit.nl/", wait_until="networkidle")

    print("Landed on:", page.url)

    page.wait_for_timeout(5000)
    browser.close()
