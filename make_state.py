from playwright.sync_api import sync_playwright

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()

        page = context.new_page()
        page.goto("https://www.asrcockpit.nl/", wait_until="domcontentloaded")

        print("\n✅ Log nu handmatig in in het geopende venster.")
        print("   Zodra je helemaal ingelogd bent (dashboard zichtbaar), druk je hier ENTER.\n")
        input()

        context.storage_state(path="state.json")
        print("✅ state.json opgeslagen.")

        browser.close()

if __name__ == "__main__":
    main()
