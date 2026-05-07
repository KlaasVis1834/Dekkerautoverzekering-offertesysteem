from playwright.sync_api import sync_playwright

START = "https://www.asrcockpit.nl/"

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(START, wait_until="domcontentloaded")

        print("Log in indien nodig.")
        print("Navigeer in de browser door tot je een URL ziet die begint met:")
        print("https://qis.asrcockpit.nl/")
        input("Druk ENTER om state.json op te slaan...")

        context.storage_state(path="state.json")
        print("✅ state.json opgeslagen.")
        browser.close()

if __name__ == "__main__":
    main()
