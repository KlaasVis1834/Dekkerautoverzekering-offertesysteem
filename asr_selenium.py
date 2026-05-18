from __future__ import annotations

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


ASR_COCKPIT_URL = "https://www.asrcockpit.nl/"


def open_asr_browser():
    options = webdriver.ChromeOptions()
    options.add_argument(r"--user-data-dir=C:\dekker-selenium\chrome-profile")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    driver.get(ASR_COCKPIT_URL)
    return driver


if __name__ == "__main__":
    driver = open_asr_browser()
    input("Log handmatig in bij ASR Cockpit. Druk daarna op Enter...")
