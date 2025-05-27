import time
import os
import sys
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException


def resource_path(relative_path):
    """
    Returns the path to a resource, preferring an external file next to the .exe when frozen.
    """
    if getattr(sys, "frozen", False):
        # We're in a PyInstaller bundle
        base_dir = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, relative_path)


def main():
    # ——————————
    # CONFIGURATION
    # ——————————
    EXCEL_INPUT = resource_path("links.xlsx")
    OUTPUT_CSV = "salesnav_titles.csv"
    OUTPUT_XLSX = "salesnav_titles.xlsx"
    LINKS_COLUMN = None
    LOAD_TIMEOUT = 10
    PAGE_TIMEOUT = 15
    DELAY_BETWEEN = 2

    # Load credentials
    load_dotenv(resource_path("cred.env"))
    EMAIL = os.getenv("LINKEDIN_EMAIL")
    PASSWORD = os.getenv("LINKEDIN_PASSWORD")

    # Selenium setup
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(PAGE_TIMEOUT)
    wait = WebDriverWait(driver, LOAD_TIMEOUT)

    # Login
    driver.get("https://www.linkedin.com/login")
    time.sleep(2)
    driver.find_element(By.ID, "username").send_keys(EMAIL)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    time.sleep(3)

    # Load links
    df_links = pd.read_excel(EXCEL_INPUT, header=0)
    links = (
        df_links.iloc[:, 0].astype(str).tolist()
        if LINKS_COLUMN is None
        else df_links[LINKS_COLUMN].astype(str).tolist()
    )

    results = []

    # Scrape
    for url in links:
        val = url.strip()
        print(f"Processing: {val}")

        if val.lower() == "no prospect linkedin":
            results.append({"Profile URL": val, "Current Title": val})
            time.sleep(DELAY_BETWEEN)
            continue

        try:
            driver.get(val)
        except (TimeoutException, WebDriverException) as e:
            print(f"  → Page load failed: {e}")
            results.append({"Profile URL": val, "Current Title": "Page Load Timeout"})
            time.sleep(DELAY_BETWEEN)
            continue

        try:
            elem = wait.until(
                EC.visibility_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "div[data-sn-view-name='lead-current-role'] span[data-anonymize='job-title']",
                    )
                )
            )
            title = elem.text.strip() or "No title found"
            print(f"  → Found title: {title}")
        except Exception as e:
            print(f"  → Title not found: {e}")
            title = "Error or Not Found"

        results.append({"Profile URL": val, "Current Title": title})
        time.sleep(DELAY_BETWEEN)

    driver.quit()

    # ——————————
    # FILTER & WRITE OUT CSV
    # ——————————
    df_out = pd.DataFrame(results, columns=["Profile URL", "Current Title"])
    # remove any timeouts
    df_out = df_out[df_out["Current Title"] != "Page Load Timeout"]

    csv_path = resource_path(OUTPUT_CSV)
    df_out.to_csv(csv_path, index=False)
    print("Saved CSV to:", csv_path)

    # ——————————
    # CONVERT CSV → XLSX
    # ——————————
    df = pd.read_csv(csv_path)
    df = df[["Profile URL", "Current Title"]]
    df.columns = ["URL", "Position Title"]
    xlsx_path = resource_path(OUTPUT_XLSX)
    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    print("Saved Excel to:", xlsx_path)


if __name__ == "__main__":
    main()
