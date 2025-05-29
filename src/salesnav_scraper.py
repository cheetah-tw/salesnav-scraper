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
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, relative_path)


def main():
    # ——————————
    # CONFIGURATION
    # ——————————
    EXCEL_INPUT = resource_path("links.xlsx")
    OUTPUT_CSV = "salesnav_prospects.csv"
    OUTPUT_XLSX = "salesnav_prospects.xlsx"
    LOAD_TIMEOUT = 10
    PAGE_TIMEOUT = 15
    DELAY_BETWEEN = 2

    # Load credentials from cred.env (at repo root)
    dotenv_path = resource_path("cred.env")
    if not os.path.exists(dotenv_path):
        raise FileNotFoundError(f"Credentials file not found: {dotenv_path}")
    load_dotenv(dotenv_path)
    EMAIL = os.getenv("LINKEDIN_EMAIL")
    PASSWORD = os.getenv("LINKEDIN_PASSWORD")
    if not EMAIL or not PASSWORD:
        raise ValueError("LINKEDIN_EMAIL and LINKEDIN_PASSWORD must be set in cred.env")

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

    # Load links from data/links.xlsx
    df_links = pd.read_excel(EXCEL_INPUT, header=0)
    links = df_links.iloc[:, 0].astype(str).tolist()

    results = []

    # Scrape each profile
    for url in links:
        val = url.strip()
        print(f"Processing: {val}")

        if val.lower() == "no prospect linkedin":
            results.append(
                {
                    "Full Name": val,
                    "Profile URL": val,
                    "Current Title": val,
                    "Company": val,
                    "Company Link": val,
                }
            )
            time.sleep(DELAY_BETWEEN)
            continue

        try:
            driver.get(val)
        except (TimeoutException, WebDriverException) as e:
            print(f"  → Page load failed: {e}")
            results.append(
                {
                    "Full Name": val,
                    "Profile URL": val,
                    "Current Title": "Page Load Timeout",
                    "Company": "Page Load Timeout",
                    "Company Link": "",
                }
            )
            time.sleep(DELAY_BETWEEN)
            continue

        # ——————————
        # FULL NAME
        # ——————————
        try:
            name_elem = wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "h1[data-anonymize='person-name']")
                )
            )
            full_name = name_elem.text.strip()
        except Exception:
            full_name = "No name found"

        # ——————————
        # CURRENT TITLE
        # ——————————
        try:
            title_elem = wait.until(
                EC.visibility_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "div[data-sn-view-name='lead-current-role'] "
                        "span[data-anonymize='job-title']",
                    )
                )
            )
            current_title = title_elem.text.strip() or "No title found"
        except Exception:
            current_title = "Error or Not Found"

        # ——————————
        # COMPANY NAME & LINK
        # ——————————
        try:
            comp_elem = wait.until(
                EC.visibility_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "div[data-sn-view-name='lead-current-role'] "
                        "a[data-anonymize='company-name']",
                    )
                )
            )
            company = comp_elem.text.strip()
            company_link = comp_elem.get_attribute("href")
        except Exception:
            company = "No company found"
            company_link = ""

        results.append(
            {
                "Full Name": full_name,
                "Profile URL": val,
                "Current Title": current_title,
                "Company": company,
                "Company Link": company_link,
            }
        )

        time.sleep(DELAY_BETWEEN)

    driver.quit()

    # ——————————
    # WRITE & FILTER CSV
    # ——————————
    df_out = pd.DataFrame(
        results,
        columns=[
            "Full Name",
            "Profile URL",
            "Current Title",
            "Company",
            "Company Link",
        ],
    )
    # Drop timeout rows
    df_out = df_out[df_out["Current Title"] != "Page Load Timeout"]

    csv_path = resource_path(OUTPUT_CSV)
    df_out.to_csv(csv_path, index=False)
    print("Saved CSV to:", csv_path)

    # ——————————
    # CONVERT CSV → XLSX
    # ——————————
    df = pd.read_csv(csv_path)
    xlsx_path = resource_path(OUTPUT_XLSX)
    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    print("Saved Excel to:", xlsx_path)


if __name__ == "__main__":
    main()
