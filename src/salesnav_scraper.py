import time
import os
import sys
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
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


def get_chrome_driver_path():
    bundled = resource_path("chromedriver.exe")
    # 1) if we’re frozen and you’ve bundled a driver, use it
    if getattr(sys, "frozen", False) and os.path.exists(bundled):
        return bundled
    # 2) otherwise let webdriver‑manager fetch the matching one
    return ChromeDriverManager().install()


def main():
    # ——————————
    # CONFIGURATION
    # ——————————
    EXCEL_INPUT = resource_path("links.xlsx")
    OUTPUT_CSV_LONG = "salesnav_prospects_long.csv"
    OUTPUT_XLSX_LONG = "salesnav_prospects_long.xlsx"
    OUTPUT_CSV_WIDE = "salesnav_prospects_wide.csv"
    OUTPUT_XLSX_WIDE = "salesnav_prospects_wide.xlsx"
    LOAD_TIMEOUT = 10
    PAGE_TIMEOUT = 15
    DELAY_BETWEEN = 2

    # Load credentials from cred.env
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
    driver_path = get_chrome_driver_path()
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(PAGE_TIMEOUT)
    wait = WebDriverWait(driver, LOAD_TIMEOUT)

    # Login
    driver.get("https://www.linkedin.com/login")
    time.sleep(2)
    driver.find_element(By.ID, "username").send_keys(EMAIL)
    driver.find_element(By.ID, "password").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    time.sleep(3)

    # Load profile URLs
    df_links = pd.read_excel(EXCEL_INPUT, header=0)
    links = df_links.iloc[:, 0].astype(str).tolist()

    long_results = []

    # Scrape each profile in original order
    for idx, url in enumerate(links):
        val = url.strip()
        print(f"Processing ({idx}): {val}")

        # Handle explicit "no prospect" entries
        if val.lower() == "no prospect linkedin":
            long_results.append(
                {
                    "ScanOrder": idx,
                    "Full Name": val,
                    "Profile URL": val,
                    "Title": val,
                    "Company": val,
                    "Link": val,
                }
            )
            time.sleep(DELAY_BETWEEN)
            continue

        # Load the page
        try:
            driver.get(val)
        except (TimeoutException, WebDriverException) as e:
            print(f"  → Page load failed: {e}")
            long_results.append(
                {
                    "ScanOrder": idx,
                    "Full Name": val,
                    "Profile URL": val,
                    "Title": "Page Load Timeout",
                    "Company": "Page Load Timeout",
                    "Link": "",
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
        # SCRAPE ALL CURRENT ROLES
        # ——————————
        try:
            container = wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div[data-sn-view-name='lead-current-role']")
                )
            )
            # grab all titles
            titles = container.find_elements(
                By.CSS_SELECTOR, "span[data-anonymize='job-title']"
            )
            # grab all company names, whether link or plain text
            companies = container.find_elements(
                By.CSS_SELECTOR, "[data-anonymize='company-name']"
            )
        except Exception:
            titles = []
            companies = []

        # emit at least one row
        if not titles:
            long_results.append(
                {
                    "ScanOrder": idx,
                    "Full Name": full_name,
                    "Profile URL": val,
                    "Title": "No title found",
                    "Company": "No company found",
                    "Link": "",
                }
            )
        else:
            for i, t in enumerate(titles):
                job = t.text.strip() or "No title found"
                # if a matching company element exists, take its text
                if i < len(companies):
                    c = companies[i]
                    comp_name = c.text.strip() or ""
                    # get href if it's an <a>, else empty
                    comp_link = c.get_attribute("href") or ""
                else:
                    comp_name = ""
                    comp_link = ""

                long_results.append(
                    {
                        "ScanOrder": idx,
                        "Full Name": full_name,
                        "Profile URL": val,
                        "Title": job,
                        "Company": comp_name,
                        "Link": comp_link,
                    }
                )

        time.sleep(DELAY_BETWEEN)

    driver.quit()

    # ——————————
    # SAVE LONG FORM
    # ——————————
    df_long = pd.DataFrame(long_results)
    df_long = df_long[df_long["Title"] != "Page Load Timeout"]

    df_long.to_csv(resource_path(OUTPUT_CSV_LONG), index=False)
    df_long.to_excel(resource_path(OUTPUT_XLSX_LONG), index=False, engine="openpyxl")
    print("Saved long-form CSV & XLSX.")

    # ——————————
    # PIVOT TO WIDE FORM (preserves scan order)
    # ——————————
    grouped = (
        df_long.groupby(
            ["ScanOrder", "Full Name", "Profile URL"], sort=False, dropna=False
        )
        .agg({"Title": list, "Company": list, "Link": list})
        .reset_index()
    )

    max_roles = grouped["Title"].map(len).max()

    for i in range(max_roles):
        grouped[f"Title_{i+1}"] = grouped["Title"].apply(
            lambda L: L[i] if i < len(L) else ""
        )
        grouped[f"Company_{i+1}"] = grouped["Company"].apply(
            lambda L: L[i] if i < len(L) else ""
        )
        grouped[f"Link_{i+1}"] = grouped["Link"].apply(
            lambda L: L[i] if i < len(L) else ""
        )

    df_wide = grouped.sort_values("ScanOrder").drop(
        columns=["ScanOrder", "Title", "Company", "Link"]
    )

    df_wide.to_csv(resource_path(OUTPUT_CSV_WIDE), index=False)
    df_wide.to_excel(resource_path(OUTPUT_XLSX_WIDE), index=False, engine="openpyxl")
    print("Saved wide-form CSV & XLSX.")


if __name__ == "__main__":
    main()
