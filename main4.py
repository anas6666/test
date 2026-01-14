import os
import time
import pandas as pd
from datetime import datetime

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service

# -----------------------------
# CONFIGURATION
# -----------------------------
print("🚀 Initializing configuration...")

BASE_DIR = os.getcwd()
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads_temp")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

MAX_URLS = 100

EXCEL_FILE = "URLS.xlsx"
URL_COLUMN = "PV"
RESULT_COLUMN = "Entreprise"
OUTPUT_FILE = "tender_results.xlsx"

# -----------------------------
# SELENIUM SETUP
# -----------------------------
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)

driver.set_page_load_timeout(60)

print("✅ WebDriver ready")

# -----------------------------
# LOAD EXCEL
# -----------------------------
df = pd.read_excel(EXCEL_FILE)

# Reset index to avoid any weird jumps
df = df.reset_index(drop=True)

# Create result column if missing
if RESULT_COLUMN not in df.columns:
    df[RESULT_COLUMN] = None

total_rows = min(len(df), MAX_URLS)
print(f"📊 Processing {total_rows} URLs (MAX={MAX_URLS})")

# -----------------------------
# SCRAPING LOOP (HARD LIMITED)
# -----------------------------
processed = 0

for index in range(total_rows):
    url = df.iloc[index][URL_COLUMN]

    if pd.isna(url):
        continue

    try:
        print(f"[{index + 1}/{total_rows}] Opening: {url}")
        driver.get(str(url))

        element = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "table-results"))
        )

        ent = (
            element.text
            .strip()
            .replace("\n", " - ")
        )

        df.at[index, RESULT_COLUMN] = ent

    except Exception as e:
        print(f"❌ Failed for URL {url}")
        df.at[index, RESULT_COLUMN] = None

    processed += 1

    if processed >= MAX_URLS:
        print("🛑 HARD STOP reached (1000 URLs)")
        break

# -----------------------------
# CLEANUP & SAVE
# -----------------------------
driver.quit()

df.to_excel(OUTPUT_FILE, index=False)

print("✅ Scraping finished successfully")
print(f"📦 Artifact ready: {OUTPUT_FILE}")
