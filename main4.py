import os
import time
import re
import shutil
import zipfile
import subprocess
import traceback
import unicodedata
import random
import pandas as pd
from datetime import datetime, timedelta

# PDF / OCR / DOC
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

# -----------------------------
# CONFIGURATION
# -----------------------------
print("🚀 Initializing configuration...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--headless=chrome")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

# ✅ CRITICAL FIX: Add User-Agent to look like a real browser
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30) # Increased to 30s
driver.set_page_load_timeout(60)
print("✅ WebDriver initialized.")

# -------- CONFIG --------
EXCEL_FILE = "URLS.xlsx"
URL_COLUMN = "PV"
RESULT_COLUMN = "Entreprise"   # new column
# ------------------------

# Read Excel
df = pd.read_excel(EXCEL_FILE)

# Create result column if it doesn't exist
if RESULT_COLUMN not in df.columns:
    df[RESULT_COLUMN] = None


for index, row in df.iterrows():
    url = row[URL_COLUMN]

    # Skip empty URLs
    if pd.isna(url):
        continue

    try:
        print(f"Opening: {url}")
        driver.get(str(url))
        time.sleep(2)

        ent = (
            driver
            .find_element(By.CLASS_NAME, "table-results")
            .text
            .strip()
            .replace('\n', ' - ')
        )

        df.at[index, RESULT_COLUMN] = ent

    except Exception as e:
        print(f"No entreprise found for {url}")
        df.at[index, RESULT_COLUMN] = None

output_file = "tender_results.xlsx"
df.to_excel(output_file, index=False)

print("✅ Scraping finished")
print(f"📦 Artifact ready: {output_file}")
