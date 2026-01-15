import os
import time
import pandas as pd
import requests
from bs4 import BeautifulSoup

# -----------------------------
# CONFIGURATION
# -----------------------------
EXCEL_FILE = "URLS.xlsx"
URL_COLUMN = "PV"
RESULT_COLUMN = "Entreprise"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

print("🚀 Starting scraping (bs4, no batches)")

# -----------------------------
# LOAD EXCEL
# -----------------------------
df = pd.read_excel(EXCEL_FILE)
df = df.reset_index(drop=True)

if RESULT_COLUMN not in df.columns:
    df[RESULT_COLUMN] = None

# -----------------------------
# SCRAPING LOOP
# -----------------------------
for index, row in df.iterrows():
    if pd.notna(row[RESULT_COLUMN]):
        continue  # already processed

    url = row[URL_COLUMN]
    if pd.isna(url):
        continue

    try:
        print(f"[{index + 1}] Fetching: {url}")

        response = requests.get(url, headers=HEADERS, timeout=15)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find("table", class_="table-results")

        if table:
            df.at[index, RESULT_COLUMN] = table.get_text(
                separator=" - ", strip=True
            )
        else:
            df.at[index, RESULT_COLUMN] = None

    except Exception as e:
        print(f"❌ Failed: {url}")
        df.at[index, RESULT_COLUMN] = None

    # SAVE AFTER EACH URL (CRASH SAFE)
    df.to_excel(EXCEL_FILE, index=False)

    # polite delay
    time.sleep(0.5)

print("✅ Scraping finished")
print(f"📦 Results saved to {EXCEL_FILE}")
