import os
import sys
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

BATCH_SIZE = 1000        # URLs per run
MAX_BATCHES = 8         # Stop after 5 batches (500 URLs)
BATCH_COUNTER_FILE = "batch_counter.txt"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

print("🚀 Starting batch scraping (bs4)")

# -----------------------------
# LOAD / INIT BATCH COUNTER
# -----------------------------
if os.path.exists(BATCH_COUNTER_FILE):
    with open(BATCH_COUNTER_FILE, "r") as f:
        batch_count = int(f.read().strip())
else:
    batch_count = 0

if batch_count >= MAX_BATCHES:
    print("🛑 Max batches reached. Stopping.")
    sys.exit(0)

print(f"🔁 Running batch {batch_count + 1} / {MAX_BATCHES}")

# -----------------------------
# LOAD EXCEL
# -----------------------------
df = pd.read_excel(EXCEL_FILE)
df = df.reset_index(drop=True)

if RESULT_COLUMN not in df.columns:
    df[RESULT_COLUMN] = None

# Find first unprocessed row
start_index = df[df[RESULT_COLUMN].isna()].index.min()

if pd.isna(start_index):
    print("✅ All URLs already processed.")
    sys.exit(0)

end_index = min(start_index + BATCH_SIZE, len(df))
print(f"📊 Processing rows {start_index + 1} → {end_index}")

# -----------------------------
# SCRAPING LOOP (BS4)
# -----------------------------
for index in range(start_index, end_index):
    url = df.at[index, URL_COLUMN]

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

# -----------------------------
# UPDATE BATCH COUNTER
# -----------------------------
batch_count += 1
with open(BATCH_COUNTER_FILE, "w") as f:
    f.write(str(batch_count))

print("✅ Batch completed successfully")
print(f"📦 Progress saved to {EXCEL_FILE}")
