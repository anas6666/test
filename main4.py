import os
import sys
import pandas as pd

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

EXCEL_FILE = "URLS.xlsx"
URL_COLUMN = "PV"
RESULT_COLUMN = "Entreprise"

BATCH_SIZE = 50        # URLs per run
MAX_BATCHES = 2         # 5 batches = 500 URLs max
BATCH_COUNTER_FILE = "batch_counter.txt"

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
# SELENIUM SETUP
# -----------------------------
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--blink-settings=imagesEnabled=false")
options.add_argument("--disable-blink-features=AutomationControlled")

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

# -----------------------------
# SCRAPING LOOP
# -----------------------------
for index in range(start_index, end_index):
    url = df.at[index, URL_COLUMN]

    if pd.isna(url):
        continue

    try:
        print(f"[{index + 1}] Opening: {url}")
        driver.get(str(url))

        element = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "table-results"))
        )

        df.at[index, RESULT_COLUMN] = element.text.strip().replace("\n", " - ")

    except Exception:
        df.at[index, RESULT_COLUMN] = None

    # SAVE AFTER EACH URL (CRASH SAFE)
    df.to_excel(EXCEL_FILE, index=False)

driver.quit()

# -----------------------------
# UPDATE BATCH COUNTER
# -----------------------------
batch_count += 1
with open(BATCH_COUNTER_FILE, "w") as f:
    f.write(str(batch_count))

print("✅ Batch finished successfully")
print(f"📦 Progress saved to {EXCEL_FILE}")
