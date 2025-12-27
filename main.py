import os
import time
import re
import shutil
import zipfile
import subprocess
import traceback
import unicodedata
import random
import requests
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
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# -----------------------------
# CONFIGURATION
# -----------------------------
print("🚀 Initializing configuration...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--headless=chrome")  # more stable on CI
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 25)
driver.set_page_load_timeout(40)
print("✅ WebDriver initialized.")

PDF_PAGE_LIMIT = 10

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def clean_extracted_text(text):
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"Page\s*\d+\s*/\s*\d+", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[\u0000-\u001f]+", "", text)
    cleaned_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    pretty = "\n".join(cleaned_lines)
    pretty = re.sub(r"\n{3,}", "\n\n", pretty)
    return pretty.strip()


def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        page_count = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(page_count):
            text += doc[i].get_text("text") + "\n"
        doc.close()
    except Exception:
        text = ""
    if len(text.strip()) < 50:
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            for page_image in pages:
                text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
        except Exception as e:
            print(f"⚠️ OCR failed for {file_path}: {e}")
    return clean_extracted_text(text)


def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return clean_extracted_text(text)
    except Exception:
        return ""


def extract_text_from_doc(file_path):
    try:
        process = subprocess.Popen(["antiword", file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        text = stdout.decode("utf-8", errors="ignore")
        return clean_extracted_text(text)
    except Exception as e:
        print(f"⚠️ Antiword failed for {file_path}: {e}")
        return ""


def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"⚠️ Failed to unzip {file_path}: {e}")
        return None


def clear_download_directory():
    for item in os.listdir(download_dir):
        path = os.path.join(download_dir, item)
        try:
            if os.path.isfile(path) or os.path.islink(path):
                os.unlink(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except Exception as e:
            print(f"⚠️ Failed to delete {path}: {e}")


def wait_for_download_complete(timeout=90):
    seconds = 0
    while seconds < timeout:
        downloading = any(f.endswith(".crdownload") for f in os.listdir(download_dir))
        if not downloading:
            files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload")]
            if files:
                return os.path.join(download_dir, files[0])
        time.sleep(1)
        seconds += 1
    return None

# -----------------------------
# MAIN SCRIPT
# -----------------------------
all_processed_tenders = []

try:
    print("\n--- Starting scraping ---")
    driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
    time.sleep(2)

    # Step 1: Open "Définir" popup
    define_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_domaineActivite_linkDisplay")))
    define_btn.click()
    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])

    # Step 2: Select Services checkbox and validate
    checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_repeaterCategorie_ctl2_idCategorie")))
    checkbox.click()
    validate_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
    validate_btn.click()
    driver.switch_to.window(driver.window_handles[0])
    print("✅ Services selected.")

    # Step 3: Set date filter to yesterday
    date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    date_input.clear()
    for char in yesterday:
        date_input.send_keys(char)
        time.sleep(random.uniform(0.05, 0.15))
    search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_button)
    time.sleep(0.8)
    search_button.click()
    time.sleep(2)

    # Step 4: Set results per page
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
    Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
    time.sleep(2)

    # Step 5: Scrape table
    rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
    data = []
    for row in rows:
        try:
            ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
            objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
            buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
            lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
            deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
            first_button = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
            data.append({
                "reference": ref,
                "objet": objet,
                "acheteur": buyer,
                "lieux_execution": lieux,
                "date_limite": deadline,
                "first_button_url": first_button
            })
        except Exception as e:
            print(f"⚠️ Error extracting row: {e}")

    df = pd.DataFrame(data)
        # Save Excel
    output_file = "marches_publics_services_2020_to_now.xlsx"
    df.to_excel(output_file, index=False)
    print(f"✅ Excel saved: {output_file}")

except Exception as e:
    print("❌ Error during execution:")
    print(e)

finally:
    try:
        driver.quit()
    except:
        pass


    
