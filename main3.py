import os
import time
import re
import shutil
import traceback
import unicodedata
import random
import pandas as pd
from datetime import datetime

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    StaleElementReferenceException
)

# -----------------------------
# CONFIGURATION
# -----------------------------
print("🚀 [INIT] Initializing configuration...")

# Create a temporary folder for downloads
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

# Chrome Options
options = webdriver.ChromeOptions()
options.add_argument("--headless=new") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

# Initialize Driver
service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)
driver.set_page_load_timeout(60)
print("✅ [INIT] WebDriver initialized.")

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def clear_download_directory():
    """Removes all files in the download directory to keep it clean."""
    if os.path.exists(download_dir):
        for filename in os.listdir(download_dir):
            file_path = os.path.join(download_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"⚠️ Failed to delete {file_path}. Reason: {e}")

# -----------------------------
# MAIN SCRIPT
# -----------------------------
all_processed_tenders = [] # List to store final data
metadata_list = [] # List to store initial data

try:
    print("\n--- [STEP 1] Starting Navigation ---")
    url = "https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons"
    driver.get(url)
    print(f"🌐 Page loaded. Title: {driver.title}")
    
    # Wait for body to ensure page isn't blank
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)

    # --- [STEP 2] Handle 'Définir' Popup ---
    print("--- [STEP 2] Handling Popup ---")
    try:
        define_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_domaineActivite_linkDisplay")))
        define_btn.click()
        print("ℹ️ 'Définir' button clicked.")
    except TimeoutException:
        print("❌ [CRITICAL] Could not find 'Définir' button. Taking screenshot.")
        driver.save_screenshot("error_step2_popup.png")
        raise

    # Switch to the popup window
    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])

    # Select 'Services' checkbox
    checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_repeaterCategorie_ctl2_idCategorie")))
    checkbox.click()
    
    # Click Validate
    validate_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
    validate_btn.click()
    
    # Switch back to main window
    driver.switch_to.window(driver.window_handles[0])
    print("✅ Services selected and popup closed.")
    time.sleep(1)

    # --- [STEP 3] Fill Search Form ---
    print("--- [STEP 3] Filling Search Form ---")
    yesterday = "01/01/2020"
    
    # Fill Date
    try:
        date_input = driver.find_element(By.NAME, "ctl0$CONTENU_PAGE$AdvancedSearch$dateMiseEnLigneStart")
        date_input.clear()
        date_input.send_keys(yesterday)
    except NoSuchElementException:
        try:
            date_input_2 = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
            date_input_2.clear()
            date_input_2.send_keys(yesterday)
        except:
            print("⚠️ Date input not found, skipping date filter.")

    # Fill Keyword
    try:
        input_field = driver.find_element(By.NAME, "ctl0$CONTENU_PAGE$AdvancedSearch$keywordSearch")
        input_field.clear()
        # input_field.send_keys("intelligence artificielle") 
        print("ℹ️ Keywords entered (or skipped).")
    except Exception as e:
        print(f"⚠️ Warning: Keyword input error: {e}")

    time.sleep(1)
    
    # Click Search
    search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_button)
    time.sleep(1)
    search_button.click()
    print("✅ Search submitted.")
    time.sleep(5)

    # --- [STEP 4] Set Page Size ---
    try:
        print("--- [STEP 4] Adjusting Page Size ---")
        wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
        Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
        print("ℹ️ Page size set to 500.")
        time.sleep(5)
    except TimeoutException:
        print("ℹ️ No results table found (possibly 0 results).")

    # --- [STEP 5] Scrape Main Table ---
    print("--- [STEP 5] Scraping Table Rows ---")
    
    try:
        rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
        print(f"📄 Found {len(rows)} rows on current page.")

        for row in rows:
            try:
                # Skip header rows
                if "table-header" in row.get_attribute("class"): 
                    continue

                ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
                
                # XPath relative to row
                objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
                buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
                lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
                deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
                first_button = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
                
                metadata_list.append({
                    "reference": ref,
                    "objet": objet,
                    "acheteur": buyer,
                    "lieux_execution": lieux,
                    "date_limite": deadline,
                    "first_button_url": first_button
                })
            except Exception as e:
                continue
    except Exception as e:
        print(f"❌ Error scraping table: {e}")

    print(f"✅ Total tenders collected: {len(metadata_list)}")

    # -----------------------------
    # STEP 5.5: IMMEDIATE CSV SAVE
    # -----------------------------
    if metadata_list:
        print("💾 [SAFETY SAVE] Saving initial list to CSV...")
        try:
            initial_df = pd.DataFrame(metadata_list)
            # Use utf-8-sig so Excel opens it correctly with accents
            initial_csv_name = "initial_tenders_list.csv"
            initial_df.to_csv(initial_csv_name, index=False, encoding='utf-8-sig')
            print(f"✅ Initial list saved: {initial_csv_name}")
        except Exception as e:
            print(f"⚠️ Failed to save initial CSV: {e}")

    # --- [STEP 6] Process Each Tender Link (SKIPPED) ---
    print("\n--- [STEP 6] Skipping Deep Scraping (User Request) ---")
    
    # We transfer the metadata list directly to the final list
    # so that Step 7 saves the data we have already found.
    all_processed_tenders = metadata_list 

    ''' 
    # COMMENTED OUT AS REQUESTED
    df = pd.DataFrame(metadata_list)

    for idx, row in df.iterrows():
        link = row['first_button_url']
        print(f"\n🔗 [{idx+1}/{len(df)}] Processing: {link}")

        try:
            driver.get(link)
        except TimeoutException:
            print("   ⚠️ Timeout loading page. Retrying...")
            driver.refresh()
            time.sleep(3)

        merged_text = "No participants found"

        # Try to find 'Extrait de PV' or Participants Table
        try:
            # Look for the PV link
            try:
                pv_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@title, "Extrait de PV") or contains(text(), "Extrait de PV")]')))
                driver.execute_script("arguments[0].scrollIntoView(true);", pv_link)
                pv_link.click()
                time.sleep(2)
            except:
                pass # Link might not exist, but table might be there

            # Scrape the specific table for participants
            try:
                table_body = driver.find_element(By.CSS_SELECTOR, '#entreprisesParticipantesIn table.table-results tbody')
                part_rows = table_body.find_elements(By.TAG_NAME, 'tr')
                companies = [r.text.strip() for r in part_rows if r.text.strip()]
                
                if companies:
                    merged_text = " | ".join(companies)
                    print(f"   ✅ Found companies: {merged_text[:50]}...")
                else:
                    print("   ℹ️ Table empty.")
            except NoSuchElementException:
                 print("   ℹ️ No participant table found.")

        except Exception as e:
            print(f"   ⚠️ Error processing details: {e}")

        # Add gathered text to the data
        tender_payload = row.to_dict()
        tender_payload["participants"] = merged_text
        all_processed_tenders.append(tender_payload)

        # Cleanup temp files and wait politely
        clear_download_directory()
        time.sleep(random.uniform(1.5, 3))
    '''

except Exception as e:
    print("\n❌ [FATAL ERROR] Script crashed.")
    print(traceback.format_exc())
    driver.save_screenshot("fatal_crash.png")

finally:
    # -----------------------------
    # STEP 7: SAVE FINAL EXCEL
    # -----------------------------
    print("\n--- [STEP 7] Saving Final Data ---")
    
    if all_processed_tenders:
        # Create a unique filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"marches_publics_companys.xlsx"
        
        df_out = pd.DataFrame(all_processed_tenders)
        
        try:
            # Save to Excel
            df_out.to_excel(excel_filename, index=False, engine='openpyxl')
            print(f"✅ SUCCESS! Data saved to: {excel_filename}")
            print(f"📊 Total Rows: {len(df_out)}")
        except Exception as e:
            print(f"❌ Failed to save Excel (Reason: {e})")
            # Fallback to CSV
            csv_filename = f"marches_publics_backup_{timestamp}.csv"
            df_out.to_csv(csv_filename, index=False, encoding='utf-8-sig')
            print(f"✅ Saved as CSV instead: {csv_filename}")
    else:
        print("⚠️ No data was collected, so no file was saved.")

    # Close Browser
    try:
        driver.quit()
        print("👋 Browser closed.")
    except:
        pass

    # Final Cleanup
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir, ignore_errors=True)
    
    print("🎉 Done.")
