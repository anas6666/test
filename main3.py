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
    StaleElementReferenceException,
    ElementNotInteractableException # Added for potential Next button issues
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
current_page_number = 1 # Added for pagination tracking

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
    yesterday = "01/01/2020" # Using a fixed date for broader results
    
    # Fill Date
    try:
        date_input = driver.find_element(By.NAME, "ctl0$CONTENU_PAGE$AdvancedSearch$dateMiseEnLigneStart")
        date_input.clear()
        date_input.send_keys(yesterday)
        print(f"ℹ️ Start date set to: {yesterday}")
    except NoSuchElementException:
        try:
            date_input_2 = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
            date_input_2.clear()
            date_input_2.send_keys(yesterday)
            print(f"ℹ️ Start date (alternative field) set to: {yesterday}")
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
    time.sleep(5) # Give it time to load the initial results

    # --- [STEP 4] Set Page Size ---
    try:
        print("--- [STEP 4] Adjusting Page Size ---")
        # Ensure the table and page size dropdown are present before interacting
        wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
        Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
        print("ℹ️ Page size set to 500.")
        time.sleep(7) # Increased sleep after changing page size, as it triggers a reload
    except TimeoutException:
        print("ℹ️ No results table or page size dropdown found (possibly 0 results).")
    except NoSuchElementException:
        print("ℹ️ Page size dropdown not found, continuing with default page size.")
        
    # --- [STEP 5] Scrape Main Table and Paginate ---
    print("\n--- [STEP 5] Scraping Table Rows and Paginating ---")

    while True: # Loop indefinitely until no more "Next" buttons
        print(f"--- Scraping Page {current_page_number} ---")
        try:
            # Wait for the table rows to be present on the page
            # This helps against StaleElementReferenceException if the page reloads
            wait.until(EC.presence_of_all_elements_located((By.XPATH, '//table[@class="table-results"]/tbody/tr[not(contains(@class, "table-header"))]')))
            
            rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr[not(contains(@class, "table-header"))]')
            
            if not rows:
                print(f"📄 No tender rows found on page {current_page_number}. Ending pagination.")
                break # Exit if no rows are found, likely end of results
            
            print(f"📄 Found {len(rows)} rows on Page {current_page_number}.")

            for row in rows:
                try:
                    ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
                    
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
                except StaleElementReferenceException:
                    print("   ⚠️ Stale element encountered, re-locating row for next iteration.")
                    continue # Skip this row, try next if the element is stale
                except Exception as e:
                    print(f"   ⚠️ Error scraping a row on page {current_page_number}: {e}")
                    continue
            
            print(f"✅ Total tenders collected so far: {len(metadata_list)}")
            
            # --- Pagination Logic ---
            # Try to find the "Next" button
            # Look for an <a> tag with title "Page suivante" or text "Suivante"
            # It's usually within a 'pagination' div. The structure might vary.
            next_button = None
            try:
                # This XPath looks for a link within the pagination section that contains "Suivante"
                # or a link that specifically has 'title="Page suivante"'.
                # Adjust this XPath if your "Next" button has a different identifier.
                next_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_resultSearch_PagerTop_ctl2")))
        
                
                # Check if the next button is enabled/active
                if "aspNetDisabled" in next_button.get_attribute("class") or next_button.get_attribute("disabled"):
                    print("ℹ️ 'Next' button is disabled or not active. Ending pagination.")
                    break # Exit loop if button is disabled
                
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(1) # Small pause before clicking
                next_button.click()
                print(f"➡️ Clicked 'Next' button to go to Page {current_page_number + 1}.")
                time.sleep(5) # Give page time to load new results
                current_page_number += 1
                
            except (NoSuchElementException, TimeoutException, ElementNotInteractableException):
                print("ℹ️ 'Next' button not found or not clickable. Ending pagination.")
                break # Exit the loop if no next button is found or interactable
            
        except TimeoutException:
            print(f"❌ Timed out waiting for table rows on page {current_page_number}. Ending pagination.")
            break # Exit the loop if rows aren't found on a new page
        except Exception as e:
            print(f"❌ Error during pagination on page {current_page_number}: {e}")
            print(traceback.format_exc())
            break # Break on other unexpected errors

    print(f"✅ Total unique tenders collected after pagination: {len(metadata_list)}")

    # -----------------------------
    # STEP 5.5: IMMEDIATE CSV SAVE
    # -----------------------------
    if metadata_list:
        print("💾 [SAFETY SAVE] Saving initial list to CSV...")
        try:
            initial_df = pd.DataFrame(metadata_list)
            # Use utf-8-sig so Excel opens it correctly with accents
            initial_csv_name = "initial_tenders_list_paginated.csv" # Changed filename
            initial_df.to_csv(initial_csv_name, index=False, encoding='utf-8-sig')
            print(f"✅ Initial list saved: {initial_csv_name}")
        except Exception as e:
            print(f"⚠️ Failed to save initial CSV: {e}")

    # --- [STEP 6] Process Each Tender Link (SKIPPED) ---
    print("\n--- [STEP 6] Skipping Deep Scraping (User Request) ---")
    
    # We transfer the metadata list directly to the final list
    # so that Step 7 saves the data we have already found.
    all_processed_tenders = metadata_list 

    # (Your commented-out deep scraping code remains here)

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
        excel_filename = f"marches_publics_companys_all_pages.xlsx" # Changed filename
        
        df_out = pd.DataFrame(all_processed_tenders)
        
        try:
            # Save to Excel
            df_out.to_excel(excel_filename, index=False, engine='openpyxl')
            print(f"✅ SUCCESS! Data saved to: {excel_filename}")
            print(f"📊 Total Rows: {len(df_out)}")
        except Exception as e:
            print(f"❌ Failed to save Excel (Reason: {e})")
            # Fallback to CSV
            csv_filename = f"marches_publics_backup_all_pages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
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
