"""
CMF Website Scraper Module
Handles interaction with the CMF website for company and document retrieval
"""
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import os
import re

from src.utils.helpers import extract_year_from_text


def init_driver():
    """Initialize Chrome WebDriver with headless options"""
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    service = Service(ChromeDriverManager().install(), log_path=os.devnull)
    return webdriver.Chrome(service=service, options=chrome_options)

def extract_year_from_text(text):
    if not text:
        return None
    
    #  Try 4-digit year first
    match = re.search(r'\b(20\d{2})\b', text)
    if match:
        return match.group(1)
    
    # 2️⃣ Fallback: last 2 digits before .pdf
    match = re.search(r'(\d{2})\.pdf$', text)
    if match:
        year_2d = int(match.group(1))
        if year_2d >= 0 and year_2d <= 25:  # assume 2000-2025
            return str(2000 + year_2d)
        else:
            return str(1900 + year_2d)  # fallback if somehow >25
    
    return None
def get_all_companies(driver):
    """Fetch all companies from the CMF dropdown"""
    print("⏳ Chargement de la liste des sociétés...")
    driver.get("https://www.cmf.tn/consultation-des-tats-financier-des-soci-t-s-faisant-ape")
    wait = WebDriverWait(driver, 20)
    
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#edit_field_societesape_value_chosen a")))
        dropdown.click()
        
        options = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "#edit_field_societesape_value_chosen .chosen-results li")))
        
        companies = [opt.text.strip() for opt in options if opt.text.strip()]
        return companies
    except Exception as e:
        print(f"❌ Erreur lors du chargement des sociétés : {e}")
        return []


def select_company_and_submit(driver, company_name):
    """Select company in dropdown and submit the form"""
    print(f"\n🔄 Sélection de '{company_name}' sur le site CMF...")
    try:
        wait = WebDriverWait(driver, 20)
        
        # Refresh the page to avoid stale elements
        driver.get("https://www.cmf.tn/consultation-des-tats-financier-des-soci-t-s-faisant-ape")
        time.sleep(2)
        
        # Click dropdown
        dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#edit_field_societesape_value_chosen a")))
        dropdown.click()
        time.sleep(1)
        
        # Find and click the company option
        options = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "#edit_field_societesape_value_chosen .chosen-results li")))
        
        selected_option = None
        for option in options:
            if company_name.lower() == option.text.strip().lower():
                selected_option = option
                break
        
        if not selected_option:
            # Try partial match
            for option in options:
                opt_text = option.text.strip()
                if opt_text and company_name.lower() in opt_text.lower():
                    selected_option = option
                    break
        
        if not selected_option:
            print(f"Impossible de sélectionner '{company_name}' dans le dropdown.")
            return False
        
        selected_option.click()
        
        # Submit form
        submit = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#edit-submit-consultation-des-tats-financier-des-soci-t-s-faisant-ape")))
        submit.click()
        time.sleep(3)  # Wait for results to load
        
        print(" Formulaire soumis, chargement des résultats...")
        return True
        
    except Exception as e:
        print(f" Erreur lors de la sélection de la société : {e}")
        return False


def scrape_document_list(driver, company_name):
    """Scrape all documents from the results page"""
    print(f"\n Récupération des documents...")
    
    try:
        pdfs = []
        page_num = 1
        
        while True:
            time.sleep(2)  # Wait for page to load
            rows = driver.find_elements(By.CSS_SELECTOR, ".view-content .views-row")
            print(f"DEBUG: Found {len(rows)} documents on page {page_num}.")
            
            for row in rows:
                try:
                    pdf_url = row.find_element(By.CSS_SELECTOR, "a[href$='.pdf']").get_attribute("href")
                    pdf_name = row.find_element(
                        By.CSS_SELECTOR, ".field-name-field-p-riode .field-item").text.strip()
                    
                    year = extract_year_from_text(pdf_url) or extract_year_from_text(pdf_name)
                    
                    if year and year.isdigit():
                        pdfs.append({
                            "url": pdf_url,
                            "nom": pdf_name,
                            "annee": year,
                            "societe": company_name
                        })
                    else:
                        print(f"DEBUG: Skipped document '{pdf_name}' (URL: {pdf_url}) - Year extracted: {year}")
                except (NoSuchElementException, StaleElementReferenceException):
                    continue
            
            # Try to go to next page
            try:
                next_btn = driver.find_element(By.CSS_SELECTOR, ".pager .next a:not(.disabled)")
                next_btn.click()
                page_num += 1
                time.sleep(3)
            except:
                break
        
        print(f"✅ {len(pdfs)} documents trouvés au total.")
        return pdfs
        
    except Exception as e:
        print(f"❌ ERREUR : {str(e)}")
        return []
