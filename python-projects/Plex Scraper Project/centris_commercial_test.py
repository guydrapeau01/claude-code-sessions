#!/usr/bin/env python3
"""
Centris COMMERCIAL Multifamily Scraper
With exact filters as specified
"""

import time
import re
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd


def setup_driver():
    """Create Chrome driver - visible so you can watch"""
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    
    try:
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        print("✓ Chrome started")
    except:
        driver = webdriver.Chrome(options=options)
        print("✓ Chrome started")
    
    return driver


def handle_popups(driver):
    """Click any privacy/cookie popups"""
    
    button_keywords = ['Accepter', 'Accept', 'Enregistrer', 'Save', 
                      'Confirmer', 'Confirm', 'Tout accepter', 'Continuer']
    
    for keyword in button_keywords:
        try:
            button = driver.find_element(By.XPATH, 
                f"//button[contains(text(), '{keyword}')]")
            driver.execute_script("arguments[0].click();", button)
            print(f"  ✓ Clicked '{keyword}' button")
            time.sleep(1)
        except:
            pass


def navigate_and_filter(driver):
    """
    Navigate to Centris Commercial section and apply filters:
    - Montréal (Île)
    - Laval  
    - Max price: $2,000,000
    - Multifamilial
    - Min 5 units
    - Listed in last 7 days
    """
    
    print("\n" + "="*70)
    print("NAVIGATING TO CENTRIS COMMERCIAL")
    print("="*70)
    
    # Go to Centris homepage
    print("\n1. Opening Centris.ca...")
    driver.get("https://www.centris.ca")
    time.sleep(3)
    
    # Handle popups
    print("2. Checking for popups...")
    handle_popups(driver)
    time.sleep(1)
    
    # Click COMMERCIAL tab
    print("3. Clicking 'Commercial' tab...")
    try:
        # Find and click Commercial link
        commercial_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, 
                "//a[contains(text(), 'Commercial') or contains(@href, 'commercial')]"))
        )
        commercial_link.click()
        print("   ✓ Clicked Commercial")
        time.sleep(2)
    except Exception as e:
        print(f"   ✗ Could not find Commercial tab: {e}")
        print("   Trying direct URL...")
        driver.get("https://www.centris.ca/fr/commercial~a-vendre")
        time.sleep(3)
    
    handle_popups(driver)
    
    # NOW APPLY FILTERS
    print("\n4. Applying filters...")
    
    # Filter: Property type = Multifamilial
    print("   Setting property type to 'Multifamilial'...")
    try:
        # Look for category/property type dropdown or checkboxes
        multifam_checkbox = driver.find_element(By.XPATH,
            "//label[contains(text(), 'Multifamilial')]//input | "
            "//input[@value='Multifamilial'] | "
            "//span[contains(text(), 'Multifamilial')]"
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", multifam_checkbox)
        time.sleep(0.5)
        multifam_checkbox.click()
        print("   ✓ Selected Multifamilial")
    except Exception as e:
        print(f"   ⚠ Could not find Multifamilial checkbox: {e}")
        print("   You may need to click it manually...")
        input("   Press ENTER when you've selected Multifamilial...")
    
    time.sleep(1)
    
    # Filter: Locations (Montreal Île, Laval)
    print("   Setting locations: Montréal (Île), Laval...")
    try:
        # Click on location/region selector
        location_input = driver.find_element(By.XPATH,
            "//input[@placeholder='Localisation'] | "
            "//input[contains(@class, 'location')] | "
            "//div[contains(@class, 'location')]//input"
        )
        location_input.click()
        time.sleep(1)
        
        # Type and select Montreal
        location_input.send_keys("Montréal")
        time.sleep(1)
        # Select from dropdown
        montreal_option = driver.find_element(By.XPATH,
            "//*[contains(text(), 'Montréal (Île)')]")
        montreal_option.click()
        time.sleep(1)
        
        # Add Laval
        location_input.send_keys("Laval")
        time.sleep(1)
        laval_option = driver.find_element(By.XPATH,
            "//*[contains(text(), 'Laval')]")
        laval_option.click()
        print("   ✓ Selected Montréal (Île) and Laval")
    except Exception as e:
        print(f"   ⚠ Could not set locations automatically: {e}")
        print("   Please select Montréal (Île) and Laval manually...")
        input("   Press ENTER when done...")
    
    time.sleep(1)
    
    # Filter: Max price $2,000,000
    print("   Setting max price to $2,000,000...")
    try:
        max_price_input = driver.find_element(By.XPATH,
            "//input[@placeholder='Max'] | "
            "//input[contains(@id, 'maxprice')] | "
            "//input[contains(@name, 'maxprice')]"
        )
        max_price_input.clear()
        max_price_input.send_keys("2000000")
        print("   ✓ Set max price")
    except Exception as e:
        print(f"   ⚠ Could not set max price: {e}")
        print("   Please enter 2,000,000 manually...")
        input("   Press ENTER when done...")
    
    time.sleep(1)
    
    # Filter: Min 5 units
    print("   Setting minimum 5 units...")
    try:
        min_units_input = driver.find_element(By.XPATH,
            "//input[contains(@id, 'minunits')] | "
            "//input[contains(@name, 'minunits')] | "
            "//label[contains(text(), 'Nombre d')]//following::input[1]"
        )
        min_units_input.clear()
        min_units_input.send_keys("5")
        print("   ✓ Set min units to 5")
    except Exception as e:
        print(f"   ⚠ Could not set min units: {e}")
        print("   Please enter 5 manually in 'Nombre d'unités min'...")
        input("   Press ENTER when done...")
    
    time.sleep(1)
    
    # Filter: New since (last 7 days)
    print("   Setting 'New since last 7 days'...")
    try:
        # Calculate date 7 days ago
        seven_days_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        
        # Find date picker or "new since" field
        date_input = driver.find_element(By.XPATH,
            "//input[contains(@placeholder, 'date')] | "
            "//input[@type='date'] | "
            "//label[contains(text(), 'Nouveau depuis')]//following::input[1]"
        )
        date_input.clear()
        date_input.send_keys(seven_days_ago)
        print(f"   ✓ Set date to {seven_days_ago}")
    except Exception as e:
        print(f"   ⚠ Could not set date filter: {e}")
        print("   Please set 'Nouveau depuis' to last 7 days manually...")
        input("   Press ENTER when done...")
    
    time.sleep(1)
    
    # Click SEARCH button
    print("\n5. Clicking SEARCH...")
    try:
        search_button = driver.find_element(By.XPATH,
            "//button[contains(text(), 'Rechercher')] | "
            "//button[@type='submit'] | "
            "//button[contains(@class, 'search')]"
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
        time.sleep(0.5)
        search_button.click()
        print("   ✓ Clicked search")
        time.sleep(4)
    except Exception as e:
        print(f"   ⚠ Could not click search: {e}")
        print("   Press the search button manually...")
        input("   Press ENTER when search results appear...")
    
    handle_popups(driver)
    
    print("\n✓ Filters applied - on search results page")


def extract_property_details(driver):
    """
    Extract data from a commercial property detail page
    Looking for:
    - Revenus bruts potentiels
    - Nombre d'unités: Résidentiel (6)
    - Taxes municipales (2026) or (2025)
    - Taxes scolaires (2026) or (2025)
    """
    
    print("      Extracting property details...")
    
    # Get all page text
    page_text = driver.find_element(By.TAG_NAME, 'body').text
    
    data = {}
    
    # ADDRESS
    try:
        h1 = driver.find_element(By.TAG_NAME, 'h1')
        data['address'] = h1.text.strip()
    except:
        data['address'] = "Address not found"
    
    # PRICE
    data['price'] = extract_price(page_text)
    
    # REVENUS BRUTS POTENTIELS (Gross Income)
    print("      Looking for 'Revenus bruts potentiels'...")
    patterns = [
        r'Revenus bruts potentiels\s*:?\s*\$?\s*([\d\s]+)',
        r'Revenus bruts\s*:?\s*\$?\s*([\d\s]+)',
        r'Revenus potentiels\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    gross_income = None
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            income_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                gross_income = float(income_str)
                if 1000 < gross_income < 2000000:
                    data['gross_income'] = gross_income
                    print(f"      ✓ Found gross income: ${gross_income:,.0f}")
                    break
            except:
                pass
    
    if not gross_income:
        print("      ⚠ Gross income not found")
        data['gross_income'] = None
    
    # NOMBRE D'UNITÉS: RÉSIDENTIEL
    print("      Looking for number of units...")
    patterns = [
        r"Nombre d['\u2019]unités\s*:?\s*Résidentiel\s*\((\d+)\)",
        r"Nombre d['\u2019]unités\s*:?\s*(\d+)",
        r"Résidentiel\s*\((\d+)\s*unités?\)",
        r"(\d+)\s*unités?\s*résidentiell?",
    ]
    
    units = None
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            units = int(match.group(1))
            if 2 <= units <= 100:
                data['num_units'] = units
                print(f"      ✓ Found {units} units")
                break
    
    if not units:
        print("      ⚠ Number of units not found")
        data['num_units'] = None
    
    # TAXES MUNICIPALES (2026) or (2025)
    print("      Looking for municipal taxes...")
    patterns = [
        r'Municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Taxes municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Municipales\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    muni_tax = None
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            tax_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                muni_tax = float(tax_str)
                if 500 < muni_tax < 200000:
                    data['municipal_tax'] = muni_tax
                    print(f"      ✓ Found municipal tax: ${muni_tax:,.0f}")
                    break
            except:
                pass
    
    if not muni_tax:
        print("      ⚠ Municipal tax not found")
        data['municipal_tax'] = None
    
    # TAXES SCOLAIRES (2026) or (2025)
    print("      Looking for school taxes...")
    patterns = [
        r'Scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Taxes scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Scolaires\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    school_tax = None
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            tax_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                school_tax = float(tax_str)
                if 100 < school_tax < 50000:
                    data['school_tax'] = school_tax
                    print(f"      ✓ Found school tax: ${school_tax:,.0f}")
                    break
            except:
                pass
    
    if not school_tax:
        print("      ⚠ School tax not found")
        data['school_tax'] = None
    
    data['url'] = driver.current_url
    data['scraped_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    return data


def extract_price(page_text):
    """Extract asking price"""
    patterns = [
        r'Prix demandé\s*:?\s*\$?\s*([\d\s]+)',
        r'Asking price\s*:?\s*\$?\s*([\d\s]+)',
        r'\$\s*([\d\s]{6,})',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            price_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                price = float(price_str)
                if 100000 < price < 10000000:
                    return price
            except:
                pass
    
    return None


def test_single_property():
    """
    Test by extracting data from ONE property only
    """
    
    driver = setup_driver()
    
    try:
        # Navigate and apply filters
        navigate_and_filter(driver)
        
        print("\n" + "="*70)
        print("TESTING: SCRAPING ONE PROPERTY")
        print("="*70)
        
        # Wait for results to load
        print("\nWaiting for search results...")
        time.sleep(3)
        
        # Scroll to load lazy content
        for _ in range(2):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # Find first property card
        print("\nLooking for first property card...")
        try:
            # Try different selectors for commercial listings
            selectors = [
                ".property-thumbnail-item a",
                "div[class*='property'] a",
                ".listing-card a",
                "a[href*='/commercial/']",
            ]
            
            first_card = None
            for selector in selectors:
                try:
                    cards = driver.find_elements(By.CSS_SELECTOR, selector)
                    if cards:
                        first_card = cards[0]
                        break
                except:
                    continue
            
            if not first_card:
                print("⚠ No property cards found")
                print("Are there any search results? Check the browser...")
                input("Press ENTER to continue anyway...")
                return None
            
            print(f"✓ Found property card")
            
            # Get URL before clicking
            href = first_card.get_attribute('href')
            print(f"   URL: {href[:80]}...")
            
            # Click to open detail page
            print("\nClicking on property...")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_card)
            time.sleep(0.5)
            
            try:
                first_card.click()
            except:
                driver.get(href)
            
            time.sleep(3)
            
            # Handle any popups
            handle_popups(driver)
            
            # Extract data
            print("\n" + "-"*70)
            data = extract_property_details(driver)
            print("-"*70)
            
            # Show results
            print("\n" + "="*70)
            print("EXTRACTED DATA:")
            print("="*70)
            print(f"Address:        {data.get('address', 'N/A')}")
            print(f"Price:          ${data.get('price', 0):,.0f}")
            print(f"Units:          {data.get('num_units', 'N/A')}")
            print(f"Gross Income:   ${data.get('gross_income', 0):,.0f}" if data.get('gross_income') else "Gross Income:   Not found")
            print(f"Municipal Tax:  ${data.get('municipal_tax', 0):,.0f}" if data.get('municipal_tax') else "Municipal Tax:  Not found")
            print(f"School Tax:     ${data.get('school_tax', 0):,.0f}" if data.get('school_tax') else "School Tax:     Not found")
            print(f"URL:            {data.get('url', '')}")
            print("="*70)
            
            # Save to CSV
            df = pd.DataFrame([data])
            filename = 'test_single_property.csv'
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            print(f"\n✓ Saved to: {filename}")
            
            return data
            
        except Exception as e:
            print(f"✗ Error: {e}")
            import traceback
            traceback.print_exc()
            return None
        
    finally:
        print("\nKeeping browser open so you can inspect...")
        input("Press ENTER to close browser...")
        driver.quit()


if __name__ == '__main__':
    print("\n╔═══════════════════════════════════════════════════════════════╗")
    print("║  CENTRIS COMMERCIAL SCRAPER - TEST MODE (1 property only)    ║")
    print("╚═══════════════════════════════════════════════════════════════╝")
    print("\nThis will:")
    print("  1. Open Centris")
    print("  2. Go to Commercial section")
    print("  3. Apply your exact filters")
    print("  4. Extract data from ONE property")
    print("\nYou'll see the browser so you can help if needed.\n")
    
    input("Press ENTER to start...")
    
    test_single_property()
