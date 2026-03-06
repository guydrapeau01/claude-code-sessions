#!/usr/bin/env python3
"""
Centris Commercial Scraper - MANUAL GUIDED VERSION

YOU do the searching and filtering.
The script does the tedious data extraction.
"""

import time
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd


def setup_driver():
    """Create Chrome driver"""
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    
    try:
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except:
        driver = webdriver.Chrome(options=options)
    
    return driver


def extract_property_data(driver):
    """
    Extract data from commercial property detail page
    
    Looking for:
    - Address (from h1)
    - Price (Prix demandé)
    - Revenus bruts potentiels
    - Nombre d'unités: Résidentiel (X)
    - Municipales (2026 or 2025)
    - Scolaires (2026 or 2025)
    """
    
    # Get all page text
    page_text = driver.find_element(By.TAG_NAME, 'body').text
    
    data = {}
    
    # ADDRESS
    try:
        h1 = driver.find_element(By.TAG_NAME, 'h1')
        data['address'] = h1.text.strip()
    except:
        data['address'] = "Unknown"
    
    # PRICE
    price_patterns = [
        r'Prix demandé\s*:?\s*\$?\s*([\d\s]+)',
        r'Asking price\s*:?\s*\$?\s*([\d\s]+)',
        r'\$\s*([\d\s]{6,})',
    ]
    
    data['price'] = None
    for pattern in price_patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            price_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                price = float(price_str)
                if 100000 < price < 10000000:
                    data['price'] = price
                    break
            except:
                pass
    
    # REVENUS BRUTS POTENTIELS
    income_patterns = [
        r'Revenus bruts potentiels\s*:?\s*\$?\s*([\d\s]+)',
        r'Revenus bruts\s*:?\s*\$?\s*([\d\s]+)',
        r'Revenus potentiels\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    data['gross_income'] = None
    for pattern in income_patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            income_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                income = float(income_str)
                if 1000 < income < 2000000:
                    data['gross_income'] = income
                    break
            except:
                pass
    
    # NOMBRE D'UNITÉS: RÉSIDENTIEL (X)
    units_patterns = [
        r"Nombre d['\u2019]unités\s*:?\s*Résidentiel\s*\((\d+)\)",
        r"Résidentiel\s*\((\d+)\)",
        r"Nombre d['\u2019]unités\s*:?\s*(\d+)",
        r"(\d+)\s*unités?\s*résidentiel",
    ]
    
    data['num_units'] = None
    for pattern in units_patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            try:
                units = int(match.group(1))
                if 2 <= units <= 100:
                    data['num_units'] = units
                    break
            except:
                pass
    
    # MUNICIPALES (2026 or 2025)
    muni_patterns = [
        r'Municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Taxes municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Municipales\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    data['municipal_tax'] = None
    for pattern in muni_patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            tax_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                tax = float(tax_str)
                if 500 < tax < 200000:
                    data['municipal_tax'] = tax
                    break
            except:
                pass
    
    # SCOLAIRES (2026 or 2025)
    school_patterns = [
        r'Scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Taxes scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        r'Scolaires\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    data['school_tax'] = None
    for pattern in school_patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            tax_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                tax = float(tax_str)
                if 100 < tax < 50000:
                    data['school_tax'] = tax
                    break
            except:
                pass
    
    data['url'] = driver.current_url
    data['scraped_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    return data


def scrape_from_manual_search(max_properties=20):
    """
    Scrape properties from search results that USER has already set up
    """
    
    print("\n" + "="*70)
    print("CENTRIS COMMERCIAL SCRAPER - MANUAL GUIDED")
    print("="*70)
    print("\nINSTRUCTIONS:")
    print("1. Chrome will open")
    print("2. YOU navigate to Centris Commercial")
    print("3. YOU apply filters:")
    print("   - Montréal (Île)")
    print("   - Laval")
    print("   - Max price: $2,000,000")
    print("   - Multifamilial")
    print("   - Min 5 units")
    print("   - New since last 7 days")
    print("4. YOU click SEARCH")
    print("5. When you see search results, come back here and press ENTER")
    print("6. The script will scrape all properties automatically")
    print("\n" + "="*70)
    
    input("\nPress ENTER to open Chrome...")
    
    driver = setup_driver()
    
    print("\n✓ Chrome opened")
    print("\nNow YOU do the search:")
    print("  1. Go to https://www.centris.ca")
    print("  2. Click 'Commercial'")
    print("  3. Apply your filters")
    print("  4. Click Search")
    print("\nWhen you see search results...")
    
    input("Press ENTER to start scraping...")
    
    all_properties = []
    
    try:
        print("\n" + "="*70)
        print(f"SCRAPING UP TO {max_properties} PROPERTIES")
        print("="*70)
        
        page_number = 1
        properties_scraped = 0
        
        while properties_scraped < max_properties:
            
            print(f"\n--- PAGE {page_number} ---")
            
            # Scroll to load all cards
            print("Scrolling to load all properties...")
            for _ in range(3):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)
            
            # Find all property cards
            print("Finding property cards...")
            selectors = [
                "a.property-thumbnail-summary-link",
                ".property-thumbnail-item a",
                "div[class*='property'] a[href*='/commercial/']",
                ".listing-card a",
            ]
            
            cards = []
            for selector in selectors:
                try:
                    cards = driver.find_elements(By.CSS_SELECTOR, selector)
                    if cards:
                        break
                except:
                    pass
            
            if not cards:
                print("⚠ No property cards found on this page")
                break
            
            print(f"Found {len(cards)} property cards on page {page_number}")
            
            # Process each card
            card_index = 0
            while card_index < len(cards) and properties_scraped < max_properties:
                
                print(f"\n[{properties_scraped + 1}/{max_properties}] Property {card_index + 1}/{len(cards)}...")
                
                # Re-find cards to avoid stale elements
                cards = []
                for selector in selectors:
                    try:
                        cards = driver.find_elements(By.CSS_SELECTOR, selector)
                        if cards:
                            break
                    except:
                        pass
                
                if card_index >= len(cards):
                    break
                
                card = cards[card_index]
                
                # Get URL
                try:
                    href = card.get_attribute('href')
                except:
                    card_index += 1
                    continue
                
                # Click property
                print(f"  Opening: {href[-50:]}")
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card)
                    time.sleep(0.5)
                    
                    try:
                        card.click()
                    except:
                        driver.get(href)
                    
                    time.sleep(2)
                    
                    # Extract data
                    data = extract_property_data(driver)
                    
                    if data.get('price'):
                        all_properties.append(data)
                        properties_scraped += 1
                        
                        # Show what we got
                        print(f"  ✓ {data['address'][:50]}")
                        print(f"    Price:      ${data.get('price', 0):,.0f}")
                        print(f"    Units:      {data.get('num_units', '?')}")
                        
                        if data.get('gross_income'):
                            print(f"    Income:     ${data['gross_income']:,.0f}")
                        else:
                            print(f"    Income:     Not found")
                        
                        if data.get('municipal_tax'):
                            print(f"    Muni Tax:   ${data['municipal_tax']:,.0f}")
                        else:
                            print(f"    Muni Tax:   Not found")
                        
                        if data.get('school_tax'):
                            print(f"    School Tax: ${data['school_tax']:,.0f}")
                        else:
                            print(f"    School Tax: Not found")
                    
                    # Go back
                    print("  Going back...")
                    driver.back()
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"  ✗ Error: {e}")
                
                card_index += 1
            
            # Try next page
            print("\nLooking for next page...")
            try:
                next_button = driver.find_element(By.XPATH,
                    "//a[@aria-label='Suivant' or @aria-label='Next'] | "
                    "//li[contains(@class, 'next')]/a"
                )
                
                if 'disabled' not in next_button.get_attribute('class'):
                    print("✓ Going to next page...")
                    next_button.click()
                    time.sleep(3)
                    page_number += 1
                else:
                    print("→ No more pages")
                    break
            except:
                print("→ No next page button")
                break
        
        print("\n" + "="*70)
        print(f"SCRAPING COMPLETE - {len(all_properties)} properties")
        print("="*70)
        
    except KeyboardInterrupt:
        print("\n\n⚠ Stopped by user")
    
    finally:
        # Save results
        if all_properties:
            df = pd.DataFrame(all_properties)
            filename = f'centris_commercial_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            
            print(f"\n✓ Saved {len(all_properties)} properties to: {filename}")
            
            # Show summary
            with_income = sum(1 for p in all_properties if p.get('gross_income'))
            with_taxes = sum(1 for p in all_properties if p.get('municipal_tax'))
            
            print(f"\nData quality:")
            print(f"  With gross income: {with_income}/{len(all_properties)}")
            print(f"  With tax data:     {with_taxes}/{len(all_properties)}")
            
            if with_income < len(all_properties):
                print(f"\n⚠ {len(all_properties) - with_income} missing gross income")
                print(f"  You'll need to add it manually in the CSV")
        else:
            print("\n⚠ No properties scraped")
        
        print("\nKeeping browser open for 10 seconds...")
        time.sleep(10)
        driver.quit()
        print("✓ Browser closed")
    
    return all_properties


if __name__ == '__main__':
    
    print("\n╔═══════════════════════════════════════════════════════════════╗")
    print("║     CENTRIS COMMERCIAL SCRAPER - MANUAL GUIDED VERSION        ║")
    print("╚═══════════════════════════════════════════════════════════════╝")
    print("\nYOU do the filtering, the SCRIPT does the data extraction.")
    
    # How many properties to scrape
    MAX_PROPERTIES = 20  # Change this if you want more/less
    
    properties = scrape_from_manual_search(max_properties=MAX_PROPERTIES)
    
    if properties:
        print(f"\n{'='*70}")
        print("NEXT STEPS")
        print(f"{'='*70}")
        print("\n1. Review the CSV file")
        print("2. Fill in any missing 'gross_income' values")
        print("3. Run: python run_analysis.py")
        print("   (Tell it to skip scraping and use your CSV)")
