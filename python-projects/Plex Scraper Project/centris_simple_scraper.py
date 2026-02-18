#!/usr/bin/env python3
"""
Centris Commercial Scraper - NO CLICKING VERSION

YOU set up the search with filters.
Script extracts data DIRECTLY from the property cards (no clicking).
"""

import time
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd


def setup_driver():
    """Create Chrome driver - NO popup handling"""
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


def parse_number(text):
    """Extract number from text like '$1 234 567' or '1,234'"""
    if not text:
        return None
    try:
        cleaned = re.sub(r'[^\d.]', '', str(text).replace(' ', '').replace('\xa0', ''))
        return float(cleaned) if cleaned else None
    except:
        return None


def extract_from_card(card):
    """
    Extract all data directly from a property card element
    WITHOUT clicking on it
    """
    
    data = {}
    card_text = card.text
    
    try:
        # ADDRESS - usually in a link or heading
        try:
            address_elem = card.find_element(By.CSS_SELECTOR, 
                "a.address, .property-address, h2, h3, a")
            data['address'] = address_elem.text.strip()
        except:
            data['address'] = "Unknown"
        
        # PRICE
        try:
            price_elem = card.find_element(By.CSS_SELECTOR, 
                ".price, [class*='price'], .asking-price")
            price = parse_number(price_elem.text)
            data['price'] = price if price and price > 50000 else None
        except:
            # Fallback: search in card text
            match = re.search(r'\$\s*([\d\s]+)', card_text)
            if match:
                data['price'] = parse_number(match.group(1))
            else:
                data['price'] = None
        
        # UNITS - might be visible on card
        units_match = re.search(r'(\d+)\s*unités?|(\d+)\s*logements?', card_text, re.IGNORECASE)
        if units_match:
            units = int(units_match.group(1) or units_match.group(2))
            data['num_units'] = units if 2 <= units <= 100 else None
        else:
            data['num_units'] = None
        
        # URL - get the link to the detail page
        try:
            link = card.find_element(By.CSS_SELECTOR, "a[href*='/commercial/']")
            data['url'] = link.get_attribute('href')
        except:
            try:
                link = card.find_element(By.TAG_NAME, "a")
                data['url'] = link.get_attribute('href')
            except:
                data['url'] = None
        
        # These are typically NOT on cards, need to visit detail page
        data['gross_income'] = None
        data['municipal_tax'] = None
        data['school_tax'] = None
        data['scraped_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        
    except Exception as e:
        print(f"    Error parsing card: {e}")
        return None
    
    return data if data.get('price') else None


def visit_and_extract_details(driver, url):
    """
    Visit a property detail page and extract:
    - Revenus bruts potentiels
    - Nombre d'unités: Résidentiel
    - Taxes municipales
    - Taxes scolaires
    """
    
    try:
        driver.get(url)
        time.sleep(2)
        
        page_text = driver.find_element(By.TAG_NAME, 'body').text
        
        details = {}
        
        # REVENUS BRUTS POTENTIELS
        patterns = [
            r'Revenus bruts potentiels\s*:?\s*\$?\s*([\d\s]+)',
            r'Revenus bruts\s*:?\s*\$?\s*([\d\s]+)',
            r'Revenus potentiels\s*:?\s*\$?\s*([\d\s]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                income = parse_number(match.group(1))
                if income and 1000 < income < 2000000:
                    details['gross_income'] = income
                    break
        
        if not details.get('gross_income'):
            details['gross_income'] = None
        
        # NOMBRE D'UNITÉS: RÉSIDENTIEL
        patterns = [
            r"Nombre d['\u2019]unités\s*:?\s*Résidentiel\s*\((\d+)\)",
            r"Résidentiel\s*\((\d+)\)",
            r"Nombre d['\u2019]unités\s*:?\s*(\d+)",
        ]
        
        for pattern in patterns:
            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                units = int(match.group(1))
                if 2 <= units <= 100:
                    details['num_units'] = units
                    break
        
        if not details.get('num_units'):
            details['num_units'] = None
        
        # MUNICIPALES (2026 or 2025)
        patterns = [
            r'Municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
            r'Taxes municipales\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                tax = parse_number(match.group(1))
                if tax and 500 < tax < 200000:
                    details['municipal_tax'] = tax
                    break
        
        if not details.get('municipal_tax'):
            details['municipal_tax'] = None
        
        # SCOLAIRES (2026 or 2025)
        patterns = [
            r'Scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
            r'Taxes scolaires\s*\(202[56]\)\s*:?\s*\$?\s*([\d\s]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                tax = parse_number(match.group(1))
                if tax and 100 < tax < 50000:
                    details['school_tax'] = tax
                    break
        
        if not details.get('school_tax'):
            details['school_tax'] = None
        
        return details
        
    except Exception as e:
        print(f"      Error visiting detail page: {e}")
        return {
            'gross_income': None,
            'num_units': None,
            'municipal_tax': None,
            'school_tax': None
        }


def scrape_current_page(driver):
    """
    Scrape all properties visible on current search results page
    WITHOUT clicking or navigating away
    """
    
    print("\n  Scrolling to load all cards...")
    # Scroll to load lazy-loaded content
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)
    
    print("  Finding property cards...")
    
    # Find all property cards
    selectors = [
        ".property-thumbnail-item",
        "div[class*='property-card']",
        "div[class*='listing']",
        ".thumbnail-item",
    ]
    
    cards = []
    for selector in selectors:
        try:
            cards = driver.find_elements(By.CSS_SELECTOR, selector)
            if cards:
                print(f"  ✓ Found {len(cards)} cards using selector: {selector}")
                break
        except:
            pass
    
    if not cards:
        print("  ✗ No cards found")
        return []
    
    # Extract data from each card
    properties = []
    
    for i, card in enumerate(cards, 1):
        print(f"\n  [{i}/{len(cards)}] Extracting from card...")
        
        data = extract_from_card(card)
        
        if data:
            print(f"    ✓ {data.get('address', 'N/A')[:45]}")
            print(f"      Price: ${data.get('price', 0):,.0f}")
            print(f"      URL: {data.get('url', 'N/A')[-50:]}")
            properties.append(data)
        else:
            print(f"    ✗ Could not extract data")
    
    return properties


def enrich_with_details(driver, properties):
    """
    Visit each property URL and get the detailed data
    (Income, units, taxes)
    """
    
    print("\n" + "="*70)
    print(f"VISITING {len(properties)} PROPERTY PAGES FOR DETAILS")
    print("="*70)
    
    for i, prop in enumerate(properties, 1):
        
        url = prop.get('url')
        if not url:
            print(f"\n[{i}/{len(properties)}] Skipping - no URL")
            continue
        
        print(f"\n[{i}/{len(properties)}] {prop.get('address', 'N/A')[:45]}")
        print(f"  Visiting: {url[-55:]}")
        
        details = visit_and_extract_details(driver, url)
        
        # Merge details into property
        prop.update(details)
        
        # Show what we found
        if details.get('gross_income'):
            print(f"  ✓ Income:     ${details['gross_income']:,.0f}")
        else:
            print(f"  ⚠ Income:     Not found")
        
        if details.get('num_units'):
            print(f"  ✓ Units:      {details['num_units']}")
        else:
            print(f"  ⚠ Units:      Not found")
        
        if details.get('municipal_tax'):
            print(f"  ✓ Muni Tax:   ${details['municipal_tax']:,.0f}")
        else:
            print(f"  ⚠ Muni Tax:   Not found")
        
        if details.get('school_tax'):
            print(f"  ✓ School Tax: ${details['school_tax']:,.0f}")
        else:
            print(f"  ⚠ School Tax: Not found")
        
        # Small delay between pages
        time.sleep(2)
    
    return properties


def main():
    
    print("\n╔═══════════════════════════════════════════════════════════════╗")
    print("║   CENTRIS COMMERCIAL SCRAPER - SIMPLE VERSION                 ║")
    print("╚═══════════════════════════════════════════════════════════════╝")
    
    print("\nSTEPS:")
    print("  1. Chrome will open")
    print("  2. YOU manually:")
    print("     - Go to Centris.ca → Commercial")
    print("     - Close any popups")
    print("     - Set all your filters")
    print("     - Click SEARCH")
    print("  3. When you SEE the search results, come back here")
    print("  4. Press ENTER")
    print("  5. Script will extract data from the cards")
    print("\n" + "="*70)
    
    input("\nPress ENTER to open Chrome...")
    
    driver = setup_driver()
    
    print("\n✓ Chrome opened")
    print("\nDO NOT CLOSE THE BROWSER!")
    print("\nIn the Chrome window:")
    print("  1. Navigate to Centris Commercial")
    print("  2. Close any popups YOURSELF") 
    print("  3. Apply your filters")
    print("  4. Click SEARCH")
    print("  5. Wait for results to load")
    
    print("\nWhen you can SEE your search results...")
    input("\n→ Press ENTER here to start extracting data...")
    
    all_properties = []
    
    try:
        page_num = 1
        
        while True:
            print(f"\n{'='*70}")
            print(f"PAGE {page_num}")
            print(f"{'='*70}")
            
            # Extract from current page
            properties = scrape_current_page(driver)
            
            if not properties:
                print("\nNo properties found on this page")
                break
            
            all_properties.extend(properties)
            
            # Ask if user wants to go to next page
            print(f"\n✓ Extracted {len(properties)} properties from page {page_num}")
            print(f"  Total so far: {len(all_properties)}")
            
            response = input("\nGo to next page? (y/n): ").lower()
            
            if response != 'y':
                break
            
            print("\nYOU click the 'Next Page' button in the browser...")
            input("Then press ENTER here when next page loads...")
            
            page_num += 1
        
        # Now visit each property for details
        if all_properties:
            response = input(f"\n\nVisit all {len(all_properties)} properties for detailed data? (y/n): ").lower()
            
            if response == 'y':
                all_properties = enrich_with_details(driver, all_properties)
        
        # Save results
        if all_properties:
            df = pd.DataFrame(all_properties)
            filename = f'centris_commercial_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            
            print(f"\n{'='*70}")
            print("COMPLETE")
            print(f"{'='*70}")
            print(f"\n✓ Saved {len(all_properties)} properties to: {filename}")
            
            # Show summary
            with_income = sum(1 for p in all_properties if p.get('gross_income'))
            with_taxes = sum(1 for p in all_properties if p.get('municipal_tax'))
            
            print(f"\nData quality:")
            print(f"  With gross income: {with_income}/{len(all_properties)}")
            print(f"  With tax data:     {with_taxes}/{len(all_properties)}")
            
            if with_income < len(all_properties):
                print(f"\n⚠ {len(all_properties) - with_income} properties missing gross income")
                print(f"  You can fill it manually in the CSV before analysis")
        
    except KeyboardInterrupt:
        print("\n\n⚠ Stopped by user")
    
    finally:
        print("\nKeeping browser open for 10 seconds...")
        time.sleep(10)
        driver.quit()
        print("✓ Browser closed")


if __name__ == '__main__':
    main()
