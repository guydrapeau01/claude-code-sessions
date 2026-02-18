#!/usr/bin/env python3
"""
Centris Scraper - ROBUST VERSION
Handles all the clicking, cookies, and navigation properly
"""

import time
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd


def setup_driver():
    """Create Chrome driver"""
    options = Options()
    # Don't run headless - we want to SEE what's happening
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


def accept_cookies(driver):
    """Accept cookies/privacy popup if it appears"""
    print("  Checking for cookie/privacy popups...")
    
    # List of possible button texts and selectors
    button_texts = [
        'Accepter',
        'Accept',
        'Tout accepter',
        'Accept all',
        'J\'accepte',
        'Continuer',
        'Continue',
        'OK',
        'Fermer',
        'Close',
        'Enregistrer',  # Save settings button
        'Save',
        'Confirmer',
        'Confirm'
    ]
    
    # Try multiple strategies
    strategies = [
        # Strategy 1: By button text
        lambda: driver.find_element(By.XPATH, 
            f"//button[{' or '.join([f'contains(text(), \"{text}\")' for text in button_texts])}]"),
        
        # Strategy 2: By ID
        lambda: driver.find_element(By.XPATH, 
            "//button[contains(@id, 'accept') or contains(@id, 'cookie') or contains(@id, 'privacy') or contains(@id, 'save') or contains(@id, 'confirm')]"),
        
        # Strategy 3: By class
        lambda: driver.find_element(By.XPATH,
            "//button[contains(@class, 'accept') or contains(@class, 'cookie') or contains(@class, 'consent') or contains(@class, 'save')]"),
        
        # Strategy 4: Any button in a popup/modal
        lambda: driver.find_element(By.XPATH,
            "//div[contains(@class, 'modal') or contains(@class, 'popup') or contains(@class, 'dialog')]//button[1]"),
        
        # Strategy 5: Specific Centris privacy button
        lambda: driver.find_element(By.CSS_SELECTOR, 
            "button[data-action='accept'], button.btn-accept, .cookie-banner button, button[type='submit']"),
    ]
    
    clicked = False
    for i, strategy in enumerate(strategies, 1):
        try:
            button = WebDriverWait(driver, 2).until(
                lambda d: strategy()
            )
            # Scroll into view and click
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            time.sleep(0.3)
            button.click()
            print(f"  ✓ Clicked popup button (strategy {i})")
            time.sleep(1)
            clicked = True
            # Don't break - there might be a second popup
        except:
            continue
    
    if not clicked:
        print("  → No popup found (or already accepted)")
    
    return clicked


def search_centris_montreal():
    """
    Go to Centris and search for plexes in Montreal/Laval
    Returns driver positioned on search results
    """
    driver = setup_driver()
    
    print("\n" + "="*70)
    print("STEP 1: OPENING CENTRIS & SEARCHING")
    print("="*70)
    
    # Go directly to plex search for Montreal
    url = "https://www.centris.ca/fr/plex~a-vendre~montreal"
    print(f"\nOpening: {url}")
    driver.get(url)
    time.sleep(3)
    
    # Accept cookies
    accept_cookies(driver)
    
    # Wait for results to load
    print("  Waiting for search results to load...")
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "property-thumbnail-item"))
        )
        print("  ✓ Search results loaded")
    except TimeoutException:
        print("  ⚠ Search results didn't load properly")
    
    return driver


def get_listing_cards(driver):
    """
    Get all property cards from current page
    Returns list of WebElements
    """
    try:
        # Scroll to load lazy-loaded content
        for i in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # Find all property cards
        cards = driver.find_elements(By.CLASS_NAME, "property-thumbnail-item")
        print(f"\n  Found {len(cards)} property cards on this page")
        return cards
    except Exception as e:
        print(f"  Error getting cards: {e}")
        return []


def click_property_card(driver, card, index):
    """
    Click on a property card to open details
    Returns True if successful
    """
    try:
        # Find the clickable link in this card
        link = card.find_element(By.CSS_SELECTOR, "a.property-thumbnail-summary-link")
        
        # Get the URL before clicking (backup)
        href = link.get_attribute('href')
        
        # Scroll card into view
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
        time.sleep(0.5)
        
        # Try to click
        try:
            link.click()
        except:
            # If click fails, navigate directly
            driver.get(href)
        
        time.sleep(2)
        
        # Wait for detail page to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "h1"))
        )
        
        return True
        
    except Exception as e:
        print(f"    ✗ Could not click card {index}: {e}")
        return False


def extract_property_data(driver):
    """
    Extract all data from property detail page
    Returns dict with property data
    """
    try:
        # Get all text on page for parsing
        page_text = driver.find_element(By.TAG_NAME, 'body').text
        
        data = {}
        
        # ADDRESS
        try:
            h1 = driver.find_element(By.TAG_NAME, 'h1')
            data['address'] = h1.text.strip()
        except:
            data['address'] = "Address not found"
        
        # PRICE - look for asking price
        data['price'] = extract_price(driver, page_text)
        
        # NUMBER OF UNITS
        data['num_units'] = extract_units(page_text)
        
        # GROSS INCOME (Revenus bruts)
        data['gross_income'] = extract_income(page_text)
        
        # MUNICIPAL TAX
        data['municipal_tax'] = extract_tax(page_text, 'municipal')
        
        # SCHOOL TAX
        data['school_tax'] = extract_tax(page_text, 'school')
        
        # UTILITIES
        data['utilities'] = extract_utilities(page_text)
        
        # URL
        data['url'] = driver.current_url
        
        # Timestamp
        data['scraped_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        
        return data
        
    except Exception as e:
        print(f"    ✗ Error extracting data: {e}")
        return None


def extract_price(driver, page_text):
    """Extract asking price"""
    
    # Try to find price element directly
    selectors = [
        ".price",
        "[class*='price']",
        ".asking-price",
    ]
    
    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for el in elements:
                text = el.text.strip()
                # Look for price pattern
                match = re.search(r'\$?\s*([\d\s]+)', text.replace('\xa0', ' '))
                if match:
                    price_str = match.group(1).replace(' ', '').replace(',', '')
                    price = float(price_str)
                    if 100000 < price < 10000000:  # Sanity check
                        return price
        except:
            continue
    
    # Fallback: regex on page text
    patterns = [
        r'Prix demandé[:\s]*\$?\s*([\d\s]+)',
        r'Asking price[:\s]*\$?\s*([\d\s]+)',
        r'\$\s*([\d\s]{3,})\s',
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
                continue
    
    return None


def extract_units(page_text):
    """Extract number of units"""
    
    patterns = [
        r'(\d+)\s*logements?',
        r'(\d+)\s*units?',
        r'(\d+)\s*appartements?',
        r'(\d+)[\s-]plex',
        r'Nombre de logements?\s*:?\s*(\d+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            units = int(match.group(1))
            if 2 <= units <= 50:  # Sanity check
                return units
    
    return None


def extract_income(page_text):
    """Extract gross annual income (revenus bruts)"""
    
    patterns = [
        r'Revenus? bruts?\s*:?\s*\$?\s*([\d\s]+)',
        r'Gross revenue\s*:?\s*\$?\s*([\d\s]+)',
        r'Revenus? annuels?\s*:?\s*\$?\s*([\d\s]+)',
        r'Total loyers\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            income_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                income = float(income_str)
                if 5000 < income < 1000000:  # Sanity check
                    return income
            except:
                continue
    
    return None


def extract_tax(page_text, tax_type):
    """Extract municipal or school tax"""
    
    if tax_type == 'municipal':
        patterns = [
            r'[Tt]axes? municipales?\s*:?\s*\$?\s*([\d\s]+)',
            r'Municipal tax[^\d]*([\d\s]+)',
            r'[Ii]mpôts? fonciers?\s*:?\s*\$?\s*([\d\s]+)',
        ]
    else:  # school
        patterns = [
            r'[Tt]axes? scolaires?\s*:?\s*\$?\s*([\d\s]+)',
            r'School tax[^\d]*([\d\s]+)',
        ]
    
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            tax_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                tax = float(tax_str)
                if 100 < tax < 100000:  # Sanity check
                    return tax
            except:
                continue
    
    return None


def extract_utilities(page_text):
    """Extract utilities (usually monthly)"""
    
    patterns = [
        r'[ÉEé]nergie\s*:?\s*\$?\s*([\d\s]+)',
        r'[Cc]hauffage\s*:?\s*\$?\s*([\d\s]+)',
        r'[Uu]tilities?\s*:?\s*\$?\s*([\d\s]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            util_str = match.group(1).replace(' ', '').replace(',', '').replace('\xa0', '')
            try:
                util = float(util_str)
                if 50 < util < 10000:  # Sanity check
                    return util
            except:
                continue
    
    return None


def go_back_to_search(driver):
    """Go back to search results"""
    print("    → Going back to search results...")
    driver.back()
    time.sleep(2)
    
    # Check for any popups again
    accept_cookies(driver)


def scrape_centris(max_properties=20):
    """
    Main scraping function
    Scrapes up to max_properties from Centris
    """
    
    print("\n" + "="*70)
    print("CENTRIS SCRAPER - INTERACTIVE VERSION")
    print("="*70)
    print(f"\nYou will see Chrome open and navigate Centris automatically.")
    print(f"Scraping up to {max_properties} properties...")
    print("\nPress Ctrl+C at any time to stop.\n")
    
    driver = search_centris_montreal()
    
    all_properties = []
    properties_scraped = 0
    
    try:
        page_number = 1
        
        while properties_scraped < max_properties:
            
            print("\n" + "-"*70)
            print(f"Page {page_number} - Getting property cards...")
            print("-"*70)
            
            cards = get_listing_cards(driver)
            
            if not cards:
                print("\n  No more property cards found.")
                break
            
            # Process each card on this page
            card_index = 0
            while card_index < len(cards) and properties_scraped < max_properties:
                
                print(f"\n[{properties_scraped + 1}/{max_properties}] Processing property {card_index + 1}/{len(cards)} on page {page_number}...")
                
                # RE-FIND the cards to avoid stale element
                cards = get_listing_cards(driver)
                if card_index >= len(cards):
                    break
                
                card = cards[card_index]
                
                # Click on property
                if click_property_card(driver, card, card_index):
                    
                    # Extract data from detail page
                    data = extract_property_data(driver)
                    
                    if data and data.get('price'):
                        all_properties.append(data)
                        properties_scraped += 1
                        
                        # Show what we got
                        print(f"    ✓ {data.get('address', 'N/A')[:50]}")
                        print(f"      Price: ${data.get('price', 0):,.0f}")
                        print(f"      Units: {data.get('num_units', '?')}")
                        income = data.get('gross_income') or 0
                        print(f"      Income: ${income:,.0f}/year" if income else "      Income: Not shown")
                        muni = data.get('municipal_tax') or 0
                        school = data.get('school_tax') or 0
                        total_tax = muni + school
                        print(f"      Taxes: ${total_tax:,.0f}/year" if total_tax else "      Taxes: Not shown")
                    
                    # Go back to search results
                    go_back_to_search(driver)
                    
                    # Small delay to be polite
                    time.sleep(2)
                
                else:
                    print(f"    ✗ Skipped (couldn't click)")
                
                card_index += 1
            
            # Try to go to next page
            print("\n  Looking for 'Next Page' button...")
            try:
                # Find and click next page button
                next_button = driver.find_element(By.XPATH, 
                    "//a[@aria-label='Suivant' or @aria-label='Next' or contains(@class, 'next')]")
                
                if 'disabled' not in next_button.get_attribute('class'):
                    print("  ✓ Going to next page...")
                    next_button.click()
                    time.sleep(3)
                    page_number += 1
                else:
                    print("  → No more pages")
                    break
            except:
                print("  → No next page button found")
                break
    
    except KeyboardInterrupt:
        print("\n\n⚠ Stopped by user (Ctrl+C)")
    
    finally:
        driver.quit()
        print("\n✓ Browser closed")
    
    return all_properties


def save_to_csv(properties, filename='scraped_properties.csv'):
    """Save scraped properties to CSV"""
    
    if not properties:
        print("\n⚠ No properties to save")
        return
    
    df = pd.DataFrame(properties)
    
    # Reorder columns
    cols = ['address', 'price', 'num_units', 'gross_income', 
            'municipal_tax', 'school_tax', 'utilities', 'url', 'scraped_at']
    df = df[[col for col in cols if col in df.columns]]
    
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    
    print(f"\n{'='*70}")
    print(f"SCRAPING COMPLETE")
    print(f"{'='*70}")
    print(f"\n✓ Scraped {len(properties)} properties")
    print(f"✓ Saved to: {filename}")
    
    # Show summary
    with_income = sum(1 for p in properties if p.get('gross_income'))
    with_taxes = sum(1 for p in properties if p.get('municipal_tax'))
    
    print(f"\nData quality:")
    print(f"  Properties with gross income: {with_income}/{len(properties)}")
    print(f"  Properties with tax data: {with_taxes}/{len(properties)}")
    
    if with_income < len(properties):
        print(f"\n⚠ {len(properties) - with_income} properties missing gross income")
        print(f"  → Open {filename} and fill in 'gross_income' column manually")


if __name__ == '__main__':
    
    # Change this number to scrape more/fewer properties
    MAX_PROPERTIES = 20
    
    print("\n")
    print("╔═══════════════════════════════════════════════════════════════╗")
    print("║  CENTRIS SCRAPER - You'll watch it work in Chrome browser    ║")
    print("╚═══════════════════════════════════════════════════════════════╝")
    print(f"\nScraping up to {MAX_PROPERTIES} properties...")
    print("Press Ctrl+C at any time to stop early.\n")
    
    input("Press ENTER to start...")
    
    # Run the scraper
    properties = scrape_centris(max_properties=MAX_PROPERTIES)
    
    # Save results
    if properties:
        save_to_csv(properties)
        
        print(f"\n{'='*70}")
        print("NEXT STEPS")
        print(f"{'='*70}")
        print("\n1. Review scraped_properties.csv")
        print("2. Fill in any missing 'gross_income' values")
        print("3. Run: python run_analysis.py")
        print("   (It will skip the scraping step and use your CSV)")
    else:
        print("\n⚠ No properties were scraped.")
        print("   Check if Centris is blocking automated access.")
        print("   Try running again or reduce MAX_PROPERTIES.")
