#!/usr/bin/env python3
"""
Property Scraper for Montreal/Laval area
Searches for 5+ unit multi-family properties
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from typing import List, Dict, Optional
from datetime import datetime
import json


class PropertyScraper:
    """Scrapes real estate listings for multi-unit properties"""
    
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
    def search_duproprio(self, min_units: int = 5, 
                        regions: List[str] = None) -> List[Dict]:
        """
        Search DuProprio for multi-unit properties
        Note: This is a template - actual implementation would need to handle
        the website's structure and respect their terms of service
        """
        
        if regions is None:
            regions = ['montreal', 'laval']
        
        properties = []
        
        # This is a placeholder - actual scraping would go here
        # For now, we'll return sample structure
        
        print("Note: Web scraping requires careful implementation to respect website terms")
        print("Consider using official APIs when available (Centris API, etc.)")
        
        return properties
    
    def search_centris(self, min_units: int = 5) -> List[Dict]:
        """
        Search Centris for multi-unit properties
        Centris has an API that can be used with proper authentication
        """
        
        print("Centris API integration would go here")
        print("Requires API credentials from Centris")
        
        return []
    
    def estimate_income(self, num_units: int, location: str = 'montreal') -> float:
        """
        Estimate gross annual income based on market rents
        Montreal/Laval: ~$1,090/month per unit
        Montérégie/Lanaudière: ~$900/month per unit
        """
        
        location_lower = location.lower()
        
        if any(region in location_lower for region in ['montreal', 'montréal', 'laval']):
            monthly_rent_per_unit = 1090
        else:
            monthly_rent_per_unit = 900
        
        annual_income = num_units * monthly_rent_per_unit * 12
        
        return annual_income
    
    def parse_price(self, price_str: str) -> Optional[float]:
        """Parse price string to float"""
        if not price_str:
            return None
        
        # Remove currency symbols and spaces
        price_str = re.sub(r'[^\d,.]', '', price_str)
        price_str = price_str.replace(',', '').replace(' ', '')
        
        try:
            return float(price_str)
        except:
            return None
    
    def extract_num_units(self, title: str, description: str = '') -> Optional[int]:
        """Extract number of units from listing text"""
        
        text = (title + ' ' + description).lower()
        
        # Look for patterns like "6-plex", "6 logements", "6 units"
        patterns = [
            r'(\d+)[\s-]?plex',
            r'(\d+)[\s-]?logements?',
            r'(\d+)[\s-]?units?',
            r'(\d+)[\s-]?appartements?'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return int(match.group(1))
        
        return None


class ManualPropertyLoader:
    """
    For manual entry of properties found through browsing
    Creates a simple input format
    """
    
    @staticmethod
    def create_template() -> pd.DataFrame:
        """Create an empty template for manual data entry"""
        
        template = pd.DataFrame(columns=[
            'Address',
            'Market_Price',
            'Num_Units',
            'Gross_Annual_Income',
            'Listing_URL',
            'Notes'
        ])
        
        return template
    
    @staticmethod
    def load_from_csv(filepath: str) -> List[Dict]:
        """Load properties from CSV file"""
        
        # Try different encodings to handle French characters
        encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
        
        df = None
        for encoding in encodings:
            try:
                df = pd.read_csv(filepath, encoding=encoding)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if df is None:
            raise ValueError(f"Could not read CSV file with any encoding. Please save as UTF-8.")
        
        properties = []
        for _, row in df.iterrows():
            prop = {
                'address': row.get('Address', ''),
                'market_price': row.get('Market_Price', 0),
                'num_units': int(row.get('Num_Units', 0)),
                'gross_annual_income': row.get('Gross_Annual_Income', 0),
                'listing_url': row.get('Listing_URL', ''),
                'notes': row.get('Notes', '')
            }
            
            # Add optional expense details
            if 'Municipal_Tax_Year' in row and pd.notna(row['Municipal_Tax_Year']):
                prop['municipal_tax_year'] = row['Municipal_Tax_Year']
            if 'School_Tax_Year' in row and pd.notna(row['School_Tax_Year']):
                prop['school_tax_year'] = row['School_Tax_Year']
            if 'Utilities_Month' in row and pd.notna(row['Utilities_Month']):
                prop['utilities_month'] = row['Utilities_Month']
            if 'Property_Mgmt_Pct' in row and pd.notna(row['Property_Mgmt_Pct']):
                prop['property_mgmt_pct'] = row['Property_Mgmt_Pct']
            if 'Maintenance_Pct' in row and pd.notna(row['Maintenance_Pct']):
                prop['maintenance_pct'] = row['Maintenance_Pct']
            if 'Insurance_Pct' in row and pd.notna(row['Insurance_Pct']):
                prop['insurance_pct'] = row['Insurance_Pct']
            if 'Other_Pct' in row and pd.notna(row['Other_Pct']):
                prop['other_pct'] = row['Other_Pct']
            
            properties.append(prop)
        
        return properties
    
    @staticmethod
    def save_template(filepath: str = 'property_input_template.csv'):
        """Save a CSV template for manual entry"""
        
        template = ManualPropertyLoader.create_template()
        
        # Add sample rows
        sample_data = [
            {
                'Address': '123 Rue Example, Montreal',
                'Market_Price': 850000,
                'Num_Units': 6,
                'Gross_Annual_Income': 60000,
                'Listing_URL': 'https://example.com/listing',
                'Notes': 'Sample property - replace with actual data'
            }
        ]
        
        template = pd.DataFrame(sample_data)
        template.to_csv(filepath, index=False)
        
        print(f"Template saved to: {filepath}")
        print("\nInstructions:")
        print("1. Open the CSV file in Excel or Google Sheets")
        print("2. Replace the sample data with actual properties")
        print("3. For Gross_Annual_Income: Use actual if known, or leave 0 to estimate")
        print("4. Save and load using ManualPropertyLoader.load_from_csv()")
        
        return filepath


def create_property_database():
    """
    Create a simple database structure to track properties over time
    """
    
    # This would connect to a SQLite database or similar
    # For now, just demonstrate the structure
    
    schema = {
        'properties': {
            'id': 'INTEGER PRIMARY KEY',
            'address': 'TEXT',
            'market_price': 'REAL',
            'num_units': 'INTEGER',
            'gross_annual_income': 'REAL',
            'listing_url': 'TEXT',
            'date_added': 'TIMESTAMP',
            'status': 'TEXT'  # active, sold, withdrawn
        },
        'analyses': {
            'id': 'INTEGER PRIMARY KEY',
            'property_id': 'INTEGER',
            'analysis_date': 'TIMESTAMP',
            'economic_value': 'REAL',
            'best_scenario': 'TEXT',
            'value_ratio': 'REAL',
            'notes': 'TEXT'
        }
    }
    
    print("Database schema for tracking properties:")
    print(json.dumps(schema, indent=2))
    
    return schema


def main():
    """Demo the scraping and loading capabilities"""
    
    print("Montreal/Laval Property Scraper\n")
    print("="*80)
    
    # Create manual entry template
    print("\n1. Creating manual entry template...")
    ManualPropertyLoader.save_template('property_input_template.csv')
    
    # Show how to estimate income
    print("\n2. Income estimation examples:")
    scraper = PropertyScraper()
    
    locations = ['Montreal', 'Laval', 'Longueuil', 'Terrebonne']
    for location in locations:
        income_6plex = scraper.estimate_income(6, location)
        print(f"   {location} - 6-plex estimated income: ${income_6plex:,.0f}/year")
    
    # Show database structure
    print("\n3. Database structure for tracking:")
    create_property_database()
    
    print("\n" + "="*80)
    print("\nNEXT STEPS:")
    print("1. Manually enter properties into 'property_input_template.csv'")
    print("2. Or develop API integration with Centris/DuProprio")
    print("3. Load properties and run analysis with property_analyzer.py")
    print("\nFor automated scraping, consider:")
    print("  - Centris Professional API (requires license)")
    print("  - DuProprio RSS feeds or API")
    print("  - Always respect robots.txt and terms of service")


if __name__ == '__main__':
    main()
