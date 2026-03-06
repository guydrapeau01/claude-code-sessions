#!/usr/bin/env python3
"""
Run Analysis on Properties from CSV File

This script loads properties from property_input_template.csv
and runs the complete economic value analysis.

Usage:
    python run_my_properties.py
"""

from main_workflow import run_complete_analysis
from property_scraper import ManualPropertyLoader

def main():
    print("="*80)
    print("LOADING PROPERTIES FROM CSV FILE")
    print("="*80)
    print("\nReading: property_input_template.csv")
    
    # Load properties from CSV
    try:
        properties = ManualPropertyLoader.load_from_csv('property_input_template.csv')
        
        if not properties:
            print("\n⚠ No properties found in CSV file!")
            print("\nMake sure property_input_template.csv has properties listed.")
            print("Example format:")
            print("Address,Market_Price,Num_Units,Gross_Annual_Income,Listing_URL,Notes")
            print("123 Main St,850000,6,65000,https://...,Great location")
            return
        
        print(f"\n✓ Found {len(properties)} properties to analyze:")
        for i, prop in enumerate(properties, 1):
            print(f"  {i}. {prop['address']} - ${prop['market_price']:,} ({prop['num_units']} units)")
        
        # Clean properties - pass all expense details
        clean_properties = []
        for prop in properties:
            clean_prop = {
                'address': prop['address'],
                'market_price': prop['market_price'],
                'num_units': prop['num_units'],
                'gross_annual_income': prop.get('gross_annual_income', 0)
            }
            
            # Add optional expense details if provided
            if 'municipal_tax_year' in prop:
                clean_prop['municipal_tax'] = prop['municipal_tax_year']
            if 'school_tax_year' in prop:
                clean_prop['school_tax'] = prop['school_tax_year']
            if 'utilities_month' in prop:
                clean_prop['utilities_monthly'] = prop['utilities_month']
            if 'property_mgmt_pct' in prop:
                clean_prop['property_mgmt_pct'] = prop['property_mgmt_pct']
            if 'maintenance_pct' in prop:
                clean_prop['maintenance_pct'] = prop['maintenance_pct']
            if 'insurance_pct' in prop:
                clean_prop['insurance_pct'] = prop['insurance_pct']
            if 'other_pct' in prop:
                clean_prop['other_pct'] = prop['other_pct']
            
            clean_properties.append(clean_prop)
        
        # Run complete analysis
        print("\n" + "="*80)
        print("RUNNING ANALYSIS")
        print("="*80 + "\n")
        
        results, summary, report_file = run_complete_analysis(
            clean_properties,
            'My_Properties_Analysis.xlsx'
        )
        
        print("\n" + "="*80)
        print("✓ ANALYSIS COMPLETE!")
        print("="*80)
        print(f"\nExcel report saved to: My_Properties_Analysis.xlsx")
        print("\nOpen the Excel file to see:")
        print("  • Summary comparison of all properties")
        print("  • Detailed analysis for each property")
        print("  • All 5 financing scenarios")
        print("  • Economic value vs market price")
        
    except FileNotFoundError:
        print("\n✗ Error: property_input_template.csv not found!")
        print("\nMake sure you're running this from the montreal_plex_analyzer folder")
        print("and that property_input_template.csv exists.")
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        print("\nCheck that your CSV file is formatted correctly:")
        print("Address,Market_Price,Num_Units,Gross_Annual_Income,Listing_URL,Notes")


if __name__ == '__main__':
    main()
