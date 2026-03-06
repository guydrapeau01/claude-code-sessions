#!/usr/bin/env python3
"""
Analyze properties from Centris Commercial CSV
"""

import pandas as pd
import sys
import os

# Import the analysis functions
from run_analysis import load_csv, analyze_property, generate_report, CONFIG

def analyze_commercial_csv(csv_file):
    """
    Load commercial CSV and run analysis
    """
    
    print("\n" + "="*70)
    print("ANALYZING CENTRIS COMMERCIAL PROPERTIES")
    print("="*70)
    print(f"\nLoading: {csv_file}")
    
    if not os.path.exists(csv_file):
        print(f"\n✗ File not found: {csv_file}")
        print("Make sure you're in the right directory")
        return
    
    # Load CSV
    try:
        for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
            try:
                df = pd.read_csv(csv_file, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
    except Exception as e:
        print(f"✗ Error reading CSV: {e}")
        return
    
    print(f"✓ Loaded {len(df)} properties")
    
    # Convert to analysis format
    properties = []
    skipped = []
    
    for _, row in df.iterrows():
        addr = row.get('address')
        price = row.get('price')
        income = row.get('gross_income')
        
        # Skip if missing critical data
        if pd.isna(addr) or pd.isna(price):
            continue
        
        if pd.isna(income) or income == 0:
            skipped.append(str(addr)[:50])
            continue
        
        properties.append({
            'address': addr,
            'price': float(price),
            'num_units': int(row.get('num_units', 0)) if not pd.isna(row.get('num_units')) else 0,
            'gross_income': float(income),
            'municipal_tax': float(row.get('municipal_tax', 0)) if not pd.isna(row.get('municipal_tax')) else 0,
            'school_tax': float(row.get('school_tax', 0)) if not pd.isna(row.get('school_tax')) else 0,
            'utilities': 0,  # Commercial properties - assume tenant-paid
            'url': row.get('url', ''),
        })
    
    print(f"\n✓ {len(properties)} properties with income data")
    
    if skipped:
        print(f"\n⚠ Skipped {len(skipped)} properties missing gross_income:")
        for s in skipped[:5]:
            print(f"  - {s}")
        if len(skipped) > 5:
            print(f"  ... and {len(skipped) - 5} more")
        print(f"\nTo analyze them, open {csv_file} and fill in 'gross_income' column")
    
    if not properties:
        print("\n✗ No properties with income data to analyze!")
        print(f"Open {csv_file} and add income values")
        return
    
    # Run analysis
    print("\n" + "="*70)
    print("CALCULATING ECONOMIC VALUE")
    print("="*70 + "\n")
    
    analyses = []
    
    for prop in properties:
        result = analyze_property(prop)
        if result:
            analyses.append(result)
            r = result['best_value_ratio']
            icon = '🟢' if r >= 1.15 else '🟡' if r >= 1.0 else '🔴'
            print(f"{icon} {prop['address'][:50]}")
            print(f"   Value Ratio: {r:.1%}  |  "
                  f"Cashflow: ${result['best_monthly_cf']:,.0f}/mo  |  "
                  f"ROI: {result['best_cash_roi']:.1%}")
    
    if not analyses:
        print("\n⚠ No properties could be analyzed (all had negative NOI)")
        return
    
    # Generate report
    print("\n" + "="*70)
    print("GENERATING EXCEL REPORT")
    print("="*70)
    
    report_file = generate_report(analyses)
    
    # Summary
    good = [a for a in analyses if a['best_value_ratio'] >= 1.15]
    ok = [a for a in analyses if 1.0 <= a['best_value_ratio'] < 1.15]
    bad = [a for a in analyses if a['best_value_ratio'] < 1.0]
    
    print("\n" + "="*70)
    print("RESULTS SUMMARY")
    print("="*70)
    print(f"\n  Excellent (>115%): {len(good)}")
    print(f"  Good (100-115%):   {len(ok)}")
    print(f"  Overpriced (<100%): {len(bad)}")
    
    if analyses:
        best = sorted(analyses, key=lambda x: x['best_value_ratio'], reverse=True)[0]
        
        print(f"\n  🏆 BEST DEAL:")
        print(f"  {best['address']}")
        print(f"  Price:          ${best['price']:,.0f}")
        print(f"  Economic Value: ${best['best_econ_value']:,.0f}")
        print(f"  Value Ratio:    {best['best_value_ratio']:.1%}")
        print(f"  Monthly CF:     ${best['best_monthly_cf']:,.0f}")
        print(f"  Cash ROI:       {best['best_cash_roi']:.1%}")
        if best.get('url'):
            print(f"  URL:            {best['url']}")
    
    print(f"\n  📊 Full report: {report_file}")
    print()


if __name__ == '__main__':
    
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]
    else:
        # Find most recent centris_commercial CSV
        import glob
        files = glob.glob('centris_commercial_*.csv')
        if files:
            csv_file = max(files, key=os.path.getctime)
            print(f"Using most recent file: {csv_file}")
        else:
            print("\nNo centris_commercial_*.csv file found!")
            print("\nUsage:")
            print("  python analyze_commercial.py your_file.csv")
            print("\nOr just run it and it will use the most recent file")
            sys.exit(1)
    
    analyze_commercial_csv(csv_file)
