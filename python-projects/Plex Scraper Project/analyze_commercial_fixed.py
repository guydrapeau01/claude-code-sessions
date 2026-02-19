#!/usr/bin/env python3
"""Analyze Centris Commercial properties"""
import pandas as pd
import glob, os
from run_analysis import analyze_property, generate_report

# Find most recent CSV file
csv_files = glob.glob('centris_commercial_*.csv')
if not csv_files:
    print("\n✗ No centris_commercial_*.csv file found!")
    print("Run the scraper first: python centris_simple_scraper.py")
    exit(1)

csv_file = max(csv_files, key=os.path.getctime)
print(f"\n✓ Using most recent file: {csv_file}")

df = pd.read_csv(csv_file, encoding='utf-8-sig')
print(f"✓ Loaded {len(df)} properties")

properties = []
for i, row in df.iterrows():
    # Use URL as address if address is missing
    addr = row.get('address') if pd.notna(row.get('address')) and str(row.get('address')).strip() else row.get('url', f'Property {i+1}')
    
    if pd.isna(row.get('gross_income')) or row.get('gross_income') == 0:
        continue
    
    properties.append({
        'address': str(addr),
        'price': float(row.get('price')),
        'num_units': int(row.get('num_units', 0)) if pd.notna(row.get('num_units')) else 0,
        'gross_income': float(row.get('gross_income')),
        'municipal_tax': float(row.get('municipal_tax', 0)) if pd.notna(row.get('municipal_tax')) else 0,
        'school_tax': float(row.get('school_tax', 0)) if pd.notna(row.get('school_tax')) else 0,
        'utilities': 0,
        'url': row.get('url', ''),
    })

print(f"✓ {len(properties)} properties ready to analyze\n")

print("="*70)
print("ANALYZING PROPERTIES")
print("="*70 + "\n")

analyses = []
for prop in properties:
    result = analyze_property(prop)
    if result:
        analyses.append(result)
        r = result['best_value_ratio']
        icon = '[GOOD]' if r >= 1.0 else '[SKIP]'
        print(f"{icon} {prop['address'][-60:]}")
        print(f"       Ratio: {r:.1%} | CF: ${result['best_monthly_cf']:,.0f}/mo | ROI: {result['best_cash_roi']:.1%}")

if analyses:
    print("\n" + "="*70)
    print("GENERATING EXCEL REPORT")
    print("="*70)
    
    report = generate_report(analyses)
    
    good = [a for a in analyses if a['best_value_ratio'] >= 1.15]
    ok = [a for a in analyses if 1.0 <= a['best_value_ratio'] < 1.15]
    bad = [a for a in analyses if a['best_value_ratio'] < 1.0]
    
    print("\n" + "="*70)
    print("RESULTS")
    print("="*70)
    print(f"\n  Excellent (>115%): {len(good)}")
    print(f"  Good (100-115%):   {len(ok)}")
    print(f"  Overpriced (<100%): {len(bad)}")
    
    best = max(analyses, key=lambda x: x['best_value_ratio'])
    print(f"\n  🏆 BEST DEAL:")
    print(f"  {best['address'][-70:]}")
    print(f"  Price:          ${best['price']:,.0f}")
    print(f"  Economic Value: ${best['best_econ_value']:,.0f}")
    print(f"  Value Ratio:    {best['best_value_ratio']:.1%}")
    print(f"  Monthly CF:     ${best['best_monthly_cf']:,.0f}")
    print(f"  Cash ROI:       {best['best_cash_roi']:.1%}")
    
    print(f"\n  📊 Excel report: {report}")
    print()
else:
    print("\n⚠ No valid analyses (all negative NOI)")
