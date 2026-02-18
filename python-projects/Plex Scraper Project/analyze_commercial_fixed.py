#!/usr/bin/env python3
"""Analyze Centris Commercial properties"""
import pandas as pd
import sys, os
from run_analysis import analyze_property, generate_report

csv_file = 'centris_commercial_20260217_180409.csv'

print(f"\nLoading: {csv_file}")
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

analyses = []
for prop in properties:
    result = analyze_property(prop)
    if result:
        analyses.append(result)
        r = result['best_value_ratio']
        icon = 'GOOD' if r >= 1.0 else 'SKIP'
        print(f"[{icon}] {prop['address'][-60:]}")
        print(f"  Ratio: {r:.1%} | CF: ${result['best_monthly_cf']:,.0f}/mo | ROI: {result['best_cash_roi']:.1%}")

if analyses:
    report = generate_report(analyses)
    best = max(analyses, key=lambda x: x['best_value_ratio'])
    print(f"\n🏆 BEST: {best['address'][-60:]}")
    print(f"   Ratio: {best['best_value_ratio']:.1%} | Price: ${best['price']:,.0f}")
    print(f"\n✓ Report: {report}")
else:
    print("\n⚠ No valid analyses")
