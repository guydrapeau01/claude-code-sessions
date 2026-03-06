# Montreal/Laval Plex Analyzer

Automated property analysis tool for multi-unit residential buildings (5+ units) in the Montreal/Laval area, using economic value methodology based on TGA (Taux Global d'Actualisation).

## Overview

This tool helps real estate investors:
1. **Search** for 5+ unit properties in Montreal/Laval
2. **Calculate** economic value using TGA methodology
3. **Compare** multiple CMHC financing scenarios
4. **Identify** properties trading below their economic value
5. **Generate** detailed Excel reports with analysis

## Key Concepts

### Economic Value (Valeur Économique)
The true value of a property based on its income-generating capacity:
```
Economic Value = Net Operating Income (NOI) / TGA
```

### TGA (Taux Global d'Actualisation)
Overall capitalization rate that accounts for debt service requirements:
```
TGA = Annuity Factor / RCD
where:
  Annuity Factor = [r(1+r)^n] / [(1+r)^n - 1]
  RCD = Ratio de Couverture de la Dette (typically 1.1)
```

### Value Ratio
```
Value Ratio = Economic Value / Market Price
```
- **> 1.0**: Property is undervalued (good opportunity)
- **< 1.0**: Property is overpriced
- **≈ 1.0**: Property is fairly priced

## Installation

### Prerequisites
```bash
Python 3.8+
pip install pandas openpyxl numpy
```

### Setup
```bash
git clone <repository>
cd montreal_plex_analyzer
```

## Usage

### Method 1: Manual Property Entry (CSV)

1. **Generate template:**
```python
python property_scraper.py
```

This creates `property_input_template.csv`

2. **Fill in your properties:**
```csv
Address,Market_Price,Num_Units,Gross_Annual_Income,Listing_URL,Notes
123 Rue Example Montreal,850000,6,65000,https://...,Good location
```

3. **Run analysis:**
```python
from main_workflow import run_complete_analysis
from property_scraper import ManualPropertyLoader

properties = ManualPropertyLoader.load_from_csv('my_properties.csv')
results, summary, report_file = run_complete_analysis(properties)
```

### Method 2: Programmatic Entry

```python
from main_workflow import run_complete_analysis

properties = [
    {
        'address': '1825, Rue Sainte-Hélène, Longueuil',
        'market_price': 1199999,
        'num_units': 6,
        'gross_annual_income': 84660  # Optional - will estimate if 0
    },
    # Add more properties...
]

results, summary, report_file = run_complete_analysis(properties)
```

### Method 3: Quick Demo

```bash
python main_workflow.py
```

This runs analysis on sample properties and generates a report.

## Files

- **property_analyzer.py**: Core analysis engine with TGA calculations
- **property_scraper.py**: Property search and data loading utilities
- **main_workflow.py**: Complete workflow orchestration
- **README.md**: This file

## CMHC Financing Scenarios

The tool analyzes 5 standard CMHC financing scenarios:

| Scenario | Down Payment | CMHC Premium | Amortization | Interest Rate |
|----------|--------------|--------------|--------------|---------------|
| 100pts | 5% | 2.55% | 50 years | 4.0% |
| 70pts | 15% | 3.30% | 45 years | 4.0% |
| 50pts | 20% | 3.30% | 40 years | 4.0% |
| SCHL | 25% | 5.50% | 40 years | 3.9% |
| Conventional | 35%+ | 5.50% | 40 years | 5.4% |

The best scenario is automatically selected for each property.

## Analysis Parameters

Default values (customizable):
- **Vacancy Rate**: 3%
- **Operating Expense Ratio**: 25% of gross income
- **RCD (Debt Coverage Ratio)**: 1.1

## Output Reports

Excel reports include:

1. **Summary Sheet**: Comparative overview of all properties
2. **Detail Sheets**: Individual property analysis (one per property)
3. **Parameters Sheet**: Analysis assumptions and financing scenarios

### Key Metrics Reported

- Economic Value
- Value Ratio (Economic Value / Market Price)
- TGA (Taux Global d'Actualisation)
- Down Payment Amount & Percentage
- Monthly Mortgage Payment
- Monthly Cashflow
- Annual Cashflow
- Cash ROI (Return on Down Payment)
- NOI (Net Operating Income)
- MRB (Multiplicateur Revenu Brut)
- MRN (Multiplicateur Revenu Net)
- Value per Unit

## Income Estimation

If gross annual income is unknown, the tool estimates based on market rents:
- **Montreal/Laval**: $1,090/unit/month
- **Montérégie/Lanaudière**: $900/unit/month

```python
estimated_income = num_units × monthly_rent × 12
```

## Examples

### Finding Good Deals

Properties with Value Ratio > 1.0 are trading below their economic value:

```
Address: 1, Rue des Pins, Chambly
Market Price: $839,000
Economic Value: $1,020,450
Value Ratio: 121.6% ✓ GOOD DEAL
Monthly Cashflow: $2,450
Cash ROI: 22.3%
```

### Avoiding Overpriced Properties

Properties with Value Ratio < 1.0 are overpriced:

```
Address: 123 Rue Expensive, Montreal
Market Price: $1,500,000
Economic Value: $1,200,000
Value Ratio: 80.0% ✗ OVERPRICED
```

## Customization

### Adjust Analysis Parameters

```python
from property_analyzer import PropertyAnalyzer

analyzer = PropertyAnalyzer(
    vacancy_rate=0.05,      # 5% vacancy
    expense_ratio=0.30      # 30% operating expenses
)
```

### Add Custom Financing Scenarios

```python
from property_analyzer import FinancingScenario, FINANCING_SCENARIOS

FINANCING_SCENARIOS['Custom'] = FinancingScenario(
    name='Custom Financing',
    rpv=0.80,                    # 80% LTV
    prime_schl_pct=0.04,         # 4% CMHC premium
    droit_schl=900,              # CMHC fee
    amortization_years=35,       # 35 year amortization
    interest_rate=0.045,         # 4.5% interest
    rcd=1.15                     # 1.15 debt coverage ratio
)
```

## Web Scraping Integration

**Note**: Web scraping should respect website terms of service.

### Recommended Approaches:

1. **Centris Professional API** (requires license)
   - Official API for Quebec real estate data
   - Provides structured property data
   - Subscription required

2. **Manual CSV Entry**
   - Most straightforward approach
   - Copy/paste from listings into CSV
   - No API restrictions

3. **DuProprio RSS/API** (if available)
   - Check for official data feeds
   - Respect rate limits and terms

### Example API Integration Template:

```python
class CentrisAPI:
    def search_properties(self, min_units=5, regions=['montreal', 'laval']):
        # Implement API calls here
        pass
```

## Advanced Features

### Batch Processing

```python
import glob
from property_scraper import ManualPropertyLoader

# Load all CSV files in a directory
csv_files = glob.glob('properties/*.csv')

all_properties = []
for csv_file in csv_files:
    properties = ManualPropertyLoader.load_from_csv(csv_file)
    all_properties.extend(properties)

results, summary, report = run_complete_analysis(all_properties)
```

### Tracking Over Time

```python
import sqlite3
from datetime import datetime

# Save analysis to database
conn = sqlite3.connect('property_tracking.db')

for result in results:
    conn.execute("""
        INSERT INTO analyses 
        (date, address, market_price, economic_value, value_ratio)
        VALUES (?, ?, ?, ?, ?)
    """, (
        datetime.now(),
        result['address'],
        result['market_price'],
        result['scenarios'][result['best_scenario']]['economic_value'],
        result['best_value_ratio']
    ))

conn.commit()
```

## Tips for Finding Properties

1. **Centris.ca**: Official Quebec real estate listing service
2. **DuProprio.com**: For-sale-by-owner listings
3. **Kijiji/Facebook Marketplace**: Sometimes good deals
4. **Real Estate Agents**: Specialized in multi-family
5. **MLS Listings**: Through licensed agents

### Search Terms:
- "6-plex Montreal"
- "multiplex Laval"
- "immeuble à revenus" (revenue building)
- "6 logements" (6 units)
- "multiunit residential"

## Troubleshooting

### "Module not found" errors
```bash
pip install pandas openpyxl numpy
```

### Income estimation seems off
Adjust the default rental rates in `property_scraper.py`:
```python
def estimate_income(self, num_units, location='montreal'):
    monthly_rent_per_unit = 1200  # Adjust this value
    ...
```

### TGA calculations differ from spreadsheet
Verify the financing scenario parameters match your original Excel model.

## Contributing

To add new features:
1. Fork the repository
2. Create a feature branch
3. Add tests
4. Submit a pull request

## License

[Your License Here]

## Disclaimer

This tool is for informational purposes only. Always:
- Verify property data independently
- Conduct proper due diligence
- Consult with real estate professionals
- Review actual income statements and expenses
- Obtain professional inspections
- Seek legal and financial advice

The analysis is only as good as the input data. Garbage in, garbage out!

## Support

For questions or issues:
- Email: [your-email]
- GitHub Issues: [repo-url]/issues

---

**Happy Investing! 🏢💰**
