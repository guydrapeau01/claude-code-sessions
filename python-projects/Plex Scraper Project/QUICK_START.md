# QUICK START GUIDE
## Montreal/Laval Plex Analyzer

### 🚀 Get Started in 3 Steps

## Step 1: Prepare Your Property Data

**Option A: Use the CSV Template (Recommended)**

1. A template file `property_input_template.csv` has been created
2. Open it in Excel or Google Sheets
3. Add your properties:

```
Address,Market_Price,Num_Units,Gross_Annual_Income,Listing_URL,Notes
123 Main St Montreal,850000,6,65000,https://...,Nice area
456 Park Ave Laval,920000,6,0,https://...,Will estimate income
```

**Option B: Create Properties Programmatically**

Edit this code snippet:

```python
properties = [
    {
        'address': 'Your property address',
        'market_price': 850000,
        'num_units': 6,
        'gross_annual_income': 65000  # Set to 0 to auto-estimate
    }
]
```

## Step 2: Run the Analysis

### Quick Method (Demo):
```bash
python main_workflow.py
```

### Custom Analysis:
```python
from main_workflow import run_complete_analysis

# Your properties here
properties = [...]

results, summary, report_file = run_complete_analysis(properties)
```

## Step 3: Review Your Report

The Excel report contains:
- **Summary Sheet**: All properties compared side-by-side
- **Detail Sheets**: Individual analysis for each property
- **Parameters Sheet**: Assumptions and financing scenarios

### 📊 What to Look For:

✅ **Value Ratio > 1.0** = Good opportunity
- Property is trading below its economic value
- The higher above 1.0, the better the deal

✅ **Positive Monthly Cashflow**
- Property generates income after all expenses

✅ **High Cash ROI**
- Good return on your down payment
- Target: >15-20% for multi-family

---

## 💡 Pro Tips

### Finding Properties:
1. Browse Centris.ca, DuProprio.com
2. Search: "6-plex Montreal", "multiplex Laval", "immeuble à revenus"
3. Copy property details into CSV template

### Income Estimation:
If you don't know actual rent:
- Montreal/Laval: ~$1,090/unit/month
- Montérégie: ~$900/unit/month

The tool will auto-estimate if you enter 0 for income.

### Interpreting Results:

**Example Good Deal:**
```
Market Price: $850,000
Economic Value: $1,020,000
Value Ratio: 120% ✓
Monthly Cashflow: +$2,400
Cash ROI: 28%
→ STRONG BUY SIGNAL
```

**Example Overpriced:**
```
Market Price: $1,200,000
Economic Value: $950,000
Value Ratio: 79% ✗
Monthly Cashflow: -$800
→ AVOID
```

---

## 🔧 Customization

### Adjust Vacancy Rate
Default: 3%

```python
from property_analyzer import PropertyAnalyzer

analyzer = PropertyAnalyzer(vacancy_rate=0.05)  # 5%
```

### Change Expense Ratio
Default: 25%

```python
analyzer = PropertyAnalyzer(expense_ratio=0.30)  # 30%
```

### Custom Rental Rates

Edit `property_scraper.py`:
```python
def estimate_income(self, num_units, location='montreal'):
    monthly_rent_per_unit = 1200  # Your rate here
```

---

## 📋 Workflow Example

### Real-World Usage:

1. **Monday Morning**: Browse Centris for new 5+ unit listings
2. **Copy to CSV**: Add 10 interesting properties to template
3. **Run Analysis**: `python main_workflow.py`
4. **Review Report**: Sort by Value Ratio
5. **Visit Top 3**: Schedule showings for properties with ratio > 1.15
6. **Due Diligence**: Verify actual rents, inspect buildings
7. **Make Offer**: Use Economic Value to negotiate price

---

## ⚠️ Important Notes

1. **Verify Income**: Always confirm actual rents with landlord
2. **Inspect Property**: Get professional building inspection
3. **Review Expenses**: Check actual tax bills, insurance, maintenance
4. **Check Market**: Economic value is theoretical - market may differ
5. **Get Advice**: Consult accountant, lawyer, mortgage broker

**The tool is an analysis aid, not investment advice!**

---

## 🆘 Troubleshooting

**"Module not found"**
```bash
pip install pandas openpyxl numpy
```

**"File not found"**
- Make sure you're in the right directory
- Check CSV file path is correct

**Unexpected results**
- Verify input data is correct
- Check income is annual (not monthly)
- Confirm price and units are accurate

---

## 📞 Next Steps

1. Start with the sample analysis (already generated!)
2. Open `Montreal_Plex_Analysis.xlsx`
3. Review how it works with sample data
4. Replace with your own properties
5. Start finding deals!

**Happy Hunting! 🏢💰**
