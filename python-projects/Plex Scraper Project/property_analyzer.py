#!/usr/bin/env python3
"""
Montreal/Laval Plex Analyzer
Searches for 5+ unit properties and calculates economic value based on TGA methodology
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Dict, Optional
import json


@dataclass
class FinancingScenario:
    """CMHC/SCHL financing scenario"""
    name: str
    rpv: float  # Ratio prêt-valeur (Loan to Value)
    prime_schl_pct: float  # CMHC insurance premium %
    droit_schl: float  # CMHC application fee
    amortization_years: int
    interest_rate: float
    rcd: float = 1.1  # Ratio de couverture de la dette (Debt Coverage Ratio)
    
    
# Define standard CMHC financing scenarios
FINANCING_SCENARIOS = {
    '100pts': FinancingScenario('100pts (5% down)', 0.95, 0.0255, 900, 50, 0.04),
    '70pts': FinancingScenario('70pts (15% down)', 0.95, 0.033, 900, 45, 0.04),
    '50pts': FinancingScenario('50pts (20% down)', 0.85, 0.033, 900, 40, 0.04),
    'SCHL': FinancingScenario('SCHL (25% down)', 0.80, 0.055, 900, 40, 0.039),
    'Conventionnel': FinancingScenario('Conventional (35%+ down)', 0.75, 0.055, 900, 40, 0.054),
}


class PropertyAnalyzer:
    """Analyzes multi-unit properties using economic value methodology"""
    
    def __init__(self, vacancy_rate: float = 0.03, expense_ratio: float = 0.25):
        self.vacancy_rate = vacancy_rate
        self.expense_ratio = expense_ratio
        
    def calculate_tga(self, scenario: FinancingScenario) -> float:
        """
        Calculate TGA (Taux Global d'Actualisation / Overall Capitalization Rate)
        TGA = Annuity Factor / RCD
        Annuity Factor = [r(1+r)^n] / [(1+r)^n - 1]
        """
        r = scenario.interest_rate
        n = scenario.amortization_years
        
        # Calculate annuity factor
        annuity_factor = (r * (1 + r)**n) / ((1 + r)**n - 1)
        
        # TGA = Annuity Factor / RCD
        tga = annuity_factor / scenario.rcd
        
        return tga
    
    def calculate_financing(self, market_price: float, economic_value: float, 
                          scenario: FinancingScenario) -> Dict:
        """Calculate financing details for a given scenario"""
        
        # Determine loan amount base (market price or economic value, whichever is lower)
        loan_base = market_price if market_price < economic_value else economic_value
        
        # Calculate loan amount before insurance
        loan_amount_base = scenario.rpv * loan_base
        
        # Calculate CMHC insurance premium
        prime_schl = loan_amount_base * scenario.prime_schl_pct
        
        # Total loan amount (including insurance premium)
        total_loan = loan_amount_base + prime_schl
        
        # Down payment
        down_payment = market_price - total_loan
        
        # Monthly payment using annuity formula
        r_monthly = scenario.interest_rate / 12
        n_months = scenario.amortization_years * 12
        monthly_payment = total_loan * (r_monthly * (1 + r_monthly)**n_months) / \
                         ((1 + r_monthly)**n_months - 1)
        
        return {
            'down_payment': down_payment,
            'down_payment_pct': down_payment / market_price,
            'loan_amount_base': loan_amount_base,
            'prime_schl': prime_schl,
            'total_loan': total_loan,
            'monthly_payment': monthly_payment,
            'annual_payment': monthly_payment * 12
        }
    
    def calculate_economic_value(self, gross_annual_income: float, 
                                scenario: FinancingScenario,
                                municipal_tax: float = None,
                                school_tax: float = None,
                                utilities_monthly: float = None,
                                property_mgmt_pct: float = 0.05,
                                maintenance_pct: float = 0.03,
                                insurance_pct: float = 0.01,
                                other_pct: float = 0.02) -> Dict:
        """
        Calculate economic value using TGA methodology
        Economic Value = Net Operating Income / TGA
        
        Can use either:
        1. Detailed expenses (municipal_tax, school_tax, etc.)
        2. Simple expense_ratio (if detailed not provided)
        """
        
        # Calculate effective gross income (after vacancy)
        effective_gross_income = gross_annual_income * (1 - self.vacancy_rate)
        
        # Calculate operating expenses
        if municipal_tax is not None and school_tax is not None:
            # Use detailed expense breakdown
            property_taxes = municipal_tax + school_tax
            utilities = utilities_monthly * 12 if utilities_monthly else 0
            property_mgmt = gross_annual_income * property_mgmt_pct
            maintenance = gross_annual_income * maintenance_pct
            insurance = gross_annual_income * insurance_pct
            other = gross_annual_income * other_pct
            
            operating_expenses = (property_taxes + utilities + property_mgmt + 
                                 maintenance + insurance + other)
            
            expense_breakdown = {
                'municipal_tax': municipal_tax,
                'school_tax': school_tax,
                'total_taxes': property_taxes,
                'utilities': utilities,
                'property_mgmt': property_mgmt,
                'maintenance': maintenance,
                'insurance': insurance,
                'other': other
            }
        else:
            # Use simple expense ratio
            operating_expenses = gross_annual_income * self.expense_ratio
            expense_breakdown = None
        
        # Net Operating Income (NOI)
        noi = effective_gross_income - operating_expenses
        
        # Calculate TGA
        tga = self.calculate_tga(scenario)
        
        # Economic Value
        economic_value = noi / tga
        
        return {
            'gross_annual_income': gross_annual_income,
            'vacancy_amount': gross_annual_income * self.vacancy_rate,
            'effective_gross_income': effective_gross_income,
            'operating_expenses': operating_expenses,
            'expense_breakdown': expense_breakdown,
            'noi': noi,
            'tga': tga,
            'economic_value': economic_value
        }
    
    def analyze_property(self, market_price: float, gross_annual_income: float,
                        num_units: int, address: str = '',
                        municipal_tax: float = None, school_tax: float = None,
                        utilities_monthly: float = None,
                        property_mgmt_pct: float = 0.05,
                        maintenance_pct: float = 0.03,
                        insurance_pct: float = 0.01,
                        other_pct: float = 0.02) -> Dict:
        """
        Complete property analysis across all financing scenarios
        """
        
        results = {
            'address': address,
            'market_price': market_price,
            'gross_annual_income': gross_annual_income,
            'num_units': num_units,
            'scenarios': {}
        }
        
        # Analyze each financing scenario
        for scenario_name, scenario in FINANCING_SCENARIOS.items():
            
            # Calculate economic value
            econ_calc = self.calculate_economic_value(
                gross_annual_income, 
                scenario,
                municipal_tax=municipal_tax,
                school_tax=school_tax,
                utilities_monthly=utilities_monthly,
                property_mgmt_pct=property_mgmt_pct,
                maintenance_pct=maintenance_pct,
                insurance_pct=insurance_pct,
                other_pct=other_pct
            )
            economic_value = econ_calc['economic_value']
            
            # Calculate financing
            financing = self.calculate_financing(market_price, economic_value, scenario)
            
            # Calculate key metrics
            value_per_unit = economic_value / num_units
            mrb = market_price / gross_annual_income  # Multiplicateur du revenu brut
            mrn = market_price / econ_calc['noi']  # Multiplicateur du revenu net
            
            # Monthly cashflow estimate
            monthly_income = gross_annual_income / 12
            monthly_expenses = econ_calc['operating_expenses'] / 12
            monthly_cashflow = monthly_income * (1 - self.vacancy_rate) - \
                             monthly_expenses - financing['monthly_payment']
            
            # ROI on down payment
            annual_cashflow = monthly_cashflow * 12
            cash_roi = annual_cashflow / financing['down_payment'] if financing['down_payment'] > 0 else 0
            
            # Value vs price ratio (higher is better)
            value_ratio = economic_value / market_price
            
            results['scenarios'][scenario_name] = {
                **econ_calc,
                **financing,
                'value_per_unit': value_per_unit,
                'mrb': mrb,
                'mrn': mrn,
                'monthly_cashflow': monthly_cashflow,
                'annual_cashflow': annual_cashflow,
                'cash_roi': cash_roi,
                'value_ratio': value_ratio,
                'scenario_name': scenario.name
            }
        
        # Determine best scenario (highest economic value)
        best_scenario = max(results['scenarios'].items(), 
                          key=lambda x: x[1]['economic_value'])
        results['best_scenario'] = best_scenario[0]
        results['best_value_ratio'] = best_scenario[1]['value_ratio']
        
        return results
    
    def create_summary_report(self, properties: List[Dict]) -> pd.DataFrame:
        """Create a summary DataFrame of all analyzed properties"""
        
        summary_data = []
        
        for prop in properties:
            # Use the best scenario for summary
            best = prop['scenarios'][prop['best_scenario']]
            
            summary_data.append({
                'Address': prop['address'],
                'Units': prop['num_units'],
                'Market Price': prop['market_price'],
                'Economic Value': best['economic_value'],
                'Value Ratio': best['value_ratio'],
                'Gross Income': prop['gross_annual_income'],
                'NOI': best['noi'],
                'TGA': best['tga'],
                'Down Payment': best['down_payment'],
                'Monthly Payment': best['monthly_payment'],
                'Monthly Cashflow': best['monthly_cashflow'],
                'Cash ROI': best['cash_roi'],
                'Best Scenario': prop['best_scenario'],
                'MRB': best['mrb'],
                'Per Unit Value': best['value_per_unit']
            })
        
        df = pd.DataFrame(summary_data)
        
        # Sort by value ratio (best deals first)
        df = df.sort_values('Value Ratio', ascending=False)
        
        return df


def main():
    """Demo with sample properties"""
    
    analyzer = PropertyAnalyzer(vacancy_rate=0.03, expense_ratio=0.25)
    
    # Sample properties (from your Liste sheet)
    sample_properties = [
        {
            'address': '1825, Rue Sainte-Hélène, Longueuil',
            'market_price': 1199999,
            'gross_annual_income': 84660,
            'num_units': 6
        },
        {
            'address': '1, Rue des Pins, Chambly',
            'market_price': 839000,
            'gross_annual_income': 57672,
            'num_units': 6
        },
        {
            'address': '69-71A, Rue Saint-Charles, Saint-Jean-sur-Richelieu',
            'market_price': 947900,
            'gross_annual_income': 68220,
            'num_units': 6
        }
    ]
    
    # Analyze all properties
    results = []
    for prop in sample_properties:
        analysis = analyzer.analyze_property(**prop)
        results.append(analysis)
        
        print(f"\n{'='*80}")
        print(f"Property: {prop['address']}")
        print(f"Market Price: ${prop['market_price']:,.0f}")
        print(f"Gross Annual Income: ${prop['gross_annual_income']:,.0f}")
        print(f"Number of Units: {prop['num_units']}")
        print(f"\nBest Scenario: {analysis['best_scenario']}")
        
        best = analysis['scenarios'][analysis['best_scenario']]
        print(f"Economic Value: ${best['economic_value']:,.0f}")
        print(f"Value Ratio: {best['value_ratio']:.2%}")
        print(f"TGA: {best['tga']:.4f}")
        print(f"Down Payment: ${best['down_payment']:,.0f} ({best['down_payment_pct']:.1%})")
        print(f"Monthly Cashflow: ${best['monthly_cashflow']:,.0f}")
        print(f"Cash ROI: {best['cash_roi']:.2%}")
    
    # Create summary report
    print(f"\n{'='*80}")
    print("SUMMARY REPORT - All Properties")
    print(f"{'='*80}\n")
    
    summary = analyzer.create_summary_report(results)
    print(summary.to_string(index=False))
    
    return results, summary


if __name__ == '__main__':
    main()
