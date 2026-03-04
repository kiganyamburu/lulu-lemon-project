# Lululemon Athletica (LULU) Pro Forma Financial Model

## Project Overview

This project builds a **pro forma financial model** for Lululemon Athletica Inc. (NASDAQ: LULU) as part of an Equity Analysis assignment. The model includes fully integrated financial statements that are balanced and logically connected.

## Financial Statements Included

- **Income Statement** - Revenue, COGS, Operating Expenses, Net Income
- **Balance Sheet** - Assets, Liabilities, Shareholders' Equity (balanced: Assets = Liabilities + Equity)
- **Cash Flow Statement** - Operating, Investing, and Financing activities

## Data Sources

| Source | Description |
|--------|-------------|
| Bloomberg Terminal | Bloomberg Macro XIDF data file for historical financials and consensus estimates |
| Lululemon 10-K Filings | SEC EDGAR annual reports for historical data |
| Wall Street Consensus | Yahoo Finance, FactSet, Seeking Alpha for revenue estimates |

## Revenue Estimates (Street Consensus)

| Fiscal Year | Revenue ($M) | Growth Rate | Source |
|-------------|--------------|-------------|--------|
| FY 2023 (Actual) | $8,110.5 | - | 10-K Filing |
| FY 2024 (Actual) | $9,619.3 | 18.6% | 10-K Filing |
| FY 2025 (Actual) | $10,588.1 | 10.1% | 10-K Filing |
| FY 2026E | $11,040.3 | 4.3% | Bloomberg Consensus |
| FY 2027E | $11,537.6 | 4.5% | Bloomberg Consensus |
| FY 2028E | $12,190.5 | 5.7% | Bloomberg Consensus |

## Project Files

| File | Description |
|------|-------------|
| `lululemon_proforma_bloomberg.py` | Main model using Bloomberg Terminal data |
| `lululemon_proforma_model.py` | Alternative model with Yahoo Finance estimates |
| `update_module7_final.py` | Script to update Module7.xlsm with LULU data |
| `update_module7_v2.py` | Version 2 of the update script |
| `update_module7.py` | Initial update script |
| `Module7.xlsm` | Excel template for the financial model |
| `Module7_LULU.xlsm` | Completed model with Lululemon data |
| `Bloomberg Macro XIDF (1).xlsm` | Bloomberg data export file |

## Key Assumptions

| Assumption | Value | Basis |
|------------|-------|-------|
| Gross Margin | ~57-59% | Historical average |
| SG&A as % of Revenue | ~33-35% | Historical trend analysis |
| Effective Tax Rate | ~27-30% | Recent 10-K filings |
| CapEx as % of Revenue | ~5% | Management guidance |
| Days Inventory Outstanding | ~120 days | Historical inventory turns |

## Requirements

- Python 3.x
- openpyxl library (`pip install openpyxl`)

## Usage

```bash
# Generate the pro forma model using Bloomberg data
python lululemon_proforma_bloomberg.py

# Update the Module7 Excel template
python update_module7_final.py
```

## Model Validation

The model ensures:
- **Balance Sheet Equation**: Assets = Liabilities + Shareholders' Equity
- **Cash Flow Integration**: Ending cash ties to balance sheet cash
- **Net Income Flow**: Net income flows through retained earnings
- **Working Capital**: Changes properly reflected in operating cash flow

## Assignment Requirements

1. Fully integrated income statement, balance sheet, and cash flow statement
2. Balanced financial statements (accounting identity maintained)
3. Revenue based on Street consensus estimates
4. Proper source citations for all estimates

## Data Accessed

- **Date**: February 2026
- **Primary Sources**: SEC EDGAR, Bloomberg Terminal, Yahoo Finance

---

*This model was created for educational purposes as part of an Equity Analysis course assignment.*
