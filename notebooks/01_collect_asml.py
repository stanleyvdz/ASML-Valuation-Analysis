import yfinance as yf
import pandas as pd
import numpy as np

print("="*60)
print("ASML FINANCIAL DATA COLLECTION - CLEANED VERSION")
print("="*60)

# Fetch ASML data
asml = yf.Ticker("ASML")

# Get financial statements
income_raw = asml.financials
balance_raw = asml.balance_sheet
cashflow_raw = asml.cashflow

print("\n1. Data fetched from Yahoo Finance")
print(f"   Date range: {income_raw.columns.min()} to {income_raw.columns.max()}")

# ============================================
# INCOME STATEMENT - Extract only needed rows
# ============================================

income_rows_needed = {
    'Total Revenue': 'Revenue',
    'Cost Of Revenue': 'Cost of Sales',
    'Gross Profit': 'Gross Profit',
    'Research And Development': 'R&D Expense',
    'Selling General And Administration': 'SG&A Expense',
    'Operating Income': 'EBIT',
    'Tax Provision': 'Income Tax',
    'Net Income': 'Net Income'
}

income_clean = pd.DataFrame()

for api_name, our_name in income_rows_needed.items():
    if api_name in income_raw.index:
        income_clean.loc[our_name, :] = income_raw.loc[api_name, :]
    else:
        print(f"   Warning: '{api_name}' not found in data")
        income_clean.loc[our_name, :] = np.nan


# ============================================
# BALANCE SHEET - Extract only needed rows
# ============================================

balance_rows_needed = {
    'Total Assets': 'Total Assets',
    'Cash Cash Equivalents And Short Term Investments': 'Cash & Equivalents',
    'Net PPE': 'Property, Plant & Equipment',
    'Total Liabilities Net Minority Interest': 'Total Liabilities',
    'Stockholders Equity': 'Total Equity',
    'Ordinary Shares Number': 'Shares Outstanding'
}

balance_clean = pd.DataFrame()
for api_name, our_name in balance_rows_needed.items():
    if api_name in balance_raw.index:
        balance_clean.loc[our_name, :] = balance_raw.loc[api_name, :]
    else:
        print(f"   Warning: '{api_name}' not found in balance sheet")
        balance_clean.loc[our_name, :] = np.nan

# ============================================
# CASH FLOW - Extract only needed rows
# ============================================

cashflow_rows_needed = {
    'Operating Cash Flow': 'Operating Cash Flow',
    'Capital Expenditure': 'CapEx'
}

cashflow_clean = pd.DataFrame()
for api_name, our_name in cashflow_rows_needed.items():
    if api_name in cashflow_raw.index:
        cashflow_clean.loc[our_name, :] = cashflow_raw.loc[api_name, :]
    else:
        print(f"   Warning: '{api_name}' not found in cash flow")
        cashflow_clean.loc[our_name, :] = np.nan

# Calculate Free Cash Flow
if 'Operating Cash Flow' in cashflow_clean.index and 'CapEx' in cashflow_clean.index:
    # CapEx is usually negative, so we add it (adding a negative = subtracting)
    cashflow_clean.loc['Free Cash Flow', :] = (
        cashflow_clean.loc['Operating Cash Flow', :] + 
        cashflow_clean.loc['CapEx', :]
    )

# ============================================
# CLEAN UP DATES AND CONVERT TO MILLIONS
# ============================================

# Convert column names to years
income_clean.columns = income_clean.columns.strftime('%Y')
balance_clean.columns = balance_clean.columns.strftime('%Y')
cashflow_clean.columns = cashflow_clean.columns.strftime('%Y')

# Convert to millions (data comes in actual values)
income_clean = income_clean / 1_000_000
balance_clean = balance_clean / 1_000_000
cashflow_clean = cashflow_clean / 1_000_000

# Convert shares to millions (they come in actual count)
if 'Shares Outstanding' in balance_clean.index:
    balance_clean.loc['Shares Outstanding', :] = balance_clean.loc['Shares Outstanding', :] / 1_000_000

# Round to whole millions
income_clean = income_clean.round(0)
balance_clean = balance_clean.round(0)
cashflow_clean = cashflow_clean.round(0)

# ============================================
# SORT COLUMNS CHRONOLOGICALLY (oldest to newest)
# ============================================

income_clean = income_clean.sort_index(axis=1)
balance_clean = balance_clean.sort_index(axis=1)
cashflow_clean = cashflow_clean.sort_index(axis=1)

# ============================================
# SAVE CLEAN DATA
# ============================================

print("\n2. Data cleaned and structured")
print(f"   Years available: {', '.join(income_clean.columns)}")
print(f"   Income statement rows: {len(income_clean)}")
print(f"   Balance sheet rows: {len(balance_clean)}")
print(f"   Cash flow rows: {len(cashflow_clean)}")

# Save to CSV
income_clean.to_csv('../data/asml_income_CLEAN.csv')
balance_clean.to_csv('../data/asml_balance_CLEAN.csv')
cashflow_clean.to_csv('../data/asml_cashflow_CLEAN.csv')

print("\n3. Files saved:")
print("   ✓ asml_income_CLEAN.csv")
print("   ✓ asml_balance_CLEAN.csv")
print("   ✓ asml_cashflow_CLEAN.csv")

# ============================================
# DISPLAY SUMMARIES
# ============================================

print("\n" + "="*60)
print("INCOME STATEMENT SUMMARY (EUR Millions)")
print("="*60)
print(income_clean.to_string())

print("\n" + "="*60)
print("BALANCE SHEET SUMMARY (EUR Millions)")
print("="*60)
print(balance_clean.to_string())

print("\n" + "="*60)
print("CASH FLOW SUMMARY (EUR Millions)")
print("="*60)
print(cashflow_clean.to_string())

# ============================================
# CALCULATE KEY METRICS
# ============================================

print("\n" + "="*60)
print("KEY FINANCIAL METRICS")
print("="*60)

metrics = pd.DataFrame()

# Margins
if 'Revenue' in income_clean.index and 'Gross Profit' in income_clean.index:
    metrics.loc['Gross Margin %', :] = (
        income_clean.loc['Gross Profit', :] / 
        income_clean.loc['Revenue', :] * 100
    ).round(1)

if 'Revenue' in income_clean.index and 'EBIT' in income_clean.index:
    metrics.loc['EBIT Margin %', :] = (
        income_clean.loc['EBIT', :] / 
        income_clean.loc['Revenue', :] * 100
    ).round(1)

if 'Revenue' in income_clean.index and 'Net Income' in income_clean.index:
    metrics.loc['Net Margin %', :] = (
        income_clean.loc['Net Income', :] / 
        income_clean.loc['Revenue', :] * 100
    ).round(1)

# Growth rates (year-over-year)
if 'Revenue' in income_clean.index:
    revenue_growth = income_clean.loc['Revenue', :].pct_change() * 100
    metrics.loc['Revenue Growth %', :] = revenue_growth.round(1)

# CapEx as % of Revenue
if 'CapEx' in cashflow_clean.index and 'Revenue' in income_clean.index:
    # CapEx is negative, so we use absolute value
    metrics.loc['CapEx % of Revenue', :] = (
        abs(cashflow_clean.loc['CapEx', :]) / 
        income_clean.loc['Revenue', :] * 100
    ).round(1)

# FCF Margin
if 'Free Cash Flow' in cashflow_clean.index and 'Revenue' in income_clean.index:
    metrics.loc['FCF Margin %', :] = (
        cashflow_clean.loc['Free Cash Flow', :] / 
        income_clean.loc['Revenue', :] * 100
    ).round(1)

print(metrics.to_string())

# Save metrics
metrics.to_csv('../data/asml_metrics.csv')
print("\n   ✓ asml_metrics.csv saved")

# ============================================
# DATA QUALITY CHECK
# ============================================

print("\n" + "="*60)
print("DATA QUALITY CHECK")
print("="*60)

# Check for missing data
missing_income = income_clean.isnull().sum().sum()
missing_balance = balance_clean.isnull().sum().sum()
missing_cashflow = cashflow_clean.isnull().sum().sum()

print(f"Missing values in Income Statement: {missing_income}")
print(f"Missing values in Balance Sheet: {missing_balance}")
print(f"Missing values in Cash Flow: {missing_cashflow}")

if missing_income + missing_balance + missing_cashflow == 0:
    print("\n✓ All data complete - no missing values!")
else:
    print("\n⚠ Some data missing - you may need to fill in manually")
    print("\nMissing data locations:")
    if missing_income > 0:
        print("\nIncome Statement:")
        print(income_clean[income_clean.isnull().any(axis=1)])
    if missing_balance > 0:
        print("\nBalance Sheet:")
        print(balance_clean[balance_clean.isnull().any(axis=1)])
    if missing_cashflow > 0:
        print("\nCash Flow:")
        print(cashflow_clean[cashflow_clean.isnull().any(axis=1)])

print("\n" + "="*60)
print("DATA COLLECTION COMPLETE!")
print("="*60)
print("\nNext steps:")
print("1. Review the output above")
print("2. Check CSV files in data/ folder")
print("3. Verify numbers against ASML annual reports if needed")
print("4. Import clean CSV files into Excel model")
print("="*60)