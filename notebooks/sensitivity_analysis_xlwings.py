"""
SENSITIVITY ANALYSIS GENERATOR - Using xlwings
This script uses xlwings to interact with Excel and generate sensitivity analysis.
Requires: pip install xlwings
"""

import xlwings as xw
import time

print("="*70)
print("SENSITIVITY ANALYSIS GENERATOR")
print("="*70)

# Configuration
FILE_PATH = '../models/ASML_DCF_Model.xlsx'  # Update if different
WACC_SHEET = 'WACC'
WACC_CELL = 'C36'
DCF_SHEET = 'DCF Calculation'
GROWTH_CELL = 'B20'
VALUE_CELL = 'D40'
SENSITIVITY_SHEET = 'Sensitivity'

# WACC values (8.0% to 13.0%)
wacc_values = [0.080, 0.085, 0.090, 0.095, 0.100, 0.105, 0.110, 0.115, 0.120, 0.125, 0.130]

# Terminal Growth values (1.5% to 5.0%)
growth_values = [0.015, 0.020, 0.025, 0.030, 0.035, 0.040, 0.045, 0.050]

print(f"\n1. Opening Excel file...")
print(f"   File: {FILE_PATH}")

try:
    # Open the workbook
    wb = xw.Book(FILE_PATH)
    print("   ✓ Workbook opened")
except Exception as e:
    print(f"   ✗ Error opening file: {e}")
    print("\nMake sure:")
    print("  - Excel is installed")
    print("  - File path is correct")
    print("  - xlwings is installed: pip install xlwings")
    exit()

# Get sheets
wacc_sheet = wb.sheets[WACC_SHEET]
dcf_sheet = wb.sheets[DCF_SHEET]
sensitivity_sheet = wb.sheets[SENSITIVITY_SHEET]

print(f"\n2. Configuration:")
print(f"   WACC cell: {WACC_SHEET}!{WACC_CELL}")
print(f"   Terminal Growth cell: {DCF_SHEET}!{GROWTH_CELL}")
print(f"   DCF Value cell: {DCF_SHEET}!{VALUE_CELL}")

# Store original values
original_wacc = wacc_sheet.range(WACC_CELL).value
original_growth = dcf_sheet.range(GROWTH_CELL).value

print(f"\n3. Current values:")
print(f"   WACC: {original_wacc:.2%}")
print(f"   Terminal Growth: {original_growth:.2%}")

# Calculate sensitivity table
print(f"\n4. Calculating {len(wacc_values)} × {len(growth_values)} = {len(wacc_values) * len(growth_values)} scenarios...")
print(f"   This will take 1-2 minutes...\n")

results = []
total_scenarios = len(wacc_values) * len(growth_values)
current_scenario = 0

for row_idx, growth in enumerate(growth_values):
    row_results = []
    
    for col_idx, wacc in enumerate(wacc_values):
        current_scenario += 1
        
        # Update progress
        if current_scenario % 10 == 0:
            print(f"   Progress: {current_scenario}/{total_scenarios} ({current_scenario/total_scenarios*100:.0f}%)")
        
        # Set WACC
        wacc_sheet.range(WACC_CELL).value = wacc
        
        # Set Terminal Growth
        dcf_sheet.range(GROWTH_CELL).value = growth
        
        # Force Excel to recalculate
        wb.app.calculate()
        
        # Give Excel a moment to recalculate
        time.sleep(0.05)
        
        # Get the DCF value
        dcf_value = dcf_sheet.range(VALUE_CELL).value
        
        row_results.append(dcf_value)
    
    results.append(row_results)

print(f"\n   ✓ All scenarios calculated")

# Restore original values
print(f"\n5. Restoring original values...")
wacc_sheet.range(WACC_CELL).value = original_wacc
dcf_sheet.range(GROWTH_CELL).value = original_growth
wb.app.calculate()
print(f"   ✓ WACC restored to {original_wacc:.2%}")
print(f"   ✓ Terminal Growth restored to {original_growth:.2%}")

# Write results to Sensitivity sheet
print(f"\n6. Writing results to Sensitivity sheet...")

# Starting cell for data (B6 in your setup)
start_row = 6
start_col = 2  # Column B

for row_idx, row_data in enumerate(results):
    for col_idx, value in enumerate(row_data):
        cell_row = start_row + row_idx
        cell_col = start_col + col_idx
        
        # Write value to cell
        sensitivity_sheet.range(cell_row, cell_col).value = value

print(f"   ✓ Results written to cells B6:L13")

# Format the cells (optional)
print(f"\n7. Formatting table...")

# Select the data range
data_range = sensitivity_sheet.range('B6:L13')

# Number format (no decimals, thousands separator)
data_range.number_format = '#,##0'

print(f"   ✓ Number formatting applied")

# Save the workbook
print(f"\n8. Saving workbook...")
wb.save()
print(f"   ✓ Workbook saved")

# Display results summary
print(f"\n" + "="*70)
print("RESULTS SUMMARY")
print("="*70)

print(f"\nValue range:")
all_values = [val for row in results for val in row]
print(f"  Minimum: €{min(all_values):,.0f}")
print(f"  Maximum: €{max(all_values):,.0f}")
print(f"  Range: €{max(all_values) - min(all_values):,.0f}")

# Find base case (9.5% WACC, 2.5% Terminal Growth)
try:
    base_wacc_idx = wacc_values.index(0.095)
    base_growth_idx = growth_values.index(0.025)
    base_value = results[base_growth_idx][base_wacc_idx]
    print(f"\nBase Case (9.5% WACC, 2.5% Growth): €{base_value:,.0f}")
except:
    print(f"\nBase case not in table")

print(f"\n" + "="*70)
print("COMPLETE!")
print("="*70)
print(f"\nYour sensitivity table has been filled in.")
print(f"Check the Sensitivity sheet in Excel.")
print(f"\nRecommended next steps:")
print(f"  1. Add conditional formatting (color scale)")
print(f"  2. Highlight your base case")
print(f"  3. Add a legend below the table")

# Close Excel (comment out if you want to keep it open)
# wb.close()

print(f"\n✓ Done!")
