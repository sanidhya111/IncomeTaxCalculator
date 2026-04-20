import pandas as pd
import numpy as np
import io
import os

# We will create an Excel file that acts as a calculator comparing Old vs New tax regimes.
# The user wants to include things like home loan, investments, vested shares, and hints.
# The Excel will have input fields, and calculated fields.

# Since we need a formula-based Excel, we can write an excel file using pandas or xlsxwriter.
import xlsxwriter

# Your defined variables (using 'r' before the string to handle Windows backslashes safely)
filepath = r"E:\4. Sanu\1. Sanidhya Documents\3. Income Tax\Income_Tax_Calculator_Python"
filename = "Tax_Calculator_FY26_27.xlsx"

# Dynamically construct the full file path
full_path = os.path.join(filepath, filename)

# Create the directory and any missing parent folders if they don't exist
if not os.path.exists(filepath):
    os.makedirs(filepath, exist_ok=True)

# Create the Excel file at the newly verified/created path
workbook = xlsxwriter.Workbook(full_path)

worksheet = workbook.add_worksheet('Tax Calculator')

# Formatting
header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
hint_format = workbook.add_format({'italic': True, 'font_color': '#808080'})
input_format = workbook.add_format({'bg_color': '#FFFFE0', 'border': 1, 'num_format': '#,##0'})
calc_format = workbook.add_format({'border': 1, 'num_format': '#,##0', 'bg_color': '#F0F8FF'})
currency_format = workbook.add_format({'num_format': '₹#,##0'})

worksheet.set_column('A:A', 35)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 50)

# Headers
worksheet.write('A1', 'Income & Deductions', header_format)
worksheet.write('B1', 'Amount (Inputs)', header_format)
worksheet.write('C1', 'New Tax Regime', header_format)
worksheet.write('D1', 'Hints / Explanations', header_format)

row = 1
# Gross Income
worksheet.write(row, 0, 'Basic Salary')
worksheet.write(row, 1, 1500000, input_format)
worksheet.write(row, 3, 'Your base fixed pay', hint_format)
row += 1

worksheet.write(row, 0, 'HRA Received')
worksheet.write(row, 1, 300000, input_format)
worksheet.write(row, 3, 'House Rent Allowance given by employer', hint_format)
row += 1

worksheet.write(row, 0, 'Special/Other Allowances')
worksheet.write(row, 1, 200000, input_format)
worksheet.write(row, 3, 'LTA, Medical, Special allowance etc.', hint_format)
row += 1

worksheet.write(row, 0, 'Vested Shares / RSU Income')
worksheet.write(row, 1, 100000, input_format)
worksheet.write(row, 3, 'Perquisite value of shares vested in the FY', hint_format)
row += 1

worksheet.write(row, 0, 'Other Income (Interest, etc.)')
worksheet.write(row, 1, 50000, input_format)
worksheet.write(row, 3, 'Interest from savings, FDs, etc.', hint_format)
row += 1

worksheet.write(row, 0, 'Gross Total Income')
worksheet.write_formula(row, 1, '=SUM(B2:B6)', calc_format)
worksheet.write_formula(row, 2, '=B7', calc_format)
worksheet.write(row, 3, 'Total income before deductions', hint_format)
row += 2

# Deductions
worksheet.write(row, 0, 'Deductions (Old Regime)', header_format)
worksheet.write(row, 1, 'Old Regime Amount', header_format)
worksheet.write(row, 2, 'New Regime Allowed', header_format)
row += 1

worksheet.write(row, 0, 'Standard Deduction')
worksheet.write(row, 1, 50000, calc_format)
worksheet.write(row, 2, 75000, calc_format)
worksheet.write(row, 3, 'Fixed flat deduction for salaried (Old: 50k, New: 75k)', hint_format)
row += 1

worksheet.write(row, 0, 'HRA Exemption (Sec 10(13A))')
worksheet.write(row, 1, 150000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Least of: Actual HRA, Rent - 10% Basic, 50%/40% of Basic', hint_format)
row += 1

worksheet.write(row, 0, 'Sec 80C (EPF, LIC, PPF, ELSS, Home Principal)')
worksheet.write(row, 1, 150000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Max 1.5L. Includes EPF, PPF, Life Insurance, Home loan principal.', hint_format)
row += 1

worksheet.write(row, 0, 'Sec 80CCD(1B) (NPS Tier 1)')
worksheet.write(row, 1, 50000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Additional 50k for NPS voluntary contribution', hint_format)
row += 1

worksheet.write(row, 0, 'Sec 80CCD(2) (Employer NPS)')
worksheet.write(row, 1, 0, input_format)
worksheet.write_formula(row, 2, '=B13', calc_format)
worksheet.write(row, 3, 'Up to 10% of basic. Allowed in BOTH regimes', hint_format)
row += 1

worksheet.write(row, 0, 'Sec 24(b) Home Loan Interest')
worksheet.write(row, 1, 200000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Up to 2L for self-occupied. Not allowed in new regime for self-occupied.', hint_format)
row += 1

worksheet.write(row, 0, 'Sec 80D (Health Insurance)')
worksheet.write(row, 1, 25000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Medical insurance premium (25k self, 50k senior parents)', hint_format)
row += 1

worksheet.write(row, 0, 'Other Deductions (80TTA, 80G, etc.)')
worksheet.write(row, 1, 10000, input_format)
worksheet.write(row, 2, 0, calc_format)
worksheet.write(row, 3, 'Donations, savings interest up to 10k, etc.', hint_format)
row += 1

worksheet.write(row, 0, 'Total Deductions')
worksheet.write_formula(row, 1, '=SUM(B10:B16)', calc_format)
worksheet.write_formula(row, 2, '=C10+C13', calc_format)
row += 2

# Taxable Income
worksheet.write(row, 0, 'Net Taxable Income', header_format)
worksheet.write_formula(row, 1, '=MAX(0, B7-B17)', header_format)
worksheet.write_formula(row, 2, '=MAX(0, C7-C17)', header_format)
row += 2

# Tax Calculation breakdown Old Regime (AY 2026-27 / FY 25-26 slabs apply)
# Using generic formulas via excel IFs for old regime slabs
# 0-2.5L: 0, 2.5L-5L: 5%, 5L-10L: 20%, >10L: 30%
old_tax_formula = """=IF(B19<=250000,0,
IF(B19<=500000,(B19-250000)*0.05,
IF(B19<=1000000,12500+(B19-500000)*0.2,
112500+(B19-1000000)*0.3)))"""
# 87A rebate for old: income <= 500000, rebate up to 12500
old_rebate = "=IF(B19<=500000, MIN(B22, 12500), 0)"

# New Tax Regime FY 26-27 (AY 27-28 slabs - same as FY 25-26 updated budget)
# 0-4L: 0, 4-8L: 5%, 8-12L: 10%, 12-16L: 15%, 16-20L: 20%, 20-24L: 25%, >24L: 30%
new_tax_formula = """=IF(C19<=400000,0,
IF(C19<=800000,(C19-400000)*0.05,
IF(C19<=1200000,20000+(C19-800000)*0.10,
IF(C19<=1600000,60000+(C19-1200000)*0.15,
IF(C19<=2000000,120000+(C19-1600000)*0.20,
IF(C19<=2400000,200000+(C19-2000000)*0.25,
300000+(C19-2400000)*0.30))))))"""
# 87A rebate for new: income <= 1200000, rebate up to 60000 (updated budget 2026 continues 12L)
new_rebate = "=IF(C19<=1200000, MIN(C22, 60000), 0)"

worksheet.write(row, 0, 'Tax Calculations', header_format)
worksheet.write(row, 1, 'Old Regime', header_format)
worksheet.write(row, 2, 'New Regime', header_format)
row += 1

worksheet.write(row, 0, 'Tax on Income')
worksheet.write_formula(row, 1, old_tax_formula.replace('\n', ''), calc_format)
worksheet.write_formula(row, 2, new_tax_formula.replace('\n', ''), calc_format)
row += 1

worksheet.write(row, 0, 'Rebate u/s 87A')
worksheet.write_formula(row, 1, old_rebate, calc_format)
worksheet.write_formula(row, 2, new_rebate, calc_format)
worksheet.write(row, 3, 'Old: Up to 5L income. New: Up to 12L income', hint_format)
row += 1

worksheet.write(row, 0, 'Tax after Rebate')
worksheet.write_formula(row, 1, '=MAX(0, B22-B23)', calc_format)
worksheet.write_formula(row, 2, '=MAX(0, C22-C23)', calc_format)
row += 1

worksheet.write(row, 0, 'Health & Education Cess (4%)')
worksheet.write_formula(row, 1, '=B24*0.04', calc_format)
worksheet.write_formula(row, 2, '=C24*0.04', calc_format)
row += 1

worksheet.write(row, 0, 'Total Tax Payable')
worksheet.write_formula(row, 1, '=B24+B25', header_format)
worksheet.write_formula(row, 2, '=C24+C25', header_format)
row += 2

worksheet.write(row, 0, 'Difference (Old - New)')
worksheet.write_formula(row, 1, '=B26-C26', header_format)
worksheet.write(row, 3, 'If positive, New Regime is better. If negative, Old Regime is better.', hint_format)

workbook.close()
print("Excel generated at", filename)