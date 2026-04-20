import os
import xlsxwriter

filepath = r"E:\4. Sanu\1. Sanidhya Documents\3. Income Tax\Income_Tax_Calculator_Python"
filename = "Tax_Calculator_FY26_27.xlsx"
full_path = os.path.join(filepath, filename)

if not os.path.exists(filepath):
    os.makedirs(filepath, exist_ok=True)

workbook = xlsxwriter.Workbook(full_path)
worksheet = workbook.add_worksheet('Tax Calculator')

# ── Formats ──────────────────────────────────────────────────────────────────
hdr   = workbook.add_format({'bold':True,  'bg_color':'#D3D3D3', 'border':1})
hint  = workbook.add_format({'italic':True,'font_color':'#808080'})
inp   = workbook.add_format({'bg_color':'#FFFFE0','border':1,'num_format':'#,##0'})
calc  = workbook.add_format({'border':1,  'num_format':'#,##0','bg_color':'#F0F8FF'})
dd    = workbook.add_format({'bg_color':'#FFFFE0','border':1})
res   = workbook.add_format({'bold':True, 'border':1,'num_format':'#,##0','bg_color':'#e2efda'})
alrt  = workbook.add_format({'bold':True, 'border':1,'num_format':'#,##0','bg_color':'#fff2cc','font_color':'#c00000'})
sub   = workbook.add_format({'border':1,  'num_format':'#,##0','bg_color':'#EAF4FB','italic':True}) # sub-item
warn  = workbook.add_format({'bold':True, 'border':1,'num_format':'#,##0','bg_color':'#FFD7D7'})    # over-limit warning
ok    = workbook.add_format({'bold':True, 'border':1,'num_format':'#,##0','bg_color':'#C6EFCE'})    # within limit

worksheet.set_column('A:A', 45)
worksheet.set_column('B:B', 22)
worksheet.set_column('C:C', 22)
worksheet.set_column('D:D', 55)

# ══════════════════════════════════════════════════════════════
# ROW 1  – City dropdown
# ══════════════════════════════════════════════════════════════
worksheet.write('A1','City Type (Metro/Non-Metro)')
worksheet.data_validation('B1',{'validate':'list','source':['Metro','Non-Metro']})
worksheet.write('B1','Metro', dd)
worksheet.write('C1','Metro = 50% Basic, Non-Metro = 40% Basic for HRA', hint)

# ══════════════════════════════════════════════════════════════
# SECTION 1 – INCOME  (rows 3-14)
# ══════════════════════════════════════════════════════════════
worksheet.write('A3','INCOME COMPONENTS', hdr)
worksheet.write('B3','Monthly Amount',    hdr)
worksheet.write('C3','Annual Amount',     hdr)
worksheet.write('D3','Hints / Explanations', hdr)

# R4  Basic
worksheet.write('A4','Basic Salary'); worksheet.write('B4',89050,inp)
worksheet.write_formula('C4','=B4*12',calc)
worksheet.write('D4','Your base fixed pay from CTC', hint)

# R5  HRA
worksheet.write('A5','HRA Received'); worksheet.write('B5',35620,inp)
worksheet.write_formula('C5','=B5*12',calc)
worksheet.write('D5','House Rent Allowance given by employer', hint)

# R6  FBP / Other Allowances
worksheet.write('A6','Other Allowance (FBP / Special)'); worksheet.write('B6',16318,inp)
worksheet.write_formula('C6','=B6*12',calc)
worksheet.write('D6','Flexible Benefit Plan, Medical, LTA, Special allowance etc.', hint)

# R7  Vested Shares / RSU  (annual only)
worksheet.write('A7','Vested Shares / RSU Income (Annual)'); worksheet.write('B7',0,inp)
worksheet.write('C7',0,inp)
worksheet.write('D7','Perquisite value of RSUs/ESOPs that vested this FY. Enter annual figure.', hint)

# R8  Other Income
worksheet.write('A8','Other Income (Interest, Rent recv., etc.)'); worksheet.write('B8',0,inp)
worksheet.write('C8',0,inp)
worksheet.write('D8','FD interest, savings interest, rental income etc. (Annual)', hint)

# R9  Gross Total Income
worksheet.write('A9','Gross Total Income', res)
worksheet.write_formula('C9','=SUM(C4:C8)', res)
worksheet.write('D9','Sum of all income before any deduction', hint)

# R10 blank / separator
# ══════════════════════════════════════════════════════════════
# SECTION 2 – DEDUCTIONS  (rows 11 onwards)
# ══════════════════════════════════════════════════════════════
worksheet.write('A11','DEDUCTIONS & EXEMPTIONS', hdr)
worksheet.write('B11','Old Regime',   hdr)
worksheet.write('C11','New Regime',   hdr)
worksheet.write('D11','Hints / Explanations', hdr)

# ── 2a. Standard Deduction ──
worksheet.write('A12','Standard Deduction (Sec 16(ia))')
worksheet.write('B12',50000, calc); worksheet.write('C12',75000, calc)
worksheet.write('D12','Auto-applied. Flat deduction: Old = ₹50,000 | New = ₹75,000', hint)

# ── 2b. Professional Tax ──
worksheet.write('A13','Professional Tax (Sec 16(iii))')
worksheet.write('B13',2400, inp); worksheet.write('C13',0, calc)
worksheet.write('D13','Deducted by employer (₹200/month). Max ₹2,500/yr. Only in Old Regime.', hint)

# ── 2c. HRA Exemption ──
worksheet.write('A14','Rent Paid per Month (for HRA calc)'); worksheet.write('B14',0,inp)
worksheet.write_formula('C14','=B14*12', calc)
worksheet.write('D14','Enter actual rent you pay. Used only to calculate HRA exemption below.', hint)

worksheet.write('A15','  ↳ HRA Exemption (Sec 10(13A)) — Auto Calculated', sub)
worksheet.write_formula('B15','=MIN(C5, MAX(0, C14-0.1*C4), IF(B1="Metro", 0.5*C4, 0.4*C4))', sub)
worksheet.write('C15',0, calc)
worksheet.write('D15','Least of: (1) Actual HRA, (2) Rent - 10% Basic, (3) 50%/40% Basic', hint)

# ── 2d. Sec 80C – fully broken out ──
worksheet.write('A16','SEC 80C INVESTMENTS (Max ₹1,50,000)', hdr)
worksheet.write('B16','Old Regime', hdr); worksheet.write('C16','New Regime', hdr)

# Sub-items
worksheet.write('A17','  ↳ EPF / PF (Employee Contribution)')
worksheet.write_formula('B17','=B4*0.12*12', sub)   # 12% of Basic annual
worksheet.write('C17',0, calc)
worksheet.write('D17','Auto-calculated: 12% of Basic Salary. Verify against payslip.', hint)

worksheet.write('A18','  ↳ PPF (Public Provident Fund)'); worksheet.write('B18',0,inp); worksheet.write('C18',0,calc)
worksheet.write('D18','Annual contribution to your PPF account. Max 80C limit applies.', hint)

worksheet.write('A19','  ↳ Life Insurance Premium (LIC / Term)'); worksheet.write('B19',0,inp); worksheet.write('C19',0,calc)
worksheet.write('D19','Premium paid for life insurance policies in your name/spouse/children.', hint)

worksheet.write('A20','  ↳ ELSS Mutual Funds'); worksheet.write('B20',0,inp); worksheet.write('C20',0,calc)
worksheet.write('D20','Equity Linked Saving Scheme. 3-yr lock-in. Best returns among 80C options.', hint)

worksheet.write('A21','  ↳ Home Loan Principal Repayment'); worksheet.write('B21',0,inp); worksheet.write('C21',0,calc)
worksheet.write('D21','Principal portion of EMI on home loan (not interest). Stamp duty also qualifies.', hint)

worksheet.write('A22','  ↳ Tuition Fees (Children)'); worksheet.write('B22',0,inp); worksheet.write('C22',0,calc)
worksheet.write('D22','Tuition fees paid for up to 2 children in Indian schools/colleges.', hint)

worksheet.write('A23','  ↳ NSC / SCSS / Tax Saver FD'); worksheet.write('B23',0,inp); worksheet.write('C23',0,calc)
worksheet.write('D23','National Savings Certificate, Senior Citizen Savings Scheme, 5-yr Tax Saver FD.', hint)

# 80C Total & status
worksheet.write('A24','  80C Total (auto-summed)')
worksheet.write_formula('B24','=SUM(B17:B23)', sub); worksheet.write('C24',0, calc)
worksheet.write('D24','Sum of all 80C items entered above.', hint)

worksheet.write('A25','  80C Eligible (capped at ₹1,50,000)')
worksheet.write_formula('B25','=MIN(B24, 150000)', ok)
worksheet.write('C25',0, calc)
worksheet.write('D25','Actual deduction claimed = min of total & ₹1.5L limit.', hint)

worksheet.write('A26','  ⚠ 80C Remaining Limit')
worksheet.write_formula('B26','=MAX(0, 150000-B24)', warn)
worksheet.write('C26',0, calc)
worksheet.write('D26','If > 0, you can still invest more to fully use the ₹1.5L limit.', hint)

# ── 2e. NPS ──
worksheet.write('A27','NPS — Sec 80CCD(1B) (Own contribution, extra 50k)')
worksheet.write('B27',0,inp); worksheet.write('C27',0,calc)
worksheet.write('D27','Additional ₹50,000 over and above 80C limit. Own voluntary NPS Tier-1 contribution.', hint)

worksheet.write('A28','NPS — Sec 80CCD(2) (Employer contribution)')
worksheet.write('B28',0,inp)
worksheet.write_formula('C28','=B28',calc)
worksheet.write('D28','Employer NPS contribution up to 10% of Basic. Allowed in BOTH regimes.', hint)

# ── 2f. Home Loan Interest ──
worksheet.write('A29','Home Loan Interest — Sec 24(b)')
worksheet.write('B29',0,inp); worksheet.write('C29',0,calc)
worksheet.write('D29','Interest portion of home loan EMI. Max ₹2L for self-occupied. Not in new regime.', hint)

# ── 2g. Health Insurance ──
worksheet.write('A30','SEC 80D — HEALTH INSURANCE', hdr)
worksheet.write('B30','Old Regime', hdr); worksheet.write('C30','New Regime', hdr)

worksheet.write('A31','  ↳ Self + Spouse + Children (below 60)')
worksheet.write('B31',0,inp); worksheet.write('C31',0,calc)
worksheet.write('D31','Premium paid for family floater. Max ₹25,000.', hint)

worksheet.write('A32','  ↳ Parents (below 60)')
worksheet.write('B32',0,inp); worksheet.write('C32',0,calc)
worksheet.write('D32','Premium for parents below 60 years. Additional ₹25,000 allowed.', hint)

worksheet.write('A33','  ↳ Parents (Senior Citizen, 60+)')
worksheet.write('B33',0,inp); worksheet.write('C33',0,calc)
worksheet.write('D33','Premium for senior citizen parents. Limit is ₹50,000 (replaces the ₹25k above).', hint)

worksheet.write('A34','  ↳ Preventive Health Check-up')
worksheet.write('B34',0,inp); worksheet.write('C34',0,calc)
worksheet.write('D34','Health check-up costs. Max ₹5,000 (within the overall 80D limit).', hint)

worksheet.write('A35','  80D Total (auto-summed)')
worksheet.write_formula('B35','=SUM(B31:B34)', sub); worksheet.write('C35',0, calc)
worksheet.write('D35','Total health insurance deduction claimed.', hint)

# ── 2h. Other deductions ──
worksheet.write('A36','SEC 80TTA/80TTB — Savings Interest', )
worksheet.write('B36',0,inp); worksheet.write('C36',0,calc)
worksheet.write('D36','Interest on savings account. Max ₹10,000 (₹50,000 for senior citizens u/s 80TTB).', hint)

worksheet.write('A37','SEC 80G — Donations to Charity')
worksheet.write('B37',0,inp); worksheet.write('C37',0,calc)
worksheet.write('D37','50% or 100% deduction depending on recipient charity. Not in new regime.', hint)

# ── Total Deductions ──
worksheet.write('A38','TOTAL DEDUCTIONS', res)
worksheet.write_formula('B38','=B12+B13+B15+B25+B27+B28+B29+B35+B36+B37', res)
worksheet.write_formula('C38','=C12+C28', res)
worksheet.write('D38','Old: all eligible deductions | New: Standard Deduction + Employer NPS only', hint)

# ══════════════════════════════════════════════════════════════
# SECTION 3 – TAXABLE INCOME  (rows 40-41)
# ══════════════════════════════════════════════════════════════
worksheet.write('A40','NET TAXABLE INCOME', hdr)
worksheet.write('B40','Old Regime', hdr); worksheet.write('C40','New Regime', hdr)
worksheet.write_formula('B41','=MAX(0, C9-B38)', res)
worksheet.write_formula('C41','=MAX(0, C9-C38)', res)

# ══════════════════════════════════════════════════════════════
# SECTION 4 – TAX CALCULATION  (rows 43-51)
# ══════════════════════════════════════════════════════════════
worksheet.write('A43','FINAL TAX CALCULATION', hdr)
worksheet.write('B43','Old Regime', hdr); worksheet.write('C43','New Regime', hdr)

old_tax = ('=IF(B41<=250000,0,'
           'IF(B41<=500000,(B41-250000)*0.05,'
           'IF(B41<=1000000,12500+(B41-500000)*0.2,'
           '112500+(B41-1000000)*0.3)))')

new_tax = ('=IF(C41<=400000,0,'
           'IF(C41<=800000,(C41-400000)*0.05,'
           'IF(C41<=1200000,20000+(C41-800000)*0.10,'
           'IF(C41<=1600000,60000+(C41-1200000)*0.15,'
           'IF(C41<=2000000,120000+(C41-1600000)*0.20,'
           'IF(C41<=2400000,200000+(C41-2000000)*0.25,'
           '300000+(C41-2400000)*0.30))))))')

worksheet.write('A44','Tax on Income (as per slabs)')
worksheet.write_formula('B44', old_tax, calc); worksheet.write_formula('C44', new_tax, calc)
worksheet.write('D44','Old: 0/5/20/30% slabs | New: 0/5/10/15/20/25/30% slabs', hint)

worksheet.write('A45','Rebate u/s 87A')
worksheet.write_formula('B45','=IF(B41<=500000, MIN(B44,12500), 0)', calc)
worksheet.write_formula('C45','=IF(C41<=1200000, MIN(C44,60000), 0)', calc)
worksheet.write('D45','Old: Zero tax if income ≤ ₹5L | New: Zero tax if income ≤ ₹12L', hint)

worksheet.write('A46','Tax after Rebate')
worksheet.write_formula('B46','=MAX(0, B44-B45)', calc); worksheet.write_formula('C46','=MAX(0, C44-C45)', calc)

worksheet.write('A47','Health & Education Cess (4%)')
worksheet.write_formula('B47','=B46*0.04', calc); worksheet.write_formula('C47','=C46*0.04', calc)
worksheet.write('D47','4% on tax after rebate. Applies in both regimes.', hint)

worksheet.write('A48','TOTAL TAX PAYABLE', res)
worksheet.write_formula('B48','=B46+B47', res); worksheet.write_formula('C48','=C46+C47', res)

worksheet.write('A49','TDS Already Deducted (from payslip)')
worksheet.write('B49',0, inp); worksheet.write('C49',0, inp)
worksheet.write('D49','Total TDS deducted by employer this FY. Check Form 16 / payslips.', hint)

worksheet.write('A50','Balance Tax Payable / (Refund Due)')
worksheet.write_formula('B50','=B48-B49', alrt); worksheet.write_formula('C50','=C48-C49', alrt)
worksheet.write('D50','Negative = refund due. Positive = you owe this much more.', hint)

worksheet.write('A52','DIFFERENCE: Old Regime Tax – New Regime Tax', hdr)
worksheet.write_formula('B52','=B48-C48', alrt)
worksheet.write('D52','Positive = New Regime saves you money. Negative = Old Regime is better.', hint)

workbook.close()

print("Excel generated at", full_path)