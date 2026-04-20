import os
import xlsxwriter

filepath = r"E:\4. Sanu\1. Sanidhya Documents\3. Income Tax\Income_Tax_Calculator_Python"
filename = "Tax_Calculator_FY26_27.xlsx"
full_path = os.path.join(filepath, filename)

if not os.path.exists(filepath):
    os.makedirs(filepath, exist_ok=True)

workbook = xlsxwriter.Workbook(full_path)
worksheet = workbook.add_worksheet('Tax Calculator')
ws = worksheet
# ── Formats ──────────────────────────────────────────────────────────────────
hdr    = workbook.add_format({'bold':True,  'bg_color':'#2E4057', 'font_color':'#FFFFFF', 'border':1, 'font_size':10})
sec    = workbook.add_format({'bold':True,  'bg_color':'#4A6FA5', 'font_color':'#FFFFFF', 'border':1, 'font_size':10})
hint   = workbook.add_format({'italic':True,'font_color':'#555555', 'font_size':9, 'text_wrap':True})
inp    = workbook.add_format({'bg_color':'#FFFFE0','border':1,'num_format':'#,##0', 'font_size':10})
calc   = workbook.add_format({'border':1,   'num_format':'#,##0','bg_color':'#EBF5FB', 'font_size':10})
res    = workbook.add_format({'bold':True,  'border':2,'num_format':'#,##0','bg_color':'#D5F5E3', 'font_size':10})
alrt   = workbook.add_format({'bold':True,  'border':2,'num_format':'#,##0','bg_color':'#FEF9E7','font_color':'#C0392B', 'font_size':11})
sub_h  = workbook.add_format({'border':1,   'num_format':'#,##0','bg_color':'#D6EAF8','bold':True,'font_size':9})
warn   = workbook.add_format({'bold':True,  'border':1,'num_format':'#,##0','bg_color':'#FDEDEC','font_color':'#C0392B','font_size':9})
ok_fmt = workbook.add_format({'bold':True,  'border':1,'num_format':'#,##0','bg_color':'#EAFAF1','font_color':'#1E8449','font_size':9})
dd     = workbook.add_format({'bg_color':'#FFFFE0','border':1,'font_size':10})
na     = workbook.add_format({'border':1,'bg_color':'#F2F3F4','font_color':'#AAA','font_size':9,'italic':True})
lta_warn = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#FDEDEC','font_color':'#C0392B','font_size':9})
lta_ok   = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#EAFAF1','font_color':'#1E8449','font_size':9})
lta_calc = workbook.add_format({'border':1,'num_format':'#,##0','bg_color':'#FEF5E7','font_size':9})
sec_sub  = workbook.add_format({'bold':True,'bg_color':'#D4E6F1','border':1,'font_size':9})

ws.set_column('A:A', 48)
ws.set_column('B:B', 22)
ws.set_column('C:C', 22)
ws.set_column('D:D', 22)
ws.set_column('E:E', 70)
ws.set_row(0, 22)

# ══════════════════════════════════════════════════════════════════════════════
# ROW 1 – Title
ws.merge_range('A1:E1',
    'India Income Tax Calculator FY 2026-27 (AY 2027-28) — Salaried Employee',
    workbook.add_format({'bold':True,'bg_color':'#2E4057','font_color':'#FFD700',
                         'font_size':13,'align':'center','valign':'vcenter','border':1}))

# ROW 2 – City dropdown
ws.write('A2', 'City Type (for HRA)')
ws.data_validation('B2', {'validate':'list','source':['Metro (Mumbai/Delhi/Chennai/Kolkata/Bengaluru/Hyderabad)','Non-Metro']})
ws.write('B2', 'Metro (Mumbai/Delhi/Chennai/Kolkata/Bengaluru/Hyderabad)', dd)
ws.write('E2', 'Metro = 50% of Basic for HRA limit; Non-Metro = 40%. (Budget 2005, unchanged)', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION A – INCOME
ws.merge_range('A4:E4', 'SECTION A: GROSS INCOME', sec)
ws.write('A5','Income Component', hdr)
ws.write('B5','Monthly (₹)',      hdr)
ws.write('C5','Annual (₹)',       hdr)
ws.write('D5','Auto Annual (₹)',  hdr)
ws.write('E5','Hints & Limits',   hdr)

# Helper note row
ws.merge_range('A6:E6',
    '⬛ Yellow = Input cells.  Enter EITHER Monthly OR Annual — the other column auto-calculates.  Blue = Auto-calculated (do not edit).',
    workbook.add_format({'italic':True,'bg_color':'#FEF9E7','border':1,'font_size':9,'text_wrap':True}))

# Row 7 – Basic
ws.write('A7','Basic Salary')
ws.write('B7', 89050, inp)
ws.write('C7', 0,     inp)
ws.write_formula('D7','=IF(C7>0,C7,B7*12)', calc)
ws.write('E7','Enter monthly OR annual basic. Auto-annual = monthly×12 if monthly entered, else uses annual directly.', hint)

# Row 8 – HRA
ws.write('A8','HRA Received (from employer)')
ws.write('B8', 35620, inp)
ws.write('C8', 0,     inp)
ws.write_formula('D8','=IF(C8>0,C8,B8*12)', calc)
ws.write('E8','HRA in your CTC. Enter monthly or annual. Exemption calculated in Section B.', hint)

# Row 9 – FBP
ws.write('A9','Special / Other Allowance (FBP)')
ws.write('B9', 16318, inp)
ws.write('C9', 0,     inp)
ws.write_formula('D9','=IF(C9>0,C9,B9*12)', calc)
ws.write('E9','Flexible Benefit Plan, Special Allowance, Meal, Phone etc. Fully taxable unless reimbursed with bills.', hint)

# Row 10 – LTA  ← BOTH MONTHLY & ANNUAL SUPPORTED
ws.write('A10','LTA Received (Leave Travel Allowance)')
ws.write('B10', 0, inp)   # Monthly input
ws.write('C10', 0, inp)   # Annual input
ws.write_formula('D10','=IF(C10>0,C10,B10*12)', calc)
ws.write('E10','Enter monthly LTA component OR annual LTA — whichever you know. Auto-Annual picks up the right one. LTA exemption & unclaimed tax calculated in Section B.', hint)

# Row 11 – Variable Pay
ws.write('A11','Variable Pay / Performance Bonus')
ws.write('B11', 0, inp)
ws.write('C11', 0, inp)
ws.write_formula('D11','=IF(C11>0,C11,B11*12)', calc)
ws.write('E11','Annual bonus, PMS payout, incentives. FULLY TAXABLE both regimes. Employer deducts TDS at slab rate when paid. Enter monthly if paid monthly, or annual lump sum.', hint)

# Row 12 – RSU / ESOP
ws.write('A12','Vested RSU / ESOP Income')
ws.write('B12', 0, inp)
ws.write('C12', 0, inp)
ws.write_formula('D12','=IF(C12>0,C12,B12*12)', calc)
ws.write('E12','FMV on vest date − exercise price = perquisite taxable as salary. Employer deducts TDS. Enter annual FMV × shares vested. (Sec 17(2), both regimes).', hint)

# Row 13 – Other Income
ws.write('A13','Other Income (FD Interest, Rent received, etc.)')
ws.write('B13', 0, inp)
ws.write('C13', 0, inp)
ws.write_formula('D13','=IF(C13>0,C13,B13*12)', calc)
ws.write('E13','FD/RD interest, savings interest, rental income. Enter monthly or annual.', hint)

# Row 14 – Gross Total
ws.write('A14','GROSS TOTAL INCOME', res)
ws.write_formula('D14','=SUM(D7:D13)', res)
ws.write('E14','Sum of all Auto-Annual income figures (column D).', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION B – EXEMPTIONS
ws.merge_range('A16:E16','SECTION B: EXEMPTIONS (Reduce Gross Income — OLD REGIME ONLY unless noted)', sec)
ws.write('A17','Exemption Component', hdr)
ws.write('B17','Old Regime (₹)',      hdr)
ws.write('C17','New Regime (₹)',      hdr)
ws.write('D17','',                    hdr)
ws.write('E17','Hints & Limits',      hdr)

# Row 18 – Standard Deduction
ws.write('A18','Standard Deduction (Sec 16(ia))')
ws.write('B18', 50000, calc); ws.write('C18', 75000, calc)
ws.write('E18','OLD: ₹50,000 (Budget 2019). NEW: ₹75,000 (Budget 2024, effective FY 2024-25 onwards).', hint)

# Row 19 – Professional Tax
ws.write('A19','Professional Tax (Sec 16(iii))')
ws.write('B19', 2400, inp); ws.write('C19', 0, na)
ws.write('E19','Deducted by employer. Max ₹2,500/year. Old Regime only. NOT in New Regime. (Unchanged since inception).', hint)

# Row 20 – Rent Paid (HRA helper)
ws.write('A20','  ↳ Rent Paid per Month (for HRA calc)')
ws.write('B20', 0, inp)
ws.write_formula('C20','=B20*12', calc)
ws.write('E20','Enter monthly rent paid to landlord. PAN of landlord required if annual rent > ₹1 lakh.', hint)

# Row 21 – HRA Exemption auto-calc
ws.write('A21','  ↳ HRA Exemption — AUTO CALC (Sec 10(13A))')
hra_f = '=IF(C20=0,0,MIN(D8,MAX(0,C20-0.1*D7),IF(ISNUMBER(SEARCH("Metro",B2)),0.5*D7,0.4*D7)))'
ws.write_formula('B21', hra_f, sub_h); ws.write('C21', 0, na)
ws.write('E21','AUTO-CALC. Least of: (1) Actual HRA received, (2) Rent paid − 10% Basic, (3) 50% Basic (Metro)/40% (Non-Metro). NOT in New Regime. (Sec 10(13A), unchanged).', hint)

# ── LTA BLOCK ────────────────────────────────────────────────────────────────
ws.write('A22','  ↳ LTA Received this FY (from Section A — AUTO)')
ws.write_formula('B22','=D10', sub_h); ws.write('C22', 0, na)
ws.write('E22','Auto-pulled from Section A Row 10 (Annual LTA received from employer this FY).', hint)

ws.write('A23','  ↳ Actual Travel Bills Submitted (Claimed LTA)')
ws.write('B23', 0, inp); ws.write('C23', 0, na)
ws.write('E23','Enter total travel bills submitted to employer this FY. Exemption = LOWER of (LTA received) or (actual bills). Max 2 trips in block 2022-2025. Domestic only. Air: Economy class. Train: AC-1. (Sec 10(5), Budget 2023 block).', hint)

ws.write('A24','  ↳ LTA Exemption — AUTO CALC (Sec 10(5))')
ws.write_formula('B24','=MIN(B22,B23)', lta_ok); ws.write('C24', 0, na)
ws.write('E24','AUTO-CALC. Exemption = lower of (LTA received) vs (bills submitted). If no bills submitted, exemption = ₹0.', hint)

ws.write('A25','  ↳ ⚠ Unclaimed LTA (Taxable)')
ws.write_formula('B25','=MAX(0,B22-B24)', lta_warn); ws.write('C25', 0, na)
ws.write('E25','AUTO-CALC. LTA Received − LTA Exemption = taxable amount. Entire LTA is taxable if you did not travel or did not submit bills to employer.', hint)

tax_on_lta = ('=IF(B25=0,0,'
              'IF(D14>1000000,B25*0.30,'
              'IF(D14>500000,B25*0.20,'
              'IF(D14>250000,B25*0.05,0))))')
ws.write('A26','  ↳ 📊 Estimated Tax on Unclaimed LTA (at marginal slab)')
ws.write_formula('B26', tax_on_lta, lta_calc); ws.write('C26', 0, na)
ws.write('E26','INDICATIVE. Tax on unclaimed LTA based on gross income slab. Actual depends on final taxable income. Submit travel bills to employer to avoid this tax!', hint)

ws.write('A27','TOTAL EXEMPTIONS', res)
ws.write_formula('B27','=B18+B19+B21+B24', res)
ws.write_formula('C27','=C18', res)
ws.write('E27','Old: Standard Ded + Prof Tax + HRA + LTA. New: Standard Deduction (₹75,000) only.', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION C – DEDUCTIONS (Chapter VI-A)
ws.merge_range('A29:E29','SECTION C: DEDUCTIONS UNDER CHAPTER VI-A (Old Regime Only unless noted)', sec)
ws.write('A30','Deduction Component', hdr); ws.write('B30','Old Regime (₹)', hdr)
ws.write('C30','New Regime (₹)',      hdr); ws.write('E30','Hints & Limits', hdr)

# 80C block header — using write not merge_range for single cell
ws.write('A31','SEC 80C / 80CCC / 80CCD(1) — Combined ceiling ₹1,50,000 | OLD REGIME ONLY (Budget 2014)', sec_sub)
ws.write('B31','', sec_sub); ws.write('C31','', sec_sub); ws.write('D31','', sec_sub); ws.write('E31','', sec_sub)

ws.write('A32','  ↳ EPF / PF — Employee Contribution (AUTO)')
ws.write_formula('B32','=ROUND(D7*0.12,0)', sub_h); ws.write('C32', 0, na)
ws.write('E32','AUTO: 12% of Annual Basic (from D7). Verify vs payslip. (EPF Act 1952, within 80C ceiling ₹1.5L).', hint)

ws.write('A33','  ↳ PPF (Public Provident Fund)'); ws.write('B33',0,inp); ws.write('C33',0,na)
ws.write('E33','Min ₹500, Max ₹1,50,000/yr. 15-yr lock-in. Rate: 7.1% p.a. (Q1 FY26-27). EEE status.', hint)

ws.write('A34','  ↳ Life Insurance Premium (LIC / Term)'); ws.write('B34',0,inp); ws.write('C34',0,na)
ws.write('E34','Premium ≤10% of sum assured (policies post Apr 2012). Term plans fully qualify.', hint)

ws.write('A35','  ↳ ELSS Mutual Funds'); ws.write('B35',0,inp); ws.write('C35',0,na)
ws.write('E35','3-year lock-in. LTCG >₹1.25L taxed at 12.5% (Budget 2024). Best equity tax-saving option.', hint)

ws.write('A36','  ↳ Home Loan — Principal Repayment'); ws.write('B36',0,inp); ws.write('C36',0,na)
ws.write('E36','Principal part of EMI. Stamp duty & registration also qualifies in payment year.', hint)

ws.write('A37','  ↳ Tuition Fees (Children)'); ws.write('B37',0,inp); ws.write('C37',0,na)
ws.write('E37','Tuition fees only for up to 2 children in full-time education at Indian institutions.', hint)

ws.write('A38','  ↳ NSC / Tax Saver FD (5-yr) / SCSS'); ws.write('B38',0,inp); ws.write('C38',0,na)
ws.write('E38','5-yr lock-in. NSC interest auto-qualifies for 80C annually. SCSS rate: 8.2% p.a. (Q1 FY26-27).', hint)

ws.write('A39','  ↳ Sukanya Samriddhi Yojana (SSY)'); ws.write('B39',0,inp); ws.write('C39',0,na)
ws.write('E39','Girl child below 10 yrs. Min ₹250, Max ₹1,50,000/yr. Rate 8.2% p.a. EEE status.', hint)

ws.write('A40','  ↳ 80CCC (Pension/Annuity Plan Premium)'); ws.write('B40',0,inp); ws.write('C40',0,na)
ws.write('E40','Pension/annuity plan premium from insurer. Within ₹1.5L combined 80CCE ceiling. (Budget 2015).', hint)

ws.write('A41','  80C Total (auto-summed)')
ws.write_formula('B41','=SUM(B32:B40)', sub_h); ws.write('C41',0,na)
ws.write('E41','Sum of all 80C investments above.', hint)

ws.write('A42','  80C Eligible (capped at ₹1,50,000)')
ws.write_formula('B42','=MIN(B41,150000)', ok_fmt); ws.write('C42',0,na)
ws.write('E42','Actual deduction = min(total, ₹1,50,000). Limit unchanged since Budget 2014.', hint)

ws.write('A43','  ⚠ 80C Remaining Limit')
ws.write_formula('B43','=MAX(0,150000-B41)', warn); ws.write('C43',0,na)
ws.write('E43','Amount still available to invest. Invest more to save additional tax.', hint)

ws.write('A44','Sec 80CCD(1B) — Own NPS Tier-1 (EXTRA ₹50,000)')
ws.write('B44',0,inp); ws.write('C44',0,na)
ws.write('E44','ADDITIONAL ₹50,000 OVER AND ABOVE 80C ₹1.5L. Own voluntary NPS Tier-1. NOT in New Regime. (Budget 2015).', hint)

ws.write('A45','Sec 80CCD(2) — Employer NPS ✅ BOTH REGIMES')
ws.write('B45',0,inp); ws.write_formula('C45','=B45',calc)
ws.write('E45','BOTH REGIMES. Old: up to 10% of Basic+DA. New: up to 14% of Basic+DA (Budget 2024). Enter actual employer NPS contribution.', hint)

ws.write('A46','Sec 24(b) — Home Loan Interest (Self-Occupied)')
ws.write('B46',0,inp); ws.write('C46',0,na)
ws.write('E46','Interest on home loan. Self-occupied: max ₹2,00,000/yr. NOT in New Regime. (Budget 2017 cap).', hint)

# 80D block header
ws.write('A47','SEC 80D — HEALTH INSURANCE | OLD REGIME ONLY (Budget 2018)', sec_sub)
ws.write('B47','', sec_sub); ws.write('C47','', sec_sub); ws.write('D47','', sec_sub); ws.write('E47','', sec_sub)

ws.write('A48','  ↳ Self + Spouse + Children (below 60 yrs)'); ws.write('B48',0,inp); ws.write('C48',0,na)
ws.write('E48','Max ₹25,000/yr including preventive health check-up (max ₹5,000 within). (Budget 2018).', hint)

ws.write('A49','  ↳ Parents (below 60 yrs)'); ws.write('B49',0,inp); ws.write('C49',0,na)
ws.write('E49','Additional ₹25,000 for parents below 60. Separate from self+family limit. (Budget 2018).', hint)

ws.write('A50','  ↳ Parents (Senior Citizen, 60+ yrs)'); ws.write('B50',0,inp); ws.write('C50',0,na)
ws.write('E50','If parents 60+, limit ₹50,000 (replaces ₹25k). Self also 60+: self limit ₹50,000. (Budget 2018).', hint)

ws.write('A51','  ↳ Preventive Health Check-up'); ws.write('B51',0,inp); ws.write('C51',0,na)
ws.write('E51','Max ₹5,000 within 80D overall limit. Cash payment allowed. (Budget 2013).', hint)

ws.write('A52','  80D Total Eligible')
ws.write_formula('B52','=MIN(B48+B49+B50+B51,IF(B50>0,75000,50000))', ok_fmt); ws.write('C52',0,na)
ws.write('E52','Auto-capped: senior citizen parents → max ₹75,000; otherwise max ₹50,000.', hint)

ws.write('A53','Sec 80TTA — Savings Account Interest'); ws.write('B53',0,inp); ws.write('C53',0,na)
ws.write('E53','Interest on savings account only (NOT FD). Max ₹10,000/yr. Senior citizens use 80TTB (₹50,000). (Budget 2013).', hint)

ws.write('A54','Sec 80G — Donations to Charity'); ws.write('B54',0,inp); ws.write('C54',0,na)
ws.write('E54','100% deduction: PM Relief Fund. 50%: other approved charities (capped at 10% of adj. gross income). Cash >₹2,000 disallowed.', hint)

ws.write('A55','Sec 80E — Education Loan Interest'); ws.write('B55',0,inp); ws.write('C55',0,na)
ws.write('E55','NO upper limit. Allowed for 8 yrs from start of repayment. Self/spouse/children. NOT in New Regime.', hint)

ws.write('A56','Sec 80EEA — First Home Loan Interest (extra ₹1.5L)'); ws.write('B56',0,inp); ws.write('C56',0,na)
ws.write('E56','Additional ₹1,50,000 over Sec 24(b). First-time buyer. Stamp duty value ≤₹45L. Loan sanctioned 1-Apr-2019 to 31-Mar-2022 only.', hint)

ws.write('A57','Sec 80GG — Rent Paid (if NO HRA in salary)'); ws.write('B57',0,inp); ws.write('C57',0,na)
ws.write('E57','ONLY if HRA not in salary. Least of: Rent−10% income, 25% income, ₹5,000/month. NOT in New Regime.', hint)

ws.write('A58','Sec 80DD — Disabled Dependent Care'); ws.write('B58',0,inp); ws.write('C58',0,na)
ws.write('E58','₹75,000 (40-79% disability) or ₹1,25,000 (80%+ severe). Dependent family member. Certificate required. (Budget 2015).', hint)

ws.write('A59','Sec 80DDB — Medical Treatment (Specified Disease)'); ws.write('B59',0,inp); ws.write('C59',0,na)
ws.write('E59','Cancer, neurological, AIDS etc. Max ₹40,000 (₹1,00,000 senior citizens). Doctor certificate required.', hint)

ws.write('A60','Sec 80U — Self with Disability'); ws.write('B60',0,inp); ws.write('C60',0,na)
ws.write('E60','₹75,000 (40-79%) or ₹1,25,000 (80%+ severe) if employee self is disabled. Medical certificate required. (Budget 2015).', hint)

ws.write('A61','TOTAL CHAPTER VI-A DEDUCTIONS', res)
ws.write_formula('B61','=B42+B44+B45+B46+B52+B53+B54+B55+B56+B57+B58+B59+B60', res)
ws.write_formula('C61','=C45', res)
ws.write('E61','Old Regime: all deductions. New Regime: only Employer NPS 80CCD(2).', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION D – NET TAXABLE INCOME
ws.merge_range('A63:E63','SECTION D: NET TAXABLE INCOME', sec)
ws.write('A64','Component', hdr); ws.write('B64','Old Regime (₹)', hdr); ws.write('C64','New Regime (₹)', hdr)

ws.write('A65','Gross Income')
ws.write_formula('B65','=D14', calc); ws.write_formula('C65','=D14', calc)
ws.write('A66','Less: Exemptions (Section B)')
ws.write_formula('B66','=B27', calc); ws.write_formula('C66','=C27', calc)
ws.write('A67','Less: Chapter VI-A Deductions (Section C)')
ws.write_formula('B67','=B61', calc); ws.write_formula('C67','=C61', calc)
ws.write('A68','NET TAXABLE INCOME', res)
ws.write_formula('B68','=MAX(0,B65-B66-B67)', res)
ws.write_formula('C68','=MAX(0,C65-C66-C67)', res)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION E – TAX CALCULATION
ws.merge_range('A70:E70','SECTION E: FINAL TAX CALCULATION', sec)
ws.write('A71','Tax Component', hdr); ws.write('B71','Old Regime (₹)', hdr)
ws.write('C71','New Regime (₹)', hdr); ws.write('E71','Notes', hdr)

old_tax = ('=IF(B68<=250000,0,'
           'IF(B68<=500000,(B68-250000)*0.05,'
           'IF(B68<=1000000,12500+(B68-500000)*0.2,'
           '112500+(B68-1000000)*0.3)))')
new_tax = ('=IF(C68<=400000,0,'
           'IF(C68<=800000,(C68-400000)*0.05,'
           'IF(C68<=1200000,20000+(C68-800000)*0.10,'
           'IF(C68<=1600000,60000+(C68-1200000)*0.15,'
           'IF(C68<=2000000,120000+(C68-1600000)*0.20,'
           'IF(C68<=2400000,200000+(C68-2000000)*0.25,'
           '300000+(C68-2400000)*0.30))))))')

ws.write('A72','Tax on Income (as per slabs)')
ws.write_formula('B72', old_tax, calc); ws.write_formula('C72', new_tax, calc)
ws.write('E72','Old: 0%≤2.5L | 5% 2.5-5L | 20% 5-10L | 30%>10L. New: 0%≤4L | 5% 4-8L | 10% 8-12L | 15% 12-16L | 20% 16-20L | 25% 20-24L | 30%>24L. (Budget 2025).', hint)

old_sur = ('=IF(B68<=5000000,0,'
           'IF(B68<=10000000,B72*0.10,'
           'IF(B68<=20000000,B72*0.15,'
           'IF(B68<=50000000,B72*0.25,'
           'B72*0.37))))')
new_sur  = ('=IF(C68<=5000000,0,'
            'IF(C68<=10000000,C72*0.10,'
            'C72*0.25))')

ws.write('A73','Surcharge')
ws.write_formula('B73', old_sur, calc); ws.write_formula('C73', new_sur, calc)
ws.write('E73','Old: 10%/>₹50L, 15%/>₹1Cr, 25%/>₹2Cr, 37%/>₹5Cr. New: capped at 25% (Budget 2023).', hint)

ws.write('A74','Tax + Surcharge')
ws.write_formula('B74','=B72+B73', calc); ws.write_formula('C74','=C72+C73', calc)

ws.write('A75','Rebate u/s 87A')
ws.write_formula('B75','=IF(B68<=500000,MIN(B74,12500),0)', calc)
ws.write_formula('C75','=IF(C68<=1200000,MIN(C74,60000),0)', calc)
ws.write('E75','Old: zero tax if income ≤₹5L. New: zero tax if income ≤₹12L (₹12.75L effective with std deduction). (Budget 2025).', hint)

ws.write('A76','Tax after Rebate')
ws.write_formula('B76','=MAX(0,B74-B75)', calc); ws.write_formula('C76','=MAX(0,C74-C75)', calc)

ws.write('A77','Health & Education Cess (4%)')
ws.write_formula('B77','=B76*0.04', calc); ws.write_formula('C77','=C76*0.04', calc)
ws.write('E77','4% on tax after rebate. Both regimes. (Budget 2018, unchanged).', hint)

ws.write('A78','TOTAL TAX PAYABLE', res)
ws.write_formula('B78','=B76+B77', res); ws.write_formula('C78','=C76+C77', res)

ws.write('A79','TDS Already Deducted by Employer (this FY)')
ws.write('B79',0,inp); ws.write('C79',0,inp)
ws.write('E79','Total TDS from salary + variable pay + RSU TDS this FY. Check Form 16 Part A or all monthly payslips.', hint)

ws.write('A80','Balance Tax Payable / (Refund Due)')
ws.write_formula('B80','=B78-B79', alrt); ws.write_formula('C80','=C78-C79', alrt)
ws.write('E80','NEGATIVE = refund when filing ITR. POSITIVE = additional tax to pay. Advance tax required if liability >₹10,000 (by Mar 15).', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION F – VERDICT
ws.merge_range('A82:E82','SECTION F: VERDICT & RECOMMENDATION', sec)

ws.write('A83','DIFFERENCE: Old Regime Tax − New Regime Tax', hdr)
ws.write_formula('B83','=B78-C78', alrt)
ws.write('E83','POSITIVE = New Regime saves you money. NEGATIVE = Old Regime is better.', hint)

ws.write('A84','Effective Tax Rate — Old Regime')
ws.write_formula('B84','=IF(D14>0,ROUND(B78/D14*100,2),0)', calc)
ws.write('E84','Tax as % of gross income. Old Regime.', hint)

ws.write('A85','Effective Tax Rate — New Regime')
ws.write_formula('B85','=IF(D14>0,ROUND(C78/D14*100,2),0)', calc)
ws.write('E85','Tax as % of gross income. New Regime.', hint)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION G – TAX SAVING CHECKLIST
ws.merge_range('A87:E87','SECTION G: TAX SAVING CHECKLIST — Did you claim everything? (Old Regime)', sec)
ws.write('A88','Opportunity', hdr); ws.write('E88','Detail & Budget Reference', hdr)

tips = [
    ('✅ Fill 80C limit of ₹1.5L fully','ELSS (3-yr lock-in), PPF (7.1%), NSC. Many miss ₹10k-₹50k. Budget 2014 — limit ₹1,50,000.'),
    ('✅ NPS 80CCD(1B) — Extra ₹50,000','Additional ₹50k beyond 80C. At 30% slab = ₹15,000 saved. Budget 2015.'),
    ('✅ NPS 80CCD(2) — Ask employer','Up to 14% of Basic in New Regime too. Best dual-benefit deduction available. Budget 2024.'),
    ('✅ 80D — Health insurance for parents','Senior citizen parents limit: ₹50,000. Buying policy saves tax + provides coverage. Budget 2018.'),
    ('✅ Sec 24(b) — Home loan interest','Submit interest certificate to employer. Joint loan: both co-owners claim ₹2L each. Budget 2017.'),
    ('✅ HRA + Home Loan both allowed','Claim HRA (renting at work city) AND home loan (property elsewhere). Not mutually exclusive.'),
    ('✅ Submit LTA bills — don\'t leave money','Unclaimed LTA = 100% taxable. 2 trips per block 2022-2025. Domestic only. Submit bills before FY end.'),
    ('✅ Sec 80E — Education loan (no cap)','NO upper limit. 8 years from repayment start. Self/spouse/children. Often forgotten.'),
    ('✅ Sec 80G — Donate to save tax','PM Relief Fund: 100% no limit. NGOs: 50%. Non-cash only if amount >₹2,000.'),
    ('✅ Meal coupons / Sodexo','₹50/meal × 22 days × 12 = ₹13,200/yr tax-free. Ask HR about meal voucher facility.'),
    ('✅ Telephone/Internet reimbursement','Actual bills reimbursed for official use = non-taxable perquisite. Submit bills to employer.'),
    ('✅ Car lease scheme','Company car lease: lower perquisite tax vs. petrol reimbursement. Check with HR.'),
    ('✅ Gratuity & Leave Encashment','Gratuity ≤₹20L tax-free (Budget 2018). Leave encashment at retirement ≤₹25L tax-free (Budget 2023).'),
    ('✅ Submit Form 12BB by Jan 15','Submit investment proofs by Jan 15 to employer. Late = large TDS spike in Feb-Mar.'),
    ('✅ Advance Tax if liability >₹10,000','Pay quarterly (Jun 15, Sep 15, Dec 15, Mar 15). Avoid interest u/s 234B/234C.'),
]
for i,(tip,detail) in enumerate(tips):
    ws.write(88+i, 0, tip,    sub_h)
    ws.write(88+i, 4, detail, hint)

workbook.close()

print("Excel generated at", full_path)