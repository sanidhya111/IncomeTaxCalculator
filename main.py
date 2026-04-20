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
# Formats
hdr    = workbook.add_format({'bold':True,'bg_color':'#2E4057','font_color':'#FFFFFF','border':1,'font_size':10})
sec    = workbook.add_format({'bold':True,'bg_color':'#4A6FA5','font_color':'#FFFFFF','border':1,'font_size':10})
hint   = workbook.add_format({'italic':True,'font_color':'#555555','font_size':9,'text_wrap':True})
inp    = workbook.add_format({'bg_color':'#FFFFE0','border':1,'num_format':'#,##0','font_size':10})
calc   = workbook.add_format({'border':1,'num_format':'#,##0','bg_color':'#EBF5FB','font_size':10})
res    = workbook.add_format({'bold':True,'border':2,'num_format':'#,##0','bg_color':'#D5F5E3','font_size':10})
alrt   = workbook.add_format({'bold':True,'border':2,'num_format':'#,##0','bg_color':'#FEF9E7','font_color':'#C0392B','font_size':11})
sub_h  = workbook.add_format({'border':1,'num_format':'#,##0','bg_color':'#D6EAF8','bold':True,'font_size':9})
warn   = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#FDEDEC','font_color':'#C0392B','font_size':9})
ok_fmt = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#EAFAF1','font_color':'#1E8449','font_size':9})
dd     = workbook.add_format({'bg_color':'#FFFFE0','border':1,'font_size':10})
na     = workbook.add_format({'border':1,'bg_color':'#F2F3F4','font_color':'#AAA','font_size':9,'italic':True})
lta_warn = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#FDEDEC','font_color':'#C0392B','font_size':9})
lta_ok   = workbook.add_format({'bold':True,'border':1,'num_format':'#,##0','bg_color':'#EAFAF1','font_color':'#1E8449','font_size':9})
lta_calc = workbook.add_format({'border':1,'num_format':'#,##0','bg_color':'#FEF5E7','font_size':9})
sec_sub  = workbook.add_format({'bold':True,'bg_color':'#D4E6F1','border':1,'font_size':9})
verdict_big = workbook.add_format({'bold':True,'font_size':13,'align':'center','valign':'vcenter','border':2,'bg_color':'#D5F5E3','font_color':'#145A32','text_wrap':True})
save_green = workbook.add_format({'bold':True,'border':2,'num_format':'#,##0','bg_color':'#D5F5E3','font_color':'#145A32','font_size':12})
rate_fmt = workbook.add_format({'border':1,'num_format':'0.00','bg_color':'#EBF5FB','font_size':10})
text_box = workbook.add_format({'border':1,'bg_color':'#F8F9F9','text_wrap':True,'valign':'top','font_size':9})
helper   = workbook.add_format({'border':1,'num_format':'#,##0.00','bg_color':'#FEF5E7','font_size':9})
helper_int = workbook.add_format({'bg_color':'#FFFFE0','border':1,'num_format':'0.00','font_size':10})

ws.set_column('A:A', 52)
ws.set_column('B:B', 24)
ws.set_column('C:C', 24)
ws.set_column('D:D', 24)
ws.set_column('E:E', 82)

# Title
ws.merge_range('A1:E1','India Income Tax Calculator FY 2026-27 (AY 2027-28) — Salaried Employee',
               workbook.add_format({'bold':True,'bg_color':'#2E4057','font_color':'#FFD700','font_size':13,'align':'center','valign':'vcenter','border':1}))

# City
ws.write('A2','City Type (for HRA)')
ws.data_validation('B2', {'validate':'list','source':['Metro (Mumbai/Delhi/Chennai/Kolkata/Bengaluru/Hyderabad)','Non-Metro']})
ws.write('B2','Metro (Mumbai/Delhi/Chennai/Kolkata/Bengaluru/Hyderabad)',dd)
ws.write('E2','Metro = 50% of Basic for HRA limit; Non-Metro = 40%. (Budget 2005, unchanged)',hint)

# Section A
ws.merge_range('A4:E4','SECTION A: GROSS INCOME',sec)
ws.write('A5','Income Component',hdr); ws.write('B5','Monthly (₹)',hdr); ws.write('C5','Annual (₹)',hdr); ws.write('D5','Auto Annual (₹)',hdr); ws.write('E5','Hints & Limits',hdr)
ws.merge_range('A6:E6','⬛ Yellow = Input cells. Enter EITHER Monthly OR Annual — the other column auto-calculates. Blue = Auto-calculated.',workbook.add_format({'italic':True,'bg_color':'#FEF9E7','border':1,'font_size':9,'text_wrap':True}))

rows = [
    (7,'Basic Salary',89050,0,'Core fixed pay. PF, Gratuity, HRA limits derive from Basic.'),
    (8,'HRA Received (from employer)',35620,0,'HRA in your CTC. Enter monthly or annual. Exemption calculated in Section B.'),
    (9,'Special / Other Allowance (FBP)',16318,0,'Flexible Benefit Plan, Special Allowance, Meal, Phone etc. Fully taxable unless reimbursed.'),
    (10,'LTA Received (Leave Travel Allowance)',0,0,'Enter monthly LTA or annual LTA. Auto-Annual picks whichever you enter. LTA exemption & unclaimed tax calculated in Section B.'),
    (11,'Variable Pay / Performance Bonus',0,0,'Bonus, PMS payout, incentives. Fully taxable in both regimes. Employer deducts TDS when paid.'),
    (12,'Vested RSU / ESOP Income',0,0,'FMV on vest date − exercise price = perquisite taxable as salary. Employer deducts TDS. (Sec 17(2)).'),
    (13,'Other Income (FD Interest, Rent received, etc.)',0,0,'FD/RD interest, savings interest, rental income. Enter monthly or annual.')
]
for r,label,m,a,h in rows:
    ws.write(f'A{r}',label)
    ws.write(f'B{r}',m,inp)
    ws.write(f'C{r}',a,inp)
    ws.write_formula(f'D{r}',f'=IF(C{r}>0,C{r},B{r}*12)',calc)
    ws.write(f'E{r}',h,hint)

ws.write('A14','GROSS TOTAL INCOME',res)
ws.write_formula('D14','=SUM(D7:D13)',res)
ws.write('E14','Sum of all Auto-Annual income figures (column D).',hint)

# Section B
ws.merge_range('A16:E16','SECTION B: EXEMPTIONS (Reduce Gross Income — OLD REGIME ONLY unless noted)',sec)
ws.write('A17','Exemption Component',hdr); ws.write('B17','Old Regime (₹)',hdr); ws.write('C17','New Regime (₹)',hdr); ws.write('D17','',hdr); ws.write('E17','Hints & Limits',hdr)
ws.write('A18','Standard Deduction (Sec 16(ia))'); ws.write('B18',50000,calc); ws.write('C18',75000,calc); ws.write('E18','OLD: ₹50,000 (Budget 2019). NEW: ₹75,000 (Budget 2024).',hint)
ws.write('A19','Professional Tax (Sec 16(iii))'); ws.write('B19',2400,inp); ws.write('C19',0,na); ws.write('E19','Max ₹2,500/year. Old Regime only. NOT in New Regime.',hint)
ws.write('A20','  ↳ Rent Paid per Month (for HRA calc)'); ws.write('B20',0,inp); ws.write_formula('C20','=B20*12',calc); ws.write('E20','Enter monthly rent paid. Landlord PAN required if annual rent > ₹1 lakh.',hint)
hra_f = '=IF(C20=0,0,MIN(D8,MAX(0,C20-0.1*D7),IF(ISNUMBER(SEARCH("Metro",B2)),0.5*D7,0.4*D7)))'
ws.write('A21','  ↳ HRA Exemption — AUTO CALC (Sec 10(13A))'); ws.write_formula('B21',hra_f,sub_h); ws.write('C21',0,na); ws.write('E21','Least of: HRA received, Rent−10% Basic, 50%/40% of Basic. NOT in New Regime.',hint)
ws.write('A22','  ↳ LTA Received this FY (from Section A — AUTO)'); ws.write_formula('B22','=D10',sub_h); ws.write('C22',0,na); ws.write('E22','Auto-pulled from Section A Row 10.',hint)
ws.write('A23','  ↳ Actual Travel Bills Submitted (Claimed LTA)'); ws.write('B23',0,inp); ws.write('C23',0,na); ws.write('E23','Enter total travel bills submitted. Exemption = lower of LTA received or actual bills. Domestic only, max 2 trips in block. (Sec 10(5)).',hint)
ws.write('A24','  ↳ LTA Exemption — AUTO CALC (Sec 10(5))'); ws.write_formula('B24','=MIN(B22,B23)',lta_ok); ws.write('C24',0,na); ws.write('E24','Exemption = lower of LTA received vs bills submitted.',hint)
ws.write('A25','  ↳ ⚠ Unclaimed LTA (Taxable)'); ws.write_formula('B25','=MAX(0,B22-B24)',lta_warn); ws.write('C25',0,na); ws.write('E25','LTA received minus exemption. This amount remains fully taxable.',hint)
ws.write('A26','  ↳ 📊 Estimated Tax on Unclaimed LTA (at marginal slab)'); ws.write_formula('B26','=IF(B25=0,0,IF(D14>1000000,B25*0.30,IF(D14>500000,B25*0.20,IF(D14>250000,B25*0.05,0))))',lta_calc); ws.write('C26',0,na); ws.write('E26','Indicative tax cost of not claiming LTA. Actual depends on final taxable income.',hint)
ws.write('A27','TOTAL EXEMPTIONS',res); ws.write_formula('B27','=B18+B19+B21+B24',res); ws.write_formula('C27','=C18',res); ws.write('E27','Old: Standard Deduction + Professional Tax + HRA + LTA. New: Standard Deduction only.',hint)

# Section C
ws.merge_range('A29:E29','SECTION C: DEDUCTIONS UNDER CHAPTER VI-A (Old Regime Only unless noted)',sec)
ws.write('A30','Deduction Component',hdr); ws.write('B30','Old Regime (₹)',hdr); ws.write('C30','New Regime (₹)',hdr); ws.write('E30','Hints & Limits',hdr)
for c in ['A','B','C','D','E']: ws.write(f'{c}31','',sec_sub)
ws.write('A31','SEC 80C / 80CCC / 80CCD(1) — Combined ceiling ₹1,50,000 | OLD REGIME ONLY (Budget 2014)',sec_sub)

ws.write('A32','  ↳ EPF / PF — Employee Contribution (AUTO)'); ws.write_formula('B32','=ROUND(D7*0.12,0)',sub_h); ws.write('C32',0,na); ws.write('E32','AUTO: 12% of Annual Basic (from D7).',hint)
ws.write('A33','  ↳ PPF (Public Provident Fund)'); ws.write('B33',0,inp); ws.write('C33',0,na); ws.write('E33','Min ₹500, Max ₹1,50,000/year. 15-year lock-in. Rate 7.1% p.a. (Q1 FY26-27).',hint)
ws.write('A34','  ↳ Life Insurance Premium (LIC / Term)'); ws.write('B34',0,inp); ws.write('C34',0,na); ws.write('E34','Premium ≤10% of sum assured for full benefit.',hint)
ws.write('A35','  ↳ ELSS Mutual Funds'); ws.write('B35',0,inp); ws.write('C35',0,na); ws.write('E35','3-year lock-in. LTCG >₹1.25L taxed at 12.5% (Budget 2024).',hint)

# Principal split block
ws.write('A36','  ↳ Home Loan Principal Paid — Actual Total (full loan)')
ws.write('B36',0,inp); ws.write('C36',0,na)
ws.write('E36','Enter total annual principal repaid for the whole loan, before splitting with spouse/co-borrower. This is the actual number from loan statement / principal certificate.',hint)
ws.write('A37','     Division Factor for Principal Share')
ws.write('B37',2,helper_int); ws.write('C37',0,na)
ws.write('E37','Enter 2 if you and your wife are sharing equally, 1 if fully yours, 3 if split among 3 claimants, etc. Use your actual ownership/repayment ratio for correct claim.',hint)
ws.write('A38','     Allowed Principal Claim — Your Share (AUTO)')
ws.write_formula('B38','=IF(B37>0,B36/B37,0)',ok_fmt); ws.write('C38',0,na)
ws.write('E38','AUTO: Actual principal ÷ division factor = your claimable principal share before the overall 80C cap is applied. This value is what flows into the 80C total.',hint)

ws.write('A39','  ↳ Tuition Fees (Children)'); ws.write('B39',0,inp); ws.write('C39',0,na); ws.write('E39','Up to 2 children, tuition only.',hint)
ws.write('A40','  ↳ NSC / Tax Saver FD (5-yr) / SCSS'); ws.write('B40',0,inp); ws.write('C40',0,na); ws.write('E40','5-year lock-in. SCSS rate 8.2% p.a. (Q1 FY26-27).',hint)
ws.write('A41','  ↳ Sukanya Samriddhi Yojana (SSY)'); ws.write('B41',0,inp); ws.write('C41',0,na); ws.write('E41','Girl child below 10. Max ₹1.5L/year. Rate 8.2% p.a.',hint)
ws.write('A42','  ↳ 80CCC (Pension/Annuity Plan Premium)'); ws.write('B42',0,inp); ws.write('C42',0,na); ws.write('E42','Within the combined ₹1.5L cap.',hint)
ws.write('A43','  80C Total (auto-summed)'); ws.write_formula('B43','=SUM(B32:B35,B38:B42)',sub_h); ws.write('C43',0,na); ws.write('E43','Total of all 80C items including only your share of home-loan principal.',hint)
ws.write('A44','  80C Eligible (capped at ₹1,50,000)'); ws.write_formula('B44','=MIN(B43,150000)',ok_fmt); ws.write('C44',0,na); ws.write('E44','Final allowed 80C deduction after applying the overall ₹1.5 lakh cap.',hint)
ws.write('A45','  ⚠ 80C Remaining Limit'); ws.write_formula('B45','=MAX(0,150000-B43)',warn); ws.write('C45',0,na); ws.write('E45','You can still invest this much under 80C.',hint)

ws.write('A46','Sec 80CCD(1B) — Own NPS Tier-1 (EXTRA ₹50,000)'); ws.write('B46',0,inp); ws.write('C46',0,na); ws.write('E46','Extra ₹50,000 over 80C. NOT in New Regime. (Budget 2015).',hint)
ws.write('A47','Sec 80CCD(2) — Employer NPS ✅ BOTH REGIMES'); ws.write('B47',0,inp); ws.write_formula('C47','=B47',calc); ws.write('E47','Old: up to 10% of Basic+DA. New: up to 14% of Basic+DA (Budget 2024).',hint)

# Interest split block
ws.write('A48','Sec 24(b) — Home Loan Interest Paid — Actual Total (full loan)')
ws.write('B48',0,inp); ws.write('C48',0,na)
ws.write('E48','Enter total annual home-loan interest for the whole loan from the interest certificate, before splitting with your wife/co-borrower.',hint)
ws.write('A49','  ↳ Division Factor for Interest Share')
ws.write('B49',2,helper_int); ws.write('C49',0,na)
ws.write('E49','Enter 2 for equal sharing with your wife, 1 if fully yours, or another factor matching the actual split/ownership/repayment arrangement.',hint)
ws.write('A50','  ↳ Your Share of Interest (AUTO)')
ws.write_formula('B50','=IF(B49>0,B48/B49,0)',helper); ws.write('C50',0,na)
ws.write('E50','AUTO: Actual total interest ÷ division factor = your personal share of annual interest.',hint)
ws.write('A51','  ↳ Allowed Home Loan Interest Claim (AUTO)')
ws.write_formula('B51','=MIN(B50,200000)',ok_fmt); ws.write('C51',0,na)
ws.write('E51','AUTO: Allowed interest deduction based on your share, subject to the cap used in this calculator for self-occupied property. This is the amount used in total deductions.',hint)

for c in ['A','B','C','D','E']: ws.write(f'{c}52','',sec_sub)
ws.write('A52','SEC 80D — HEALTH INSURANCE | OLD REGIME ONLY (Budget 2018)',sec_sub)
ws.write('A53','  ↳ Self + Spouse + Children (below 60 yrs)'); ws.write('B53',0,inp); ws.write('C53',0,na); ws.write('E53','Max ₹25,000/year including preventive check-up.',hint)
ws.write('A54','  ↳ Parents (below 60 yrs)'); ws.write('B54',0,inp); ws.write('C54',0,na); ws.write('E54','Additional ₹25,000 for parents below 60.',hint)
ws.write('A55','  ↳ Parents (Senior Citizen, 60+ yrs)'); ws.write('B55',0,inp); ws.write('C55',0,na); ws.write('E55','₹50,000 if parents are senior citizens.',hint)
ws.write('A56','  ↳ Preventive Health Check-up'); ws.write('B56',0,inp); ws.write('C56',0,na); ws.write('E56','Max ₹5,000 within overall 80D limit.',hint)
ws.write('A57','  80D Total Eligible'); ws.write_formula('B57','=MIN(B53+B54+B55+B56,IF(B55>0,75000,50000))',ok_fmt); ws.write('C57',0,na); ws.write('E57','Auto-capped to ₹50k/₹75k as applicable.',hint)
ws.write('A58','Sec 80TTA — Savings Account Interest'); ws.write('B58',0,inp); ws.write('C58',0,na); ws.write('E58','Max ₹10,000/year on savings interest. (Budget 2013).',hint)
ws.write('A59','Sec 80G — Donations to Charity'); ws.write('B59',0,inp); ws.write('C59',0,na); ws.write('E59','50% or 100% deduction depending on institution. Cash >₹2,000 not allowed.',hint)
ws.write('A60','Sec 80E — Education Loan Interest'); ws.write('B60',0,inp); ws.write('C60',0,na); ws.write('E60','No upper limit. Allowed for 8 years. NOT in New Regime.',hint)
ws.write('A61','Sec 80EEA — First Home Loan Interest (extra ₹1.5L)'); ws.write('B61',0,inp); ws.write('C61',0,na); ws.write('E61','Extra ₹1.5L for eligible first-time buyers within sanction-date rules.',hint)
ws.write('A62','Sec 80GG — Rent Paid (if NO HRA in salary)'); ws.write('B62',0,inp); ws.write('C62',0,na); ws.write('E62','Only if HRA not part of salary. NOT in New Regime.',hint)
ws.write('A63','Sec 80DD — Disabled Dependent Care'); ws.write('B63',0,inp); ws.write('C63',0,na); ws.write('E63','₹75,000 / ₹1,25,000 depending on disability level.',hint)
ws.write('A64','Sec 80DDB — Medical Treatment (Specified Disease)'); ws.write('B64',0,inp); ws.write('C64',0,na); ws.write('E64','Specified disease treatment deduction with caps.',hint)
ws.write('A65','Sec 80U — Self with Disability'); ws.write('B65',0,inp); ws.write('C65',0,na); ws.write('E65','₹75,000 / ₹1,25,000 for self disability.',hint)
ws.write('A66','TOTAL CHAPTER VI-A DEDUCTIONS',res)
ws.write_formula('B66','=B44+B46+B47+B51+B57+B58+B59+B60+B61+B62+B63+B64+B65',res)
ws.write_formula('C66','=C47',res)
ws.write('E66','Old: all eligible deductions including only your share of principal and interest. New: only Employer NPS 80CCD(2).',hint)

# Section D
ws.merge_range('A68:E68','SECTION D: NET TAXABLE INCOME',sec)
ws.write('A69','Component',hdr); ws.write('B69','Old Regime (₹)',hdr); ws.write('C69','New Regime (₹)',hdr)
ws.write('A70','Gross Income'); ws.write_formula('B70','=D14',calc); ws.write_formula('C70','=D14',calc)
ws.write('A71','Less: Exemptions (Section B)'); ws.write_formula('B71','=B27',calc); ws.write_formula('C71','=C27',calc)
ws.write('A72','Less: Chapter VI-A Deductions (Section C)'); ws.write_formula('B72','=B66',calc); ws.write_formula('C72','=C66',calc)
ws.write('A73','NET TAXABLE INCOME',res); ws.write_formula('B73','=MAX(0,B70-B71-B72)',res); ws.write_formula('C73','=MAX(0,C70-C71-C72)',res)

# Section E
ws.merge_range('A75:E75','SECTION E: FINAL TAX CALCULATION',sec)
ws.write('A76','Tax Component',hdr); ws.write('B76','Old Regime (₹)',hdr); ws.write('C76','New Regime (₹)',hdr); ws.write('E76','Notes',hdr)
old_tax = '=IF(B73<=250000,0,IF(B73<=500000,(B73-250000)*0.05,IF(B73<=1000000,12500+(B73-500000)*0.2,112500+(B73-1000000)*0.3)))'
new_tax = '=IF(C73<=400000,0,IF(C73<=800000,(C73-400000)*0.05,IF(C73<=1200000,20000+(C73-800000)*0.10,IF(C73<=1600000,60000+(C73-1200000)*0.15,IF(C73<=2000000,120000+(C73-1600000)*0.20,IF(C73<=2400000,200000+(C73-2000000)*0.25,300000+(C73-2400000)*0.30))))))'
ws.write('A77','Tax on Income (as per slabs)'); ws.write_formula('B77',old_tax,calc); ws.write_formula('C77',new_tax,calc); ws.write('E77','Slabs reflect FY 2026-27 assumptions used in this calculator. (Budget 2025).',hint)
old_sur = '=IF(B73<=5000000,0,IF(B73<=10000000,B77*0.10,IF(B73<=20000000,B77*0.15,IF(B73<=50000000,B77*0.25,B77*0.37))))'
new_sur = '=IF(C73<=5000000,0,IF(C73<=10000000,C77*0.10,C77*0.25))'
ws.write('A78','Surcharge'); ws.write_formula('B78',old_sur,calc); ws.write_formula('C78',new_sur,calc); ws.write('E78','New Regime surcharge capped at 25%. (Budget 2023).',hint)
ws.write('A79','Tax + Surcharge'); ws.write_formula('B79','=B77+B78',calc); ws.write_formula('C79','=C77+C78',calc)
ws.write('A80','Rebate u/s 87A'); ws.write_formula('B80','=IF(B73<=500000,MIN(B79,12500),0)',calc); ws.write_formula('C80','=IF(C73<=1200000,MIN(C79,60000),0)',calc); ws.write('E80','Old: up to ₹12,500 if income ≤₹5L. New: up to ₹60,000 if income ≤₹12L. (Budget 2025).',hint)
ws.write('A81','Tax after Rebate'); ws.write_formula('B81','=MAX(0,B79-B80)',calc); ws.write_formula('C81','=MAX(0,C79-C80)',calc)
ws.write('A82','Health & Education Cess (4%)'); ws.write_formula('B82','=B81*0.04',calc); ws.write_formula('C82','=C81*0.04',calc); ws.write('E82','4% cess on tax after rebate. (Budget 2018, unchanged).',hint)
ws.write('A83','TOTAL TAX PAYABLE',res); ws.write_formula('B83','=B81+B82',res); ws.write_formula('C83','=C81+C82',res)
ws.write('A84','TDS Already Deducted by Employer (this FY)'); ws.write('B84',0,inp); ws.write('C84',0,inp); ws.write('E84','Enter total TDS already deducted from salary, bonus, RSU, etc.',hint)
ws.write('A85','Balance Tax Payable / (Refund Due)'); ws.write_formula('B85','=B83-B84',alrt); ws.write_formula('C85','=C83-C84',alrt); ws.write('E85','Negative = refund. Positive = extra tax payable.',hint)

# Section F improved
ws.merge_range('A87:E87','SECTION F: VERDICT & RECOMMENDATION',sec)
ws.write('A88','RECOMMENDED TAX REGIME',hdr)
ws.merge_range('B88:D88','=IF(B83<C83,"SELECT OLD REGIME",IF(C83<B83,"SELECT NEW REGIME","BOTH GIVE SAME TAX"))',verdict_big)
ws.write('E88','This is the actual decision output based on total tax payable comparison.',hint)
ws.write('A89','Why this regime is better',hdr)
ws.merge_range('B89:D89','=IF(B83<C83,"Old Regime tax is lower by ₹"&TEXT(C83-B83,"#,##0")&". Choose Old Regime.",IF(C83<B83,"New Regime tax is lower by ₹"&TEXT(B83-C83,"#,##0")&". Choose New Regime.","Both regimes give the same total tax. Choose based on simplicity / deduction preference."))',text_box)
ws.write('E89','Plain-English explanation of the recommendation.',hint)
ws.write('A90','Tax Saved by Recommended Regime',hdr)
ws.write_formula('B90','=ABS(B83-C83)',save_green)
ws.write('E90','Absolute rupee benefit of choosing the better regime.',hint)
ws.write('A91','Total Tax — Old Regime'); ws.write_formula('B91','=B83',calc)
ws.write('C91','Total Tax — New Regime'); ws.write_formula('D91','=C83',calc)
ws.write('E91','Shown together for immediate comparison.',hint)
ws.write('A92','Effective Tax Rate — Old Regime'); ws.write_formula('B92','=IF(D14>0,ROUND(B83/D14*100,2),0)',rate_fmt)
ws.write('C92','Effective Tax Rate — New Regime'); ws.write_formula('D92','=IF(D14>0,ROUND(C83/D14*100,2),0)',rate_fmt)
ws.write('E92','Tax as % of gross income under each regime.',hint)

# Section G
ws.merge_range('A94:E94','SECTION G: TAX SAVING CHECKLIST — Did you claim everything? (Old Regime)',sec)
ws.write('A95','Opportunity',hdr); ws.write('E95','Detail & Budget Reference',hdr)
tips = [
    ('✅ Fill 80C limit of ₹1.5L fully','ELSS, PPF, NSC can help fully use the ₹1.5L cap. Budget 2014.'),
    ('✅ NPS 80CCD(1B) — Extra ₹50,000','Extra deduction over 80C. Budget 2015.'),
    ('✅ NPS 80CCD(2) — Ask employer','Still allowed in New Regime too. Budget 2024 enhanced limit.'),
    ('✅ 80D — Health insurance for parents','Senior citizen parent limit ₹50,000. Budget 2018.'),
    ('✅ Split home loan correctly','Use actual total + division factor so only your share gets claimed.'),
    ('✅ HRA + Home Loan both allowed','Both can be claimed together in valid cases.'),
    ('✅ Submit LTA bills','Unclaimed LTA is fully taxable. Don’t leave this unclaimed.'),
    ('✅ Sec 80E — Education loan','No upper cap on interest deduction for eligible years.'),
    ('✅ Sec 80G — Donations','Use only approved institutions and non-cash >₹2,000.'),
    ('✅ Meal coupons / Sodexo','Potential tax-efficient structuring through employer.'),
    ('✅ Telephone/Internet reimbursement','Official-use reimbursements can reduce taxable salary.'),
    ('✅ Car lease scheme','May reduce tax versus standard car reimbursements.'),
    ('✅ Gratuity & Leave Encashment','Retirement-related tax exemptions matter in planning.'),
    ('✅ Submit Form 12BB on time','Avoid excess TDS due to late proof submission.'),
    ('✅ Advance Tax if needed','Avoid interest if residual liability exceeds threshold.'),
]
for i,(tip,detail) in enumerate(tips):
    ws.write(95+i,0,tip,sub_h)
    ws.write(95+i,4,detail,hint)

workbook.close()

print("Excel generated at", full_path)