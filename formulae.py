# Let me check my old tax regime slabs in formula.
# Old:
# 0-2.5L = 0
# 2.5L-5L = 5% => max 12,500
# 5L-10L = 20% => max 1,00,000
# >10L = 30%
# The old_tax_formula was: 112500+(B19-1000000)*0.3 => correct.

# New (FY 26-27):
# 0-4L = 0
# 4L-8L = 5% => max 20,000
# 8L-12L = 10% => max 40,000
# 12L-16L = 15% => max 60,000
# 16L-20L = 20% => max 80,000
# 20L-24L = 25% => max 100,000
# >24L = 30%
# Sum of max tax before each slab:
# at 8L: 20,000
# at 12L: 20k + 40k = 60,000 (my formula had 20000 + (C19-8L)*0.1 -> Wait, this gives 60k at 12L)
# at 16L: 60k + 60k = 120,000 (my formula had 60000 + (C19-12L)*0.15)
# at 20L: 120k + 80k = 200,000 (my formula had 120000 + (C19-16L)*0.20)
# at 24L: 200k + 100k = 300,000 (my formula had 200000 + (C19-20L)*0.25)
# >24L: 300000 + (C19-24L)*0.30

# The formulas in the Excel sheet are correct.
# Wait, let's verify standard deduction in New tax regime.
# For FY26-27, standard deduction is 75,000 [web:7].
# Let's verify 87A rebate for New tax regime.
# Resident individuals with taxable income up to 12 lakh pay zero income tax under the new tax regime (rebate up to Rs. 60,000) [web:9].
# In old regime, rebate up to 12,500 if income <= 5L.
# Both are correct.
print("Formulas verified.")