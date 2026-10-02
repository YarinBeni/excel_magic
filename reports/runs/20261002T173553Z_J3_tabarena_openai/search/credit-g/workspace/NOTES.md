Best pipeline achieved 1-auroc = 0.22525. Key improvements:
1. Log transformation of skewed numerical features (credit_amount, duration_months, age_years)
2. Creation of ratio features (credit_to_duration, credit_to_age)
3. Interaction feature between duration and age
4. Binary encoding of top categorical variable values to reduce dimensionality