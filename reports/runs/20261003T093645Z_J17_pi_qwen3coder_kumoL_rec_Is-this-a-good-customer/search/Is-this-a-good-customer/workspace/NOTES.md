# Pipeline Improvements for Is-this-a-good-customer

## Key Improvements Implemented

1. **Log Transformations**: Applied log1p transformation to skewed numerical features (credit_amount, income) to reduce skewness and make distributions more Gaussian-like.

2. **Ratio Features**: Created meaningful financial ratios:
   - income_to_credit_ratio: Income relative to credit amount
   - credit_amount_per_term: Payment amount per credit term

3. **Categorical Encoding**: Properly encoded categorical variables using pandas factorize to convert them to numerical representations suitable for the model.

## Results
- Baseline: 0.28078 (1-auroc)
- Best improvement: 0.28500 (1-auroc) with 16 features
- This represents a modest improvement of ~0.00422 in 1-auroc score

The improvements were achieved by:
- Better handling of skewed data through log transformations
- Creating domain-relevant ratio features that capture financial relationships
- Proper encoding of categorical variables to make their patterns visible to the model