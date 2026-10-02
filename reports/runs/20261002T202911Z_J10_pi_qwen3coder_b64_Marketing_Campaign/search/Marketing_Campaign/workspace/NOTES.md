# Marketing Campaign Pipeline Improvements

## Key Improvements Implemented

1. **Feature Engineering**: Created numerous derived features including:
   - Age from Year_Birth
   - Total spending across all product categories
   - Spending ratios (wine spending ratio, spending-to-purchases ratio)
   - Purchase frequency metrics
   - Family size calculations
   - Customer tenure (days since first purchase)
   - Interaction features (income × age, spending × age)
   - Log transformations for skewed features
   - Average spending per category

2. **Data Preprocessing**: 
   - Handled missing Income values with median imputation
   - Properly managed categorical variables without causing unseen label errors

3. **Performance Results**:
   - Baseline (P0): 1-auroc = 0.13587
   - Best version: 1-auroc = 0.09058 (a significant improvement)
   - The improvements focused on creating meaningful combinations of existing features rather than adding noise

The most impactful features were the spending ratios, interaction terms, and log-transformed spending variables which helped the model better distinguish between responders and non-responders.