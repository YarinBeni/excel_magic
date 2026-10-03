# Pipeline Improvements for Marketing_Campaign Dataset

## Key Improvements that Helped:

1. **Feature Engineering**: Created meaningful derived features:
   - Age from Year_Birth 
   - Total spending across all product categories
   - Spending ratios (wine spending to total spending)
   - Purchase frequency metrics
   - Customer lifetime calculations

2. **Handling Missing Values**: Properly imputed missing Income values with median

3. **Log Transformations**: Applied log1p transformation to skewed features (Income, Total_Spending, Customer_Lifetime_Days)

4. **Categorical Encoding**: Converted Education and Marital_Status to dummy variables

5. **Interaction Features**: Created meaningful interactions like Income × Age

The best performing pipeline achieved 1-auroc = 0.11053 with 46 features, representing a 11.3% improvement over the baseline score of 0.12470.

Further attempts to add more complex features (like binning, additional ratios) resulted in slightly worse performance, suggesting the current set of features is already quite effective for this particular dataset and model combination.