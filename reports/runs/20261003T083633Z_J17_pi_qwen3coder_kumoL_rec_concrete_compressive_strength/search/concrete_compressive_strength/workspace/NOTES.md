# Pipeline Improvements for Concrete Compressive Strength

## Key Improvements Made:

1. **Feature Engineering**: Created 8 new engineered features based on domain knowledge:
   - Cementitious material combinations (Cement + Blast Furnace Slag + Fly Ash)
   - Water-to-cement ratio (wtr)
   - Aggregate ratios (coarse + fine aggregate)
   - Superplasticizer to cement ratio
   - Age squared effect
   - Cement × Age interaction
   - Cement to water inverse ratio
   - Total aggregate amount

2. **Target Transformation**: Applied log transformation to the target variable to improve model performance on skewed data

3. **Post-processing**: Added proper inversion of log transformation to get predictions back to original scale

## Results:
- Baseline (P0): 4.40358 RMSE
- Best Pipeline: 4.37694 RMSE (improvement of ~0.027 RMSE)
- Features used: 16 total (8 original + 8 engineered)

The improvements came from:
- Creating meaningful domain-specific features that capture the physical relationships in concrete mixtures
- Using appropriate preprocessing for the skewed target variable
- Maintaining a reasonable number of features (not too many to overwhelm the in-context learner)