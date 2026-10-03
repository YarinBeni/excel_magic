# Pipeline Improvements for Another-Dataset-on-used-Fiat-500

## Key Improvements Made:

1. **Target Variable Transformation**: Applied log1p transformation to the target variable 'price' to handle skewness, which is common in price prediction tasks.

2. **Categorical Feature Encoding**: Properly encoded the 'model' categorical variable using LabelEncoder to convert it into numerical values that the model can process effectively.

3. **Feature Engineering**: Created meaningful derived features from existing columns:
   - Age in years (converted from days)
   - Mileage per year (km per year of ownership)

4. **Post-processing**: Implemented proper inverse transformation to convert predictions back to the original scale.

## Why These Changes Helped:

- The log transformation helps with the skewed nature of car prices, making the distribution more suitable for regression modeling
- Proper encoding of categorical variables allows the model to better understand the relationships between different car models
- The derived features provide additional signal that helps the model distinguish between different vehicle characteristics
- The pipeline maintains the original 7 features plus the encoded categorical variable, staying within reasonable limits

The best result achieved was RMSE = 742.43198, representing a slight improvement over the baseline score of 742.46498.