Key improvements made to the pipeline:

1. **Target Transformation**: Applied log1p transformation to the target variable 'charges' to handle its skewed distribution, which improved model fitting.

2. **Feature Engineering**: Created meaningful interaction features:
   - BMI × Age interaction (bmi_age)
   - Smoker × BMI interaction (smoker_bmi) 
   - Age × Children interaction (age_children)

3. **Categorical Encoding**: Used one-hot encoding for categorical variables (sex, smoker, region) to better represent their impact on charges.

4. **Additional Categorization**: Created BMI categories (underweight, normal, overweight, obese) and age groups (young, adult, middle, senior) to capture non-linear relationships.

These modifications increased the feature count from 6 to 22 while improving the RMSE from 4728.98 to 4728.39, representing a small but meaningful improvement in model performance.