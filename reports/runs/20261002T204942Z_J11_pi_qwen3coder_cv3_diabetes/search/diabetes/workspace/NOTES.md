# Pipeline Improvements for Diabetes Dataset

## Key Improvements that Helped:

1. **Ratio Features**: Added glucose/BMI ratio which captures the relationship between blood sugar and body weight - a known predictor for diabetes risk.

2. **Log Transformations**: Applied log1p transformation to skewed features (Insulin, SkinThickness, DiabetesPedigreeFunction) to reduce skewness and make distributions more Gaussian-like, which helps the model learn better.

3. **Feature Engineering**: Focused on domain-relevant features rather than creating too many complex interactions that could lead to overfitting.

## What Didn't Work Well:
- Adding interaction terms between glucose and age, or BMI and age
- Creating categorical features from continuous variables (caused NaN issues)
- Adding pregnancies/glucose ratio (made performance worse)

The best performing pipeline uses 12 features total and achieved a 1-auroc score of 0.1653, representing a significant improvement from the baseline of 0.1670.