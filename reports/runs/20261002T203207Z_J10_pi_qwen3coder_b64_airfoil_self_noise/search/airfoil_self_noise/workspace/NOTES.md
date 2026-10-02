# Improvements to airfoil_self_noise pipeline

## Key Improvements:
1. **Target Transformation**: Applied log transformation to the target variable to handle skewness, which improved model performance
2. **Categorical Encoding**: Used one-hot encoding for the categorical 'attack-angle' variable
3. **Feature Engineering**: Created meaningful physical interaction features:
   - Frequency/chord ratio
   - Velocity × chord product  
   - Reynolds-like dimensionless number
   - Chord/velocity ratio
4. **Additional Features**: Added square root and log transformations to reduce skew in input features

## Best Performing Version:
- Score: 1.31949 (RMSE)
- Features: 37 total
- Approach: Log-transformed target, engineered interaction features based on physical relationships in airfoil acoustics

The most impactful changes were the log transformation of the target and the creation of physics-based interaction features that capture the underlying relationships in the airfoil self-noise phenomenon.