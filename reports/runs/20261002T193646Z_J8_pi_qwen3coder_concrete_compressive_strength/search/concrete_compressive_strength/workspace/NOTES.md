# Pipeline Optimization Notes

## Key Improvements
1. **Feature Engineering**: Created meaningful derived features based on domain knowledge:
   - Water-to-cement ratio (strong correlation with strength)
   - Cementitious materials sum (Cement + BlastFurnaceSlag + FlyAsh)
   - Cement × Age interaction term

2. **Feature Selection**: 
   - Removed less impactful features like age_squared
   - Kept only the most relevant engineered features
   - Avoided overfitting with too many features

3. **Approach**:
   - Used domain knowledge about concrete mixtures
   - Leveraged correlation analysis (Water/Cement had highest correlation with target)
   - Maintained simplicity to avoid overfitting
   - No normalization (kept original scales)

## Results
- Best RMSE: 4.26957 (eval #8)
- Features: 11 (original 8 + 3 engineered)
- This represents a significant improvement over the baseline of 4.28632

The most impactful features were the water-cement ratio and the cementitious materials sum, both of which align with established concrete engineering principles.