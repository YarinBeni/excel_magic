# Pipeline Improvements Summary

## Best Performing Configuration
Score: 0.00593 (1-auroc) - achieved with 35 features

## Key Improvements Made
1. **Symmetric Ratio Features**: Created ratios between symmetric S-parameter pairs (sij/sji) which are meaningful in microwave sensing applications
2. **Group Sum Features**: Calculated sums of S-parameters grouped by their first digit to capture overall patterns
3. **Selective Feature Engineering**: Used only the most promising combinations to avoid overfitting

## Approach Details
- Used ratios between symmetric pairs: (s12/s21), (s13/s31), (s14/s41)
- Created group sums for three major parameter groups
- Avoided excessive feature creation to maintain model generalization
- Kept the pipeline simple and interpretable while improving performance

## Why This Worked
The S-parameter data from microwave sensing has inherent symmetry properties that make ratios between symmetric pairs informative. Grouping parameters by their indices captures meaningful patterns in the data that the base model can leverage effectively.