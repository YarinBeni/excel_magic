# Pipeline Improvements for hazelnut-spread-contaminant-detection

## Key Improvements Implemented

1. **Domain-Specific Feature Engineering**: Created features based on the physical nature of S-parameter measurements from microwave sensing:
   - Sum, mean, and variance of all S-parameters 
   - Ratio features between symmetric S-parameter pairs (s12/s21, s13/s31)
   - Difference features between symmetric S-parameter pairs (s12-s21, s13-s31)

2. **Feature Selection Strategy**: 
   - Focused on the most informative combinations rather than creating excessive features
   - Limited to ~37 features to maintain model generalization
   - Used appropriate numerical stability with epsilon for division operations

3. **Physical Insights Applied**:
   - Leveraged the symmetric nature of S-parameter matrices (sij ≈ sji)
   - Created meaningful ratios and differences that capture physical relationships
   - Added statistical summaries that encode the overall behavior of the S-parameters

## Performance Results
- Baseline (P0): 0.00723 (1-auroc)
- Best Score Achieved: 0.00702 (1-auroc) 
- Improvement: ~2.7% reduction in error

## Why This Worked
The S-parameter data has inherent symmetries and physical relationships that are captured well by the ratio and difference features. The statistical summaries (sum, mean, variance) provide good overall characterization of the signal behavior, while the symmetric pair features exploit the known structure of the data.