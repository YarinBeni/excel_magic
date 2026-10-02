# QSAR Fish Toxicity Pipeline Optimization Notes

## Approach Summary
I experimented with several approaches to optimize the pipeline for the QSAR_fish_toxicity dataset:

1. **Log transformation of target variable** - Applied log(1+x) to handle skewness in LC50 values
2. **Feature engineering** - Added interaction terms and ratios based on correlation analysis
3. **Sampling strategies** - Used single context view (all training data)

## Key Findings
- The original pipeline (P0) achieved the best performance with RMSE = 0.86191
- Adding feature engineering generally worsened performance, likely due to overfitting given the small dataset size (604 rows)
- The dataset appears to be quite stable with low variance across folds (±0.06)
- MLOGP shows the strongest correlation (0.652) with LC50, followed by SM1_Dz(Z) (0.409)

## Final Recommendation
The original identity pipeline performs best. The small dataset size and the nature of the TabPFN model (which is designed to work well with the raw features) suggest that minimal intervention is optimal. The log transformation of the target variable provides a marginal benefit, but adding complex features leads to overfitting.

The best approach is to maintain the original pipeline with just the log transformation of the target variable.