# QSAR Fish Toxicity Pipeline Optimization Notes

## Initial Baseline
- P0 (identity pipeline) achieved RMSE = 0.84090
- 6 features: CIC0, SM1_Dz(Z), GATS1i, NdsCH, NdssC, MLOGP

## Attempts Made

1. **Log transformation of target variable** - RMSE increased slightly to 0.8455
2. **Feature engineering with ratios and interactions** - RMSE varied between 0.8454-0.8475
3. **Standardization of features** - RMSE remained at baseline 0.8409
4. **Combined approaches** - No improvement over baseline

## Conclusion
The simplest approach (identity pipeline) performed best. This suggests:
- The frozen model (kumo-tabular-l) is already quite effective with the original features
- The additional engineered features either don't help or potentially introduce noise
- The dataset may already contain sufficient signal in the original features
- The 3-fold CV score of 0.8409 is already quite good for this task

The optimal strategy appears to be maintaining the original feature set without modification.