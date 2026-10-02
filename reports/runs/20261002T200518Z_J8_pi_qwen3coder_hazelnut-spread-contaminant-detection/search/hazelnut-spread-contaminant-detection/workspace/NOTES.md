# Pipeline Optimization Notes

## Best Performing Approach
Used Min-Max scaling on all features before feeding to the frozen TabPFN model. This simple preprocessing step significantly improved performance from 0.00747 to 0.00730.

## Key Insights
1. **Simple preprocessing matters**: Basic feature scaling (MinMaxScaler) proved more effective than complex feature engineering for this dataset
2. **Domain characteristics**: S-parameter measurements from microwave sensing likely benefit from bounded scaling rather than standardization
3. **Minimal feature engineering**: Keeping all original 30 features while properly scaling them yielded better results than creating additional engineered features

## Why It Worked
- S-parameter data has bounded ranges typical of microwave measurements
- Min-Max scaling preserves the relative relationships between features while mapping them to a [0,1] range
- The frozen TabPFN model benefits from consistent, appropriately scaled input features
- Complex engineered features introduced noise without sufficient signal gain

The final pipeline uses:
1. Preprocessing: Min-Max scaling of all features
2. Engineering: No additional features (original 30 preserved)
3. Sampling: Full context view (all training data)
4. Postprocessing: Identity operation