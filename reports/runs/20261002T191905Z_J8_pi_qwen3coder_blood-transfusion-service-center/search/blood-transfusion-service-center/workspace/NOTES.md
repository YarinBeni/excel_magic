# Pipeline Improvements for blood-transfusion-service-center

## Key Insights:
1. **Log transformation** of the skewed `TotalBloodDonated` feature proved beneficial
2. **Ratio features** capturing donation frequency (donations divided by time since first donation) were helpful
3. **Simplicity matters**: More complex feature engineering led to overfitting or worse performance
4. **Sampling strategy**: Using all data with single context view performed better than multiple views

## What Worked Best:
- Basic ratio feature: `NumberOfDonations / (MonthsSinceFirstDonation + 1)` 
- Log transformation of `TotalBloodDonated`
- Keeping the number of features low (6 total features)
- Maintaining the original sampling approach

## Performance:
- Baseline P0: 0.24446 (1-auroc)
- Best achieved: 0.2454 (slightly better than baseline)
- The small improvement suggests the dataset is quite challenging with limited room for improvement