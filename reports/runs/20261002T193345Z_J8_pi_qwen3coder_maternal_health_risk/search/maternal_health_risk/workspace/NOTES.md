# Pipeline Improvements for Maternal Health Risk Dataset

## Best Configuration Found
The best performing pipeline uses:
1. **Minimal feature engineering**: Only one engineered feature (BP_Diff = SystolicBP - DiastolicBP)
2. **No feature scaling**: Keeping original scales proved better for this particular dataset
3. **Full training set usage**: Using all training data rather than sampling

## Key Insights
- Simple engineered features often outperform complex ones for this type of medical data
- The BP_Diff feature captures a clinically meaningful relationship between blood pressure measurements
- Using all training data instead of stratified sampling or multiple views gave the best results
- The frozen model benefits from having more complete information rather than multiple smaller views

## Performance
- Best logloss achieved: 0.53181 (compared to baseline P0 = 0.53212)
- This represents a slight improvement over the identity pipeline