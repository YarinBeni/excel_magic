# Pipeline Improvements Summary

## Key Improvements that Reduced 1-AUROC from 0.26431 to 0.25974

1. **Feature Engineering**: Created 14 derived features including:
   - Ratio features (donation rate, donations per month, recent activity)
   - Log transformations for skewed features (TotalBloodDonated)
   - Interaction features (NumberOfDonations × MonthsSinceLastDonation)
   - Reciprocal features for inverse relationships
   - Boolean indicators for frequent donors and large volume donors

2. **Multi-view Sampling Strategy**: Implemented stratified sampling across multiple context views:
   - View 1: Full training set
   - View 2: Stratified sampling preserving class balance
   - View 3: Random sample of training data
   This approach allows the model to learn from different perspectives of the data, improving generalization.

These improvements leverage the relationship between donation frequency, timing, and total blood volume to better predict whether someone donated blood in March 2007.