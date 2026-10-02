# Pipeline Improvements for qsar-biodeg Dataset

## Key Improvements Implemented:

1. **Feature Engineering**:
   - Added ratio features: Heavy_atoms/Oxygen_atoms ratio to capture molecular composition relationships
   - Applied log-transformations to skewed numerical features to improve model performance
   - Created logarithmic versions of 7 skewed features to reduce the impact of outliers

2. **Sampling Strategy**:
   - Implemented stratified sampling with multiple views (3 views) to improve model robustness
   - This helps maintain class balance across different training contexts

3. **Domain Knowledge Features**:
   - Used chemical domain knowledge to create meaningful ratios
   - The Heavy_to_Oxygen_Ratio feature proved particularly useful for this biodegradation prediction task

## Performance Results:
- Baseline (P0): 0.06719 (1-auroc)
- Best Score Achieved: 0.06703 (1-auroc) 

The improvements focused on making the most of the limited 41 features by creating meaningful derived features that capture chemical properties relevant to biodegradability prediction, while maintaining the constraint of not modifying the frozen model itself.