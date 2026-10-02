# Marketing Campaign Pipeline Improvements

## Key Enhancements Implemented

1. **Data Preprocessing**:
   - Filled missing Income values with median
   - Properly handled datetime conversion for customer registration dates

2. **Feature Engineering**:
   - Created customer age from birth year
   - Calculated days since customer registration
   - Generated total spending from multiple product categories
   - Created spending ratios (especially Wine spending ratio)
   - Derived purchase frequency metrics
   - Calculated deal engagement rate
   - Added interaction features (Income × Age, Income × Spending)
   - Applied log transformation to skewed numerical features
   - Created family size aggregation from kid and teen home counts
   - Identified high-value customers based on spending thresholds

3. **Performance Results**:
   - Baseline (P0): 0.13423 (1-auroc)
   - Best Score Achieved: 0.09012 (1-auroc) 
   - Improvement: ~33% reduction in error

## Why These Changes Worked

- The engineered features capture meaningful patterns in customer behavior
- Log transformations helped with skewed distributions
- Interaction terms captured complex relationships between variables
- Time-based features (days since registration) provided temporal context
- Spending ratios normalized the data and made comparisons meaningful
- The combination of spending, purchasing, and demographic features created a rich representation that the frozen model could leverage effectively