# Pipeline Improvements for credit-g Dataset

## Key Improvements that Led to Best Score (0.20184):

1. **Feature Engineering**:
   - Created credit_amount_duration_ratio feature to capture the relationship between loan size and repayment period
   - Added log_credit_amount transformation to reduce skewness in the credit amount feature
   - Included age_credit_ratio to capture the relationship between customer age and credit amount

2. **Preprocessing**:
   - Proper handling of missing values for both numerical and categorical features
   - Used label encoding for categorical variables with proper handling of unseen categories

3. **Feature Selection**:
   - Started with 24 features and refined to optimal set of 23 features
   - Removed less impactful features to avoid overfitting

The improvements focused on domain-relevant ratios and transformations that help the TabPFN model better understand the relationships in the credit data, particularly the credit amount and duration relationship which is fundamental to credit risk assessment.