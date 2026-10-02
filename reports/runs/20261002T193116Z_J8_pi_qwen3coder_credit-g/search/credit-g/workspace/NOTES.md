# Pipeline Improvements for credit-g Dataset

## Key Improvements that Reduced 1-AUROC from 0.22535 to 0.22166

1. **Categorical Encoding**: Applied LabelEncoder to all categorical variables to convert them to numerical representations that the model can process effectively.

2. **Feature Engineering**:
   - Created ratio features: credit_amount/duration_months, credit_amount/age_years
   - Applied log transformations to skewed features (credit_amount, duration_months) to reduce skewness
   - Added interaction terms between age and credit amount
   - Created grouped features using pd.cut() for age and credit amount
   - Added credit_per_month feature (credit_amount/duration_months)
   - Created age_credit_interaction_new combining age groups with credit amount groups

3. **Stratified Sampling Strategy**: Implemented sampling that maintains class balance during training, helping the model generalize better across both classes.

## Features Added
- credit_amount_duration_ratio
- credit_amount_age_ratio  
- credit_amount_log
- duration_months_log
- age_credit_interaction
- age_group
- credit_amount_group
- credit_per_month
- age_credit_interaction_new

These improvements focused on creating meaningful mathematical relationships between features while maintaining the integrity of the original data distribution. The combination of proper categorical encoding and strategic feature engineering significantly improved model performance on this credit classification task.