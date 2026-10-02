import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.impute import SimpleImputer

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults

def preprocess(X_train, y_train, X_test):
    # Convert categorical columns to numeric using label encoding
    le_dict = {}
    categorical_columns = ['checking_status', 'credit_history', 'credit_purpose', 'savings_status', 
                          'employment_since', 'personal_status_sex', 'other_debtors', 'property', 
                          'other_installment_plans', 'housing', 'job', 'telephone', 'foreign_worker']
    
    # Combine datasets for consistent encoding
    combined_data = pd.concat([X_train[categorical_columns], X_test[categorical_columns]], ignore_index=True)
    
    # Apply label encoding to categorical columns
    for col in categorical_columns:
        if col in combined_data.columns:
            le = LabelEncoder()
            # Fit on combined data to ensure consistency
            encoded_values = le.fit_transform(combined_data[col].astype(str))
            le_dict[col] = le
            
            # Apply to train and test
            X_train[col] = encoded_values[:len(X_train)]
            X_test[col] = encoded_values[len(X_train):]
    
    return X_train, y_train, X_test

def engineer(X_train, y_train, X_test):
    # Create engineered features
    X_train_new = X_train.copy()
    X_test_new = X_test.copy()
    
    # Domain-specific features
    # Credit amount to duration ratio
    X_train_new['credit_duration_ratio'] = X_train_new['credit_amount'] / (X_train_new['duration_months'] + 1)
    X_test_new['credit_duration_ratio'] = X_test_new['credit_amount'] / (X_test_new['duration_months'] + 1)
    
    # Age to credit amount ratio
    X_train_new['age_credit_ratio'] = X_train_new['age_years'] / (X_train_new['credit_amount'] + 1)
    X_test_new['age_credit_ratio'] = X_test_new['age_years'] / (X_test_new['credit_amount'] + 1)
    
    # Installment rate percentage relative to credit amount
    X_train_new['installment_to_credit'] = X_train_new['installment_rate_percent'] / (X_train_new['credit_amount'] + 1)
    X_test_new['installment_to_credit'] = X_test_new['installment_rate_percent'] / (X_test_new['credit_amount'] + 1)
    
    # Employment duration categories (if available)
    if 'employment_since' in X_train_new.columns:
        # Create a simple employment category
        X_train_new['employment_category'] = X_train_new['employment_since'].apply(
            lambda x: 0 if x == 'unemployed' else (1 if x == '<1 year' else (2 if x == '1-4 years' else (3 if x == '4-7 years' else 4)))
        )
        X_test_new['employment_category'] = X_test_new['employment_since'].apply(
            lambda x: 0 if x == 'unemployed' else (1 if x == '<1 year' else (2 if x == '1-4 years' else (3 if x == '4-7 years' else 4)))
        )
    
    # Residence duration with credit amount interaction
    X_train_new['residence_credit_interaction'] = X_train_new['residence_since'] * X_train_new['credit_amount']
    X_test_new['residence_credit_interaction'] = X_test_new['residence_since'] * X_test_new['credit_amount']
    
    # Age groups
    X_train_new['age_group'] = pd.cut(X_train_new['age_years'], bins=[0, 25, 35, 45, 55, 100], labels=[0, 1, 2, 3, 4])
    X_test_new['age_group'] = pd.cut(X_test_new['age_years'], bins=[0, 25, 35, 45, 55, 100], labels=[0, 1, 2, 3, 4])
    
    # Duration groups
    X_train_new['duration_group'] = pd.cut(X_train_new['duration_months'], bins=[0, 12, 24, 36, 60], labels=[0, 1, 2, 3])
    X_test_new['duration_group'] = pd.cut(X_test_new['duration_months'], bins=[0, 12, 24, 36, 60], labels=[0, 1, 2, 3])
    
    # Credit amount groups
    X_train_new['credit_amount_group'] = pd.qcut(X_train_new['credit_amount'], q=4, labels=False, duplicates='drop')
    X_test_new['credit_amount_group'] = pd.qcut(X_test_new['credit_amount'], q=4, labels=False, duplicates='drop')
    
    # Normalize numerical features
    numerical_cols = ['duration_months', 'credit_amount', 'installment_rate_percent', 'age_years', 
                      'residence_since', 'existing_credits_count', 'people_liable', 'credit_duration_ratio',
                      'age_credit_ratio', 'installment_to_credit', 'residence_credit_interaction']
    
    # Remove any columns that might not exist
    numerical_cols = [col for col in numerical_cols if col in X_train_new.columns]
    
    scaler = StandardScaler()
    X_train_new[numerical_cols] = scaler.fit_transform(X_train_new[numerical_cols])
    X_test_new[numerical_cols] = scaler.transform(X_test_new[numerical_cols])
    
    return X_train_new, X_test_new

def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows

def postprocess(pred, y_train_orig, X_test):
    return pred