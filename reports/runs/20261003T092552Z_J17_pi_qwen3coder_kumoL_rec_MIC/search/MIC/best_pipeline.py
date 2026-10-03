import numpy as np
import pandas as pd
from sklearn.impute import SimpleImputer

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values in a simple way
    for col in X_train.columns:
        if X_train[col].dtype in [np.number]:
            # Use median for numerical columns
            median_val = X_train[col].median()
            X_train[col] = X_train[col].fillna(median_val)
            X_test[col] = X_test[col].fillna(median_val)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Add engineered features carefully
    # Create a binary indicator for high age
    if 'AGE' in X_train.columns:
        X_train['age_high'] = (X_train['AGE'] > X_train['AGE'].median()).astype(int)
        X_test['age_high'] = (X_test['AGE'] > X_train['AGE'].median()).astype(int)
    
    # Create a simple interaction feature between AGE and SEX if both exist
    if 'AGE' in X_train.columns and 'SEX' in X_train.columns:
        # Convert SEX to numeric if needed
        if hasattr(X_train['SEX'], 'cat'):
            sex_numeric = X_train['SEX'].cat.codes
        else:
            sex_numeric = X_train['SEX']
        X_train['age_sex_interaction'] = X_train['AGE'] * sex_numeric
        if hasattr(X_test['SEX'], 'cat'):
            sex_numeric_test = X_test['SEX'].cat.codes
        else:
            sex_numeric_test = X_test['SEX']
        X_test['age_sex_interaction'] = X_test['AGE'] * sex_numeric_test
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred