import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Log transform target variable to reduce skewness
    y_train_log = np.log(y_train + 1)  # Adding 1 to avoid log(0)
    return X_train, y_train_log, X_test


def engineer(X_train, y_train, X_test):
    # Simple approach - add some engineered features
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Age in years instead of days
    X_train['age_years'] = X_train['age_in_days'] / 365.25
    X_test['age_years'] = X_test['age_in_days'] / 365.25
    
    # Standardize numerical features
    numerical_cols = ['engine_power', 'age_in_days', 'km', 'previous_owners', 'age_years']
    for col in numerical_cols:
        if col in X_train.columns:
            mean_val = X_train[col].mean()
            std_val = X_train[col].std()
            if std_val > 0:  # Avoid division by zero
                X_train[col] = (X_train[col] - mean_val) / std_val
                X_test[col] = (X_test[col] - mean_val) / std_val
    
    # Handle categorical variables - just keep as is for now
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert log transformation
    return np.exp(pred) - 1