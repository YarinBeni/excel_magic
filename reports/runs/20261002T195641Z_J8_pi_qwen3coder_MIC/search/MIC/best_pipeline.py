import numpy as np
import pandas as pd
from sklearn.impute import SimpleImputer

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values properly
    # Fill numeric columns with median
    numeric_columns = X_train.select_dtypes(include=[np.number]).columns
    imputer = SimpleImputer(strategy='median')
    X_train[numeric_columns] = imputer.fit_transform(X_train[numeric_columns])
    X_test[numeric_columns] = imputer.transform(X_test[numeric_columns])
    
    # For categorical columns, fill with most frequent
    categorical_columns = X_train.select_dtypes(include=['category']).columns
    for col in categorical_columns:
        if X_train[col].isnull().any():
            mode_value = X_train[col].mode()[0] if not X_train[col].mode().empty else 'missing'
            X_train[col] = X_train[col].fillna(mode_value)
            X_test[col] = X_test[col].fillna(mode_value)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create interaction features between key variables
    # Age with sex
    X_train['AGE_SEX'] = X_train['AGE'] * (X_train['SEX'] == 1).astype(int)
    X_test['AGE_SEX'] = X_test['AGE'] * (X_test['SEX'] == 1).astype(int)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred