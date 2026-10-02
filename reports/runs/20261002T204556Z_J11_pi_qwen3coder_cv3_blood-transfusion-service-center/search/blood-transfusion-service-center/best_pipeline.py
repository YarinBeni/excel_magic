import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Apply log transformation to skewed features
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Log transform TotalBloodDonated to reduce skewness
    X_train['TotalBloodDonated_log'] = np.log1p(X_train['TotalBloodDonated'])
    X_test['TotalBloodDonated_log'] = np.log1p(X_test['TotalBloodDonated'])
    
    # Remove original TotalBloodDonated column
    X_train = X_train.drop('TotalBloodDonated', axis=1)
    X_test = X_test.drop('TotalBloodDonated', axis=1)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # No additional engineering - keep it simple
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    # Use single context view (original approach)
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred