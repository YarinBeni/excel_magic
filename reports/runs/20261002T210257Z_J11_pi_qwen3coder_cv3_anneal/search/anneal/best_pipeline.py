import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Convert categorical columns to numeric using pandas' built-in methods
    for col in X_train.columns:
        if X_train[col].dtype == 'category':
            X_train[col] = X_train[col].cat.codes
            X_test[col] = X_test[col].cat.codes
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create a few meaningful interaction features
    # Based on domain knowledge of steel properties
    
    # Create interaction features
    if 'strength' in X_train.columns and 'carbon' in X_train.columns:
        X_train['strength_x_carbon'] = X_train['strength'] * X_train['carbon']
        X_test['strength_x_carbon'] = X_test['strength'] * X_test['carbon']
        
    if 'hardness' in X_train.columns and 'carbon' in X_train.columns:
        X_train['hardness_x_carbon'] = X_train['hardness'] * X_train['carbon']
        X_test['hardness_x_carbon'] = X_test['hardness'] * X_test['carbon']
    
    # Create a ratio feature
    if 'strength' in X_train.columns and 'carbon' in X_train.columns:
        X_train['strength_over_carbon'] = X_train['strength'] / (X_train['carbon'] + 1e-8)
        X_test['strength_over_carbon'] = X_test['strength'] / (X_test['carbon'] + 1e-8)
        
    # Create a combined feature from multiple variables
    if 'thickness' in X_train.columns and 'width' in X_train.columns:
        X_train['thickness_width_product'] = X_train['thickness'] * X_train['width']
        X_test['thickness_width_product'] = X_test['thickness'] * X_test['width']
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred