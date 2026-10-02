"""Candidate data pipeline around a FROZEN tabular foundation model.

The harness calls, in order:
  1. preprocess(X_train, y_train, X_test)        -> X_train, y_train, X_test   (cleaning, target transform)
  2. engineer(X_train, y_train, X_test)          -> X_train, X_test            (feature engineering, <= 500 cols)
  3. sample(X_train, y_train, X_test, max_rows)  -> list of index arrays       (context views; each view is fit separately
                                                                                 and predictions are averaged)
  4. frozen model fit/predict per view (never edit this part; MODEL_KWARGS tunes the constructor)
  5. postprocess(pred, y_train_orig, X_test)     -> pred                       (invert target transforms, calibrate)

`pred` is a pandas DataFrame of class probabilities (columns = class labels) for classification, or a
pandas Series for regression. `y_train_orig` is the untransformed training target. All inputs are pandas
objects; X_train / X_test share columns. Do not read any file other than the ones in this directory.
"""
import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying originals
    X_train_new = X_train.copy()
    X_test_new = X_test.copy()
    
    # Add target to train set for groupby operations
    y_train_series = y_train.reset_index(drop=True) if hasattr(y_train, 'reset_index') else pd.Series(y_train)
    
    # Combine train and test for consistent feature engineering
    df_train = X_train.copy()
    df_train['approved'] = y_train_series
    df_test = X_test.copy()
    df_combined = pd.concat([df_train, df_test], ignore_index=True)
    
    # Feature engineering based on domain knowledge
    # 1. Frequency of (role, resource) pairs - this is key to the problem
    role_resource_freq = df_combined.groupby(['role_id', 'resource_id']).size().reset_index(name='role_resource_count')
    df_combined = df_combined.merge(role_resource_freq, on=['role_id', 'resource_id'], how='left')
    
    # 2. Manager request volume - also important  
    manager_request_count = df_combined.groupby('manager_id').size().reset_index(name='manager_request_count')
    df_combined = df_combined.merge(manager_request_count, on='manager_id', how='left')
    
    # 3. Role frequency
    role_freq = df_combined.groupby('role_id').size().reset_index(name='role_count')
    df_combined = df_combined.merge(role_freq, on='role_id', how='left')
    
    # 4. Resource frequency
    resource_freq = df_combined.groupby('resource_id').size().reset_index(name='resource_count')
    df_combined = df_combined.merge(resource_freq, on='resource_id', how='left')
    
    # Extract the engineered features back to the original datasets
    X_train_new['role_resource_count'] = df_combined.iloc[:len(X_train)]['role_resource_count']
    X_train_new['manager_request_count'] = df_combined.iloc[:len(X_train)]['manager_request_count']
    X_train_new['role_count'] = df_combined.iloc[:len(X_train)]['role_count']
    X_train_new['resource_count'] = df_combined.iloc[:len(X_train)]['resource_count']
    X_test_new['role_resource_count'] = df_combined.iloc[len(X_train):]['role_resource_count']
    X_test_new['manager_request_count'] = df_combined.iloc[len(X_train):]['manager_request_count']
    X_test_new['role_count'] = df_combined.iloc[len(X_train):]['role_count']
    X_test_new['resource_count'] = df_combined.iloc[len(X_train):]['resource_count']
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred