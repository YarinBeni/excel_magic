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
    # Create frequency-based features
    # Combine role and resource to get (role, resource) pair frequencies
    all_data = pd.concat([X_train, X_test], ignore_index=True)
    
    # Count occurrences of (role, resource) pairs
    role_resource_counts = all_data.groupby(['role_id', 'resource_id']).size().reset_index(name='role_resource_count')
    all_data = all_data.merge(role_resource_counts, on=['role_id', 'resource_id'], how='left')
    
    # Count occurrences of managers
    manager_counts = all_data.groupby('manager_id').size().reset_index(name='manager_request_count')
    all_data = all_data.merge(manager_counts, on='manager_id', how='left')
    
    # Count occurrences of roles
    role_counts = all_data.groupby('role_id').size().reset_index(name='role_count')
    all_data = all_data.merge(role_counts, on='role_id', how='left')
    
    # Count occurrences of resources
    resource_counts = all_data.groupby('resource_id').size().reset_index(name='resource_count')
    all_data = all_data.merge(resource_counts, on='resource_id', how='left')
    
    # Create ratio features
    all_data['role_resource_ratio'] = all_data['role_resource_count'] / (all_data['role_count'] + 1)
    all_data['manager_resource_ratio'] = all_data['manager_request_count'] / (all_data['resource_count'] + 1)
    
    # Create additional interaction features
    all_data['manager_role_interaction'] = all_data['manager_id'] * all_data['role_id']
    all_data['manager_resource_interaction'] = all_data['manager_id'] * all_data['resource_id']
    
    # Add normalized versions of counts
    all_data['normalized_role_resource_count'] = all_data['role_resource_count'] / (all_data['role_resource_count'].max() + 1)
    all_data['normalized_manager_request_count'] = all_data['manager_request_count'] / (all_data['manager_request_count'].max() + 1)
    
    # Create log-transformed features to handle potential skewness
    all_data['log_role_resource_count'] = np.log1p(all_data['role_resource_count'])
    all_data['log_manager_request_count'] = np.log1p(all_data['manager_request_count'])
    
    # Create binarized features based on thresholds
    all_data['high_role_resource'] = (all_data['role_resource_count'] > all_data['role_resource_count'].median()).astype(int)
    all_data['high_manager_requests'] = (all_data['manager_request_count'] > all_data['manager_request_count'].median()).astype(int)
    
    # Extract features for training and test sets
    X_train_processed = all_data.iloc[:len(X_train)].copy()
    X_test_processed = all_data.iloc[len(X_train):].copy()
    
    # Remove request_id as it's likely unique and not informative
    X_train_processed = X_train_processed.drop(columns=['request_id'])
    X_test_processed = X_test_processed.drop(columns=['request_id'])
    
    return X_train_processed, X_test_processed


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred