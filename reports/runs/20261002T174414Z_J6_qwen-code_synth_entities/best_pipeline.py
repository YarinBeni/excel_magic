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
    # Create a combined dataframe for easier manipulation
    X_train_combined = X_train.copy()
    X_test_combined = X_test.copy()
    
    # Add the target variable to the training set for groupby operations
    X_train_combined['approved'] = y_train
    
    # 1. Create interaction features between role and resource
    X_train_combined['role_resource_interaction'] = X_train_combined['role_id'] * X_train_combined['resource_id']
    X_test_combined['role_resource_interaction'] = X_test_combined['role_id'] * X_test_combined['resource_id']
    
    # 2. Create a ratio feature (this might capture the relationship better)
    # For role-resource interaction, we can normalize by max values
    X_train_combined['role_resource_ratio'] = X_train_combined['role_id'] / (X_train_combined['resource_id'] + 1)  # +1 to avoid division by zero
    X_test_combined['role_resource_ratio'] = X_test_combined['role_id'] / (X_test_combined['resource_id'] + 1)
    
    # 3. Create a difference feature
    X_train_combined['role_resource_diff'] = np.abs(X_train_combined['role_id'] - X_train_combined['resource_id'])
    X_test_combined['role_resource_diff'] = np.abs(X_test_combined['role_id'] - X_test_combined['resource_id'])
    
    # 4. Frequency of manager requests (as before)
    manager_counts = X_train_combined.groupby('manager_id').size().reset_index(name='manager_request_count')
    X_train_combined = X_train_combined.merge(manager_counts, on='manager_id', how='left')
    X_test_combined = X_test_combined.merge(manager_counts, on='manager_id', how='left')
    
    # 5. Create a normalized manager request count
    X_train_combined['norm_manager_request_count'] = X_train_combined['manager_request_count'] / (len(X_train_combined) + 1)
    X_test_combined['norm_manager_request_count'] = X_test_combined['manager_request_count'] / (len(X_train_combined) + 1)
    
    # Remove the target variable from the training set
    y_train_new = X_train_combined.pop('approved')
    
    # Return updated dataframes
    return X_train_combined, X_test_combined


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
