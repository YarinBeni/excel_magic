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
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Engineer features based on domain knowledge
    
    # 1. Manager request count (how many requests each manager has)
    manager_counts = pd.concat([X_train_eng[['manager_id']], X_test_eng[['manager_id']]]).groupby('manager_id').size()
    X_train_eng['manager_request_count'] = X_train_eng['manager_id'].map(manager_counts)
    X_test_eng['manager_request_count'] = X_test_eng['manager_id'].map(manager_counts)
    
    # 2. Role-Resource pair frequency (how often each role-resource combination occurs)
    role_resource_pairs = pd.concat([
        X_train_eng[['role_id', 'resource_id']].assign(temp=1),
        X_test_eng[['role_id', 'resource_id']].assign(temp=1)
    ]).groupby(['role_id', 'resource_id'])['temp'].sum()
    
    X_train_eng['role_resource_frequency'] = X_train_eng.set_index(['role_id', 'resource_id']).index.map(role_resource_pairs)
    X_test_eng['role_resource_frequency'] = X_test_eng.set_index(['role_id', 'resource_id']).index.map(role_resource_pairs)
    
    # 3. Resource request count (how many requests each resource has)
    resource_counts = pd.concat([X_train_eng[['resource_id']], X_test_eng[['resource_id']]]).groupby('resource_id').size()
    X_train_eng['resource_request_count'] = X_train_eng['resource_id'].map(resource_counts)
    X_test_eng['resource_request_count'] = X_test_eng['resource_id'].map(resource_counts)
    
    # 4. Role request count (how many requests each role has)
    role_counts = pd.concat([X_train_eng[['role_id']], X_test_eng[['role_id']]]).groupby('role_id').size()
    X_train_eng['role_request_count'] = X_train_eng['role_id'].map(role_counts)
    X_test_eng['role_request_count'] = X_test_eng['role_id'].map(role_counts)
    
    # 5. Manager-Resource interaction frequency
    manager_resource_pairs = pd.concat([
        X_train_eng[['manager_id', 'resource_id']].assign(temp=1),
        X_test_eng[['manager_id', 'resource_id']].assign(temp=1)
    ]).groupby(['manager_id', 'resource_id'])['temp'].sum()
    
    X_train_eng['manager_resource_frequency'] = X_train_eng.set_index(['manager_id', 'resource_id']).index.map(manager_resource_pairs)
    X_test_eng['manager_resource_frequency'] = X_test_eng.set_index(['manager_id', 'resource_id']).index.map(manager_resource_pairs)
    
    # 6. Manager-Role interaction frequency
    manager_role_pairs = pd.concat([
        X_train_eng[['manager_id', 'role_id']].assign(temp=1),
        X_test_eng[['manager_id', 'role_id']].assign(temp=1)
    ]).groupby(['manager_id', 'role_id'])['temp'].sum()
    
    X_train_eng['manager_role_frequency'] = X_train_eng.set_index(['manager_id', 'role_id']).index.map(manager_role_pairs)
    X_test_eng['manager_role_frequency'] = X_test_eng.set_index(['manager_id', 'role_id']).index.map(manager_role_pairs)
    
    # 7. Ratios and derived features
    X_train_eng['manager_role_ratio'] = X_train_eng['manager_request_count'] / (X_train_eng['role_request_count'] + 1)
    X_test_eng['manager_role_ratio'] = X_test_eng['manager_request_count'] / (X_test_eng['role_request_count'] + 1)
    
    X_train_eng['resource_role_ratio'] = X_train_eng['resource_request_count'] / (X_train_eng['role_request_count'] + 1)
    X_test_eng['resource_role_ratio'] = X_test_eng['resource_request_count'] / (X_test_eng['role_request_count'] + 1)
    
    # 8. Binary indicators for rare/common combinations
    rare_threshold = role_resource_pairs.quantile(0.25)
    X_train_eng['rare_role_resource'] = (X_train_eng['role_resource_frequency'] <= rare_threshold).astype(int)
    X_test_eng['rare_role_resource'] = (X_test_eng['role_resource_frequency'] <= rare_threshold).astype(int)
    
    high_freq_threshold = manager_counts.quantile(0.75)
    X_train_eng['high_freq_manager'] = (X_train_eng['manager_request_count'] >= high_freq_threshold).astype(int)
    X_test_eng['high_freq_manager'] = (X_test_eng['manager_request_count'] >= high_freq_threshold).astype(int)
    
    # 9. Request ID frequency (how often each specific request appears)
    request_counts = pd.concat([X_train_eng[['request_id']], X_test_eng[['request_id']]]).groupby('request_id').size()
    X_train_eng['request_frequency'] = X_train_eng['request_id'].map(request_counts)
    X_test_eng['request_frequency'] = X_test_eng['request_id'].map(request_counts)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    # Try a different sampling strategy - stratified sampling by target
    n_samples = len(X_train)
    
    # Get indices for positive and negative classes
    positive_indices = y_train[y_train == 1].index
    negative_indices = y_train[y_train == 0].index
    
    # Sample from each class to maintain balance
    np.random.seed(42)  # for reproducibility
    n_positive = min(len(positive_indices), max_rows // 4)  # use 25% for positive class
    n_negative = min(max_rows - n_positive, len(negative_indices))
    
    sampled_positive = np.random.choice(positive_indices, n_positive, replace=False)
    sampled_negative = np.random.choice(negative_indices, n_negative, replace=False)
    
    return [np.concatenate([sampled_positive, sampled_negative])]


def postprocess(pred, y_train_orig, X_test):
    return pred