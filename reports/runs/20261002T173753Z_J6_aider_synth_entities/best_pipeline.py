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
from sklearn.preprocessing import LabelEncoder
from sklearn.calibration import CalibratedClassifierCV

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # No preprocessing needed for this task
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create feature engineering based on the dataset description
    # The approval depends on (role, resource) pair frequency and manager request count
    
    # Combine training and test data temporarily for consistent encoding
    X_all = pd.concat([X_train, X_test], ignore_index=True)
    
    # Create frequency features for role-resource pairs
    role_resource_counts = X_all.groupby(['role_id', 'resource_id']).size().reset_index(name='role_resource_count')
    X_all = X_all.merge(role_resource_counts, on=['role_id', 'resource_id'], how='left')
    
    # Create frequency features for managers
    manager_counts = X_all.groupby('manager_id').size().reset_index(name='manager_request_count')
    X_all = X_all.merge(manager_counts, on='manager_id', how='left')
    
    # Create frequency features for roles
    role_counts = X_all.groupby('role_id').size().reset_index(name='role_count')
    X_all = X_all.merge(role_counts, on='role_id', how='left')
    
    # Create frequency features for resources
    resource_counts = X_all.groupby('resource_id').size().reset_index(name='resource_count')
    X_all = X_all.merge(resource_counts, on='resource_id', how='left')
    
    # Create interaction features
    X_all['role_resource_manager'] = X_all['role_id'].astype(str) + '_' + \
                                     X_all['resource_id'].astype(str) + '_' + \
                                     X_all['manager_id'].astype(str)
    
    # Encode the interaction feature
    le = LabelEncoder()
    X_all['role_resource_manager_encoded'] = le.fit_transform(X_all['role_resource_manager'])
    
    # Create ratio features
    X_all['role_resource_ratio'] = X_all['role_resource_count'] / (X_all['role_count'] + 1)
    X_all['resource_manager_ratio'] = X_all['role_resource_count'] / (X_all['manager_request_count'] + 1)
    
    # Create binary features indicating high-frequency combinations
    threshold = X_all['role_resource_count'].quantile(0.75)
    X_all['high_freq_role_resource'] = (X_all['role_resource_count'] >= threshold).astype(int)
    
    # Add some additional features that might help
    # Log transformed counts to handle skewness
    X_all['log_role_resource_count'] = np.log1p(X_all['role_resource_count'])
    X_all['log_manager_request_count'] = np.log1p(X_all['manager_request_count'])
    
    # Create more interaction features
    X_all['role_resource_interaction'] = X_all['role_id'] * X_all['resource_id']
    X_all['role_manager_interaction'] = X_all['role_id'] * X_all['manager_id']
    X_all['resource_manager_interaction'] = X_all['resource_id'] * X_all['manager_id']
    
    # Normalize some features
    X_all['normalized_role_resource_count'] = X_all['role_resource_count'] / (X_all['resource_count'] + 1)
    X_all['normalized_manager_request_count'] = X_all['manager_request_count'] / (X_all['role_count'] + 1)
    
    # Split back into train and test
    X_train_out = X_all.iloc[:len(X_train)].copy()
    X_test_out = X_all.iloc[len(X_train):].copy()
    
    return X_train_out, X_test_out


def sample(X_train, y_train, X_test, max_rows):
    # Create stratified sampling to ensure balanced representation
    # Since we're dealing with a binary classification, we want to maintain class balance
    n_samples = min(max_rows, len(X_train))
    
    # Create multiple context views with different stratifications
    # This helps the model learn from diverse perspectives
    
    # First view: all data
    view1 = np.arange(len(X_train))
    
    # Second view: stratified sampling to balance classes
    pos_indices = np.where(y_train == 1)[0]
    neg_indices = np.where(y_train == 0)[0]
    
    # Sample equal numbers from both classes
    n_pos = len(pos_indices)
    n_neg = len(neg_indices)
    n_sample_per_class = min(n_pos, n_neg, n_samples // 2)
    
    if n_sample_per_class > 0:
        sampled_pos = np.random.choice(pos_indices, size=n_sample_per_class, replace=False)
        sampled_neg = np.random.choice(neg_indices, size=n_sample_per_class, replace=False)
        view2 = np.concatenate([sampled_pos, sampled_neg])
    else:
        view2 = np.arange(len(X_train))
    
    # Third view: random sampling
    view3 = np.random.choice(len(X_train), size=min(n_samples, len(X_train)), replace=False)
    
    # Fourth view: oversample minority class
    if len(pos_indices) < len(neg_indices):
        # Oversample positive class
        oversampled_pos = np.random.choice(pos_indices, size=len(neg_indices), replace=True)
        view4 = np.concatenate([oversampled_pos, neg_indices])
    else:
        # Oversample negative class
        oversampled_neg = np.random.choice(neg_indices, size=len(pos_indices), replace=True)
        view4 = np.concatenate([pos_indices, oversampled_neg])
    
    return [view1, view2, view3, view4]


def postprocess(pred, y_train_orig, X_test):
    # Apply calibration to improve probability estimates
    # For binary classification, we can use isotonic regression or Platt scaling
    if isinstance(pred, pd.DataFrame):
        # If pred is a DataFrame, get the probability of positive class (assuming it's the second column)
        pred_probs = pred.iloc[:, 1] if len(pred.columns) > 1 else pred.iloc[:, 0]
    else:
        # If pred is a Series, use it directly
        pred_probs = pred
    
    # Apply sigmoid calibration to improve probability estimates
    # This helps reduce extreme probability values that might hurt AUROC
    pred_probs = 1 / (1 + np.exp(-5 * (pred_probs - 0.5)))
    
    # Return calibrated predictions
    if isinstance(pred, pd.DataFrame):
        result = pred.copy()
        result.iloc[:, 1] = pred_probs
        return result
    else:
        return pred_probs
