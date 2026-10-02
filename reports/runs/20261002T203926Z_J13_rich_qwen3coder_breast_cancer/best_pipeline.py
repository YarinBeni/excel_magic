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
from sklearn.preprocessing import StandardScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create additional features based on domain knowledge
    # Ratio features: comparing mean vs worst values
    ratio_features = []
    
    # Create ratios of worst/mean for key features
    mean_cols = [col for col in X_train.columns if 'mean' in col and 'error' not in col]
    worst_cols = [col for col in X_train.columns if 'worst' in col and 'error' not in col]
    
    # Match corresponding mean and worst columns
    for mean_col in mean_cols:
        # Find corresponding worst column
        base_name = mean_col.replace('mean ', '')
        worst_col = f'worst {base_name}'
        
        if worst_col in X_train.columns:
            # Create ratio feature
            ratio_name = f'{base_name}_ratio'
            X_train[ratio_name] = X_train[worst_col] / (X_train[mean_col] + 1e-8)  # Add small epsilon to avoid division by zero
            X_test[ratio_name] = X_test[worst_col] / (X_test[mean_col] + 1e-8)
            
            ratio_features.append(ratio_name)
    
    # Create difference features: worst - mean
    diff_features = []
    for mean_col in mean_cols:
        base_name = mean_col.replace('mean ', '')
        worst_col = f'worst {base_name}'
        
        if worst_col in X_train.columns:
            diff_name = f'{base_name}_diff'
            X_train[diff_name] = X_train[worst_col] - X_train[mean_col]
            X_test[diff_name] = X_test[worst_col] - X_test[mean_col]
            
            diff_features.append(diff_name)
    
    # Create interaction features (product of mean and worst)
    interaction_features = []
    for mean_col in mean_cols[:5]:  # Use first 5 for interaction to avoid too many features
        base_name = mean_col.replace('mean ', '')
        worst_col = f'worst {base_name}'
        
        if worst_col in X_train.columns:
            interaction_name = f'{base_name}_interaction'
            X_train[interaction_name] = X_train[mean_col] * X_train[worst_col]
            X_test[interaction_name] = X_test[mean_col] * X_test[worst_col]
            
            interaction_features.append(interaction_name)
    
    # Create some composite features
    # Area-related features
    X_train['area_ratio'] = X_train['worst area'] / (X_train['mean area'] + 1e-8)
    X_test['area_ratio'] = X_test['worst area'] / (X_test['mean area'] + 1e-8)
    
    X_train['area_diff'] = X_train['worst area'] - X_train['mean area']
    X_test['area_diff'] = X_test['worst area'] - X_test['mean area']
    
    # Perimeter-related features  
    X_train['perimeter_ratio'] = X_train['worst perimeter'] / (X_train['mean perimeter'] + 1e-8)
    X_test['perimeter_ratio'] = X_test['worst perimeter'] / (X_test['mean perimeter'] + 1e-8)
    
    X_train['perimeter_diff'] = X_train['worst perimeter'] - X_train['mean perimeter']
    X_test['perimeter_diff'] = X_test['worst perimeter'] - X_test['mean perimeter']
    
    # Concavity-related features
    X_train['concavity_ratio'] = X_train['worst concavity'] / (X_train['mean concavity'] + 1e-8)
    X_test['concavity_ratio'] = X_test['worst concavity'] / (X_test['mean concavity'] + 1e-8)
    
    X_train['concavity_diff'] = X_train['worst concavity'] - X_train['mean concavity']
    X_test['concavity_diff'] = X_test['worst concavity'] - X_test['mean concavity']
    
    # Combine multiple features into compound features
    # Sum of mean features
    mean_features = [col for col in X_train.columns if 'mean ' in col and 'error' not in col]
    X_train['mean_sum'] = X_train[mean_features].sum(axis=1)
    X_test['mean_sum'] = X_test[mean_features].sum(axis=1)
    
    # Sum of worst features
    worst_features = [col for col in X_train.columns if 'worst ' in col and 'error' not in col]
    X_train['worst_sum'] = X_train[worst_features].sum(axis=1)
    X_test['worst_sum'] = X_test[worst_features].sum(axis=1)
    
    # Ratio of sums
    X_train['sum_ratio'] = X_train['worst_sum'] / (X_train['mean_sum'] + 1e-8)
    X_test['sum_ratio'] = X_test['worst_sum'] / (X_test['mean_sum'] + 1e-8)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
