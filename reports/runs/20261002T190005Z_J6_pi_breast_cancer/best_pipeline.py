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
    # Create domain-based features and ratios that are commonly used in medical diagnostics
    # Ratios between worst and mean values for key metrics
    ratio_features = []
    
    # Create ratios for important features
    for metric in ['radius', 'texture', 'perimeter', 'area', 'compactness', 'concavity', 'concave points']:
        if f'worst {metric}' in X_train.columns and f'mean {metric}' in X_train.columns:
            ratio_col = f'worst_to_mean_{metric}_ratio'
            X_train[ratio_col] = X_train[f'worst {metric}'] / (X_train[f'mean {metric}'] + 1e-8)
            X_test[ratio_col] = X_test[f'worst {metric}'] / (X_test[f'mean {metric}'] + 1e-8)
            ratio_features.append(ratio_col)
    
    # Create differences between worst and mean values
    diff_features = []
    for metric in ['radius', 'texture', 'perimeter', 'area', 'smoothness', 'compactness', 'concavity', 'concave points', 'symmetry']:
        if f'worst {metric}' in X_train.columns and f'mean {metric}' in X_train.columns:
            diff_col = f'worst_minus_mean_{metric}'
            X_train[diff_col] = X_train[f'worst {metric}'] - X_train[f'mean {metric}']
            X_test[diff_col] = X_test[f'worst {metric}'] - X_test[f'mean {metric}']
            diff_features.append(diff_col)
    
    # Log transforms for highly skewed features
    log_features = ['mean area', 'worst area', 'mean perimeter', 'worst perimeter']
    for feature in log_features:
        if feature in X_train.columns:
            X_train[f'log_{feature}'] = np.log1p(X_train[feature])
            X_test[f'log_{feature}'] = np.log1p(X_test[feature])
            
    # Add some domain-specific features
    # Compactness can be defined as perimeter^2 / area - 1
    if 'mean perimeter' in X_train.columns and 'mean area' in X_train.columns:
        X_train['mean_perimeter_area_ratio'] = X_train['mean perimeter'] / (X_train['mean area'] + 1e-8)
        X_test['mean_perimeter_area_ratio'] = X_test['mean perimeter'] / (X_test['mean area'] + 1e-8)
        
    # Create interaction terms (product of two features)
    if 'mean concavity' in X_train.columns and 'mean concave points' in X_train.columns:
        X_train['concavity_concave_points_interaction'] = X_train['mean concavity'] * X_train['mean concave points']
        X_test['concavity_concave_points_interaction'] = X_test['mean concavity'] * X_test['mean concave points']
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
