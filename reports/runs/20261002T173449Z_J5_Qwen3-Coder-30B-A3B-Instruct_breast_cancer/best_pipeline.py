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
    # Create ratios and derived features
    # Ratio of worst to mean features
    ratio_features = []
    for i in range(10):
        mean_col = f'mean {["radius", "texture", "perimeter", "area", "smoothness", "compactness", "concavity", "concave points", "symmetry", "fractal dimension"][i]}'
        worst_col = f'worst {["radius", "texture", "perimeter", "area", "smoothness", "compactness", "concavity", "concave points", "symmetry", "fractal dimension"][i]}'
        
        if mean_col in X_train.columns and worst_col in X_train.columns:
            ratio_name = f'{mean_col.split()[-1]}_ratio'
            X_train[ratio_name] = X_train[worst_col] / (X_train[mean_col] + 1e-8)
            X_test[ratio_name] = X_test[worst_col] / (X_test[mean_col] + 1e-8)
            ratio_features.append(ratio_name)
    
    # Create some interaction terms
    # Area and perimeter ratios
    if 'mean area' in X_train.columns and 'mean perimeter' in X_train.columns:
        X_train['area_perimeter_ratio'] = X_train['mean area'] / (X_train['mean perimeter'] + 1e-8)
        X_test['area_perimeter_ratio'] = X_test['mean area'] / (X_test['mean perimeter'] + 1e-8)
    
    # Concavity and concave points interaction
    if 'mean concavity' in X_train.columns and 'mean concave points' in X_train.columns:
        X_train['concavity_concave_points_interaction'] = X_train['mean concavity'] * X_train['mean concave points']
        X_test['concavity_concave_points_interaction'] = X_test['mean concavity'] * X_test['mean concave points']
    
    # Error ratios
    if 'radius error' in X_train.columns and 'mean radius' in X_train.columns:
        X_train['radius_error_ratio'] = X_train['radius error'] / (X_train['mean radius'] + 1e-8)
        X_test['radius_error_ratio'] = X_test['radius error'] / (X_test['mean radius'] + 1e-8)
    
    # Additional ratios based on error features
    if 'texture error' in X_train.columns and 'mean texture' in X_train.columns:
        X_train['texture_error_ratio'] = X_train['texture error'] / (X_train['mean texture'] + 1e-8)
        X_test['texture_error_ratio'] = X_test['texture error'] / (X_test['mean texture'] + 1e-8)
    
    # Log transformations for skewed features
    skewed_features = ['mean area', 'mean perimeter', 'mean radius', 'mean concavity', 'mean concave points']
    for feat in skewed_features:
        if feat in X_train.columns:
            X_train[f'{feat}_log'] = np.log1p(X_train[feat])
            X_test[f'{feat}_log'] = np.log1p(X_test[feat])
    
    # Difference between worst and mean features
    diff_features = []
    for i in range(10):
        mean_col = f'mean {["radius", "texture", "perimeter", "area", "smoothness", "compactness", "concavity", "concave points", "symmetry", "fractal dimension"][i]}'
        worst_col = f'worst {["radius", "texture", "perimeter", "area", "smoothness", "compactness", "concavity", "concave points", "symmetry", "fractal dimension"][i]}'
        
        if mean_col in X_train.columns and worst_col in X_train.columns:
            diff_name = f'{mean_col.split()[-1]}_diff'
            X_train[diff_name] = X_train[worst_col] - X_train[mean_col]
            X_test[diff_name] = X_test[worst_col] - X_test[mean_col]
            diff_features.append(diff_name)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred