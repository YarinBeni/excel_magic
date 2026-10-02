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
    # No preprocessing needed for this dataset
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create engineered features based on domain knowledge and statistical relationships
    
    # Combine mean and worst features to capture both average and extreme characteristics
    mean_features = [col for col in X_train.columns if 'mean' in col]
    worst_features = [col for col in X_train.columns if 'worst' in col]
    
    # Create ratios between mean and worst features
    ratio_cols = []
    for i, (mean_col, worst_col) in enumerate(zip(mean_features, worst_features)):
        ratio_name = f"{mean_col}_to_{worst_col}".replace(' ', '_')
        X_train[ratio_name] = X_train[mean_col] / (X_train[worst_col] + 1e-8)  # Adding small value to avoid division by zero
        X_test[ratio_name] = X_test[mean_col] / (X_test[worst_col] + 1e-8)
        ratio_cols.append(ratio_name)
    
    # Create difference features between mean and worst
    diff_cols = []
    for i, (mean_col, worst_col) in enumerate(zip(mean_features, worst_features)):
        diff_name = f"{mean_col}_diff_{worst_col}".replace(' ', '_')
        X_train[diff_name] = X_train[mean_col] - X_train[worst_col]
        X_test[diff_name] = X_test[mean_col] - X_test[worst_col]
        diff_cols.append(diff_name)
    
    # Create compound features that might capture relationships in the data
    # Area and perimeter relationships (both are measures of size)
    area_perimeter_ratio = 'area_perimeter_ratio'
    X_train[area_perimeter_ratio] = X_train['mean area'] / (X_train['mean perimeter'] + 1e-8)
    X_test[area_perimeter_ratio] = X_test['mean area'] / (X_test['mean perimeter'] + 1e-8)
    
    # Compactness and concavity relationships
    compactness_concavity_ratio = 'compactness_concavity_ratio'
    X_train[compactness_concavity_ratio] = X_train['mean compactness'] / (X_train['mean concavity'] + 1e-8)
    X_test[compactness_concavity_ratio] = X_test['mean compactness'] / (X_test['mean concavity'] + 1e-8)
    
    # Create features representing the spread of measurements
    # Difference between worst and mean for key features
    spread_features = []
    for col in ['mean radius', 'mean texture', 'mean perimeter', 'mean area']:
        spread_col = f"{col.replace('mean ', '')}_spread"
        X_train[spread_col] = X_train[f'worst {col.replace("mean ", "")}'] - X_train[col]
        X_test[spread_col] = X_test[f'worst {col.replace("mean ", "")}'] - X_test[col]
        spread_features.append(spread_col)
    
    # Create log-transformed features for skewed distributions
    skewed_features = ['mean area', 'mean perimeter', 'mean radius']
    log_features = []
    for feat in skewed_features:
        log_feat = f"log_{feat}"
        X_train[log_feat] = np.log1p(X_train[feat])
        X_test[log_feat] = np.log1p(X_test[feat])
        log_features.append(log_feat)
        
    # Combine features that might be correlated in a meaningful way
    # Create a "complexity" feature based on multiple concavity and concave points metrics
    complexity_feature = 'complexity_score'
    X_train[complexity_feature] = (X_train['mean concavity'] * X_train['mean concave points']) / (X_train['mean compactness'] + 1e-8)
    X_test[complexity_feature] = (X_test['mean concavity'] * X_test['mean concave points']) / (X_test['mean compactness'] + 1e-8)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
