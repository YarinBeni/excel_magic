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

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Simple preprocessing: fill missing values and encode categoricals
    # Fill missing values for categorical variables
    categorical_columns = X_train.select_dtypes(include=['object']).columns
    
    for col in categorical_columns:
        if col in X_train.columns:
            X_train[col] = X_train[col].fillna('missing')
            X_test[col] = X_test[col].fillna('missing')
    
    # For numeric columns, fill with median
    numeric_columns = X_train.select_dtypes(include=[np.number]).columns
    for col in numeric_columns:
        if col in X_train.columns:
            median_val = X_train[col].median()
            X_train[col] = X_train[col].fillna(median_val)
            X_test[col] = X_test[col].fillna(median_val)
    
    # Encode categorical variables as integers
    for col in categorical_columns:
        if col in X_train.columns:
            le = LabelEncoder()
            # Fit on combined data to avoid unseen categories
            combined_data = pd.concat([X_train[col], X_test[col]], ignore_index=True)
            le.fit(combined_data)
            X_train[col] = le.transform(X_train[col])
            X_test[col] = le.transform(X_test[col])
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create only the most promising features
    
    # Ratio features
    if 'carbon' in X_train.columns and 'hardness' in X_train.columns:
        X_train['carbon_hardness_ratio'] = X_train['carbon'] / (X_train['hardness'] + 1e-8)
        X_test['carbon_hardness_ratio'] = X_test['carbon'] / (X_test['hardness'] + 1e-8)
    
    # Log transforms for some skewed features
    # Only keep the most informative ones
    if 'carbon' in X_train.columns:
        X_train['carbon_log'] = np.log1p(X_train['carbon'])
        X_test['carbon_log'] = np.log1p(X_test['carbon'])
    
    if 'hardness' in X_train.columns:
        X_train['hardness_log'] = np.log1p(X_train['hardness'])
        X_test['hardness_log'] = np.log1p(X_test['hardness'])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
