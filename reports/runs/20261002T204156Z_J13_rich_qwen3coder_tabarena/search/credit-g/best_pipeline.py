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
    # No preprocessing needed for this dataset
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create some engineered features
    # Copy data to avoid modifying originals
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Ratio features
    X_train['credit_amount_duration_ratio'] = X_train['credit_amount'] / (X_train['duration_months'] + 1)
    X_test['credit_amount_duration_ratio'] = X_test['credit_amount'] / (X_test['duration_months'] + 1)
    
    # Age to credit amount ratio
    X_train['age_credit_ratio'] = X_train['age_years'] / (X_train['credit_amount'] + 1)
    X_test['age_credit_ratio'] = X_test['age_years'] / (X_test['credit_amount'] + 1)
    
    # Residence time to age ratio
    X_train['residence_age_ratio'] = X_train['residence_since'] / (X_train['age_years'] + 1)
    X_test['residence_age_ratio'] = X_test['residence_since'] / (X_test['age_years'] + 1)
    
    # Log transformations for skewed features
    X_train['log_credit_amount'] = np.log1p(X_train['credit_amount'])
    X_test['log_credit_amount'] = np.log1p(X_test['credit_amount'])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
