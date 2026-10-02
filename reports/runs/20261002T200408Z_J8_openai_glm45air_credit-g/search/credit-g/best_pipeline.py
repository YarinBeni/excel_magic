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
    # Keep exactly as original
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Add just the one feature that worked best: debt-to-age proxy
    # This captures the credit burden relative to borrower's age/experience
    X_train['debt_to_age_proxy'] = X_train['credit_amount'] / (X_train['age_years'] + 1)
    X_test['debt_to_age_proxy'] = X_test['credit_amount'] / (X_test['age_years'] + 1)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    # Keep exactly as original
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred