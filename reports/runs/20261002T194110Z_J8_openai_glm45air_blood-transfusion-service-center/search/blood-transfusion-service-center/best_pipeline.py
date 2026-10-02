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
    # Apply minimal preprocessing - log transform to handle skewness
    X_train_proc = X_train.copy()
    X_test_proc = X_test.copy()
    
    # Log transform the skewed TotalBloodDonated feature
    X_train_proc['TotalBloodDonated'] = np.log1p(X_train_proc['TotalBloodDonated'])
    X_test_proc['TotalBloodDonated'] = np.log1p(X_test_proc['TotalBloodDonated'])
    
    return X_train_proc, y_train, X_test_proc


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # The successful recent donor feature (2-month threshold)
    X_train_eng['RecentDonor'] = (X_train_eng['MonthsSinceLastDonation'] <= 2).astype(int)
    X_test_eng['RecentDonor'] = (X_test_eng['MonthsSinceLastDonation'] <= 2).astype(int)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred