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
    # Try log transform of target variable to reduce skewness
    y_train_log = np.log(y_train + 1e-8)  # Add small epsilon to avoid log(0)
    return X_train, y_train_log, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying originals
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Simple feature combinations that might be physically meaningful
    # Ratio of chord length to suction side displacement thickness
    X_train_eng['chord_to_thickness_ratio'] = X_train_eng['chord-length'] / (X_train_eng['suction-side-displacement-thickness'] + 1e-8)
    X_test_eng['chord_to_thickness_ratio'] = X_test_eng['chord-length'] / (X_test_eng['suction-side-displacement-thickness'] + 1e-8)
    
    # Product of velocity and frequency (might relate to sound generation)
    X_train_eng['velocity_frequency_product'] = X_train_eng['free-stream-velocity'] * X_train_eng['frequency']
    X_test_eng['velocity_frequency_product'] = X_test_eng['free-stream-velocity'] * X_test_eng['frequency']
    
    # Log transform of frequency (it has a wide range)
    X_train_eng['log_frequency'] = np.log1p(X_train_eng['frequency'])
    X_test_eng['log_frequency'] = np.log1p(X_test_eng['frequency'])
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert log transform
    return np.exp(pred) - 1e-8
