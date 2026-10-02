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
    # Create a simple engineered feature based on physical relationships
    # Based on the description: "depends on the Strouhal-like ratio frequency*chord/velocity, 
    # scaled by thickness, plus an angle term"
    # Let's create a feature that combines frequency, chord, and velocity (Strouhal-like ratio)
    X_train['strouhal_like'] = X_train['frequency_hz'] * X_train['chord_length_m'] / (X_train['free_stream_velocity_m_s'] + 1e-10)
    X_test['strouhal_like'] = X_test['frequency_hz'] * X_test['chord_length_m'] / (X_test['free_stream_velocity_m_s'] + 1e-10)
    
    # Ensure no infinities or NaNs
    X_train['strouhal_like'] = X_train['strouhal_like'].fillna(0)
    X_test['strouhal_like'] = X_test['strouhal_like'].fillna(0)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred
