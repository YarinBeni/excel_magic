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
    # Handle any potential missing values
    X_train = X_train.fillna(0)
    X_test = X_test.fillna(0)
    
    # Scale the features to have zero mean and unit variance
    scaler = StandardScaler()
    X_train_scaled = pd.DataFrame(scaler.fit_transform(X_train), columns=X_train.columns, index=X_train.index)
    X_test_scaled = pd.DataFrame(scaler.transform(X_test), columns=X_test.columns, index=X_test.index)
    
    return X_train_scaled, y_train, X_test_scaled


def engineer(X_train, y_train, X_test):
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Add the two features that worked well
    X_train_eng['total_energy'] = (X_train ** 2).sum(axis=1)
    X_test_eng['total_energy'] = (X_test ** 2).sum(axis=1)
    
    X_train_eng['global_std'] = X_train.std(axis=1)
    X_test_eng['global_std'] = X_test.std(axis=1)
    
    # Try adding one more simple feature: the median absolute deviation
    X_train_eng['global_mad'] = X_train.subtract(X_train.median(axis=1), axis=0).abs().median(axis=1)
    X_test_eng['global_mad'] = X_test.subtract(X_test.median(axis=1), axis=0).abs().median(axis=1)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred