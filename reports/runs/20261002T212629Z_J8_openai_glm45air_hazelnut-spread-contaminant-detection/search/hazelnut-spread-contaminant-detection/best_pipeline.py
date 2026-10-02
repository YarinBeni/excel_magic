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
from sklearn.preprocessing import StandardScaler, PowerTransformer

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Apply Yeo-Johnson transformation to handle potential skewness
    # This is more robust than log transform for data that can be negative
    pt = PowerTransformer(method='yeo-johnson', standardize=True)
    
    # Fit and transform training data
    X_train_transformed = pd.DataFrame(
        pt.fit_transform(X_train), 
        columns=X_train.columns, 
        index=X_train.index
    )
    
    # Transform test data
    X_test_transformed = pd.DataFrame(
        pt.transform(X_test), 
        columns=X_test.columns, 
        index=X_test.index
    )
    
    return X_train_transformed, y_train, X_test_transformed


def engineer(X_train, y_train, X_test):
    # Keep it simple - just return the original features
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred