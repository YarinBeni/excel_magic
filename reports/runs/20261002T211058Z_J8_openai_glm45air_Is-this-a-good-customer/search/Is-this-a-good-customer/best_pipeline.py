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
    # Simple preprocessing: ensure no missing values and basic cleanup
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Handle any potential infinite values in debt-to-income ratio calculation
    X_train['income'] = X_train['income'].replace(0, 1)  # Avoid division by zero
    X_test['income'] = X_test['income'].replace(0, 1)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Key credit risk feature with zero protection
    X_train['debt_to_income_ratio'] = X_train['credit_amount'] / (X_train['income'] + 1)
    X_test['debt_to_income_ratio'] = X_test['credit_amount'] / (X_test['income'] + 1)
    
    # Add a simple binary feature for high credit amount
    credit_threshold = X_train['credit_amount'].median()
    X_train['high_credit'] = (X_train['credit_amount'] > credit_threshold).astype(int)
    X_test['high_credit'] = (X_test['credit_amount'] > credit_threshold).astype(int)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    # Create multiple context views to improve robustness
    views = []
    
    # View 1: All data
    views.append(np.arange(len(X_train)))
    
    # View 2: Random subset (70% of data)
    subset_size = int(0.7 * len(X_train))
    subset_indices = np.random.choice(len(X_train), subset_size, replace=False)
    views.append(subset_indices)
    
    return views


def postprocess(pred, y_train_orig, X_test):
    return pred