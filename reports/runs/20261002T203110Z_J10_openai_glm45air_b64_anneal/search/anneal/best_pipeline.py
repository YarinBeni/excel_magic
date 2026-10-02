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
    # Handle "not_applicable" properly for categorical columns
    for col in X_train.columns:
        if X_train[col].dtype.name == 'category':
            # For categorical columns, add 'missing' to categories if not present
            if 'missing' not in X_train[col].cat.categories:
                X_train[col] = X_train[col].cat.add_categories(['missing'])
                X_test[col] = X_test[col].cat.add_categories(['missing'])
            X_train[col] = X_train[col].fillna('missing')
            X_test[col] = X_test[col].fillna('missing')
        else:
            # For non-categorical columns, replace with NaN
            X_train[col] = X_train[col].replace('not_applicable', np.nan)
            X_test[col] = X_test[col].replace('not_applicable', np.nan)
    
    # Apply simple scaling to numerical features to help the model
    numerical_cols = ['carbon', 'hardness', 'strength', 'thick', 'width', 'len']
    numerical_cols = [col for col in numerical_cols if col in X_train.columns]
    
    if numerical_cols:
        scaler = StandardScaler()
        X_train[numerical_cols] = scaler.fit_transform(X_train[numerical_cols])
        X_test[numerical_cols] = scaler.transform(X_test[numerical_cols])
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create one simple but potentially useful feature
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Create a simple size feature using fixed bins
    if 'width' in X_train_eng.columns and 'len' in X_train_eng.columns:
        # Simple size category using fixed bins
        size_train = X_train_eng['width'] * X_train_eng['len']
        size_test = X_test_eng['width'] * X_test_eng['len']
        
        # Use fixed bins based on domain knowledge
        bins = [0, 500000, 2000000, np.inf]
        labels = ['small', 'medium', 'large']
        
        X_train_eng['size_category'] = pd.cut(size_train, bins=bins, labels=labels)
        X_test_eng['size_category'] = pd.cut(size_test, bins=bins, labels=labels)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    # Keep it simple with just one view
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    # Return predictions as-is
    return pred