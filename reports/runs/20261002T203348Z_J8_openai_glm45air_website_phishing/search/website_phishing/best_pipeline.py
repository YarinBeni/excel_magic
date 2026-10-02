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
from sklearn.preprocessing import OneHotEncoder, OrdinalEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Convert categorical columns to string type
    categorical_cols = X_train.columns.tolist()
    X_train_str = X_train[categorical_cols].astype(str)
    X_test_str = X_test[categorical_cols].astype(str)
    
    return X_train_str, y_train, X_test_str


def engineer(X_train, y_train, X_test):
    # Create both ordinal and one-hot encoded features
    categorical_cols = X_train.columns.tolist()
    
    # 1. Ordinal encoding (captures ordinal relationships if they exist)
    ordinal_encoder = OrdinalEncoder(handle_unknown='use_encoded_value', unknown_value=-1)
    X_train_ordinal = pd.DataFrame(
        ordinal_encoder.fit_transform(X_train[categorical_cols]),
        columns=[f"{col}_ordinal" for col in categorical_cols],
        index=X_train.index
    )
    X_test_ordinal = pd.DataFrame(
        ordinal_encoder.transform(X_test[categorical_cols]),
        columns=[f"{col}_ordinal" for col in categorical_cols],
        index=X_test.index
    )
    
    # 2. One-hot encoding for features with 4 or fewer unique values
    onehot_cols = []
    for col in categorical_cols:
        n_unique = X_train[col].nunique()
        if n_unique <= 4:  # Expanded to include features with 4 unique values
            onehot_cols.append(col)
    
    if onehot_cols:
        onehot_encoder = OneHotEncoder(sparse_output=False, handle_unknown='ignore')
        X_train_onehot = onehot_encoder.fit_transform(X_train[onehot_cols])
        X_test_onehot = onehot_encoder.transform(X_test[onehot_cols])
        
        # Create feature names
        feature_names = []
        for i, col in enumerate(onehot_cols):
            for category in onehot_encoder.categories_[i]:
                feature_names.append(f"{col}_{category}")
        
        # Convert to DataFrame
        X_train_onehot_df = pd.DataFrame(X_train_onehot, columns=feature_names, index=X_train.index)
        X_test_onehot_df = pd.DataFrame(X_test_onehot, columns=feature_names, index=X_test.index)
        
        # Combine ordinal and one-hot features
        X_train_combined = pd.concat([X_train_ordinal, X_train_onehot_df], axis=1)
        X_test_combined = pd.concat([X_test_ordinal, X_test_onehot_df], axis=1)
    else:
        # If no one-hot encoding, just use ordinal
        X_train_combined = X_train_ordinal
        X_test_combined = X_test_ordinal
    
    return X_train_combined, X_test_combined


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred