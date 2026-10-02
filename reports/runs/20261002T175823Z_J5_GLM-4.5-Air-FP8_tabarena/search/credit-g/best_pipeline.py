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
    # Target encoding for categorical variables
    X_train_processed = X_train.copy()
    X_test_processed = X_test.copy()
    
    # Get categorical columns
    categorical_cols = X_train.select_dtypes(include=['category']).columns
    
    # Target encoding with smoothing
    for col in categorical_cols:
        # Calculate target mean for each category with smoothing
        smoothing = 10  # smoothing parameter
        target_mean = y_train.mean()
        
        # Calculate category statistics
        cat_stats = pd.DataFrame({'category': X_train[col], 'target': y_train}).groupby('category')['target'].agg(['count', 'mean'])
        
        # Apply smoothing
        cat_stats['smooth_mean'] = (cat_stats['count'] * cat_stats['mean'] + smoothing * target_mean) / (cat_stats['count'] + smoothing)
        
        # Convert column to numeric type first, then apply target encoding
        X_train_processed[col] = X_train[col].astype(str)
        X_test_processed[col] = X_test[col].astype(str)
        
        # Map to train and test
        X_train_processed[col] = X_train_processed[col].map(cat_stats['smooth_mean']).fillna(target_mean)
        X_test_processed[col] = X_test_processed[col].map(cat_stats['smooth_mean']).fillna(target_mean)
    
    return X_train_processed, y_train, X_test_processed


def engineer(X_train, y_train, X_test):
    X_train_engineered = X_train.copy()
    X_test_engineered = X_test.copy()
    
    # Log transform credit_amount (most promising given the skew)
    X_train_engineered['credit_amount_log'] = np.log1p(X_train['credit_amount'])
    X_test_engineered['credit_amount_log'] = np.log1p(X_test['credit_amount'])
    
    # Simple credit affordability ratio
    X_train_engineered['credit_per_month'] = X_train['credit_amount'] / (X_train['duration_months'] + 1)
    X_test_engineered['credit_per_month'] = X_test['credit_amount'] / (X_test['duration_months'] + 1)
    
    # Age normalized by credit amount (risk indicator)
    X_train_engineered['age_credit_ratio'] = X_train['age_years'] / (X_train['credit_amount'] + 1)
    X_test_engineered['age_credit_ratio'] = X_test['age_years'] / (X_test['credit_amount'] + 1)
    
    return X_train_engineered, X_test_engineered


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred