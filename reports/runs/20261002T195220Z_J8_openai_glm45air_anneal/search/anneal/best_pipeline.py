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
    # Handle "not_applicable" values - replace with a special category
    X_train = X_train.replace('not_applicable', 'missing')
    X_test = X_test.replace('not_applicable', 'missing')
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Keep the simple geometric features that worked
    X_train_eng['area'] = X_train_eng['width'] * X_train_eng['len']
    X_test_eng['area'] = X_test_eng['width'] * X_test_eng['len']
    
    X_train_eng['volume'] = X_train_eng['area'] * X_train_eng['thick']
    X_test_eng['volume'] = X_test_eng['area'] * X_test_eng['thick']
    
    # Binary features for important categorical variables
    X_train_eng['is_steel_A'] = (X_train_eng['steel'] == 'A').astype(int)
    X_test_eng['is_steel_A'] = (X_test_eng['steel'] == 'A').astype(int)
    
    X_train_eng['is_steel_V'] = (X_train_eng['steel'] == 'V').astype(int)
    X_test_eng['is_steel_V'] = (X_test_eng['steel'] == 'V').astype(int)
    
    X_train_eng['is_shape_COIL'] = (X_train_eng['shape'] == 'COIL').astype(int)
    X_test_eng['is_shape_COIL'] = (X_test_eng['shape'] == 'COIL').astype(int)
    
    X_train_eng['is_shape_SHEET'] = (X_train_eng['shape'] == 'SHEET').astype(int)
    X_test_eng['is_shape_SHEET'] = (X_test_eng['shape'] == 'SHEET').astype(int)
    
    X_train_eng['has_condition_S'] = (X_train_eng['condition'] == 'S').astype(int)
    X_test_eng['has_condition_S'] = (X_test_eng['condition'] == 'S').astype(int)
    
    X_train_eng['has_formability_2'] = (X_train_eng['formability'] == '2').astype(int)
    X_test_eng['has_formability_2'] = (X_test_eng['formability'] == '2').astype(int)
    
    X_train_eng['has_surface_quality_E'] = (X_train_eng['surface-quality'] == 'E').astype(int)
    X_test_eng['has_surface_quality_E'] = (X_test_eng['surface-quality'] == 'E').astype(int)
    
    # Count features
    binary_features = ['temper_rolling', 'bc', 'bf', 'bt', 'bl', 'chrom', 'phos', 'cbond', 
                      'marvi', 'exptl', 'ferro', 'corr', 'blue_bright_varn_clean', 'lustre']
    
    X_train_eng['binary_feature_count'] = X_train_eng[binary_features].apply(
        lambda row: sum(1 for val in row if val == 'Y'), axis=1
    )
    X_test_eng['binary_feature_count'] = X_test_eng[binary_features].apply(
        lambda row: sum(1 for val in row if val == 'Y'), axis=1
    )
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred