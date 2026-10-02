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
    # Convert categorical columns to numeric
    X_train['attack-angle'] = pd.to_numeric(X_train['attack-angle'], downcast='float')
    X_test['attack-angle'] = pd.to_numeric(X_test['attack-angle'], downcast='float')
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create engineered features based on domain knowledge
    # Key aerodynamic ratios and interactions
    X_train['freq_chord_ratio'] = X_train['frequency'] / (X_train['chord-length'] + 1e-8)
    X_test['freq_chord_ratio'] = X_test['frequency'] / (X_test['chord-length'] + 1e-8)
    
    X_train['attack_chord_interaction'] = X_train['attack-angle'] * X_train['chord-length']
    X_test['attack_chord_interaction'] = X_test['attack-angle'] * X_test['chord-length']
    
    X_train['thickness_chord_ratio'] = X_train['suction-side-displacement-thickness'] / (X_train['chord-length'] + 1e-8)
    X_test['thickness_chord_ratio'] = X_test['suction-side-displacement-thickness'] / (X_test['chord-length'] + 1e-8)
    
    X_train['freq_vel_product'] = X_train['frequency'] * X_train['free-stream-velocity']
    X_test['freq_vel_product'] = X_test['frequency'] * X_test['free-stream-velocity']
    
    # Log transformation of frequency
    X_train['log_frequency'] = np.log1p(X_train['frequency'])
    X_test['log_frequency'] = np.log1p(X_test['frequency'])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred