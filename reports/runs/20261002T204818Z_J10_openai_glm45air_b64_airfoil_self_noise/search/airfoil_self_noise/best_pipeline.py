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
    # Convert attack-angle from categorical to numeric
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Convert attack-angle to numeric
    if 'attack-angle' in X_train.columns:
        X_train['attack-angle'] = pd.to_numeric(X_train['attack-angle'], errors='coerce')
        X_test['attack-angle'] = pd.to_numeric(X_test['attack-angle'], errors='coerce')
    
    # Apply log transforms to potentially skewed features
    # Frequency might be skewed (range 200-20000)
    if 'frequency' in X_train.columns:
        X_train['frequency_log'] = np.log1p(X_train['frequency'])
        X_test['frequency_log'] = np.log1p(X_test['frequency'])
    
    # Chord length might benefit from log transform
    if 'chord-length' in X_train.columns:
        X_train['chord-length_log'] = np.log1p(X_train['chord-length'])
        X_test['chord-length_log'] = np.log1p(X_test['chord-length'])
    
    # Thickness might also be skewed
    if 'suction-side-displacement-thickness' in X_train.columns:
        X_train['thickness_log'] = np.log1p(X_train['suction-side-displacement-thickness'])
        X_test['thickness_log'] = np.log1p(X_test['suction-side-displacement-thickness'])
    
    # Apply log transform to target (might help with multiplicative relationships)
    y_train_transformed = np.log1p(y_train)
    
    return X_train, y_train_transformed, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Simple interaction terms with log-transformed features
    if 'frequency_log' in X_train.columns and 'free-stream-velocity' in X_train.columns:
        X_train['freq_log_x_velocity'] = X_train['frequency_log'] * X_train['free-stream-velocity']
        X_test['freq_log_x_velocity'] = X_test['frequency_log'] * X_test['free-stream-velocity']
    
    if 'chord-length_log' in X_train.columns and 'attack-angle' in X_train.columns:
        X_train['chord_log_x_attack'] = X_train['chord-length_log'] * X_train['attack-angle']
        X_test['chord_log_x_attack'] = X_test['chord-length_log'] * X_test['attack-angle']
    
    if 'thickness_log' in X_train.columns and 'free-stream-velocity' in X_train.columns:
        X_train['thickness_log_x_velocity'] = X_train['thickness_log'] * X_train['free-stream-velocity']
        X_test['thickness_log_x_velocity'] = X_test['thickness_log'] * X_test['free-stream-velocity']
    
    # Simple ratio features
    if 'frequency_log' in X_train.columns and 'chord-length_log' in X_train.columns:
        X_train['freq_chord_ratio'] = X_train['frequency_log'] / (X_train['chord-length_log'] + 1e-8)
        X_test['freq_chord_ratio'] = X_test['frequency_log'] / (X_test['chord-length_log'] + 1e-8)
    
    # Additional domain-specific features for aerodynamic noise
    if 'free-stream-velocity' in X_train.columns and 'chord-length' in X_train.columns:
        # Characteristic frequency (velocity / chord)
        X_train['char_freq'] = X_train['free-stream-velocity'] / (X_train['chord-length'] + 1e-8)
        X_test['char_freq'] = X_test['free-stream-velocity'] / (X_test['chord-length'] + 1e-8)
    
    if 'attack-angle' in X_train.columns and 'thickness_log' in X_train.columns:
        # Angle of attack effect on thickness
        X_train['attack_x_thickness'] = X_train['attack-angle'] * X_train['thickness_log']
        X_test['attack_x_thickness'] = X_test['attack-angle'] * X_test['thickness_log']
    
    if 'frequency' in X_train.columns and 'suction-side-displacement-thickness' in X_train.columns:
        # Frequency-thickness relationship
        X_train['freq_x_thickness_orig'] = X_train['frequency'] * X_train['suction-side-displacement-thickness']
        X_test['freq_x_thickness_orig'] = X_test['frequency'] * X_test['suction-side-displacement-thickness']
    
    # More aerodynamic features
    if 'free-stream-velocity' in X_train.columns and 'attack-angle' in X_train.columns:
        # Velocity-angle interaction (important for noise generation)
        X_train['velocity_x_attack'] = X_train['free-stream-velocity'] * X_train['attack-angle']
        X_test['velocity_x_attack'] = X_test['free-stream-velocity'] * X_test['attack-angle']
    
    if 'chord-length' in X_train.columns and 'suction-side-displacement-thickness' in X_train.columns:
        # Chord-thickness ratio (aerodynamic efficiency)
        X_train['chord_thickness_ratio'] = X_train['chord-length'] / (X_train['suction-side-displacement-thickness'] + 1e-8)
        X_test['chord_thickness_ratio'] = X_test['chord-length'] / (X_test['suction-side-displacement-thickness'] + 1e-8)
    
    if 'frequency_log' in X_train.columns and 'attack-angle' in X_train.columns:
        # Frequency-attack angle interaction
        X_train['freq_log_x_attack'] = X_train['frequency_log'] * X_train['attack-angle']
        X_test['freq_log_x_attack'] = X_test['frequency_log'] * X_test['attack-angle']
    
    # Try some square root transforms for different variables
    if 'frequency' in X_train.columns:
        X_train['frequency_sqrt'] = np.sqrt(X_train['frequency'])
        X_test['frequency_sqrt'] = np.sqrt(X_test['frequency'])
    
    if 'free-stream-velocity' in X_train.columns:
        X_train['velocity_sqrt'] = np.sqrt(X_train['free-stream-velocity'])
        X_test['velocity_sqrt'] = np.sqrt(X_test['free-stream-velocity'])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert the log transform
    return np.expm1(pred)