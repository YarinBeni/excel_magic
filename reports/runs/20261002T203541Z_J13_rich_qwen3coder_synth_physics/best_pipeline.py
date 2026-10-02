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
    # No preprocessing needed for this task
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create domain-inspired features based on the physics description
    # Strouhal-like ratio: frequency * chord / velocity
    X_train['strouhal_ratio'] = (X_train['frequency_hz'] * X_train['chord_length_m']) / X_train['free_stream_velocity_m_s']
    X_test['strouhal_ratio'] = (X_test['frequency_hz'] * X_test['chord_length_m']) / X_test['free_stream_velocity_m_s']
    
    # Thickness scaling factor
    X_train['thickness_scaled'] = X_train['suction_side_thickness_m'] * X_train['chord_length_m']
    X_test['thickness_scaled'] = X_test['suction_side_thickness_m'] * X_test['chord_length_m']
    
    # Angle term (could be sine/cosine of angle)
    X_train['angle_rad'] = np.radians(X_train['attack_angle_deg'])
    X_test['angle_rad'] = np.radians(X_test['attack_angle_deg'])
    X_train['angle_sin'] = np.sin(X_train['angle_rad'])
    X_test['angle_sin'] = np.sin(X_test['angle_rad'])
    X_train['angle_cos'] = np.cos(X_train['angle_rad'])
    X_test['angle_cos'] = np.cos(X_test['angle_rad'])
    
    # Ratio features
    X_train['freq_vel_ratio'] = X_train['frequency_hz'] / X_train['free_stream_velocity_m_s']
    X_test['freq_vel_ratio'] = X_test['frequency_hz'] / X_test['free_stream_velocity_m_s']
    
    # Product features
    X_train['chord_vel_product'] = X_train['chord_length_m'] * X_train['free_stream_velocity_m_s']
    X_test['chord_vel_product'] = X_test['chord_length_m'] * X_test['free_stream_velocity_m_s']
    
    # Log transformations for skewed features
    X_train['log_chord'] = np.log1p(X_train['chord_length_m'])
    X_test['log_chord'] = np.log1p(X_test['chord_length_m'])
    
    # Normalize some key features
    scaler = StandardScaler()
    feature_cols = ['strouhal_ratio', 'thickness_scaled', 'angle_sin', 'angle_cos', 
                   'freq_vel_ratio', 'chord_vel_product', 'log_chord']
    
    X_train[feature_cols] = scaler.fit_transform(X_train[feature_cols])
    X_test[feature_cols] = scaler.transform(X_test[feature_cols])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # No postprocessing needed
    return pred