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
    # No target transform for now - keep it simple
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Domain-specific features based on aeroacoustic physics
    # 1. Strouhal-like ratio: frequency * chord / velocity (with small epsilon for stability)
    epsilon = 1e-6
    strouhal_train = X_train_eng['frequency_hz'] * X_train_eng['chord_length_m'] / (X_train_eng['free_stream_velocity_m_s'] + epsilon)
    strouhal_test = X_test_eng['frequency_hz'] * X_test_eng['chord_length_m'] / (X_test_eng['free_stream_velocity_m_s'] + epsilon)
    
    # 2. Scaled by thickness (as mentioned in task description)
    strouhal_thickness_train = strouhal_train * X_train_eng['suction_side_thickness_m']
    strouhal_thickness_test = strouhal_test * X_test_eng['suction_side_thickness_m']
    
    # 3. Simple angle interaction
    angle_interaction_train = X_train_eng['attack_angle_deg'] * 0.1  # Scale down
    angle_interaction_test = X_test_eng['attack_angle_deg'] * 0.1
    
    # 4. Velocity-chord ratio
    velocity_chord_train = X_train_eng['free_stream_velocity_m_s'] / (X_train_eng['chord_length_m'] + epsilon)
    velocity_chord_test = X_test_eng['free_stream_velocity_m_s'] / (X_test_eng['chord_length_m'] + epsilon)
    
    # Add engineered features
    X_train_eng['strouhal_ratio'] = strouhal_train
    X_train_eng['strouhal_thickness'] = strouhal_thickness_train
    X_train_eng['angle_scaled'] = angle_interaction_train
    X_train_eng['velocity_chord_ratio'] = velocity_chord_train
    
    X_test_eng['strouhal_ratio'] = strouhal_test
    X_test_eng['strouhal_thickness'] = strouhal_thickness_test
    X_test_eng['angle_scaled'] = angle_interaction_test
    X_test_eng['velocity_chord_ratio'] = velocity_chord_test
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred