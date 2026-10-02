import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Log transform the target variable to handle skewness
    y_train_log = np.log(y_train + 1e-8)  # Add small epsilon to avoid log(0)
    return X_train, y_train_log, X_test


def engineer(X_train, y_train, X_test):
    # Create interaction features and ratios
    X_train_new = X_train.copy()
    X_test_new = X_test.copy()
    
    # Create some meaningful ratios and combinations
    X_train_new['frequency_chord_ratio'] = X_train_new['frequency'] / (X_train_new['chord-length'] + 1e-8)
    X_test_new['frequency_chord_ratio'] = X_test_new['frequency'] / (X_test_new['chord-length'] + 1e-8)
    
    X_train_new['velocity_frequency_ratio'] = X_train_new['free-stream-velocity'] / (X_train_new['frequency'] + 1e-8)
    X_test_new['velocity_frequency_ratio'] = X_test_new['free-stream-velocity'] / (X_test_new['frequency'] + 1e-8)
    
    # Product features
    X_train_new['frequency_velocity_product'] = X_train_new['frequency'] * X_train_new['free-stream-velocity']
    X_test_new['frequency_velocity_product'] = X_test_new['frequency'] * X_test_new['free-stream-velocity']
    
    # Add polynomial features  
    X_train_new['chord_squared'] = X_train_new['chord-length'] ** 2
    X_test_new['chord_squared'] = X_test_new['chord-length'] ** 2
    
    # Add some domain-specific features based on physics
    # Reynolds number approximation (simplified)
    X_train_new['reynolds_number'] = X_train_new['free-stream-velocity'] * X_train_new['chord-length'] / 1.5e-5  # approx viscosity
    X_test_new['reynolds_number'] = X_test_new['free-stream-velocity'] * X_test_new['chord-length'] / 1.5e-5
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert the log transformation
    return np.exp(pred) - 1e-8