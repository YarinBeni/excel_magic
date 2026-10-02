import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults

def preprocess(X_train, y_train, X_test):
    # Log transform the target variable to handle skewness
    y_train_log = np.log(y_train + 1e-8)  # Add small epsilon to avoid log(0)
    return X_train, y_train_log, X_test

def engineer(X_train, y_train, X_test):
    # Encode categorical variables with one-hot encoding
    X_train_encoded = pd.get_dummies(X_train, columns=['attack-angle'], prefix='attack')
    X_test_encoded = pd.get_dummies(X_test, columns=['attack-angle'], prefix='attack')
    
    # Align test set columns with train set columns
    for col in X_train_encoded.columns:
        if col not in X_test_encoded.columns:
            X_test_encoded[col] = 0
    
    # Ensure same column order
    X_test_encoded = X_test_encoded[X_train_encoded.columns]
    
    # Feature engineering: create interaction terms
    # Ratio features - these often capture meaningful relationships
    X_train_encoded['frequency/chord'] = X_train_encoded['frequency'] / (X_train_encoded['chord-length'] + 1e-8)
    X_test_encoded['frequency/chord'] = X_test_encoded['frequency'] / (X_test_encoded['chord-length'] + 1e-8)
    
    # Product features - combining physical quantities
    X_train_encoded['velocity_chord'] = X_train_encoded['free-stream-velocity'] * X_train_encoded['chord-length']
    X_test_encoded['velocity_chord'] = X_test_encoded['free-stream-velocity'] * X_test_encoded['chord-length']
    
    # Reynolds-like number (dimensionless quantity common in fluid dynamics)
    X_train_encoded['reynolds_like'] = X_train_encoded['frequency'] * X_train_encoded['chord-length'] / (X_train_encoded['free-stream-velocity'] + 1e-8)
    X_test_encoded['reynolds_like'] = X_test_encoded['frequency'] * X_test_encoded['chord-length'] / (X_test_encoded['free-stream-velocity'] + 1e-8)
    
    # Additional useful ratio
    X_train_encoded['chord/velocity'] = X_train_encoded['chord-length'] / (X_train_encoded['free-stream-velocity'] + 1e-8)
    X_test_encoded['chord/velocity'] = X_test_encoded['chord-length'] / (X_test_encoded['free-stream-velocity'] + 1e-8)
    
    # Square root of some features to reduce skew
    X_train_encoded['sqrt_velocity'] = np.sqrt(X_train_encoded['free-stream-velocity'])
    X_test_encoded['sqrt_velocity'] = np.sqrt(X_test_encoded['free-stream-velocity'])
    
    # Log of some features
    X_train_encoded['log_frequency'] = np.log1p(X_train_encoded['frequency'])
    X_test_encoded['log_frequency'] = np.log1p(X_test_encoded['frequency'])
    
    return X_train_encoded, X_test_encoded

def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows

def postprocess(pred, y_train_orig, X_test):
    # Invert log transformation
    return np.exp(pred) - 1e-8