import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Very simple feature engineering - just add a basic interaction
    # Ratio of frequency to velocity (a simple physics-based ratio)
    X_train['freq_vel_ratio'] = X_train['frequency_hz'] / (X_train['free_stream_velocity_m_s'] + 1e-8)
    X_test['freq_vel_ratio'] = X_test['frequency_hz'] / (X_test['free_stream_velocity_m_s'] + 1e-8)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred