import numpy as np
import pandas as pd
from sklearn.preprocessing import MinMaxScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Use Min-Max scaling which can be effective for bounded data like S-parameters
    scaler = MinMaxScaler()
    X_train_scaled = X_train.copy()
    X_test_scaled = X_test.copy()
    
    # Only scale the feature columns (not the target)
    feature_cols = [col for col in X_train.columns if col != 'Contaminated']
    
    X_train_scaled[feature_cols] = scaler.fit_transform(X_train[feature_cols])
    X_test_scaled[feature_cols] = scaler.transform(X_test[feature_cols])
    
    return X_train_scaled, y_train, X_test_scaled


def engineer(X_train, y_train, X_test):
    # Return original data - simple and effective
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred