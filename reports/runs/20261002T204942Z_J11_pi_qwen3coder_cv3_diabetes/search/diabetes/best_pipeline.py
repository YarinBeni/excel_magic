import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create domain-specific features
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Add glucose to BMI ratio
    X_train['glucose_bmi_ratio'] = X_train['Glucose'] / (X_train['BMI'] + 1e-8)
    X_test['glucose_bmi_ratio'] = X_test['Glucose'] / (X_test['BMI'] + 1e-8)
    
    # Log transform skewed features
    skewed_features = ['Insulin', 'SkinThickness', 'DiabetesPedigreeFunction']
    for feature in skewed_features:
        if feature in X_train.columns:
            X_train[f'{feature}_log'] = np.log1p(X_train[feature])
            X_test[f'{feature}_log'] = np.log1p(X_test[feature])
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred