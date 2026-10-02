import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults

def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test

def engineer(X_train, y_train, X_test):
    # Minimal feature engineering - just a couple of useful combinations
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Add BP difference which might be clinically relevant
    X_train['BP_Diff'] = X_train['SystolicBP'] - X_train['DiastolicBP']
    X_test['BP_Diff'] = X_test['SystolicBP'] - X_test['DiastolicBP']
    
    return X_train, X_test

def sample(X_train, y_train, X_test, max_rows):
    # Use all training data - sometimes this works better
    return [np.arange(len(X_train))]  # one context view with all training rows

def postprocess(pred, y_train_orig, X_test):
    return pred