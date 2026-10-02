import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Log transform the target variable to handle skewness
    y_train_log = np.log1p(y_train)
    return X_train, y_train_log, X_test


def engineer(X_train, y_train, X_test):
    # Create additional features
    # BMI categories
    X_train['bmi_category'] = pd.cut(X_train['bmi'], bins=[0, 18.5, 25, 30, 100], 
                                     labels=['underweight', 'normal', 'overweight', 'obese'])
    X_test['bmi_category'] = pd.cut(X_test['bmi'], bins=[0, 18.5, 25, 30, 100], 
                                    labels=['underweight', 'normal', 'overweight', 'obese'])
    
    # Age groups
    X_train['age_group'] = pd.cut(X_train['age'], bins=[0, 18, 35, 50, 100], 
                                  labels=['teen', 'young_adult', 'adult', 'senior'])
    X_test['age_group'] = pd.cut(X_test['age'], bins=[0, 18, 35, 50, 100], 
                                 labels=['teen', 'young_adult', 'adult', 'senior'])
    
    # BMI * age interaction
    X_train['bmi_age_interaction'] = X_train['bmi'] * X_train['age']
    X_test['bmi_age_interaction'] = X_test['bmi'] * X_test['age']
    
    # Children * smoker interaction
    X_train['children_smoker_interaction'] = X_train['children'] * (X_train['smoker'] == 'yes').astype(int)
    X_test['children_smoker_interaction'] = X_test['children'] * (X_test['smoker'] == 'yes').astype(int)
    
    # Smoker * bmi interaction
    X_train['smoker_bmi_interaction'] = (X_train['smoker'] == 'yes').astype(int) * X_train['bmi']
    X_test['smoker_bmi_interaction'] = (X_test['smoker'] == 'yes').astype(int) * X_test['bmi']
    
    # Create ratio features
    X_train['bmi_children_ratio'] = X_train['bmi'] / (X_train['children'] + 1)  # Avoid division by zero
    X_test['bmi_children_ratio'] = X_test['bmi'] / (X_test['children'] + 1)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert the log transformation
    pred_inverted = np.expm1(pred)
    return pred_inverted