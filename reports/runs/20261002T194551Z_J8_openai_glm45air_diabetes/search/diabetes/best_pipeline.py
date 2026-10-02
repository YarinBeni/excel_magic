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
    # Handle missing values coded as 0
    # For columns where 0 doesn't make physiological sense
    zero_cols = ['Glucose', 'BloodPressure', 'SkinThickness', 'Insulin']
    
    # Replace 0s with NaN for these columns
    for col in zero_cols:
        if col in X_train.columns:
            X_train[col] = X_train[col].replace(0, np.nan)
            X_test[col] = X_test[col].replace(0, np.nan)
    
    # Try different missing value strategy - use mean instead of median
    for col in zero_cols:
        if col in X_train.columns:
            mean_val = X_train[col].mean()
            X_train[col] = X_train[col].fillna(mean_val)
            X_test[col] = X_test[col].fillna(mean_val)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Key interaction features for diabetes prediction
    X_train_eng['BMI_Age'] = X_train_eng['BMI'] * X_train_eng['Age']
    X_test_eng['BMI_Age'] = X_test_eng['BMI'] * X_test_eng['Age']
    
    X_train_eng['Glucose_BMI'] = X_train_eng['Glucose'] * X_train_eng['BMI']
    X_test_eng['Glucose_BMI'] = X_test_eng['Glucose'] * X_test_eng['BMI']
    
    # Log transform of skewed features
    skewed_features = ['Insulin', 'DiabetesPedigreeFunction']
    for feature in skewed_features:
        if feature in X_train_eng.columns:
            # Add small constant to avoid log(0)
            X_train_eng[f'{feature}_log'] = np.log1p(X_train_eng[feature])
            X_test_eng[f'{feature}_log'] = np.log1p(X_test_eng[feature])
    
    # BMI categories (simple binary)
    X_train_eng['BMI_Obese'] = (X_train_eng['BMI'] >= 30).astype(int)
    X_test_eng['BMI_Obese'] = (X_test_eng['BMI'] >= 30).astype(int)
    
    # High glucose indicator
    X_train_eng['High_Glucose'] = (X_train_eng['Glucose'] >= 140).astype(int)
    X_test_eng['High_Glucose'] = (X_test_eng['Glucose'] >= 140).astype(int)
    
    # Age groups (simple binary)
    X_train_eng['Age_Middle_Age'] = ((X_train_eng['Age'] >= 35) & (X_train_eng['Age'] < 55)).astype(int)
    X_test_eng['Age_Middle_Age'] = ((X_test_eng['Age'] >= 35) & (X_test_eng['Age'] < 55)).astype(int)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    # Try multiple context views to improve generalization
    # Create stratified samples based on key features
    n_samples = 3
    samples = []
    
    # Sample 1: All data (original approach)
    samples.append(np.arange(len(X_train)))
    
    # Sample 2: Focus on high-risk patients
    high_risk_mask = (X_train['BMI'] >= 30) | (X_train['Glucose'] >= 140) | (X_train['Age'] >= 45)
    high_risk_indices = np.where(high_risk_mask)[0]
    if len(high_risk_indices) > 0:
        samples.append(high_risk_indices)
    
    # Sample 3: Focus on younger patients
    young_mask = X_train['Age'] < 35
    young_indices = np.where(young_mask)[0]
    if len(young_indices) > 0:
        samples.append(young_indices)
    
    return samples


def postprocess(pred, y_train_orig, X_test):
    # If we have multiple predictions from different views, average them
    if isinstance(pred, list):
        # Average the predictions
        avg_pred = pred[0].copy()
        for i in range(1, len(pred)):
            avg_pred = avg_pred + pred[i]
        avg_pred = avg_pred / len(pred)
        return avg_pred
    else:
        return pred