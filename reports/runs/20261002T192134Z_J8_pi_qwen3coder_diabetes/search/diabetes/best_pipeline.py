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
    # No preprocessing needed for this dataset
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create new features based on domain knowledge
    # Ratio features that might be indicative of diabetes
    X_train_new = X_train.copy()
    X_test_new = X_test.copy()
    
    # Create BMI categories
    X_train_new['BMI_Category'] = pd.cut(X_train_new['BMI'], bins=[0, 25, 30, 100], labels=['Normal', 'Overweight', 'Obese'])
    X_test_new['BMI_Category'] = pd.cut(X_test_new['BMI'], bins=[0, 25, 30, 100], labels=['Normal', 'Overweight', 'Obese'])
    
    # Create glucose to BMI ratio (higher ratios might indicate insulin resistance)
    X_train_new['Glucose_BMI_Ratio'] = X_train_new['Glucose'] / (X_train_new['BMI'] + 1e-8)
    X_test_new['Glucose_BMI_Ratio'] = X_test_new['Glucose'] / (X_test_new['BMI'] + 1e-8)
    
    # Create insulin/glucose ratio (indicator of insulin resistance)
    X_train_new['Insulin_Glucose_Ratio'] = X_train_new['Insulin'] / (X_train_new['Glucose'] + 1e-8)
    X_test_new['Insulin_Glucose_Ratio'] = X_test_new['Insulin'] / (X_test_new['Glucose'] + 1e-8)
    
    # Create glucose squared (non-linear relationship)
    X_train_new['Glucose_Squared'] = X_train_new['Glucose'] ** 2
    X_test_new['Glucose_Squared'] = X_test_new['Glucose'] ** 2
    
    # Log transformations for skewed features
    skewed_features = ['Glucose', 'BloodPressure', 'SkinThickness', 'Insulin', 'BMI', 'DiabetesPedigreeFunction']
    for feature in skewed_features:
        if feature in X_train_new.columns:
            X_train_new[f'log_{feature}'] = np.log1p(X_train_new[feature])
            X_test_new[f'log_{feature}'] = np.log1p(X_test_new[feature])
    
    # Convert categorical variables to numeric
    X_train_new = pd.get_dummies(X_train_new, columns=['BMI_Category'], drop_first=True)
    X_test_new = pd.get_dummies(X_test_new, columns=['BMI_Category'], drop_first=True)
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred