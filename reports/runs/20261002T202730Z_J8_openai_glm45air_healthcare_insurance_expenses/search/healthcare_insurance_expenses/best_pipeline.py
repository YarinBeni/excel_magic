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
from sklearn.preprocessing import LabelEncoder

MODEL_KWARGS = {'n_estimators': 16}  # Try more estimators


def preprocess(X_train, y_train, X_test):
    # Log transform the target to handle right-skewness
    y_train_transformed = np.log1p(y_train)
    
    # Encode categorical variables with one-hot encoding
    X_train_processed = X_train.copy()
    X_test_processed = X_test.copy()
    
    # One-hot encode categorical columns
    categorical_cols = ['sex', 'smoker', 'region']
    
    for col in categorical_cols:
        # One-hot encode
        train_dummies = pd.get_dummies(X_train[col], prefix=col, drop_first=True)
        test_dummies = pd.get_dummies(X_test[col], prefix=col, drop_first=True)
        
        # Ensure both train and test have same columns
        for col_name in train_dummies.columns:
            if col_name not in test_dummies.columns:
                test_dummies[col_name] = 0
        for col_name in test_dummies.columns:
            if col_name not in train_dummies.columns:
                train_dummies[col_name] = 0
        
        # Add to processed data
        X_train_processed = pd.concat([X_train_processed.drop(col, axis=1), train_dummies], axis=1)
        X_test_processed = pd.concat([X_test_processed.drop(col, axis=1), test_dummies], axis=1)
    
    return X_train_processed, y_train_transformed, X_test_processed


def engineer(X_train, y_train, X_test):
    X_train_engineered = X_train.copy()
    X_test_engineered = X_test.copy()
    
    # BMI categories
    def bmi_category(bmi):
        if bmi < 18.5:
            return 0  # underweight
        elif bmi < 25:
            return 1  # normal
        elif bmi < 30:
            return 2  # overweight
        else:
            return 3  # obese
    
    X_train_engineered['bmi_category'] = X_train['bmi'].apply(bmi_category)
    X_test_engineered['bmi_category'] = X_test['bmi'].apply(bmi_category)
    
    # Age groups
    def age_group(age):
        if age < 30:
            return 0
        elif age < 40:
            return 1
        elif age < 50:
            return 2
        else:
            return 3
    
    X_train_engineered['age_group'] = X_train['age'].apply(age_group)
    X_test_engineered['age_group'] = X_test['age'].apply(age_group)
    
    # BMI squared for non-linear relationship
    X_train_engineered['bmi_squared'] = X_train['bmi'] ** 2
    X_test_engineered['bmi_squared'] = X_test['bmi'] ** 2
    
    # Age squared for non-linear relationship
    X_train_engineered['age_squared'] = X_train['age'] ** 2
    X_test_engineered['age_squared'] = X_test['age'] ** 2
    
    # Interaction features that are safe (numeric combinations)
    X_train_engineered['age_bmi_interaction'] = X_train['age'] * X_train['bmi']
    X_test_engineered['age_bmi_interaction'] = X_test['age'] * X_test['bmi']
    
    # Get smoker column
    smoker_col = [col for col in X_train.columns if col.startswith('smoker_')][0]
    
    X_train_engineered['smoker_age'] = X_train[smoker_col] * X_train['age']
    X_test_engineered['smoker_age'] = X_test[smoker_col] * X_test['age']
    
    X_train_engineered['smoker_bmi'] = X_train[smoker_col] * X_train['bmi']
    X_test_engineered['smoker_bmi'] = X_test[smoker_col] * X_test['bmi']
    
    return X_train_engineered, X_test_engineered


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Convert log-transformed predictions back to original scale
    pred_original = np.expm1(pred)
    return pred_original