"""Candidate data pipeline around a FROZEN tabular foundation model.

The harness calls, in order:
  1. preprocess(X_train, y_train, X_test)        -> X_train, y_train, X_test   (cleaning, target transform)
  2. engineer(X_train, y_train, X_test)          -> X_train, X_test            (feature engineering, <= 500 cols)
  3. sample(X_train, y_train, X_test, max_rows)  -> list of index arrays       (context views; each view is fit separately
                                                                                 and predictions are averaged)
  4. frozen model fit/predict per view (never edit this part; MODEL_KWARGS tunes the constructor)
  5. postprocess(pred, y_train_orig, X_test)     -> pred                       (inverse target transforms, calibrate)

`pred` is a pandas DataFrame of class probabilities (columns = class labels) for classification, or a
pandas Series for regression. `y_train_orig` is the untransformed training target. All inputs are pandas
objects; X_train / X_test share columns. Do not read any file other than the ones in this directory.
"""
import numpy as np
import pandas as pd
from sklearn.preprocessing import LabelEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values in weight column with median
    weight_median = X_train['weight'].median()
    X_train['weight'] = X_train['weight'].fillna(weight_median)
    X_test['weight'] = X_test['weight'].fillna(weight_median)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Feature engineering
    # 1. Weight per month of membership (ratio feature)
    X_train_eng['weight_per_month'] = X_train_eng['weight'] / (X_train_eng['months_as_member'] + 1)
    X_test_eng['weight_per_month'] = X_test_eng['weight'] / (X_test_eng['months_as_member'] + 1)
    
    # 2. Log transform of months_as_member (skewed distribution)
    X_train_eng['log_months'] = np.log1p(X_train_eng['months_as_member'])
    X_test_eng['log_months'] = np.log1p(X_test_eng['months_as_member'])
    
    # 3. Days before squared (capture non-linear effects)
    X_train_eng['days_before_squared'] = X_train_eng['days_before'] ** 2
    X_test_eng['days_before_squared'] = X_test_eng['days_before'] ** 2
    
    # 4. Polynomial features for key numerical variables
    # Months as member cubed
    X_train_eng['months_cubed'] = X_train_eng['months_as_member'] ** 3
    X_test_eng['months_cubed'] = X_test_eng['months_as_member'] ** 3
    
    # Square root of weight
    X_train_eng['weight_sqrt'] = np.sqrt(X_train_eng['weight'])
    X_test_eng['weight_sqrt'] = np.sqrt(X_test_eng['weight'])
    
    # 5. Encode categorical variables
    categorical_cols = ['day_of_week', 'time', 'category']
    le_dict = {}
    
    for col in categorical_cols:
        le = LabelEncoder()
        # Fit on combined train+test to handle all possible values
        combined_values = pd.concat([X_train_eng[col], X_test_eng[col]]).astype(str)
        le.fit(combined_values)
        
        X_train_eng[col + '_encoded'] = le.transform(X_train_eng[col].astype(str))
        X_test_eng[col + '_encoded'] = le.transform(X_test_eng[col].astype(str))
        le_dict[col] = le
    
    # 6. Interaction features
    # Time and category interaction
    X_train_eng['time_category'] = X_train_eng['time_encoded'] * X_train_eng['category_encoded']
    X_test_eng['time_category'] = X_test_eng['time_encoded'] * X_test_eng['category_encoded']
    
    # Day of week and days before interaction
    X_train_eng['day_days_interaction'] = X_train_eng['day_of_week_encoded'] * X_train_eng['days_before']
    X_test_eng['day_days_interaction'] = X_test_eng['day_of_week_encoded'] * X_test_eng['days_before']
    
    # Weight and membership interaction
    X_train_eng['weight_months_interaction'] = X_train_eng['weight'] * X_train_eng['months_as_member']
    X_test_eng['weight_months_interaction'] = X_test_eng['weight'] * X_test_eng['months_as_member']
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    n_samples = len(X_train)
    
    # Create multiple context views for better generalization
    views = []
    
    # View 1: All training data (original)
    views.append(np.arange(n_samples))
    
    # View 2: Random subset (50% of data)
    np.random.seed(42)
    views.append(np.random.choice(n_samples, size=n_samples // 2, replace=False))
    
    # View 3: Stratified sample by target variable
    pos_indices = np.where(y_train == 1)[0]
    neg_indices = np.where(y_train == 0)[0]
    
    # Sample equal number from each class
    n_pos = min(len(pos_indices), 150)
    n_neg = min(len(neg_indices), 150)
    
    stratified_sample = np.concatenate([
        np.random.choice(pos_indices, size=n_pos, replace=False),
        np.random.choice(neg_indices, size=n_neg, replace=False)
    ])
    views.append(stratified_sample)
    
    # View 4: Another random subset with different seed
    np.random.seed(123)
    views.append(np.random.choice(n_samples, size=n_samples // 2, replace=False))
    
    # View 5: Focus on recent members (less than 12 months)
    recent_members = np.where(X_train['months_as_member'] <= 12)[0]
    if len(recent_members) > 0:
        views.append(np.random.choice(recent_members, size=min(200, len(recent_members)), replace=False))
    
    return views


def postprocess(pred, y_train_orig, X_test):
    return pred