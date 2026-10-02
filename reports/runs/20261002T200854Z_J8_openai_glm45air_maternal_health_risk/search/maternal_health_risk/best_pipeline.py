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
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Create frequency features to handle duplicates
    # Combine all features to create a unique row identifier
    train_rows = X_train_eng.astype(str).agg('|'.join, axis=1)
    test_rows = X_test_eng.astype(str).agg('|'.join, axis=1)
    
    # Calculate frequencies
    train_freq = train_rows.value_counts()
    all_freq = pd.concat([train_rows, test_rows]).value_counts()
    
    # Add frequency features
    X_train_eng['RowFrequency'] = train_rows.map(train_freq).values
    X_test_eng['RowFrequency'] = test_rows.map(all_freq).values
    
    # Blood pressure categories
    def bp_category(row):
        systolic = row['SystolicBP']
        diastolic = row['DiastolicBP']
        
        if systolic < 120 and diastolic < 80:
            return 0  # Normal
        elif 120 <= systolic <= 129 and diastolic < 80:
            return 1  # Elevated
        elif (130 <= systolic <= 139) or (80 <= diastolic <= 89):
            return 2  # Hypertension Stage 1
        else:
            return 3  # Hypertension Stage 2
    
    X_train_eng['BPCategory'] = X_train_eng.apply(bp_category, axis=1)
    X_test_eng['BPCategory'] = X_test_eng.apply(bp_category, axis=1)
    
    # Pulse pressure (Systolic - Diastolic)
    X_train_eng['PulsePressure'] = X_train_eng['SystolicBP'] - X_train_eng['DiastolicBP']
    X_test_eng['PulsePressure'] = X_test_eng['SystolicBP'] - X_test_eng['DiastolicBP']
    
    # Blood sugar categories
    def bs_category(bs):
        if bs < 100:
            return 0  # Normal
        elif 100 <= bs <= 125:
            return 1  # Prediabetes
        else:
            return 2  # Diabetes
    
    X_train_eng['BSCategory'] = X_train_eng['BS'].apply(bs_category)
    X_test_eng['BSCategory'] = X_test_eng['BS'].apply(bs_category)
    
    # Age groups
    def age_group(age):
        if age < 20:
            return 0  # Adolescent
        elif age <= 35:
            return 1  # Young adult
        else:
            return 2  # Advanced maternal age
    
    X_train_eng['AgeGroup'] = X_train_eng['Age'].apply(age_group)
    X_test_eng['AgeGroup'] = X_test_eng['Age'].apply(age_group)
    
    # Heart rate categories
    def hr_category(hr):
        if hr < 60:
            return 0  # Bradycardia
        elif hr <= 100:
            return 1  # Normal
        else:
            return 2  # Tachycardia
    
    X_train_eng['HRCategory'] = X_train_eng['HeartRate'].apply(hr_category)
    X_test_eng['HRCategory'] = X_test_eng['HeartRate'].apply(hr_category)
    
    # Body temperature categories
    def temp_category(temp):
        if temp < 97:
            return 0  # Low
        elif temp <= 99:
            return 1  # Normal
        elif temp <= 101:
            return 2  # Slight fever
        else:
            return 3  # High fever
    
    X_train_eng['TempCategory'] = X_train_eng['BodyTemp'].apply(temp_category)
    X_test_eng['TempCategory'] = X_test_eng['BodyTemp'].apply(temp_category)
    
    # Risk indicators (count of abnormal values)
    def risk_indicators(row):
        count = 0
        # High blood pressure
        if row['SystolicBP'] >= 140 or row['DiastolicBP'] >= 90:
            count += 1
        # High blood sugar
        if row['BS'] >= 126:
            count += 1
        # High heart rate
        if row['HeartRate'] > 100:
            count += 1
        # High temperature
        if row['BodyTemp'] > 101:
            count += 1
        # Advanced maternal age
        if row['Age'] >= 35:
            count += 1
        return count
    
    X_train_eng['RiskIndicators'] = X_train_eng.apply(risk_indicators, axis=1)
    X_test_eng['RiskIndicators'] = X_test_eng.apply(risk_indicators, axis=1)
    
    # Simple target encoding without reindexing
    categorical_features = ['BPCategory', 'BSCategory', 'AgeGroup', 'HRCategory', 'TempCategory']
    
    for feature in categorical_features:
        # Create a mapping from category to target mean
        temp_df = pd.DataFrame({feature: X_train_eng[feature], 'target': y_train})
        target_means = temp_df.groupby(feature)['target'].mean()
        
        # Apply mapping with fallback to overall mean
        overall_mean = y_train.mean()
        X_train_eng[f'{feature}_TargetEnc'] = X_train_eng[feature].map(target_means).fillna(overall_mean)
        X_test_eng[f'{feature}_TargetEnc'] = X_test_eng[feature].map(target_means).fillna(overall_mean)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred