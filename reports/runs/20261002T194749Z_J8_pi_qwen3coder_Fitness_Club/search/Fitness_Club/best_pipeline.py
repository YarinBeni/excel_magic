import numpy as np
import pandas as pd
from sklearn.preprocessing import LabelEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults

def preprocess(X_train, y_train, X_test):
    # Handle missing values in weight column
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Fill missing weight values with median
    weight_median = X_train['weight'].median()
    X_train['weight'] = X_train['weight'].fillna(weight_median)
    X_test['weight'] = X_test['weight'].fillna(weight_median)
    
    return X_train, y_train, X_test

def engineer(X_train, y_train, X_test):
    # Create engineered features
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Encode categorical variables
    le_day = LabelEncoder()
    le_time = LabelEncoder()
    le_category = LabelEncoder()
    
    X_train['day_of_week_encoded'] = le_day.fit_transform(X_train['day_of_week'])
    X_test['day_of_week_encoded'] = le_day.transform(X_test['day_of_week'])
    
    X_train['time_encoded'] = le_time.fit_transform(X_train['time'])
    X_test['time_encoded'] = le_time.transform(X_test['time'])
    
    X_train['category_encoded'] = le_category.fit_transform(X_train['category'])
    X_test['category_encoded'] = le_category.transform(X_test['category'])
    
    # Create interaction features
    X_train['weight_months_ratio'] = X_train['weight'] / (X_train['months_as_member'] + 1)
    X_test['weight_months_ratio'] = X_test['weight'] / (X_test['months_as_member'] + 1)
    
    # Create time-based features
    X_train['is_weekend'] = (X_train['day_of_week'] == 'Sat') | (X_train['day_of_week'] == 'Sun')
    X_test['is_weekend'] = (X_test['day_of_week'] == 'Sat') | (X_test['day_of_week'] == 'Sun')
    
    # Create binary features for time
    X_train['is_PM'] = (X_train['time'] == 'PM')
    X_test['is_PM'] = (X_test['time'] == 'PM')
    
    # Drop original categorical columns
    X_train = X_train.drop(['day_of_week', 'time', 'category'], axis=1)
    X_test = X_test.drop(['day_of_week', 'time', 'category'], axis=1)
    
    return X_train, X_test

def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows

def postprocess(pred, y_train_orig, X_test):
    return pred