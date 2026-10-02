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
from sklearn.preprocessing import OneHotEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Apply sqrt transformation to target (price) to handle skewness
    y_train_sqrt = np.sqrt(y_train)
    return X_train, y_train_sqrt, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Convert age to years for better interpretability
    X_train_eng['age_in_years'] = X_train_eng['age_in_days'] / 365.25
    X_test_eng['age_in_years'] = X_test_eng['age_in_days'] / 365.25
    
    # Create mileage per year feature (important for car valuation)
    X_train_eng['km_per_year'] = X_train_eng['km'] / X_train_eng['age_in_years']
    X_test_eng['km_per_year'] = X_test_eng['km'] / X_test_eng['age_in_years']
    
    # Handle potential division by zero for very new cars
    X_train_eng['km_per_year'] = X_train_eng['km_per_year'].fillna(0)
    X_test_eng['km_per_year'] = X_test_eng['km_per_year'].fillna(0)
    
    # One-hot encode the model column
    model_encoder = OneHotEncoder(sparse_output=False, handle_unknown='ignore')
    
    # Fit on training data and transform both train and test
    model_train_encoded = model_encoder.fit_transform(X_train_eng[['model']])
    model_test_encoded = model_encoder.transform(X_test_eng[['model']])
    
    # Get feature names for the encoded columns
    model_features = [f'model_{cat}' for cat in model_encoder.categories_[0]]
    
    # Add encoded features to the dataframes
    for i, feature in enumerate(model_features):
        X_train_eng[feature] = model_train_encoded[:, i]
        X_test_eng[feature] = model_test_encoded[:, i]
    
    # Drop the original model column since it's now encoded
    X_train_eng = X_train_eng.drop('model', axis=1)
    X_test_eng = X_test_eng.drop('model', axis=1)
    
    # Create interaction features
    X_train_eng['engine_power_age_interaction'] = X_train_eng['engine_power'] * X_train_eng['age_in_years']
    X_test_eng['engine_power_age_interaction'] = X_test_eng['engine_power'] * X_test_eng['age_in_years']
    
    X_train_eng['engine_power_km_interaction'] = X_train_eng['engine_power'] * X_train_eng['km']
    X_test_eng['engine_power_km_interaction'] = X_test_eng['engine_power'] * X_test_eng['km']
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Convert sqrt predictions back to original scale
    pred_original = pred ** 2
    return pred_original