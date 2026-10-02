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
    # Apply log transform to target (concrete strength often follows log-normal distribution)
    y_train_transformed = np.log1p(y_train)
    
    # Add key ratios in preprocess
    X_train_proc = X_train.copy()
    X_test_proc = X_test.copy()
    
    # Most important concrete ratios
    X_train_proc['WaterCementRatio'] = X_train_proc['Water'] / (X_train_proc['Cement'] + 1)
    X_test_proc['WaterCementRatio'] = X_test_proc['Water'] / (X_test_proc['Cement'] + 1)
    
    X_train_proc['TotalBinder'] = X_train_proc['Cement'] + X_train_proc['BlastFurnaceSlag'] + X_train_proc['FlyAsh']
    X_test_proc['TotalBinder'] = X_test_proc['Cement'] + X_test_proc['BlastFurnaceSlag'] + X_test_proc['FlyAsh']
    
    X_train_proc['WaterBinderRatio'] = X_train_proc['Water'] / (X_train_proc['TotalBinder'] + 1)
    X_test_proc['WaterBinderRatio'] = X_test_proc['Water'] / (X_test_proc['TotalBinder'] + 1)
    
    return X_train_proc, y_train_transformed, X_test_proc


def engineer(X_train, y_train, X_test):
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Log transforms for skewed features
    for feature in ['BlastFurnaceSlag', 'FlyAsh', 'Superplasticizer', 'Age']:
        X_train_eng[f'Log{feature}'] = np.log1p(X_train_eng[feature])
        X_test_eng[f'Log{feature}'] = np.log1p(X_test_eng[feature])
    
    # Age-cement interaction (strength development over time)
    X_train_eng['AgeCementInteraction'] = X_train_eng['Age'] * X_train_eng['Cement'] / 1000
    X_test_eng['AgeCementInteraction'] = X_test_eng['Age'] * X_test_eng['Cement'] / 1000
    
    # Aggregate properties
    X_train_eng['TotalAggregate'] = X_train_eng['CoarseAggregate'] + X_train_eng['FineAggregate']
    X_test_eng['TotalAggregate'] = X_test_eng['CoarseAggregate'] + X_test_eng['FineAggregate']
    
    X_train_eng['AggregateRatio'] = X_train_eng['FineAggregate'] / (X_train_eng['CoarseAggregate'] + 1)
    X_test_eng['AggregateRatio'] = X_test_eng['FineAggregate'] / (X_test_eng['CoarseAggregate'] + 1)
    
    # Paste to aggregate ratio (cementitious materials to total aggregate)
    X_train_eng['PasteAggregateRatio'] = X_train_eng['TotalBinder'] / (X_train_eng['TotalAggregate'] + 1)
    X_test_eng['PasteAggregateRatio'] = X_test_eng['TotalBinder'] / (X_test_eng['TotalAggregate'] + 1)
    
    # Age-related features for strength development modeling
    X_train_eng['LogAge'] = np.log1p(X_train_eng['Age'])
    X_test_eng['LogAge'] = np.log1p(X_test_eng['Age'])
    
    # Cement content ratio
    X_train_eng['CementRatio'] = X_train_eng['Cement'] / (X_train_eng['TotalBinder'] + 1)
    X_test_eng['CementRatio'] = X_test_eng['Cement'] / (X_test_eng['TotalBinder'] + 1)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Invert the log transform
    return np.expm1(pred)