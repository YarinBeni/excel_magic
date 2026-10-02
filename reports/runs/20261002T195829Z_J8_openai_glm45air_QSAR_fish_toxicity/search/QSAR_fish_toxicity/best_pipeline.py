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
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    # Handle duplicates by creating multiple context views
    # Since there are 15% duplicates when ignoring target, create stratified views
    
    # Create a hash of feature columns to identify duplicates
    feature_hash = (X_train['CIC0'].astype(str) + '_' + 
                   X_train['SM1_Dz(Z)'].astype(str) + '_' +
                   X_train['GATS1i'].astype(str) + '_' +
                   X_train['NdsCH'].astype(str) + '_' +
                   X_train['NdssC'].astype(str) + '_' +
                   X_train['MLOGP'].astype(str))
    
    # Get unique feature combinations
    unique_features = feature_hash.unique()
    
    # Create multiple views by sampling unique feature combinations
    views = []
    n_views = 3  # Create 3 views to handle duplicates
    
    for i in range(n_views):
        # Sample unique feature combinations for this view
        if len(unique_features) > max_rows // 2:
            sampled_features = np.random.choice(unique_features, size=max_rows//2, replace=False)
        else:
            sampled_features = unique_features
            
        # Get indices for this view
        view_indices = []
        for feat in sampled_features:
            view_indices.extend(np.where(feature_hash == feat)[0])
        
        views.append(np.array(view_indices))
    
    return views


def postprocess(pred, y_train_orig, X_test):
    # Average predictions from multiple views
    return pred.mean(axis=0) if hasattr(pred, 'shape') and len(pred.shape) > 1 else pred