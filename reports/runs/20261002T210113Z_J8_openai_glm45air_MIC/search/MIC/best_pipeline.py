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
from sklearn.preprocessing import KBinsDiscretizer

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Blood pressure ratios (clinically meaningful)
    # Systolic/Diastolic ratios for both measurements
    mask_kbrig = (X_train_eng['S_AD_KBRIG'].notna()) & (X_train_eng['D_AD_KBRIG'].notna())
    X_train_eng['BP_RATIO_KBRIG'] = np.where(mask_kbrig, 
                                           X_train_eng['S_AD_KBRIG'] / X_train_eng['D_AD_KBRIG'], 
                                           np.nan)
    X_test_eng['BP_RATIO_KBRIG'] = np.where((X_test_eng['S_AD_KBRIG'].notna()) & (X_test_eng['D_AD_KBRIG'].notna()),
                                          X_test_eng['S_AD_KBRIG'] / X_test_eng['D_AD_KBRIG'],
                                          np.nan)
    
    mask_orit = (X_train_eng['S_AD_ORIT'].notna()) & (X_train_eng['D_AD_ORIT'].notna())
    X_train_eng['BP_RATIO_ORIT'] = np.where(mask_orit,
                                          X_train_eng['S_AD_ORIT'] / X_train_eng['D_AD_ORIT'],
                                          np.nan)
    X_test_eng['BP_RATIO_ORIT'] = np.where((X_test_eng['S_AD_ORIT'].notna()) & (X_test_eng['D_AD_ORIT'].notna()),
                                         X_test_eng['S_AD_ORIT'] / X_test_eng['D_AD_ORIT'],
                                         np.nan)
    
    # Missing value indicators for columns with high missing rates
    high_missing_cols = ['S_AD_KBRIG', 'D_AD_KBRIG', 'S_AD_ORIT', 'D_AD_ORIT', 'K_BLOOD', 
                        'NA_BLOOD', 'ALT_BLOOD', 'AST_BLOOD', 'KFK_BLOOD', 'L_BLOOD', 'ROE']
    
    for col in high_missing_cols:
        if col in X_train_eng.columns:
            X_train_eng[f'{col}_MISSING'] = X_train_eng[col].isna().astype(int)
            X_test_eng[f'{col}_MISSING'] = X_test_eng[col].isna().astype(int)
    
    # Age groups (clinically relevant age categories)
    age_bins = [0, 40, 55, 65, 75, 100]
    age_labels = ['young', 'middle', 'senior', 'elderly', 'very_old']
    
    X_train_eng['AGE_GROUP'] = pd.cut(X_train_eng['AGE'], bins=age_bins, labels=age_labels, right=False)
    X_test_eng['AGE_GROUP'] = pd.cut(X_test_eng['AGE'], bins=age_bins, labels=age_labels, right=False)
    
    # Convert to categorical if not already
    if 'AGE_GROUP' in X_train_eng.columns:
        X_train_eng['AGE_GROUP'] = X_train_eng['AGE_GROUP'].astype('category')
        X_test_eng['AGE_GROUP'] = X_test_eng['AGE_GROUP'].astype('category')
    
    # Count of abnormal indicators - handle categorical data properly
    abnormal_cols = ['INF_ANAM', 'STENOK_AN', 'ZSN_A', 'SIM_GIPERT', 'DLIT_AG', 
                    'ant_im', 'lat_im', 'inf_im', 'post_im']
    
    abnormal_cols_existing = [col for col in abnormal_cols if col in X_train_eng.columns]
    if abnormal_cols_existing:
        # Convert to numeric safely - handle both numeric and categorical
        abnormal_train = pd.DataFrame()
        abnormal_test = pd.DataFrame()
        
        for col in abnormal_cols_existing:
            if pd.api.types.is_numeric_dtype(X_train_eng[col]):
                abnormal_train[col] = (X_train_eng[col] > 0).astype(int)
                abnormal_test[col] = (X_test_eng[col] > 0).astype(int)
            else:
                # For categorical, assume non-zero/non-null values are abnormal
                abnormal_train[col] = (~X_train_eng[col].isin([0, '0', np.nan])).astype(int)
                abnormal_test[col] = (~X_test_eng[col].isin([0, '0', np.nan])).astype(int)
        
        X_train_eng['ABNORMAL_COUNT'] = abnormal_train.sum(axis=1)
        X_test_eng['ABNORMAL_COUNT'] = abnormal_test.sum(axis=1)
    
    # Log transforms for skewed numerical features
    skewed_features = ['ALT_BLOOD', 'AST_BLOOD', 'KFK_BLOOD', 'ROE']
    
    for feature in skewed_features:
        if feature in X_train_eng.columns:
            # Add small constant to handle zeros and apply log transform
            train_vals = X_train_eng[feature].fillna(0)
            test_vals = X_test_eng[feature].fillna(0)
            
            # Use log1p to handle zeros better
            X_train_eng[f'{feature}_LOG'] = np.log1p(train_vals)
            X_test_eng[f'{feature}_LOG'] = np.log1p(test_vals)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred