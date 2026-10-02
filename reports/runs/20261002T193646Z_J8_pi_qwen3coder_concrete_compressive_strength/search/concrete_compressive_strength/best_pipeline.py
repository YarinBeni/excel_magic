import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # No preprocessing needed for this dataset
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Combine training and test data for consistent feature engineering
    n_train = len(X_train)
    combined = pd.concat([X_train, X_test], ignore_index=True)
    
    # Create key features based on domain knowledge and correlations
    # Water-to-cement ratio (strongly correlated with strength)
    combined['water_cement_ratio'] = combined['Water'] / (combined['Cement'] + 1e-8)
    
    # Cementitious materials sum (important for strength)
    combined['cementitious'] = (combined['Cement'] + 
                               combined['BlastFurnaceSlag'] + 
                               combined['FlyAsh'])
    
    # Interaction terms
    combined['cement_age'] = combined['Cement'] * combined['Age']
    
    # Split back into train and test
    X_train_new = combined.iloc[:n_train]
    X_test_new = combined.iloc[n_train:]
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred