import numpy as np
import pandas as pd
from sklearn.preprocessing import LabelEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # No preprocessing needed for this task
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Combine train and test for consistent processing
    X_full = pd.concat([X_train, X_test], ignore_index=True)
    
    # Handle categorical variables by converting to string first, then encoding
    categorical_columns = ['sex', 'education', 'product_type', 'having_children_flg', 'region', 
                          'family_status', 'phone_operator', 'is_client']
    
    for col in categorical_columns:
        if col in X_full.columns:
            # Convert to string to handle any potential issues, then encode
            X_full[col] = X_full[col].astype(str)
            le = LabelEncoder()
            X_full[col] = le.fit_transform(X_full[col])
    
    # Separate back into train and test
    X_train_new = X_full.iloc[:len(X_train)]
    X_test_new = X_full.iloc[len(X_train):]
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred