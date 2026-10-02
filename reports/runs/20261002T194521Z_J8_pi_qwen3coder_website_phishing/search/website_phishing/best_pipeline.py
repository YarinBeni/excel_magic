import numpy as np
import pandas as pd
from sklearn.preprocessing import LabelEncoder

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Convert categorical columns to numeric using label encoding
    le = LabelEncoder()
    
    # Fit on combined data to ensure consistency
    combined_data = pd.concat([X_train, X_test], ignore_index=True)
    
    categorical_columns = X_train.select_dtypes('category').columns
    
    for col in categorical_columns:
        if col in combined_data.columns:
            # Encode the combined data
            combined_data[col] = le.fit_transform(combined_data[col].astype(str))
    
    # Split back into train and test
    X_train_encoded = combined_data.iloc[:len(X_train)]
    X_test_encoded = combined_data.iloc[len(X_train):]
    
    return X_train_encoded, y_train, X_test_encoded


def engineer(X_train, y_train, X_test):
    # Create new features based on domain knowledge
    X_train_new = X_train.copy()
    X_test_new = X_test.copy()
    
    # Create binary features for common patterns
    # Check if certain combinations are indicative of phishing
    X_train_new['has_suspicious_popup'] = (X_train_new['popUpWidnow'] == 'Suspicious').astype(int)
    X_test_new['has_suspicious_popup'] = (X_test_new['popUpWidnow'] == 'Suspicious').astype(int)
    
    X_train_new['has_phishy_ssl'] = (X_train_new['SSLfinal_State'] == 'Phishy').astype(int)
    X_test_new['has_phishy_ssl'] = (X_test_new['SSLfinal_State'] == 'Phishy').astype(int)
    
    X_train_new['has_suspicious_url_length'] = (X_train_new['URL_Length'] == 'Suspicious').astype(int)
    X_test_new['has_suspicious_url_length'] = (X_test_new['URL_Length'] == 'Suspicious').astype(int)
    
    # Age of domain feature - combine legitimate and phishy
    X_train_new['domain_age_legitimate'] = (X_train_new['age_of_domain'] == 'Legitimate').astype(int)
    X_test_new['domain_age_legitimate'] = (X_test_new['age_of_domain'] == 'Legitimate').astype(int)
    
    # IP address feature 
    X_train_new['has_ip_address'] = (X_train_new['having_IP_Address'] == 'Legitimate').astype(int)
    X_test_new['has_ip_address'] = (X_test_new['having_IP_Address'] == 'Legitimate').astype(int)
    
    # Create some interaction features
    # Combine URL-related features
    url_features = ['SFH', 'Request_URL', 'URL_of_Anchor', 'web_traffic']
    for i, feat1 in enumerate(url_features):
        for j, feat2 in enumerate(url_features[i+1:], i+1):
            if feat1 != feat2:
                feature_name = f'{feat1}_vs_{feat2}'
                X_train_new[feature_name] = (X_train_new[feat1] == X_train_new[feat2]).astype(int)
                X_test_new[feature_name] = (X_test_new[feat1] == X_test_new[feat2]).astype(int)
    
    # Create more complex interaction features
    # Count how many suspicious features we have
    suspicious_features = ['has_suspicious_popup', 'has_phishy_ssl', 'has_suspicious_url_length']
    X_train_new['suspicious_count'] = X_train_new[suspicious_features].sum(axis=1)
    X_test_new['suspicious_count'] = X_test_new[suspicious_features].sum(axis=1)
    
    # Create a feature indicating if multiple suspicious elements are present
    X_train_new['high_risk'] = (X_train_new['suspicious_count'] >= 2).astype(int)
    X_test_new['high_risk'] = (X_test_new['suspicious_count'] >= 2).astype(int)
    
    return X_train_new, X_test_new


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred