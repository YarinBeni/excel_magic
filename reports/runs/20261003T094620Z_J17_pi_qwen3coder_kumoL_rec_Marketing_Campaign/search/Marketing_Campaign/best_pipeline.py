import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values in Income column
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Fill missing Income with median
    income_median = X_train['Income'].median()
    X_train['Income'].fillna(income_median, inplace=True)
    X_test['Income'].fillna(income_median, inplace=True)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying originals
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Feature engineering
    # Age feature from Year_Birth
    X_train['Age'] = 2023 - X_train['Year_Birth']
    X_test['Age'] = 2023 - X_test['Year_Birth']
    
    # Total spending across all product categories
    spending_cols = ['MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 
                     'MntSweetProducts', 'MntGoldProds']
    X_train['Total_Spending'] = X_train[spending_cols].sum(axis=1)
    X_test['Total_Spending'] = X_test[spending_cols].sum(axis=1)
    
    # Average spending per category
    X_train['Avg_Spending_Per_Category'] = X_train['Total_Spending'] / len(spending_cols)
    X_test['Avg_Spending_Per_Category'] = X_test['Total_Spending'] / len(spending_cols)
    
    # Ratio features
    X_train['Spending_Ratio_Wines'] = X_train['MntWines'] / (X_train['Total_Spending'] + 1e-8)
    X_test['Spending_Ratio_Wines'] = X_test['MntWines'] / (X_test['Total_Spending'] + 1e-8)
    
    # Purchase frequency features
    purchase_cols = ['NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']
    X_train['Total_Purchases'] = X_train[purchase_cols].sum(axis=1)
    X_test['Total_Purchases'] = X_test[purchase_cols].sum(axis=1)
    
    # Deal purchases ratio
    X_train['Deal_Purchase_Ratio'] = X_train['NumDealsPurchases'] / (X_train['Total_Purchases'] + 1e-8)
    X_test['Deal_Purchase_Ratio'] = X_test['NumDealsPurchases'] / (X_test['Total_Purchases'] + 1e-8)
    
    # Customer lifetime (days since registration)
    X_train['Customer_Lifetime_Days'] = (pd.to_datetime('2023-01-01') - pd.to_datetime(X_train['Dt_Customer'])).dt.days
    X_test['Customer_Lifetime_Days'] = (pd.to_datetime('2023-01-01') - pd.to_datetime(X_test['Dt_Customer'])).dt.days
    
    # Interaction features
    X_train['Income_Age'] = X_train['Income'] * X_train['Age']
    X_test['Income_Age'] = X_test['Income'] * X_test['Age']
    
    # Log transformations for skewed features
    skewed_features = ['Income', 'Total_Spending', 'Customer_Lifetime_Days']
    for feature in skewed_features:
        if feature in X_train.columns:
            X_train[f'{feature}_log'] = np.log1p(X_train[feature])
            X_test[f'{feature}_log'] = np.log1p(X_test[feature])
    
    # Drop original date column that's not needed anymore
    X_train.drop(['Dt_Customer'], axis=1, inplace=True)
    X_test.drop(['Dt_Customer'], axis=1, inplace=True)
    
    # Convert categorical features to numeric
    categorical_columns = ['Education', 'Marital_Status']
    for col in categorical_columns:
        if col in X_train.columns:
            # Create dummy variables
            dummies = pd.get_dummies(X_train[col], prefix=col)
            X_train = pd.concat([X_train, dummies], axis=1)
            X_train.drop(col, axis=1, inplace=True)
            
            # Apply same transformation to test set
            dummies_test = pd.get_dummies(X_test[col], prefix=col)
            X_test = pd.concat([X_test, dummies_test], axis=1)
            X_test.drop(col, axis=1, inplace=True)
    
    # Ensure consistent columns between train and test sets
    train_cols = set(X_train.columns)
    test_cols = set(X_test.columns)
    common_cols = train_cols.intersection(test_cols)
    
    # Add missing columns to test set
    for col in train_cols - test_cols:
        X_test[col] = 0
        
    # Remove extra columns from test set
    for col in test_cols - train_cols:
        X_test.drop(col, axis=1, inplace=True)
        
    # Reorder columns to match training set
    X_test = X_test.reindex(columns=X_train.columns, fill_value=0)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred