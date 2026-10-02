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
from datetime import datetime

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values in Income - fill with median
    income_median = X_train['Income'].median()
    X_train['Income'] = X_train['Income'].fillna(income_median)
    X_test['Income'] = X_test['Income'].fillna(income_median)
    
    # Convert Dt_Customer to datetime
    X_train['Dt_Customer'] = pd.to_datetime(X_train['Dt_Customer'])
    X_test['Dt_Customer'] = pd.to_datetime(X_test['Dt_Customer'])
    
    # Convert categorical campaign columns to numeric
    campaign_cols = ['AcceptedCmp1', 'AcceptedCmp2', 'AcceptedCmp3', 'AcceptedCmp4', 'AcceptedCmp5']
    for col in campaign_cols:
        X_train[col] = pd.to_numeric(X_train[col], errors='coerce').fillna(0).astype(int)
        X_test[col] = pd.to_numeric(X_test[col], errors='coerce').fillna(0).astype(int)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Age from Year_Birth
    current_year = datetime.now().year
    X_train_eng['Age'] = current_year - X_train_eng['Year_Birth']
    X_test_eng['Age'] = current_year - X_test_eng['Year_Birth']
    
    # Date features from Dt_Customer
    X_train_eng['Customer_Days'] = (datetime.now() - X_train_eng['Dt_Customer']).dt.days
    X_test_eng['Customer_Days'] = (datetime.now() - X_test_eng['Dt_Customer']).dt.days
    
    X_train_eng['Customer_Year'] = X_train_eng['Dt_Customer'].dt.year
    X_test_eng['Customer_Year'] = X_test_eng['Dt_Customer'].dt.year
    
    X_train_eng['Customer_Month'] = X_train_eng['Dt_Customer'].dt.month
    X_test_eng['Customer_Month'] = X_test_eng['Dt_Customer'].dt.month
    
    # Family features
    X_train_eng['Total_Children'] = X_train_eng['Kidhome'] + X_train_eng['Teenhome']
    X_test_eng['Total_Children'] = X_test_eng['Kidhome'] + X_test_eng['Teenhome']
    
    X_train_eng['Family_Size'] = X_train_eng['Total_Children'] + 2  # Assuming 2 adults
    X_test_eng['Family_Size'] = X_test_eng['Total_Children'] + 2
    
    # Total spending across all categories
    spending_cols = ['MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 
                    'MntSweetProducts', 'MntGoldProds']
    X_train_eng['Total_Spending'] = X_train_eng[spending_cols].sum(axis=1)
    X_test_eng['Total_Spending'] = X_test_eng[spending_cols].sum(axis=1)
    
    # Log transforms for skewed spending variables
    for col in spending_cols + ['Total_Spending']:
        # Add 1 to avoid log(0)
        X_train_eng[f'{col}_log'] = np.log1p(X_train_eng[col])
        X_test_eng[f'{col}_log'] = np.log1p(X_test_eng[col])
    
    # Log transform Income (add small constant to handle 0 values)
    X_train_eng['Income_log'] = np.log1p(X_train_eng['Income'])
    X_test_eng['Income_log'] = np.log1p(X_test_eng['Income'])
    
    # Polynomial features for key variables
    X_train_eng['Age_squared'] = X_train_eng['Age'] ** 2
    X_test_eng['Age_squared'] = X_test_eng['Age'] ** 2
    
    X_train_eng['Income_squared'] = X_train_eng['Income'] / 10000  # Scale down first
    X_test_eng['Income_squared'] = X_test_eng['Income'] / 10000
    
    X_train_eng['Income_squared'] = X_train_eng['Income_squared'] ** 2
    X_test_eng['Income_squared'] = X_test_eng['Income_squared'] ** 2
    
    X_train_eng['Total_Spending_squared'] = X_train_eng['Total_Spending'] / 1000  # Scale down
    X_test_eng['Total_Spending_squared'] = X_test_eng['Total_Spending'] / 1000
    
    X_train_eng['Total_Spending_squared'] = X_train_eng['Total_Spending_squared'] ** 2
    X_test_eng['Total_Spending_squared'] = X_test_eng['Total_Spending_squared'] ** 2
    
    # Average spending per category
    X_train_eng['Avg_Spending_Per_Category'] = X_train_eng['Total_Spending'] / len(spending_cols)
    X_test_eng['Avg_Spending_Per_Category'] = X_test_eng['Total_Spending'] / len(spending_cols)
    
    # Spending ratios
    X_train_eng['Wines_Ratio'] = X_train_eng['MntWines'] / (X_train_eng['Total_Spending'] + 1)
    X_test_eng['Wines_Ratio'] = X_test_eng['MntWines'] / (X_test_eng['Total_Spending'] + 1)
    
    X_train_eng['Meat_Ratio'] = X_train_eng['MntMeatProducts'] / (X_train_eng['Total_Spending'] + 1)
    X_test_eng['Meat_Ratio'] = X_test_eng['MntMeatProducts'] / (X_test_eng['Total_Spending'] + 1)
    
    X_train_eng['Premium_Ratio'] = X_train_eng['MntGoldProds'] / (X_train_eng['Total_Spending'] + 1)
    X_test_eng['Premium_Ratio'] = X_test_eng['MntGoldProds'] / (X_test_eng['Total_Spending'] + 1)
    
    # Total purchases (need this before ratio features)
    purchase_cols = ['NumDealsPurchases', 'NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']
    X_train_eng['Total_Purchases'] = X_train_eng[purchase_cols].sum(axis=1)
    X_test_eng['Total_Purchases'] = X_test_eng[purchase_cols].sum(axis=1)
    
    # Additional ratio features
    X_train_eng['Web_To_Store_Ratio'] = X_train_eng['NumWebPurchases'] / (X_train_eng['NumStorePurchases'] + 1)
    X_test_eng['Web_To_Store_Ratio'] = X_test_eng['NumWebPurchases'] / (X_test_eng['NumStorePurchases'] + 1)
    
    X_train_eng['Catalog_To_Total_Ratio'] = X_train_eng['NumCatalogPurchases'] / (X_train_eng['Total_Purchases'] + 1)
    X_test_eng['Catalog_To_Total_Ratio'] = X_test_eng['NumCatalogPurchases'] / (X_test_eng['Total_Purchases'] + 1)
    
    # Purchase frequency
    X_train_eng['Purchase_Frequency'] = X_train_eng['Total_Purchases'] / (X_train_eng['Recency'] + 1)
    X_test_eng['Purchase_Frequency'] = X_test_eng['Total_Purchases'] / (X_test_eng['Recency'] + 1)
    
    # Web engagement
    X_train_eng['Web_Engagement_Ratio'] = X_train_eng['NumWebVisitsMonth'] / (X_train_eng['NumWebPurchases'] + 1)
    X_test_eng['Web_Engagement_Ratio'] = X_test_eng['NumWebVisitsMonth'] / (X_test_eng['NumWebPurchases'] + 1)
    
    # Campaign acceptance features
    campaign_cols = ['AcceptedCmp1', 'AcceptedCmp2', 'AcceptedCmp3', 'AcceptedCmp4', 'AcceptedCmp5']
    X_train_eng['Total_Campaign_Accepted'] = X_train_eng[campaign_cols].sum(axis=1)
    X_test_eng['Total_Campaign_Accepted'] = X_test_eng[campaign_cols].sum(axis=1)
    
    X_train_eng['Campaign_Acceptance_Rate'] = X_train_eng['Total_Campaign_Accepted'] / len(campaign_cols)
    X_test_eng['Campaign_Acceptance_Rate'] = X_test_eng['Total_Campaign_Accepted'] / len(campaign_cols)
    
    # Income-based features
    X_train_eng['Income_Per_Child'] = X_train_eng['Income'] / (X_train_eng['Total_Children'] + 1)
    X_test_eng['Income_Per_Child'] = X_test_eng['Income'] / (X_test_eng['Total_Children'] + 1)
    
    X_train_eng['Income_Per_Family_Member'] = X_train_eng['Income'] / X_train_eng['Family_Size']
    X_test_eng['Income_Per_Family_Member'] = X_test_eng['Income'] / X_test_eng['Family_Size']
    
    # Spending power (spending relative to income)
    X_train_eng['Spending_To_Income_Ratio'] = X_train_eng['Total_Spending'] / (X_train_eng['Income'] + 1)
    X_test_eng['Spending_To_Income_Ratio'] = X_test_eng['Total_Spending'] / (X_test_eng['Income'] + 1)
    
    # Interaction features
    X_train_eng['Age_Income_Interaction'] = X_train_eng['Age'] * X_train_eng['Income'] / 100000
    X_test_eng['Age_Income_Interaction'] = X_test_eng['Age'] * X_test_eng['Income'] / 100000
    
    X_train_eng['Spending_Income_Interaction'] = X_train_eng['Total_Spending'] * X_train_eng['Income'] / 1000000
    X_test_eng['Spending_Income_Interaction'] = X_test_eng['Total_Spending'] * X_test_eng['Income'] / 1000000
    
    X_train_eng['Family_Spending_Interaction'] = X_train_eng['Family_Size'] * X_train_eng['Total_Spending']
    X_test_eng['Family_Spending_Interaction'] = X_test_eng['Family_Size'] * X_test_eng['Total_Spending']
    
    # Drop original date column and Year_Birth (replaced by Age)
    X_train_eng = X_train_eng.drop(['Dt_Customer', 'Year_Birth'], axis=1)
    X_test_eng = X_test_eng.drop(['Dt_Customer', 'Year_Birth'], axis=1)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    # Simple probability calibration for imbalanced dataset
    if isinstance(pred, pd.DataFrame):
        # Get positive class probabilities (assuming column 1 is positive class)
        pos_probs = pred.iloc[:, 1] if pred.shape[1] > 1 else pred.iloc[:, 0]
        
        # Apply temperature scaling to calibrate probabilities
        # Temperature > 1 makes probabilities more confident, < 1 makes them less confident
        temperature = 1.2  # Slight temperature adjustment for imbalanced data
        
        # Apply temperature scaling
        calibrated_probs = np.power(pos_probs, 1.0/temperature)
        
        # Ensure probabilities are in valid range
        calibrated_probs = np.clip(calibrated_probs, 1e-7, 1 - 1e-7)
        
        # Create new DataFrame with calibrated probabilities
        calibrated_pred = pred.copy()
        calibrated_pred.iloc[:, 1] = calibrated_probs if pred.shape[1] > 1 else calibrated_probs
        
        return calibrated_pred
    else:
        return pred