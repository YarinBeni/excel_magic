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
    # Handle missing Income values with median
    income_median = X_train['Income'].median()
    X_train = X_train.copy()
    X_test = X_test.copy()
    X_train['Income'] = X_train['Income'].fillna(income_median)
    X_test['Income'] = X_test['Income'].fillna(income_median)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Convert categorical campaign columns to numeric (0/1)
    campaign_cols = ['AcceptedCmp1', 'AcceptedCmp2', 'AcceptedCmp3', 'AcceptedCmp4', 'AcceptedCmp5', 'Complain']
    for col in campaign_cols:
        if col in X_train.columns:
            X_train[col] = (X_train[col] == 1).astype(int)
            X_test[col] = (X_test[col] == 1).astype(int)
    
    # Convert Dt_Customer to datetime and extract features
    for df in [X_train, X_test]:
        df['Dt_Customer'] = pd.to_datetime(df['Dt_Customer'])
        df['Customer_Days'] = (datetime.now() - df['Dt_Customer']).dt.days
        df['Customer_Year'] = df['Dt_Customer'].dt.year
        df['Customer_Month'] = df['Dt_Customer'].dt.month
        df['Customer_DayOfWeek'] = df['Dt_Customer'].dt.dayofweek
    
    # Calculate age from Year_Birth
    current_year = datetime.now().year
    X_train['Age'] = current_year - X_train['Year_Birth']
    X_test['Age'] = current_year - X_test['Year_Birth']
    
    # Family size
    X_train['Family_Size'] = X_train['Kidhome'] + X_train['Teenhome']
    X_test['Family_Size'] = X_test['Kidhome'] + X_test['Teenhome']
    
    # Total spending across all categories
    spending_cols = ['MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 'MntSweetProducts', 'MntGoldProds']
    X_train['Total_Spending'] = X_train[spending_cols].sum(axis=1)
    X_test['Total_Spending'] = X_test[spending_cols].sum(axis=1)
    
    # Average spending per category
    X_train['Avg_Spending'] = X_train['Total_Spending'] / 6
    X_test['Avg_Spending'] = X_test['Total_Spending'] / 6
    
    # Spending ratios
    X_train['Wines_Ratio'] = X_train['MntWines'] / (X_train['Total_Spending'] + 1)
    X_test['Wines_Ratio'] = X_test['MntWines'] / (X_test['Total_Spending'] + 1)
    
    X_train['Meat_Ratio'] = X_train['MntMeatProducts'] / (X_train['Total_Spending'] + 1)
    X_test['Meat_Ratio'] = X_test['MntMeatProducts'] / (X_test['Total_Spending'] + 1)
    
    X_train['Premium_Ratio'] = X_train['MntGoldProds'] / (X_train['Total_Spending'] + 1)
    X_test['Premium_Ratio'] = X_test['MntGoldProds'] / (X_test['Total_Spending'] + 1)
    
    # Income-based features
    X_train['Income_per_Family'] = X_train['Income'] / (X_train['Family_Size'] + 1)
    X_test['Income_per_Family'] = X_test['Income'] / (X_test['Family_Size'] + 1)
    
    X_train['Spending_to_Income_Ratio'] = X_train['Total_Spending'] / (X_train['Income'] + 1)
    X_test['Spending_to_Income_Ratio'] = X_test['Total_Spending'] / (X_test['Income'] + 1)
    
    # Total purchases
    purchase_cols = ['NumDealsPurchases', 'NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']
    X_train['Total_Purchases'] = X_train[purchase_cols].sum(axis=1)
    X_test['Total_Purchases'] = X_test[purchase_cols].sum(axis=1)
    
    # Purchase channel preferences
    X_train['Web_Ratio'] = X_train['NumWebPurchases'] / (X_train['Total_Purchases'] + 1)
    X_test['Web_Ratio'] = X_test['NumWebPurchases'] / (X_test['Total_Purchases'] + 1)
    
    X_train['Store_Ratio'] = X_train['NumStorePurchases'] / (X_train['Total_Purchases'] + 1)
    X_test['Store_Ratio'] = X_test['NumStorePurchases'] / (X_test['Total_Purchases'] + 1)
    
    X_train['Catalog_Ratio'] = X_train['NumCatalogPurchases'] / (X_train['Total_Purchases'] + 1)
    X_test['Catalog_Ratio'] = X_test['NumCatalogPurchases'] / (X_test['Total_Purchases'] + 1)
    
    # Campaign acceptance features
    campaign_cols = ['AcceptedCmp1', 'AcceptedCmp2', 'AcceptedCmp3', 'AcceptedCmp4', 'AcceptedCmp5']
    X_train['Total_Accepted_Campaigns'] = X_train[campaign_cols].sum(axis=1)
    X_test['Total_Accepted_Campaigns'] = X_test[campaign_cols].sum(axis=1)
    
    # Customer engagement score
    X_train['Engagement_Score'] = X_train['Total_Purchases'] + X_train['Total_Accepted_Campaigns'] * 2
    X_test['Engagement_Score'] = X_test['Total_Purchases'] + X_test['Total_Accepted_Campaigns'] * 2
    
    # Recency-based features
    X_train['High_Recency'] = (X_train['Recency'] > 60).astype(int)
    X_test['High_Recency'] = (X_test['Recency'] > 60).astype(int)
    
    X_train['Low_Recency'] = (X_train['Recency'] <= 30).astype(int)
    X_test['Low_Recency'] = (X_test['Recency'] <= 30).astype(int)
    
    # Age groups
    X_train['Age_Group'] = pd.cut(X_train['Age'], bins=[0, 30, 40, 50, 60, 100], labels=['Young', 'Adult', 'Middle', 'Senior', 'Elderly'])
    X_test['Age_Group'] = pd.cut(X_test['Age'], bins=[0, 30, 40, 50, 60, 100], labels=['Young', 'Adult', 'Middle', 'Senior', 'Elderly'])
    
    # Interaction features
    X_train['Age_Income_Interaction'] = X_train['Age'] * X_train['Income'] / 100000
    X_test['Age_Income_Interaction'] = X_test['Age'] * X_test['Income'] / 100000
    
    X_train['Spending_Age_Ratio'] = X_train['Total_Spending'] / (X_train['Age'] + 1)
    X_test['Spending_Age_Ratio'] = X_test['Total_Spending'] / (X_test['Age'] + 1)
    
    # Customer value indicators
    X_train['Customer_Value_Score'] = X_train['Total_Spending'] * (1 / (X_train['Recency'] + 1))
    X_test['Customer_Value_Score'] = X_test['Total_Spending'] * (1 / (X_test['Recency'] + 1))
    
    # Web engagement features
    X_train['Web_Engagement'] = X_train['NumWebPurchases'] * X_train['NumWebVisitsMonth']
    X_test['Web_Engagement'] = X_test['NumWebPurchases'] * X_test['NumWebVisitsMonth']
    
    # Drop original date column
    X_train = X_train.drop('Dt_Customer', axis=1)
    X_test = X_test.drop('Dt_Customer', axis=1)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred