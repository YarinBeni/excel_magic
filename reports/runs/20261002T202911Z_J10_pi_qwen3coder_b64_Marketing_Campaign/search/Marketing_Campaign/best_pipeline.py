import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values in Income column
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Fill missing income with median
    income_median = X_train['Income'].median()
    X_train['Income'] = X_train['Income'].fillna(income_median)
    X_test['Income'] = X_test['Income'].fillna(income_median)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Create age from Year_Birth
    X_train['Age'] = 2023 - X_train['Year_Birth']
    X_test['Age'] = 2023 - X_test['Year_Birth']
    
    # Create total spending
    spending_cols = ['MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 'MntSweetProducts', 'MntGoldProds']
    X_train['Total_Spending'] = X_train[spending_cols].sum(axis=1)
    X_test['Total_Spending'] = X_test[spending_cols].sum(axis=1)
    
    # Create spending ratio features
    X_train['Wine_Spending_Ratio'] = X_train['MntWines'] / (X_train['Total_Spending'] + 1)
    X_test['Wine_Spending_Ratio'] = X_test['MntWines'] / (X_test['Total_Spending'] + 1)
    
    # Create purchase frequency features
    purchase_cols = ['NumDealsPurchases', 'NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']
    X_train['Total_Purchases'] = X_train[purchase_cols].sum(axis=1)
    X_test['Total_Purchases'] = X_test[purchase_cols].sum(axis=1)
    
    # Create spending per purchase
    X_train['Spending_Per_Purchase'] = X_train['Total_Spending'] / (X_train['Total_Purchases'] + 1)
    X_test['Spending_Per_Purchase'] = X_test['Total_Spending'] / (X_test['Total_Purchases'] + 1)
    
    # Create family size
    X_train['Family_Size'] = X_train['Kidhome'] + X_train['Teenhome'] + 1
    X_test['Family_Size'] = X_test['Kidhome'] + X_test['Teenhome'] + 1
    
    # Create customer tenure
    X_train['Customer_Tenure'] = (pd.to_datetime('2023-01-01') - pd.to_datetime(X_train['Dt_Customer'])).dt.days
    X_test['Customer_Tenure'] = (pd.to_datetime('2023-01-01') - pd.to_datetime(X_test['Dt_Customer'])).dt.days
    
    # Create interaction features
    X_train['Income_Age'] = X_train['Income'] * X_train['Age']
    X_test['Income_Age'] = X_test['Income'] * X_test['Age']
    
    # Create log transformations for skewed features
    X_train['Log_Total_Spending'] = np.log1p(X_train['Total_Spending'])
    X_test['Log_Total_Spending'] = np.log1p(X_test['Total_Spending'])
    
    # Create ratio features that might be more informative
    X_train['Spending_To_Purchases_Ratio'] = X_train['Total_Spending'] / (X_train['Total_Purchases'] + 1)
    X_test['Spending_To_Purchases_Ratio'] = X_test['Total_Spending'] / (X_test['Total_Purchases'] + 1)
    
    # Create additional derived features
    X_train['Avg_Spending_Per_Category'] = X_train['Total_Spending'] / len(spending_cols)
    X_test['Avg_Spending_Per_Category'] = X_test['Total_Spending'] / len(spending_cols)
    
    # Create recency-based features
    X_train['Recent_Purchases'] = X_train['Recency'] / 100.0  # Normalize recency
    X_test['Recent_Purchases'] = X_test['Recency'] / 100.0
    
    # Drop original categorical columns that might cause issues
    X_train = X_train.drop(['Dt_Customer'], axis=1)
    X_test = X_test.drop(['Dt_Customer'], axis=1)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred