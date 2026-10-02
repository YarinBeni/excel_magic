import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Fill missing income with median
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    income_median = X_train['Income'].median()
    X_train['Income'] = X_train['Income'].fillna(income_median)
    X_test['Income'] = X_test['Income'].fillna(income_median)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying originals
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Convert date to days since customer registration
    X_train['Dt_Customer'] = pd.to_datetime(X_train['Dt_Customer'])
    X_test['Dt_Customer'] = pd.to_datetime(X_test['Dt_Customer'])
    
    # Calculate days since customer registration
    reference_date = X_train['Dt_Customer'].max()
    X_train['Days_Since_Customer'] = (reference_date - X_train['Dt_Customer']).dt.days
    X_test['Days_Since_Customer'] = (reference_date - X_test['Dt_Customer']).dt.days
    
    # Create age from birth year
    X_train['Age'] = 2022 - X_train['Year_Birth']
    X_test['Age'] = 2022 - X_test['Year_Birth']
    
    # Create total spending
    spending_cols = ['MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 'MntSweetProducts', 'MntGoldProds']
    X_train['Total_Spending'] = X_train[spending_cols].sum(axis=1)
    X_test['Total_Spending'] = X_test[spending_cols].sum(axis=1)
    
    # Create spending ratios
    X_train['Spending_Ratio_Wine'] = X_train['MntWines'] / (X_train['Total_Spending'] + 1e-8)
    X_test['Spending_Ratio_Wine'] = X_test['MntWines'] / (X_test['Total_Spending'] + 1e-8)
    
    X_train['Spending_Ratio_Fruit'] = X_train['MntFruits'] / (X_train['Total_Spending'] + 1e-8)
    X_test['Spending_Ratio_Fruit'] = X_test['MntFruits'] / (X_test['Total_Spending'] + 1e-8)
    
    # Create purchase frequency features
    X_train['Total_Purchases'] = X_train[['NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']].sum(axis=1)
    X_test['Total_Purchases'] = X_test[['NumWebPurchases', 'NumCatalogPurchases', 'NumStorePurchases']].sum(axis=1)
    
    # Create deal engagement
    X_train['Deal_Engagement'] = X_train['NumDealsPurchases'] / (X_train['Total_Purchases'] + 1e-8)
    X_test['Deal_Engagement'] = X_test['NumDealsPurchases'] / (X_test['Total_Purchases'] + 1e-8)
    
    # Create interaction features
    X_train['Income_Age'] = X_train['Income'] * X_train['Age']
    X_test['Income_Age'] = X_test['Income'] * X_test['Age']
    
    X_train['Income_Spending'] = X_train['Income'] * X_train['Total_Spending']
    X_test['Income_Spending'] = X_test['Income'] * X_test['Total_Spending']
    
    # Log transform skewed features
    skewed_features = ['Income', 'Total_Spending', 'MntWines', 'MntFruits', 'MntMeatProducts', 'MntFishProducts', 'MntSweetProducts', 'MntGoldProds']
    for col in skewed_features:
        if col in X_train.columns:
            X_train[f'{col}_log'] = np.log1p(X_train[col])
            X_test[f'{col}_log'] = np.log1p(X_test[col])
    
    # Create categorical aggregations
    # Create family size from kids and teens
    X_train['Family_Size'] = X_train['Kidhome'] + X_train['Teenhome']
    X_test['Family_Size'] = X_test['Kidhome'] + X_test['Teenhome']
    
    # Create high-value customer indicator
    X_train['High_Value_Customer'] = (X_train['Total_Spending'] > X_train['Total_Spending'].median()).astype(int)
    X_test['High_Value_Customer'] = (X_test['Total_Spending'] > X_test['Total_Spending'].median()).astype(int)
    
    # Remove original date column
    X_train = X_train.drop(['Dt_Customer'], axis=1)
    X_test = X_test.drop(['Dt_Customer'], axis=1)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred