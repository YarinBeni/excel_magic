import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults

def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test

def engineer(X_train, y_train, X_test):
    # Create derived features that might be informative
    
    # Combine features to create ratios and relationships
    X_train_combined = X_train.copy()
    X_test_combined = X_test.copy()
    
    # Ratio features
    X_train_combined['donation_rate'] = X_train_combined['NumberOfDonations'] / (X_train_combined['MonthsSinceFirstDonation'] + 1)
    X_test_combined['donation_rate'] = X_test_combined['NumberOfDonations'] / (X_test_combined['MonthsSinceFirstDonation'] + 1)
    
    # Average donation amount
    X_train_combined['avg_donation_amount'] = X_train_combined['TotalBloodDonated'] / (X_train_combined['NumberOfDonations'] + 1)
    X_test_combined['avg_donation_amount'] = X_test_combined['TotalBloodDonated'] / (X_test_combined['NumberOfDonations'] + 1)
    
    # Recent donation activity
    X_train_combined['recent_activity'] = X_train_combined['MonthsSinceLastDonation'] / (X_train_combined['MonthsSinceFirstDonation'] + 1)
    X_test_combined['recent_activity'] = X_test_combined['MonthsSinceLastDonation'] / (X_test_combined['MonthsSinceFirstDonation'] + 1)
    
    # Log transformations for skewed features
    X_train_combined['log_total_donated'] = np.log1p(X_train_combined['TotalBloodDonated'])
    X_test_combined['log_total_donated'] = np.log1p(X_test_combined['TotalBloodDonated'])
    
    # Interaction features
    X_train_combined['donation_interaction'] = X_train_combined['NumberOfDonations'] * X_train_combined['MonthsSinceLastDonation']
    X_test_combined['donation_interaction'] = X_test_combined['NumberOfDonations'] * X_test_combined['MonthsSinceLastDonation']
    
    # Reciprocal features for inverse relationships
    X_train_combined['reciprocal_months_last'] = 1 / (X_train_combined['MonthsSinceLastDonation'] + 1)
    X_test_combined['reciprocal_months_last'] = 1 / (X_test_combined['MonthsSinceLastDonation'] + 1)
    
    # Additional ratio features
    X_train_combined['donations_per_month'] = X_train_combined['NumberOfDonations'] / (X_train_combined['MonthsSinceLastDonation'] + 1)
    X_test_combined['donations_per_month'] = X_test_combined['NumberOfDonations'] / (X_test_combined['MonthsSinceLastDonation'] + 1)
    
    # Polynomial features
    X_train_combined['months_squared'] = X_train_combined['MonthsSinceLastDonation'] ** 2
    X_test_combined['months_squared'] = X_test_combined['MonthsSinceLastDonation'] ** 2
    
    # Boolean features based on thresholds
    X_train_combined['frequent_donor'] = (X_train_combined['NumberOfDonations'] > 5).astype(int)
    X_test_combined['frequent_donor'] = (X_test_combined['NumberOfDonations'] > 5).astype(int)
    
    X_train_combined['large_volume_donor'] = (X_train_combined['TotalBloodDonated'] > 2000).astype(int)
    X_test_combined['large_volume_donor'] = (X_test_combined['TotalBloodDonated'] > 2000).astype(int)
    
    return X_train_combined, X_test_combined

def sample(X_train, y_train, X_test, max_rows):
    # Create multiple context views for better generalization
    n_samples = len(X_train)
    
    # View 1: All data
    view1 = np.arange(n_samples)
    
    # View 2: Stratified sampling by target
    # Separate indices by target class
    idx_0 = np.where(y_train == 0)[0]
    idx_1 = np.where(y_train == 1)[0]
    
    # Sample from each class
    n_0 = len(idx_0)
    n_1 = len(idx_1)
    
    # Take all of class 0, and a random sample of class 1
    view2_idx_1 = np.random.choice(idx_1, size=min(20, n_1), replace=False) if n_1 > 0 else np.array([])
    view2 = np.concatenate([idx_0, view2_idx_1])
    
    # View 3: Random sample
    view3 = np.random.choice(n_samples, size=min(200, n_samples), replace=False)
    
    return [view1, view2, view3]

def postprocess(pred, y_train_orig, X_test):
    return pred