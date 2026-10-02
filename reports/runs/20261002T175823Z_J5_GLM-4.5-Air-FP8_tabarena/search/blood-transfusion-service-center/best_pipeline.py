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

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create a copy to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # RFM-based features (the ones that worked)
    
    # 1. Monetary value per donation
    avg_donation_train = X_train_eng['TotalBloodDonated'] / X_train_eng['NumberOfDonations'].replace(0, np.nan)
    avg_donation_test = X_test_eng['TotalBloodDonated'] / X_test_eng['NumberOfDonations'].replace(0, np.nan)
    X_train_eng['AvgDonationAmount'] = avg_donation_train.fillna(0)
    X_test_eng['AvgDonationAmount'] = avg_donation_test.fillna(0)
    
    # 2. Donation frequency
    freq_train = X_train_eng['NumberOfDonations'] / X_train_eng['MonthsSinceFirstDonation'].replace(0, np.nan)
    freq_test = X_test_eng['NumberOfDonations'] / X_test_eng['MonthsSinceFirstDonation'].replace(0, np.nan)
    X_train_eng['DonationFrequency'] = freq_train.fillna(0)
    X_test_eng['DonationFrequency'] = freq_test.fillna(0)
    
    # 3. Recency-Frequency interaction
    X_train_eng['RecencyFreqInteraction'] = X_train_eng['MonthsSinceLastDonation'] * X_train_eng['NumberOfDonations']
    X_test_eng['RecencyFreqInteraction'] = X_test_eng['MonthsSinceLastDonation'] * X_test_eng['NumberOfDonations']
    
    # 4. Tenure-Frequency interaction
    X_train_eng['TenureFreqInteraction'] = X_train_eng['MonthsSinceFirstDonation'] * X_train_eng['NumberOfDonations']
    X_test_eng['TenureFreqInteraction'] = X_test_eng['MonthsSinceFirstDonation'] * X_test_eng['NumberOfDonations']
    
    # 5. Ratio of recent to total donation period
    X_train_eng['RecentToTotalPeriod'] = X_train_eng['MonthsSinceLastDonation'] / X_train_eng['MonthsSinceFirstDonation'].replace(0, np.nan)
    X_test_eng['RecentToTotalPeriod'] = X_test_eng['MonthsSinceLastDonation'] / X_test_eng['MonthsSinceFirstDonation'].replace(0, np.nan)
    X_train_eng['RecentToTotalPeriod'] = X_train_eng['RecentToTotalPeriod'].fillna(1.0)
    X_test_eng['RecentToTotalPeriod'] = X_test_eng['RecentToTotalPeriod'].fillna(1.0)
    
    # 6. Time since last donation (squared)
    X_train_eng['MonthsSinceLastDonationSquared'] = X_train_eng['MonthsSinceLastDonation'] ** 2
    X_test_eng['MonthsSinceLastDonationSquared'] = X_test_eng['MonthsSinceLastDonation'] ** 2
    
    # 7. Regular donor indicator (optimized threshold)
    avg_months_between_train = X_train_eng['MonthsSinceFirstDonation'] / X_train_eng['NumberOfDonations'].replace(0, np.nan)
    avg_months_between_test = X_test_eng['MonthsSinceFirstDonation'] / X_test_eng['NumberOfDonations'].replace(0, np.nan)
    X_train_eng['RegularDonor'] = (avg_months_between_train <= 4).astype(int)
    X_test_eng['RegularDonor'] = (avg_months_between_test <= 4).astype(int)
    
    # 8. High value donor indicator
    median_donation = X_train_eng['AvgDonationAmount'].median()
    X_train_eng['HighValueDonor'] = (X_train_eng['AvgDonationAmount'] > median_donation).astype(int)
    X_test_eng['HighValueDonor'] = (X_test_eng['AvgDonationAmount'] > median_donation).astype(int)
    
    # 9. Recent donor indicator (optimized threshold)
    X_train_eng['RecentDonor'] = (X_train_eng['MonthsSinceLastDonation'] <= 6).astype(int)
    X_test_eng['RecentDonor'] = (X_test_eng['MonthsSinceLastDonation'] <= 6).astype(int)
    
    # 10. Log transform of average donation amount
    X_train_eng['LogAvgDonationAmount'] = np.log1p(X_train_eng['AvgDonationAmount'])
    X_test_eng['LogAvgDonationAmount'] = np.log1p(X_test_eng['AvgDonationAmount'])
    
    # 11. Donor consistency
    expected_donations_train = X_train_eng['MonthsSinceFirstDonation'] / 4
    expected_donations_test = X_test_eng['MonthsSinceFirstDonation'] / 4
    X_train_eng['DonorConsistency'] = X_train_eng['NumberOfDonations'] / expected_donations_train.replace(0, np.nan)
    X_test_eng['DonorConsistency'] = X_test_eng['NumberOfDonations'] / expected_donations_test.replace(0, np.nan)
    X_train_eng['DonorConsistency'] = X_train_eng['DonorConsistency'].fillna(1.0)
    X_test_eng['DonorConsistency'] = X_test_eng['DonorConsistency'].fillna(1.0)
    
    # 12. Recent high-value donor interaction
    X_train_eng['RecentHighValueDonor'] = X_train_eng['RecentDonor'] * X_train_eng['HighValueDonor']
    X_test_eng['RecentHighValueDonor'] = X_test_eng['RecentDonor'] * X_test_eng['HighValueDonor']
    
    # Optimized threshold features that worked
    X_train_eng['VeryRecentDonor'] = (X_train_eng['MonthsSinceLastDonation'] <= 3).astype(int)
    X_test_eng['VeryRecentDonor'] = (X_test_eng['MonthsSinceLastDonation'] <= 3).astype(int)
    
    X_train_eng['StrictRegularDonor'] = (avg_months_between_train <= 3).astype(int)
    X_test_eng['StrictRegularDonor'] = (avg_months_between_test <= 3).astype(int)
    
    median_freq = X_train_eng['DonationFrequency'].median()
    X_train_eng['HighFrequencyDonor'] = (X_train_eng['DonationFrequency'] > median_freq).astype(int)
    X_test_eng['HighFrequencyDonor'] = (X_test_eng['DonationFrequency'] > median_freq).astype(int)
    
    # Additional features to try
    
    # 13. Ultra recent donor (within 1 month)
    X_train_eng['UltraRecentDonor'] = (X_train_eng['MonthsSinceLastDonation'] <= 1).astype(int)
    X_test_eng['UltraRecentDonor'] = (X_test_eng['MonthsSinceLastDonation'] <= 1).astype(int)
    
    # 14. Very high value donor (above 75th percentile)
    q75_donation = X_train_eng['AvgDonationAmount'].quantile(0.75)
    X_train_eng['VeryHighValueDonor'] = (X_train_eng['AvgDonationAmount'] > q75_donation).astype(int)
    X_test_eng['VeryHighValueDonor'] = (X_test_eng['AvgDonationAmount'] > q75_donation).astype(int)
    
    # 15. Super regular donor (every 2 months)
    X_train_eng['SuperRegularDonor'] = (avg_months_between_train <= 2).astype(int)
    X_test_eng['SuperRegularDonor'] = (avg_months_between_test <= 2).astype(int)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred