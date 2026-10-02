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
from sklearn.preprocessing import FunctionTransformer
from sklearn.feature_selection import SelectKBest, f_classif

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Handle missing values (zeros that are likely missing in medical data)
    # Replace zeros with NaN for columns where zero doesn't make sense
    missing_cols = ['Glucose', 'BloodPressure', 'SkinThickness', 'Insulin', 'BMI']
    
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    for col in missing_cols:
        if col in X_train.columns:
            X_train[col] = X_train[col].replace(0, np.nan)
            X_test[col] = X_test[col].replace(0, np.nan)
    
    # Use mean for missing value imputation
    for col in missing_cols:
        if col in X_train.columns:
            mean_val = X_train[col].mean()
            X_train[col] = X_train[col].fillna(mean_val)
            X_test[col] = X_test[col].fillna(mean_val)
    
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()
    
    # Create interaction features and domain-specific features for diabetes
    # BMI categories
    X_train['BMI_Category'] = pd.cut(X_train['BMI'], bins=[0, 18.5, 25, 30, 100], labels=['Underweight', 'Normal', 'Overweight', 'Obese'])
    X_test['BMI_Category'] = pd.cut(X_test['BMI'], bins=[0, 18.5, 25, 30, 100], labels=['Underweight', 'Normal', 'Overweight', 'Obese'])
    
    # Age groups
    X_train['Age_Group'] = pd.cut(X_train['Age'], bins=[0, 30, 40, 50, 60, 100], labels=['<30', '30-40', '40-50', '50-60', '60+'])
    X_test['Age_Group'] = pd.cut(X_test['Age'], bins=[0, 30, 40, 50, 60, 100], labels=['<30', '30-40', '40-50', '50-60', '60+'])
    
    # Glucose to BMI ratio (indicates insulin resistance)
    X_train['Glucose_BMI_Ratio'] = X_train['Glucose'] / X_train['BMI']
    X_test['Glucose_BMI_Ratio'] = X_test['Glucose'] / X_test['BMI']
    
    # Insulin resistance indicator
    X_train['Insulin_Resistance'] = (X_train['Glucose'] * X_train['Insulin']) / (X_train['BMI'] * 100)
    X_test['Insulin_Resistance'] = (X_test['Glucose'] * X_test['Insulin']) / (X_test['BMI'] * 100)
    
    # Pregnancy and age interaction
    X_train['Pregnancy_Age_Ratio'] = X_train['Pregnancies'] / X_train['Age']
    X_test['Pregnancy_Age_Ratio'] = X_test['Pregnancies'] / X_test['Age']
    
    # Blood pressure categories
    X_train['BP_Category'] = pd.cut(X_train['BloodPressure'], bins=[0, 80, 90, 120, 200], labels=['Normal', 'High-Normal', 'Hypertension-1', 'Hypertension-2'])
    X_test['BP_Category'] = pd.cut(X_test['BloodPressure'], bins=[0, 80, 90, 120, 200], labels=['Normal', 'High-Normal', 'Hypertension-1', 'Hypertension-2'])
    
    # Try different transformations for skewed features
    # For Insulin: try square root instead of log
    if 'Insulin' in X_train.columns:
        X_train['Insulin_sqrt'] = np.sqrt(X_train['Insulin'])
        X_test['Insulin_sqrt'] = np.sqrt(X_test['Insulin'])
    
    # For DiabetesPedigreeFunction: try box-cox like transformation (log + shift)
    if 'DiabetesPedigreeFunction' in X_train.columns:
        X_train['DiabetesPedigreeFunction_transformed'] = np.log1p(X_train['DiabetesPedigreeFunction'] * 10)
        X_test['DiabetesPedigreeFunction_transformed'] = np.log1p(X_test['DiabetesPedigreeFunction'] * 10)
    
    # Convert categorical features to numeric using one-hot encoding
    categorical_features = ['BMI_Category', 'Age_Group', 'BP_Category']
    for feature in categorical_features:
        if feature in X_train.columns:
            # One-hot encode
            dummies_train = pd.get_dummies(X_train[feature], prefix=feature)
            dummies_test = pd.get_dummies(X_test[feature], prefix=feature)
            
            # Align columns
            all_columns = set(dummies_train.columns) | set(dummies_test.columns)
            for col in all_columns:
                if col not in dummies_train.columns:
                    dummies_train[col] = 0
                if col not in dummies_test.columns:
                    dummies_test[col] = 0
            
            # Add to dataframe
            X_train = pd.concat([X_train, dummies_train], axis=1)
            X_test = pd.concat([X_test, dummies_test], axis=1)
            
            # Drop original categorical column
            X_train = X_train.drop(feature, axis=1)
            X_test = X_test.drop(feature, axis=1)
    
    # Feature selection - keep top 20 features based on ANOVA F-value
    feature_cols = [col for col in X_train.columns if col != 'TestedPositiveForDiabetes']
    if len(feature_cols) > 20:
        selector = SelectKBest(score_func=f_classif, k=20)
        X_train_selected = selector.fit_transform(X_train[feature_cols], y_train)
        X_test_selected = selector.transform(X_test[feature_cols])
        
        # Get selected feature names
        selected_features = [feature_cols[i] for i in selector.get_support(indices=True)]
        
        # Create new dataframes with selected features
        X_train = pd.DataFrame(X_train_selected, columns=selected_features, index=X_train.index)
        X_test = pd.DataFrame(X_test_selected, columns=selected_features, index=X_test.index)
    
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred