"""Candidate data pipeline around a FROZEN tabular foundation model.

The harness calls, in order:
  1. preprocess(X_train, y_train, X_test)        -> X_train, y_train, X_test   (cleaning, target transform)
  2. engineer(X_train, y_train, X_test)          -> X_train, X_test            (feature engineering, <= 500 cols)
  3. sample(X_train, y_train, X_test, max_rows)  -> list of index arrays       (context views; each view is fit separately
                                                                                 and predictions are averaged)
  4. frozen model fit/predict per view (never edit this part; MODEL_KWARGS tunes the constructor)
  5. postprocess(pred, y_train_orig, X_test)     -> pred                       (inverse target transforms, calibrate)

`pred` is a pandas DataFrame of class probabilities (columns = class labels) for classification, or a
pandas Series for regression. `y_train_orig` is the untransformed training target. All inputs are pandas
objects; X_train / X_test share columns. Do not read any file other than the ones in this directory.
"""
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    X_train_processed = X_train.copy()
    X_test_processed = X_test.copy()
    
    # Log transform skewed features (based on the describe output)
    skewed_features = [
        'Sum_dO_EStates',  # Range 0 to 71, likely skewed
        'Laplace_MoharIndex2',  # Range 0.579 to 17.537, likely skewed
        'Weighted_HyperWiener_Index_Burden_Matrix',  # Range 1.544 to 5.692
        'Weighted_LeadingEigenvalue_Burden_Matrix',  # Range 2.267 to 10.692
        'Weighted_SpectralMoment6_Burden_Matrix'  # Range 4.917 to 14.7
    ]
    
    for feature in skewed_features:
        if feature in X_train_processed.columns:
            # Add small constant to avoid log(0)
            X_train_processed[f'{feature}_log'] = np.log1p(X_train_processed[feature])
            X_test_processed[f'{feature}_log'] = np.log1p(X_test_processed[feature])
    
    # Add square root transformations for some features
    sqrt_features = [
        'Num_Heavy_Atoms',
        'Num_Circuits', 
        'Num_Substituted_BenzeneC',
        'Num_Terminal_PrimaryC'
    ]
    
    for feature in sqrt_features:
        if feature in X_train_processed.columns:
            X_train_processed[f'{feature}_sqrt'] = np.sqrt(X_train_processed[feature])
            X_test_processed[f'{feature}_sqrt'] = np.sqrt(X_test_processed[feature])
    
    return X_train_processed, y_train, X_test_processed


def engineer(X_train, y_train, X_test):
    X_train_engineered = X_train.copy()
    X_test_engineered = X_test.copy()
    
    # Domain-specific chemical features for biodegradability
    if 'Num_Halogen_Atoms' in X_train.columns and 'Num_Oxygen_Atoms' in X_train.columns:
        X_train_engineered['Halogen_O_Ratio'] = X_train['Num_Halogen_Atoms'] / (X_train['Num_Oxygen_Atoms'] + 1)
        X_test_engineered['Halogen_O_Ratio'] = X_test['Num_Halogen_Atoms'] / (X_test['Num_Oxygen_Atoms'] + 1)
    
    if 'Num_Nitrogen_Atoms' in X_train.columns and 'Num_Heavy_Atoms' in X_train.columns:
        X_train_engineered['N_Heavy_Ratio'] = X_train['Num_Nitrogen_Atoms'] / (X_train['Num_Heavy_Atoms'] + 1)
        X_test_engineered['N_Heavy_Ratio'] = X_test['Num_Nitrogen_Atoms'] / (X_test['Num_Heavy_Atoms'] + 1)
    
    if 'Num_Circuits' in X_train.columns and 'Num_Heavy_Atoms' in X_train.columns:
        X_train_engineered['Ring_Heavy_Ratio'] = X_train['Num_Circuits'] / (X_train['Num_Heavy_Atoms'] + 1)
        X_test_engineered['Ring_Heavy_Ratio'] = X_test['Num_Circuits'] / (X_test['Num_Heavy_Atoms'] + 1)
    
    # Polarity indicators
    if 'Num_Oxygen_Atoms' in X_train.columns and 'Num_Nitrogen_Atoms' in X_train.columns and 'Num_Heavy_Atoms' in X_train.columns:
        X_train_engineered['Polarity_Index'] = (X_train['Num_Oxygen_Atoms'] + X_train['Num_Nitrogen_Atoms']) / (X_train['Num_Heavy_Atoms'] + 1)
        X_test_engineered['Polarity_Index'] = (X_test['Num_Oxygen_Atoms'] + X_test['Num_Nitrogen_Atoms']) / (X_test['Num_Heavy_Atoms'] + 1)
    
    # Branching indicators
    if 'Num_Terminal_PrimaryC' in X_train.columns and 'Num_Heavy_Atoms' in X_train.columns:
        X_train_engineered['Branching_Index'] = X_train['Num_Terminal_PrimaryC'] / (X_train['Num_Heavy_Atoms'] + 1)
        X_test_engineered['Branching_Index'] = X_test['Num_Terminal_PrimaryC'] / (X_test['Num_Heavy_Atoms'] + 1)
    
    # Aromaticity
    if 'Num_Aromatic_Nitro_Groups' in X_train.columns and 'Num_Heavy_Atoms' in X_train.columns:
        X_train_engineered['Aromatic_Ratio'] = X_train['Num_Aromatic_Nitro_Groups'] / (X_train['Num_Heavy_Atoms'] + 1)
        X_test_engineered['Aromatic_Ratio'] = X_test['Num_Aromatic_Nitro_Groups'] / (X_test['Num_Heavy_Atoms'] + 1)
    
    # Interaction terms between domain features
    if 'Halogen_O_Ratio' in X_train_engineered.columns and 'Polarity_Index' in X_train_engineered.columns:
        X_train_engineered['Halogen_Polarity_Interaction'] = X_train_engineered['Halogen_O_Ratio'] * X_train_engineered['Polarity_Index']
        X_test_engineered['Halogen_Polarity_Interaction'] = X_test_engineered['Halogen_O_Ratio'] * X_test_engineered['Polarity_Index']
    
    if 'Ring_Heavy_Ratio' in X_train_engineered.columns and 'Branching_Index' in X_train_engineered.columns:
        X_train_engineered['Ring_Branching_Interaction'] = X_train_engineered['Ring_Heavy_Ratio'] * X_train_engineered['Branching_Index']
        X_test_engineered['Ring_Branching_Interaction'] = X_test_engineered['Ring_Heavy_Ratio'] * X_test_engineered['Branching_Index']
    
    if 'N_Heavy_Ratio' in X_train_engineered.columns and 'Aromatic_Ratio' in X_train_engineered.columns:
        X_train_engineered['N_Aromatic_Interaction'] = X_train_engineered['N_Heavy_Ratio'] * X_train_engineered['Aromatic_Ratio']
        X_test_engineered['N_Aromatic_Interaction'] = X_test_engineered['N_Heavy_Ratio'] * X_test_engineered['Aromatic_Ratio']
    
    # Add a simple feature count that might capture complexity
    atom_features = ['Num_Heavy_Atoms', 'Num_Oxygen_Atoms', 'Num_Nitrogen_Atoms', 'Num_Halogen_Atoms', 
                    'Num_Carbon_Atoms', 'Num_Hydrogen_Atoms']
    present_atom_features = [f for f in atom_features if f in X_train.columns]
    if present_atom_features:
        X_train_engineered['Atom_Feature_Count'] = X_train[present_atom_features].sum(axis=1)
        X_test_engineered['Atom_Feature_Count'] = X_test[present_atom_features].sum(axis=1)
    
    return X_train_engineered, X_test_engineered


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred