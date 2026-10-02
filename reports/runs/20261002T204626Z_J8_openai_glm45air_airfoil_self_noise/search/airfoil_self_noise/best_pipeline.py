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
from sklearn.preprocessing import StandardScaler, LabelEncoder

MODEL_KWARGS = {'n_estimators': 16, 'device': 'cuda'}  # Increased n_estimators for better performance


def preprocess(X_train, y_train, X_test):
    # Handle categorical attack-angle - encode it
    X_train_processed = X_train.copy()
    X_test_processed = X_test.copy()
    
    # Encode attack-angle as numeric
    le = LabelEncoder()
    X_train_processed['attack-angle'] = le.fit_transform(X_train_processed['attack-angle'].astype(str))
    X_test_processed['attack-angle'] = le.transform(X_test_processed['attack-angle'].astype(str))
    
    # Scale numerical features (except the encoded attack-angle)
    numeric_features = ['frequency', 'chord-length', 'free-stream-velocity', 'suction-side-displacement-thickness']
    scaler = StandardScaler()
    X_train_processed[numeric_features] = scaler.fit_transform(X_train_processed[numeric_features])
    X_test_processed[numeric_features] = scaler.transform(X_test_processed[numeric_features])
    
    return X_train_processed, y_train, X_test_processed


def engineer(X_train, y_train, X_test):
    X_train_engineered = X_train.copy()
    X_test_engineered = X_test.copy()
    
    # Create domain-relevant features for airfoil self-noise prediction
    
    # 1. Log transform frequency (likely skewed)
    X_train_engineered['log_frequency'] = np.log1p(X_train_engineered['frequency'])
    X_test_engineered['log_frequency'] = np.log1p(X_test_engineered['frequency'])
    
    # 2. Aerodynamic ratios and interactions
    # Reynolds number related features (velocity * chord / thickness)
    X_train_engineered['reynolds_like'] = (X_train_engineered['free-stream-velocity'] * 
                                          X_train_engineered['chord-length'] / 
                                          (X_train_engineered['suction-side-displacement-thickness'] + 1e-8))
    X_test_engineered['reynolds_like'] = (X_test_engineered['free-stream-velocity'] * 
                                         X_test_engineered['chord-length'] / 
                                         (X_test_engineered['suction-side-displacement-thickness'] + 1e-8))
    
    # 3. Attack angle effects (non-linear relationship expected)
    X_train_engineered['attack_angle_squared'] = X_train_engineered['attack-angle'] ** 2
    X_test_engineered['attack_angle_squared'] = X_test_engineered['attack-angle'] ** 2
    
    # 4. Frequency-velocity interaction (important for noise)
    X_train_engineered['freq_velocity_interaction'] = (X_train_engineered['frequency'] * 
                                                      X_train_engineered['free-stream-velocity'])
    X_test_engineered['freq_velocity_interaction'] = (X_test_engineered['frequency'] * 
                                                     X_test_engineered['free-stream-velocity'])
    
    # 5. Chord-thickness ratio
    X_train_engineered['chord_thickness_ratio'] = (X_train_engineered['chord-length'] / 
                                                  (X_train_engineered['suction-side-displacement-thickness'] + 1e-8))
    X_test_engineered['chord_thickness_ratio'] = (X_test_engineered['chord-length'] / 
                                                 (X_test_engineered['suction-side-displacement-thickness'] + 1e-8))
    
    # 6. Additional aerodynamic features
    # Strouhal number approximation (frequency * thickness / velocity)
    X_train_engineered['strouhal_like'] = (X_train_engineered['frequency'] * 
                                          X_train_engineered['suction-side-displacement-thickness'] / 
                                          (X_train_engineered['free-stream-velocity'] + 1e-8))
    X_test_engineered['strouhal_like'] = (X_test_engineered['frequency'] * 
                                         X_test_engineered['suction-side-displacement-thickness'] / 
                                         (X_test_engineered['free-stream-velocity'] + 1e-8))
    
    # 7. Attack angle with velocity interaction
    X_train_engineered['attack_velocity_interaction'] = (X_train_engineered['attack-angle'] * 
                                                         X_train_engineered['free-stream-velocity'])
    X_test_engineered['attack_velocity_interaction'] = (X_test_engineered['attack-angle'] * 
                                                        X_test_engineered['free-stream-velocity'])
    
    # 8. Log transform of thickness (often beneficial for small positive values)
    X_train_engineered['log_thickness'] = np.log1p(X_train_engineered['suction-side-displacement-thickness'])
    X_test_engineered['log_thickness'] = np.log1p(X_test_engineered['suction-side-displacement-thickness'])
    
    # 9. Additional domain-specific feature: velocity-thickness ratio
    X_train_engineered['velocity_thickness_ratio'] = (X_train_engineered['free-stream-velocity'] / 
                                                     (X_train_engineered['suction-side-displacement-thickness'] + 1e-8))
    X_test_engineered['velocity_thickness_ratio'] = (X_test_engineered['free-stream-velocity'] / 
                                                    (X_test_engineered['suction-side-displacement-thickness'] + 1e-8))
    
    # 10. Frequency-thickness interaction
    X_train_engineered['freq_thickness_interaction'] = (X_train_engineered['frequency'] * 
                                                        X_train_engineered['suction-side-displacement-thickness'])
    X_test_engineered['freq_thickness_interaction'] = (X_test_engineered['frequency'] * 
                                                       X_test_engineered['suction-side-displacement-thickness'])
    
    return X_train_engineered, X_test_engineered


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]  # one context view with all training rows


def postprocess(pred, y_train_orig, X_test):
    return pred