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
    
    # Log transform frequency (spans several orders of magnitude)
    X_train_eng['log_frequency'] = np.log1p(X_train_eng['frequency'])
    X_test_eng['log_frequency'] = np.log1p(X_test_eng['frequency'])
    
    # Create physically meaningful interaction features
    # Reynolds number related: velocity * chord_length / thickness
    X_train_eng['reynolds_like'] = X_train_eng['free-stream-velocity'] * X_train_eng['chord-length'] / (X_train_eng['suction-side-displacement-thickness'] + 1e-8)
    X_test_eng['reynolds_like'] = X_test_eng['free-stream-velocity'] * X_test_eng['chord-length'] / (X_test_eng['suction-side-displacement-thickness'] + 1e-8)
    
    # Frequency-chord interaction (higher frequency with smaller chord = more noise)
    X_train_eng['freq_chord_ratio'] = X_train_eng['frequency'] / (X_train_eng['chord-length'] + 1e-8)
    X_test_eng['freq_chord_ratio'] = X_test_eng['frequency'] / (X_test_eng['chord-length'] + 1e-8)
    
    # Velocity-thickness interaction
    X_train_eng['velocity_thickness'] = X_train_eng['free-stream-velocity'] * X_train_eng['suction-side-displacement-thickness']
    X_test_eng['velocity_thickness'] = X_test_eng['free-stream-velocity'] * X_test_eng['suction-side-displacement-thickness']
    
    # Attack-angle encoded as numeric (assuming it's ordered)
    X_train_eng['attack_angle_numeric'] = pd.to_numeric(X_train_eng['attack-angle'], errors='coerce')
    X_test_eng['attack_angle_numeric'] = pd.to_numeric(X_test_eng['attack-angle'], errors='coerce')
    
    # Frequency-velocity interaction
    X_train_eng['freq_velocity'] = X_train_eng['frequency'] * X_train_eng['free-stream-velocity']
    X_test_eng['freq_velocity'] = X_test_eng['frequency'] * X_test_eng['free-stream-velocity']
    
    # Thickness log transform
    X_train_eng['thickness_log'] = np.log1p(X_train_eng['suction-side-displacement-thickness'])
    X_test_eng['thickness_log'] = np.log1p(X_test_eng['suction-side-displacement-thickness'])
    
    # Add back the helpful chord_length_squared
    X_train_eng['chord_length_squared'] = X_train_eng['chord-length'] ** 2
    X_test_eng['chord_length_squared'] = X_test_eng['chord-length'] ** 2
    
    # Domain-specific aerodynamic features
    # Strouhal number-like feature (frequency * thickness / velocity)
    X_train_eng['strouhal_like'] = X_train_eng['frequency'] * X_train_eng['suction-side-displacement-thickness'] / (X_train_eng['free-stream-velocity'] + 1e-8)
    X_test_eng['strouhal_like'] = X_test_eng['frequency'] * X_test_eng['suction-side-displacement-thickness'] / (X_test_eng['free-stream-velocity'] + 1e-8)
    
    # Pressure coefficient-like feature (velocity^2 / frequency)
    X_train_eng['pressure_coeff_like'] = (X_train_eng['free-stream-velocity'] ** 2) / (X_train_eng['frequency'] + 1e-8)
    X_test_eng['pressure_coeff_like'] = (X_test_eng['free-stream-velocity'] ** 2) / (X_test_eng['frequency'] + 1e-8)
    
    # Try some additional ratio features that might be physically meaningful
    # Chord-length to velocity ratio
    X_train_eng['chord_velocity_ratio'] = X_train_eng['chord-length'] / (X_train_eng['free-stream-velocity'] + 1e-8)
    X_test_eng['chord_velocity_ratio'] = X_test_eng['chord-length'] / (X_test_eng['free-stream-velocity'] + 1e-8)
    
    # Frequency to thickness ratio
    X_train_eng['freq_thickness_ratio'] = X_train_eng['frequency'] / (X_train_eng['suction-side-displacement-thickness'] + 1e-8)
    X_test_eng['freq_thickness_ratio'] = X_test_eng['frequency'] / (X_test_eng['suction-side-displacement-thickness'] + 1e-8)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    # Create multiple context views for better generalization
    n_samples = len(X_train)
    
    # View 1: All samples (original)
    view1 = np.arange(n_samples)
    
    # View 2: Stratified sampling based on target quantiles
    y_quantiles = pd.qcut(y_train, q=3, labels=False)
    view2 = []
    for quantile in range(3):
        quantile_indices = np.where(y_quantiles == quantile)[0]
        view2.extend(quantile_indices)
    view2 = np.array(view2)
    
    # View 3: Random subset for diversity
    np.random.seed(42)
    view3 = np.random.choice(n_samples, size=int(n_samples * 0.8), replace=False)
    
    # View 4: Stratified by frequency ranges
    freq_quantiles = pd.qcut(X_train['frequency'], q=3, labels=False)
    view4 = []
    for freq_quantile in range(3):
        freq_indices = np.where(freq_quantiles == freq_quantile)[0]
        view4.extend(freq_indices)
    view4 = np.array(view4)
    
    # View 5: Stratified by velocity ranges
    vel_quantiles = pd.qcut(X_train['free-stream-velocity'], q=3, labels=False)
    view5 = []
    for vel_quantile in range(3):
        vel_indices = np.where(vel_quantiles == vel_quantile)[0]
        view5.extend(vel_indices)
    view5 = np.array(view5)
    
    return [view1, view2, view3, view4, view5]


def postprocess(pred, y_train_orig, X_test):
    # Average predictions from multiple views
    if isinstance(pred, list):
        # If pred is a list of predictions from different views
        avg_pred = np.mean(pred, axis=0)
    else:
        avg_pred = pred
    
    return avg_pred