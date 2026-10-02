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
from sklearn.preprocessing import StandardScaler, PowerTransformer
from sklearn.utils import resample

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # Apply power transformation to handle skewness, then standard scaling
    pt = PowerTransformer(method='yeo-johnson', standardize=False)
    scaler = StandardScaler()
    
    # Fit power transformer on training data
    X_train_transformed = pd.DataFrame(
        pt.fit_transform(X_train), 
        columns=X_train.columns, 
        index=X_train.index
    )
    X_test_transformed = pd.DataFrame(
        pt.transform(X_test), 
        columns=X_test.columns, 
        index=X_test.index
    )
    
    # Apply standard scaling
    X_train_scaled = pd.DataFrame(
        scaler.fit_transform(X_train_transformed), 
        columns=X_train.columns, 
        index=X_train.index
    )
    X_test_scaled = pd.DataFrame(
        scaler.transform(X_test_transformed), 
        columns=X_test.columns, 
        index=X_test.index
    )
    
    return X_train_scaled, y_train, X_test_scaled


def engineer(X_train, y_train, X_test):
    # Create copies to avoid modifying original data
    X_train_eng = X_train.copy()
    X_test_eng = X_test.copy()
    
    # Feature engineering: Create ratios and differences between related features
    feature_groups = {
        'radius': ['mean radius', 'radius error', 'worst radius'],
        'texture': ['mean texture', 'texture error', 'worst texture'],
        'perimeter': ['mean perimeter', 'perimeter error', 'worst perimeter'],
        'area': ['mean area', 'area error', 'worst area'],
        'smoothness': ['mean smoothness', 'smoothness error', 'worst smoothness'],
        'compactness': ['mean compactness', 'compactness error', 'worst compactness'],
        'concavity': ['mean concavity', 'concavity error', 'worst concavity'],
        'concave points': ['mean concave points', 'concave points error', 'worst concave points'],
        'symmetry': ['mean symmetry', 'symmetry error', 'worst symmetry'],
        'fractal dimension': ['mean fractal dimension', 'fractal dimension error', 'worst fractal dimension']
    }
    
    # Create ratios and differences for each feature group
    for group_name, features in feature_groups.items():
        if all(f in X_train.columns for f in features):
            # Mean to worst ratio (important for cancer detection)
            X_train_eng[f'{group_name}_mean_to_worst_ratio'] = X_train[features[0]] / (X_train[features[2]] + 1e-8)
            X_test_eng[f'{group_name}_mean_to_worst_ratio'] = X_test[features[0]] / (X_test[features[2]] + 1e-8)
            
            # Error to mean ratio (measurement reliability)
            X_train_eng[f'{group_name}_error_to_mean_ratio'] = X_train[features[1]] / (X_train[features[0]] + 1e-8)
            X_test_eng[f'{group_name}_error_to_mean_ratio'] = X_test[features[1]] / (X_test[features[0]] + 1e-8)
            
            # Worst minus mean (abnormality measure)
            X_train_eng[f'{group_name}_worst_minus_mean'] = X_train[features[2]] - X_train[features[0]]
            X_test_eng[f'{group_name}_worst_minus_mean'] = X_test[features[2]] - X_test[features[0]]
            
            # Worst to mean ratio (abnormality indicator)
            X_train_eng[f'{group_name}_worst_to_mean_ratio'] = X_train[features[2]] / (X_train[features[0]] + 1e-8)
            X_test_eng[f'{group_name}_worst_to_mean_ratio'] = X_test[features[2]] / (X_test[features[0]] + 1e-8)
    
    # Create aggregate features across all measurements
    mean_features = [col for col in X_train.columns if 'mean' in col]
    error_features = [col for col in X_train.columns if 'error' in col]
    worst_features = [col for col in X_train.columns if 'worst' in col]
    
    if mean_features:
        X_train_eng['mean_sum'] = X_train[mean_features].sum(axis=1)
        X_test_eng['mean_sum'] = X_test[mean_features].sum(axis=1)
        X_train_eng['mean_max'] = X_train[mean_features].max(axis=1)
        X_test_eng['mean_max'] = X_test[mean_features].max(axis=1)
        X_train_eng['mean_std'] = X_train[mean_features].std(axis=1)
        X_test_eng['mean_std'] = X_test[mean_features].std(axis=1)
        X_train_eng['mean_range'] = X_train[mean_features].max(axis=1) - X_train[mean_features].min(axis=1)
        X_test_eng['mean_range'] = X_test[mean_features].max(axis=1) - X_test[mean_features].min(axis=1)
    
    if worst_features:
        X_train_eng['worst_sum'] = X_train[worst_features].sum(axis=1)
        X_test_eng['worst_sum'] = X_test[worst_features].sum(axis=1)
        X_train_eng['worst_max'] = X_train[worst_features].max(axis=1)
        X_test_eng['worst_max'] = X_test[worst_features].max(axis=1)
        X_train_eng['worst_std'] = X_train[worst_features].std(axis=1)
        X_test_eng['worst_std'] = X_test[worst_features].std(axis=1)
    
    if error_features:
        X_train_eng['error_sum'] = X_train[error_features].sum(axis=1)
        X_test_eng['error_sum'] = X_test[error_features].sum(axis=1)
        X_train_eng['error_max'] = X_train[error_features].max(axis=1)
        X_test_eng['error_max'] = X_test[error_features].max(axis=1)
    
    # Compactness and geometric features
    if 'mean perimeter' in X_train.columns and 'mean area' in X_train.columns:
        X_train_eng['mean_circularity'] = 4 * np.pi * X_train['mean area'] / (X_train['mean perimeter'] ** 2 + 1e-8)
        X_test_eng['mean_circularity'] = 4 * np.pi * X_test['mean area'] / (X_test['mean perimeter'] ** 2 + 1e-8)
        X_train_eng['mean_compactness_corrected'] = (X_train['mean perimeter'] ** 2) / (4 * np.pi * X_train['mean area'] + 1e-8)
        X_test_eng['mean_compactness_corrected'] = (X_test['mean perimeter'] ** 2) / (4 * np.pi * X_test['mean area'] + 1e-8)
    
    if 'worst perimeter' in X_train.columns and 'worst area' in X_train.columns:
        X_train_eng['worst_circularity'] = 4 * np.pi * X_train['worst area'] / (X_train['worst perimeter'] ** 2 + 1e-8)
        X_test_eng['worst_circularity'] = 4 * np.pi * X_test['worst area'] / (X_test['worst perimeter'] ** 2 + 1e-8)
        X_train_eng['worst_compactness_corrected'] = (X_train['worst perimeter'] ** 2) / (4 * np.pi * X_train['worst area'] + 1e-8)
        X_test_eng['worst_compactness_corrected'] = (X_test['worst perimeter'] ** 2) / (4 * np.pi * X_test['worst area'] + 1e-8)
    
    # Create interaction features between important medical measurements
    important_pairs = [
        ('mean radius', 'mean area'),
        ('mean perimeter', 'mean area'),
        ('mean radius', 'mean concave points'),
        ('worst radius', 'worst area'),
        ('worst perimeter', 'worst area'),
        ('worst radius', 'worst concave points'),
    ]
    
    for feat1, feat2 in important_pairs:
        if feat1 in X_train.columns and feat2 in X_train.columns:
            X_train_eng[f'{feat1}_x_{feat2}'] = X_train[feat1] * X_train[feat2]
            X_test_eng[f'{feat1}_x_{feat2}'] = X_test[feat1] * X_test[feat2]
            X_train_eng[f'{feat1}_div_{feat2}'] = X_train[feat1] / (X_train[feat2] + 1e-8)
            X_test_eng[f'{feat1}_div_{feat2}'] = X_test[feat1] / (X_test[feat2] + 1e-8)
    
    # Create tumor volume estimation (assuming spherical approximation)
    if 'mean radius' in X_train.columns:
        X_train_eng['estimated_volume_mean'] = (4/3) * np.pi * (X_train['mean radius'] ** 3)
        X_test_eng['estimated_volume_mean'] = (4/3) * np.pi * (X_test['mean radius'] ** 3)
    
    if 'worst radius' in X_train.columns:
        X_train_eng['estimated_volume_worst'] = (4/3) * np.pi * (X_train['worst radius'] ** 3)
        X_test_eng['estimated_volume_worst'] = (4/3) * np.pi * (X_test['worst radius'] ** 3)
    
    return X_train_eng, X_test_eng


def sample(X_train, y_train, X_test, max_rows):
    # Create multiple views to handle class imbalance and provide different perspectives
    views = []
    
    # View 1: All training data (original view)
    views.append(np.arange(len(X_train)))
    
    # View 2: Balanced subsample (undersample majority class)
    benign_indices = np.where(y_train == 1)[0]
    malignant_indices = np.where(y_train == 0)[0]
    
    # Take equal number from each class
    n_balanced = min(len(benign_indices), len(malignant_indices))
    balanced_indices = np.concatenate([
        np.random.choice(benign_indices, n_balanced, replace=False),
        np.random.choice(malignant_indices, n_balanced, replace=False)
    ])
    views.append(balanced_indices)
    
    # View 3: Focus on difficult cases (higher uncertainty regions)
    # Use a simple heuristic: samples with extreme feature values
    extreme_mask = (X_train.abs() > X_train.quantile(0.95)).any(axis=1)
    if extreme_mask.sum() > 10:  # Ensure we have enough samples
        extreme_indices = np.where(extreme_mask)[0]
        if len(extreme_indices) > max_rows // 4:
            extreme_indices = np.random.choice(extreme_indices, max_rows // 4, replace=False)
        views.append(extreme_indices)
    
    # View 4: High-risk cases (samples with high abnormality scores)
    abnormality_features = [col for col in X_train.columns if 'worst_minus_mean' in col or 'worst_to_mean_ratio' in col]
    if abnormality_features:
        abnormality_scores = X_train[abnormality_features].sum(axis=1)
        high_risk_mask = abnormality_scores > abnormality_scores.quantile(0.8)
        if high_risk_mask.sum() > 10:
            high_risk_indices = np.where(high_risk_mask)[0]
            if len(high_risk_indices) > max_rows // 4:
                high_risk_indices = np.random.choice(high_risk_indices, max_rows // 4, replace=False)
            views.append(high_risk_indices)
    
    return views


def postprocess(pred, y_train_orig, X_test):
    # Apply conservative calibration for medical diagnosis
    # Adjust probabilities to be more conservative (lower false positives)
    if isinstance(pred, pd.DataFrame):
        # For binary classification, adjust the positive class probability
        if 1 in pred.columns:
            # Make predictions more conservative (reduce false positives)
            pred[1] = pred[1] * 0.95 + 0.025  # Slight shift toward neutral
            pred[0] = 1 - pred[1]  # Ensure probabilities sum to 1
    return pred