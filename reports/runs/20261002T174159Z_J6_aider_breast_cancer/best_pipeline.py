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
from sklearn.preprocessing import StandardScaler
from sklearn.feature_selection import SelectKBest, f_classif
from sklearn.decomposition import PCA

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # No preprocessing needed for this dataset
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Create ratios and derived features based on domain knowledge
    # Focus on ratios that are known to be important in breast cancer diagnosis
    ratio_features = []
    
    # Create ratios for key features that are commonly used in medical diagnostics
    ratio_pairs = [
        ('worst radius', 'mean radius'),
        ('worst texture', 'mean texture'),
        ('worst perimeter', 'mean perimeter'),
        ('worst area', 'mean area'),
        ('worst concavity', 'mean concavity'),
        ('worst concave points', 'mean concave points'),
        ('worst compactness', 'mean compactness'),
        ('worst smoothness', 'mean smoothness'),
        ('worst symmetry', 'mean symmetry'),
        ('worst fractal dimension', 'mean fractal dimension')
    ]
    
    for worst_col, mean_col in ratio_pairs:
        if worst_col in X_train.columns and mean_col in X_train.columns:
            ratio_name = f"{worst_col}_to_{mean_col}"
            X_train[ratio_name] = X_train[worst_col] / (X_train[mean_col] + 1e-8)  # Add small value to avoid division by zero
            X_test[ratio_name] = X_test[worst_col] / (X_test[mean_col] + 1e-8)
            ratio_features.append(ratio_name)
    
    # Create interaction terms (product of features) - focus on clinically relevant combinations
    interaction_features = []
    # Important pairs based on medical domain knowledge
    important_features = ['mean radius', 'mean texture', 'mean area', 'mean concavity', 'mean concave points']
    
    for i, feat1 in enumerate(important_features):
        for j, feat2 in enumerate(important_features[i+1:], i+1):
            if feat1 in X_train.columns and feat2 in X_train.columns:
                interaction_name = f"{feat1}_x_{feat2}"
                X_train[interaction_name] = X_train[feat1] * X_train[feat2]
                X_test[interaction_name] = X_test[feat1] * X_test[feat2]
                interaction_features.append(interaction_name)
    
    # Create polynomial features (squared terms) for non-linear relationships
    poly_features = []
    for col in ['mean radius', 'mean texture', 'mean area', 'mean concavity', 'mean concave points']:
        if col in X_train.columns:
            poly_name = f"{col}_squared"
            X_train[poly_name] = X_train[col] ** 2
            X_test[poly_name] = X_test[col] ** 2
            poly_features.append(poly_name)
    
    # Create log-transformed features for skewed data
    log_features = []
    skewed_features = ['mean area', 'worst area', 'mean perimeter', 'worst perimeter']
    for col in skewed_features:
        if col in X_train.columns:
            log_name = f"log_{col}"
            X_train[log_name] = np.log1p(X_train[col])
            X_test[log_name] = np.log1p(X_test[col])
            log_features.append(log_name)
    
    # Create additional domain-specific features
    # Compactness ratio (perimeter squared over area)
    if 'mean perimeter' in X_train.columns and 'mean area' in X_train.columns:
        comp_ratio_name = 'compactness_ratio'
        X_train[comp_ratio_name] = (X_train['mean perimeter'] ** 2) / (X_train['mean area'] + 1e-8)
        X_test[comp_ratio_name] = (X_test['mean perimeter'] ** 2) / (X_test['mean area'] + 1e-8)
        ratio_features.append(comp_ratio_name)
    
    # Concavity to area ratio
    if 'mean concavity' in X_train.columns and 'mean area' in X_train.columns:
        conc_area_ratio_name = 'concavity_to_area_ratio'
        X_train[conc_area_ratio_name] = X_train['mean concavity'] / (X_train['mean area'] + 1e-8)
        X_test[conc_area_ratio_name] = X_test['mean concavity'] / (X_test['mean area'] + 1e-8)
        ratio_features.append(conc_area_ratio_name)
    
    # Combine all new features
    new_features = ratio_features + interaction_features + poly_features + log_features
    
    # Scale all features
    scaler = StandardScaler()
    X_train_scaled = pd.DataFrame(
        scaler.fit_transform(X_train[new_features]), 
        columns=new_features, 
        index=X_train.index
    )
    X_test_scaled = pd.DataFrame(
        scaler.transform(X_test[new_features]), 
        columns=new_features, 
        index=X_test.index
    )
    
    # Combine with original features
    X_train_final = pd.concat([X_train.drop(columns=new_features), X_train_scaled], axis=1)
    X_test_final = pd.concat([X_test.drop(columns=new_features), X_test_scaled], axis=1)
    
    # Apply PCA for dimensionality reduction but keep more components
    # Since we're working with a small dataset, let's be more conservative with PCA
    n_components = min(30, len(X_train_final.columns))  # Keep more components for this small dataset
    if n_components < len(X_train_final.columns):
        pca = PCA(n_components=n_components)
        X_train_pca = pca.fit_transform(X_train_final)
        X_test_pca = pca.transform(X_test_final)
        
        # Create feature names for PCA components
        pca_feature_names = [f"pca_{i}" for i in range(X_train_pca.shape[1])]
        
        # Convert back to DataFrames
        X_train_final = pd.DataFrame(X_train_pca, columns=pca_feature_names, index=X_train_final.index)
        X_test_final = pd.DataFrame(X_test_pca, columns=pca_feature_names, index=X_test_final.index)
    
    # Feature selection to keep top features
    selector = SelectKBest(score_func=f_classif, k=min(50, len(X_train_final.columns)))
    X_train_selected = selector.fit_transform(X_train_final, y_train)
    selected_features = X_train_final.columns[selector.get_support()]
    
    X_train_final = pd.DataFrame(X_train_selected, columns=selected_features, index=X_train_final.index)
    X_test_final = X_test_final[selected_features]
    
    return X_train_final, X_test_final


def sample(X_train, y_train, X_test, max_rows):
    # Create stratified sampling to ensure balanced views
    # Since we have a binary classification, create balanced context views
    n_samples = len(X_train)
    
    # Create multiple context views with balanced sampling
    # This helps the frozen model learn better from different perspectives
    if n_samples <= max_rows:
        # If we have fewer samples than max_rows, just use all
        return [np.arange(n_samples)]
    else:
        # Create multiple overlapping views with better balance
        views = []
        n_views = min(5, n_samples // max_rows + 1)  # Create at most 5 views
        
        # For binary classification, ensure both classes are represented
        malignant_indices = np.where(y_train == 0)[0]
        benign_indices = np.where(y_train == 1)[0]
        
        # Sample from each class proportionally
        for i in range(n_views):
            # Shuffle indices
            np.random.shuffle(malignant_indices)
            np.random.shuffle(benign_indices)
            
            # Take roughly equal amounts from each class
            n_malignant = len(malignant_indices) // 2
            n_benign = len(benign_indices) // 2
            
            # Ensure we don't exceed available samples
            n_malignant = min(n_malignant, len(malignant_indices))
            n_benign = min(n_benign, len(benign_indices))
            
            # Take samples from both classes
            sampled_indices = np.concatenate([
                malignant_indices[:n_malignant],
                benign_indices[:n_benign]
            ])
            
            # Shuffle the final indices
            np.random.shuffle(sampled_indices)
            views.append(sampled_indices)
        
        return views


def postprocess(pred, y_train_orig, X_test):
    # Apply temperature scaling for better calibration
    if isinstance(pred, pd.DataFrame):
        # For binary classification, we need to handle the probability properly
        # Get the probability of the positive class (benign)
        prob_positive = pred.iloc[:, 1]  # Assuming second column is positive class
    else:
        # For regression-like outputs or single column
        prob_positive = pred
    
    # Use a simpler but more effective calibration approach
    # Apply sigmoid transformation to push probabilities toward extremes
    # This helps with better separation between classes
    calibrated_prob = 1.0 / (1.0 + np.exp(-5 * (prob_positive - 0.5)))
    
    # Return calibrated predictions
    if isinstance(pred, pd.DataFrame):
        result = pred.copy()
        result.iloc[:, 1] = calibrated_prob
        result.iloc[:, 0] = 1 - calibrated_prob
        return result
    else:
        return calibrated_prob
