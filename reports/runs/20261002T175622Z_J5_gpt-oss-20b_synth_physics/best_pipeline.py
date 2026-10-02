import numpy as np
import pandas as pd
from sklearn.preprocessing import PolynomialFeatures

MODEL_KWARGS = {}


def preprocess(X_train, y_train, X_test):
    # Transform target to reduce skewness
    offset = abs(y_train.min()) + 1e-6
    y_train = np.log1p(y_train + offset)
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    def add_features(df):
        df = df.copy()
        # Basic physical features
        df['strouhal'] = df['frequency_hz'] * df['chord_length_m'] / df['free_stream_velocity_m_s']
        df['strouhal_thick'] = df['strouhal'] * df['suction_side_thickness_m']
        df['attack_angle_rad'] = np.deg2rad(df['attack_angle_deg'])
        df['attack_angle_sq'] = df['attack_angle_deg'] ** 2
        df['strouhal_attack'] = df['strouhal'] * df['attack_angle_deg']
        df['strouhal_thick_attack'] = df['strouhal_thick'] * df['attack_angle_deg']
        # Log transforms
        df['log_frequency'] = np.log1p(df['frequency_hz'])
        df['log_velocity'] = np.log1p(df['free_stream_velocity_m_s'])
        df['log_chord'] = np.log1p(df['chord_length_m'])
        df['log_thickness'] = np.log1p(df['suction_side_thickness_m'])
        # Ratios
        df['chord_thickness_ratio'] = df['chord_length_m'] / (df['suction_side_thickness_m'] + 1e-6)
        df['freq_velocity_ratio'] = df['frequency_hz'] / (df['free_stream_velocity_m_s'] + 1e-6)
        # Trigonometric angle features
        df['sin_angle'] = np.sin(df['attack_angle_rad'])
        df['cos_angle'] = np.cos(df['attack_angle_rad'])
        # Prepare columns for polynomial expansion
        numeric_cols = [
            'frequency_hz', 'free_stream_velocity_m_s', 'chord_length_m',
            'attack_angle_deg', 'suction_side_thickness_m',
            'strouhal', 'strouhal_thick', 'attack_angle_rad',
            'attack_angle_sq', 'strouhal_attack', 'strouhal_thick_attack',
            'log_frequency', 'log_velocity', 'log_chord', 'log_thickness',
            'chord_thickness_ratio', 'freq_velocity_ratio', 'sin_angle', 'cos_angle'
        ]
        poly = PolynomialFeatures(degree=2, include_bias=False, interaction_only=False)
        X_poly = poly.fit_transform(df[numeric_cols])
        # Remove original columns from polynomial output
        X_poly = X_poly[:, len(numeric_cols):]
        poly_cols = poly.get_feature_names_out(numeric_cols)[len(numeric_cols):]
        poly_cols = [f"poly_{c}" for c in poly_cols]
        df_poly = pd.DataFrame(X_poly, columns=poly_cols, index=df.index)
        df = pd.concat([df, df_poly], axis=1)
        return df

    X_train = add_features(X_train)
    X_test = add_features(X_test)
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    # Invert target transform
    offset = abs(y_train_orig.min()) + 1e-6
    return np.expm1(pred) - offset
