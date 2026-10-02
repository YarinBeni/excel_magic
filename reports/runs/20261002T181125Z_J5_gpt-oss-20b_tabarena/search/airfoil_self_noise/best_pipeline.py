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


# Convert attack-angle to numeric float and add engineered features

def preprocess(X_train, y_train, X_test):
    # Convert attack-angle to float (it is categorical but numeric values)
    for df in (X_train, X_test):
        if 'attack-angle' in df.columns:
            df['attack-angle'] = df['attack-angle'].astype(float)
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Helper to add features to a dataframe
    def add_features(df):
        df = df.copy()
        # Basic numeric features
        freq = df['frequency']
        ang = df['attack-angle']
        chord = df['chord-length']
        vel = df['free-stream-velocity']
        disp = df['suction-side-displacement-thickness']

        # Interaction terms
        df['freq*ang'] = freq * ang
        df['freq*chord'] = freq * chord
        df['freq*vel'] = freq * vel
        df['freq*disp'] = freq * disp
        df['ang*chord'] = ang * chord
        df['ang*vel'] = ang * vel
        df['ang*disp'] = ang * disp
        df['chord*vel'] = chord * vel
        df['chord*disp'] = chord * disp
        df['vel*disp'] = vel * disp

        # Squared terms
        df['freq_sq'] = freq ** 2
        df['ang_sq'] = ang ** 2
        df['chord_sq'] = chord ** 2
        df['vel_sq'] = vel ** 2
        df['disp_sq'] = disp ** 2

        # Log transforms (add 1 to avoid log(0))
        df['log_freq'] = np.log1p(freq)
        df['log_chord'] = np.log1p(chord)
        df['log_vel'] = np.log1p(vel)
        df['log_disp'] = np.log1p(disp)

        return df

    X_train = add_features(X_train)
    X_test = add_features(X_test)
    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    # Use all rows as a single context view
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred
