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

# ---------------------------------------------------------------------------
# 1. Preprocess: nothing special needed for this dataset.
# ---------------------------------------------------------------------------

def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test

# ---------------------------------------------------------------------------
# 2. Feature engineering: moderate set of ratios, logs, squares and pairwise
#    interactions between mean, worst and error columns. Keeps feature count
#    well below 500.
# ---------------------------------------------------------------------------

def engineer(X_train, y_train, X_test):
    def add_features(df):
        # Base columns
        mean_cols = [c for c in df.columns if c.startswith('mean ')]
        worst_cols = [c.replace('mean ', 'worst ') for c in mean_cols]
        error_cols = [c.replace('mean ', '') + ' error' for c in mean_cols]

        # Basic transformations
        for m, w, e in zip(mean_cols, worst_cols, error_cols):
            df[m + ' / ' + w] = df[m] / (df[w] + 1e-6)
            df[e + ' / ' + m] = df[e] / (df[m] + 1e-6)
            df[e + ' / ' + w] = df[e] / (df[w] + 1e-6)
            df[m + ' * ' + w] = df[m] * df[w]
            df[m + ' * ' + e] = df[m] * df[e]
            df[w + ' * ' + e] = df[w] * df[e]
            df['log ' + m] = np.log1p(df[m])
            df['log ' + w] = np.log1p(df[w])
            df['log ' + e] = np.log1p(df[e])

        # Squares
        for m in mean_cols:
            df[m + '^2'] = df[m] ** 2
        for w in worst_cols:
            df[w + '^2'] = df[w] ** 2
        for e in error_cols:
            df[e + '^2'] = df[e] ** 2

        # Pairwise interactions between mean and worst (10x10)
        for m in mean_cols:
            for w in worst_cols:
                df[m + ' * ' + w] = df[m] * df[w]

        # Pairwise interactions between mean and error (10x10)
        for m in mean_cols:
            for e in error_cols:
                df[m + ' * ' + e] = df[m] * df[e]

        return df

    X_train = add_features(X_train.copy())
    X_test = add_features(X_test.copy())
    return X_train, X_test

# ---------------------------------------------------------------------------
# 3. Sampling: single context view with all training rows.
# ---------------------------------------------------------------------------

def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]

# ---------------------------------------------------------------------------
# 4. Postprocess: no target transformation needed; return raw probabilities.
# ---------------------------------------------------------------------------

def postprocess(pred, y_train_orig, X_test):
    return pred
