import numpy as np
import pandas as pd

MODEL_KWARGS = {}


def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Compute counts on training data
    X_train_with_target = X_train.copy()
    X_train_with_target['approved'] = y_train.values

    rr_counts = X_train_with_target.groupby(['role_id', 'resource_id']).size().reset_index(name='rr_count')
    m_counts = X_train_with_target.groupby('manager_id').size().reset_index(name='m_count')
    r_counts = X_train_with_target.groupby('role_id').size().reset_index(name='r_count')
    res_counts = X_train_with_target.groupby('resource_id').size().reset_index(name='res_count')
    mr_counts = X_train_with_target.groupby(['manager_id', 'role_id']).size().reset_index(name='mr_count')
    mr_res_counts = X_train_with_target.groupby(['manager_id', 'resource_id']).size().reset_index(name='mr_res_count')

    def merge_counts(df, counts_df, left_keys, right_keys, new_col):
        return df.merge(counts_df, left_on=left_keys, right_on=right_keys, how='left')

    X_train = merge_counts(X_train, rr_counts, ['role_id', 'resource_id'], ['role_id', 'resource_id'], 'rr_count')
    X_test = merge_counts(X_test, rr_counts, ['role_id', 'resource_id'], ['role_id', 'resource_id'], 'rr_count')

    X_train = merge_counts(X_train, m_counts, ['manager_id'], ['manager_id'], 'm_count')
    X_test = merge_counts(X_test, m_counts, ['manager_id'], ['manager_id'], 'm_count')

    X_train = merge_counts(X_train, r_counts, ['role_id'], ['role_id'], 'r_count')
    X_test = merge_counts(X_test, r_counts, ['role_id'], ['role_id'], 'r_count')

    X_train = merge_counts(X_train, res_counts, ['resource_id'], ['resource_id'], 'res_count')
    X_test = merge_counts(X_test, res_counts, ['resource_id'], ['resource_id'], 'res_count')

    X_train = merge_counts(X_train, mr_counts, ['manager_id', 'role_id'], ['manager_id', 'role_id'], 'mr_count')
    X_test = merge_counts(X_test, mr_counts, ['manager_id', 'role_id'], ['manager_id', 'role_id'], 'mr_count')

    X_train = merge_counts(X_train, mr_res_counts, ['manager_id', 'resource_id'], ['manager_id', 'resource_id'], 'mr_res_count')
    X_test = merge_counts(X_test, mr_res_counts, ['manager_id', 'resource_id'], ['manager_id', 'resource_id'], 'mr_res_count')

    count_cols = ['rr_count', 'm_count', 'r_count', 'res_count', 'mr_count', 'mr_res_count']
    X_train[count_cols] = X_train[count_cols].fillna(0)
    X_test[count_cols] = X_test[count_cols].fillna(0)

    eps = 1e-6
    # ratio features
    X_train['rr_to_r'] = X_train['rr_count'] / (X_train['r_count'] + eps)
    X_test['rr_to_r'] = X_test['rr_count'] / (X_test['r_count'] + eps)

    X_train['m_to_r'] = X_train['m_count'] / (X_train['r_count'] + eps)
    X_test['m_to_r'] = X_test['m_count'] / (X_test['r_count'] + eps)

    X_train['mr_to_m'] = X_train['mr_count'] / (X_train['m_count'] + eps)
    X_test['mr_to_m'] = X_test['mr_count'] / (X_test['m_count'] + eps)

    X_train['mr_res_to_m'] = X_train['mr_res_count'] / (X_train['m_count'] + eps)
    X_test['mr_res_to_m'] = X_test['mr_res_count'] / (X_test['m_count'] + eps)

    # log transforms of counts
    for col in count_cols:
        X_train[f'log_{col}'] = np.log1p(X_train[col])
        X_test[f'log_{col}'] = np.log1p(X_test[col])

    # log transforms of ratios
    ratio_cols = ['rr_to_r', 'm_to_r', 'mr_to_m', 'mr_res_to_m']
    for col in ratio_cols:
        X_train[f'log_{col}'] = np.log1p(X_train[col])
        X_test[f'log_{col}'] = np.log1p(X_test[col])

    # drop original target if present
    if 'approved' in X_train.columns:
        X_train = X_train.drop(columns=['approved'])
    if 'approved' in X_test.columns:
        X_test = X_test.drop(columns=['approved'])

    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred
