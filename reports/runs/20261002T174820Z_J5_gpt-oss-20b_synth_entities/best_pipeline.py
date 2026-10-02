import numpy as np
import pandas as pd

MODEL_KWARGS = {}  # constructor settings for the frozen model; {} keeps the defaults


def preprocess(X_train, y_train, X_test):
    # No target transformation needed for binary classification
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    # Combine target with training features for counting
    train_df = X_train.copy()
    train_df['approved'] = y_train
    test_df = X_test.copy()

    # Create combined role-resource pair string
    train_df['role_resource'] = train_df['role_id'].astype(str) + '_' + train_df['resource_id'].astype(str)
    test_df['role_resource'] = test_df['role_id'].astype(str) + '_' + test_df['resource_id'].astype(str)

    # Compute counts and approval rates per role-resource pair
    rr_counts = train_df.groupby('role_resource')['approved'].agg(['count', 'sum']).rename(columns={'count':'rr_count', 'sum':'rr_approved'}).reset_index()
    rr_counts['rr_rate'] = rr_counts['rr_approved'] / rr_counts['rr_count']
    rr_counts = rr_counts[['role_resource', 'rr_count', 'rr_rate']]
    train_df = train_df.merge(rr_counts, on='role_resource', how='left')
    test_df = test_df.merge(rr_counts, on='role_resource', how='left')

    # Compute counts and approval rates per manager
    mgr_counts = train_df.groupby('manager_id')['approved'].agg(['count', 'sum']).rename(columns={'count':'mgr_count', 'sum':'mgr_approved'}).reset_index()
    mgr_counts['mgr_rate'] = mgr_counts['mgr_approved'] / mgr_counts['mgr_count']
    mgr_counts = mgr_counts[['manager_id', 'mgr_count', 'mgr_rate']]
    train_df = train_df.merge(mgr_counts, on='manager_id', how='left')
    test_df = test_df.merge(mgr_counts, on='manager_id', how='left')

    # Compute counts per role and resource for additional context
    role_counts = train_df.groupby('role_id')['approved'].agg(['count', 'sum']).rename(columns={'count':'role_count', 'sum':'role_approved'}).reset_index()
    role_counts['role_rate'] = role_counts['role_approved'] / role_counts['role_count']
    role_counts = role_counts[['role_id', 'role_count', 'role_rate']]
    train_df = train_df.merge(role_counts, on='role_id', how='left')
    test_df = test_df.merge(role_counts, on='role_id', how='left')

    res_counts = train_df.groupby('resource_id')['approved'].agg(['count', 'sum']).rename(columns={'count':'res_count', 'sum':'res_approved'}).reset_index()
    res_counts['res_rate'] = res_counts['res_approved'] / res_counts['res_count']
    res_counts = res_counts[['resource_id', 'res_count', 'res_rate']]
    train_df = train_df.merge(res_counts, on='resource_id', how='left')
    test_df = test_df.merge(res_counts, on='resource_id', how='left')

    # Log transform counts to reduce skew
    for col in ['rr_count', 'mgr_count', 'role_count', 'res_count']:
        train_df[col] = np.log1p(train_df[col])
        test_df[col] = np.log1p(test_df[col])

    # Fill any missing values with -1 (shouldn't happen but safe)
    train_df = train_df.fillna(-1)
    test_df = test_df.fillna(-1)

    # Drop helper columns that are not needed for the model
    train_df = train_df.drop(columns=['role_resource', 'approved'])
    test_df = test_df.drop(columns=['role_resource'])

    return train_df, test_df


def sample(X_train, y_train, X_test, max_rows):
    # Use a single view with all rows
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred
