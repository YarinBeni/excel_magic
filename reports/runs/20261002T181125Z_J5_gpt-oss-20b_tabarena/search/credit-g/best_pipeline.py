import numpy as np
import pandas as pd

MODEL_KWARGS = {}


def preprocess(X_train, y_train, X_test):
    X_train = X_train.drop(columns=['good_or_bad_customer'], errors='ignore').copy()
    X_test = X_test.drop(columns=['good_or_bad_customer'], errors='ignore').copy()
    return X_train, y_train, X_test


def engineer(X_train, y_train, X_test):
    X_train = X_train.copy()
    X_test = X_test.copy()

    # numeric derived features
    X_train['credit_amount_log'] = np.log1p(X_train['credit_amount'])
    X_test['credit_amount_log'] = np.log1p(X_test['credit_amount'])

    X_train['credit_amount_per_duration'] = X_train['credit_amount'] / X_train['duration_months']
    X_test['credit_amount_per_duration'] = X_test['credit_amount'] / X_test['duration_months']

    X_train['credit_amount_per_people'] = X_train['credit_amount'] / X_train['people_liable']
    X_test['credit_amount_per_people'] = X_test['credit_amount'] / X_test['people_liable']

    X_train['credit_amount_per_age'] = X_train['credit_amount'] / X_train['age_years']
    X_test['credit_amount_per_age'] = X_test['credit_amount'] / X_test['age_years']

    X_train['installment_rate_times_amount'] = X_train['installment_rate_percent'] * X_train['credit_amount']
    X_test['installment_rate_times_amount'] = X_test['installment_rate_percent'] * X_test['credit_amount']

    return X_train, X_test


def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]


def postprocess(pred, y_train_orig, X_test):
    return pred
