import pandas as pd
import numpy as np
import xgboost as xgb
from sklearn.model_selection import TimeSeriesSplit
from scipy.stats import t

def custom_dm_test(errors_A, errors_B, h=1, power=2):
    loss_A = np.abs(errors_A)**power
    loss_B = np.abs(errors_B)**power
    d_t = loss_A - loss_B
    mean_d = np.mean(d_t)
    T = len(d_t)
    gamma_0 = np.mean((d_t - mean_d)**2)
    gamma_k_sum = 0.0
    for k in range(1, h + 1):
        gamma_k = np.mean((d_t[k:] - mean_d) * (d_t[:-k] - mean_d))
        weight = 1 - (k / (h + 1))
        gamma_k_sum += weight * gamma_k
    long_run_variance = gamma_0 + 2 * gamma_k_sum
    if long_run_variance <= 0:
        return np.nan, np.nan
    dm_stat = mean_d / np.sqrt(long_run_variance / T)
    p_value = 2 * (1 - t.cdf(np.abs(dm_stat), df=T - 1))
    return dm_stat, p_value

df = pd.read_csv('DM/xgboost_data.csv')
df = df.dropna(subset=['price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']).copy()

model_a_features = [
    'price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
    'score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5'
]
model_b_features = {
    'sneakers': ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch3', 'score_ch4'],
    'cards':    ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch1'],
    'lego':     ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch1'],
}

tscv = TimeSeriesSplit(n_splits=4, test_size=6)

# A안: 마지막 폴드만 (원본 방식)
plan_a_results = []
# B안: 전체 폴드 합산
plan_b_results = []

for item_id in df['item_id'].unique():
    subset = df[df['item_id'] == item_id].copy()
    subset = subset.sort_values('year_month').reset_index(drop=True)
    asset_type = subset['asset_type'].iloc[0]

    X_a = subset[model_a_features]
    X_b = subset[model_b_features[asset_type]]
    y = subset['mean_price']

    all_actual, all_pred_a, all_pred_b = [], [], []
    last_actual, last_pred_a, last_pred_b = None, None, None

    for train_idx, test_idx in tscv.split(subset):
        if len(test_idx) == 0:
            continue
        X_train_a, X_test_a = X_a.iloc[train_idx], X_a.iloc[test_idx]
        X_train_b, X_test_b = X_b.iloc[train_idx], X_b.iloc[test_idx]
        y_train, y_test = y.iloc[train_idx], y.iloc[test_idx]

        model_a = xgb.XGBRegressor(n_estimators=100, max_depth=3, learning_rate=0.05,
                                   random_state=42, eval_metric='rmse', verbosity=0)
        model_a.fit(X_train_a, y_train)
        model_b = xgb.XGBRegressor(n_estimators=100, max_depth=3, learning_rate=0.05,
                                   random_state=42, eval_metric='rmse', verbosity=0)
        model_b.fit(X_train_b, y_train)

        pred_a = model_a.predict(X_test_a)
        pred_b = model_b.predict(X_test_b)

        all_actual.extend(y_test.values)
        all_pred_a.extend(pred_a)
        all_pred_b.extend(pred_b)

        last_actual = y_test.values
        last_pred_a = pred_a
        last_pred_b = pred_b

    def interpret(dm_stat, p_value):
        if np.isnan(dm_stat):
            return "검정 불가"
        if p_value < 0.05:
            return "Model A 우수" if dm_stat < 0 else "Model B 우수"
        if p_value < 0.10:
            return "Model A 우세(경계)" if dm_stat < 0 else "Model B 우세(경계)"
        return "유의미한 차이 없음"

    # A안
    ea, eb = last_actual - last_pred_a, last_actual - last_pred_b
    dm_a, p_a = custom_dm_test(ea, eb)
    plan_a_results.append({
        'item_id': item_id, 'T': len(last_actual),
        'DM_stat': round(dm_a, 3), 'p': round(p_a, 4), 'result': interpret(dm_a, p_a)
    })

    # B안
    arr_actual = np.array(all_actual)
    arr_a = np.array(all_pred_a)
    arr_b = np.array(all_pred_b)
    ea2, eb2 = arr_actual - arr_a, arr_actual - arr_b
    dm_b, p_b = custom_dm_test(ea2, eb2)
    plan_b_results.append({
        'item_id': item_id, 'T': len(arr_actual),
        'DM_stat': round(dm_b, 3), 'p': round(p_b, 4), 'result': interpret(dm_b, p_b)
    })

df_a = pd.DataFrame(plan_a_results)
df_b = pd.DataFrame(plan_b_results)

print("=" * 65)
print("A안: 마지막 폴드만 (T=6)")
print("=" * 65)
print(df_a.to_string(index=False))

print("\n" + "=" * 65)
print("B안: 전체 폴드 합산 (T=24)")
print("=" * 65)
print(df_b.to_string(index=False))

df_a.to_csv('DM/dm_results_plan_a.csv', index=False)
df_b.to_csv('DM/dm_results_plan_b.csv', index=False)
print("\n결과 저장 완료: DM/dm_results_plan_a.csv, DM/dm_results_plan_b.csv")
