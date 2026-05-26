"""STEP 09b — XGBoost 하이퍼파라미터 튜닝 + DM 재검증

목적:
  기존 `DM/dm_plan_b.py`는 고정 하이퍼파라미터(n_estimators=100, max_depth=3,
  learning_rate=0.05)를 사용. 튜닝되지 않은 비교는 Model A·B 사이의 진정한
  성능 차이를 측정하지 못할 수 있음 (한쪽이 default에 우연히 더 잘 맞을 위험).
  본 스크립트는 Optuna로 자산별 하이퍼파라미터를 탐색한 뒤,
  Model A, B-original, B-DH 세 가지를 동일 조건에서 DM 검정 재실행.

튜닝 단위: 자산별 (아이템 단위는 N=48로 과적합 위험).
  자산 내 모든 아이템의 평균 RMSE를 목적함수로 minimize.

비교 모델:
  Model A          전채널 (CH1~5 + 통제 4)        9 피처
  Model B-original 기존 자산-평균 Granger 선별   자산별 다름
  Model B-DH       DH 패널 검정 결과 반영        자산별 다름

산출:
  results/xgboost_tuned_params.csv    자산별 튜닝된 하이퍼파라미터
  results/xgboost_tuned_rmse.csv      아이템×모델별 평균 RMSE
  results/dm_tuned.csv                튜닝 후 DM 검정 결과
"""
import os, warnings, json
import numpy as np
import pandas as pd
import optuna
import xgboost as xgb
from sklearn.model_selection import TimeSeriesSplit
from scipy.stats import t

warnings.filterwarnings('ignore')
optuna.logging.set_verbosity(optuna.logging.WARNING)
os.makedirs('results', exist_ok=True)

RANDOM_STATE = 42
N_TRIALS     = 60               # 자산당 trial (60 × 3 = 180회)
N_SPLITS     = 4
TEST_SIZE    = 6

# ── 데이터 ─────────────────────────────────────────────────────────────────────
df = pd.read_csv('DM/xgboost_data.csv').dropna(
    subset=['price_chg_lag1','price_chg_lag2','price_chg_lag3']).copy()

CTRL = ['price_vs_ma3','price_chg_lag1','price_chg_lag2','price_chg_lag3']
ALL_CH = ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']

FEATURES_A = CTRL + ALL_CH

# Model B-original (기존 자산-평균 Granger raw p<0.05 기준)
B_ORIGINAL = {
    'sneakers': CTRL + ['score_ch3','score_ch4'],
    'cards':    CTRL + ['score_ch1'],
    'lego':     CTRL + ['score_ch1'],     # MARGINAL이라 기존 분석에서 포함
}
# Model B-DH (DH 패널 검정 p<0.05 기준; MARGINAL은 제외)
B_DH = {
    'sneakers': CTRL + ['score_ch4'],
    'cards':    CTRL + ['score_ch1','score_ch2'],
    'lego':     CTRL,                      # DH에서 모든 채널 비유의
}

ASSETS = ['sneakers','cards','lego']

# ── 자산별 RMSE 평가 함수 ──────────────────────────────────────────────────────
def asset_mean_rmse(params, features, asset):
    """주어진 하이퍼파라미터+피처로 자산 내 모든 아이템 평균 RMSE."""
    sub = df[df['asset_type'] == asset]
    rmses = []
    for it in sub['item_id'].unique():
        d = sub[sub['item_id'] == it].sort_values('year_month').reset_index(drop=True)
        X, y = d[features], d['mean_price']
        tscv = TimeSeriesSplit(n_splits=N_SPLITS, test_size=TEST_SIZE)
        errs = []
        for tr, te in tscv.split(d):
            if len(te) == 0:
                continue
            m = xgb.XGBRegressor(**params, random_state=RANDOM_STATE,
                                 verbosity=0, eval_metric='rmse')
            m.fit(X.iloc[tr], y.iloc[tr])
            pred = m.predict(X.iloc[te])
            errs.extend((y.iloc[te].values - pred) ** 2)
        if errs:
            rmses.append(np.sqrt(np.mean(errs)))
    return float(np.mean(rmses))

# ── Optuna 목적함수 ────────────────────────────────────────────────────────────
def make_objective(features, asset):
    def objective(trial):
        params = {
            'n_estimators'    : trial.suggest_int('n_estimators', 50, 300, step=50),
            'max_depth'       : trial.suggest_int('max_depth', 2, 6),
            'learning_rate'   : trial.suggest_float('learning_rate', 0.01, 0.2, log=True),
            'subsample'       : trial.suggest_float('subsample', 0.6, 1.0),
            'colsample_bytree': trial.suggest_float('colsample_bytree', 0.6, 1.0),
            'min_child_weight': trial.suggest_int('min_child_weight', 1, 5),
            'reg_lambda'      : trial.suggest_float('reg_lambda', 0.0, 5.0),
            'reg_alpha'       : trial.suggest_float('reg_alpha', 0.0, 1.0),
        }
        return asset_mean_rmse(params, features, asset)
    return objective

# ── DM 검정 (기존 dm_plan_b.py 동일) ──────────────────────────────────────────
def dm_test(errors_A, errors_B, h=1):
    loss_A = errors_A ** 2
    loss_B = errors_B ** 2
    d_t = loss_A - loss_B
    mean_d = float(np.mean(d_t))
    T = len(d_t)
    gamma_0 = float(np.mean((d_t - mean_d) ** 2))
    gamma_k_sum = 0.0
    for k in range(1, h + 1):
        gamma_k = float(np.mean((d_t[k:] - mean_d) * (d_t[:-k] - mean_d)))
        gamma_k_sum += (1 - k / (h + 1)) * gamma_k
    lrv = gamma_0 + 2 * gamma_k_sum
    if lrv <= 0:
        return np.nan, np.nan
    dm_stat = mean_d / np.sqrt(lrv / T)
    p_val = 2 * (1 - t.cdf(abs(dm_stat), df=T - 1))
    return dm_stat, p_val

# ── 메인: 자산별 Model A·B-original·B-DH 튜닝 ────────────────────────────────
tuned_params = {}         # (asset, model_name) -> dict
tuned_rmse   = {}         # (asset, model_name) -> mean RMSE

models_to_tune = [
    ('A',          {a: FEATURES_A for a in ASSETS}),
    ('B-original', B_ORIGINAL),
    ('B-DH',       B_DH),
]

for model_name, feat_map in models_to_tune:
    print(f'\n=== 튜닝: Model {model_name} ===')
    for asset in ASSETS:
        feats = feat_map[asset]
        study = optuna.create_study(
            direction='minimize',
            sampler=optuna.samplers.TPESampler(seed=RANDOM_STATE))
        study.optimize(make_objective(feats, asset), n_trials=N_TRIALS, show_progress_bar=False)
        tuned_params[(asset, model_name)] = study.best_params
        tuned_rmse[(asset, model_name)]   = study.best_value
        print(f'  {asset:10s} best RMSE={study.best_value:.3f}  '
              f'lr={study.best_params["learning_rate"]:.3f}  '
              f'depth={study.best_params["max_depth"]}  '
              f'n_est={study.best_params["n_estimators"]}')

# 튜닝 결과 저장
param_rows = []
for (asset, mname), params in tuned_params.items():
    row = {'asset_type': asset, 'model': mname, 'tuned_rmse': round(tuned_rmse[(asset,mname)], 3)}
    row.update({k: (round(v,4) if isinstance(v,float) else v) for k, v in params.items()})
    param_rows.append(row)
pd.DataFrame(param_rows).to_csv('results/xgboost_tuned_params.csv', index=False)
print(f'\nSaved → results/xgboost_tuned_params.csv')

# ── 튜닝된 파라미터로 아이템별 예측 + DM ──────────────────────────────────────
def fit_predict(asset, item, features, params):
    d = df[df['item_id'] == item].sort_values('year_month').reset_index(drop=True)
    X, y = d[features], d['mean_price']
    tscv = TimeSeriesSplit(n_splits=N_SPLITS, test_size=TEST_SIZE)
    actuals, preds = [], []
    for tr, te in tscv.split(d):
        if len(te) == 0:
            continue
        m = xgb.XGBRegressor(**params, random_state=RANDOM_STATE,
                             verbosity=0, eval_metric='rmse')
        m.fit(X.iloc[tr], y.iloc[tr])
        actuals.extend(y.iloc[te].values)
        preds.extend(m.predict(X.iloc[te]))
    return np.array(actuals), np.array(preds)

dm_rows, rmse_rows = [], []
for item in df['item_id'].unique():
    asset = df[df['item_id'] == item]['asset_type'].iloc[0]
    feats_A = FEATURES_A
    feats_Bo = B_ORIGINAL[asset]
    feats_Bd = B_DH[asset]
    params_A  = tuned_params[(asset, 'A')]
    params_Bo = tuned_params[(asset, 'B-original')]
    params_Bd = tuned_params[(asset, 'B-DH')]

    y_act, pred_A  = fit_predict(asset, item, feats_A,  params_A)
    _,     pred_Bo = fit_predict(asset, item, feats_Bo, params_Bo)
    _,     pred_Bd = fit_predict(asset, item, feats_Bd, params_Bd)

    eA  = y_act - pred_A
    eBo = y_act - pred_Bo
    eBd = y_act - pred_Bd

    rmse_A  = float(np.sqrt(np.mean(eA**2)))
    rmse_Bo = float(np.sqrt(np.mean(eBo**2)))
    rmse_Bd = float(np.sqrt(np.mean(eBd**2)))
    rmse_rows.append({'item_id': item, 'asset_type': asset,
                      'rmse_A': round(rmse_A,2),
                      'rmse_B_original': round(rmse_Bo,2),
                      'rmse_B_DH': round(rmse_Bd,2),
                      'T': len(y_act)})

    # DM: A vs B-original, A vs B-DH, B-original vs B-DH
    for label, e1, e2 in [
        ('A_vs_B-original', eA, eBo),
        ('A_vs_B-DH',       eA, eBd),
        ('B-original_vs_B-DH', eBo, eBd),
    ]:
        s, p = dm_test(e1, e2)
        # 부호 해석: dm<0 → 첫 번째가 더 작은 손실 → 첫 번째 우수
        if np.isnan(s):
            interp = '검정 불가'
        elif p < 0.05:
            interp = f'{label.split("_vs_")[0]} 우수' if s < 0 else f'{label.split("_vs_")[1]} 우수'
        elif p < 0.10:
            interp = '경계'
        else:
            interp = '차이 없음'
        dm_rows.append({'item_id': item, 'asset_type': asset,
                        'comparison': label,
                        'DM_stat': round(s,3) if not np.isnan(s) else None,
                        'p': round(p,4) if not np.isnan(p) else None,
                        'verdict': interp})

df_rmse = pd.DataFrame(rmse_rows)
df_dm   = pd.DataFrame(dm_rows)
df_rmse.to_csv('results/xgboost_tuned_rmse.csv', index=False)
df_dm.to_csv('results/dm_tuned.csv', index=False)
print(f'Saved → results/xgboost_tuned_rmse.csv  ({len(df_rmse)} items)')
print(f'Saved → results/dm_tuned.csv  ({len(df_dm)} rows)')

# ── 요약 ──────────────────────────────────────────────────────────────────────
print('\n=== 튜닝 후 자산별 평균 RMSE 비교 ===')
print(df_rmse.groupby('asset_type')[['rmse_A','rmse_B_original','rmse_B_DH']].mean().round(2))

print('\n=== DM 판정 분포 ===')
print(df_dm.groupby(['comparison','verdict']).size().unstack(fill_value=0))

print('\n=== A vs B-original 비교 (튜닝 vs 기존 결과) ===')
sub = df_dm[df_dm['comparison'] == 'A_vs_B-original']
print(sub[['item_id','asset_type','DM_stat','p','verdict']].to_string(index=False))
