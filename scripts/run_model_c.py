"""
Model C: SHAP 기반 채널 선택 (데이터 누출 없음 — fold별 train에서만 SHAP 계산)

각 TimeSeriesSplit fold에서:
  1. 전채널 Model A를 train fold로 훈련
  2. train fold의 SHAP 중요도로 상위 N_TOP개 채널 선택
  3. 선택된 채널 + price lag으로 Model C 훈련
  4. test fold에서 AUC + Brier 오차 계산

DM 검정: Model C vs Model A (Brier score 기준)
저장: results/model_c_results.csv
"""

import pandas as pd
import numpy as np
import shap
from collections import Counter
from xgboost import XGBClassifier
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import roc_auc_score
from scipy import stats

PRICE_LAGS = ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']
CHANNELS   = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
ALL_FEATS  = PRICE_LAGS + CHANNELS
TARGET     = 'price_direction'
N_SPLITS   = 5
N_TOP      = 2  # 전 자산 동일하게 상위 2개 채널 선택


def load_params(asset_type, params_df):
    row = params_df[
        (params_df['asset_type'] == asset_type) & (params_df['model'] == 'A')
    ].iloc[0]
    return {
        'n_estimators':     int(row['n_estimators']),
        'max_depth':        int(row['max_depth']),
        'learning_rate':    float(row['learning_rate']),
        'subsample':        float(row['subsample']),
        'colsample_bytree': float(row['colsample_bytree']),
        'min_child_weight': int(row['min_child_weight']),
        'reg_lambda':       float(row['reg_lambda']),
        'reg_alpha':        float(row['reg_alpha']),
        'eval_metric':      'logloss',
        'random_state':     42,
        'verbosity':        0,
    }


def dm_test(e_c, e_a):
    n = min(len(e_c), len(e_a))
    d = e_c[:n] - e_a[:n]
    mean_d = np.mean(d)
    gamma0 = np.var(d, ddof=1)
    gamma1 = np.cov(d[1:], d[:-1])[0, 1] if n > 2 else 0.0
    nw_var  = gamma0 + 2.0 * gamma1
    dm_stat = mean_d / np.sqrt(max(nw_var / n, 1e-12))
    p_val   = 2.0 * (1.0 - stats.t.cdf(abs(dm_stat), df=n - 1))

    if p_val < 0.10 and dm_stat > 0:
        verdict = 'A better'
    elif p_val < 0.10 and dm_stat < 0:
        verdict = 'C better'
    else:
        verdict = 'no diff'

    return round(dm_stat, 3), round(p_val, 4), verdict


def run_model_c(data, y, params):
    tscv = TimeSeriesSplit(n_splits=N_SPLITS)
    aucs_a, aucs_c = [], []
    errors_a, errors_c = [], []
    ch_log = []

    for fold, (train_idx, test_idx) in enumerate(tscv.split(data)):
        X_tr = data.iloc[train_idx]
        X_te = data.iloc[test_idx]
        y_tr = y.iloc[train_idx]
        y_te = y.iloc[test_idx]

        # ── Step 1: Model A (전채널) train fold 훈련 ──────────
        clf_a = XGBClassifier(**params)
        clf_a.fit(X_tr[ALL_FEATS], y_tr)

        # ── Step 2: SHAP — train fold에서만 계산 (누출 없음) ──
        explainer = shap.TreeExplainer(clf_a)
        shap_vals = explainer.shap_values(X_tr[ALL_FEATS])
        mean_abs  = np.abs(shap_vals).mean(axis=0)

        # price lag SHAP 제외, 채널 SHAP만 추출
        ch_shap = dict(zip(CHANNELS, mean_abs[len(PRICE_LAGS):]))
        top_channels = sorted(ch_shap, key=ch_shap.get, reverse=True)[:N_TOP]
        ch_log.append({'fold': fold, 'selected': top_channels, 'shap': ch_shap})

        # ── Step 3: Model C (price lag + SHAP 상위 채널) ─────
        feats_c = PRICE_LAGS + top_channels
        clf_c = XGBClassifier(**params)
        clf_c.fit(X_tr[feats_c], y_tr)

        # ── Step 4: 평가 ─────────────────────────────────────
        if len(np.unique(y_te)) < 2:
            continue

        prob_a = clf_a.predict_proba(X_te[ALL_FEATS])[:, 1]
        prob_c = clf_c.predict_proba(X_te[feats_c])[:, 1]

        aucs_a.append(roc_auc_score(y_te, prob_a))
        aucs_c.append(roc_auc_score(y_te, prob_c))
        errors_a.extend((y_te.values - prob_a) ** 2)
        errors_c.extend((y_te.values - prob_c) ** 2)

    return (
        np.mean(aucs_a) if aucs_a else np.nan,
        np.mean(aucs_c) if aucs_c else np.nan,
        np.array(errors_a),
        np.array(errors_c),
        ch_log,
    )


def main():
    df        = pd.read_csv('data/processed/panel_monthly_scaled.csv')
    params_df = pd.read_csv('results/xgboost_tuned_params.csv')

    results = []

    for asset in ['sneakers', 'cards', 'lego']:
        print(f"\n{'='*50}  {asset.upper()}")

        data = (
            df[df['asset_type'] == asset]
            .sort_values('year_month')
            .dropna(subset=PRICE_LAGS + [TARGET])
            .reset_index(drop=True)
        )
        y      = data[TARGET]
        params = load_params(asset, params_df)

        auc_a, auc_c, err_a, err_c, ch_log = run_model_c(data, y, params)
        dm_stat, dm_p, verdict = dm_test(err_c, err_a)

        # fold별 선택된 채널 빈도 요약
        all_selected = [ch for fold in ch_log for ch in fold['selected']]
        freq = Counter(all_selected).most_common()
        freq_str = ', '.join(f"{ch}({cnt}/{N_SPLITS})" for ch, cnt in freq)

        print(f"  Model A  AUC={auc_a:.4f}")
        print(f"  Model C  AUC={auc_c:.4f}  |  DM stat={dm_stat}  p={dm_p}  → {verdict}")
        print(f"  선택 채널 빈도: {freq_str}")

        results.append({
            'asset_type':   asset,
            'auc_A':        round(auc_a, 4),
            'auc_C':        round(auc_c, 4),
            'delta_auc':    round(auc_c - auc_a, 4),
            'dm_stat':      dm_stat,
            'dm_p':         dm_p,
            'verdict':      verdict,
            'top_channels': freq_str,
            'n_top':        N_TOP,
        })

    out = pd.DataFrame(results)
    out.to_csv('results/model_c_results.csv', index=False, encoding='utf-8-sig')
    print(f"\n저장: results/model_c_results.csv")
    print(out[['asset_type', 'auc_A', 'auc_C', 'delta_auc', 'dm_p', 'verdict', 'top_channels']].to_string(index=False))


if __name__ == '__main__':
    main()
