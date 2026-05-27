"""
Ablation study: P1 구성요소 기여도 + A-dropGranger

모델 목록 (asset별):
  A              : 가격래그 + CH1~5 전부 (기준 모델)
  Baseline       : 가격래그만 (채널 없음)
  Channels-only  : CH1~5만 (가격래그 없음)
  CH1-only ~ CH5-only : 가격래그 + 채널 1개
  A-dropGranger  : 가격래그 + Granger 비유의 채널만

평가: AUC (TimeSeriesSplit n=5)
DM검정: Brier score 기준, vs Model A
"""

import pandas as pd
import numpy as np
from xgboost import XGBClassifier
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import roc_auc_score
from scipy import stats

# ── 상수 ──────────────────────────────────────────────
PRICE_LAGS = ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']
CHANNELS   = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
TARGET     = 'price_direction'
N_SPLITS   = 5

# Granger 유의 채널 (CLAUDE.md 확정값 기준)
GRANGER_SIG = {
    'sneakers': ['score_ch1', 'score_ch3', 'score_ch4'],
    'cards':    ['score_ch1'],
    'lego':     [],
}

# ── 모델 구성 정의 ─────────────────────────────────────
def get_configs(asset_type: str) -> dict:
    sig     = GRANGER_SIG[asset_type]
    non_sig = [c for c in CHANNELS if c not in sig]

    configs = {
        'A':             PRICE_LAGS + CHANNELS,
        'Baseline':      PRICE_LAGS,
        'Channels-only': CHANNELS,
        'CH1-only':      PRICE_LAGS + ['score_ch1'],
        'CH2-only':      PRICE_LAGS + ['score_ch2'],
        'CH3-only':      PRICE_LAGS + ['score_ch3'],
        'CH4-only':      PRICE_LAGS + ['score_ch4'],
        'CH5-only':      PRICE_LAGS + ['score_ch5'],
        'A-dropGranger': PRICE_LAGS + non_sig,   # lego는 non_sig=CH1~5 → A와 동일
    }
    return configs

# ── XGBoost 파라미터 로드 (기존 Model A 튜닝값 재사용) ─
def load_params(asset_type: str, params_df: pd.DataFrame) -> dict:
    row = params_df[
        (params_df['asset_type'] == asset_type) & (params_df['model'] == 'A')
    ].iloc[0]
    return {
        'n_estimators':    int(row['n_estimators']),
        'max_depth':       int(row['max_depth']),
        'learning_rate':   float(row['learning_rate']),
        'subsample':       float(row['subsample']),
        'colsample_bytree':float(row['colsample_bytree']),
        'min_child_weight':int(row['min_child_weight']),
        'reg_lambda':      float(row['reg_lambda']),
        'reg_alpha':       float(row['reg_alpha']),
        'eval_metric':     'logloss',
        'random_state':    42,
        'verbosity':       0,
    }

# ── TimeSeriesSplit CV ─────────────────────────────────
def run_cv(X: pd.DataFrame, y: pd.Series, params: dict):
    """
    Returns:
        mean_auc : float
        oof_errors : np.ndarray  - Brier score per test sample (DM 검정용)
        test_indices : list      - 각 fold의 test index (정렬 일치 확인용)
    """
    tscv = TimeSeriesSplit(n_splits=N_SPLITS)
    aucs, oof_errors, test_indices = [], [], []

    for train_idx, test_idx in tscv.split(X):
        X_tr, X_te = X.iloc[train_idx], X.iloc[test_idx]
        y_tr, y_te = y.iloc[train_idx], y.iloc[test_idx]

        clf = XGBClassifier(**params)
        clf.fit(X_tr, y_tr)

        prob = clf.predict_proba(X_te)[:, 1]

        if len(np.unique(y_te)) > 1:
            aucs.append(roc_auc_score(y_te, prob))

        brier = (y_te.values - prob) ** 2
        oof_errors.extend(brier)
        test_indices.extend(test_idx.tolist())

    return np.mean(aucs) if aucs else np.nan, np.array(oof_errors), test_indices

# ── DM 검정 (Newey-West, 단측: ablation이 A보다 나쁜지) ─
def dm_test(e_ablation: np.ndarray, e_A: np.ndarray):
    """d > 0 이면 ablation 오차가 A보다 큼 → A가 더 좋음"""
    n = min(len(e_ablation), len(e_A))
    d = e_ablation[:n] - e_A[:n]
    mean_d = np.mean(d)

    gamma0 = np.var(d, ddof=1)
    gamma1 = np.cov(d[1:], d[:-1])[0, 1] if n > 2 else 0.0
    nw_var  = gamma0 + 2.0 * gamma1
    dm_stat = mean_d / np.sqrt(max(nw_var / n, 1e-12))
    p_val   = 2.0 * (1.0 - stats.t.cdf(abs(dm_stat), df=n - 1))

    if p_val < 0.10 and dm_stat > 0:
        verdict = 'A better'
    elif p_val < 0.10 and dm_stat < 0:
        verdict = 'ablation better'
    else:
        verdict = 'no diff'

    return round(dm_stat, 3), round(p_val, 4), verdict

# ── 메인 ──────────────────────────────────────────────
def main():
    df        = pd.read_csv('data/processed/panel_monthly_scaled.csv')
    params_df = pd.read_csv('results/xgboost_tuned_params.csv')

    all_results = []

    for asset in ['sneakers', 'cards', 'lego']:
        print(f"\n{'='*50}")
        print(f"  {asset.upper()}")
        print(f"{'='*50}")

        data = (
            df[df['asset_type'] == asset]
            .sort_values('year_month')
            .dropna(subset=PRICE_LAGS + [TARGET])
            .reset_index(drop=True)
        )
        y       = data[TARGET]
        params  = load_params(asset, params_df)
        configs = get_configs(asset)

        # Model A를 먼저 돌려서 기준 오차 확보
        _, errors_A, _ = run_cv(data[configs['A']], y, params)

        for model_name, features in configs.items():
            note = ''
            if model_name == 'A-dropGranger' and not GRANGER_SIG[asset]:
                note = 'Granger유의채널없음→A동일'

            auc, errors_abl, _ = run_cv(data[features], y, params)
            dm_stat, dm_p, verdict = dm_test(errors_abl, errors_A)

            row = {
                'asset_type': asset,
                'model':      model_name,
                'n_features': len(features),
                'features':   ' + '.join(features),
                'auc':        round(auc, 4),
                'dm_stat':    dm_stat,
                'dm_p':       dm_p,
                'vs_A':       verdict,
                'note':       note,
            }
            all_results.append(row)

            flag = '★' if model_name in ('Baseline', 'A-dropGranger') else ' '
            print(f"  {flag} {model_name:20s} | AUC={auc:.4f} | DM p={dm_p:.4f} | {verdict}  {note}")

    out = pd.DataFrame(all_results)
    out.to_csv('results/ablation_results.csv', index=False, encoding='utf-8-sig')
    print(f"\n\n저장 완료: results/ablation_results.csv")

    # 요약 출력
    print("\n[요약] Model A AUC 대비")
    for asset in ['sneakers', 'cards', 'lego']:
        sub = out[out['asset_type'] == asset][['model', 'auc', 'vs_A']]
        auc_A = sub[sub['model'] == 'A']['auc'].values[0]
        print(f"\n  {asset} (A baseline AUC={auc_A:.4f})")
        for _, r in sub.iterrows():
            delta = r['auc'] - auc_A
            sign  = '+' if delta >= 0 else ''
            print(f"    {r['model']:20s}  AUC={r['auc']:.4f}  ({sign}{delta:.4f})  {r['vs_A']}")

if __name__ == '__main__':
    main()
