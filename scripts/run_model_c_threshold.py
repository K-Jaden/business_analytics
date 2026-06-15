"""
Model C 변형 — SHAP "누적 중요도 임계값" 기반 채널 선택 (수치 기준 도입)

기존 run_model_c.py 는 N_TOP=2 (상위 2개)를 임의로 고정 → "왜 2개냐?" 비판 가능.
본 스크립트는 선행연구(누적 중요도 cumulative importance) 방식으로 N을 데이터가 정하게 함:

  각 fold(train)에서:
    1. Model A(전채널) 훈련 → 채널별 mean|SHAP| 산출 (price-lag 제외, 채널만)
    2. 채널 SHAP를 내림차순 정렬 후 누적합 / 전체합 = 누적 비율
    3. 누적 비율이 THRESHOLD(기본 0.80) 이상이 될 때까지 채널을 채택
       → 자산·fold마다 선택 개수 N이 자동으로 달라짐 (임의의 k 없음)
    4. 선택 채널 + price lag 으로 Model C 훈련
    5. test fold AUC + Brier 오차 → DM 검정 (vs Model A)

비교용으로 "평균 이상(above-mean)" 규칙도 함께 산출.

Usage : python scripts/run_model_c_threshold.py [threshold]
Output: results/model_c_threshold_results.csv
        results/model_c_channel_selection.csv  (자산별 글로벌 SHAP 기준 선택 요약)
"""

import sys
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
THRESHOLD  = float(sys.argv[1]) if len(sys.argv) > 1 else 0.80  # 누적 SHAP 임계값

CH_LABEL = {
    'score_ch1': 'CH1', 'score_ch2': 'CH2', 'score_ch3': 'CH3',
    'score_ch4': 'CH4', 'score_ch5': 'CH5',
}


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
    nw_var = gamma0 + 2.0 * gamma1
    dm_stat = mean_d / np.sqrt(max(nw_var / n, 1e-12))
    p_val = 2.0 * (1.0 - stats.t.cdf(abs(dm_stat), df=n - 1))
    if p_val < 0.10 and dm_stat > 0:
        verdict = 'A better'
    elif p_val < 0.10 and dm_stat < 0:
        verdict = 'C better'
    else:
        verdict = 'no diff'
    return round(dm_stat, 3), round(p_val, 4), verdict


def select_by_cumulative(ch_shap: dict, threshold: float):
    """채널 SHAP 딕셔너리 → 누적 비율 threshold 이상이 될 때까지 채택."""
    total = sum(ch_shap.values())
    if total <= 0:
        # 전 채널 기여 0 → 가장 큰 1개만(사실상 무의미) 채택
        return [max(ch_shap, key=ch_shap.get)]
    ordered = sorted(ch_shap, key=ch_shap.get, reverse=True)
    cum, picked = 0.0, []
    for ch in ordered:
        picked.append(ch)
        cum += ch_shap[ch] / total
        if cum >= threshold:
            break
    return picked


def select_above_mean(ch_shap: dict):
    """채널 SHAP가 '채널 평균 중요도' 이상인 채널만 채택 (sklearn SelectFromModel 기본값='mean')."""
    total = sum(ch_shap.values())
    if total <= 0:
        return [max(ch_shap, key=ch_shap.get)]
    mean_imp = total / len(ch_shap)
    picked = [c for c in sorted(ch_shap, key=ch_shap.get, reverse=True)
              if ch_shap[c] >= mean_imp]
    return picked or [max(ch_shap, key=ch_shap.get)]


def _fit_eval(X_tr, X_te, y_tr, y_te, feats, params):
    clf = XGBClassifier(**params)
    clf.fit(X_tr[feats], y_tr)
    prob = clf.predict_proba(X_te[feats])[:, 1]
    auc = roc_auc_score(y_te, prob)
    err = (y_te.values - prob) ** 2
    return auc, err


def run_model_c(data, y, params, threshold):
    """동일 fold에서 두 선택 규칙(누적 임계 / 평균 이상)을 함께 평가."""
    tscv = TimeSeriesSplit(n_splits=N_SPLITS)
    res = {
        'cum':  dict(auc_c=[], err_a=[], err_c=[], ch=[], n=[]),
        'mean': dict(auc_c=[], err_a=[], err_c=[], ch=[], n=[]),
    }
    aucs_a = []

    for fold, (train_idx, test_idx) in enumerate(tscv.split(data)):
        X_tr, X_te = data.iloc[train_idx], data.iloc[test_idx]
        y_tr, y_te = y.iloc[train_idx], y.iloc[test_idx]

        clf_a = XGBClassifier(**params)
        clf_a.fit(X_tr[ALL_FEATS], y_tr)
        explainer = shap.TreeExplainer(clf_a)
        mean_abs = np.abs(explainer.shap_values(X_tr[ALL_FEATS])).mean(axis=0)
        ch_shap = dict(zip(CHANNELS, mean_abs[len(PRICE_LAGS):]))

        if len(np.unique(y_te)) < 2:
            continue
        prob_a = clf_a.predict_proba(X_te[ALL_FEATS])[:, 1]
        auc_a = roc_auc_score(y_te, prob_a)
        err_a = (y_te.values - prob_a) ** 2
        aucs_a.append(auc_a)

        for key, selector in (('cum', lambda d: select_by_cumulative(d, threshold)),
                              ('mean', select_above_mean)):
            sel = selector(ch_shap)
            auc_c, err_c = _fit_eval(X_tr, X_te, y_tr, y_te, PRICE_LAGS + sel, params)
            res[key]['auc_c'].append(auc_c)
            res[key]['err_a'].extend(err_a)
            res[key]['err_c'].extend(err_c)
            res[key]['ch'].extend(sel)
            res[key]['n'].append(len(sel))

    auc_a_mean = np.mean(aucs_a) if aucs_a else np.nan
    out = {}
    for key in ('cum', 'mean'):
        r = res[key]
        out[key] = (
            auc_a_mean,
            np.mean(r['auc_c']) if r['auc_c'] else np.nan,
            np.array(r['err_a']), np.array(r['err_c']),
            r['ch'], r['n'],
        )
    return out


def global_shap_selection(data, y, params, threshold):
    """전체 데이터 기준(글로벌) SHAP → 슬라이드/논문용 자산별 대표 선택."""
    clf = XGBClassifier(**params)
    clf.fit(data[ALL_FEATS], y)
    explainer = shap.TreeExplainer(clf)
    sv = explainer.shap_values(data[ALL_FEATS])
    mean_abs = np.abs(sv).mean(axis=0)
    ch_shap = dict(zip(CHANNELS, mean_abs[len(PRICE_LAGS):]))
    total = sum(ch_shap.values())
    mean_imp = total / len(CHANNELS)
    cum_sel = select_by_cumulative(ch_shap, threshold)
    above_mean = [c for c in sorted(ch_shap, key=ch_shap.get, reverse=True)
                  if ch_shap[c] >= mean_imp]
    return ch_shap, total, cum_sel, above_mean


def main():
    df = pd.read_csv('data/processed/panel_monthly_scaled.csv')
    params_df = pd.read_csv('results/xgboost_tuned_params.csv')

    print(f"누적 SHAP 임계값 = {THRESHOLD:.0%}\n")
    results, sel_rows = [], []

    for asset in ['sneakers', 'cards', 'lego']:
        print(f"{'='*54}  {asset.upper()}")
        data = (
            df[df['asset_type'] == asset]
            .sort_values('year_month')
            .dropna(subset=PRICE_LAGS + [TARGET])
            .reset_index(drop=True)
        )
        y = data[TARGET]
        params = load_params(asset, params_df)

        out = run_model_c(data, y, params, THRESHOLD)
        ch_shap, total, cum_sel, above_mean = global_shap_selection(data, y, params, THRESHOLD)

        print(f"  [글로벌] 채널 SHAP 합={total:.3f}")
        for c in sorted(ch_shap, key=ch_shap.get, reverse=True):
            pct = ch_shap[c] / total * 100 if total else 0
            print(f"      {CH_LABEL[c]}: {ch_shap[c]:.4f}  ({pct:4.1f}%)")

        for rule_key, rule_name, glob_sel in (
            ('cum',  f'누적{THRESHOLD:.0%}', cum_sel),
            ('mean', '평균이상(SelectFromModel 기본)', above_mean),
        ):
            auc_a, auc_c, err_a, err_c, ch_log, n_log = out[rule_key]
            dm_stat, dm_p, verdict = dm_test(err_c, err_a)
            freq = Counter(ch_log).most_common()
            freq_str = ', '.join(f"{CH_LABEL[c]}({n}/{N_SPLITS})" for c, n in freq)
            print(f"  ── 규칙[{rule_name}]  글로벌선택={[CH_LABEL[c] for c in glob_sel]}")
            print(f"     fold별 채널수={n_log} (평균 {np.mean(n_log):.1f})  빈도: {freq_str}")
            print(f"     A AUC={auc_a:.4f}  C AUC={auc_c:.4f}  DM p={dm_p}  -> {verdict}")
            results.append({
                'asset_type': asset, 'rule': rule_name, 'threshold': THRESHOLD,
                'auc_A': round(auc_a, 4), 'auc_C': round(auc_c, 4),
                'delta_auc': round(auc_c - auc_a, 4),
                'dm_stat': dm_stat, 'dm_p': dm_p, 'verdict': verdict,
                'n_min': min(n_log), 'n_max': max(n_log),
                'n_mean': round(float(np.mean(n_log)), 2),
                'global_sel': '+'.join(CH_LABEL[c] for c in glob_sel),
                'fold_freq': freq_str,
            })
        print()
        sel_rows.append({
            'asset_type': asset,
            'channel_shap_total': round(total, 4),
            'cum80_channels': '+'.join(CH_LABEL[c] for c in cum_sel),
            'cum80_n': len(cum_sel),
            'above_mean_channels': '+'.join(CH_LABEL[c] for c in above_mean),
            'above_mean_n': len(above_mean),
            **{CH_LABEL[c]: round(ch_shap[c], 4) for c in CHANNELS},
        })

    out = pd.DataFrame(results)
    out.to_csv('results/model_c_threshold_results.csv', index=False, encoding='utf-8-sig')
    sel = pd.DataFrame(sel_rows)
    sel.to_csv('results/model_c_channel_selection.csv', index=False, encoding='utf-8-sig')
    print("저장: results/model_c_threshold_results.csv")
    print("저장: results/model_c_channel_selection.csv")
    print()
    print(out[['asset_type', 'rule', 'auc_A', 'auc_C', 'delta_auc', 'dm_p', 'verdict',
               'n_mean', 'global_sel']].to_string(index=False))


if __name__ == '__main__':
    main()
