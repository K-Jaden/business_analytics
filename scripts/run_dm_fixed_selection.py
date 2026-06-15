"""21쪽(shap_summary.csv) Model A SHAP의 above-mean 선택으로 고정한 채 DM 검정.
선택: sneakers CH1+CH5+CH4 / cards CH4+CH1 / lego CH4+CH1  (평균 임계값, 21쪽 값 기준)
평가: 튜닝 모델 + TimeSeriesSplit(5), Brier 오차로 DM (vs Model A).
출력: results/dm_fixed_selection.csv
"""
import pandas as pd
import numpy as np
from xgboost import XGBClassifier
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import roc_auc_score
from scipy import stats

PRICE = ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']
ALL = PRICE + ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
SEL = {
    'sneakers': ['score_ch1', 'score_ch5', 'score_ch4'],
    'cards':    ['score_ch4', 'score_ch1'],
    'lego':     ['score_ch4', 'score_ch1'],
}
df = pd.read_csv('data/processed/panel_monthly_scaled.csv')
pp = pd.read_csv('results/xgboost_tuned_params.csv')


def params(a):
    r = pp[(pp.asset_type == a) & (pp.model == 'A')].iloc[0]
    return dict(n_estimators=int(r.n_estimators), max_depth=int(r.max_depth),
               learning_rate=float(r.learning_rate), subsample=float(r.subsample),
               colsample_bytree=float(r.colsample_bytree), min_child_weight=int(r.min_child_weight),
               reg_lambda=float(r.reg_lambda), reg_alpha=float(r.reg_alpha),
               eval_metric='logloss', random_state=42, verbosity=0)


def dm(ec, ea):
    n = min(len(ec), len(ea)); d = ec[:n] - ea[:n]
    md = np.mean(d); g0 = np.var(d, ddof=1)
    g1 = np.cov(d[1:], d[:-1])[0, 1] if n > 2 else 0.0
    stat = md / np.sqrt(max((g0 + 2 * g1) / n, 1e-12))
    return round(stat, 3), round(2 * (1 - stats.t.cdf(abs(stat), df=n - 1)), 4)


rows = []
for a in ['sneakers', 'cards', 'lego']:
    d = df[df.asset_type == a].sort_values('year_month').dropna(subset=PRICE + ['price_direction']).reset_index(drop=True)
    y = d['price_direction']; p = params(a)
    feats_c = PRICE + SEL[a]
    tscv = TimeSeriesSplit(n_splits=5)
    aa, ac, ea, ec = [], [], [], []
    for tr, te in tscv.split(d):
        if len(np.unique(y.iloc[te])) < 2:
            continue
        cA = XGBClassifier(**p).fit(d[ALL].iloc[tr], y.iloc[tr])
        cC = XGBClassifier(**p).fit(d[feats_c].iloc[tr], y.iloc[tr])
        pA = cA.predict_proba(d[ALL].iloc[te])[:, 1]
        pC = cC.predict_proba(d[feats_c].iloc[te])[:, 1]
        aa.append(roc_auc_score(y.iloc[te], pA)); ac.append(roc_auc_score(y.iloc[te], pC))
        ea.extend((y.iloc[te].values - pA) ** 2); ec.extend((y.iloc[te].values - pC) ** 2)
    st, pv = dm(np.array(ec), np.array(ea))
    sel_lbl = '+'.join('CH' + c[-1] for c in SEL[a])
    rows.append(dict(asset_type=a, selected=sel_lbl, auc_A=round(np.mean(aa), 4),
                     auc_C=round(np.mean(ac), 4), dm_stat=st, dm_p=pv,
                     verdict='no diff' if pv >= 0.10 else ('A better' if st > 0 else 'C better')))
    print(f"{a:9} sel={sel_lbl:14} A={np.mean(aa):.4f} C={np.mean(ac):.4f} DM stat={st} p={pv}")

pd.DataFrame(rows).to_csv('results/dm_fixed_selection.csv', index=False, encoding='utf-8-sig')
print("저장: results/dm_fixed_selection.csv")
