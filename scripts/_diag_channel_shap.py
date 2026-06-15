"""진단: (a) Model A 내부 채널 SHAP vs (b) 각 채널 단독(가격래그+1채널) SHAP 비교."""
import pandas as pd
import numpy as np
import shap
from xgboost import XGBClassifier

PRICE = ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']
CH = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
LBL = {c: f'CH{i+1}' for i, c in enumerate(CH)}

df = pd.read_csv('data/processed/panel_monthly_scaled.csv')
pp = pd.read_csv('results/xgboost_tuned_params.csv')


def params(a):
    r = pp[(pp.asset_type == a) & (pp.model == 'A')].iloc[0]
    return dict(n_estimators=int(r.n_estimators), max_depth=int(r.max_depth),
               learning_rate=float(r.learning_rate), subsample=float(r.subsample),
               colsample_bytree=float(r.colsample_bytree), min_child_weight=int(r.min_child_weight),
               reg_lambda=float(r.reg_lambda), reg_alpha=float(r.reg_alpha),
               eval_metric='logloss', random_state=42, verbosity=0)


def msh(clf, X):
    return np.abs(shap.TreeExplainer(clf).shap_values(X)).mean(axis=0)


for a in ['sneakers', 'cards', 'lego']:
    d = df[df.asset_type == a].sort_values('year_month').dropna(subset=PRICE + ['price_direction']).reset_index(drop=True)
    y = d['price_direction']
    p = params(a)
    # (a) Model A 내부
    clfA = XGBClassifier(**p); clfA.fit(d[PRICE + CH], y)
    shA = dict(zip(CH, msh(clfA, d[PRICE + CH])[len(PRICE):]))
    # (b) 각 채널 단독
    solo = {}
    for c in CH:
        clf = XGBClassifier(**p); clf.fit(d[PRICE + [c]], y)
        solo[c] = msh(clf, d[PRICE + [c]])[len(PRICE)]
    print(f'\n=== {a.upper()} ===')
    print(f'{"채널":5} {"ModelA내부":>12} {"단독":>10}')
    for c in CH:
        print(f'{LBL[c]:5} {shA[c]:12.4f} {solo[c]:10.4f}')
