"""4모델 비교 — Robustness Check (발표 3막용)

모델: XGBoost, CatBoost, RandomForest, ElasticNet(LogisticRegression)
피처: Model A (PRICE_LAGS + CH1~5 전부)
평가: AUC (TimeSeriesSplit n_splits=5, 자산 내 아이템 평균)
산출:
  results/model_comparison_auc.csv
  results/figures/model_comparison.png  (4모델 × 3자산 = 12막대)
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from xgboost import XGBClassifier
from catboost import CatBoostClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import roc_auc_score
import os, warnings

warnings.filterwarnings('ignore')
os.makedirs('results/figures', exist_ok=True)

PRICE_LAGS = ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']
CHANNELS   = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
FEATURES   = PRICE_LAGS + CHANNELS
TARGET     = 'price_direction'
N_SPLITS   = 5
ASSETS     = ['sneakers', 'cards', 'lego']

# ── XGBoost 튜닝 파라미터 로드 ──────────────────────────────────────────────────
def load_xgb_params(asset: str, params_df: pd.DataFrame) -> dict:
    row = params_df[(params_df['asset_type'] == asset) & (params_df['model'] == 'A')].iloc[0]
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

# ── 모델 인스턴스 생성 ─────────────────────────────────────────────────────────
def get_models(xgb_params: dict) -> dict:
    return {
        'XGBoost':     XGBClassifier(**xgb_params),
        'CatBoost':    CatBoostClassifier(iterations=200, depth=4, learning_rate=0.05,
                                          random_seed=42, verbose=0),
        'RandomForest':RandomForestClassifier(n_estimators=200, max_depth=4,
                                              random_state=42, n_jobs=-1),
        'ElasticNet':  LogisticRegression(penalty='elasticnet', solver='saga',
                                          l1_ratio=0.5, C=1.0, max_iter=1000,
                                          random_state=42),
    }

# ── TimeSeriesSplit AUC ────────────────────────────────────────────────────────
def cv_auc(model, X: pd.DataFrame, y: pd.Series) -> float:
    tscv = TimeSeriesSplit(n_splits=N_SPLITS)
    aucs = []
    for tr, te in tscv.split(X):
        X_tr, X_te = X.iloc[tr], X.iloc[te]
        y_tr, y_te = y.iloc[tr], y.iloc[te]
        if len(np.unique(y_te)) < 2:
            continue
        model.fit(X_tr, y_tr)
        prob = model.predict_proba(X_te)[:, 1]
        aucs.append(roc_auc_score(y_te, prob))
    return float(np.mean(aucs)) if aucs else np.nan

# ── 메인 ──────────────────────────────────────────────────────────────────────
df        = pd.read_csv('data/processed/panel_monthly_scaled.csv')
params_df = pd.read_csv('results/xgboost_tuned_params.csv')

rows = []
for asset in ASSETS:
    print(f'\n=== {asset.upper()} ===')
    # ablation.py 방식: 자산 전체 데이터를 한 번에 사용 (5아이템×48월 ≈ 240행)
    data = (
        df[df['asset_type'] == asset]
        .sort_values('year_month')
        .dropna(subset=PRICE_LAGS + [TARGET])
        .reset_index(drop=True)
    )
    X = data[FEATURES]
    y = data[TARGET]
    xgb_params = load_xgb_params(asset, params_df)
    models = get_models(xgb_params)

    for model_name, model in models.items():
        auc = cv_auc(model, X, y)
        rows.append({'asset_type': asset, 'model': model_name, 'auc': round(auc, 4)})
        print(f'  {model_name:13s} | AUC={auc:.4f}')

# 결과 저장
result_df = pd.DataFrame(rows)
result_df.to_csv('results/model_comparison_auc.csv', index=False, encoding='utf-8-sig')
print(f'\n저장 → results/model_comparison_auc.csv')

summary = result_df.copy()
summary['auc'] = summary['auc'].round(4)
print('\n=== 자산별 평균 AUC ===')
print(summary.pivot(index='asset_type', columns='model', values='auc').to_string())

# ── 막대그래프 ─────────────────────────────────────────────────────────────────
# 한글 폰트
for fname in ['Malgun Gothic', 'NanumGothic', 'AppleGothic', 'DejaVu Sans']:
    try:
        fm.findfont(fm.FontProperties(family=fname), fallback_to_default=False)
        plt.rcParams['font.family'] = fname
        break
    except Exception:
        continue
plt.rcParams['axes.unicode_minus'] = False

MODEL_ORDER = ['XGBoost', 'CatBoost', 'RandomForest', 'ElasticNet']
ASSET_LABELS = {'sneakers': 'Sneakers', 'cards': 'Cards', 'lego': 'LEGO'}
COLORS = ['#2196F3', '#FF9800', '#4CAF50', '#9C27B0']

fig, axes = plt.subplots(1, 3, figsize=(13, 5), sharey=False)
fig.suptitle('4-Model AUC Comparison (Model A: All Channels)', fontsize=13, fontweight='bold')

for ax, asset in zip(axes, ASSETS):
    sub = summary[summary['asset_type'] == asset]
    aucs = [sub[sub['model'] == m]['auc'].values[0] if len(sub[sub['model'] == m]) > 0 else np.nan
            for m in MODEL_ORDER]

    bars = ax.bar(MODEL_ORDER, aucs, color=COLORS, width=0.6, edgecolor='white', linewidth=0.8)
    ax.axhline(0.5, color='grey', linestyle='--', linewidth=0.8, label='Random (0.5)')

    for bar, val in zip(bars, aucs):
        if not np.isnan(val):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.005,
                    f'{val:.3f}', ha='center', va='bottom', fontsize=8.5, fontweight='bold')

    ax.set_title(ASSET_LABELS[asset], fontsize=11, fontweight='bold')
    ax.set_ylim(0.4, min(1.0, max(aucs) + 0.08) if not all(np.isnan(aucs)) else 1.0)
    ax.set_ylabel('AUC' if ax == axes[0] else '')
    ax.set_xticklabels(MODEL_ORDER, rotation=20, ha='right', fontsize=9)
    ax.spines[['top', 'right']].set_visible(False)

axes[0].legend(fontsize=8, loc='lower right')
plt.tight_layout()
fig.savefig('results/figures/model_comparison.png', dpi=150, bbox_inches='tight')
print('저장 → results/figures/model_comparison.png')
