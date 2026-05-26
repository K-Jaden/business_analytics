"""Robustness check — 래그 고정(1, 2, 3) Granger 검정

기존 run_granger.py가 VAR BIC 최적 래그 하나만 보고하는 것을 보완.
래그를 1, 2, 3으로 각각 고정하여 동일 채널의 유의성이
래그 선택에 무관하게 안정적인지 확인.

출력:
  results/granger_robustness.csv   자산 × 채널 × 래그(1/2/3) 결과
"""
import os, warnings
import pandas as pd
import numpy as np
from statsmodels.tsa.stattools import grangercausalitytests

warnings.filterwarnings('ignore')
os.makedirs('results', exist_ok=True)

for path in ['data/processed/asset_series.csv', 'results/stationarity.csv']:
    if not os.path.exists(path):
        raise FileNotFoundError(f'먼저 이전 스텝 실행 필요: {path}')

asset_series = pd.read_csv('data/processed/asset_series.csv')
stationarity = pd.read_csv('results/stationarity.csv')

ASSETS   = ['sneakers', 'cards', 'lego']
CHANNELS = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
LAGS     = [1, 2, 3]

def get_stationary_series(asset, col, df_asset):
    s = df_asset[col].copy()
    row = stationarity[(stationarity['asset_type'] == asset) &
                       (stationarity['series'] == col)]
    if not row.empty and row.iloc[0]['differenced'] == 1:
        s = s.diff()
    return s

rows = []

for asset in ASSETS:
    df_asset = (asset_series[asset_series['asset_type'] == asset]
                .sort_values('year_month')
                .reset_index(drop=True))

    price_s = get_stationary_series(asset, 'mean_price', df_asset)

    for ch in CHANNELS:
        if ch not in df_asset.columns:
            continue

        ch_s = get_stationary_series(asset, ch, df_asset)
        combined = pd.concat([price_s, ch_s], axis=1).dropna()
        combined.columns = ['price', 'channel']

        row = {'asset_type': asset, 'channel': ch}

        for lag in LAGS:
            if len(combined) < lag + 5:
                row[f'F_lag{lag}'] = None
                row[f'p_lag{lag}'] = None
                continue
            try:
                res = grangercausalitytests(combined.values, maxlag=lag, verbose=False)
                f_stat, p_val, _, _ = res[lag][0]['ssr_ftest']
                row[f'F_lag{lag}'] = round(f_stat, 3)
                row[f'p_lag{lag}'] = round(p_val, 4)
            except Exception as e:
                print(f'  [ERROR] {asset}/{ch}/lag{lag}: {e}')
                row[f'F_lag{lag}'] = None
                row[f'p_lag{lag}'] = None

        # 래그별 유의(p<0.05) 판정
        sig_lags = [l for l in LAGS if row.get(f'p_lag{l}') is not None and row[f'p_lag{l}'] < 0.05]
        row['sig_lags'] = str(sig_lags) if sig_lags else '-'
        row['consistent'] = len(sig_lags) >= 2  # 2개 이상 래그에서 유의

        rows.append(row)

        # 출력
        parts = []
        for lag in LAGS:
            F = row.get(f'F_lag{lag}')
            p = row.get(f'p_lag{lag}')
            if F is not None:
                sig = '**' if p < 0.05 else ('*' if p < 0.10 else '')
                parts.append(f'lag{lag}: F={F:.2f} p={p:.4f}{sig}')
            else:
                parts.append(f'lag{lag}: N/A')
        print(f'{asset}/{ch}:  {" | ".join(parts)}')

df = pd.DataFrame(rows)
df.to_csv('results/granger_robustness.csv', index=False)
print(f'\nSaved → results/granger_robustness.csv')

# 요약 출력
print('\n=== 래그 견고성 요약 ===')
print('(** p<0.05, * p<0.10)\n')
col_order = ['asset_type', 'channel',
             'F_lag1', 'p_lag1', 'F_lag2', 'p_lag2', 'F_lag3', 'p_lag3',
             'sig_lags', 'consistent']
print(df[col_order].to_string(index=False))
