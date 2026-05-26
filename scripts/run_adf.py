"""STEP 07 — ADF 정상성 검정 → results/stationarity.csv

각 (asset_type, series) 에 대해:
  p > 0.05 → 1차 차분 후 재검정
  결과: adf_p, adf_p_diff, differenced (0/1)
"""
import os
import pandas as pd
import numpy as np
from statsmodels.tsa.stattools import adfuller

os.makedirs('results', exist_ok=True)

if not os.path.exists('data/processed/asset_series.csv'):
    raise FileNotFoundError('먼저 build_panel.py 실행 필요')

asset_series = pd.read_csv('data/processed/asset_series.csv')
ASSETS  = ['sneakers', 'cards', 'lego']
SERIES  = ['mean_price', 'score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']

def run_adf(series):
    """ADF 검정. 결측 제거 후 실행."""
    s = series.dropna()
    if len(s) < 8:
        return None, None
    if s.std() == 0:  # 상수 시리즈 → adfuller LinAlgError 방지
        return None, None
    try:
        result = adfuller(s, autolag='AIC')
        return result[0], result[1]  # (adf_stat, p_value)
    except Exception as e:
        print(f'    [ADF ERROR] {e}')
        return None, None

rows = []
for asset in ASSETS:
    df = (asset_series[asset_series['asset_type'] == asset]
          .sort_values('year_month')
          .reset_index(drop=True))

    for col in SERIES:
        if col not in df.columns:
            continue
        series = df[col]
        stat, p = run_adf(series)
        if stat is None:
            print(f'[SKIP] {asset}/{col}: 유효 데이터 부족')
            rows.append({'asset_type': asset, 'series': col,
                         'adf_stat': None, 'adf_p': None,
                         'adf_p_diff': None, 'differenced': None, 'n_obs': int(series.notna().sum())})
            continue

        diff_needed = p > 0.05
        stat_d, p_d = (None, None)
        if diff_needed:
            stat_d, p_d = run_adf(series.diff())

        rows.append({
            'asset_type': asset,
            'series':     col,
            'adf_stat':   round(stat, 4),
            'adf_p':      round(p, 4),
            'adf_p_diff': round(p_d, 4) if p_d is not None else None,
            'differenced': int(diff_needed),
            'n_obs':       int(series.notna().sum()),
        })
        status = '[DIFF]' if diff_needed else '[STAT]'
        p_str = f'p={p:.4f}'
        if diff_needed:
            p_str += f' → diff p={p_d:.4f}'
        print(f'{status} {asset}/{col}: {p_str}')

df_result = pd.DataFrame(rows)
df_result.to_csv('results/stationarity.csv', index=False)
print(f'\nSaved → results/stationarity.csv')
print(df_result[['asset_type','series','adf_p','differenced']].to_string(index=False))
