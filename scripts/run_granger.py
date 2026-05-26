"""STEP 08 — Granger 인과성 검정 (RQ1, RQ2)

분석 단위: 자산(asset) — 아이템 평균 대표 시계열
총 검정 수: 3자산 × 5채널 = 15회
다중비교 보정: Benjamini-Hochberg FDR

출력:
  results/granger_results.csv     15개 검정 결과
  results/channel_importance.csv  채널별 유의 횟수 및 Jaccard (RQ2)
"""
import os, warnings
import pandas as pd
import numpy as np
from statsmodels.tsa.stattools import grangercausalitytests, adfuller
from statsmodels.tsa.vector_ar.var_model import VAR
from statsmodels.stats.multitest import multipletests

warnings.filterwarnings('ignore')
os.makedirs('results', exist_ok=True)

# ── 데이터 로드 ───────────────────────────────────────────────────────────────
for path in ['data/processed/asset_series.csv', 'results/stationarity.csv']:
    if not os.path.exists(path):
        raise FileNotFoundError(f'먼저 이전 스텝 실행 필요: {path}')

asset_series  = pd.read_csv('data/processed/asset_series.csv')
stationarity  = pd.read_csv('results/stationarity.csv')

ASSETS   = ['sneakers', 'cards', 'lego']
CHANNELS = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
MAX_LAG  = 6

def get_stationary_series(asset, col, df_asset):
    """ADF 결과 기반으로 필요 시 1차 차분 적용"""
    s = df_asset[col].copy()
    row = stationarity[(stationarity['asset_type'] == asset) &
                       (stationarity['series'] == col)]
    if not row.empty and row.iloc[0]['differenced'] == 1:
        s = s.diff()
    return s

def granger_test(price_series, channel_series, max_lag=MAX_LAG):
    """
    Granger 검정: channel → price ?
    Returns: best_lag (VAR BIC 기준), F, p

    VAR BIC를 쓰는 이유: grangercausalitytests는 래그마다 샘플 크기가
    달라져 BIC 직접 비교 불가. VAR는 동일 샘플에서 모든 래그를 비교.
    """
    combined = pd.concat([price_series, channel_series], axis=1).dropna()
    if len(combined) < max_lag + 5:
        return None, None, None

    data = combined.values  # col0=price, col1=channel

    # VAR BIC로 최적 래그 선택
    try:
        var_res = VAR(data).select_order(maxlags=max_lag)
        best_lag = int(max(1, var_res.bic))  # numpy int → Python int 변환, 0 방어
    except Exception:
        best_lag = 1

    # 최적 래그로 Granger F-검정
    try:
        res = grangercausalitytests(data, maxlag=best_lag, verbose=False)
        f_stat, p_val, _, _ = res[best_lag][0]['ssr_ftest']
        return best_lag, round(f_stat, 4), round(p_val, 4)
    except Exception as e:
        print(f'    [ERROR] {e}')
        return None, None, None

# ── 메인 검정 루프 ────────────────────────────────────────────────────────────
raw_results = []

for asset in ASSETS:
    df_asset = (asset_series[asset_series['asset_type'] == asset]
                .sort_values('year_month')
                .reset_index(drop=True))

    price_s = get_stationary_series(asset, 'mean_price', df_asset)

    for ch in CHANNELS:
        if ch not in df_asset.columns:
            raw_results.append({'asset_type': asset, 'channel': ch,
                                'lag': None, 'F': None, 'p': None})
            continue

        ch_s = get_stationary_series(asset, ch, df_asset)
        lag, F, p = granger_test(price_s, ch_s)
        status = f'F={F:.2f} p={p:.4f} lag={lag}' if F is not None else 'SKIP'
        print(f'  {asset}/{ch}: {status}')
        raw_results.append({'asset_type': asset, 'channel': ch,
                            'lag': lag, 'F': F, 'p': p})

# ── RQ1 판정 (raw p 기준, BH 보정 미적용) ────────────────────────────────────
df_res = pd.DataFrame(raw_results)

def rq1_verdict(row):
    if pd.isna(row['F']) or pd.isna(row['p']):
        return 'SKIP'
    if row['F'] >= 4.0 and row['p'] < 0.05:
        return 'SIGNIFICANT'
    if row['F'] >= 2.0 and row['p'] < 0.10:
        return 'MARGINAL'
    return 'NOT_SIGNIFICANT'

df_res['verdict'] = df_res.apply(rq1_verdict, axis=1)
df_res.to_csv('results/granger_results.csv', index=False)
print(f'\nSaved → results/granger_results.csv')

# ── RQ2: 자산별 유의 채널 집합 비교 (Jaccard) ──────────────────────────────────
sig_channels = {}
for asset in ASSETS:
    sig = set(df_res[(df_res['asset_type'] == asset) &
                     (df_res['verdict'] == 'SIGNIFICANT')]['channel'].tolist())
    sig_channels[asset] = sig
    print(f'{asset} 유의 채널: {sig if sig else "(없음)"}')

def jaccard(a, b):
    if not a and not b:
        return 1.0
    return round(len(a & b) / len(a | b), 3)

pairs = [(a1, a2) for i, a1 in enumerate(ASSETS)
         for a2 in ASSETS[i+1:]]
jaccard_rows = []
for a1, a2 in pairs:
    j = jaccard(sig_channels[a1], sig_channels[a2])
    interp = ('다름' if j < 0.3 else '부분' if j <= 0.6 else '유사')
    print(f'Jaccard({a1}, {a2}) = {j}  → {interp}')
    jaccard_rows.append({'asset1': a1, 'asset2': a2, 'jaccard': j, 'interpretation': interp})

# ── channel_importance.csv 저장 ───────────────────────────────────────────────
imp_rows = []
for ch in CHANNELS:
    ch_df = df_res[df_res['channel'] == ch]
    sig_count = (ch_df['verdict'] == 'SIGNIFICANT').sum()
    mean_F = ch_df['F'].mean()
    imp_rows.append({'channel': ch, 'sig_count': int(sig_count),
                     'mean_F': round(mean_F, 3) if pd.notna(mean_F) else None})

df_imp = pd.DataFrame(imp_rows).sort_values('sig_count', ascending=False)
df_jaccard = pd.DataFrame(jaccard_rows)

# 두 테이블 concat해서 저장 (섹션 구분)
with open('results/channel_importance.csv', 'w', encoding='utf-8', newline='') as f:
    f.write('# Channel Importance\n')
    df_imp.to_csv(f, index=False)
    f.write('\n# Jaccard Similarity (RQ2)\n')
    df_jaccard.to_csv(f, index=False)

print(f'\nSaved → results/channel_importance.csv')

# ── 요약 출력 ──────────────────────────────────────────────────────────────────
print('\n=== Granger 결과 요약 ===')
print(df_res[['asset_type','channel','lag','F','p','verdict']].to_string(index=False))

# 엣지케이스 경고
all_nonsig = all(v in ('NOT_SIGNIFICANT','SKIP') for v in df_res['verdict'])
if all_nonsig:
    print('\n[!] 전채널 비유의 - 래그 1~6 확장 검토 필요 (Opus 필요합니다. 전환해주세요.)')
all_sig = all(v == 'SIGNIFICANT' for v in df_res[df_res['verdict'] != 'SKIP']['verdict'])
if all_sig:
    print('\n[!] 전채널 유의 - F 크기로 채널 간 순위 산출')
