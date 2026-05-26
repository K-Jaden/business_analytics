"""STEP 10 — 다변수 VAR + Granger 블록·개별 인과 검정 (RQ1 보강)

목적:
  기존 단독(bivariate) Granger와 DH 패널 Granger는 모두 채널을 하나씩만 투입.
  채널 간 다중공선성(예: cards CH2-CH3 r=0.763)이 존재하면 단독 검정의 유의
  판정이 다른 채널의 효과를 흡수했을 가능성 배제 불가.
  본 스크립트는 자산-평균 시계열에 5채널을 동시 투입한 다변수 VAR을 적합한 뒤
  두 종류의 인과 검정을 수행:

  1) 블록 검정: H0 = "전 5채널이 동시에 가격을 Granger-cause하지 않는다"
     → 다중비교 문제를 단일 검정으로 통합

  2) 개별 검정: H0 = "채널 c가 다른 4개 채널을 통제했을 때 가격을 cause하지 않는다"
     → 단독 Granger와 비교 가능한 형태로 보고

분석 단위:
  자산-평균 시계열 (DH에서 이미 아이템 단위는 보완).
  6변수 VAR (mean_price + score_ch1~5), T=48.
  래그: VAR BIC로 선택 (maxlags=4).

산출:
  results/var_block.csv      자산별 블록 + 개별 검정 결과
  results/var_compare.csv    단독 Granger·DH·VAR 3종 비교
"""
import os, warnings
import numpy as np
import pandas as pd
from statsmodels.tsa.vector_ar.var_model import VAR

warnings.filterwarnings('ignore')
os.makedirs('results', exist_ok=True)

# ── 데이터 로드 ───────────────────────────────────────────────────────────────
asset_series = pd.read_csv('data/processed/asset_series.csv')
stationarity = pd.read_csv('results/stationarity.csv')
granger_avg  = pd.read_csv('results/granger_results.csv')
granger_dh   = pd.read_csv('results/granger_dh.csv')

ASSETS   = ['sneakers', 'cards', 'lego']
CHANNELS = ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']
MAX_LAG  = 4

def maybe_diff(series, asset, col):
    row = stationarity[(stationarity['asset_type'] == asset) &
                       (stationarity['series'] == col)]
    if not row.empty and int(row.iloc[0]['differenced']) == 1:
        return series.diff()
    return series

# ── 메인 ──────────────────────────────────────────────────────────────────────
block_rows = []
individual_rows = []
sample_size_rows = []

for asset in ASSETS:
    d = (asset_series[asset_series['asset_type'] == asset]
         .sort_values('year_month').reset_index(drop=True))

    # 정상성 처리
    cols = ['mean_price'] + CHANNELS
    df_v = pd.DataFrame({c: maybe_diff(d[c], asset, c) for c in cols}).dropna()
    T_eff = len(df_v)

    # VAR 적합 + 래그 선택 (DataFrame 넘겨서 컬럼명 보존)
    model = VAR(df_v)
    sel = model.select_order(maxlags=MAX_LAG)
    best_lag = max(1, int(sel.bic))
    fit = model.fit(best_lag)

    sample_size_rows.append({'asset_type': asset, 'T_eff': T_eff,
                             'best_lag': best_lag, 'aic': round(fit.aic, 3),
                             'bic': round(fit.bic, 3)})

    # 변수 인덱스 매핑 (df_v의 컬럼 순서대로)
    var_names = list(df_v.columns)
    price_idx = var_names.index('mean_price')

    # 1) 블록 검정: 5채널 → 가격
    test = fit.test_causality(caused='mean_price', causing=CHANNELS, kind='f')
    block_rows.append({
        'asset_type': asset, 'lag': best_lag, 'T_eff': T_eff,
        'F_stat'   : round(float(test.test_statistic), 4),
        'p'        : round(float(test.pvalue), 6),
        'df_num'   : int(test.df[0]) if hasattr(test, 'df') else None,
        'df_den'   : int(test.df[1]) if hasattr(test, 'df') else None,
        'verdict'  : ('SIG' if test.pvalue < 0.05
                      else 'MARG' if test.pvalue < 0.10 else 'NS'),
    })

    # 2) 개별 검정: 각 채널이 가격을 cause (다른 4채널 자동 통제 — VAR 내 부분 F)
    for ch in CHANNELS:
        test_ind = fit.test_causality(caused='mean_price', causing=[ch], kind='f')
        individual_rows.append({
            'asset_type': asset, 'channel': ch, 'lag': best_lag,
            'F_stat'    : round(float(test_ind.test_statistic), 4),
            'p'         : round(float(test_ind.pvalue), 6),
            'verdict'   : ('SIG' if test_ind.pvalue < 0.05
                           else 'MARG' if test_ind.pvalue < 0.10 else 'NS'),
        })

# ── 저장 ──────────────────────────────────────────────────────────────────────
df_block = pd.DataFrame(block_rows)
df_ind   = pd.DataFrame(individual_rows)
df_size  = pd.DataFrame(sample_size_rows)

df_block.to_csv('results/var_block.csv', index=False)
df_ind.to_csv('results/var_individual.csv', index=False)

# ── 3종 비교표 (단독 Granger / DH / VAR 개별) ─────────────────────────────────
compare_rows = []
for asset in ASSETS:
    for ch in CHANNELS:
        # 단독 Granger
        sg = granger_avg[(granger_avg['asset_type']==asset) &
                         (granger_avg['channel']==ch)]
        sg_p = float(sg.iloc[0]['p']) if not sg.empty else None
        sg_v = sg.iloc[0]['verdict'] if not sg.empty else None

        # DH (래그 1)
        dh = granger_dh[(granger_dh['asset_type']==asset) &
                        (granger_dh['channel']==ch) &
                        (granger_dh['lag']==1)]
        dh_p = float(dh.iloc[0]['p_Z_tilde']) if not dh.empty else None
        dh_v = dh.iloc[0]['verdict'] if not dh.empty else None

        # VAR 개별
        vi = df_ind[(df_ind['asset_type']==asset) & (df_ind['channel']==ch)]
        var_p = float(vi.iloc[0]['p']) if not vi.empty else None
        var_v = vi.iloc[0]['verdict'] if not vi.empty else None

        compare_rows.append({
            'asset_type': asset, 'channel': ch,
            'single_p'  : sg_p, 'single_verdict': sg_v,
            'dh_p'      : dh_p, 'dh_verdict'    : dh_v,
            'var_p'     : var_p, 'var_verdict'  : var_v,
        })

df_cmp = pd.DataFrame(compare_rows)
df_cmp.to_csv('results/var_compare.csv', index=False)

# ── 요약 출력 ─────────────────────────────────────────────────────────────────
print('=== VAR 적합 ===')
print(df_size.to_string(index=False))

print('\n=== 블록 검정 (5채널 → 가격) ===')
print(df_block.to_string(index=False))

print('\n=== 개별 검정 (다른 4채널 통제 후 각 채널 → 가격) ===')
print(df_ind.to_string(index=False))

print('\n=== 3종 비교 (단독 vs DH vs VAR 개별) ===')
print(df_cmp[['asset_type','channel',
              'single_p','single_verdict',
              'dh_p','dh_verdict',
              'var_p','var_verdict']].to_string(index=False))

# ── BH 보정 (VAR 개별 15개) ────────────────────────────────────────────────────
from statsmodels.stats.multitest import multipletests
p_vals = df_ind['p'].values
_, p_bh, _, _ = multipletests(p_vals, alpha=0.05, method='fdr_bh')
df_ind['p_bh'] = p_bh.round(4)
df_ind['verdict_bh'] = df_ind['p_bh'].apply(
    lambda p: 'SIG' if p<0.05 else ('MARG' if p<0.10 else 'NS'))
df_ind.to_csv('results/var_individual.csv', index=False)

print('\n=== VAR 개별 BH 보정 후 ===')
print(df_ind[['asset_type','channel','p','p_bh','verdict','verdict_bh']].to_string(index=False))

print(f'\nSaved → results/var_block.csv, var_individual.csv, var_compare.csv')
