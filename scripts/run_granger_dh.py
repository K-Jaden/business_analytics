"""STEP 08b — Dumitrescu-Hurlin (2012) 패널 Granger 검정

목적:
  기존 자산-평균 Granger(scripts/run_granger.py)는 아이템 5개의 평균 시계열을 사용해
  정보 손실 + XGBoost(아이템 단위)와 분석 단위 불일치 문제가 있음.
  DH 검정은 자산별 5개 아이템 각각에 Granger를 돌린 뒤 통계량을 합산해
  자산 차원의 패널 인과성 결론을 도출 — 분석 단위 일관성 확보.

DH 통계량 (Dumitrescu & Hurlin 2012, Economic Modelling 29, 1450-1460):
  각 아이템 i × 채널 c × 래그 K마다 개별 Granger F-stat W_i 산출
  W_bar = (1/N) * Σ W_i
  Z_bar = sqrt(N/(2K)) * (W_bar - K)                                      ~ N(0,1)  (T → ∞)
  Z_tilde = sqrt(N/(2K)) * (T-3K-5)/(T-3K-3) * [(T-3K-3)/(T-3K-1)*W_bar - K] ~ N(0,1)  (유한 T)
  유한표본 보정을 위해 Z_tilde를 주 통계량으로 보고.

검정 단위:
  3자산 × 5채널 × 래그(1,2) = 30개 DH 검정
  (자산별로 묶음 — 아이템 5개가 패널의 cross-section)

정상성 처리:
  results/stationarity.csv에서 자산-평균 시계열에 대해 차분 여부가 결정됨.
  일관성을 위해 동일 규칙을 아이템 시계열에도 적용 (자산 내 5개 아이템 모두 동일 처리).

산출:
  results/granger_dh.csv     30개 DH 결과
  results/granger_dh_compare.csv  자산-평균 Granger와의 일치/불일치 표
"""
import os, warnings
import numpy as np
import pandas as pd
from statsmodels.tsa.stattools import grangercausalitytests
from scipy.stats import norm

warnings.filterwarnings('ignore')
os.makedirs('results', exist_ok=True)

# ── 데이터 로드 ───────────────────────────────────────────────────────────────
PANEL_PATH       = 'data/processed/panel_monthly.csv'
STATIONARITY_PATH = 'results/stationarity.csv'
GRANGER_PATH     = 'results/granger_results.csv'

for p in [PANEL_PATH, STATIONARITY_PATH, GRANGER_PATH]:
    if not os.path.exists(p):
        raise FileNotFoundError(f'먼저 이전 스텝 실행 필요: {p}')

panel        = pd.read_csv(PANEL_PATH).sort_values(['item_id', 'year_month']).reset_index(drop=True)
stationarity = pd.read_csv(STATIONARITY_PATH)
granger_avg  = pd.read_csv(GRANGER_PATH)

ASSETS   = ['sneakers', 'cards', 'lego']
CHANNELS = ['score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5']
LAGS     = [1, 2]                  # DH 검정 래그 — 기존 Granger 결과(주 래그 1, 일부 2) 반영
N_PER_ASSET = 5                    # 자산당 아이템 수

# ── 헬퍼: 자산-평균 stationarity 기준으로 아이템 시계열 차분 ───────────────────
def maybe_diff(series, asset, col):
    row = stationarity[(stationarity['asset_type'] == asset) &
                       (stationarity['series'] == col)]
    if not row.empty and int(row.iloc[0]['differenced']) == 1:
        return series.diff()
    return series

# ── 헬퍼: 단일 아이템 Granger F-stat 계산 ────────────────────────────────────
def item_granger_f(price_s, ch_s, lag):
    """아이템 1개에 대해 Granger F-stat 반환 (실패 시 None)."""
    df = pd.concat([price_s, ch_s], axis=1).dropna()
    if len(df) < lag + 5:
        return None
    try:
        res = grangercausalitytests(df.values, maxlag=lag, verbose=False)
        f_stat = res[lag][0]['ssr_ftest'][0]
        return float(f_stat)
    except Exception:
        return None

# ── 헬퍼: DH 통계량 ───────────────────────────────────────────────────────────
def dh_statistic(W_list, T, K):
    """W_list = 개별 아이템 F-stat 리스트, T = 시계열 길이, K = 래그."""
    W = [w for w in W_list if w is not None and np.isfinite(w)]
    N = len(W)
    if N < 2:
        return None
    W_bar = float(np.mean(W))
    # T - 3K - 1, T - 3K - 3, T - 3K - 5 모두 양수여야 함
    if T - 3*K - 5 <= 0:
        z_tilde = None
    else:
        adj1 = (T - 3*K - 3) / (T - 3*K - 1)
        adj2 = (T - 3*K - 5) / (T - 3*K - 3)
        z_tilde = np.sqrt(N / (2.0 * K)) * adj2 * (adj1 * W_bar - K)
    z_bar = np.sqrt(N / (2.0 * K)) * (W_bar - K)
    # 단측 검정 (upper tail) — F-stat은 양수
    p_tilde = 1 - norm.cdf(z_tilde) if z_tilde is not None else None
    p_bar   = 1 - norm.cdf(z_bar)
    return {
        'N': N, 'W_bar': round(W_bar, 4),
        'Z_bar': round(z_bar, 4), 'p_Z_bar': round(p_bar, 6),
        'Z_tilde': round(z_tilde, 4) if z_tilde is not None else None,
        'p_Z_tilde': round(p_tilde, 6) if p_tilde is not None else None,
    }

# ── 메인 루프 ─────────────────────────────────────────────────────────────────
records = []
T_full = panel.groupby('item_id').size().min()  # = 48
print(f'[INFO] T = {T_full}, 자산당 N = {N_PER_ASSET}, 래그 = {LAGS}')
print(f'[INFO] DH 가정: T-3K-5 > 0 → K=1: 40, K=2: 37, K=3: 34 → 충족\n')

for asset in ASSETS:
    items = sorted(panel[panel['asset_type'] == asset]['item_id'].unique())
    assert len(items) == N_PER_ASSET, f'{asset}: {len(items)} != 5'

    for ch in CHANNELS:
        for K in LAGS:
            W_list = []
            for it in items:
                df_it = panel[panel['item_id'] == it].sort_values('year_month').reset_index(drop=True)
                price_s = maybe_diff(df_it['mean_price'], asset, 'mean_price')
                ch_s    = maybe_diff(df_it[ch],          asset, ch)
                f = item_granger_f(price_s, ch_s, K)
                W_list.append(f)

            stat = dh_statistic(W_list, T=T_full, K=K)
            if stat is None:
                print(f'  {asset}/{ch}/K={K}: SKIP (N<2)')
                records.append({'asset_type': asset, 'channel': ch, 'lag': K,
                                'N': 0, 'W_bar': None,
                                'Z_bar': None, 'p_Z_bar': None,
                                'Z_tilde': None, 'p_Z_tilde': None,
                                'individual_F': W_list})
                continue

            print(f'  {asset}/{ch}/K={K}: W_bar={stat["W_bar"]:.3f}  '
                  f'Z_tilde={stat["Z_tilde"]}  p={stat["p_Z_tilde"]}')
            records.append({'asset_type': asset, 'channel': ch, 'lag': K,
                            **stat, 'individual_F': W_list})

# ── 판정 ──────────────────────────────────────────────────────────────────────
df = pd.DataFrame(records)

def dh_verdict(row):
    p = row.get('p_Z_tilde')
    if p is None or pd.isna(p):
        return 'SKIP'
    if p < 0.05:
        return 'SIGNIFICANT'
    if p < 0.10:
        return 'MARGINAL'
    return 'NOT_SIGNIFICANT'

df['verdict'] = df.apply(dh_verdict, axis=1)

# 개별 F는 별도 컬럼으로 펼침
for i in range(N_PER_ASSET):
    df[f'F_item{i+1}'] = df['individual_F'].apply(
        lambda lst: round(lst[i], 3) if (i < len(lst) and lst[i] is not None) else None)
df = df.drop(columns=['individual_F'])

df.to_csv('results/granger_dh.csv', index=False)
print(f'\nSaved → results/granger_dh.csv  ({len(df)} rows)')

# ── 자산-평균 Granger와 비교 표 ────────────────────────────────────────────────
compare_rows = []
for _, dh_row in df.iterrows():
    # 자산-평균 결과에서 같은 (asset, channel) 행 찾기 — 단, 평균은 래그 1 위주
    avg_row = granger_avg[(granger_avg['asset_type'] == dh_row['asset_type']) &
                          (granger_avg['channel']    == dh_row['channel'])]
    if avg_row.empty:
        avg_lag = avg_p = avg_verdict = None
    else:
        avg_lag     = int(avg_row.iloc[0]['lag']) if pd.notna(avg_row.iloc[0]['lag']) else None
        avg_p       = avg_row.iloc[0]['p']
        avg_verdict = avg_row.iloc[0]['verdict']

    compare_rows.append({
        'asset_type'   : dh_row['asset_type'],
        'channel'      : dh_row['channel'],
        'dh_lag'       : dh_row['lag'],
        'dh_W_bar'     : dh_row['W_bar'],
        'dh_Z_tilde'   : dh_row['Z_tilde'],
        'dh_p'         : dh_row['p_Z_tilde'],
        'dh_verdict'   : dh_row['verdict'],
        'avg_lag'      : avg_lag,
        'avg_p'        : avg_p,
        'avg_verdict'  : avg_verdict,
        'agreement'    : ('SAME' if avg_verdict == dh_row['verdict']
                          else f'{avg_verdict}→{dh_row["verdict"]}')
    })

df_cmp = pd.DataFrame(compare_rows)
df_cmp.to_csv('results/granger_dh_compare.csv', index=False)
print(f'Saved → results/granger_dh_compare.csv  ({len(df_cmp)} rows)')

# ── 요약 ──────────────────────────────────────────────────────────────────────
print('\n=== DH 검정 결과 요약 ===')
print(df[['asset_type','channel','lag','W_bar','Z_tilde','p_Z_tilde','verdict']]
      .to_string(index=False))

print('\n=== 자산-평균 Granger와 비교 (래그 1만) ===')
print(df_cmp[df_cmp['dh_lag'] == 1][
    ['asset_type','channel','dh_p','dh_verdict','avg_p','avg_verdict','agreement']
].to_string(index=False))

sig_count = (df['verdict'] == 'SIGNIFICANT').sum()
marg_count = (df['verdict'] == 'MARGINAL').sum()
print(f'\n총 30 검정 중: SIGNIFICANT={sig_count}, MARGINAL={marg_count}')
