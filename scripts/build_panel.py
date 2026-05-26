"""STEP 06 — 패널 통합 + 자산별 대표 시계열 생성

출력:
  data/processed/panel_monthly.csv       아이템 레벨 (price + CH1~5 + price_direction)
  data/processed/asset_series.csv        자산 레벨 월별 평균 (Granger 투입)
  data/processed/panel_monthly_scaled.csv  z-score 통일본 (XGBoost 전용)
"""
import os
import pandas as pd
import numpy as np

ITEMS = {
    'sneakers': ['sneakers_jordan1','sneakers_panda','sneakers_yeezy','sneakers_travis','sneakers_nb550'],
    'cards':    ['cards_charizard1','cards_charizard2','cards_umbreon','cards_rayquaza','cards_pikachu'],
    'lego':     ['lego_falcon','lego_hogwarts','lego_titanic','lego_porsche','lego_bugatti'],
}
ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022, 2026) for m in range(1, 13)]
SCORE_COLS = ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']
os.makedirs('data/processed', exist_ok=True)
os.makedirs('results', exist_ok=True)

# ── 1. 가격 데이터 로드 ────────────────────────────────────────────────────────
price_frames = []
for asset_type, item_list in ITEMS.items():
    for item_id in item_list:
        path = f'data/raw/prices/{item_id}.csv'
        if not os.path.exists(path):
            print(f'[WARN] price missing: {path}')
            continue
        df = pd.read_csv(path)
        df['item_id']    = item_id
        df['asset_type'] = asset_type
        price_frames.append(df)

prices = pd.concat(price_frames, ignore_index=True)

# ── 2. 채널 점수 로드 ──────────────────────────────────────────────────────────
if not os.path.exists('data/processed/channel_scores.csv'):
    raise FileNotFoundError('먼저 build_channel_scores.py 실행 필요')

ch_scores = pd.read_csv('data/processed/channel_scores.csv')

# ── 3. 병합 ───────────────────────────────────────────────────────────────────
panel = prices.merge(ch_scores[['item_id','year_month'] + SCORE_COLS],
                     on=['item_id','year_month'], how='left')

# ── 4. 선형 보간 (연속 3개월↑ 결측 시 보간 안 함) ────────────────────────────
def interpolate_with_limit(series, max_gap=2):
    """연속 max_gap개월 이하 결측만 보간 (limit_direction='forward' + backward)"""
    return series.interpolate(method='linear', limit=max_gap, limit_direction='both')

for col in ['mean_price'] + SCORE_COLS:
    panel[col] = panel.groupby('item_id')[col].transform(interpolate_with_limit)

# ── 5. tx_count < 3 처리 ──────────────────────────────────────────────────────
# 가격은 이미 보간됨. tx_count=0 은 CLAUDE.md 기준으로 limitation 명시용 플래그
if 'tx_count' in panel.columns:
    low_tx = (panel['tx_count'] < 3).sum()
    if low_tx:
        print(f'[INFO] tx_count < 3: {low_tx} rows (보간 또는 limitation)')

# ── 6. price_direction 타겟 변수 ───────────────────────────────────────────────
panel['price_direction'] = (
    panel.groupby('item_id')['mean_price'].diff() > 0
).astype('Int64')  # nullable int (첫 행 NaT 허용)

# ── 7. 행 수 확인 (panel_monthly.csv 저장은 통제 피처 추가 후 Step 9에서 수행) ──
print(f'panel_monthly: {len(panel)} rows (통제 피처 추가 후 저장)')

# ── 8. 자산별 대표 시계열 (Granger 투입) ──────────────────────────────────────
asset_series = panel.groupby(['asset_type','year_month'])[
    ['mean_price'] + SCORE_COLS
].mean().reset_index()
asset_series.sort_values(['asset_type','year_month'], inplace=True)
asset_series.to_csv('data/processed/asset_series.csv', index=False)
print(f'asset_series: {len(asset_series)} rows → data/processed/asset_series.csv')

# ── 9. 통제 피처 추가 (XGBoost 투입 전 필요) ──────────────────────────────────
# price_vs_ma3: 현재 가격 / 3개월 이동평균 (추세 대비 현재 위치)
panel['price_vs_ma3'] = panel.groupby('item_id')['mean_price'].transform(
    lambda x: x / x.rolling(3, min_periods=1).mean()
)
# price_chg_lag1~3: t-lag 시점의 전월 대비 가격 변화율 (shift 필수 — 미래 누설 방지)
for lag in [1, 2, 3]:
    panel[f'price_chg_lag{lag}'] = panel.groupby('item_id')['mean_price'].transform(
        lambda x, l=lag: x.pct_change(1).shift(l)
    )
# price_vs_ma3 초기 NaN은 없음 (min_periods=1). lag 피처 초기 NaN은 XGBoost가 처리.
panel.to_csv('data/processed/panel_monthly.csv', index=False)  # 통제 피처 포함 덮어쓰기
print(f'panel_monthly (통제 피처 추가) → data/processed/panel_monthly.csv')

# ── 10. z-score 통일본 (XGBoost 전용) ─────────────────────────────────────────
panel_scaled = panel.copy()
for col in SCORE_COLS:
    vals = panel[col].fillna(0)
    mean, std = vals.mean(), vals.std()
    panel_scaled[col] = (vals - mean) / (std if std > 0 else 1.0)
# 통제 피처는 이미 비율/변화율 형태라 z-score 불필요 — 원본 그대로 유지
panel_scaled.to_csv('data/processed/panel_monthly_scaled.csv', index=False)
print(f'panel_monthly_scaled → data/processed/panel_monthly_scaled.csv')

# ── 11. 채널 간 상관 확인 (다중공선성 점검) ───────────────────────────────────
print('\n채널 간 상관행렬 (CH2·CH3 다중공선성 주목):')
corr = panel[SCORE_COLS].corr().round(3)
print(corr)

high_corr = [(c1, c2, corr.loc[c1, c2])
             for i, c1 in enumerate(SCORE_COLS)
             for j, c2 in enumerate(SCORE_COLS)
             if i < j and abs(corr.loc[c1, c2]) >= 0.7]
if high_corr:
    print('\n[!] 상관계수 0.7↑ (논문에 VIF와 함께 보고):')
    for c1, c2, r in high_corr:
        print(f'    {c1} - {c2}: {r}')
