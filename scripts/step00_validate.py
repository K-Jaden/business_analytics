"""STEP 00 — 데이터 연속성 사전 탐색
체크 항목:
  - 가격: 48개월 커버리지, tx_count<3 월수, 연속 결측, ±50% 급등락
  - Google Trends: 48개월 커버리지, zero값 비율
  - GDELT vol/tone: 48개월 커버리지, blank(NaN) 비율
  - YouTube views_raw: 현재 수집 월수 (미완료 아이템 포함)
출력: results/step00_report.csv + 콘솔 요약
"""
import os, csv
import pandas as pd

ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022, 2026) for m in range(1, 13)]
TOTAL_MONTHS = len(ALL_MONTHS)  # 48

ITEMS = {
    'sneakers': ['sneakers_jordan1', 'sneakers_panda', 'sneakers_yeezy',
                 'sneakers_travis', 'sneakers_nb550'],
    'cards':    ['cards_charizard1', 'cards_umbreon', 'cards_rayquaza',
                 'cards_pikachu', 'cards_charizard2'],
    'lego':     ['lego_falcon', 'lego_hogwarts', 'lego_titanic',
                 'lego_porsche', 'lego_bugatti'],
}

PRICE_DIR  = 'data/raw/prices'
TREND_DIR  = 'data/raw/google_trends'
GDELT_DIR  = 'data/raw/gdelt'
YT_DIR     = 'data/raw/youtube'

os.makedirs('results', exist_ok=True)

rows = []
flags = []

def flag(item_id, check, msg):
    flags.append({'item_id': item_id, 'check': check, 'message': msg})
    print(f'  [FLAG] {check}: {msg}')

# ── 가격 검증 ──────────────────────────────────────────────────────────────────
def check_price(item_id):
    path = f'{PRICE_DIR}/{item_id}.csv'
    if not os.path.exists(path):
        flag(item_id, 'PRICE', '파일 없음')
        return {}

    df = pd.read_csv(path)
    df['year_month'] = df['year_month'].astype(str)
    covered = set(df['year_month'])
    missing = [m for m in ALL_MONTHS if m not in covered]
    coverage = len(covered) / TOTAL_MONTHS

    # tx_count 분석
    low_tx = df[df['tx_count'] < 3]
    zero_tx = df[df['tx_count'] == 0]

    # 연속 결측 탐지 (price NaN 또는 월 자체 없음)
    df_full = pd.DataFrame({'year_month': ALL_MONTHS})
    df_full = df_full.merge(df, on='year_month', how='left')
    is_missing = df_full['mean_price'].isna()
    max_consec_missing = 0
    cur = 0
    for v in is_missing:
        cur = cur + 1 if v else 0
        max_consec_missing = max(max_consec_missing, cur)

    # tx_count 연속 <3
    df_full2 = df_full.copy()
    df_full2['low'] = (df_full2['tx_count'].fillna(0) < 3)
    max_consec_lowtx = 0
    cur2 = 0
    for v in df_full2['low']:
        cur2 = cur2 + 1 if v else 0
        max_consec_lowtx = max(max_consec_lowtx, cur2)

    # ±50% 급등락
    df_sorted = df.sort_values('year_month')
    df_sorted['pct_chg'] = df_sorted['mean_price'].pct_change().abs()
    spike_months = df_sorted[df_sorted['pct_chg'] > 0.5]['year_month'].tolist()

    if coverage < 0.83:
        flag(item_id, 'PRICE_COV', f'커버리지 {coverage:.0%} < 83% (missing={missing})')
    if max_consec_missing >= 3:
        flag(item_id, 'PRICE_CONSEC_MISSING', f'연속 결측 {max_consec_missing}개월')
    if max_consec_lowtx >= 3:
        flag(item_id, 'PRICE_LOWTX_CONSEC', f'tx_count<3 연속 {max_consec_lowtx}개월')
    if spike_months:
        flag(item_id, 'PRICE_SPIKE', f'±50% 초과: {spike_months}')

    return {
        'price_coverage': f'{len(covered)}/{TOTAL_MONTHS}',
        'price_missing_months': len(missing),
        'price_consec_missing': max_consec_missing,
        'price_lowtx_months': len(low_tx),
        'price_zerotx_months': len(zero_tx),
        'price_consec_lowtx': max_consec_lowtx,
        'price_spike_months': len(spike_months),
    }

# ── Google Trends 검증 ────────────────────────────────────────────────────────
def check_trends(item_id):
    path = f'{TREND_DIR}/{item_id}.csv'
    if not os.path.exists(path):
        flag(item_id, 'TRENDS', '파일 없음')
        return {}

    df = pd.read_csv(path)
    df['year_month'] = df['year_month'].astype(str)
    covered = set(df['year_month'])
    coverage = len(covered) / TOTAL_MONTHS

    df_full = pd.DataFrame({'year_month': ALL_MONTHS}).merge(df, on='year_month', how='left')
    zeros = (df_full['interest'].fillna(0) == 0).sum()
    zero_ratio = zeros / TOTAL_MONTHS
    missing = df_full['interest'].isna().sum()

    if coverage < 0.83:
        flag(item_id, 'TRENDS_COV', f'커버리지 {coverage:.0%}')
    if zero_ratio > 0.20:
        flag(item_id, 'TRENDS_ZEROS', f'zero값 {zeros}/{TOTAL_MONTHS} ({zero_ratio:.0%}) > 20%')

    return {
        'ch1_coverage': f'{len(covered)}/{TOTAL_MONTHS}',
        'ch1_zero_months': int(zeros),
        'ch1_zero_ratio': f'{zero_ratio:.0%}',
        'ch1_missing': int(missing),
    }

# ── GDELT 검증 ────────────────────────────────────────────────────────────────
def check_gdelt(item_id):
    result = {}
    for suffix, col, ch in [('vol', 'vol_norm', 'CH2'), ('tone', 'tone', 'CH3')]:
        path = f'{GDELT_DIR}/{item_id}_{suffix}.csv'
        if not os.path.exists(path):
            flag(item_id, f'GDELT_{ch}', '파일 없음')
            result[f'{ch.lower()}_coverage'] = 'N/A'
            result[f'{ch.lower()}_blank_months'] = 'N/A'
            continue

        df = pd.read_csv(path)
        df['year_month'] = df['year_month'].astype(str)
        df_full = pd.DataFrame({'year_month': ALL_MONTHS}).merge(df, on='year_month', how='left')

        blank = df_full[col].isna().sum()
        blank_ratio = blank / TOTAL_MONTHS
        covered = df_full[col].notna().sum()

        if blank_ratio > 0.20:
            flag(item_id, f'GDELT_{ch}_BLANK', f'blank {blank}/{TOTAL_MONTHS} ({blank_ratio:.0%}) > 20%')

        result[f'{ch.lower()}_coverage'] = f'{covered}/{TOTAL_MONTHS}'
        result[f'{ch.lower()}_blank_months'] = int(blank)

    return result

# ── YouTube 검증 ──────────────────────────────────────────────────────────────
def check_youtube(item_id):
    path = f'{YT_DIR}/{item_id}_views_raw.csv'
    if not os.path.exists(path):
        return {
            'ch4_collected_months': 0,
            'ch4_zero_video_months': 0,
            'ch4_status': '미수집',
        }

    df = pd.read_csv(path)
    collected = len(df)
    zero_vid = (df['video_count'] == 0).sum()

    # 최종 views.csv 존재 여부 (48개월 완료 + score_ch4 계산됨)
    final_exists = os.path.exists(f'{YT_DIR}/{item_id}_views.csv')
    status = '완료' if final_exists else f'{collected}/48 수집중'

    if final_exists and zero_vid / TOTAL_MONTHS > 0.20:
        flag(item_id, 'CH4_ZEROS', f'영상 0개 월: {zero_vid}/{TOTAL_MONTHS} ({zero_vid/TOTAL_MONTHS:.0%}) > 20%')

    return {
        'ch4_collected_months': collected,
        'ch4_zero_video_months': int(zero_vid),
        'ch4_status': status,
    }

# ── 메인 루프 ────────────────────────────────────────────────────────────────
print('=' * 65)
print('STEP 00 - 데이터 연속성 사전 탐색')
print('=' * 65)

for asset, items in ITEMS.items():
    print(f'\n[{asset.upper()}]')
    for item_id in items:
        print(f'\n  {item_id}')
        row = {'item_id': item_id, 'asset_type': asset}
        row.update(check_price(item_id))
        row.update(check_trends(item_id))
        row.update(check_gdelt(item_id))
        row.update(check_youtube(item_id))
        rows.append(row)

# ── 결과 저장 ────────────────────────────────────────────────────────────────
df_report = pd.DataFrame(rows)
df_report.to_csv('results/step00_report.csv', index=False, encoding='utf-8-sig')

df_flags = pd.DataFrame(flags) if flags else pd.DataFrame(columns=['item_id','check','message'])
df_flags.to_csv('results/step00_flags.csv', index=False, encoding='utf-8-sig')

# ── 요약 ──────────────────────────────────────────────────────────────────────
print('\n' + '=' * 65)
print(f'총 플래그: {len(flags)}개')
if flags:
    print('\n[플래그 목록]')
    for f in flags:
        print(f"  {f['item_id']:25s} [{f['check']:25s}] {f['message']}")

print(f'\nSaved → results/step00_report.csv')
print(f'Saved → results/step00_flags.csv')

# ── 아이템별 요약 테이블 ──────────────────────────────────────────────────────
print('\n[아이템별 요약]')
print(f"{'item_id':25s} {'가격':^10} {'CH1zeros':^10} {'CH2blank':^10} {'CH3blank':^10} {'CH4':^12}")
print('-' * 80)
for _, r in df_report.iterrows():
    price_cov = r.get('price_coverage', 'N/A')
    ch1_z     = r.get('ch1_zero_months', 'N/A')
    ch2_b     = r.get('ch2_blank_months', 'N/A')
    ch3_b     = r.get('ch3_blank_months', 'N/A')
    ch4_s     = r.get('ch4_status', 'N/A')
    print(f"{r['item_id']:25s} {str(price_cov):^10} {str(ch1_z):^10} {str(ch2_b):^10} {str(ch3_b):^10} {str(ch4_s):^12}")
