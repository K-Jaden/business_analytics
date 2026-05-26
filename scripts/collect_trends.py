"""STEP 02 — Google Trends 월별 수집 (pytrends, geo='', 아이템별 단독)"""
import time, os, csv, sys
from pytrends.request import TrendReq
import pandas as pd

# 아이템별 키워드 (단독 검색어 — 너무 길면 데이터 0)
KEYWORDS = {
    'sneakers_jordan1': 'Jordan 1 Bordeaux',
    'sneakers_panda':   'Nike Dunk Panda',
    'sneakers_yeezy':   'Yeezy Zebra',
    'sneakers_travis':  'Travis Scott Jordan 1',
    'sneakers_nb550':   'New Balance 550',
    'cards_charizard1': 'Charizard VMAX Shining Fates',
    'cards_umbreon':    'Umbreon VMAX Alt Art',
    'cards_rayquaza':   'Rayquaza VMAX Alt Art',
    'cards_pikachu':    'Pikachu VMAX',
    'cards_charizard2': 'Charizard GX Hidden Fates',
    'lego_falcon':      'LEGO Millennium Falcon',
    'lego_hogwarts':    'LEGO Hogwarts Castle',
    'lego_titanic':     'LEGO Titanic 10294',
    'lego_porsche':     'LEGO Porsche 911 42096',
    'lego_bugatti':     'LEGO Bugatti Chiron 42083',
}

TIMEFRAME = '2022-01-01 2025-12-31'
OUT_DIR = 'data/raw/google_trends'
os.makedirs(OUT_DIR, exist_ok=True)

pytrends = TrendReq(hl='en-US', tz=0, timeout=(10, 25))

def fetch_item(item_id, keyword):
    out = f'{OUT_DIR}/{item_id}.csv'
    if os.path.exists(out):
        print(f'[{item_id}] already exists, skip')
        return True

    print(f'[{item_id}] "{keyword}"', flush=True)
    for attempt in range(3):
        try:
            pytrends.build_payload([keyword], cat=0, timeframe=TIMEFRAME, geo='', gprop='')
            df = pytrends.interest_over_time()

            if df.empty:
                print(f'  attempt {attempt+1}: empty response')
                time.sleep(15)
                continue

            # 주별 → 월별 평균 (floor('MS') = 월 시작일로 그루핑)
            df.index = pd.to_datetime(df.index)
            monthly = (df[keyword]
                       .resample('MS')
                       .mean()
                       .round(1))

            # 2022-01 ~ 2025-12 필터
            monthly = monthly[
                (monthly.index >= '2022-01-01') &
                (monthly.index <= '2025-12-31')
            ]

            rows = [{'year_month': d.strftime('%Y-%m'), 'interest': v}
                    for d, v in monthly.items() if not pd.isna(v)]

            with open(out, 'w', newline='', encoding='utf-8') as f:
                w = csv.DictWriter(f, fieldnames=['year_month', 'interest'])
                w.writeheader(); w.writerows(rows)

            missing = [f'{y}-{m:02d}' for y in range(2022,2026)
                       for m in range(1,13)
                       if f'{y}-{m:02d}' not in {r["year_month"] for r in rows}]
            print(f'  OK: {len(rows)}/48 months  missing={missing}')
            return True

        except Exception as e:
            print(f'  attempt {attempt+1} error: {e}')
            time.sleep(20)

    print(f'  [{item_id}] FAILED after 3 attempts')
    return False

results = {}
items = list(KEYWORDS.items())
for i, (item_id, keyword) in enumerate(items):
    ok = fetch_item(item_id, keyword)
    results[item_id] = ok
    if ok and i < len(items) - 1:
        print(f'  sleeping 12s...')
        time.sleep(12)
    elif not ok:
        time.sleep(5)

print('\n=== SUMMARY ===')
passed = sum(results.values())
print(f'{passed}/{len(results)} collected')
for iid, ok in results.items():
    if not ok:
        print(f'  FAILED: {iid}')
