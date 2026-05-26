"""STEP 03 — GDELT CH2(뉴스량) + CH3(뉴스감성) 동시 수집
전체 기간 한 번에 fetch → 월별 집계 (30 API calls total)
"""
import time, os, csv, json
import urllib.request, urllib.parse, urllib.error
import pandas as pd

OUT_DIR = 'data/raw/gdelt'
os.makedirs(OUT_DIR, exist_ok=True)

QUERIES = {
    'sneakers_jordan1': '"Air Jordan 1" Bordeaux',
    'sneakers_panda':   '"Nike Dunk" Panda resell',
    'sneakers_yeezy':   '"Yeezy Zebra" resell',
    'sneakers_travis':  '"Travis Scott Jordan" resell',
    'sneakers_nb550':   '"New Balance 550" resell',
    'cards_charizard1': '"Charizard VMAX" "Shining Fates"',
    'cards_umbreon':    '"Umbreon VMAX" "Alt Art"',
    'cards_rayquaza':   '"Rayquaza VMAX" "Alt Art"',
    'cards_pikachu':    '"Pikachu VMAX" pokemon card',
    'cards_charizard2': '"Charizard GX" "Hidden Fates"',
    'lego_falcon':      'LEGO "Millennium Falcon" 75192',
    'lego_hogwarts':    'LEGO "Hogwarts Castle" 71043',
    'lego_titanic':     'LEGO Titanic 10294',
    'lego_porsche':     'LEGO "Porsche 911" 42096',
    'lego_bugatti':     'LEGO "Bugatti Chiron" 42083',
}

BASE = 'https://api.gdeltproject.org/api/v2/doc/doc'
TIMEFRAME = ('20220101000000', '20251231235959')
ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]

def fetch_gdelt(query, mode):
    """전체 기간 한 번에 fetch, 일별 시계열 반환"""
    params = {
        'query':         query,
        'mode':          mode,
        'format':        'json',
        'startdatetime': TIMEFRAME[0],
        'enddatetime':   TIMEFRAME[1],
    }
    url = BASE + '?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as r:
        body = r.read().decode('utf-8')
    # GDELT는 body에 HTTP 헤더를 포함 → 빈 줄 이후가 실제 JSON
    if '\n\n' in body:
        body = body.split('\n\n', 1)[1]
    data = json.loads(body)
    series = data.get('timeline', [{}])[0].get('data', [])
    return series  # [{date: 'YYYYMMDDHHMMSS', value: float}, ...]

def series_to_monthly(series):
    """일별 시계열 → 월별 평균"""
    if not series:
        return {}
    records = []
    for pt in series:
        date_str = str(pt.get('date', ''))[:8]  # YYYYMMDD
        val = pt.get('value')
        if len(date_str) == 8 and val is not None:
            try:
                ym = f'{date_str[:4]}-{date_str[4:6]}'
                records.append({'ym': ym, 'val': float(val)})
            except Exception:
                pass
    if not records:
        return {}
    df = pd.DataFrame(records)
    monthly = df.groupby('ym')['val'].mean().round(6)
    return {ym: v for ym, v in monthly.items() if '2022-01' <= ym <= '2025-12'}

def collect_item(item_id, query):
    vol_out  = f'{OUT_DIR}/{item_id}_vol.csv'
    tone_out = f'{OUT_DIR}/{item_id}_tone.csv'
    if os.path.exists(vol_out) and os.path.exists(tone_out):
        print(f'[{item_id}] already exists, skip')
        return True

    print(f'[{item_id}] "{query}"', flush=True)
    results = {}

    for mode, out_path, col in [
        ('timelinevol',  vol_out,  'vol_norm'),
        ('timelinetone', tone_out, 'tone'),
    ]:
        if os.path.exists(out_path):
            print(f'  {mode}: already exists')
            continue
        for attempt in range(4):
            try:
                series = fetch_gdelt(query, mode)
                monthly = series_to_monthly(series)
                if mode == 'timelinetone':
                    monthly = {ym: round(v/10, 6) for ym, v in monthly.items()}
                elif mode == 'timelinevol':
                    # raw count → 0~1 min-max 정규화 (아이템 내)
                    vals = [v for v in monthly.values() if v is not None]
                    max_v = max(vals) if vals else 1
                    if max_v == 0: max_v = 1
                    monthly = {ym: round(v/max_v, 6) for ym, v in monthly.items()}
                rows = [{'year_month': ym, col: monthly.get(ym, None)}
                        for ym in ALL_MONTHS]
                with open(out_path, 'w', newline='', encoding='utf-8') as f:
                    w = csv.DictWriter(f, fieldnames=['year_month', col])
                    w.writeheader(); w.writerows(rows)
                ok = sum(1 for r in rows if r[col] is not None)
                print(f'  {mode}: {ok}/48 months OK')
                results[mode] = ok
                break
            except urllib.error.HTTPError as e:
                wait = 120 if e.code == 429 else 45 * (2 ** attempt)
                print(f'  {mode} attempt {attempt+1} HTTP {e.code}, sleep {wait}s')
                time.sleep(wait)
            except Exception as e:
                wait = 45 * (2 ** attempt)
                print(f'  {mode} attempt {attempt+1} error: {e}, sleep {wait}s')
                time.sleep(wait)
        time.sleep(30)

    return True

all_ok = True
for i, (item_id, query) in enumerate(QUERIES.items()):
    ok = collect_item(item_id, query)
    all_ok = all_ok and ok
    if i < len(QUERIES) - 1:
        time.sleep(30)

print(f'\n=== GDELT DONE --- {"ALL OK" if all_ok else "CHECK LOGS"} ===')
