"""panda CH2(timelinevol) 단독 재수집"""
import time, csv, json, urllib.request, urllib.parse
import pandas as pd

OUT = 'data/raw/gdelt/sneakers_panda_vol.csv'
QUERY = '"Nike Dunk" Panda resell'
ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]

def fetch():
    params = {
        'query': QUERY, 'mode': 'timelinevol', 'format': 'json',
        'startdatetime': '20220101000000', 'enddatetime': '20251231235959',
    }
    url = 'https://api.gdeltproject.org/api/v2/doc/doc?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as r:
        body = r.read().decode('utf-8')
    if '\n\n' in body:
        body = body.split('\n\n', 1)[1]
    data = json.loads(body)
    return data.get('timeline', [{}])[0].get('data', [])

for attempt in range(5):
    try:
        print(f'attempt {attempt+1}...')
        series = fetch()
        monthly = {}
        for pt in series:
            ds = str(pt.get('date',''))[:8]
            if len(ds) == 8:
                ym = f'{ds[:4]}-{ds[4:6]}'
                if '2022-01' <= ym <= '2025-12':
                    monthly.setdefault(ym, []).append(float(pt['value']))
        monthly = {ym: sum(v)/len(v) for ym, v in monthly.items()}
        vals = list(monthly.values())
        max_v = max(vals) if vals else 1
        if max_v == 0: max_v = 1
        monthly = {ym: round(v/max_v, 6) for ym, v in monthly.items()}
        rows = [{'year_month': ym, 'vol_norm': monthly.get(ym)} for ym in ALL_MONTHS]
        with open(OUT, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['year_month','vol_norm'])
            w.writeheader(); w.writerows(rows)
        ok = sum(1 for r in rows if r['vol_norm'] is not None)
        print(f'완료: {ok}/48 months OK')
        break
    except urllib.error.HTTPError as e:
        wait = 120 if e.code == 429 else 60
        print(f'HTTP {e.code}, {wait}s 대기...')
        time.sleep(wait)
    except Exception as e:
        print(f'오류: {e}, 60s 대기...')
        time.sleep(60)
