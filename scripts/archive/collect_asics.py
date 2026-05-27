"""ASICS 후보 커버리지 테스트 + 최적 선택"""
import undetected_chromedriver as uc
import time, json, os, csv
from datetime import datetime
from collections import defaultdict

CHROME_VERSION = 147
KRW_TO_USD = 1380

# 테스트 우선순위: Gel-Kayano 14 (문화적 의미) → Gel-1130 (폭넓은 커버리지)
CANDIDATES = [
    ('sneakers_asics', 'asics-gel-kayano-14-cream-sweet-pink'),
    ('sneakers_asics', 'asics-gel-kayano-14-white-graphite-grey'),
    ('sneakers_asics', 'asics-gel-kayano-14-black-pure-silver'),
    ('sneakers_asics', 'asics-gel-1130-white-pure-silver'),
    ('sneakers_asics', 'asics-gel-1130-black-pure-silver'),
    ('sneakers_asics', 'asics-gel-nyc-graphite-grey-black'),
    ('sneakers_asics', 'asics-gt-2160-black-pure-silver'),
]

def parse_xvalue(xv):
    try:
        return datetime.utcfromtimestamp(int(xv) / 1000)
    except (ValueError, TypeError):
        pass
    try:
        return datetime.fromisoformat(str(xv)[:19])
    except Exception:
        return None

def detect_currency(series):
    return 'KRW' if any(pt.get('yValue', 0) > 10000 for pt in series[:5]) else 'USD'

def make_driver():
    opts = uc.ChromeOptions()
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--lang=en-US')
    opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    return uc.Chrome(options=opts, version_main=CHROME_VERSION, headless=False)

def collect_slug(slug, retries=2):
    driver = make_driver()
    try:
        driver.get('https://stockx.com')
        time.sleep(15)
        for attempt in range(retries):
            driver.get_log('performance')
            driver.get(f'https://stockx.com/{slug}')
            time.sleep(15)
            for y in [500, 1000, 1500, 2000]:
                driver.execute_script(f'window.scrollTo(0, {y})')
                time.sleep(2)
            time.sleep(5)
            series = []
            for log in driver.get_log('performance'):
                try:
                    msg = json.loads(log['message'])
                    if msg['message']['method'] != 'Network.requestWillBeSent':
                        continue
                    post = msg['message']['params'].get('request', {}).get('postData', '')
                    rid = msg['message']['params'].get('requestId')
                    if not post: continue
                    gql = json.loads(post)
                    if gql.get('operationName') != 'fetchSalesGraph': continue
                    body = driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': rid})
                    raw = json.loads(body.get('body', '{}'))
                    series = raw.get('data', {}).get('product', {}).get('salesChart', {}).get('series', [])
                    break
                except Exception:
                    pass
            if series:
                return series
            time.sleep(20)
    except Exception as e:
        print(f'  driver error: {e}')
    finally:
        try: driver.quit()
        except: pass
    return []

def check_coverage(series, currency):
    monthly = defaultdict(list)
    for pt in series:
        d = parse_xvalue(pt['xValue'])
        if not d: continue
        ym = d.strftime('%Y-%m')
        price = pt['yValue']
        if currency == 'KRW': price = round(price / KRW_TO_USD, 2)
        monthly[ym].append(price)

    all_months = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]
    present = [m for m in all_months if m in monthly]
    missing = [m for m in all_months if m not in monthly]
    gap = max_gap = 0
    for m in all_months:
        if m not in monthly: gap += 1; max_gap = max(max_gap, gap)
        else: gap = 0
    return monthly, len(present), missing, max_gap

def save_csv(item_id, monthly):
    all_months = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]
    rows = [{'year_month': ym,
             'mean_price': round(sum(monthly[ym])/len(monthly[ym]), 2),
             'tx_count': len(monthly[ym])}
            for ym in all_months if ym in monthly]
    os.makedirs('data/raw/prices', exist_ok=True)
    with open(f'data/raw/prices/{item_id}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader(); w.writerows(rows)
    return len(rows)

for item_id, slug in CANDIDATES:
    if os.path.exists(f'data/raw/prices/{item_id}.csv'):
        print(f'[{item_id}] already exists, skip')
        break

    print(f'\n[{slug}]')
    series = collect_slug(slug)
    print(f'  Points: {len(series)}')
    if not series:
        print('  No data, next...')
        time.sleep(10)
        continue

    currency = detect_currency(series)
    monthly, n, missing, max_gap = check_coverage(series, currency)
    print(f'  Coverage: {n}/48  max_gap={max_gap}  currency={currency}')
    print(f'  Missing: {missing}')

    if n >= 44 and max_gap <= 2:
        saved = save_csv(item_id, monthly)
        print(f'  SELECTED -> {saved}/48 months saved (slug={slug})')
        break
    elif n >= 40 and max_gap < 3:
        saved = save_csv(item_id, monthly)
        print(f'  PASS (marginal) -> {saved}/48 months saved (slug={slug})')
        break
    else:
        print(f'  Insufficient ({n}/48 max_gap={max_gap}), trying next...')
        time.sleep(10)
else:
    print('\nAll candidates exhausted')
