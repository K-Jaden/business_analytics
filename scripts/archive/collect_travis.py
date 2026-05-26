import undetected_chromedriver as uc
import time, json, os, csv
from datetime import datetime
from collections import defaultdict

CHROME_VERSION = 147
KRW_TO_USD = 1380
SLUG = 'air-jordan-1-retro-high-travis-scott'
ITEM_ID = 'sneakers_travis'

def parse_xvalue(xv):
    try:
        return datetime.utcfromtimestamp(int(xv) / 1000)
    except (ValueError, TypeError):
        pass
    try:
        return datetime.fromisoformat(str(xv)[:19])
    except Exception:
        return None

def make_driver():
    opts = uc.ChromeOptions()
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--lang=en-US')
    opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    return uc.Chrome(options=opts, version_main=CHROME_VERSION, headless=False)

driver = make_driver()
try:
    driver.get('https://stockx.com')
    print('StockX home... 15s')
    time.sleep(15)

    series = []
    for attempt in range(3):
        driver.get_log('performance')
        driver.get(f'https://stockx.com/{SLUG}')
        time.sleep(15)
        for y in [500, 1000, 1500, 2000, 2500]:
            driver.execute_script(f'window.scrollTo(0, {y})')
            time.sleep(2)
        time.sleep(5)
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
            print(f'  attempt {attempt+1}: OK ({len(series)} points)')
            break
        print(f'  attempt {attempt+1}: not found, retry...')
        time.sleep(20)
finally:
    try: driver.quit()
    except: pass

if not series:
    print('FAILED'); exit(1)

currency = 'KRW' if any(pt.get('yValue', 0) > 10000 for pt in series[:5]) else 'USD'
print(f'  currency: {currency}')

monthly = defaultdict(list)
for pt in series:
    d = parse_xvalue(pt['xValue'])
    if not d: continue
    ym = d.strftime('%Y-%m')
    price = pt['yValue']
    if currency == 'KRW': price = round(price / KRW_TO_USD, 2)
    monthly[ym].append(price)

all_months = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]
rows = [{'year_month': ym,
         'mean_price': round(sum(monthly[ym])/len(monthly[ym]), 2),
         'tx_count': len(monthly[ym])}
        for ym in all_months if ym in monthly]

os.makedirs('data/raw/prices', exist_ok=True)
with open(f'data/raw/prices/{ITEM_ID}.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
    w.writeheader(); w.writerows(rows)

missing = [m for m in all_months if m not in monthly]
print(f'  {len(rows)}/48 months saved  missing={missing}')
