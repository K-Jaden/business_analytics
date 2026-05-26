import undetected_chromedriver as uc
import time, json, os, csv
from datetime import datetime
from collections import defaultdict

CHROME_VERSION = 147

SNEAKERS = {
    'sneakers_jordan1': 'air-jordan-1-retro-high-bordeaux',
    'sneakers_panda':   'nike-dunk-low-retro-white-black-2021',
    'sneakers_yeezy':   'adidas-yeezy-boost-350-v2-white-core-black-red',
    'sneakers_travis':  'air-jordan-1-retro-high-travis-scott',
    'sneakers_nb550':   'new-balance-550-white-green',
}

KRW_TO_USD = 1380

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
    if not series:
        return 'USD'
    return 'KRW' if any(pt.get('yValue', 0) > 10000 for pt in series[:5]) else 'USD'

def make_driver():
    opts = uc.ChromeOptions()
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--lang=en-US')
    opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    return uc.Chrome(options=opts, version_main=CHROME_VERSION, headless=False)

def collect_item(slug, item_id, retries=3):
    """아이템 1개당 드라이버를 새로 시작 (창 닫힘 방지)"""
    driver = make_driver()
    try:
        driver.get('https://stockx.com')
        print(f'  [{item_id}] StockX home loaded, waiting 15s...')
        time.sleep(15)

        for attempt in range(retries):
            driver.get_log('performance')
            driver.get(f'https://stockx.com/{slug}')
            time.sleep(15)
            for y in [500, 1000, 1500, 2000, 2500]:
                driver.execute_script(f'window.scrollTo(0, {y})')
                time.sleep(2)
            time.sleep(5)

            series = []
            for log in driver.get_log('performance'):
                try:
                    msg = json.loads(log['message'])
                    if msg['message']['method'] != 'Network.requestWillBeSent':
                        continue
                    req = msg['message']['params'].get('request', {})
                    post = req.get('postData', '')
                    rid = msg['message']['params'].get('requestId')
                    if not post:
                        continue
                    gql = json.loads(post)
                    if gql.get('operationName') != 'fetchSalesGraph':
                        continue
                    body = driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': rid})
                    raw = json.loads(body.get('body', '{}'))
                    series = (raw.get('data', {})
                                 .get('product', {})
                                 .get('salesChart', {})
                                 .get('series', []))
                    break
                except Exception:
                    pass

            if series:
                print(f'  [{item_id}] attempt {attempt+1}: OK ({len(series)} points)')
                break
            print(f'  [{item_id}] attempt {attempt+1}: fetchSalesGraph not found, retry...')
            time.sleep(20)
    except Exception as e:
        print(f'  [{item_id}] driver error: {e}')
        series = []
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    if not series:
        print(f'  [{item_id}] FAILED')
        return False

    currency = detect_currency(series)
    print(f'  [{item_id}] currency: {currency}')

    monthly = defaultdict(list)
    for pt in series:
        d = parse_xvalue(pt['xValue'])
        if not d:
            continue
        ym = d.strftime('%Y-%m')
        price = pt['yValue']
        if currency == 'KRW':
            price = round(price / KRW_TO_USD, 2)
        monthly[ym].append(price)

    os.makedirs('data/raw/prices', exist_ok=True)
    rows = [{'year_month': ym,
             'mean_price': round(sum(monthly[ym]) / len(monthly[ym]), 2),
             'tx_count': len(monthly[ym])}
            for ym in sorted(monthly)
            if '2022-01' <= ym <= '2025-12']

    out = f'data/raw/prices/{item_id}.csv'
    with open(out, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader()
        w.writerows(rows)
    print(f'  [{item_id}] {len(rows)}/48 months saved -> {out}')
    return True

def main():
    for item_id, slug in SNEAKERS.items():
        out = f'data/raw/prices/{item_id}.csv'
        if os.path.exists(out):
            print(f'[{item_id}] already exists, skip')
            continue
        print(f'\n[{item_id}] slug={slug}')
        ok = collect_item(slug, item_id)
        if ok:
            print(f'  [{item_id}] done, waiting 15s before next...')
            time.sleep(15)
        else:
            print(f'  [{item_id}] failed, continuing to next item')
            time.sleep(5)

if __name__ == '__main__':
    main()
