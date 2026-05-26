import undetected_chromedriver as uc
import time, json, os, csv
from datetime import datetime
from collections import defaultdict

CHROME_VERSION = 147
KRW_TO_USD = 1380

# Yeezy: re-collect with intervals=200 via browser fetch
YEEZY_SLUG = 'adidas-yeezy-boost-350-v2-white-core-black-red'

# ASICS candidates (순서대로 시도, 첫 번째 성공 사용)
ASICS_CANDIDATES = [
    'asics-gel-kayano-14-white-sage',
    'asics-gel-kayano-14-birch-midnight',
    'asics-gel-kayano-14-white-white',
    'asics-gel-kayano-14',
    'asics-gel-kayano-14-white-pure-silver',
]
ASICS_ITEM_ID = 'sneakers_asics'

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

def fetch_via_browser(driver, slug, intervals=200):
    """브라우저 컨텍스트에서 fetch 실행 (StockX 쿠키 자동 포함)"""
    js = """
    var done = arguments[arguments.length - 1];
    var slug = arguments[0];
    var intervals = arguments[1];
    fetch('https://stockx.com/api/p/e', {
        method: 'POST',
        credentials: 'include',
        headers: {
            'Content-Type': 'application/json',
            'apollographql-client-name': 'Iron'
        },
        body: JSON.stringify({
            operationName: 'fetchSalesGraph',
            variables: {
                productId: slug,
                countryCode: 'KR',
                currencyCode: 'KRW',
                startDate: '2022-01-01',
                endDate: '2025-12-31',
                intervals: intervals
            },
            query: 'query fetchSalesGraph($productId: String!, $countryCode: String, $currencyCode: CurrencyCode, $startDate: String, $endDate: String, $intervals: Int) { product(id: $productId) { salesChart(countryCode: $countryCode, currencyCode: $currencyCode, startDate: $startDate, endDate: $endDate, intervals: $intervals) { series { xValue yValue } } } }'
        })
    }).then(function(r) { return r.json(); })
      .then(function(d) { done(JSON.stringify(d)); })
      .catch(function(e) { done(JSON.stringify({error: e.toString()})); });
    """
    result = driver.execute_async_script(js, slug, intervals)
    if not result:
        return []
    parsed = json.loads(result)
    if 'error' in parsed:
        print(f'  fetch error: {parsed["error"]}')
        return []
    return (parsed.get('data', {})
                  .get('product', {})
                  .get('salesChart', {})
                  .get('series', []))

def intercept_normal(driver, slug):
    """기존 방식: 페이지 로드 후 network log 인터셉트"""
    driver.get_log('performance')
    driver.get(f'https://stockx.com/{slug}')
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
            if not post:
                continue
            gql = json.loads(post)
            if gql.get('operationName') != 'fetchSalesGraph':
                continue
            body = driver.execute_cdp_cmd('Network.getResponseBody', {'requestId': rid})
            raw = json.loads(body.get('body', '{}'))
            return raw.get('data', {}).get('product', {}).get('salesChart', {}).get('series', [])
        except Exception:
            pass
    return []

def series_to_monthly(series, currency):
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
    return monthly

def check_coverage(monthly):
    all_months = [f'{y}-{m:02d}' for y in range(2022, 2026) for m in range(1, 13)]
    present = [m for m in all_months if m in monthly]
    missing = [m for m in all_months if m not in monthly]
    gap = max_gap = 0
    for m in all_months:
        if m not in monthly:
            gap += 1; max_gap = max(max_gap, gap)
        else:
            gap = 0
    return len(present), missing, max_gap

def save_csv(item_id, monthly):
    all_months = [f'{y}-{m:02d}' for y in range(2022, 2026) for m in range(1, 13)]
    rows = [{'year_month': ym,
             'mean_price': round(sum(monthly[ym]) / len(monthly[ym]), 2),
             'tx_count': len(monthly[ym])}
            for ym in all_months if ym in monthly]
    os.makedirs('data/raw/prices', exist_ok=True)
    with open(f'data/raw/prices/{item_id}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader(); w.writerows(rows)
    return len(rows)

# ── 1. Yeezy (intervals=200) ────────────────────────────────────────────────
def collect_yeezy():
    print('\n[sneakers_yeezy] Adidas Yeezy 350 V2 Zebra (intervals=200)')
    driver = make_driver()
    try:
        driver.get('https://stockx.com')
        print('  StockX home... waiting 15s')
        time.sleep(15)

        # 페이지 로드해서 세션 확립
        driver.get(f'https://stockx.com/{YEEZY_SLUG}')
        print('  Product page loaded... waiting 15s')
        time.sleep(15)

        # intervals=200으로 브라우저 fetch
        print('  Fetching with intervals=200...')
        series = fetch_via_browser(driver, YEEZY_SLUG, intervals=200)
        print(f'  intervals=200: {len(series)} points')

        if len(series) < 80:
            # fallback: 인터셉트 방식
            print('  Fallback to intercept...')
            series2 = intercept_normal(driver, YEEZY_SLUG)
            print(f'  intercept: {len(series2)} points')
            if len(series2) > len(series):
                series = series2
    except Exception as e:
        print(f'  Error: {e}')
        series = []
    finally:
        try: driver.quit()
        except: pass

    if not series:
        print('  FAILED')
        return False

    currency = detect_currency(series)
    monthly = series_to_monthly(series, currency)
    n, missing, max_gap = check_coverage(monthly)
    print(f'  Coverage: {n}/48  missing={missing}  max_gap={max_gap}')
    saved = save_csv('sneakers_yeezy', monthly)
    print(f'  Saved: {saved}/48 months')
    return n >= 48

# ── 2. ASICS (슬러그 탐색 + 수집) ──────────────────────────────────────────
def collect_asics():
    print('\n[sneakers_asics] ASICS Gel-Kayano 14 - slug search')
    for slug in ASICS_CANDIDATES:
        print(f'\n  Trying: {slug}')
        driver = make_driver()
        try:
            driver.get('https://stockx.com')
            time.sleep(15)
            series = intercept_normal(driver, slug)
            print(f'  Points: {len(series)}')
        except Exception as e:
            print(f'  Error: {e}')
            series = []
        finally:
            try: driver.quit()
            except: pass

        if not series:
            print(f'  No data for {slug}, trying next...')
            time.sleep(10)
            continue

        currency = detect_currency(series)
        monthly = series_to_monthly(series, currency)
        n, missing, max_gap = check_coverage(monthly)
        print(f'  Coverage: {n}/48  missing={missing}  max_gap={max_gap}  currency={currency}')

        if n >= 40 and max_gap < 3:
            saved = save_csv(ASICS_ITEM_ID, monthly)
            print(f'  PASS -> saved {saved}/48  slug={slug}')
            return slug
        else:
            print(f'  Insufficient coverage ({n}/48), trying next slug...')
            time.sleep(10)

    print('  All ASICS candidates failed')
    return None

if __name__ == '__main__':
    yeezy_ok = collect_yeezy()
    time.sleep(15)
    asics_slug = collect_asics()

    print('\n=== SUMMARY ===')
    print(f'Yeezy: {"OK" if yeezy_ok else "NEEDS CHECK"}')
    print(f'ASICS: {asics_slug if asics_slug else "FAILED"}')
