"""StockX ASICS 브라우저 탐색 + Yeezy 300-interval 재시도"""
import undetected_chromedriver as uc
import time, json, re

CHROME_VERSION = 147

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

    # ASICS 브랜드 페이지
    driver.get('https://stockx.com/sneakers/asics?sort=featured&order=DESC')
    print('ASICS page... 10s')
    time.sleep(10)

    # 페이지 소스에서 슬러그 추출
    src = driver.page_source

    # href="/에서 asics 포함 URL 추출
    slugs = re.findall(r'href="/(asics-[^"?]+)"', src)
    slugs = list(dict.fromkeys(slugs))  # 중복 제거
    print(f'\nFound {len(slugs)} ASICS slugs:')
    for s in slugs[:30]:
        print(f'  {s}')

    # Yeezy: network log에서 실제 쿼리 캡처
    print('\n\nCapturing Yeezy fetchSalesGraph query...')
    driver.get_log('performance')
    driver.get('https://stockx.com/adidas-yeezy-boost-350-v2-white-core-black-red')
    time.sleep(15)
    for y in [500,1000,1500,2000]:
        driver.execute_script(f'window.scrollTo(0,{y})')
        time.sleep(2)
    time.sleep(5)

    captured_query = None
    captured_vars = None
    for log in driver.get_log('performance'):
        try:
            msg = json.loads(log['message'])
            if msg['message']['method'] != 'Network.requestWillBeSent':
                continue
            req = msg['message']['params'].get('request', {})
            post = req.get('postData', '')
            if not post:
                continue
            gql = json.loads(post)
            if gql.get('operationName') != 'fetchSalesGraph':
                continue
            captured_query = gql.get('query', '')
            captured_vars = gql.get('variables', {})
            print(f'  Captured query! vars={captured_vars}')
            break
        except Exception:
            pass

    if captured_query:
        # 실제 쿼리로 intervals=300 재요청
        print('  Re-requesting with intervals=300 via browser fetch...')
        new_vars = dict(captured_vars)
        new_vars['intervals'] = 300
        js = """
        var done = arguments[arguments.length-1];
        var q = arguments[0];
        var v = arguments[1];
        fetch('https://stockx.com/api/p/e', {
            method:'POST', credentials:'include',
            headers:{'Content-Type':'application/json','apollographql-client-name':'Iron'},
            body: JSON.stringify({operationName:'fetchSalesGraph', variables:v, query:q})
        }).then(r=>r.json())
          .then(d=>done(JSON.stringify(d)))
          .catch(e=>done(JSON.stringify({error:e.toString()})));
        """
        result = driver.execute_async_script(js, captured_query, new_vars)
        if result:
            parsed = json.loads(result)
            series = (parsed.get('data',{}).get('product',{})
                           .get('salesChart',{}).get('series',[]))
            print(f'  intervals=300 result: {len(series)} points')
            # 월별 커버리지 체크
            months = set()
            for pt in series:
                xv = pt['xValue'][:7]  # YYYY-MM
                if '2022-01' <= xv <= '2025-12':
                    months.add(xv)
            print(f'  Unique months covered: {len(months)}/48')
            if len(series) > 0:
                with open('yeezy_series_300.json','w') as f:
                    json.dump(series, f)
                print('  Saved to yeezy_series_300.json')
        else:
            print('  No result from intervals=300 fetch')
    else:
        print('  fetchSalesGraph not captured')

finally:
    try: driver.quit()
    except: pass
