# 데이터 수집 가이드

## 환경
```bash
pip install requests pandas beautifulsoup4 greenlet --only-binary :all:
pip install undetected-chromedriver selenium
```
Chrome 버전 확인: `chrome://version` → 숫자 메모 (version_main 파라미터에 사용)

---

## 아이템 목록

### 스니커즈 — StockX (US Size 10)
| item_id | slug |
|---------|------|
| sneakers_jordan1 | `air-jordan-1-retro-high-bordeaux` |
| sneakers_panda | `nike-dunk-low-retro-white-black-2021` |
| sneakers_yeezy | `adidas-yeezy-boost-350-v2-white-core-black-red` |
| sneakers_travis | `air-jordan-1-retro-high-travis-scott` |
| sneakers_nb550 | `new-balance-550-white-green` |

### 카드 — PriceCharting (PSA 10 graded)
| item_id | url_path |
|---------|----------|
| cards_charizard1 | `/game/pokemon-shining-fates/charizard-vmax-sv107` |
| cards_umbreon | `/game/pokemon-evolving-skies/umbreon-vmax-215` |
| cards_rayquaza | `/game/pokemon-evolving-skies/rayquaza-vmax-218` |
| cards_pikachu | `/game/pokemon-vivid-voltage/pikachu-vmax-188` |
| cards_charizard2 | `/game/pokemon-hidden-fates/charizard-gx-sv49` |

### 레고 — BrickRanker (New Sealed)
| item_id | url_path |
|---------|----------|
| lego_atat | `/rankings/set/75313-1/at-at` |
| lego_taj | `/rankings/set/10256-1/taj-mahal` |
| lego_homealone | `/rankings/set/21330-1/home-alone` |
| lego_stranger | `/rankings/set/75810-1/stranger-things` |
| lego_haunted | `/rankings/set/10273-1/haunted-house` |

출력 형식: `data/raw/prices/{item_id}.csv` — 컬럼: `year_month, mean_price, tx_count`
범위: 2022-01 ~ 2025-12

---

## 수집 코드

### collect_sneakers.py
```python
import undetected_chromedriver as uc
import time, json, os, csv
from datetime import datetime
from collections import defaultdict

CHROME_VERSION = 147  # chrome://version 확인 후 수정

SNEAKERS = {
    'sneakers_jordan1': 'air-jordan-1-retro-high-bordeaux',
    'sneakers_panda':   'nike-dunk-low-retro-white-black-2021',
    'sneakers_yeezy':   'adidas-yeezy-boost-350-v2-white-core-black-red',
    'sneakers_travis':  'air-jordan-1-retro-high-travis-scott',
    'sneakers_nb550':   'new-balance-550-white-green',
}

def parse_xvalue(xv):
    try:
        return datetime.utcfromtimestamp(int(xv) / 1000)
    except (ValueError, TypeError):
        return datetime.fromisoformat(str(xv)[:19])

def collect_item(driver, slug, item_id, retries=2):
    for attempt in range(retries):
        driver.get_log('performance')
        driver.get(f'https://stockx.com/{slug}')
        time.sleep(14)
        for y in [500, 1000, 1500, 2000]:
            driver.execute_script(f'window.scrollTo(0, {y})')
            time.sleep(3)
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
                series = json.loads(body.get('body', '{}')).get('data', {}).get('product', {}).get('salesChart', {}).get('series', [])
                break
            except Exception:
                pass

        if series:
            break
        print(f'  [{item_id}] attempt {attempt+1} failed, retrying...')
        time.sleep(15)

    if not series:
        print(f'  [{item_id}] FAILED — no fetchSalesGraph data')
        return

    monthly = defaultdict(list)
    for pt in series:
        d = parse_xvalue(pt['xValue'])
        if d:
            monthly[d.strftime('%Y-%m')].append(pt['yValue'])

    os.makedirs('data/raw/prices', exist_ok=True)
    rows = [{'year_month': ym, 'mean_price': round(sum(v)/len(v), 2), 'tx_count': len(v)}
            for ym in sorted(monthly) if '2022-01' <= ym <= '2025-12']

    with open(f'data/raw/prices/{item_id}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader(); w.writerows(rows)
    print(f'  [{item_id}] {len(rows)}/48 months saved')

def main():
    opts = uc.ChromeOptions()
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--lang=en-US')
    opts.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    driver = uc.Chrome(options=opts, version_main=CHROME_VERSION, headless=False)
    try:
        driver.get('https://stockx.com')
        time.sleep(12)
        for item_id, slug in SNEAKERS.items():
            print(f'\n[{item_id}]')
            collect_item(driver, slug, item_id)
            time.sleep(12)
    finally:
        driver.quit()

if __name__ == '__main__':
    main()
```

### collect_cards.py
```python
import requests, re, json, time, os, csv
from datetime import datetime

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
BASE = 'https://www.pricecharting.com'

CARDS = {
    'cards_charizard1': '/game/pokemon-shining-fates/charizard-vmax-sv107',
    'cards_umbreon':    '/game/pokemon-evolving-skies/umbreon-vmax-215',
    'cards_rayquaza':   '/game/pokemon-evolving-skies/rayquaza-vmax-218',
    'cards_pikachu':    '/game/pokemon-vivid-voltage/pikachu-vmax-188',
    'cards_charizard2': '/game/pokemon-hidden-fates/charizard-gx-sv49',
}

def collect(item_id, url_path):
    r = requests.get(BASE + url_path, headers=HEADERS, timeout=20)
    if r.status_code != 200:
        print(f'{item_id}: HTTP {r.status_code}'); return
    m = re.search(r'VGPC\.chart_data\s*=\s*(\{.*?\});', r.text, re.DOTALL)
    if not m:
        print(f'{item_id}: chart_data not found'); return
    data = json.loads(m.group(1))
    graded = [(ts, v) for ts, v in data.get('graded', []) if v > 0]
    os.makedirs('data/raw/prices', exist_ok=True)
    rows = [{'year_month': datetime.utcfromtimestamp(ts/1000).strftime('%Y-%m'),
             'mean_price': round(v/100, 2), 'tx_count': 1}
            for ts, v in graded
            if '2022-01' <= datetime.utcfromtimestamp(ts/1000).strftime('%Y-%m') <= '2025-12']
    with open(f'data/raw/prices/{item_id}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader(); w.writerows(rows)
    print(f'{item_id}: {len(rows)}/48 months saved')

for item_id, path in CARDS.items():
    collect(item_id, path)
    time.sleep(3)
```

### collect_lego.py
```python
import requests, re, time, os, csv
from datetime import datetime

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
BASE = 'https://brickranker.com'

LEGO = {
    'lego_atat':      '/rankings/set/75313-1/at-at',
    'lego_taj':       '/rankings/set/10256-1/taj-mahal',
    'lego_homealone': '/rankings/set/21330-1/home-alone',
    'lego_stranger':  '/rankings/set/75810-1/stranger-things',
    'lego_haunted':   '/rankings/set/10273-1/haunted-house',
}

def label_to_ym(label):
    try:
        return datetime.strptime(label.strip(), '%b %y').strftime('%Y-%m')
    except Exception:
        return None

def collect(item_id, url_path):
    r = requests.get(BASE + url_path, headers=HEADERS, timeout=20)
    if r.status_code != 200:
        print(f'{item_id}: HTTP {r.status_code}'); return
    ml = re.search(r'labels\s*=\s*\[([^\]]+)\]', r.text)
    mp = re.search(r'(?:var\s+)?prices\s*=\s*\[([^\]]+)\]', r.text)
    if not ml or not mp:
        print(f'{item_id}: labels/prices not found'); return
    labels = re.findall(r"'([^']+)'", ml.group(1))
    prices = re.findall(r'[\d.]+', mp.group(1))
    os.makedirs('data/raw/prices', exist_ok=True)
    rows = [{'year_month': ym, 'mean_price': float(p), 'tx_count': 1}
            for lbl, p in zip(labels, prices)
            if (ym := label_to_ym(lbl)) and '2022-01' <= ym <= '2025-12']
    with open(f'data/raw/prices/{item_id}.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month', 'mean_price', 'tx_count'])
        w.writeheader(); w.writerows(rows)
    print(f'{item_id}: {len(rows)}/48 months saved')

for item_id, path in LEGO.items():
    collect(item_id, path)
    time.sleep(2)
```

---

## 실행 순서

```bash
mkdir -p data/raw/prices
python collect_cards.py   # ~1분
python collect_lego.py    # ~30초
python collect_sneakers.py  # ~15분, Chrome 창 뜸 — 닫으면 안 됨
```

---

## 트러블슈팅

| 증상 | 해결 |
|------|------|
| `SessionNotCreatedException: version 14X` | `CHROME_VERSION` 을 `chrome://version` 숫자로 수정 |
| `fetchSalesGraph NOT found` | sleep 14→20, y=2500 추가 스크롤 |
| StockX yValue 비정상적으로 큼 | KRW 단위 → `/1300` 환산 또는 응답 `currencyCode` 확인 |
| PriceCharting `chart_data not found` | URL 경로 재확인, 리다이렉트 여부 체크 |
| BrickRanker `labels/prices not found` | HTML에서 변수명 직접 확인 (구조 변경 가능) |

---

## 검증 (수집 후)

```python
import pandas as pd, os

def validate(item_id):
    path = f'data/raw/prices/{item_id}.csv'
    if not os.path.exists(path):
        print(f'{item_id}: FILE MISSING'); return False
    df = pd.read_csv(path)
    df = df[(df.year_month >= '2022-01') & (df.year_month <= '2025-12')]
    n = len(df)
    all_months = pd.date_range('2022-01', periods=48, freq='MS').strftime('%Y-%m').tolist()
    consec = max_c = 0
    for m in all_months:
        if m not in df.year_month.values:
            consec += 1; max_c = max(max_c, consec)
        else:
            consec = 0
    ok = n >= 40 and max_c < 3
    print(f'{item_id}: {n}/48 max_gap={max_c} -> {"PASS" if ok else "FAIL"}')
    return ok

items = ['sneakers_jordan1','sneakers_panda','sneakers_yeezy','sneakers_travis','sneakers_nb550',
         'cards_charizard1','cards_umbreon','cards_rayquaza','cards_pikachu','cards_charizard2',
         'lego_atat','lego_taj','lego_homealone','lego_stranger','lego_haunted']
results = [validate(i) for i in items]
print(f'\n{sum(results)}/15 PASS')
```
