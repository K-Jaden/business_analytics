import urllib.request, urllib.parse, json, time

PY = "C:\\Users\\이승현\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"
BASE = 'https://api.gdeltproject.org/api/v2/doc/doc'

def fetch(query, mode='timelinetone'):
    params = {
        'query': query,
        'mode': mode,
        'format': 'json',
        'startdatetime': '20220101000000',
        'enddatetime': '20220331235959',
    }
    url = BASE + '?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.loads(r.read())
    timeline = data.get('timeline', [])
    if timeline:
        d = timeline[0].get('data', [])
        return len(d), d[:3]
    return 0, []

# 1. 원래 쿼리
print('=== 원래 쿼리: "Jordan 1 Bordeaux" resell ===')
try:
    n, sample = fetch('"Jordan 1 Bordeaux" resell')
    print(f'data points: {n}')
    print(f'sample: {sample}')
except Exception as e:
    print(f'ERROR: {e}')

time.sleep(12)

# 2. 따옴표 없이 넓은 쿼리
print('\n=== 넓은 쿼리: Jordan Bordeaux resell ===')
try:
    n, sample = fetch('Jordan Bordeaux resell')
    print(f'data points: {n}')
    print(f'sample: {sample}')
except Exception as e:
    print(f'ERROR: {e}')

time.sleep(12)

# 3. 더 넓게: Jordan 1 Bordeaux만
print('\n=== 더 넓은 쿼리: "Jordan 1 Bordeaux" ===')
try:
    n, sample = fetch('"Jordan 1 Bordeaux"')
    print(f'data points: {n}')
    print(f'sample: {sample}')
except Exception as e:
    print(f'ERROR: {e}')
