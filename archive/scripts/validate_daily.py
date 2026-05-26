"""매일 수집 후 데이터 품질 점검
출력: 각 채널·아이템의 현황, 문제 항목 명시
"""
import os, csv, json
from collections import defaultdict

ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]
ITEMS = ['sneakers_jordan1','sneakers_panda','sneakers_yeezy','sneakers_travis','sneakers_nb550',
         'cards_charizard1','cards_umbreon','cards_rayquaza','cards_pikachu','cards_charizard2',
         'lego_atat','lego_taj','lego_homealone','lego_stranger','lego_haunted']

PASS = '✓'; WARN = '!'; FAIL = '✗'

issues = []

def read_csv_col(path, col):
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        return {r['year_month']: r[col] for r in csv.DictReader(f) if col in r}

def coverage(data_dict):
    present = [m for m in ALL_MONTHS if m in data_dict and data_dict[m] not in (None,'','None')]
    gaps, max_gap, run = [], 0, 0
    for m in ALL_MONTHS:
        if m not in present or data_dict.get(m) in (None,'','None'):
            run += 1; max_gap = max(max_gap, run)
        else:
            run = 0
    missing = [m for m in ALL_MONTHS if m not in present]
    return len(present), missing, max_gap

print('='*60)
print('DAILY DATA VALIDATION REPORT')
print('='*60)

# ── CH1: Google Trends ───────────────────────────────────────────────────────
print('\n[CH1] Google Trends')
ch1_ok = 0
for item in ITEMS:
    d = read_csv_col(f'data/raw/google_trends/{item}.csv', 'interest')
    n, missing, gap = coverage(d)
    flag = PASS if n == 48 else (WARN if n >= 40 else FAIL)
    if flag != PASS:
        issues.append(f'CH1 {item}: {n}/48')
    else:
        ch1_ok += 1
print(f'  {ch1_ok}/15 items complete (48/48 months)')

# ── CH2: GDELT 뉴스량 ─────────────────────────────────────────────────────────
print('\n[CH2] GDELT 뉴스량 (vol_norm)')
ch2_ok = 0
for item in ITEMS:
    d = read_csv_col(f'data/raw/gdelt/{item}_vol.csv', 'vol_norm')
    n, missing, gap = coverage(d)
    vals = [float(v) for v in d.values() if v not in (None,'','None')]
    zero_pct = round(sum(1 for v in vals if float(v)==0)/len(vals)*100, 0) if vals else 100
    flag = PASS if n==48 and zero_pct < 80 else (WARN if n>=40 else FAIL)
    if flag != PASS:
        issues.append(f'CH2 {item}: {n}/48, zeros={zero_pct}%')
    if n == 48: ch2_ok += 1
    elif n > 0:
        print(f'  {flag} {item}: {n}/48 months  zeros={zero_pct}%')
print(f'  {ch2_ok}/15 items complete')

# ── CH3: GDELT 감성 ───────────────────────────────────────────────────────────
print('\n[CH3] GDELT 뉴스감성 (tone)')
ch3_ok = 0
for item in ITEMS:
    d = read_csv_col(f'data/raw/gdelt/{item}_tone.csv', 'tone')
    n, missing, gap = coverage(d)
    vals = [float(v) for v in d.values() if v not in (None,'','None')]
    out_of_range = sum(1 for v in vals if not (-1<=v<=1)) if vals else 0
    flag = PASS if n==48 and out_of_range==0 else (WARN if n>=40 else FAIL)
    if flag != PASS:
        issues.append(f'CH3 {item}: {n}/48, out_of_range={out_of_range}')
    if n == 48: ch3_ok += 1
    elif n > 0:
        print(f'  {flag} {item}: {n}/48 months')
print(f'  {ch3_ok}/15 items complete')

# ── CH4: YouTube 조회수 ───────────────────────────────────────────────────────
print('\n[CH4] YouTube 조회수 (views_raw)')
ch4_complete, ch4_partial, ch4_missing_items = 0, 0, []
progress = {}
if os.path.exists('data/logs/youtube_progress.json'):
    with open('data/logs/youtube_progress.json') as f:
        progress = json.load(f)

for item in ITEMS:
    raw_path = f'data/raw/youtube/{item}_views_raw.csv'
    d = read_csv_col(raw_path, 'raw_views')
    n, missing, gap = coverage(d)
    done_months = sum(1 for s in progress.get(item,{}).values() if s=='done')
    err_months  = sum(1 for s in progress.get(item,{}).values() if s=='error')

    zero_months = sum(1 for v in d.values() if v not in (None,'','None') and int(v)==0)
    zero_pct = round(zero_months/n*100,0) if n > 0 else 0

    if n == 48:
        ch4_complete += 1
        if zero_pct > 50:
            issues.append(f'CH4 {item}: 48/48 but {zero_pct}% zero-view months (keyword 확인)')
    elif n > 0:
        ch4_partial += 1
        print(f'  ~ {item}: {n}/48 months  zeros={zero_pct}%  errors={err_months}')
        if err_months > 0:
            issues.append(f'CH4 {item}: {err_months} error months (내일 재시도)')
    else:
        ch4_missing_items.append(item)

total_done = sum(1 for i in progress.values() for s in i.values() if s=='done')
total_err  = sum(1 for i in progress.values() for s in i.values() if s=='error')
print(f'  {ch4_complete}/15 complete  |  {ch4_partial} partial  |  {len(ch4_missing_items)} not started')
print(f'  Overall: {total_done}/720 done, {total_err} errors (내일 재시도)')

# ── CH5: YouTube 댓글 ─────────────────────────────────────────────────────────
print('\n[CH5] YouTube 댓글 (comments)')
ch5_ok = 0
for item in ITEMS:
    path = f'data/raw/youtube/{item}_comments.csv'
    if not os.path.exists(path):
        continue
    with open(path, encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
    months_with_comments = len(set(r['year_month'] for r in rows))
    avg_per_month = len(rows) / months_with_comments if months_with_comments else 0
    if months_with_comments == 48: ch5_ok += 1
    elif months_with_comments > 0:
        print(f'  ~ {item}: {months_with_comments}/48 months, avg {avg_per_month:.0f} comments/month')

print(f'  {ch5_ok}/15 items complete')

# ── 가격 데이터 ───────────────────────────────────────────────────────────────
print('\n[Price] 가격 데이터')
price_ok = 0
for item in ITEMS:
    d = read_csv_col(f'data/raw/prices/{item}.csv', 'mean_price')
    n, _, _ = coverage(d)
    if n >= 44: price_ok += 1
print(f'  {price_ok}/15 items OK (≥44/48 months)')

# ── 전체 요약 ─────────────────────────────────────────────────────────────────
print('\n' + '='*60)
print('SUMMARY')
print('='*60)
channels = {
    'CH1 Google Trends': ch1_ok,
    'CH2 GDELT 뉴스량':  ch2_ok,
    'CH3 GDELT 감성':    ch3_ok,
    'CH4 YouTube 조회수 (완료)': ch4_complete,
    'CH5 YouTube 댓글 (완료)':   ch5_ok,
    'Price':             price_ok,
}
for ch, n in channels.items():
    bar = '█' * n + '░' * (15-n)
    flag = PASS if n==15 else (WARN if n>0 else FAIL)
    print(f'  {flag} {ch:25s} {bar} {n}/15')

pipeline_ready = ch1_ok==15 and ch2_ok==15 and ch3_ok==15 and ch4_complete==15 and price_ok==15
print(f'\n  파이프라인 준비: {"✓ 모든 채널 완료 — STEP 05~09 진행 가능" if pipeline_ready else "수집 진행 중..."}')

if issues:
    print('\n[요주의 항목]')
    for iss in issues:
        print(f'  ! {iss}')
else:
    print('\n  문제 항목 없음')

# ── error 항목 자동 리셋 (내일 재시도용) ─────────────────────────────────────
if total_err > 0:
    reset_count = 0
    for item in progress:
        for ym in list(progress[item]):
            if progress[item][ym] == 'error':
                del progress[item][ym]
                reset_count += 1
    with open('data/logs/youtube_progress.json','w') as f:
        json.dump(progress, f, indent=2)
    print(f'\n  [{reset_count} error 항목 → 미수집으로 리셋. 내일 실행 시 자동 재시도]')
