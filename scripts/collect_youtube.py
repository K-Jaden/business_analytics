"""STEP 04 — YouTube 조회수·댓글 수집
설계 원칙:
  - 각 item-month 완료 즉시 CSV에 append + progress 저장 (크래시 안전)
  - score_ch4(z-score)는 48개월 완료 후 별도 계산
  - ConnectionResetError 포함 모든 네트워크 오류 retry
  - 116 units/item-month × 480 = 55,680 units ≈ 6일 (sneakers+cards, lego 교체 후 추가)
"""
import os, csv, json, time, math
from calendar import monthrange
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import socket

API_KEY       = os.environ.get('YOUTUBE_API_KEY', '')  # export YOUTUBE_API_KEY=your_key
DAILY_LIMIT   = 9800
PROGRESS_FILE = 'data/logs/youtube_progress.json'
VIEWS_DIR     = 'data/raw/youtube'
os.makedirs(VIEWS_DIR, exist_ok=True)
os.makedirs('data/logs', exist_ok=True)

ITEM_KEYWORDS = {
    # sneakers (StockX, US Size 10)
    'sneakers_jordan1': 'Jordan 1 Bordeaux resell review',
    'sneakers_panda':   'Nike Dunk Low Panda resell review',
    'sneakers_yeezy':   'Yeezy 350 V2 Zebra resell price',
    'sneakers_travis':  'Travis Scott Jordan 1 resell legit check',
    'sneakers_nb550':   'New Balance 550 White Green resell review',
    # cards (PriceCharting, PSA10)
    'cards_charizard1': 'Charizard VMAX Shining Fates PSA price',
    'cards_umbreon':    'Umbreon VMAX Alt Art PSA price',
    'cards_rayquaza':   'Rayquaza VMAX Alt Art PSA price',
    'cards_pikachu':    'Pikachu VMAX PSA price review',
    'cards_charizard2': 'Charizard GX Hidden Fates PSA price',
    # lego (BrickRanker, New Sealed) — 3차 교체 확정 5종
    'lego_falcon':      'LEGO Millennium Falcon 75192 review sealed price',
    'lego_hogwarts':    'LEGO Hogwarts Castle 71043 review sealed price',
    'lego_titanic':     'LEGO Titanic 10294 review sealed price',
    'lego_porsche':     'LEGO Porsche 911 42096 review sealed price',
    'lego_bugatti':     'LEGO Bugatti Chiron 42083 review sealed price',
}

COLLECT_ITEMS = list(ITEM_KEYWORDS.keys())  # sneakers는 progress에서 자동 skip

ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022,2026) for m in range(1,13)]
ITEMS = list(ITEM_KEYWORDS.keys())

youtube = build('youtube', 'v3', developerKey=API_KEY)

# ── 진행 상태 ────────────────────────────────────────────────────────────────
def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, encoding='utf-8') as f:
            return json.load(f)
    return {item: {} for item in ITEMS}

def save_progress(prog):
    with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
        json.dump(prog, f, indent=2)

# ── 유닛 추적 ────────────────────────────────────────────────────────────────
units_used = 0

def use_units(n):
    global units_used
    units_used += n
    if units_used >= DAILY_LIMIT:
        print(f'\n[QUOTA] {units_used} units used -- stopping. Resume tomorrow.')
        save_progress(progress)
        finalize_scores()
        exit(0)

# ── API 호출 (네트워크 오류 포함 retry) ─────────────────────────────────────
def api_call(func, retries=4):
    """어떤 네트워크/API 오류든 exponential backoff retry"""
    for attempt in range(retries):
        try:
            return func()
        except HttpError as e:
            code = e.resp.status
            if code in (403, 400):
                raise  # 할당량/권한 오류는 재시도 불필요
            wait = 15 * (2 ** attempt)
            print(f'  HttpError {code}, retry {attempt+1}/{retries} in {wait}s')
            time.sleep(wait)
        except (ConnectionResetError, socket.error, OSError) as e:
            wait = 20 * (2 ** attempt)
            print(f'  Network error: {e}, retry {attempt+1}/{retries} in {wait}s')
            time.sleep(wait)
    raise RuntimeError(f'All {retries} retries failed')

# ── YouTube API 헬퍼 ─────────────────────────────────────────────────────────
def ym_to_range(ym):
    y, m = int(ym[:4]), int(ym[5:])
    last = monthrange(y, m)[1]
    return (f'{y}-{m:02d}-01T00:00:00Z',
            f'{y}-{m:02d}-{last:02d}T23:59:59Z')

def search_videos(keyword, after, before):
    use_units(100)
    resp = api_call(lambda: youtube.search().list(
        q=keyword, type='video', part='id',
        maxResults=50, order='relevance',
        publishedAfter=after, publishedBefore=before,
    ).execute())
    return [item['id']['videoId'] for item in resp.get('items', [])]

def get_view_counts(video_ids):
    if not video_ids:
        return {}
    use_units(1)
    resp = api_call(lambda: youtube.videos().list(
        part='statistics,snippet',
        id=','.join(video_ids),
    ).execute())
    result = {}
    for item in resp.get('items', []):
        vid = item['id']
        published = item['snippet'].get('publishedAt', '')
        views = int(item['statistics'].get('viewCount', 0))
        result[vid] = {'views': views, 'publishedAt': published}
    return result

def get_comments(video_id):
    use_units(1)
    try:
        resp = api_call(lambda: youtube.commentThreads().list(
            part='snippet', videoId=video_id,
            maxResults=30, order='relevance',
            textFormat='plainText',
        ).execute())
        return [item['snippet']['topLevelComment']['snippet']['textDisplay']
                for item in resp.get('items', [])]
    except HttpError:
        return []

# ── 언어 감지 ────────────────────────────────────────────────────────────────
try:
    from langdetect import detect
    def is_english(t):
        try: return detect(t) == 'en'
        except: return False
except ImportError:
    def is_english(t): return True

# ── 즉시 저장 (크래시 안전) ──────────────────────────────────────────────────
def append_views(item_id, ym, raw_views, video_count):
    path = f'{VIEWS_DIR}/{item_id}_views_raw.csv'
    write_header = not os.path.exists(path)
    with open(path, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month','raw_views','video_count'])
        if write_header: w.writeheader()
        w.writerow({'year_month': ym, 'raw_views': raw_views, 'video_count': video_count})

def append_comments(item_id, ym, comments):
    if not comments:
        return
    path = f'{VIEWS_DIR}/{item_id}_comments.csv'
    write_header = not os.path.exists(path)
    with open(path, 'a', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=['year_month','comment'])
        if write_header: w.writeheader()
        for c in comments:
            w.writerow({'year_month': ym, 'comment': c})

# ── 월 단위 수집 ─────────────────────────────────────────────────────────────
def collect_month(item_id, ym, keyword):
    after, before = ym_to_range(ym)
    try:
        video_ids = search_videos(keyword, after, before)
    except HttpError as e:
        if 'quotaExceeded' in str(e):
            print(f'  [QUOTA EXCEEDED] stopping immediately.')
            save_progress(progress)
            finalize_scores()
            exit(0)
        print(f'  search error: {e}'); return False

    if not video_ids:
        append_views(item_id, ym, 0, 0)
        return True

    try:
        stats = get_view_counts(video_ids)
    except HttpError as e:
        print(f'  videos error: {e}'); return False

    # publishedAt 검증 — 해당 월 벗어난 영상 제외
    valid = {vid: s for vid, s in stats.items()
             if s['publishedAt'][:7] == ym}
    total_views = sum(s['views'] for s in valid.values())

    # 즉시 raw 저장
    append_views(item_id, ym, total_views, len(valid))

    # 조회수 상위 15개 댓글 수집 (각 1 unit)
    top15 = sorted(valid.items(), key=lambda x: x[1]['views'], reverse=True)[:15]
    comments = []
    for vid, _ in top15:
        comments.extend([c for c in get_comments(vid) if is_english(c)])
        time.sleep(0.5)

    append_comments(item_id, ym, comments)
    print(f'  views={total_views}  videos={len(valid)}  comments={len(comments)}')
    return True

# ── z-score 최종 계산 (48개월 완료 후) ──────────────────────────────────────
def finalize_scores():
    """raw CSV → score_ch4 포함 최종 views CSV 생성"""
    for item_id in ITEMS:
        raw_path   = f'{VIEWS_DIR}/{item_id}_views_raw.csv'
        final_path = f'{VIEWS_DIR}/{item_id}_views.csv'
        if not os.path.exists(raw_path):
            continue
        with open(raw_path, encoding='utf-8') as f:
            rows = {r['year_month']: r for r in csv.DictReader(f)}

        if len(rows) < 48:
            continue  # 아직 미완료

        log_vals = [math.log1p(int(rows[m]['raw_views'])) for m in ALL_MONTHS if m in rows]
        if not log_vals:
            continue
        mean_l = sum(log_vals) / len(log_vals)
        std_l  = (sum((x - mean_l)**2 for x in log_vals) / len(log_vals)) ** 0.5 or 1.0

        with open(final_path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['year_month','raw_views','video_count','log1p_views','score_ch4'])
            w.writeheader()
            for ym in ALL_MONTHS:
                if ym not in rows:
                    continue
                rv = int(rows[ym]['raw_views'])
                vc = int(rows[ym]['video_count'])
                lv = math.log1p(rv)
                w.writerow({'year_month': ym, 'raw_views': rv, 'video_count': vc,
                            'log1p_views': round(lv, 4),
                            'score_ch4':   round((lv - mean_l) / std_l, 4)})
        print(f'  [{item_id}] score_ch4 finalized -> {final_path}')

# ── 메인 ─────────────────────────────────────────────────────────────────────
progress = load_progress()

todo = [(item_id, ym)
        for item_id in COLLECT_ITEMS
        for ym in ALL_MONTHS
        if progress.get(item_id, {}).get(ym) != 'done']

total = len(todo)
print(f'Collecting: {COLLECT_ITEMS}')
print(f'Remaining: {total} item-months  ({total*116:,} units)')
print(f'Today can process: ~{DAILY_LIMIT//116} item-months\n')

for idx, (item_id, ym) in enumerate(todo):
    print(f'[{idx+1}/{total}] {item_id} {ym}  (units used: {units_used})', flush=True)

    ok = collect_month(item_id, ym, ITEM_KEYWORDS[item_id])

    if ok:
        progress.setdefault(item_id, {})[ym] = 'done'
    else:
        progress.setdefault(item_id, {})[ym] = 'error'

    # 5개마다 progress 저장
    if (idx + 1) % 5 == 0:
        save_progress(progress)

    time.sleep(5)  # YouTube rate limit 방지

save_progress(progress)
finalize_scores()

done_count = sum(1 for i in progress.values() for s in i.values() if s == 'done')
print(f'\nDone today. Units used: {units_used}')
print(f'Total progress: {done_count}/720 item-months')
if done_count < 720:
    print('Run again tomorrow to continue.')
else:
    print('All 720 item-months complete!')
