"""STEP 05 — 채널 점수 산출 → data/processed/channel_scores.csv

CH1: Google Trends interest (0~100, as-is)
CH2: GDELT vol (이미 0~1 정규화됨)
CH3: GDELT tone (이미 -1~+1 정규화됨)
CH4: YouTube views score_ch4 (이미 z-score)
CH5: FinBERT P(pos)-P(neg) 월평균 from YouTube comments

사전 설치:
  pip install torch transformers langdetect
  (GPU 없으면 CPU 동작, 약 10~30분 소요)
"""
import os, csv, sys
from collections import defaultdict
import pandas as pd

ITEMS = {
    'sneakers': ['sneakers_jordan1','sneakers_panda','sneakers_yeezy','sneakers_travis','sneakers_nb550'],
    'cards':    ['cards_charizard1','cards_charizard2','cards_umbreon','cards_rayquaza','cards_pikachu'],
    'lego':     ['lego_falcon','lego_hogwarts','lego_titanic','lego_porsche','lego_bugatti'],
}
ALL_MONTHS = [f'{y}-{m:02d}' for y in range(2022, 2026) for m in range(1, 13)]

# ── FinBERT 로드 ──────────────────────────────────────────────────────────────
try:
    import torch
    from transformers import AutoTokenizer, AutoModelForSequenceClassification
    print('Loading FinBERT (ProsusAI/finbert)...')
    _tokenizer = AutoTokenizer.from_pretrained('ProsusAI/finbert')
    _model     = AutoModelForSequenceClassification.from_pretrained('ProsusAI/finbert')
    _model.eval()
    _device = 'cuda' if torch.cuda.is_available() else 'cpu'
    _model.to(_device)
    print(f'  FinBERT ready on {_device}')
    FINBERT_OK = True
except ImportError:
    print('[WARN] torch/transformers not installed - CH5 will be NaN')
    print('       pip install torch transformers')
    FINBERT_OK = False

def _finbert_batch(texts, batch_size=32):
    """texts → list of [P(pos), P(neg), P(neu)]"""
    import torch
    results = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        enc = _tokenizer(batch, return_tensors='pt', truncation=True,
                         max_length=128, padding=True).to(_device)
        with torch.no_grad():
            logits = _model(**enc).logits
        probs = torch.softmax(logits, dim=-1).cpu().tolist()
        results.extend(probs)
    return results  # [[pos, neg, neu], ...]

def ch5_for_item(item_id):
    """comments CSV → {year_month: score_ch5}"""
    if not FINBERT_OK:
        return {}
    path = f'data/raw/youtube/{item_id}_comments.csv'
    if not os.path.exists(path):
        return {}

    monthly_texts = defaultdict(list)
    with open(path, encoding='utf-8') as f:
        for row in csv.DictReader(f):
            text = row['comment'].strip()
            if text and len(text) > 3:
                monthly_texts[row['year_month']].append(text)

    scores = {}
    for ym, texts in monthly_texts.items():
        if not texts:
            continue
        probs = _finbert_batch(texts)
        scores[ym] = round(sum(p[0] - p[1] for p in probs) / len(probs), 6)
    return scores

# ── CH1~4 로더 ────────────────────────────────────────────────────────────────
def _load_csv_col(path, col):
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        return {r['year_month']: (float(r[col]) if r[col] not in ('', None) else None)
                for r in csv.DictReader(f)}

def load_ch1(item_id):
    return _load_csv_col(f'data/raw/google_trends/{item_id}.csv', 'interest')

def load_ch2(item_id):
    return _load_csv_col(f'data/raw/gdelt/{item_id}_vol.csv', 'vol_norm')

def load_ch3(item_id):
    return _load_csv_col(f'data/raw/gdelt/{item_id}_tone.csv', 'tone')

def load_ch4(item_id):
    return _load_csv_col(f'data/raw/youtube/{item_id}_views.csv', 'score_ch4')

# ── 기존 CH5 캐시 로드 (FinBERT 재실행 방지) ─────────────────────────────────
CACHE_PATH = 'data/processed/channel_scores.csv'
cached_ch5 = {}  # {item_id: {year_month: score_ch5}}
if os.path.exists(CACHE_PATH):
    cache_df = pd.read_csv(CACHE_PATH)
    if 'score_ch5' in cache_df.columns:
        for _, row in cache_df.iterrows():
            cached_ch5.setdefault(row['item_id'], {})[row['year_month']] = (
                None if pd.isna(row['score_ch5']) else row['score_ch5']
            )
        print(f'CH5 캐시 로드 완료 ({len(cached_ch5)}개 아이템) — FinBERT 재실행 생략')

# ── 메인 ──────────────────────────────────────────────────────────────────────
os.makedirs('data/processed', exist_ok=True)
rows = []

for asset_type, item_list in ITEMS.items():
    for item_id in item_list:
        print(f'  {item_id}...', end=' ', flush=True)
        ch1 = load_ch1(item_id)
        ch2 = load_ch2(item_id)
        ch3 = load_ch3(item_id)
        ch4 = load_ch4(item_id)
        # CH5: 캐시 있으면 재사용, 없으면 FinBERT 실행
        ch5 = cached_ch5.get(item_id) if item_id in cached_ch5 else ch5_for_item(item_id)
        for ym in ALL_MONTHS:
            rows.append({
                'item_id':    item_id,
                'asset_type': asset_type,
                'year_month': ym,
                'score_ch1':  ch1.get(ym),
                'score_ch2':  ch2.get(ym),
                'score_ch3':  ch3.get(ym),
                'score_ch4':  ch4.get(ym),
                'score_ch5':  ch5.get(ym) if ch5 else None,
            })
        print('done')

df = pd.DataFrame(rows)
df.to_csv('data/processed/channel_scores.csv', index=False)
print(f'\nSaved {len(df)} rows → data/processed/channel_scores.csv')

print('\n결측 요약 (아이템 레벨):')
for col in ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']:
    n_miss = df[col].isna().sum()
    pct = 100 * n_miss / len(df)
    print(f'  {col}: {n_miss}/720 missing ({pct:.1f}%)')
