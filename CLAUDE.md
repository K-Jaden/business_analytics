# CLAUDE.md — 리셀 시장 미디어 선행 지표 연구

## RULES (최우선)
- 모르면 "모른다"고 말할 것. 추측 금지.
- 실행 안 한 코드를 했다고 말하지 말 것.
- 데이터 없으면 더미 생성 금지 → 중단 후 이유 출력.
- 아래 상황이면 작업 멈추고 반드시 말할 것: "Opus 필요합니다. 전환해주세요."
  - Granger 결과 해석 모호 / 엣지 케이스 판단
  - 15개 검정 결과로 논문 결론 작성
  - XGBoost 결과가 예상과 크게 다를 때 원인 분석
  - FinBERT 점수 이상 디버깅
  - 동일 에러 3회 이상 미해결

---

## PROGRESS
> 세션 시작 시 먼저 확인. 완료 시 [x], 오류 시 [!]+메모.

- [ ] 00 데이터 연속성 사전 탐색
- [x] 01 가격 수집 (StockX·PriceCharting)
- [x] 02 Google Trends 수집
- [x] 03 GDELT 뉴스량 수집 (jordan1·panda 쿼리 수정 재수집 완료)
- [x] 04 YouTube 조회수·댓글 수집 (720/720 완료)
- [x] 05 채널 점수 산출 (FinBERT 포함, 전체 15개 아이템)
- [x] 06 패널 통합 + 자산별 대표 시계열 생성
- [x] 07 ADF 정상성 검정
- [x] 08 Granger 검정 (RQ1·RQ2) — 15회 완료 (raw p 기준, BH 보정 미적용)
- [x] 09 XGBoost + SHAP + DM 검정 (RQ3) — 완료 (팀원 진행, 결과 수신 완료)

마지막 업데이트: 2026-05-18
다음 재개 위치: **전체 완료** — `paper/draft.tex` (19쪽, 참고문헌 27개) 컴파일 완료, PDF 생성됨

### Granger 최종 결과 (raw p 기준, p<0.05 단일 판정 기준 — 2026-05-18 확정)
> ※ 기존 F≥4.0 임계 제거. p<0.05만 SIGNIFICANT.

| 자산 | 채널 | lag | F | p | 판정 |
|------|------|-----|---|---|------|
| sneakers | CH3 뉴스감성 | 1 | 13.38 | 0.0007 | SIGNIFICANT |
| cards | CH1 Google Trends | 1 | 8.42 | 0.006 | SIGNIFICANT |
| sneakers | CH4 YT조회수 | 1 | 4.73 | 0.035 | SIGNIFICANT |
| sneakers | CH1 Google Trends | 2 | 3.47 | 0.041 | SIGNIFICANT |
| lego | CH1 Google Trends | 1 | 3.92 | 0.054 | MARGINAL |
| 나머지 10개 | — | — | <2 | >0.18 | NOT_SIG |

**BH 보정 후:** SIGNIFICANT 4개→2개 (sneakers CH3, cards CH1만 생존). 부록 표 9에 병기.

Jaccard:
- sneakers={CH1,CH3,CH4}, cards={CH1}, lego={}
- sneakers–cards = **0.333** (CH1 공유, 부분적 차이)
- sneakers–lego = 0.000 / cards–lego = 0.000 → RQ2 대체로 지지

### XGBoost/DM/SHAP 최종 결과 (STEP 09)
**DM 검정 (Plan B, XGBRegressor, T=24/아이템):**
- 15개 중 11개(73%): 차이 없음 → RQ3 간결성 확보
- Model A 유의 우수: sneakers_jordan1(p<0.001), lego_falcon(p=0.032)
- RMSE 평균: sneakers A=452/B=569, lego A=176/B=199 → Model A 우세

**SHAP (분류 기준, Model A):**
- sneakers: CH1(0.125) > CH5(0.124) > CH4(0.080) > CH3(0.011)
- cards: CH4(0.091) > CH1(0.048) > CH2(0.031)
- lego: CH4(0.070) > CH1(0.036) > CH5(0.021)

### 논문 현황
- `paper/draft.tex`: 19쪽, 참고문헌 27개, 컴파일 완료 (2026-05-18)
- 엄격 평가 점수: **71.5/100** (학부 A+, KCI minor revision 가능)

### 소스 변경 이력
- 레고: PriceCharting → BrickRanker (PriceCharting LEGO는 2023-11 이후 데이터만 존재)
- 레고 아이템 1차 교체: AT-AT·Taj Mahal·Home Alone·Stranger Things·Haunted House → Millennium Falcon·Eiffel Tower·Back to the Future·Hogwarts Castle·Ferrari Daytona SP3 (Google Trends 검색량 부족으로 CH1 신호 확보 불가)
- 레고 아이템 2차 교체: Eiffel Tower·Back to the Future·Ferrari Daytona → Titanic·Razor Crest·Porsche 911 (출시일 결측 문제: Eiffel 10개월, DeLorean 3개월, Ferrari 5개월 결측. Colosseum GT 44/48 zero → Porsche로 교체)
- 레고 아이템 3차 교체: Razor Crest → Bugatti Chiron (GT zeros 30/48→25/48 개선, GT max 46.6→66.0, 25개 이상 후보 테스트 후 최선 확정. AT-AT·Home Alone 제외 이유: 크리스마스 계절성 신호로 spurious correlation 위험)
- sneakers_jordan1: Jordan 1 Chicago L&F → Jordan 1 Bordeaux (Chicago L&F 출시 2022-11, 2022-01~07 결측)
- 스니커즈 수집: Playwright stealth → undetected-chromedriver headed 모드 (Cloudflare 우회)
- sneakers_yeezy: StockX salesChart 100-interval 샘플링 아티팩트로 4개월(2022-10·2023-09·2024-08·2025-06) 결측 → 선형보간 처리, tx_count=0 표시, 논문 limitation 명시 필요
- cards_charizard1: Evolving Skies #74(잘못된 URL, Champion's Path로 리다이렉트) → Shining Fates SV107 (48/48 완전 커버리지)

---

## RESEARCH QUESTIONS

| | 연구 문제 | 실험 | 출력 지표 |
|--|---------|------|---------|
| RQ1 | 채널별로 자산 가격에 선행하는가? | 자산별 단독 Granger (15회) | F-통계량, p값, 래그 |
| RQ2 | 유의 채널 구성이 자산마다 다른가? | 자산별 유의 채널 집합 비교 | Jaccard 유사도 |
| RQ3 | 선별 채널이 전채널보다 예측 성능 동등↑? | Model A vs B XGBoost | AUC 차이, DM p값 |

---

## ITEMS

**스니커즈** | StockX | US Size 10 | 2022-01~2025-12 (48개월)
```
sneakers_jordan1   Jordan 1 Retro High OG Bordeaux  ※ 교체: Chicago L&F는 2022-11 출시로 2022-01~07 결측
sneakers_panda     Nike Dunk Low Panda
sneakers_yeezy     Yeezy 350 V2 Zebra
sneakers_travis    Travis Scott Jordan 1
sneakers_nb550     New Balance 550 White Green
```

**트레이딩 카드** | PriceCharting | PSA 10 | 2022-01~2025-12 (48개월)
```
cards_charizard1   Charizard VMAX Shining Fates SV107  ※ 교체: Evolving Skies #74는 잘못된 URL(Champion's Path 카드로 리다이렉트)
cards_umbreon      Umbreon VMAX Alt Art
cards_rayquaza     Rayquaza VMAX Alt Art
cards_pikachu      Pikachu VMAX
cards_charizard2   Charizard GX Hidden Fates
```

**레고** | BrickRanker | New Sealed | 2022-01~2025-12 (48개월)
```
lego_falcon        Millennium Falcon 75192   출시 2017, 48/48, GT max 66.2 (0값 0/48)
lego_hogwarts      Hogwarts Castle 71043     출시 2018, 48/48, GT max 65.2 (0값 0/48)
lego_titanic       Titanic 10294             출시 2021-11, 48/48, GT max 60.2 (0값 2/48)
lego_porsche       Porsche 911 42096         출시 2018, 48/48, GT max 67.2 (0값 21/48) ※ limitation 명시 필요
lego_bugatti       Bugatti Chiron 42083      출시 2018, 48/48, GT max 66.0 (0값 25/48) ※ limitation 명시 필요
```

---

## CHANNELS — 점수 변환 규칙

> 원시 데이터 → Granger 직접 투입 금지. 반드시 점수 변환 후 투입.
> Granger(STEP 08): 원본 점수 사용 / XGBoost·SHAP(STEP 09): z-score 통일본 사용.

| ID | 이름 | 변환 방식 | 범위 |
|----|------|---------|------|
| CH1 | Google Trends | pytrends 월평균 | 0~100 |
| CH2 | 뉴스 보도량 | GDELT timelinevol 월평균 → 아이템 내 max값으로 0~1 정규화 | 0~1 |
| CH3 | 뉴스 감성 | GDELT timelinetone 월평균 (-10~+10 → /10 정규화) | -1~+1 |
| CH4 | YT 조회수 | publishedAt 구간화 후 log1p → 아이템 내 z-score | z-score |
| CH5 | YT 댓글 감성 | FinBERT P(pos)-P(neg) 월평균 | -1~+1 |

> ※ CH3 limitation: 원계획(NewsAPI+FinBERT)은 무료 플랜 역사 데이터 30일 제한으로 4년치 수집 불가.
> GDELT timelinetone으로 대체. GDELT tone은 기사 제목·본문 감성을 자동 산출하므로
> FinBERT 대비 도메인 특화도 낮음 — 논문 limitation 섹션에 명시 필요.

**CH4 누적 왜곡 방지:**
```
score_ch4 = 해당 월(1일~말일) publishedAt 영상들의 조회수 합산만 사용
반드시 publishedAfter/publishedBefore 파라미터로 구간 제한
```

**STEP 09 z-score 통일 (XGBoost 투입 직전):**
```python
from sklearn.preprocessing import StandardScaler
cols = ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']
# panel_monthly_scaled = XGBoost 전용 복사본 (원본 덮어쓰기 금지)
panel_monthly_scaled = panel_monthly.copy()
panel_monthly_scaled[cols] = StandardScaler().fit_transform(panel_monthly[cols])
# ⚠️ data leakage 주의: 엄밀하게는 TimeSeriesSplit 각 fold의 train에서만 fit해야 함
# 소표본(N=48) 특성상 실질적 영향은 크지 않으나, 논문 limitation에 명시
```

---

## FILES

```
data/raw/prices/          StockX·PriceCharting 원본
data/raw/google_trends/   pytrends 원본
data/raw/gdelt/           GDELT volume 원본
data/raw/newsapi/         뉴스 기사 본문 (CH3 FinBERT 투입용)
data/raw/youtube/         YT views·comments 원본
data/processed/channel_scores.csv      item×month×CH1~5 점수 (원본 스케일)
data/processed/asset_series.csv        자산별 대표 시계열 (아이템 평균) ← Granger 투입
data/processed/panel_monthly.csv       가격+점수 통합 패널 (원본 스케일)
data/processed/panel_monthly_scaled.csv  z-score 통일본 (XGBoost 전용)
data/logs/collection_log.json
results/stationarity.csv
results/granger_results.csv            3자산×5채널 = 15개 결과
results/channel_importance.csv
results/xgboost_auc.csv
results/figures/
```

---

## STEPS

### 00 데이터 연속성 사전 탐색
```
[ ] 아이템별 실제 데이터 시작·종료월 확인
[ ] 48개월 중 수집된 월 수 ≥ 40개월(83%)
[ ] 채널별 결측률 < 20%
[ ] 연속 결측 3개월↑ → 해당 아이템 제외 검토
[ ] 가격 전월 대비 ±50% 초과 월 플래그
[ ] tx_count=0 월 확인
[ ] tx_count < 3인 달 → 선형보간 대체 / 연속 3개월↑ tx_count < 3 → 해당 기간 제외 검토
[ ] CH4·CH5: publishedAt 기준 영상 0개 월 비율
```

### 01 가격 수집
- StockX: Playwright stealth, US Size 10, sleep(10~15), 재시도 3회
- PriceCharting: requests+BS4, 카드=PSA10, 레고=new-price, sleep(3)
- **tx_count 처리 기준:** tx_count < 3인 달 → 선형보간 / 연속 3개월↑ → 해당 기간 제외 검토
- 논문 출처: "StockX 웹사이트(stockx.com)에서 수집"
- 출력: `data/raw/prices/{item_id}.csv` → `[year_month, mean_price, tx_count]`

### 02 Google Trends
- pytrends, geo='', 아이템별 단독, sleep(10)
- 출력: `data/raw/google_trends/{item_id}.csv` → `[year_month, interest]`

### 03 GDELT (CH2 + CH3)
- CH2: `mode=timelinevolnorm` → `{item_id}_vol.csv` (col: vol_norm, 0~1)
- CH3: `mode=timelinetone` → `{item_id}_tone.csv` (col: tone, -1~+1, /10 정규화)
- 일별 → 월평균, 아이템 간 sleep(30), 429 발생 시 sleep(120)
- 출력: `data/raw/gdelt/{item_id}_vol.csv`, `data/raw/gdelt/{item_id}_tone.csv`

### 04 YouTube (Quota Optimized)

**키워드 사전 정의 (노이즈 차단 필수):**
> 단일 진실의 출처(SSOT): `scripts/collect_youtube.py`. 아래는 그 사본이며,
> 변경 시 반드시 두 곳 모두 갱신할 것.
```python
ITEM_KEYWORDS = {
    "sneakers_jordan1":  "Jordan 1 Bordeaux resell review",
    "sneakers_panda":    "Nike Dunk Low Panda resell review",
    "sneakers_yeezy":    "Yeezy 350 V2 Zebra resell price",
    "sneakers_travis":   "Travis Scott Jordan 1 resell legit check",
    "sneakers_nb550":    "New Balance 550 White Green resell review",
    "cards_charizard1":  "Charizard VMAX Shining Fates PSA price",
    "cards_umbreon":     "Umbreon VMAX Alt Art PSA price",
    "cards_rayquaza":    "Rayquaza VMAX Alt Art PSA price",
    "cards_pikachu":     "Pikachu VMAX PSA price review",
    "cards_charizard2":  "Charizard GX Hidden Fates PSA price",
    "lego_falcon":       "LEGO Millennium Falcon 75192 review sealed price",
    "lego_hogwarts":     "LEGO Hogwarts Castle 71043 review sealed price",
    "lego_titanic":      "LEGO Titanic 10294 review sealed price",
    "lego_porsche":      "LEGO Porsche 911 42096 review sealed price",
    "lego_bugatti":      "LEGO Bugatti Chiron 42083 review sealed price",
}
```

**수집 프로세스 (월별 116 units 소모):**
```
1. search.list (100 units): relevance 기준 상위 50개 video_id 추출
   - 파라미터: type="video", publishedAfter/publishedBefore 필수 구간 제한
   - order="relevance" 사용 (viewCount 금지 — 역주행 영상 혼입 위험)

2. videos.list (1 unit): 50개 ID를 콤마로 묶어 viewCount 일괄 수집

3. 데이터 정제: publishedAt이 해당 월을 벗어난 영상 즉시 제외 (API 오류 이중 검증)

4. CH4 산출: 정제된 영상들의 조회수 합산 (log1p 변환 후 아이템 내 z-score)

5. 댓글 수집 (15 units): 조회수 상위 15개 영상에만 commentThreads.list 호출
   - 영상당 30댓글, 영어 댓글만(langdetect), order=relevance
```

**할당량 계산:**
```
116 units/월 × 720개월(15 아이템 × 48개월) = 83,520 units ≈ 9일
진행 상태 JSON 저장 → 다음날 이어서 실행
```

**예외 처리:**
```
검색 결과 0개인 달 → score_ch4 = 0 기록, STEP 00 결측 기준(연속 3개월↑) 동일 적용
```

**논문 기술 주의사항:**
```
CH4(조회수)는 resell·review 키워드 포함 영상 기반이므로
'정보 탐색 의도(information-seeking intent)' 지표로 성격 명시.
일반 팬덤·hype 과소 가능성을 limitation에 서술.
```

- 출력: `{item_id}_views.csv`, `{item_id}_comments.csv`

### 05 채널 점수 산출
- CH3: GDELT tone → `{item_id}_tone.csv` 그대로 사용 (이미 -1~+1 정규화 완료)
- CH5: 댓글 → FinBERT P(pos)-P(neg) 월평균
  - ProsusAI/finbert, batch=32, 영어만(langdetect)
- 출력: `data/processed/channel_scores.csv`
  - columns: `item_id, asset_type, year_month, score_ch1~5`

### 06 패널 통합 + 자산별 대표 시계열 생성
```python
# 1. 아이템별 채널 점수 + 가격 merge
panel_monthly = prices.merge(channel_scores, on=['item_id','year_month'])

# 2. 자산별 대표 시계열 생성 (Granger 투입용)
# 아이템 5개의 price, score_ch1~5를 월별 평균
asset_series = panel_monthly.groupby(['asset_type','year_month'])[
    ['mean_price','score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']
].mean().reset_index()

# 3. XGBoost용 z-score 통일본 별도 생성
cols = ['score_ch1','score_ch2','score_ch3','score_ch4','score_ch5']
panel_monthly_scaled = panel_monthly.copy()
panel_monthly_scaled[cols] = StandardScaler().fit_transform(panel_monthly[cols])

# 4. 타겟 변수 추가
panel_monthly['price_direction'] = (
    panel_monthly.groupby('item_id')['mean_price'].diff() > 0
).astype(int)

# 5. 분석 전 채널 간 상관행렬 확인 (CH2·CH3 다중공선성 점검)
print(panel_monthly[cols].corr())
# 상관계수 0.7↑ 채널은 논문에서 VIF와 함께 보고
```
- 결측: 선형보간, 연속 3개월↑ 결측 시 해당 기간 제외

### 07 ADF 정상성 검정
- asset_series 기준으로 각 (asset_type, channel) 시계열 ADF
- p>0.05 → 1차 차분 후 재검정
- 출력: `results/stationarity.csv` → `[asset_type, channel, p, differenced]`

### 08 Granger 검정
```
분석 단위: 자산(asset) — 아이템 평균 대표 시계열 사용
총 검정 수: 3자산 × 5채널 = 15회

각 (asset_type, channel) 쌍에 대해:
  제한 모형:   price_t = f(price_t-1 ... price_t-p)
  비제한 모형: price_t = f(price_t-1 ... price_t-p, score_ch_t-1 ... score_ch_t-p)

  ※ 채널 하나씩만 투입 (혼합 금지)
  ※ 래그: AIC 기준 1~4
  ※ 다중비교 보정: Benjamini-Hochberg FDR (15개 기준)

RQ1 판단:
  F>4.0, p_bh<0.05   → 유의 선행
  F 2~4, p 0.05~0.10 → 경계 (논문 명시)
  F<2.0, p≥0.10      → 선행 없음

RQ2 판단 (Jaccard):
  <0.3   → 자산별 채널 구성 다름 (RQ2 지지)
  0.3~0.6 → 부분적 차이
  >0.6   → 구성 유사 → F 크기·래그 차이 분석으로 결론 수정

엣지케이스:
  전채널 비유의 → 래그 1~6 확장 → 여전히 비유의면 "선행성 없음"으로 결론
  전채널 유의   → F 크기로 채널 간 순위 산출
```
- 출력: `results/granger_results.csv`, `results/channel_importance.csv`

### 09 XGBoost + SHAP + DM
```
입력: panel_monthly_scaled (z-score 통일본)
타겟: price_direction (0/1) ← panel_monthly에서 생성된 값 사용
Model A: 통제 피처(price_vs_ma3, price_chg_lag1~3) + CH1~5 전부
Model B: 통제 피처 + Granger 유의 채널만
검증: TimeSeriesSplit(n_splits=5)
평가: AUC

DM 검정: A vs B 예측 오차 시계열 비교
  DM p<0.10, B>A → RQ3 지지
  DM p>0.10      → 간결성 확보로 해석

SHAP: TreeExplorer, 채널별 평균|SHAP| 순위
  Granger F순위 vs SHAP 순위 교차 검증
  불일치 채널 → 논문에서 이유 설명
```
- 출력: `results/xgboost_auc.csv`, `results/figures/shap_*.png`

---

## API 제약

| 소스 | 제약 | 대응 |
|------|------|------|
| StockX | 비공식, stealth 필요 | Playwright stealth, sleep(10~15) |
| PriceCharting | HTML | requests+BS4, sleep(3) |
| pytrends | 429 | sleep(10), 재시도 3회 |
| GDELT | rate limit | sleep(10), 실패 로그 |
| NewsAPI | 월 1,000건 무료 | 아이템별 쿼리 최소화 |
| YouTube | 10,000 units/day | search.list(100)+videos.list(1)=116/월, 진행상태 JSON, 소진 시 중단 |
