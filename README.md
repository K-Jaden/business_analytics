# 미디어 채널은 리셀 시장 가격을 선행하는가?

자산 유형별 Granger 인과성 분석과 머신러닝 검증 — 스니커즈 · 트레이딩 카드 · 레고 48개월 패널 연구

**저자**: 이승현 · 윤재은 · 최형호 (3인 팀 연구)

---

## 한 줄 요약

미디어 채널(검색·뉴스·유튜브)이 리셀 가격을 **표본 내에서는** 선행하는 것처럼 보이지만, **표본 외 예측에서는 증분 기여가 사실상 없다**는 것을 5단계 검증으로 보인 연구입니다.

---

## 연구 문제

| | 연구 문제 | 방법 | 결과 |
|---|---|---|---|
| **RQ1** | 미디어 채널이 자산 가격을 선행하는가? | 자산별 단독 Granger 15회 | raw p<0.05 **4건** → BH 보정 후 **2건** 생존 |
| **RQ2** | 유의 채널 구성이 자산마다 다른가? | 유의 채널 집합의 Jaccard 유사도 | 스니커즈–카드 0.333, 레고와는 0.000 → **자산별로 다름** |
| **RQ3** | 선별 채널이 전채널보다 예측 성능이 좋은가? | XGBoost A/B + Diebold–Mariano 검정 | 15개 중 **11개(73%) 차이 없음** |

---

## 핵심 결과 — 표본 내 유의성과 표본 외 예측력의 괴리

이 연구의 중심 결론은 "미디어가 가격을 예측한다"가 **아니라**, 그 통념이 검증을 통과하지 못한다는 것입니다.

### 1. Granger 검정은 소수만 통과 (`results/granger_results.csv`)

| 자산 | 채널 | lag | F | p |
|---|---|--:|--:|--:|
| sneakers | CH3 뉴스감성 | 1 | 13.38 | 0.0007 |
| cards | CH1 검색트렌드 | 1 | 8.42 | 0.0058 |
| sneakers | CH4 YT조회수 | 1 | 4.73 | 0.0352 |
| sneakers | CH1 검색트렌드 | 2 | 3.47 | 0.0407 |
| 나머지 11건 | — | — | <2 | >0.18 |

**Benjamini–Hochberg FDR 보정(15개 기준) 후 생존은 2건**(sneakers–CH3, cards–CH1)뿐입니다.

### 2. 강건성 검정 3종으로 재차 축소

- **래그 견고성**: sneakers–CH3는 3개 래그에서 견고. sneakers–CH5는 L=1에서 무의미하나 L=2,3에서 유의 → **지연 효과 신규 발견**
- **Dumitrescu–Hurlin 패널 Granger**: sneakers–CH3가 비유의로 약화(자산 평균 집계가 신호를 증폭했을 가능성). BH 보정 후 일관 생존은 **cards–CH1 한 조합**
- **VAR 블록 인과**(채널 간 다중공선성 통제): sneakers(p=0.0009)·cards(p=0.0017) 유의, lego 비유의 (`results/var_block.csv`)
- **충격반응함수(IRF)**: 둘 다 1~2개월 선행 후 4개월 내 감쇄. **부호가 정반대** — cards–CH1 양(+10.80), sneakers–CH3 음(−2.82), 95% 부트스트랩 CI가 0을 포함하지 않음

### 3. 절제 실험 — 채널의 증분 기여는 0에 수렴 (`results/ablation_results.csv`)

| 모델 | 피처 | sneakers AUC | cards AUC | lego AUC |
|---|---|--:|--:|--:|
| **Channels-only** (채널 5개만) | 5 | **0.505** | **0.474** | **0.532** |
| **Baseline** (가격 시차만) | 4 | 0.837 | 0.863 | 0.963 |
| **Model A** (가격시차 + 전채널) | 9 | 0.841 | 0.859 | 0.960 |
| **A-dropGranger** (Granger 유의채널 제거) | 6~9 | 0.847 | 0.867 | 0.960 |

- 채널만으로는 **AUC 0.47~0.53 = 동전 던지기 수준**
- 가격 자기시차만 쓴 Baseline과 전채널 Model A가 **DM 검정상 구분되지 않음**
- Granger 유의 채널을 **빼도 성능이 떨어지지 않음** (cards·lego는 오히려 소폭 상승)

> **결론: Granger 유의 ≠ 표본 외 예측 유용성.** 표본 내 시계열 인과성이 실무적 예측 가치를 담보하지 않는다는 것을 같은 데이터에서 직접 보였습니다.

### 4. 벤치마크 비교 (`results/model_comparison_auc.csv`)

| 자산 | XGBoost | CatBoost | RandomForest | ElasticNet |
|---|--:|--:|--:|--:|
| sneakers | 0.841 | **0.960** | 0.923 | 0.601 |
| cards | 0.859 | **0.975** | 0.952 | 0.663 |
| lego | **0.960** | 0.956 | 0.937 | 0.706 |

트리 앙상블이 선형 기준선을 크게 상회 — 가격 시차의 비선형 구조가 신호의 대부분입니다.

---

## 데이터

**15개 아이템 × 48개월 (2022-01 ~ 2025-12) 월별 패널**, 전량 직접 수집했습니다.

| 자산군 | 출처 | 조건 | 아이템 |
|---|---|---|---|
| 스니커즈 | StockX | US Size 10 | Jordan 1 Bordeaux, Dunk Panda, Yeezy 350 Zebra, Travis Scott Jordan 1, NB 550 |
| 트레이딩 카드 | PriceCharting | PSA 10 | Charizard VMAX, Umbreon VMAX, Rayquaza VMAX, Pikachu VMAX, Charizard GX |
| 레고 | BrickRanker | New Sealed | Millennium Falcon, Hogwarts Castle, Titanic, Porsche 911, Bugatti Chiron |

### 미디어 채널 5종

| ID | 채널 | 변환 | 범위 |
|---|---|---|---|
| CH1 | Google Trends | pytrends 월평균 | 0~100 |
| CH2 | 뉴스 보도량 | GDELT timelinevol 월평균 → 아이템 내 정규화 | 0~1 |
| CH3 | 뉴스 감성 | GDELT timelinetone 월평균 | −1~+1 |
| CH4 | YouTube 조회수 | 해당 월 게시 영상 조회수 합 → log1p → 아이템 내 z-score | z |
| CH5 | YouTube 댓글 감성 | FinBERT P(pos)−P(neg) 월평균 | −1~+1 |

수집 난점과 해결은 `CLAUDE.md`의 **소스 변경 이력**에 전부 기록돼 있습니다 — 레고 아이템 3차례 교체(검색량 부족·출시일 결측), StockX Cloudflare 대응, YouTube API 할당량 설계(116 units/월 × 720개월 ≈ 9일 분할 실행) 등.

---

## 실행 순서

```bash
# 0. 데이터 연속성 사전 검증
python scripts/step00_validate.py

# 1. 수집
python scripts/collect_trends.py       # CH1 Google Trends
python scripts/collect_gdelt.py        # CH2·CH3 GDELT
python scripts/collect_youtube.py      # CH4·CH5 YouTube (진행상태 JSON 재개 지원)
python scripts/collect_sneakers.py     # StockX
node   scripts/collect_cards.js        # PriceCharting
node   scripts/collect_lego.js         # BrickRanker

# 2. 점수화 · 패널 구축
python scripts/build_channel_scores.py # FinBERT 포함 CH1~5 점수
python scripts/build_panel.py          # 패널 + 자산별 대표 시계열 + z-score 통일본

# 3. RQ1·RQ2 — 시계열 인과성
python scripts/run_adf.py              # ADF 정상성
python scripts/run_granger.py          # Granger 15회 + BH 보정
python scripts/run_granger_robustness.py  # 래그 견고성
python scripts/run_granger_dh.py       # Dumitrescu-Hurlin 패널
python scripts/run_var_block.py        # VAR 블록 인과
python scripts/run_irf.py              # 충격반응함수

# 4. RQ3 — 머신러닝 검증
python scripts/tune_xgboost.py         # 하이퍼파라미터 튜닝
python scripts/run_model_c.py          # 폴드 내 SHAP 선별 (Model C)
python scripts/run_dm_fixed_selection.py  # 고정선별 DM 검정
python scripts/run_model_comparison.py # CatBoost·RF·ElasticNet 벤치마크
python scripts/ablation.py             # 절제 실험

# 5. 논문 그림
python paper/make_figures.py
```

---

## 저장소 구조

```
paper/
  draft.tex          논문 원고 (2,357줄 · 약 35쪽 · 부록 4종)
  draft.pdf          컴파일본
  references.bib
  fig/               논문 게재 그림 13종
scripts/             분석 파이프라인 (수집 → 점수화 → 인과성 → ML)
  presentation/      발표자료 생성 스크립트 (python-pptx)
results/             검정 결과 CSV 25종 + 그림 81종
data/processed/      패널 데이터 (원본 스케일 / z-score 통일본)
DM/                  Diebold-Mariano 검정 (팀 협업 산출물)
CLAUDE.md            연구 전 과정의 의사결정 로그 · 채널 정의 · API 제약
```

> `data/raw/`(수집 원본)와 발표자료 pptx는 용량 문제로 저장소에 포함하지 않았습니다. 수집 스크립트로 재현 가능합니다.

---

## 방법론상 유의점 (논문 limitation 명시분)

1. **소표본**: 자산별 T=48개월. Granger·VAR의 검정력이 제한적이며, 이 때문에 BH 보정 후 생존 조합이 급감합니다.
2. **z-score 누설**: `StandardScaler`를 전체 표본에서 fit — 엄밀하게는 TimeSeriesSplit의 각 fold train에서만 fit해야 합니다. N=48 특성상 실질 영향은 작다고 판단했으나 한계로 명시했습니다.
3. **CH3 대체**: 원계획(NewsAPI + FinBERT)은 무료 플랜의 30일 히스토리 제한으로 4년치 수집이 불가해 GDELT tone으로 대체 — 도메인 특화도가 낮습니다.
4. **CH4 성격**: `resell`·`review` 키워드 기반이라 '정보 탐색 의도' 지표에 가깝고, 일반 팬덤·하이프를 과소 포착할 수 있습니다.
5. **자산 평균 집계**: Granger는 아이템 5개를 평균한 대표 시계열로 수행 — DH 패널 검정에서 sneakers–CH3가 약화된 것은 이 집계가 신호를 증폭했을 가능성을 시사합니다.
6. **가격 결측 보간**: sneakers_yeezy는 StockX 샘플링 아티팩트로 4개월 결측 → 선형보간. lego_porsche·bugatti는 Google Trends 0값이 각각 21/48, 25/48로 CH1 신호가 약합니다.

---

## 기술 스택

Python (pandas, numpy, statsmodels, scikit-learn, xgboost, catboost, shap, transformers/FinBERT, pytrends, matplotlib) · Node.js (Playwright, undetected-chromedriver) · LaTeX (kotex)

---

## 라이선스

MIT (`LICENSE` 참조). 논문 원고 및 그림의 저작권은 저자 3인에게 있습니다.
