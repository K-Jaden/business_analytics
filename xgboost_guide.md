# XGBoost 분석 가이드

## 데이터 파일

`data/processed/xgboost_data.csv`

- 705 rows (15개 아이템 × 47개월, 2022-02 ~ 2025-12)
- 2022-01 제외: 첫 달은 price_direction 정의 불가(이전 달 데이터 없음)
- 모든 채널 점수는 **z-score 정규화 완료** — 추가 스케일링 불필요

---

## 컬럼 설명

### 식별자
| 컬럼 | 설명 |
|------|------|
| `item_id` | 아이템 고유 ID (예: sneakers_jordan1, lego_falcon) |
| `asset_type` | 자산 유형: sneakers / cards / lego |
| `year_month` | 관측 월 (YYYY-MM) |

### 참고용 (피처 미사용 권장)
| 컬럼 | 설명 |
|------|------|
| `mean_price` | 해당 월 평균 거래가격 (원본 스케일, 달러) |
| `tx_count` | 해당 월 거래 건수 (3 미만은 선형보간 처리됨) |

### 타겟 변수
| 컬럼 | 설명 |
|------|------|
| `price_direction` | **예측 대상** — 전월 대비 가격 상승이면 1, 하락/동일이면 0 |

### 통제 피처 (Control Features)
| 컬럼 | 설명 |
|------|------|
| `price_vs_ma3` | 현재 가격 / 3개월 이동평균 (가격 추세 위치) |
| `price_chg_lag1` | 1개월 전 가격 변화율 |
| `price_chg_lag2` | 2개월 전 가격 변화율 |
| `price_chg_lag3` | 3개월 전 가격 변화율 |

### 채널 피처 (Channel Features) — z-score 정규화됨
| 컬럼 | 채널 | 원본 범위 | 설명 |
|------|------|----------|------|
| `score_ch1` | Google Trends | 0~100 | 해당 아이템 검색량 (pytrends 월평균) |
| `score_ch2` | 뉴스 보도량 | 0~1 | GDELT timelinevol 월평균 → 아이템 내 max 정규화 |
| `score_ch3` | 뉴스 감성 | -1~+1 | GDELT timelinetone 월평균 ÷ 10 |
| `score_ch4` | YouTube 조회수 | z-score | 키워드 관련 영상 조회수 합산 → log1p → z-score |
| `score_ch5` | YouTube 댓글 감성 | -1~+1 | FinBERT P(pos) - P(neg) 월평균 |

### 모델 설정 참고
| 컬럼 | 설명 |
|------|------|
| `model_b_channels` | 해당 자산의 Granger 유의 채널 목록 (Model B 피처 선택용) |

---

## 모델 설계

### Model A — 전채널 모델
```
피처: price_vs_ma3, price_chg_lag1~3, score_ch1~5
타겟: price_direction
```

### Model B — Granger 선별 채널 모델
```
피처: price_vs_ma3, price_chg_lag1~3 + Granger 유의 채널만
타겟: price_direction
```

**자산별 Model B 채널 (model_b_channels 컬럼 참고):**

| 자산 | 유의 채널 | Granger F | p값 | 판정 |
|------|----------|-----------|-----|------|
| sneakers | score_ch3, score_ch4 | 13.38, 4.73 | 0.0007, 0.035 | SIGNIFICANT |
| cards | score_ch1 | 8.42 | 0.006 | SIGNIFICANT |
| lego | score_ch1 | 3.92 | 0.054 | MARGINAL (경계) |

---

## 권장 구현 방향

### 1. 데이터 분할
```python
from sklearn.model_selection import TimeSeriesSplit

# 아이템별로 시간 순 정렬 필수
df = df.sort_values(['item_id', 'year_month'])

# TimeSeriesSplit (n_splits=5)
tscv = TimeSeriesSplit(n_splits=5)
```

### 2. XGBoost 학습
```python
import xgboost as xgb
from sklearn.metrics import roc_auc_score

model_a_features = [
    'price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
    'score_ch1', 'score_ch2', 'score_ch3', 'score_ch4', 'score_ch5'
]
model_b_features = {
    'sneakers': ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch3', 'score_ch4'],
    'cards':    ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch1'],
    'lego':     ['price_vs_ma3', 'price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3',
                 'score_ch1'],
}

clf = xgb.XGBClassifier(n_estimators=100, max_depth=3, learning_rate=0.05,
                         random_state=42, eval_metric='auc')
```

### 3. 평가 지표
- **AUC** (ROC-AUC): 주요 지표
- DM 검정 (Diebold-Mariano): Model A vs B 예측 오차 비교
  - DM p < 0.10 이고 B AUC >= A AUC → RQ3 지지 (선별 채널이 전채널 이상)
  - DM p > 0.10 → "간결성 확보"로 해석 가능

### 4. SHAP 분석
```python
import shap

explainer = shap.TreeExplainer(clf)
shap_values = explainer.shap_values(X_test)

# 채널별 평균 |SHAP| 순위 산출
# Granger F 순위와 비교 — 불일치 채널은 논문에서 원인 서술
```

---

## 주의사항

1. **z-score 이미 적용됨** — score_ch1~5에 추가 StandardScaler 적용 금지
2. **시계열 분할 필수** — 랜덤 split 사용 시 data leakage 발생
3. **아이템별 독립 학습 vs 풀링 선택** — 논문 설계에 따라 결정
   - 풀링 (전체 720행): 샘플 수 확보, 아이템 간 이질성 무시
   - 아이템별 (48행): 이질성 반영, 소표본 문제
4. **lego Model B**: Granger MARGINAL(p=0.054) 채널 포함 — 논문에서 근거 명시 필요
5. **결측 처리**: score_ch2·ch3 일부 NaN 존재 — XGBoost는 NaN 자체 처리 가능하므로 별도 imputation 불필요
6. **출력 저장 위치**: `results/xgboost_auc.csv`, `results/figures/shap_*.png`
