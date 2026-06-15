# -*- coding: utf-8 -*-
"""
수업 발표용 추가 그림 생성
  - fig_gap_matrix.png  : 0막 연구 갭 시각화
  - fig_ablation2.png   : ablation 핵심 4모형 비교 (상대 스케일)
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG_DIR = os.path.join(ROOT, "paper", "fig")
os.makedirs(FIG_DIR, exist_ok=True)

plt.rcParams["font.family"] = ["Malgun Gothic", "AppleGothic", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False

# ── Colors ──────────────────────────────────────────────────────────
C_EXIST  = "#5B9BD5"   # 기존 연구 있음 (파란)
C_GAP    = "#E8E8E8"   # 연구 공백 (회색)
C_OURS   = "#C00000"   # 본 연구 (레드)
C_NA     = "#D5D5D5"   # 해당 없음 (연회색)
C_INK    = "#1A1A1A"
C_WHITE  = "#FFFFFF"

# ── Cell data ───────────────────────────────────────────────────────
# 행(위→아래): 멀티채널, 리셀가격, 미디어→가격
# 열(좌→우): 전통금융, 스니커즈, 카드, 레고
# (메인텍스트, 소텍스트, 배경색, 글자색)
cells = [
    [  # row 0 — 멀티채널 미디어 분석
        ("✓ 기존",     "Bollen (2011)",       C_EXIST, C_WHITE),
        ("★ 본 연구",  "최초 시도",            C_OURS,  C_WHITE),
        ("★ 본 연구",  "최초 시도",            C_OURS,  C_WHITE),
        ("★ 본 연구",  "최초 시도",            C_OURS,  C_WHITE),
    ],
    [  # row 1 — 리셀/대체자산 가격 연구
        ("─",          "해당 없음",            C_NA,    C_INK),
        ("✓ 기존",     "StockX 연구",          C_EXIST, C_WHITE),
        ("✓ 기존",     "카드 가격 연구",        C_EXIST, C_WHITE),
        ("✓ 기존",     "Dobrynskaya (2018)",   C_EXIST, C_WHITE),
    ],
    [  # row 2 — 미디어→가격 선행성 (bottom)
        ("✓ 기존",     "Tetlock 2007\nDa et al. 2011",  C_EXIST, C_WHITE),
        ("─",          "연구 공백",             C_GAP,   C_INK),
        ("─",          "연구 공백",             C_GAP,   C_INK),
        ("─",          "연구 공백",             C_GAP,   C_INK),
    ],
]

row_labels = ["멀티채널\n미디어 분석", "리셀/대체자산\n가격 연구", "미디어→가격\n선행성 연구"]
col_labels = ["전통 금융\n(주식)", "스니커즈", "트레이딩\n카드", "레고"]

nrows, ncols = 3, 4
cw, ch = 1.0, 1.0          # cell unit in data coords
L = 1.5                     # left margin for row labels

fig, ax = plt.subplots(figsize=(13, 5.6))
ax.set_xlim(-L, ncols * cw + 0.15)
ax.set_ylim(-0.5, nrows * ch + 0.55)
ax.axis("off")
fig.patch.set_facecolor("white")

# ── Draw cells ──────────────────────────────────────────────────────
for ri in range(nrows):
    for ci in range(ncols):
        main, sub, bg, tc = cells[ri][ci]
        # flip row so row 0 (멀티채널) is at top
        data_y = (nrows - 1 - ri) * ch
        data_x = ci * cw

        # Cell background
        rect = FancyBboxPatch(
            (data_x + 0.03, data_y + 0.03), cw - 0.06, ch - 0.06,
            boxstyle="round,pad=0.03",
            facecolor=bg, edgecolor=C_WHITE, linewidth=2.5,
        )
        ax.add_patch(rect)

        # Main text
        ax.text(data_x + cw / 2, data_y + ch * 0.62, main,
                ha="center", va="center", fontsize=12.5,
                fontweight="bold", color=tc)
        # Sub text
        ax.text(data_x + cw / 2, data_y + ch * 0.27, sub,
                ha="center", va="center", fontsize=8.5,
                color=tc, alpha=0.92, linespacing=1.4)

# ── Column headers ───────────────────────────────────────────────────
for ci, label in enumerate(col_labels):
    ax.text(ci * cw + cw / 2, nrows * ch + 0.08, label,
            ha="center", va="bottom", fontsize=11,
            fontweight="bold", color=C_INK)

# ── Row labels ───────────────────────────────────────────────────────
for ri, label in enumerate(row_labels):
    data_y = (nrows - 1 - ri) * ch + ch / 2
    ax.text(-0.08, data_y, label,
            ha="right", va="center", fontsize=10,
            fontweight="bold", color=C_INK, linespacing=1.4)

# ── Highlight box around "본 연구" zone ──────────────────────────────
hl = FancyBboxPatch(
    (1 * cw + 0.01, (nrows - 1 - 0) * ch + 0.01),
    3 * cw - 0.02, ch - 0.02,
    boxstyle="round,pad=0.02",
    facecolor="none", edgecolor=C_OURS, linewidth=3.5,
)
ax.add_patch(hl)

# Label for highlight
ax.text(1 * cw + 3 * cw / 2, (nrows - 1 - 0) * ch + ch + 0.03,
        "← 본 연구 (리셀 × 멀티채널, 최초 시도)",
        ha="center", va="bottom", fontsize=9,
        color=C_OURS, fontweight="bold")

# ── Legend ───────────────────────────────────────────────────────────
legend_handles = [
    mpatches.Patch(facecolor=C_EXIST, label="기존 연구 있음"),
    mpatches.Patch(facecolor=C_GAP,   label="연구 공백 (없음)"),
    mpatches.Patch(facecolor=C_OURS,  label="★ 본 연구 (최초)"),
    mpatches.Patch(facecolor=C_NA,    label="해당 없음"),
]
ax.legend(handles=legend_handles, loc="lower left", fontsize=9.5,
          framealpha=0.95, edgecolor="#cccccc", ncol=4)

plt.tight_layout()
out = os.path.join(FIG_DIR, "fig_gap_matrix.png")
plt.savefig(out, dpi=150, bbox_inches="tight", facecolor="white")
print("Saved:", out)

# ======================================================================
# fig_ablation2.png — Ablation 핵심 4모형 비교 (상대 스케일)
# ======================================================================
# 4개 모형만 추출: Model A / 가격래그만(Baseline) / 채널만 / A-dropGranger
# 왼쪽 패널: 전체 스케일 (Channels-only의 극적 하락 강조)
# 오른쪽 패널: 상대 스케일 (AUC - Model A, 미세 차이 부각)
# ──────────────────────────────────────────────────────────────────────
data = {
    "sneakers": {"A": 0.8413, "Baseline": 0.8366, "Channels": 0.5047, "dropGranger": 0.8472},
    "cards":    {"A": 0.8587, "Baseline": 0.8631, "Channels": 0.4736, "dropGranger": 0.8671},
    "lego":     {"A": 0.9601, "Baseline": 0.9628, "Channels": 0.5318, "dropGranger": 0.9601},
}
assets    = ["Sneakers", "Cards", "LEGO"]
keys      = ["A", "Baseline", "Channels", "dropGranger"]
colors    = ["#185FA5", "#5B9BD5", "#AAAAAA", "#C00000"]
labels    = ["Model A\n(전채널)", "가격래그만\n(채널 없음)", "채널만\n(래그 없음)", "A-dropGranger\n(Granger채널 제거)"]
hatches   = ["", "", "xx", ""]

fig2, (ax_l, ax_r) = plt.subplots(1, 2, figsize=(14, 5.5),
                                   gridspec_kw={"width_ratios": [1.1, 1]})
fig2.patch.set_facecolor("white")
plt.rcParams["font.family"] = ["Malgun Gothic", "sans-serif"]

n_assets = len(assets)
n_models = len(keys)
bw = 0.18
x = np.arange(n_assets)

# ── 왼쪽: 전체 스케일 ──────────────────────────────────────────────
for mi, (key, col, lbl, hatch) in enumerate(zip(keys, colors, labels, hatches)):
    vals = [data[a.lower()][key] for a in assets]
    xpos = x + (mi - (n_models - 1) / 2) * bw
    bars = ax_l.bar(xpos, vals, bw * 0.88, color=col,
                    hatch=hatch, edgecolor="white", linewidth=0.8, label=lbl)
    for bar, v in zip(bars, vals):
        yoff = -0.04 if v < 0.6 else 0.008
        va   = "top" if v < 0.6 else "bottom"
        fw   = "bold" if (key == "Channels") else "normal"
        ax_l.text(bar.get_x() + bar.get_width() / 2, v + yoff,
                  f"{v:.3f}", ha="center", va=va,
                  fontsize=8.5, fontweight=fw,
                  color="#C00000" if key == "Channels" else "#1A1A1A")

ax_l.axhline(0.5, color="#C00000", linestyle="--", linewidth=1.2, alpha=0.6)
ax_l.text(n_assets - 0.05, 0.502, "무작위 수준 (0.5)", fontsize=8.5,
          color="#C00000", ha="right", va="bottom")
ax_l.set_xticks(x); ax_l.set_xticklabels(assets, fontsize=12, fontweight="bold")
ax_l.set_ylim(0.40, 1.02)
ax_l.set_ylabel("AUC", fontsize=11); ax_l.set_title("① 가격 래그의 역할", fontsize=13, fontweight="bold")
ax_l.spines["top"].set_visible(False); ax_l.spines["right"].set_visible(False)
ax_l.legend(fontsize=8.5, loc="lower right", framealpha=0.9)
ax_l.set_facecolor("white")

# ── 오른쪽: 상대 스케일 (AUC - Model A) ──────────────────────────────
rel_keys   = ["Baseline", "dropGranger"]
rel_colors = ["#5B9BD5", "#C00000"]
rel_labels = ["가격래그만 − Model A", "A-dropGranger − Model A"]
bw2 = 0.28

for mi, (key, col, lbl) in enumerate(zip(rel_keys, rel_colors, rel_labels)):
    diffs = [data[a.lower()][key] - data[a.lower()]["A"] for a in assets]
    xpos  = x + (mi - 0.5) * bw2
    bars2 = ax_r.bar(xpos, diffs, bw2 * 0.88, color=col,
                     edgecolor="white", linewidth=0.8, label=lbl)
    for bar, d, a in zip(bars2, diffs, assets):
        yoff = 0.0008 if d >= 0 else -0.0008
        va2  = "bottom" if d >= 0 else "top"
        is_sig = (key == "dropGranger" and a == "Cards")
        ax_r.text(bar.get_x() + bar.get_width() / 2,
                  d + yoff,
                  f"{d:+.4f}" + (" ★p=.004" if is_sig else ""),
                  ha="center", va=va2, fontsize=8.5,
                  fontweight="bold" if is_sig else "normal",
                  color="#C00000" if is_sig else "#1A1A1A")

ax_r.axhline(0, color="#1A1A1A", linewidth=1.0)
ax_r.set_xticks(x); ax_r.set_xticklabels(assets, fontsize=12, fontweight="bold")
ax_r.set_ylabel("AUC 차이 (vs Model A)", fontsize=11)
ax_r.set_title("② Granger 채널 제거 시 변화", fontsize=13, fontweight="bold")
ax_r.spines["top"].set_visible(False); ax_r.spines["right"].set_visible(False)
ax_r.legend(fontsize=9, loc="upper left", framealpha=0.9)
ax_r.set_facecolor("white")

# Cards dropGranger 강조 화살표
diff_cards = data["cards"]["dropGranger"] - data["cards"]["A"]
xpos_cards = 1 + 0.5 * bw2
ax_r.annotate("Granger 채널 제거\n→ AUC 향상 (유의)",
              xy=(xpos_cards, diff_cards),
              xytext=(xpos_cards + 0.45, diff_cards + 0.003),
              fontsize=8.5, color="#C00000", fontweight="bold",
              arrowprops=dict(arrowstyle="->", color="#C00000", lw=1.5))

plt.tight_layout(pad=2.0)
out2 = os.path.join(FIG_DIR, "fig_ablation2.png")
plt.savefig(out2, dpi=150, bbox_inches="tight", facecolor="white")
print("Saved:", out2)
