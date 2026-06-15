#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RQ1 Granger 검정 — 자산별(스니커즈/카드/레고) 그룹 막대그래프
y축 = -log10(p) (라벨에는 "유의성 강도"로 표기) → lag와 무관하게 단일 기준선(p=0.05)
하나로 색상(유의/비유의)이 그대로 설명됨 (F값은 lag마다 임계값이 달라 색상과
높이가 안 맞아 보이는 문제 회피). F·lag 값은 우측 표에서 별도 확인.
Usage : python scripts/make_fig_granger_by_asset.py
Output: results/figures/granger_by_asset.png
"""
import os
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "results", "figures", "granger_by_asset.png")

for f in ["Malgun Gothic", "맑은 고딕"]:
    if any(f.lower() in fn.name.lower() for fn in fm.fontManager.ttflist):
        plt.rcParams["font.family"] = f
        break
plt.rcParams["axes.unicode_minus"] = False

GREEN = "#1E7E34"
GRAY = "#AAB1BC"

# (asset, channel, lag, p, verdict) — 레고 CH1(p=0.054)도 단순화를 위해 비유의로 통일
data = [
    ("스니커즈", "CH1", 2, 0.041, "SIG"),
    ("스니커즈", "CH2", 1, 0.220, "NOT"),
    ("스니커즈", "CH3", 1, 0.0007, "SIG"),
    ("스니커즈", "CH4", 1, 0.035, "SIG"),
    ("스니커즈", "CH5", 1, 0.764, "NOT"),
    ("카드", "CH1", 1, 0.006, "SIG"),
    ("카드", "CH2", 1, 0.398, "NOT"),
    ("카드", "CH3", 1, 0.242, "NOT"),
    ("카드", "CH4", 1, 0.622, "NOT"),
    ("카드", "CH5", 1, 0.438, "NOT"),
    ("레고", "CH1", 1, 0.054, "NOT"),
    ("레고", "CH2", 1, 0.576, "NOT"),
    ("레고", "CH3", 1, 0.530, "NOT"),
    ("레고", "CH4", 1, 0.183, "NOT"),
    ("레고", "CH5", 1, 0.680, "NOT"),
]
COLOR = {"SIG": GREEN, "NOT": GRAY}

fig, ax = plt.subplots(figsize=(7.3, 4.6), dpi=200)

assets = ["스니커즈", "카드", "레고"]
group_gap = 1.2
bar_w = 0.8
xs, heights, colors, labels, ps = [], [], [], [], []
xticks, xticklabels = [], []
show_p = []
x0 = 0
for a in assets:
    rows = [d for d in data if d[0] == a]
    start = x0
    for (asset, ch, lag, p, verdict) in rows:
        xs.append(x0)
        heights.append(-math.log10(p))
        colors.append(COLOR[verdict])
        labels.append(ch)
        ps.append(p)
        # p값 라벨: 유의(SIG) 막대 + 레고 CH1(경계, p=0.054)만 표시
        show_p.append(verdict == "SIG" or (asset == "레고" and ch == "CH1"))
        x0 += 1
    xticks.append((start + x0 - 1) / 2)
    xticklabels.append(a)
    x0 += group_gap

bars = ax.bar(xs, heights, width=bar_w, color=colors, edgecolor="white")

for i, (x, h, lab, p, sp) in enumerate(zip(xs, heights, labels, ps, show_p)):
    if sp:
        p_txt = "p<0.001" if p < 0.001 else f"p={p:.3f}"
        off = 0.06 if (i % 2 == 0) else 0.24   # 인접 막대 라벨 겹침 방지(높이 비슷한 경우 교대 배치)
        ax.text(x, h + off, p_txt, ha="center", va="bottom", fontsize=8.5, fontweight="bold")
    ax.text(x, -0.12, lab, ha="center", va="top", fontsize=9, color="#444444")

# 단일 기준선 (lag와 무관) — 색상과 높이가 그대로 일치
LOG05 = -math.log10(0.05)   # ≈1.301
ax.axhline(LOG05, color="#2E6DA4", linestyle="--", linewidth=1.2)
ax.axhline(0, color="#888888", linewidth=0.8)

ax.set_xticks([])
for xt, lab in zip(xticks, xticklabels):
    ax.text(xt, -0.78, lab, ha="center", va="top", fontsize=13, fontweight="bold")

ax.set_ylabel("유의성 강도  (위로 갈수록 더 유의함)", fontsize=11)
ax.set_ylim(-1.0, 3.8)
ax.set_xlim(-1, x0 - group_gap)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.spines["bottom"].set_visible(False)
ax.spines["left"].set_bounds(0, 3.8)
ax.set_yticks([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5])
ax.tick_params(axis="x", length=0)

from matplotlib.lines import Line2D
handles = [
    plt.Rectangle((0, 0), 1, 1, color=GREEN, label="유의 (p<0.05)"),
    plt.Rectangle((0, 0), 1, 1, color=GRAY, label="비유의 (p≥0.05)"),
    Line2D([0], [0], color="#2E6DA4", linestyle="--", linewidth=1.2, label="p = 0.05 기준선"),
]
ax.legend(handles=handles, loc="upper right", fontsize=9, frameon=False)

plt.subplots_adjust(top=0.97, left=0.10, right=0.98)
plt.savefig(OUT, bbox_inches="tight", pad_inches=0.1)
print(f"[OK] {OUT}")
