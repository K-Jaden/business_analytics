"""
Ablation study figure for draft.tex §4.2
논문용 그림: 9개 모델 × 3자산 AUC 비교 (grouped bar chart)
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# --- 한글 폰트 ---
from matplotlib import font_manager
for fname in font_manager.findSystemFonts():
    if 'malgun' in fname.lower() or 'Malgun' in fname:
        font_manager.fontManager.addfont(fname)
        plt.rcParams['font.family'] = 'Malgun Gothic'
        break

plt.rcParams.update({
    'font.size': 8,
    'axes.titlesize': 9,
    'axes.labelsize': 8,
    'xtick.labelsize': 7.5,
    'ytick.labelsize': 8,
    'figure.dpi': 300,
})

df = pd.read_csv('results/ablation_results.csv')

MODEL_ORDER = ['A', 'Baseline', 'Channels-only',
               'CH1-only', 'CH2-only', 'CH3-only', 'CH4-only', 'CH5-only',
               'A-dropGranger']
MODEL_LABELS = ['Model A\n(기준)', 'Baseline\n(가격래그만)', 'Channels-only\n(채널만)',
                'CH1-only', 'CH2-only', 'CH3-only', 'CH4-only', 'CH5-only',
                'A-drop\nGranger']
ASSETS = ['sneakers', 'cards', 'lego']
ASSET_LABELS = ['Sneakers', 'Cards', 'Lego']

# 색상 설정
COLOR_A      = '#2166ac'   # 파랑 - Model A
COLOR_DROP   = '#d6604d'   # 빨강 - A-dropGranger (핵심 비교)
COLOR_CONLY  = '#969696'   # 회색 - Channels-only (실패 케이스)
COLORS_ASSET = ['#4393c3', '#f4a582', '#92c47a']  # sneakers/cards/lego

fig, axes = plt.subplots(1, 3, figsize=(7.0, 2.9), sharey=False)
fig.subplots_adjust(wspace=0.42, left=0.07, right=0.97, top=0.88, bottom=0.22)

for ax, asset, asset_label, col in zip(axes, ASSETS, ASSET_LABELS, COLORS_ASSET):
    sub = df[df['asset_type'] == asset].set_index('model')
    aucs = [sub.loc[m, 'auc'] for m in MODEL_ORDER]
    verdicts = [sub.loc[m, 'vs_A'] for m in MODEL_ORDER]

    bar_colors = []
    for m, v in zip(MODEL_ORDER, verdicts):
        if m == 'A':
            bar_colors.append(COLOR_A)
        elif m == 'A-dropGranger':
            bar_colors.append(COLOR_DROP)
        elif m == 'Channels-only':
            bar_colors.append(COLOR_CONLY)
        else:
            bar_colors.append(col)

    x = np.arange(len(MODEL_ORDER))
    bars = ax.bar(x, aucs, color=bar_colors, width=0.65, edgecolor='white', linewidth=0.4)

    # Model A 참조선
    auc_A = sub.loc['A', 'auc']
    ax.axhline(auc_A, color=COLOR_A, linestyle='--', linewidth=0.8, alpha=0.6)

    # A-dropGranger에 DM 유의성 별표
    drop_idx = MODEL_ORDER.index('A-dropGranger')
    verdict_drop = verdicts[drop_idx]
    auc_drop = aucs[drop_idx]
    if verdict_drop == 'ablation better':
        ax.text(drop_idx, auc_drop + 0.003, '★', ha='center', va='bottom',
                fontsize=7, color=COLOR_DROP, fontweight='bold')

    # Channels-only에 ✕ 표시
    conly_idx = MODEL_ORDER.index('Channels-only')
    ax.text(conly_idx, aucs[conly_idx] + 0.003, 'X', ha='center', va='bottom',
            fontsize=7, color='#555555', fontweight='bold')

    # Y축 범위
    ymin = min(aucs) - 0.03
    ymax = max(aucs) + 0.04
    ax.set_ylim(ymin, ymax)

    ax.set_title(asset_label, fontweight='bold', pad=4)
    ax.set_xticks(x)
    ax.set_xticklabels(MODEL_LABELS, rotation=55, ha='right', fontsize=6.8)
    ax.set_ylabel('AUC' if asset == 'sneakers' else '', labelpad=2)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:.3f}'))
    ax.tick_params(axis='y', labelsize=7)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=0.3, linewidth=0.5)

# 범례
legend_handles = [
    mpatches.Patch(color=COLOR_A, label='Model A (기준)'),
    mpatches.Patch(color=COLOR_DROP, label='A-dropGranger'),
    mpatches.Patch(color=COLOR_CONLY, label='Channels-only'),
    mpatches.Patch(color='#4393c3', label='CH단독 모델'),
]
fig.legend(handles=legend_handles, loc='lower center', ncol=4,
           fontsize=7, frameon=False, bbox_to_anchor=(0.52, -0.01))

out = 'results/figures/ablation_auc.png'
plt.savefig(out, dpi=300, bbox_inches='tight')
print(f"저장 완료: {out}")
plt.close()
