import json, base64, os

os.makedirs('results/figures', exist_ok=True)

with open('biz_anl_(2).ipynb', encoding='utf-8') as f:
    nb = json.load(f)

# 셀별 이미지 레이블 정의
cell_labels = {
    4:  'clf_shap_asset',    # asset별 분류 SHAP
    6:  'clf_shap_item',     # item별 분류 SHAP
    9:  'reg_rmse_bar',      # RMSE 막대그래프
    10: 'reg_shap_asset',    # asset별 회귀 SHAP
    11: 'reg_shap_item',     # item별 회귀 SHAP
}

saved = []
for cell_idx, label in cell_labels.items():
    cell = nb['cells'][cell_idx]
    img_count = 0
    for out in cell.get('outputs', []):
        png = out.get('data', {}).get('image/png')
        if not png:
            continue
        fname = f'results/figures/{label}_{img_count:02d}.png'
        with open(fname, 'wb') as f:
            f.write(base64.b64decode(png))
        saved.append(fname)
        img_count += 1

print(f"저장 완료: {len(saved)}개")
for s in saved:
    print(' ', s)
