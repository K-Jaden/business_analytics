import numpy as np
import pandas as pd
import xgboost as xgb
import shap

df = pd.read_csv('data/processed/xgboost_data.csv')
df = df.dropna(subset=['price_chg_lag1', 'price_chg_lag2', 'price_chg_lag3']).copy()

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

rows = []

for asset in ['sneakers', 'cards', 'lego']:
    subset = df[df['asset_type'] == asset].copy().sort_values('year_month').reset_index(drop=True)
    y = subset['price_direction']

    for model_name, features in [('Model A', model_a_features), ('Model B', model_b_features[asset])]:
        X = subset[features]
        clf = xgb.XGBClassifier(n_estimators=100, max_depth=3, learning_rate=0.05,
                                 random_state=42, eval_metric='auc', verbosity=0)
        clf.fit(X, y)

        explainer = shap.TreeExplainer(clf)
        shap_values = explainer.shap_values(X)
        mean_abs_shap = np.abs(shap_values).mean(axis=0)

        for feat, val in zip(features, mean_abs_shap):
            rows.append({
                'asset_type': asset,
                'model': model_name,
                'feature': feat,
                'mean_abs_shap': round(val, 6)
            })

result = pd.DataFrame(rows)
result.to_csv('results/shap_summary.csv', index=False)

print(result.to_string(index=False))
print("\nSaved: results/shap_summary.csv")
