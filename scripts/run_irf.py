"""
Impulse Response Function (IRF) for the two Granger-significant
asset–channel pairs (BH-survived): sneakers↔CH3, cards↔CH1.

Output:
    results/irf_results.csv   point estimates with bootstrap 95% CI
    results/figures/irf_sneakers_ch3.png
    results/figures/irf_cards_ch1.png
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from statsmodels.tsa.api import VAR

mpl.rcParams["font.family"] = "Malgun Gothic"
mpl.rcParams["axes.unicode_minus"] = False

ROOT = Path(__file__).resolve().parents[1]
ASSET_SERIES = ROOT / "data" / "processed" / "asset_series.csv"
STATIONARITY = ROOT / "results" / "stationarity.csv"
GRANGER = ROOT / "results" / "granger_results.csv"
OUT_CSV = ROOT / "results" / "irf_results.csv"
FIG_DIR = ROOT / "results" / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

HORIZON = 12
BOOT_REPS = 1000
SEED = 42

PAIRS = [
    ("sneakers", "score_ch3", "뉴스 감성 (CH3)"),
    ("cards",    "score_ch1", "Google Trends (CH1)"),
]


def _prepared_series(asset: str, channel: str) -> tuple[pd.DataFrame, dict]:
    """Return (var_input_df, info) — apply 1st-differencing per stationarity table."""
    series = pd.read_csv(ASSET_SERIES)
    series = series[series["asset_type"] == asset].sort_values("year_month").reset_index(drop=True)

    stat = pd.read_csv(STATIONARITY)
    stat = stat[stat["asset_type"] == asset].set_index("series")

    def _diff_if_needed(col: str) -> tuple[pd.Series, bool]:
        diffed = bool(stat.loc[col, "differenced"])
        return (series[col].diff() if diffed else series[col]), diffed

    price, price_diffed = _diff_if_needed("mean_price")
    chan,  chan_diffed  = _diff_if_needed(channel)

    df = pd.concat([price.rename("price"), chan.rename("channel")], axis=1).dropna().reset_index(drop=True)
    info = {
        "asset": asset,
        "channel": channel,
        "price_differenced": price_diffed,
        "channel_differenced": chan_diffed,
        "n_obs": len(df),
    }
    return df, info


def _granger_lag(asset: str, channel: str) -> int:
    g = pd.read_csv(GRANGER)
    row = g[(g["asset_type"] == asset) & (g["channel"] == channel)]
    if row.empty:
        return 1
    return int(row.iloc[0]["lag"])


def _residual_bootstrap_irf(data: np.ndarray, lag: int, horizon: int,
                            reps: int, seed: int):
    """Residual bootstrap for orthogonalised IRF 95% CI.

    statsmodels' errband_mc returns lower==upper in our environment, so we
    re-implement the standard residual bootstrap of Lütkepohl (2005, §3.5).
    """
    fit = VAR(data).fit(maxlags=lag, ic=None, trend="c")
    point = fit.irf(horizon).orth_irfs  # (h+1, k, k)

    coefs = fit.coefs            # (lag, k, k)
    intercept = fit.intercept    # (k,)
    resid = fit.resid            # (n-lag, k)
    n, k = data.shape

    rng = np.random.default_rng(seed)
    collected: list[np.ndarray] = []

    for _ in range(reps):
        idx = rng.integers(0, len(resid), size=len(resid))
        boot_resid = resid[idx]
        sim = np.zeros_like(data)
        sim[:lag] = data[:lag]
        for t in range(lag, n):
            val = intercept.copy()
            for p in range(lag):
                val = val + coefs[p] @ sim[t - 1 - p]
            val = val + boot_resid[t - lag]
            sim[t] = val
        try:
            boot_fit = VAR(sim).fit(maxlags=lag, ic=None, trend="c")
            collected.append(boot_fit.irf(horizon).orth_irfs)
        except Exception:
            continue

    boot = np.asarray(collected)
    lower = np.percentile(boot, 2.5, axis=0)
    upper = np.percentile(boot, 97.5, axis=0)
    return point, lower, upper


def _plot(point, lower, upper, response_idx, impulse_idx,
          title: str, ylabel: str, out_path: Path) -> None:
    h = point.shape[0]
    x = np.arange(h)
    y = point[:, response_idx, impulse_idx]
    lo = lower[:, response_idx, impulse_idx]
    hi = upper[:, response_idx, impulse_idx]

    fig, ax = plt.subplots(figsize=(7, 4))
    ax.axhline(0.0, color="grey", lw=0.8, ls="--")
    ax.fill_between(x, lo, hi, alpha=0.25, label="95% bootstrap CI")
    ax.plot(x, y, lw=2.0, color="C0", label="Orth. IRF")
    ax.set_xlabel("월 후 (horizon)")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.set_xticks(np.arange(0, h, 2))
    ax.legend(loc="best", frameon=False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)


def main() -> None:
    rows = []

    for asset, channel, channel_label in PAIRS:
        df, info = _prepared_series(asset, channel)
        lag = _granger_lag(asset, channel)
        data = df[["price", "channel"]].values
        point, lower, upper = _residual_bootstrap_irf(data, lag, HORIZON, BOOT_REPS, SEED)

        # response = price (index 0), impulse = channel (index 1)
        for h in range(point.shape[0]):
            rows.append({
                "asset": asset,
                "channel": channel,
                "horizon": h,
                "irf_price_to_channel": point[h, 0, 1],
                "ci_lo": lower[h, 0, 1],
                "ci_hi": upper[h, 0, 1],
                "lag": lag,
                "price_differenced": info["price_differenced"],
                "channel_differenced": info["channel_differenced"],
                "n_obs": info["n_obs"],
            })

        # plot price response to a one-SD orthogonalised shock in channel
        price_label = "Δ가격" if info["price_differenced"] else "가격"
        chan_in_diff = " (1차 차분)" if info["channel_differenced"] else ""
        title = f"{asset.capitalize()}: {channel_label}{chan_in_diff} 충격 → {price_label} 반응"
        out_png = FIG_DIR / f"irf_{asset}_{channel.split('_')[1]}.png"
        _plot(point, lower, upper, 0, 1, title, price_label, out_png)
        print(f"saved {out_png.name} | n={info['n_obs']} lag={lag}")

    pd.DataFrame(rows).to_csv(OUT_CSV, index=False)
    print(f"wrote {OUT_CSV}")


if __name__ == "__main__":
    sys.exit(main())
