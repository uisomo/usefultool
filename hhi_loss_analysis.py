#!/usr/bin/env python3
"""
HHI vs Portfolio Loss Rate Analysis via Monte Carlo Simulation.

Demonstrates how portfolio concentration (measured by HHI) affects
loss distribution under purely idiosyncratic risk (zero correlation).

Conditions:
  - 10 obligors
  - PD = 0.03% (AA rating)
  - LGD = 100%
  - Correlation = 0
  - 1,000,000 simulation trials per HHI point
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

# ── Japanese font setup ──────────────────────────────────────────────────
font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
font_manager.fontManager.addfont(font_path)
plt.rcParams["font.family"] = "IPAGothic"

# ── Configuration ────────────────────────────────────────────────────────
N_OBLIGORS = 10
PD = 0.0003           # 0.03% (AA rating)
LGD = 1.0             # 100%
N_SIMS = 1_000_000
N_HHI_POINTS = 50


def generate_weights(alpha, n=N_OBLIGORS):
    """Generate weight vector with concentration parameter alpha in [0, 1].

    alpha=0 -> equal weights (HHI = 1/N)
    alpha=1 -> full concentration on obligor 0 (HHI = 1.0)
    """
    equal = np.ones(n) / n
    concentrated = np.zeros(n)
    concentrated[0] = 1.0
    return (1 - alpha) * equal + alpha * concentrated


def compute_hhi(weights):
    return np.sum(weights ** 2)


def run_simulation(weights, n_sims=N_SIMS):
    """Run Monte Carlo and return portfolio loss rates array."""
    n = len(weights)
    defaults = (np.random.random((n_sims, n)) < PD).astype(np.float64)
    return defaults @ (weights * LGD)


def compute_stats(loss_rates):
    return {
        "expected_loss": np.mean(loss_rates),
        "std_dev": np.std(loss_rates),
        "var_999": np.percentile(loss_rates, 99.9),
        "var_99": np.percentile(loss_rates, 99.0),
        "max_loss": np.max(loss_rates),
    }


def theoretical_std(hhi_values):
    return np.sqrt(hhi_values * PD * (1 - PD))


def create_chart(hhi_values, stats_list):
    fig, ax = plt.subplots(figsize=(12, 8))

    hhis = np.array(hhi_values)
    std_devs = np.array([s["std_dev"] for s in stats_list])
    var_999 = np.array([s["var_999"] for s in stats_list])
    max_losses = np.array([s["max_loss"] for s in stats_list])
    expected = np.array([s["expected_loss"] for s in stats_list])

    # Simulated metrics
    ax.plot(hhis, expected * 100, "o-", label="期待損失率 (E[L])",
            markersize=3, color="#2196F3")
    ax.plot(hhis, std_devs * 100, "s-", label="標準偏差 (σ)",
            markersize=3, color="#FF9800")
    ax.plot(hhis, var_999 * 100, "^-", label="VaR 99.9%",
            markersize=4, color="#E91E63")
    ax.plot(hhis, max_losses * 100, "D-", label="最大損失率",
            markersize=3, color="#9C27B0", alpha=0.6)

    # Theoretical curves
    hhi_fine = np.linspace(hhis.min(), hhis.max(), 200)
    ax.plot(hhi_fine, theoretical_std(hhi_fine) * 100, "--",
            color="red", label="理論値 σ = √(HHI·PD·(1-PD))", linewidth=2)
    ax.axhline(y=PD * 100, color="gray", linestyle=":",
               label=f"理論値 E[L] = {PD*100:.3f}%", linewidth=1.5)

    ax.set_xlabel("HHI (ハーフィンダール・ハーシュマン指数)", fontsize=13)
    ax.set_ylabel("損失率 (%)", fontsize=13)
    ax.set_title(
        "HHI（ポートフォリオ集中度）と損失率の関係\n"
        f"債務者数={N_OBLIGORS}, PD={PD*100:.3f}% (AA格), "
        f"LGD={LGD*100:.0f}%, 相関=0, "
        f"シミュレーション={N_SIMS:,}回",
        fontsize=13,
    )
    ax.legend(fontsize=10, loc="upper left")
    ax.grid(True, alpha=0.3)
    ax.set_xlim(0.05, 1.05)

    plt.tight_layout()
    output_path = "hhi_loss_analysis.png"
    plt.savefig(output_path, dpi=150, bbox_inches="tight")
    print(f"Saved: {output_path}")


def main():
    np.random.seed(42)

    alphas = np.linspace(0, 0.995, N_HHI_POINTS)
    hhi_values = []
    stats_list = []

    print(f"Running Monte Carlo: {N_HHI_POINTS} HHI points x "
          f"{N_SIMS:,} simulations each")
    print(f"  PD={PD*100:.3f}%, LGD={LGD*100:.0f}%, "
          f"N_obligors={N_OBLIGORS}, correlation=0\n")

    for i, alpha in enumerate(alphas):
        weights = generate_weights(alpha)
        hhi = compute_hhi(weights)
        hhi_values.append(hhi)

        loss_rates = run_simulation(weights)
        stats = compute_stats(loss_rates)
        stats_list.append(stats)

        if (i + 1) % 10 == 0 or i == 0:
            print(f"  [{i+1:2d}/{N_HHI_POINTS}] HHI={hhi:.4f}  "
                  f"E[L]={stats['expected_loss']*100:.4f}%  "
                  f"σ={stats['std_dev']*100:.4f}%  "
                  f"VaR99.9={stats['var_999']*100:.2f}%")

    print("\nGenerating chart...")
    create_chart(hhi_values, stats_list)
    print("Done!")


if __name__ == "__main__":
    main()
