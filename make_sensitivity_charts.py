#!/usr/bin/env python3
"""读取《600887_毛利率WACC敏感性》工作簿，生成毛利率 × WACC 敏感性图。

输出：
  08 敏感性 毛利率与WACC（热力图，退出乘数法）
  09 敏感性 毛利率情景折线（每股价值 vs WACC，含现价参考线）
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_sens"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import openpyxl

XLSX = Path(__file__).resolve().parents[2] / "数据文件" / "敏感性分析" / "600887_毛利率WACC敏感性_20261026_Rev00.xlsx"
OUTDIR = Path(__file__).resolve().parents[2] / "可视化图表" / "600887_DCF估值图"
REV = "20261026_Rev00"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

NAVY, RED, GRAY = "#1F4E79", "#C00000", "#8C8C8C"
SERIES = ["#C00000", "#E67E22", "#404040", "#2E7D32", "#1F4E79"]
LABELS = ["毛利率 −2pct", "毛利率 −1pct", "基准", "毛利率 +1pct", "毛利率 +2pct"]

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["毛利率WACC敏感性"]

WACC = np.array([ws.cell(37, c).value for c in range(3, 10)], dtype=float) * 100
DELTA = np.array([ws.cell(r, 2).value for r in range(38, 43)], dtype=float) * 100
GRID = np.array([[ws.cell(r, c).value for c in range(3, 10)] for r in range(38, 43)], dtype=float)
BASE_PRICE = float(ws["C45"].value)
SPOT = float(ws["C47"].value)
GM = float(ws["C7"].value) * 100
W0 = float(ws["C15"].value) * 100


def save(fig, name, subtitle):
    fig.text(0.02, 0.965, subtitle, fontsize=9, color="#666666", ha="left", va="top")
    out = OUTDIR / name
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存:", out.name)


def heat(ax, grid, xlab, ylab, fmt="{:,.2f}", cmap="YlGnBu"):
    im = ax.imshow(grid, cmap=cmap, aspect="auto", origin="lower")
    for i in range(grid.shape[0]):
        for j in range(grid.shape[1]):
            v = grid[i, j]
            shade = "white" if v > np.nanmean(grid) else "#333333"
            ax.text(j, i, fmt.format(v), ha="center", va="center",
                    fontsize=9.5, color=shade, fontweight="bold")
    ax.set_xticks(range(grid.shape[1]))
    ax.set_yticks(range(grid.shape[0]))
    ax.set_xlabel(xlab, fontsize=10)
    ax.set_ylabel(ylab, fontsize=10)
    for s in ax.spines.values():
        s.set_visible(False)
    return im


def fig08():
    fig, ax = plt.subplots(figsize=(8.8, 4.6), dpi=150)
    im = heat(ax, GRID, "WACC（%）", "预测期毛利率变动（pct）")
    ax.set_xticklabels([f"{v:.2f}" for v in WACC], fontsize=9.5)
    ax.set_yticklabels(["基准" if v == 0 else f"{v:+.0f}" for v in DELTA], fontsize=9.5)
    ib = int(np.argmin(np.abs(DELTA)))
    jb = int(np.argmin(np.abs(WACC - W0)))
    ax.add_patch(mpatches.Rectangle((jb - 0.5, ib - 0.5), 1, 1, fill=False,
                                    edgecolor=RED, linewidth=1.8, zorder=5))
    fig.colorbar(im, ax=ax, shrink=0.85, label="元/股")
    save(fig, f"600887_DCF估值图_08_敏感性_毛利率与WACC_{REV}.png",
         f"基准毛利率 {GM:.2f}%、WACC {W0:.2f}%，基准每股 {BASE_PRICE:.2f} 元（退出乘数法）")


def fig09():
    fig, ax = plt.subplots(figsize=(9.2, 5.0), dpi=150)
    for i in [0, 1, 3, 4]:
        ax.plot(WACC, GRID[i], marker="o", markersize=4.5, linewidth=1.6,
                color=SERIES[i], label=LABELS[i], zorder=3)
    ax.plot(WACC, GRID[2], marker="o", markersize=5.5, linewidth=2.6,
            color="#404040", label=LABELS[2], zorder=4)
    ax.axhline(SPOT, color=RED, linestyle=(0, (5, 3)), linewidth=1.3, zorder=2)
    ax.annotate(f"现价 {SPOT:.2f} 元", (WACC[0], SPOT), textcoords="offset points",
                xytext=(9, -14), fontsize=9, color=RED, fontweight="bold", va="top",
                bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "none", "alpha": 0.9})
    ax.axvline(W0, color=GRAY, linestyle=":", linewidth=1.0, zorder=1)
    idx = np.argmin(np.abs(WACC - W0))
    ax.annotate(f"{BASE_PRICE:.2f}", (WACC[idx], GRID[2][idx]), textcoords="offset points",
                xytext=(0, 11), ha="center", fontsize=9.5, color="#404040", fontweight="bold")
    ax.set_ylim(GRID.min() - 1.2, GRID.max() + 2.6)
    ax.set_yticks(np.arange(22, 40, 2))
    ax.set_xlabel("WACC（%）", fontsize=10)
    ax.set_ylabel("每股价值（元）", fontsize=10)
    ax.set_xticks(WACC)
    ax.set_xticklabels([f"{v:.2f}" for v in WACC], fontsize=9.5)
    ax.grid(axis="y", color="#E4E4E4", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")
    handles, labels = ax.get_legend_handles_labels()
    order = [4, 3, 2, 1, 0]
    ax.legend([handles[i] for i in order], [labels[i] for i in order],
              fontsize=9, frameon=False, ncol=5, loc="upper center",
              handlelength=1.6, columnspacing=1.4)
    save(fig, f"600887_DCF估值图_09_敏感性_毛利率情景折线_{REV}.png",
         f"退出乘数法，终值 = 2030E EBITDA × {float(ws['C14'].value):g}×；虚线为现价、点线为基准 WACC")


if __name__ == "__main__":
    fig08()
    fig09()
    print("\n网格范围 %.2f ~ %.2f 元 | 基准 %.2f 元" % (GRID.min(), GRID.max(), BASE_PRICE))
