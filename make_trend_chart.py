#!/usr/bin/env python3
"""从基础趋势表读数据，生成伊利股份基础趋势图（四面板，无双轴重叠）。"""

from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl

XLSX = Path(__file__).resolve().parents[2] / "数据文件" / "基础趋势表" / "600887_伊利股份_基础趋势表_20261011_Rev01.xlsx"
OUT = Path(__file__).resolve().parents[2] / "可视化图表" / "600887_基础趋势图" / "600887_基础趋势图_20261020_Rev01.png"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

LABELS = ["2021A", "2022A", "2023A", "2024A", "2025A", "2026H1"]
COLS = [3, 4, 5, 6, 8, 9]          # 利润表/现金流量表列：C D E F H I（跳过 G=2025H1）
YEAR, HALF = "#2E5C8A", "#9DC3E6"
C_REVYOY, C_NIYOY, C_GM, C_NM, C_OCF = "#C00000", "#1F6FB2", "#2E7D32", "#8E44AD", "#E67E22"


def read_data():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    lp, cf = wb["利润表"], wb["现金流量表"]
    g = lambda ws, r, c: ws.cell(r, c).value
    rev = np.array([g(lp, 3, c) for c in COLS]) / 1e8
    cogs = np.array([g(lp, 10, c) for c in COLS]) / 1e8
    ni = np.array([g(lp, 51, c) for c in COLS]) / 1e8
    ocf = np.array([g(cf, 27, c) for c in COLS]) / 1e8
    rev_base = [g(lp, 3, 2) / 1e8] + list(rev[:-1])      # 2021 基数取 2020 年报
    ni_base = [g(lp, 51, 2) / 1e8] + list(ni[:-1])
    rev_base[5] = g(lp, 3, 7) / 1e8                      # 2026H1 对比 2025H1
    ni_base[5] = g(lp, 51, 7) / 1e8
    return dict(rev=rev, ni=ni,
                gm=(rev - cogs) / rev * 100, nm=ni / rev * 100, ocf_ni=ocf / ni,
                rev_yoy=(rev / np.array(rev_base) - 1) * 100,
                ni_yoy=(ni / np.array(ni_base) - 1) * 100)


def style_ax(ax):
    ax.grid(axis="y", color="#DDDDDD", linewidth=0.7)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")


def mark_half_year(ax, axmax):
    """标注 2026H1 为半年数（仅用于绝对额面板）"""
    ax.axvspan(4.5, 5.62, color="#FFF3E0", zorder=0)
    ax.axvline(4.5, color="#C00000", linestyle=(0, (4, 3)), linewidth=1.2, zorder=4)
    ax.text(5.05, axmax * 0.965, "半年数", ha="center", va="top",
            fontsize=9.5, color="#C00000")


def bar_panel(ax, vals, title, ylabel, imax):
    style_ax(ax)
    x = np.arange(6)
    cols = [YEAR] * 5 + [HALF]
    bars = ax.bar(x, vals, width=0.55, color=cols, zorder=3,
                  edgecolor=["none"] * 5 + [YEAR])
    bars[5].set_hatch("//")
    labels = ax.bar_label(bars, labels=[f"{v:,.0f}" for v in vals], label_type="center",
                          fontsize=10, color="white", fontweight="bold")
    labels[5].set_color("#1B3A5C")           # 2026H1 柱子浅色，白字看不清，改深色
    top = max(vals) * imax
    ax.set_ylim(0, top)
    ax.set_ylabel(ylabel, fontsize=10)
    ax.set_title(title, fontsize=12, fontweight="bold", loc="left", pad=10)
    mark_half_year(ax, top)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS, fontsize=10)


def growth_panel(ax, rev_yoy, ni_yoy):
    style_ax(ax)
    x = np.arange(6)
    ax.axhline(0, color="#999999", linewidth=0.9, linestyle=":")
    ax.plot(x, rev_yoy, color=C_REVYOY, marker="o", linewidth=2, markersize=6, label="营收增速")
    ax.plot(x, ni_yoy, color=C_NIYOY, marker="s", linewidth=2, markersize=6, label="归母净利增速")
    # 每个标签放在背离另一条线的方向，避免互压
    for i in x:
        rev_above = rev_yoy[i] > ni_yoy[i]
        ax.annotate(f"{rev_yoy[i]:+.1f}%", (i, rev_yoy[i]), textcoords="offset points",
                    xytext=(0, 12 if rev_above else -20), ha="center",
                    fontsize=9, color=C_REVYOY)
        ax.annotate(f"{ni_yoy[i]:+.1f}%", (i, ni_yoy[i]), textcoords="offset points",
                    xytext=(0, -20 if rev_above else 12), ha="center",
                    fontsize=9, color=C_NIYOY)
    ax.set_ylim(-36, 56)
    ax.set_ylabel("同比增速（%）", fontsize=10)
    ax.set_title("营收增速与归母净利增速（均为同比）", fontsize=12, fontweight="bold",
                 loc="left", pad=10)
    ax.legend(loc="lower left", fontsize=9.5, frameon=False, ncol=2)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS, fontsize=10)


def ratio_panel(ax, gm, nm, ocf_ni):
    style_ax(ax)
    x = np.arange(6)
    l1, = ax.plot(x, gm, color=C_GM, marker="s", linewidth=2.2, markersize=6, label="毛利率")
    l2, = ax.plot(x, nm, color=C_NM, marker="^", linewidth=2.2, markersize=6, label="归母净利率")
    ax.set_ylim(0, 46)
    ax.set_ylabel("毛利率 / 归母净利率（%）", fontsize=10)
    for i in x:
        ax.annotate(f"{gm[i]:.2f}%", (i, gm[i]), textcoords="offset points", xytext=(0, 10),
                    ha="center", fontsize=9, color=C_GM)
        ax.annotate(f"{nm[i]:.2f}%", (i, nm[i]), textcoords="offset points", xytext=(0, -18),
                    ha="center", fontsize=9, color=C_NM)
    ax2 = ax.twinx()
    bars = ax2.bar(x, ocf_ni, width=0.30, color=C_OCF, alpha=0.25, zorder=1)
    ax2.set_ylim(0, 4.6)
    ax2.set_ylabel("经营现金流 / 归母净利润（倍）", fontsize=10, color=C_OCF)
    ax2.tick_params(axis="y", colors=C_OCF)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_color(C_OCF)
    for i in x:
        ax2.annotate(f"{ocf_ni[i]:.2f}", (i, ocf_ni[i]), textcoords="offset points",
                     xytext=(0, 5), ha="center", fontsize=9, color="#B35C00")
    ax.set_title("毛利率 / 归母净利率 / 经营现金流对归母净利润的覆盖", fontsize=12,
                 fontweight="bold", loc="left", pad=26)
    ax.legend([l1, l2, bars], ["毛利率（左轴）", "归母净利率（左轴）", "经营现金流/归母净利润（右轴）"],
              loc="lower left", bbox_to_anchor=(0, 1.005), ncol=3, fontsize=9.5, frameon=False)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS, fontsize=10)


def main() -> None:
    d = read_data()
    fig, axes = plt.subplots(4, 1, figsize=(11, 15.5), dpi=150)
    fig.subplots_adjust(top=0.905, bottom=0.035, left=0.085, right=0.915, hspace=0.42)

    bar_panel(axes[0], d["rev"], "营业收入", "营业收入（亿元）", 1.20)
    bar_panel(axes[1], d["ni"], "归母净利润", "归母净利润（亿元）", 1.20)
    growth_panel(axes[2], d["rev_yoy"], d["ni_yoy"])
    ratio_panel(axes[3], d["gm"], d["nm"], d["ocf_ni"])

    fig.text(0.085, 0.968,
             "金额单位：亿元　|　"
             "上方两图的 2026H1 为半年累计，绝对额不可与年度比较；下方两图为增速与比率，跨期可比",
             fontsize=9.5, color="#555555", ha="left")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUT, bbox_inches="tight", facecolor="white")
    print("已保存:", OUT)
    for k, v in d.items():
        print(f"  {k:8}", np.round(v, 3).tolist())


if __name__ == "__main__":
    main()
