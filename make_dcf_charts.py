#!/usr/bin/env python3
"""从 600887_DCF模型 Rev07 读数据，按类型生成估值图表（每张独立保存，研报风格）。

输出：
  01 营业收入与增速      02 归母净利润与利润率    03 自由现金流与资本开支
  04 FCF 构成            05 企业价值构成与两法结果
  06 敏感性 WACC×g       07 敏感性 WACC×退出乘数
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_dcf"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.ticker import FuncFormatter

XLSX = Path(__file__).resolve().parents[2] / "数据文件" / "DCF模型" / "600887_DCF模型_20261012_Rev07.xlsx"
OUTDIR = Path(__file__).resolve().parents[2] / "可视化图表" / "600887_DCF估值图"
REV = "20261020_Rev01"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

NAVY, RED, GREEN, ORANGE, GRAY = "#1F4E79", "#C00000", "#2E7D32", "#E67E22", "#8C8C8C"
HIST = "#1F4E79"
FORE = "#9DC3E6"
LABELS = ["2021A", "2022A", "2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
SPLIT = 4.5  # 2025A 与 2026E 之间

wb = openpyxl.load_workbook(XLSX, data_only=True)
RP, DCF = wb["报表预测"], wb["DCF"]
HCOLS, FCOLS = ["G", "H", "I", "J", "K"], ["L", "M", "N", "O", "P"]
ALL = HCOLS + FCOLS


def row(sheet, r, cols=ALL):
    return np.array([sheet["%s%d" % (c, r)].value for c in cols], dtype=float)


def style(ax, ylab="", xticks=True):
    ax.grid(axis="y", color="#E4E4E4", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")
    if xticks:
        ax.set_xticks(range(10))
        ax.set_xticklabels(LABELS, fontsize=9.5)
        ax.axvline(SPLIT, color=RED, linestyle=(0, (4, 3)), linewidth=1.0, zorder=1)
    if ylab:
        ax.set_ylabel(ylab, fontsize=10)


def save(fig, name, title, subtitle):
    fig.text(0.055, 0.955, subtitle, fontsize=9, color="#666666", ha="left", va="top")
    out = OUTDIR / name
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存:", out.name)


SRC_NOTE = ""


def fig01():
    rev = row(RP, 8)
    yoy = row(RP, 9) * 100
    fig, ax = plt.subplots(figsize=(10.5, 5.6), dpi=150)
    style(ax, "营业收入（亿元）")
    x = np.arange(10)
    bars = ax.bar(x, rev, width=0.6, color=[HIST] * 5 + [FORE] * 5, zorder=3)
    for i in range(5, 10):
        bars[i].set_edgecolor(NAVY)
    for i in x:
        ax.annotate(f"{rev[i]:,.0f}", (i, rev[i]), textcoords="offset points", xytext=(0, 5),
                    ha="center", fontsize=9, color=NAVY, fontweight="bold")
    ax.set_ylim(0, rev.max() * 1.22)
    ax2 = ax.twinx()
    ax2.plot(x[1:], yoy[1:], color=RED, marker="o", linewidth=2.0, markersize=5.5, zorder=5)
    for i in x[1:]:
        ax2.annotate(f"{yoy[i]:+.1f}%", (i, yoy[i]), textcoords="offset points",
                     xytext=(0, -16 if yoy[i] < 0 else 9), ha="center", fontsize=8.5, color=RED)
    ax2.axhline(0, color="#AAAAAA", linewidth=0.8, linestyle=":")
    ax2.set_ylim(-22, 26)
    ax2.set_ylabel("营业收入同比（%）", fontsize=10, color=RED)
    ax2.tick_params(axis="y", colors=RED)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_color(RED)
    ax.text(0.02, 0.94, "深色＝历史，浅色＝预测", transform=ax.transAxes, fontsize=9, color="#666666")
    save(fig, f"600887_DCF估值图_01_营业收入与增速_{REV}.png", "营业收入与增速",
         "柱＝营业收入（左轴）；线＝同比增速（右轴）")


def fig02():
    ni = row(RP, 34)
    gm = row(RP, 11) * 100
    nm = ni / row(RP, 8) * 100
    fig, ax = plt.subplots(figsize=(10.5, 5.6), dpi=150)
    style(ax, "归母净利润（亿元）")
    x = np.arange(10)
    bars = ax.bar(x, ni, width=0.6, color=[HIST] * 5 + [FORE] * 5, zorder=3)
    for i in range(5, 10):
        bars[i].set_edgecolor(NAVY)
    for i in x:
        ax.annotate(f"{ni[i]:,.1f}", (i, ni[i]), textcoords="offset points", xytext=(0, 5),
                    ha="center", fontsize=9, color=NAVY, fontweight="bold")
    ax.set_ylim(0, ni.max() * 1.28)
    ax2 = ax.twinx()
    for y, c, mk, tag in ((gm, GREEN, "s", "毛利率"), (nm, ORANGE, "^", "归母净利率")):
        ax2.plot(x, y, color=c, marker=mk, linewidth=2.0, markersize=5.5, label=tag, zorder=5)
    for i in (0, 4, 5, 9):
        ax2.annotate(f"{gm[i]:.2f}%", (i, gm[i]), textcoords="offset points", xytext=(0, 9),
                     ha="center", fontsize=8.5, color=GREEN)
        ax2.annotate(f"{nm[i]:.2f}%", (i, nm[i]), textcoords="offset points", xytext=(0, -17),
                     ha="center", fontsize=8.5, color=ORANGE)
    ax2.set_ylim(0, 42)
    ax2.set_ylabel("毛利率 / 归母净利率（%）", fontsize=10)
    ax2.spines["top"].set_visible(False)
    ax2.legend(loc="upper left", fontsize=9.5, frameon=False, ncol=2)
    save(fig, f"600887_DCF估值图_02_归母净利润与利润率_{REV}.png", "归母净利润与利润率",
         "柱＝归母净利润（左轴）；线＝毛利率、归母净利率（右轴）")


def fig03():
    ocf = row(RP, 71)
    capex = -row(RP, 72)
    fcf = row(DCF, 13, ["G", "H", "I", "J", "K"])
    fig, ax = plt.subplots(figsize=(10.5, 5.6), dpi=150)
    style(ax, "亿元")
    x = np.arange(5)
    ax.bar(x - 0.19, ocf[5:], width=0.36, color=NAVY, label="经营活动现金流", zorder=3)
    ax.bar(x + 0.19, capex[5:], width=0.36, color=ORANGE, label="资本支出", zorder=3)
    for i in x:
        ax.annotate(f"{ocf[5 + i]:,.0f}", (i - 0.19, ocf[5 + i]), textcoords="offset points",
                    xytext=(0, 4), ha="center", fontsize=9, color=NAVY, fontweight="bold")
        ax.annotate(f"{capex[5 + i]:,.0f}", (i + 0.19, capex[5 + i]), textcoords="offset points",
                    xytext=(0, -14), ha="center", fontsize=9, color="white", fontweight="bold")
    ax.plot(x, fcf, color=RED, marker="D", linewidth=2.2, markersize=6, label="自由现金流 FCF", zorder=5)
    for i in x:
        ax.annotate(f"{fcf[i]:,.1f}", (i, fcf[i]), textcoords="offset points", xytext=(0, 10),
                    ha="center", fontsize=9, color=RED, fontweight="bold",
                    bbox=dict(facecolor="white", edgecolor="none", pad=0.8, alpha=0.85))
    ax.set_ylim(0, max(ocf[5:].max(), capex[5:].max()) * 1.30)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS[5:], fontsize=10)
    ax.legend(loc="upper left", fontsize=9.5, frameon=False, ncol=3)
    save(fig, f"600887_DCF估值图_03_自由现金流与资本开支_{REV}.png",
         "经营现金流、资本开支与自由现金流（2026E–2030E）",
         "FCF = EBIAT + 折旧摊销 − 资本支出 − 营运资本变动")


def fig04():
    ebit = row(DCF, 9, ["G", "H", "I", "J", "K"])
    da = row(DCF, 10, ["G", "H", "I", "J", "K"])
    capex = row(DCF, 11, ["G", "H", "I", "J", "K"])
    nwc = row(DCF, 12, ["G", "H", "I", "J", "K"])
    fcf = row(DCF, 13, ["G", "H", "I", "J", "K"])
    fig, ax = plt.subplots(figsize=(10.5, 5.6), dpi=150)
    style(ax, "亿元")
    x = np.arange(5)
    # 各项目对 FCF 的贡献（正负号统一到「对现金流的影响」）
    comps = [
        (ebit, NAVY, "EBIAT（息前税后利润）"),
        (da, GREEN, "加：折旧摊销"),
        (-capex, ORANGE, "减：资本支出"),
        (-nwc, GRAY, "减：营运资本变动（负值＝占款释放、现金流入）"),
    ]
    cum = np.zeros(5)          # 瀑布式：每段的起点＝前一段的终点
    for vals, color, tag in comps:
        ax.bar(x, vals, bottom=cum, width=0.5, color=color, label=tag, zorder=3)
        cum = cum + vals
    ax.plot(x, fcf, color=RED, marker="D", linewidth=0, markersize=8, zorder=6)
    for i in x:
        ax.annotate(f"{fcf[i]:,.1f}", (i, fcf[i]), textcoords="offset points", xytext=(0, 12),
                    ha="center", fontsize=10, color=RED, fontweight="bold")
    ax.annotate("自由现金流 FCF", (0.02, 0.86), xycoords="axes fraction", fontsize=10,
                color=RED, fontweight="bold")
    ax.set_ylim(0, float(cum.max()) * 1.28)
    ax.set_xticks(x)
    ax.set_xticklabels(LABELS[5:], fontsize=10)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.11), fontsize=9, frameon=False, ncol=2)
    save(fig, f"600887_DCF估值图_04_FCF构成_{REV}.png", "自由现金流构成",
         "堆叠柱＝各项目对 FCF 的贡献，红点＝FCF 合计；资本支出与营运资本变动按「对现金流的影响」取号")


def fig05():
    pv_stage = float(DCF["G18"].value)
    tv_perp = float(DCF["G25"].value)
    tv_mult = float(DCF["G43"].value)
    ps_perp = float(DCF["G35"].value)
    ps_mult = float(DCF["G53"].value)
    implied = float(DCF["G55"].value)
    tv_share = float(DCF["G57"].value) * 100
    g_base = float(DCF["G23"].value) * 100
    mult_base = float(DCF["G41"].value)
    fig, axes = plt.subplots(1, 2, figsize=(11.5, 5.4), dpi=150,
                             gridspec_kw={"width_ratios": [1.05, 1]})
    ax = axes[0]
    style(ax, "亿元", xticks=False)
    x = np.arange(2)
    ax.bar(x, [pv_stage, pv_stage], width=0.5, color=NAVY, label="预测期 FCF 现值", zorder=3)
    ax.bar(x, [tv_perp, tv_mult], bottom=[pv_stage, pv_stage], width=0.5, color=FORE,
           edgecolor=NAVY, label="终值现值", zorder=3)
    labs = [("永续增长法", tv_perp, ps_perp), ("退出乘数法", tv_mult, ps_mult)]
    for i, (tag, tv, ps) in enumerate(labs):
        ax.annotate(f"{pv_stage:,.0f}", (i, pv_stage / 2), ha="center", va="center",
                    fontsize=9.5, color="white", fontweight="bold")
        ax.annotate(f"{tv:,.0f}", (i, pv_stage + tv / 2), ha="center", va="center",
                    fontsize=9.5, color=NAVY, fontweight="bold")
        ax.annotate(f"每股 {ps:,.2f} 元", (i, pv_stage + tv), textcoords="offset points",
                    xytext=(0, 8), ha="center", fontsize=10, color=RED, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels([f"永续增长法（g={g_base:.1f}%）", f"退出乘数法（{mult_base:g}×）"], fontsize=10)
    ax.set_ylim(0, (pv_stage + tv_perp) * 1.22)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), fontsize=9, frameon=False, ncol=2)
    ax = axes[1]
    style(ax, "元 / 股", xticks=False)
    names = ["永续增长法", "退出乘数法", "现价（2026-10-09）"]
    vals = [ps_perp, ps_mult, 27.66]
    cols = [RED, NAVY, GRAY]
    bars = ax.bar(np.arange(3), vals, width=0.5, color=cols, zorder=3)
    for i, v in enumerate(vals):
        ax.annotate(f"{v:,.2f}", (i, v), textcoords="offset points", xytext=(0, 5),
                    ha="center", fontsize=10.5, color=cols[i], fontweight="bold")
    ax.set_xticks(np.arange(3))
    ax.set_xticklabels(names, fontsize=10)
    ax.set_ylim(0, max(vals) * 1.25)
    ax.set_xlim(-0.6, 2.6)
    ax.text(0.02, 0.92, f"终值现值占企业价值 {tv_share:.1f}%\n永续法隐含退出倍数 {implied:.2f}×",
            transform=ax.transAxes, fontsize=9, color="#555555")
    save(fig, f"600887_DCF估值图_05_企业价值构成与两法结果_{REV}.png",
         "企业价值构成与两法估值结果",
         "股权价值 = 企业价值 − 有息负债 508.06 − 少数股东权益 38.89 + 现金 194.95（亿元）；股本 63.25 亿股")


def heat(ax, grid, xlab, ylab, fmt="{:,.1f}", cmap="YlGnBu"):
    im = ax.imshow(grid, cmap=cmap, aspect="auto", origin="lower")
    for i in range(grid.shape[0]):
        for j in range(grid.shape[1]):
            v = grid[i, j]
            if np.isnan(v):
                continue
            shade = "white" if v > np.nanmean(grid) else "#333333"
            ax.text(j, i, fmt.format(v), ha="center", va="center", fontsize=9.5,
                    color=shade, fontweight="bold")
    ax.set_xticks(range(grid.shape[1]))
    ax.set_yticks(range(grid.shape[0]))
    ax.set_xlabel(xlab, fontsize=10)
    ax.set_ylabel(ylab, fontsize=10)
    for s in ax.spines.values():
        s.set_visible(False)
    return im


def fig06():
    wacc = [float(DCF["%s23" % c].value) * 100 for c in "KLMNO"]
    g = [float(DCF["J%d" % r].value) * 100 for r in range(24, 29)]
    grid = np.array([[float(DCF["%s%d" % (c, r)].value) for c in "KLMNO"] for r in range(24, 29)])
    fig, ax = plt.subplots(figsize=(8.6, 5.0), dpi=150)
    im = heat(ax, grid, "WACC（%）", "永续增长率 g（%）")
    ax.set_xticklabels([f"{v:.2f}" for v in wacc], fontsize=9.5)
    ax.set_yticklabels([f"{v:.2f}" for v in g], fontsize=9.5)
    fig.colorbar(im, ax=ax, shrink=0.85, label="元/股")
    save(fig, f"600887_DCF估值图_06_敏感性_WACC与g_{REV}.png",
         "永续增长法：每股价值敏感性（元/股）",
         f"基准 WACC {wacc[1]:.2f}%、g {g[2]:.1f}%，基准每股 {DCF['G35'].value:.2f} 元")


def fig07():
    wacc = [float(DCF["%s41" % c].value) * 100 for c in "KLMNO"]
    mult = [DCF["J%d" % r].value for r in range(42, 47)]
    grid = np.array([[float(DCF["%s%d" % (c, r)].value) for c in "KLMNO"] for r in range(42, 47)])
    fig, ax = plt.subplots(figsize=(8.6, 5.0), dpi=150)
    im = heat(ax, grid, "WACC（%）", "退出乘数（EV/EBITDA，倍）", cmap="YlOrRd")
    ax.set_xticklabels([f"{v:.2f}" for v in wacc], fontsize=9.5)
    ax.set_yticklabels([f"{v:g}" for v in mult], fontsize=9.5)
    fig.colorbar(im, ax=ax, shrink=0.85, label="元/股")
    save(fig, f"600887_DCF估值图_07_敏感性_WACC与退出乘数_{REV}.png",
         "退出乘数法：每股价值敏感性（元/股）",
         f"基准 WACC {wacc[1]:.2f}%、退出乘数 {float(DCF['G41'].value):g}×，基准每股 {DCF['G53'].value:.2f} 元")


def main() -> None:
    print("== 生成 DCF 估值图 ==")
    fig01(); fig02(); fig03(); fig04(); fig05(); fig06(); fig07()
    print("\n== 关键数 ==")
    print("  永续法每股 %.2f 元 | 乘数法每股 %.2f 元 | 终值占比 %.1f%% | 隐含退出倍数 %.2f×"
          % (DCF["G35"].value, DCF["G53"].value, DCF["G57"].value * 100, DCF["G55"].value))


if __name__ == "__main__":
    main()
