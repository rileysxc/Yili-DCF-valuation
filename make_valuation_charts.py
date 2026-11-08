#!/usr/bin/env python3
"""生成 10.15 任务 5 的估值图表（每张独立保存，研报风格）。

输出到 可视化图表/600887_估值图/：
    01 PE/PB 历史分位带    02 同业估值对比    03 DCF 两法对照
    04 报告版 DCF 区间结论图（仅区间/中枢/现价，不含敏感性矩阵）
    05 合理价值区间合成（各方法端点 + 最终区间，足球场图）

数据来源（全部为已核产物）：
    - 原始数据/07_清洗后数据/<公司>/<代码>_估值_20261011_Rev05.csv     （5 年日频 PE/PB 序列）
    - 数据文件/估值分位表/三家分位表_20261015_Rev00.xlsx              （分位点与当前分位）
    - 数据文件/同业对比表/同业对比表_20261015_Rev00.xlsx              （可比公司估值与 ROE）
    - 数据文件/隐含倍数对照表/隐含倍数对照表_20261015_Rev02.xlsx      （EV/EBITDA，三家统一按 07 层市值）
    - 数据文件/DCF模型/600887_DCF模型_20261012_Rev07.xlsx             （两法结果与两张敏感性矩阵）

用法（需 bundled Python，含 matplotlib）：
    "<bundled python>" make_valuation_charts.py              # 三张全出
    "<bundled python>" make_valuation_charts.py --only 3     # 只重出图 3
"""

from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_val"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

BASE = Path(__file__).resolve().parents[2]
CLEAN = BASE / "原始数据" / "07_清洗后数据"
DATA = BASE / "数据文件"
OUTDIR = BASE / "可视化图表" / "600887_估值图"
REV_DATA = "20261015_Rev00"          # 输入数据文件（分位表、同业对比表）的版本
PEER_REV = "20261016_Rev01"          # 同业对比表版本（Rev01 仅表头更名，数值未变）
# 输出图版本：图 2 于 2026-10-16 修正分组标签与口径说明后重出 Rev01；
# 全部三张于 2026-10-16 因「估值章由第五章改为第四章」把图内标题由「图 5-x」改为「图 4-x」，齐升一版。
# 图内新标题：估值图 01（PE/PB 分位带）→ 图 4-4；02（同业估值对比）→ 图 4-5；03（DCF 两法对照与敏感性）→ 图 4-3。
FIG_REV = {"1": "20261020_Rev01", "2": "20261020_Rev01", "3": "20261020_Rev01"}

CODES = [("伊利", "600887", "600887_伊利股份", "#1F4E79"),
         ("光明", "600597", "600597_光明乳业", "#E67E22"),
         ("新乳业", "002946", "002946_新乳业", "#2E7D32")]
DCF_XLSX = DATA / "DCF模型" / "600887_DCF模型_20261012_Rev07.xlsx"
PCT_XLSX = DATA / "估值分位表" / f"三家分位表_{REV_DATA}.xlsx"
PEER_XLSX = DATA / "同业对比表" / f"同业对比表_{PEER_REV}.xlsx"
IMP_XLSX = DATA / "隐含倍数对照表" / "隐含倍数对照表_20261015_Rev02.xlsx"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False
NAVY, RED, GREEN, ORANGE, GRAY = "#1F4E79", "#C00000", "#2E7D32", "#E67E22", "#8C8C8C"
LIGHT, SOFT = "#9DC3E6", "#DCE6F1"
BAND = ["#F2F2F2", "#DCE6F1", "#F2F2F2"]  # P10–P25 / P25–P75 / P75–P90


def daily(code_dir: str):
    df = pd.read_csv(CLEAN / code_dir / f"{code_dir.split('_')[0]}_估值_20261011_Rev05.csv")
    df["date"] = pd.to_datetime(df["date"])
    return df.sort_values("date")


def pct_row(ws, tag: str):
    """从分位表读一行，返回 dict。"""
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == tag:
            return {ws.cell(1, c).value: ws.cell(r, c).value for c in range(2, ws.max_column + 1)}
    raise KeyError(tag)


SHORT = {"伊利股份": "伊利", "光明乳业": "光明", "新乳业": "新乳业"}


def short(v) -> str:
    v = str(v)
    return SHORT.get(v, v)


# ---------------------------------------------------------------- 图 1
def fig_percentile_band(rev: str | None = None):
    rev = rev or FIG_REV["1"]
    wb = openpyxl.load_workbook(PCT_XLSX, data_only=True)
    ws = wb["分位表_5年日频_主"]
    fig, axes = plt.subplots(2, 3, figsize=(16.5, 8.4), sharex=False)
    for j, (name, code, folder, color) in enumerate(CODES):
        df = daily(folder)
        for i, (col, tag_metric, title) in enumerate(
                [("PE(TTM)", "PE(TTM)", "PE(TTM)（倍）"), ("市净率", "PB", "PB(LF)（倍）")]):
            ax = axes[i][j]
            s = df[["date", col]].dropna()
            valid = s[s[col] > 0]
            p = pct_row(ws, f"{name}_{tag_metric}")
            q = [p["P10"], p["P25"], p["P50"], p["P75"], p["P90"]]
            ax.axhspan(q[0], q[1], color=BAND[0], zorder=0)
            ax.axhspan(q[1], q[3], color=BAND[1], zorder=0)
            ax.axhspan(q[3], q[4], color=BAND[2], zorder=0)
            for y, lab in zip(q, ["P10", "P25", "P50", "P75", "P90"]):
                ax.axhline(y, color=GRAY, lw=0.7, ls="--", zorder=1)
                ax.text(df["date"].iloc[-1], y, f" {lab} {y:.2f}", va="center", ha="left",
                        fontsize=7.5, color=GRAY)
            ax.plot(valid["date"], valid[col], color=color, lw=1.1, zorder=3)
            if len(valid) < len(s):                      # 有被剔除的非正值：标出剔除区间
                ax.axvspan(valid["date"].iloc[-1], s["date"].iloc[-1],
                           color=RED, alpha=0.07, zorder=0)
            if tag_metric == "PB" or name != "光明":
                ax.scatter([valid["date"].iloc[-1]], [valid[col].iloc[-1]], color=RED, s=34,
                           zorder=5, label=f"当前 {valid[col].iloc[-1]:.2f}")
            else:
                ax.scatter([valid["date"].iloc[-1]], [valid[col].iloc[-1]], color=RED, s=34,
                           zorder=5, marker="D", label=f"最后有效值 {valid[col].iloc[-1]:.2f}")
                ax.annotate("2026-03-31 起 PE 转负\n（红色区间：128 个交易日已剔除）",
                            xy=(valid["date"].iloc[-1], valid[col].iloc[-1]),
                            xytext=(0.05, 0.28), textcoords="axes fraction", fontsize=8,
                            color=RED, arrowprops=dict(arrowstyle="->", color=RED, lw=0.8))
            cur_pct = p["当前分位(%)"]
            lab = f"当前分位 {cur_pct}%" if cur_pct == cur_pct else ""
            if name == "光明" and tag_metric == "PE(TTM)":
                lab = "当前分位不适用"
            ax.set_title(f"{name}　{title}　{lab}", fontsize=10, color=NAVY)
            ax.grid(alpha=0.18, lw=0.6)
            ax.legend(fontsize=7.5, loc="upper right", framealpha=0.85)
            ax.tick_params(axis="x", labelsize=7.5, rotation=0)
            ax.tick_params(axis="y", labelsize=8)
    for ax in axes[0]:
        ax.set_ylim(bottom=0)
    fig.text(0.5, 0.022,
             "深蓝阴影 = 历史 P25–P75 区间，浅灰阴影 = P10–P25 与 P75–P90，虚线为五个分位点；红点为当前值。"
             "光明 PE(TTM) 自 2026-03-31 起为负（剔除 128 个交易日、占 9.18%），其当前分位不适用，"
             "该段在图 4-5 中以红色竖带标出、不参与曲线。",
             ha="center", fontsize=8, color=GRAY)
    fig.tight_layout(rect=(0, 0.05, 1, 0.99))
    OUTDIR.mkdir(parents=True, exist_ok=True)
    out = OUTDIR / f"600887_估值图_01_PE-PB历史分位带_{rev}.png"
    fig.savefig(out, dpi=170, facecolor="white")
    plt.close(fig)
    print("已保存", out)


# ---------------------------------------------------------------- 图 2
def fig_peer(rev: str | None = None):
    rev = rev or FIG_REV["2"]
    wb = openpyxl.load_workbook(PEER_XLSX, data_only=True)
    ws = wb["可比公司"]
    rows = {}
    for r in (6, 7, 8):
        rows[short(ws.cell(r, 2).value)] = {
            "code": ws.cell(r, 3).value, "price": ws.cell(r, 4).value,
            "mv": ws.cell(r, 5).value, "pe": ws.cell(r, 6).value, "pe_static": ws.cell(r, 7).value,
            "pb": ws.cell(r, 8).value, "roe": ws.cell(r, 9).value,
            "pe_p5": ws.cell(r, 10).value, "pb_p5": ws.cell(r, 11).value,
            "pe_p10": ws.cell(r, 12).value, "pb_p10": ws.cell(r, 13).value}

    imp = openpyxl.load_workbook(IMP_XLSX, data_only=True)["隐含倍数对照表"]
    ev_rows = {short(imp.cell(r, 2).value):
               [imp.cell(r, c).value for c in range(3, 8)] for r in (4, 5, 6)}

    fig = plt.figure(figsize=(16.5, 8.6))
    gs = fig.add_gridspec(2, 2, height_ratios=[1, 1], hspace=0.42, wspace=0.20)
    names = [n for n in ["伊利", "光明", "新乳业"]]
    x = np.arange(len(names))

    # A：当期估值倍数
    ax = fig.add_subplot(gs[0, 0])
    w = 0.30
    pe = [rows[n]["pe"] if isinstance(rows[n]["pe"], (int, float)) else np.nan for n in names]
    pb = [rows[n]["pb"] for n in names]
    ax.bar(x - w / 2, pe, w, color=NAVY, label="PE(TTM)（倍）")
    ax.bar(x + w / 2, pb, w, color=ORANGE, label="PB(LF)（倍）")
    for xi, (a, b) in enumerate(zip(pe, pb)):
        if a == a:
            ax.text(xi - w / 2, a + 0.3, f"{a:.2f}", ha="center", fontsize=9)
        ax.text(xi + w / 2, b + 0.3, f"{b:.2f}", ha="center", fontsize=9)
    ax.annotate("PE(TTM) 为负\n（−83.74 倍）", xy=(x[1] - w / 2, 0.35),
                xytext=(x[1] - w / 2, 6.5), ha="center", fontsize=8, color=RED,
                arrowprops=dict(arrowstyle="->", color=RED, lw=0.8))
    ax.set_xticks(x, names)
    ax.set_title("A　2026-10-09 当期估值倍数", fontsize=11, color=NAVY)
    ax.legend(fontsize=8.5, framealpha=0.85)
    ax.grid(axis="y", alpha=0.2, lw=0.6)
    ax.set_ylim(0, 20)

    # B：分位
    ax = fig.add_subplot(gs[0, 1])
    # 标签只按窗口命名：「主」「稳健」是第五章区间合成里的角色（PE 5 年日频进合成、PE 10 年周度作稳健性），
    # PB 只作辅证、不进合成，故不冠「主／稳健」字样（用户 2026-10-16 决定）。
    groups = [("PE 日频(主)", "pe_p5"), ("PB 日频", "pb_p5"),
              ("PE 周度(稳健)", "pe_p10"), ("PB 周度", "pb_p10")]
    w = 0.2
    for k, (lab, key) in enumerate(groups):
        vals = [rows[n][key] if isinstance(rows[n][key], (int, float)) else np.nan for n in names]
        pos = x + (k - 1.5) * w
        ax.bar(pos, vals, w, label=lab,
               color=[NAVY, ORANGE, LIGHT, GREEN][k])
        for xi, v in zip(pos, vals):
            if v == v:
                ax.text(xi, v + 0.6, f"{v:.1f}", ha="center", fontsize=7.5)
    for dx in (-1.5 * w, 0.5 * w):
        ax.text(x[1] + dx, 1.0, "PE 不适用", ha="center", fontsize=6.8, color=RED)
    ax.set_xticks(x, names)
    ax.set_ylim(0, 46)
    ax.set_title("B　PE/PB 历史分位（%，截至 2026-10-09）", fontsize=11, color=NAVY)
    ax.legend(fontsize=8, ncol=2, framealpha=0.85, loc="upper center")
    ax.grid(axis="y", alpha=0.2, lw=0.6)

    # C：EV/EBITDA
    ax = fig.add_subplot(gs[1, 0])
    years = ["2021A", "2022A", "2023A", "2024A", "2025A"]
    for (name, code, folder, color) in CODES:
        key = name
        ax.plot(years, ev_rows[key], marker="o", ms=5, lw=1.8, color=color, label=name)
        for xi, v in enumerate(ev_rows[key]):
            ax.text(xi, v + 0.45, f"{v:.2f}", ha="center", fontsize=7.5, color=color)
    ax.set_title("C　EV/EBITDA（倍，三家统一 07 层市值）", fontsize=11, color=NAVY)
    ax.legend(fontsize=8.5, framealpha=0.85, loc="upper right")
    ax.grid(alpha=0.2, lw=0.6)
    ax.set_ylim(0, 24)

    # D：ROE 与 ROE 差
    ax = fig.add_subplot(gs[1, 1])
    roe = [rows[n]["roe"] * 100 for n in names]
    bars = ax.bar(x, roe, 0.45, color=[NAVY, ORANGE, GREEN])
    for xi, v in zip(x, roe):
        ax.text(xi, v + (0.6 if v >= 0 else -1.9), f"{v:.2f}%", ha="center", fontsize=10)
    ax.axhline(0, color=GRAY, lw=0.9)
    ax.set_xticks(x, names)
    ax.set_title("D　2025A ROE（期末归母净资产口径，%）", fontsize=11, color=NAVY)
    ax.grid(axis="y", alpha=0.2, lw=0.6)
    ax.set_ylim(-4, 26)

    fig.subplots_adjust(top=0.94)
    fig.text(0.5, 0.012,
             "总市值（2026-10-09）：伊利 1749.59 / 光明 97.04 / 新乳业 131.60 亿元。"
             "ROE 三家统一为期末归母净资产口径；伊利与新乳业 2025A 差 2.39pct（2021A 为 6.41pct）。"
             "PB 分位在第五章只作辅证、不参与区间合成。EV/EBITDA 三家统一按 07 层市值，"
             "比 DCF 模型缓存值（按 63.25 亿股）高约 0.0006 倍。",
             ha="center", fontsize=8, color=GRAY)
    out = OUTDIR / f"600887_估值图_02_同业估值对比_{rev}.png"
    fig.savefig(out, dpi=170, facecolor="white")
    plt.close(fig)
    print("已保存", out)


# ---------------------------------------------------------------- 图 3
def fig_dcf(rev: str | None = None):
    rev = rev or FIG_REV["3"]
    wb = openpyxl.load_workbook(DCF_XLSX, data_only=True)["DCF"]
    grid = lambda r0, r1, c0, c1: np.array(
        [[wb.cell(r, c).value for c in range(c0, c1 + 1)] for r in range(r0, r1 + 1)], dtype=float)
    perp = grid(24, 28, 11, 15)       # K24:O28  WACC×g
    mult = grid(42, 46, 11, 15)       # K42:O46  WACC×退出乘数
    wacc = [wb[f"{c}23"].value for c in "KLMNO"]
    gs_ = [wb[f"J{r}"].value for r in range(24, 29)]
    mults = [wb[f"J{r}"].value for r in range(42, 47)]
    pv_growth, pv_mult, price = wb["G35"].value, wb["G53"].value, 27.66
    g_base, mult_base = wb["G23"].value, wb["G41"].value
    wacc_base = wb["P16"].value

    fig = plt.figure(figsize=(16.5, 6.4))
    gs = fig.add_gridspec(1, 3, width_ratios=[1, 1.15, 1.15], wspace=0.32)

    ax = fig.add_subplot(gs[0, 0])
    vals = [pv_growth, pv_mult, price]
    cols = [NAVY, ORANGE, RED]
    labs = [f"永续增长法\n{pv_growth:.2f} 元", f"退出乘数法\n{pv_mult:.2f} 元", "现价\n27.66 元"]
    b = ax.bar([0, 1, 2], vals, 0.52, color=cols)
    for xi, v in zip([0, 1, 2], vals):
        ax.text(xi, v + 0.6, f"{v:.2f}", ha="center", fontsize=12, fontweight="bold")
    ax.set_xticks([0, 1, 2], labs, fontsize=9.5)
    ax.set_ylim(0, max(vals) * 1.18)
    ax.axhline(price, color=RED, lw=1.0, ls="--", alpha=0.7)
    ax.set_title("A　DCF 两法结果与现价（元/股）", fontsize=11, color=NAVY)
    ax.grid(axis="y", alpha=0.2, lw=0.6)
    diff = pv_growth - pv_mult
    diff_pct = diff / ((pv_growth + pv_mult) / 2)
    ax.text(0.5, 0.90, f"两法差 {diff:.2f} 元（{diff_pct:.1%}，分母为两法均值）", transform=ax.transAxes,
            ha="center", fontsize=9, color=GRAY)

    for k, (mat, ylab, ytick, ttl, colbar) in enumerate(
            [(perp, "永续增长率 g", gs_, "B　永续增长法：WACC × g（每股价值，元）", "Blues"),
             (mult, "退出乘数（倍）", mults, "C　退出乘数法：WACC × 退出乘数（每股价值，元）", "Oranges")]):
        ax = fig.add_subplot(gs[0, k + 1])
        im = ax.imshow(mat, cmap=colbar, aspect="auto", origin="lower")
        for i in range(mat.shape[0]):
            for j in range(mat.shape[1]):
                ax.text(j, i, f"{mat[i, j]:.2f}", ha="center", va="center", fontsize=9,
                        color="white" if (mat[i, j] - mat.min()) / (mat.max() - mat.min()) > 0.62 else "#333333")
        ax.set_xticks(range(5), [f"{w * 100:.2f}%" for w in wacc], fontsize=9)
        ax.set_yticks(range(5), [f"{g * 100:.2f}%" for g in ytick] if k == 0 else [str(m) for m in ytick],
                      fontsize=9)
        ax.set_xlabel("WACC", fontsize=9.5)
        ax.set_ylabel(ylab, fontsize=9.5)
        ax.set_title(ttl, fontsize=11, color=NAVY)
        y_base = int(np.argmin(np.abs(np.array(ytick) - (g_base if k == 0 else mult_base))))
        x_base = int(np.argmin(np.abs(np.array(wacc) - wacc_base)))
        ax.scatter([x_base], [y_base], s=170, facecolor="none", edgecolor=RED, lw=2.0)
        fig.colorbar(im, ax=ax, fraction=0.046, pad=0.03)

    fig.text(0.5, 0.015,
             f"基准组合：WACC {wacc_base:.2%}、g {g_base:.2%}、退出乘数 {mult_base:g}×；红框为基准格。"
             f"永续增长法结果区间 {perp.min():.2f} ~ {perp.max():.2f} 元、退出乘数法 {mult.min():.2f} ~ {mult.max():.2f} 元（矩阵覆盖范围）。"
             "",
             ha="center", fontsize=8, color=GRAY)
    fig.tight_layout(rect=(0, 0.045, 1, 0.99))
    out = OUTDIR / f"600887_估值图_03_DCF两法对照_{rev}.png"
    fig.savefig(out, dpi=170, facecolor="white")
    plt.close(fig)
    print("已保存", out)


# ---------------------------------------------------------------- 图 4（报告版）
def fig_dcf_report(rev: str | None = None):
    """报告版专用的 DCF 结论图：只保留区间、中枢、现价，不含敏感性矩阵。"""
    rev = rev or "20261023_Rev00"
    ws = openpyxl.load_workbook(DCF_XLSX, data_only=True)["DCF"]
    perp, mult = ws["G35"].value, ws["G53"].value
    mid = (perp + mult) / 2
    price = 27.66
    low_gap = price / mult - 1

    fig, ax = plt.subplots(figsize=(10.8, 3.6), dpi=170)
    ax.set_xlim(24, 41)
    ax.set_ylim(0, 1)
    ax.set_yticks([])
    ax.set_xlabel("每股价值（元）", fontsize=10)
    ax.grid(axis="x", color="#E6E6E6", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#BBBBBB")
    ax.tick_params(axis="x", labelsize=9.5, colors="#333333")

    ax.barh(0.50, perp - mult, left=mult, height=0.20, color=SOFT,
            edgecolor=NAVY, linewidth=1.1, zorder=2)
    ax.plot([mid, mid], [0.36, 0.64], color=NAVY, linestyle=(0, (4, 3)),
            linewidth=1.3, zorder=3)
    ax.plot([price, price], [0.06, 0.94], color=RED, linewidth=1.6, zorder=4)

    ax.text(mid, 0.66, f"区间 {mult:.2f}–{perp:.2f} 元",
            ha="center", va="bottom", fontsize=9, color=GRAY)
    ax.text(mid, 0.28, f"中枢 {mid:.2f}", ha="center", va="top",
            fontsize=9.5, color=NAVY)
    ax.text(mult, 0.20, f"退出乘数法\n{mult:.2f} 元（12×）", ha="center",
            va="top", fontsize=9, color="#333333")
    ax.text(perp, 0.20, f"永续增长法\n{perp:.2f} 元（g 2.00%）", ha="center",
            va="top", fontsize=9, color="#333333")
    ax.annotate(f"现价 {price:.2f} 元\n低于下限 {abs(low_gap):.2%}",
                xy=(price, 0.94), xytext=(price - 0.9, 0.90),
                ha="right", va="top", fontsize=9, color=RED,
                arrowprops=dict(arrowstyle="-", color=RED, lw=0.9))

    fig.tight_layout()
    OUTDIR.mkdir(parents=True, exist_ok=True)
    out = OUTDIR / f"600887_估值图_04_DCF区间_{rev}.png"
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存", out)


# ---------------------------------------------------------------- 图 5
def fig_football_field(rev: str | None = None):
    """4.3.3 合理价值区间合成：各方法端点 + 最终区间。"""
    rev = rev or "20261023_Rev00"
    price = 27.66
    # 端点值：DCF 两法取模型 G35/G53；PE/PB 取表 4-7 的 P25–P75 × EPS 1.6004 / BVPS 8.4914；
    # 可比公司取 32.31 元。与面试版表 4-7／4-8 一致。
    rows = [
        ("DCF 永续增长法", "point", 38.30, NAVY),
        ("DCF 退出乘数法", "point", 29.19, ORANGE),
        ("PE 5年日频（主）", "range", (25.48, 39.74), "#2E75B6"),
        ("PE 10年周度（稳健）", "range", (31.54, 43.32), SOFT),
        ("PB 5年日频（辅证）", "range", (26.75, 41.01), "#BFBFBF"),
        ("PB 10年周度（辅证）", "range", (28.11, 60.29), "#D9D9D9"),
        ("可比公司法（TTM）", "point", 32.31, ORANGE),
        ("DCF 最终区间", "range", (29.19, 38.30), NAVY),
    ]

    fig, ax = plt.subplots(figsize=(13.4, 6.6), dpi=170)
    y = np.arange(len(rows))
    for i, (label, kind, val, color) in enumerate(rows):
        is_final = label == "DCF 最终区间"
        if kind == "range":
            lo, hi = val
            ax.barh(i, hi - lo, left=lo, height=0.46, color=color,
                    alpha=0.22 if is_final else 0.55,
                    edgecolor=NAVY if is_final else "#8C8C8C",
                    linewidth=1.8 if is_final else 0.8, zorder=3)
            ax.text(hi + 0.45, i, f"{lo:.2f}–{hi:.2f}", va="center", ha="left",
                    fontsize=9.5, color=NAVY if is_final else "#333333",
                    fontweight="bold" if is_final else "normal")
        else:
            ax.scatter([val], [i], s=110, color=color, edgecolor="white",
                       linewidth=1.0, zorder=5)
            ax.text(val + 0.5, i, f"{val:.2f}", va="center", ha="left",
                    fontsize=9.5, color="#333333")

    ax.axvline(price, color=RED, linewidth=1.6, zorder=4)
    ax.text(price - 0.35, -0.92, f"现价 {price:.2f}", color=RED,
            fontsize=9.5, ha="right", va="center")

    ax.set_yticks(y, [r[0] for r in rows], fontsize=10.5)
    ax.tick_params(axis="y", length=0)
    ax.set_xlim(22, 62)
    ax.set_ylim(len(rows) - 0.35, -1.15)
    ax.set_xticks(np.arange(22, 63, 4))
    ax.set_xlabel("每股价值（元）", fontsize=10)
    ax.grid(axis="x", color="#E6E6E6", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#BBBBBB")
    ax.tick_params(axis="x", labelsize=9.5, colors="#333333")

    legend = [
        Patch(facecolor=NAVY, alpha=0.22, edgecolor=NAVY, label="DCF 最终区间"),
        Patch(facecolor="#2E75B6", alpha=0.55, edgecolor="#8C8C8C", label="PE 区间"),
        Patch(facecolor="#BFBFBF", alpha=0.55, edgecolor="#8C8C8C", label="PB 区间（辅证）"),
        Line2D([0], [0], marker="o", color="w", markerfacecolor=ORANGE,
               markersize=8, label="单点方法"),
        Line2D([0], [0], color=RED, linewidth=1.6, label="现价"),
    ]
    ax.legend(handles=legend, loc="lower center", bbox_to_anchor=(0.5, -0.24),
              ncol=5, frameon=False, fontsize=9)

    fig.tight_layout()
    OUTDIR.mkdir(parents=True, exist_ok=True)
    out = OUTDIR / f"600887_估值图_05_合理价值区间合成_{rev}.png"
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存", out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成第四章估值图表（图 1–图 5 各自独立保存）")
    ap.add_argument("--only", default="1,2,3,4,5", help="只重出指定图，逗号分隔，例如 --only 3")
    ap.add_argument("--rev", default=None, help="覆盖输出版本号，例如 --rev 20261019_Rev00")
    args = ap.parse_args()

    figs = {"1": fig_percentile_band, "2": fig_peer, "3": fig_dcf,
            "4": fig_dcf_report, "5": fig_football_field}
    for key in sorted({k.strip() for k in args.only.split(",") if k.strip()}):
        figs[key](args.rev)
