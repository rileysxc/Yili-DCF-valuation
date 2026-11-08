#!/usr/bin/env python3
"""生成 10.14 任务 5 的财务质量图表（每张独立保存，研报风格）。

输出到 可视化图表/600887_财务质量图/：
    01 利润与现金流对照      02 减值与扣非还原      03 ROE 三家对比

数据来源（全部为已核产物，不另取数）：
    - 数据文件/财务分析附表/600887_利润与现金流对照表_20261014_Rev00.xlsx
    - 数据文件/财务分析附表/600887_指标驱动分解表_20261014_Rev00.xlsx（三家对比页）
    - 原始数据/07_清洗后数据/600887_伊利股份/600887_利润表_20261011_Rev05.csv（两项减值损失）
"""

from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_fq"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl

BASE = Path(__file__).resolve().parents[2]
TABLES = BASE / "数据文件" / "财务分析附表"
CLEAN = BASE / "原始数据" / "07_清洗后数据" / "600887_伊利股份"
OUTDIR = BASE / "可视化图表" / "600887_财务质量图"
REV = "20261014_Rev00"
OUT_REV = "20261020_Rev01"

CASH = TABLES / f"600887_利润与现金流对照表_{REV}.xlsx"
DRIVER = TABLES / f"600887_指标驱动分解表_{REV}.xlsx"
PL_CSV = CLEAN / "600887_利润表_20261011_Rev05.csv"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

NAVY, RED, GREEN, ORANGE, GRAY = "#1F4E79", "#C00000", "#2E7D32", "#E67E22", "#8C8C8C"
LIGHT = "#9DC3E6"
LABELS = ["2020A", "2021A", "2022A", "2023A", "2024A", "2025H1", "2025A", "2026H1"]
PERIODS = ["2020-12-31", "2021-12-31", "2022-12-31", "2023-12-31", "2024-12-31",
           "2025-06-30", "2025-12-31", "2026-06-30"]
X = np.arange(len(LABELS))
SPLIT = 4.5  # 年度块（2020A–2024A）与「2025H1 / 2025A / 2026H1」之间
BOX = {"boxstyle": "round,pad=0.12", "facecolor": "white", "edgecolor": "none", "alpha": 0.82}


def find_row(ws, prefix: str) -> int:
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).startswith(prefix):
            return r
    raise KeyError(prefix)


def series(ws, prefix: str) -> np.ndarray:
    r = find_row(ws, prefix)
    return np.array([ws.cell(r, c).value for c in range(2, 10)], dtype=float)


def pl_series(subject: str) -> np.ndarray:
    with PL_CSV.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    idx = rows[0].index(subject)
    by_period = {r[0].strip(): r for r in rows[1:]}
    return np.array([float(by_period[p][idx]) for p in PERIODS])


def style(ax, ylab="", xticks=True, split=True):
    ax.grid(axis="y", color="#E4E4E4", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")
    if xticks:
        ax.set_xticks(X)
        ax.set_xticklabels(LABELS, fontsize=9.5)
    if split:
        ax.axvline(SPLIT, color=RED, linestyle=(0, (4, 3)), linewidth=1.0, zorder=1)
    if ylab:
        ax.set_ylabel(ylab, fontsize=10)


def save(fig, name, title, subtitle):
    fig.text(0.055, 0.955, subtitle, fontsize=9, color="#666666", ha="left", va="top")
    OUTDIR.mkdir(parents=True, exist_ok=True)
    out = OUTDIR / name
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存:", out.name)


# ------------------------------------------------------------------ 图 3-1


def fig01():
    ws = openpyxl.load_workbook(CASH, data_only=True)["对照表"]
    np_parent = series(ws, "归母净利润")
    deducted = series(ws, "扣非归母净利润")
    ocf = series(ws, "经营活动产生的现金流量净额")
    ratio = series(ws, "经营现金流 ÷ 归母净利润")

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10.6, 7.4), dpi=150,
                                   gridspec_kw={"height_ratios": [2.05, 1], "hspace": 0.34})
    fig.subplots_adjust(top=0.86)

    style(ax1, "亿元")
    ax1.bar(X - 0.19, np_parent, width=0.37, color=NAVY, label="归母净利润", zorder=3)
    ax1.bar(X + 0.19, deducted, width=0.37, color=LIGHT, edgecolor=NAVY, linewidth=0.6,
            label="扣非归母净利润", zorder=3)
    ax1.plot(X, ocf, color=ORANGE, marker="o", markersize=6, linewidth=2.2,
             label="经营活动现金流量净额", zorder=5)
    for i in X:
        ax1.annotate(f"{np_parent[i]:,.0f}", (i - 0.19, np_parent[i]), textcoords="offset points",
                     xytext=(0, 4), ha="center", fontsize=8.5, color=NAVY)
        ax1.annotate(f"{deducted[i]:,.0f}", (i + 0.19, deducted[i]), textcoords="offset points",
                     xytext=(0, 4), ha="center", fontsize=8.5, color="#5B7FA6")
        ax1.annotate(f"{ocf[i]:,.0f}", (i, ocf[i]), textcoords="offset points", xytext=(0, 8),
                     ha="center", fontsize=8.5, color=ORANGE, fontweight="bold", bbox=BOX)
    ax1.set_ylim(0, max(ocf.max(), np_parent.max()) * 1.24)
    ax1.legend(loc="upper left", frameon=False, fontsize=9.5, ncol=3)

    style(ax2, "倍")
    bars = ax2.bar(X, ratio, width=0.5, color=[GREEN if v >= 1 else "#D98880" for v in ratio], zorder=3)
    ax2.axhline(1.0, color=GRAY, linewidth=1.0, linestyle=(0, (5, 3)), zorder=2)
    for i, _b in zip(X, bars):
        ax2.annotate(f"{ratio[i]:.2f}", (i, ratio[i]), textcoords="offset points", xytext=(0, 4),
                     ha="center", fontsize=8.5, color="#444444")
    ax2.set_ylim(0, ratio.max() * 1.28)
    ax2.text(len(LABELS) - 0.45, ratio.max() * 1.17, "虚线 = 1.0 参考线", fontsize=8.5,
             color=GRAY, ha="right")

    save(fig, f"600887_财务质量图_01_利润与现金流对照_{OUT_REV}.png",
         "伊利股份：利润与现金流对照（8 期）",
         "上：归母净利润、扣非归母净利润、经营活动现金流量净额（亿元）｜下：经营活动现金流 ÷ 归母净利润（倍）"
         "｜2025H1 与 2026H1 为半年累计值，只与同期比")


# ------------------------------------------------------------------ 图 3-2


def fig02():
    ws = openpyxl.load_workbook(CASH, data_only=True)["对照表"]
    np_parent = series(ws, "归母净利润")
    deducted = series(ws, "扣非归母净利润")
    a_imp = np.abs(pl_series("资产减值损失")) / 1e8
    c_imp = np.abs(pl_series("信用减值损失")) / 1e8

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10.6, 7.4), dpi=150,
                                   gridspec_kw={"height_ratios": [1.35, 1], "hspace": 0.34})
    fig.subplots_adjust(top=0.86)

    style(ax1, "亿元")
    ax1.bar(X - 0.19, np_parent, width=0.37, color=NAVY, label="归母净利润", zorder=3)
    ax1.bar(X + 0.19, deducted, width=0.37, color=LIGHT, edgecolor=NAVY, linewidth=0.6,
            label="扣非归母净利润", zorder=3)
    for i in X:
        ax1.annotate(f"{np_parent[i]:,.0f}", (i - 0.19, np_parent[i]), textcoords="offset points",
                     xytext=(0, 4), ha="center", fontsize=8.5, color=NAVY)
        ax1.annotate(f"{deducted[i]:,.0f}", (i + 0.19, deducted[i]), textcoords="offset points",
                     xytext=(0, 4), ha="center", fontsize=8.5, color="#5B7FA6")
    ax1.set_ylim(0, np_parent.max() * 1.3)
    ax1.legend(loc="upper left", frameon=False, fontsize=9.5, ncol=2)

    style(ax2, "亿元")
    ax2.bar(X - 0.19, a_imp, width=0.37, color=RED, label="资产减值损失", zorder=3)
    ax2.bar(X + 0.19, c_imp, width=0.37, color="#F0A3A3", edgecolor=RED, linewidth=0.6,
            label="信用减值损失", zorder=3)
    for i in X:
        ax2.annotate(f"{a_imp[i]:,.1f}", (i - 0.19, a_imp[i]), textcoords="offset points", xytext=(0, 4),
                     ha="center", fontsize=8.5, color=RED)
        if c_imp[i] > 0.05:
            ax2.annotate(f"{c_imp[i]:,.1f}", (i + 0.19, c_imp[i]), textcoords="offset points", xytext=(0, 4),
                         ha="center", fontsize=8.5, color="#B05A5A")
    ax2.set_ylim(0, a_imp.max() * 1.32)
    ax2.legend(loc="upper left", frameon=False, fontsize=9.5, ncol=2)

    save(fig, f"600887_财务质量图_02_减值与扣非还原_{OUT_REV}.png",
         "伊利股份：减值与扣非还原（8 期）",
         "上：归母净利润与扣非归母净利润（亿元）｜下：资产减值损失与信用减值损失（亿元，取绝对值）"
         "｜2026H1 资产减值 24.56 亿元")


# ------------------------------------------------------------------ 图 3-3


def fig03():
    ws = openpyxl.load_workbook(DRIVER, data_only=True)["三家对比"]
    labels = ["2021A", "2022A", "2023A", "2024A", "2025A", "2026H1"]
    x = np.arange(len(labels))
    rows: dict[str, np.ndarray] = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not v or "——" not in str(v):
            continue
        name = str(v).split("——")[-1].strip()
        if name in ("伊利股份", "光明乳业", "新乳业"):
            rows[name] = np.array([ws.cell(r, c).value for c in range(2, 8)], dtype=float) * 100

    fig, ax = plt.subplots(figsize=(10.6, 5.8), dpi=150)
    style(ax, "ROE（%）", xticks=False, split=False)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.axvline(4.5, color=RED, linestyle=(0, (4, 3)), linewidth=1.0, zorder=1)

    plan = [("伊利股份", NAVY, 2.8, 7.5), ("新乳业", GREEN, 1.7, 5.0), ("光明乳业", GRAY, 1.7, 5.0)]
    for name, color, lw, ms in plan:
        ax.plot(x, rows[name], color=color, marker="o", linewidth=lw, markersize=ms,
                label=name, zorder=5 if name == "伊利股份" else 4)
    # 标签用横向避让：伊利在点左侧、新乳业在点右侧、光明在点上方，避免同一点位叠字
    placement = {"伊利股份": (-9, 0, "right"), "新乳业": (9, 4, "left"), "光明乳业": (0, 9, "center")}
    for name, color, _lw, _ms in plan:
        y = rows[name]
        dx, dy, ha = placement[name]
        for i in x:
            ax.annotate(f"{y[i]:.2f}", (i, y[i]), textcoords="offset points", xytext=(dx, dy),
                        ha=ha, fontsize=8.2, color=color,
                        fontweight="bold" if name == "伊利股份" else "normal", bbox=BOX, zorder=8)
    ax.axhline(0, color="#AAAAAA", linewidth=0.9, linestyle=":")
    ax.set_ylim(-5, 27)
    ax.legend(loc="upper right", frameon=False, fontsize=10)
    ax.text(4.56, -4.4, "2026H1 为半年值", fontsize=8.5, color=RED, ha="left")

    save(fig, f"600887_财务质量图_03_ROE三家对比_{OUT_REV}.png",
         "伊利、光明、新乳业：ROE 对比（2021A–2025A、2026H1）",
         "ROE ＝ 归母净利润 ÷ 期末归母净资产（三家统一口径）｜2026H1 为半年值，不与年度比较、不做年化")


if __name__ == "__main__":
    fig01()
    fig02()
    fig03()
    print("输出目录:", OUTDIR)
