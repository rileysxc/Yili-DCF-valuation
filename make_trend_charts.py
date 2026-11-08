#!/usr/bin/env python3
"""从三家基础趋势表读数据，生成「伊利为主 + 同业对比」的单面板趋势图，每张图独立保存。

输出 5 张 PNG：
  01 营业收入 / 02 归母净利润 / 03 增速 / 04 毛利率·归母净利率·现金流覆盖 / 05 ROE

口径（与《基础趋势表》口径说明一致）：
  营业收入＝利润表「营业收入」（不含利息收入）；归母净利率＝归母净利润 ÷ 营业收入；
  ROE＝归母净利润 ÷ 期末归母净资产；经营现金流覆盖＝经营活动现金流量净额 ÷ 归母净利润。
  2026H1 为半年累计，绝对额与年度不可比，ROE 亦为半年值、不做年化。
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

# matplotlib 默认缓存目录在部分环境不可写，先指到临时目录（须在 import matplotlib 之前）
os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_trend"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl

BASE = Path(__file__).resolve().parents[2] / "数据文件" / "基础趋势表"
OUTDIR = Path(__file__).resolve().parents[2] / "可视化图表" / "600887_基础趋势图"
REV = "20261020_Rev01"

FILES = {
    "伊利": BASE / "600887_伊利股份_基础趋势表_20261011_Rev01.xlsx",
    "光明": BASE / "600597_光明乳业_基础趋势表_20261011_Rev00.xlsx",
    "新乳业": BASE / "002946_新乳业_基础趋势表_20261011_Rev00.xlsx",
}

# 基础趋势表列：C=2021-12-31 D=2022 E=2023 F=2024 H=2025-12-31 I=2026-06-30
COLS = ["C", "D", "E", "F", "H", "I"]
PREV = {"C": "B", "D": "C", "E": "D", "F": "E", "H": "F", "I": "G"}
LABELS = ["2021A", "2022A", "2023A", "2024A", "2025A", "2026H1"]

COLOR = {"伊利": "#C00000", "光明": "#7F8C8D", "新乳业": "#2E7D32"}
LW = {"伊利": 3.0, "光明": 1.8, "新乳业": 1.8}
MARKER = {"伊利": "o", "光明": "s", "新乳业": "^"}

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def ci(letter: str) -> int:
    return openpyxl.utils.column_index_from_string(letter)


def read_company(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    lp, cf, bs = wb["利润表"], wb["现金流量表"], wb["资产负债表"]
    n = len(COLS)
    rev = np.array([lp.cell(3, ci(c)).value for c in COLS]) / 1e8
    cogs = np.array([lp.cell(10, ci(c)).value for c in COLS]) / 1e8
    ni = np.array([lp.cell(51, ci(c)).value for c in COLS]) / 1e8
    ocf = np.array([cf.cell(27, ci(c)).value for c in COLS]) / 1e8
    eq = np.array([bs.cell(138, ci(c)).value for c in COLS]) / 1e8
    rev_base = np.array([lp.cell(3, ci(PREV[c])).value for c in COLS]) / 1e8
    ni_base = np.array([lp.cell(51, ci(PREV[c])).value for c in COLS]) / 1e8
    return dict(
        rev=rev,
        ni=ni,
        gm=(rev - cogs) / rev * 100,
        nm=ni / rev * 100,
        ocf_ni=np.divide(ocf, ni),
        roe=ni / eq * 100,
        rev_yoy=(rev / rev_base - 1) * 100,
        ni_yoy=(ni / ni_base - 1) * 100,
    )


def style_ax(ax):
    ax.grid(axis="y", color="#DDDDDD", linewidth=0.7)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BBBBBB")


def mark_half(ax, ymax, x=5):
    ax.axvline(4.5, color="#C00000", linestyle=(0, (4, 3)), linewidth=1.1, zorder=1)
    ax.text(5.0, ymax, "半年数", ha="center", va="top", fontsize=9.5, color="#C00000")


def finish(fig, ax, out_name: str, title: str, subtitle: str):
    ax.set_xticks(range(6))
    ax.set_xticklabels(LABELS, fontsize=10)
    fig.text(0.075, 0.95, subtitle, fontsize=9.5, color="#555555", ha="left")
    fig.subplots_adjust(top=0.88, bottom=0.10, left=0.075, right=0.90)
    out = OUTDIR / out_name
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存:", out.name)


def peer_tail(ax, x, y, color, dy=6):
    """只标同业最后一个点，避免图面拥挤"""
    i = int(np.asarray(x)[-1])
    ax.annotate(f"{y[i]:.2f}", (i, y[i]), textcoords="offset points", xytext=(0, dy),
                ha="center", fontsize=9, color=color)


def fig_amount(data, key, title, fname, unit="亿元", note_half=True):
    fig, ax = plt.subplots(figsize=(11, 6.4), dpi=150)
    style_ax(ax)
    x = np.arange(6)
    vals = data["伊利"][key]
    bars = ax.bar(x, vals, width=0.55, color="#2E5C8A", zorder=3)
    bars[5].set_color("#9DC3E6")
    bars[5].set_edgecolor("#2E5C8A")
    bars[5].set_hatch("//")
    labs = ax.bar_label(bars, labels=[f"{v:,.1f}" for v in vals], label_type="center",
                        fontsize=10, color="white", fontweight="bold")
    labs[5].set_color("#1B3A5C")
    top = max(vals) * 1.20
    ax.set_ylim(0, top)
    if note_half:
        mark_half(ax, top * 0.97)
    ax.set_ylabel(f"伊利股份（{unit}）", fontsize=10)

    ax2 = ax.twinx()
    ax2.set_axisbelow(True)
    for name in ("光明", "新乳业"):
        y = data[name][key]
        ax2.plot(x, y, color=COLOR[name], linewidth=LW[name], marker=MARKER[name],
                 markersize=6, label=f"{name}（右轴）")
        peer_tail(ax2, x, y, COLOR[name])
    ax2.set_ylabel(f"光明 / 新乳业（{unit}）", fontsize=10, color="#555555")
    ax2.tick_params(axis="y", colors="#555555")
    ax2.spines["top"].set_visible(False)
    ax2.legend(loc="upper left", fontsize=9.5, frameon=False, ncol=1)

    subtitle = ("柱＝伊利（左轴），线＝光明/新乳业（右轴）"
                "　|　2026H1 为半年累计，绝对额不可与年度比较")
    finish(fig, ax, fname, title, subtitle)


def fig_growth(data):
    fig, ax = plt.subplots(figsize=(11, 6.4), dpi=150)
    style_ax(ax)
    x = np.arange(6)
    ax.axhline(0, color="#999999", linewidth=0.9, linestyle=":")
    handles = []
    for name in ("伊利", "光明", "新乳业"):
        for key, ls, mk in (("rev_yoy", "-", "o"), ("ni_yoy", "--", "s")):
            y = data[name][key]
            ln, = ax.plot(x, y, color=COLOR[name], linewidth=LW[name], linestyle=ls,
                          marker=mk, markersize=5.5 if name == "伊利" else 4.5, alpha=1.0 if name == "伊利" else 0.85)
            if name == "伊利":
                for i in x:
                    ax.annotate(f"{y[i]:+.1f}%", (i, y[i]), textcoords="offset points",
                                xytext=(0, 10 if ls == "-" else -18), ha="center",
                                fontsize=9, color=COLOR["伊利"])
            handles.append((ln, f"{name} 营收增速" if ls == "-" else f"{name} 归母净利增速"))
    # 纵轴按三家全部数据自适应，避免把光大/新乳业的极端值截到图外
    allv = np.concatenate([data[n][k] for n in data for k in ("rev_yoy", "ni_yoy")])
    lo, hi = float(allv.min()), float(allv.max())
    span = hi - lo
    ax.set_ylim(lo - span * 0.08, hi + span * 0.12)
    ax.set_ylabel("同比增速（%）", fontsize=10)
    ax.legend([h[0] for h in handles], [h[1] for h in handles], loc="upper left",
              fontsize=9, frameon=False, ncol=2)
    subtitle = ("颜色＝公司（伊利加粗），实线＝营收增速，虚线＝归母净利增速"
                "　|　2026H1 对比 2025H1")
    finish(fig, ax, f"600887_基础趋势图_03_增速_{REV}.png", "营收增速与归母净利增速（同比，三家对比）", subtitle)
    print("  伊利 营收增速:", np.round(data["伊利"]["rev_yoy"], 2).tolist())
    print("  伊利 净利增速:", np.round(data["伊利"]["ni_yoy"], 2).tolist())


def fig_ratio(data):
    fig, ax = plt.subplots(figsize=(11, 6.4), dpi=150)
    style_ax(ax)
    x = np.arange(6)
    handles = []
    for name in ("伊利", "光明", "新乳业"):
        for key, ls, mk, tag in (("gm", "-", "s", "毛利率"), ("nm", "--", "^", "归母净利率")):
            y = data[name][key]
            ln, = ax.plot(x, y, color=COLOR[name], linewidth=LW[name], linestyle=ls,
                          marker=mk, markersize=5.5 if name == "伊利" else 4.5,
                          alpha=1.0 if name == "伊利" else 0.85)
            if name == "伊利":
                for i in x:
                    ax.annotate(f"{y[i]:.2f}%", (i, y[i]), textcoords="offset points",
                                xytext=(0, 10 if ls == "-" else -18), ha="center",
                                fontsize=9, color=COLOR["伊利"])
            handles.append((ln, f"{name} {tag}"))
    allgm = np.concatenate([data[n]["gm"] for n in data] + [data[n]["nm"] for n in data])
    ax.set_ylim(min(0.0, float(allgm.min()) - 1.0), float(allgm.max()) * 1.18)
    ax.set_ylabel("毛利率 / 归母净利率（%）", fontsize=10)
    ax2 = ax.twinx()
    y = data["伊利"]["ocf_ni"]
    bars = ax2.bar(x, y, width=0.26, color="#E67E22", alpha=0.28, zorder=0)
    ax2.set_ylim(0, 5.0)
    ax2.set_ylabel("经营现金流 / 归母净利润（倍）", fontsize=10, color="#B35C00")
    ax2.tick_params(axis="y", colors="#B35C00")
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_color("#B35C00")
    for i in x:
        ax2.annotate(f"{y[i]:.2f}", (i, y[i]), textcoords="offset points", xytext=(0, 4),
                     ha="center", fontsize=8.5, color="#B35C00")
    handles.append((bars, "伊利 经营现金流/归母净利润（右轴）"))
    ax.legend([h[0] for h in handles], [h[1] for h in handles], loc="upper left",
              fontsize=9, frameon=False, ncol=2)
    subtitle = ("颜色＝公司（伊利加粗），实线＝毛利率，虚线＝归母净利率，柱＝伊利现金流覆盖（右轴）"
                "　|　2026H1 为半年累计，比率可跨期比较")
    finish(fig, ax, f"600887_基础趋势图_04_毛利率净利率现金流_{REV}.png",
           "毛利率、归母净利率与经营现金流覆盖（三家对比）", subtitle)


def fig_roe(data):
    fig, ax = plt.subplots(figsize=(11, 6.4), dpi=150)
    style_ax(ax)
    x = np.arange(6)
    ax.axhline(0, color="#999999", linewidth=0.9, linestyle=":")
    for name in ("光明", "新乳业"):
        y = data[name]["roe"]
        ax.plot(x, y, color=COLOR[name], linewidth=LW[name], marker=MARKER[name],
                markersize=6, label=f"{name}")
        peer_tail(ax, x, y, COLOR[name])
    y = data["伊利"]["roe"]
    ax.plot(x, y, color=COLOR["伊利"], linewidth=LW["伊利"], marker="o", markersize=7,
            label="伊利", zorder=5)
    ax.plot(x, y, color=COLOR["伊利"], linewidth=0, marker="o", markersize=7, zorder=6)
    for i in x:
        ax.annotate(f"{y[i]:.2f}%", (i, y[i]), textcoords="offset points", xytext=(0, -20 if i in (3, 5) else 12),
                    ha="center", fontsize=9.5, color=COLOR["伊利"], fontweight="bold")
    top = max(max(data[n]["roe"]) for n in data) * 1.30
    ax.set_ylim(min(0, min(min(data[n]["roe"]) for n in data) * 1.4), top)
    ax.set_ylabel("ROE（%），归母净利润 ÷ 期末归母净资产", fontsize=10)
    ax.legend(loc="upper left", fontsize=10, frameon=False, ncol=3)
    subtitle = ("口径：归母净利润 ÷ 期末归母净资产（期末口径，非平均口径）"
                "　|　2026H1 为半年值，不与年度 ROE 比较、不做年化")
    finish(fig, ax, f"600887_基础趋势图_05_ROE_{REV}.png", "ROE（三家对比）", subtitle)
    for name in data:
        print("  %s ROE:" % name, np.round(data[name]["roe"], 2).tolist())


def main() -> None:
    data = {name: read_company(path) for name, path in FILES.items()}
    print("== 生成单面板图表（伊利为主 + 同业对比）==")
    fig_amount(data, "rev", "营业收入（三家对比）",
               f"600887_基础趋势图_01_营业收入_{REV}.png")
    fig_amount(data, "ni", "归母净利润（三家对比）",
               f"600887_基础趋势图_02_归母净利润_{REV}.png")
    fig_growth(data)
    fig_ratio(data)
    fig_roe(data)
    print("\n== 伊利关键数据 ==")
    for k in ("rev", "ni", "gm", "nm", "ocf_ni", "roe"):
        print("  %-7s %s" % (k, np.round(data["伊利"][k], 2).tolist()))


if __name__ == "__main__":
    main()
