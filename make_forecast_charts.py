#!/usr/bin/env python3
"""生成第四章《盈利预测》的正文配图（2 张，每张独立保存，研报风格）。

输出到 可视化图表/600887_盈利预测图/：
    01 营业收入与增速   02 营业成本与毛利率
    03–06（四项费用率／资本开支与折旧摊销／FCF 构成／归母净利润与净利率）
    保留为内部备用，不再默认生成。

数据来源：数值取自数据文件/DCF模型/600887_DCF模型_20261012_Rev07.xlsx 的《报表预测》页，
以常量固化在脚本内（本脚本不直接读 Excel，模型升版后需同步下方常量）。
历史 2021A-2025A 等于年报口径。2021A 同比的分母 2020A 营业收入 965.24 亿元
取自 07 清洗层利润表。

用法（需 bundled Python，含 matplotlib）：
    "<bundled python>" make_forecast_charts.py
    "<bundled python>" make_forecast_charts.py --rev 20261019_Rev00
"""

from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_fc"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

BASE = Path(__file__).resolve().parents[2]
OUTDIR = BASE / "可视化图表" / "600887_盈利预测图"
OUT_REV = "20261020_Rev01"          # 输出图版本号（收入序列用变量名 REV，见下）

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False
NAVY, ORANGE, RED, GREEN, GRAY, HIST, FORE = (
    "#1F4E79", "#E67E22", "#C00000", "#2E7D32", "#8C8C8C", "#8FAADC", "#E67E22")

YEARS = ["2021A", "2022A", "2023A", "2024A", "2025A",
         "2026E", "2027E", "2028E", "2029E", "2030E"]
X = np.arange(len(YEARS))
N_HIST = 5

REV = [1101.44, 1226.98, 1257.58, 1153.93, 1156.36, 1214.18, 1280.96, 1357.82, 1432.50, 1504.12]
REV_2020 = 965.24
COST = [764.17, 831.19, 847.89, 762.99, 756.54, 816.14, 861.03, 912.69, 962.89, 1011.03]
GM = [0.30621, 0.32258, 0.32578, 0.33879, 0.34576, 0.32782, 0.32782, 0.32782, 0.32782, 0.32782]
R_TAX = [0.006025, 0.006046, 0.005826, 0.006350, 0.005816, 0.006012, 0.006012, 0.006012, 0.006012, 0.006012]
R_SELL = [0.175360, 0.186704, 0.179483, 0.190499, 0.186313, 0.186313, 0.186313, 0.186313, 0.186313, 0.186313]
R_ADM = [0.038378, 0.043545, 0.040985, 0.038695, 0.040607, 0.040607, 0.040607, 0.040607, 0.040607, 0.040607]
R_RD = [0.005457, 0.006696, 0.006760, 0.007539, 0.007913, 0.007913, 0.007913, 0.007913, 0.007913, 0.007913]
CAPEX = [66.83, 66.46, 69.56, 39.78, 30.37, 60.71, 64.05, 67.89, 71.62, 75.21]
DA = [34.21, 40.68, 44.31, 46.12, 45.95, 38.43, 42.70, 47.23, 52.00, 57.02]
NI = [87.05, 94.31, 104.29, 84.53, 115.65, 77.98, 82.36, 87.40, 92.30, 96.99]
FCF_P = [175.13, 67.82, 74.36, 80.09, 85.94]
EBIAT_P = [79.21, 83.56, 88.58, 93.45, 98.12]
NWC_P = [118.19, 5.60, 6.45, 6.26, 6.01]      # -营运资本变动（正值＝现金释放）
CAPEX_P = CAPEX[5:]
DA_P = DA[5:]


def style(ax):
    ax.grid(axis="y", alpha=0.20, lw=0.6)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)


def head(fig, title, subtitle):
    fig.text(0.055, 0.955, subtitle, fontsize=9, color="#666666", ha="left", va="top")


def foot(fig, text):
    fig.text(0.055, 0.012, text, fontsize=8, color=GRAY, ha="left", va="bottom")


def save(fig, name):
    OUTDIR.mkdir(parents=True, exist_ok=True)
    out = OUTDIR / name
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("已保存", out)


def fig01(rev):
    fig, ax = plt.subplots(figsize=(11, 5.4))
    ax.bar(X, REV, width=0.6, color=[HIST] * N_HIST + [FORE] * 5, zorder=3)
    for xi, v in zip(X, REV):
        ax.text(xi, v - 78, f"{v:,.0f}", ha="center", fontsize=8.5, color="white")
    yoy = [(REV[0] / REV_2020 - 1)] + [REV[i] / REV[i - 1] - 1 for i in range(1, 10)]
    ax2 = ax.twinx()
    ax2.plot(X, [y * 100 for y in yoy], color=RED, marker="o", lw=2.0, ms=5.5, zorder=5,
             label="同比（右轴，%）")
    for xi, y in zip(X, yoy):
        ax2.text(xi, y * 100 + 1.4, f"{y*100:+.1f}", ha="center", fontsize=8, color=RED)
    ax2.axhline(0, color="#AAAAAA", lw=0.8, ls=":")
    ax2.set_ylim(-14, 24)
    ax2.set_ylabel("同比（%）", fontsize=9.5)
    ax.set_xticks(X, YEARS, fontsize=9)
    ax.set_ylabel("营业收入（亿元）", fontsize=9.5)
    ax.set_ylim(0, 1750)
    style(ax)
    head(fig, "图 4-1　营业收入与同比增速（2021A–2030E）",
         "浅蓝＝历史实际（2021A–2025A），橙色＝预测（2026E–2030E）；红线为同比增速（右轴）")
    foot(fig, "单位：亿元；预测期增速假设为 5.0% / 5.5% / 6.0% / 5.5% / 5.0%；2021A 同比分母为 2020A 营业收入 965.24 亿元。")
    fig.tight_layout(rect=(0.03, 0.06, 0.97, 0.90))
    save(fig, f"600887_盈利预测图_01_营业收入与增速_{rev}.png")


def fig02(rev):
    fig, ax = plt.subplots(figsize=(11, 5.4))
    ax.bar(X, COST, width=0.6, color=[HIST] * N_HIST + [FORE] * 5, zorder=3, label="营业成本（左轴）")
    for xi, v in zip(X, COST):
        ax.text(xi, v - 55, f"{v:,.0f}", ha="center", fontsize=8.5, color="white")
    ax.set_xticks(X, YEARS, fontsize=9)
    ax.set_ylabel("营业成本（亿元）", fontsize=9.5)
    ax.set_ylim(0, 1250)
    style(ax)
    ax2 = ax.twinx()
    ax2.plot(X, [g * 100 for g in GM], color=RED, marker="o", lw=2.0, ms=5.5, zorder=5,
             label="毛利率（右轴，%）")
    for xi, g in zip(X, GM):
        ax2.text(xi, g * 100 + 0.28, f"{g*100:.2f}", ha="center", fontsize=8, color=RED)
    ax2.axhline(32.7824, color=GREEN, lw=1.4, ls="--", zorder=4)
    ax2.text(0.02, 28.72, "绿虚线＝预测期假设 32.78%（＝2021–2025 五年均值）",
             fontsize=8.5, color=GREEN, va="bottom")
    ax2.set_ylim(28, 38)
    ax2.set_ylabel("毛利率（%）", fontsize=9.5)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=8.5, loc="upper left", framealpha=0.9)
    head(fig, "图 4-2　营业成本与毛利率（2021A–2030E）",
         "柱＝营业成本（左轴）；红线＝毛利率（右轴）；绿虚线＝预测期取值 32.78%")
    foot(fig, "毛利率 ＝ 1 − 营业成本 ÷ 营业收入，一律用利润表「营业成本」（财务摘要的「营业成本」实为营业总成本）。")
    fig.tight_layout(rect=(0.03, 0.06, 0.97, 0.90))
    save(fig, f"600887_盈利预测图_02_营业成本与毛利率_{rev}.png")


def fig03(rev):
    """备用图：四项费用率。默认不生成。"""
    fig, ax = plt.subplots(figsize=(11, 5.4))
    series = [("营业税金及附加率", R_TAX, "#8C8C8C"), ("销售费用率", R_SELL, NAVY),
              ("管理费用率", R_ADM, ORANGE), ("研发费用率", R_RD, GREEN)]
    for lab, s, c in series:
        ax.plot(X[:N_HIST], [v * 100 for v in s[:N_HIST]], color=c, marker="o", lw=2.0, ms=5,
                label=f"{lab}（实际）")
        ax.plot(X[N_HIST - 1:], [v * 100 for v in s[N_HIST - 1:]], color=c, marker="o", lw=2.0, ms=5,
                ls="--", label=f"{lab}（预测）")
        if lab != "营业税金及附加率":
            ax.text(9.15, s[-1] * 100, f" {s[-1]*100:.4f}%".rstrip("0").rstrip("."), fontsize=8.5,
                    color=c, va="center")
    ax.axvline(4.5, color="#BBBBBB", lw=1.0, ls=":")
    ax.text(4.62, 20.6, "预测期起点", fontsize=8.5, color=GRAY)
    ax.set_xticks(X, YEARS, fontsize=9)
    ax.set_ylabel("占营业收入比（%）", fontsize=9.5)
    ax.set_ylim(0, 22)
    ax.set_xlim(-0.5, 10.6)
    style(ax)
    ax.legend(fontsize=7.8, ncol=2, loc="center left", bbox_to_anchor=(0.015, 0.58), framealpha=0.9)
    head(fig, "图 4-3　四项费用率（2021A–2030E）",
         "实线＝历史实际；虚线＝预测（三项费用率取 2025A 实际值并全期不变，税金及附加率取五年均值）")
    foot(fig, "销售费用 18.6313%、管理费用 4.0607%、研发费用 0.7913%（均取 2025A 实际）、营业税金及附加率 0.6012%（五年均值）。")
    fig.tight_layout(rect=(0.03, 0.06, 0.97, 0.90))
    save(fig, f"600887_盈利预测图_03_费用率_{rev}.png")


def fig04(rev):
    """备用图：资本开支与折旧摊销。默认不生成。"""
    fig, axes = plt.subplots(1, 2, figsize=(12.2, 5.4), gridspec_kw={"wspace": 0.28})
    ax = axes[0]
    ax.bar(X - 0.19, CAPEX, width=0.36, color=[HIST] * N_HIST + [FORE] * 5, zorder=3, label="资本开支")
    ax.bar(X + 0.19, DA, width=0.36, color=[GREEN] * N_HIST + ["#7FBF7F"] * 5, zorder=3, label="折旧摊销")
    for xi, (a, b) in enumerate(zip(CAPEX, DA)):
        ax.text(xi - 0.19, a + 1.2, f"{a:.1f}", ha="center", fontsize=7.6, color=NAVY)
        ax.text(xi + 0.19, b + 1.2, f"{b:.1f}", ha="center", fontsize=7.6, color=GREEN)
    ax.set_xticks(X, YEARS, fontsize=8.5)
    ax.set_ylabel("亿元", fontsize=9.5)
    ax.set_ylim(0, 88)
    style(ax)
    ax.legend(fontsize=8.5, loc="upper left", framealpha=0.9)
    ax.set_title("A　资本开支与折旧摊销（亿元）", fontsize=10.5, color=NAVY)

    ax = axes[1]
    ratio = [c / r * 100 for c, r in zip(CAPEX, REV)]
    ax.bar(X, ratio, width=0.6, color=[HIST] * N_HIST + [FORE] * 5, zorder=3)
    for xi, v in zip(X, ratio):
        ax.text(xi, v + 0.10, f"{v:.2f}%", ha="center", fontsize=8, color=NAVY)
    ax.axhline(5.0, color=RED, lw=1.4, ls="--", zorder=4)
    ax.text(0.02, 5.12, "预测期假设 5%", fontsize=8.5, color=RED, va="bottom")
    ax.set_xticks(X, YEARS, fontsize=8.5)
    ax.set_ylabel("资本开支 ÷ 营业收入（%）", fontsize=9.5)
    ax.set_ylim(0, 7.2)
    style(ax)
    ax.set_title("B　资本开支占营业收入比（%）", fontsize=10.5, color=NAVY)
    head(fig, "图 4-4　资本开支与折旧摊销（2021A–2030E）",
         "A：绝对额（浅蓝＝历史、橙＝预测的资本开支；绿＝折旧摊销）；B：资本开支占营业收入比与 5% 假设")
    foot(fig, "预测期折旧摊销只含固定资产折旧（现有固定资产 ÷ 10 ＋ 各年资本开支 ÷ 15）；2025A 实际还含使用权资产摊销 1.52、无形资产摊销 3.58、"
              "长期待摊费用摊销 0.15 亿元。资本开支按营业收入 × 5% 假设（依据 2025 年报「2026 年计划投资 64.04 亿元」）。")
    fig.subplots_adjust(top=0.895, bottom=0.165, left=0.06, right=0.965)
    save(fig, f"600887_盈利预测图_04_资本开支与折旧摊销_{rev}.png")


def fig05(rev):
    """备用图：FCF 构成。默认不生成。"""
    lab = YEARS[N_HIST:]
    x = np.arange(5)
    e, d, n, c = np.array(EBIAT_P), np.array(DA_P), np.array(NWC_P), np.array(CAPEX_P)
    pos = e + d
    fig, axes = plt.subplots(1, 2, figsize=(12.6, 5.6),
                             gridspec_kw={"width_ratios": [1.5, 1], "wspace": 0.24})

    ax = axes[0]
    ax.bar(x, e, width=0.52, color=NAVY, zorder=3, label="EBIAT（＝EBIT×(1−25%)）")
    ax.bar(x, d, width=0.52, bottom=e, color="#7FBF7F", zorder=3, label="折旧摊销")
    ax.bar(x, n, width=0.52, bottom=pos, color=ORANGE, zorder=3, label="−营运资本变动（正值＝现金释放）")
    ax.bar(x, -c, width=0.52, color=RED, zorder=3, label="−资本开支")
    for xi, v in zip(x, e):
        ax.text(xi, v / 2, f"{v:.1f}", ha="center", fontsize=8.5, color="white")
    for xi, v in zip(x, d):
        ax.text(xi, e[xi] + v / 2, f"{v:.1f}", ha="center", fontsize=8.5, color="#1B4D1B")
    for xi, v in zip(x, n):
        ax.text(xi, pos[xi] + v / 2, f"{v:.1f}", ha="center", fontsize=8.5, color="#7A4A00")
    for xi, v in zip(x, c):
        ax.text(xi, -v / 2, f"{v:.1f}", ha="center", fontsize=8.5, color="white")
    for xi, v in zip(x, FCF_P):
        ax.text(xi, pos[xi] + n[xi] + 7, f"FCF {v:.2f}", ha="center", fontsize=9.5,
                color=RED, fontweight="bold")
    ax.axhline(0, color="#888888", lw=1.0)
    ax.set_xticks(x, lab, fontsize=9.5)
    ax.set_ylabel("亿元", fontsize=9.5)
    ax.set_ylim(-95, 228)
    style(ax)
    ax.legend(fontsize=8.2, ncol=2, loc="upper center", bbox_to_anchor=(0.5, -0.155),
              frameon=False)
    ax.set_title("A　FCF 的构成（亿元）", fontsize=10.5, color=NAVY, pad=8)

    ax = axes[1]
    ax.bar(x, FCF_P, width=0.55, color=[ORANGE] + [NAVY] * 4, zorder=3)
    for xi, v in zip(x, FCF_P):
        ax.text(xi, v + 3.5, f"{v:.2f}", ha="center", fontsize=9.5, color=NAVY)
    ax.set_xticks(x, lab, fontsize=9.5)
    ax.set_ylabel("自由现金流（亿元）", fontsize=9.5)
    ax.set_ylim(0, 205)
    style(ax)
    ax.set_title("B　自由现金流（亿元）", fontsize=10.5, color=NAVY, pad=8)
    ax.annotate("含营运资本释放 +118.19 亿元", xy=(0.30, 178), xytext=(1.35, 197),
                fontsize=8.5, color=RED, ha="left",
                arrowprops=dict(arrowstyle="->", color=RED, lw=0.9))

    head(fig, "图 4-5　自由现金流构成与走势（2026E–2030E）",
         "FCF ＝ EBIAT ＋ 折旧摊销 ＋（−营运资本变动）− 资本开支")
    foot(fig, "2026E 的营运资本变动为 −118.19 亿元（营运资本减少、释放现金），对 FCF 形成 +118.19 亿元的正贡献；"
              "2027E–2030E 收窄至 5.60/6.45/6.26/6.01 亿元。")
    fig.subplots_adjust(top=0.885, bottom=0.265, left=0.055, right=0.975)
    save(fig, f"600887_盈利预测图_05_FCF构成_{rev}.png")


def fig06(rev):
    """备用图：归母净利润与净利率。默认不生成。"""
    fig, ax = plt.subplots(figsize=(11, 5.4))
    ax.bar(X, NI, width=0.6, color=[HIST] * N_HIST + [FORE] * 5, zorder=3)
    for xi, v in zip(X, NI):
        ax.text(xi, v - 7.0, f"{v:,.2f}", ha="center", fontsize=8.5, color="white")
    ax.set_xticks(X, YEARS, fontsize=9)
    ax.set_ylabel("归母净利润（亿元）", fontsize=9.5)
    ax.set_ylim(0, 148)
    style(ax)
    ax2 = ax.twinx()
    m = [n / r * 100 for n, r in zip(NI, REV)]
    ax2.plot(X, m, color=RED, marker="o", lw=2.0, ms=5.5, zorder=5, label="归母净利率（右轴，%）")
    for xi, v in zip(X, m):
        ax2.text(xi, v + 0.30, f"{v:.2f}%", ha="center", fontsize=8, color=RED)
    ax2.set_ylim(4.5, 12.5)
    ax2.set_ylabel("归母净利率（%）", fontsize=9.5)
    head(fig, "图 4-6　归母净利润与归母净利率（2021A–2030E）",
         "柱＝归母净利润（左轴）；红线＝归母净利率（右轴）")
    foot(fig, "2026E 归母净利润 77.98 亿元＝2025A 115.65 亿元的 67.4%（−37.67 亿元）：利润总额 −26.51 亿元（三费 −13.58、财务费用净收益归零 −6.91、"
              "其他收益等五项归零 −3.87、毛利 −1.79、营业外净额 +3.16）＋ 所得税增加 11.06 亿元（有效税率 11.39% → 25.00%）。")
    fig.tight_layout(rect=(0.03, 0.06, 0.97, 0.90))
    save(fig, f"600887_盈利预测图_06_归母净利润与净利率_{rev}.png")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成第四章盈利预测正文配图（01/02 两张；03–06 备用）")
    ap.add_argument("--rev", default=OUT_REV, help="输出版本号，例如 20261019_Rev00")
    a = ap.parse_args()
    for fn in (fig01, fig02):
        fn(a.rev)
