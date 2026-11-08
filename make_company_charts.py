#!/usr/bin/env python3
"""生成第二章公司图：2026H1 分品类毛利率。"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "mplconfig_company"))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

OUTDIR = Path(__file__).resolve().parents[2] / "可视化图表" / "600887_公司图"
REV = "20261020_Rev01"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

NAVY, ORANGE, GREEN, GRAY = "#1F4E79", "#E67E22", "#2E7D32", "#8C8C8C"

labels = ["液体乳", "奶粉及奶制品", "冷饮"]
values = [34.16, 41.25, 39.06]
total = 36.43

fig, ax = plt.subplots(figsize=(9.0, 5.2), dpi=170)
bars = ax.bar(labels, values, width=0.55, color=[NAVY, ORANGE, GREEN], zorder=3)
ax.axhline(total, color=GRAY, linestyle="--", linewidth=1.2, zorder=2)
ax.text(2.42, total + 0.7, f"综合毛利率 {total:.2f}%", color=GRAY, fontsize=9,
        ha="right", va="bottom")

for bar, v in zip(bars, values):
    ax.text(bar.get_x() + bar.get_width() / 2, v + 0.7, f"{v:.2f}%",
            ha="center", va="bottom", fontsize=11, fontweight="bold", color=NAVY)

ax.set_ylim(0, 50)
ax.set_ylabel("毛利率（%）", fontsize=10)
ax.grid(axis="y", color="#E4E4E4", linewidth=0.8)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
ax.spines["left"].set_color("#BBBBBB")
ax.spines["bottom"].set_color("#BBBBBB")

fig.text(0.055, 0.015,
         "毛利率＝(营业收入−营业成本)÷营业收入。",
         fontsize=8, color=GRAY)

OUTDIR.mkdir(parents=True, exist_ok=True)
out = OUTDIR / f"600887_公司图_01_2026H1分品类毛利率_{REV}.png"
fig.savefig(out, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("saved:", out)
