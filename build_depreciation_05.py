# -*- coding: utf-8 -*-
"""从年报/半年报 PDF 抽取「现金流量表补充资料」的折旧摊销，落 05 层 CSV。

口径与 600887_折旧摊销_*.csv 完全一致：
  首列  固定资产折旧及油气资产折耗等  = 「固定资产折旧、油气资产折耗、生产性生物资产折旧」合并行，
                                        若年报分行列示则相加（固定资产折旧 ＋ 生产性生物资产折旧 ＋ 油气资产折耗）
  其余  使用权资产摊销/折旧、无形资产摊销、长期待摊费用摊销
  ⚠ 不含「股权激励费用摊销」（非资产折旧摊销）

每份年报的补充资料含「本期／上期」两列，因此：
  2021 年报→2021A(本期)+2020A(上期)；2022–2025 年报同理；2026 半年报→2026H1(本期)+2025H1(上期)

用法：python build_depreciation_05.py [--write]
      需用 bundled Python（依赖 pypdf）。
"""
import argparse
import csv
import os
import re
import sys
from pathlib import Path

from pypdf import PdfReader

sys.stdout.reconfigure(encoding="utf-8")

ROOT = str(Path(__file__).resolve().parents[2] / "原始资料")
OUT = str(Path(__file__).resolve().parents[2] / "原始数据" / "05_手工整理数据")

COMBINED = "固定资产折旧、油气资产折耗、生产性生物资产折旧"
LABELS = {
    "fixed": COMBINED,
    "fixed_alt": "固定资产折旧",
    "bio": "生产性生物资产折旧",
    "oil": "油气资产折耗",
    "rou": ["使用权资产摊销", "使用权资产折旧"],
    "intang": "无形资产摊销",
    "lt_prepaid": "长期待摊费用摊销",
}

COMPANIES = {
    "600597": ("光明乳业", "光明乳业年报"),
    "002946": ("新乳业", "新乳业年报"),
}

# (报告期, 文件名, 用本期还是上期)
PERIODS = []
for y in range(2021, 2026):
    PERIODS.append((f"{y}-12-31", f"{y}年年度报告.pdf", "本期"))
PERIODS.append(("2020-12-31", "2021年年度报告.pdf", "上期"))
PERIODS.append(("2026-06-30", "2026年半年报告.pdf", "本期"))
PERIODS.append(("2025-06-30", "2026年半年报告.pdf", "上期"))

NUM = re.compile(r"\(?-?[\d,]+(?:\.\d+)?\)?")


def to_num(s):
    s = s.strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "")
    if s in ("", "-", "--"):
        return 0.0
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def parse_supplement(page_text):
    """截出「将净利润调节为经营活动现金流量」到「经营活动产生的现金流量净额」之间的段落。

    注意：**不能**把行内空白压掉——本期/上期两列就是靠空格分隔的，压掉会串列。
    """
    if "将净利润调节为经营活动现金流量" not in page_text:
        return None
    seg = page_text.split("将净利润调节为经营活动现金流量", 1)[1]
    return seg.split("经营活动产生的现金流量净额", 1)[0]


def _pair(seg, label):
    """在段落里找 label 之后的两个金额（本期、上期）。标签可能跨行，故逐字允许空白。"""
    pat = "".join(re.escape(ch) + r"\s*" for ch in label)
    for m in re.finditer(pat, seg):
        lines = seg[m.end():].split("\n")
        nums = NUM.findall(lines[0])
        if not nums:                       # 金额换到下一行（表头接行的情况）
            for extra in lines[1:3]:
                if re.search(r"[\u4e00-\u9fff]", extra):
                    break
                nums = NUM.findall(extra)
                if nums:
                    break
        if nums:
            vals = [to_num(n) for n in nums[:2]]
            while len(vals) < 2:           # 上期为空时补 0
                vals.append(0.0)
            return vals
    return None


def components(seg, idx):
    """返回 (固定资产折旧及油气资产折耗等, 使用权资产摊销, 无形资产摊销, 长期待摊费用摊销, 构成说明)。"""
    fixed = _pair(seg, COMBINED)
    src = COMBINED
    if fixed is None:
        fixed = [0.0, 0.0]
        used = []
        for lab in (LABELS["fixed_alt"], LABELS["bio"], LABELS["oil"]):
            v = _pair(seg, lab)
            if v:
                fixed = [a + b for a, b in zip(fixed, v)]
                used.append(lab)
        src = "＋".join(used)
    rou, src_rou = None, ""
    for lab in LABELS["rou"]:
        rou = _pair(seg, lab)
        if rou:
            src_rou = lab
            break
    rou = rou or [0.0, 0.0]
    itg = _pair(seg, LABELS["intang"]) or [0.0, 0.0]
    ltp = _pair(seg, LABELS["lt_prepaid"]) or [0.0, 0.0]
    return fixed[idx], rou[idx], itg[idx], ltp[idx], src, src_rou


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--selftest", action="store_true",
                    help="用同一解析器跑伊利，与既有 600887_折旧摊销_*.csv 对账")
    args = ap.parse_args()

    if args.selftest:
        import pandas as pd
        ref = None
        for f in sorted(os.listdir(OUT)):
            if f.startswith("600887_折旧摊销_") and f.endswith(".csv"):
                ref = f
        ref = pd.read_csv(os.path.join(OUT, ref), encoding="utf-8-sig")
        ref["报告期"] = pd.to_datetime(ref["报告期"]).dt.strftime("%Y-%m-%d")
        mine = {}
        cache = {}
        for period, fname, which in PERIODS:
            pdf = os.path.join(ROOT, "伊利年报", fname)
            if pdf not in cache:
                reader = PdfReader(pdf)
                for i, pg in enumerate(reader.pages):
                    seg = parse_supplement(pg.extract_text() or "")
                    if seg:
                        cache[pdf] = (seg, i + 1)
                        break
            seg, page = cache[pdf]
            idx = 0 if which == "本期" else 1
            total = sum(components(seg, idx)[:4])
            mine[period] = round(total, 2)
        print("① 解析器自测（伊利，与既有 05 层 CSV 对账）")
        for _, r in ref.iterrows():
            p = r["报告期"]
            ok = abs((mine.get(p) or 0) - round(r["折旧摊销合计"], 2)) < 0.05
            print(f"   {p}  解析 {mine.get(p):>17.2f} | CSV {round(r['折旧摊销合计'], 2):>17.2f} | "
                  f"{'一致' if ok else '不一致'}")

    for code, (name, folder) in COMPANIES.items():
        rows = []
        cache = {}
        for period, fname, which in PERIODS:
            pdf = os.path.join(ROOT, folder, fname)
            if pdf not in cache:
                reader = PdfReader(pdf)
                found = (None, None)
                for i, pg in enumerate(reader.pages):
                    seg = parse_supplement(pg.extract_text() or "")
                    if seg:
                        found = (seg, i + 1)
                        break
                cache[pdf] = found
            seg, page = cache[pdf]
            if seg is None:
                print(f"  !! {name} {fname} 未找到补充资料")
                continue
            idx = 0 if which == "本期" else 1
            fixed, rou, itg, ltp, fixed_src, rou_src = components(seg, idx)
            total = fixed + rou + itg + ltp
            note = f"{fname[:-4]} PDF p.{page}（{which}）"
            if fixed_src != COMBINED:
                note += f"；首列＝{fixed_src}"
            if rou_src and rou_src != "使用权资产摊销":
                note += f"；{rou_src}"
            rows.append([period, round(fixed, 2), round(rou, 2), round(itg, 2),
                         round(ltp, 2), round(total, 2), note])
            print(f"  {name} {period} 合计={total / 1e8:>8.4f} 亿  {note}")
        rows.sort(key=lambda r: r[0])
        path = os.path.join(OUT, f"{code}_折旧摊销_20261015_Rev00.csv")
        if args.write:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["报告期", "固定资产折旧及油气资产折耗等", "使用权资产摊销",
                            "无形资产摊销", "长期待摊费用摊销", "折旧摊销合计", "来源"])
                w.writerows(rows)
            print(f"  已落盘 {path}\n")


if __name__ == "__main__":
    main()
