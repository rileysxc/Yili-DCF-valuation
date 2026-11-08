#!/usr/bin/env python3
"""扫描原始数据层，生成字段级索引（数据字典的机器可读部分）。

用法：
    python build_field_index.py
python build_field_index.py --root "..\\..\\原始数据" --rev Rev01

产出：
    00_索引/字段索引_<YYYYMMDD>_<rev>.csv

列：一级目录 / 公司 / 文件 / 数据行数 / 字段序号 / 字段名 / 数据类型 / 空值数 / 空值率 / 样例值
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

import pandas as pd

DEFAULT_ROOT = Path(__file__).resolve().parents[2] / "原始数据"
SKIP_DIRS = {"00_索引", "05_手工整理数据"}

# 同名表往往有多份取数/清洗版本（_20261010_Rev00、_20261011_Rev05 …）。
# 索引只描述**当前层**，所以每张表只取最新一份：先比取数日、再比 Rev。
STAMP_PAT = re.compile(r"^(?P<base>.+?)_(?P<date>\d{8})_Rev(?P<rev>\d+)\.csv$")


def latest_only(csvs: list[Path], root: Path) -> list[Path]:
    """同一张表只保留 (取数日, Rev) 最大的那一份。"""
    best: dict[tuple[str, str], tuple[str, int, Path]] = {}
    others: list[Path] = []
    for csv in csvs:
        m = STAMP_PAT.match(csv.name)
        if not m:
            others.append(csv)
            continue
        rel = csv.relative_to(root)
        key = (rel.parent.as_posix(), m.group("base"))
        stamp = (m.group("date"), int(m.group("rev")))
        if key not in best or stamp > best[key][:2]:
            best[key] = (*stamp, csv)
    return sorted(others + [v[2] for v in best.values()])


def read_csv(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("无法解码")


def scan(root: Path) -> tuple[pd.DataFrame, list[str]]:
    rows, problems = [], []
    csvs = [p for p in root.rglob("*.csv")
            if p.relative_to(root).parts[0] not in SKIP_DIRS]
    for csv in latest_only(csvs, root):
        rel = csv.relative_to(root)
        parts = rel.parts
        layer = parts[0]
        company = parts[1] if len(parts) > 2 else "—"
        try:
            df = read_csv(csv)
        except Exception as exc:  # noqa: BLE001
            problems.append(f"{rel.as_posix()} 读取失败：{type(exc).__name__}: {exc}")
            continue
        n = len(df)
        for i, col in enumerate(df.columns, start=1):
            s = df[col]
            nulls = int(s.isna().sum())
            sample = ""
            non_null = s.dropna()
            if len(non_null):
                v = non_null.iloc[-1] if not isinstance(non_null.iloc[-1], float) else non_null.iloc[0]
                sample = str(v)[:40]
            rows.append({
                "一级目录": layer,
                "公司": company,
                "文件": rel.name,
                "数据行数": n,
                "字段序号": i,
                "字段名": col,
                "数据类型": str(s.dtype),
                "空值数": nulls,
                "空值率": round(nulls / n, 4) if n else 0,
                "样例值": sample,
            })
    return pd.DataFrame(rows), problems


def main() -> int:
    parser = argparse.ArgumentParser(description="生成字段级索引")
    parser.add_argument("--root", default=str(DEFAULT_ROOT), help="原始数据目录")
    parser.add_argument("--rev", default="Rev00", help="版本号，默认 Rev00")
    parser.add_argument("--stamp", default=None, help="文件名日期，默认今天 YYYYMMDD")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not root.exists():
        print(f"[ERROR] 目录不存在：{root}", file=sys.stderr)
        return 2

    index, problems = scan(root)
    if index.empty:
        print("[ERROR] 未扫描到任何 CSV", file=sys.stderr)
        return 3

    stamp = args.stamp or dt.date.today().strftime("%Y%m%d")
    out_dir = root / "00_索引"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"字段索引_{stamp}_{args.rev}.csv"
    if out.exists():
        print(f"[ERROR] 已存在，未覆盖：{out.name}（提高 --rev 重跑）", file=sys.stderr)
        return 4
    index.to_csv(out, index=False, encoding="utf-8-sig")

    print(f"扫描目录：{root}")
    print(f"文件数  ：{index['文件'].nunique()}")
    print(f"字段数  ：{len(index)}")
    print(f"产出    ：{out}")
    print("\n按数据层统计：")
    g = index.groupby("一级目录").agg(文件数=("文件", "nunique"), 字段数=("字段名", "size"))
    print(g.to_string())
    for msg in problems:
        print(f"  [WARN] {msg}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
