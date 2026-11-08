# -*- coding: utf-8 -*-
"""构建 07_清洗后数据 层。

对三家（伊利 600887 / 光明 600597 / 新乳业 002946）的时间序列文件做清洗，
统一输出到 原始数据/07_清洗后数据/<代码>_<简称>/，作为下游唯一消费层。

已完成（2026-10-11 轮次）
    1. 日期统一：首列一律 date + ISO 字符串 YYYY-MM-DD（伊利原本就是这个格式）
    2. 市场类窗口对齐：行情、估值 裁到 2021-01-01 ~ 2026-10-09，时间升序
    3. 伊利镜像：伊利各表在 01/02/03 层已经合规，原样镜像进 07
    4. PE/PB 序列过滤：估值十年PE、估值十年PB 只保留数值 > 0 的行
       （亏损期 PE 为负，负 PE 没有比较意义）
    5. 报表类收口：三表、财务指标 筛选到 8 个报告期 + 时间升序
    6. 财务摘要裁列：只保留「选项、指标」+ 8 个报告期列
    7. PE/PB 数值列名统一：伊利原本叫 value，改成 pe_ttm / pb

报告期（8 期）
    2020-12-31 年度基数，2021~2025 五个年报，2025-06-30 中期基数，2026-06-30

上游
    原始数据/01_报表原表/<代码>_<简称>/
    原始数据/02_行情数据/<代码>_<简称>/
    原始数据/03_估值数据/<代码>_<简称>/
下游
    原始数据/07_清洗后数据/<代码>_<简称>/<代码>_<表名>_<运行日>_Rev<nn>.csv
    原始数据/07_清洗后数据/00_清洗日志_<运行日>_Rev<nn>.csv

用法
    python build_clean_layer.py            # 输出 Rev00
    python build_clean_layer.py --rev 2    # 输出 Rev02（旧版保留）
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parents[2] / "原始数据"
LAYER_REPORT = BASE / "01_报表原表"
LAYER_QUOTE = BASE / "02_行情数据"
LAYER_VALUATION = BASE / "03_估值数据"
LAYER_CLEAN = BASE / "07_清洗后数据"

RUN_DATE = "20261011"

COMPANIES = {
    "600887": "600887_伊利股份",
    "600597": "600597_光明乳业",
    "002946": "002946_新乳业",
}

TARGET_COL = "date"
RENAME_MAP = {"报告日": TARGET_COL, "日期": TARGET_COL, "数据日期": TARGET_COL}

# 市场类：按报告统一窗口裁剪并升序；报表类此轮只统一日期，报告期筛选留到下一轮
MARKET_TABLES = {"行情", "估值"}
WINDOW_START = pd.Timestamp("2021-01-01")
WINDOW_END = pd.Timestamp("2026-10-09")

# PE/PB 序列：不做窗口裁剪（裁了就失去长历史），只保留数值 > 0 的行，并统一数值列名
PE_PB_TABLES = {"估值十年PE": "pe_ttm", "估值十年PB": "pb"}

PERIODS = ["20201231", "20211231", "20221231", "20231231",
           "20241231", "20250630", "20251231", "20260630"]
PERIODS_ISO = [f"{p[:4]}-{p[4:6]}-{p[6:]}" for p in PERIODS]

SUMMARY_PREFIX = "财务摘要"
SKIP_TABLES = {"财务摘要_全量"}   # 全量归档，不属于可比口径，01 层保留

FILE_PATTERN = re.compile(r"^(\d{6})_(.+?)_(\d{8})_Rev(\d+)\.csv$")
ISO_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def to_iso(series: pd.Series) -> pd.Series:
    """把 20260630 / 2026-06-30 一律转成 2026-06-30。"""
    digits = series.astype(str).str.strip().str.replace(r"[^0-9]", "", regex=True)
    parsed = pd.to_datetime(digits, format="%Y%m%d", errors="coerce")
    if parsed.isna().any():
        bad = series[parsed.isna()].unique()[:5].tolist()
        raise ValueError(f"存在无法解析的日期：{bad}")
    return parsed.dt.strftime("%Y-%m-%d")


def collect_inputs() -> list[Path]:
    """同一公司同一张表可能有多份取数版本，只挑最新的一份（先比取数日，再比 Rev）。"""
    latest: dict[tuple[str, str], tuple[str, int, Path]] = {}
    for folder in (LAYER_REPORT, LAYER_QUOTE, LAYER_VALUATION):
        for code, name in COMPANIES.items():
            sub = folder / name
            if not sub.is_dir():
                continue
            for path in sorted(sub.glob(f"{code}_*.csv")):
                match = FILE_PATTERN.match(path.name)
                if not match:
                    continue
                key = (code, match.group(2))
                stamp = (match.group(3), int(match.group(4)))
                if key not in latest or stamp > latest[key][:2]:
                    latest[key] = (*stamp, path)
    return sorted(entry[2] for entry in latest.values())


def table_name(path: Path) -> str:
    match = FILE_PATTERN.match(path.name)
    return match.group(2) if match else path.stem


def finish(df: pd.DataFrame, table: str, code: str, rev: str, record: dict) -> dict:
    """写出到 07 层，再把文件读回来复核一遍。"""
    record["输出行数"] = len(df)
    out_dir = LAYER_CLEAN / COMPANIES[code]
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{code}_{table}_{RUN_DATE}_Rev{rev}.csv"
    df.to_csv(out_path, index=False, encoding="utf-8-sig")

    reloaded = pd.read_csv(out_path, encoding="utf-8-sig")
    if list(reloaded.columns) != list(df.columns):
        raise AssertionError(f"{out_path.name}: 回读列名与写出前不一致")
    if len(reloaded) != len(df):
        raise AssertionError(f"{out_path.name}: 回读行数与写出前不一致")
    if TARGET_COL in df.columns:
        if not reloaded[TARGET_COL].astype(str).str.match(ISO_PATTERN).all():
            raise AssertionError(f"{out_path.name}: 回读首列不是 YYYY-MM-DD")

    record["输出文件"] = out_path.name
    return record


def process(path: Path, rev: str) -> dict:
    code = path.name[:6]
    table = table_name(path)
    df = pd.read_csv(path, encoding="utf-8-sig")
    src_col = df.columns[0]
    record = {
        "代码": code,
        "简称": COMPANIES[code],
        "表名": table,
        "源文件": path.name,
        "源首列名": src_col,
        "源行数": len(df),
        "输出行数": "",
        "校验": "",
        "状态": "",
        "输出文件": "",
    }

    if table in SKIP_TABLES:
        record["状态"] = "跳过（全量归档，不属于可比口径；01 层保留）"
        return record

    if table.startswith(SUMMARY_PREFIX):
        labels = list(df.columns[:2])
        keep = labels + [p for p in PERIODS if p in df.columns]
        missing = [p for p in PERIODS if p not in df.columns]
        df = df[keep].copy()
        record["校验"] = "列裁剪" + (f"；缺列 {missing}" if missing else "；8 期齐全")
        record["状态"] = f"财务摘要裁列：{len(keep)} 列（选项/指标 + 8 期，升序）"
        return finish(df, table, code, rev, record)

    renamed = src_col in RENAME_MAP
    if renamed:
        iso = to_iso(df[src_col])
        if not (len(iso) == len(df)):
            raise AssertionError(f"{path.name}: 行数不一致")
        if len(set(df[src_col].astype(str))) != len(set(iso)):
            raise AssertionError(f"{path.name}: 唯一期数发生变化")
        if df[src_col].notna().sum() != iso.notna().sum():
            raise AssertionError(f"{path.name}: 转换后出现空值")
        df = df.assign(**{src_col: iso}).rename(columns={src_col: TARGET_COL})
    elif src_col != TARGET_COL or not df[src_col].astype(str).str.match(ISO_PATTERN).all():
        record["状态"] = f"跳过（首列 {src_col} 既非日期列，也不符合目标格式）"
        return record

    if not df[TARGET_COL].astype(str).str.match(ISO_PATTERN).all():
        raise AssertionError(f"{path.name}: 首列不是 YYYY-MM-DD")
    record["校验"] = "行数/唯一期数/空值/日期格式 通过"

    if table in MARKET_TABLES:
        ts = pd.to_datetime(df[TARGET_COL])
        before = len(df)
        df = (
            df[(ts >= WINDOW_START) & (ts <= WINDOW_END)]
            .sort_values(TARGET_COL)
            .reset_index(drop=True)
        )
        if df.empty:
            raise AssertionError(f"{path.name}: 窗口内没有数据")
        record["状态"] = (
            f"市场类：窗口对齐 {WINDOW_START:%Y-%m-%d}~{WINDOW_END:%Y-%m-%d}"
            f"（{before} -> {len(df)} 行，升序）"
        )
    elif table in PE_PB_TABLES:
        target_value_col = PE_PB_TABLES[table]
        value_col = df.columns[-1]
        note = ""
        if value_col != target_value_col:
            df = df.rename(columns={value_col: target_value_col})
            note = f"列名 {value_col} -> {target_value_col}；"
        values = pd.to_numeric(df[target_value_col], errors="coerce")
        before = len(df)
        df = df[values > 0].sort_values(TARGET_COL).reset_index(drop=True)
        if df.empty:
            raise AssertionError(f"{path.name}: 过滤后没有数据")
        if not (pd.to_numeric(df[target_value_col], errors="coerce") > 0).all():
            raise AssertionError(f"{path.name}: 过滤后仍存在非正值")
        record["状态"] = (
            f"PE/PB：{note}只留 {target_value_col} > 0"
            f"（{before} -> {len(df)} 行，剔除 {before - len(df)} 行非正值，升序）"
        )
    else:
        before = len(df)
        df = (
            df[df[TARGET_COL].isin(PERIODS_ISO)]
            .sort_values(TARGET_COL)
            .reset_index(drop=True)
        )
        if df.empty:
            raise AssertionError(f"{path.name}: 8 个报告期内没有数据")
        missing = [p for p in PERIODS_ISO if p not in set(df[TARGET_COL])]
        head = "已统一首列 %s -> date" % src_col if renamed else "镜像（首列已符合目标格式）"
        record["状态"] = f"{head}；筛选到 8 个报告期（{before} -> {len(df)} 行，升序）"
        if missing:
            record["校验"] = f"{record['校验']}；缺报告期 {missing}"

    return finish(df, table, code, rev, record)


def main() -> None:
    parser = argparse.ArgumentParser(description="构建 07_清洗后数据 层")
    parser.add_argument("--rev", type=int, default=0, help="输出文件版本号，默认 0")
    args = parser.parse_args()
    rev = f"{args.rev:02d}"

    LAYER_CLEAN.mkdir(parents=True, exist_ok=True)
    records = [process(path, rev) for path in collect_inputs()]

    log = pd.DataFrame(records)
    log_path = LAYER_CLEAN / f"00_清洗日志_{RUN_DATE}_Rev{rev}.csv"
    log.to_csv(log_path, index=False, encoding="utf-8-sig")

    written = int((log["输出文件"] != "").sum())
    print(f"inputs={len(records)} written={written}")
    print(f"log={log_path.name}")


if __name__ == "__main__":
    main()
