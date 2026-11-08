# -*- coding: utf-8 -*-
"""对 07_清洗后数据 做缺失值与异常值体检，输出供人工判断的表。

输出
    07_清洗后数据/00_缺失与异常汇总_<运行日>_Rev<nn>.csv   按 文件 x 列 汇总
    07_清洗后数据/00_缺失与异常明细_<运行日>_Rev<nn>.csv   每个具体异常点一行

    文件名里的 Rev 是「被检视的 07 层版本号」，不是报告自身的版本号。

异常值判据（分开列示，便于逐条判断）
    逻辑矛盾    与分布无关的硬性矛盾，例如 high < low、流通市值 > 总市值、总股本 <= 0
    日变动越界  行情用 close 算日收益、估值用自带的当日涨跌幅，|变动| > 10.5% 即越过
                A 股 10% 涨跌停，属数据或口径问题
    序列跳变    PE/PB 稀疏序列相邻采样点变动 > 30%
    数量级异常  只用于报表类。同一科目跨期相差 50 倍以上（典型是单位或口径搞错），
                或该科目中位数为 0 却又有非零期（零值混杂，可能是缺失被填 0）

    报表类刻意不做分布离群检测：8 个报告期里混着 5 个年报和 3 个半年报，
    半年报的金额天然是年报的一半量级，用离群规则只会把半年报判成异常
    （实测会把光明乳业的销售费用、管理费用全部标红）。
    也不把「同比大幅变动」列为异常：那是财务分析内容，不是数据质量问题，
    按 50% 阈值会产生 900 多条，无法用于判断。
    也刻意不对价格、市值、估值这类水平值时序做离群检测：它们非平稳，只会把
    整段牛熊区间判成异常（实测伊利市净率 1395 个交易日里 273 天会被标红）。

缺失值分两类，因为处置方式完全不同
    结构性缺失  整列全空。通常是该主体根本没有这个科目（比如乳企报表里的保险、
                银行专用行项目），属于正常现象，忽略即可。按 代码 x 表名 聚合成一行。
    局部缺失    只有部分报告期为空。这一批才是需要逐条判断的。

关注级别
    高    逻辑矛盾，或属于报告核心科目（营收、成本、归母净利、经营现金流、
          资产总计、货币资金、存货、固定资产、有息负债、PE、PB 等）
    中    非核心科目

刻意不做的事
    - 不填补缺失值，不剔除异常值。本脚本只出报告，处置由人工决定。
    - 财务摘要的报告期在列名上，行是指标，故只做缺失值统计，不做离群统计。

用法
    python audit_data_quality.py --rev 4
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parents[2] / "原始数据"
LAYER_CLEAN = BASE / "07_清洗后数据"

RUN_DATE = "20261011"
TARGET_COL = "date"

# 逻辑规则的取值范围，键为 (表名, 列名)
LOGIC_RULES: list[tuple[str, str, str, str]] = [
    ("行情", "high", "<", "low"),
    ("行情", "close", ">", "high"),
    ("行情", "close", "<", "low"),
    ("估值", "流通市值", ">", "总市值"),
    ("估值", "流通股本", ">", "总股本"),
]

# 必须为正的列
POSITIVE_COLS: list[tuple[str, str]] = [
    ("行情", "close"), ("行情", "volume"), ("行情", "amount"),
    ("行情", "high"), ("行情", "low"), ("行情", "open"),
    ("估值", "总市值"), ("估值", "总股本"), ("估值", "当日收盘价"),
    ("资产负债表", "资产总计"),
]

SUMMARY_PREFIX = "财务摘要"
PE_PB_TABLES = {"估值十年PE", "估值十年PB"}

SERIES_JUMP_THRESHOLD = 0.30   # PE/PB 相邻采样点变动阈值
MAGNITUDE_RATIO = 50           # 同一科目跨期相差倍数阈值
MAGNITUDE_MIN_SAMPLES = 4      # 该列至少要有多于 4 个非空期才判

EXAMPLE_LIMIT = 3
STRUCTURAL_LIST_LIMIT = 6

# 报告核心科目：这些列一旦缺失或离群，必须逐条看
KEY_COLS = {
    "营业收入", "营业总收入", "营业总成本", "营业成本", "销售费用", "管理费用",
    "财务费用", "研发费用", "营业利润", "利润总额", "净利润",
    "归属于母公司所有者的净利润", "经营活动产生的现金流量净额",
    "货币资金", "应收票据及应收账款", "应收账款", "存货", "固定资产",
    "资产总计", "负债合计", "归属于母公司所有者权益合计",
    "短期借款", "长期借款", "应付债券", "一年内到期的非流动负债",
    "当日收盘价", "总市值", "PE(TTM)", "市净率",
    "close", "volume", "pe_ttm", "pb",
}


def attention(table: str, col: str, is_logic: bool = False) -> str:
    if is_logic or col in KEY_COLS:
        return "高"
    return "中"


def load_rev_files(rev: str) -> list[Path]:
    return sorted(LAYER_CLEAN.glob(f"*/*_Rev{rev}.csv"))


def numeric_columns(df: pd.DataFrame) -> list[str]:
    cols = []
    for col in df.columns:
        if col == TARGET_COL:
            continue
        if pd.to_numeric(df[col], errors="coerce").notna().any():
            cols.append(col)
    return cols


def collect_missing(code: str, table: str, df: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    """整列全空的聚合成一行（结构性缺失），其余逐列列出（局部缺失）。"""
    structural: list[str] = []
    rows: list[dict] = []
    detail: list[dict] = []
    for col in df.columns:
        mask = df[col].isna()
        missing = int(mask.sum())
        if missing == 0:
            continue
        if missing == len(df):
            structural.append(col)
            continue
        rows.append({
            "代码": code, "表名": table, "列名": col,
            "问题类型": "缺失值-局部", "规则": "空值计数",
            "关注级别": attention(table, col),
            "总行数": len(df), "非空数": int(df[col].notna().sum()),
            "问题数": missing, "占比": round(missing / len(df), 4),
            "最小值": None, "最大值": None, "阈值": None, "示例": "",
        })
        for i in df.index[mask]:
            if TARGET_COL in df.columns:
                date_field = str(df.at[i, TARGET_COL])
                note = f"{col} 该期为空"
            else:
                date_field = col
                note = f"{df.at[i, df.columns[1]]} 在该期为空"
            detail.append({
                "代码": code, "表名": table, "列名": col, "问题类型": "缺失值-局部",
                "规则": "空值计数", "关注级别": attention(table, col),
                "日期": date_field, "值": None, "阈值": "—", "说明": note,
            })
    if structural:
        head = "、".join(structural[:STRUCTURAL_LIST_LIMIT])
        suffix = f" 等 {len(structural)} 列" if len(structural) > STRUCTURAL_LIST_LIMIT else ""
        rows.append({
            "代码": code, "表名": table, "列名": "多列整列为空",
            "问题类型": "缺失值-结构性", "规则": "整列空值",
            "关注级别": "高" if set(structural) & KEY_COLS else "中",
            "总行数": len(df), "非空数": 0, "问题数": len(df) * len(structural),
            "占比": 1.0, "最小值": None, "最大值": None, "阈值": None,
            "示例": f"{head}{suffix}",
        })
    return rows, detail
    rows: list[dict] = []
    for col in df.columns:
        values = df[col]
        missing = int(values.isna().sum())
        if missing == 0:
            continue
        rows.append({
            "代码": code, "表名": table, "列名": col, "问题类型": "缺失值",
            "规则": "空值计数", "总行数": len(df), "非空数": int(values.notna().sum()),
            "问题数": missing, "占比": round(missing / len(df), 4) if len(df) else None,
            "最小值": None, "最大值": None, "阈值": None, "示例": "",
        })
    return rows


def collect_logic(code: str, table: str, df: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    summary: list[dict] = []
    detail: list[dict] = []

    def flag(col: str, rule: str, mask: pd.Series, values: pd.Series, threshold: str) -> None:
        hits = int(mask.sum())
        if not hits:
            return
        samples = df.loc[mask, TARGET_COL].astype(str).head(EXAMPLE_LIMIT).tolist()
        summary.append({
            "代码": code, "表名": table, "列名": col, "问题类型": "异常值-逻辑",
            "规则": rule, "关注级别": attention(table, col, is_logic=True),
            "总行数": len(df), "非空数": int(values.notna().sum()),
            "问题数": hits, "占比": round(hits / len(df), 4),
            "最小值": pd.to_numeric(values, errors="coerce").min(),
            "最大值": pd.to_numeric(values, errors="coerce").max(),
            "阈值": threshold, "示例": ", ".join(samples),
        })
        for idx in df.index[mask]:
            detail.append({
                "代码": code, "表名": table, "列名": col, "问题类型": "异常值-逻辑",
                "规则": rule, "关注级别": "高", "日期": str(df.at[idx, TARGET_COL]),
                "值": values.at[idx], "阈值": threshold,
                "说明": f"{col} {rule}",
            })

    for rule_table, left, op, right in LOGIC_RULES:
        if table != rule_table or left not in df.columns or right not in df.columns:
            continue
        lv = pd.to_numeric(df[left], errors="coerce")
        rv = pd.to_numeric(df[right], errors="coerce")
        mask = (lv < rv) if op == "<" else (lv > rv)
        flag(left, f"{left} {op} {right}", mask.fillna(False), lv, f"{left} {op} {right}")

    for rule_table, col in POSITIVE_COLS:
        if table != rule_table or col not in df.columns:
            continue
        values = pd.to_numeric(df[col], errors="coerce")
        flag(col, f"{col} <= 0", (values <= 0).fillna(False), values, f"{col} <= 0")

    return summary, detail


def emit(code: str, table: str, col: str, rule: str, mask: pd.Series, values: pd.Series,
         threshold: str, df: pd.DataFrame, source_col: str | None = None,
         is_logic: bool = False, note: str | None = None) -> tuple[list[dict], list[dict]]:
    """把命中的行写成 1 条汇总 + N 条明细。"""
    hits = int(mask.sum())
    if not hits:
        return [], []
    kinds = "异常值-逻辑" if is_logic else "异常值-离群"
    level = "高" if is_logic else attention(table, col)
    base = pd.to_numeric(df[source_col or col], errors="coerce")
    hit_values = pd.to_numeric(values[mask], errors="coerce")
    dates = df.loc[mask, TARGET_COL].astype(str)
    examples = ", ".join(
        f"{d}={v:.4g}" for d, v in zip(dates.head(EXAMPLE_LIMIT), hit_values.head(EXAMPLE_LIMIT))
    )
    summary = [{
        "代码": code, "表名": table, "列名": col, "问题类型": kinds,
        "规则": rule, "关注级别": level,
        "总行数": len(df), "非空数": int(base.notna().sum()),
        "问题数": hits, "占比": round(hits / len(df), 4),
        "最小值": hit_values.min(), "最大值": hit_values.max(),
        "阈值": threshold, "示例": examples,
    }]
    detail = [{
        "代码": code, "表名": table, "列名": col, "问题类型": kinds,
        "规则": rule, "关注级别": level,
        "日期": str(df.at[i, TARGET_COL]), "值": values.at[i],
        "阈值": threshold, "说明": note or f"{col} {rule}",
    } for i in df.index[mask]]
    return summary, detail


def collect_outliers(code: str, table: str, df: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    """按表的语义分派异常值规则；水平值时序不做离群检测。"""
    if table.startswith(SUMMARY_PREFIX):
        return [], []

    if table == "行情":
        results = []
        close = pd.to_numeric(df["close"], errors="coerce")
        ret = close.pct_change()
        results.append(emit(code, table, "close", "日收益绝对值>10.5%",
                            (ret.abs() > 0.105).fillna(False), ret,
                            "|日收益| > 10.5%", df, source_col="close",
                            note="日收益越过 A 股 10% 涨跌停"))
        volume = pd.to_numeric(df["volume"], errors="coerce")
        results.append(emit(code, table, "volume", "成交量为 0",
                            (volume == 0).fillna(False), volume,
                            "volume == 0", df, note="疑似停牌或数据缺失"))
    elif table == "估值":
        col = "当日涨跌幅"
        values = pd.to_numeric(df[col], errors="coerce")
        results = [emit(code, table, col, "当日涨跌幅绝对值>10.5%",
                        (values.abs() > 10.5).fillna(False), values,
                        "|当日涨跌幅| > 10.5", df, note="越过 A 股 10% 涨跌停")]
    elif table in PE_PB_TABLES:
        col = df.columns[-1]
        values = pd.to_numeric(df[col], errors="coerce")
        ret = values.pct_change()
        results = [emit(code, table, col, "相邻采样点变动>30%",
                        (ret.abs() > SERIES_JUMP_THRESHOLD).fillna(False), ret,
                        f"|相邻采样点变动| > {SERIES_JUMP_THRESHOLD:.0%}", df, source_col=col,
                        note="PE/PB 序列跳变，可能是盈利基数切换")]
    else:
        results = [magnitude_outliers(code, table, df)]

    summary: list[dict] = []
    detail: list[dict] = []
    for part_summary, part_detail in results:
        summary.extend(part_summary)
        detail.extend(part_detail)
    return summary, detail


def magnitude_outliers(code: str, table: str, df: pd.DataFrame) -> tuple[list[dict], list[dict]]:
    """报表类的数量级异常与零值混杂。"""
    summary: list[dict] = []
    detail: list[dict] = []

    for col in numeric_columns(df):
        values = pd.to_numeric(df[col], errors="coerce")
        valid = values.dropna()
        if len(valid) < MAGNITUDE_MIN_SAMPLES:
            continue
        median = valid.abs().median()

        if median == 0:
            nonzero = int((valid != 0).sum())
            if 0 < nonzero < len(valid):
                dates = ", ".join(
                    f"{df.at[i, TARGET_COL]}={values.at[i]:.4g}"
                    for i in valid.index[valid != 0][:EXAMPLE_LIMIT]
                )
                summary.append({
                    "代码": code, "表名": table, "列名": col, "问题类型": "异常值-零值混杂",
                    "规则": "中位数为0但存在非零期", "关注级别": attention(table, col),
                    "总行数": len(df), "非空数": int(len(valid)), "问题数": nonzero,
                    "占比": round(nonzero / len(df), 4),
                    "最小值": valid.min(), "最大值": valid.max(),
                    "阈值": "该列中位数绝对值为 0", "示例": dates,
                })
            continue

        mask = ((valid.abs() > MAGNITUDE_RATIO * median)
                | (valid.abs() < median / MAGNITUDE_RATIO))
        mask = mask.reindex(df.index, fill_value=False)
        part_summary, part_detail = emit(
            code, table, col, f"跨期相差>{MAGNITUDE_RATIO}倍", mask, values,
            f"中位数绝对值 {median:.6g}，偏离 {MAGNITUDE_RATIO} 倍以上", df,
            note="数量级异常，疑似单位或口径错误",
        )
        summary.extend(part_summary)
        detail.extend(part_detail)

    return summary, detail


def main() -> None:
    parser = argparse.ArgumentParser(description="07 层缺失值与异常值体检")
    parser.add_argument("--rev", default="5", help="被检视的 07 层版本号，默认取当前最新的 5")
    args = parser.parse_args()
    rev = f"{int(args.rev):02d}"

    summary_rows: list[dict] = []
    detail_rows: list[dict] = []
    for path in load_rev_files(rev):
        code = path.name[:6]
        table = re.sub(r"^\d{6}_|_\d{8}_Rev\d+\.csv$", "", path.name)
        df = pd.read_csv(path, encoding="utf-8-sig")

        miss_sum, miss_detail = collect_missing(code, table, df)
        logic_sum, logic_detail = collect_logic(code, table, df)
        out_sum, out_detail = collect_outliers(code, table, df)
        summary_rows.extend(miss_sum + logic_sum + out_sum)
        detail_rows.extend(miss_detail + logic_detail + out_detail)

    summary = pd.DataFrame(summary_rows)
    if not summary.empty:
        summary = summary.sort_values(
            ["关注级别", "问题类型", "问题数"], ascending=[True, True, False]
        )
    detail = pd.DataFrame(detail_rows)

    sum_path = LAYER_CLEAN / f"00_缺失与异常汇总_{RUN_DATE}_Rev{rev}.csv"
    detail_path = LAYER_CLEAN / f"00_缺失与异常明细_{RUN_DATE}_Rev{rev}.csv"
    summary.to_csv(sum_path, index=False, encoding="utf-8-sig")
    detail.to_csv(detail_path, index=False, encoding="utf-8-sig")

    print(f"files={len(load_rev_files(rev))} summary={len(summary)} detail={len(detail)}")
    print(f"summary={sum_path.name}")
    print(f"detail={detail_path.name}")


if __name__ == "__main__":
    main()
