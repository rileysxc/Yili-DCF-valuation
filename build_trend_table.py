# -*- coding: utf-8 -*-
"""生成公司基础趋势表工作簿（版式与伊利定稿完全一致）。

结构（5 个页签，顺序固定）
    基础趋势表   9 个指标 × 6 期（2021A–2025A + 2026H1），**全部是引用附表的公式**
    资产负债表   科目为行、8 期为列，取自 07 清洗层
    利润表       同上
    现金流量表   同上
    口径说明     口径与陷阱写在文件里，便于交付后自查

数据来源
    原始数据/07_清洗后数据/<代码>_<简称>/<代码>_<表名>_<运行日>_Rev<nn>.csv

输出
    数据文件/基础趋势表/<代码>_<简称>_基础趋势表_<运行日>_Rev<nn>.xlsx

用法
    python build_trend_table.py --symbols 600887,600597,002946
    python build_trend_table.py --symbols 600597 --rev 1
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parents[2] / "原始数据"
LAYER_CLEAN = BASE / "07_清洗后数据"
OUT_DIR = Path(__file__).resolve().parents[2] / "数据文件" / "基础趋势表"

RUN_DATE = "20261011"
SRC_REV = "04"

COMPANIES = {
    "600887": "600887_伊利股份",
    "600597": "600597_光明乳业",
    "002946": "002946_新乳业",
}

TABLES = ["资产负债表", "利润表", "现金流量表"]
SHEET_ORDER = ["基础趋势表", *TABLES, "口径说明"]

PERIODS = ["2020-12-31", "2021-12-31", "2022-12-31", "2023-12-31",
           "2024-12-31", "2025-06-30", "2025-12-31", "2026-06-30"]

# 趋势表的列：(表头, 本期报告期, 上年同期报告期)
# 用报告期名取附表列，不写死列号——写死过一次，2021A 指到了 2020-12-31
TREND_COLS = [
    ("2021A", "2021-12-31", "2020-12-31"),
    ("2022A", "2022-12-31", "2021-12-31"),
    ("2023A", "2023-12-31", "2022-12-31"),
    ("2024A", "2024-12-31", "2023-12-31"),
    ("2025A", "2025-12-31", "2024-12-31"),
    ("2026H1(与2025H1对比)", "2026-06-30", "2025-06-30"),
]

# 这些字段在源 CSV 里是数字，但原稿存的是文本，按文本写以保持一致
TEXT_FIELDS = {"数据源", "是否审计", "公告日期", "币种", "类型", "更新日期"}

MONEY_FMT = "#,##0.00_);[Red](#,##0.00)"
PCT_FMT = "0.00%"
COL_WIDTHS = {"资产负债表": 26.4, "利润表": 20.86, "现金流量表": 31.07}

ROW_REVENUE = "营业收入"
ROW_COST = "营业成本"
ROW_NET_PROFIT = "归属于母公司所有者的净利润"
ROW_OCF = "经营活动产生的现金流量净额"
ROW_EQUITY = "归属于母公司股东权益合计"


def source_path(code: str, table: str) -> Path:
    return LAYER_CLEAN / COMPANIES[code] / f"{code}_{table}_{RUN_DATE}_Rev{SRC_REV}.csv"


def load_tables(code: str) -> dict[str, pd.DataFrame]:
    tables: dict[str, pd.DataFrame] = {}
    for table in TABLES:
        path = source_path(code, table)
        if not path.exists():
            raise FileNotFoundError(f"缺 07 层文件：{path}")
        df = pd.read_csv(path, encoding="utf-8-sig")
        if list(df["date"]) != PERIODS:
            raise AssertionError(f"{path.name}: 报告期不是预期 8 期或未升序")
        tables[table] = df
    return tables


def write_table_sheet(wb: Workbook, table: str, df: pd.DataFrame) -> tuple[dict[str, int], dict[str, str]]:
    """写附表。返回关键科目行号（1-based）与「报告期 -> 列字母」映射。"""
    ws = wb.create_sheet(table)
    ws.append(["科目", *PERIODS])
    for col in df.columns:
        if col == "date":
            continue
        values = df[col].tolist()
        if col in TEXT_FIELDS:
            values = [None if pd.isna(v) else str(v) for v in values]
        ws.append([col, *values])
    ws.column_dimensions["A"].width = COL_WIDTHS[table]
    for i in range(2, 2 + len(PERIODS)):
        ws.column_dimensions[get_column_letter(i)].width = 14.73
    ws.freeze_panes = "B2"

    # 表头必须是预期的 8 期，位置从 B 列起；否则后面的公式会整体错位
    header = [ws.cell(row=1, column=i).value for i in range(2, 2 + len(PERIODS))]
    if header != PERIODS:
        raise AssertionError(f"{table}: 附表表头与预期报告期不一致 {header}")
    period_letter = {p: get_column_letter(i) for i, p in enumerate(PERIODS, start=2)}

    wanted = {ROW_REVENUE, ROW_COST, ROW_NET_PROFIT, ROW_OCF, ROW_EQUITY}
    found: dict[str, int] = {}
    for i, (name, *_rest) in enumerate(ws.iter_rows(values_only=True), start=1):
        if name in wanted:
            found[name] = i
    return found, period_letter


def write_trend_sheet(wb: Workbook, rows: dict[str, int], pl: dict[str, str]) -> None:
    ws = wb.create_sheet("基础趋势表", 0)
    letters = [pl[cur] for _n, cur, _prev in TREND_COLS]

    ws.append(["指标", *[c[0] for c in TREND_COLS]])
    ws.append(["营业收入/亿元", *[
        f"=利润表!{L}{rows[ROW_REVENUE]}/100000000" for L in letters]])
    ws.append(["营收增速", *[
        f"=利润表!{pl[cur]}{rows[ROW_REVENUE]}/利润表!{pl[prev]}{rows[ROW_REVENUE]}-1"
        for _n, cur, prev in TREND_COLS]])
    ws.append(["归母净利润/亿元", *[
        f"=利润表!{L}{rows[ROW_NET_PROFIT]}/100000000" for L in letters]])
    ws.append(["归母净利增速", *[
        f"=利润表!{pl[cur]}{rows[ROW_NET_PROFIT]}/利润表!{pl[prev]}{rows[ROW_NET_PROFIT]}-1"
        for _n, cur, prev in TREND_COLS]])
    ws.append(["毛利率", *[
        f"=(利润表!{L}{rows[ROW_REVENUE]}-利润表!{L}{rows[ROW_COST]})"
        f"/利润表!{L}{rows[ROW_REVENUE]}" for L in letters]])
    ws.append(["归母净利率", *[
        f"=利润表!{L}{rows[ROW_NET_PROFIT]}/利润表!{L}{rows[ROW_REVENUE]}" for L in letters]])
    ws.append(["经营现金流/归母净利润", *[
        f"=现金流量表!{L}{rows[ROW_OCF]}/利润表!{L}{rows[ROW_NET_PROFIT]}" for L in letters]])
    ws.append(["平均净资产/亿元", *[
        f"=(资产负债表!{pl[prev]}{rows[ROW_EQUITY]}"
        f"+资产负债表!{pl[cur]}{rows[ROW_EQUITY]})/2/100000000"
        for _n, cur, prev in TREND_COLS]])
    # ROE 引用趋势表自己的第 4、9 行，所以用趋势表列字母（B 起），不是附表列字母
    ws.append(["ROE", *[
        f"={get_column_letter(2 + i)}4/{get_column_letter(2 + i)}9"
        for i in range(len(TREND_COLS))]])

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "center")
    ws.row_dimensions[1].height = 40.5
    for r in (2, 4, 8, 9):
        for c in range(2, 2 + len(TREND_COLS)):
            ws.cell(row=r, column=c).number_format = MONEY_FMT
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")
    for r in (3, 5, 6, 7, 10):
        for c in range(2, 2 + len(TREND_COLS)):
            ws.cell(row=r, column=c).number_format = PCT_FMT
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="right")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 12
    ws.freeze_panes = "B2"


def build_notes(tables: dict[str, pd.DataFrame], extra: list[tuple[str, str]]) -> list[tuple[str, str]]:
    pl = tables["利润表"]
    last = pl.index[-1]
    revenue = pl.at[last, ROW_REVENUE]
    net = pl.at[last, ROW_NET_PROFIT]
    cost = pl.at[last, ROW_COST]
    return [
        ("数据版本", f"07 清洗层 Rev{SRC_REV}（三家口径完全一致：首列 date + ISO 字符串、"
                 "报表 8 期、市场类 2021-01-01~2026-10-09）"),
        ("取数日期", "2026-10-10 / 10-11 取数；市场数据（行情、估值）截止 2026-10-09"),
        ("数据来源", "AKShare：新浪财经（三张报表）、东方财富（估值日频）、百度（估值十年）"),
        ("报告期", "8 期：2020-12-31 作年度基数、2025-06-30 作中期基数，"
                "加 2021–2025 五个年报与 2026 半年报"),
        ("2026H1 口径", "6 个月累计。只能与 2025H1 对比，不可与年度数直接比较——"
                    "本表比率行（毛利率/净利率/现金流比）可比，绝对额行（营业收入/归母净利润）不可比"),
        ("营收口径", "营业收入（不含利息收入）；含利息收入的营业总收入仅作对照"),
        ("毛利率口径", "（营业收入 − 营业成本）÷ 营业收入，取利润表「营业成本」。"
                  "财务摘要的「营业成本」实为营业总成本，用它算会得到错误结果"),
        ("归母净利率口径", f"归母净利润 ÷ 营业收入（本表 2026H1 = {net / revenue:.2%}）"),
        ("净利润口径提醒", "同期「净利润 ÷ 营业收入」（含少数股东损益）是另一个口径、数值不同，"
                     "全篇须统一并加脚注"),
        ("经营现金流口径", "经营活动产生的现金流量净额 ÷ 归母净利润（分母是归母净利润，不是净利润）"),
        ("增速口径", "一律同比：(本期 − 上年同期) ÷ 上年同期。2026H1 的上年同期是 2025H1"),
        ("平均净资产口径", "归母净资产（期初 + 期末）÷ 2。若改用期末净资产口径，结果差别较大，"
                    "报告须选定一种并写明"),
        ("ROE 半年值", "2026H1 的 ROE 用 2025H1 与 2026H1 的平均净资产计算，是半年值，"
                   "只作标注，不与年度 ROE 比较、不做年化"),
        ("金额单位", "亿元。附表的原始数据为元，本表金额行已除以 100000000"),
        ("公式缓存", "趋势表全部为引用附表的公式，已设置打开时强制重算。"
                 "若用不重算公式的查看器打开，趋势表可能显示空白"),
    ] + extra


def write_notes_sheet(wb: Workbook, notes: list[tuple[str, str]]) -> None:
    ws = wb.create_sheet("口径说明")
    ws.append(["项目", "说明"])
    for row in notes:
        ws.append(list(row))
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 78
    ws.freeze_panes = "A2"


EXTRA_NOTES: dict[str, list[tuple[str, str]]] = {
    "600887": [(
        "2021 年股本与净资产变动",
        "出处：《2021 年年度报告·财务报表附注·公司基本情况》，报告页码 96（PDF 第 100 页）。"
        "2021 年 7 月回购注销 447,200 股限制性股票，总股本变 6,082,177,633 股；"
        "经证监许可[2021]3196 号批准，2021 年 11 月非公开发行 317,953,285 股，总股本变 6,400,130,918 股。"
        "定增募资使 2021 年归母净资产由 303.84 亿元升至 477.08 亿元，ROE 因此被摊薄，"
        "读 2021 年 ROE 时须注意",
    )],
}


def build_workbook(code: str) -> Workbook:
    tables = load_tables(code)
    wb = Workbook()
    wb.remove(wb.active)

    rows: dict[str, int] = {}
    period_letter: dict[str, str] = {}
    for table in TABLES:
        found, letters = write_table_sheet(wb, table, tables[table])
        rows.update(found)
        period_letter.update(letters)
    missing = {ROW_REVENUE, ROW_COST, ROW_NET_PROFIT, ROW_OCF, ROW_EQUITY} - set(rows)
    if missing:
        raise AssertionError(f"{code}: 附表中找不到科目 {missing}")

    write_trend_sheet(wb, rows, period_letter)
    write_notes_sheet(wb, build_notes(tables, EXTRA_NOTES.get(code, [])))
    wb._sheets = [wb[name] for name in SHEET_ORDER]
    wb.calculation.fullCalcOnLoad = True
    return wb


def check_column_consistency(codes: list[str]) -> None:
    """三家附表字段名与顺序必须一致，否则行号语义会错位。"""
    ref: dict[str, list[str]] = {}
    for code in codes:
        for table in TABLES:
            cols = list(pd.read_csv(source_path(code, table),
                                    encoding="utf-8-sig", nrows=0).columns)
            if table in ref and ref[table] != cols:
                raise AssertionError(f"{code} 的 {table} 列名/顺序与首家不一致")
            ref.setdefault(table, cols)
    print(f"列一致性通过：{len(ref)} 张表，字段名与顺序三家一致")


def main() -> None:
    parser = argparse.ArgumentParser(description="生成基础趋势表工作簿")
    parser.add_argument("--symbols", default="600887,600597,002946")
    parser.add_argument("--rev", type=int, default=0)
    args = parser.parse_args()
    rev = f"{args.rev:02d}"
    codes = [s.strip() for s in args.symbols.split(",") if s.strip()]

    check_column_consistency(codes)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for code in codes:
        out = OUT_DIR / f"{COMPANIES[code]}_基础趋势表_{RUN_DATE}_Rev{rev}.xlsx"
        build_workbook(code).save(out)
        print(f"{code} -> {out.name}")


if __name__ == "__main__":
    main()
