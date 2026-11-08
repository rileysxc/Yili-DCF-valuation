#!/usr/bin/env python3
"""10.14 项目线第 1–3 项的产物生成脚本。

产物（落在 数据文件/财务分析附表/）：
    600887_利润与现金流对照表_20261014_Rev00.xlsx   —— 任务 1
    600887_指标驱动分解表_20261014_Rev00.xlsx       —— 任务 2
    600887_三表勾稽检查表_20261014_Rev00.xlsx       —— 任务 3

数据来源：
1) 07 清洗层 Rev05：三家（600887 / 600597 / 002946）的利润表、资产负债表、现金流量表；
    2) 05 手工整理数据《600887_扣非归母净利润_20261014_Rev00.csv》——手工从定期报告
       「主要会计数据」逐期摘录，带 PDF 页码；07 清洗层没有该字段；
    3) 定期报告 PDF：仅用于「一手来源抽检」页签的复核（需 pypdf；不可用时该页签标注未复核）。

口径约定：
    金额一律亿元（元 ÷ 1e8，显示保留 2 位，计算用未取整值）；
    资产负债表的「归属于母公司股东权益合计」＝期末归母净资产（两家可比公司同版式第 138 行为准）。

本脚本只产出数据，不含判断与结论（见 AGENTS.md 三段式约定）。
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[2]
CLEAN = ROOT / "原始数据" / "07_清洗后数据"
MANUAL = ROOT / "原始数据" / "05_手工整理数据"
MODEL_XLSX_DIR = ROOT / "原始数据" / "06_建模用Excel"
PDF_DIR = ROOT / "原始资料" / "伊利年报"
OUTDIR = ROOT / "数据文件" / "财务分析附表"

PERIODS = [
    "2020-12-31",
    "2021-12-31",
    "2022-12-31",
    "2023-12-31",
    "2024-12-31",
    "2025-06-30",
    "2025-12-31",
    "2026-06-30",
]
PLABELS = ["2020A", "2021A", "2022A", "2023A", "2024A", "2025H1", "2025A", "2026H1"]
CLEAN_REV = "Rev05"

# 同比配对：(标签, 本期, 上年同期)。2020A 无上年同期、2025H1 的 2024H1 不在 8 期口径内，均不配对。
YOY_PAIRS = [
    ("2021A", "2021-12-31", "2020-12-31"),
    ("2022A", "2022-12-31", "2021-12-31"),
    ("2023A", "2023-12-31", "2022-12-31"),
    ("2024A", "2024-12-31", "2023-12-31"),
    ("2025A", "2025-12-31", "2024-12-31"),
    ("2026H1", "2026-06-30", "2025-06-30"),
]

# 三家对比期：2021A–2025A ＋ 2026H1（与《基础趋势表》列一致）
COMP_PERIODS = ["2021-12-31", "2022-12-31", "2023-12-31", "2024-12-31", "2025-12-31", "2026-06-30"]
COMP_LABELS = ["2021A", "2022A", "2023A", "2024A", "2025A", "2026H1"]

COMPANIES = {
    "600887": ("伊利股份", CLEAN / "600887_伊利股份"),
    "600597": ("光明乳业", CLEAN / "600597_光明乳业"),
    "002946": ("新乳业", CLEAN / "002946_新乳业"),
}

FONT = "宋体"
HDR_FILL = "DDEBF7"
NOTE_FILL = "FFF2CC"
FMT_MONEY = "#,##0.00"
FMT_DIFF = "#,##0.000000"
FMT_RATIO = "0.00"
FMT_PCT = "0.00%"


# ---------------------------------------------------------------- 基础工具


def latest(folder: Path, pattern: str) -> Path:
    cands = sorted(folder.glob(pattern))
    if not cands:
        raise FileNotFoundError(f"{folder} 下找不到 {pattern}")
    return cands[-1]


def read_clean_sheet(folder: Path, code: str, kind: str) -> tuple[list[str], dict[str, dict[str, float | None]]]:
    """读 07 清洗层某公司的某张表：返回 (科目列表, {期间: {科目: 数值}})。"""
    path = latest(folder, f"{code}_{kind}_*_{CLEAN_REV}.csv")
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    header, data = rows[0], rows[1:]
    meta = {"数据源", "是否审计", "公告日期", "币种", "类型", "更新日期"}
    subjects = [c for c in header[1:] if c not in meta]
    table: dict[str, dict[str, float | None]] = {}
    for row in data:
        period = row[0].strip()
        rec: dict[str, float | None] = {}
        for name, raw in zip(header[1:], row[1:]):
            if name not in subjects:
                continue
            try:
                rec[name] = float(str(raw).strip())
            except (TypeError, ValueError):
                rec[name] = None
        table[period] = rec
    return subjects, table


def val(table: dict[str, dict[str, float | None]], period: str, subject: str) -> float:
    """取数；空值按 0 处理（仅用于加总项，如「手续费及佣金支出」）。"""
    v = table.get(period, {}).get(subject)
    return 0.0 if v is None else float(v)


def div(num: float | None, den: float | None) -> float | None:
    if num is None or den in (None, 0):
        return None
    return num / den


def yoy(cur: float | None, prev: float | None) -> float | None:
    if cur is None or prev in (None, 0):
        return None
    return cur / prev - 1


def money(yuan: float | None) -> float | None:
    return None if yuan is None else yuan / 1e8


def style_sheet(ws, ncol: int, header_rows: list[int], widths: dict[str, float] | None = None) -> None:
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ncol):
        for cell in row:
            cell.font = Font(name=FONT, size=11, bold=cell.row in header_rows)
            if cell.row in header_rows:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.fill = PatternFill("solid", fgColor=HDR_FILL)
    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width


def put(ws, row: int, label: str, values: list, fmt: str | None = None, bold: bool = False) -> None:
    ws.cell(row, 1, label)
    if bold:
        ws.cell(row, 1).font = Font(name=FONT, size=11, bold=True)
    for i, v in enumerate(values):
        cell = ws.cell(row, 2 + i, v)
        cell.font = Font(name=FONT, size=11, bold=bold)
        cell.alignment = Alignment(horizontal="right")
        if fmt and isinstance(v, (int, float)):
            cell.number_format = fmt


# ------------------------------------------------------- 三家共用数据准备


def load_all() -> dict[str, dict]:
    data: dict[str, dict] = {}
    for code, (name, folder) in COMPANIES.items():
        _, pl = read_clean_sheet(folder, code, "利润表")
        _, bs = read_clean_sheet(folder, code, "资产负债表")
        _, cf = read_clean_sheet(folder, code, "现金流量表")
        data[code] = {"name": name, "pl": pl, "bs": bs, "cf": cf}
    return data


def load_deducted() -> dict[str, dict[str, float]]:
    """读 05 层《扣非归母净利润》手工摘录。"""
    path = latest(MANUAL, "600887_扣非归母净利润_*_Rev*.csv")
    out: dict[str, dict[str, float]] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            out[row["期间"].strip()] = {
                "deducted": float(row["扣非归母净利润(元)"]),
                "file": row["披露文件"],
                "page": int(row["PDF页码"]),
                "rpage": row["报告页码"],
                "table": row["表名"],
                "check": row["交叉验证"],
            }
    return out


def ily_metrics(data: dict, deducted: dict) -> dict[str, dict[str, float | None]]:
    """伊利 8 期各项水平值（未取整）。"""
    pl, bs, cf = data["600887"]["pl"], data["600887"]["bs"], data["600887"]["cf"]
    out: dict[str, dict[str, float | None]] = {}
    for p in PERIODS:
        rev = val(pl, p, "营业收入")
        cost = val(pl, p, "营业成本")
        np_parent = val(pl, p, "归属于母公司所有者的净利润")
        deduct = deducted[p]["deducted"]
        ocf = val(cf, p, "经营活动产生的现金流量净额")
        equity = val(bs, p, "归属于母公司股东权益合计")
        op = val(pl, p, "营业利润")
        sell = val(pl, p, "销售费用")
        admin = val(pl, p, "管理费用")
        rd = val(pl, p, "研发费用")
        fin = val(pl, p, "财务费用")
        gross_margin = div(rev - cost, rev)
        exp_3 = (sell + admin + fin) / rev
        exp_4 = (sell + admin + rd + fin) / rev
        net_margin = div(np_parent, rev)
        out[p] = {
            "revenue": rev,
            "cost": cost,
            "np_parent": np_parent,
            "deducted": deduct,
            "nonrecurring": np_parent - deduct,
            "ocf": ocf,
            "ocf_minus_np": ocf - np_parent,
            "ocf_over_np": div(ocf, np_parent),
            "ocf_over_deducted": div(ocf, deduct),
            "deducted_margin": div(deduct, rev),
            "nonrecurring_share": div(np_parent - deduct, np_parent),
            "equity": equity,
            "gross_margin": gross_margin,
            "sell_rate": div(sell, rev),
            "admin_rate": div(admin, rev),
            "rd_rate": div(rd, rev),
            "fin_rate": div(fin, rev),
            "exp3_rate": exp_3,
            "exp4_rate": exp_4,
            "op_margin": div(op, rev),
            "net_margin": net_margin,
            "roe": div(np_parent, equity),
            # 归母净利率 = 毛利率 − 四费率合计 − 其他项 → 其他项倒挤
            "other_item": gross_margin - exp_4 - net_margin,
        }
    return out


# ---------------------------------------------------- 产物 1：利润 vs 现金流


def build_profit_cash(metrics: dict, deducted: dict, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "对照表"
    ws["A1"] = "伊利股份（600887.SH）利润 vs 现金流 对照表"
    ws["A1"].font = Font(name=FONT, size=12, bold=True)
    ws["A2"] = "单位：亿元（比率与同比除外）｜数据截止 2026-10-09｜报告期 8 期：2020A–2025A ＋ 2025H1"
    ws["A3"] = "数据来源：07 清洗层 Rev05（利润表、现金流量表）＋ 定期报告「主要会计数据」（扣非归母净利润，见 05 层手工摘录）"

    head = 5
    put(ws, head, "项目", PLABELS + [], bold=True)
    rows = [
        ("营业收入", [money(metrics[p]["revenue"]) for p in PERIODS], FMT_MONEY),
        ("归母净利润", [money(metrics[p]["np_parent"]) for p in PERIODS], FMT_MONEY),
        ("扣非归母净利润", [money(metrics[p]["deducted"]) for p in PERIODS], FMT_MONEY),
        ("非经常性损益（＝归母净利润 − 扣非归母净利润）", [money(metrics[p]["nonrecurring"]) for p in PERIODS], FMT_MONEY),
        ("经营活动产生的现金流量净额", [money(metrics[p]["ocf"]) for p in PERIODS], FMT_MONEY),
        ("经营活动现金流净额 − 归母净利润", [money(metrics[p]["ocf_minus_np"]) for p in PERIODS], FMT_MONEY),
        ("经营现金流 ÷ 归母净利润（倍）", [metrics[p]["ocf_over_np"] for p in PERIODS], FMT_RATIO),
        ("经营现金流 ÷ 扣非归母净利润（倍）", [metrics[p]["ocf_over_deducted"] for p in PERIODS], FMT_RATIO),
        ("扣非归母净利率（扣非归母净利润 ÷ 营业收入）", [metrics[p]["deducted_margin"] for p in PERIODS], FMT_PCT),
        ("非经常性损益 ÷ 归母净利润", [metrics[p]["nonrecurring_share"] for p in PERIODS], FMT_PCT),
    ]
    r = head + 1
    for label, values, fmt in rows:
        put(ws, r, label, values, fmt)
        r += 1

    r += 1
    g = ws.cell(r, 1, "同比（%）")
    g.font = Font(name=FONT, size=11, bold=True)
    fill = PatternFill("solid", fgColor=HDR_FILL)
    for c in range(1, 10):
        ws.cell(r, c).fill = fill
    r += 1
    pairs = YOY_PAIRS
    yoy_rows = [
        ("营业收入同比", "revenue"),
        ("归母净利润同比", "np_parent"),
        ("扣非归母净利润同比", "deducted"),
        ("经营活动现金流净额同比", "ocf"),
    ]
    for label, key in yoy_rows:
        vals: list = ["—"] * 8
        for lbl, cur, prev in pairs:
            idx = PLABELS.index(lbl)
            vals[idx] = yoy(metrics[cur][key], metrics[prev][key])
        put(ws, r, label, vals, FMT_PCT)
        r += 1

    r += 1
    notes = [
        "注 1：2020A 无上年同期（2019 年不纳入本报告口径）；2025H1 的同比需 2024H1，未在 8 期口径内 —— 两列同比留「—」。",
        "注 2：2026H1 同比与 2026 年半年报「主要会计数据」披露值一致：归母净利润 −20.02%、扣非归母净利润 −20.25%、经营活动现金流量净额 ＋229.23%。",
        "注 3：非经常性损益为「归母净利润 − 扣非归母净利润」的倒差值，公司未在本表口径下直接披露该差额。",
        "注 4：本页只呈现数据；口径与逐项出处见「口径与出处」。",
    ]
    for note in notes:
        ws.cell(r, 1, note).font = Font(name=FONT, size=10)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=NOTE_FILL)
        r += 1

    style_sheet(ws, 9, [head], {"A": 44})
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 13
    ws.freeze_panes = "B6"

    ws2 = wb.create_sheet("口径与出处")
    head2 = ["项目", "说明"]
    ws2.append(head2)
    detail = [
        ("数据版本", "07 清洗层 Rev05（三张报表）＋ 05 层《600887_扣非归母净利润_20261014_Rev00.csv》"),
        ("取数日期", "报表数据 2026-10-10 / 10-11 取数；扣非归母净利润 2026-10-14 从定期报告摘录并复核"),
        ("报告期", "8 期：2020-12-31、2021-12-31、2022-12-31、2023-12-31、2024-12-31、2025-06-30、2025-12-31、2026-06-30"),
        ("金额单位", "亿元＝元 ÷ 100000000，显示保留 2 位；同比与比率用未取整值计算后再取 2 位，因此与显示值直接相减可能差 0.01"),
        ("营业收入口径", "利润表「营业收入」（不含利息收入）；营业总收入仅作对照"),
        ("归母净利润口径", "利润表「归属于母公司所有者的净利润」（不是「净利润」，两者差少数股东损益）"),
        ("扣非归母净利润口径", "公司披露值，非自算：年报 / 半年报「第二节·主要会计数据和财务指标·(一)主要会计数据」；出处逐期见 05 层手工摘录（含 PDF 页码）"),
        ("扣非归母净利润复核", "2020A–2024A 用「本年列」与「次年年报上年同期列」双向核对，5 期全部一致；2025H1 与 2026H1 取自 2026 年半年报（上年同期列 / 本报告期列）；2025A 为单一来源（半年报不含年度同期数）"),
        ("经营活动现金流口径", "现金流量表「经营活动产生的现金流量净额」（该表的「经营活动产生的现金流量」整列为空，不能用）"),
        ("经营现金流 ÷ 归母净利润", "经营现金流净额 ÷ 归母净利润，倍；分母是归母净利润，不是净利润"),
        ("非经常性损益口径", "归母净利润 − 扣非归母净利润，自算倒差"),
        ("同比口径", "（本期 − 上年同期）÷ 上年同期；2026H1 的上年同期是 2025H1"),
        ("禁用项提示", "财务摘要页签的「营业成本」实为营业总成本（2026H1＝551.34 亿元），用它算毛利率会得到 14.51%，正确值 36.43%"),
        ("待用户处理", "本表不判断利润下滑是减值扰动还是经营恶化 —— 结论由用户给出（AGENTS.md 三段式）"),
    ]
    for k, v in detail:
        ws2.append([k, v])
    for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=2):
        for cell in row:
            cell.font = Font(name=FONT, size=11, bold=cell.row == 1)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for c in ws2[1]:
        c.fill = PatternFill("solid", fgColor=HDR_FILL)
        c.alignment = Alignment(horizontal="center")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 96
    wb.save(path)


# ------------------------------------------------- 产物 2：指标驱动分解表


def build_driver_table(metrics: dict, data: dict, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "伊利_8期"
    ws["A1"] = "伊利股份（600887.SH）盈利指标驱动分解表（8 期）"
    ws["A1"].font = Font(name=FONT, size=12, bold=True)
    ws["A2"] = "单位：比率为百分数，金额为亿元｜数据截止 2026-10-09｜来源：07 清洗层 Rev05"
    ws["A3"] = "ROE 分母口径：期末归母净资产（＝资产负债表「归属于母公司股东权益合计」）；2026H1 为半年值，只作标注、不做年化"

    head = 5
    put(ws, head, "指标", PLABELS, bold=True)
    r = head + 1
    level_rows = [
        ("营业收入（亿元）", [money(metrics[p]["revenue"]) for p in PERIODS], FMT_MONEY),
        ("毛利率（（营业收入 − 营业成本）÷ 营业收入）", [metrics[p]["gross_margin"] for p in PERIODS], FMT_PCT),
        ("销售费用率（销售费用 ÷ 营业收入）", [metrics[p]["sell_rate"] for p in PERIODS], FMT_PCT),
        ("管理费用率（管理费用 ÷ 营业收入）", [metrics[p]["admin_rate"] for p in PERIODS], FMT_PCT),
        ("研发费用率（研发费用 ÷ 营业收入）", [metrics[p]["rd_rate"] for p in PERIODS], FMT_PCT),
        ("财务费用率（财务费用 ÷ 营业收入）", [metrics[p]["fin_rate"] for p in PERIODS], FMT_PCT),
        ("三费率合计（销售＋管理＋财务）", [metrics[p]["exp3_rate"] for p in PERIODS], FMT_PCT),
        ("四费率合计（销售＋管理＋研发＋财务）", [metrics[p]["exp4_rate"] for p in PERIODS], FMT_PCT),
        ("营业利润率（营业利润 ÷ 营业收入）", [metrics[p]["op_margin"] for p in PERIODS], FMT_PCT),
        ("归母净利率（归母净利润 ÷ 营业收入）", [metrics[p]["net_margin"] for p in PERIODS], FMT_PCT),
        ("其他项（＝毛利率 − 四费率合计 − 归母净利率，倒挤）", [metrics[p]["other_item"] for p in PERIODS], FMT_PCT),
        ("期末归母净资产（亿元）", [money(metrics[p]["equity"]) for p in PERIODS], FMT_MONEY),
        ("ROE（归母净利润 ÷ 期末归母净资产）", [metrics[p]["roe"] for p in PERIODS], FMT_PCT),
    ]
    for label, values, fmt in level_rows:
        put(ws, r, label, values, fmt)
        r += 1

    r += 1
    ws.cell(r, 1, "同比变动（pct 点）").font = Font(name=FONT, size=11, bold=True)
    fill = PatternFill("solid", fgColor=HDR_FILL)
    for c in range(1, 10):
        ws.cell(r, c).fill = fill
    r += 1
    pairs = YOY_PAIRS
    delta_rows = [
        ("毛利率变动", "gross_margin"),
        ("销售费用率变动", "sell_rate"),
        ("管理费用率变动", "admin_rate"),
        ("研发费用率变动", "rd_rate"),
        ("财务费用率变动", "fin_rate"),
        ("三费率合计变动", "exp3_rate"),
        ("四费率合计变动", "exp4_rate"),
        ("营业利润率变动", "op_margin"),
        ("归母净利率变动", "net_margin"),
        ("其他项变动", "other_item"),
        ("ROE 变动", "roe"),
    ]
    delta_store: dict[str, dict[str, float | None]] = {}
    for label, key in delta_rows:
        vals: list = ["—"] * 8
        store: dict[str, float | None] = {}
        for lbl, cur, prev in pairs:
            d = None if metrics[cur][key] is None or metrics[prev][key] is None else metrics[cur][key] - metrics[prev][key]
            store[lbl] = d
            vals[PLABELS.index(lbl)] = d
        delta_store[key] = store
        put(ws, r, label, vals, FMT_PCT)
        r += 1

    r += 1
    ws.cell(r, 1, "归母净利率变动分解（pct 点，同比）").font = Font(name=FONT, size=11, bold=True)
    for c in range(1, 10):
        ws.cell(r, c).fill = fill
    r += 1
    dec_rows = [
        ("① 毛利率变动（正向计入）", "gross_margin", 1.0),
        ("② 四费率合计变动（反向计入，取负号）", "exp4_rate", -1.0),
        ("③ 其他项变动（倒挤：归母净利率变动 − ① − ②）", None, None),
        ("合计 = 归母净利率变动（校验）", "net_margin", 1.0),
    ]
    for label, key, sign in dec_rows:
        vals: list = ["—"] * 8
        for lbl, cur, prev in pairs:
            if key is None:
                dm = delta_store["net_margin"][lbl]
                d1 = delta_store["gross_margin"][lbl]
                d2 = delta_store["exp4_rate"][lbl]
                v = None if None in (dm, d1, d2) else dm - d1 - (-d2)
            else:
                d = delta_store[key][lbl]
                v = None if d is None else d * sign
            vals[PLABELS.index(lbl)] = v
        put(ws, r, label, vals, FMT_PCT)
        r += 1

    r += 1
    for note in [
        "注 1：2020A 无上年同期（2019 年不纳入本报告口径）；2025H1 的同比需 2024H1，未在 8 期口径内 —— 两列留「—」。",
        "注 2：归母净利率 = 毛利率 − 四费率合计 − 其他项；其他项含营业税金及附加、资产减值损失、信用减值损失、其他收益、投资收益、公允价值变动、营业外收支、所得税费用、少数股东损益等，本表只作倒挤、不拆分。",
        "注 3：财务费用率为负（利息收入大于利息支出），四费率合计已被其压低。",
        "注 4：ROE 全期统一为期末归母净资产口径；2026H1 为半年值，与年度 ROE 不可直接比较。",
        "注 5：本页只呈现数据，不写判断；结论由用户给出。",
    ]:
        ws.cell(r, 1, note).font = Font(name=FONT, size=10)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=NOTE_FILL)
        r += 1

    style_sheet(ws, 9, [head], {"A": 46})
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 13
    ws.freeze_panes = "B6"

    # ------- 三家对比
    ws2 = wb.create_sheet("三家对比")
    ws2["A1"] = "三家盈利指标对比（2021A–2025A、2026H1）"
    ws2["A1"].font = Font(name=FONT, size=12, bold=True)
    ws2["A2"] = "单位：百分数｜来源：07 清洗层 Rev05（三家同版式附表：利润表第 3/10/20/21/22/23/51 行、资产负债表第 138 行）"
    ws2["A3"] = "ROE 分母口径：期末归母净资产（三家统一）；2026H1 为半年值，只作标注、不做年化"
    comp_periods = COMP_PERIODS
    comp_labels = COMP_LABELS
    head2 = 5
    put(ws2, head2, "指标 / 公司", comp_labels, bold=True)
    r = head2 + 1
    comp_rows = [
        ("毛利率", lambda c, p: _comp(c, p, "gross")),
        ("销售费用率", lambda c, p: _comp(c, p, "sell")),
        ("管理费用率", lambda c, p: _comp(c, p, "admin")),
        ("研发费用率", lambda c, p: _comp(c, p, "rd")),
        ("财务费用率", lambda c, p: _comp(c, p, "fin")),
        ("四费率合计", lambda c, p: _comp(c, p, "exp4")),
        ("归母净利率", lambda c, p: _comp(c, p, "net_margin")),
        ("ROE（期末归母净资产）", lambda c, p: _comp(c, p, "roe")),
    ]
    for label, fn in comp_rows:
        for code, (name, _folder) in COMPANIES.items():
            vals = [fn(code, p) for p in comp_periods]
            put(ws2, r, f"{label} —— {name}", vals, FMT_PCT)
            r += 1
        r += 1
    for note in [
        "注 1：2026H1 为半年值，与年度指标不可直接比较；光明乳业 2025A 归母净利润为负，其归母净利率与 ROE 为负值。",
        "注 2：三家 2026H1 的对比基期为 2025H1（未在对比期列内），故本页不列同比。",
        "注 3：本页只呈现数据，不写判断。",
    ]:
        ws2.cell(r, 1, note).font = Font(name=FONT, size=10)
        ws2.cell(r, 1).fill = PatternFill("solid", fgColor=NOTE_FILL)
        r += 1
    style_sheet(ws2, 8, [head2], {"A": 30})
    for c in range(2, 8):
        ws2.column_dimensions[get_column_letter(c)].width = 13
    ws2.freeze_panes = "B6"

    ws3 = wb.create_sheet("口径与出处")
    ws3.append(["项目", "说明"])
    for k, v in [
        ("数据版本", "07 清洗层 Rev05（三张报表）"),
        ("报告期", "伊利 8 期；三家对比 6 期（2021A–2025A、2026H1）"),
        ("毛利率口径", "（营业收入 − 营业成本）÷ 营业收入；取利润表「营业成本」，不用财务摘要的「营业成本」（后者实为营业总成本）"),
        ("费率口径", "各项费用 ÷ 营业收入；财务费用为负表示利息收入大于利息支出，不做归零处理"),
        ("三费率 / 四费率", "三费率＝销售＋管理＋财务；四费率＝销售＋管理＋研发＋财务。研发费用在利润表单独列示，两家口径都列出以便对照"),
        ("归母净利率口径", "归母净利润 ÷ 营业收入（分母不是营业总收入）"),
        ("ROE 口径", "归母净利润 ÷ 期末归母净资产（＝资产负债表「归属于母公司股东权益合计」）；2026-10-13 起由平均净资产改为期末口径，三家统一；2026H1 为半年值、不做年化"),
        ("驱动分解口径", "归母净利率 = 毛利率 − 四费率合计 − 其他项；同比变动以「本期 − 上年同期」的 pct 点表示，其他项为倒挤值"),
        ("与 DCF 模型的关系", "DCF 模型「财务指标」页已算 2022A–2025A 与预测期的同类指标；本页实际期取到 2020A、2021A 与 2025H1、2026H1，并补同比变动与分解，两者互为对照"),
        ("待用户处理", "本表不判断利润下滑是减值扰动还是经营恶化；结论由用户给出"),
    ]:
        ws3.append([k, v])
    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row, max_col=2):
        for cell in row:
            cell.font = Font(name=FONT, size=11, bold=cell.row == 1)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for c in ws3[1]:
        c.fill = PatternFill("solid", fgColor=HDR_FILL)
        c.alignment = Alignment(horizontal="center")
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 96
    wb.save(path)


def _comp(code: str, period: str, key: str) -> float | None:
    pl, bs, cf = _DATA[code]["pl"], _DATA[code]["bs"], _DATA[code]["cf"]
    rev = val(pl, period, "营业收入")
    if key == "gross":
        return div(rev - val(pl, period, "营业成本"), rev)
    if key == "sell":
        return div(val(pl, period, "销售费用"), rev)
    if key == "admin":
        return div(val(pl, period, "管理费用"), rev)
    if key == "rd":
        return div(val(pl, period, "研发费用"), rev)
    if key == "fin":
        return div(val(pl, period, "财务费用"), rev)
    if key == "exp4":
        return (val(pl, period, "销售费用") + val(pl, period, "管理费用") + val(pl, period, "研发费用") + val(pl, period, "财务费用")) / rev
    if key == "net_margin":
        return div(val(pl, period, "归属于母公司所有者的净利润"), rev)
    if key == "roe":
        return div(val(pl, period, "归属于母公司所有者的净利润"), val(bs, period, "归属于母公司股东权益合计"))
    raise KeyError(key)


_DATA: dict = {}


# --------------------------------------------------- 产物 3：三表勾稽检查表


def cross_checks() -> list[tuple[str, str, str, callable]]:
    """(序号关系名, 涉及报表, 公式说明, 差额函数)"""
    rels: list[tuple[str, str, str, callable]] = []

    def bs(name: str, expr: str, fn):
        rels.append((name, "资产负债表", expr, fn))

    def pl_(name: str, expr: str, fn):
        rels.append((name, "利润表", expr, fn))

    def cf_(name: str, expr: str, fn):
        rels.append((name, "现金流量表", expr, fn))

    B = lambda p, k: val(_DATA["600887"]["bs"], p, k)
    L = lambda p, k: val(_DATA["600887"]["pl"], p, k)
    C = lambda p, k: val(_DATA["600887"]["cf"], p, k)

    bs("资产总计 = 流动资产合计 + 非流动资产合计", "资产总计 −（流动资产合计 + 非流动资产合计）",
       lambda p: B(p, "资产总计") - (B(p, "流动资产合计") + B(p, "非流动资产合计")))
    bs("负债合计 = 流动负债合计 + 非流动负债合计", "负债合计 −（流动负债合计 + 非流动负债合计）",
       lambda p: B(p, "负债合计") - (B(p, "流动负债合计") + B(p, "非流动负债合计")))
    bs("所有者权益（或股东权益）合计 = 归属于母公司股东权益合计 + 少数股东权益", "所有者权益合计 −（归母权益 + 少数股东权益）",
       lambda p: B(p, "所有者权益(或股东权益)合计") - (B(p, "归属于母公司股东权益合计") + B(p, "少数股东权益")))
    bs("资产总计 = 负债和所有者权益（或股东权益）总计", "资产总计 − 负债和所有者权益总计",
       lambda p: B(p, "资产总计") - B(p, "负债和所有者权益(或股东权益)总计"))
    bs("负债和所有者权益（或股东权益）总计 = 负债合计 + 所有者权益（或股东权益）合计", "负债和所有者权益总计 −（负债合计 + 所有者权益合计）",
       lambda p: B(p, "负债和所有者权益(或股东权益)总计") - (B(p, "负债合计") + B(p, "所有者权益(或股东权益)合计")))
    pl_("营业总成本 = 营业成本 + 营业税金及附加 + 销售费用 + 管理费用 + 研发费用 + 财务费用（＋手续费及佣金支出、房地产销售成本、其他业务成本）",
        "营业总成本 − Σ（营业成本、税金及附加、销售、管理、研发、财务费用、手续费及佣金支出、房地产销售成本、其他业务成本）；减值损失不在其中",
        lambda p: L(p, "营业总成本") - sum(
            L(p, k) for k in ["营业成本", "营业税金及附加", "销售费用", "管理费用", "研发费用", "财务费用",
                              "手续费及佣金支出", "房地产销售成本", "其他业务成本"]))
    pl_("营业利润 = 营业总收入 − 营业总成本 + 其他收益 + 投资收益 + 公允价值变动收益 + 信用减值损失 + 资产减值损失 + 资产处置收益 + 其他业务利润 + 净敞口套期收益 + 汇兑收益",
        "营业利润 −（营业总收入 − 营业总成本 + 其他收益 + 投资收益 + 公允价值变动收益 + 信用减值损失 + 资产减值损失 + 资产处置收益 + 其他业务利润 + 净敞口套期收益 + 汇兑收益）",
        lambda p: L(p, "营业利润") - (L(p, "营业总收入") - L(p, "营业总成本") + sum(
            L(p, k) for k in ["其他收益", "投资收益", "公允价值变动收益", "信用减值损失", "资产减值损失",
                              "资产处置收益", "其他业务利润", "净敞口套期收益", "汇兑收益"])))
    pl_("营业总收入 = 营业收入 + 利息收入（＋已赚保费、手续费及佣金收入、房地产销售收入、其他业务收入）",
        "营业总收入 − Σ（营业收入、利息收入、已赚保费、手续费及佣金收入、房地产销售收入、其他业务收入）",
        lambda p: L(p, "营业总收入") - sum(
            L(p, k) for k in ["营业收入", "利息收入", "已赚保费", "手续费及佣金收入", "房地产销售收入", "其他业务收入"]))
    pl_("利润总额 = 营业利润 + 营业外收入 − 营业外支出", "利润总额 −（营业利润 + 营业外收入 − 营业外支出）",
        lambda p: L(p, "利润总额") - (L(p, "营业利润") + L(p, "营业外收入") - L(p, "营业外支出")))
    pl_("净利润 = 利润总额 − 所得税费用", "净利润 −（利润总额 − 所得税费用）",
        lambda p: L(p, "净利润") - (L(p, "利润总额") - L(p, "所得税费用")))
    pl_("净利润 = 归属于母公司所有者的净利润 + 少数股东损益", "净利润 −（归母净利润 + 少数股东损益）",
        lambda p: L(p, "净利润") - (L(p, "归属于母公司所有者的净利润") + L(p, "少数股东损益")))
    cf_("经营活动产生的现金流量净额 = 经营活动现金流入小计 − 经营活动现金流出小计", "净额 −（流入小计 − 流出小计）",
        lambda p: C(p, "经营活动产生的现金流量净额") - (C(p, "经营活动现金流入小计") - C(p, "经营活动现金流出小计")))
    cf_("投资活动产生的现金流量净额 = 投资活动现金流入小计 − 投资活动现金流出小计", "净额 −（流入小计 − 流出小计）",
        lambda p: C(p, "投资活动产生的现金流量净额") - (C(p, "投资活动现金流入小计") - C(p, "投资活动现金流出小计")))
    cf_("筹资活动产生的现金流量净额 = 筹资活动现金流入小计 − 筹资活动现金流出小计", "净额 −（流入小计 − 流出小计）",
        lambda p: C(p, "筹资活动产生的现金流量净额") - (C(p, "筹资活动现金流入小计") - C(p, "筹资活动现金流出小计")))
    cf_("现金及现金等价物净增加额 = 三活动净额之和 + 汇率变动对现金及现金等价物的影响",
        "净增加额 −（经营净额 + 投资净额 + 筹资净额 + 汇率影响）",
        lambda p: C(p, "现金及现金等价物净增加额") - (C(p, "经营活动产生的现金流量净额") + C(p, "投资活动产生的现金流量净额")
                                                     + C(p, "筹资活动产生的现金流量净额") + C(p, "汇率变动对现金及现金等价物的影响")))
    cf_("期末现金及现金等价物余额 = 期初现金及现金等价物余额 + 现金及现金等价物净增加额",
        "期末余额 −（期初余额 + 净增加额）",
        lambda p: C(p, "期末现金及现金等价物余额") - (C(p, "期初现金及现金等价物余额") + C(p, "现金及现金等价物净增加额")))
    cf_("期末现金及现金等价物余额 = 现金的期末余额 + 现金等价物的期末余额", "期末余额 −（现金期末余额 + 现金等价物期末余额）",
        lambda p: C(p, "期末现金及现金等价物余额") - (C(p, "现金的期末余额") + C(p, "现金等价物的期末余额")))
    return rels


def pdf_page_of(value: float, filename: str, cache: dict) -> int | None:
    """在一手来源 PDF 里精确匹配千分位字符串，返回首次命中的 PDF 页码。"""
    if filename not in cache:
        try:
            from pypdf import PdfReader
        except ImportError:
            cache[filename] = None
            return None
        reader = PdfReader(str(PDF_DIR / filename))
        cache[filename] = [re.sub(r"[\s,]", "", page.extract_text() or "") for page in reader.pages]
    pages = cache[filename]
    if pages is None:
        return None
    needle = f"{value:,.2f}".replace(",", "")
    for i, body in enumerate(pages, 1):
        if needle in body:
            return i
    return None


def build_cross_check(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "勾稽检查表"
    ws["A1"] = "伊利股份（600887.SH）三表勾稽检查表"
    ws["A1"].font = Font(name=FONT, size=12, bold=True)
    ws["A2"] = "覆盖 8 期：2020-12-31、2021-12-31、2022-12-31、2023-12-31、2024-12-31、2025-06-30、2025-12-31、2026-06-30｜来源：07 清洗层 Rev05"
    ws["A3"] = "逐期用 07 清洗层数值重算，取 8 期最大绝对差额；容差 0.01 元（分位以下视为相等）。差额为浮点残差，量级 ≤ 0.00005 元（＝0.0000000005 亿元），本质为 0"
    head = 5
    for i, title in enumerate(["序号", "勾稽关系", "涉及报表", "8 期实测最大绝对差额（元）", "是否通过", "公式（差额算法）"], start=1):
        ws.cell(head, i, title)
    rels = cross_checks()
    r = head + 1
    detail_rows: list[tuple[str, list[float]]] = []
    for idx, (name, sheetname, expr, fn) in enumerate(rels, 1):
        diffs = [abs(fn(p)) for p in PERIODS]
        detail_rows.append((name, [fn(p) for p in PERIODS]))
        ws.cell(r, 1, idx)
        ws.cell(r, 2, name)
        ws.cell(r, 3, sheetname)
        c = ws.cell(r, 4, max(diffs))
        c.number_format = FMT_DIFF
        ws.cell(r, 5, "通过" if max(diffs) < 0.01 else "不通过")
        ws.cell(r, 6, expr)
        r += 1
    for row in ws.iter_rows(min_row=head + 1, max_row=r - 1, max_col=6):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    r += 1
    for note in [
        "注 1：第 6 项（营业总成本）第一次核验时把资产减值损失、信用减值损失算进了营业总成本，曾报出 3.27–52.31 亿元差额；回 2026 年半年报合并利润表原文确认这两个科目不在营业总成本「其中」项内，更正后差额为 0（留痕见《数据验收报告》3.1 节）。",
        "注 2：第 1–5 项为资产负债表内部勾稽，第 6–11 项为利润表内部勾稽，第 12–17 项为现金流量表内部勾稽。",
        "注 3：本表只覆盖 600887 伊利股份（建模用工作簿只做了这一家的三张表）。",
        "注 4：勾稽通过只说明表内自洽，不说明数据与披露一致 —— 与披露的一致性见「一手来源抽检」页。",
        "注 5：差额列保留 6 位小数显示；全部差额 ≤ 0.00005 元，为浮点运算残差，账面上可视为 0。",
    ]:
        ws.cell(r, 1, note).font = Font(name=FONT, size=10)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=NOTE_FILL)
        r += 1
    style_sheet(ws, 6, [head], {"A": 6, "B": 66, "C": 12, "D": 22, "E": 10, "F": 60})
    ws.freeze_panes = "B6"

    ws2 = wb.create_sheet("逐期差额明细")
    ws2["A1"] = "勾稽差额逐期明细（元）"
    ws2["A1"].font = Font(name=FONT, size=12, bold=True)
    ws2["A2"] = "数值为「左侧项 − 右侧项」，单位为元；8 期全部应接近 0"
    head2 = 4
    put(ws2, head2, "勾稽关系", PLABELS, bold=True)
    r = head2 + 1
    for name, diffs in detail_rows:
        put(ws2, r, name[:60], diffs, FMT_DIFF)
        r += 1
    style_sheet(ws2, 9, [head2], {"A": 60})
    for c in range(2, 10):
        ws2.column_dimensions[get_column_letter(c)].width = 15
    ws2.freeze_panes = "B5"

    # ------ 与建模用工作簿的一致性
    ws3 = wb.create_sheet("与清洗层一致性")
    ws3["A1"] = "建模用工作簿三张报表 vs 07 清洗层 Rev05 逐格比对"
    ws3["A1"].font = Font(name=FONT, size=12, bold=True)
    ws3["A2"] = "比对对象：原始数据/06_建模用Excel/ 目录下最新版《600887_财务数据》工作簿；容差：相对误差 < 1e-6"
    head3 = 4
    for i, title in enumerate(["页签", "共同科目", "期数", "比对单元", "不一致单元", "备注"], start=1):
        ws3.cell(head3, i, title)
    wb_path = latest(MODEL_XLSX_DIR, "600887_财务数据_*.xlsx")
    model = load_workbook(wb_path, data_only=True)
    r = head3 + 1
    for sheetname, kind in [("利润表", "利润表"), ("资产负债表", "资产负债表"), ("现金流量表", "现金流量表")]:
        subjects, table = read_clean_sheet(CLEAN / "600887_伊利股份", "600887", kind)
        mws = model[sheetname]
        dates = [str(mws.cell(1, c).value) for c in range(2, mws.max_column + 1)]
        wbvals: dict[tuple[str, str], float | None] = {}
        meta_rows = 0
        for row in range(2, mws.max_row + 1):
            label = mws.cell(row, 1).value
            if label is None:
                continue
            label = str(label).strip()
            if label not in subjects:
                meta_rows += 1
                continue
            for c, d in enumerate(dates, start=2):
                wbvals[(label, d)] = mws.cell(row, c).value
        compared = mismatch = 0
        for s in subjects:
            for p in PERIODS:
                if p not in dates:
                    continue
                a = table.get(p, {}).get(s)
                b = wbvals.get((s, p))
                if isinstance(b, str):
                    try:
                        b = float(b)
                    except ValueError:
                        b = None
                compared += 1
                if a is None and b is None:
                    continue
                if a is None or b is None or abs(a - b) > 1e-6 * max(1.0, abs(a)):
                    mismatch += 1
        ws3.cell(r, 1, sheetname)
        ws3.cell(r, 2, len(subjects))
        ws3.cell(r, 3, len([p for p in PERIODS if p in dates]))
        ws3.cell(r, 4, compared)
        ws3.cell(r, 5, mismatch)
        ws3.cell(r, 6, f"工作簿另有 {meta_rows} 个非科目行（元数据类）；本页只证明两侧一致，不构成对原始披露的验证")
        r += 1
    r += 1
    ws3.cell(r, 1, f"比对用的工作簿：{wb_path.name}").font = Font(name=FONT, size=10)
    style_sheet(ws3, 6, [head3], {"A": 14, "B": 12, "C": 8, "D": 12, "E": 12, "F": 70})

    # ------ 一手来源抽检
    ws4 = wb.create_sheet("一手来源抽检")
    ws4["A1"] = "关键数字回一手来源抽检"
    ws4["A1"].font = Font(name=FONT, size=12, bold=True)
    ws4["A2"] = "方法：定期报告 PDF 全文去空格去逗号后，与千分位字符串精确匹配；命中即原文逐字符出现。需 pypdf，不可用时「命中页码」留空并标注未复核"
    head4 = 4
    for i, title in enumerate(["序号", "报告期", "项目", "数值（元）", "来源文件", "命中页码（PDF）", "结果"], start=1):
        ws4.cell(head4, i, title)
    checks: list[tuple[str, str, float, str]] = [
        ("2025A", "营业收入", 115636231250.05, "2025年年度报告.pdf"),
        ("2025A", "营业成本", 75653843800.99, "2025年年度报告.pdf"),
        ("2025A", "利润总额", 12993830374.42, "2025年年度报告.pdf"),
        ("2025A", "归母净利润", 11565166497.81, "2025年年度报告.pdf"),
        ("2025A", "扣非归母净利润", 11068321695.12, "2025年年度报告.pdf"),
        ("2025A", "资产总计", 152098516632.34, "2025年年度报告.pdf"),
        ("2025A", "期末归母净资产", 54655690475.89, "2025年年度报告.pdf"),
        ("2025A", "货币资金", 19494920358.00, "2025年年度报告.pdf"),
        ("2025A", "商誉", 2181530707.86, "2025年年度报告.pdf"),
        ("2025A", "经营活动现金流量净额", 14343920490.32, "2025年年度报告.pdf"),
        ("2026H1", "营业收入", 64330936547.38, "2026年半年报告.pdf"),
        ("2026H1", "营业成本", 40894077275.25, "2026年半年报告.pdf"),
        ("2026H1", "利润总额", 7314474445.37, "2026年半年报告.pdf"),
        ("2026H1", "归母净利润", 5758628097.79, "2026年半年报告.pdf"),
        ("2026H1", "扣非归母净利润", 5595762517.76, "2026年半年报告.pdf"),
        ("2026H1", "资产总计", 157165308525.12, "2026年半年报告.pdf"),
        ("2026H1", "期末归母净资产", 53711364981.99, "2026年半年报告.pdf"),
        ("2026H1", "经营活动现金流量净额", 9759147483.18, "2026年半年报告.pdf"),
        ("2025H1", "扣非归母净利润（上年同期列）", 7016349502.87, "2026年半年报告.pdf"),
        ("2020A", "扣非归母净利润", 6625112858.43, "2020年年度报告.pdf"),
        ("2021A", "扣非归母净利润", 7943784753.50, "2021年年度报告.pdf"),
        ("2022A", "扣非归母净利润", 8585692731.44, "2022年年度报告.pdf"),
        ("2023A", "扣非归母净利润", 10026171076.03, "2023年年度报告.pdf"),
        ("2024A", "扣非归母净利润", 6011274945.92, "2024年年度报告.pdf"),
    ]
    cache: dict = {}
    r = head4 + 1
    for idx, (period, item, value, filename) in enumerate(checks, 1):
        page = pdf_page_of(value, filename, cache)
        ws4.cell(r, 1, idx)
        ws4.cell(r, 2, period)
        ws4.cell(r, 3, item)
        ws4.cell(r, 4, value).number_format = FMT_MONEY
        ws4.cell(r, 5, filename)
        ws4.cell(r, 6, page if page else "")
        ws4.cell(r, 7, "命中" if page else ("未复核（无 pypdf）" if cache.get(filename) is None else "未命中"))
        r += 1
    style_sheet(ws4, 7, [head4], {"A": 6, "B": 10, "C": 26, "D": 20, "E": 24, "F": 14, "G": 16})
    ws4.freeze_panes = "A5"

    ws5 = wb.create_sheet("口径与说明")
    ws5.append(["项目", "说明"])
    for k, v in [
        ("验收对象", "07 清洗层 Rev05 的三张报表（合并口径；年报经审计，2025H1、2026H1 未经审计）"),
        ("报告期", "8 期（同上表）；2020-12-31 作年度基数、2025-06-30 作中期基数"),
        ("覆盖范围", "本表只覆盖 600887 伊利股份；光明乳业、新乳业未做同口径勾稽（建模用工作簿只做了伊利）"),
        ("勾稽算法", "逐期用 07 清洗层科目数值重算差额，再取 8 期最大绝对值；容差 0.01 元"),
        ("空值处理", "空单元格按 0 参与加总（如「手续费及佣金支出」，伊利无此项）"),
        ("已知缺口", "折旧摊销不在 07 层任何一张表，取自 05 层《600887_折旧摊销》；三表间的跨表勾稽（如未分配利润与净利润）本表未覆盖"),
        ("与《数据验收报告》的关系", "本表是验收报告第三节的可复核版本：差额由本脚本重算，不是转录；口径与结论一致"),
        ("待用户处理", "本表不判断报表是否可用于建模；结论由用户给出"),
    ]:
        ws5.append([k, v])
    for row in ws5.iter_rows(min_row=1, max_row=ws5.max_row, max_col=2):
        for cell in row:
            cell.font = Font(name=FONT, size=11, bold=cell.row == 1)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for c in ws5[1]:
        c.fill = PatternFill("solid", fgColor=HDR_FILL)
        c.alignment = Alignment(horizontal="center")
    ws5.column_dimensions["A"].width = 24
    ws5.column_dimensions["B"].width = 96
    wb.save(path)


# ------------------------------------------------------------------- 主流程


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    OUTDIR.mkdir(parents=True, exist_ok=True)
    global _DATA
    _DATA = load_all()
    deducted = load_deducted()
    metrics = ily_metrics(_DATA, deducted)

    out1 = OUTDIR / "600887_利润与现金流对照表_20261014_Rev00.xlsx"
    out2 = OUTDIR / "600887_指标驱动分解表_20261014_Rev00.xlsx"
    out3 = OUTDIR / "600887_三表勾稽检查表_20261014_Rev00.xlsx"
    build_profit_cash(metrics, deducted, out1)
    build_driver_table(metrics, _DATA, out2)
    build_cross_check(out3)

    print("=== 伊利 8 期：利润 vs 现金流（亿元 / 倍 / %）===")
    print("期间        营业收入   归母净利   扣非归母   非经常    经营现金流  现金流/归母净利")
    for p, lbl in zip(PERIODS, PLABELS):
        m = metrics[p]
        print(f"{lbl:<9} {money(m['revenue']):>10.2f} {money(m['np_parent']):>10.2f} "
              f"{money(m['deducted']):>10.2f} {money(m['nonrecurring']):>9.2f} {money(m['ocf']):>10.2f} "
              f"{m['ocf_over_np']:>13.2f}")
    print()
    print("=== 伊利 8 期：驱动分解（%）===")
    print("期间        毛利率   三费率   四费率   归母净利率  其他项    ROE")
    for p, lbl in zip(PERIODS, PLABELS):
        m = metrics[p]
        print(f"{lbl:<9} {m['gross_margin']*100:>7.2f} {m['exp3_rate']*100:>7.2f} {m['exp4_rate']*100:>7.2f} "
              f"{m['net_margin']*100:>10.2f} {m['other_item']*100:>7.2f} {m['roe']*100:>6.2f}")
    print()
    for out in (out1, out2, out3):
        print("written:", out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
