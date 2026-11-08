#!/usr/bin/env python3
"""伊利股份（600887.SH）数据流水线：取数 -> 统一日期 -> 筛选 -> 导出 CSV + Excel。

一条命令跑完全流程，不依赖 notebook。

用法：
    python 600887_pipeline.py
python 600887_pipeline.py --out "..\\..\\原始数据"
    python 600887_pipeline.py --rev Rev01

产出（默认输出到项目的「原始数据」目录）：
    00_索引/export_manifest_<YYYYMMDD>_<rev>.json   导出清单
    01_报表原表/600887_伊利股份/600887_<内容>_<YYYYMMDD>_<rev>.csv
    02_行情数据/600887_伊利股份/600887_行情_<YYYYMMDD>_<rev>.csv
    03_估值数据/600887_伊利股份/600887_估值*_<YYYYMMDD>_<rev>.csv
    06_建模用Excel/600887_财务数据_<YYYYMMDD>_<rev>.xlsx   共 7 个页签

版本约定：文件名带日期与版本号，重出请**提高 --rev**（Rev01、Rev02……）。
**旧版本一律保留，不覆盖、不删除。**

Excel 页签：
    利润表 / 资产负债表 / 现金流量表 / 财务指标
        —— 已转置为「科目为行、报告期为列」，可直接搭 DCF
    财务摘要        —— 指标为行、报告期为列
    估值日频        —— 含总股本、市值，供 WACC 使用
    口径说明        —— 口径与陷阱写入文件本身

数据源（接口名一并写入 manifest）：
    三张报表     stock_financial_report_sina         新浪
    财务摘要     stock_financial_abstract            新浪
    财务指标     stock_financial_analysis_indicator  新浪
    行情         stock_zh_a_daily                    新浪，前复权，2021-01-01 起
    估值日频     stock_value_em                      东方财富
    估值十年     stock_zh_valuation_baidu            百度，PE(TTM) 与 PB 分开落盘

硬性约束：
    - 任一数据源失败即中止，不导出部分数据。
    - 导出前用「一手锚点」校验关键数字（来自 2026 年半年报），对不上就中止。
    - 同名文件不覆盖，跳过并在 manifest 里记为 skipped_existing。
    - 不做单位换算；金额单位保持「元」。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import pandas as pd

SYMBOL = "600887"
SINA_PREFIXED = "sh600887"

# 报告期共 8 期。两个基数不能省：
#   20201231 —— 年度基数，用于算 2021 年同比增速与 ROE 期初净资产
#   20250630 —— 中期基数，用于算 2026H1 同比增速（中期只能与中期比，不能与年度比）
PERIODS = ["20201231", "20211231", "20221231", "20231231",
           "20241231", "20250630", "20251231", "20260630"]

MARKET_CUTOFF = "2026-10-09"   # 市场数据（行情、估值）截止日
PRICE_START = "20210101"

# 数据分层：内容 -> 一级目录；公司名下再分一层
COMPANY_DIR = "600887_伊利股份"
LAYER_BY_CONTENT = {
    "资产负债表": "01_报表原表", "利润表": "01_报表原表", "现金流量表": "01_报表原表",
    "财务摘要": "01_报表原表", "财务摘要_全量": "01_报表原表", "财务指标": "01_报表原表",
    "行情": "02_行情数据",
    "估值": "03_估值数据", "估值十年PE": "03_估值数据", "估值十年PB": "03_估值数据",
}
EXCEL_LAYER = "06_建模用Excel"
INDEX_LAYER = "00_索引"
PER_COMPANY_LAYERS = ("01_报表原表", "02_行情数据", "03_估值数据")


def layer_dir(out_dir: Path, layer: str) -> Path:
    """返回分层目录并确保存在；报表/行情/估值三层再按公司分一层"""
    p = out_dir / layer
    if layer in PER_COMPANY_LAYERS:
        p = p / COMPANY_DIR
    p.mkdir(parents=True, exist_ok=True)
    return p


# 一手锚点：来源为 2026 年半年度报告（p6 / p10），单位亿元
# (数据框, 字段, 报告期, 期望值, 容差)
ANCHORS = [
    ("df_profit", "营业收入", "20260630", 643.31, 0.02),
    ("df_profit", "营业总收入", "20260630", 644.90, 0.02),
    ("df_profit", "营业成本", "20260630", 408.94, 0.02),
    ("df_cashflow", "经营活动产生的现金流量净额", "20260630", 97.59, 0.02),
    ("df_balance", "归属于母公司股东权益合计", "20260630", 537.11, 0.02),
]


# ---------------------------------------------------------------- 取数

def fetch_all(ak) -> tuple[dict, list[str]]:
    """拉取全部数据源；返回 (数据字典, 失败信息列表)。任一失败都会被记录。"""
    jobs = [
        ("df_balance", "资产负债表", "stock_financial_report_sina", "新浪",
         lambda: ak.stock_financial_report_sina(stock=SINA_PREFIXED, symbol="资产负债表")),
        ("df_profit", "利润表", "stock_financial_report_sina", "新浪",
         lambda: ak.stock_financial_report_sina(stock=SINA_PREFIXED, symbol="利润表")),
        ("df_cashflow", "现金流量表", "stock_financial_report_sina", "新浪",
         lambda: ak.stock_financial_report_sina(stock=SINA_PREFIXED, symbol="现金流量表")),
        ("financial_data", "财务指标", "stock_financial_analysis_indicator", "新浪",
         lambda: ak.stock_financial_analysis_indicator(symbol=SYMBOL, start_year="2020")),
        ("financial_abstract", "财务摘要", "stock_financial_abstract", "新浪",
         lambda: ak.stock_financial_abstract(symbol=SYMBOL)),
        ("daily_stock", "行情", "stock_zh_a_daily", "新浪（前复权）",
         lambda: ak.stock_zh_a_daily(symbol=SINA_PREFIXED, start_date=PRICE_START,
                                     end_date=MARKET_CUTOFF.replace("-", ""), adjust="qfq")),
        ("daily_value", "估值日频", "stock_value_em", "东方财富",
         lambda: ak.stock_value_em(symbol=SYMBOL)),
        ("pe_history", "估值十年PE", "stock_zh_valuation_baidu", "百度",
         lambda: ak.stock_zh_valuation_baidu(symbol=SYMBOL, indicator="市盈率(TTM)",
                                             period="近十年")),
        ("pb_history", "估值十年PB", "stock_zh_valuation_baidu", "百度",
         lambda: ak.stock_zh_valuation_baidu(symbol=SYMBOL, indicator="市净率",
                                             period="近十年")),
    ]

    data: dict[str, pd.DataFrame] = {}
    problems: list[str] = []
    sources: dict[str, dict] = {}

    for var, label, api, src, fn in jobs:
        try:
            df = fn()
            if df is None or len(df) == 0:
                raise ValueError("接口返回空表")
            data[var] = df
            sources[var] = {"label": label, "api": api, "source": src,
                            "rows": int(df.shape[0]), "cols": int(df.shape[1]),
                            "status": "ok"}
            print(f"  [OK]   {label:12} {api:36} {df.shape}")
        except Exception as exc:  # noqa: BLE001 - 逐个记录后统一中止
            msg = f"{label} ({api}) 取数失败：{type(exc).__name__}: {exc}"
            problems.append(msg)
            sources[var] = {"label": label, "api": api, "source": src,
                            "status": "failed", "error": f"{type(exc).__name__}: {exc}"}
            print(f"  [FAIL] {label:12} {api:36} {type(exc).__name__}")

    return {"data": data, "sources": sources}, problems


# ------------------------------------------------- 统一日期 + 筛选

def norm_date(df: pd.DataFrame, col: str, fmt: str | None = None) -> pd.DataFrame:
    """统一日期列：转 datetime64[ns]、改名为 date，并置于第一列"""
    df = df.copy()
    converted = pd.to_datetime(df[col], format=fmt)
    df = df.drop(columns=[col])
    df.insert(0, "date", converted)
    return df


def filter_periods(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["date"].isin(pd.to_datetime(PERIODS, format="%Y%m%d"))].reset_index(drop=True)


def filter_daily(df: pd.DataFrame) -> pd.DataFrame:
    start = pd.Timestamp(PRICE_START)
    end = pd.Timestamp(MARKET_CUTOFF)
    return df[(df["date"] >= start) & (df["date"] <= end)].reset_index(drop=True)


def clean(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    d = data
    d["df_balance"] = filter_periods(norm_date(d["df_balance"], "报告日", fmt="%Y%m%d"))
    d["df_profit"] = filter_periods(norm_date(d["df_profit"], "报告日", fmt="%Y%m%d"))
    d["df_cashflow"] = filter_periods(norm_date(d["df_cashflow"], "报告日", fmt="%Y%m%d"))
    d["financial_data"] = filter_periods(norm_date(d["financial_data"], "日期"))
    d["daily_stock"] = filter_daily(norm_date(d["daily_stock"], "date"))
    d["daily_value"] = filter_daily(norm_date(d["daily_value"], "数据日期"))
    d["pe_history"] = norm_date(d["pe_history"], "date")
    d["pb_history"] = norm_date(d["pb_history"], "date")
    keep = [c for c in PERIODS if c in d["financial_abstract"].columns]
    d["fin_abs_wide"] = d["financial_abstract"][["选项", "指标"] + keep]
    d["keep"] = keep
    return d


def check_anchors(d: dict[str, pd.DataFrame]) -> list[str]:
    """用一手披露数字校验关键字段；返回问题列表（空列表表示通过）"""
    problems = []
    for var, field, period, expect, tol in ANCHORS:
        df = d[var]
        target = pd.Timestamp(period)
        row = df[df["date"] == target]
        if row.empty:
            problems.append(f"{var} 缺 {period} 报告期")
            continue
        if field not in df.columns:
            problems.append(f"{var} 缺字段「{field}」")
            continue
        actual = row.iloc[0][field] / 1e8
        if abs(actual - expect) > tol:
            problems.append(f"{var}.{field} @{period} = {actual:.2f} 亿元，"
                            f"应约 {expect} 亿元（中报口径）")
    return problems


# ---------------------------------------------------------------- 导出

def periods_as_columns(df: pd.DataFrame, index_name: str) -> pd.DataFrame:
    """报告期转置成列：科目为行、报告期为列，时间升序"""
    out = df.set_index("date").sort_index().T
    out.columns = [c.strftime("%Y-%m-%d") for c in out.columns]
    out.index = [str(i) for i in out.index]
    out.index.name = index_name
    return out


def export(d: dict[str, pd.DataFrame], out_dir: Path, stamp: str, rev: str) -> tuple[list, list]:
    """导出 CSV 与 Excel。同名文件不覆盖（旧版一律保留）；返回 (写入清单, 跳过清单)"""
    written, skipped = [], []

    def save_csv(df: pd.DataFrame, content: str) -> None:
        s = df.sort_values("date").reset_index(drop=True) if "date" in df.columns else df
        path = layer_dir(out_dir, LAYER_BY_CONTENT[content]) / f"600887_{content}_{stamp}_{rev}.csv"
        rel = path.relative_to(out_dir).as_posix()
        if path.exists():
            skipped.append(rel)
            return
        s.to_csv(path, index=False, encoding="utf-8-sig")
        written.append((rel, s.shape))

    for content, key in [("资产负债表", "df_balance"), ("利润表", "df_profit"),
                         ("现金流量表", "df_cashflow"), ("财务指标", "financial_data"),
                         ("财务摘要_全量", "financial_abstract"), ("财务摘要", "fin_abs_wide"),
                         ("行情", "daily_stock"), ("估值", "daily_value"),
                         ("估值十年PE", "pe_history"), ("估值十年PB", "pb_history")]:
        save_csv(d[key], content)

    fin_abs_model = d["fin_abs_wide"].copy()
    fin_abs_model.columns = ["选项", "指标"] + [
        f"{c[:4]}-{c[4:6]}-{c[6:]}" for c in d["keep"]]

    notes = pd.DataFrame([
        ["数据截止", f"市场数据（行情、估值）截至 {MARKET_CUTOFF}；财报最新期为 2026-06-30"],
        ["营收主口径", "营业收入（不含利息收入），2026H1 = 643.31 亿元"],
        ["对照口径", "营业总收入（含利息收入），2026H1 = 644.90 亿元；不作主口径"],
        ["报告期", "8 期：2020-12-31、2021-12-31、2022-12-31、2023-12-31、"
                   "2024-12-31、2025-06-30、2025-12-31、2026-06-30"],
        ["两个基数的用途", "2020-12-31 算 2021 年同比与 ROE 期初净资产；"
                            "2025-06-30 算 2026H1 同比。中期只能与中期比，不能与年度比"],
        ["金额单位", "元（未换算；报告正文统一用亿元）"],
        ["毛利率口径", "必须用利润表的「营业成本」。财务摘要的「营业成本」实为营业总成本，"
                       "用它算毛利率会得到 14.51%，正确值 36.43%"],
        ["经营现金流口径", "必须用「经营活动产生的现金流量净额」。"
                            "「经营活动产生的现金流量」那一列整列为空"],
        ["股本口径", "行情表的 outstanding_share 是流通股本（62.94 亿股）；"
                     "总股本 63.25 亿股见「估值日频」表"],
        ["日期格式", "统一为 datetime64[ns]，列名 date；导出文件按时间升序"],
        ["数据源", "新浪财经（三表、财务摘要、财务指标、行情）；东方财富（估值日频）；"
                   "百度（估值十年）"],
    ], columns=["项目", "说明"])

    xlsx = layer_dir(out_dir, EXCEL_LAYER) / f"600887_财务数据_{stamp}_{rev}.xlsx"
    xrel = xlsx.relative_to(out_dir).as_posix()
    if xlsx.exists():
        skipped.append(xrel)
    else:
        with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
            periods_as_columns(d["df_profit"], "科目").to_excel(xw, sheet_name="利润表")
            periods_as_columns(d["df_balance"], "科目").to_excel(xw, sheet_name="资产负债表")
            periods_as_columns(d["df_cashflow"], "科目").to_excel(xw, sheet_name="现金流量表")
            periods_as_columns(d["financial_data"], "指标").to_excel(xw, sheet_name="财务指标")
            fin_abs_model.to_excel(xw, sheet_name="财务摘要", index=False)
            d["daily_value"].to_excel(xw, sheet_name="估值日频", index=False)
            notes.to_excel(xw, sheet_name="口径说明", index=False)
        written.append((xrel, "7 sheets"))

    return written, skipped


def main() -> int:
    parser = argparse.ArgumentParser(description="伊利 600887 数据流水线")
    parser.add_argument("--out", default=str(
        Path(__file__).resolve().parents[2] / "原始数据"), help="输出目录")
    parser.add_argument("--rev", default="Rev00", help="版本号，默认 Rev00；重出请提高")
    parser.add_argument("--stamp", default=None, help="文件名日期，默认今天 YYYYMMDD")
    args = parser.parse_args()

    try:
        import akshare as ak
    except ImportError:
        print("[ERROR] 未安装 AKShare。请用 C:\\Python314\\python.exe 运行，"
              "或先 pip install akshare", file=sys.stderr)
        return 3

    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = args.stamp or dt.date.today().strftime("%Y%m%d")

    print(f"AKShare {getattr(ak, '__version__', '?')} | 输出目录 {out_dir}")
    print("\n[1/4] 取数")
    result, problems = fetch_all(ak)
    if problems:
        print("\n[中止] 有数据源取数失败，不导出部分数据：", file=sys.stderr)
        for msg in problems:
            print(f"  - {msg}", file=sys.stderr)
        return 1

    print("\n[2/4] 统一日期 + 筛选")
    d = clean(result["data"])
    for k in ["df_balance", "df_profit", "df_cashflow", "financial_data"]:
        print(f"  {k:16} {d[k].shape}")
    for k in ["daily_stock", "daily_value", "pe_history", "pb_history"]:
        print(f"  {k:16} {d[k].shape}  {d[k]['date'].min().date()} ~ {d[k]['date'].max().date()}")

    print("\n[3/4] 一手锚点校验（2026 年半年报）")
    anchor_problems = check_anchors(d)
    if anchor_problems:
        print("\n[中止] 关键数字与半年报对不上，不导出：", file=sys.stderr)
        for msg in anchor_problems:
            print(f"  - {msg}", file=sys.stderr)
        return 2
    for _, field, period, expect, _ in ANCHORS:
        print(f"  [OK] {field} @{period} = {expect} 亿元")

    print("\n[4/4] 导出")
    written, skipped = export(d, out_dir, stamp, args.rev)

    mpath = layer_dir(out_dir, INDEX_LAYER) / f"export_manifest_{stamp}_{args.rev}.json"
    mrel = mpath.relative_to(out_dir).as_posix()

    # files 记录本次涉及的全部文件（含已存在被跳过的），而不只是本次新写入的
    file_entries = [{"name": n, "status": "written",
                     "shape": list(s) if isinstance(s, tuple) else s}
                    for n, s in written]
    file_entries += [{"name": n, "status": "skipped_existing"} for n in skipped]
    file_entries.append({"name": mrel,
                         "status": "skipped_existing" if mpath.exists() else "written"})

    manifest = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "symbol": SYMBOL,
        "rev": args.rev,
        "periods": PERIODS,
        "market_data_cutoff": MARKET_CUTOFF,
        "price_start": PRICE_START,
        "sorted_ascending": True,
        "anchors_passed": [f"{v}.{f}@{p}={e}" for v, f, p, e, _ in ANCHORS],
        "sources": result["sources"],
        "excel_sheets": ["利润表", "资产负债表", "现金流量表", "财务指标",
                         "财务摘要", "估值日频", "口径说明"],
        "file_count": len(file_entries),
        "files": file_entries,
        "skipped_existing": skipped,
    }
    if mpath.exists():
        skipped.append(mrel)
    else:
        mpath.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        written.append((mrel, "-"))

    for n, s in written:
        print(f"  [写入] {n:58} {s}")
    for n in skipped:
        print(f"  [跳过·同名已存在，旧版保留] {n}")
    print(f"\n完成：写入 {len(written)} 个文件，跳过 {len(skipped)} 个")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
