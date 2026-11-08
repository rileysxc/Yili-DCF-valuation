#!/usr/bin/env python3
"""投研项目取数脚本（AKShare，A 股）。

拉取标的的财务报表、财务指标、日行情与估值历史，按
「<代码>_<内容>_<YYYYMMDD>_<rev>.csv」命名落盘，并生成 fetch_manifest。

用法：
    python fetch_stock_data.py --symbols 600597,002946 --out "<原始数据目录>"
    python fetch_stock_data.py --symbols 600597 --rev Rev01 --out "<原始数据目录>"

数据源（接口名一并写入 manifest）：
    财务摘要     stock_financial_abstract            新浪
    财务指标     stock_financial_analysis_indicator  新浪
    三大报表     stock_financial_report_sina         新浪
    行情         stock_zh_a_daily                    新浪，前复权，默认回溯 10 年
    估值         stock_value_em                      东方财富，日频（PE(TTM)/PE(静)/PB/市值/股本）
    估值十年PE   stock_zh_valuation_baidu            百度（PE(TTM)）
    估值十年PB   stock_zh_valuation_baidu            百度（PB）

注意：百度估值接口返回点数有上限，密度随 period 变化：近三年≈日频、近五年≈隔日、
近十年≈隔日、全部≈月度。因此「估值十年」不是日频，样本区间与频率必须写进
DATA_NOTES，报告里写分位数时要一并说明。另外该接口的 PE 与 PB 采样日期互不相同，
合并会造成大量空值，故 PE、PB 分开落盘。

约定：
    - 只取数落盘，不做单位换算与口径统一。
    - 按数据层落盘：报表原表 / 行情数据 / 估值数据（每层再按「代码_简称」分公司），
      manifest 写入 00_索引。公司简称自动从 AKShare 的 A 股名称表解析，取不到则用纯代码。
    - 单一数据源失败不影响其他数据源，失败原因写入 manifest 的 warnings。
    - 同名文件默认拒绝覆盖；重取请提高 --rev（Rev01、Rev02……）。
    - 不填补空值，不用估计值替代缺失数据。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import pandas as pd

# 数据分层：内容 -> 一级目录；报表/行情/估值三层再按公司分一层
LAYER_BY_CONTENT = {
    "资产负债表": "01_报表原表", "利润表": "01_报表原表", "现金流量表": "01_报表原表",
    "财务摘要": "01_报表原表", "财务指标": "01_报表原表",
    "行情": "02_行情数据",
    "估值": "03_估值数据", "估值十年PE": "03_估值数据", "估值十年PB": "03_估值数据",
}
INDEX_LAYER = "00_索引"
PER_COMPANY_LAYERS = ("01_报表原表", "02_行情数据", "03_估值数据")


def layer_dir(out_dir: Path, layer: str, company_dir: str) -> Path:
    """返回分层目录并确保存在"""
    p = out_dir / layer
    if layer in PER_COMPANY_LAYERS:
        p = p / company_dir
    p.mkdir(parents=True, exist_ok=True)
    return p


def resolve_company_dirs(ak, symbols: list[str]) -> dict[str, str]:
    """用 AKShare 的 A 股代码名称表生成「代码_简称」目录名；取不到则退回纯代码"""
    try:
        table = ak.stock_info_a_code_name()
        name_map = dict(zip(table["code"].astype(str), table["name"].astype(str)))
    except Exception as exc:  # noqa: BLE001 - 名称不是必需品
        print(f"[WARN] 未取到股票简称表（{type(exc).__name__}），分层目录将只用代码")
        name_map = {}
    return {s: (f"{s}_{name_map[s]}" if s in name_map else s) for s in symbols}


def a_prefix(symbol: str) -> str:
    if symbol.startswith(("6", "9")):
        return "sh"
    if symbol.startswith(("0", "3")):
        return "sz"
    if symbol.startswith(("4", "8")):
        return "bj"
    return "sh"


def build_name(symbol: str, content: str, rev: str) -> str:
    return f"{symbol}_{content}_{dt.date.today():%Y%m%d}_{rev}.csv"


class Writer:
    """统一落盘：拒绝覆盖同名文件，并把结果写进 manifest。"""

    def __init__(self, out_dir: Path, rev: str, company_dirs: dict[str, str]):
        self.out_dir = out_dir
        self.rev = rev
        self.company_dirs = company_dirs
        self.entries: list[dict] = []
        self.warnings: list[str] = []

    def _record(self, symbol: str, content: str, source: str, api: str,
                df: pd.DataFrame, filename: str, date_col: str | None) -> None:
        entry = {"symbol": symbol, "content": content, "status": "ok", "file": filename,
                 "api": api, "source": source, "rows": int(df.shape[0]),
                 "cols": int(df.shape[1])}
        if date_col and date_col in df.columns and len(df):
            series = pd.to_datetime(df[date_col], errors="coerce").dropna()
            if len(series):
                span_years = max((series.max() - series.min()).days / 365.25, 1.0)
                entry["date_range"] = [str(series.min().date()), str(series.max().date())]
                entry["points_per_year"] = round(df.shape[0] / span_years, 1)
        self.entries.append(entry)

    def save(self, symbol: str, content: str, source: str, api: str, df: pd.DataFrame,
             date_col: str | None = None) -> None:
        target = layer_dir(self.out_dir, LAYER_BY_CONTENT[content],
                           self.company_dirs.get(symbol, symbol))
        path = target / build_name(symbol, content, self.rev)
        rel = path.relative_to(self.out_dir).as_posix()
        if path.exists():
            self.warnings.append(f"{rel} 已存在，本次跳过（提高 --rev 重新取数）")
            self.entries.append({"symbol": symbol, "content": content,
                                 "status": "skipped_exists", "file": rel, "api": api})
            return
        df.to_csv(path, index=False, encoding="utf-8-sig")
        self._record(symbol, content, source, api, df, rel, date_col)

    def fail(self, symbol: str, content: str, api: str, exc: Exception) -> None:
        self.warnings.append(f"{symbol} {content} ({api}) 取数失败：{type(exc).__name__}: {exc}")
        self.entries.append({"symbol": symbol, "content": content, "status": "failed",
                             "api": api, "error": f"{type(exc).__name__}: {exc}"})


def fetch_baidu_valuation(ak, symbol: str, indicator: str, label: str) -> pd.DataFrame:
    """百度估值序列（近十年）。PE 与 PB 的采样日期不同，故各自单独落盘。"""
    df = ak.stock_zh_valuation_baidu(symbol=symbol, indicator=indicator, period="近十年")
    df = df.rename(columns={"value": label})
    df[label] = pd.to_numeric(df[label], errors="coerce")
    return df


def collect(ak, symbol: str, start_date: str, end_date: str, w: Writer) -> None:
    prefixed = a_prefix(symbol) + symbol
    start_year = str(int(start_date[:4]))

    jobs = [
        ("财务摘要", "stock_financial_abstract", "新浪", None,
         lambda: ak.stock_financial_abstract(symbol=symbol)),
        ("财务指标", "stock_financial_analysis_indicator", "新浪", None,
         lambda: ak.stock_financial_analysis_indicator(symbol=symbol, start_year=start_year)),
        ("资产负债表", "stock_financial_report_sina", "新浪", None,
         lambda: ak.stock_financial_report_sina(stock=prefixed, symbol="资产负债表")),
        ("利润表", "stock_financial_report_sina", "新浪", None,
         lambda: ak.stock_financial_report_sina(stock=prefixed, symbol="利润表")),
        ("现金流量表", "stock_financial_report_sina", "新浪", None,
         lambda: ak.stock_financial_report_sina(stock=prefixed, symbol="现金流量表")),
        ("行情", "stock_zh_a_daily", "新浪（前复权）", "date",
         lambda: ak.stock_zh_a_daily(symbol=prefixed, start_date=start_date,
                                     end_date=end_date, adjust="qfq")),
        ("估值", "stock_value_em", "东方财富（日频）", "数据日期",
         lambda: ak.stock_value_em(symbol=symbol)),
        ("估值十年PE", "stock_zh_valuation_baidu", "百度", "date",
         lambda: fetch_baidu_valuation(ak, symbol, "市盈率(TTM)", "pe_ttm")),
        ("估值十年PB", "stock_zh_valuation_baidu", "百度", "date",
         lambda: fetch_baidu_valuation(ak, symbol, "市净率", "pb")),
    ]
    for content, api, source, date_col, fn in jobs:
        try:
            df = fn()
            if df is None or len(df) == 0:
                raise ValueError("接口返回空表")
            w.save(symbol, content, source, api, df, date_col=date_col)
        except Exception as exc:  # noqa: BLE001 - 记录后继续下一个数据源
            w.fail(symbol, content, api, exc)


def main() -> int:
    parser = argparse.ArgumentParser(description="AKShare A 股投研取数脚本")
    parser.add_argument("--symbols", required=True, help="逗号分隔的 6 位 A 股代码")
    parser.add_argument("--out", required=True, help="原始数据输出目录")
    parser.add_argument("--rev", default="Rev00", help="版本号，默认 Rev00")
    parser.add_argument("--years", type=int, default=10, help="行情回溯年数，默认 10")
    args = parser.parse_args()

    try:
        import akshare as ak
    except ImportError:
        print("[ERROR] 未安装 AKShare，无法取数。请先安装或改由用户提供导出数据。", file=sys.stderr)
        return 3

    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    end_date = dt.date.today()
    start_date = end_date.replace(year=end_date.year - args.years)
    start_s, end_s = start_date.strftime("%Y%m%d"), end_date.strftime("%Y%m%d")

    symbols = [s.strip() for s in args.symbols.split(",") if s.strip()]
    company_dirs = resolve_company_dirs(ak, symbols)
    writer = Writer(out_dir, args.rev, company_dirs)
    print(f"AKShare {getattr(ak, '__version__', '?')} | 输出目录 {out_dir} | 行情区间 {start_s}~{end_s}")
    for s in symbols:
        print(f"  {s} -> {company_dirs.get(s, s)}")

    for symbol in symbols:
        if len(symbol) != 6:
            writer.warnings.append(f"跳过非 6 位 A 股代码：{symbol}")
            print(f"[SKIP] 非 6 位 A 股代码：{symbol}")
            continue
        print(f"--- {symbol} ---")
        collect(ak, symbol, start_s, end_s, writer)

    manifest = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "akshare_version": getattr(ak, "__version__", None),
        "rev": args.rev,
        "output_dir": str(out_dir),
        "price_window": {"start": start_s, "end": end_s, "years": args.years},
        "entries": writer.entries,
        "warnings": writer.warnings,
    }
    manifest_path = layer_dir(out_dir, INDEX_LAYER, "") / \
        f"fetch_manifest_{dt.date.today():%Y%m%d}_{args.rev}.json"
    if manifest_path.exists():
        print(f"[ERROR] manifest 已存在，未覆盖：{manifest_path.name}（提高 --rev 重跑）", file=sys.stderr)
        return 4
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    ok = sum(1 for e in writer.entries if e.get("status") == "ok")
    failed = sum(1 for e in writer.entries if e.get("status") == "failed")
    print(f"\n完成：成功 {ok} 项，失败 {failed} 项，告警 {len(writer.warnings)} 条")
    for msg in writer.warnings:
        print(f"  [WARN] {msg}")
    print(f"manifest: {manifest_path}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
