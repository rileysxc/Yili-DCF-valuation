# -*- coding: utf-8 -*-
"""三家 PE(TTM)/PB(LF) 历史分位表（伊利 600887、光明 600597、新乳业 002946）。

主口径  ：近 5 年日频（2021-01-01 ~ 2026-10-09），剔除 PE/PB ≤ 0 与空值。
稳健性  ：近 10 年周度（2016-10-11 ~ 2026-10-09）。
当前分位：P = 样本中 ≤ 当前值的个数 ÷ 样本总数（占比法，含并列值）。

数据全部取 原始数据/07_清洗后数据/：
  *日频  <代码>_估值_<日期>_RevNN.csv          列 PE(TTM) / 市净率
  *周度  <代码>_估值十年PE/PB_<日期>_RevNN.csv  列 pe_ttm / pb
  周度 PE 的负值在清洗阶段已被上游剔除，剔除数只能取 00_清洗日志_<日期>_RevNN.csv。

用法：python build_valuation_percentiles.py
      python build_valuation_percentiles.py --rev Rev01
"""
import argparse
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

BASE = str(Path(__file__).resolve().parents[2] / "原始数据" / "07_清洗后数据")
OUT_DIR = str(Path(__file__).resolve().parents[2] / "数据文件" / "估值分位表")

DAILY_START, DAILY_END = "2021-01-01", "2026-10-09"
LAST_TRADING_DAY = "2026-10-09"

COMPANIES = {
    "600887": ("伊利", "600887_伊利股份"),
    "600597": ("光明", "600597_光明乳业"),
    "002946": ("新乳业", "002946_新乳业"),
}

COLS = ["样本数", "剔除数", "剔除比例(%)", "当前值", "当前值日期", "当前分位(%)",
        "max", "min", "mean", "median", "P10", "P25", "P50", "P75", "P90", "备注"]


def locate(folder, stem):
    """取该报表的最新 Rev 文件。"""
    d = os.path.join(BASE, folder)
    hits = sorted(f for f in os.listdir(d) if f.startswith(stem) and f.endswith(".csv"))
    if not hits:
        raise FileNotFoundError(f"{folder}/{stem}* 不存在")
    return os.path.join(d, hits[-1])


def stats(values, asof, extra=None):
    """按占比法算分位与描述统计。"""
    v = pd.Series(values).dropna()
    cur = v.iloc[-1]                      # 序列已按日期升序
    row = {
        "当前值": round(cur, 4),
        "当前值日期": asof,
        "当前分位(%)": round((v <= cur).mean() * 100, 1),
        "max": round(v.max(), 2), "min": round(v.min(), 2),
        "mean": round(v.mean(), 2), "median": round(v.median(), 2),
    }
    for q, name in [(0.10, "P10"), (0.25, "P25"), (0.50, "P50"), (0.75, "P75"), (0.90, "P90")]:
        row[name] = round(v.quantile(q), 2)
    row.update(extra or {})
    return row


def daily_table(rev_drop_cols):
    rows = []
    for code, (name, folder) in COMPANIES.items():
        df = pd.read_csv(locate(folder, f"{code}_估值_"))
        df["date"] = pd.to_datetime(df["date"])
        df = df[(df["date"] >= DAILY_START) & (df["date"] <= DAILY_END)].sort_values("date")
        n_all = len(df)
        for col, tag in [("PE(TTM)", "PE(TTM)"), ("市净率", "PB")]:
            keep = df[col].notna() & (df[col] > 0)
            n_drop = int((~keep).sum())
            d = df.loc[keep]
            asof = d["date"].iloc[-1].strftime("%Y-%m-%d")
            rows.append({
                "指标": f"{name}_{tag}",
                "样本数": len(d),
                "剔除数": n_drop,
                "剔除比例(%)": round(n_drop / n_all * 100, 2),
                **stats(d[col].tolist(), asof,
                        extra={"备注": "" if asof == LAST_TRADING_DAY
                               else f"最新交易日无有效值，当前值取 {asof}，当前分位不适用"}),
            })
    return pd.DataFrame(rows).set_index("指标")[COLS]


def weekly_drop():
    """周度 PE/PB 的负值剔除数：只能取清洗日志（上游已剔除），且分表名统计。"""
    log = pd.read_csv(locate("", "00_清洗日志_"))
    log.columns = [c.strip() for c in log.columns]
    table2tag = {"估值十年PE": "PE(TTM)", "估值十年PB": "PB"}
    out = {}
    for _, r in log[log["表名"].astype(str).str.strip().isin(table2tag)].iterrows():
        key = (str(r["代码"]).zfill(6), table2tag[str(r["表名"]).strip()])
        n_out = 0 if pd.isna(r["输出行数"]) else int(r["输出行数"])
        out[key] = int(r["源行数"]) - n_out
    return out


def weekly_table():
    drops = weekly_drop()
    rows = []
    for code, (name, folder) in COMPANIES.items():
        for stem, col, tag in [(f"{code}_估值十年PE_", "pe_ttm", "PE(TTM)"),
                               (f"{code}_估值十年PB_", "pb", "PB")]:
            df = pd.read_csv(locate(folder, stem))
            df["date"] = pd.to_datetime(df["date"])
            df = df.sort_values("date")
            n_drop = drops[(code, tag)]
            asof = df["date"].iloc[-1].strftime("%Y-%m-%d")
            rows.append({
                "指标": f"{name}_{tag}",
                "样本数": len(df),
                "剔除数": n_drop,
                "剔除比例(%)": round(n_drop / (len(df) + n_drop) * 100, 2) if n_drop else 0.0,
                **stats(df[col].tolist(), asof,
                        extra={"备注": "" if asof in ("2026-10-09", "2026-10-10")
                               else "最新周度无有效值（PE 转负），当前分位不适用"}),
            })
    return pd.DataFrame(rows).set_index("指标")[COLS]


def cross_check(daily, weekly):
    """三种方法复算分位 + 与 07 层末值对照。"""
    print("\n① 与 07 层原始末值对照")
    for code, (name, folder) in COMPANIES.items():
        raw = pd.read_csv(locate(folder, f"{code}_估值_")).iloc[-1]
        for tag, col in [("PE(TTM)", "PE(TTM)"), ("PB", "市净率")]:
            print(f"   {name}_{tag}: 本表 {daily.loc[f'{name}_{tag}', '当前值']:<10} "
                  f"07层末行 {round(raw[col], 4)}")

    print("\n② 占比法 / 排序法 / rank 法 复算")
    for code, (name, folder) in COMPANIES.items():
        for stem, col in [(f"{code}_估值_", "PE(TTM)"), (f"{code}_估值十年PE_", "pe_ttm")]:
            df = pd.read_csv(locate(folder, stem))
            df["date"] = pd.to_datetime(df["date"])
            df = df[df[col].notna() & (df[col] > 0)].sort_values("date")
            v = df[col]
            cur = v.iloc[-1]
            a = (v <= cur).mean() * 100
            b = np.searchsorted(np.sort(v.values), cur, side="right") / len(v) * 100
            c = v.rank(pct=True).iloc[-1] * 100
            print(f"   {name} {stem.split('_')[-1]:<6} 占比 {round(a, 2):>6} | 排序 {round(b, 2):>6} | "
                  f"rank {round(c, 2):>6} | 并列 {(v == cur).sum()}")

    print("\n③ 样本数")
    print(daily[["样本数", "剔除数"]].to_string())
    print(weekly[["样本数", "剔除数"]].to_string())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rev", default="Rev00")
    ap.add_argument("--date", default="20261015")
    ap.add_argument("--no-write", action="store_true")
    args = ap.parse_args()

    daily = daily_table(None)
    weekly = weekly_table()
    cross_check(daily, weekly)

    path = os.path.join(OUT_DIR, f"三家分位表_{args.date}_{args.rev}.xlsx")
    if args.no_write:
        print(f"\n预览模式，未落盘（将写入 {path}）")
        return
    notes = pd.DataFrame([
        ("数据来源", "07 层 <代码>_估值_*_Rev*.csv（日频）、<代码>_估值十年PE/PB_*_Rev*.csv（周度）、00_清洗日志_*.csv"),
        ("主口径窗口", f"{DAILY_START} ~ {DAILY_END}，三家各 1395 个交易日"),
        ("稳健性窗口", "2016-10-11 ~ 2026-10-09 周度；新乳业受上市日限制起点 2019-01-28（约 7.7 年）；光明十年 PE 最新值停在 2026-03-29"),
        ("指标口径", "PE 取 PE(TTM)；PB 取「市净率」（LF 口径）"),
        ("剔除规则", "PE(TTM) 或 PB ≤ 0 或空值剔除。日频剔除数由本表统计；周度 PE 的负值已在上游剔除，剔除数取清洗日志（源行数 − 输出行数）"),
        ("当前分位定义", "P = 样本中 ≤ 当前值的个数 ÷ 样本总数（占比法，含并列值）"),
        ("已知口径差异", "末行日期不统一：伊利十年 PE/PB 为 2026-10-09；光明十年 PB、新乳业十年 PE/PB 为 2026-10-10（周度按自然周结束日标注）"),
        ("局限", "光明 2025A 归母净利为负，2026-03-31 起 PE(TTM) 转负，故光明 PE 的当前分位不适用，仅 PB 可比"),
    ], columns=["项目", "内容"])
    os.makedirs(OUT_DIR, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        daily.to_excel(w, sheet_name="分位表_5年日频_主")
        weekly.to_excel(w, sheet_name="分位表_10年周度_稳健性")
        notes.to_excel(w, sheet_name="口径说明", index=False)
    print(f"\n已落盘: {path}")


if __name__ == "__main__":
    main()
