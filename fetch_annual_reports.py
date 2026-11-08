# -*- coding: utf-8 -*-
"""下载可比公司年报/半年报 PDF（来源：巨潮资讯网 www.cninfo.com.cn）。

用途：光明乳业(600597)、新乳业(002946) 的年报，补齐 EBITDA 所需的
「现金流量表补充资料」（折旧摊销），与 原始资料/伊利年报/ 保持同一套文件。

用法：python fetch_annual_reports.py            # 列出候选，不下载
      python fetch_annual_reports.py --write    # 列出并下载到 原始资料/
"""
import os
import re
import sys
import time
from urllib.parse import urlparse, parse_qs

import akshare as ak
import requests

sys.stdout.reconfigure(encoding="utf-8")

ROOT = str(Path(__file__).resolve().parents[2] / "原始资料")
TARGETS = [
    ("600597", "光明乳业", "光明乳业年报"),
    ("002946", "新乳业", "新乳业年报"),
]
YEARS = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
HALF_YEAR = 2026
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def query(code, category, start, end):
    df = ak.stock_zh_a_disclosure_report_cninfo(
        symbol=code, market="沪深京", category=category,
        start_date=start, end_date=end,
    )
    return df


def pick_annual(df, year):
    """挑出某一年度的报告正文（排除摘要、英文版、更正、已取消）。"""
    pat = re.compile(rf"{year}\s*年\s*年度报告")
    bad = re.compile(r"摘要|英文|English|更正|取消|问询|意见|说明|审计报告")
    hit = df[df["公告标题"].astype(str).apply(lambda t: bool(pat.search(t)) and not bad.search(t))]
    if hit.empty:
        return None
    return hit.sort_values("公告时间").iloc[-1]


def pick_half(df, year):
    pat = re.compile(rf"{year}\s*年\s*半年度报告")
    bad = re.compile(r"摘要|英文|English|更正|取消|问询|意见|说明")
    hit = df[df["公告标题"].astype(str).apply(lambda t: bool(pat.search(t)) and not bad.search(t))]
    if hit.empty:
        return None
    return hit.sort_values("公告时间").iloc[-1]


def pdf_url(announce_url):
    q = parse_qs(urlparse(announce_url).query)
    aid = q["announcementId"][0]
    raw = q["announcementTime"][0]
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        ds = raw                                   # akshare 新版直接给日期串
    else:
        ds = time.strftime("%Y-%m-%d", time.localtime(int(raw) / 1000))
    return f"http://static.cninfo.com.cn/finalpage/{ds}/{aid}.PDF"


def download(url, path):
    for u in (url, url[:-4] + ".pdf"):
        try:
            r = requests.get(u, headers=HEADERS, timeout=60)
            if r.status_code == 200 and r.content[:4] == b"%PDF":
                with open(path, "wb") as fh:
                    fh.write(r.content)
                return len(r.content)
        except Exception as exc:  # noqa: BLE001
            print("    下载失败", type(exc).__name__, exc)
    return 0


def main(write):
    for code, name, folder in TARGETS:
        out = os.path.join(ROOT, folder)
        os.makedirs(out, exist_ok=True)
        print(f"\n===== {name} {code} -> {out}")
        ann = query(code, "年报", "20200101", "20261231")
        half = query(code, "半年报", f"{HALF_YEAR}0101", f"{HALF_YEAR}1231")
        jobs = []
        for y in YEARS:
            row = pick_annual(ann, y)
            jobs.append((f"{y}年年度报告.pdf", row))
        row = pick_half(half, HALF_YEAR)
        jobs.append((f"{HALF_YEAR}年半年报告.pdf", row))
        for fname, row in jobs:
            if row is None:
                print(f"  [缺] {fname}  未在公告列表中找到")
                continue
            url = pdf_url(row["公告链接"])
            print(f"  [有] {fname}  {row['公告标题']}  {row['公告时间']}")
            if write:
                p = os.path.join(out, fname)
                if os.path.exists(p) and os.path.getsize(p) > 100_000:
                    print("       已存在，跳过")
                    continue
                n = download(url, p)
                print(f"       -> {'OK ' + str(round(n / 1024 / 1024, 2)) + ' MB' if n else 'FAILED'}", url)
                time.sleep(1)


if __name__ == "__main__":
    main(write="--write" in sys.argv)
