# -*- coding: utf-8 -*-
"""抓取农业农村部畜牧兽医局「监测预警」栏目的生鲜乳周度价格。

为什么要单独爬
    AKShare 没有农业农村部接口；国家统计局「农产品集贸市场价格」下 27 个指标
    也不含生鲜乳。原奶周期维度只剩这一个一手官方源头。

数据源
    列表页 https://xmsyj.moa.gov.cn/jcyj/ （翻页规则 index_N.htm）
    目标文章《X 月第 N 周畜产品和饲料集贸市场价格情况》每周一篇
    取值处：正文「生鲜乳价格。内蒙古、河北等 10 个主产省份生鲜乳平均价格 X.XX 元/公斤，
            比前一周上涨/下跌/持平 X.X%，同比上涨/下跌/持平 X.X%。」

输出
    原始数据/04_行业数据/生鲜乳价格/宏观_生鲜乳价格周度_<运行日>_Rev<nn>.csv      长表，一行一周
    原始数据/04_行业数据/生鲜乳价格/生鲜乳收购价周数据_全量_<运行日>_Rev<nn>.xlsx  与用户原表同版式

用法
    python fetch_moa_milk_price.py
    python fetch_moa_milk_price.py --since 2016-01-01 --rev Rev01
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = Path(__file__).resolve().parents[2] / "原始数据"
INDUSTRY_DIR = BASE / "04_行业数据"
# 04 层按主题分两个子文件夹（2026-10-11 起）：宏观统计 / 生鲜乳价格
MILK_DIR = INDUSTRY_DIR / "生鲜乳价格"

LIST_URL = "https://xmsyj.moa.gov.cn/jcyj/"
TITLE_KEY = "畜产品和饲料集贸市场价格情况"
UA = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                     "(KHTML, like Gecko) Chrome/125.0 Safari/537.36")}

MAX_PAGES = 30

# 老标题写「12月份第3周」，新标题写「9月第5周」，两种都要认
WEEK_RE = re.compile(r"(\d{1,2})月份?第(\d)周")
DATE_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
SAMPLE_RE = re.compile(r"采集日为(\d{1,2})月(\d{1,2})日")
PRICE_RE = re.compile(r"生鲜乳平均价格([0-9.]+)元\s*/\s*公斤")
# 「比前一周上涨 x%」与「与前一周持平」都要认
MOM_RE = re.compile(r"[比与]前一周(上涨|上升|下跌|下降|持平)([0-9.]*)%?")
YOY_RE = re.compile(r"同比(上涨|上升|下跌|下降|持平)([0-9.]*)%?")


def get(url, session):
    try:
        resp = session.get(url, headers=UA, timeout=25)
        resp.encoding = resp.apparent_encoding or resp.encoding
        return resp.text if resp.status_code == 200 else None
    except requests.RequestException:
        return None


def signed(direction, number):
    if direction == "持平":
        return 0.0
    if not number:
        return None
    value = float(number)
    return -value if direction in ("下跌", "下降") else value


def parse_article(html, title, url):
    text = re.sub(r"\s+", "", BeautifulSoup(html, "lxml").get_text(" "))
    price = PRICE_RE.search(text)
    if not price:
        return None
    week = WEEK_RE.search(title)
    published = DATE_RE.search(title) or DATE_RE.search(text)
    if not week or not published:
        return None

    # 正文形态是「生鲜乳价格。内蒙古、河北等……同比上涨 2.0%。饲料价格。……」
    # 小标题自己带一个句号，所以要从「生鲜乳价格。」之后开始切，再切到下一个句号，
    # 否则会越过句子边界，把玉米、豆粕的环比同比抓成生鲜乳的
    marker = "生鲜乳价格。"
    body_start = (text.index(marker) + len(marker)) if marker in text else text.find("生鲜乳价格")
    stop = text.find("。", body_start)
    sentence = text[body_start:stop + 1] if stop > body_start else text[body_start:body_start + 200]
    mom = MOM_RE.search(sentence)
    yoy = YOY_RE.search(sentence)

    pub_year, pub_month, pub_day = (int(x) for x in published.groups())
    month, nth = int(week.group(1)), int(week.group(2))
    year = pub_year - 1 if month - pub_month > 6 else pub_year

    sample = SAMPLE_RE.search(text)
    sample_date = None
    if sample:
        s_month, s_day = int(sample.group(1)), int(sample.group(2))
        # 采集日一定不晚于该周所在月。1 月第 1 周的采集日会落在上一年 12 月，
        # 例如「2026年1月第1周（采集日为12月31日）」应为 2025-12-31
        s_year = year - 1 if s_month > month else year
        sample_date = f"{s_year:04d}-{s_month:02d}-{s_day:02d}"

    return {
        "周次": f"{year}年{month}月第{nth}周",
        "采集日": sample_date,
        "发布日期": f"{pub_year:04d}-{pub_month:02d}-{pub_day:02d}",
        "价格_元每公斤": float(price.group(1)),
        "环比_pct": signed(*mom.groups()) if mom else None,
        "同比_pct": signed(*yoy.groups()) if yoy else None,
        "标题": title,
        "来源URL": url,
    }


def collect_links(session, since):
    found = {}
    for page in range(MAX_PAGES):
        url = LIST_URL if page == 0 else f"{LIST_URL}index_{page}.htm"
        html = get(url, session)
        if not html:
            print(f"  列表第 {page + 1} 页取不到，停止翻页")
            break
        soup = BeautifulSoup(html, "lxml")
        fresh = 0
        oldest = None
        for anchor in soup.find_all("a"):
            title = anchor.get_text(strip=True)
            href = anchor.get("href")
            if not href or TITLE_KEY not in title:
                continue
            date = DATE_RE.search(title)
            if not date:
                continue
            published = dt.date(*(int(x) for x in date.groups()))
            oldest = published if oldest is None else min(oldest, published)
            if published < since:
                continue
            full = requests.compat.urljoin(url, href)
            if full not in found:
                found[full] = title
                fresh += 1
        print(f"  列表第 {page + 1} 页：新增 {fresh} 篇，本页最早 {oldest}")
        if oldest and oldest < since:
            break
    return found


def main():
    parser = argparse.ArgumentParser(description="抓取农业农村部生鲜乳周度价格")
    parser.add_argument("--since", default="2020-01-01")
    parser.add_argument("--rev", default="Rev00")
    parser.add_argument("--delay", type=float, default=0.15,
                        help="每篇文章之间的停顿秒数；被站点限流时调大")
    args = parser.parse_args()
    since = dt.date.fromisoformat(args.since)

    session = requests.Session()
    print(f"=== 收集文章列表（回溯到 {since}）===")
    links = collect_links(session, since)
    print(f"待抓取 {len(links)} 篇")

    rows, failed = [], []
    for i, (url, title) in enumerate(sorted(links.items()), 1):
        html = get(url, session)
        row = parse_article(html, title, url) if html else None
        if row:
            rows.append(row)
        else:
            failed.append(url)
        if i % 25 == 0:
            print(f"  已处理 {i}/{len(links)}")
        time.sleep(args.delay)

    MILK_DIR.mkdir(parents=True, exist_ok=True)
    # 必须按采集日排。按「周次」这类中文字符串排会把 2022年10月 排到 2022年1月 前面
    df = pd.DataFrame(rows).sort_values("采集日").reset_index(drop=True)
    csv_path = MILK_DIR / f"宏观_生鲜乳价格周度_{dt.date.today():%Y%m%d}_{args.rev}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    # 自检：采集日不能晚于发布日期，也不能晚于今天
    today = dt.date.today().isoformat()
    bad = df[(df["采集日"] > df["发布日期"]) | (df["采集日"] > today)]
    if len(bad):
        print(f"  [警告] {len(bad)} 行的采集日不合理，请复核：")
        print(bad[["周次", "采集日", "发布日期"]].head(10).to_string(index=False))
    else:
        print("  自检通过：采集日均不晚于发布日期与今天")

    # 宽表按时间倒序放列（最新在左），与用户原表版式一致
    weeks = df["周次"].tolist()[::-1]
    wide = [
        ["时间", *weeks],
        ["价格/元/公斤", *df["价格_元每公斤"].tolist()[::-1]],
        ["同比", *[round(v / 100, 4) if pd.notna(v) else None for v in df["同比_pct"][::-1]]],
        ["环比", *[round(v / 100, 4) if pd.notna(v) else None for v in df["环比_pct"][::-1]]],
        ["数据来源", "中华人民共和国农业农村部畜牧兽医局", *[""] * (len(df) - 1)],
    ]
    xlsx_path = MILK_DIR / f"生鲜乳收购价周数据_全量_{dt.date.today():%Y%m%d}_{args.rev}.xlsx"
    pd.DataFrame(wide).to_excel(xlsx_path, index=False, header=False)

    print(f"\n解析成功 {len(df)} 周，失败 {len(failed)} 篇")
    if len(df):
        print(f"  区间：{df['周次'].iloc[0]} ~ {df['周次'].iloc[-1]}")
        print(f"  价格区间：{df['价格_元每公斤'].min():.2f} ~ {df['价格_元每公斤'].max():.2f} 元/公斤")
        print(df.tail(6)[["周次", "采集日", "价格_元每公斤", "环比_pct", "同比_pct"]].to_string(index=False))
    print(f"  长表：{csv_path.name}")
    print(f"  宽表：{xlsx_path.name}")


if __name__ == "__main__":
    main()
