# -*- coding: utf-8 -*-
"""行业分析用的宏观与行业官方统计取数（AKShare）。

为什么要取这些
    报告的行业分析「需求」维度需要行业总量数据。按证据规则，官方统计（C 级）
    高于券商研报（D 级），所以先把能拿到的官方口径落盘，再谈研报对照。

取数目标
    社零总额、消费者信心指数、CPI、工业增加值增长      —— 东方财富经济数据（原始来源：国家统计局）
    乳制品产量、社零分项                              —— 国家统计局月度数据（通用接口，需给菜单路径）

落盘
    原始数据/04_行业数据/宏观统计/宏观_<指标名>_<运行日>_Rev<nn>.csv
    原始数据/00_索引/macro_manifest_<运行日>_Rev<nn>.json

用法
    python fetch_macro_data.py --discover          # 只打印统计局菜单树，不取数
    python fetch_macro_data.py                     # 取数并落盘
    python fetch_macro_data.py --rev 1             # 提高版本号重取
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parents[2] / "原始数据"
INDUSTRY_DIR = BASE / "04_行业数据"
# 04 层按主题分两个子文件夹（2026-10-11 起）：宏观统计 / 生鲜乳价格
OUT_DIR = INDUSTRY_DIR / "宏观统计"
INDEX_DIR = BASE / "00_索引"
DOC_DIR = Path(__file__).resolve().parents[2] / "数据文件"

# 东方财富经济数据接口：名字 -> (akshare 函数名, 出处)
EM_JOBS: dict[str, tuple[str, str]] = {
    "社零总额": ("macro_china_consumer_goods_retail",
              "东方财富-经济数据-社会消费品零售总额（原始来源：国家统计局）"),
    "消费者信心指数": ("macro_china_xfzxx",
                "东方财富-经济数据一览-消费者信心指数（原始来源：国家统计局）"),
    "CPI": ("macro_china_cpi",
            "东方财富-中国居民消费价格指数（原始来源：国家统计局）"),
    "工业增加值增长": ("macro_china_gyzjz",
                "东方财富-经济数据-工业增加值增长（原始来源：国家统计局）"),
}

# 国家统计局通用接口：名字 -> (kind, path, period, 出处)。路径按统计局菜单名用 > 连接，
# 菜单名由 --discover 实取，不要凭印象写
NBS_JOBS: dict[str, tuple[str, str, str, str]] = {
    "乳制品产量": ("月度数据", "工业>工业主要产品产量>乳制品", "2016-",
               "国家统计局-月度数据-工业-工业主要产品产量-乳制品（产量，万吨）"),
    "社零总额_统计局": ("月度数据", "国内贸易>社会消费品零售总额", "2016-",
                  "国家统计局-月度数据-国内贸易-社会消费品零售总额"),
    "社零按消费类型": ("月度数据", "国内贸易>按消费类型分社会消费品零售总额", "2016-",
                 "国家统计局-月度数据-国内贸易-按消费类型分社会消费品零售总额"),
    "限上粮油食品饮料烟酒类零售": ("月度数据",
                      "国内贸易>限上单位商品零售类值>粮油、食品、饮料、烟酒类", "2016-",
                      "国家统计局-月度数据-国内贸易-限上单位商品零售类值-粮油食品饮料烟酒类"),
    "网上零售额": ("月度数据", "国内贸易>网上零售额", "2016-",
               "国家统计局-月度数据-国内贸易-网上零售额"),
    "畜产品集贸市场价格": ("月度数据", "价格指数>农产品集贸市场价格>畜产品集贸市场价格", "2016-",
                   "国家统计局-月度数据-价格指数-农产品集贸市场价格-畜产品（可能含生鲜乳）"),
}

NBS_DISCOVER_PATHS = [
    "工业",
    "国内贸易",
    "价格指数",
    # 取回空表的三个数据集，往下再钻一层看有没有子目录
    "工业>工业主要产品产量",
    "国内贸易>限上单位商品零售类值",
    "价格指数>农产品集贸市场价格",
]


def discover() -> None:
    """按 NBS_DISCOVER_PATHS 逐条解析并打印其下一层目录，确认 path 怎么写。"""
    from akshare.economic import macro_china_nbs as eco

    config = eco._KIND_CONFIG["月度数据"]
    code, route = int(config["code"]), str(config["route"])
    root = eco._get_nbs_root_id(code, route)
    for path in NBS_DISCOVER_PATHS:
        node_id = root
        trail: list[str] = []
        try:
            for part in [p.strip() for p in path.split(">") if p.strip()]:
                nodes = eco._get_nbs_tree(node_id, code, route)
                target = eco._find_nbs_node_by_name(nodes, part)
                node_id = target["_id"]
                trail.append(part)
            kids = eco._get_nbs_tree(node_id, code, route)
        except Exception as exc:  # noqa: BLE001
            print(f"\n{path} 解析失败：{type(exc).__name__}: {exc}")
            continue
        print(f"\n{' > '.join(trail)} 下 {len(kids)} 个：")
        for kid in kids[:40]:
            print(f"    - {kid['name']}")


def save(name: str, df: pd.DataFrame, source: str, rev: str,
         entries: list[dict], warnings: list[str]) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"宏观_{name}_{dt.date.today():%Y%m%d}_{rev}.csv"
    if path.exists():
        warnings.append(f"{path.name} 已存在，跳过（提高 --rev 重取）")
        entries.append({"指标": name, "status": "skipped_exists", "file": path.name})
        return
    # 国家统计局接口把指标名放在索引里，写盘时必须带上索引，否则行标签全丢
    has_index = not isinstance(df.index, pd.RangeIndex)
    out = df.copy()
    if has_index:
        out.index.name = "指标"
    out.to_csv(path, index=has_index, encoding="utf-8-sig")
    entries.append({
        "指标": name, "status": "ok", "file": path.name,
        "source": source, "rows": int(df.shape[0]), "cols": int(df.shape[1]),
        "columns": [str(c) for c in df.columns],
    })
    print(f"  [OK] {name}: {df.shape[0]} 行 x {df.shape[1]} 列 -> {path.name}")


def main() -> None:
    parser = argparse.ArgumentParser(description="行业分析用宏观与行业统计取数")
    parser.add_argument("--rev", default="Rev00")
    parser.add_argument("--discover", action="store_true", help="只打印统计局菜单树")
    args = parser.parse_args()

    if args.discover:
        discover()
        return

    import akshare as ak

    entries: list[dict] = []
    warnings: list[str] = []

    print("=== 东方财富经济数据 ===")
    for name, (func, source) in EM_JOBS.items():
        try:
            df = getattr(ak, func)()
            save(name, df, source, args.rev, entries, warnings)
        except Exception as exc:  # noqa: BLE001 - 单个接口失败不影响其他
            warnings.append(f"{name}（{func}）失败：{type(exc).__name__}: {exc}")
            entries.append({"指标": name, "status": "failed", "api": func,
                            "error": f"{type(exc).__name__}: {exc}"})
            print(f"  [失败] {name}: {type(exc).__name__}")

    print("=== 国家统计局月度数据 ===")
    for name, (kind, path, period, source) in NBS_JOBS.items():
        try:
            df = ak.macro_china_nbs_nation(kind=kind, path=path, period=period)
            if df is None or df.empty:
                # 有些数据集不吃「2016-」这种区间写法，回退到 LAST10 再试一次
                df = ak.macro_china_nbs_nation(kind=kind, path=path, period="LAST10")
            if df is None or df.empty:
                raise ValueError(f"返回空表（period={period} 与 LAST10 都为空）")
            save(name, df, source, args.rev, entries, warnings)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{name}（path={path}）失败：{type(exc).__name__}: {exc}")
            entries.append({"指标": name, "status": "failed", "path": path,
                            "error": f"{type(exc).__name__}: {exc}"})
            print(f"  [失败] {name}: {type(exc).__name__}: {exc}")

    INDEX_DIR.mkdir(parents=True, exist_ok=True)
    manifest = INDEX_DIR / f"macro_manifest_{dt.date.today():%Y%m%d}_{args.rev}.json"
    manifest.write_text(json.dumps({
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "rev": args.rev,
        "entries": entries,
        "warnings": warnings,
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    ok = sum(1 for e in entries if e.get("status") == "ok")
    skipped = sum(1 for e in entries if e.get("status") == "skipped_exists")
    failed = sum(1 for e in entries if e.get("status") == "failed")
    print(f"\n成功 {ok} 项，跳过（已存在）{skipped} 项，失败 {failed} 项")
    print(f"manifest: {manifest.name}")


if __name__ == "__main__":
    main()
