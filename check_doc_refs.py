#!/usr/bin/env python3
"""复核文档里的文件引用是否与实际文件对得上。

用法：
    python check_doc_refs.py

检查范围：
    ../数据文件/*.md、./README.md

检查内容：
    1. 文档中反引号引用的文件名是否存在（支持分层目录内的文件名匹配）
    2. 文档声明的「对应数据版本」与原始数据层实际的最大版本是否一致

说明：
    「对应数据版本」只统计接口取数的数据层（00–04、06）。`05_手工整理数据/` 是手工整理的
    独立数据集，版本号按文件族各自递增（如 WACC 参数 Rev00 → Rev01），不参与主版本判定。

退出码：0 全部通过；1 存在过期引用
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent


def find_project_root(start: Path) -> Path:
    """从脚本位置向上找项目根：同时含「原始数据」与「数据文件」两级目录的那个"""
    for p in (start, *start.parents):
        if (p / "原始数据").is_dir() and (p / "数据文件").is_dir():
            return p
    raise SystemExit(f"未找到项目根目录（应含 原始数据/ 与 数据文件/）：{start}")


PROJ = find_project_root(REPO)
RAW = PROJ / "原始数据"
DOC_DIR = PROJ / "数据文件"

FILE_PAT = re.compile(r"`([^`\n]+\.(?:csv|json|xlsx|md|py|pdf|ipynb))`")
VER_PAT = re.compile(r"_(\d{8})_(Rev\d+)")
DECLARED_PAT = re.compile(r"对应数据版本\*{0,2}\s*[:：][^\n]{0,40}?(Rev\d+)")
SKIP_PREFIX = ("python ", "pip ")
MANUAL_LAYER = "05_手工整理数据"


def rev_num(text: str) -> int:
    """把 RevNN 取成整数比较。字符串比较会把 Rev9 排在 Rev10 后面，绝不能这么比。"""
    m = re.search(r"Rev(\d+)", text)
    return int(m.group(1)) if m else -1


def actual_versions() -> Counter:
    """统计原始数据层里各 (日期, 版本) 出现次数；手工整理层不计入"""
    c: Counter = Counter()
    for f in RAW.rglob("*"):
        if f.is_file():
            if MANUAL_LAYER in f.relative_to(RAW).parts:
                continue
            m = VER_PAT.search(f.name)
            if m:
                c[f"{m.group(1)}_{m.group(2)}"] += 1
    return c


def doc_files() -> list[Path]:
    """每个文档族只取最新版本（旧版是历史记录，不参与检查），外加 README。

    文档按族分装在 数据文件/ 的子文件夹里，所以要递归扫描；
    递归时同时看文件名与所在文件夹，避免不同族的同名文件互相覆盖。
    """
    families: dict[str, Path] = {}
    for d in sorted(DOC_DIR.rglob("*.md")):
        m = VER_PAT.search(d.name)
        fam = f"{d.parent.name}/{VER_PAT.sub('', d.name)}"
        if not m:
            families[fam] = d
            continue
        cur = families.get(fam)
        if cur is None or rev_num(m.group(2)) > rev_num(VER_PAT.search(cur.name).group(2)):
            families[fam] = d
    out = list(families.values())
    readme = REPO / "README.md"
    if readme.exists():
        out.append(readme)
    return out


def ref_exists(ref: str, all_names: set[str]) -> bool:
    p = Path(ref)
    if p.is_absolute():
        return p.exists()
    return Path(ref).name in all_names


def latest_of_family(name: str) -> str | None:
    """返回同一文档族里版本号最高的文件名；该文件不是版本化文档时返回 None"""
    m = VER_PAT.search(name)
    if not m:
        return None
    fam = VER_PAT.sub("", name)
    best = name
    for d in DOC_DIR.rglob("*.md"):
        if VER_PAT.sub("", d.name) == fam:
            cur = VER_PAT.search(d.name)
            if cur and rev_num(cur.group(2)) > rev_num(VER_PAT.search(best).group(2)):
                best = d.name
    return best


def main() -> int:
    versions = actual_versions()
    print("原始数据层实际版本分布：")
    for k, v in sorted(versions.items()):
        print(f"  {k}  共 {v} 个文件")
    # 层版本 = 树内最大 Rev 号。
    # 不能用「日期最新那一组的 Rev」：新数据族从 Rev00/Rev01 起步，那样会把层版本号算低，
    # 反过来逼所有文档把声明从 Rev05 改成 Rev01（2026-10-12 修正）。
    layer_rev = max((rev_num(k) for k in versions), default=0)
    newest_date = max((k.split("_")[0] for k in versions), default=None)
    main_ver = f"Rev{layer_rev:02d}"
    print(f"\n原始数据层版本 = {main_ver}（树内最大 Rev 号）；最新取数日期 = {newest_date}\n")

    problems = 0
    all_names = {f.name for f in PROJ.rglob("*") if f.is_file()}

    for doc in doc_files():
        txt = doc.read_text(encoding="utf-8")
        # 「上一版：`xxx`」指向自己的历史版本，属正常，不参与最新版本校验
        txt_no_prev = re.sub(r"上一版本?[：:]\s*`[^`]+`", "", txt)
        issues = []

        # 1. 文件引用是否存在
        for ref in sorted(set(FILE_PAT.findall(txt))):
            if ref.startswith(SKIP_PREFIX) or "<" in ref or "*" in ref:
                continue
            if not ref_exists(ref, all_names):
                issues.append(f"引用不存在：{ref}")
                continue
            latest = latest_of_family(Path(ref).name)
            if latest and latest != Path(ref).name and ref in txt_no_prev:
                issues.append(f"引用的是旧版本 {Path(ref).name}，最新为 {latest}")

        # 2. 声明的对应数据版本
        m = DECLARED_PAT.search(txt)
        if not m:
            issues.append("未声明「对应数据版本」")
        elif main_ver and rev_num(m.group(1)) != rev_num(main_ver):
            issues.append(f"声明为 {m.group(1)}，实际最新为 {main_ver}")

        if issues:
            problems += len(issues)
            print(f"[需处理] {doc.name}")
            for it in issues:
                print(f"    - {it}")
        else:
            print(f"[通过]   {doc.name}")

    print(f"\n{'发现 ' + str(problems) + ' 处待处理' if problems else '全部通过'}")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
