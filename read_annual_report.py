#!/usr/bin/env python3
"""读取年报 / 半年报 PDF：提取文本、按关键词检索、按页查看。

用法：
    # 1) 先把 PDF 提取成文本缓存（只需跑一次）
    python read_annual_report.py extract --pdf "2026年半年报告.pdf" --out cache_hy2026.txt

    # 2) 按关键词检索，输出命中处上下文与页码
    python read_annual_report.py find --txt cache_hy2026.txt --kw 利益联结 --kw 生鲜乳价格

    # 3) 按页查看（页码为 PDF 页序）
    python read_annual_report.py pages --txt cache_hy2026.txt --frm 8 --to 10

说明：
    - extract 依赖 pdfplumber（bundled Python 已装，`C:\\Python314` 未装）。
    - find / pages 只用标准库，任何 Python 都能跑。
    - PowerShell 终端直接看中文可能乱码，建议加 `--out` 写成 UTF-8 文件后再看。
    - 页码一律指 **PDF 物理页序**；脚本会同时给出报告页脚页码（如 6/255）。
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

PAGE_SPLIT = re.compile(r"\n===== PAGE (\d+) =====\n")


def load_pages(txt: Path) -> dict[int, str]:
    parts = PAGE_SPLIT.split(txt.read_text(encoding="utf-8"))
    return {int(parts[i]): parts[i + 1] for i in range(1, len(parts), 2)}


def page_label(body: str) -> str:
    """提取页脚的报告页码，如 6/255"""
    labels = re.findall(r"(\d{1,3})\s*/\s*(\d{2,3})", body)
    return f"{labels[-1][0]}/{labels[-1][1]}" if labels else "?"


def write(text: str, out: str | None) -> None:
    if out:
        Path(out).write_text(text, encoding="utf-8")
        print(f"written: {out}")
    else:
        print(text)


def cmd_extract(args: argparse.Namespace) -> int:
    import pdfplumber

    src = Path(args.pdf)
    out = []
    with pdfplumber.open(src) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            out.append(f"\n===== PAGE {i} =====\n" + (page.extract_text() or ""))
    text = "".join(out)
    Path(args.out).write_text(text, encoding="utf-8")
    print(f"pages={len(out)} chars={len(text)} -> {args.out}")
    return 0


def cmd_find(args: argparse.Namespace) -> int:
    pages = load_pages(Path(args.txt))
    chunks = []
    total = 0
    for kw in args.kw:
        chunks.append(f"\n########## {kw} ##########")
        hits = 0
        for pno in sorted(pages):
            body = pages[pno]
            for m in re.finditer(re.escape(kw), body):
                if hits >= args.max:
                    break
                hits += 1
                total += 1
                s, e = max(0, m.start() - 150), min(len(body), m.end() + 250)
                chunks.append(f"--- PDF p.{pno}（报告页码 {page_label(body)}）\n{body[s:e].strip()}")
        if hits == 0:
            chunks.append("（无命中）")
    write("\n".join(chunks), args.out)
    print(f"hits={total}")
    return 0


def cmd_pages(args: argparse.Namespace) -> int:
    pages = load_pages(Path(args.txt))
    sel = []
    for n in range(args.frm, args.to + 1):
        if n in pages:
            sel.append(f"\n##### PDF p.{n}（报告页码 {page_label(pages[n])}） #####\n{pages[n]}")
    write("".join(sel), args.out)
    print(f"pages={len(sel)}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="年报 PDF 读取工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("extract", help="PDF 转文本缓存")
    p1.add_argument("--pdf", required=True)
    p1.add_argument("--out", required=True)
    p1.set_defaults(func=cmd_extract)

    p2 = sub.add_parser("find", help="按关键词检索")
    p2.add_argument("--txt", required=True)
    p2.add_argument("--kw", action="append", required=True)
    p2.add_argument("--max", type=int, default=3, help="每个关键词最多几条命中")
    p2.add_argument("--out")
    p2.set_defaults(func=cmd_find)

    p3 = sub.add_parser("pages", help="按页查看")
    p3.add_argument("--txt", required=True)
    p3.add_argument("--frm", type=int, required=True)
    p3.add_argument("--to", type=int, required=True)
    p3.add_argument("--out")
    p3.set_defaults(func=cmd_pages)

    args = ap.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
