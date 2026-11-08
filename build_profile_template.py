#!/usr/bin/env python3
"""生成《公司画像与行业情况》单页签资料表。

Rev01 起改为简表：7 个模块、每模块一段内容 + 来源页码，不再拆成 52 条待填要点。
内容以填报人阅读 2026 年半年报所得的结论为主；标 ★ 的行是脚本补充的公司基本信息，
风险一栏留空待填，仅附公司披露的风险条目供参考。

用法：
    python build_profile_template.py
    python build_profile_template.py --rev Rev02
    python build_profile_template.py --out "<输出目录>" --asof 20261010

约定：同名文件已存在时中止，请提高 --rev 重出，旧版不覆盖。
"""

from __future__ import annotations

import argparse
import math
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_OUT = Path(__file__).resolve().parents[2] / "数据文件" / "公司画像与行业分析"
STOCK = "600887"
STOCK_NAME = "伊利股份"

TITLE_FILL = PatternFill("solid", fgColor="1B3A5C")
HEAD_FILL = PatternFill("solid", fgColor="2E5C8A")
INPUT_FILL = PatternFill("solid", fgColor="FFFDF5")
TODO_FILL = PatternFill("solid", fgColor="FDE9E9")
NOTE_FILL = PatternFill("solid", fgColor="F2F2F2")
TITLE_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=13)
HEAD_FONT = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
BODY_FONT = Font(name="微软雅黑", size=10.5)
HINT_FONT = Font(name="微软雅黑", size=10, color="404040")
SRC_FONT = Font(name="微软雅黑", size=9.5, color="595959")
SIDE = Side(style="thin", color="BFBFBF")
BORDER = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEADERS = ["模块", "内容", "来源（页码为 PDF 页序）", "备注"]
WIDTHS = [17, 84, 33, 26]

TITLE = f"{STOCK_NAME}（{STOCK}.SH）公司画像与行业情况"
HINT = ("内容以阅读《2026 年半年度报告》所得结论为主；标 ★ 的行由脚本按年报补充，未经确认。"
        "红色底为待填 / 待核。页码 = PDF 物理页序，括号内为公司报告页脚页码。"
        "行业优势第 (5) 条属分析师口径（开源证券研报），与公司口径分开列示。")

ROWS: list[tuple[str, str, str, str, bool]] = [
    (
        "公司基本情况 ★",
        "全称：内蒙古伊利实业集团股份有限公司（简称「伊利股份」，外文 YILI）\n"
        "股票：A 股 600887，上海证券交易所\n"
        "上市：1996-01-25 首次公开发行 1,715 万股，1996-03-12 挂牌交易\n"
        "法定代表人：潘刚；董事会秘书：邱向敏\n"
        "注册地址 / 办公地址：内蒙古自治区呼和浩特市敕勒川乳业开发区伊利大街 1 号（邮编 010110）\n"
        "员工：62,666 人（母公司 26,095 + 主要子公司 36,571，2025 年末）",
        "2025 年报 p.6–7（报告页码 2–3）：公司简介、股票简况\n"
        "2025 年报 p.46（报告页码 42）：员工情况\n"
        "2025 年报 p.98（报告页码 94）：公司概况（上市日期）",
        "★ 我按年报补齐的基本信息，请复核；如写进报告需再引一次页码。",
        False,
    ),
    (
        "经营范围（主营业务）",
        "公司主要从事各类乳品及健康饮品的生产与销售，经营液态奶、乳饮料、奶粉、酸奶、冷冻饮品、"
        "奶酪、乳脂、高端乳蛋白、包装饮用水等产品。国内为主、海外为辅。",
        "2026 半年报 p.8「第三节 一、(一) 经营范围」（报告页码 5/215）",
        "与 2025 年报 p.10（6/255）一致，仅年报无「高端乳蛋白」项。",
        False,
    ),
    (
        "经营模式",
        "（1）六大产品业务群（液态奶、婴幼儿营养品、成人营养品、冷冻饮品、酸奶、奶酪），事业部制，"
        "各事业部在本领域内开展产、供、销\n"
        "（2）采购：原料奶建立利益联结机制和嵌入式服务；其他原辅材料通过招标集中采购\n"
        "（3）生产与物流：全球布局生产基地（亚洲、大洋洲、欧洲），大数据算法科学规划仓储布局与"
        "物流线路，与第三方物流服务商合作\n"
        "（4）营销与销售：母子品牌管理体系，数字化营销和内容营销，经销为主、直营为辅，"
        "覆盖商超、便利店、零食店、会员店、电商、即时零售、餐饮等全渠道\n"
        "（5）海外业务：实行战略管控，海外子公司生产销售冷冻饮品、乳铁蛋白、酪蛋白、乳清蛋白、"
        "乳脂、奶粉、奶酪等",
        "2026 半年报 p.8「(二) 经营模式」（报告页码 5/215）\n"
        "同一章节见 2025 年报 p.10（6/255）",
        "括号内细节是我按半年报原文补的，可留可删。",
        False,
    ),
    (
        "主要业绩驱动因素",
        "（1）消费者导向：深入洞察消费者需求，推出新品，推动乳制品结构持续升级\n"
        "（2）产业链关系：深化上游原料奶供应商利益联结机制，强化与下游经销商的合作关系\n"
        "（3）数智融合：数业深度融合、升级数智化业务运营体系，推动人工智能落地应用",
        "2026 半年报 p.8–9「(三) 主要业绩驱动因素」（报告页码 5–6/215）\n"
        "同期新品收入占比 15.8%（2025 年报为 16.4%）",
        "新品占比可作为该驱动因素的量化锚点。",
        False,
    ),
    (
        "行业情况—优势",
        "（1）宏观：社零保持增长、消费者信心回升、CPI 温和回升、PMI 向好；线上消费占比提升；"
        "乡村消费增速快于城镇；人工智能发展迅速\n"
        "（2）行业趋势：乳品供需矛盾进一步缓解（生鲜乳价格企稳）；高品质产品与基础产品需求同步增长；"
        "品类创新活跃、新产品差异化增强；乳精深加工产业快速发展\n"
        "（3）品类：消费者健康意识提升；户外文旅消费等场景持续发展，带动茶饮、咖啡、烘焙、西餐等业态\n"
        "（4）政策：商务部对原产于欧盟的进口乳制品实施反补贴（2025-12 临时，2026-02 终裁，为期 5 年）；"
        "2026-07《扩大消费「十五五」规划》批复\n"
        "（5）生鲜乳价格：底部基本确立，原料价格温和上涨 〔分析师口径〕",
        "（1）–（4）：2026 半年报 p.9–10「(四) 行业情况说明」（报告页码 6–7/215）\n"
        "（5）开源证券《食品饮料行业投资策略》2026-09-17 p.16（分析师口径）",
        "第 (5) 条不是公司口径，见下方红行。",
        False,
    ),
    (
        "行业情况—风险",
        "（1）供应风险：美伊冲突带来石油价格上涨后的物流成本上升，进口原辅材料价格整体上升；"
        "GLP-1 类药物市场快速发展，带动进口乳清蛋白价格快速上涨\n"
        "（2）财务风险：海外业务受汇率、贸易政策及关税波动影响加大\n"
        "（3）产品质量风险：确保产品质量与安全",
        "2026 半年报 p.25「(一) 可能面对的风险」（报告页码 22/215）\n"
        "同期亦见 2025 年报 p.36（报告页码 32/255）",
        "均为公司口径。半年报较年报有更新：年报称进口原辅材料价格「稳中有降」，"
        "半年报改为「整体呈现上行趋势」，写报告时以半年报为准。",
        False,
    ),
    (
        "口径标注（公司口径 vs 分析师口径）",
        "① 「奶价底部基本确立、原料价格温和上涨」出自开源证券《消费筑底，寒梅将绽——行业投资策略》"
        "（2026-09-17，PDF p.16），原文为「奶价底部基本确立，预计 2026 年奶价企稳回升」"
        "「预计 2026 年下半年奶价呈现温和上涨趋势」；公司口径为「国内生鲜乳价格企稳」"
        "（2026 半年报 p.9）。按证据规则，研报仅作市场一致预期对照，写进报告须注明「（分析师口径）」，"
        "不得与公司口径混写。\n"
        "② 「高端乳蛋白」是半年报口径，2025 年报无此项。\n"
        "③ 员工数为 2025 年末数，2026 半年报未重述。",
        "2026 半年报 p.9\n开源证券行业策略 2026-09-17 p.16\n2025 年报 p.46",
        "另：资料名「2026Q2食品饮料行业研报.pdf」与内容不符——实际是开源证券 2026-09-17 的"
        "行业投资策略（27 页），建议改名。",
        True,
    ),
]


def text_lines(text: str, width: int) -> int:
    """估算文本在给定列宽下占几行（列宽单位为半角字符）"""
    if not text:
        return 1
    per_line = max(4.0, width / 2.05)
    lines = 0
    for seg in text.split("\n"):
        eff = sum(1 if ord(ch) > 127 else 0.55 for ch in seg)
        lines += max(1, math.ceil(eff / per_line))
    return lines


def main() -> int:
    ap = argparse.ArgumentParser(description="生成公司画像与行业情况简表")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="输出目录")
    ap.add_argument("--asof", default=date.today().strftime("%Y%m%d"), help="日期标签 YYYYMMDD")
    ap.add_argument("--rev", default="Rev01", help="版本号，重出时提高")
    args = ap.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{STOCK}_公司画像与行业分析资料表_{args.asof}_{args.rev}.xlsx"
    if path.exists():
        print(f"[中止] 同名文件已存在：{path.name}\n        请用 --rev Rev02 重出，旧版保留不覆盖。")
        return 1

    wb = Workbook()
    ws = wb.active
    ws.title = "公司画像与行业情况"
    last = get_column_letter(len(HEADERS))

    ws.merge_cells(f"A1:{last}1")
    c = ws.cell(1, 1, TITLE)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{last}2")
    c = ws.cell(2, 1, HINT)
    c.font = HINT_FONT
    c.fill = NOTE_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    for i, name in enumerate(HEADERS, start=1):
        c = ws.cell(3, i, name)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[3].height = 20

    for n, (module, content, source, note, todo) in enumerate(ROWS, start=1):
        r = 3 + n
        fill = TODO_FILL if todo else INPUT_FILL
        for i, val in enumerate([module, content, source, note], start=1):
            c = ws.cell(r, i, val)
            c.font = BODY_FONT if i <= 2 else SRC_FONT
            c.alignment = WRAP if i != 1 else Alignment(
                horizontal="left", vertical="top", wrap_text=True)
            c.border = BORDER
            if i in (2, 3):
                c.fill = fill
        ws.cell(r, 1).font = Font(name="微软雅黑", bold=True, size=10.5)
        ws.row_dimensions[r].height = max(
            text_lines(content, WIDTHS[1]), text_lines(source, WIDTHS[2]),
            text_lines(note, WIDTHS[3])) * 15 + 8

    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B4"
    ws.sheet_view.showGridLines = False
    wb.save(path)

    print(f"已生成：{path}")
    print(f"  模块 {len(ROWS)} 行；空行（待填）{sum(1 for r in ROWS if not r[1])} 行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
