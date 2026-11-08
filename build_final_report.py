"""Build the final DOCX for the Yili valuation report.

Reads the frozen report body (Rev02) and the DCF model (Rev07), then writes a
print-ready A4 DOCX with cover, TOC field, figures and four appendix tables.
All statement figures are read from the model; nothing is retyped by hand.
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = Path(__file__).resolve().parents[2]
REPORT_MD = BASE / "报告文件/整合稿/600887_伊利股份估值报告_报告版_20261020_Rev02.md"
MODEL = BASE / "数据文件/DCF模型/600887_DCF模型_20261012_Rev07.xlsx"
OUT_DIR = BASE / "报告文件/最终报告"
OUT_DOCX = OUT_DIR / "600887_伊利股份估值报告_最终版_20261025_Rev02.docx"

BODY_FONT = "宋体"
HEAD_FONT = "黑体"
NUM_FONT = "Times New Roman"


# ---------------------------------------------------------------------------
# model data
# ---------------------------------------------------------------------------

YEAR_COLS = "JKLMNOP"
YEAR_LABELS = ["2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]


def load_model() -> tuple[dict, dict]:
    wb = openpyxl.load_workbook(MODEL, data_only=True)
    fc = wb["报表预测"]
    fi = wb["财务指标"]

    fc_rows = {}
    for row in range(1, fc.max_row + 1):
        label = fc[f"B{row}"].value
        if not isinstance(label, str) or not label.strip():
            continue
        values = [fc[f"{col}{row}"].value for col in YEAR_COLS]
        if all(v is None for v in values):
            continue
        fc_rows[row] = (label.strip(), values)

    fi_rows = {}
    for row in range(1, fi.max_row + 1):
        label = fi[f"B{row}"].value
        if not isinstance(label, str) or not label.strip():
            continue
        values = [fi[f"{col}{row}"].value for col in "IJKLMNO"]
        if all(v is None for v in values):
            continue
        fi_rows[row] = (label.strip(), values)
    return fc_rows, fi_rows


PL_ROWS = [
    (7, "营业总收入", True),
    (8, "营业收入", True),
    (10, "营业成本", False),
    (12, "营业税金及附加", False),
    (14, "销售费用", False),
    (16, "管理费用", False),
    (18, "研发费用", False),
    (20, "财务费用", False),
    (21, "其他收益", False),
    (22, "资产减值损失（含信用减值）", False),
    (23, "资产处置收益", False),
    (24, "公允价值变动收益", False),
    (25, "投资收益", False),
    (26, "营业利润（EBIT）", True),
    (27, "营业外收入", False),
    (28, "营业外支出", False),
    (29, "利润总额", True),
    (30, "所得税", False),
    (31, "净利润", True),
    (32, "少数股东损益", False),
    (34, "归母净利润", True),
    (35, "折旧摊销", False),
    (36, "EBITDA", True),
    (37, "EPS（元/股）", False),
]

BS_ROWS = [
    (40, "流动资产", True),
    (41, "　现金及现金等价物", False),
    (42, "　应收账款", False),
    (43, "　应收票据", False),
    (44, "　存货", False),
    (45, "　预付账款", False),
    (46, "　合同资产", False),
    (47, "　其他流动资产", False),
    (48, "非流动资产", True),
    (49, "　长期投资", False),
    (50, "　固定资产", False),
    (51, "　无形资产", False),
    (52, "　其他长期资产", False),
    (53, "资产合计", True),
    (54, "流动负债", True),
    (55, "　短期借款", False),
    (56, "　应付账款", False),
    (57, "　其他流动负债", False),
    (58, "非流动负债", True),
    (59, "　长期借款", False),
    (60, "　其他长期负债", False),
    (61, "负债合计", True),
    (62, "少数股东权益", False),
    (63, "归属母公司股东权益", True),
    (64, "负债和股东权益合计", True),
]

CF_ROWS = [
    (67, "净利润", True),
    (68, "折旧摊销", False),
    (69, "营运资金变动", False),
    (70, "其他", False),
    (71, "经营活动现金流", True),
    (72, "资本支出", False),
    (73, "投资变动", False),
    (74, "其他", False),
    (75, "投资活动现金流", True),
    (76, "银行借款", False),
    (77, "分配股利", False),
    (79, "筹资活动现金流", True),
    (80, "净现金流", True),
    (81, "期初现金及现金等价物", False),
    (82, "期末现金及现金等价物", True),
]

FIN_GROUPS = [
    (
        "成长能力",
        [
            (8, "营业收入增长率", "pct"),
            (9, "营业利润增长率", "pct"),
            (10, "归母净利润增长率", "pct"),
            (11, "息税前利润增长率", "pct"),
            (12, "息税折旧前利润增长率", "pct"),
            (13, "EPS 增长率", "pct"),
        ],
    ),
    (
        "获利能力",
        [
            (15, "息税前利润率", "pct"),
            (16, "营业利润率", "pct"),
            (17, "毛利率", "pct"),
            (18, "归母净利润率", "pct"),
            (19, "ROE", "pct"),
            (20, "ROA", "pct"),
        ],
    ),
    (
        "偿债能力",
        [
            (22, "资产负债率", "pct_dec"),
            (23, "净负债权益比", "pct_dec"),
            (24, "流动比率", "x"),
        ],
    ),
    (
        "营运能力",
        [
            (26, "总资产周转率", "turn"),
            (27, "应收账款周转天数", "days"),
            (28, "应付账款周转天数", "days"),
        ],
    ),
    (
        "费用率",
        [
            (30, "销售费用率", "pct"),
            (31, "管理费用率", "pct"),
            (32, "研发费用率", "pct"),
            (33, "财务费用率", "pct"),
        ],
    ),
    (
        "每股指标",
        [
            (35, "每股收益（最新摊薄）", "money"),
            (36, "每股经营现金流", "money"),
            (37, "每股净资产", "money"),
            (38, "每股股息", "money"),
        ],
    ),
    (
        "估值比率",
        [
            (40, "PE（最新摊薄）", "x"),
            (41, "PB（最新摊薄）", "x"),
            (42, "EV/EBITDA", "x"),
            (43, "价格/现金流", "x"),
        ],
    ),
]


def fmt_amount(value) -> str:
    if value is None:
        return "—"
    if abs(value) < 0.005:
        return "0.00"
    return f"{value:,.2f}"


def fmt_fin(value, kind: str) -> str:
    if value is None:
        return "—"
    if kind == "pct":
        return f"{value:.2f}%"
    if kind == "pct_dec":
        return f"{value * 100:.2f}%"
    if kind == "turn":
        return f"{value / 100:.2f}"
    if kind in {"x", "days", "money"}:
        return f"{value:,.2f}"
    raise ValueError(kind)


# ---------------------------------------------------------------------------
# low-level docx helpers
# ---------------------------------------------------------------------------


def set_run_font(run, size: float, bold: bool = False, color=None, font=BODY_FONT):
    run.font.name = NUM_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), NUM_FONT)
    rfonts.set(qn("w:hAnsi"), NUM_FONT)
    rfonts.set(qn("w:eastAsia"), font)


def set_style_font(style, size: float, bold: bool = False, font=BODY_FONT, color=None):
    style.font.name = NUM_FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), NUM_FONT)
    rfonts.set(qn("w:hAnsi"), NUM_FONT)
    rfonts.set(qn("w:eastAsia"), font)


def add_paragraph_border(paragraph, color="808080", size=6):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    ppr.append(pbdr)


def add_field(paragraph, instruction: str, placeholder: str = "1", size=8):
    run = paragraph.add_run()
    set_run_font(run, size, font=BODY_FONT)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)


def add_toc_field(doc):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-2" \\h \\z \\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "目录将在 Word 中自动更新"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)


def shade_cell(cell, fill="F2F2F2"):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def set_table_borders(table):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "12")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    for edge in ("left", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblpr.append(borders)

    cellmar = OxmlElement("w:tblCellMar")
    for side, width in (("top", 20), ("left", 60), ("bottom", 20), ("right", 60)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(width))
        el.set(qn("w:type"), "dxa")
        cellmar.append(el)
    tblpr.append(cellmar)


def set_col_widths(table, widths_cm):
    table.autofit = False
    tblpr = table._tbl.tblPr
    tblw = OxmlElement("w:tblW")
    tblw.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tblw.set(qn("w:type"), "dxa")
    tblpr.append(tblw)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    grid = table._tbl.tblGrid
    for index, width in enumerate(widths_cm):
        grid.gridCol_lst[index].set(qn("w:w"), str(int(width * 567)))
    for row in table.rows:
        for index, width in enumerate(widths_cm):
            row.cells[index].width = Cm(width)


def set_header_row(table):
    trpr = table.rows[0]._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    trpr.append(header)
    for cell in table.rows[0].cells:
        tcpr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        borders.append(bottom)
        tcpr.append(borders)


def write_cell(cell, text, bold=False, align="right", size=7.5, indent_cm=0.0,
               color=None, font=BODY_FONT):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.left_indent = Cm(indent_cm)
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    run = paragraph.add_run(str(text))
    set_run_font(run, size, bold=bold, color=color, font=font)


def add_data_table(doc, headers, rows, widths, font_size=7.5, group_rows=None,
                   bold_rows=None, align_rest="right"):
    """rows: list of list[str]; bold_rows: set of row indices (0-based in rows)."""
    group_rows = group_rows or {}
    bold_rows = bold_rows or set()
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    set_col_widths(table, widths)
    for index, header in enumerate(headers):
        write_cell(table.rows[0].cells[index], header, bold=True,
                   align="left" if index == 0 or align_rest == "left" else "right",
                   size=font_size)
    set_header_row(table)

    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        trpr = table.rows[-1]._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        trpr.append(cant_split)
        if row_index in group_rows:
            merged = cells[0].merge(cells[-1])
            write_cell(merged, row[0], bold=True, align="left", size=font_size)
            shade_cell(merged, "F2F2F2")
            continue
        for index, value in enumerate(row):
            label = value
            indent = 0.0
            if index == 0:
                if label.startswith("　"):
                    indent = 0.45
                    label = label.replace("　", "", 1)
                write_cell(cells[index], label,
                           bold=row_index in bold_rows,
                           align="left", size=font_size, indent_cm=indent)
            else:
                write_cell(cells[index], label,
                           bold=row_index in bold_rows,
                           align=align_rest, size=font_size)
    return table


def add_caption(doc, text, kind="table"):
    paragraph = doc.add_paragraph(style="CaptionCN")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6 if kind == "table" else 3)
    paragraph.paragraph_format.space_after = Pt(3 if kind == "table" else 10)
    paragraph.paragraph_format.keep_with_next = kind == "table"
    run = paragraph.add_run(text)
    set_run_font(run, 9, bold=True, font=BODY_FONT)


def add_source_note(doc, text):
    paragraph = doc.add_paragraph(style="SourceNote")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run(text)
    set_run_font(run, 7.5, color=RGBColor(0x59, 0x59, 0x59), font=BODY_FONT)


def add_body_paragraph(doc, text, first_indent=True):
    paragraph = doc.add_paragraph(style="BodyCN")
    if first_indent:
        paragraph.paragraph_format.first_line_indent = Pt(21)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, 10.5, bold=True)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, 10.5)


# ---------------------------------------------------------------------------
# document construction
# ---------------------------------------------------------------------------


def setup_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.header_distance = Cm(1.1)
    section.footer_distance = Cm(0.9)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    set_style_font(normal, 10.5)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    for name in ("Heading 1", "Heading 2"):
        style = doc.styles[name]
        set_style_font(style, 16 if name == "Heading 1" else 12, bold=True,
                       font=HEAD_FONT, color=RGBColor(0, 0, 0))
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.space_before = Pt(18 if name == "Heading 1" else 12)
        style.paragraph_format.space_after = Pt(8 if name == "Heading 1" else 6)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.2

    body = doc.styles.add_style("BodyCN", 1)
    body.base_style = doc.styles["Normal"]
    set_style_font(body, 10.5)
    body.paragraph_format.line_spacing = 1.5
    body.paragraph_format.space_after = Pt(6)

    caption = doc.styles.add_style("CaptionCN", 1)
    caption.base_style = doc.styles["Normal"]
    set_style_font(caption, 9, bold=True)
    caption.paragraph_format.line_spacing = 1.0

    source = doc.styles.add_style("SourceNote", 1)
    source.base_style = doc.styles["Normal"]
    set_style_font(source, 7.5, color=RGBColor(0x59, 0x59, 0x59))
    source.paragraph_format.line_spacing = 1.0
    source.paragraph_format.space_after = Pt(10)

    bullet = doc.styles["List Bullet"]
    set_style_font(bullet, 10.5)
    bullet.paragraph_format.line_spacing = 1.5
    bullet.paragraph_format.space_after = Pt(3)

    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = paragraph.add_run("伊利股份（600887.SH）｜利润下滑是减值，经营稳健是底色")
    set_run_font(run, 7.5, color=RGBColor(0x59, 0x59, 0x59))
    add_paragraph_border(paragraph, size=4)

    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("第 ")
    set_run_font(run, 8, color=RGBColor(0x59, 0x59, 0x59))
    add_field(paragraph, "PAGE", "1", size=8)
    run = paragraph.add_run(" 页 / 共 ")
    set_run_font(run, 8, color=RGBColor(0x59, 0x59, 0x59))
    add_field(paragraph, "NUMPAGES", "1", size=8)
    run = paragraph.add_run(" 页")
    set_run_font(run, 8, color=RGBColor(0x59, 0x59, 0x59))
    return doc


def build_cover(doc):
    for _ in range(5):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(18)
    run = title.add_run("利润下滑是减值，经营稳健是底色")
    set_run_font(run, 24, bold=True, font=HEAD_FONT)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_before = Pt(10)
    run = subtitle.add_run("伊利股份（600887.SH）估值研究报告")
    set_run_font(run, 14, bold=False, font=HEAD_FONT)

    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(14)
    add_paragraph_border(rule, color="000000", size=8)

    for _ in range(2):
        doc.add_paragraph()

    metrics = [
        ("评级", "买入"),
        ("目标区间", "29.19–38.30 元"),
        ("现价（2026-10-09）", "27.66 元"),
        ("报告日期", "2026 年 10 月 25 日"),
    ]
    table = doc.add_table(rows=len(metrics), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_col_widths(table, [5.0, 7.0])
    for index, (key, value) in enumerate(metrics):
        write_cell(table.rows[index].cells[0], key, bold=False, align="right",
                   size=11)
        write_cell(table.rows[index].cells[1], value, bold=True, align="left",
                   size=12)

    for _ in range(5):
        doc.add_paragraph()

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("数据截止 2026-10-09｜金额单位：亿元（除特别注明）")
    set_run_font(run, 9, color=RGBColor(0x59, 0x59, 0x59))

    doc.add_page_break()


def build_toc(doc):
    heading = doc.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    heading.paragraph_format.space_before = Pt(6)
    heading.paragraph_format.space_after = Pt(14)
    run = heading.add_run("目　录")
    set_run_font(run, 16, bold=True, font=HEAD_FONT)
    add_toc_field(doc)
    doc.add_page_break()


def parse_inline_table(lines):
    text = "\n".join(lines)

    def clean(cell: str) -> str:
        return html.unescape(re.sub(r"<[^>]+>", "", cell)).strip()

    header = []
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.S):
        th = re.findall(r"<th[^>]*>(.*?)</th>", tr, flags=re.S)
        td = re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S)
        if th:
            header = [clean(cell) for cell in th]
        if td:
            rows.append([clean(cell) for cell in td])
    return header, rows


def add_inline_table(doc, header, rows):
    widths = [4.0] + [1.55] * (len(header) - 1)
    table = doc.add_table(rows=1, cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    set_col_widths(table, widths)
    for index, text in enumerate(header):
        write_cell(table.rows[0].cells[index], text, bold=True,
                   align="left" if index == 0 else "right", size=7)
    set_header_row(table)
    for row in rows:
        cells = table.add_row().cells
        for index, text in enumerate(row):
            write_cell(cells[index], text, align="left" if index == 0 else "right",
                       size=7)
    return table


def build_body(doc):
    lines = REPORT_MD.read_text(encoding="utf-8").splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith("## 一、结论"))
    lines = lines[start:]

    index = 0
    while index < len(lines):
        line = lines[index].rstrip()
        if not line.strip() or line.strip() == "---":
            index += 1
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=2)
            index += 1
            continue
        if line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=1)
            index += 1
            continue
        if line.startswith("!["):
            match = re.match(r"!\[[^\]]*\]\(([^)]+)\)", line)
            if match:
                picture = doc.add_paragraph()
                picture.alignment = WD_ALIGN_PARAGRAPH.CENTER
                picture.paragraph_format.space_before = Pt(8)
                picture.paragraph_format.space_after = Pt(2)
                picture.add_run().add_picture(match.group(1), width=Cm(15.4))
            index += 1
            continue
        if line.strip().startswith("<div"):
            text = re.sub(r"<[^>]+>", "", line).strip()
            add_caption(doc, text, kind="figure" if text.startswith("图") else "table")
            index += 1
            continue
        if line.strip().startswith("<table"):
            block = [line]
            while "</table>" not in block[-1] and index + 1 < len(lines):
                index += 1
                block.append(lines[index])
            header, rows = parse_inline_table(block)
            add_inline_table(doc, header, rows)
            index += 1
            continue
        if line.startswith("- "):
            paragraph = doc.add_paragraph(style="List Bullet")
            text = line[2:].strip()
            parts = re.split(r"(\*\*.+?\*\*)", text)
            for part in parts:
                if not part:
                    continue
                if part.startswith("**") and part.endswith("**"):
                    run = paragraph.add_run(part[2:-2])
                    set_run_font(run, 10.5, bold=True)
                else:
                    run = paragraph.add_run(part)
                    set_run_font(run, 10.5)
            index += 1
            continue
        add_body_paragraph(doc, line.strip())
        index += 1


def build_appendix(doc, fc_rows, fi_rows):
    doc.add_page_break()
    doc.add_heading("附录　财务报表与财务指标", level=1)

    definitions = [
        ("附表1　利润表（2024A–2030E，亿元）", "利润表", PL_ROWS, "fc",
         "数据来源：DCF 模型 Rev07《报表预测》页。2024A–2025A 为年报实际值，2026E–2030E 为模型预测值。"),
        ("附表2　资产负债表（2024A–2030E，亿元）", "资产负债表", BS_ROWS, "fc",
         "数据来源：DCF 模型 Rev07《报表预测》页。资产合计＝负债和股东权益合计，7 期均勾稽。"),
        ("附表3　现金流量表（2024A–2030E，亿元）", "现金流量表", CF_ROWS, "fc",
         "数据来源：DCF 模型 Rev07《报表预测》页。期末现金＝期初现金＋净现金流。"),
    ]

    for caption, sheet_name, row_defs, source, note in definitions:
        add_caption(doc, caption, kind="table")
        header = ["项目"] + YEAR_LABELS
        rows = []
        bold_rows = set()
        for row_index, (row_number, label, bold) in enumerate(row_defs):
            if row_number not in fc_rows:
                raise KeyError(f"missing row {row_number} in 报表预测")
            values = fc_rows[row_number][1]
            rows.append([label] + [fmt_amount(v) for v in values])
            if bold:
                bold_rows.add(row_index)
        add_data_table(doc, header, rows, [5.0] + [1.65] * 7,
                       font_size=7.5, bold_rows=bold_rows)
        doc.add_page_break()

    add_caption(doc, "附表4　财务指标（2024A–2030E）", kind="table")
    header = ["指标"] + YEAR_LABELS
    rows = []
    group_rows = {}
    for group_name, items in FIN_GROUPS:
        group_rows[len(rows)] = group_name
        rows.append([group_name] + [""] * 7)
        for row_number, label, kind in items:
            if row_number not in fi_rows:
                raise KeyError(f"missing row {row_number} in 财务指标")
            values = fi_rows[row_number][1]
            rows.append([label] + [fmt_fin(v, kind) for v in values])
    add_data_table(doc, header, rows, [5.0] + [1.65] * 7,
                   font_size=7.5, group_rows=group_rows)


def build_rating(doc):
    heading = doc.add_heading("公司投资评级", level=1)
    heading.paragraph_format.page_break_before = True
    add_caption(doc, "附表5　投资评级标准", kind="table")
    rows = [
        ["买入", "预计该公司股价在未来 6–12 个月内超越基准指数 20% 以上。"],
        ["增持", "预计该公司股价在未来 6–12 个月内超越基准指数 10%–20%。"],
        ["中性", "预计该公司股价在未来 6–12 个月内相对基准指数变动幅度在 −10%–10% 之间。"],
        ["减持", "预计该公司股价在未来 6–12 个月内相对基准指数跌幅在 10% 以上。"],
        ["未有评级", "因无法获取必要的资料或者其他原因，未能给出明确的投资评级。"],
    ]
    add_data_table(doc, ["评级", "定义"], rows, [2.6, 13.9], font_size=9,
                   align_rest="left")


def build_disclaimer(doc):
    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(26)
    heading.paragraph_format.space_after = Pt(4)
    run = heading.add_run("免责声明")
    set_run_font(run, 10, bold=True, font=HEAD_FONT)

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.4
    run = paragraph.add_run(
        "本报告为内部研究练习稿，基于公开资料与作者测算，不构成任何投资建议。"
        "报告中的预测与估值依赖既定假设，实际结果可能与之存在重大差异。"
    )
    set_run_font(run, 8, color=RGBColor(0x59, 0x59, 0x59))


def validate(fc_rows, fi_rows):
    for col_index, year in enumerate(YEAR_LABELS):
        assets = fc_rows[53][1][col_index]
        liab_equity = fc_rows[64][1][col_index]
        assert abs(assets - liab_equity) < 1e-6, f"balance sheet mismatch {year}"
        net_cash = fc_rows[80][1][col_index]
        begin = fc_rows[81][1][col_index]
        end = fc_rows[82][1][col_index]
        assert abs(begin + net_cash - end) < 1e-6, f"cash flow mismatch {year}"
    assert abs(fc_rows[8][1][0] - 1153.9331097669) < 1e-6
    assert abs(fc_rows[34][1][6] - 96.99253074782068) < 1e-6
    assert abs(fi_rows[19][1][6] - 16.272335328813554) < 1e-6


def set_bullet_color(doc):
    numbering = doc.part.numbering_part.element
    for level in numbering.iter(qn("w:lvl")):
        num_format = level.find(qn("w:numFmt"))
        if num_format is None or num_format.get(qn("w:val")) != "bullet":
            continue
        rpr = level.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            level.append(rpr)
        color = rpr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            rpr.append(color)
        color.set(qn("w:val"), "000000")


def main():
    if OUT_DOCX.exists() and "--force" not in sys.argv:
        raise SystemExit(
            f"{OUT_DOCX.name} 已存在。该文件是定稿，可能含手工修改；"
            "确需重写请加 --force。"
        )
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fc_rows, fi_rows = load_model()
    validate(fc_rows, fi_rows)

    doc = setup_document()
    doc.core_properties.title = "利润下滑是减值，经营稳健是底色——伊利股份估值研究报告"
    doc.core_properties.subject = "伊利股份（600887.SH）估值研究"
    doc.core_properties.author = "内部研究"
    doc.core_properties.comments = "数据截止 2026-10-09；模型 Rev07。"

    build_cover(doc)
    build_toc(doc)
    build_body(doc)
    build_appendix(doc, fc_rows, fi_rows)
    build_rating(doc)
    build_disclaimer(doc)
    set_bullet_color(doc)
    doc.save(OUT_DOCX)
    print(f"saved: {OUT_DOCX}")
    print(f"paragraphs={len(doc.paragraphs)} tables={len(doc.tables)}")


if __name__ == "__main__":
    main()
