"""Build a 深圳大学实验报告 (DOCX) from a simple Markdown content file.

The script fills the school template that ships in ``assets/base.docx``:
cover fields, the seven report sections inside the template table, the
preview-quiz page and the raw-data page.

Usage
-----
    python report_builder.py --content content.md --out report.docx \
        [--kind preview|full] [--meta meta.json] [--quiz quiz1.png --quiz quiz2.png] \
        [--rawdata-image photo.jpg] [--experiment NAME --location ROOM \
         --date 2026-09-23 --submit-date 2026-09-29]

``--quiz`` / ``--rawdata-image`` 都可重复；路径相对 ``--content`` 所在目录解析。
--quiz 的图自动排进第 2 页（多张纵向堆叠、只占一页），--rawdata-image 的手写照片
自动排到最后一页、单独起页。

Content file format (full contract in reference/content-guide.md)
----------------------------------------------------------------
    # 一、实验目的
    1 了解……
    # 二、实验原理
    ## 1 玻尔原子理论。
    正文段落，首行自动缩进两个字。
    $$h\\nu=E_2-E_1$$ (1)
    ![图1 能级跃迁图。](figures/fig01.png)
    表1 峰值电压与第一激发电位表
    | 峰值次数 n | 峰值电压 Um (V) |
    | --- | --- |
    | 1 | 9.1 |
    # 原始数据记录表
    | 阳极电压 Ua (V) | 阳极电流 Ia (μA) |
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from latex2omml import latex_to_omml  # noqa: E402

# --------------------------------------------------------------------------
# constants
# --------------------------------------------------------------------------
SECTION_ORDER = [
    ("purpose", ("实验目的",), "一、实验目的"),
    ("principle", ("实验原理",), "二、实验原理"),
    ("apparatus", ("实验仪器",), "三、实验仪器"),
    ("procedure", ("实验内容与步骤", "实验内容及步骤", "实验步骤"), "四、实验内容与步骤"),
    ("data", ("数据处理",), "五、数据处理"),
    ("conclusion", ("结果陈述", "结果与讨论", "实验总结"), "六、结果陈述"),
    ("questions", ("思考题",), "七、思考题"),
]

BODY_PT = 10.5        # 五号
CAPTION_PT = 9.0      # 小五
HEADING_PT = 10.5
TABLE_PT = 9.0
BODY_LINE = 1.0       # 参考报告正文为单倍行距
COVER_LINE_TWIPS = 1200   # 封面每行固定 60 pt（≈ 模板的 3.75 倍行距）
MAX_IMAGE_CM = 14.4   # text column width (A4 minus the template margins)
MAX_IMAGE_H_CM = 9.0  # tall figures push their caption onto the next page
RAW_IMAGE_CM = 15.0   # 原始数据记录表照片/预习试卷页：尽量占满整页（批注24）
PAGE_IMAGE_H_CM = 23.0
QUIZ_STACK_GAP_CM = 0.2   # 第 2 页放多张试卷截图时，每张之间预留的行高


# --------------------------------------------------------------------------
# run / paragraph / table helpers
# --------------------------------------------------------------------------
def set_run(run, size=BODY_PT, bold=False, cn="宋体", en="Times New Roman",
            italic=False, underline=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.name = en
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    fonts.set(qn("w:ascii"), en)
    fonts.set(qn("w:hAnsi"), en)
    fonts.set(qn("w:eastAsia"), cn)
    fonts.set(qn("w:cs"), en)
    return run


def set_paragraph(p, align=None, first_line_chars=0, line_spacing=BODY_LINE,
                  space_before=0, space_after=0, keep_with_next=False):
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    if line_spacing:
        pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    ppr = p._p.get_or_add_pPr()
    existing = ppr.find(qn("w:ind"))
    if existing is not None:
        ppr.remove(existing)
    if first_line_chars:
        # get_or_add_ind keeps the child order Word requires inside w:pPr
        ind = ppr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(first_line_chars * BODY_PT * 20)))
    if keep_with_next:
        ppr.get_or_add_keepNext()
    return p


def add_run(p, text, **kw):
    return set_run(p.add_run(text), **kw)


def set_paragraph_mark_size(paragraph, pt):
    """把段落标记（¶）本身的字号设成 pt。

    空段落的行高由"段落标记"的字符格式决定；把标记设成 1 pt，空段落就几乎不占
    行高。第 2 页（预习试卷截图页）靠它留白：用户往上粘一张接近整页的图时，
    剩下这 1 pt 也不会把图挤到下一页、凭空多出一张白页。
    注意**不要**同时把行距设成"固定值(exact)"——那会把行内图片裁掉。
    """
    ppr = paragraph._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)          # pPr 里 w:rPr 排在靠后位置，无 sectPr 时追加即合规
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set(qn("w:val"), str(max(2, int(round(pt * 2)))))
    return paragraph


TBLPR_ORDER = (
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange",
)


def set_tbl_prop(tbl_pr, tag, element):
    """Insert/replace a w:tblPr child at its schema position."""
    existing = tbl_pr.find(qn(tag))
    if existing is not None:
        tbl_pr.remove(existing)
    position = TBLPR_ORDER.index(tag)
    for child in tbl_pr:
        child_tag = "w:" + child.tag.split("}")[-1]
        if child_tag in TBLPR_ORDER and TBLPR_ORDER.index(child_tag) > position:
            child.addprevious(element)
            return
    tbl_pr.append(element)


def set_cell_borders(table, sz=4):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:%s" % edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    set_tbl_prop(table._tbl.tblPr, "w:tblBorders", borders)


def set_table_full_width(table):
    tbl_pr = table._tbl.tblPr
    width = OxmlElement("w:tblW")
    width.set(qn("w:w"), "5000")
    width.set(qn("w:type"), "pct")
    set_tbl_prop(tbl_pr, "w:tblW", width)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "autofit")
    set_tbl_prop(tbl_pr, "w:tblLayout", layout)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER


def cell_vertical_center(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tc_pr.append(va)


# --------------------------------------------------------------------------
# content parsing
# --------------------------------------------------------------------------
class Block:
    def __init__(self, kind, **kw):
        self.kind = kind
        self.__dict__.update(kw)

    def __repr__(self):  # pragma: no cover - debugging aid
        return "<%s %s>" % (self.kind, {k: v for k, v in self.__dict__.items() if k != "kind"})


EQ_RE = re.compile(r"^\$\$(?P<body>.+?)\$\$\s*(?:\((?P<num>[^)]*)\))?\s*$")
FIG_RE = re.compile(r"^!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]+)\)\s*"
                    r"(?:\{\s*w\s*=\s*(?P<width>[0-9.]+)\s*(?P<unit>cm|mm|in)?\s*\})?\s*$")
TABLE_CAP_RE = re.compile(r"^表\s*\d+")
FIG_CAP_RE = re.compile(r"^图\s*\d+")
# 以序号开头的条目（1 / 1、 / （1） / ① / 一、）：与实验目的、实验原理的条目一样顶格
LIST_ITEM_RE = re.compile(
    r"^\s*(?:\d+\s*[\.、)）]?|[（(]\s*\d+\s*[）)]|[①-⑩]"
    r"|[一二三四五六七八九十]+\s*[、.．)）])\s*\S")
# 把条目前的"阿拉伯数字序号"切出来加粗（1、2、3…）；「（1）（2）」这类括号序号不加粗
BOLD_MARKER_RE = re.compile(r"^(\s*\d{1,2}\s*(?:[、.．]\s*|\s+))(?=[^\d\s])")


def normalise_section_key(title):
    clean = re.sub(r"^[一二三四五六七八九十]+[、.．]\s*", "", title.strip()).replace(" ", "")
    if clean.startswith("原始数据"):
        return "rawdata"
    for key, aliases, _title in SECTION_ORDER:
        for alias in aliases:
            if clean.startswith(alias):
                return key
    return None


def parse_content(text):
    """Parse the Markdown content into ``{section_key: [Block, ...]}``."""
    sections, current = {}, None
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line or line.startswith("<!--"):
            continue
        if line.startswith("### "):
            if current is None:
                raise SystemExit("content.md: sub-heading before any section: %r" % line)
            sections[current].append(Block("h2", text=line[4:].strip()))
            continue
        if line.startswith("## "):
            if current is None:
                raise SystemExit("content.md: sub-heading before any section: %r" % line)
            sections[current].append(Block("h2", text=line[3:].strip()))
            continue
        if line.startswith("# "):
            key = normalise_section_key(line[2:])
            if key is None:
                raise SystemExit("content.md: unknown section heading %r" % line)
            current = key
            sections.setdefault(current, [])
            continue
        if current is None:
            raise SystemExit("content.md: text before the first section: %r" % line)
        eq = EQ_RE.match(line)
        if eq:
            sections[current].append(Block("eq", latex=eq.group("body").strip(),
                                           number=(eq.group("num") or "").strip()))
            continue
        fig = FIG_RE.match(line)
        if fig:
            width_cm = None
            if fig.group("width"):
                unit = fig.group("unit") or "cm"
                factor = {"cm": 1.0, "mm": 0.1, "in": 2.54}[unit]
                width_cm = float(fig.group("width")) * factor
            sections[current].append(Block("figure", path=fig.group("path").strip(),
                                           caption=fig.group("caption").strip(),
                                           width_cm=width_cm))
            continue
        if line.startswith("|"):
            rows = []
            while i - 1 < len(lines) and lines[i - 1].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i - 1].strip().strip("|").split("|")])
                i += 1
            i -= 1
            rows = [r for r in rows if not all(set(c) <= set("-: ") for c in r)]
            caption = ""
            if sections[current] and sections[current][-1].kind == "caption":
                caption = sections[current].pop().text
            sections[current].append(Block("table", rows=rows, caption=caption))
            continue
        if TABLE_CAP_RE.match(line):
            sections[current].append(Block("caption", text=line))
            continue
        sections[current].append(Block("p", text=line))
    return sections


# --------------------------------------------------------------------------
# block rendering (works for a table cell or the document body)
# --------------------------------------------------------------------------
def image_size(path):
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        return None


def resolve_path(path, base_dir):
    if os.path.isabs(path):
        return path
    return os.path.normpath(os.path.join(base_dir or ".", path))


def add_figure_paragraph(container, path, caption, max_w=MAX_IMAGE_CM,
                         max_h=MAX_IMAGE_H_CM, width_cm=None, paragraph=None):
    """插入一张居中插图，图题另起一段居中写在图的下方（批注12）。

    图题必须**单独成段**：如果和图挤在同一段里，遇到接近通栏宽的大图，
    LibreOffice 会把图题挤到图旁边的窄缝里竖着排；分成两段则一定在图下面。
    为防分页把图与图题拆开，图片段落带 keepNext（Word 里会与图题同页）。
    """
    p = paragraph if paragraph is not None else container.add_paragraph()
    set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                  space_before=6, space_after=0, keep_with_next=True)
    size = image_size(path)
    if width_cm:
        width = Cm(min(float(width_cm), max_w))
    else:
        width = Cm(max_w)
        if size and size[0]:
            natural_cm = size[0] / 96.0 * 2.54
            width = Cm(min(max_w, natural_cm))
    if size and size[0]:
        ratio = size[1] / float(size[0])
        if width.cm * ratio > max_h:
            width = Cm(max_h / ratio)
    p.add_run().add_picture(path, width=width)
    if caption and container is not None:
        cap = container.add_paragraph()
        set_paragraph(cap, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                      space_before=0, space_after=6)
        add_run(cap, caption, size=CAPTION_PT)
    return p


def add_picture_paragraph(container, path, max_w=MAX_IMAGE_CM, max_h=MAX_IMAGE_H_CM,
                          width_cm=None, caption=""):
    """Standalone centred picture; keeps a caption on the same paragraph."""
    return add_figure_paragraph(container, path, caption, max_w, max_h, width_cm)


def render_blocks(container, blocks, base_dir, section=""):
    """Render a block list into ``container`` (a _Cell or the Document body).

    ``section`` carries the section key so that a couple of sections can follow
    the sample report's own conventions (实验目的的条目在范文里是加粗、不缩进的).
    """
    for block in blocks:
        if block.kind == "h2":
            p = container.add_paragraph()
            set_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=BODY_LINE,
                          space_before=0, space_after=0, keep_with_next=True)
            add_run(p, block.text, size=BODY_PT, bold=True)
        elif block.kind == "p":
            p = container.add_paragraph()
            text = block.text
            if section == "questions":
                # 思考题不写"答："，直接写答案
                text = re.sub(r"^\s*答\s*[:：]\s*", "", text)
            emphasis = section == "purpose"
            is_item = bool(LIST_ITEM_RE.match(text))
            if section in ("conclusion", "questions"):
                # 结果陈述与思考题里的（1）（2）… 前面空两格
                first_line = 2
            elif emphasis or is_item:
                first_line = 0
            else:
                first_line = 2
            set_paragraph(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                          first_line_chars=first_line,
                          line_spacing=BODY_LINE)
            marker = BOLD_MARKER_RE.match(text) if is_item else None
            if marker:
                # 只有"1 2 3"这种阿拉伯数字序号加粗；（1）（2）不加粗
                add_run(p, marker.group(1), size=BODY_PT, bold=True)
                add_run(p, text[marker.end():], size=BODY_PT, bold=emphasis)
            else:
                add_run(p, text, size=BODY_PT, bold=emphasis)
        elif block.kind == "eq":
            p = container.add_paragraph()
            set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=BODY_LINE,
                          space_before=2, space_after=2)
            p._p.append(latex_to_omml(block.latex))
            if block.number:
                add_run(p, "\u3000\u3000\u3000\u3000(%s)" % block.number, size=BODY_PT)
        elif block.kind == "figure":
            path = resolve_path(block.path, base_dir)
            if not os.path.exists(path):
                print("WARN figure not found: %s" % path, file=sys.stderr)
                p = container.add_paragraph()
                set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0)
                add_run(p, "[缺少图片: %s]" % block.path, size=CAPTION_PT)
                if block.caption:
                    add_run(p, block.caption, size=CAPTION_PT)
            else:
                add_figure_paragraph(container, path, block.caption,
                                     width_cm=block.width_cm)
        elif block.kind == "caption":
            p = container.add_paragraph()
            set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                          space_before=6, space_after=0, keep_with_next=True)
            add_run(p, block.text, size=CAPTION_PT)
        elif block.kind == "table":
            render_table(container, block.caption, block.rows)


def render_table(container, caption, rows):
    if caption:
        cap = container.add_paragraph()
        set_paragraph(cap, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                      space_before=6, space_after=0, keep_with_next=True)
        add_run(cap, caption, size=CAPTION_PT)
    if not rows:
        return None
    ncols = max(len(r) for r in rows)
    table = container.add_table(rows=len(rows), cols=ncols)
    set_table_full_width(table)
    set_cell_borders(table)
    for r_i, row in enumerate(rows):
        for c_i in range(ncols):
            target = table.cell(r_i, c_i)
            cell_vertical_center(target)
            para = target.paragraphs[0]
            set_paragraph(para, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                          space_before=1, space_after=1)
            add_run(para, row[c_i] if c_i < len(row) else "",
                    size=TABLE_PT, bold=(r_i == 0))
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))
    return table


# --------------------------------------------------------------------------
# cover
# --------------------------------------------------------------------------
def display_width(text):
    return sum(2 if ord(ch) > 0x2000 else 1 for ch in text)


def underline_segments(paragraph):
    segments, current = [], []
    for run in paragraph.runs:
        if run.font.underline:
            current.append(run)
        elif current:
            segments.append(current)
            current = []
    if current:
        segments.append(current)
    return segments


def fill_segment(runs, value, min_width=None):
    """Centre a value inside a blank, preserving its width.

    A 全角冒号 that the template placed inside the blank (指导教师：…) stays
    hard against the label instead of floating into the middle.
    """
    value = str(value)
    original = "".join(r.text for r in runs)
    prefix = "：" if original.lstrip().startswith("：") else ""
    width = sum(display_width(r.text) for r in runs)
    if min_width:
        width = max(width, min_width)
    region = width - display_width(prefix)
    vw = display_width(value)
    if vw + 2 > region:
        region = vw + 2
    pad = region - vw
    left = pad // 2
    if value:
        runs[0].text = prefix + " " * left + value + " " * (pad - left)
    else:
        # a blank field must still show its underline: plain trailing spaces get
        # trimmed by the layout engine, so end the blank with one nbsp (which
        # some CJK fonts draw full width, hence the extra unit of slack)
        runs[0].text = prefix + " " * max(region - 2, 0) + "\u00a0"
    for run in runs[1:]:
        run.text = ""


def find_paragraph(doc, prefix):
    """Locate a cover line by label, tolerating a swallowed 全角冒号."""
    wanted = prefix.replace(" ", "").rstrip("：")
    for p in doc.paragraphs:
        text = p.text.replace(" ", "")
        if text.startswith(wanted):
            return p
    return None


def set_cover_line_spacing(paragraph, twips=1200):
    """封面每行固定行高（60 pt）。

    参考报告用的是模板的 3.75 倍行距；"auto" 在不同阅读器里算出的行高不一样
    （Word ≈ 60 pt，LibreOffice ≈ 86 pt 会把最后两行挤到第二页），所以统一成
    固定值，Word 里看到的疏密与参考报告一致，又不会溢出。
    """
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = ppr.makeelement(qn("w:spacing"), {})
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(twips))
    spacing.set(qn("w:lineRule"), "exact")


def fill_cover(doc, meta):
    specs = [
        ("课程名称：", [meta.get("course", "")]),
        ("实验名称：", [meta.get("experiment", "")]),
        ("学院：", [meta.get("college", "")]),
        ("指导教师：", [meta.get("teacher", "")]),
        ("报告人：", [meta.get("reporter", ""), meta.get("group", "")]),
        ("学号", [meta.get("student_id", ""), meta.get("location", "")]),
        ("实验时间：", list(meta.get("experiment_date_parts", ["", "", ""]))),
        ("提交时间：", list(meta.get("submit_date_parts", ["", "", ""]))),
    ]
    for label, values in specs:
        p = find_paragraph(doc, label)
        if p is None:
            print("WARN cover line not found: %s" % label, file=sys.stderr)
            continue
        segments = underline_segments(p)
        if len(segments) < len(values):
            print("WARN cover line %s has %d blanks for %d values"
                  % (label, len(segments), len(values)), file=sys.stderr)
        for seg, value in zip(segments, values):
            # keep the blank's own width so the underlines stay flush as designed
            fill_segment(seg, value)
        for seg in segments[len(values):]:
            fill_segment(seg, "")
        set_cover_line_spacing(p)


# --------------------------------------------------------------------------
# date helpers
# --------------------------------------------------------------------------
def date_parts(value, default=None):
    """Accept 2026-09-23 / 2026/9/23 / 26-9-23 and return ['26', '9', '23']."""
    if not value:
        return list(default or ["", "", ""])
    if isinstance(value, (list, tuple)):
        return [str(v) for v in value]
    text = str(value).strip()
    m = re.match(r"^(\d{2,4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if not m:
        return [text, "", ""]
    year, month, day = m.groups()
    return [year[-2:] if len(year) == 4 else year, str(int(month)), str(int(day))]


# --------------------------------------------------------------------------
# raw-data page and quiz page
# --------------------------------------------------------------------------
def set_page_break_before(paragraph):
    """Start this paragraph on a fresh page without adding a blank page."""
    ppr = paragraph._p.get_or_add_pPr()
    ppr.get_or_add_pageBreakBefore()
    return paragraph


def insert_page_break_before(doc, anchor_element):
    """Put a page-break paragraph immediately before ``anchor_element``."""
    br = doc.add_paragraph()
    set_paragraph(br, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=1.0)
    br.add_run().add_break(WD_BREAK.PAGE)
    br._p.getparent().remove(br._p)
    anchor_element.addprevious(br._p)
    return br


def fill_rawdata_page(doc, rows, images, base_dir, with_table=False, with_page=False):
    """最后一页（原始数据记录表）。

    默认**整页删掉**（老师要的原始数据记录表是另外打印/手写后附上的）；
    需要保留"原始数据记录表 + 组号 姓名"这一页时加 --rawdata-page，
    再想把课件的空白数据表印上去时加 --rawdata-table。
    手写照片（--rawdata-image）永远单独占一页，附在最后。
    """
    placeholder = heading = byline = None
    for p in doc.paragraphs:
        text = p.text.strip()
        if "表格自拟" in text:
            placeholder = p
        elif text == "原始数据记录表":
            heading = p
        elif text.replace(" ", "").startswith("组号"):
            byline = p
    if not (with_page or with_table):
        for target in (placeholder, heading, byline):
            if target is not None:
                target._p.getparent().remove(target._p)
    else:
        if placeholder is not None:
            placeholder._p.getparent().remove(placeholder._p)
        if heading is not None:
            set_page_break_before(heading)
        if rows and with_table:
            render_table(doc, "", rows)
    first = True
    for path in images:
        full = resolve_path(path, base_dir)
        if not os.path.exists(full):
            print("WARN raw-data image missing: %s" % full, file=sys.stderr)
            continue
        p = doc.add_paragraph()
        set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                      space_before=6, space_after=6)
        if first:
            set_page_break_before(p)
            first = False
        size = image_size(full)
        width = Cm(RAW_IMAGE_CM)
        if size and size[0]:
            ratio = size[1] / float(size[0])
            if RAW_IMAGE_CM * ratio > PAGE_IMAGE_H_CM:
                width = Cm(PAGE_IMAGE_H_CM / ratio)
        p.add_run().add_picture(full, width=width)


def insert_quiz_page(doc, body_table, images, base_dir=None):
    """封面之后单独留一页给预习试卷（选择题）截图。

    这一页上**只放图片段落**，而且段落标记只有 1 pt：页面几乎整页可用。用户自己
    往这一页粘图时，图片占满整页也不会因为残留空行被挤到下一页、多出一张白页
    （这是"粘图后凭空多一页"的根因）。<br>
    给了 --quiz 图片（可重复）就自动排上去；**多张一律纵向排在同一个第 2 页里**，
    总高超过一页时整体等比缩小——保证"封面 / 试卷 / 正文"永远只有三处分页。
    没给图就留一整页空白（页内不放任何提示文字）。
    """
    paths = [resolve_path(p, base_dir) for p in images]
    for path in paths:
        if not os.path.exists(path):
            print("WARN quiz image missing: %s" % path, file=sys.stderr)
    paths = [p for p in paths if os.path.exists(p)]

    # 先按整栏宽算出每张图的自然高度，再判断一整页装不装得下；装不下就整体等比缩
    widths = [RAW_IMAGE_CM] * len(paths)
    heights = []
    for path in paths:
        size = image_size(path)
        heights.append(RAW_IMAGE_CM * (size[1] / float(size[0]))
                       if size and size[0] else None)
    known = [h for h in heights if h is not None]
    budget = PAGE_IMAGE_H_CM - QUIZ_STACK_GAP_CM * max(0, len(paths) - 1)
    if known and sum(known) > budget:
        factor = budget / sum(known)
        widths = [w * factor for w in widths]
        heights = [None if h is None else h * factor for h in heights]

    slots = []
    if not paths:
        slots.append(doc.add_paragraph())        # 没图：留一整页空白
    else:
        for index, path in enumerate(paths):
            p = doc.add_paragraph()
            p.add_run().add_picture(path, width=Cm(widths[index]))
            slots.append(p)
    for p in slots:
        set_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, line_spacing=1.0,
                      space_before=0, space_after=0)
        set_paragraph_mark_size(p, 1)
    set_page_break_before(slots[0])              # 封面之后从新的一页开始（不用单独的分页段）

    # 把这一页的所有段落按原顺序整体挪到正文表格之前
    anchor = body_table._tbl
    parent = anchor.getparent()
    index = list(parent).index(anchor)
    for offset, p in enumerate(slots):
        p._p.getparent().remove(p._p)
        parent.insert(index + offset, p._p)


def start_body_on_new_page(doc, body_table):
    """正文七节从新的一页开始。

    **不用"分页空段落"**：那种做法在页面被占满时会反噬——第 2 页粘满预习试卷
    截图后，那个分页段落会被挤到下一页，于是正文前凭空多出一张空白页（用户
    2026-10-13 反馈的问题）。改成给正文表格首个单元格的第一段设
    `w:pageBreakBefore`，Word/WPS 会在进入表格处直接分页，不留任何多余段落。
    """
    try:
        first = body_table.rows[0].cells[0].paragraphs[0]
    except (IndexError, AttributeError):
        return
    set_page_break_before(first)


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def load_meta(args):
    profile = {}
    if args.profile and os.path.exists(args.profile):
        with open(args.profile, encoding="utf-8") as fh:
            profile = json.load(fh)
    profile.pop("formatting", None)
    meta = dict(profile)
    if args.meta and os.path.exists(args.meta):
        with open(args.meta, encoding="utf-8") as fh:
            meta.update(json.load(fh))
    for key, value in {
        "experiment": args.experiment, "location": args.location,
        "course": args.course, "teacher": args.teacher, "college": args.college,
        "group": args.group, "reporter": args.reporter,
        "student_id": args.student_id,
        "experiment_date": args.date, "submit_date": args.submit_date,
    }.items():
        if value:
            meta[key] = value
    meta["experiment_date_parts"] = date_parts(
        meta.get("experiment_date"), meta.get("experiment_date_parts"))
    meta["submit_date_parts"] = date_parts(
        meta.get("submit_date"), meta.get("submit_date_parts"))
    return meta


def _render_cell(cell, title, blocks, base_dir, kind, key=""):
    for paragraph in cell.paragraphs[1:]:
        paragraph._p.getparent().remove(paragraph._p)
    heading = cell.paragraphs[0]
    for run in list(heading.runs):
        run._element.getparent().remove(run._element)
    ppr = heading._p.get_or_add_pPr()
    numpr = ppr.find(qn("w:numPr"))
    if numpr is not None:
        ppr.remove(numpr)
    # 不加 keepNext：老师的模板里各小节是连续排版的，标题若"与下段同页"
    # 会让整行被顶到下一页，出现"实验目的独占一页"的难看的空白
    set_paragraph(heading, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=BODY_LINE,
                  space_before=0, space_after=2, keep_with_next=False)
    add_run(heading, title, size=HEADING_PT, bold=True)
    if blocks:
        render_blocks(cell, blocks, base_dir, section=key)
        # 每个板块结束后空一行：正文七节之间各留一个空行（用户 2026-10-13 要求）
        gap = cell.add_paragraph()
        set_paragraph(gap, align=WD_ALIGN_PARAGRAPH.LEFT, line_spacing=BODY_LINE)
        add_run(gap, "", size=BODY_PT)
    else:
        p = cell.add_paragraph()
        set_paragraph(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_chars=2,
                      line_spacing=BODY_LINE)
        add_run(p, "", size=BODY_PT)


def main(argv=None):
    here = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(description="Build a 深圳大学实验报告 DOCX")
    parser.add_argument("--content", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--base", default=os.path.join(here, "..", "assets", "base.docx"))
    parser.add_argument("--profile", default=os.path.join(here, "..", "assets", "profile.json"))
    parser.add_argument("--meta")
    parser.add_argument("--kind", choices=("preview", "full"), default="preview")
    parser.add_argument("--quiz", action="append", default=[])
    parser.add_argument("--rawdata-image", action="append", default=[])
    parser.add_argument("--rawdata-page", action="store_true",
                        help="保留最后一页「原始数据记录表 + 组号 姓名」（默认整页删掉）")
    parser.add_argument("--rawdata-table", action="store_true",
                        help="把 content.md 里的空白数据表也排到原始数据记录表页（隐含 --rawdata-page）")
    parser.add_argument("--experiment")
    parser.add_argument("--location")
    parser.add_argument("--course")
    parser.add_argument("--teacher")
    parser.add_argument("--college")
    parser.add_argument("--group")
    parser.add_argument("--reporter")
    parser.add_argument("--student-id")
    parser.add_argument("--date")
    parser.add_argument("--submit-date")
    args = parser.parse_args(argv)

    meta = load_meta(args)
    missing = [k for k in ("experiment", "location", "experiment_date", "submit_date")
               if not meta.get(k)]
    if missing:
        print("WARN meta fields missing: %s" % ", ".join(missing), file=sys.stderr)

    content_path = os.path.abspath(args.content)
    base_dir = os.path.dirname(content_path)
    with open(content_path, encoding="utf-8") as fh:
        sections = parse_content(fh.read())

    doc = Document(args.base)
    if len(doc.tables) < 2 or len(doc.tables[1].rows) < 7:
        raise SystemExit("base.docx does not look like the report template")

    fill_cover(doc, meta)

    body_table = doc.tables[1]
    for index, (key, _aliases, title) in enumerate(SECTION_ORDER):
        _render_cell(body_table.rows[index].cells[0], title,
                     sections.get(key, []), base_dir, args.kind, key=key)

    raw_rows = [b.rows for b in sections.get("rawdata", []) if b.kind == "table"]
    raw_images = list(args.rawdata_image) + list(meta.get("rawdata_images", []))
    fill_rawdata_page(doc, raw_rows[0] if raw_rows else [], raw_images, base_dir,
                      with_table=args.rawdata_table,
                      with_page=args.rawdata_page or args.rawdata_table)

    quiz = list(args.quiz) + list(meta.get("quiz_images", []))
    insert_quiz_page(doc, body_table, quiz, base_dir)
    start_body_on_new_page(doc, body_table)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    doc.save(args.out)
    print("WROTE %s" % os.path.abspath(args.out))
    missing_sections = [title for key, _a, title in SECTION_ORDER
                        if key not in sections]
    if missing_sections:
        print("EMPTY SECTIONS: %s" % ", ".join(missing_sections))
    return 0


if __name__ == "__main__":
    sys.exit(main())
