#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""check_report.py -- 深圳大学实验报告 (.docx) 格式与完整性检查。

用法:
    python check_report.py <report.docx> [--json <out.json>] [--kind preview|full]

按严重程度分组打印: ERROR(会扣分/结构性错误) / WARN(很可能违反教师批注) / INFO(观察)
末行固定为  SUMMARY errors=N warnings=N
退出码: 无 ERROR -> 0 ; 有 ERROR -> 1

文档模型(深圳大学实验报告模板): 封面(2x3 得分表 + 标题 + 带下划线的字段段) ->
正文 9 行 x 1 列表格(实验目的/实验原理/实验仪器/四、实验内容与步骤/五、数据处理/
六、结果陈述/七、思考题/指导教师批阅意见/成绩评定) -> 页脚注 -> 原始数据记录表独立页。
"""

import argparse
import datetime
import io
import json
import os
import re
import sys
import traceback

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
O_MATH, O_MATHPARA = '{%s}oMath' % M_NS, '{%s}oMathPara' % M_NS
M_T = '{%s}t' % M_NS
W_T, W_P, W_TBL = qn('w:t'), qn('w:p'), qn('w:tbl')
W_DRW, W_PICT = qn('w:drawing'), qn('w:pict')
W_BR, W_SECTPR = qn('w:br'), qn('w:sectPr')

CN_NUM = '一二三四五六七八九十'
SECTION_NAMES = ['实验目的', '实验原理', '实验仪器', '实验内容与步骤',
                 '数据处理', '结果陈述', '思考题']
TEACHER_ROWS = ['指导教师批阅意见', '成绩评定']
# 模板自带的成绩评定评分表格(表头), 属模板内容而非学生填写
RUBRIC_HEADERS = ('预习（20分）', '操作及记录（40分）', '数据处理与结果陈述（30分）',
                  '思考题（10分）', '报告整体印象', '总分')
RUBRIC_ALT = ('操作及记录', '思考题')      # 模板中分数可能占另一行的两栏
RUBRIC_NORM = [re.sub(r'\s+', '', h) for h in RUBRIC_HEADERS]
RUBRIC_ALT_NORM = [re.sub(r'\s+', '', h) for h in RUBRIC_ALT]
CJK_FONT_HINTS = ('宋体', 'SimSun', 'NSimSun')
# 模板自带提示语, 不计入"内容非空"
PLACEHOLDER = re.compile(r'注[:：]|需从原始数据记录表|表格自拟|请填写|请勿|删除本')

COVER_FIELDS = ['课程名称', '实验名称', '学院', '指导教师', '报告人', '学号',
                '实验地点', '实验时间', '提交时间']
COVER_REQUIRED = ('实验名称', '实验地点', '实验时间', '提交时间')
COVER_LABEL_RE = re.compile(r'课\s*程\s*名\s*称|实\s*验\s*名\s*称|学\s*院|指\s*导\s*教\s*师|'
                            r'报\s*告\s*人|学\s*号|实\s*验\s*地\s*点|实\s*验\s*时\s*间|'
                            r'提\s*交\s*时\s*间|组\s*号')
COVER_NOISE = re.compile(r'^[\s:：.·_\-—－~～/／\\|]+$')

CAP_FIG = re.compile(r'^\s*图\s*(\d+)')
CAP_TBL = re.compile(r'^\s*表\s*(\d+)')
REF_FIG = re.compile(r'图\s*(\d+)')
REF_TBL = re.compile(r'表\s*(\d+)')
EQ_NUM = re.compile(r'[（(]\s*(\d+)\s*[)）]\s*$')
EQ_REF = re.compile(r'式\s*[（(]\s*(\d+)\s*[)）]')
DATE_RE = re.compile(r'(\d{2,4})\s*年\s*(\d{1,2})?\s*月\s*(\d{1,2})?\s*日')
Q_NUM = re.compile(r'^\s*(\d{1,2})\s*[\.、,，)）:：]?\s*(.+)$')
ANS = re.compile(r'^\s*(答|答案|解)\s*[:：]')
RAW_MARK = re.compile(r'^\s*原\s*始\s*数\s*据\s*记\s*录\s*表\s*[（(]?[^。；;]{0,12}[）)]?\s*$')
RAW_TEXT = re.compile(r'原\s*始\s*数\s*据\s*记\s*录\s*表')


# ---------------------------------------------------------------- findings
class Report(object):
    ORDER = {'ERROR': 0, 'WARN': 1, 'INFO': 2}

    def __init__(self):
        self.items = []

    def add(self, sev, rule, msg, where=''):
        for it in self.items:          # 相同规则+相同消息只留一条, 避免刷屏
            if it['rule'] == rule and it['message'] == msg:
                return
        self.items.append({'severity': sev, 'rule': rule, 'message': msg, 'where': where})

    def error(self, r, m, w=''):
        self.add('ERROR', r, m, w)

    def warn(self, r, m, w=''):
        self.add('WARN', r, m, w)

    def info(self, r, m, w=''):
        self.add('INFO', r, m, w)

    def counts(self):
        return (len([i for i in self.items if i['severity'] == 'ERROR']),
                len([i for i in self.items if i['severity'] == 'WARN']))

    def sorted_items(self):
        return sorted(self.items, key=lambda i: (self.ORDER.get(i['severity'], 9), i['rule']))


def guard(report, rule, where=''):
    """单条规则抛异常时降级为 INFO, 绝不影响其它规则。"""
    def deco(fn):
        def wrapper(*a, **kw):
            try:
                return fn(*a, **kw)
            except Exception as exc:  # noqa: BLE001 - 故意兜底
                report.info(rule + '.crash', '该规则检查时异常, 结果不可信: %s: %s'
                            % (type(exc).__name__, exc), where)
        return wrapper
    return deco


# ---------------------------------------------------------------- OOXML 工具
def p_text(p_el, with_math=True):
    """段落文字; with_math 时把 OMML 的 m:t 也拼进来(公式变量常不在 w:t 里)。"""
    parts = []
    for node in p_el.iter():
        if node.tag == W_T or (with_math and node.tag == M_T):
            parts.append(node.text or '')
    return re.sub(r'[ \t\u00a0\u3000]+', ' ', ''.join(parts)).strip()


def is_img(p_el):
    return bool(p_el.findall('.//' + W_DRW) or p_el.findall('.//' + W_PICT))


def n_img(p_el):
    return len(p_el.findall('.//' + W_DRW)) + len(p_el.findall('.//' + W_PICT))


def is_math(p_el):
    return bool(p_el.findall('.//' + O_MATH) or p_el.findall('.//' + O_MATHPARA))


def run_texts(p_el):
    """按文档顺序返回段落内所有 run 的文字, 图片所在的 run 记为 None。

    用来区分"图片之前"和"图片之后"的文字: 生成器会把图片与题注放进同一段,
    中间用 w:br 换行(避免 Word 把题注挤到下一页)。
    """
    out = []
    for child in p_el.iterchildren():
        if child.tag != qn('w:r'):
            continue          # 跳过 bookmarks / 批注锚点等, 但仍会递归查找图片
        has_img = bool(child.findall('.//' + W_DRW) or child.findall('.//' + W_PICT))
        txt = ''.join((t.text or '') for t in child.iter(W_T))
        if has_img and not txt.strip():
            out.append(None)
        elif txt:
            out.append(txt)
    return out


def text_after_image(p_el):
    """图片(或其后换行)之后的文字 = 同段题注。没有图片或图片后无文字则返回 ''。"""
    if not is_img(p_el):
        return ''
    seq = run_texts(p_el)
    idx = None
    for i, x in enumerate(seq):
        if x is None:
            idx = i
    if idx is None:
        return ''
    return re.sub(r'[ \t\u00a0\u3000]+', ' ', ''.join(seq[idx + 1:])).strip()


def own_caption(p_el):
    """段落"自带"的题注文字(图题注与图片同段时)。"""
    return text_after_image(p_el)


def resolve_fig_caption(paras, pi):
    """返回第 pi 段的图题注文字: 优先同段(图片后的文字), 否则下一段。

    返回 (题注文字, 布局) ; 布局为 'same' / 'next' / None。
    """
    same = own_caption(paras[pi]._p)
    if CAP_FIG.match(same):
        return same, 'same'
    if pi + 1 < len(paras):
        nxt = p_text(paras[pi + 1]._p)
        if CAP_FIG.match(nxt):
            return nxt, 'next'
    return '', None


def pv(txt, n=40):
    txt = re.sub(r'\s+', ' ', txt or '').strip()
    return (txt[:n] + '…') if len(txt) > n else (txt or '(空段落)')


def run_size(run):
    rPr = run._r.find(qn('w:rPr'))
    sz = rPr.find(qn('w:sz')) if rPr is not None else None
    try:
        return int(sz.get(qn('w:val'))) / 2.0 if sz is not None else None
    except (TypeError, ValueError):
        return None


def run_font(run):
    rPr = run._r.find(qn('w:rPr'))
    rf = rPr.find(qn('w:rFonts')) if rPr is not None else None
    return None if rf is None else (rf.get(qn('w:eastAsia')) or rf.get(qn('w:ascii'))
                                    or rf.get(qn('w:hAnsi')))


def img_anchors(p_el):
    """每张浮动图片的 (top, height), 单位 twips; 用于检测图片压住题注。"""
    out = []
    for dr in p_el.findall('.//' + W_DRW):
        anchor = dr.find(qn('wp:anchor'))
        if anchor is None:
            continue
        ext = dr.find('.//' + qn('wp:extent'))
        try:
            cy = float(ext.get('cy')) / 635.0 if ext is not None else None
        except (TypeError, ValueError):
            cy = None
        top, pv_el = None, anchor.find(qn('wp:positionV'))
        off = pv_el.find(qn('wp:posOffset')) if pv_el is not None else None
        if off is not None and off.text:
            try:
                top = float(off.text) / 635.0
            except ValueError:
                top = None
        out.append((top, cy))
    return out


def frame_y(p_el):
    pPr = p_el.find(qn('w:pPr'))
    fr = pPr.find(qn('w:framePr')) if pPr is not None else None
    if fr is None:
        return None
    try:
        return float(fr.get(qn('w:y')))
    except (TypeError, ValueError):
        return None


def cell_blocks(tc):
    return [(('p' if c.tag == W_P else 'tbl'),
             Paragraph(c, None) if c.tag == W_P else Table(c, None))
            for c in tc.iterchildren() if c.tag in (W_P, W_TBL)]


def parse_date(text):
    m = DATE_RE.search(re.sub(r'\s+', ' ', text or ''))
    if not m or not m.group(2) or not m.group(3):
        return None
    y = int(m.group(1))
    try:
        return datetime.date(y + 2000 if y < 100 else y, int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


# ---------------------------------------------------------------- 文档模型
class DocModel(object):
    def __init__(self, path):
        self.path = path
        self.doc = Document(path)
        self.body = self.doc.element.body
        self.main = self._find_main_table()
        self.main_el = self.main._tbl if self.main is not None else None
        self.cover_paras, self.cover_table = self._collect_cover()
        self.sections = self._collect_sections()
        self.body_p_els = list(self.main_el.iter(W_P)) if self.main_el is not None else []

    # -- 顺序工具: 不能缓存 lxml 元素的 id() (代理对象会被回收) ----------
    def tail_els(self):
        out, seen = [], self.main_el is None
        for child in self.body.iterchildren():
            if child is self.main_el:
                seen = True
            elif seen:
                out.append(child)
        return out

    def head_els(self):
        out = []
        for child in self.body.iterchildren():
            if child is self.main_el:
                break
            out.append(child)
        return out

    def doc_paras(self):
        return [Paragraph(el, None) for el in self.body.iter(W_P)]

    def _find_main_table(self):
        tbls = [Table(el, None) for el in self.body.iter(W_TBL)]
        for want_cols in (True, False):
            for t in tbls:
                try:
                    if len(t.rows) == 9 and (not want_cols or len(t.columns) == 1):
                        return t
                except Exception:  # noqa: BLE001
                    continue
        return None

    def _collect_cover(self):
        paras, tbl = [], None
        for el in self.head_els():
            if el.tag == W_P:
                paras.append(Paragraph(el, None))
            elif el.tag == W_TBL and tbl is None:
                try:
                    t = Table(el, None)
                    if len(t.rows) >= 2 and len(t.columns) >= 2:
                        tbl = t
                except Exception:  # noqa: BLE001
                    pass
        return paras, tbl

    def _collect_sections(self):
        secs = []
        if self.main is None:
            return secs
        for ri, row in enumerate(self.main.rows):
            blocks = cell_blocks(row.cells[0]._tc)
            secs.append({'row': ri, 'tc': row.cells[0]._tc, 'blocks': blocks,
                         'paras': [b for k, b in blocks if k == 'p'],
                         'tables': [b for k, b in blocks if k == 'tbl']})
        return secs

    def heading_of(self, si):
        sec = self.sections[si] if si < len(self.sections) else None
        return re.sub(r'\s+', '', p_text(sec['paras'][0]._p)) if sec and sec['paras'] else ''

    def content_of(self, si):
        sec = self.sections[si] if si < len(self.sections) else None
        return sec['paras'][1:] if sec else []

    def content_items(self, si):
        """章节的实质内容 [(段落, 文字)]; 空段/模板提示语被剔除。"""
        out = []
        for p in self.content_of(si):
            t = p_text(p._p)
            if not t:
                if is_img(p._p) or is_math(p._p):
                    out.append((p, '（图片/公式）'))
            elif not PLACEHOLDER.search(t):
                out.append((p, t))
        return out

    def section_is_empty(self, si):
        return not self.content_items(si)


def cover_values(model):
    """{标签(去空格): 值文本} -- 值可能跨段(如 实验时间/提交时间 分两行)。"""
    items = [(i, p_text(p._p)) for i, p in enumerate(model.cover_paras)]
    bounds = sorted([(i, m.start(), m.group(0)) for i, t in items
                     for m in COVER_LABEL_RE.finditer(t)], key=lambda x: (x[0], x[1]))
    values = {}
    for bi, (pi, off, lab) in enumerate(bounds):
        nxt = bounds[bi + 1] if bi + 1 < len(bounds) else None
        chunks = [items[pi][1][off + len(lab):(nxt[1] if nxt and nxt[0] == pi else None)]]
        if nxt is None:
            chunks += [items[j][1] for j in range(pi + 1, len(items))]
        elif nxt[0] != pi:
            chunks += [items[j][1] for j in range(pi + 1, nxt[0])]
        values.setdefault(re.sub(r'\s+', '', lab),
                          ' '.join(c for c in chunks if c).strip())
    return values


def get_field(values, name):
    key = re.sub(r'\s+', '', name)
    for k, v in values.items():
        if key in k:
            return v
    return None


def is_rubric_cell(text):
    """单元格是否"恰好"是模板评分表头的一部分。

    只按精确取值匹配(忽略空白/全角半角括号差异), 不按包含关系 -- 否则
    「预习（20分）已完成」这类学生补字会被误判为模板内容。多输出了"分数行"
    (如只写「40分」)也算模板的一部分。
    """
    norm = re.sub(r'[\s\u00a0\u3000]+', '', text or '').replace('(', '（').replace(')', '）')
    if not norm:
        return False
    if norm in RUBRIC_NORM or norm in RUBRIC_ALT_NORM:
        return True
    m = re.fullmatch(r'（(\d+)分）', norm)
    return bool(m and m.group(1) in ('20', '40', '30', '10'))


def rubric_violations(model, si):
    """返回成绩评定行里"不属于模板"的文字。

    模板内容 = 「成绩评定」标题 + 评分表格的表头单元格。表头本身不告警;
    但学生若在表头单元格里补字(如「预习（20分）已完成」)、在评分行填入分数,
    或在表格外另写文字, 都算违规内容。
    """
    bad = []
    for k, blk in model.sections[si]['blocks']:
        if k == 'p':
            t = p_text(blk._p)
            if t and t != '成绩评定':
                bad.append(t)
            continue
        for ri, row in enumerate(blk.rows):
            for ci, cell in enumerate(row.cells):
                t = (cell.text or '').strip()
                if not t:
                    continue
                if ri == 0 and is_rubric_cell(t):
                    continue          # 模板评分表头
                bad.append('r%dc%d %s' % (ri + 1, ci + 1, t))
    return bad


def collect_figures(model):
    """{图号: (行号, 题注段落, 是否与图片同段)}

    题注有两种合法布局: 与图片同段(图片后 w:br 换行再写题注), 或紧随其后的独立段落。
    """
    figs = {}
    for si in range(len(model.sections)):
        paras = model.sections[si]['paras']
        for pi, p in enumerate(paras):
            same = own_caption(p._p)
            if CAP_FIG.match(same):                      # 同段题注
                figs.setdefault(CAP_FIG.match(same).group(1), (si, p, True))
                continue
            if is_img(p._p):                             # 独立段落题注
                nxt = p_text(paras[pi + 1]._p) if pi + 1 < len(paras) else ''
                m = CAP_FIG.match(nxt)
                if m:
                    figs.setdefault(m.group(1), (si, paras[pi + 1], False))
            elif CAP_FIG.match(p_text(p._p)):            # 追认: 图片段落无题注但本段有
                figs.setdefault(CAP_FIG.match(p_text(p._p)).group(1), (si, p, False))
    return figs


def collect_tbl_captions(model):
    """{表号: (行号, 段落, 是否紧接表格)}"""
    caps = {}
    for si in range(len(model.sections)):
        blocks = model.sections[si]['blocks']
        for bi, (k, blk) in enumerate(blocks):
            if k != 'p':
                continue
            m = CAP_TBL.match(p_text(blk._p))
            if m:
                nxt_tbl = bi + 1 < len(blocks) and blocks[bi + 1][0] == 'tbl'
                caps.setdefault(m.group(1), (si, blk, nxt_tbl))
    return caps


# ---------------------------------------------------------------- 规则 1-2 封面
def rule_cover(model, report, values):
    def blank(val):
        return val is None or not val.strip() or bool(COVER_NOISE.match(val))

    if not values:
        report.error('cover.parse', '封面未识别到任何字段标签, 封面结构可能已被改动', '封面')
        return
    for disp in COVER_FIELDS:
        val = get_field(values, disp)
        if val is None:
            report.error('cover.missing-label', '封面缺少字段标签「%s」' % disp, '封面')
        elif blank(val):
            msg = '封面「%s」未填写(空白)' % disp
            (report.error if disp in COVER_REQUIRED else report.warn)('cover.blank', msg, '封面')
        else:
            report.info('cover.value', '封面「%s」= %s' % (disp, pv(val, 44)), '封面')
    rp = get_field(values, '报告人') or ''
    m = re.search(r'组\s*号\s*[:：]?\s*(\S*)', rp)
    if m is not None and not m.group(1).strip():
        report.error('cover.blank', '封面「组号」未填写(空白)', '封面')
    if not re.search(r'\d{6,}', get_field(values, '学号') or ''):
        report.error('cover.blank', '封面「学号」未填写或不是有效学号(需 6 位以上数字)', '封面')
    if model.cover_table is None:
        report.warn('cover.table', '未找到封面 2x3「得分/教师签名/批改日期」表格', '封面')
    else:
        hdr = [c.text.strip() for c in model.cover_table.rows[0].cells]
        for need in ('得分', '教师签名', '批改日期'):
            if not any(need in h for h in hdr):
                report.warn('cover.table', '封面表格缺少表头「%s」(实有: %s)'
                            % (need, '/'.join(hdr)), '封面')


def rule_dates(model, report, values):
    where = '封面·日期'
    exp, sub = parse_date(get_field(values, '实验时间')), parse_date(get_field(values, '提交时间'))
    if exp is None:
        report.warn('date.parse', '实验时间无法解析为日期: %s'
                    % pv(get_field(values, '实验时间') or '', 30), where)
    if sub is None:
        report.warn('date.parse', '提交时间无法解析为日期(年/月/日可能未填全): %s'
                    % pv(get_field(values, '提交时间') or '', 30), where)
    if exp is None or sub is None:
        return
    delta = (sub - exp).days
    if delta < 0:
        report.warn('date.order', '提交时间(%s)早于实验时间(%s), 相差 %d 天'
                    % (sub, exp, -delta), where)
    elif delta > 7:
        report.warn('date.window', '提交时间(%s)距实验时间(%s) %d 天, 超过 7 天上限'
                    % (sub, exp, delta), where)
    else:
        report.info('date.order', '提交时间(%s)距实验时间(%s) %d 天, 符合 ≤7 天要求'
                    % (sub, exp, delta), where)


# ---------------------------------------------------------------- 规则 3-4 章节
def rule_sections(model, report):
    if model.main is None:
        report.error('section.table', '未找到正文 9 行 x 1 列章节表格, 结构与模板不符')
        return
    if len(model.sections) < 9:
        report.warn('section.table', '正文表格只有 %d 行(模板为 9 行)' % len(model.sections))
    found, missing = [], []
    for i, name in enumerate(SECTION_NAMES):
        head, hit = '', None
        for si in range(min(7, len(model.sections))):
            h = model.heading_of(si)
            if name in h:
                head, hit = h, si
                break
        if hit is None:
            missing.append('%s、%s' % (CN_NUM[i], name))
        else:
            found.append((i, hit, head, head.startswith(CN_NUM[i] + '、')))
    if missing:
        report.error('section.missing', '缺少章节: %s' % '、'.join(missing), '正文表格')
    rows = [f[1] for f in found]
    if rows != sorted(rows):
        report.warn('section.order', '章节出现顺序与模板不一致: %s'
                    % ' → '.join(f[2] for f in found), '正文表格')
    unnum = [SECTION_NAMES[f[0]] for f in found if not f[3]]
    if unnum and len(unnum) < len(found):
        report.warn('section.numbering',
                    '章节编号不一致: 「%s」无「一、二、…」序号, 其余章节省略(模板前三节自带无编号, '
                    '建议统一补上)' % '、'.join(unnum), '正文表格')
    elif unnum:
        report.warn('section.numbering', '全部章节均无「一、…七、」序号: %s' % '、'.join(unnum),
                    '正文表格')
    elif found:
        report.info('section.numbering', '章节编号统一(一、~七、)')


def rule_emptiness(model, report, kind):
    if model.main is None:
        return
    for i in range(min(7, len(model.sections))):
        where = '正文表格 第%d行' % (i + 1)
        items = model.content_items(i)
        label = '%s、%s' % (CN_NUM[i], SECTION_NAMES[i])
        if items:
            first = next((t for _p, t in items if t and t != '（图片/公式）'), '')
            report.info('section.content', '%s 有 %d 段实质内容%s'
                        % (label, len(items), ('，首段: ' + pv(first, 30)) if first else ''), where)
        elif i < 4:
            report.error('section.empty', '%s 为空(preview/full 均不应为空)' % label, where)
        elif kind == 'full':
            report.error('section.empty', '%s 为空(full 完整报告要求填写)' % label, where)
        else:
            report.info('section.empty', '%s 为空(--kind preview 下允许, 交报告前需补写)' % label, where)
    for i, name in enumerate(TEACHER_ROWS):
        ri = 7 + i
        if ri >= len(model.sections):
            continue
        if i == 1:                      # 成绩评定: 模板自带评分表格, 需排除表头
            bad = rubric_violations(model, ri)
            if bad:
                report.warn('section.teacher',
                            '第%d行「%s」的评分表格属模板内容, 但该行还出现了其它文字(学生不应填写): %s'
                            % (ri + 1, name, '; '.join(pv(b, 26) for b in bad[:3])),
                            '正文表格 第%d行' % (ri + 1))
        elif model.content_items(ri):
            report.warn('section.teacher', '第%d行「%s」应为教师填写, 学生不应写入内容'
                        % (ri + 1, name), '正文表格 第%d行' % (ri + 1))


# ---------------------------------------------------------------- 规则 5 图表
def rule_figures(model, report):
    n_paras, n_same, n_next, no_cap, no_period = 0, 0, 0, [], []
    for si in range(min(7, len(model.sections))):
        paras = model.sections[si]['paras']
        for pi, p in enumerate(paras):
            if not is_img(p._p):
                continue
            n_paras += 1
            cap, layout = resolve_fig_caption(paras, pi)
            if layout is None:
                no_cap.append('第%d行 第%d段' % (si + 1, pi + 1))
                continue
            n_same += 1 if layout == 'same' else 0
            n_next += 1 if layout == 'next' else 0
            if not cap.rstrip().endswith('。'):
                no_period.append('图%s(%s)' % (CAP_FIG.match(cap).group(1), pv(cap, 26)))
    if n_paras == 0:
        report.warn('figure.none', '正文(第1~7行)未检测到任何图片')
    else:
        report.info('figure.count', '正文共 %d 个含图段落(第1~7行); 题注布局: 同段 %d 个, 独立成段 %d 个'
                    % (n_paras, n_same, n_next))
    if no_cap:
        report.error('figure.caption-missing',
                     '%d 张图片的同段/下一段题注都不是以「图N」开头——图题里必须出现「图」字, '
                     '写法为「图N 说明文字。」: %s'
                     % (len(no_cap), '; '.join(no_cap[:6])))
    if no_period:
        report.warn('figure.caption-period', '%d 条图题注未以句号「。」结尾: %s'
                    % (len(no_period), '; '.join(no_period[:6])))
    # 浮动图片与其题注同位置 -> 题注被图片盖住
    for si in range(min(7, len(model.sections))):
        paras = model.sections[si]['paras']
        for pi, p in enumerate(paras):
            anchors = img_anchors(p._p)
            if not anchors or pi + 1 >= len(paras):
                continue
            nxt = paras[pi + 1]
            if not CAP_FIG.match(p_text(nxt._p)):
                continue
            y = frame_y(nxt._p)
            if y is None:
                continue
            for top, cy in anchors:
                if top is not None and cy and abs(top - y) < min(cy, 400) * 0.5:
                    report.warn('figure.caption-overlap',
                                '图题注「%s」与图片重叠(图 top=%.0f / 题注 y=%.0f twips), 题注可能被遮挡'
                                % (pv(p_text(nxt._p), 20), top, y), '正文表格 第%d行' % (si + 1))
    # 表题注须在表格之前
    caps = collect_tbl_captions(model)
    n_tbls = 0
    for si in range(min(7, len(model.sections))):
        blocks = model.sections[si]['blocks']
        for bi, (k, _blk) in enumerate(blocks):
            if k != 'tbl':
                continue
            n_tbls += 1
            prev = p_text(blocks[bi - 1][1]._p) if bi > 0 and blocks[bi - 1][0] == 'p' else ''
            if not CAP_TBL.match(prev):
                report.error('table.caption-position',
                             '正文第%d行中的表格前面没有以「表N」开头的题注——表题必须写在表格'
                             '上方且含「表」字, 前一段: %s' % (si + 1, pv(prev, 26)))
    if n_tbls and not caps:
        report.warn('table.caption-missing', '正文有 %d 个表格但没有任何「表N」题注' % n_tbls)


# ---------------------------------------------------------------- 规则 6 引用
def rule_references(model, report):
    figs, tbls = collect_figures(model), collect_tbl_captions(model)
    fig_nums, tbl_nums = set(figs), set(tbls)
    cap_els = set([id(p._p) for _n, (_s, p, _same) in figs.items()] +
                  [id(p._p) for _n, (_s, p, _x) in tbls.items()])
    ref_fig, ref_tbl = set(), set()
    for el in model.body_p_els:
        if id(el) in cap_els:
            continue
        t = p_text(el)
        if t:
            ref_fig |= set(REF_FIG.findall(t))
            ref_tbl |= set(REF_TBL.findall(t))
    key = lambda s: int(s)

    def list_of(prefix, nums):
        nums = sorted(nums, key=key)
        return '、'.join(prefix + n for n in nums) if nums else '无'

    unref = sorted(fig_nums - ref_fig, key=key)
    if unref:
        report.warn('ref.figure-unreferenced', '图题注 %s 在正文中从未被引用(应写「如图N所示」等)'
                    % '、'.join('图' + n for n in unref))
    undef = sorted(ref_fig - fig_nums, key=key)
    if undef:
        report.error('ref.figure-undefined', '正文引用了不存在的图 %s (全文图题注: %s)'
                     % ('、'.join('图' + n for n in undef), list_of('图', fig_nums)))
    if tbl_nums:
        unref_t = sorted(tbl_nums - ref_tbl, key=key)
        if unref_t:
            report.warn('ref.table-unreferenced', '表题注 %s 在正文中从未被引用'
                        % '、'.join('表' + n for n in unref_t))
    else:
        report.info('ref.table', '正文中没有「表N」题注')
    undef_t = sorted(ref_tbl - tbl_nums, key=key)
    if undef_t:
        report.error('ref.table-undefined', '正文引用了不存在的表 %s (全文表题注: %s)'
                     % ('、'.join('表' + n for n in undef_t), list_of('表', tbl_nums)))
    if fig_nums:
        report.info('ref.figure', '图题注: %s ; 已引用: %s'
                    % (list_of('图', fig_nums), list_of('图', ref_fig)))


def nearest_prev_text(paras, pi):
    """从第 pi 段往前找最近的"正文"段落(跳过空段、图题/表题、图片段)。"""
    for j in range(pi - 1, -1, -1):
        t = p_text(paras[j]._p)
        if is_img(paras[j]._p) or not t:
            continue
        if CAP_FIG.match(t) or CAP_TBL.match(t):
            continue
        return t
    return ''


def rule_prereference(model, report):
    """每个图/表/公式的**前一段**必须引用它（如图N所示 / 如表N所示 / 如式（N）所示）。

    用户 2026-10-13 要求：图表公式出现之前一定要有解说文字，格式为
    「如图1」「如表1」「如式（1）所示」。
    """
    where = '正文表格'
    for si in range(min(7, len(model.sections))):
        paras = model.sections[si]['paras']
        for pi, p in enumerate(paras):
            if is_img(p._p):
                cap, layout = resolve_fig_caption(paras, pi)
                m = CAP_FIG.match(cap) if (cap and layout) else None
                if not m:
                    continue
                prev = nearest_prev_text(paras, pi)
                if ('图' + m.group(1)) not in prev:
                    report.warn('ref.figure-preface',
                                '图%s 的前一段没有提到它(应写「…如图%s所示」): 前文 = %s'
                                % (m.group(1), m.group(1), pv(prev, 30)),
                                '%s 第%d行' % (where, si + 1))
        blocks = model.sections[si]['blocks']
        for bi, (k, _blk) in enumerate(blocks):
            if k != 'tbl':
                continue
            if bi == 0 or blocks[bi - 1][0] != 'p':
                continue
            cap_txt = p_text(blocks[bi - 1][1]._p)
            m = CAP_TBL.match(cap_txt)
            if not m:
                continue
            prev = ''
            for bj in range(bi - 2, -1, -1):
                if blocks[bj][0] != 'p':
                    continue
                t = p_text(blocks[bj][1]._p)
                if not t or CAP_TBL.match(t) or CAP_FIG.match(t):
                    continue
                prev = t
                break
            if ('表' + m.group(1)) not in prev:
                report.warn('ref.table-preface',
                            '表%s 的前一段没有提到它(应写「…如表%s所示」): 前文 = %s'
                            % (m.group(1), m.group(1), pv(prev, 30)),
                            '%s 第%d行' % (where, si + 1))
        for pi, p in enumerate(paras):
            if not is_math(p._p):
                continue
            m = EQ_NUM.search(p_text(p._p))
            if not m:
                continue
            prev = nearest_prev_text(paras, pi)
            if not re.search(r'式\s*[（(]\s*%s\s*[)）]' % re.escape(m.group(1)), prev):
                report.warn('ref.equation-preface',
                            '式（%s）的前一段没有提到它(应写「…如式（%s）所示」): 前文 = %s'
                            % (m.group(1), m.group(1), pv(prev, 30)),
                            '%s 第%d行' % (where, si + 1))
    return


# ---------------------------------------------------------------- 规则 7 公式
def rule_equations(model, report):
    eqs = [(si, p) for si in range(min(7, len(model.sections)))
           for p in model.sections[si]['paras'] if is_math(p._p)]
    if not eqs:
        report.info('equation.count', '未检测到 OMML 公式')
        return
    report.info('equation.count', '检测到 %d 个含 OMML 公式的段落(m:oMath/m:oMathPara)' % len(eqs))
    unnum, nums = [], set()
    for si, p in eqs:
        t = p_text(p._p)
        m = EQ_NUM.search(t)
        if m:
            nums.add(m.group(1))
        tag = '第%d行 %s' % (si + 1, pv(t, 22))
        if not m and len(EQ_NUM.sub('', t).strip()) <= 3:   # 纯公式段必须有编号
            unnum.append(tag)
    # 「连续公式段落」必须**真的相邻**：同一节里第 pi-1 段也是公式段才算。
    # （旧写法只记住"上一个公式"就配对，等于恒报「公式数-1」处，是误报，2026-10-13 修正。）
    consec = []
    for si in range(min(7, len(model.sections))):
        paras = model.sections[si]['paras']
        for pi in range(1, len(paras)):
            if is_math(paras[pi - 1]._p) and is_math(paras[pi]._p):
                consec.append(('第%d行 %s' % (si + 1, pv(p_text(paras[pi - 1]._p), 22)),
                               '第%d行 %s' % (si + 1, pv(p_text(paras[pi]._p), 22))))
    if unnum:
        report.warn('equation.number', '%d 个纯公式段落没有末尾括号编号(如「（1）」): %s'
                    % (len(unnum), '; '.join(unnum[:6])))
    else:
        report.info('equation.number', '纯公式段落均带末尾编号')
    refs = set()
    for p in model.doc_paras():
        refs |= set(EQ_REF.findall(p_text(p._p)))
    if nums:
        unref = sorted(nums - refs, key=int)
        if unref:
            report.warn('equation.ref-unreferenced', '编号公式 %s 在正文中从未被引用(建议写「由式（N）得」)'
                        % '、'.join('式（%s）' % n for n in unref))
        else:
            report.info('equation.ref', '全部编号公式均被正文引用')
        bad = sorted(refs - nums, key=int)
        if bad:
            report.warn('equation.ref', '正文引用「式（%s）」，但没有对应编号的公式'
                        % '）、式（'.join(bad))
    if consec:
        report.info('equation.consecutive', '存在连续公式段落而无说明文字(%d 处): %s'
                    % (len(consec), '; '.join('%s → %s' % (a, b) for a, b in consec[:4])))


# ---------------------------------------------------------------- 规则: 六节不出现公式
def rule_conclusion_plain(model, report, kind):
    """六、结果陈述 只写文字 + 数值, 不放公式(用户 2026-10-13 要求)。

    数据处理(含求平均值、相对误差)全部集中在五、数据处理; 六节出现
    OMML 公式、未转换的 $$LaTeX$$ 一律 ERROR, 引用「式（N）」则 WARN。
    """
    if model.main is None:
        return
    si = None
    for i in range(min(7, len(model.sections))):
        if '结果陈述' in model.heading_of(i):
            si = i
            break
    if si is None:
        return
    row = '正文表格 第%d行' % (si + 1)
    hard, soft = [], []
    for p in model.sections[si]['paras']:
        t = p_text(p._p)
        tag = '%s / %s' % (row, pv(t, 26) or '（空段）')
        if is_math(p._p):
            hard.append(tag)
        elif re.search(r'\$\$|\\frac|\\sum|\\bar|\\overline|\\Delta|\\times', t):
            hard.append(tag + ' ← 未转换的 LaTeX')
        elif EQ_REF.search(t):
            soft.append(tag)
    if hard:
        report.error('section.equation-in-conclusion',
                     '六、结果陈述 里出现 %d 处公式(用户要求本节不出公式): %s。'
                     '请把算式、代入过程全部挪到五、数据处理, 本节只留文字 + 最终数值'
                     % (len(hard), '; '.join(hard[:4])), row)
    if soft:
        report.warn('section.equation-in-conclusion',
                    '六、结果陈述 里引用了「式（N）」%d 处: %s。'
                    '本节宜直接给数值, 不必再回指公式' % (len(soft), '; '.join(soft[:3])), row)
    if not hard and not soft:
        report.info('section.equation-in-conclusion', '六、结果陈述 无公式(符合要求)')


# ---------------------------------------------------------------- 规则 8 字体
def rule_fonts(model, report, kind):
    combos, big, elevated = {}, [], []
    n_runs = n_no_size = 0
    for si in range(min(7, len(model.sections))):
        paras = model.sections[si]['paras']
        for pi, p in enumerate(paras):
            t = p_text(p._p)
            caption = bool(CAP_FIG.match(t) or CAP_TBL.match(t))
            for run in p.runs:
                txt = (run.text or '').strip()
                if not txt:
                    continue
                font, size = run_font(run), run_size(run)
                if (font or '') in ('Cambria Math', 'Symbol'):
                    continue                     # 公式字体, 不算正文
                key = (font or '(继承Normal)', size)
                combos.setdefault(key, 0)
                combos[key] += 1
                n_runs += 1
                if size is None:
                    n_no_size += 1
                elif size > 14:
                    big.append((size, '第%d行 %s' % (si + 1, pv(txt, 24))))
                elif size > 10.5 and not caption and pi > 0:
                    elevated.append((size, '第%d行 %s' % (si + 1, pv(txt, 24))))
    if not combos:
        report.info('font.scan', '正文未找到带文字的 run, 跳过字体检查')
        return
    normal_pt = 10.5
    try:
        if model.doc.styles['Normal'].font.size:
            normal_pt = model.doc.styles['Normal'].font.size.pt
    except Exception:  # noqa: BLE001
        pass
    ordered = sorted(combos.items(), key=lambda kv: (-kv[1], str(kv[0])))
    desc = '; '.join('%s/%s x%d' % (k[0], ('%.1fpt' % k[1]) if k[1] else '无显式字号', n)
                     for k, n in ordered[:8])
    if len(ordered) > 8:
        desc += '; …(共 %d 种组合)' % len(ordered)
    report.info('font.combos', '正文段落 run 的(字体,字号)组合: %s' % desc)
    if n_no_size:
        report.info('font.inherit', '%d/%d 个 run 未显式设字号, 继承 Normal 样式(%.1fpt), 符合要求'
                    % (n_no_size, n_runs, normal_pt))
    if big:
        report.warn('font.size-large', '正文出现大于 14pt 的字号: %s (示例: %s)'
                    % ('、'.join('%.0fpt' % s for s in sorted({b[0] for b in big}, reverse=True)),
                       '; '.join(b[1] for b in big[:4])))
    else:
        report.info('font.size-large', '正文未发现 >14pt 的显式字号')
    if elevated:
        report.info('font.size-elevated',
                    '有 %d 个正文 run 显式字号为 %s, 大于正文标准 10.5pt(五号);'
                    ' 若属标题/图表名则合规, 否则建议统一'
                    % (len(elevated),
                       '、'.join('%.1fpt' % s for s in sorted({e[0] for e in elevated}, reverse=True))))
    if kind == 'full':
        weird = [(k, n) for k, n in ordered
                 if k[0] != '(继承Normal)' and not any(h in k[0] for h in CJK_FONT_HINTS)]
        if weird:
            report.info('font.ea', '非宋体东亚字体的 run: %s'
                        % '; '.join('%s x%d' % (k[0], n) for k, n in weird[:6]))


# ---------------------------------------------------------------- 规则 9 原始数据页
def rule_rawdata(model, report):
    in_table_rows, before = [], []
    for si, sec in enumerate(model.sections):
        if any(RAW_TEXT.search(p_text(el)) for el in sec['tc'].iter(W_P)):
            in_table_rows.append(si + 1)
    for el in model.head_els():
        if el.tag == W_P and RAW_TEXT.search(p_text(el)):
            before.append(p_text(el))
    tail = [(el, p_text(el)) for el in model.tail_els() if el.tag == W_P]
    markers = [i for i, (_el, t) in enumerate(tail) if RAW_MARK.match(t)]
    if not markers:
        extra = []
        if in_table_rows:
            extra.append('该字样仅出现在正文表格内部第 %s 行(模板提示语)'
                         % '、'.join(str(r) for r in in_table_rows))
        if before:
            extra.append('或出现在正文表格之前')
        notes = [t for _el, t in tail if RAW_TEXT.search(t)]
        if notes and not markers:
            extra.append('表格后虽有相关字样但不是独立页标题: %s' % pv(notes[0], 34))
        report.warn('rawdata.missing', '报告内没有「原始数据记录表」页(默认就不排这一页，'
                    '原始数据记录表另行打印/手写后附上即可；要排进报告用 --rawdata-page)'
                     + ('；' + '；'.join(extra) if extra else ''), '正文表格之后')
        return
    mi = markers[0]
    report.info('rawdata.ok', '「原始数据记录表」页位于正文表格之后(表后第 %d 段)' % (mi + 1),
                '正文表格之后')
    rest = ' '.join(t for _el, t in tail[mi:])
    if not re.search(r'组\s*号', rest) or not re.search(r'姓\s*名', rest):
        report.warn('rawdata.head', '「原始数据记录表」页缺少「组号 … 姓名 …」标识行',
                    '正文表格之后')
    has_tbl = any(el.tag == W_TBL for el in model.tail_els())
    if not has_tbl and not any(t.strip() for _el, t in tail[mi + 1:]):
        report.warn('rawdata.empty', '「原始数据记录表」页除标题外没有任何内容',
                    '正文表格之后')


# ---------------------------------------------------------------- 规则 10 思考题
def rule_thinking(model, report):
    if model.main is None or len(model.sections) < 7:
        return
    where = '正文表格 第7行'
    texts = [(p, p_text(p._p)) for p in model.content_of(6)]
    qs = [(i, t) for i, (_p, t) in enumerate(texts)
          if t and Q_NUM.match(t) and len(Q_NUM.match(t).group(2)) >= 4]
    if not qs:
        body = ' '.join(t for _p, t in texts if t).strip()
        if body and not model.section_is_empty(6):
            report.warn('thinking.question-missing',
                        '七、思考题 中未发现以「1」「2」…编号的题干, 只有答案/正文(缺题干)', where)
        return
    no_ans = []
    for qi, (i, t) in enumerate(qs):
        end = qs[qi + 1][0] if qi + 1 < len(qs) else len(texts)
        answered = any(texts[j][1] and len(texts[j][1].strip()) >= 8
                       for j in range(i + 1, end))
        if not answered:
            no_ans.append(pv(t, 26))
    if no_ans:
        report.warn('thinking.no-answer', '%d 道思考题后面没有作答内容: %s'
                    % (len(no_ans), '; '.join(no_ans[:5])), where)
    else:
        report.info('thinking.answered',
                    '七、思考题 %d 道题均已作答（本模板不写「答：」字样）' % len(qs), where)


# ---------------------------------------------------------------- 规则 11 统计
def rule_stats(model, report):
    n_p = len(list(model.body.iter(W_P)))
    n_tbl = len(list(model.body.iter(W_TBL)))
    n_img_total = sum(n_img(el) for el in model.body.iter(W_P))
    page_brs = [b for b in model.body.iter(W_BR) if b.get(qn('w:type')) == 'page']
    report.info('stats.counts', '全文段落 %d 个, 表格 %d 个(正文表格内 %d), 图片 %d 张'
                % (n_p, n_tbl,
                   len(list(model.main_el.iter(W_TBL))) if model.main_el is not None else 0,
                   n_img_total))
    report.info('stats.pages', '分页代理: 显式分页符 %d 个, 节(sectPr) %d 个, Word 已渲染分页标记 %d 个'
                % (len(page_brs), len(list(model.body.iter(W_SECTPR))),
                   len(list(model.body.iter(qn('w:lastRenderedPageBreak'))))))
    report.info('stats.images', '文档关系中的图片资源 %d 个'
                % len([r for r in model.doc.part.rels.values() if 'image' in r.reltype]))


# ---------------------------------------------------------------- 主流程
def check(path, kind):
    report = Report()
    try:
        model = DocModel(path)
    except Exception as exc:  # noqa: BLE001
        report.error('doc.open', '无法打开/解析文档: %s: %s' % (type(exc).__name__, exc), path)
        return report
    try:
        values = cover_values(model)
    except Exception:  # noqa: BLE001
        values = {}
    steps = [
        ('cover', lambda: rule_cover(model, report, values)),
        ('date', lambda: rule_dates(model, report, values)),
        ('section', lambda: rule_sections(model, report)),
        ('section.empty', lambda: rule_emptiness(model, report, kind)),
        ('figure', lambda: rule_figures(model, report)),
        ('ref', lambda: rule_references(model, report)),
        ('ref.preface', lambda: rule_prereference(model, report)),
        ('equation', lambda: rule_equations(model, report)),
        ('equation.conclusion', lambda: rule_conclusion_plain(model, report, kind)),
        ('font', lambda: rule_fonts(model, report, kind)),
        ('rawdata', lambda: rule_rawdata(model, report)),
        ('thinking', lambda: rule_thinking(model, report)),
        ('stats', lambda: rule_stats(model, report)),
    ]
    for name, fn in steps:
        guard(report, name)(fn)()
    return report


def render(report, path, kind):
    out = io.StringIO()
    e, w = report.counts()
    out.write('=' * 72 + '\n实验报告检查: %s\n模式: --kind %s\n' % (os.path.basename(path), kind)
              + '=' * 72 + '\n')
    for sev in ('ERROR', 'WARN', 'INFO'):
        items = [i for i in report.sorted_items() if i['severity'] == sev]
        out.write('\n--- %s (%d) ---\n' % (sev, len(items)))
        if not items:
            out.write('  (无)\n')
        for it in items:
            out.write('  * %s: %s%s\n' % (it['rule'], it['message'],
                                          ('  [%s]' % it['where']) if it['where'] else ''))
    out.write('\nSUMMARY errors=%d warnings=%d\n' % (e, w))
    return out.getvalue()


def main(argv=None):
    ap = argparse.ArgumentParser(description='深圳大学实验报告 .docx 格式与完整性检查')
    ap.add_argument('docx', help='待检查的 .docx 文件')
    ap.add_argument('--json', dest='json_out', default=None, help='把 findings 写为 JSON 数组')
    ap.add_argument('--kind', dest='kind', choices=['preview', 'full'], default='full',
                    help='preview=预习报告(五/六/七允许为空), full=完整报告(默认)')
    args = ap.parse_args(argv)
    # Windows 控制台多为 GBK, 直接打印 ∆/⋅/… 会 UnicodeEncodeError
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except (AttributeError, ValueError):
            pass
    if not os.path.isfile(args.docx):
        sys.stderr.write('找不到文件: %s\n' % args.docx)
        return 2
    report = check(args.docx, args.kind)
    sys.stdout.write(render(report, args.docx, args.kind))
    if args.json_out:
        try:
            with io.open(args.json_out, 'w', encoding='utf-8') as fh:
                json.dump(report.sorted_items(), fh, ensure_ascii=False, indent=2)
            sys.stdout.write('JSON findings 已写入: %s\n' % args.json_out)
        except Exception as exc:  # noqa: BLE001
            sys.stderr.write('写 JSON 失败: %s\n' % exc)
    return 1 if report.counts()[0] else 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        sys.exit(2)
