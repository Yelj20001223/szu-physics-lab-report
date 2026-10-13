#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""清掉 .docx / .xlsx 里的"遗留文档属性"，只留你指定的署名。

为什么需要它：Office 文档的属性里常年躺着别人的痕迹——模板原作者的
`Company`、累计编辑时长 `TotalTime`、上次打印时间 `lastPrinted`、以及
修订/批注的 `w:author`。这些不一定是"个人隐私"，但公开分享模板或报告前
清一遍更干净；关键词扫描（`scan-pii.py`）查不到这类结构性元数据。

处理项：
  docProps/core.xml  : dc:creator / cp:lastModifiedBy -> 改成 --creator
                       cp:lastPrinted                -> 删除
                       dc:title                      -> 改成 --title（给了才改）
  docProps/app.xml   : Company                       -> 删除
                       TotalTime                     -> 置 0
  docProps/custom.xml: 整份删除（多为 WPS 版本号之类的生成器信息）
                       同时清掉 [Content_Types].xml 与 *.rels 里对它的引用
                       —— 不清引用 Word 会报"文件已损坏"
  正文 part          : 删掉 w:author / w:initials / w:date（修订与批注作者标记）

注意：docProps 不参与排版，改这些字**不会**影响版式、页码、字体。

用法:
  python scrub_docx_props.py "assets/base.docx" --dry-run
  python scrub_docx_props.py "报告.docx" --creator 张三 --title 某某实验报告
"""
import argparse
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

DROP_PARTS = ("docProps/custom.xml",)
PROPS_REL = "custom-properties"


def strip_tag(xml, tag, reason):
    """删掉 <tag .../> 与 <tag ...>...</tag>。"""
    xml, n1 = re.subn(r"<%s(?:\s[^>]*)?/>" % tag, "", xml)
    xml, n2 = re.subn(r"<%s(?:\s[^>]*)?>.*?</%s>" % (tag, tag), "", xml, flags=re.S)
    if n1 + n2:
        reason.append("删除 <%s> x%d" % (tag, n1 + n2))
    return xml


def set_tag(xml, tag, value, reason):
    pat = re.compile(r"(<%s(?:\s[^>]*)?>)(.*?)(</%s>)" % (tag, tag), re.S)
    m = pat.search(xml)
    if not m or m.group(2) == value:      # 没有该标签 / 值本来就一样，都不算改动
        return xml
    xml = pat.sub(lambda mm: mm.group(1) + value + mm.group(3), xml)
    reason.append("<%s> -> %r" % (tag, value))
    return xml


def drop_attr(xml, attr, reason):
    xml, n = re.subn(r'\s%s="[^"]*"' % attr, "", xml)
    if n:
        reason.append("去属性 %s x%d" % (attr, n))
    return xml


def drop_refs(xml, part, reason, where=""):
    """从 [Content_Types].xml / *.rels 里删掉指向 part 的引用（含关系与类型声明）。"""
    target = part.lstrip("/")
    n = 0
    # <Override PartName="/docProps/custom.xml" .../>
    xml, k = re.subn(r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(target), "", xml)
    n += k
    # <Relationship ... Target="docProps/custom.xml" .../>
    xml, k = re.subn(r'<Relationship[^>]*Target="%s"[^>]*/>' % re.escape(target), "", xml)
    n += k
    # <Relationship ... Type=".../custom-properties" .../>（Target 写法可能不同）
    xml, k = re.subn(r'<Relationship[^>]*Type="[^"]*%s"[^>]*/>' % PROPS_REL, "", xml)
    n += k
    if n:
        reason.append("清引用 %s 对 %s 的 %d 处引用" % (where or "?", part, n))
    return xml


def scrub(path, creator, title, dry_run):
    path = Path(path)
    if not path.is_file():
        print("找不到文件:", path)
        return 1
    src = zipfile.ZipFile(path)
    reasons, out = [], {}

    for item in src.infolist():
        name = item.filename
        data = src.read(name)
        if name in DROP_PARTS:
            reasons.append("删除 %s" % name)
            continue
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue

        xml = before = data.decode("utf8")
        if name == "docProps/core.xml":
            xml = set_tag(xml, "dc:creator", creator, reasons)
            xml = set_tag(xml, "cp:lastModifiedBy", creator, reasons)
            xml = strip_tag(xml, "cp:lastPrinted", reasons)
            if title is not None:
                xml = set_tag(xml, "dc:title", title, reasons)
        elif name == "docProps/app.xml":
            xml = strip_tag(xml, "Company", reasons)
            xml = set_tag(xml, "TotalTime", "0", reasons)
        elif name == "[Content_Types].xml" or name.endswith(".rels"):
            # 被删部件的引用必须一并清掉
            for part in DROP_PARTS:
                if part in xml:
                    xml = drop_refs(xml, part, reasons, name)
        if name.startswith("word/document.xml") or name.startswith("word/header") \
                or name.startswith("word/footer"):
            xml = drop_attr(xml, "w:author", reasons)
            xml = drop_attr(xml, "w:initials", reasons)
            xml = drop_attr(xml, "w:date", reasons)
        if xml != before:
            out[name] = xml.encode("utf8")

    if not out:
        print("无需改动:", path)
        return 0
    print("将改动 %s:" % path)
    for r in dict.fromkeys(reasons):
        print("   -", r)
    if dry_run:
        print("(--dry-run，未写盘)")
        return 0

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = path.with_suffix(path.suffix + ".bak-" + ts)
    shutil.copy2(path, bak)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename in DROP_PARTS:
                continue
            dst.writestr(item, out.get(item.filename, src.read(item.filename)))
    src.close()
    tmp.replace(path)
    print("已写盘（备份 %s）" % bak.name)
    return 0


def main():
    ap = argparse.ArgumentParser(description="清理 docx/xlsx 的遗留文档属性")
    ap.add_argument("path", help="要处理的文件")
    ap.add_argument("--creator", default="szu-physics-lab-report",
                    help="署名写入 dc:creator / cp:lastModifiedBy（默认 szu-physics-lab-report）")
    ap.add_argument("--title", default=None, help="顺便把 dc:title 改成这个（不给则不动）")
    ap.add_argument("--dry-run", action="store_true", help="只打印要改什么，不写盘")
    a = ap.parse_args()
    return scrub(a.path, a.creator, a.title, a.dry_run)


if __name__ == "__main__":
    sys.exit(main())
