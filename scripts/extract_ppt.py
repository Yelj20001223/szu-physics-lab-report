"""Extract lecture-slide content and figures from a PPTX for report writing.

Produces, inside ``--out-dir``:

* ``outline.md``      per-slide text in reading order, tables as Markdown,
                      speaker notes, and one line per extracted picture;
* ``figures/``        every picture, named ``slideNN-figNN-<sha1>.png``;
* ``figures/index.json``  slide number, size in px, byte size, sha1 for each one.

Usage
-----
    python extract_ppt.py --pptx lecture.pptx --out-dir work/ppt
    python extract_ppt.py --pptx lecture.pptx --out-dir work/ppt --slides 4-14
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys

from pptx import Presentation
from pptx.util import Emu

SKIP_PREFIX = ("Picture ", "图片 ", "Image ", "WordArt ")
MIN_USEFUL_PX = 90          # ignore bullets, icons and other decorations
GROUP_TYPE = 6


def emu_to_px(value):
    return int(round(Emu(value).inches * 96)) if value is not None else None


def shape_rect(shape):
    try:
        return (shape.top or 0, shape.left or 0)
    except Exception:
        return (0, 0)


def iter_shapes(shapes):
    """Yield shapes depth-first, recursing into groups."""
    ordered = sorted(list(shapes), key=shape_rect)
    for shape in ordered:
        if shape.shape_type == GROUP_TYPE or shape.__class__.__name__ == "GroupShape":
            try:
                for sub in iter_shapes(shape.shapes):
                    yield sub
                continue
            except Exception:
                pass
        yield shape


def collect_slide(slide, slide_no, figures_dir, index, seen, min_px):
    lines = []
    try:
        title = slide.shapes.title
        if title is not None and title.text.strip():
            lines.append("**标题：%s**" % title.text.strip())
    except Exception:
        pass
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs).strip()
                if text:
                    lines.append(("  " * para.level) + text)
        if getattr(shape, "has_table", False) and shape.has_table:
            rows = []
            for row in shape.table.rows:
                rows.append([cell.text.replace("\n", " ").strip() for cell in row.cells])
            lines.append("")
            lines.append("表格（幻灯片原表，可用于原始数据记录表）：")
            lines.append("| " + " | ".join(rows[0]) + " |")
            lines.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
            for row in rows[1:]:
                lines.append("| " + " | ".join(row) + " |")
            lines.append("")
        if shape.shape_type == 13:  # PICTURE
            try:
                image = shape.image
            except Exception:
                continue
            blob = image.blob
            sha = hashlib.sha1(blob).hexdigest()
            width = emu_to_px(shape.width)
            height = emu_to_px(shape.height)
            px = image.size
            useful = px[0] >= min_px and px[1] >= min_px
            if sha in seen:
                lines.append("- [图片] %s（与 %s 相同，未重复导出）"
                             % (sha[:8], seen[sha]))
                continue
            fig_no = 0
            for entry in index:
                if entry["slide"] == slide_no:
                    fig_no += 1
            ext = (image.ext or "png").lower()
            if ext == "jpeg":
                ext = "jpg"
            name = "slide%02d-fig%02d-%s.%s" % (slide_no, fig_no + 1, sha[:8], ext)
            path = os.path.join(figures_dir, name)
            with open(path, "wb") as fh:
                fh.write(blob)
            seen[sha] = name
            index.append({
                "slide": slide_no, "file": name, "px": list(px),
                "bytes": len(blob), "sha1": sha, "useful": useful,
            })
            note = "" if useful else "（尺寸很小，多半是图标/装饰）"
            lines.append("- [图片] `figures/%s` %dx%d px %s" % (name, px[0], px[1], note))
        if shape.shape_type == 7:  # EMBEDDED_OLE_OBJECT
            lines.append("- [公式对象] 幻灯片里是 OLE 公式，内容需按上下文重写为 LaTeX。")
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            lines.append("")
            lines.append("> 备注：%s" % notes.replace("\n", " "))
    return lines


def parse_range(spec, total):
    if not spec:
        return list(range(1, total + 1))
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            out.extend(range(int(start), int(end) + 1))
        else:
            out.append(int(part))
    return [n for n in out if 1 <= n <= total]


def main(argv=None):
    parser = argparse.ArgumentParser(description="Extract PPTX text and figures")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--slides", help="e.g. 4-14 or 4,6,8 (1-based)")
    parser.add_argument("--min-px", type=int, default=MIN_USEFUL_PX,
                        help="pictures smaller than this are flagged as decoration")
    args = parser.parse_args(argv)

    prs = Presentation(args.pptx)
    total = len(prs.slides)
    wanted = set(parse_range(args.slides, total))
    figures_dir = os.path.join(args.out_dir, "figures")
    os.makedirs(figures_dir, exist_ok=True)

    index, seen, out = [], {}, []
    out.append("# %s" % os.path.basename(args.pptx))
    out.append("")
    out.append("幻灯片 %d 页；已导出图片见 `figures/`。正文按阅读顺序（上→下、左→右）排列。"
               % total)
    for number, slide in enumerate(prs.slides, 1):
        if number not in wanted:
            continue
        out.append("")
        out.append("## 第 %d 页" % number)
        out.append("")
        out.extend(collect_slide(slide, number, figures_dir, index, seen, args.min_px))

    with io.open(os.path.join(args.out_dir, "figures", "index.json"), "w",
                 encoding="utf-8") as fh:
        json.dump(index, fh, ensure_ascii=False, indent=1)
    with io.open(os.path.join(args.out_dir, "outline.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(out) + "\n")

    useful = [e for e in index if e["useful"]]
    print("WROTE %s" % os.path.join(os.path.abspath(args.out_dir), "outline.md"))
    print("FIGURES %d total, %d likely useful" % (len(index), len(useful)))
    for entry in index:
        flag = "  " if entry["useful"] else " ?"
        print("%s slide %2d  %-40s %dx%d" % (flag, entry["slide"], entry["file"],
                                             entry["px"][0], entry["px"][1]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
