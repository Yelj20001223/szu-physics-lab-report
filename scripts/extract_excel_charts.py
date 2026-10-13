"""把 Excel 里已经做好的图表导出成图片，供实验报告使用（不要自己重新画图）。

用法
----
1) 打印每个图表的位置和要执行的渲染命令：

       python extract_excel_charts.py --xlsx 数据.xlsx

   脚本会读出每个图表的锚点单元格与尺寸，算出一个略大于图表的 A1 区域，
   并打印一条 libreoffice-kit 的 render 命令（渲染指定区域为 PNG）。

2) 按提示执行 render（需要能跑 LibreOffice；DSH 沙箱下要一次性开完全权限），
   然后自动裁剪白边、按顺序输出图片：

       python extract_excel_charts.py --crop-dir work/charts --out-dir figures

裁剪后的文件名是 excel-chart-1.png、excel-chart-2.png……顺序与 plan 打印的顺序一致。
"""

from __future__ import annotations

import argparse
import glob
import math
import os
import sys

import openpyxl

# 经验换算：默认列宽 8.43 字符 ≈ 1.7 cm，默认行高 ≈ 0.53 cm
COL_CM = 1.7
ROW_CM = 0.53


def col_letter(index):
    """0 -> A, 25 -> Z, 26 -> AA"""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def plan(xlsx, dpi=150):
    wb = openpyxl.load_workbook(xlsx)
    items = []
    for ws in wb.worksheets:
        for chart in ws._charts:
            anchor = chart.anchor
            if not hasattr(anchor, "_from"):
                items.append((ws.title, None, None, chart.width, chart.height))
                continue
            frm = anchor._from
            start_col, start_row = frm.col, frm.row
            ncols = int(math.ceil((chart.width or 15) / COL_CM)) + 2
            nrows = int(math.ceil((chart.height or 7.5) / ROW_CM)) + 3
            rng = "%s%d:%s%d" % (col_letter(start_col), start_row + 1,
                                 col_letter(start_col + ncols), start_row + nrows)
            items.append((ws.title, rng, (start_col, start_row), chart.width, chart.height))
    sheets = {ws.title for ws in wb.worksheets}
    for i, (sheet, rng, _pos, w, h) in enumerate(items, 1):
        print("图表 %d: sheet=%s  区域=%s  尺寸=%.1fx%.1fcm" % (i, sheet, rng, w or 0, h or 0))
    print()
    print("渲染命令（一次渲染全部图表，输出目录必须是新的）：")
    print('  node <libreofficeKit.node> <libreofficeKit.cli> render \\')
    print('      --input "%s" \\' % os.path.abspath(xlsx))
    print('      --output-dir "<workdir>/charts" \\')
    if len(sheets) == 1 and items and items[0][1]:
        only = list(sheets)[0]
        print('      --sheet %s --range %s --dpi %d' % (only, items[0][1], dpi))
        if len(items) > 1:
            print("  （同一张工作表上有多个图表时，逐个改成上面的区域分别渲染，") 
            print("    例如再跑一次 --range %s）" % items[1][1])
    else:
        print("      --sheet <工作表名> --range <区域> --dpi %d" % dpi)
    return 0


def crop(crop_dir, out_dir, pad=6, threshold=245):
    from PIL import Image
    files = sorted(glob.glob(os.path.join(crop_dir, "*.png")))
    if not files:
        print("没有找到 PNG：%s" % crop_dir, file=sys.stderr)
        return 1
    os.makedirs(out_dir, exist_ok=True)
    for i, path in enumerate(files, 1):
        im = Image.open(path).convert("RGB")
        gray = im.convert("L")
        # 找出非白像素的范围，裁掉四周白边
        mask = gray.point(lambda v: 0 if v >= threshold else 255)
        box = mask.getbbox()
        if box is None:
            print("WARN %s 全是白的，跳过" % path, file=sys.stderr)
            continue
        left = max(box[0] - pad, 0)
        top = max(box[1] - pad, 0)
        right = min(box[2] + pad, im.width)
        bottom = min(box[3] + pad, im.height)
        out = os.path.join(out_dir, "excel-chart-%d.png" % i)
        im.crop((left, top, right, bottom)).save(out)
        print("WROTE %s  (%dx%d, 源 %s)" % (out, right - left, bottom - top, os.path.basename(path)))
    return 0


def crop_box(page, box, out, pad=0):
    """按像素框裁剪渲染出来的页面，取出其中一张图表。"""
    from PIL import Image
    im = Image.open(page).convert("RGB")
    x0, y0, x1, y1 = box
    x0 = max(int(x0) - pad, 0)
    y0 = max(int(y0) - pad, 0)
    x1 = min(int(x1) + pad, im.width)
    y1 = min(int(y1) + pad, im.height)
    if x1 <= x0 or y1 <= y0:
        print("裁剪框无效: %s" % (box,), file=sys.stderr)
        return 1
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    im.crop((x0, y0, x1, y1)).save(out)
    print("WROTE %s  (%dx%d, 源 %s)" % (out, x1 - x0, y1 - y0, os.path.basename(page)))
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="导出 Excel 里已做好的图表")
    parser.add_argument("--xlsx", help="plan 模式：工作簿路径")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--crop-dir", help="crop 模式：渲染出来的 PNG 目录")
    parser.add_argument("--out-dir", default="figures")
    parser.add_argument("--pad", type=int, default=6)
    parser.add_argument("--page", help="cropbox 模式：渲染出来的某一页 PNG")
    parser.add_argument("--box", help="cropbox 模式：x0,y0,x1,y1（像素）")
    parser.add_argument("--out", help="cropbox 模式的输出文件")
    args = parser.parse_args(argv)
    if args.page and args.box and args.out:
        box = [int(v) for v in args.box.split(",")]
        if len(box) != 4:
            parser.error("--box 需要 x0,y0,x1,y1 四个数")
        return crop_box(args.page, box, args.out, args.pad)
    if args.crop_dir:
        return crop(args.crop_dir, args.out_dir, args.pad)
    if args.xlsx:
        return plan(args.xlsx, args.dpi)
    parser.error("需要 --xlsx（plan）、--crop-dir（crop）或 --page/--box/--out（cropbox）")


if __name__ == "__main__":
    sys.exit(main())
