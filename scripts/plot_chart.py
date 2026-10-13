#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Render physics-lab data charts to PNG with Pillow + numpy only (no matplotlib).

    python plot_chart.py --spec spec.json --out out.png

Writes the PNG and prints one JSON object to stdout:
    {"out": "...", "fit": {"slope":...,"intercept":...,"r2":...}|null, "points": <int>}

Spec keys, all optional unless noted: "kind" = scatter_fit|line|scatter|multi
(default line), "x_label", "y_label", "title", "x"/"y" (required except for
"multi"), "series" [{"name","x","y","style"}], "fit" {"type","x0",
"show_equation","show_r2"}, "x_log", "y_log", "size" [w,h], "dpi", "grid",
"legend". Put units in the axis labels, e.g. "峰值电压 U (V)". With fit.x0 the
line is pinned through (x0, y(x0)) and only the slope is fitted.
"""

import argparse
import json
import math
import os
import sys

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps

FONT_CANDIDATES = [r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\msyh.ttc",
                   r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\times.ttf"]
COLORS = [(31, 78, 168), (200, 60, 40), (30, 140, 70), (140, 70, 170), (200, 130, 20)]
BLACK, GREY, WHITE = (0, 0, 0), (216, 216, 216), (255, 255, 255)
# Glyphs that common CJK fonts map to .notdef: rewritten to ASCII, never tofu.
ASCII_MAP = {"\u00b2": "^2", "\u00b3": "^3", "\u00b9": "^1", "\u2070": "^0", "\u2074": "^4",
             "\u2075": "^5", "\u2076": "^6", "\u2077": "^7", "\u2078": "^8", "\u2079": "^9",
             "\u207b": "^-", "\u207a": "^+", "\u2103": "\u00b0C", "\u00b5": "u", "\u03bc": "u",
             "\u03a9": "ohm", "\u00b7": ".", "\uff5e": "~", "\u2014": "-", "\u2013": "-",
             "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'}
_font_cache = {}


class SpecError(ValueError):
    """Bad input: reported on stderr with a non-zero exit status."""


# --- text/font -------------------------------------------------------------

def load_font(size):
    """CJK-capable truetype font at `size` px: first candidate that loads."""
    size = max(8, int(size))
    if size not in _font_cache:
        font = None
        for path in FONT_CANDIDATES:
            try:
                font = ImageFont.truetype(path, size)
                break
            except Exception:
                continue
        _font_cache[size] = font or ImageFont.load_default(size)
    return _font_cache[size]


def text_wh(font, s):
    """Ink width/height of `s` in pixels, from the font metrics."""
    if not s:
        return 0.0, 0.0
    try:
        b = font.getbbox(s)
        return float(b[2] - b[0]), float(b[3] - b[1])
    except Exception:
        return float(font.getlength(s)), float(font.size)


def clean_text(s, font):
    """Drop/substitute characters this font cannot draw, so labels never show tofu."""
    def ok(ch):
        try:
            m = font.getmask(ch)
            return (m[0] if isinstance(m, tuple) else m).getbbox() is not None
        except Exception:
            return False
    return "".join(ch if ord(ch) < 127 or ok(ch) else ASCII_MAP.get(ch, "") for ch in str(s))


def text_rotated(img, cx, cy, s, font, fill=BLACK):
    """Draw `s` rotated 90 deg CCW, centred on (cx, cy)."""
    w, h = text_wh(font, s)
    mask = Image.new("L", (int(w) + 8, int(h) + 8), 255)
    ImageDraw.Draw(mask).text((4, 4), s, font=font, fill=0)
    inv = ImageOps.invert(mask).rotate(90, expand=True)
    img.paste(Image.new("RGB", inv.size, fill),
              (int(round(cx - inv.width / 2.0)), int(round(cy - inv.height / 2.0))), inv)


# --- numbers, ticks, formatting --------------------------------------------

def fmt_num(v, prec=None, sig=5):
    """Compact number without float noise."""
    if not np.isfinite(v):
        return "nan"
    if v == 0:
        return "0"
    if prec is not None:
        s = f"{v:.{max(0, int(prec))}f}"
        return s[1:] if s.startswith("-") and float(s) == 0 else s
    if abs(v) >= 1e5 or abs(v) < 1e-4:
        return f"{v:.3e}".replace("e+0", "e+").replace("e-0", "e-")
    return f"{v:.{sig}g}"


def nice_ticks(lo, hi, target=7):
    """'Nice' tick values (1/2/2.5/5 x 10^k) spanning [lo, hi]."""
    if not (np.isfinite(lo) and np.isfinite(hi)) or hi <= lo:
        return [lo]
    raw = (hi - lo) / max(1, target)
    mag = 10.0 ** math.floor(math.log10(raw))
    step = next((m * mag for m in (1.0, 2.0, 2.5, 5.0, 10.0) if raw <= m * mag * 1.001),
                10.0 * mag)
    start, stop = math.ceil(lo / step - 1e-9), math.floor(hi / step + 1e-9)
    if stop < start:
        return [lo, hi]
    if stop - start > 60:
        return nice_ticks(lo, hi, target * 2)
    return [step * (start + i) for i in range(int(round(stop - start)) + 1)]


def axis_ticks(lo, hi, log, target=7):
    """-> (tick positions, labels, minor positions) in the plotted coordinates."""
    if not log:
        pos = np.array(nice_ticks(lo, hi, target), dtype=float)
        step = (pos[1] - pos[0]) if pos.size > 1 else 1.0
        dec = 3 if step <= 0 else max(0, min(12, -int(math.floor(math.log10(step)))))
        return pos, [fmt_num(v, dec) for v in pos], np.array([])
    e0 = int(math.floor(math.log10(lo))) if lo > 0 else 0
    e1 = int(math.ceil(math.log10(hi))) if hi > 0 else 0
    step = (e1 - e0) // 12 + 1 if e1 - e0 > 14 else 1      # label every k-th decade
    majors = [t for t in (10.0 ** e for e in range(e0, e1 + 1, step))
              if lo - 1e-12 <= t <= hi + 1e-12] or [10.0 ** e0]
    minors = [m * t for t in majors for m in (2, 3, 4, 5, 6, 7, 8, 9) if lo < m * t < hi]
    return (np.log10(majors), [fmt_num(v, sig=3) for v in majors],
            np.log10(minors) if minors else np.array([]))


# --- spec parsing ----------------------------------------------------------

def _txt(spec, key, default=""):
    v = spec.get(key)
    if v is None:
        return default
    if isinstance(v, str):
        return v
    if isinstance(v, (int, float, bool)):
        return str(v)
    raise SpecError(f"'{key}' must be a string")


def _as_xy(xs, ys, ctx):
    if xs is None or ys is None:
        raise SpecError(f"{ctx}: 'x' and 'y' are both required")
    try:
        x = np.asarray(xs, dtype=float).ravel()
        y = np.asarray(ys, dtype=float).ravel()
    except (TypeError, ValueError):
        raise SpecError(f"{ctx}: 'x' and 'y' must be lists of numbers")
    if x.size != y.size:
        raise SpecError(f"{ctx}: length mismatch (x={x.size}, y={y.size})")
    keep = np.isfinite(x) & np.isfinite(y)
    if not keep.all():
        print(f"warning: {ctx}: dropped {int((~keep).sum())} non-finite point(s)",
              file=sys.stderr)
        x, y = x[keep], y[keep]
    if x.size == 0:
        raise SpecError(f"{ctx}: no finite data points")
    if x.size == 1:
        print(f"warning: {ctx}: only one point; nothing to fit or connect", file=sys.stderr)
    return x, y


def _parse_fit(fit):
    """Normalise the optional 'fit' block; raises SpecError on nonsense."""
    if fit is True:
        fit = {}
    if fit is not None and fit is not False and not isinstance(fit, (dict, int, float)):
        raise SpecError("'fit' must be an object or null")
    if not isinstance(fit, dict):
        return None
    ftype = _txt(fit, "type", "linear").lower()
    if ftype not in ("linear", "line", "leastsq"):
        raise SpecError(f"unsupported fit type {ftype!r} (only 'linear')")
    x0 = fit.get("x0")
    if x0 is not None:
        try:
            x0 = float(x0)
        except (TypeError, ValueError):
            raise SpecError("fit.x0 must be a number")
        if not np.isfinite(x0):
            raise SpecError("fit.x0 must be finite")
    return {"x0": x0, "show_equation": bool(fit.get("show_equation", True)),
            "show_r2": bool(fit.get("show_r2", True))}


def parse_spec(spec):
    if not isinstance(spec, dict):
        raise SpecError("spec must be a JSON object")
    kind = (_txt(spec, "kind").strip().lower()
            or ("multi" if isinstance(spec.get("series"), list) and spec["series"] else "line"))
    if kind not in ("line", "scatter", "scatter_fit", "multi"):
        raise SpecError(f"unknown kind {kind!r} (expected line|scatter|scatter_fit|multi)")
    groups = []
    if kind == "multi":
        for i, s in enumerate(spec.get("series") or []):
            if not isinstance(s, dict):
                raise SpecError(f"series[{i}] must be an object")
            x, y = _as_xy(s.get("x"), s.get("y"), f"series[{i}]")
            style = _txt(s, "style", "line").lower()
            groups.append({"name": _txt(s, "name") or f"S{i + 1}", "x": x, "y": y,
                           "style": style if style in ("line", "scatter") else "line"})
        if not groups:
            raise SpecError("kind 'multi' requires a non-empty 'series' list")
    else:
        x, y = _as_xy(spec.get("x"), spec.get("y"), "data")
        groups.append({"name": "", "x": x, "y": y,
                       "style": "scatter" if kind in ("scatter", "scatter_fit") else "line"})
    fit = _parse_fit(spec.get("fit"))
    if fit is not None and kind == "multi":
        raise SpecError("'fit' is not supported with kind 'multi'")
    if fit is None and kind == "scatter_fit":
        fit = {"x0": None, "show_equation": True, "show_r2": True}
    size = spec.get("size") or [1000, 700]
    try:
        width, height = max(240, int(size[0])), max(200, int(size[1]))
        dpi = max(36, min(1200, int(spec.get("dpi") or 150)))
    except (TypeError, ValueError, IndexError, KeyError):
        raise SpecError("'size' must be [width, height] and 'dpi' an integer")
    return {"kind": kind, "groups": groups, "x_label": _txt(spec, "x_label") or "x",
            "y_label": _txt(spec, "y_label") or "y", "title": _txt(spec, "title"),
            "x_log": bool(spec.get("x_log", False)), "y_log": bool(spec.get("y_log", False)),
            "grid": bool(spec.get("grid", True)), "legend": bool(spec.get("legend", True)),
            "fit": fit, "width": width, "height": height, "dpi": dpi}


# --- fitting ---------------------------------------------------------------

def fit_series(x, y, x0=None):
    """Least-squares line on the (already log-transformed) data."""
    if x.size < 2:
        raise SpecError("fit requires at least 2 data points")
    mx, my = float(x.mean()), float(y.mean())
    if x0 is None:
        var = float(((x - mx) ** 2).sum())
        if var <= 1e-300:
            raise SpecError("cannot fit: all x values are identical")
        slope = float(((x - mx) * (y - my)).sum() / var)
    else:
        d = x - float(x0)
        den = float((d * d).sum())
        if den <= 1e-300:
            raise SpecError("cannot fit through x0: all x values equal x0")
        slope = float((d * (y - my)).sum() / den)   # pinned so the mean residual is 0
    intercept = my - slope * (mx if x0 is None else float(x0))
    ss_tot = float(((y - my) ** 2).sum())
    if ss_tot <= 1e-300:
        print("warning: y values are all equal; degenerate fit (R2 reported as 1)",
              file=sys.stderr)
        r2 = 1.0
    else:
        r2 = 1.0 - float(((y - (slope * x + intercept)) ** 2).sum()) / ss_tot
    if not (np.isfinite(slope) and np.isfinite(intercept)):
        raise SpecError("fit produced non-finite coefficients")
    return slope, intercept, r2, x0 is not None


def equation_text(slope, intercept, x_log):
    unit = abs(abs(slope) - 1.0) < 1e-12
    head = ("-" if slope < 0 else "") if unit else fmt_num(slope, sig=4)
    out = f"y = {head}{'lg x' if x_log else 'x'}"
    if abs(intercept) > 1e-12:
        out += f" {'+' if intercept > 0 else '-'} {fmt_num(abs(intercept), sig=4)}"
    return out


def dashed_line(draw, p0, p1, fill, dash=9.0, gap=5.0, width=2):
    length = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
    if length <= 1e-9:
        return
    ux, uy = (p1[0] - p0[0]) / length, (p1[1] - p0[1]) / length
    pos = 0.0
    while pos < length:
        end = min(pos + dash, length)
        draw.line([(p0[0] + ux * pos, p0[1] + uy * pos),
                   (p0[0] + ux * end, p0[1] + uy * end)], fill=fill, width=width)
        pos = end + gap


def corner_box(bw, bh, box, pts, prefer):
    """Top-left corner for a bw x bh label box: the emptiest of the four corners."""
    pl, pt, pr, pb = box
    cands = [(pr - 10 - bw, pt + 10), (pl + 10, pt + 10),
             (pr - 10 - bw, pb - 10 - bh), (pl + 10, pb - 10 - bh)]
    if bw > (pr - pl) - 20 or bh > (pb - pt) - 20:
        cands = [(pl + 10, pt + 10)]
    best = None
    for rank, idx in enumerate(prefer):
        x0, y0 = cands[idx]
        cover = sum(1 for px, py in pts
                    if x0 - 6 <= px <= x0 + bw + 6 and y0 - 6 <= py <= y0 + bh + 6)
        if best is None or (cover, rank) < best[0]:
            best = ((cover, rank), x0, y0)
    return best[1], best[2]


# --- render ----------------------------------------------------------------

def render(cfg):
    W, H, dpi = cfg["width"], cfg["height"], cfg["dpi"]
    base = int(round(15 * max(0.8, min(1.8, dpi / 150.0))))
    f_tick, f_axis = load_font(base), load_font(int(round(base * 1.15)))
    f_eq, f_leg = load_font(int(round(base * 1.05))), load_font(int(round(base * 0.95)))
    PAD = max(12, int(round(16 * max(0.8, min(1.8, dpi / 150.0)))))  # >= 12 px padding
    TICK, TGAP, LGAP = max(4, base // 3), max(3, base // 4), max(4, base // 3)
    x_log, y_log, fit_cfg = cfg["x_log"], cfg["y_log"], cfg["fit"]

    def transform(v, log, name):
        if log and np.any(v <= 0):
            raise SpecError(f"{cfg['kind']}: {name}_log requires all {name} > 0")
        return np.log10(v) if log else v

    series = [(g, transform(g["x"], x_log, "x"), transform(g["y"], y_log, "y"),
               clean_text(g["name"], f_leg) or "数据") for g in cfg["groups"]]
    all_x = np.concatenate([s[1] for s in series])
    all_y = np.concatenate([s[2] for s in series])
    slope = intercept = r2 = None
    fixed = False
    if fit_cfg is not None:
        slope, intercept, r2, fixed = fit_series(series[0][1], series[0][2], fit_cfg["x0"])
        fx = np.array([float(all_x.min()), float(all_x.max())])
        all_x = np.concatenate([all_x, fx])
        all_y = np.concatenate([all_y, slope * fx + intercept])
    limits = []
    for vals, log in ((all_x, x_log), (all_y, y_log)):
        lo, hi = float(vals.min()), float(vals.max())
        if hi - lo <= 1e-12:                        # one distinct value: pad it out
            lo, hi = lo - max(abs(lo) * 0.05, 0.5), hi + max(abs(lo) * 0.05, 0.5)
        m = (hi - lo) * 0.06
        limits.append((min(lo - m, lo - 0.03) if log else lo - m, hi + m))
    (x_lo, x_hi), (y_lo, y_hi) = limits
    x_pos, x_lab, x_minor = axis_ticks(x_lo, x_hi, x_log)
    y_pos, y_lab, y_minor = axis_ticks(y_lo, y_hi, y_log)
    x_name = clean_text(cfg["x_label"], f_axis) + (" (log)" if x_log else "")
    y_name = clean_text(cfg["y_label"], f_axis) + (" (log)" if y_log else "")
    title = clean_text(cfg["title"], f_axis)
    names = [s[3] for s in series]
    show_legend = cfg["legend"] and (len(series) > 1 or cfg["kind"] == "multi")
    leg_w = (max(text_wh(f_leg, n)[0] for n in names) + base * 2.6) if show_legend else 0.0
    leg_h = (len(names) * (base + 6) + 8) if show_legend else 0.0
    w_yt = max(text_wh(f_tick, s)[0] for s in y_lab)
    w_xt = max(text_wh(f_tick, s)[0] for s in x_lab)
    h_xt = max(text_wh(f_tick, s)[1] for s in x_lab)
    h_xl, h_yl = text_wh(f_axis, x_name)[1], text_wh(f_axis, y_name)[1]
    h_title = text_wh(f_axis, title)[1] + 10 if title else 0.0
    pl = int(round(PAD + h_yl + LGAP + w_yt + TGAP + TICK + 2))   # rotated label first
    pr = int(round(W - PAD - max(w_xt / 2.0 + 4.0, 6.0)))
    pt = int(round(PAD + TICK + TGAP + h_title))
    pb = int(round(H - PAD - TICK - TGAP - h_xt - LGAP - h_xl - 4))
    if pr - pl < 80 or pb - pt < 80:
        raise SpecError(f"canvas {W}x{H} is too small for these labels; enlarge 'size'")

    img = Image.new("RGB", (W, H), WHITE)
    draw = ImageDraw.Draw(img)
    span_x, span_y = (x_hi - x_lo) or 1.0, (y_hi - y_lo) or 1.0
    X = lambda v: pl + (float(v) - x_lo) / span_x * (pr - pl)      # noqa: E731
    Y = lambda v: pb - (float(v) - y_lo) / span_y * (pb - pt)      # noqa: E731
    txt = lambda xy, s, f, anc="la": draw.text(xy, s, font=f, fill=BLACK, anchor=anc)
    if cfg["grid"]:
        for v in x_pos:
            draw.line([(X(v), pt), (X(v), pb)], fill=GREY, width=1)
        for v in y_pos:
            draw.line([(pl, Y(v)), (pr, Y(v))], fill=GREY, width=1)
    for v in x_minor:
        draw.line([(X(v), pb), (X(v), pb + TICK // 2)], fill=(120, 120, 120), width=1)
    for v in y_minor:
        draw.line([(pl - TICK // 2, Y(v)), (pl, Y(v))], fill=(120, 120, 120), width=1)
    for v, s in zip(x_pos, x_lab):
        draw.line([(X(v), pb), (X(v), pb + TICK)], fill=BLACK, width=1)
        txt((X(v), pb + TICK + TGAP), s, f_tick, "ma")
    for v, s in zip(y_pos, y_lab):
        draw.line([(pl - TICK, Y(v)), (pl, Y(v))], fill=BLACK, width=1)
        txt((pl - TICK - TGAP, Y(v)), s, f_tick, "rm")
    draw.rectangle([pl, pt, pr, pb], outline=BLACK, width=1)
    txt(((pl + pr) / 2.0, pb + TICK + TGAP + h_xt + LGAP), x_name, f_axis, "ma")
    text_rotated(img, PAD + h_yl / 2.0, (pt + pb) / 2.0, y_name, f_axis)
    if title:
        txt(((pl + pr) / 2.0, max(2, PAD - 2)), title, f_axis, "ma")

    data_px = []
    for i, (g, tx, ty, _) in enumerate(series):
        col = COLORS[i % len(COLORS)]
        pts = [(X(a), Y(b)) for a, b in zip(tx, ty)]
        data_px += pts
        if g["style"] == "line" and len(pts) > 1:
            draw.line(pts, fill=col, width=2, joint="curve")
        for cx, cy in pts:
            draw.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], fill=col, outline=WHITE)
    if slope is not None:
        ends = [(X(x_lo), Y(slope * x_lo + intercept)),
                (X(x_hi), Y(slope * x_hi + intercept))]
        dashed_line(draw, *sorted(ends), fill=(20, 20, 20))

    box = (pl, pt, pr, pb)
    if show_legend:
        bx, by = corner_box(leg_w + 8, leg_h, box, data_px, [0, 2, 1, 3])
        draw.rectangle([bx, by, bx + leg_w + 8, by + leg_h], fill=WHITE, outline=(140,) * 3)
        for i, n in enumerate(names):
            col, cy = COLORS[i % len(COLORS)], by + 4 + (i + 0.5) * (base + 6)
            draw.line([(bx + 8, cy), (bx + 8 + base * 1.6, cy)], fill=col, width=2)
            draw.ellipse([bx + 7 + base * 0.8, cy - 3, bx + 9 + base * 0.8 + 2, cy + 3],
                         fill=col, outline=WHITE)
            txt((bx + 8 + base * 1.9, cy), n, f_leg, "lm")
    lines = []
    if slope is not None:
        if fit_cfg["show_equation"]:
            lines.append(equation_text(slope, intercept, x_log))
            if fixed:
                lines.append(f"(\u8fc7\u5b9a\u70b9 x0 = {fmt_num(fit_cfg['x0'], sig=4)})")
        if fit_cfg["show_r2"]:
            lines.append(f"R\u00b2 = {r2:.4f}")
    if lines:
        lw = max(text_wh(f_eq, s)[0] for s in lines)
        lh = max(text_wh(f_eq, s)[1] for s in lines)
        bw, bh = lw + base * 1.4, len(lines) * (lh + 6) + base
        bx, by = corner_box(bw, bh, box, data_px, [0, 1, 2, 3])
        draw.rectangle([bx, by, bx + bw, by + bh], fill=WHITE, outline=(60, 60, 60))
        for i, s in enumerate(lines):
            txt((bx + bw / 2.0, by + base * 0.5 + i * (lh + 6)), s, f_eq, "ma")
    return img, (slope, intercept, r2, fixed)


# ---------------------------------------------------------------------------- main

def main(argv=None):
    ap = argparse.ArgumentParser(description="Render physics-lab charts to PNG (Pillow only).")
    ap.add_argument("--spec", required=True, help="path to the JSON spec file")
    ap.add_argument("--out", required=True, help="path of the PNG to write")
    args = ap.parse_args(argv)
    try:
        with open(args.spec, "r", encoding="utf-8-sig") as fh:
            spec = json.load(fh)
        cfg = parse_spec(spec)
        img, (slope, intercept, r2, fixed) = render(cfg)
    except FileNotFoundError:
        print(f"error: spec file not found: {args.spec}", file=sys.stderr)
        return 2
    except json.JSONDecodeError as exc:
        print(f"error: invalid JSON in {args.spec}: {exc}", file=sys.stderr)
        return 2
    except SpecError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 3
    except OSError as exc:
        print(f"error: cannot read {args.spec}: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:                        # never write a half-drawn chart
        print(f"error: rendering failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 4
    out = os.path.abspath(args.out)
    try:
        parent = os.path.dirname(out)
        if parent:
            os.makedirs(parent, exist_ok=True)
        img.save(out, format="PNG", dpi=(cfg["dpi"], cfg["dpi"]))
    except OSError as exc:
        print(f"error: cannot write {out}: {exc}", file=sys.stderr)
        return 5
    print(json.dumps({
        "out": out,
        "fit": None if slope is None else {
            "slope": round(float(slope), 6), "intercept": round(float(intercept), 6),
            "r2": round(float(r2), 6), "fixed_at_x0": bool(fixed),
            "x0": float(cfg["fit"]["x0"]) if fixed else None,
            "on": "log10(y)" if cfg["y_log"] else "y"},
        "points": int(sum(g["x"].size for g in cfg["groups"]))}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
