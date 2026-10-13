"""Minimal LaTeX -> OMML (Word 公式) converter.

Only the subset needed by undergraduate physics reports is supported. The
output is a ``m:oMath`` element that can be appended directly to a
``w:p`` paragraph element.

Supported syntax
----------------
* groups ``{...}``
* ``\\frac{a}{b}``, ``\\dfrac``, ``\\tfrac``
* ``\\sqrt{a}``, ``\\sqrt[n]{a}``
* ``^{}``/``^{...}`` and ``_{}``/``_{...}`` (separately or combined)
* ``\\left( ... \\right)`` and friends (auto sized delimiters)
* ``\\overline{}``, ``\\bar{}``, ``\\vec{}``, ``\\hat{}``, ``\\dot{}``
* ``\\text{}``, ``\\mathrm{}``, ``\\rm``, ``\\operatorname{}`` (upright)
* greek letters, operators, arrows and common physics symbols
* ``\\times``, ``\\cdot``, ``\\pm``, ``\\le``, ``\\ge``, ``\\approx`` ...
* spacing: ``\\ `` ``\\,`` ``\\;`` ``\\quad`` ``\\qquad`` ``\\!``
* upright multi-letter function names (``\\lg``, ``\\ln``, ``\\sin`` ...)
"""

from __future__ import annotations

import re

from docx.oxml.ns import qn
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"m": M, "w": W}

# Symbols that map directly onto a Unicode character.
SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ",
    "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ",
    "Omega": "Ω",
    "pm": "±", "mp": "∓", "times": "×", "div": "÷", "cdot": "·",
    "ast": "*", "star": "⋆", "circ": "∘", "bullet": "∙",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "approx": "≈", "equiv": "≡", "sim": "∼", "propto": "∝", "infty": "∞",
    "rightarrow": "→", "to": "→", "leftarrow": "←", "leftrightarrow": "↔",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
    "uparrow": "↑", "downarrow": "↓",
    "sum": "∑", "prod": "∏", "int": "∫", "oint": "∮", "partial": "∂",
    "nabla": "∇", "perp": "⊥", "parallel": "∥", "angle": "∠", "degree": "°",
    "prime": "′", "ldots": "…", "cdots": "⋯", "dots": "…", "vdots": "⋮",
    "in": "∈", "notin": "∉", "subset": "⊂", "supset": "⊃", "cup": "∪",
    "cap": "∩", "emptyset": "∅", "forall": "∀", "exists": "∃",
    "because": "∵", "therefore": "∴", "ll": "≪", "gg": "≫",
    "langle": "⟨", "rangle": "⟩", "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ",
    "Im": "ℑ", "aleph": "ℵ", "surd": "√", "neg": "¬", "land": "∧",
    "lor": "∨", "oplus": "⊕", "otimes": "⊗", "dagger": "†",
    "quad": "\u2003", "qquad": "\u2003\u2003", " ": " ", ",": "\u2009",
    ";": "\u2005", "!": "", ":": "\u2005",
}

# 这些符号 Cambria Math 在部分阅读器里缺字形（会显示成 ¿），改用正文字体
TEXT_FONT_SYMBOLS = {"%", "‰", "℃", "°"}

# Upright function / operator names.
UPRIGHT_NAMES = {
    "lg", "ln", "log", "exp", "sin", "cos", "tan", "cot", "sec", "csc",
    "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "lim", "max",
    "min", "sup", "inf", "det", "dim", "deg", "gcd", "mod", "bmod",
}

DELIMS = {"(": ")", "[": "]", "\\{": "\\}", "\\lbrace": "\\rbrace",
          "|": "|", "\\|": "\\|", "\\langle": "\\rangle",
          ".": "."}

_ACCENTS = {
    "overline": "\u0305", "bar": "\u0304", "vec": "\u20d7",
    "hat": "\u0302", "dot": "\u0307", "ddot": "\u0308",
    "tilde": "\u0303", "widehat": "\u0302",
}


class LatexError(ValueError):
    """Raised when the LaTeX subset parser meets unsupported syntax."""


# --------------------------------------------------------------------------
# OMML element helpers
# --------------------------------------------------------------------------
def _new(tag: str, **attrs):
    el = etree.Element("{%s}%s" % (M, tag), nsmap=NSMAP)
    for key, value in attrs.items():
        el.set("{%s}%s" % (M, key), value)
    return el


def _run(text: str, upright: bool = False, italic: bool | None = None,
         font: str = "Cambria Math"):
    """Create an ``m:r`` math run."""
    run = _new("r")
    if upright or italic is not None:
        rpr = _new("rPr")
        if upright:
            rpr.append(_new("sty", val="p"))
        elif italic is False:
            rpr.append(_new("sty", val="p"))
        run.append(rpr)
    wrpr = etree.SubElement(run, "{%s}rPr" % W)
    fonts = etree.SubElement(wrpr, "{%s}rFonts" % W)
    fonts.set("{%s}ascii" % W, font)
    fonts.set("{%s}hAnsi" % W, font)
    fonts.set("{%s}eastAsia" % W, font)
    text_el = etree.SubElement(run, "{%s}t" % M)
    text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_el.text = text
    return run


def _wrap(tag: str, children):
    el = _new(tag)
    for child in children:
        el.append(child)
    return el


class _Node:
    """A parsed math fragment: a list of OMML elements."""

    __slots__ = ("elements",)

    def __init__(self, elements=None):
        self.elements = list(elements or [])

    def text(self) -> str:
        out = []
        for el in self.elements:
            for t in el.iter("{%s}t" % M):
                out.append(t.text or "")
        return "".join(out)

    def is_single_run(self):
        return len(self.elements) == 1 and self.elements[0].tag == "{%s}r" % M


# --------------------------------------------------------------------------
# Tokenizer / parser
# --------------------------------------------------------------------------
_TOKEN_RE = re.compile(
    r"""
      (?P<cmd>\\[A-Za-z]+|\\.)          # backslash command or escaped char
    | (?P<group>\{)                     # group open
    | (?P<groupend>\})
    | (?P<sup>\^)
    | (?P<sub>_)
    | (?P<space>\s+)
    | (?P<char>[^\\{}^_\s])
    """,
    re.VERBOSE | re.DOTALL,
)


class _Parser:
    def __init__(self, source: str):
        self.tokens = list(_TOKEN_RE.finditer(source))
        self.pos = 0

    # -- token helpers ---------------------------------------------------
    def _peek(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def _next(self):
        tok = self._peek()
        if tok is not None:
            self.pos += 1
        return tok

    # -- grammar ---------------------------------------------------------
    def parse(self, stop_at_group_end: bool = False) -> _Node:
        node = _Node()
        while True:
            tok = self._peek()
            if tok is None:
                break
            kind = tok.lastgroup
            if kind == "groupend":
                if stop_at_group_end:
                    self._next()
                    break
                raise LatexError("unbalanced '}' in %r" % self._source())
            if kind == "group":
                self._next()
                inner = self.parse(stop_at_group_end=True)
                node.elements.extend(inner.elements)
                continue
            if kind == "space":
                self._next()
                continue
            if kind == "sup":
                self._next()
                node = self._attach_script(node, "sup")
                continue
            if kind == "sub":
                self._next()
                node = self._attach_script(node, "sub")
                continue
            if kind == "cmd":
                self._next()
                self._apply_command(tok.group("cmd"), node)
                continue
            self._next()
            node.elements.append(_run(tok.group("char")))
        return node

    def _source(self):
        return "".join(t.group(0) for t in self.tokens)

    # -- pieces ----------------------------------------------------------
    def _one_token_node(self) -> _Node:
        """Parse exactly one atom (used for ^ and _ arguments)."""
        tok = self._peek()
        if tok is None:
            raise LatexError("missing script argument")
        if tok.lastgroup == "group":
            self._next()
            return self.parse(stop_at_group_end=True)
        if tok.lastgroup == "cmd":
            self._next()
            node = _Node()
            self._apply_command(tok.group("cmd"), node)
            return node
        self._next()
        return _Node([_run(tok.group("char"))])

    def _attach_script(self, base: _Node, kind: str) -> _Node:
        script = self._one_token_node()
        if not base.elements:
            # A script with no base, e.g. a leading ^2 (rare) - keep the script.
            base = _Node([_run("")])
        if kind == "sub":
            # base_{sub} possibly followed by ^{sup}
            nxt = self._peek()
            if nxt is not None and nxt.lastgroup == "sup":
                self._next()
                sup = self._one_token_node()
                el = _wrap("sSubSup", [
                    _wrap("e", base.elements),
                    _wrap("sub", script.elements),
                    _wrap("sup", sup.elements),
                ])
                return _Node([el])
            el = _wrap("sSub", [_wrap("e", base.elements), _wrap("sub", script.elements)])
            return _Node([el])
        # sup, possibly followed by _{sub}
        nxt = self._peek()
        if nxt is not None and nxt.lastgroup == "sub":
            self._next()
            sub = self._one_token_node()
            el = _wrap("sSubSup", [
                _wrap("e", base.elements),
                _wrap("sub", sub.elements),
                _wrap("sup", script.elements),
            ])
            return _Node([el])
        el = _wrap("sSup", [_wrap("e", base.elements), _wrap("sup", script.elements)])
        return _Node([el])

    def _command_arg(self) -> _Node:
        return self._one_token_node()

    def _apply_command(self, cmd: str, node: _Node) -> None:
        name = cmd[1:]

        if name in ("frac", "dfrac", "tfrac"):
            num = self._command_arg()
            den = self._command_arg()
            node.elements.append(_wrap("f", [
                _wrap("num", num.elements), _wrap("den", den.elements)]))
            return

        if name == "sqrt":
            degree = None
            tok = self._peek()
            if tok is not None and tok.lastgroup == "char" and tok.group("char") == "[":
                # optional degree: read raw until matching ]
                self._next()
                raw = []
                depth = 1
                while True:
                    tk = self._next()
                    if tk is None:
                        raise LatexError("unterminated \\sqrt[...]")
                    if tk.lastgroup == "char" and tk.group("char") == "[":
                        depth += 1
                    if tk.lastgroup == "char" and tk.group("char") == "]":
                        depth -= 1
                        if depth == 0:
                            break
                    raw.append(tk.group(0))
                degree = _Parser("".join(raw)).parse()
            radicand = self._command_arg()
            parts = [_wrap("radPr", [])]
            if degree is not None:
                parts.append(_wrap("deg", degree.elements))
            else:
                parts.append(_new("deg"))
            parts.append(_wrap("e", radicand.elements))
            node.elements.append(_wrap("rad", parts))
            return

        if name in ("text", "textrm", "mathrm", "operatorname", "mbox"):
            arg = self._command_arg()
            node.elements.append(_run(arg.text(), upright=True))
            return

        if name == "rm":
            return  # style switch, applies to nothing in our subset

        if name in _ACCENTS:
            arg = self._command_arg()
            accent = _new("acc", **{"chr": _ACCENTS[name]})
            accent.append(_wrap("e", arg.elements))
            node.elements.append(accent)
            return

        if name in ("left", "right", "big", "Big", "bigg", "Bigg", "middle"):
            if name == "left":
                open_tok = self._next()
                if open_tok is None:
                    raise LatexError("dangling \\left")
                beg = open_tok.group(0)
                inner = self._read_until_right()
                end = inner.pop("end")
                dpr = _new("dPr")
                b = _new("begChr")
                # 起始符号就是 \left 后面那个字符；DELIMS 只用来查它配对的右符号
                b.set("{%s}val" % M, beg)
                e = _new("endChr")
                e.set("{%s}val" % M, DELIMS.get(beg, end))
                dpr.append(b)
                dpr.append(e)
                node.elements.append(_wrap("d", [dpr, _wrap("e", inner["elements"])]))
                return
            return  # \right handled by _read_until_right

        if name in SYMBOLS or name in TEXT_FONT_SYMBOLS:
            value = SYMBOLS.get(name, name)
            if value:
                if name in TEXT_FONT_SYMBOLS:
                    node.elements.append(_run(value, upright=True,
                                              font="Times New Roman"))
                else:
                    node.elements.append(
                        _run(value, upright=name in ("quad", "qquad")))
            return

        if name in UPRIGHT_NAMES:
            node.elements.append(_run(name, upright=True))
            return

        # Unknown command: emit it verbatim rather than failing the report.
        node.elements.append(_run("\\" + name, upright=True))

    def _read_until_right(self) -> dict:
        elements = []
        while True:
            tok = self._peek()
            if tok is None:
                raise LatexError("missing \\right")
            if tok.lastgroup == "cmd" and tok.group("cmd") == "\\right":
                self._next()
                end_tok = self._next()
                end = end_tok.group(0) if end_tok is not None else "."
                return {"elements": elements, "end": end}
            if tok.lastgroup == "group":
                self._next()
                sub = self.parse(stop_at_group_end=True)
                elements.extend(sub.elements)
                continue
            if tok.lastgroup == "space":
                self._next()
                continue
            if tok.lastgroup == "sup":
                self._next()
                cur = _Node(elements)
                elements = self._attach_script(cur, "sup").elements
                continue
            if tok.lastgroup == "sub":
                self._next()
                cur = _Node(elements)
                elements = self._attach_script(cur, "sub").elements
                continue
            if tok.lastgroup == "cmd":
                self._next()
                sub = _Node()
                self._apply_command(tok.group("cmd"), sub)
                elements.extend(sub.elements)
                continue
            self._next()
            elements.append(_run(tok.group("char")))


def latex_to_omml(latex: str):
    """Convert a LaTeX string into an ``m:oMath`` lxml element."""
    source = latex.strip()
    # \tag{1} and trailing "(1)" markers belong to the layout, not the maths.
    source = re.sub(r"\\tag\{[^}]*\}", "", source)
    parser = _Parser(source)
    node = parser.parse()
    math = _new("oMath")
    for el in node.elements:
        math.append(el)
    return math


def _selftest():  # pragma: no cover - manual smoke test
    samples = [
        r"h\nu = E_2 - E_1",
        r"\lambda = \frac{h}{p}",
        r"V_1 = \frac{1}{6}\left(\frac{U_7-U_1}{3}+\frac{U_8-U_2}{3}\right)",
        r"eU = \frac{1}{2}mv^2",
        r"\lg I_a = \lg I + \frac{e}{2.303kT}\sqrt{\frac{E_a}{4\pi\varepsilon_0}}",
        r"\overline{U} = 4.85 \pm 0.06\ \mathrm{V}",
    ]
    for s in samples:
        el = latex_to_omml(s)
        text = "".join(t.text or "" for t in el.iter("{%s}t" % M))
        print("%-70s -> %s" % (s, text))


if __name__ == "__main__":  # pragma: no cover
    _selftest()
