"""Rich-text token model and LaTeX parser shared by the TeX converter."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional

GREEK = {
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ε",
    r"\varepsilon": "ε",
    r"\eta": "η",
    r"\theta": "θ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\nu": "ν",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\phi": "φ",
    r"\chi": "χ",
    r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ",
    r"\Delta": "Δ",
    r"\Theta": "Θ",
    r"\Lambda": "Λ",
    r"\Sigma": "Σ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
    r"\times": "×",
    r"\cdot": "·",
    r"\cdots": "⋯",
    r"\ldots": "…",
    r"\pm": "±",
    r"\mp": "∓",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\neq": "≠",
    r"\approx": "≈",
    r"\sim": "∼",
    r"\infty": "∞",
    r"\ell": "ℓ",
    r"\circ": "°",
    r"\deg": "°",
    r"\AA": "Å",
    r"\star": "★",
    r"\ast": "*",
    r"\lvert": "|",
    r"\rvert": "|",
    r"\langle": "⟨",
    r"\rangle": "⟩",
    r"\sum": "∑",
    r"\in": "∈",
    r"\setminus": "∖",
    r"\cup": "∪",
    r"\gcd": "gcd",
    r"\arg": "arg",
    r"\min": "min",
    r"\gets": "←",
    r"\leftarrow": "←",
    r"\longleftarrow": "⟵",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\longrightarrow": "⟶",
}

# Commands whose textual content is skipped entirely (kept for compatibility
# with hand-authored .tex sources).
DISCARD_COMMANDS = (
    r"\State",
    r"\Require",
    r"\Ensure",
    r"\For",
    r"\EndFor",
    r"\While",
    r"\EndWhile",
    r"\If",
    r"\EndIf",
    r"\Return",
    r"\Function",
    r"\EndFunction",
)


@dataclass
class Span:
    """One run of text with formatting flags."""

    text: str
    bold: bool = False
    italic: bool = False
    superscript: bool = False
    subscript: bool = False
    color: str | None = None  # hex like "FF0000"
    highlight: str | None = None  # highlight color name for OOXML e.g. "yellow"


@dataclass
class TableCell:
    """Parsed LaTeX table cell with optional span metadata."""

    text: str
    rowspan: int = 1
    colspan: int = 1


def _find_matching_brace(s: str, start: int) -> int:
    """Return index past the closing `}` that matches the `{` at *start*."""
    assert s[start] == "{"
    depth = 1
    pos = start + 1
    while pos < len(s) and depth > 0:
        if s[pos] == "{":
            depth += 1
        elif s[pos] == "}":
            depth -= 1
        pos += 1
    return pos


def tokenize_tex(raw: str, resolve_ref: Callable[[str], str] | None = None) -> list[Span]:
    """Parse LaTeX-ish text into a flat list of Span objects.

    *resolve_ref* maps ``\\ref{key}``/``\\eqref{key}`` labels to display
    strings; when omitted the label text is kept as-is.
    """
    s = _preprocess(raw, resolve_ref)
    spans: list[Span] = []

    def _emit(text: str, **kw) -> None:
        if text:
            spans.append(Span(text=text, **kw))

    def _parse_math(s: str, bold=False, sup=False, sub=False, color=None, hl=None) -> None:
        """Parse math mode: letters are italic, digits are upright."""
        i = 0
        buf = ""
        buf_italic = False

        def flush() -> None:
            nonlocal buf, buf_italic
            if buf:
                _emit(
                    buf,
                    bold=bold,
                    italic=buf_italic,
                    superscript=sup,
                    subscript=sub,
                    color=color,
                    highlight=hl,
                )
                buf = ""
                buf_italic = False

        while i < len(s):
            ch = s[i]
            if ch == "Δ":
                # Upright capital delta is used as an increment operator.
                if buf and buf_italic:
                    flush()
                buf += ch
                buf_italic = False
                i += 1
            elif ch.isalpha():
                if buf and not buf_italic:
                    flush()
                buf += ch
                buf_italic = True
                i += 1
            elif ch.isdigit():
                if buf and buf_italic:
                    flush()
                buf += ch
                buf_italic = False
                i += 1
            elif ch == "{":
                flush()
                brace_end = _find_matching_brace(s, i)
                inner = s[i + 1 : brace_end - 1]
                _parse_math(inner, bold=bold, sup=sup, sub=sub, color=color, hl=hl)
                i = brace_end
            elif ch == "}":
                i += 1
            elif ch in " \t":
                # TeX ignores ordinary spaces in math mode.
                flush()
                i += 1
            elif ch in "+-=<>()[].,;:!?":
                if buf and buf_italic:
                    flush()
                buf += ch
                buf_italic = False
                i += 1
            elif ch == "\\":
                flush()
                m = re.match(r"\\[a-zA-Z]+\*?", s[i:])
                if m:
                    cmd = m.group(0)
                    cmd_end = i + len(cmd)
                    if cmd in (r"\frac", r"\tfrac", r"\dfrac"):
                        if cmd_end < len(s) and s[cmd_end] == "{":
                            num_end = _find_matching_brace(s, cmd_end)
                            num = s[cmd_end + 1 : num_end - 1]
                            den_start = num_end
                            if den_start < len(s) and s[den_start] == "{":
                                den_end = _find_matching_brace(s, den_start)
                                den = s[den_start + 1 : den_end - 1]
                                _parse_math(num, bold=bold, sup=sup, sub=sub, color=color, hl=hl)
                                _emit(
                                    "/",
                                    bold=bold,
                                    italic=False,
                                    superscript=sup,
                                    subscript=sub,
                                    color=color,
                                    highlight=hl,
                                )
                                _parse_math(den, bold=bold, sup=sup, sub=sub, color=color, hl=hl)
                                i = den_end
                                continue
                    if cmd in GREEK:
                        _emit(
                            GREEK[cmd],
                            bold=bold,
                            italic=False,
                            superscript=sup,
                            subscript=sub,
                            color=color,
                            highlight=hl,
                        )
                        i = cmd_end
                        continue
                    if cmd_end < len(s) and s[cmd_end] == "{":
                        brace_end = _find_matching_brace(s, cmd_end)
                        _parse(
                            s[i:brace_end],
                            bold=bold,
                            italic=True,
                            sup=sup,
                            sub=sub,
                            color=color,
                            hl=hl,
                        )
                        i = brace_end
                    else:
                        _parse(
                            s[i:cmd_end],
                            bold=bold,
                            italic=True,
                            sup=sup,
                            sub=sub,
                            color=color,
                            hl=hl,
                        )
                        i = cmd_end
                else:
                    buf += ch
                    i += 1
            elif ch in "^_":
                flush()
                tag = "sup" if ch == "^" else "sub"
                if i + 1 < len(s) and s[i + 1] == "{":
                    brace_end = _find_matching_brace(s, i + 1)
                    inner = s[i + 2 : brace_end - 1]
                    _parse_math(inner, bold=bold, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                    i = brace_end
                elif i + 1 < len(s):
                    m_cmd = re.match(r"\\(text|mathrm|mathbf|emph|textit|textbf)\{", s[i + 1 :])
                    if m_cmd:
                        cmd_start = i + 1
                        brace_start = cmd_start + m_cmd.end() - 1
                        brace_end = _find_matching_brace(s, brace_start)
                        inner = s[cmd_start:brace_end]
                        _parse(inner, bold=bold, italic=True, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                        i = brace_end
                    else:
                        _parse_math(
                            s[i + 1],
                            bold=bold,
                            sup=tag == "sup",
                            sub=tag == "sub",
                            color=color,
                            hl=hl,
                        )
                        i += 2
                else:
                    i += 1
                continue
            else:
                buf += ch
                i += 1
        flush()

    def _parse(s: str, bold=False, italic=False, sup=False, sub=False, color=None, hl=None) -> None:
        i = 0
        buf = ""

        def flush() -> None:
            nonlocal buf
            if buf:
                _emit(
                    buf,
                    bold=bold,
                    italic=italic,
                    superscript=sup,
                    subscript=sub,
                    color=color,
                    highlight=hl,
                )
                buf = ""

        while i < len(s):
            ch = s[i]
            m_hl = re.match(r"\\(highlight|aigc|gmy|textcolor|colorbox)\{", s[i:])
            if m_hl:
                flush()
                cmd = m_hl.group(1)
                if cmd == "textcolor":
                    brace1_start = i + m_hl.end() - 1
                    brace1_end = _find_matching_brace(s, brace1_start)
                    color_arg = s[brace1_start + 1 : brace1_end - 1].strip()
                    if brace1_end < len(s) and s[brace1_end] == "{":
                        brace2_end = _find_matching_brace(s, brace1_end)
                        inner = s[brace1_end + 1 : brace2_end - 1]
                        cval = _tex_color_to_hex(color_arg)
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=cval, hl=hl)
                        i = brace2_end
                    else:
                        i = brace1_end
                    continue
                elif cmd == "colorbox":
                    brace1_start = i + m_hl.end() - 1
                    brace1_end = _find_matching_brace(s, brace1_start)
                    bg_arg = s[brace1_start + 1 : brace1_end - 1].strip()
                    if brace1_end < len(s) and s[brace1_end] == "{":
                        brace2_end = _find_matching_brace(s, brace1_end)
                        inner = s[brace1_end + 1 : brace2_end - 1]
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=_tex_hl_name(bg_arg))
                        i = brace2_end
                    else:
                        i = brace1_end
                    continue
                else:
                    brace_start = i + m_hl.end() - 1
                    brace_end = _find_matching_brace(s, brace_start)
                    inner = s[brace_start + 1 : brace_end - 1]
                    if cmd == "highlight":
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl="yellow")
                    elif cmd == "aigc":
                        _emit(
                            "[",
                            bold=bold,
                            italic=italic,
                            superscript=sup,
                            subscript=sub,
                            color="FF0000",
                            highlight=hl,
                        )
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color="FF0000", hl=hl)
                        _emit(
                            "]",
                            bold=bold,
                            italic=italic,
                            superscript=sup,
                            subscript=sub,
                            color="FF0000",
                            highlight=hl,
                        )
                    elif cmd == "gmy":
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color="4472C4", hl=hl)
                    i = brace_end
                    continue
            m_comment = re.match(r"\\Comment\{", s[i:])
            if m_comment:
                flush()
                brace_start = i + m_comment.end() - 1
                brace_end = _find_matching_brace(s, brace_start)
                inner = s[brace_start + 1 : brace_end - 1]
                _emit(
                    " // ",
                    bold=bold,
                    italic=True,
                    superscript=sup,
                    subscript=sub,
                    color="808080",
                    highlight=hl,
                )
                _parse(inner, bold=bold, italic=True, sup=sup, sub=sub, color="808080", hl=hl)
                i = brace_end
                continue
            m_cmd = re.match(
                r"\\(emph|textit|textbf|textsc|texttt|textrm|textsf|text|mathrm|mathbf)\{",
                s[i:],
            )
            if m_cmd:
                flush()
                cmd_name = m_cmd.group(1)
                brace_start = i + m_cmd.end() - 1
                brace_end = _find_matching_brace(s, brace_start)
                inner = s[brace_start + 1 : brace_end - 1]
                if cmd_name in ("emph", "textit"):
                    _parse(inner, bold=bold, italic=True, sup=sup, sub=sub, color=color, hl=hl)
                elif cmd_name in ("textbf", "mathbf"):
                    _parse(inner, bold=True, italic=italic, sup=sup, sub=sub, color=color, hl=hl)
                elif cmd_name == "mathrm":
                    _parse(inner, bold=bold, italic=False, sup=sup, sub=sub, color=color, hl=hl)
                else:
                    _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl)
                i = brace_end
                continue
            if ch == "^" or ch == "_":
                flush()
                tag = "sup" if ch == "^" else "sub"
                if i + 1 < len(s) and s[i + 1] == "{":
                    brace_end = _find_matching_brace(s, i + 1)
                    inner = s[i + 2 : brace_end - 1]
                    _parse(inner, bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                    i = brace_end
                elif i + 1 < len(s):
                    m_cmd = re.match(r"\\(text|mathrm|mathbf|emph|textit|textbf)\{", s[i + 1 :])
                    if m_cmd:
                        cmd_start = i + 1
                        brace_start = cmd_start + m_cmd.end() - 1
                        brace_end = _find_matching_brace(s, brace_start)
                        inner = s[cmd_start:brace_end]
                        _parse(inner, bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                        i = brace_end
                    else:
                        _parse(s[i + 1], bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                        i += 2
                else:
                    i += 1
                continue
            if ch == "\\" and i + 1 < len(s) and s[i + 1] == "_":
                flush()
                buf += "_"
                i += 2
                continue
            if ch == "$":
                flush()
                end = s.find("$", i + 1)
                if end == -1:
                    end = len(s)
                inner = s[i + 1 : end]
                _parse_math(inner, bold=bold, sup=sup, sub=sub, color=color, hl=hl)
                i = end + 1
                continue
            if ch == "\\":
                m_unk = re.match(r"\\([a-zA-Z]+)\*?", s[i:])
                if m_unk:
                    if m_unk.group(0) in DISCARD_COMMANDS:
                        i += len(m_unk.group(0))
                        continue
                    flush()
                    cmd_full = m_unk.group(0)
                    after = i + len(cmd_full)
                    if after < len(s) and s[after] == "{":
                        brace_end = _find_matching_brace(s, after)
                        inner = s[after + 1 : brace_end - 1]
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl)
                        i = brace_end
                    else:
                        i = after
                    continue
                buf += ch
                i += 1
                continue
            if ch == "{":
                flush()
                brace_end = _find_matching_brace(s, i)
                inner = s[i + 1 : brace_end - 1]
                _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl)
                i = brace_end
                continue
            if ch == "}":
                i += 1
                continue
            buf += ch
            i += 1
        flush()

    _parse(s)
    merged: list[Span] = []
    for sp in spans:
        if (
            merged
            and merged[-1].bold == sp.bold
            and merged[-1].italic == sp.italic
            and merged[-1].superscript == sp.superscript
            and merged[-1].subscript == sp.subscript
            and merged[-1].color == sp.color
            and merged[-1].highlight == sp.highlight
        ):
            merged[-1].text += sp.text
        else:
            merged.append(sp)
    return merged


def spans_to_plain(spans: list[Span]) -> str:
    return "".join(s.text for s in spans)


def plain_tex(raw: str) -> str:
    """Quick plain-text fallback (for bibliography entries etc.)."""
    return spans_to_plain(tokenize_tex(raw))


def _tex_color_to_hex(name: str) -> str:
    mapping = {
        "red": "FF0000",
        "blue": "0000FF",
        "green": "008000",
        "gray": "808080",
        "black": "000000",
        "white": "FFFFFF",
    }
    return mapping.get(name.strip().lower(), name.strip())


def _tex_hl_name(name: str) -> str:
    mapping = {"yellow": "yellow", "orange": "orange", "gray": "gray", "cyan": "cyan"}
    return mapping.get(name.strip().lower(), "yellow")


def _preprocess(s: str, resolve_ref: Callable[[str], str] | None = None) -> str:
    s = s.replace("~", "\u00A0")
    s = re.sub(r"\\bar\{([^{}]*)\}", lambda m: "".join(ch + "\u0305" for ch in m.group(1)), s)
    s = s.replace("---", "—")
    s = s.replace("--", "–")
    s = s.replace("``", "\u201C")
    s = s.replace("''", "\u201D")
    s = s.replace("`", "\u2018")
    s = re.sub(r"\$([^$]+)\$", lambda m: "$" + m.group(1).replace("-", "–") + "$", s)
    s = re.sub(r"\\vdet\b", lambda _: "V_{\\mathrm{det}}", s)
    s = re.sub(r"\\etasq\b", lambda _: "\\eta^{2}", s)
    s = s.replace("\\%", "%")
    s = s.replace("\\{", "{")
    s = s.replace("\\}", "}")
    s = s.replace("\\ ", " ")
    s = s.replace("\\(", "(")
    s = s.replace("\\)", ")")
    for cmd in DISCARD_COMMANDS:
        s = re.sub(re.escape(cmd) + r"\b\s*", "", s)
    s = re.sub(r"\\protect\s*", "", s)
    s = re.sub(r"\\label\{[^}]*\}", "", s)

    def _resolve_ref(m: re.Match) -> str:
        key = m.group(1)
        return resolve_ref(key) if resolve_ref else key

    s = re.sub(r"\\ref\*?\{([^}]*)\}", _resolve_ref, s)
    s = re.sub(r"\\eqref\*?\{([^}]*)\}", _resolve_ref, s)
    s = re.sub(r"\\href\{[^}]*\}\{([^}]*)\}", r"\1", s)
    s = re.sub(r"\\url\{([^}]*)\}", r"\1", s)
    s = re.sub(r"\\\\\s*(?:\[[\d.]+em\])?", " ", s)
    s = s.replace(r"\,", "\u2009")
    s = s.replace(r"\;", " ")
    s = s.replace(r"\!", "")
    s = re.sub(r"\\quad\b", "  ", s)
    s = re.sub(r"\\qquad\b", "    ", s)
    s = re.sub(r"\\left\s*([(\[|.])", r"\1", s)
    s = re.sub(r"\\right\s*([)\]|.])", r"\1", s)
    s = re.sub(r"\\mathbb\{R\}", "ℝ", s)
    for cmd, uni in sorted(GREEK.items(), key=lambda item: len(item[0]), reverse=True):
        s = s.replace(cmd, uni)
    s = re.sub(r"\\binom\{([^}]*)\}\{([^}]*)\}", r"C(\1,\2)", s)
    return s


__all__ = ["Span", "TableCell", "tokenize_tex", "spans_to_plain", "plain_tex", "GREEK"]