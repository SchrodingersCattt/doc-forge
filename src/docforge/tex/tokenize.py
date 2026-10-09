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
    r"\kappa": "κ",
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
    r"\Pi": "Π",
    r"\Sigma": "Σ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
    r"\Xi": "Ξ",
    r"\times": "×",
    r"\otimes": "⊗",
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
    r"\max": "max",
    r"\min": "min",
    r"\gets": "←",
    r"\leftarrow": "←",
    r"\longleftarrow": "⟵",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\longrightarrow": "⟶",
}

# Relations take a space on each side. TeX inserts that space and then
# discards the author's spaces, so a converter that drops math spaces
# glues "a = b" into "a=b".
SPACED_RELATIONS = {
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\neq": "≠",
    r"\approx": "≈",
    r"\sim": "∼",
}
RELATION_CHARS = "=<>≤≥≠≈∼"

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
    mono: bool = False  # \texttt; rendered as Consolas in DOCX


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


def _replace_display_argument(
    text: str,
    command: str,
    *,
    argument_count: int = 2,
    optional_prefix: bool = False,
    keep_index: int = -1,
) -> str:
    r"""Replace a metadata-bearing macro with its displayed argument.

    A regular expression such as ``\\href\{[^}]*\}\{([^}]*)\}`` only works
    for flat arguments.  TeX arguments commonly contain another command (for
    example ``\\href{url}{\\textbf{link}}``), so use the same balanced-brace
    reader as the main tokenizer.  ``optional_prefix`` handles commands such
    as ``\\hyperref[label]{visible text}``, whose optional label is metadata.
    Commands with no complete argument list are left untouched for the normal
    parser to handle.
    """
    pattern = re.compile(r"\\" + re.escape(command) + r"(?![A-Za-z])")
    cursor = 0
    pieces: list[str] = []
    while True:
        match = pattern.search(text, cursor)
        if match is None:
            pieces.append(text[cursor:])
            break
        pos = match.end()
        had_optional = False
        if optional_prefix:
            while pos < len(text) and text[pos] in " \t\r\n":
                pos += 1
            if pos < len(text) and text[pos] == "[":
                end = text.find("]", pos + 1)
                if end < 0:
                    pieces.append(text[cursor:])
                    break
                pos = end + 1
                had_optional = True
        args: list[str] = []
        probe = pos
        valid = True
        expected_args = 1 if had_optional else argument_count
        for _ in range(expected_args):
            while probe < len(text) and text[probe] in " \t\r\n":
                probe += 1
            if probe >= len(text) or text[probe] != "{":
                valid = False
                break
            end = _find_matching_brace(text, probe)
            if end <= probe or end > len(text) or text[end - 1] != "}":
                valid = False
                break
            args.append(text[probe + 1 : end - 1])
            probe = end
        if not valid:
            # Do not consume a malformed invocation.  Continue searching after
            # the command so a later valid invocation can still be normalized.
            pieces.append(text[cursor : match.end()])
            cursor = match.end()
            continue
        pieces.append(text[cursor : match.start()])
        pieces.append(args[keep_index])
        cursor = probe
    return "".join(pieces)


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
            elif ch in RELATION_CHARS:
                flush()
                _emit(
                    f" {ch} ",
                    bold=bold,
                    italic=False,
                    superscript=sup,
                    subscript=sub,
                    color=color,
                    highlight=hl,
                )
                i += 1
            elif ch in "+-()[].,;:!?":
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
                    if cmd == r"\highlight" and cmd_end < len(s) and s[cmd_end] == "{":
                        brace_end = _find_matching_brace(s, cmd_end)
                        _parse_math(
                            s[cmd_end + 1 : brace_end - 1],
                            bold=bold,
                            sup=sup,
                            sub=sub,
                            color=color,
                            hl="yellow",
                        )
                        i = brace_end
                        continue
                    if cmd in SPACED_RELATIONS:
                        _emit(
                            f" {SPACED_RELATIONS[cmd]} ",
                            bold=bold,
                            italic=False,
                            superscript=sup,
                            subscript=sub,
                            color=color,
                            highlight=hl,
                        )
                        i = cmd_end
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
                    if cmd == r"\boldsymbol" and cmd_end < len(s):
                        if s[cmd_end] == "{":
                            brace_end = _find_matching_brace(s, cmd_end)
                            _parse_math(s[cmd_end + 1 : brace_end - 1], bold=True, sup=sup, sub=sub, color=color, hl=hl)
                            i = brace_end
                        else:
                            following = re.match(r"\\[a-zA-Z]+\*?", s[cmd_end:]) if s[cmd_end] == "\\" else None
                            end = cmd_end + (len(following.group(0)) if following else 1)
                            _parse_math(s[cmd_end:end], bold=True, sup=sup, sub=sub, color=color, hl=hl)
                            i = end
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
                    elif s[i + 1] == "\\" and (macro := re.match(r"\\[a-zA-Z]+\*?", s[i + 1 :])):
                        _parse_math(macro.group(0), bold=bold, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl)
                        i += 1 + len(macro.group(0))
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

    def _parse(s: str, bold=False, italic=False, sup=False, sub=False, color=None, hl=None, mono=False) -> None:
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
                    mono=mono,
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
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=cval, hl=hl, mono=mono)
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
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=_tex_hl_name(bg_arg), mono=mono)
                        i = brace2_end
                    else:
                        i = brace1_end
                    continue
                else:
                    brace_start = i + m_hl.end() - 1
                    brace_end = _find_matching_brace(s, brace_start)
                    inner = s[brace_start + 1 : brace_end - 1]
                    if cmd == "highlight":
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl="yellow", mono=mono)
                    elif cmd == "aigc":
                        _emit(
                            "[",
                            bold=bold,
                            italic=italic,
                            superscript=sup,
                            subscript=sub,
                            color="FF0000",
                            highlight=hl,
                            mono=mono,
                        )
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color="FF0000", hl=hl, mono=mono)
                        _emit(
                            "]",
                            bold=bold,
                            italic=italic,
                            superscript=sup,
                            subscript=sub,
                            color="FF0000",
                            highlight=hl,
                            mono=mono,
                        )
                    elif cmd == "gmy":
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color="4472C4", hl=hl, mono=mono)
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
                _parse(inner, bold=bold, italic=True, sup=sup, sub=sub, color="808080", hl=hl, mono=mono)
                i = brace_end
                continue
            m_cmd = re.match(
                r"\\(textsuperscript|emph|textit|textbf|textsc|texttt|textrm|textsf|text|mathrm|mathbf)\{",
                s[i:],
            )
            if m_cmd:
                flush()
                cmd_name = m_cmd.group(1)
                brace_start = i + m_cmd.end() - 1
                brace_end = _find_matching_brace(s, brace_start)
                inner = s[brace_start + 1 : brace_end - 1]
                if cmd_name == "textsuperscript":
                    _parse(inner, bold=bold, italic=italic, sup=True, sub=sub, color=color, hl=hl, mono=mono)
                elif cmd_name in ("emph", "textit"):
                    _parse(inner, bold=bold, italic=True, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                elif cmd_name in ("textbf", "mathbf"):
                    _parse(inner, bold=True, italic=italic, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                elif cmd_name == "mathrm":
                    _parse(inner, bold=bold, italic=False, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                elif cmd_name == "texttt":
                    _parse(inner, bold=bold, italic=False, sup=sup, sub=sub, color=color, hl=hl, mono=True)
                else:
                    _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                i = brace_end
                continue
            if ch == "^" or ch == "_":
                flush()
                tag = "sup" if ch == "^" else "sub"
                if i + 1 < len(s) and s[i + 1] == "{":
                    brace_end = _find_matching_brace(s, i + 1)
                    inner = s[i + 2 : brace_end - 1]
                    _parse(inner, bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl, mono=mono)
                    i = brace_end
                elif i + 1 < len(s):
                    m_cmd = re.match(r"\\(text|mathrm|mathbf|emph|textit|textbf)\{", s[i + 1 :])
                    if m_cmd:
                        cmd_start = i + 1
                        brace_start = cmd_start + m_cmd.end() - 1
                        brace_end = _find_matching_brace(s, brace_start)
                        inner = s[cmd_start:brace_end]
                        _parse(inner, bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl, mono=mono)
                        i = brace_end
                    else:
                        _parse(s[i + 1], bold=bold, italic=italic, sup=tag == "sup", sub=tag == "sub", color=color, hl=hl, mono=mono)
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
                    cmd_name = m_unk.group(1)
                    cmd_full = m_unk.group(0)
                    after = i + len(cmd_full)
                    if cmd_name in {"vspace", "hspace"} and after < len(s) and s[after] == "{":
                        i = _find_matching_brace(s, after)
                        continue
                    if cmd_name in {"begin", "end"} and after < len(s) and s[after] == "{":
                        brace_end = _find_matching_brace(s, after)
                        env_name = s[after + 1 : brace_end - 1].strip()
                        if env_name in {"center", "minipage", "flushleft", "flushright"}:
                            i = brace_end
                            if (
                                cmd_name == "begin"
                                and env_name == "minipage"
                                and i < len(s)
                                and s[i] == "{"
                            ):
                                i = _find_matching_brace(s, i)
                            continue
                    flush()
                    cmd_full = m_unk.group(0)
                    after = i + len(cmd_full)
                    if after < len(s) and s[after] == "{":
                        brace_end = _find_matching_brace(s, after)
                        inner = s[after + 1 : brace_end - 1]
                        _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                        i = brace_end
                        # hyperref prints only the typeset argument; the second
                        # brace is the PDF-bookmark fallback and must not appear.
                        if cmd_full == r"\texorpdfstring":
                            while i < len(s) and s[i] in " \t":
                                i += 1
                            if i < len(s) and s[i] == "{":
                                i = _find_matching_brace(s, i)
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
                _parse(inner, bold=bold, italic=italic, sup=sup, sub=sub, color=color, hl=hl, mono=mono)
                i = brace_end
                continue
            if ch == "}":
                i += 1
                continue
            buf += ch
            i += 1
        flush()

    _parse(s)
    spans = _tighten_relation_spaces(spans)
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
            and merged[-1].mono == sp.mono
        ):
            merged[-1].text += sp.text
        else:
            merged.append(sp)
    return merged


def _tighten_relation_spaces(spans: list[Span]) -> list[Span]:
    """Drop a relation's padding space when the neighboring run already has one.

    A relation that is its own math group, as in ``X $>$ A``, sits between
    prose spaces. Padding it again would print a double space. A relation
    inside a larger expression has no neighboring spaces, because math mode
    discarded them, and keeps both padding spaces.
    """
    symbols = set(SPACED_RELATIONS.values()) | set(RELATION_CHARS)
    for index, span in enumerate(spans):
        core = span.text.strip()
        if span.text != f" {core} " or core not in symbols:
            continue
        left = "" if index == 0 or spans[index - 1].text.endswith((" ", "\n")) else " "
        right = "" if index + 1 == len(spans) or spans[index + 1].text.startswith((" ", "\n")) else " "
        span.text = f"{left}{core}{right}"
    return spans


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


_MATH_SPAN_RE = re.compile(
    r"\$\$[\s\S]*?\$\$"
    r"|\$[^$]*\$"
    r"|\\\[[\s\S]*?\\\]"
    r"|\\begin\{(?:equation|align|eqnarray|gather|multline)\*?\}[\s\S]*?"
    r"\\end\{(?:equation|align|eqnarray|gather|multline)\*?\}"
)


def _replace_prose_apostrophes(s: str) -> str:
    """Turn prose apostrophes into closing single quotes, leaving math primes."""
    pieces: list[str] = []
    cursor = 0
    for match in _MATH_SPAN_RE.finditer(s):
        pieces.append(s[cursor:match.start()].replace("'", "\u2019"))
        pieces.append(match.group(0))
        cursor = match.end()
    pieces.append(s[cursor:].replace("'", "\u2019"))
    return "".join(pieces)


def unpaired_quote_errors(text: str) -> list[str]:
    """Return quote-pairing failures in reader-facing text.

    Opening quotes must close with the matching curly mark. A straight
    apostrophe or straight double quote may not close an open quotation.
    A closing single quote with no open quote is an apostrophe.
    """
    closers = {"\u2019": "\u2018", "\u201d": "\u201c"}
    straight = {"'": "\u2018", '"': "\u201c"}
    names = {"\u2018": "single", "\u201c": "double"}
    stack: list[str] = []
    errors: list[str] = []
    for index, char in enumerate(text):
        if char in closers:
            expected = closers[char]
            if stack and stack[-1] == expected:
                stack.pop()
            elif char == "\u201d":
                errors.append(f"closing double quote without an opening quote at offset {index}")
        elif char in ("\u2018", "\u201c"):
            stack.append(char)
        elif char in straight:
            if stack and stack[-1] == straight[char]:
                snippet = text[max(0, index - 20):index + 12].replace("\n", " ")
                errors.append(
                    f"straight mark closes a {names[straight[char]]} quote near {snippet!r}"
                )
                stack.pop()
            elif char == '"':
                snippet = text[max(0, index - 20):index + 12].replace("\n", " ")
                errors.append(f"straight double quote near {snippet!r}")
    if stack:
        opened = ", ".join(names[char] for char in stack)
        errors.append(f"unclosed {opened} quote(s)")
    return errors


def _collapse_consecutive_number_lists(text: str) -> str:
    """Collapse ``S9, S10, S11 and S12`` to ``S9–S12`` after references resolve.

    A run needs three or more integers that share a prefix and increase by one.
    Decimal values, sample identifiers such as ``S1-2``, and panel letters stay
    as written.
    """
    item = re.compile(r"(?<![\w.\-])([A-Za-z]*)(\d+)(?![\w.\-])")
    separator = re.compile(r"(?:\s*,\s*|\s+and\s+)", re.IGNORECASE)
    matches = list(item.finditer(text))
    if len(matches) < 3:
        return text
    pieces: list[str] = []
    cursor = 0
    index = 0
    while index < len(matches):
        run = [matches[index]]
        probe = index
        while probe + 1 < len(matches):
            gap = text[run[-1].end():matches[probe + 1].start()]
            if separator.fullmatch(gap) is None:
                break
            previous, current = run[-1], matches[probe + 1]
            same_prefix = previous.group(1) == current.group(1)
            consecutive = int(current.group(2)) == int(previous.group(2)) + 1
            if not (same_prefix and consecutive):
                break
            run.append(current)
            probe += 1
        if len(run) >= 3:
            pieces.append(text[cursor:run[0].start()])
            prefix = run[0].group(1)
            pieces.append(f"{prefix}{run[0].group(2)}–{prefix}{run[-1].group(2)}")
            cursor = run[-1].end()
            index = probe + 1
            continue
        index += 1
    pieces.append(text[cursor:])
    return "".join(pieces)


def _preprocess(s: str, resolve_ref: Callable[[str], str] | None = None) -> str:
    s = s.replace("~", "\u00A0")
    s = re.sub(r"\\bar\{\\mathbf\s+([A-Za-z])\}", lambda m: r"\mathbf{" + m.group(1) + "\u0305}", s)
    s = re.sub(r"\\bar\{([^{}]*)\}", lambda m: "".join(ch + "\u0305" for ch in m.group(1)), s)
    s = s.replace("---", "—")
    s = s.replace("--", "–")
    s = s.replace("``", "\u201C")
    s = s.replace("''", "\u201D")
    s = s.replace("`", "\u2018")
    s = _replace_prose_apostrophes(s)
    s = re.sub(r"\$([^$]+)\$", lambda m: "$" + m.group(1).replace("-", "–") + "$", s)
    s = re.sub(r"\\vdet\b", lambda _: "V_{\\mathrm{det}}", s)
    s = re.sub(r"\\etasq\b", lambda _: "\\eta^{2}", s)
    s = s.replace("\\%", "%")
    s = s.replace("\\#", "#")
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
    s = _collapse_consecutive_number_lists(s)
    # Link targets and bookmark labels are metadata.  Keep the displayed
    # argument even when it contains nested TeX commands; flat regular
    # expressions would leave the target (or fallback text) in the output.
    s = _replace_display_argument(s, "href")
    s = _replace_display_argument(s, "hyperlink")
    s = _replace_display_argument(s, "hyperref", optional_prefix=True)
    # ``texorpdfstring`` receives the typeset form first and a PDF bookmark
    # fallback second.  DOCX has no bookmark-text rendering mode, so retain
    # only the first argument (including nested formatting commands).
    s = _replace_display_argument(s, "texorpdfstring", keep_index=0)
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


__all__ = [
    "Span", "TableCell", "tokenize_tex", "spans_to_plain", "plain_tex",
    "unpaired_quote_errors", "GREEK",
]
