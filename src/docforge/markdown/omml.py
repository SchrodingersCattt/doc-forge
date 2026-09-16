"""OMML (Office Math Markup Language) builders for inline and display math.

These helpers construct native Word equation objects from the conservative
math subset that appears in Markdown/LaTeX sources: fractions, subscripts,
superscripts, Unicode symbols, and ``\\mathrm``/``\\text``/``\\mathbf``
commands.  They are the OMML counterpart to ``docforge.tex.tokenize`` and
are used by both the Markdown renderer and the TeX converter.
"""

from __future__ import annotations

import re
from lxml import etree

MATH_COMMANDS = {
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\Delta": "Δ",
    r"\epsilon": "ε",
    r"\eta": "η",
    r"\theta": "θ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\Gamma": "Γ",
    r"\Phi": "Φ",
    r"\phi": "φ",
    r"\omega": "ω",
    r"\sum": "∑",
    r"\in": "∈",
    r"\times": "×",
    r"\cdot": "·",
    r"\pm": "±",
    r"\neq": "≠",
    r"\approx": "≈",
    r"\rightarrow": "→",
    r"\longrightarrow": "→",
    r"\AA": "Å",
}

MATH_CONTEXT = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def math_run(text: str, normal: bool = False) -> etree._Element:
    """A single ``m:r`` run; *normal* marks upright (``m:nor``) text."""
    run = etree.Element(f"{{{MATH_CONTEXT}}}r")
    if normal:
        run_properties = etree.Element(f"{{{MATH_CONTEXT}}}rPr")
        run_properties.append(etree.Element(f"{{{MATH_CONTEXT}}}nor"))
        run.append(run_properties)
    node = etree.Element(f"{{{MATH_CONTEXT}}}t")
    node.set(XML_SPACE, "preserve")
    node.text = text
    run.append(node)
    return run


def math_text_elements(text: str, normal: bool = False) -> list:
    """Split *text* into runs, applying italic (variable) vs upright (number/operator) mode."""
    if not text:
        return []
    if normal:
        return [math_run(text, normal=True)]
    elements: list = []
    buffer = ""
    buffer_normal = True
    for char in text:
        char_normal = not char.isalpha()
        if buffer and char_normal != buffer_normal:
            elements.append(math_run(buffer, normal=buffer_normal))
            buffer = ""
        buffer += char
        buffer_normal = char_normal
    if buffer:
        elements.append(math_run(buffer, normal=buffer_normal))
    return elements


def math_arg(tag: str, elements: list) -> etree._Element:
    argument = etree.Element(f"{{{MATH_CONTEXT}}}{tag}")
    if elements:
        for element in elements:
            argument.append(element)
    else:
        argument.append(math_run("", normal=True))
    return argument


def math_fraction(numerator: list, denominator: list) -> etree._Element:
    fraction = etree.Element(f"{{{MATH_CONTEXT}}}f")
    fraction.append(math_arg("num", numerator))
    fraction.append(math_arg("den", denominator))
    return fraction


def math_matrix(rows: list[list[list]]) -> etree._Element:
    """Build an OMML matrix from rows of parsed cell elements."""
    matrix = etree.Element(f"{{{MATH_CONTEXT}}}m")
    for row in rows:
        matrix_row = etree.Element(f"{{{MATH_CONTEXT}}}mr")
        for cell in row:
            matrix_row.append(math_arg("e", cell))
        matrix.append(matrix_row)
    return matrix


def math_parenthesized(element: etree._Element) -> etree._Element:
    """Wrap one math element in scalable parentheses."""
    delimiter = etree.Element(f"{{{MATH_CONTEXT}}}d")
    properties = etree.Element(f"{{{MATH_CONTEXT}}}dPr")
    begin = etree.Element(f"{{{MATH_CONTEXT}}}begChr")
    begin.set(f"{{{MATH_CONTEXT}}}val", "(")
    end = etree.Element(f"{{{MATH_CONTEXT}}}endChr")
    end.set(f"{{{MATH_CONTEXT}}}val", ")")
    properties.extend((begin, end))
    delimiter.append(properties)
    delimiter.append(math_arg("e", [element]))
    return delimiter


def math_script(base, subscript=None, superscript=None) -> etree._Element:
    """Build sSub/sSup/sSubSup around a base run or element."""
    if subscript is not None and superscript is not None:
        script = etree.Element(f"{{{MATH_CONTEXT}}}sSubSup")
        script.append(math_arg("e", [base]))
        script.append(math_arg("sub", subscript))
        script.append(math_arg("sup", superscript))
        return script
    if subscript is not None:
        script = etree.Element(f"{{{MATH_CONTEXT}}}sSub")
        script.append(math_arg("e", [base]))
        script.append(math_arg("sub", subscript))
        return script
    script = etree.Element(f"{{{MATH_CONTEXT}}}sSup")
    script.append(math_arg("e", [base]))
    script.append(math_arg("sup", superscript or []))
    return script


def find_matching_brace(text: str, start: int) -> int:
    """Return the index just past the ``}`` matching the ``{`` at *start*."""
    if start >= len(text) or text[start] != "{":
        raise ValueError("Expected an opening brace")
    depth = 1
    index = start + 1
    while index < len(text) and depth:
        if text[index] == "{":
            depth += 1
        elif text[index] == "}":
            depth -= 1
        index += 1
    if depth:
        raise ValueError("Unclosed brace in math expression")
    return index


def parse_command_arg(text: str, index: int) -> tuple[str | None, int]:
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != "{":
        return None, index
    end = find_matching_brace(text, index)
    return text[index + 1 : end - 1], index + len(text[index + 1 : end - 1]) + 2


def parse_script_arg(text: str, index: int, normal: bool = False) -> tuple[list, int]:
    if index >= len(text):
        return [], index
    if text[index] == "{":
        end = find_matching_brace(text, index)
        return parse_math_omml(text[index + 1 : end - 1], normal=normal), end
    if text[index] == "\\":
        match = re.match(r"\\(?:[A-Za-z]+|.)", text[index:])
        if match:
            token = match.group(0)
            return parse_math_omml(token, normal=normal), index + len(token)
    return parse_math_omml(text[index], normal=normal), index + 1


def parse_math_omml(text: str, normal: bool = False) -> list:
    """Parse the documented math subset into a flat list of OMML elements."""
    elements: list = []
    index = 0
    while index < len(text):
        char = text[index]
        if char.isspace():
            elements.extend(math_text_elements(" ", normal=True))
            index += 1
            continue
        if char == "{":
            end = find_matching_brace(text, index)
            elements.extend(parse_math_omml(text[index + 1 : end - 1], normal=normal))
            index = end
            continue
        if char in "_^" and elements:
            base = elements.pop()
            subscript = None
            superscript = None
            while index < len(text) and text[index] in "_^":
                kind = text[index]
                parsed, index = parse_script_arg(text, index + 1, normal=normal)
                if kind == "_":
                    subscript = parsed
                else:
                    superscript = parsed
            elements.append(math_script(base, subscript=subscript, superscript=superscript))
            continue
        if char == "\\":
            matrix_match = re.match(
                r"\\begin\{pmatrix\}(.*?)\\end\{pmatrix\}",
                text[index:],
                flags=re.DOTALL,
            )
            if matrix_match:
                rows = [
                    [parse_math_omml(cell.strip(), normal=normal) for cell in row.split("&")]
                    for row in re.split(r"\\\\", matrix_match.group(1))
                    if row.strip()
                ]
                elements.append(math_parenthesized(math_matrix(rows)))
                index += matrix_match.end()
                continue
            match = re.match(r"\\(?:[A-Za-z]+|.)", text[index:])
            if not match:
                index += 1
                continue
            command = match.group(0)
            command_end = index + len(command)
            if command in (r"\frac", r"\tfrac", r"\dfrac"):
                numerator, num_end = parse_command_arg(text, command_end)
                denominator, den_end = parse_command_arg(text, num_end)
                if numerator is not None and denominator is not None:
                    elements.append(
                        math_fraction(
                            parse_math_omml(numerator, normal=normal),
                            parse_math_omml(denominator, normal=normal),
                        )
                    )
                    index = den_end
                    continue
            if command in (r"\mathrm", r"\text", r"\textrm"):
                inner, end = parse_command_arg(text, command_end)
                if inner is not None:
                    elements.extend(parse_math_omml(inner, normal=True))
                    index = end
                    continue
            if command == r"\mathbf":
                inner, end = parse_command_arg(text, command_end)
                if inner is not None:
                    elements.extend(parse_math_omml(inner, normal=normal))
                    index = end
                    continue
            if command == r"\mathbb":
                inner, end = parse_command_arg(text, command_end)
                if inner is not None:
                    elements.extend(math_text_elements("ℝ" if inner == "R" else inner, normal=True))
                    index = end
                    continue
            if command in MATH_COMMANDS:
                elements.extend(math_text_elements(MATH_COMMANDS[command], normal=normal))
                index = command_end
                continue
            if command in (r"\,", r"\;", r"\!", r"\ "):
                elements.extend(math_text_elements(" ", normal=True))
                index = command_end
                continue
            elements.extend(math_text_elements(command.lstrip("\\"), normal=True))
            index = command_end
            continue
        elements.extend(math_text_elements(char, normal=normal))
        index += 1
    return elements


def normalize_math_source(text: str) -> str:
    """Strip spacing and grouping commands that OMML does not need."""
    value = text.strip()
    value = re.sub(r"\\left\s*([([|.])", r"\1", value)
    value = re.sub(r"\\right\s*([)\]|.])", r"\1", value)
    value = value.replace(r"\,", " ")
    value = re.sub(r"\\qquad\b", "    ", value)
    value = re.sub(r"\\quad\b", "  ", value)
    return value.strip().rstrip(",")


__all__ = [
    "MATH_COMMANDS",
    "math_run",
    "math_text_elements",
    "math_arg",
    "math_fraction",
    "math_matrix",
    "math_parenthesized",
    "math_script",
    "find_matching_brace",
    "parse_command_arg",
    "parse_script_arg",
    "parse_math_omml",
    "normalize_math_source",
]