"""StyleProfile v1.

A profile is format-independent. ``template`` is an optional DOCX whose
style definitions are the baseline for the styles the profile names.
``restyle_docx`` copies those definitions, then applies role overrides.
``RoleSpec.select`` names an existing Word style whose paragraphs should
receive this role. A role that matches no paragraph is unused.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Mapping


STYLE_SCHEMA = "docforge.style.v1"
_ALIGNMENTS = frozenset({"left", "center", "right", "justify"})
_MISSING = frozenset({"error", "create"})
_DIRECT = frozenset({"preserve", "clear_conflicting", "clear_all"})
_NATIVE = frozenset({"fail", "warn"})


class StyleProfileError(ValueError):
    """A profile that does not match the v1 schema."""


@dataclass(frozen=True)
class FontSpec:
    latin: str | None = None
    east_asia: str | None = None
    complex: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None

    def explicit(self) -> dict[str, object]:
        return {key: value for key, value in self.to_dict().items() if value is not None}

    def to_dict(self) -> dict[str, object]:
        return {
            "latin": self.latin,
            "east_asia": self.east_asia,
            "complex": self.complex,
            "size_pt": self.size_pt,
            "bold": self.bold,
            "italic": self.italic,
            "color": self.color,
        }


@dataclass(frozen=True)
class ParagraphSpec:
    alignment: str | None = None
    line_spacing: float | None = None
    before_pt: float | None = None
    after_pt: float | None = None
    first_line_pt: float | None = None
    keep_with_next: bool | None = None
    page_break_before: bool | None = None

    def explicit(self) -> dict[str, object]:
        return {key: value for key, value in self.to_dict().items() if value is not None}

    def to_dict(self) -> dict[str, object]:
        return {
            "alignment": self.alignment,
            "line_spacing": self.line_spacing,
            "before_pt": self.before_pt,
            "after_pt": self.after_pt,
            "first_line_pt": self.first_line_pt,
            "keep_with_next": self.keep_with_next,
            "page_break_before": self.page_break_before,
        }


@dataclass(frozen=True)
class RoleSpec:
    word_style: str | None = None
    select: str | None = None
    font: FontSpec = field(default_factory=FontSpec)
    paragraph: ParagraphSpec = field(default_factory=ParagraphSpec)

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {}
        if self.word_style:
            payload["word_style"] = self.word_style
        if self.select:
            payload["select"] = self.select
        font = self.font.explicit()
        paragraph = self.paragraph.explicit()
        if font:
            payload["font"] = font
        if paragraph:
            payload["paragraph"] = paragraph
        return payload


@dataclass(frozen=True)
class PageSpec:
    size: str | None = None
    margins_mm: Mapping[str, float] = field(default_factory=dict)

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {}
        if self.size:
            payload["size"] = self.size
        if self.margins_mm:
            payload["margins_mm"] = dict(self.margins_mm)
        return payload


@dataclass(frozen=True)
class StylePolicy:
    missing_style: str = "error"
    direct_formatting: str = "clear_conflicting"
    unsupported_native_parts: str = "fail"

    def to_dict(self) -> dict[str, str]:
        return {
            "missing_style": self.missing_style,
            "direct_formatting": self.direct_formatting,
            "unsupported_native_parts": self.unsupported_native_parts,
        }


@dataclass(frozen=True)
class StyleProfile:
    roles: Mapping[str, RoleSpec] = field(default_factory=dict)
    page: PageSpec = field(default_factory=PageSpec)
    policy: StylePolicy = field(default_factory=StylePolicy)
    template: str | None = None
    schema: str = STYLE_SCHEMA
    template_base: Path | None = field(default=None, compare=False, repr=False)

    def __post_init__(self) -> None:
        if self.schema != STYLE_SCHEMA:
            raise StyleProfileError(f"unsupported style schema: {self.schema}")
        if self.policy.missing_style not in _MISSING:
            raise StyleProfileError("policy.missing_style must be error or create")
        if self.policy.direct_formatting not in _DIRECT:
            raise StyleProfileError("policy.direct_formatting is not a known policy")
        if self.policy.unsupported_native_parts not in _NATIVE:
            raise StyleProfileError("policy.unsupported_native_parts must be fail or warn")
        for name, role in self.roles.items():
            if not str(name).strip():
                raise StyleProfileError("role name must not be empty")
            alignment = role.paragraph.alignment
            if alignment is not None and alignment not in _ALIGNMENTS:
                raise StyleProfileError(f"role {name} has unknown alignment {alignment}")
            if role.font.color is not None and not _hex_color(role.font.color):
                raise StyleProfileError(f"role {name} color must be a hex RGB value")

    def role(self, name: str) -> RoleSpec:
        try:
            return self.roles[name]
        except KeyError as exc:
            raise KeyError(name) from exc

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {"schema": self.schema, "roles": {name: role.to_dict() for name, role in self.roles.items()}}
        if self.template:
            payload["template"] = self.template
        page = self.page.to_dict()
        if page:
            payload["page"] = page
        payload["policy"] = self.policy.to_dict()
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> StyleProfile:
        unknown = set(value) - {"schema", "template", "roles", "page", "policy"}
        if unknown:
            raise StyleProfileError(f"unknown style profile fields: {sorted(unknown)}")
        roles_raw = value.get("roles") or {}
        if not isinstance(roles_raw, Mapping):
            raise StyleProfileError("roles must be a mapping")
        roles = {str(name): _role(name, spec) for name, spec in roles_raw.items()}
        return cls(
            roles=roles,
            page=_page(value.get("page") or {}),
            policy=_policy(value.get("policy") or {}),
            template=str(value["template"]) if value.get("template") else None,
            schema=str(value.get("schema", STYLE_SCHEMA)),
        )


def load_profile(source: StyleProfile | Path | str | Mapping[str, object]) -> StyleProfile:
    if isinstance(source, StyleProfile):
        return source
    if isinstance(source, Mapping):
        return StyleProfile.from_dict(source)
    path = Path(source)
    text = path.read_text(encoding="utf-8-sig")
    payload = json.loads(text) if text.lstrip().startswith("{") else _yaml(text)
    if not isinstance(payload, dict):
        raise StyleProfileError("style profile must be a mapping")
    # Relative template paths are resolved next to the profile file.
    return replace(StyleProfile.from_dict(payload), template_base=path.expanduser().resolve().parent)


def _role(name: object, spec: object) -> RoleSpec:
    if not isinstance(spec, Mapping):
        raise StyleProfileError(f"role {name} must be a mapping")
    unknown = set(spec) - {"word_style", "select", "font", "paragraph"}
    if unknown:
        raise StyleProfileError(f"role {name} has unknown fields: {sorted(unknown)}")
    return RoleSpec(
        word_style=str(spec["word_style"]) if spec.get("word_style") else None,
        select=str(spec["select"]) if spec.get("select") else None,
        font=_font(spec.get("font") or {}),
        paragraph=_paragraph(spec.get("paragraph") or {}),
    )


def _font(spec: object) -> FontSpec:
    if not isinstance(spec, Mapping):
        raise StyleProfileError("font must be a mapping")
    unknown = set(spec) - {"latin", "east_asia", "complex", "size_pt", "bold", "italic", "color"}
    if unknown:
        raise StyleProfileError(f"font has unknown fields: {sorted(unknown)}")
    return FontSpec(
        latin=_optional_str(spec.get("latin")),
        east_asia=_optional_str(spec.get("east_asia")),
        complex=_optional_str(spec.get("complex")),
        size_pt=_optional_float(spec.get("size_pt")),
        bold=_optional_bool(spec.get("bold")),
        italic=_optional_bool(spec.get("italic")),
        color=_optional_str(spec.get("color")),
    )


def _paragraph(spec: object) -> ParagraphSpec:
    if not isinstance(spec, Mapping):
        raise StyleProfileError("paragraph must be a mapping")
    unknown = set(spec) - {"alignment", "line_spacing", "before_pt", "after_pt", "first_line_pt", "keep_with_next", "page_break_before"}
    if unknown:
        raise StyleProfileError(f"paragraph has unknown fields: {sorted(unknown)}")
    return ParagraphSpec(
        alignment=_optional_str(spec.get("alignment")),
        line_spacing=_optional_float(spec.get("line_spacing")),
        before_pt=_optional_float(spec.get("before_pt")),
        after_pt=_optional_float(spec.get("after_pt")),
        first_line_pt=_optional_float(spec.get("first_line_pt")),
        keep_with_next=_optional_bool(spec.get("keep_with_next")),
        page_break_before=_optional_bool(spec.get("page_break_before")),
    )


def _page(spec: object) -> PageSpec:
    if not spec:
        return PageSpec()
    if not isinstance(spec, Mapping):
        raise StyleProfileError("page must be a mapping")
    unknown = set(spec) - {"size", "margins_mm"}
    if unknown:
        raise StyleProfileError(f"page has unknown fields: {sorted(unknown)}")
    margins = spec.get("margins_mm") or {}
    if not isinstance(margins, Mapping):
        raise StyleProfileError("margins_mm must be a mapping")
    return PageSpec(
        size=_optional_str(spec.get("size")),
        margins_mm={str(key): float(value) for key, value in margins.items()},
    )


def _policy(spec: object) -> StylePolicy:
    if not spec:
        return StylePolicy()
    if not isinstance(spec, Mapping):
        raise StyleProfileError("policy must be a mapping")
    unknown = set(spec) - {"missing_style", "direct_formatting", "unsupported_native_parts"}
    if unknown:
        raise StyleProfileError(f"policy has unknown fields: {sorted(unknown)}")
    current = StylePolicy()
    return StylePolicy(
        missing_style=str(spec.get("missing_style", current.missing_style)),
        direct_formatting=str(spec.get("direct_formatting", current.direct_formatting)),
        unsupported_native_parts=str(spec.get("unsupported_native_parts", current.unsupported_native_parts)),
    )


def _optional_str(value: object) -> str | None:
    if value is None:
        return None
    return str(value)


def _optional_float(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise StyleProfileError(f"expected a number, got {value!r}")
    return float(value)


def _optional_bool(value: object) -> bool | None:
    if value is None or value == "":
        return None
    if not isinstance(value, bool):
        raise StyleProfileError(f"expected a boolean, got {value!r}")
    return value


def _hex_color(value: str) -> bool:
    text = value.removeprefix("#")
    return len(text) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in text)


def _yaml(text: str) -> dict:
    """Parse the StyleProfile subset: nested mappings, scalars, and flow mappings."""

    lines = []
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        lines.append(raw.replace("\t", "    "))
    return _yaml_block(lines, 0)[0]


def _yaml_block(lines: list[str], indent: int) -> tuple[dict, int]:
    mapping: dict = {}
    index = 0
    while index < len(lines):
        line = lines[index]
        current = len(line) - len(line.lstrip(" "))
        if current < indent:
            break
        if current > indent:
            raise StyleProfileError("style YAML indentation jumps inward")
        key, separator, rest = line.strip().partition(":")
        if not separator or not key:
            raise StyleProfileError(f"style YAML entry is not a mapping: {line.strip()}")
        rest = rest.split(" #", 1)[0].strip()
        index += 1
        if rest:
            mapping[key] = _flow(rest) if rest.startswith("{") else _scalar(rest)
            continue
        if index >= len(lines):
            mapping[key] = {}
            continue
        next_indent = len(lines[index]) - len(lines[index].lstrip(" "))
        if next_indent <= indent:
            mapping[key] = {}
            continue
        child, consumed = _yaml_block(lines[index:], next_indent)
        mapping[key] = child
        index += consumed
    return mapping, index


def _flow(text: str) -> dict:
    body = text.strip()[1:-1].strip()
    if not body:
        return {}
    result = {}
    for part in body.split(","):
        key, separator, value = part.partition(":")
        if not separator:
            raise StyleProfileError(f"style YAML flow mapping is invalid: {text}")
        result[key.strip()] = _scalar(value.strip())
    return result


def _scalar(text: str) -> object:
    if text in {"true", "True"}:
        return True
    if text in {"false", "False"}:
        return False
    if text in {"null", "~"}:
        return None
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        return text[1:-1]
    try:
        if "." in text:
            return float(text)
        return int(text)
    except ValueError:
        return text


__all__ = [
    "STYLE_SCHEMA", "FontSpec", "PageSpec", "ParagraphSpec", "RoleSpec", "StylePolicy",
    "StyleProfile", "StyleProfileError", "load_profile",
]
