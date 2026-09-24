"""Shared citation numbering and inherited-map resolution."""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import Any, Literal

from .formatters import format_entry
from .model import BibliographyEntry

NumberingPolicy = Literal["first-citation", "source-order"]


def _sort_key(label: str) -> tuple[int, int, str]:
    match = re.fullmatch(r"([A-Za-z]*)(\d+)", label)
    if not match:
        return (2, 0, label)
    prefix, number = match.groups()
    return (0 if not prefix else 1, int(number), prefix)


def _compress(labels: Iterable[str]) -> str:
    unique = sorted(set(labels), key=_sort_key)
    if not unique:
        return ""
    groups: list[list[str]] = [[unique[0]]]
    for label in unique[1:]:
        previous = groups[-1][-1]
        left = re.fullmatch(r"([A-Za-z]*)(\d+)", previous)
        right = re.fullmatch(r"([A-Za-z]*)(\d+)", label)
        if left and right and left.group(1) == right.group(1) and int(right.group(2)) == int(left.group(2)) + 1:
            groups[-1].append(label)
        else:
            groups.append([label])
    parts = [group[0] if len(group) < 3 else f"{group[0]}–{group[-1]}" for group in groups]
    return "[" + ",".join(parts) + "]"


class CitationResolver:
    """Resolve citations independently of Markdown, TeX, or DOCX."""

    def __init__(
        self,
        entries: Mapping[str, BibliographyEntry | Mapping[str, Any] | str],
        *,
        inherited_numbers: Mapping[str, int] | None = None,
        local_prefix: str = "",
        numbering: NumberingPolicy = "first-citation",
        strict: bool = False,
    ) -> None:
        self.entries = {
            key: value if isinstance(value, BibliographyEntry) else BibliographyEntry.from_mapping(key, value)
            for key, value in entries.items()
        }
        self.bib = self.entries
        self.inherited_numbers = dict(inherited_numbers or {})
        self.local_prefix = local_prefix
        self.numbering = numbering
        self.strict = strict
        self.order: list[str] = []
        self._key2num: dict[str, int] = {}

    def validate(self, keys: Iterable[str]) -> None:
        missing = sorted(set(keys) - set(self.entries))
        if missing:
            raise ValueError(f"Unknown citation key(s): {', '.join(missing)}")

    def plan(
        self,
        cited: Iterable[str],
        *,
        inherited_numbers: Mapping[str, int] | None = None,
        local_prefix: str | None = None,
        supplement_prefix: str = "S",
        numbering: NumberingPolicy | None = None,
    ) -> tuple[tuple[str, ...], dict[str, str | int]]:
        keys = list(dict.fromkeys(cited))
        self.validate(keys)
        inherited = dict(self.inherited_numbers if inherited_numbers is None else inherited_numbers)
        policy = numbering or self.numbering
        prefix = self.local_prefix if local_prefix is None else local_prefix
        shared = [key for key in keys if key in inherited]
        local = [key for key in keys if key not in inherited]
        if policy == "source-order":
            source_rank = {key: index for index, key in enumerate(self.entries)}
            local.sort(key=lambda key: source_rank[key])
        mapping: dict[str, str | int] = {key: inherited[key] for key in shared}
        mapping.update({key: f"{prefix or supplement_prefix}{index}" if prefix or inherited else index for index, key in enumerate(local, 1)})
        return tuple(shared + local), mapping

    def _label_for_key(self, key: str, *, use_inherited: bool = True) -> str:
        if self.strict and key not in self.entries:
            raise ValueError(f"Unknown citation key: {key}")
        if use_inherited and key in self.inherited_numbers:
            return str(self.inherited_numbers[key])
        if key not in self._key2num:
            self._key2num[key] = len(self.order) + 1
            self.order.append(key)
        return f"{self.local_prefix}{self._key2num[key]}"

    def labels_for(self, keys: str | Iterable[str], *, use_inherited: bool = True) -> list[str]:
        values = [item.strip() for item in keys.split(",")] if isinstance(keys, str) else list(keys)
        return [self._label_for_key(key, use_inherited=use_inherited) for key in values if key]

    def resolve(self, keys: str | Iterable[str], *, use_inherited: bool = True) -> str:
        return _compress(self.labels_for(keys, use_inherited=use_inherited))

    def citation_parts(self, keys: str | Iterable[str], *, use_inherited: bool = True) -> list[tuple[str, str | None]]:
        labels = self.labels_for(keys, use_inherited=use_inherited)
        parts: list[tuple[str, str | None]] = []
        for index, label in enumerate(labels):
            if index:
                parts.append((",", None))
            parts.append((label, label))
        return parts

    def citation_numbers(self) -> dict[str, int]:
        return dict(self._key2num)

    def reference_items(self, *, profile: str = "plain") -> list[tuple[str, str]]:
        return [
            (f"{self.local_prefix}{self._key2num[key]}", format_entry(self.entries[key], profile=profile))
            for key in self.order
            if key in self._key2num and key in self.entries
        ]

