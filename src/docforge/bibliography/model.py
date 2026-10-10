"""Backend-neutral bibliography entries."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any, Mapping


def _empty_fields() -> Mapping[str, Any]:
    # Python 3.11 treats MappingProxyType as an unhashable default and rejects it.
    return MappingProxyType({})


def _normalise_field(name: str, value: Any) -> Any:
    """Canonicalize fields shared by JSON and BibTeX formatters.

    BibTeX loaders naturally produce strings, while JSON commonly stores a
    publication year as an integer.  Keep the public entry model consistent
    for the common numeric form without rejecting legitimate non-numeric year
    values such as ``in press`` or ``2024a``.
    """
    if name == "year" and isinstance(value, str):
        candidate = value.strip()
        if re.fullmatch(r"\d+", candidate):
            return int(candidate)
    return value


@dataclass(frozen=True)
class BibliographyEntry:
    """A bibliography record that preserves source fields and entry type."""

    key: str
    entry_type: str = "misc"
    fields: Mapping[str, Any] = field(default_factory=_empty_fields)
    source_format: str = "unknown"

    @classmethod
    def from_mapping(
        cls,
        key: str,
        value: Mapping[str, Any] | str,
        *,
        source_format: str = "unknown",
    ) -> "BibliographyEntry":
        if isinstance(value, str):
            if not value.strip():
                raise ValueError(f"Bibliography entry must be non-empty: {key!r}")
            fields = {"raw": value.strip()}
            return cls(key, "misc", MappingProxyType(fields), source_format)
        if not isinstance(value, Mapping):
            raise ValueError(f"Bibliography entry must be a string or object: {key!r}")
        normalized = {
            str(name).lower(): _normalise_field(str(name).lower(), item)
            for name, item in value.items()
        }
        entry_type = str(normalized.pop("_type", normalized.pop("entry_type", "misc")))
        return cls(key, entry_type, MappingProxyType(normalized), source_format)

    def get(self, name: str, default: Any = "") -> Any:
        return self.fields.get(name.lower(), default)

    def __getitem__(self, name: str) -> Any:
        return self.fields[name.lower()]

