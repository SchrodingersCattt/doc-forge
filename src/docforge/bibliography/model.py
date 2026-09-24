"""Backend-neutral bibliography entries."""

from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType
from typing import Any, Mapping


@dataclass(frozen=True)
class BibliographyEntry:
    """A bibliography record that preserves source fields and entry type."""

    key: str
    entry_type: str = "misc"
    fields: Mapping[str, Any] = MappingProxyType({})
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
        normalized = {str(name).lower(): item for name, item in value.items()}
        entry_type = str(normalized.pop("_type", normalized.pop("entry_type", "misc")))
        return cls(key, entry_type, MappingProxyType(normalized), source_format)

    def get(self, name: str, default: Any = "") -> Any:
        return self.fields.get(name.lower(), default)

    def __getitem__(self, name: str) -> Any:
        return self.fields[name.lower()]

