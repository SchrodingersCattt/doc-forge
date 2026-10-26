"""Typed SemanticDocument v1.

The document tree stores meaning and structure. It is not a lossless copy of
OOXML, Markdown source layout, or a TeX preamble. Adapters own parsing and
rendering; this module imports neither Pandoc nor python-docx.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from typing import Iterator, Mapping


SCHEMA = "docforge.semantic.v1"

BLOCK_KINDS = frozenset({
    "paragraph", "heading", "list", "list_item", "table", "table_row", "table_cell",
    "figure", "caption", "math", "code", "quote", "thematic_break",
    "bibliography", "reference_entry", "opaque",
})
INLINE_KINDS = frozenset({
    "text", "emph", "strong", "code", "link", "math", "break",
    "subscript", "superscript", "underline", "strike", "span", "citation", "opaque",
})
_ABS_PATH = re.compile(r"^(?:[A-Za-z]:[\\/]|/)")


def _pairs(items: Mapping[str, str] | tuple[tuple[str, str], ...] | None) -> tuple[tuple[str, str], ...]:
    if not items:
        return ()
    if isinstance(items, Mapping):
        pairs = ((str(key), str(value)) for key, value in items.items())
    else:
        pairs = ((str(key), str(value)) for key, value in items)
    return tuple(sorted(pairs))


def _attr(attrs: tuple[tuple[str, str], ...], key: str, default: str = "") -> str:
    for name, value in attrs:
        if name == key:
            return value
    return default


@dataclass(frozen=True)
class Provenance:
    """Portable origin of a node. Paths are file names, never absolute paths."""

    source: str = ""
    anchor: str = ""
    start: int | None = None
    end: int | None = None

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {}
        if self.source:
            payload["source"] = self.source
        if self.anchor:
            payload["anchor"] = self.anchor
        if self.start is not None:
            payload["start"] = self.start
        if self.end is not None:
            payload["end"] = self.end
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object] | None) -> Provenance | None:
        if not value:
            return None
        return cls(
            source=str(value.get("source", "")),
            anchor=str(value.get("anchor", "")),
            start=_optional_int(value.get("start")),
            end=_optional_int(value.get("end")),
        )


@dataclass(frozen=True)
class InlineNode:
    kind: str
    text: str = ""
    children: tuple[InlineNode, ...] = ()
    attrs: tuple[tuple[str, str], ...] = ()
    id: str = ""

    def __post_init__(self) -> None:
        if self.kind not in INLINE_KINDS:
            raise ValueError(f"unknown inline kind: {self.kind}")
        object.__setattr__(self, "attrs", _pairs(self.attrs))

    def attr(self, key: str, default: str = "") -> str:
        return _attr(self.attrs, key, default)

    def plain(self) -> str:
        if self.kind in {"text", "code", "math"}:
            return self.text
        if self.kind == "break":
            return "\n"
        return "".join(child.plain() for child in self.children) or self.text

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {"kind": self.kind}
        if self.id:
            payload["id"] = self.id
        if self.text:
            payload["text"] = self.text
        if self.attrs:
            payload["attrs"] = {key: value for key, value in self.attrs}
        if self.children:
            payload["children"] = [child.to_dict() for child in self.children]
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> InlineNode:
        children = tuple(cls.from_dict(item) for item in value.get("children", ()))  # type: ignore[union-attr]
        return cls(
            kind=str(value["kind"]),
            text=str(value.get("text", "")),
            children=children,
            attrs=_pairs(value.get("attrs") or ()),  # type: ignore[arg-type]
            id=str(value.get("id", "")),
        )

    def project(self) -> dict[str, object]:
        payload: dict[str, object] = {"kind": self.kind, "text": self.text}
        if self.attrs:
            payload["attrs"] = {key: value for key, value in self.attrs}
        if self.children:
            payload["children"] = [child.project() for child in self.children]
        return payload


@dataclass(frozen=True)
class BlockNode:
    id: str
    kind: str
    children: tuple[BlockNode, ...] = ()
    inlines: tuple[InlineNode, ...] = ()
    attrs: tuple[tuple[str, str], ...] = ()
    provenance: Provenance | None = None

    def __post_init__(self) -> None:
        if self.kind not in BLOCK_KINDS:
            raise ValueError(f"unknown block kind: {self.kind}")
        if not self.id:
            raise ValueError("block id must not be empty")
        object.__setattr__(self, "attrs", _pairs(self.attrs))

    def attr(self, key: str, default: str = "") -> str:
        return _attr(self.attrs, key, default)

    def plain(self) -> str:
        parts = [node.plain() for node in self.inlines]
        parts.extend(child.plain() for child in self.children)
        if self.kind == "math":
            parts.append(self.attr("tex"))
        return "".join(parts)

    def walk(self) -> Iterator[BlockNode]:
        yield self
        for child in self.children:
            yield from child.walk()

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {"id": self.id, "kind": self.kind}
        if self.attrs:
            payload["attrs"] = {key: value for key, value in self.attrs}
        if self.inlines:
            payload["inlines"] = [node.to_dict() for node in self.inlines]
        if self.children:
            payload["children"] = [child.to_dict() for child in self.children]
        if self.provenance is not None:
            payload["provenance"] = self.provenance.to_dict()
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> BlockNode:
        return cls(
            id=str(value["id"]),
            kind=str(value["kind"]),
            children=tuple(cls.from_dict(item) for item in value.get("children", ())),  # type: ignore[union-attr]
            inlines=tuple(InlineNode.from_dict(item) for item in value.get("inlines", ())),  # type: ignore[union-attr]
            attrs=_pairs(value.get("attrs") or ()),  # type: ignore[arg-type]
            provenance=Provenance.from_dict(value.get("provenance")),  # type: ignore[arg-type]
        )

    def project(self) -> dict[str, object]:
        payload: dict[str, object] = {
            "kind": self.kind,
            "attrs": {key: value for key, value in self.attrs},
            "inlines": [node.project() for node in self.inlines],
            "children": [child.project() for child in self.children],
        }
        return payload


@dataclass(frozen=True)
class Reference:
    id: str
    kind: str
    key: str
    target: str = ""

    def to_dict(self) -> dict[str, str]:
        payload = {"id": self.id, "kind": self.kind, "key": self.key}
        if self.target:
            payload["target"] = self.target
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> Reference:
        return cls(str(value["id"]), str(value["kind"]), str(value["key"]), str(value.get("target", "")))

    def project(self) -> dict[str, str]:
        return {"kind": self.kind, "key": self.key, "target": self.target}


@dataclass(frozen=True)
class ResourceDescriptor:
    id: str
    media_type: str
    sha256: str = ""
    path: str = ""

    def __post_init__(self) -> None:
        if self.path and (_ABS_PATH.match(self.path) or ".." in self.path.split("/")):
            raise ValueError("resource path must be a relative portable path")

    def to_dict(self) -> dict[str, str]:
        payload = {"id": self.id, "media_type": self.media_type}
        if self.sha256:
            payload["sha256"] = self.sha256
        if self.path:
            payload["path"] = self.path
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> ResourceDescriptor:
        return cls(
            str(value["id"]),
            str(value.get("media_type", "application/octet-stream")),
            str(value.get("sha256", "")),
            str(value.get("path", "")),
        )

    def project(self) -> dict[str, str]:
        return {"media_type": self.media_type, "sha256": self.sha256}


@dataclass(frozen=True)
class DocumentMetadata:
    title: str = ""
    authors: tuple[str, ...] = ()
    affiliations: tuple[str, ...] = ()
    language: str = ""
    extra: tuple[tuple[str, str], ...] = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "extra", _pairs(self.extra))

    def to_dict(self) -> dict[str, object]:
        payload: dict[str, object] = {"title": self.title, "authors": list(self.authors)}
        if self.affiliations:
            payload["affiliations"] = list(self.affiliations)
        if self.language:
            payload["language"] = self.language
        if self.extra:
            payload["extra"] = {key: value for key, value in self.extra}
        return payload

    @classmethod
    def from_dict(cls, value: Mapping[str, object] | None) -> DocumentMetadata:
        if not value:
            return cls()
        return cls(
            title=str(value.get("title", "")),
            authors=tuple(str(item) for item in value.get("authors", ())),  # type: ignore[union-attr]
            affiliations=tuple(str(item) for item in value.get("affiliations", ())),  # type: ignore[union-attr]
            language=str(value.get("language", "")),
            extra=_pairs(value.get("extra") or ()),  # type: ignore[arg-type]
        )

    def project(self) -> dict[str, object]:
        return self.to_dict()


@dataclass(frozen=True)
class SemanticDocument:
    """Canonical, versioned document tree."""

    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    blocks: tuple[BlockNode, ...] = ()
    references: tuple[Reference, ...] = ()
    resources: tuple[ResourceDescriptor, ...] = ()
    schema: str = SCHEMA

    def __post_init__(self) -> None:
        if self.schema != SCHEMA:
            raise ValueError(f"unsupported semantic schema: {self.schema}")

    def resource(self, resource_id: str) -> ResourceDescriptor | None:
        for item in self.resources:
            if item.id == resource_id:
                return item
        return None

    def walk(self) -> Iterator[BlockNode]:
        for block in self.blocks:
            yield from block.walk()

    def to_dict(self) -> dict[str, object]:
        return {
            "schema": self.schema,
            "metadata": self.metadata.to_dict(),
            "blocks": [block.to_dict() for block in self.blocks],
            "references": [item.to_dict() for item in self.references],
            "resources": [item.to_dict() for item in self.resources],
        }

    @classmethod
    def from_dict(cls, value: Mapping[str, object]) -> SemanticDocument:
        schema = str(value.get("schema", ""))
        if schema != SCHEMA:
            raise ValueError(f"unsupported semantic schema: {schema}")
        return cls(
            metadata=DocumentMetadata.from_dict(value.get("metadata")),  # type: ignore[arg-type]
            blocks=tuple(BlockNode.from_dict(item) for item in value.get("blocks", ())),  # type: ignore[union-attr]
            references=tuple(Reference.from_dict(item) for item in value.get("references", ())),  # type: ignore[union-attr]
            resources=tuple(ResourceDescriptor.from_dict(item) for item in value.get("resources", ())),  # type: ignore[union-attr]
            schema=schema,
        )

    def projection(self) -> dict[str, object]:
        """Normalized meaning used for equivalence. Provenance and ids are omitted."""

        return {
            "schema": self.schema,
            "metadata": self.metadata.project(),
            "blocks": [block.project() for block in self.blocks],
            "references": [item.project() for item in self.references],
            "resources": {item.id: item.project() for item in self.resources},
        }

    def semantic_hash(self) -> str:
        payload = json.dumps(self.projection(), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def validate(self) -> tuple[str, ...]:
        """Return integrity errors. An empty tuple means the tree is consistent."""

        errors: list[str] = []
        ids: list[str] = []
        for block in self.walk():
            if block.id in ids:
                errors.append(f"duplicate block id: {block.id}")
            ids.append(block.id)
            if block.kind == "figure":
                resource_id = block.attr("resource")
                if resource_id and self.resource(resource_id) is None:
                    errors.append(f"figure {block.id} references missing resource {resource_id}")
            if block.kind == "table_cell":
                for name in ("colspan", "rowspan"):
                    raw = block.attr(name)
                    if raw and (not raw.isdigit() or int(raw) < 1):
                        errors.append(f"{block.id} has invalid {name}")
        resource_ids = [item.id for item in self.resources]
        if len(resource_ids) != len(set(resource_ids)):
            errors.append("duplicate resource id")
        for item in self.references:
            if not item.key:
                errors.append(f"reference {item.id} has an empty key")
        return tuple(errors)


def _optional_int(value: object) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError("provenance offsets must be integers")
    return value


def text_node(text: str, *, id: str = "") -> InlineNode:
    return InlineNode("text", text=text, id=id)


__all__ = [
    "SCHEMA", "BLOCK_KINDS", "INLINE_KINDS", "Provenance", "InlineNode", "BlockNode",
    "Reference", "ResourceDescriptor", "DocumentMetadata", "SemanticDocument", "text_node",
]
