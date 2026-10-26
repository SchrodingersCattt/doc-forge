"""Canonical semantic document model."""

from .model import (
    BLOCK_KINDS,
    INLINE_KINDS,
    SCHEMA,
    BlockNode,
    DocumentMetadata,
    InlineNode,
    Provenance,
    Reference,
    ResourceDescriptor,
    SemanticDocument,
    text_node,
)

__all__ = [
    "SCHEMA", "BLOCK_KINDS", "INLINE_KINDS", "Provenance", "InlineNode", "BlockNode",
    "Reference", "ResourceDescriptor", "DocumentMetadata", "SemanticDocument", "text_node",
]
