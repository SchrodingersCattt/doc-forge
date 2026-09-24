"""Project-neutral TeX figure/table declaration and reference gate."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

LABEL_RE = re.compile(r"\\label\{(?P<label>(?P<kind>fig|tab):[^}]+)\}")
REF_RE = re.compile(r"\\ref\*?\{(?P<label>[^}]+)\}")
IGNORED_KINDS = {"sec", "si", "alg", "eq", "sub"}


@dataclass(frozen=True)
class ReferenceIssue:
    kind: str
    label: str
    message: str


def strip_comments(text: str) -> str:
    lines = []
    for line in text.splitlines():
        escaped = False
        cut = len(line)
        for index, char in enumerate(line):
            if char == "%" and not escaped:
                cut = index
                break
            escaped = char == "\\" and not escaped
            if char != "\\":
                escaped = False
        lines.append(line[:cut])
    return "\n".join(lines)


def audit_references(root: Path, config: dict[str, Any]) -> list[ReferenceIssue]:
    documents = config.get("documents")
    if not isinstance(documents, dict) or not documents:
        raise ValueError("reference gate requires a non-empty documents mapping")
    root = root.resolve()
    texts: dict[str, str] = {}
    for relative in documents:
        path = (root / relative).resolve()
        path.relative_to(root)
        texts[relative] = strip_comments(path.read_text(encoding="utf-8"))
    labels: dict[str, tuple[str, str]] = {}
    references: dict[str, list[str]] = {}
    declarations: dict[str, list[str]] = {name: [] for name in texts}
    for name, text in texts.items():
        for match in LABEL_RE.finditer(text):
            label = match.group("label")
            if label in labels:
                raise ValueError(f"duplicate label: {label}")
            labels[label] = (name, match.group("kind"))
            declarations[name].append(label)
        for match in REF_RE.finditer(text):
            references.setdefault(match.group("label"), []).append(name)
    issues: list[ReferenceIssue] = []
    for label in sorted(set(references) - set(labels)):
        if ":" in label and label.split(":", 1)[0] not in IGNORED_KINDS:
            issues.append(ReferenceIssue("reference", label, "references an undefined label"))
    main = str(config.get("main_document", next(iter(texts))))
    if main not in texts:
        raise ValueError(f"main_document is not listed in documents: {main}")
    for label in declarations[main]:
        if label.startswith(("fig:", "tab:")) and main not in references.get(label, []):
            issues.append(ReferenceIssue(label.split(":", 1)[0], label, "main figure/table is never cited in main prose"))
    return issues

