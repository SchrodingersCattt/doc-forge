"""Configurable abbreviation gate with project-specific whitelists."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

BUILTIN_WHITELIST = {"AI", "API", "DOI", "GPU", "ORCID", "PDF", "SI", "URL"}
ACRONYM_RE = re.compile(r"\b[A-Z][A-Z0-9-]{1,}\b")


@dataclass(frozen=True)
class AbbreviationFinding:
    path: str
    line: int
    acronym: str
    rule: str
    message: str


def load_abbreviation_whitelist(path: Path | None = None, *, project: str | None = None) -> set[str]:
    """Load ``abbr-whitelist`` from a JSON gate config.

    Accepted forms are a top-level list, ``{"abbr-whitelist": [...]}``, or a
    project map under ``projects``.  The package defaults remain deliberately
    small and language-independent.
    """
    allowed = set(BUILTIN_WHITELIST)
    if path is None:
        return allowed
    payload: Any = json.loads(path.read_text(encoding="utf-8-sig"))
    values: Any = payload
    if isinstance(payload, dict):
        values = payload.get("abbr-whitelist", payload.get("abbreviation_whitelist", []))
        if project and isinstance(payload.get("projects"), dict):
            project_config = payload["projects"].get(project, {})
            if isinstance(project_config, dict):
                values = list(values or []) + list(project_config.get("abbr-whitelist", project_config.get("abbreviation_whitelist", [])) or [])
    if not isinstance(values, list) or not all(isinstance(item, str) for item in values):
        raise ValueError("abbr-whitelist must be a list of strings")
    allowed.update(item.strip() for item in values if item.strip())
    return allowed


def _plain_text(line: str) -> str:
    line = re.sub(r"(?<!\\)%.*$", "", line)
    line = re.sub(r"\\(?:cite\w*|ref|eqref|label|url|path|href)\*?(?:\[[^]]*\])?\{[^{}]*\}", " ", line)
    line = re.sub(r"\$[^$]*\$|\\\([^)]*\\\)|\\\[[^]]*\\\]", " ", line)
    line = re.sub(r"\\[A-Za-z@]+\*?(?:\[[^]]*\])?", " ", line)
    return re.sub(r"\s+", " ", line.replace("\\%", "%")).strip()


def check_abbreviations(
    paths: list[Path],
    *,
    whitelist: set[str] | None = None,
    require_definition: bool = True,
    require_reuse: bool = False,
) -> list[AbbreviationFinding]:
    allowed = set(BUILTIN_WHITELIST if whitelist is None else whitelist)
    findings: list[AbbreviationFinding] = []
    for path in paths:
        text = path.read_text(encoding="utf-8-sig")
        lines = [_plain_text(line) for line in text.splitlines()]
        occurrences: dict[str, list[tuple[int, str]]] = {}
        for number, line in enumerate(lines, 1):
            for match in ACRONYM_RE.finditer(line):
                acronym = match.group(0)
                if acronym in allowed or acronym.isdigit():
                    continue
                occurrences.setdefault(acronym, []).append((number, line))
        for acronym, seen in occurrences.items():
            defined = any(
                re.search(rf"[A-Za-z][A-Za-z0-9 -]{{2,}}\(\s*{re.escape(acronym)}\s*\)", line)
                for _, line in seen
            )
            if require_definition and not defined:
                line, excerpt = seen[0]
                findings.append(AbbreviationFinding(str(path), line, acronym, "ABBR_UNDEFINED", f"Define abbreviation {acronym} at first use."))
            if require_reuse and len(seen) == 1 and defined:
                line, excerpt = seen[0]
                findings.append(AbbreviationFinding(str(path), line, acronym, "ABBR_UNUSED", f"Remove abbreviation {acronym}; it is defined but not reused."))
    return findings

