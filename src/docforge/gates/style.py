"""Configurable, project-neutral prose style gate."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

DEFAULT_RULES = {
    "EM_DASH": (r"—|---", "Use an en dash for ranges or relations."),
    "NUMERIC_HYPHEN_RANGE": (r"(?<![\w.])\d+(?:\.\d+)?\s-\s\d+(?:\.\d+)?(?!\w)", "Use an en dash for numeric ranges."),
    "UNICODE_SCRIPT": (r"[⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉]", "Use native markup for subscripts and superscripts."),
}


@dataclass(frozen=True)
class StyleFinding:
    path: str
    line: int
    rule: str
    message: str
    excerpt: str
    token: str

    @property
    def fingerprint(self) -> str:
        payload = "|".join((self.path, self.rule, re.sub(r"\s+", " ", self.excerpt).strip(), self.token))
        return f"{self.rule}:{hashlib.sha256(payload.encode('utf-8')).hexdigest()[:20]}"


def load_style_config(path: Path) -> dict[str, Any]:
    config = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(config, dict):
        raise ValueError("style gate config must be a JSON object")
    return config


def _plain_line(line: str) -> str:
    line = re.sub(r"(?<!\\)%.*$", "", line)
    line = re.sub(r"\\(?:cite\w*|ref|eqref|label|url|path|includegraphics|input|bibliography)\*?(?:\[[^]]*\])?\{[^{}]*\}", " ", line)
    line = re.sub(r"\$[^$]*\$|\\\([^)]*\\\)|\\\[[^]]*\\\]", " ", line)
    line = re.sub(r"\\[A-Za-z@]+\*?(?:\[[^]]*\])?", " ", line)
    return re.sub(r"\s+", " ", line.replace("\\%", "%")).strip()


def check_style(paths: list[Path], *, config: dict[str, Any] | None = None) -> list[StyleFinding]:
    config = config or {}
    rules = dict(DEFAULT_RULES)
    for name, rule in config.get("rules", {}).items():
        if isinstance(rule, dict) and rule.get("pattern"):
            rules[name] = (str(rule["pattern"]), str(rule.get("message", name)))
    findings: list[StyleFinding] = []
    for path in paths:
        for number, raw in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
            line = _plain_line(raw)
            if not line:
                continue
            for rule, (pattern, message) in rules.items():
                for match in re.finditer(pattern, line):
                    findings.append(StyleFinding(str(path), number, rule, message, line, match.group(0)))
    return findings


def baseline_payload(findings: list[StyleFinding]) -> dict[str, Any]:
    return {"version": 1, "findings": [
        {"fingerprint": item.fingerprint, "path": item.path, "line": item.line, "rule": item.rule, "excerpt": item.excerpt[:220]}
        for item in findings
    ]}


def unresolved(findings: list[StyleFinding], baseline: dict[str, Any] | None = None) -> list[StyleFinding]:
    allowed = {str(item.get("fingerprint")) for item in (baseline or {}).get("findings", []) if isinstance(item, dict)}
    return [item for item in findings if item.fingerprint not in allowed]

