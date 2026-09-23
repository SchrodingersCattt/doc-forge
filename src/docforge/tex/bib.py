"""BibTeX parsing and TeX citation resolution."""
from __future__ import annotations
import re
from pathlib import Path
from .tokenize import plain_tex

def parse_bib(path: Path) -> dict[str, dict]:
    """Parse a BibTeX file into {key: field-dict} (fields lower-cased)."""
    text = path.read_text(encoding="utf-8")
    entries: dict[str, dict] = {}
    for m in re.finditer(r"@(\w+)\{(\w+),\s*(.*?)\n\}", text, re.DOTALL):
        entry_type, key, body = m.group(1), m.group(2), m.group(3)
        fields: dict[str, str] = {"_type": entry_type}
        for fm in re.finditer(r"(\w+)\s*=\s*\{(.*?)\}(?:\s*,|\s*$)", body, re.DOTALL):
            fname = fm.group(1).lower()
            fval = re.sub(r"\s+", " ", fm.group(2).strip())
            fields[fname] = fval
        entries[key] = fields
    return entries



def _author_short(raw: str) -> str:
    raw = raw.replace(" and others", " et al.")
    parts = [a.strip() for a in raw.split(" and ") if a.strip()]
    if not parts:
        return raw
    first = parts[0].split(",")[0].strip()
    if len(parts) == 1:
        return first
    if len(parts) == 2:
        if parts[1].strip().startswith("et al"):
            return f"{first} et al."
        return f"{first} & {parts[1].split(',')[0].strip()}"
    return f"{first} et al."


# ═══════════════════════════════════════════════════════════════════════════
#  Citation resolver
# ═══════════════════════════════════════════════════════════════════════════

class CitationResolver:
    def __init__(self, bib: dict[str, dict],
                 inherited_numbers: dict[str, int] | None = None,
                 local_prefix: str = ""):
        self.bib = bib
        self.order: list[str] = []
        self._key2num: dict[str, int] = {}
        self.inherited_numbers = inherited_numbers or {}
        self.local_prefix = local_prefix

    def resolve(self, keys_str: str, use_inherited: bool = True) -> str:
        return _compress_labels(self.labels_for(keys_str, use_inherited=use_inherited))

    def labels_for(self, keys_str: str, use_inherited: bool = True) -> list[str]:
        keys = [k.strip() for k in keys_str.split(",")]
        labels = []
        for k in keys:
            if not k:
                continue
            labels.append(self._label_for_key(k, use_inherited=use_inherited))
        return labels

    def citation_parts(self, keys_str: str, use_inherited: bool = True) -> list[tuple[str, str | None]]:
        """Return display chunks for an in-text citation.

        Each tuple is (text, target_label). Separators have target_label=None;
        ranges link to their first reference.
        """
        labels = self.labels_for(keys_str, use_inherited=use_inherited)
        if not labels:
            return []
        groups = _group_labels(labels)
        parts: list[tuple[str, str | None]] = []
        for gi, group in enumerate(groups):
            if gi > 0:
                parts.append((",", None))
            if len(group) <= 2:
                for li, item in enumerate(group):
                    if li > 0:
                        parts.append((",", None))
                    parts.append((item[2], item[2]))
            else:
                parts.append((f"{group[0][2]}–{group[-1][2]}", group[0][2]))
        return parts

    def _label_for_key(self, key: str, use_inherited: bool = True) -> str:
        if use_inherited and key in self.inherited_numbers:
            return str(self.inherited_numbers[key])
        if key not in self._key2num:
            self._key2num[key] = len(self.order) + 1
            self.order.append(key)
        return f"{self.local_prefix}{self._key2num[key]}"

    def reference_list(self) -> list[str]:
        return [ref for _, ref in self.reference_items()]

    def reference_items(self) -> list[tuple[str, str]]:
        refs: list[str] = []
        items: list[tuple[str, str]] = []
        seen = set()
        for key in self.order:
            if key in seen or key not in self.bib:
                continue
            seen.add(key)
            e = self.bib[key]
            author = _author_short(e.get("author", ""))
            title = plain_tex(e.get("title", ""))
            journal = plain_tex(e.get("journal", e.get("booktitle", "")))
            vol = e.get("volume", "")
            pages = plain_tex(e.get("pages", ""))
            year = e.get("year", "")
            doi = e.get("doi", "")
            parts = [f"{author}.", f"{title}."]
            if journal:
                j = journal
                if vol:
                    j += f" {vol}"
                if pages:
                    j += f", {pages}"
                j += f" ({year})."
                parts.append(j)
            elif year:
                parts.append(f"({year}).")
            if doi:
                parts.append(f"https://doi.org/{doi}")
            label = f"{self.local_prefix}{self._key2num[key]}"
            items.append((label, " ".join(parts)))
        return items

    def citation_numbers(self) -> dict[str, int]:
        return dict(self._key2num)


def _citation_sort_key(label: str) -> tuple[int, int, str]:
    m = re.match(r"^([A-Za-z]*)(\d+)$", label)
    if not m:
        return (2, 0, label)
    prefix, num = m.group(1), int(m.group(2))
    if not prefix:
        return (0, num, label)
    return (1, num, prefix)


def _group_labels(labels: list[str]) -> list[list[tuple[str, int, str]]]:
    if not labels:
        return []
    unique = sorted(set(labels), key=_citation_sort_key)
    parsed: list[tuple[str, int, str]] = []
    for label in unique:
        m = re.match(r"^([A-Za-z]*)(\d+)$", label)
        if m:
            parsed.append((m.group(1), int(m.group(2)), label))
        else:
            parsed.append((label, -1, label))
    groups: list[list[tuple[str, int, str]]] = [[parsed[0]]]
    for item in parsed[1:]:
        prev = groups[-1][-1]
        if item[0] == prev[0] and item[1] == prev[1] + 1 and item[1] > 0:
            groups[-1].append(item)
        else:
            groups.append([item])
    return groups


def _compress_labels(labels: list[str]) -> str:
    groups = _group_labels(labels)
    if not groups:
        return ""
    parts = []
    for g in groups:
        if len(g) <= 2:
            parts.extend(item[2] for item in g)
        else:
            parts.append(f"{g[0][2]}–{g[-1][2]}")
    return "[" + ",".join(parts) + "]"


def _resolve_cites_in_text(text: str, resolver: CitationResolver) -> str:
    text = re.sub(r"\\mainref\{([^}]+)\}", lambda m: resolver.resolve(m.group(1)), text)
    return re.sub(r"\\cite\{([^}]+)\}", lambda m: resolver.resolve(m.group(1), use_inherited=False), text)

__all__=["parse_bib","CitationResolver","_author_short","_compress_labels"]
