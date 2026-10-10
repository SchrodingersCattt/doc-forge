"""Profile-driven, transactional delivery of Markdown manuscripts.

The delivery command is intentionally the only place where a reviewed DOCX is
allowed to flow back into source files. A reviewed document is converted in a
scratch directory, then its mapped line ranges are applied to staged copies of
the existing Markdown files. Originals and delivery outputs are published
together only after every artifact has passed preflight and built successfully.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from lxml import etree

from ..output import validate_output_path
from .docx_export import docx_to_markdown, load_section_map, split_markdown_sections
from .template import AssemblyResult, assemble_markdown_template, write_assembly_sidecars


@dataclass(frozen=True)
class DeliveryArtifact:
    """One generated delivery document and its audit sidecars."""

    name: str
    output: Path
    manifest: Path
    checksum: Path
    inputs: tuple[Path, ...]


@dataclass(frozen=True)
class DeliveryResult:
    """Paths produced by :func:`deliver`."""

    delivery_dir: Path
    artifacts: tuple[DeliveryArtifact, ...]
    manifest: Path


@dataclass(frozen=True)
class _PreparedReview:
    inputs: tuple[Path, ...]
    logical_inputs: tuple[Path, ...]
    updates: tuple[tuple[Path, Path], ...]
    source_map: tuple[dict[str, Any], ...]


_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_REVISION_NAMES = frozenset(
    {"ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "trPrChange"}
)
_ROLE_ALIASES = {"si": "supplement", "supporting_information": "supplement"}
_INPUT_ALIASES = ("inputs", "sources", "markdown")
_REVIEW_ALIASES = ("reviewed_docx", "reviewed", "review")
_OPTION_KEYS = (
    "title", "keep_comments", "skip_images", "columns", "figure_span", "font_family",
    "east_asia_font", "style_profile", "line_numbers", "include_title",
    "strip_level_one_headings", "heading_before", "numbering_prefix", "bibliography_scope",
    "citation_numbering", "bibliography_profile", "include_metadata_back_matter", "native_toc",
    "restart_heading_numbering", "body_first_line_chars", "page_break_before_h1",
    "body_font_size", "abstract_font_size", "caption_font_size", "reference_font_size",
)


def _canonical_entry(key: str, value: Mapping[str, Any]) -> tuple[str, dict[str, Any]]:
    entry = dict(value)
    for canonical, aliases in (("inputs", _INPUT_ALIASES), ("reviewed_docx", _REVIEW_ALIASES)):
        if canonical not in entry:
            for alias in aliases:
                if alias in entry:
                    entry[canonical] = entry[alias]
                    break
    for canonical, aliases in (("section_map", ("section_map_path",)), ("source_dir", ("sources_dir",)),
                               ("source_markdown", ("source",))):
        if canonical not in entry:
            for alias in aliases:
                if alias in entry:
                    entry[canonical] = entry[alias]
                    break
    return key, entry


def _canonical_profile(profile: Mapping[str, Any]) -> dict[str, Any]:
    """Normalize historical aliases to one versioned delivery profile."""
    payload = dict(profile)
    schema = payload.get("schema", payload.get("profile_version"))
    if schema is not None and str(schema) not in {"docforge.delivery.v1", "1", "1.0"}:
        raise ValueError(f"unsupported delivery profile schema: {schema}")
    payload["schema"] = "docforge.delivery.v1"
    if "delivery_dir" not in payload and "output_dir" in payload:
        payload["delivery_dir"] = payload["output_dir"]
    if "names" not in payload and "delivery_names" in payload:
        payload["names"] = payload["delivery_names"]
    if "reviewed_docx" not in payload:
        for alias in _REVIEW_ALIASES[1:]:
            if alias in payload:
                payload["reviewed_docx"] = payload[alias]
                break
    result: list[tuple[str, dict[str, Any]]] = []
    artifacts = payload.get("artifacts")
    if artifacts is not None:
        if not isinstance(artifacts, list) or not artifacts:
            raise ValueError("delivery profile 'artifacts' must be a non-empty list")
        for index, item in enumerate(artifacts):
            if not isinstance(item, Mapping):
                raise ValueError(f"delivery artifact {index} must be an object")
            value = dict(item)
            key = str(value.pop("id", value.pop("name", f"artifact_{index + 1}")))
            result.append(_canonical_entry(key, value))
    else:
        for original in ("article", "supplement", "si", "supporting_information"):
            value = payload.get(original)
            if value is None:
                continue
            if not isinstance(value, Mapping):
                raise ValueError(f"delivery profile '{original}' must be an object")
            key = _ROLE_ALIASES.get(original, original)
            if any(item[0] == key for item in result):
                raise ValueError(f"delivery profile defines '{key}' more than once")
            result.append((key, _canonical_entry(key, value)[1]))
    if not result:
        raise ValueError("delivery profile requires article/supplement or artifacts")
    canonical_artifacts = [{"id": key, **entry} for key, entry in result]
    # Preserve the concise root-level review form while still presenting one
    # canonical per-artifact shape to the rest of the implementation.
    for item in canonical_artifacts:
        key = str(item["id"])
        for field in (
            "reviewed_docx", "section_map", "source_paragraph_map", "revision_map",
            "source_dir", "source_markdown",
        ):
            if field in item or field not in payload:
                continue
            value = payload[field]
            item[field] = value.get(key) if isinstance(value, Mapping) else value
    payload["artifacts"] = canonical_artifacts
    for key in ("article", "supplement", "si", "supporting_information"):
        payload.pop(key, None)
    return payload


def _load_profile(profile: Mapping[str, Any] | Path | str) -> tuple[dict[str, Any], Path, str]:
    if isinstance(profile, (str, Path)):
        profile_path = Path(profile)
        payload = json.loads(profile_path.read_text(encoding="utf-8"))
        source, base = str(profile_path), profile_path.parent.resolve()
    else:
        payload, source, base = dict(profile), "<mapping>", Path.cwd().resolve()
    if not isinstance(payload, dict):
        raise ValueError("delivery profile must be a JSON object")
    return _canonical_profile(payload), base, source


def _path(value: Any, base: Path, *, label: str) -> Path:
    if not isinstance(value, (str, Path)) or not str(value).strip():
        raise ValueError(f"delivery profile requires {label}")
    candidate = Path(value)
    return candidate if candidate.is_absolute() else base / candidate


def _entries(profile: Mapping[str, Any]) -> list[tuple[str, dict[str, Any]]]:
    artifacts = profile.get("artifacts")
    if not isinstance(artifacts, list) or not artifacts:
        raise ValueError("delivery profile requires a non-empty canonical 'artifacts' list")
    result: list[tuple[str, dict[str, Any]]] = []
    for index, item in enumerate(artifacts):
        if not isinstance(item, Mapping):
            raise ValueError(f"delivery artifact {index} must be an object")
        value = dict(item)
        key = str(value.pop("id", value.pop("name", f"artifact_{index + 1}")))
        result.append((key, _canonical_entry(key, value)[1]))
    return result


def _resolve_inputs(value: Any, base: Path, *, label: str) -> list[Path]:
    if isinstance(value, (str, Path)):
        values: Sequence[Any] = [value]
    elif isinstance(value, Sequence) and not isinstance(value, (bytes, bytearray)):
        values = value
    else:
        raise ValueError(f"{label} must be a path or list of paths")
    result = [_path(item, base, label=label) for item in values]
    if not result:
        raise ValueError(f"{label} must not be empty")
    missing = [str(path) for path in result if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Markdown input does not exist: {missing[0]}")
    return result


def _safe_relative(value: Any, *, label: str) -> Path:
    if not isinstance(value, (str, Path)) or not str(value).strip():
        raise ValueError(f"{label} must be a non-empty relative path")
    candidate = Path(str(value))
    if candidate.is_absolute() or ".." in candidate.parts or candidate == Path("."):
        raise ValueError(f"{label} must remain confined to its delivery directory: {value}")
    return candidate


def _output_name(entry: Mapping[str, Any], key: str, profile: Mapping[str, Any], *, timestamp: str) -> Path:
    names = profile.get("names", {})
    named = names.get(key) if isinstance(names, Mapping) else None
    value = entry.get("delivery_name") or entry.get("output_name") or entry.get("output") or named
    if value is None:
        pattern = entry.get("output_pattern", profile.get("output_pattern"))
        if pattern:
            try:
                value = str(pattern).format(role=key, name=key, key=key, kind=key, timestamp=timestamp)
            except (KeyError, ValueError) as exc:
                raise ValueError(f"invalid delivery output pattern: {pattern!r}") from exc
        else:
            value = f"{key}.docx"
    name = _safe_relative(value, label=f"delivery {key}")
    if name.suffix.lower() != ".docx":
        name = name.with_suffix(name.suffix + ".docx" if name.suffix else ".docx")
    validate_output_path(name, label=f"delivery {key}")
    return name


def _sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _source_map(inputs: Sequence[Path]) -> list[dict[str, Any]]:
    """Return a line-addressable, hashed map of every source paragraph."""
    result: list[dict[str, Any]] = []
    for path in inputs:
        raw = path.read_text(encoding="utf-8-sig")
        lines = raw.splitlines(keepends=True)
        index, paragraph = 0, 0
        found = False
        while index < len(lines):
            while index < len(lines) and not lines[index].strip():
                index += 1
            if index >= len(lines):
                break
            found, start = True, index
            while index < len(lines) and lines[index].strip():
                index += 1
            segment = "".join(lines[start:index])
            result.append({
                "path": str(path.resolve()), "source": str(path),
                "start_line": start + 1, "end_line": index,
                "hash": _sha256_text(segment), "paragraph": paragraph,
                "text": "\n".join(line.rstrip("\r\n") for line in lines[start:index]),
            })
            paragraph += 1
        if not found:
            result.append({
                "path": str(path.resolve()), "source": str(path), "start_line": 1,
                "end_line": 0, "hash": _sha256_text(raw), "paragraph": None,
                "text": "", "line_count": len(lines),
            })
    return result


def _docx_has_revisions(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            for name in archive.namelist():
                if not name.startswith("word/") or not name.endswith(".xml"):
                    continue
                try:
                    root = etree.fromstring(archive.read(name))
                except etree.XMLSyntaxError as exc:
                    raise ValueError(f"invalid XML in reviewed DOCX part {name}") from exc
                if any(etree.QName(node).localname in _REVISION_NAMES for node in root.iter()):
                    return True
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError(f"cannot read reviewed DOCX safely: {path}") from exc
    return False


def _docx_paragraph_text(paragraph: etree._Element, *, view: str) -> str:
    """Read the visible text in one DOCX paragraph for a revision view."""
    pieces: list[str] = []

    def walk(node: etree._Element, hidden: bool = False) -> None:
        local = etree.QName(node).localname
        if local in {"del", "moveFrom"}:
            hidden = view == "accepted"
        elif local in {"ins", "moveTo"}:
            hidden = view == "original"
        if hidden:
            return
        if local in {"t", "delText"}:
            if local == "t" or view == "original":
                pieces.append(node.text or "")
            return
        if local == "tab":
            pieces.append("\t")
            return
        if local in {"br", "cr"}:
            pieces.append("\n")
            return
        for child in node:
            walk(child, hidden)

    walk(paragraph)
    return "".join(pieces)


def _docx_paragraph_records(path: Path) -> dict[str, dict[str, Any]]:
    """Return stable ``w14:paraId`` records from the reviewed document."""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            try:
                root = etree.fromstring(archive.read("word/document.xml"))
            except KeyError as exc:
                raise ValueError("reviewed DOCX has no word/document.xml") from exc
            except etree.XMLSyntaxError as exc:
                raise ValueError("invalid XML in reviewed DOCX part word/document.xml") from exc
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError(f"cannot read reviewed DOCX safely: {path}") from exc
    records: dict[str, dict[str, Any]] = {}
    body = root.find(f".//{{{_W}}}body")
    if body is None:
        raise ValueError("reviewed DOCX has no document body")
    for ordinal, paragraph in enumerate(body.iter(f"{{{_W}}}p")):
        identifier = paragraph.get(f"{{{_W14}}}paraId")
        if not identifier:
            continue
        revisions = {
            etree.QName(node).localname
            for node in paragraph.iter()
            if etree.QName(node).localname in {"ins", "del"}
        }
        records[str(identifier).casefold()] = {
            "id": str(identifier),
            "ordinal": ordinal,
            "revisions": revisions,
            "original_text": _docx_paragraph_text(paragraph, view="original"),
            "accepted_text": _docx_paragraph_text(paragraph, view="accepted"),
        }
    return records


def _reviewed_identity(item: Mapping[str, Any]) -> str | None:
    value = item.get("reviewed_paragraph", item.get("reviewed_id"))
    if isinstance(value, Mapping):
        value = value.get("id", value.get("paragraph_id", value.get("para_id")))
    if value is None or not str(value).strip():
        return None
    return str(value)


def _normalise_edit_text(value: str) -> str:
    """Compare Markdown and DOCX paragraph text without markup syntax."""
    value = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"[`*_#>~]", "", value)
    return re.sub(r"\s+", " ", value).strip()


def _local_media_references(markdown: str) -> list[str]:
    """Extract local Markdown/HTML image references from a candidate chunk."""
    references = re.findall(r"!\[[^\]]*\]\((?:<([^>]+)>|([^\s)]+))", markdown)
    values = [first or second for first, second in references]
    values.extend(re.findall(r"<img\b[^>]*\bsrc=[\"']([^\"']+)[\"']", markdown, flags=re.IGNORECASE))
    result: list[str] = []
    for value in values:
        value = value.split("#", 1)[0].split("?", 1)[0].strip()
        if not value or re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", value) or value.startswith("//"):
            continue
        result.append(value)
    return result


def _stamp_output_paragraph_ids(path: Path, review: _PreparedReview | None) -> dict[tuple[str, int], dict[str, Any]]:
    """Carry reviewed paragraph IDs into the generated DOCX when possible."""
    if review is None or not review.source_map:
        return {}
    try:
        with zipfile.ZipFile(path, "r") as archive:
            files = {name: archive.read(name) for name in archive.namelist()}
        root = etree.fromstring(files["word/document.xml"])
    except (KeyError, zipfile.BadZipFile, OSError, etree.XMLSyntaxError):
        # Test doubles and non-DOCX renderers cannot carry package identities;
        # the source map still contains the reviewed identity in that case.
        return {}
    body = root.find(f".//{{{_W}}}body")
    if body is None:
        return {}
    paragraphs = list(body.iter(f"{{{_W}}}p"))
    available: dict[str, list[dict[str, Any]]] = {}
    for record in review.source_map:
        identity = record.get("docx_paragraph_id")
        if not identity:
            continue
        expected = _normalise_edit_text(str(record.get("reviewed_text", "")))
        if expected:
            available.setdefault(expected, []).append(record)
    result: dict[tuple[str, int], dict[str, Any]] = {}
    used: set[int] = set()
    for ordinal, paragraph in enumerate(paragraphs):
        text = _normalise_edit_text(_docx_paragraph_text(paragraph, view="accepted"))
        if not text:
            continue
        match = next(
            (record for expected, records in available.items() if expected in text or text in expected for record in records),
            None,
        )
        if match is None or ordinal in used:
            continue
        identity = str(match["docx_paragraph_id"])
        paragraph.set(f"{{{_W14}}}paraId", identity)
        used.add(ordinal)
        result[(str(match["path"]), int(match["start_line"]))] = {
            "id": identity,
            "ordinal": ordinal,
            "start": ordinal + 1,
            "end": ordinal + 1,
        }
    if not result:
        return {}
    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    temporary = path.with_suffix(path.suffix + ".ids")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, payload in files.items():
                archive.writestr(name, payload)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    return result


def _confined(path: Path, root: Path, *, label: str) -> Path:
    root, candidate = root.resolve(), path.resolve()
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"{label} escapes its configured directory: {path}") from exc
    return candidate


def _review_chunks(markdown: str, entries: list[dict[str, Any]]) -> list[tuple[str, str]]:
    if any(item.get("reviewed_start_line") is None or item.get("reviewed_end_line") is None for item in entries):
        raise ValueError("section map requires reviewed_start_line/reviewed_end_line for every reviewed section")
    lines = markdown.splitlines(keepends=True)
    reviewed_ranges: list[tuple[int, int, str]] = []
    for item in entries:
        filename = str(item.get("file", "<unknown>"))
        start, end = int(item["reviewed_start_line"]), int(item["reviewed_end_line"])
        if start < 1 or end < start or end > len(lines):
            raise ValueError(f"reviewed section range is invalid for {filename!r}: {start}-{end}")
        if any(not (end < previous_start or start > previous_end) for previous_start, previous_end, _ in reviewed_ranges):
            raise ValueError(f"reviewed section ranges overlap for {filename!r}: {start}-{end}")
        reviewed_ranges.append((start, end, filename))
    covered = {line for start, end, _ in reviewed_ranges for line in range(start, end + 1)}
    unmatched = [index for index, line in enumerate(lines, 1) if line.strip() and index not in covered]
    if unmatched:
        shown = ", ".join(str(index) for index in unmatched[:8])
        suffix = "..." if len(unmatched) > 8 else ""
        raise ValueError(f"reviewed DOCX contains unmatched paragraphs at lines {shown}{suffix}")
    if any(item.get("start") is not None or item.get("end") is not None for item in entries):
        return split_markdown_sections(markdown, entries)
    chunks: list[tuple[str, str]] = []
    for item in entries:
        filename = str(item["file"])
        start = item.get("reviewed_start_line", item.get("start_line"))
        end = item.get("reviewed_end_line", item.get("end_line"))
        if start is None or end is None:
            raise ValueError(f"section map entry {filename!r} needs reviewed line range")
        start, end = int(start), int(end)
        if start < 1 or end < start or end > len(lines):
            raise ValueError(f"reviewed section range is invalid for {filename!r}: {start}-{end}")
        chunks.append((filename, "".join(lines[start - 1:end])))
    return chunks


def _source_range(item: Mapping[str, Any], lines: list[str], *, next_start: int | None = None) -> tuple[int, int]:
    """Resolve legacy heading maps to concrete source line ranges."""
    if item.get("start_line") is not None or item.get("end_line") is not None:
        start = int(item.get("start_line", 1))
        end = int(item.get("end_line", len(lines)))
        return start, end
    headings: list[tuple[str, int]] = []
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("#"):
            value = stripped.lstrip("#").strip().rstrip(":").strip().casefold()
            if value:
                headings.append((value, index + 1))

    def resolve(reference: Any, default: int) -> int:
        if reference is None:
            return default
        wanted = str(reference.get("heading") if isinstance(reference, Mapping) else reference).rstrip(":").strip().casefold()
        occurrence = int(reference.get("occurrence", 1)) if isinstance(reference, Mapping) else 1
        matches = [line for value, line in headings if value == wanted]
        if occurrence < 1 or occurrence > len(matches):
            raise ValueError(f"source map heading not found: {wanted!r}")
        return matches[occurrence - 1]

    start = resolve(item.get("start"), 1)
    end_default = (next_start - 1) if next_start is not None else len(lines)
    end = resolve(item.get("end"), end_default)
    if end < start:
        raise ValueError(f"source map heading range is empty: {item.get('file')}")
    return start, end


def _prepare_review(entry: Mapping[str, Any], base: Path, stage_root: Path) -> _PreparedReview:
    reviewed_path = _path(entry["reviewed_docx"], base, label="reviewed_docx")
    revision_map_value = entry.get("revision_map") or entry.get("source_paragraph_map")
    if revision_map_value is None:
        raise ValueError("--accept-revisions requires a source_paragraph_map or revision_map")
    section_map_path = _path(entry.get("section_map", revision_map_value), base, label="section_map")
    source_dir = _path(entry.get("source_dir"), base, label="source_dir").resolve()
    if not reviewed_path.is_file() or not section_map_path.is_file() or not source_dir.is_dir():
        raise FileNotFoundError("reviewed_docx, section_map, and source_dir must exist")
    mapped = load_section_map(section_map_path)
    missing = [
        str(item.get("file", "<unknown>"))
        for item in mapped
        if (
            item.get("start_line") is None
            or item.get("end_line") is None
            or item.get("hash", item.get("sha256")) is None
            or item.get("reviewed_start_line") is None
            or item.get("reviewed_end_line") is None
            or _reviewed_identity(item) is None
            or (
                item.get("reviewed_hash", item.get("accepted_hash")) is None
                and not (
                    isinstance(item.get("reviewed_paragraph", item.get("reviewed_id")), Mapping)
                    and item.get("reviewed_paragraph", item.get("reviewed_id")).get("hash")
                )
            )
        )
    ]
    if missing:
        raise ValueError("revision map requires source path/range/hash and reviewed paragraph identity/hash: " + ", ".join(missing))
    full_md = stage_root / "reviewed" / f"{hashlib.sha1(str(reviewed_path).encode()).hexdigest()}.md"
    full_md.parent.mkdir(parents=True, exist_ok=True)
    docx_to_markdown(reviewed_path, output=full_md, track_changes="accept", force=True)
    chunks = _review_chunks(full_md.read_text(encoding="utf-8"), mapped)
    reviewed_paragraphs = _docx_paragraph_records(reviewed_path)
    matched_paragraphs: set[str] = set()
    for item, (_name, replacement) in zip(mapped, chunks):
        identity = _reviewed_identity(item)
        assert identity is not None
        record = reviewed_paragraphs.get(identity.casefold())
        if record is None:
            raise ValueError(f"reviewed paragraph identity is absent from DOCX: {identity}")
        if not record["revisions"].intersection({"ins", "del"}):
            raise ValueError(f"reviewed paragraph has no w:ins/w:del revision: {identity}")
        if identity.casefold() in matched_paragraphs:
            raise ValueError(f"reviewed paragraph identity is mapped more than once: {identity}")
        matched_paragraphs.add(identity.casefold())
        reviewed_text = item.get("reviewed_text")
        reviewed_expected = _normalise_edit_text(str(reviewed_text)) if reviewed_text is not None else _normalise_edit_text(replacement)
        accepted_text = _normalise_edit_text(record["accepted_text"])
        if reviewed_expected and reviewed_expected not in accepted_text and accepted_text not in reviewed_expected:
            raise RuntimeError(f"reviewed inserted text does not match mapped paragraph: {identity}")
    # Preserve the source directory layout in staging so relative figures,
    # tables, and media links resolve exactly as they do from the profile.
    staged_root = stage_root / "sources" / hashlib.sha1(str(source_dir).encode()).hexdigest()[:12]
    staged_root.mkdir(parents=True, exist_ok=True)
    staged_by_original: dict[Path, Path] = {}
    ranges: dict[Path, list[tuple[int, int]]] = {}
    records: list[dict[str, Any]] = []
    plans: list[tuple[Path, int, int, str, dict[str, Any]]] = []
    for map_index, (item, (_name, replacement)) in enumerate(zip(mapped, chunks)):
        target = _confined(source_dir / str(item["file"]), source_dir, label="source map file")
        if not target.is_file():
            raise FileNotFoundError(f"mapped Markdown source does not exist: {target}")
        raw = target.read_text(encoding="utf-8-sig")
        old_lines = raw.splitlines(keepends=True)
        next_start = None
        if map_index + 1 < len(mapped) and mapped[map_index + 1].get("file") == item.get("file"):
            if mapped[map_index + 1].get("start_line") is not None:
                next_start = int(mapped[map_index + 1]["start_line"])
            else:
                next_start, _ = _source_range(mapped[map_index + 1], [line.rstrip("\r\n") for line in old_lines])
        start, end = _source_range(item, [line.rstrip("\r\n") for line in old_lines], next_start=next_start)
        if start < 1 or end < start or end > len(old_lines):
            raise ValueError(f"source map line range is invalid for {target}: {start}-{end}")
        if any(not (end < a or start > b) for a, b in ranges.setdefault(target, [])):
            raise ValueError(f"source map ranges overlap for {target}")
        ranges[target].append((start, end))
        expected = item.get("hash", item.get("sha256"))
        old_segment = "".join(old_lines[start - 1:end])
        if expected is not None and str(expected) != _sha256_text(old_segment):
            raise RuntimeError(f"source map hash mismatch for {target}:{start}-{end}")
        identity = _reviewed_identity(item)
        assert identity is not None
        original_text = _normalise_edit_text(reviewed_paragraphs[identity.casefold()]["original_text"])
        source_expected = _normalise_edit_text(str(item.get("source_text", old_segment)))
        if source_expected and source_expected not in original_text and original_text not in source_expected:
            raise RuntimeError(f"reviewed deleted text does not match source paragraph: {identity}")
        if not replacement.endswith("\n"):
            replacement += "\n"
        reviewed_identity = item.get("reviewed_paragraph", item.get("reviewed_id"))
        reviewed_hash = item.get("reviewed_hash", item.get("accepted_hash"))
        if reviewed_hash is None and isinstance(reviewed_identity, Mapping):
            reviewed_hash = reviewed_identity.get("hash")
        if reviewed_hash is not None and str(reviewed_hash) != _sha256_text(replacement):
            raise RuntimeError(f"reviewed section hash mismatch for {target}:{start}-{end}")
        if target not in staged_by_original:
            staged = staged_root / target.relative_to(source_dir)
            staged_by_original[target] = staged
        plans.append((target, start, end, replacement, dict(item)))
    if entry.get("inputs") is not None:
        logical = _resolve_inputs(entry["inputs"], base, label="inputs")
        for input_path in logical:
            input_target = _confined(input_path, source_dir, label="delivery input")
            if input_target not in staged_by_original:
                raise ValueError(f"delivery input is absent from the revision map: {input_target}")
    else:
        logical = list(staged_by_original)
    # Copy only mapped source files. In particular, do not recurse through a
    # project checkout and accidentally stage .git, virtualenvs, or unrelated
    # private files. Every resolved source path remains confined to source_dir.
    for target, staged in staged_by_original.items():
        target = _confined(target, source_dir, label="mapped source")
        if not target.is_file():
            raise FileNotFoundError(f"mapped Markdown source does not exist: {target}")
        staged.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(target, staged)
    # Preserve local media links from both the source segment and the reviewed
    # candidate. Candidate media is extracted by docx_to_markdown below its
    # scratch directory; copy it into the staged source-relative location.
    for target, staged in staged_by_original.items():
        source_text = target.read_text(encoding="utf-8-sig")
        target_media = _local_media_references(source_text)
        for _map_target, _start, _end, replacement, _item in plans:
            if _map_target == target:
                target_media.extend(_local_media_references(replacement))
        for reference in target_media:
            relative_ref = Path(reference)
            if relative_ref.is_absolute() or ".." in relative_ref.parts:
                raise ValueError(f"local media reference escapes source directory: {reference}")
            extracted_candidate = full_md.parent / relative_ref
            source_candidate = target.parent / relative_ref
            if extracted_candidate.is_file():
                media_source = _confined(extracted_candidate, full_md.parent, label="extracted media")
            elif source_candidate.is_file():
                media_source = _confined(source_candidate, source_dir, label="source media")
            else:
                raise FileNotFoundError(f"local media referenced by reviewed Markdown does not exist: {reference}")
            media_destination = staged_root / target.relative_to(source_dir).parent / relative_ref
            media_destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(media_source, media_destination)
    # Every range is interpreted against the immutable source snapshot. Apply
    # from the end of each file so replacing a multi-line range cannot shift
    # the coordinates of an earlier range in the same source.
    for target in sorted({plan[0] for plan in plans}, key=str):
        staged = staged_by_original[target]
        staged_lines = staged.read_text(encoding="utf-8-sig").splitlines(keepends=True)
        target_plans = sorted((plan for plan in plans if plan[0] == target), key=lambda plan: plan[1], reverse=True)
        for _target, start, end, replacement, item in target_plans:
            staged_lines[start - 1:end] = [replacement]
            identity = _reviewed_identity(item)
            assert identity is not None
            reviewed_record = reviewed_paragraphs[identity.casefold()]
            reviewed_value = item.get("reviewed_paragraph", item.get("reviewed_id"))
            reviewed_hash = item.get("reviewed_hash", item.get("accepted_hash"))
            if reviewed_hash is None and isinstance(reviewed_value, Mapping):
                reviewed_hash = reviewed_value.get("hash")
            records.append({
                "path": str(target), "start_line": start, "end_line": start + replacement.count("\n") - 1,
                "hash": _sha256_text(replacement), "original_hash": str(item.get("hash", item.get("sha256", ""))),
                "file": str(item["file"]),
                "reviewed_start_line": int(item["reviewed_start_line"]),
                "reviewed_end_line": int(item["reviewed_end_line"]),
                "reviewed_paragraph": reviewed_value,
                "reviewed_hash": reviewed_hash,
                "reviewed_text": replacement,
                "docx_paragraph_id": reviewed_record["id"],
                "docx_paragraph_range": {
                    "start": int(item["reviewed_start_line"]),
                    "end": int(item["reviewed_end_line"]),
                },
                "reviewed_docx_paragraph": {
                    "id": reviewed_record["id"],
                    "ordinal": reviewed_record["ordinal"],
                    "revision_kinds": sorted(reviewed_record["revisions"]),
                },
            })
        staged.write_text("".join(staged_lines), encoding="utf-8")
    if not staged_by_original:
        raise ValueError("section map produced no source updates")
    staged_inputs = [staged_by_original.get(path.resolve(), path) for path in logical]
    return _PreparedReview(tuple(staged_inputs), tuple(logical), tuple(staged_by_original.items()), tuple(records))


def _rewrite_paths(value: Any, replacements: Mapping[str, str]) -> Any:
    if isinstance(value, str):
        return replacements.get(value, value)
    if isinstance(value, list):
        return [_rewrite_paths(item, replacements) for item in value]
    if isinstance(value, dict):
        return {key: _rewrite_paths(item, replacements) for key, item in value.items()}
    return value


def _post_review_source_map(
    review: _PreparedReview,
    *,
    output_paragraphs: Mapping[tuple[str, int], Mapping[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Describe the staged post-review bytes under their published paths."""
    logical_by_staged = {str(staged.resolve()): logical for logical, staged in review.updates}
    result: list[dict[str, Any]] = []
    for staged in review.inputs:
        records = _source_map([staged])
        logical = logical_by_staged.get(str(staged.resolve()))
        for record in records:
            if logical is not None:
                record["path"] = str(logical.resolve())
                record["source"] = str(logical)
            record["provenance"] = "reviewed-docx-accepted"
            for changed in review.source_map:
                if (
                    changed.get("path") == str(logical or staged)
                    and changed.get("start_line") == record.get("start_line")
                ):
                    for key in (
                        "original_hash", "file", "reviewed_start_line", "reviewed_end_line",
                        "reviewed_paragraph", "reviewed_hash", "reviewed_text",
                        "docx_paragraph_id", "docx_paragraph_range", "reviewed_docx_paragraph",
                    ):
                        if key in changed:
                            record[key] = changed[key]
                    record.setdefault("file", str(logical or staged))
                    if output_paragraphs:
                        output = output_paragraphs.get((str(logical or staged), int(record["start_line"])))
                        if output is not None:
                            record["docx_paragraph_id"] = output["id"]
                            record["docx_paragraph_range"] = {
                                "start": output["start"], "end": output["end"],
                            }
        result.extend(records)
    return result


def _commit_transaction(pairs: Sequence[tuple[Path, Path]], *, force: bool) -> None:
    unique, seen = [], set()
    for source, destination in pairs:
        destination = destination.resolve()
        if destination in seen:
            continue
        seen.add(destination)
        unique.append((source, destination))
    backup_root = Path(tempfile.mkdtemp(prefix="docforge-commit-backup-"))
    backups, created = {}, []
    try:
        for _source, destination in unique:
            if destination.exists():
                if not force:
                    raise FileExistsError(f"Output exists; pass --force to overwrite: {destination}")
                backup = backup_root / str(len(backups))
                shutil.copy2(destination, backup)
                backups[destination] = backup
        for source, destination in unique:
            destination.parent.mkdir(parents=True, exist_ok=True)
            if not destination.exists():
                created.append(destination)
            try:
                os.replace(source, destination)
            except OSError:
                shutil.copy2(source, destination)
                source.unlink(missing_ok=True)
    except Exception:
        for destination in created:
            destination.unlink(missing_ok=True)
        for destination, backup in backups.items():
            shutil.copy2(backup, destination)
        raise
    finally:
        shutil.rmtree(backup_root, ignore_errors=True)


def deliver(profile: Mapping[str, Any] | Path | str, *, accept_revisions: bool = False, force: bool = False) -> DeliveryResult:
    """Build all artifacts and publish outputs and source updates transactionally."""
    payload, base, source = _load_profile(profile)
    accept_revisions = accept_revisions or bool(payload.get("accept_revisions", False))
    delivery_dir = _path(payload.get("delivery_dir", "delivery"), base, label="delivery_dir").resolve()
    shared = payload.get("shared", {})
    if not isinstance(shared, Mapping):
        raise ValueError("delivery profile 'shared' must be an object")
    entries = _entries(payload)
    entries.sort(key=lambda item: (0 if item[0] == "article" else 1))
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    if accept_revisions and not any(entry.get("reviewed_docx") for _key, entry in entries):
        raise ValueError("--accept-revisions requires reviewed_docx/reviewed input in the delivery profile")
    output_targets: set[Path] = set()

    def register_delivery_target(path: Path, *, label: str) -> Path:
        """Register every eventual delivery path before publishing anything."""
        resolved = _confined(path, delivery_dir, label=label)
        if resolved in output_targets:
            raise ValueError(f"delivery targets collide at output path: {path}")
        output_targets.add(resolved)
        return resolved

    aggregate_rel = _safe_relative(payload.get("manifest", "delivery.manifest.json"), label="delivery manifest")
    register_delivery_target(delivery_dir / aggregate_rel, label="delivery manifest")
    for key, entry in entries:
        reviewed = entry.get("reviewed_docx")
        if reviewed is not None:
            reviewed_path = _path(reviewed, base, label="reviewed_docx")
            if not reviewed_path.is_file():
                raise FileNotFoundError(reviewed_path)
            if not accept_revisions and _docx_has_revisions(reviewed_path):
                raise ValueError(f"{key} reviewed DOCX contains unapplied revisions; pass --accept-revisions")
        output = delivery_dir / _output_name(entry, key, payload, timestamp=timestamp)
        register_delivery_target(output, label=f"delivery {key}")
        register_delivery_target(output.with_suffix(".manifest.json"), label=f"delivery {key} manifest")
        register_delivery_target(output.with_suffix(".sha256"), label=f"delivery {key} checksum")
        if output.exists() and not force:
            raise FileExistsError(f"Output exists; pass --force to overwrite: {output}")
        template = _path(entry.get("template", shared.get("template")), base, label=f"{key}.template")
        if not template.is_file():
            raise FileNotFoundError(template)
        for field in ("metadata", "metadata_path", "bibliography"):
            value = entry.get(field, shared.get(field))
            if value is not None and not _path(value, base, label=f"{key}.{field}").is_file():
                raise FileNotFoundError(_path(value, base, label=f"{key}.{field}"))
    delivery_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".docforge-delivery-", dir=delivery_dir.parent) as temporary:
        stage_root, stage_delivery = Path(temporary), Path(temporary) / "delivery"
        stage_delivery.mkdir()
        source_updates: dict[Path, Path] = {}
        source_targets: set[Path] = set()
        prepared = []
        for key, raw_entry in entries:
            entry = dict(raw_entry)
            review = _prepare_review(entry, base, stage_root) if accept_revisions and entry.get("reviewed_docx") else None
            if review is not None:
                for logical, staged in review.updates:
                    destination = logical.resolve()
                    if destination in output_targets:
                        raise ValueError(f"delivery source update collides at output path: {logical}")
                    if destination in source_targets:
                        raise ValueError(f"delivery source updates collide at path: {logical}")
                    source_targets.add(destination)
                    source_updates[logical] = staged
                inputs = list(review.inputs)
            else:
                if entry.get("inputs") is None:
                    raise ValueError(f"delivery artifact '{key}' requires inputs")
                review, inputs = None, _resolve_inputs(entry["inputs"], base, label=f"{key}.inputs")
            prepared.append((key, {**entry, "_inputs": inputs}, review, _output_name(entry, key, payload, timestamp=timestamp)))
        built, source_records, article_manifest = [], [], None
        stage_manifests: list[Path] = []
        replacements: dict[str, str] = {}
        for key, entry, review, output_rel in prepared:
            inputs = list(entry["_inputs"])
            template = _path(entry.get("template", shared.get("template")), base, label=f"{key}.template")
            metadata_value = entry.get("metadata", entry.get("metadata_path", shared.get("metadata")))
            metadata = _path(metadata_value, base, label=f"{key}.metadata") if metadata_value is not None else None
            bibliography_value = entry.get("bibliography", shared.get("bibliography"))
            bibliography = _path(bibliography_value, base, label=f"{key}.bibliography") if bibliography_value is not None else None
            citation_value = entry.get("citation_base", shared.get("citation_base", payload.get("citation_base")))
            if key == "supplement" and citation_value is None and article_manifest is not None:
                citation_value = article_manifest
            citation_base = _path(citation_value, base, label=f"{key}.citation_base") if citation_value is not None else None
            options = dict(shared.get("options", {})) if isinstance(shared.get("options", {}), Mapping) else {}
            if isinstance(payload.get("options"), Mapping):
                options.update(payload["options"])
            if isinstance(entry.get("options"), Mapping):
                options.update(entry["options"])
            for option in _OPTION_KEYS:
                if option in entry:
                    options[option] = entry[option]
            stage_output = stage_delivery / output_rel
            stage_output.parent.mkdir(parents=True, exist_ok=True)
            result: AssemblyResult = assemble_markdown_template(inputs, template_path=template, output=stage_output, metadata_path=metadata, bibliography_path=bibliography, citation_base_path=citation_base, force=True, **options)
            output_paragraphs = _stamp_output_paragraph_ids(stage_output, review)
            stage_manifest, stage_checksum = write_assembly_sidecars(result, inputs=inputs, template_path=template, metadata_path=metadata, bibliography_path=bibliography, citation_base_path=citation_base, command=["docforge", "deliver", "--profile", source, *( ["--accept-revisions"] if accept_revisions else [] )])
            stage_manifests.append(stage_manifest)
            final_output = delivery_dir / output_rel
            final_manifest, final_checksum = final_output.with_suffix(".manifest.json"), final_output.with_suffix(".sha256")
            replacements.update({str(stage_output.resolve()): str(final_output), str(stage_manifest.resolve()): str(final_manifest), str(stage_checksum.resolve()): str(final_checksum)})
            if review is not None:
                replacements.update({str(staged.resolve()): str(logical.resolve()) for logical, staged in review.updates})
            built.append(DeliveryArtifact(key, final_output, final_manifest, final_checksum, tuple(review.logical_inputs if review else inputs)))
            source_records.extend(_post_review_source_map(review, output_paragraphs=output_paragraphs) if review is not None else _source_map(inputs))
            if key == "article":
                # The supplement needs a readable citation base while both
                # artifacts are still staged. It is rewritten to the final
                # path below before publication.
                article_manifest = stage_manifest
        aggregate = stage_delivery / aggregate_rel
        final_aggregate = delivery_dir / aggregate_rel
        validate_output_path(final_aggregate, label="delivery manifest")
        review_sections = [item for item in source_records if item.get("reviewed_paragraph") is not None]
        aggregate.write_text(json.dumps({"schema": "docforge.delivery.v1", "generated_at_utc": datetime.now(timezone.utc).isoformat(), "profile": source, "accept_revisions": accept_revisions, "delivery_dir": str(delivery_dir), "artifacts": [{"name": x.name, "output": str(x.output), "manifest": str(x.manifest), "checksum": str(x.checksum), "inputs": [str(p) for p in x.inputs]} for x in built], "source_paragraph_map": source_records, "sections": review_sections}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        for manifest in stage_manifests:
            data = json.loads(manifest.read_text(encoding="utf-8"))
            manifest.write_text(json.dumps(_rewrite_paths(data, replacements), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        pairs = [(file, delivery_dir / file.relative_to(stage_delivery)) for file in stage_delivery.rglob("*") if file.is_file()]
        pairs.extend((staged, logical) for logical, staged in source_updates.items())
        _commit_transaction(pairs, force=force)
        return DeliveryResult(delivery_dir, tuple(built), final_aggregate)


__all__ = ["DeliveryArtifact", "DeliveryResult", "deliver"]
