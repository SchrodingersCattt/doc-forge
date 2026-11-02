"""Conservative, source-backed DOCX/TeX review round trips.

Only uniquely matched, plain, one-line paragraphs are editable. Unsupported
review changes and edits outside those paragraphs fail before publication.
The private bundle stores the original packages and hashes for audit; an
unchanged TeX document reuses the reviewed DOCX byte for byte.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import shutil
import tempfile
import zipfile
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

from lxml import etree

from ..adapters.docx import audit_docx
from ..docxdiff.package import Package
from ..docxdiff.redline import W, _accepted_revision_view, visible_text
from ..roundtrip import ensure_distinct_paths

_SCHEMA = "docforge.tex-review-roundtrip.v1"
_W = f"{{{W}}}"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_FORBIDDEN_TEX = frozenset("\\{}$#%&_^~\r\n")
_COMMENTS = ("word/comments", "word/people.xml")
_NONCONTENT_PARTS = frozenset({"word/settings.xml", "word/webSettings.xml"})
_ALLOWED_CHILDREN = frozenset({"pPr", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd"})
_ALLOWED_RUN_CHILDREN = frozenset({"rPr", "t", "commentReference"})


class ReviewRoundtripError(ValueError):
    """A review or TeX edit is unsafe to map without losing information."""


def _digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _c14n(node: etree._Element | None) -> bytes:
    return b"" if node is None else etree.tostring(node, method="c14n", exclusive=True)


def _root(path: Path) -> etree._Element:
    return _accepted_revision_view(Package.load(path).xml("word/document.xml"))


def _validate_source_docx(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise ReviewRoundtripError("review inputs must be DOCX files")
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)) or archive.testzip() is not None:
            raise ReviewRoundtripError("invalid or ambiguous DOCX package")
    if not audit_docx(path).ok:
        raise ReviewRoundtripError("review DOCX failed package audit")


def _children(root: etree._Element) -> list[etree._Element]:
    body = root.find(_W + "body")
    if body is None:
        raise ReviewRoundtripError("DOCX has no document body")
    return list(body)


def _text_runs(paragraph: etree._Element) -> list[tuple[str, bytes]]:
    """Project each character's explicit run properties for comparison."""
    runs: list[tuple[str, bytes]] = []
    for node in paragraph.iter(_W + "t"):
        ancestor = node.getparent()
        while ancestor is not None and ancestor is not paragraph and ancestor.tag != _W + "r":
            ancestor = ancestor.getparent()
        props = ancestor.find(_W + "rPr") if ancestor is not None and ancestor.tag == _W + "r" else None
        runs.append((node.text or "", _c14n(props)))
    return runs


def _styles(paragraph: etree._Element) -> list[bytes]:
    return [style for text, style in _text_runs(paragraph) for _ in text]


def _simple(paragraph: etree._Element) -> bool:
    if paragraph.tag != _W + "p":
        return False
    for child in paragraph:
        if not isinstance(child.tag, str) or etree.QName(child).namespace != W:
            return False
        name = etree.QName(child).localname
        if name not in _ALLOWED_CHILDREN:
            return False
        if name == "r":
            for item in child:
                if not isinstance(item.tag, str) or etree.QName(item).namespace != W:
                    return False
                if etree.QName(item).localname not in _ALLOWED_RUN_CHILDREN:
                    return False
    return True


def _protected(paragraph: etree._Element) -> list[bytes]:
    """Compare complex content; comment anchors may legitimately be added."""
    ignored = {"pPr", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd"}
    result: list[bytes] = []
    for child in paragraph:
        if not isinstance(child.tag, str):
            result.append(_c14n(child))
            continue
        name = etree.QName(child).localname
        if name in {"bookmarkStart", "bookmarkEnd"}:
            result.append(_c14n(child))
        elif name not in ignored:
            result.append(_c14n(child))
        elif name == "r":
            for item in child:
                if isinstance(item.tag, str) and etree.QName(item).localname not in {"rPr", "t", "commentReference"}:
                    result.append(_c14n(item))
    return result


def _relationship_bytes(data: bytes) -> tuple[bytes, ...]:
    root = etree.fromstring(data)
    return tuple(sorted(
        _c14n(item) for item in root
        if "comment" not in item.get("Type", "").lower()
        and not item.get("Type", "").lower().endswith("/people")
    ))


def _check_other_parts(before: Path, after: Path) -> None:
    with zipfile.ZipFile(before) as a, zipfile.ZipFile(after) as b:
        def considered(name: str) -> bool:
            return (name.startswith("word/") and name != "word/document.xml"
                    and not name.startswith(_COMMENTS) and name not in _NONCONTENT_PARTS)

        left = {name for name in a.namelist() if considered(name)}
        right = {name for name in b.namelist() if considered(name)}
        if left != right:
            raise ReviewRoundtripError("unsupported change to DOCX package parts")
        for name in left:
            old, new = a.read(name), b.read(name)
            if name == "word/_rels/document.xml.rels":
                if _relationship_bytes(old) != _relationship_bytes(new):
                    raise ReviewRoundtripError("unsupported relationship change")
            elif old != new:
                raise ReviewRoundtripError("unsupported change to DOCX package parts")


def _source_span(source: str, text: str) -> tuple[int, int] | None:
    if not text or any(char in _FORBIDDEN_TEX for char in text):
        return None
    start = source.find(text)
    if start < 0 or source.find(text, start + 1) >= 0:
        return None
    end = start + len(text)
    line_start = source.rfind("\n", 0, start) + 1
    line_end = source.find("\n", end)
    if line_end < 0:
        line_end = len(source)
    if source[line_start:start].strip() or source[end:line_end].strip():
        return None
    # Plain lines inside listings, floats, equations, or arbitrary environments
    # are not prose anchors. Do not claim to expand macros or TeX conditionals.
    prefix = source[:line_start]
    document_start = prefix.find(r"\begin{document}")
    if document_start < 0:
        return None
    body_prefix = prefix[document_start:]
    if re.search(r"\\(?:if(?:false|true)|else|fi|catcode)\b", body_prefix):
        return None
    stack: list[str] = []
    for line in body_prefix.splitlines():
        line = re.split(r"(?<!\\)%", line, maxsplit=1)[0]
        for marker in re.finditer(r"\\(begin|end)\{([^}]+)\}", line):
            action, name = marker.groups()
            if action == "begin":
                stack.append(name)
            elif not stack or stack.pop() != name:
                return None
    if stack != ["document"]:
        return None
    return start, end


def _check_paragraphs(old: etree._Element, new: etree._Element, index: int) -> tuple[str, str]:
    old_text, new_text = visible_text(old), visible_text(new)
    if _c14n(old.find(_W + "pPr")) != _c14n(new.find(_W + "pPr")):
        raise ReviewRoundtripError(f"paragraph {index}: paragraph formatting changed")
    if _protected(old) != _protected(new):
        raise ReviewRoundtripError(f"paragraph {index}: field, bookmark, drawing or other object changed")
    old_styles, new_styles = _styles(old), _styles(new)
    if len(old_styles) != len(old_text) or len(new_styles) != len(new_text):
        if _c14n(old) != _c14n(new):
            raise ReviewRoundtripError(f"paragraph {index}: complex text cannot be mapped")
        return old_text, new_text
    if old_text == new_text:
        if old_styles != new_styles:
            raise ReviewRoundtripError(f"paragraph {index}: run formatting changed")
        return old_text, new_text
    if not _simple(old) or not _simple(new):
        raise ReviewRoundtripError(f"paragraph {index}: only simple text changes are supported")
    if len(set(old_styles + new_styles)) > 1:
        raise ReviewRoundtripError(f"paragraph {index}: mixed formatting is not supported")
    if SequenceMatcher(None, old_text, new_text, autojunk=False).ratio() < 0.5:
        raise ReviewRoundtripError(f"paragraph {index}: change is too ambiguous to align")
    if any(char in _FORBIDDEN_TEX for char in new_text):
        raise ReviewRoundtripError(f"paragraph {index}: text requires TeX escaping")
    return old_text, new_text


def export_review_bundle(baseline: Path, reviewed: Path, source_tex: Path, workdir: Path) -> dict:
    """Create an auditable bundle after applying supported reviewer edits to TeX.

    Unmatched changes, changed complex objects, reordered blocks, duplicate
    source matches and formatting changes abort without publishing a bundle.
    """
    baseline, reviewed, source_tex, workdir = map(Path, (baseline, reviewed, source_tex, workdir))
    if workdir.exists():
        raise FileExistsError(f"bundle already exists: {workdir}")
    if any(path.resolve() == workdir.resolve() for path in (baseline, reviewed, source_tex)):
        raise ReviewRoundtripError("bundle collides with an input")
    _validate_source_docx(baseline)
    _validate_source_docx(reviewed)
    source_bytes = source_tex.read_bytes()
    try:
        source = source_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ReviewRoundtripError("source TeX must be UTF-8") from exc
    _check_other_parts(baseline, reviewed)
    old_children, new_children = _children(_root(baseline)), _children(_root(reviewed))
    if len(old_children) != len(new_children) or any(a.tag != b.tag for a, b in zip(old_children, new_children)):
        raise ReviewRoundtripError("inserted, deleted or moved DOCX blocks are unsupported")
    paragraphs: list[tuple[int, str, str, bool]] = []
    for index, (old, new) in enumerate(zip(old_children, new_children)):
        if old.tag != _W + "p":
            if _c14n(old) != _c14n(new):
                raise ReviewRoundtripError(f"block {index}: non-paragraph content changed")
            continue
        old_text, new_text = _check_paragraphs(old, new, index)
        paragraphs.append((index, old_text, new_text, _simple(old) and _simple(new)))
    changed = [(i, a, b) for i, a, b, _simple_value in paragraphs if a != b]
    # With more than one changed paragraph, Word's durable paragraph ids are
    # required to rule out silent positional pairing after a reorder.
    if len(changed) > 1:
        key = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
        ids = [(old_children[i].get(key), new_children[i].get(key)) for i, _, _ in changed]
        if any(not a or a != b for a, b in ids) or len({a for a, _ in ids}) != len(ids):
            raise ReviewRoundtripError("multiple changes require stable Word paragraph ids")
    counts = Counter(text for _, text, _, simple in paragraphs if simple and text)
    candidates = []
    for index, old_text, new_text, simple in paragraphs:
        span = _source_span(source, old_text) if simple and counts[old_text] == 1 else None
        if span is None:
            if old_text != new_text:
                raise ReviewRoundtripError(f"paragraph {index}: no unique plain TeX source span")
            continue
        candidates.append((span[0], span[1], index, old_text, new_text))
    candidates.sort()
    if any(a[1] > b[0] for a, b in zip(candidates, candidates[1:])):
        raise ReviewRoundtripError("source paragraph spans overlap")
    pieces, blocks, cursor, offset = [], [], 0, 0
    for start, end, index, old_text, new_text in candidates:
        pieces.extend((source[cursor:start], new_text))
        new_start = start + offset
        blocks.append({"body_index": index, "start": new_start, "end": new_start + len(new_text),
                       "text_sha256": _digest(new_text.encode("utf-8"))})
        offset += len(new_text) - len(old_text)
        cursor = end
    pieces.append(source[cursor:])
    mapped = "".join(pieces).encode("utf-8")
    manifest = {
        "schema": _SCHEMA,
        "sha256": {"baseline.docx": _digest(baseline.read_bytes()),
                   "reviewed.docx": _digest(reviewed.read_bytes()),
                   "original.tex": _digest(source_bytes), "baseline.tex": _digest(mapped)},
        "blocks": blocks,
        "reviewed_changed_paragraphs": len(changed),
    }
    workdir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".tex-review-", dir=workdir.parent) as temporary:
        stage = Path(temporary) / "bundle"
        stage.mkdir()
        shutil.copyfile(baseline, stage / "baseline.docx")
        shutil.copyfile(reviewed, stage / "reviewed.docx")
        (stage / "original.tex").write_bytes(source_bytes)
        (stage / "baseline.tex").write_bytes(mapped)
        (stage / "document.tex").write_bytes(mapped)
        (stage / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
        for name, expected in manifest["sha256"].items():
            if _digest((stage / name).read_bytes()) != expected:
                raise ReviewRoundtripError(f"input changed while staging: {name}")
        os.replace(stage, workdir)
    return manifest


def _visible_nodes(paragraph: etree._Element) -> list[etree._Element]:
    result = []
    for node in paragraph.iter(_W + "t"):
        if any(parent.tag in {_W + "del", _W + "moveFrom"} for parent in node.iterancestors() if parent is not paragraph):
            continue
        result.append(node)
    return result


def _patch_text(paragraph: etree._Element, expected: str, replacement: str, index: int) -> None:
    nodes = _visible_nodes(paragraph)
    if "".join(node.text or "" for node in nodes) != expected:
        raise ReviewRoundtripError(f"paragraph {index}: visible text changed since export")
    offsets = []
    pos = 0
    for node in nodes:
        end = pos + len(node.text or "")
        offsets.append((pos, end, node))
        pos = end
    # One contiguous replacement avoids splitting a word at coincidental
    # character matches and never crosses a run or comment boundary.
    prefix = 0
    while prefix < min(len(expected), len(replacement)) and expected[prefix] == replacement[prefix]:
        prefix += 1
    suffix = 0
    while (suffix < len(expected) - prefix and suffix < len(replacement) - prefix
           and expected[len(expected) - 1 - suffix] == replacement[len(replacement) - 1 - suffix]):
        suffix += 1
    old_end = len(expected) - suffix
    new_end = len(replacement) - suffix
    matches = [(start, end, node) for start, end, node in offsets if start <= prefix <= old_end <= end]
    if len(matches) != 1:
        raise ReviewRoundtripError(f"paragraph {index}: edit crosses a run or anchor boundary")
    start, _end, node = matches[0]
    text = node.text or ""
    node.text = text[:prefix - start] + replacement[prefix:new_end] + text[old_end - start:]
    if node.text[:1].isspace() or node.text[-1:].isspace():
        node.set(_XML_SPACE, "preserve")
    if visible_text(_accepted_revision_view(paragraph)) != replacement:
        raise ReviewRoundtripError(f"paragraph {index}: patched text failed verification")


def build_reviewed_docx(workdir: Path, output: Path, *, force: bool = False) -> dict:
    """Rebuild a reviewed DOCX from the bundle's current document.tex."""
    workdir, output = Path(workdir), Path(output)
    manifest = json.loads((workdir / "manifest.json").read_text(encoding="utf-8"))
    if manifest.get("schema") != _SCHEMA:
        raise ReviewRoundtripError("unsupported review bundle schema")
    paths = [workdir / name for name in ("baseline.docx", "reviewed.docx", "original.tex", "baseline.tex", "document.tex", "manifest.json")]
    ensure_distinct_paths(output, *paths)
    if output.exists() and not force:
        raise FileExistsError(f"output exists: {output}")
    hashes = manifest.get("sha256", {})
    for name in ("baseline.docx", "reviewed.docx", "original.tex", "baseline.tex"):
        if _digest((workdir / name).read_bytes()) != hashes.get(name):
            raise ReviewRoundtripError(f"bundle snapshot hash mismatch: {name}")
    base = (workdir / "baseline.tex").read_bytes().decode("utf-8")
    edited = (workdir / "document.tex").read_bytes().decode("utf-8")
    reviewed = workdir / "reviewed.docx"
    output.parent.mkdir(parents=True, exist_ok=True)
    if base == edited:
        descriptor, name = tempfile.mkstemp(prefix=".tex-review-", suffix=".docx", dir=output.parent)
        os.close(descriptor)
        temp = Path(name)
        try:
            shutil.copyfile(reviewed, temp)
            if _digest(temp.read_bytes()) != hashes["reviewed.docx"]:
                raise ReviewRoundtripError("exact DOCX reuse verification failed")
            os.replace(temp, output)
        finally:
            temp.unlink(missing_ok=True)
        return {"mode": "EXACT_REUSE", "changed_paragraphs": 0}
    blocks = manifest.get("blocks")
    if not isinstance(blocks, list):
        raise ReviewRoundtripError("bundle has no paragraph map")
    previous_end = 0
    for block in blocks:
        start, end = block["start"], block["end"]
        if not (isinstance(start, int) and isinstance(end, int) and previous_end <= start < end <= len(base)):
            raise ReviewRoundtripError("invalid paragraph map ranges")
        if _digest(base[start:end].encode("utf-8")) != block["text_sha256"]:
            raise ReviewRoundtripError("paragraph map hash mismatch")
        previous_end = end
    changes: dict[int, list[tuple[int, int, str]]] = {}
    for tag, a, b, c, d in SequenceMatcher(None, base, edited, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        candidates = [block for block in blocks if block["start"] <= a <= b <= block["end"]]
        if len(candidates) != 1:
            raise ReviewRoundtripError("TeX edit is outside one uniquely mapped paragraph")
        block = candidates[0]
        value = edited[c:d]
        if any(char in _FORBIDDEN_TEX for char in value):
            raise ReviewRoundtripError("TeX edit introduces unsupported syntax")
        changes.setdefault(block["body_index"], []).append((a - block["start"], b - block["start"], value))
    package = Package.load(reviewed)
    raw_root = package.xml("word/document.xml")
    original_root = copy.deepcopy(raw_root)
    children = _children(raw_root)
    accepted_children = _children(_accepted_revision_view(Package.load(reviewed).xml("word/document.xml")))
    by_index = {block["body_index"]: block for block in blocks}
    for index, operations in changes.items():
        if not isinstance(index, int) or index < 0 or index >= len(children) or len(children) != len(accepted_children):
            raise ReviewRoundtripError("DOCX paragraph anchor is invalid")
        block = by_index[index]
        original = base[block["start"]:block["end"]]
        if visible_text(accepted_children[index]) != original:
            raise ReviewRoundtripError(f"paragraph {index}: reviewed source no longer matches the map")
        replacement = original
        for start, end, value in reversed(operations):
            replacement = replacement[:start] + value + replacement[end:]
        _patch_text(children[index], original, replacement, index)
    # Verify the complete XML tree apart from text in the mapped paragraphs.
    # Run/style properties, fields, drawings, markers and all other content
    # must have remained structurally identical.
    projected_original, projected_patched = copy.deepcopy(original_root), copy.deepcopy(raw_root)
    for tree in (projected_original, projected_patched):
        projection_children = _children(tree)
        for index in changes:
            for node in projection_children[index].iter(_W + "t"):
                node.text = ""
                node.attrib.pop(_XML_SPACE, None)
    if _c14n(projected_original) != _c14n(projected_patched):
        raise ReviewRoundtripError("patch changed content outside mapped text runs")
    package.set_xml("word/document.xml", raw_root)
    descriptor, name = tempfile.mkstemp(prefix=".tex-review-", suffix=".docx", dir=output.parent)
    os.close(descriptor)
    temp = Path(name)
    try:
        package.write(temp, overwrite=True)
        result = audit_docx(temp)
        if not result.ok:
            raise ReviewRoundtripError("patched DOCX failed package audit")
        os.replace(temp, output)
    finally:
        temp.unlink(missing_ok=True)
    return {"mode": "SOURCE_PATCH", "changed_paragraphs": len(changes)}


__all__ = ["ReviewRoundtripError", "export_review_bundle", "build_reviewed_docx"]
