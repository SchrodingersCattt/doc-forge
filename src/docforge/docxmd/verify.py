"""Round-trip checks for the DOCX <-> Markdown bundle.

:func:`docx_model` describes what a reader of the document sees: block
order, paragraph properties, every character with its resolved visible
formatting, tracked changes, comments, pictures, tables and the page setup.
Two DOCX files with equal models render the same text the same way.
"""

from __future__ import annotations

import difflib
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path, PurePosixPath

from lxml import etree

from ..docxdiff.package import Package
from . import drawing as drawing_mod
from .common import clean_tokens, top_name
from .ooxml import KNOWN_NAMESPACES, PKG_REL, R, W, flatten
from .reader import kept_bookmarks
from .styles import StyleSheet

W15 = KNOWN_NAMESPACES["w15"]
_SKIP = {"bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd", "lastRenderedPageBreak"}
_SHARED_PARTS = (
    "word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/fontTable.xml",
    "word/theme/theme1.xml", "word/footnotes.xml", "word/endnotes.xml",
)


def _local(node) -> str:
    return etree.QName(node).localname if isinstance(node.tag, str) else ""


def _category(char: str) -> str:
    code = ord(char)
    if char.isspace():
        return "s"
    if "a" <= char.lower() <= "z":
        return "a"
    if code < 0x80:
        return "p"
    if 0x2E80 <= code:
        return "c"
    return "o"


class _Model:
    def __init__(self, path: Path) -> None:
        self.package = Package.load(path)
        self.root = self.package.xml("word/document.xml")
        styles = self.package.parts.get("word/styles.xml")
        self.styles = StyleSheet(etree.fromstring(styles) if styles else None)
        self.rels = {}
        rels = self.package.parts.get("word/_rels/document.xml.rels")
        if rels:
            for rel in etree.fromstring(rels).findall(f"{{{PKG_REL}}}Relationship"):
                self.rels[rel.get("Id")] = (rel.get("Type").rsplit("/", 1)[-1], rel.get("Target"))
        self.comment_order: dict[str, int] = {}
        self.cache: dict = {}

    def _target(self, rel_id: str) -> str:
        kind, target = self.rels.get(rel_id, ("?", rel_id))
        if kind == "image":
            part = re.sub(r"[^/]+/\.\./", "", str(PurePosixPath("word") / target))
            payload = self.package.parts.get(part, b"")
            return "image:" + hashlib.sha1(payload).hexdigest()[:16]
        return f"{kind}:{target}"

    def _comment(self, raw: str) -> int:
        return self.comment_order.setdefault(raw, len(self.comment_order))

    def _signature(self, node: etree._Element) -> tuple:
        attrs = []
        for key, value in node.attrib.items():
            local = etree.QName(key).localname
            if local.startswith("rsid") or local in {"paraId", "textId", "anchorId", "editId"}:
                continue
            if local == "id" and _local(node) in {"docPr", "cNvPr", "ins", "del", "pPrChange", "rPrChange"}:
                continue
            if etree.QName(key).namespace == R:
                value = self._target(value)
            attrs.append((key, value))
        return (node.tag, tuple(sorted(attrs)), (node.text or ""), tuple(
            self._signature(child) for child in node if isinstance(child.tag, str) and _local(child) not in _SKIP
        ))

    def visible(self, style, table_style, rpr, text):
        key = (style, table_style, tuple(rpr), text)
        if key not in self.cache:
            self.cache[key] = self.styles.visible(style, rpr, text, table_style)
        return self.cache[key]

    def paragraph(self, p: etree._Element, table_style: str | None = None) -> tuple:
        ppr = clean_tokens(flatten(p.find(f"./{{{W}}}pPr")))
        style = next((value for key, value in ppr if key == "pStyle@val"), None)
        mark = [(key[4:], value) for key, value in ppr if key.startswith("rPr.")]
        props = tuple(item for item in ppr if not item[0].startswith("rPr."))
        mark_visible = self.visible(style, table_style, mark, "a中")
        content: list[tuple] = []
        self._content(p, style, table_style, (), content)
        merged: list[tuple] = []
        for item in content:
            if merged and item[0] == "text" and merged[-1][0] == "text" and merged[-1][2:] == item[2:]:
                merged[-1] = ("text", merged[-1][1] + item[1]) + item[2:]
            else:
                merged.append(item)
        return ("p", props, mark_visible, tuple(merged))

    def _content(self, node, style, table_style, wrappers, out) -> None:
        for child in node:
            name = _local(child)
            if name in _SKIP or name in {"pPr", ""}:
                continue
            if name == "r":
                rpr = clean_tokens(flatten(child.find(f"./{{{W}}}rPr")))
                for item in child:
                    kind = _local(item)
                    if kind in _SKIP or kind in {"rPr", ""}:
                        continue
                    if kind in {"t", "delText"}:
                        text = item.text or ""
                        start = 0
                        for index in range(1, len(text) + 1):
                            if index == len(text) or _category(text[index]) != _category(text[start]):
                                chunk = text[start:index]
                                out.append(("text", chunk, self.visible(style, table_style, rpr, chunk), wrappers))
                                start = index
                    elif kind == "commentReference":
                        out.append(("comment-ref", self._comment(item.get(f"{{{W}}}id")), wrappers))
                    elif kind == "drawing":
                        described = drawing_mod.describe(item)
                        if described:
                            attrs, rel_id = described
                            out.append(("image", tuple(attrs), self._target(rel_id), wrappers))
                        else:
                            out.append(("raw", self._signature(item), wrappers))
                    else:
                        out.append(("atom", self._signature(item), self.visible(style, table_style, rpr, ""), wrappers))
            elif name in {"ins", "del", "moveTo", "moveFrom"}:
                kind = "ins" if name in {"ins", "moveTo"} else "del"
                mark = (kind, child.get(f"{{{W}}}author"), child.get(f"{{{W}}}date"))
                self._content(child, style, table_style, wrappers + (mark,), out)
            elif name == "hyperlink":
                rel = child.get(f"{{{R}}}id")
                href = self._target(rel) if rel else "#" + (child.get(f"{{{W}}}anchor") or "")
                self._content(child, style, table_style, wrappers + (("link", href),), out)
            elif name in {"commentRangeStart", "commentRangeEnd"}:
                out.append((name, self._comment(child.get(f"{{{W}}}id"))))
            elif name in {"smartTag", "customXml"}:
                self._content(child, style, table_style, wrappers, out)
            else:
                out.append(("raw", self._signature(child), wrappers))

    def table(self, tbl: etree._Element) -> tuple:
        tbl_pr = tuple(clean_tokens(flatten(tbl.find(f"./{{{W}}}tblPr"))))
        style = next((value for key, value in tbl_pr if key == "tblStyle@val"), None)
        grid = tuple(col.get(f"{{{W}}}w") for col in tbl.findall(f"./{{{W}}}tblGrid/{{{W}}}gridCol"))
        rows = []
        for tr in tbl.findall(f"./{{{W}}}tr"):
            cells = []
            for tc in tr.findall(f"./{{{W}}}tc"):
                cells.append((
                    tuple(clean_tokens(flatten(tc.find(f"./{{{W}}}tcPr")))),
                    tuple(self.paragraph(p, style) for p in tc.findall(f"./{{{W}}}p")),
                ))
            rows.append((tuple(clean_tokens(flatten(tr.find(f"./{{{W}}}trPr")))), tuple(cells)))
        return ("tbl", tbl_pr, grid, tuple(rows))

    def blocks(self) -> list[tuple]:
        body = self.root.find(f"./{{{W}}}body")
        result = []
        for child in body:
            name = _local(child)
            if name in _SKIP:
                continue
            if name == "p":
                result.append(self.paragraph(child))
            elif name == "tbl":
                result.append(self.table(child))
            elif name == "sectPr":
                result.append(("sectPr", self._signature(child)))
            else:
                result.append(("raw", self._signature(child)))
        background = self.root.find(f"./{{{W}}}background")
        if background is not None:
            result.insert(0, ("background", self._signature(background)))
        return result

    def comments(self) -> list[tuple]:
        if "word/comments.xml" not in self.package.parts:
            return []
        root = etree.fromstring(self.package.parts["word/comments.xml"])
        done: dict[str, str] = {}
        if "word/commentsExtended.xml" in self.package.parts:
            for item in etree.fromstring(self.package.parts["word/commentsExtended.xml"]).iter(f"{{{W15}}}commentEx"):
                done[item.get(f"{{{W15}}}paraId")] = item.get(f"{{{W15}}}done", "0")
        result = []
        for comment in root.findall(f"{{{W}}}comment"):
            raw = comment.get(f"{{{W}}}id")
            if raw not in self.comment_order:
                continue
            paragraphs = comment.findall(f"{{{W}}}p")
            last = paragraphs[-1].get(f"{{{KNOWN_NAMESPACES['w14']}}}paraId") if paragraphs else None
            result.append((
                self.comment_order[raw],
                comment.get(f"{{{W}}}author"), comment.get(f"{{{W}}}initials"), comment.get(f"{{{W}}}date"),
                done.get(last, "0"),
                tuple(self.paragraph(p) for p in paragraphs),
            ))
        return sorted(result)


def _block_label(block: tuple) -> str:
    def text_of(paragraph):
        return "".join(item[1] for item in paragraph[3] if item[0] == "text")

    if block[0] == "p":
        return "p: " + text_of(block)[:70]
    if block[0] == "tbl":
        return "table"
    return block[0]


def compare_docx(original: Path, rebuilt: Path) -> list[str]:
    """Differences a reader could notice between two DOCX files (empty when equal)."""
    left, right = _Model(original), _Model(rebuilt)
    a, b = left.blocks(), right.blocks()
    issues: list[str] = []
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        for offset in range(max(i2 - i1, j2 - j1)):
            old = a[i1 + offset] if i1 + offset < i2 else None
            new = b[j1 + offset] if j1 + offset < j2 else None
            issues.append(f"block {i1 + offset}: {_describe(old, new)}")
    if left.comments() != right.comments():
        issues.append("comments differ")
    kept_left = sorted(kept_bookmarks(left.root).values())
    kept_right = sorted(kept_bookmarks(right.root).values())
    if kept_left != kept_right:
        issues.append(f"bookmarks differ: {sorted(set(kept_left) ^ set(kept_right))}")
    for part in _SHARED_PARTS:
        if left.package.parts.get(part) != right.package.parts.get(part):
            issues.append(f"part {part} differs")
    return issues


def _describe(old: tuple | None, new: tuple | None) -> str:
    if old is None:
        return f"extra {_block_label(new)}"
    if new is None:
        return f"missing {_block_label(old)}"
    if old[0] != new[0]:
        return f"{_block_label(old)} became {_block_label(new)}"
    if old[0] == "p":
        if old[1] != new[1]:
            only_old = [item for item in old[1] if item not in new[1]]
            only_new = [item for item in new[1] if item not in old[1]]
            return f"paragraph properties {only_old} -> {only_new} in {_block_label(old)}"
        if old[2] != new[2]:
            return f"paragraph mark {old[2]} -> {new[2]} in {_block_label(old)}"
        for x, y in zip(old[3], new[3]):
            if x != y:
                return f"content {x!r} -> {y!r} in {_block_label(old)}"
        return f"content length {len(old[3])} -> {len(new[3])} in {_block_label(old)}"
    return f"{_block_label(old)} differs"


def redline_changes(original: Path, rebuilt: Path) -> dict[str, int]:
    from ..docxdiff.redline import create_tracked_docx

    with tempfile.TemporaryDirectory() as tmp:
        return create_tracked_docx(original, rebuilt, Path(tmp) / "redline.docx", author="roundtrip", workers=1)


def roundtrip(
    source: Path, work: Path, doc_id: str = "main", *, split: list[dict] | None = None,
    bibliography: str = "references.json", bib_style: str | None = None,
    output: Path | None = None, section_map: Path | None = None,
    track_changes: str = "accept", baseline: Path | None = None,
    redline: Path | None = None, force: bool = False,
) -> dict:
    """Round-trip a DOCX while preserving its source package.

    The legacy ``docxmd`` bundle path (``split``/``doc_id``) still exports and
    rebuilds through the reversible writer.  When ``section_map`` is supplied,
    the exact workflow used by ``docforge roundtrip`` is selected: DOCX-origin
    Markdown is exported with :func:`docx_to_markdown`, unchanged hashes copy
    the source package byte-for-byte, and edits are applied with
    :func:`apply_markdown_delta` so layout-bearing OOXML is retained.
    """
    if section_map is not None:
        return _exact_roundtrip(
            source, work, section_map=section_map, track_changes=track_changes,
            output=output, baseline=baseline, redline=redline, force=force,
        )
    from .reader import ExportOptions, export_docx
    from .writer import build_docx

    export = export_docx(source, work, ExportOptions(doc_id, split, bibliography, bib_style))
    output = output or work / f"_roundtrip/{doc_id}.docx"
    output.parent.mkdir(parents=True, exist_ok=True)
    build_docx(work, doc_id, output, overwrite=True)
    issues = compare_docx(source, output)
    summary = redline_changes(source, output)
    return {
        "export": {key: (dict(value) if hasattr(value, "items") else value) for key, value in export.items()},
        "rebuilt": str(output),
        "issues": issues,
        "redline": summary,
        "ok": not issues and summary.get("changed", 0) == summary.get("inserted", 0) == summary.get("deleted", 0) == 0,
    }


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _bundle_path(root: Path, relative: str, *, label: str) -> Path:
    """Resolve a manifest path while keeping it inside the bundle."""
    bundle = root.resolve()
    candidate = (bundle / relative).resolve()
    try:
        candidate.relative_to(bundle)
    except ValueError as exc:
        raise ValueError(f"roundtrip manifest {label} escapes the workdir: {relative}") from exc
    return candidate


def _exact_roundtrip(
    source: Path,
    work: Path,
    *,
    section_map: Path,
    track_changes: str,
    output: Path | None,
    baseline: Path | None,
    redline: Path | None,
    force: bool,
) -> dict:
    """Implementation for the package-preserving command workflow."""
    from ..markdown.docx_export import docx_to_markdown, reuse_unchanged_roundtrip_source
    from ..markdown.source_edit import apply_markdown_delta

    source, work = Path(source), Path(work)
    if not source.is_file():
        raise FileNotFoundError(source)
    if track_changes not in {"accept", "reject", "all"}:
        raise ValueError("track_changes must be accept, reject, or all")
    work.mkdir(parents=True, exist_ok=True)
    manifest_path = work / "manifest.json"
    if not manifest_path.is_file():
        docx_to_markdown(
            source,
            split_dir=work,
            section_map_path=section_map,
            track_changes=track_changes,
            force=force,
        )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("schema") != "docforge.docx2md.v1":
        raise ValueError("roundtrip workdir manifest must use schema docforge.docx2md.v1")
    if manifest.get("source_sha256") != _file_sha256(source):
        raise ValueError("roundtrip workdir was created from a different input DOCX")
    snapshot = _bundle_path(work, str(manifest.get("source_copy", "")), label="source_copy")
    if not snapshot.is_file() or _file_sha256(snapshot) != manifest.get("source_copy_sha256"):
        raise ValueError("roundtrip source snapshot is missing or its SHA-256 does not match")
    output = Path(output) if output is not None else work / "roundtrip.docx"
    output.parent.mkdir(parents=True, exist_ok=True)
    audit_path = output.with_suffix(".manifest.json")
    if audit_path.exists() and not force:
        raise FileExistsError(f"Output exists; pass --force to overwrite: {audit_path}")

    # The existing helper verifies every section and media digest and performs
    # the byte-for-byte copy for a true no-op.
    unchanged = reuse_unchanged_roundtrip_source(manifest_path, output, force=force)
    section_hashes = dict(manifest.get("section_sha256", {}))
    media_hashes = dict(manifest.get("media_sha256", {}))
    current_section_hashes = {
        name: _file_sha256(_bundle_path(work, name, label="section"))
        for name in section_hashes
        if _bundle_path(work, name, label="section").is_file()
    }
    changed_sections = sorted(name for name, expected in section_hashes.items() if current_section_hashes.get(name) != expected)
    changed_media = sorted(
        name for name, expected in media_hashes.items()
        if not _bundle_path(work, name, label="media").is_file() or _file_sha256(_bundle_path(work, name, label="media")) != expected
    )
    if unchanged:
        mode = "exact-source-reuse"
        summary = {"changed": 0, "inserted": 0, "deleted": 0}
    else:
        if changed_media:
            raise ValueError("DOCX-origin media changed; source package patching cannot redraw images")
        # Export the original snapshot to a temporary bundle to obtain the
        # exact baseline text while leaving the user's edited bundle intact.
        with tempfile.TemporaryDirectory(prefix="docforge-roundtrip-baseline-") as temporary:
            baseline_dir = Path(temporary)
            docx_to_markdown(
                snapshot,
                split_dir=baseline_dir,
                section_map_path=section_map,
                track_changes=track_changes,
                force=True,
            )
            names = list(section_hashes)
            baseline_md = baseline if baseline is not None and baseline.suffix.lower() != ".docx" else baseline_dir / "_baseline.md"
            edited_md = baseline_dir / "_edited.md"
            if baseline is None or baseline.suffix.lower() == ".docx":
                baseline_md.write_text("\n\n".join((baseline_dir / name).read_text(encoding="utf-8") for name in names), encoding="utf-8")
            edited_md.write_text(
                "\n\n".join(_bundle_path(work, name, label="section").read_text(encoding="utf-8") for name in names),
                encoding="utf-8",
            )
            source_edit_baseline = baseline_md
            source_edit_edited = edited_md
            delta_summary = apply_markdown_delta(
                snapshot,
                source_edit_baseline,
                source_edit_edited,
                output,
                overwrite=True,
            )
            summary = {
                "changed": int(delta_summary.get("replaced", 0)),
                "inserted": int(delta_summary.get("inserted", 0)),
                "deleted": int(delta_summary.get("deleted", 0)),
                "replaced": int(delta_summary.get("replaced", 0)),
            }
        mode = "source-package-patch"

    redline_summary = None
    redline_path = None
    if redline is not None:
        from ..docxdiff.redline import create_tracked_docx

        redline_path = Path(redline)
        redline_base = baseline if baseline is not None and baseline.suffix.lower() == ".docx" else source
        redline_summary = create_tracked_docx(redline_base, output, redline_path, author="roundtrip", overwrite=force, workers=1)

    bundle_manifest_sha256 = _file_sha256(manifest_path)
    audit = {
        "schema": "docforge.roundtrip.v1",
        "mode": "exact-reuse" if unchanged else "rebuild",
        "legacy_mode": mode,
        "exact_source_reused": unchanged,
        "source_sha256": _file_sha256(source),
        "manifest_sha256": bundle_manifest_sha256,
        "source": {"path": source.name, "sha256": _file_sha256(source)},
        "source_snapshot": {"path": str(Path(manifest["source_copy"]).as_posix()), "sha256": manifest["source_copy_sha256"]},
        "output": {"path": output.name, "sha256": _file_sha256(output)},
        "sections": {"sha256": {name: current_section_hashes.get(name) for name in section_hashes}, "changed": changed_sections},
        "media": {"sha256": media_hashes, "changed": changed_media},
        "revisions": redline_summary or summary,
        "track_changes": track_changes,
    }
    if redline_path is not None:
        audit["redline"] = {"path": redline_path.name, "sha256": _file_sha256(redline_path), "summary": redline_summary}
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    output_sha256 = _file_sha256(output)
    return {
        "mode": audit["mode"],
        "legacy_mode": mode,
        "exact_source_reused": unchanged,
        "output": str(output),
        "output_sha256": output_sha256,
        "source_sha256": audit["source_sha256"],
        "manifest_sha256": bundle_manifest_sha256,
        "manifest": str(audit_path),
        "changed_sections": changed_sections,
        "redline": redline_summary,
        "revisions": summary,
        "ok": True,
    }


def _soffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("soffice.com")
    if not found:
        for candidate in (r"C:\Program Files\LibreOffice\program\soffice.com", "/usr/bin/soffice",
                          "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
            if Path(candidate).exists():
                return candidate
    return found


def _word_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def render_pdfs(docs: list[Path], out_dir: Path, engine: str = "auto") -> list[Path]:
    """Render DOCX files to PDF with Word (``engine="word"``) or LibreOffice.

    Word is the reference layout. LibreOffice shapes text per run, so
    splitting or merging runs with identical formatting can move its line
    breaks even though Word lays the text out the same way.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    if engine == "auto":
        engine = "word" if _word_available() else "libreoffice"
    outputs = [out_dir / (Path(doc).stem + ".pdf") for doc in docs]
    if engine == "word":
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        # Word works on copies (the originals may be open elsewhere) and keeps
        # them locked for a while after quitting, so cleanup is best-effort.
        tmp = Path(tempfile.mkdtemp(prefix=".render-", dir=out_dir))
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            for index, (doc, pdf) in enumerate(zip(docs, outputs)):
                copy = tmp / f"{index}" / Path(doc).name
                copy.parent.mkdir()
                shutil.copy(doc, copy)
                opened = word.Documents.Open(str(copy.resolve()), ReadOnly=True, AddToRecentFiles=False, Visible=False)
                try:
                    opened.ExportAsFixedFormat(str(pdf.resolve()), 17)
                finally:
                    opened.Close(0)
        finally:
            word.Quit()
            shutil.rmtree(tmp, ignore_errors=True)
        return outputs
    soffice = _soffice()
    if not soffice:
        raise FileNotFoundError("LibreOffice (soffice) is required to render DOCX")
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), *map(str, docs)],
        check=True, capture_output=True, timeout=1200,
    )
    return outputs


def render_pdf(docx: Path, out_dir: Path, engine: str = "auto") -> Path:
    return render_pdfs([docx], out_dir, engine)[0]


def _glyphs(page) -> list[tuple]:
    glyphs = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for char in span["chars"]:
                    if char["c"].strip():
                        glyphs.append((char["c"], span["font"], round(span["size"], 1), span["color"], char["origin"]))
    return glyphs


def compare_pdfs(left: Path, right: Path, dpi: int = 60, tolerance: float = 0.5) -> list[dict]:
    """Per page: glyph identity, the largest glyph shift (pt), image count and pixel difference.

    ``same`` requires every glyph to keep its character, font, size and
    colour, to move by at most ``tolerance`` points, and the page to keep its
    image count.
    """
    import fitz

    a, b = fitz.open(left), fitz.open(right)
    pages = []
    for index in range(max(len(a), len(b))):
        if index >= len(a) or index >= len(b):
            pages.append({"page": index + 1, "same": False, "missing": True})
            continue
        ga, gb = _glyphs(a[index]), _glyphs(b[index])
        glyphs_equal = [item[:4] for item in ga] == [item[:4] for item in gb]
        shift = max(
            (max(abs(x[4][0] - y[4][0]), abs(x[4][1] - y[4][1])) for x, y in zip(ga, gb)), default=0.0,
        )
        images = (len(a[index].get_images()), len(b[index].get_images()))
        pa = a[index].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
        pb = b[index].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
        if (pa.width, pa.height) == (pb.width, pb.height):
            different = sum(1 for x, y in zip(pa.samples, pb.samples) if abs(x - y) > 32)
            pixels = different / max(len(pa.samples), 1)
        else:
            pixels = 1.0
        pages.append({
            "page": index + 1,
            "same": glyphs_equal and shift <= tolerance and images[0] == images[1],
            "glyphs_equal": glyphs_equal,
            "max_shift_pt": round(shift, 3),
            "images": list(images),
            "pixel_diff": round(pixels, 5),
        })
    return pages


__all__ = ["compare_docx", "compare_pdfs", "redline_changes", "render_pdf", "render_pdfs", "roundtrip"]
