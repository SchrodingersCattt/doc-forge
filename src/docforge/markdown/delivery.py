"""Profile-driven delivery of article and supplement DOCX files.

The delivery layer deliberately stays project-neutral.  A JSON profile names
Markdown inputs, a template, and delivery names; the same citation manifest is
then passed from the article to the supplement.  When ``accept_revisions`` is
requested, reviewed DOCX files are converted to mapped Markdown sections
before any document is built.  Missing mapping information fails closed so a
review cannot be silently omitted.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from ..output import validate_output_path
from .docx_export import docx_to_markdown, load_section_map
from .launcher import parse_markdown
from .template import AssemblyResult, assemble_markdown_template, sha256_file, write_assembly_sidecars


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


def _load_profile(profile: Mapping[str, Any] | Path | str) -> tuple[dict[str, Any], Path, str]:
    if isinstance(profile, (str, Path)):
        profile_path = Path(profile)
        payload = json.loads(profile_path.read_text(encoding="utf-8"))
        source = str(profile_path)
        base = profile_path.parent.resolve()
    else:
        payload = dict(profile)
        source = "<mapping>"
        base = Path.cwd().resolve()
    if not isinstance(payload, dict):
        raise ValueError("delivery profile must be a JSON object")
    return payload, base, source


def _path(value: Any, base: Path, *, label: str) -> Path:
    if not isinstance(value, (str, Path)) or not str(value).strip():
        raise ValueError(f"delivery profile requires {label}")
    candidate = Path(value)
    return candidate if candidate.is_absolute() else base / candidate


def _entries(profile: Mapping[str, Any]) -> list[tuple[str, dict[str, Any]]]:
    artifacts = profile.get("artifacts")
    if artifacts is not None:
        if not isinstance(artifacts, list) or not artifacts:
            raise ValueError("delivery profile 'artifacts' must be a non-empty list")
        result: list[tuple[str, dict[str, Any]]] = []
        for index, item in enumerate(artifacts):
            if not isinstance(item, Mapping):
                raise ValueError(f"delivery artifact {index} must be an object")
            value = dict(item)
            key = str(value.pop("id", value.pop("name", f"artifact_{index + 1}")))
            result.append((key, value))
        return result
    result = []
    for key in ("article", "supplement", "si", "supporting_information"):
        value = profile.get(key)
        if value is None:
            continue
        if not isinstance(value, Mapping):
            raise ValueError(f"delivery profile '{key}' must be an object")
        canonical = "supplement" if key in {"si", "supporting_information"} else key
        result.append((canonical, dict(value)))
    if not result:
        raise ValueError("delivery profile requires article/supplement or artifacts")
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


def _output_name(entry: Mapping[str, Any], key: str, profile: Mapping[str, Any]) -> str:
    names = profile.get("names", profile.get("delivery_names", {}))
    named = names.get(key) if isinstance(names, Mapping) else None
    value = entry.get("delivery_name") or entry.get("output_name") or entry.get("output") or named
    if value is None:
        pattern = entry.get("output_pattern", profile.get("output_pattern"))
        if pattern:
            value = str(pattern).format(name=key, key=key, kind=key)
        else:
            value = f"{key}.docx"
    name = Path(str(value)).name
    if not name.lower().endswith(".docx"):
        name += ".docx"
    validate_output_path(Path(name), label=f"delivery {key}")
    return name


def _apply_review(entry: dict[str, Any], base: Path, *, force: bool) -> list[Path] | None:
    reviewed = entry.get("reviewed_docx") or entry.get("reviewed") or entry.get("review")
    if reviewed is None:
        return None
    reviewed_path = _path(reviewed, base, label="reviewed_docx")
    if not reviewed_path.is_file():
        raise FileNotFoundError(f"Reviewed DOCX does not exist: {reviewed_path}")
    section_map_value = entry.get("section_map") or entry.get("section_map_path")
    source_dir_value = entry.get("source_dir") or entry.get("sources_dir")
    source_value = entry.get("source_markdown") or entry.get("source")
    if section_map_value is not None or source_dir_value is not None:
        if section_map_value is None or source_dir_value is None:
            raise ValueError("accepting reviewed revisions requires both section_map and source_dir")
        section_map = _path(section_map_value, base, label="section_map")
        source_dir = _path(source_dir_value, base, label="source_dir")
        load_section_map(section_map)  # validate before changing any files
        docx_to_markdown(
            reviewed_path,
            split_dir=source_dir,
            section_map_path=section_map,
            track_changes="accept",
            force=force,
        )
        mapped = load_section_map(section_map)
        inputs = [source_dir / str(item["file"]) for item in mapped]
        missing = [str(path) for path in inputs if not path.is_file()]
        if missing:
            raise RuntimeError(f"review conversion did not produce mapped source: {missing[0]}")
        entry["inputs"] = [str(path) for path in inputs]
        return inputs
    if source_value is not None:
        source = _path(source_value, base, label="source_markdown")
        docx_to_markdown(reviewed_path, output=source, track_changes="accept", force=force)
        entry["inputs"] = [str(source)]
        return [source]
    raise ValueError(
        "--accept-revisions requires reviewed_docx plus section_map/source_dir "
        "(or source_markdown) for each reviewed artifact"
    )


def _source_map(inputs: Sequence[Path]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in inputs:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        try:
            blocks = parse_markdown(path)
            paragraphs = [block.text for block in blocks]
        except (ValueError, RuntimeError):
            paragraphs = []
        for index, text in enumerate(paragraphs):
            result.append({"source": str(path), "paragraph": index, "text": text})
        # Keep an entry even for an empty or parser-unsupported source so the
        # manifest remains a complete map of the delivery inputs.
        if not paragraphs:
            result.append({"source": str(path), "paragraph": None, "line_count": len(lines)})
    return result


def deliver(
    profile: Mapping[str, Any] | Path | str,
    *,
    accept_revisions: bool = False,
    force: bool = False,
) -> DeliveryResult:
    """Build all documents declared by a delivery profile.

    ``article`` is built first.  Unless explicitly supplied, a supplement's
    ``citation_base`` points at the article's manifest, ensuring shared citation
    numbering.  Relative paths resolve next to the profile file, and every
    delivery name is validated before writing so accidental ``final`` names
    cannot be published.
    """
    payload, base, source = _load_profile(profile)
    accept_revisions = accept_revisions or bool(payload.get("accept_revisions", False))
    delivery_dir = _path(payload.get("delivery_dir", payload.get("output_dir", "delivery")), base, label="delivery_dir")
    delivery_dir.mkdir(parents=True, exist_ok=True)
    shared = payload.get("shared", {})
    if not isinstance(shared, Mapping):
        raise ValueError("delivery profile 'shared' must be an object")
    entries = _entries(payload)
    # Citation inheritance is deterministic even when a profile uses the
    # generic ``artifacts`` list and happens to list the supplement first.
    entries.sort(key=lambda item: (0 if item[0] == "article" else 1))
    if accept_revisions:
        has_review = any(
            isinstance(entry, Mapping)
            and (entry.get("reviewed_docx") or entry.get("reviewed") or entry.get("review"))
            for _key, entry in entries
        ) or bool(payload.get("reviewed_docx") or payload.get("reviewed"))
        if not has_review:
            raise ValueError("--accept-revisions requires reviewed_docx/reviewed input in the delivery profile")
    built: list[DeliveryArtifact] = []
    aggregate_map: list[dict[str, Any]] = []
    article_manifest: Path | None = None
    for key, raw_entry in entries:
        entry = dict(raw_entry)
        if accept_revisions:
            # Allow a top-level reviewed mapping for concise article +
            # supplement profiles while retaining per-artifact overrides.
            reviewed_defaults = payload.get("reviewed_docx") or payload.get("reviewed")
            if "reviewed_docx" not in entry and "reviewed" not in entry and reviewed_defaults is not None:
                if isinstance(reviewed_defaults, Mapping):
                    if key in reviewed_defaults:
                        entry["reviewed_docx"] = reviewed_defaults[key]
                else:
                    entry["reviewed_docx"] = reviewed_defaults
            for field in ("section_map", "section_map_path", "source_dir", "sources_dir", "source_markdown", "source"):
                if field not in entry and isinstance(payload.get(field), Mapping) and key in payload[field]:
                    entry[field] = payload[field][key]
                elif field not in entry and field in payload and not isinstance(payload[field], Mapping):
                    entry[field] = payload[field]
            reviewed_inputs = _apply_review(entry, base, force=force)
        else:
            reviewed_inputs = None
        input_value = entry.get("inputs", entry.get("sources", entry.get("markdown")))
        if input_value is None:
            raise ValueError(f"delivery artifact '{key}' requires inputs")
        inputs = reviewed_inputs or _resolve_inputs(input_value, base, label=f"{key}.inputs")
        template_value = entry.get("template", shared.get("template"))
        template = _path(template_value, base, label=f"{key}.template")
        metadata_value = entry.get("metadata", entry.get("metadata_path", shared.get("metadata")))
        metadata = _path(metadata_value, base, label=f"{key}.metadata") if metadata_value is not None else None
        bibliography_value = entry.get("bibliography", shared.get("bibliography"))
        bibliography = _path(bibliography_value, base, label=f"{key}.bibliography") if bibliography_value is not None else None
        citation_base_value = entry.get("citation_base", shared.get("citation_base", payload.get("citation_base")))
        if key == "supplement" and citation_base_value is None and article_manifest is not None:
            citation_base_value = article_manifest
        citation_base = _path(citation_base_value, base, label=f"{key}.citation_base") if citation_base_value is not None else None
        options: dict[str, Any] = dict(shared.get("options", {})) if isinstance(shared.get("options", {}), Mapping) else {}
        if isinstance(payload.get("options"), Mapping):
            options.update(payload["options"])
        if isinstance(entry.get("options"), Mapping):
            options.update(entry["options"])
        # Profile keys are also accepted as direct assemble options.
        for option in (
            "title", "keep_comments", "skip_images", "columns", "figure_span", "font_family",
            "east_asia_font", "style_profile", "line_numbers", "include_title",
            "strip_level_one_headings", "heading_before", "numbering_prefix", "bibliography_scope",
            "citation_numbering", "bibliography_profile", "include_metadata_back_matter", "native_toc",
            "restart_heading_numbering", "body_first_line_chars", "page_break_before_h1",
            "body_font_size", "abstract_font_size", "caption_font_size", "reference_font_size",
        ):
            if option in entry:
                options[option] = entry[option]
        output = delivery_dir / _output_name(entry, key, payload)
        validate_output_path(output, label=f"delivery {key}")
        result: AssemblyResult = assemble_markdown_template(
            inputs,
            template_path=template,
            output=output,
            metadata_path=metadata,
            bibliography_path=bibliography,
            citation_base_path=citation_base,
            force=force,
            **options,
        )
        manifest, checksum = write_assembly_sidecars(
            result,
            inputs=inputs,
            template_path=template,
            metadata_path=metadata,
            bibliography_path=bibliography,
            citation_base_path=citation_base,
            command=[
                "docforge",
                "deliver",
                "--profile",
                source,
                *( ["--accept-revisions"] if accept_revisions else [] ),
            ],
        )
        built.append(DeliveryArtifact(key, output, manifest, checksum, tuple(inputs)))
        aggregate_map.extend(_source_map(inputs))
        if key == "article":
            article_manifest = manifest
    aggregate = delivery_dir / str(payload.get("manifest", "delivery.manifest.json"))
    validate_output_path(aggregate, label="delivery manifest")
    aggregate.write_text(
        json.dumps(
            {
                "schema": "docforge.delivery.v1",
                "generated_at_utc": datetime.now(timezone.utc).isoformat(),
                "profile": source,
                "accept_revisions": accept_revisions,
                "delivery_dir": str(delivery_dir),
                "artifacts": [
                    {
                        "name": item.name,
                        "output": str(item.output),
                        "manifest": str(item.manifest),
                        "checksum": str(item.checksum),
                        "inputs": [str(path) for path in item.inputs],
                    }
                    for item in built
                ],
                "source_paragraph_map": aggregate_map,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    return DeliveryResult(delivery_dir, tuple(built), aggregate)


__all__ = ["DeliveryArtifact", "DeliveryResult", "deliver"]
