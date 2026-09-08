"""Front-matter-driven image generation with auditable sidecars."""

from __future__ import annotations

import base64
import hashlib
import json
import mimetypes
import os
import re
import time
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from PIL import Image

ALLOWED_FRONT_MATTER = {
    "output_name",
    "aspect_ratio",
    "image_size",
    "status",
    "model",
    "structure_slots",
}
PROMPT_FRONT_MATTER_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n(.*)$", re.DOTALL)


@dataclass(frozen=True)
class PromptSpec:
    source: str
    source_sha256: str
    output_name: str
    aspect_ratio: str
    image_size: str
    status: str
    model: str | None
    prompt: str


@dataclass(frozen=True)
class GenerateResult:
    output: Path
    metadata: dict[str, Any]


def load_prompt(path: Path, *, root: Path | None = None) -> PromptSpec:
    raw = path.read_text(encoding="utf-8")
    metadata: dict[str, Any] = {}
    match = PROMPT_FRONT_MATTER_RE.match(raw)
    if match:
        header = match.group(1)
        body = match.group(2)
        end = len(raw)
    else:
        header = ""
        body = raw
        end = len(raw)
    if header:
        for line in header.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if ":" not in stripped:
                raise ValueError(f"Invalid front-matter line in {path}: {line}")
            key, value = stripped.split(":", 1)
            key = key.strip()
            if key not in ALLOWED_FRONT_MATTER:
                raise ValueError(f"Unsupported front-matter key {key!r} in {path}")
            metadata[key] = parse_scalar(value)
    prompt = body.strip()
    if not prompt:
        raise ValueError(f"Prompt body is empty: {path}")
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    source = str(path.relative_to(root)) if root else str(path)
    return PromptSpec(
        source=source,
        source_sha256=digest,
        output_name=metadata.get("output_name", path.stem),
        aspect_ratio=metadata.get("aspect_ratio", "16:9"),
        image_size=metadata.get("image_size", "2K"),
        status=metadata.get("status", "unspecified"),
        model=metadata.get("model"),
        prompt=prompt,
    )


def parse_scalar(value: str) -> Any:
    value = value.strip()
    if value in ("true", "True"):
        return True
    if value in ("false", "False"):
        return False
    if value.isdigit():
        return int(value)
    try:
        return float(value)
    except ValueError:
        return value


def discover_prompts(prompt_dir: Path) -> list[Path]:
    return sorted(p for p in prompt_dir.glob("*.md") if p.is_file())


def load_configuration(env_file: Path | None = None) -> dict[str, str]:
    try:
        from dotenv import load_dotenv
    except ImportError:  # pragma: no cover
        load_dotenv = None  # type: ignore[assignment]
    if load_dotenv is not None:
        load_dotenv(env_file, override=False) if env_file else load_dotenv(override=False)
    config = {
        "base_url": os.getenv("BASE_URL", "").strip(),
        "litellm_key": os.getenv("LITELLM_KEY", "").strip(),
        "gemini_key": os.getenv("GEMINI_API_KEY", "").strip()
        or os.getenv("GOOGLE_API_KEY", "").strip(),
        "image_model": os.getenv("IMAGE_MODEL", "").strip(),
    }
    return config


def safe_endpoint_host(base_url: str) -> str:
    parsed = urllib.parse.urlparse(base_url)
    if parsed.username or parsed.password or parsed.query:
        raise ValueError("BASE_URL must not contain credentials or query parameters")
    return parsed.netloc


def normalize_openai_base_url(base_url: str) -> str:
    """Return the API root accepted by an OpenAI-compatible client.

    Some LiteLLM deployments expose the model catalogue at ``/models`` and use
    the same URL in local configuration. The OpenAI client expects the API root,
    otherwise it appends ``/images/generations`` below ``/models`` and receives
    a 404 response.
    """
    parsed = urllib.parse.urlparse(base_url.rstrip("/"))
    path = parsed.path.rstrip("/")
    if path.endswith("/models"):
        path = path[: -len("/models")]
    normalized = parsed._replace(path=path, params="", query="", fragment="")
    return urllib.parse.urlunparse(normalized).rstrip("/")


def decode_data_url(value: str) -> tuple[bytes, str]:
    match = re.fullmatch(r"data:([^;,]+);base64,(.+)", value, flags=re.DOTALL)
    if not match:
        raise ValueError("Unsupported image data URL")
    return base64.b64decode(match.group(2), validate=True), match.group(1)


def download_image(url: str, max_bytes: int = 30 * 1024 * 1024) -> tuple[bytes, str]:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != "https":
        raise ValueError("Only HTTPS image URLs are accepted")
    request = urllib.request.Request(url, headers={"User-Agent": "docforge-figure-generator/1.0"})
    with urllib.request.urlopen(request, timeout=120) as response:
        content_type = response.headers.get_content_type()
        if not content_type.startswith("image/"):
            raise ValueError(f"Remote response is not an image: {content_type}")
        content_length = response.headers.get("Content-Length")
        if content_length and int(content_length) > max_bytes:
            raise ValueError("Remote image exceeds size limit")
        data = response.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValueError("Remote image exceeds size limit")
    return data, content_type


def generate_litellm(spec: PromptSpec, config: dict[str, str]) -> tuple[bytes, str, dict[str, Any]]:
    from openai import OpenAI

    if not config["base_url"]:
        raise RuntimeError("BASE_URL is missing from .env")
    if not config["litellm_key"]:
        raise RuntimeError("LITELLM_KEY is missing from .env")
    model = spec.model or config["image_model"]
    if not model:
        raise RuntimeError("IMAGE_MODEL is missing from prompt front matter and .env")
    host = safe_endpoint_host(config["base_url"])
    api_root = normalize_openai_base_url(config["base_url"])
    client = OpenAI(
        api_key=config["litellm_key"],
        base_url=api_root,
        timeout=300,
        max_retries=2,
    )
    kwargs: dict[str, Any] = {
        "model": model,
        "prompt": spec.prompt,
        "n": 1,
        "response_format": "b64_json",
    }
    response = client.images.generate(**kwargs)
    if not response.data:
        raise RuntimeError("Image endpoint returned no images")
    item = response.data[0]
    if getattr(item, "b64_json", None):
        data = base64.b64decode(item.b64_json, validate=True)
        mime = "image/png"
    elif getattr(item, "url", None):
        value = item.url
        if value.startswith("data:"):
            data, mime = decode_data_url(value)
        else:
            data, mime = download_image(value)
    else:
        raise RuntimeError("Image response contained neither b64_json nor URL")
    return data, mime, {
        "backend": "litellm",
        "endpoint_host": host,
        "requested_model": model,
        "response_created": getattr(response, "created", None),
    }


def generate_google(spec: PromptSpec, config: dict[str, str]) -> tuple[bytes, str, dict[str, Any]]:
    from google import genai
    from google.genai import types

    if not config["gemini_key"]:
        raise RuntimeError("GEMINI_API_KEY or GOOGLE_API_KEY is missing from .env")
    model = spec.model or config["image_model"] or "gemini-3-pro-image"
    client = genai.Client(api_key=config["gemini_key"])
    response = client.models.generate_content(
        model=model,
        contents=spec.prompt,
        config=types.GenerateContentConfig(
            response_modalities=["IMAGE"],
            image_config=types.ImageConfig(
                aspect_ratio=spec.aspect_ratio,
                image_size=spec.image_size,
            ),
        ),
    )
    candidates = getattr(response, "candidates", None) or []
    for candidate in candidates:
        content = getattr(candidate, "content", None)
        for part in getattr(content, "parts", None) or []:
            inline = getattr(part, "inline_data", None)
            if inline and getattr(inline, "data", None):
                raw = inline.data
                data = raw if isinstance(raw, bytes) else base64.b64decode(raw)
                mime = getattr(inline, "mime_type", None) or "image/png"
                return data, mime, {
                    "backend": "google",
                    "endpoint_host": "generativelanguage.googleapis.com",
                    "requested_model": model,
                }
    raise RuntimeError("Google response contained no image data")


def infer_extension(mime_type: str, data: bytes) -> str:
    guessed = mimetypes.guess_extension(mime_type) or ".png"
    if guessed == ".jpe":
        guessed = ".jpg"
    try:
        with Image.open(BytesIO(data)) as image:
            fmt = (image.format or "").lower()
        if fmt == "jpeg":
            return ".jpg"
        if fmt:
            return f".{fmt}"
    except Exception:
        pass
    return guessed


def verify_image(data: bytes) -> tuple[int | None, int | None, str | None]:
    with Image.open(BytesIO(data)) as image:
        image.verify()
    with Image.open(BytesIO(data)) as image:
        width, height = image.size
        fmt = image.format
    return width, height, fmt


def unique_output_path(output_dir: Path, base_name: str, extension: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    candidate = output_dir / f"{base_name}{extension}"
    if not candidate.exists():
        return candidate
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return output_dir / f"{base_name}_{timestamp}{extension}"


def write_atomic(path: Path, data: bytes) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_bytes(data)
    os.replace(temp, path)


def generate_one(
    spec: PromptSpec,
    backend: str,
    config: dict[str, str],
    *,
    output_dir: Path,
    overwrite: bool = False,
    root: Path | None = None,
) -> GenerateResult:
    started = datetime.now(timezone.utc)
    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            if backend == "litellm":
                data, mime, provider_meta = generate_litellm(spec, config)
            elif backend == "google":
                data, mime, provider_meta = generate_google(spec, config)
            else:
                raise ValueError(f"Unknown backend: {backend}")
            width, height, image_format = verify_image(data)
            extension = infer_extension(mime, data)
            output_path = (
                output_dir / f"{spec.output_name}{extension}"
                if overwrite
                else unique_output_path(output_dir, spec.output_name, extension)
            )
            write_atomic(output_path, data)
            metadata = {
                "schema_version": 1,
                "created_utc": datetime.now(timezone.utc).isoformat(),
                "prompt": {k: v for k, v in asdict(spec).items() if k != "prompt"},
                "requested": {
                    "backend": backend,
                    "aspect_ratio": spec.aspect_ratio,
                    "image_size": spec.image_size,
                },
                "provider": provider_meta,
                "artifact": {
                    "path": str(output_path.relative_to(root)) if root else str(output_path),
                    "mime_type": mime,
                    "format": image_format,
                    "width": width,
                    "height": height,
                    "byte_length": len(data),
                    "sha256": hashlib.sha256(data).hexdigest(),
                },
                "attempts": attempt,
                "elapsed_seconds": (datetime.now(timezone.utc) - started).total_seconds(),
            }
            meta_path = output_path.with_suffix(output_path.suffix + ".json")
            write_atomic(meta_path, json.dumps(metadata, ensure_ascii=False, indent=2).encode("utf-8"))
            return GenerateResult(output=output_path, metadata=metadata)
        except Exception as exc:  # noqa: BLE001 - sanitize before display
            last_error = exc
            if attempt == 3:
                break
            time.sleep(min(2**attempt, 8))
    assert last_error is not None
    raise RuntimeError(f"Generation failed after 3 attempts: {type(last_error).__name__}: {last_error}")


def select_prompts(paths: list[Path], selectors: list[str], all_prompts: bool) -> list[Path]:
    if all_prompts:
        return paths
    if not selectors:
        return []
    selected: list[Path] = []
    for selector in selectors:
        matches = [p for p in paths if p.stem == selector or p.name == selector]
        if not matches:
            matches = [p for p in paths if selector in p.stem]
        if len(matches) != 1:
            raise ValueError(f"Selector {selector!r} matched {len(matches)} prompt files")
        if matches[0] not in selected:
            selected.append(matches[0])
    return selected


__all__ = [
    "PromptSpec",
    "GenerateResult",
    "load_prompt",
    "discover_prompts",
    "load_configuration",
    "generate_one",
    "select_prompts",
    "write_atomic",
    "verify_image",
]