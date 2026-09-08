"""Front-matter-driven image generation with auditable sidecars."""

from .generator import (
    PromptSpec,
    GenerateResult,
    load_prompt,
    discover_prompts,
    load_configuration,
    generate_one,
    select_prompts,
    write_atomic,
    verify_image,
)

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