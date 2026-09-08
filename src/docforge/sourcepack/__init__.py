"""Auditable TeX source packaging."""

from .packer import package_files, sha256_bytes, git_short_hash

__all__ = ["package_files", "sha256_bytes", "git_short_hash"]