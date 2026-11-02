from __future__ import annotations

import os
from pathlib import Path

import pytest


@pytest.fixture
def directory_link(tmp_path: Path):
    """Create real directory links, using unprivileged junctions on Windows."""
    links: list[Path] = []

    def create(link: Path, target: Path) -> Path:
        link = link.absolute()
        target = target.resolve()
        assert link.is_relative_to(tmp_path.resolve())
        assert target.is_relative_to(tmp_path.resolve())
        if os.name == "nt":
            import _winapi

            creator = getattr(_winapi, "CreateJunction", None)
            if creator is None:
                pytest.skip("Windows directory junction creation is unavailable")
            creator(str(target), str(link))
        else:
            link.symlink_to(target, target_is_directory=True)
        links.append(link)
        return link

    yield create
    for link in reversed(links):
        if link.is_symlink():
            link.unlink()
        elif link.exists():
            os.rmdir(link)
