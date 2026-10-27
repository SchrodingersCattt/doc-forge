"""Controlled conversion bench for small, medium, and complex fixtures.

Each operation runs in its own process so peak RSS is that process, not a
shared suite. Publish the median of three iterations. A missing before-run
is not a regression pass; compare only measured pairs.

    python -m docforge.bench --iterations 3
"""

from __future__ import annotations

import argparse
import json
import resource
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path


KINDS = ("small", "medium", "complex")
OPS = ("md-docx", "exact-reuse", "restyle")


def markdown_for(kind: str, root: Path) -> str:
    """Return synthetic Markdown. The image stays inside ``root``."""

    if kind == "small":
        return "Alpha paragraph.\n\nBeta paragraph.\n"
    if kind == "medium":
        lines = [
            "# Section\n",
            "Body with **bold** and *italic*.\n",
            "| A | B |\n| --- | --- |\n| 1 | 2 |\n",
        ]
        lines.extend(f"Paragraph {index} keeps a short synthetic sentence.\n" for index in range(40))
        return "\n".join(lines)
    from PIL import Image

    image = root / "mark.png"
    Image.new("RGB", (8, 8), (1, 2, 3)).save(image)
    lines = [
        "# Complex\n",
        f"![Mark]({image.name})\n",
        "See *term* and **name**.\n",
        "| H | I |\n| --- | --- |\n| a | b |\n| c | d |\n",
    ]
    lines.extend(
        f"## Part {index % 7}\n\nComplex paragraph {index} with a number {index * 3}.\n"
        for index in range(120)
    )
    return "\n".join(lines)


def run_case(op: str, kind: str, root: Path) -> dict[str, object]:
    """Time one operation. Exact reuse and restyle build the source first."""

    from .api import convert, restyle

    if op not in OPS or kind not in KINDS:
        raise ValueError(f"unknown bench case: {op} {kind}")
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    source = root / f"{kind}.md"
    source.write_text(markdown_for(kind, root), encoding="utf-8")
    out = root / f"{kind}.docx"
    profile = {
        "schema": "docforge.style.v1",
        "roles": {"body": {"word_style": "Normal", "font": {"size_pt": 12}}},
        "policy": {"missing_style": "error", "direct_formatting": "clear_conflicting"},
    }
    if op == "md-docx":
        start = time.perf_counter()
        convert(source, out, force=True)
        target = out
    elif op == "exact-reuse":
        convert(source, out, force=True)
        target = root / "copy.docx"
        start = time.perf_counter()
        convert(out, target, force=True)
    else:
        convert(source, out, force=True)
        target = root / "styled.docx"
        start = time.perf_counter()
        restyle(out, target, profile=profile, force=True)
    return {
        "op": op,
        "kind": kind,
        "seconds": time.perf_counter() - start,
        "rss_kb": resource.getrusage(resource.RUSAGE_SELF).ru_maxrss,
        "bytes": target.stat().st_size,
    }


def measure(*, iterations: int = 3) -> list[dict[str, object]]:
    """Run every case in a fresh interpreter and return one row per iteration."""

    rows: list[dict[str, object]] = []
    for iteration in range(iterations):
        for kind in KINDS:
            for op in OPS:
                with tempfile.TemporaryDirectory(prefix="docforge-bench-") as directory:
                    completed = subprocess.run(
                        [sys.executable, "-m", "docforge.bench", "--worker", op, kind, directory],
                        check=False,
                        capture_output=True,
                        text=True,
                    )
                    if completed.returncode != 0:
                        raise RuntimeError(completed.stderr[-2000:] or completed.stdout[-2000:])
                    row = json.loads(completed.stdout.strip().splitlines()[-1])
                    row["iteration"] = iteration
                    rows.append(row)
    return rows


def median_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[tuple[str, str], list[dict[str, object]]] = {}
    for row in rows:
        grouped.setdefault((str(row["op"]), str(row["kind"])), []).append(row)
    summary = []
    for (op, kind), items in grouped.items():
        summary.append({
            "op": op,
            "kind": kind,
            "seconds": statistics.median(float(item["seconds"]) for item in items),
            "rss_kb": statistics.median(int(item["rss_kb"]) for item in items),
            "bytes": statistics.median(int(item["bytes"]) for item in items),
            "runs": len(items),
        })
    return summary


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Time fixed docforge conversion fixtures")
    parser.add_argument("--iterations", type=int, default=3)
    parser.add_argument("--worker", nargs=3, metavar=("OP", "KIND", "ROOT"))
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.worker:
        op, kind, root = args.worker
        print(json.dumps(run_case(op, kind, Path(root))))
        return 0
    if args.iterations < 1:
        raise SystemExit("iterations must be positive")
    rows = measure(iterations=args.iterations)
    print(json.dumps({"rows": rows, "median": median_rows(rows)}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = ["KINDS", "OPS", "markdown_for", "measure", "median_rows", "run_case"]
