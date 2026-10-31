"""Advisory reminder when src/ changes without a docs note.

The process exits 0 when it warns. A missing note is a warning annotation and
a pull request comment. It does not fail the build.

The checklist prompt in ``.github/pull_request_template.md`` names the token
``Docs: N/A``. That prompt is not a declaration. A declaration is any other
occurrence of that exact token, usually with a one-line reason.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

MARKER = "<!-- docforge-docs-reminder -->"
TEMPLATE_PROMPT = "Docs paths touched, or `Docs: N/A` and a one-line reason:"
_COMMENT = re.compile(r"<!--.*?-->", re.DOTALL)

COMMENT_BODY = f"""{MARKER}

**Advisory.** This note does not fail the build.

`src/` changed, this pull request does not change `docs/**` or `AGENTS.md`, and the description has no `Docs: N/A` note.

- Out of scope (pure refactor, test-only or CI-only change, private helper, or docs-only pull request): add `Docs: N/A` and a one-line reason.
- In scope (public API, CLI, Markdown/TeX/DOCX contracts, migration, or ADR-level design): update the stable docs in this pull request.

See `AGENTS.md`.
"""


def declaration_present(body: str | None) -> bool:
    """Return true when the author wrote ``Docs: N/A`` outside the template prompt."""
    text = _COMMENT.sub("", body or "")
    text = text.replace(TEMPLATE_PROMPT, "", 1)
    return "Docs: N/A" in text


def _is_src(path: str) -> bool:
    return path == "src" or path.startswith("src/")


def _is_docs_or_agents(path: str) -> bool:
    return path == "AGENTS.md" or path == "docs" or path.startswith("docs/")


def should_warn(paths: list[str], body: str | None) -> bool:
    """Warn when src/ changed, docs were not touched, and the body has no note."""
    if not any(_is_src(path) for path in paths):
        return False
    if any(_is_docs_or_agents(path) for path in paths):
        return False
    return not declaration_present(body)


def paths_from_entries(entries: list[dict]) -> list[str]:
    paths: list[str] = []
    for entry in entries:
        filename = entry.get("filename") or ""
        if filename:
            paths.append(filename)
        previous = entry.get("previous_filename") or ""
        if previous:
            paths.append(previous)
    return paths


def paths_from_pages(pages: list) -> list[str]:
    entries: list[dict] = []
    for page in pages:
        if isinstance(page, list):
            entries.extend(item for item in page if isinstance(item, dict))
        elif isinstance(page, dict):
            entries.append(page)
    return paths_from_entries(entries)


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _self_check() -> None:
    root = _repo_root()
    template = (root / ".github" / "pull_request_template.md").read_text(encoding="utf-8")
    if template.count(TEMPLATE_PROMPT) != 1:
        raise SystemExit("pull request template prompt drifted from the docs reminder")
    if template.count("Docs: N/A") != 1:
        raise SystemExit("template must mention Docs: N/A only in the checklist prompt")
    if declaration_present(template):
        raise SystemExit("unfilled template must not count as Docs: N/A")

    src = "src/docforge/api.py"
    cases = [
        ([src], "", True),
        ([src], None, True),
        ([src, "docs/pipeline.md"], "", False),
        ([src, "docs/adr/0002-example.md"], "", False),
        ([src, "AGENTS.md"], "", False),
        ([src], "Docs: N/A — private helper", False),
        ([src], "See Docs: N/A for the internal refactor.", False),
        ([src], "<!-- Docs: N/A — hidden -->", True),
        ([src], template, True),
        ([src], template + "\nDocs: N/A — test-only change\n", False),
        ([src], template.replace(TEMPLATE_PROMPT, TEMPLATE_PROMPT + " Docs: N/A — refactor"), False),
        (["tests/test_api.py"], "", False),
        (["docs/pipeline.md"], "", False),
        (["README.md", "src-extra/notes.md"], "", False),
        ([], "Docs: N/A — nothing", False),
    ]
    for paths, body, expected in cases:
        actual = should_warn(paths, body)
        if actual is not expected:
            raise SystemExit(f"should_warn({paths!r}, {body!r}) -> {actual}, expected {expected}")

    renamed_out_of_docs = paths_from_entries(
        [{"filename": src, "previous_filename": "docs/old.md"}]
    )
    if should_warn(renamed_out_of_docs, ""):
        raise SystemExit("a rename that touches docs/ must not warn")
    renamed_inside_src = paths_from_entries(
        [{"filename": "src/docforge/b.py", "previous_filename": "src/docforge/a.py"}]
    )
    if not should_warn(renamed_inside_src, ""):
        raise SystemExit("a rename inside src/ must warn when docs are untouched")
    pages = [[{"filename": src}], [{"filename": "tests/test_api.py"}]]
    if paths_from_pages(pages) != [src, "tests/test_api.py"]:
        raise SystemExit("paginated file lists must flatten in order")

    marker_body = f"quoted marker\n{MARKER}\n"
    bot = {
        "id": 10,
        "body": marker_body,
        "user": {"login": "github-actions[bot]", "type": "Bot"},
    }
    human = {
        "id": 11,
        "body": marker_body,
        "user": {"login": "reviewer", "type": "User"},
    }
    other_bot = {
        "id": 12,
        "body": marker_body,
        "user": {"login": "example[bot]", "type": "Bot"},
    }
    missing_type = {
        "id": 13,
        "body": marker_body,
        "user": {"login": "github-actions[bot]"},
    }
    no_marker = {
        "id": 14,
        "body": "no marker here",
        "user": {"login": "github-actions[bot]", "type": "Bot"},
    }
    if reminder_comment_ids([[human, bot, other_bot, missing_type, no_marker]]) != ["10"]:
        raise SystemExit("only a github-actions[bot] comment with the marker may be edited")


def _gh(args: list[str], *, payload: dict | None = None) -> subprocess.CompletedProcess[str]:
    command = ["gh", "api", *args]
    if payload is not None:
        command.extend(["--input", "-"])
    return subprocess.run(
        command,
        input=json.dumps(payload) if payload is not None else None,
        text=True,
        capture_output=True,
        check=False,
        env=os.environ,
    )


def _fetch_pages(endpoint: str) -> list:
    result = _gh(["--paginate", "--slurp", endpoint])
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "request failed").strip()
        raise RuntimeError(detail.splitlines()[-1][:500])
    if not result.stdout.strip():
        return []
    data = json.loads(result.stdout)
    if not isinstance(data, list):
        raise RuntimeError("unexpected API payload")
    return data


def owned_reminder_comment(item: dict) -> bool:
    """True only for this workflow's comment, posted by github-actions[bot]."""
    if not isinstance(item, dict) or item.get("id") is None:
        return False
    if MARKER not in (item.get("body") or ""):
        return False
    user = item.get("user")
    if not isinstance(user, dict):
        return False
    return user.get("type") == "Bot" and user.get("login") == "github-actions[bot]"


def reminder_comment_ids(pages: list) -> list[str]:
    found: list[str] = []
    for page in pages:
        items = page if isinstance(page, list) else [page]
        for item in items:
            if isinstance(item, dict) and owned_reminder_comment(item):
                found.append(str(item["id"]))
    return found


def _comment_ids(repo: str, number: str) -> list[str]:
    return reminder_comment_ids(_fetch_pages(f"repos/{repo}/issues/{number}/comments"))


def _upsert_comment(repo: str, number: str) -> None:
    ids = _comment_ids(repo, number)
    if not ids:
        result = _gh(
            ["--method", "POST", f"repos/{repo}/issues/{number}/comments"],
            payload={"body": COMMENT_BODY},
        )
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "could not comment").strip()
            raise RuntimeError(detail.splitlines()[-1][:500])
        return
    first, *rest = ids
    result = _gh(
        ["--method", "PATCH", f"repos/{repo}/issues/comments/{first}"],
        payload={"body": COMMENT_BODY},
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "could not update comment").strip()
        raise RuntimeError(detail.splitlines()[-1][:500])
    for extra in rest:
        _gh(["--method", "DELETE", f"repos/{repo}/issues/comments/{extra}"])


def _clear_comments(repo: str, number: str) -> None:
    for comment_id in _comment_ids(repo, number):
        result = _gh(["--method", "DELETE", f"repos/{repo}/issues/comments/{comment_id}"])
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "could not delete comment").strip()
            raise RuntimeError(detail.splitlines()[-1][:500])


def run() -> None:
    repo = os.environ["REPO"]
    number = os.environ["PR_NUMBER"]
    body = os.environ.get("PR_BODY") or ""
    pages = _fetch_pages(f"repos/{repo}/pulls/{number}/files")
    paths = paths_from_pages(pages)
    if should_warn(paths, body):
        print(
            "::warning::src/ changed without docs/** or AGENTS.md, "
            "and the pull request body has no Docs: N/A note. See AGENTS.md."
        )
        try:
            _upsert_comment(repo, number)
        except Exception as exc:
            print(f"::warning::docs reminder could not comment ({exc})")
        return
    try:
        _clear_comments(repo, number)
    except Exception as exc:
        print(f"::warning::docs reminder could not update its comment ({exc})")


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--self-check" in args:
        _self_check()
        print("docs reminder self-check ok")
        return 0
    try:
        run()
    except Exception as exc:
        print(f"::warning::docs reminder skipped ({exc})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
