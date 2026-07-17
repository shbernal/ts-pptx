#!/usr/bin/env python3
"""Generate `docs/doc-index.md`: a read_when discovery index.

Every narrative doc declares a `read_when:` list in its frontmatter, but nothing
aggregates them, so there is no single place to answer "which doc covers this
task?". This script walks `docs/`, collects each page that carries a non-empty
`read_when`, and emits a generated `docs/doc-index.md` grouping every page under its
scenario hints.

Generated pages without `read_when` (e.g. the typedoc `reference/api/` tree) are
skipped by construction. The output is a generated artifact (gitignored, like
`reference/api/` and `public/llms*.txt`); regenerate it with `pnpm run
docs:index`. It is validated by `scripts/docs-check.py` (frontmatter + links),
which runs after generation in `docs:build`.

The frontmatter parser mirrors `docs-list.py` / `docs-check.py` (the repo keeps
these small scripts self-contained rather than sharing a module).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


EXCLUDED_DIRS = {"archive", "research"}
OUTPUT_NAME = "doc-index.md"


def strip_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def compact_strings(values: list[object]) -> list[str]:
    out: list[str] = []
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            out.append(text)
    return out


def parse_inline_list(value: str) -> list[str]:
    try:
        parsed = json.loads(value.replace("'", '"'))
    except json.JSONDecodeError:
        return []
    return compact_strings(parsed) if isinstance(parsed, list) else []


def parse_frontmatter(path: Path) -> dict[str, object]:
    raw = path.read_text(encoding="utf-8")
    if not raw.startswith("---\n") and not raw.startswith("---\r\n"):
        return {}

    lines = raw.splitlines()
    end_index = None
    for index, line in enumerate(lines[1:], start=1):
        if line.strip() in {"---", "..."}:
            end_index = index
            break
    if end_index is None:
        return {}

    data: dict[str, object] = {}
    collecting: str | None = None
    for raw_line in lines[1:end_index]:
        line = raw_line.strip()
        if not line:
            continue

        if collecting and line.startswith("- "):
            current = data.setdefault(collecting, [])
            if isinstance(current, list):
                current.append(strip_quotes(line[2:].strip()))
            continue

        collecting = None
        if ":" not in line:
            continue

        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if not value:
            data[key] = []
            collecting = key
        elif value.startswith("[") and value.endswith("]"):
            data[key] = parse_inline_list(value)
        else:
            data[key] = strip_quotes(value)

    return data


def walk_docs(docs_dir: Path) -> list[Path]:
    files: list[Path] = []
    for path in docs_dir.rglob("*.md"):
        rel_parts = path.relative_to(docs_dir).parts
        if any(part.startswith(".") or part in EXCLUDED_DIRS for part in rel_parts):
            continue
        if path.name == OUTPUT_NAME:  # never index the generated index itself
            continue
        files.append(path)
    return sorted(files, key=lambda item: item.relative_to(docs_dir).as_posix())


def render(entries: list[dict[str, object]]) -> str:
    lines = [
        "---",
        "doc-schema-version: 1",
        'title: "Documentation Index"',
        'summary: "Generated read_when discovery index: every guide and the scenarios that should send you to it."',
        "read_when:",
        "  - Looking for which doc covers a task or scenario",
        "  - Discovering documentation by when to read it",
        'doc_type: "reference"',
        "---",
        "",
        "<!-- GENERATED FILE. Do not edit by hand.",
        "     Regenerate with `pnpm run docs:index` (runs in `docs:prepare`).",
        "     Source: the `read_when:` frontmatter across docs/. -->",
        "",
        "# Documentation Index",
        "",
        "Each doc below declares, in its frontmatter, the situations in which you",
        "should read it. This page aggregates those `read_when` hints so you can find",
        "the right doc by task. When a task matches a hint, read that doc before",
        "coding and update it when behavior changes.",
        "",
    ]

    for entry in entries:
        rel = str(entry["rel"])
        title = str(entry["title"])
        summary = str(entry["summary"]).strip()
        read_when = entry["read_when"]
        lines.append(f"## [{title}]({rel})")
        lines.append("")
        if summary:
            lines.append(summary)
            lines.append("")
        lines.append("Read when:")
        lines.append("")
        if isinstance(read_when, list):
            for hint in read_when:
                lines.append(f"- {hint}")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    docs_dir = Path("docs")
    if not docs_dir.exists() or not docs_dir.is_dir():
        print("docs:index: missing docs directory. Run from repo root.", file=sys.stderr)
        return 1

    entries: list[dict[str, object]] = []
    for path in walk_docs(docs_dir):
        frontmatter = parse_frontmatter(path)
        read_when = frontmatter.get("read_when", [])
        hints = compact_strings(read_when) if isinstance(read_when, list) else []
        if not hints:  # generated/api pages and any page without hints
            continue
        rel = path.relative_to(docs_dir).as_posix()
        title = str(frontmatter.get("title", "")).strip() or rel
        entries.append(
            {
                "rel": rel,
                "title": title,
                "summary": str(frontmatter.get("summary", "")).strip(),
                "read_when": hints,
            }
        )

    output_path = docs_dir / OUTPUT_NAME
    output_path.write_text(render(entries), encoding="utf-8")
    print(f"docs:index: wrote {output_path.as_posix()} ({len(entries)} doc(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
