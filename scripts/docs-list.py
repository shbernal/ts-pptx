#!/usr/bin/env python3
from __future__ import annotations

import json
import sys
from pathlib import Path


EXCLUDED_DIRS = {"archive", "research"}


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


def parse_frontmatter(path: Path) -> tuple[dict[str, object], str | None]:
    raw = path.read_text(encoding="utf-8")
    if not raw.startswith("---\n") and not raw.startswith("---\r\n"):
        return {}, "missing front matter"

    lines = raw.splitlines()
    end_index = None
    for index, line in enumerate(lines[1:], start=1):
        if line.strip() in {"---", "..."}:
            end_index = index
            break

    if end_index is None:
        return {}, "unterminated front matter"

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

    return data, None


def walk_docs(docs_dir: Path) -> list[Path]:
    files: list[Path] = []
    for path in docs_dir.rglob("*.md"):
        if any(part.startswith(".") or part in EXCLUDED_DIRS for part in path.relative_to(docs_dir).parts):
            continue
        files.append(path)
    return sorted(files, key=lambda item: item.relative_to(docs_dir).as_posix())


def main() -> int:
    docs_dir = Path("docs")
    if not docs_dir.exists():
        print("docs:list: missing docs directory. Run from repo root.", file=sys.stderr)
        return 1
    if not docs_dir.is_dir():
        print("docs:list: docs path is not a directory.", file=sys.stderr)
        return 1

    print("Listing all markdown files in docs folder:")
    for path in walk_docs(docs_dir):
        rel = path.relative_to(docs_dir).as_posix()
        frontmatter, error = parse_frontmatter(path)
        summary = str(frontmatter.get("summary", "")).strip()
        read_when = frontmatter.get("read_when", [])
        read_when_values = compact_strings(read_when) if isinstance(read_when, list) else []

        if summary:
            print(f"{rel} - {summary}")
            if read_when_values:
                print(f"  Read when: {'; '.join(read_when_values)}")
        else:
            reason = error or "summary key missing"
            print(f"{rel} - [{reason}]")

    print(
        '\nReminder: when a task matches any "Read when" hint above, read that doc before coding and update docs when behavior changes.'
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
