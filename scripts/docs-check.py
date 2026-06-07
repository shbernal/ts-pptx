#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import urlparse


ALLOWED_DOC_TYPES = {
    "overview",
    "architecture",
    "guide",
    "reference",
    "troubleshooting",
    "decision",
    "runbook",
}
REQUIRED_FIELDS = {"doc-schema-version", "title", "summary", "read_when", "doc_type"}
EXCLUDED_DIRS = {"archive", "research"}
MARKDOWN_LINK_RE = re.compile(r"(?<!!)\[[^\]]+\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")


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


def route_for(relative_path: str) -> set[str]:
    stem = re.sub(r"\.md$", "", relative_path)
    routes = {normalize_route(stem)}
    if stem == "index":
        routes.add("/")
    if stem.endswith("/index"):
        routes.add(normalize_route(stem[: -len("/index")]))
    return routes


def normalize_route(route: str) -> str:
    route = route.split("#", 1)[0].split("?", 1)[0].strip("/")
    return f"/{route}" if route else "/"


def collect_nav_pages(value: object) -> list[str]:
    pages: list[str] = []
    if isinstance(value, list):
        for item in value:
            pages.extend(collect_nav_pages(item))
    elif isinstance(value, dict):
        for key, item in value.items():
            if key == "pages" and isinstance(item, list):
                for page in item:
                    if isinstance(page, str):
                        pages.append(page)
                    else:
                        pages.extend(collect_nav_pages(page))
            else:
                pages.extend(collect_nav_pages(item))
    return pages


def check_frontmatter(path: Path, docs_dir: Path) -> list[str]:
    rel = path.relative_to(docs_dir).as_posix()
    frontmatter, parse_error = parse_frontmatter(path)
    errors: list[str] = []
    if parse_error:
        return [f"{rel}: {parse_error}"]

    missing = sorted(REQUIRED_FIELDS - set(frontmatter))
    for field in missing:
        errors.append(f"{rel}: missing frontmatter field `{field}`")

    schema_version = str(frontmatter.get("doc-schema-version", "")).strip()
    if schema_version != "1":
        errors.append(f"{rel}: `doc-schema-version` must be 1")

    for field in ["title", "summary", "doc_type"]:
        value = str(frontmatter.get(field, "")).strip()
        if not value:
            errors.append(f"{rel}: `{field}` must be non-empty")

    doc_type = str(frontmatter.get("doc_type", "")).strip()
    if doc_type and doc_type not in ALLOWED_DOC_TYPES:
        allowed = ", ".join(sorted(ALLOWED_DOC_TYPES))
        errors.append(f"{rel}: `doc_type` must be one of: {allowed}")

    read_when = frontmatter.get("read_when", [])
    if not isinstance(read_when, list) or not compact_strings(read_when):
        errors.append(f"{rel}: `read_when` must contain at least one hint")

    return errors


def check_code_fences(path: Path, docs_dir: Path) -> list[str]:
    rel = path.relative_to(docs_dir).as_posix()
    fence_count = 0
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.lstrip().startswith("```"):
            fence_count += 1
    return [f"{rel}: unbalanced fenced code block"] if fence_count % 2 else []


def check_docs_json(docs_dir: Path, markdown_files: list[Path]) -> list[str]:
    config_path = docs_dir / "docs.json"
    if not config_path.exists():
        return ["docs/docs.json: missing docs navigation file"]

    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        return [f"docs/docs.json: invalid JSON at line {error.lineno}: {error.msg}"]

    errors: list[str] = []
    page_paths = {path.relative_to(docs_dir).as_posix() for path in markdown_files}
    page_keys = {re.sub(r"\.md$", "", path) for path in page_paths}
    for page in collect_nav_pages(config.get("navigation", [])):
        key = page.strip().strip("/")
        if key not in page_keys:
            errors.append(f"docs/docs.json: navigation page `{page}` has no matching docs page")

    return errors


def is_external(target: str) -> bool:
    parsed = urlparse(target)
    return bool(parsed.scheme) or target.startswith(("mailto:", "tel:"))


def check_links(path: Path, docs_dir: Path, routes: set[str]) -> list[str]:
    rel = path.relative_to(docs_dir).as_posix()
    text = path.read_text(encoding="utf-8")
    errors: list[str] = []

    for match in MARKDOWN_LINK_RE.finditer(text):
        target = match.group(1).strip()
        if not target or target.startswith("#") or is_external(target):
            continue

        target_path = target.split("#", 1)[0].split("?", 1)[0]
        if target_path.startswith("/"):
            route = normalize_route(target_path)
            if route not in routes:
                errors.append(f"{rel}: broken docs route `{target}`")
            continue

        resolved = (path.parent / target_path).resolve()
        if not str(resolved).startswith(str(docs_dir.resolve())):
            continue
        if target_path.endswith(".md") and not resolved.exists():
            errors.append(f"{rel}: broken relative link `{target}`")

    return errors


def main() -> int:
    docs_dir = Path("docs")
    if not docs_dir.exists() or not docs_dir.is_dir():
        print("docs:check: missing docs directory. Run from repo root.", file=sys.stderr)
        return 1

    markdown_files = walk_docs(docs_dir)
    routes: set[str] = set()
    for path in markdown_files:
        routes.update(route_for(path.relative_to(docs_dir).as_posix()))

    errors: list[str] = []
    errors.extend(check_docs_json(docs_dir, markdown_files))
    for path in markdown_files:
        errors.extend(check_frontmatter(path, docs_dir))
        errors.extend(check_code_fences(path, docs_dir))
        errors.extend(check_links(path, docs_dir, routes))

    if errors:
        for error in errors:
            print(f"docs:check: {error}", file=sys.stderr)
        return 1

    print(f"docs:check: ok ({len(markdown_files)} docs page(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
