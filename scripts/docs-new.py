#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


ALLOWED_DOC_TYPES = {
    "overview",
    "architecture",
    "guide",
    "reference",
    "troubleshooting",
    "decision",
    "runbook",
}


def title_from_slug(slug: str) -> str:
    name = Path(slug).stem if slug.endswith(".md") else Path(slug).name
    return " ".join(part.capitalize() for part in re.split(r"[-_]+", name) if part)


def page_key(slug: str) -> str:
    key = slug.strip().strip("/")
    if key.startswith("docs/"):
        key = key[len("docs/") :]
    if key.endswith(".md"):
        key = key[:-3]
    return key


def page_path(docs_dir: Path, slug: str) -> Path:
    key = page_key(slug)
    path = docs_dir / f"{key}.md"
    resolved = path.resolve()
    if not str(resolved).startswith(str(docs_dir.resolve())):
        raise ValueError("page slug must stay under docs/")
    return path


def frontmatter(title: str, summary: str, doc_type: str, read_when: list[str]) -> str:
    hints = "\n".join(f"  - {hint}" for hint in read_when)
    return f"""---
doc-schema-version: 1
title: "{title}"
summary: "{summary}"
read_when:
{hints}
doc_type: "{doc_type}"
---

# {title}

Add the useful project-specific content here.
"""


def collect_groups(value: object) -> list[dict[str, object]]:
    groups: list[dict[str, object]] = []
    if isinstance(value, list):
        for item in value:
            groups.extend(collect_groups(item))
    elif isinstance(value, dict):
        if isinstance(value.get("group"), str) and isinstance(value.get("pages"), list):
            groups.append(value)
        for item in value.values():
            groups.extend(collect_groups(item))
    return groups


def add_to_nav(docs_json_path: Path, key: str, group_name: str) -> None:
    if not docs_json_path.exists():
        return

    config = json.loads(docs_json_path.read_text(encoding="utf-8"))
    navigation = config.setdefault("navigation", [])
    if not isinstance(navigation, list):
        raise ValueError("docs/docs.json navigation must be a list")

    groups = collect_groups(navigation)
    target_group = next((group for group in groups if group.get("group") == group_name), None)
    if target_group is None:
        target_group = {"group": group_name, "pages": []}
        navigation.append(target_group)

    pages = target_group.setdefault("pages", [])
    if not isinstance(pages, list):
        raise ValueError(f"docs/docs.json group `{group_name}` pages must be a list")
    if key not in pages:
        pages.append(key)

    docs_json_path.write_text(json.dumps(config, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a docs page with standard frontmatter.")
    parser.add_argument("slug", help="Docs slug, for example `guides/setup` or `reference/cli`.")
    parser.add_argument("--title", help="Display title. Defaults to a title derived from the slug.")
    parser.add_argument("--summary", help="One-sentence page summary.")
    parser.add_argument("--type", choices=sorted(ALLOWED_DOC_TYPES), default="guide", dest="doc_type")
    parser.add_argument("--read-when", action="append", default=[], help="Routing hint. May be repeated.")
    parser.add_argument("--nav-group", help="Append the page to this docs/docs.json navigation group.")
    parser.add_argument("--force", action="store_true", help="Overwrite an existing docs page.")
    args = parser.parse_args()

    docs_dir = Path("docs")
    docs_dir.mkdir(parents=True, exist_ok=True)

    key = page_key(args.slug)
    title = args.title or title_from_slug(key)
    summary = args.summary or f"Documentation page for {title}."
    read_when = args.read_when or [f"Working on {title}"]
    target = page_path(docs_dir, key)

    if target.exists() and not args.force:
        print(f"docs:new: {target} already exists; pass --force to overwrite", file=sys.stderr)
        return 1

    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(frontmatter(title, summary, args.doc_type, read_when), encoding="utf-8")

    if args.nav_group:
        add_to_nav(docs_dir / "docs.json", key, args.nav_group)

    print(f"docs:new: wrote {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
