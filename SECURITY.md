# Security Policy

## Reporting a vulnerability

**Please do not report security vulnerabilities through public GitHub issues,
pull requests, or discussions.**

Instead, use GitHub's private vulnerability reporting:

1. Go to the repository's **Security** tab.
2. Click **Report a vulnerability** (under *Advisories*).
3. Describe the issue, including a minimal reproduction if possible (a small
   script and, where relevant, the resulting `.pptx`).

This routes the report privately to the maintainer. Please allow a reasonable
window for a fix before any public disclosure.

## Scope

PptxGenJS is a Node-first library that generates PowerPoint `.pptx` packages from
untrusted input (text, colors, image data, table data). Relevant classes of
issue include, for example:

- Crafted input that causes the generator to emit a malformed package, hang, or
  consume unbounded memory/CPU.
- Path- or entry-name handling in the read/round-trip subsystem (`src/read/`)
  that could escape the package boundary.
- Dependency vulnerabilities that are reachable through this package's API.

Out of scope: issues that only manifest in the two out-of-active-scope domains
documented in [AGENTS.md](AGENTS.md) (live-DOM/browser-layout features and
third-party office-suite interop quirks), unless they represent a genuine
security defect in the OOXML this library itself emits.

## Supported versions

This project ships from `master` and releases roll forward; fixes land in the
next release rather than being backported. Please verify against the latest
published version before reporting.
