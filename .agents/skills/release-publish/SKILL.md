---
name: release-publish
description: Use to cut and publish a new PptxGenJS version (any "do a release", "minor/major/patch release", "publish vX.Y.Z", "ship a release" request in this repo). Encodes the exact release flow — version bump in three files, CHANGELOG, tag, and the GitHub Release that triggers CI. IMPORTANT — publishing to npm is done by CI (trusted publishing), never by running `npm publish` locally. Do not run `npm publish`, `npm login`, or `npm token` for a release.
---

# Releasing & Publishing PptxGenJS

npm publishing is **fully automated in CI**. A release is finished by creating a
**GitHub Release** for a `vX.Y.Z` tag; that fires `.github/workflows/publish.yml`,
which publishes to npm via **trusted publishing (OIDC)** — no npm token, no
`npm login`, no local `npm publish`. If you ever type `npm publish` for a release,
you are doing it wrong: stop and create the GitHub Release instead.

## The one rule that trips people up

Do **not** publish locally. The publish workflow guards on two things and will
fail the run if either is off:

- the release tag must equal `v<package.json version>` exactly, and
- that version must not already exist on npm.

So the whole job is: get the version consistent across three files, tag it, and
cut a Release. CI does the rest (and re-runs every lint/typecheck/test/pack gate
before it publishes).

## Choosing the version (SemVer)

`package.json` is on major `10.x`. Pick the bump from what landed since the last
release (`git log --oneline vLAST..HEAD`) — but the user's explicit ask wins:

- **patch** (`X.Y.Z+1`) — only `fix:` / internal changes, no API surface change.
- **minor** (`X.Y+1.0`) — any `feat:` / additive API, or a fix the user wants
  shipped as a minor.
- **major** (`X+1.0.0`) — a breaking API change (this fork allows them; see
  `CHANGELOG.md` / AGENTS.md "API Evolution Policy").

If the user says "minor release," honor that even for a lone fix — don't second-guess
into a patch.

## Steps

Assume today's date is available; use `YYYY-MM-DD` in the CHANGELOG.

### 1. Confirm the working tree

- Run `git status`. **Leave scratch/plan files untracked** — never `git add -A`.
  Only the three release files below go in the release commit.
- `git log --oneline vLAST..HEAD` — every commit here must be represented in the
  CHANGELOG `[Unreleased]` section. If a fix/feature landed without a CHANGELOG
  entry, write one now (match the dense, prose style of existing entries).

### 2. Bump the version in THREE places (all must match)

- `package.json` → `"version": "X.Y.Z"`
- `src/pptxgen.ts` → `const VERSION = 'X.Y.Z'`
- `CHANGELOG.md` → convert the `## [Unreleased]` heading region into a released
  section, keeping an empty `## [Unreleased]` above it:

  ```
  ## [Unreleased]

  ## [X.Y.Z](https://github.com/shbernal/PptxGenJS/releases/tag/vX.Y.Z) - YYYY-MM-DD

  ### Fixed
  - ...
  ### Added
  - ...
  ```

### 3. Verify locally before committing

The publish workflow runs the full gate set (lint, format:check, typecheck,
typecheck:scripts, test:coverage, build, package:lint, pack:check, test:package,
test:demos). Catch failures now, not in a half-finished Release run:

```bash
pnpm run build && pnpm run typecheck && pnpm test
```

Run `pnpm run lint && pnpm run format:check` too if you touched anything beyond the
version strings. (The commit's own pre-commit/pre-push hooks also run eslint,
prettier, and typecheck.)

### 4. Commit, tag, push

```bash
git add CHANGELOG.md package.json src/pptxgen.ts
git commit -m "chore(release): vX.Y.Z"
git tag vX.Y.Z -m "vX.Y.Z"
git push origin master
git push origin vX.Y.Z
```

Commit-message rules: subject `chore(release): vX.Y.Z`, **no** AI/Claude
attribution footer and **no** cross-repo file references (see the repo's commit
conventions). Tag is annotated, matching prior release tags.

### 5. Create the GitHub Release (this is what publishes)

Body = the version's CHANGELOG section followed by a full-changelog link. Match
prior releases (`gh release view vLAST`):

```bash
cat > /tmp/notes.md <<'EOF'
### Fixed

- <the changelog bullet(s) for this version>

**Full changelog:** https://github.com/shbernal/PptxGenJS/blob/vX.Y.Z/CHANGELOG.md
EOF
gh release create vX.Y.Z --title vX.Y.Z --notes-file /tmp/notes.md
```

### 6. Watch the publish run

```bash
gh run list --workflow=publish.yml --limit 1
gh run watch <run-id>      # or: gh run view <run-id> --log-failed
```

Confirm it reaches "Publish to npm" and succeeds. Only then is the release done.
Sanity check: `npm view @shbernal/pptxgenjs version` should report `X.Y.Z`.

## If the publish run fails

- **Tag/version mismatch** or **already published** → the three files disagree or
  you reused a version. Fix the version, re-tag, and cut a new Release. Never work
  around it by publishing locally.
- **A lint/test/pack gate failed** → fix on `master`, then either bump to the next
  patch and re-release, or (if the tag content is still what you want) delete and
  recreate the tag + Release after pushing the fix. Prefer a fresh patch version
  over force-moving a published-looking tag.
- The workflow can also be re-run manually via `workflow_dispatch` from the tag if
  the failure was transient (`gh workflow run publish.yml --ref vX.Y.Z`).
