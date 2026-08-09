---
name: release-publish
description: Use to cut and publish a new ts-pptx version (any "do a release", "minor/major/patch release", "publish vX.Y.Z", "ship a release" request in this repo). Encodes the exact release flow — CHANGELOG, the `pnpm version` bump that writes the other two version files and tags, and the GitHub Release that triggers CI. IMPORTANT — publishing to npm is done by CI (trusted publishing), never by running `npm publish` locally. Do not run `npm publish`, `npm login`, or `npm token` for a release.
metadata:
  # For working *on* ts-pptx, not *with* it. `npx skills add shbernal/ts-pptx` walks
  # .claude/skills/ (a symlink to this tree) as well as the published skills/, and this flag
  # is what keeps it out of the menu a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install
  # it anyway.
  internal: true
---

# Releasing & Publishing ts-pptx

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

So the whole job is: write the CHANGELOG, let `pnpm version` bump and tag, and cut
a Release. CI does the rest (and re-runs every lint/typecheck/test/pack gate before
it publishes).

### The one historical exception (already spent — do not repeat it)

npm cannot configure a trusted publisher for a package that does not yet exist:
the setting lives on the package's settings page, so the package has to be on the
registry first ([npm/cli#8544](https://github.com/npm/cli/issues/8544)). Bootstrapping
`@shbernal/ts-pptx` therefore took one local `npm publish` of a throwaway `0.0.1`
(since deprecated) purely to create the package; trusted publishing was enabled
immediately after, and `1.0.0` onward went through CI with provenance.

That exception is spent. It applies again only if this project is ever published
under a **new name or scope**. For every release of the existing package, the rule
above holds without qualification: no local `npm publish`.

## Choosing the version (SemVer)

`package.json` is on major `1.x`. Pick the bump from what landed since the last
release (`git log --oneline vLAST..HEAD`) — but the user's explicit ask wins:

- **patch** (`X.Y.Z+1`) — only `fix:` / internal changes, no API surface change.
- **minor** (`X.Y+1.0`) — any `feat:` / additive API, or a fix the user wants
  shipped as a minor.
- **major** (`X+1.0.0`) — a breaking API change (this project allows them; see
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

### 2. Write the CHANGELOG entry and stage it

The version lands in three files, but you only edit one. `package.json` is the
version of record and `src/presentation.ts` is derived from it — `pnpm version` in
step 4 bumps the first and rewrites the second. **Do not hand-edit either one**;
doing so puts you on a path where `pnpm version` refuses to run.

- `CHANGELOG.md` → convert the `## [Unreleased]` heading region into a released
  section, keeping an empty `## [Unreleased]` above it:

  ```
  ## [Unreleased]

  ## [X.Y.Z](https://github.com/shbernal/ts-pptx/releases/tag/vX.Y.Z) - YYYY-MM-DD

  ### Fixed
  - ...
  ### Added
  - ...
  ```

Stage it, and only it:

```bash
git add CHANGELOG.md
```

### 3. Verify locally before committing

The publish workflow does not define its own gate — it calls `ci.yml` via
`workflow_call` and publishes only if that passes. So the thing to reproduce
locally is CI, and CI is two commands. Catch failures now, not in a half-finished
Release run:

```bash
pnpm run check:static && pnpm run verify:full
```

(The commit's own pre-commit/pre-push hooks also run oxlint, oxfmt, and
typecheck, so `check:static` should be a formality.)

### 4. Bump, commit and tag — one command

```bash
pnpm version <major|minor|patch> --message 'chore(release): v%s' --no-git-checks
git push origin master
git push origin vX.Y.Z
```

`pnpm version` bumps `package.json`, runs the `version` lifecycle script
(`scripts/sync-version.mjs`) which rewrites `const VERSION` in
`src/presentation.ts` and stages it, then makes **one** commit holding all three
files — the staged `CHANGELOG.md` rides along — and creates the annotated tag.
Do not bump by hand; that is what step 2 was avoiding.

Both flags matter, and both were checked rather than assumed:

- **`--message`** — pnpm's default subject is the bare version (`3.2.0`). The
  flag produces `chore(release): vX.Y.Z`, and the annotated tag gets the same
  text as its message. pnpm does **not** read npm's `message` config, so putting
  it in an `.npmrc` will not work.
- **`--no-git-checks`** — pnpm refuses to run against anything but a spotless
  tree (`ERR_PNPM_UNCLEAN_WORKING_TREE`), and the `CHANGELOG.md` you staged in
  step 2 counts as unclean. Without this flag the release becomes two commits.
  It waives the check for *everything*, which is why step 1 insists on a clean
  `git status` and no `git add -A`.

Commit-message rules: subject `chore(release): vX.Y.Z`, **no** AI/Claude
attribution footer and **no** cross-repo file references (see the repo's commit
conventions).

If `pnpm version` fails partway, check `git status` before retrying — the bump,
the constant rewrite and the commit are separate steps, so it can leave the two
version files ahead of the commit. `pnpm run version:check` says whether the
constant and the manifest agree.

### 5. Create the GitHub Release (this is what publishes)

Body = the version's CHANGELOG section followed by a full-changelog link. Match
prior releases (`gh release view vLAST`).

Write the notes to a file with your file-writing tool, then point `gh` at it —
never build the body with a shell here-doc. This is the same rule, and the same
reason, as the commit-message convention above: the POSIX and PowerShell dialects
disagree (`<<'EOF'` vs `@'…'@`), picking the wrong one does not error, and the
delimiter ends up as literal text in the published release notes.

```bash
gh release create vX.Y.Z --title vX.Y.Z --notes-file <path/to/notes.md>
```

Notes body:

```markdown
### Fixed

- <the changelog bullet(s) for this version>

**Full changelog:** https://github.com/shbernal/ts-pptx/blob/vX.Y.Z/CHANGELOG.md
```

### 6. Watch the publish run

```bash
gh run list --workflow=publish.yml --limit 1
gh run watch <run-id>      # or: gh run view <run-id> --log-failed
```

Confirm it reaches "Publish to npm" and succeeds. Only then is the release done.
Sanity check: `npm view @shbernal/ts-pptx version` should report `X.Y.Z`.

## If the publish run fails

- **Tag/version mismatch** or **already published** → the tag and `package.json`
  disagree, or you reused a version. (`pnpm run version:check` covers the other
  pair, the manifest against `src/presentation.ts`.) Fix the version, re-tag, and
  cut a new Release. Never work around it by publishing locally.
- **A lint/test/pack gate failed** → fix on `master`, then either bump to the next
  patch and re-release, or (if the tag content is still what you want) delete and
  recreate the tag + Release after pushing the fix. Prefer a fresh patch version
  over force-moving a published-looking tag.
- The workflow can also be re-run manually via `workflow_dispatch` from the tag if
  the failure was transient (`gh workflow run publish.yml --ref vX.Y.Z`).
