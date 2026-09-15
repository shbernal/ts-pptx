---
name: release-publish
description: Use to cut and publish a new ts-pptx version (any "do a release", "minor/major/patch release", "publish vX.Y.Z", "ship a release" request in this repo). Encodes the exact release flow: the comparison refresh, the CHANGELOG, the `pnpm version` bump that writes the other two version files and tags, and the GitHub Release that triggers CI. IMPORTANT: publishing to npm is done by CI (trusted publishing), never by running `npm publish` locally. Do not run `npm publish`, `npm login`, or `npm token` for a release.
metadata:
  # For working *on* ts-pptx, not *with* it. `npx skills add shbernal/ts-pptx` walks
  # .claude/skills/ (a symlink to this tree) as well as the published skills/, and this flag
  # is what keeps it out of the menu a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install
  # it anyway.
  internal: true
---

# Releasing and publishing ts-pptx

CI publishes to npm. A release is finished by publishing a **GitHub Release** for a `vX.Y.Z` tag,
which fires `.github/workflows/publish.yml`. The workflow authenticates through trusted publishing
(OIDC), so there is no npm token and no `npm login`.

**Never run `npm publish` locally**, not for a release and not to repair a failed one. If you are
about to type it, stop and use the workflow instead.

A release publishes `pptx-ts` and then `@shbernal/ts-pptx`, the same build staged under a second
name by `scripts/alias-package.mjs`. You do not bump, tag or stage the alias, and there is one
changelog.

`docs/contributing/releasing.md` explains why each step is shaped the way it is: the snapshot
review rubric, the `pnpm version` flags, the workflow guards. This skill is the sequence.

## Choosing the version (SemVer)

Read the current version from `package.json`, and pick the bump from what landed since the last
release (`git log --oneline vLAST..HEAD`). The user's explicit ask wins:

- **patch** (`X.Y.Z+1`): only `fix:` or internal changes, no API change.
- **minor** (`X.Y+1.0`): any `feat:` or additive API, or a fix the user wants shipped as a minor.
- **major** (`X+1.0.0`): a breaking API change (allowed here; see AGENTS.md "API evolution policy").

If the user says "minor release", do that even for a lone fix.

## Steps

Use today's date as `YYYY-MM-DD` in the CHANGELOG.

### 1. Confirm the working tree

- Run `git status`. Leave scratch and plan files untracked, and never `git add -A`. Only the
  release files go in the release commit.
- `git log --oneline vLAST..HEAD`: every commit must be represented under `[Unreleased]` in
  `CHANGELOG.md`. Write a missing entry now, in the dense prose style of the existing ones.

### 2. Refresh the comparison

```bash
pnpm run comparison:measure      # network: npm install, GitHub and npm APIs; about a minute
pnpm run comparison:render
```

Leave the machine idle while `comparison:measure` runs, because it times both libraries. Never
pass `--allow-unavailable`; if a fetch fails, re-run later.

Read the diff of `scripts/comparison/snapshot.json` against "What to stop on in the snapshot diff"
in `docs/contributing/releasing.md` before staging. Stop and report to the user on a flipped
coverage row, a moved validity count, an upstream version bump, or a timing row that changed sign.

```bash
git add scripts/comparison/snapshot.json docs/comparison.md docs/comparison-method.md docs/comparison-syntax.md README.md
```

### 3. Write the CHANGELOG entry and stage it

Do not hand-edit `package.json` or `src/presentation.ts`. `pnpm version` in step 5 bumps the first
and rewrites the second, and a hand edit puts you on a path where it refuses to run.

Convert the `## [Unreleased]` region into a released section, keeping an empty `## [Unreleased]`
above it:

```
## [Unreleased]

## [X.Y.Z](https://github.com/shbernal/ts-pptx/releases/tag/vX.Y.Z) - YYYY-MM-DD

### Fixed
- ...
### Added
- ...
```

```bash
git add CHANGELOG.md
```

### 4. Run the local gate

The publish workflow's gate is `ci.yml`, and CI is these two commands:

```bash
pnpm run check:static && pnpm run verify:full
```

Then confirm the version is new under both names. Both commands should fail with a registry 404:

```bash
npm view pptx-ts@X.Y.Z version
npm view @shbernal/ts-pptx@X.Y.Z version
```

### 5. Bump, commit and tag

```bash
pnpm version <major|minor|patch> --message 'chore(release): v%s' --no-git-checks
```

This makes one commit holding `package.json`, the rewritten `VERSION` constant and the staged
`CHANGELOG.md`, and creates the annotated tag. The commit subject is `chore(release): vX.Y.Z`, with
no AI attribution footer and no cross-repo file references.

If it fails partway, check `git status` before retrying: the bump, the constant rewrite and the
commit are separate steps, so the version files can end up ahead of the commit.
`pnpm run version:check` says whether the constant and the manifest agree.

### 6. Push

```bash
git push origin master
git push origin vX.Y.Z
```

### 7. Create the GitHub Release (this is what publishes)

The body is the version's CHANGELOG section followed by a full-changelog link. Match prior releases
(`gh release view vLAST`). Write the notes to a file with your file-writing tool and pass the file.
Never build the body with a shell here-doc: the POSIX and PowerShell dialects disagree, and a
delimiter in the wrong one lands in the published notes.

```bash
gh release create vX.Y.Z --title vX.Y.Z --notes-file <path/to/notes.md>
```

```markdown
### Fixed

- <the changelog bullet(s) for this version>

**Full changelog:** https://github.com/shbernal/ts-pptx/blob/vX.Y.Z/CHANGELOG.md
```

### 8. Watch the publish run

```bash
gh run list --workflow=publish.yml --limit 1
gh run watch <run-id>      # or: gh run view <run-id> --log-failed
```

The release is done when "Publish to npm" and "Publish scoped alias" both succeed. A step reported
as skipped means that name already had the version, which is the expected shape of a re-dispatch.

### 9. Post-publish checks

```bash
npm view pptx-ts@X.Y.Z version dist-tags --json
npm view @shbernal/ts-pptx@X.Y.Z version dist-tags --json
gh release view vX.Y.Z --repo shbernal/ts-pptx
```

Comment on every issue the release closes (the CHANGELOG entry cites them):

```bash
gh issue comment <N> --repo shbernal/ts-pptx --body "Released in X.Y.Z."
```

After a major release, open an issue on `pptx-html` asking for a matching release. Do not hold the
release for it.

## If the publish run fails

- **Tag and version mismatch, or both names already published**: the tag and `package.json`
  disagree, or the version was reused. Fix the version, re-tag, and cut a new Release.
- **A CI gate failed**: fix on `master`, then bump to the next patch and release that. Prefer a
  fresh patch over force-moving a tag that has been pushed.
- **The canonical publish succeeded and the alias failed**: re-dispatch the workflow on the same
  tag. The guard passes because the alias still lacks the version, and the `pptx-ts` step skips.

  ```bash
  gh workflow run publish.yml --repo shbernal/ts-pptx --ref vX.Y.Z
  ```

- **A transient failure**: the same `gh workflow run` command re-runs the workflow from the tag.
