---
name: charcheck-upstream
description: Report a charcheck bug, a silent miss, or a wrong `--fix` rewrite to its GitHub tracker from a project that depends on it. Use when charcheck passes a file you know holds a banned character, when it flags a region its scope claims not to read, when `--fix` changes prose you did not want changed, when it crashes on a file valid for its language, when a finding's line or column does not point at the character, and above all whenever you are about to add an `exclude` glob, write a suppression comment, or drop a rule to `raw` so that charcheck stops reporting something. Filing the bug is the fix; the suppression is the stopgap.
metadata:
  # Local addition, not upstream's. Everything under .agents/skills/ carries this so that
  # `npx skills add shbernal/ts-pptx` does not offer it to a ts-pptx consumer: this skill
  # is about *charcheck's* tracker, and a consumer of ts-pptx need not depend on charcheck
  # at all. `npx skills update charcheck-upstream` overwrites the file wholesale, so
  # re-apply these three lines after every refresh.
  internal: true
---

# Reporting a charcheck problem upstream

You are in a project that _uses_ `charcheck`, not the project that builds it. This skill is
how a defect you hit here becomes a permanent regression test there.

**Nothing will throw to send you here.** charcheck's characteristic failure is silence: a
rule that reads the wrong part of a file, or that opens no file at all, reports a clean run
and exits 0. A scan that reads nothing looks exactly like a scan that passed. So the moment
that should load this skill is not an exception, it is a suspicion, and most often it is
this one:

> I am about to add an `exclude` glob, write a suppression comment, or change a rule's
> scope to `raw`, so that charcheck stops reporting this.

That is the workaround moment, and it is the one worth stopping at. Silencing a true
finding and silencing a bug look identical from inside this repository.

Filing here is unusually cheap, because a reproduction is a config and one short file. That
cuts both ways: the cost of a report is minutes, and the tracker fills with non-bugs unless
the triage below actually happens. Most of this skill is that triage.

## 1. Decide whether it is actually ours

### The four steps, in order

**Step 1: did the rule open the file at all?** This is the commonest answer by a distance,
and charcheck already tells you on stderr:

```
charcheck: rule "no-em-dash-in-markup" matched no files: site/**/*.vue. Check the
globs; a dotted directory is only entered when a pattern names it.
```

A dotted directory is the usual cause. `site/**/*.vue` never reaches
`site/.vitepress/theme/Card.vue`, and `docs/**` never walks into `.github`. Name the
directory: `site/.vitepress/**/*.vue`. Two related traps that produce the same silence:
positional paths are intersected with each rule's `include` rather than substituted for it,
so `charcheck src/` scans nothing when no rule targets `src/`; and the warning is not raised
under `--staged` for a rule that merely had nothing staged, since it describes the globs
rather than the commit. Full text under
[The rule never opened the file](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#the-rule-never-opened-the-file).

**Step 2: can the rule's `scope` see that region?** A scope is an allowlist of what a rule
may match inside, and text outside it is not a miss, it is the contract:

| Scope           | Reads                                                | Never reads                                                 |
| --------------- | ---------------------------------------------------- | ----------------------------------------------------------- |
| `raw` (default) | The whole file                                       | Nothing is out of reach                                     |
| `strings`       | String and template literals                         | Comments, identifiers, JSX text                             |
| `markup`        | Template text, allowlisted attributes, script blocks | `<style>`, custom blocks, `v-html`                          |
| `markdown`      | Prose                                                | Fenced and inline code, link targets, HTML blocks           |
| `html`          | Page text, allowlisted attributes, script literals   | `<style>`, comments, `<code>`, `<pre>`, unlisted attributes |

The two that catch people: a `strings` rule cannot see a comment, and a `markdown` rule
cannot see a fenced block. See
[The scope cannot see that region](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#the-scope-cannot-see-that-region).

**Step 3: is it already a stated limitation?**
[Limitations](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md) is
unusually complete, and it covers most apparent misses. The ones reached for most often:

| You saw                                          | Stated at                                                                                                                                            |
| ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| Text in an HTML block inside a `.md` went unread | [HTML block inside Markdown](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#an-html-block-inside-markdown-is-skipped)           |
| A fence's `title="..."` went unread              | [Language tag and meta](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#a-fences-language-tag-and-meta-are-skipped)              |
| Text in a CSS `content` property went unread     | [`<style>` is never read](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#style-is-never-read)                                   |
| An `<i18n>` block went unread                    | [Vue custom blocks](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#vue-custom-blocks-are-skipped)                               |
| A `.mdx` file was not scanned                    | [`.mdx` is not reachable](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#mdx-is-not-reachable)                                  |
| `v-html` content went unread                     | [`v-html`](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#v-html-content-is-not-reachable)                                      |
| A Jinja or Handlebars expression was flagged     | [Template languages in `html`](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#a-template-language-inside-html-is-read-as-prose) |
| A `.tsx` file was refused, exit code 2           | [JSX on TypeScript 7](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#tsx-and-jsx-are-refused-on-typescript-7)                   |
| A pattern after a labelled block was misread     | [Labelled block](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#a-labelled-block-is-read-as-an-object-literal)                  |
| A staged-then-deleted file was not scanned       | [Git](https://github.com/shbernal/charcheck/blob/main/docs/limitations.md#a-file-staged-as-added-and-then-deleted-is-not-scanned)                    |

**Step 4: only now is it ours.** Three steps ruled out, and the behaviour still contradicts
what a scope or a document promises. File it.

### Always ours, file without further deliberation

- **`UnknownTokenKindError`.** The closest thing charcheck has to an internal error. It
  means a TypeScript major renamed a `SyntaxKind`, and the alternative to this error is a
  scan that quietly finds nothing. The message asks for the TypeScript version; give it.
- **`UnsupportedPeerDependencyError` when the installed TypeScript is _newer_ than the
  supported range.** The parser moved again and charcheck has not caught up. That is
  upstream work, never yours. (An _older_ TypeScript is your install to fix.)
- **Any wrong `--fix` output.** A linter that rewrites prose incorrectly is worse than one
  that misses, and this has no equivalent in a tool that only reports. If `--fix` produced a
  comma splice, broke a sentence, changed a line ending, or touched anything outside the
  matched span, that is a bug of the highest severity here. Report it even if you have
  already reverted.
- **A false positive under any scope other than `raw`.** The scope said it would not read
  that region and it did. Always an extractor bug.
- **A crash on a file that is valid for its language.** Valid meaning its own parser accepts
  it, not meaning it is tidy.
- **A finding whose `line` or `column` does not point at the character.** Positions are
  1-based UTF-16 code units and are supposed to agree with an editor's cursor. An escape
  sequence, a CRLF file, or an astral character nearby are the usual triggers.

### Yours, not ours

These carry no tracker pointer on purpose. Filing one costs a maintainer the same triage as
a real report and ends in a close:

| Error                                | What to do                                                           |
| ------------------------------------ | -------------------------------------------------------------------- |
| `ConfigError`, `ConfigNotFoundError` | Fix or find the config. The message names the key.                   |
| `RuleError`                          | Fix the rule. Usually a pattern the scope's extensions cannot cover. |
| `UnsupportedScopeError`              | The scope cannot read that file type. Split the rule per surface.    |
| `MissingPeerDependencyError`         | Install the peer the scope needs, as a devDependency.                |
| `JsxUnsupportedError`                | A stated limitation. Use TypeScript 5 or 6, or exclude those files.  |
| `GitError`                           | Environment. `--staged` needs a repository and a readable index.     |
| A usage error, exit code 2           | Read the message. Conflicting flags are named explicitly.            |

### Decided against, so not a bug and not a feature request

- **Vocabulary rules.** Word lists, "delve" detectors, anything about which words are used.
  charcheck is about characters and ships no vocabulary opinions.
- **Natural-language understanding.** Anything requiring charcheck to know what a sentence
  means.

A **missing scope** is different: it is wanted, and it is a feature request rather than a
bug. Svelte is openly asked for. Use
[the feature request form](https://github.com/shbernal/charcheck/issues/new?template=feature_request.yml),
not the one below.

## 2. Collect the facts

One command collects nearly all of them, already in the shape section 5 files:

```bash
pnpm exec charcheck --report-issue > <a path this repo already ignores>/charcheck-report.md
```

It writes the charcheck, Node, operating system and peer versions, and then the fact that
decides most reports: **each rule as it resolved, not the config as it was written**,
including **how many files each one matched**. A pasted config hides the two things that
explain nearly every silent miss, which are what the globs actually reached and which rules
matched nothing at all. If the matched-no-files warning fired, that count is zero and it is
the finding.

It reads no file's content, exits 0 whatever your tree holds, and refuses to combine with any
flag that would select files or shape a report of findings. Add `--config <path>` if the
config is not the one found from here.

The sections you have to write yourself are left as bracketed placeholders. Two of them are
the ones the command cannot know: **the exact command you ran, and its full output verbatim
with the exit code.** Do not paraphrase the output. The throw site is often identifiable from
the exact wording, and the stderr warnings are where the matched-no-files case lives.

### The globs are anonymized for you, and there is nothing to approve

A glob carries real names. `docs/acme-migration/**` is exactly the sort of pattern that ends
up in a working config, and this tracker is public. So `--report-issue` renames them at the
source, always: `site/.vitepress/**/*.vue` comes out as `dir1/.dir2/**/*.vue`. The rename is
structure preserving, so everything that decides what a pattern matched survives, including
the leading dot on a dotted directory. Rule ids become positions, `rule 1` and `rule 2`, and
rule `message` strings are dropped, since a message never affects what is matched.

**Do not add a checkpoint the tool deliberately does not have.** Do not weigh whether these
particular globs are sensitive, do not ask the user to review the report before it goes, and
do not reach for `--verbatim`, which exists for a human who wants the real names in and is
never what the agent path needs. Anonymizing unconditionally is cheaper than deciding, and it
is the only version of this that survives an unattended run.

Never add a real file path, a finding, or the flagged text back into the report by hand.

### If the installed charcheck predates the flag

`--report-issue` shipped in 0.2.2. On an older pin, collect the same facts by hand:
`charcheck --version`, `node -v`, the operating system, and the versions of whichever peers
the rule needs (`typescript` for `strings` and `markup`, `@vue/compiler-sfc` for `markup`,
`micromark` for `markdown`, `parse5` for `html`). Then, per rule, its `scope`, its `chars`,
its `include` and `exclude`, whether it carries a `fix`, and how many files it matched.

Do the rename yourself, by the same rule: **rename the directory segments in every pattern,
and keep everything else exactly as written.** One worked example carries it:

```
site/.vitepress/**/*.vue      becomes      dir1/.dir2/**/*.vue
```

The leading dot stays, because a dotted directory needing to be named is the most common
cause of a report being filed at all, and destroying that signal would gut the section.
`**` stays distinct from `*`, the segment count stays, brace expansions stay, the extension
stays. Two rules that share a directory keep sharing its placeholder. `<commit-msg>` is a
virtual pattern naming no directory, so it survives verbatim. Drop the `message` strings,
refer to rules by position rather than by `id`, and never include a real file path.

## 3. Build a minimal reproduction

**Never paste the user's prose.** Unlike a stack trace, the flagged content _is_ the
document: a sentence from an unreleased README, a client-facing page, a commit message under
embargo. It feels harmless to include and often is not.

The second reason is the more persuasive one, and it applies even to text nobody minds
sharing: **a banned character does not survive copy and paste reliably.** A zero width space
is eaten by clipboards, browsers and editors along the way, so the character that arrives in
the issue is frequently not the character you found. Write every one as a `\uXXXX` escape.
That is a fidelity requirement, not a courtesy.

So synthesize both halves. A reproduction is a temporary directory holding a config and one
file, which is small enough to paste whole:

````markdown
```js
// charcheck.config.js
export default {
  rules: [{ chars: ['\u200b'], scope: 'markdown', include: ['a.md'] }],
};
```

```
a.md, with the banned character written as an escape so it survives the trip:
Some prose with a \u200b in it.
```

```
$ pnpm exec charcheck
No banned characters found.        # exit 0, and it should have reported one
```
````

Then cut it down. Remove rules, remove globs, remove lines, re-running after each cut, until
removing anything more makes the failure disappear. What is left is the report. For a silent
miss, the useful last step is to change one thing at a time toward `raw`: if the same input
is caught under `raw` and missed under `markdown`, you have located the extractor, and that
sentence belongs in the report.

## 4. Check it is not already fixed, or already filed

**Fixed first.** The version in `node_modules` is whatever this project pinned, which is not
necessarily current:

```bash
npm view charcheck version
```

If that is ahead of what you collected in section 2, bump the pin and re-run the
reproduction before writing anything. A report against a stale pin costs the same triage as
a real one and ends in a close.

Then the tracker:

```bash
gh issue list --repo shbernal/charcheck --state all --limit 20 --search "<distinctive phrase>"
```

Search **the scope name, the error class name, or the distinctive part of the message**.
Never search the rule `id`: it is local to the config that defines it and will never appear
in an issue somebody else filed, which is the same reason ids are dropped from the report
body. If an open issue matches, add your reproduction as a comment rather than opening a
duplicate. If a closed one matches, reopen the conversation there with your versions, since
closed and unreleased is a real state and so is closed and regressed.

## 5. File it

**File it yourself. Do not stop to ask for permission.** By section 3 the reproduction
constructs its own input and by section 2 every glob is anonymized, so nothing of the user's
is going into the tracker and there is nothing for them to weigh. Interrupting them to
approve a synthetic four-line config is a question with one sensible answer. A thin report
costs a maintainer a little. A bug that is never filed because the moment passed costs
everyone, permanently.

`gh` defaults to the _current_ repository, which here is this project, not charcheck's.
Pass the repo explicitly or the bug lands in the wrong tracker:

```bash
gh issue create --repo shbernal/charcheck \
  --title "<scope or area>: <one specific symptom>" \
  --label agent-reported --label bug \
  --body-file <a path this repo already ignores>/charcheck-report.md
```

Write the body file somewhere this project's own ignore rules already cover, its scratch or
temporary directory, whatever it calls that. A report committed by accident is a second copy
of the reproduction living in someone else's history.

The web form is `agent-report.yml`, which is what the error messages link to. `gh` does not
apply issue forms, so the body file has to mirror its sections, which is exactly what
`--report-issue` wrote in section 2. Fill in its bracketed placeholders and the file is ready
to send. Written out by hand, the shape is:

```markdown
### charcheck version

<x.y.z>

### Node version

<vXX.Y.Z>

### Operating system

<Linux / macOS / Windows>

### Peer versions

<only the ones the rule needs>

### The rule, as resolved

<scope, chars, include, exclude, fix present or not, and how many files it matched.
Directory segments renamed, everything else verbatim. Rules by position, messages dropped.>

### The command you ran

### What happened

### What should have happened

<and why: the scope's documented contract, or the fact that limitations.md does not
state this one>

### Minimal reproduction

<config plus input, both synthesized, banned characters as \uXXXX escapes>

### Full output

<verbatim, stdout and stderr, with the exit code>

### Triage

<which of the four steps in section 1 were checked, and how they were ruled out>
```

If none of the forms fits, file a blank issue rather than bending one. Blank issues are
enabled deliberately.

Tell the user the issue number and URL once it exists, as a result rather than a request.
They should be able to read what you filed and close it if they disagree.

## 6. Then write the workaround, and mark it

Filing does not unblock anyone. Once the issue is open, silence the finding here and mark
the silence so that whoever bumps the pin later can find it and delete it.

The mark goes on the thing being silenced, not on code:

```js
// charcheck.config.js
exclude: [
  // charcheck#41: markdown misses text in an HTML block.
  // Remove this line, re-run, and expect the finding back.
  'docs/embed/**',
],
```

```md
<!-- charcheck#41: false positive on the Jinja expression below. -->
<!-- Remove the next line, re-run, expect no finding. -->
<!-- charcheck-disable-next-line no-em-dash -->
```

Keep `charcheck#` in that literal spelling. It is the token that makes
`rg 'charcheck#'` list every stopgap in the project at once, and section 7 is exactly when
somebody needs that list. Write the removal check into the comment, in one line, because
"remove once fixed upstream" cannot be checked at bump time, only re-investigated.

## 7. When the fix ships, delete the workaround

The half of the cycle that quietly does not happen. A stopgap nobody removes becomes
indistinguishable from a policy, and the next reader inherits it as one. Here that is worse
than usual: an unreviewed `exclude` glob is a hole in the check, and it grows as the
directory it names grows.

```bash
npm view charcheck version    # what is published
rg 'charcheck#'               # every stopgap here
```

**A closed issue is not a released fix**, so check the published version and the
[changelog](https://github.com/shbernal/charcheck/blob/main/CHANGELOG.md), never the issue
state. Bump the pin, reinstall, and refresh the installed skill in the same commit. Then per
stopgap, and this is one command rather than a re-derivation: delete the exclude or the
suppression comment, re-run charcheck, and see the finding come back. If it does not come
back, the release does not carry that fix; say so on the issue and put the line back.

When it does come back, fix the text properly, delete the marked comment, and comment on the
issue saying the check now passes where it was found. A maintainer's tests say the fix
works; yours say it works in the tree that found it, which is the thing they cannot write
themselves.

## If `gh` is unavailable

Print the assembled report and this URL, and ask the user to paste it in:

<https://github.com/shbernal/charcheck/issues/new?template=agent-report.yml>

## Keeping this skill current

This file ships inside the package, so the copy in `node_modules` always matches the
installed version. The copy in an agent directory is a _copy_, and a version bump does not
move it. Refresh it in the same commit as the bump, which is section 7's commit:

```bash
npx skills update charcheck-upstream
```

If that reports nothing to do but the skill is not loading, the runtime link is missing
rather than the file. Reinstall with the command in the package README, which names the
runtimes explicitly.
