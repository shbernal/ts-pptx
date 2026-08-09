---
name: ts-pptx-upstream
description: Report a ts-pptx (@shbernal/ts-pptx) bug, gap, or wrong output to its GitHub tracker from a project that depends on it. Use when ts-pptx throws an InternalError or an error whose message points at the issue tracker, when it cannot read a .pptx that opens cleanly in PowerPoint, when it writes a deck PowerPoint repairs or renders wrong, when a construct does not survive a round trip, when its types block correct code, or whenever you are about to write a workaround for ts-pptx behaving incorrectly. Filing the bug is the fix; the workaround is the stopgap.
---

# Reporting a ts-pptx problem upstream

You are in a project that *uses* `@shbernal/ts-pptx`, not the project that builds it.
This skill is how a defect you hit here becomes a permanent regression test there.

The library's maintainers turn reports with a minimal reproduction into fixtures and
regression cases, which is the only mechanism that guarantees a bug is never
reintroduced. A report without a reproduction is a wish; a report with one is a fix.

**The single most valuable thing you can do is reduce the failure to a script that
builds its own deck.** Everything below is in service of that.

## 1. Decide whether it is actually ours

Every failure ts-pptx raises is a `TsPptxError` in one of five classes, and the class
already answers *whose problem is this*. Catch it and read `err.name` and `err.code`
— `code` is a stable `area/condition` string and is API; the message is not, so never
branch on its wording.

| class                     | whose bug        | report it?                                                          |
| ------------------------- | ---------------- | ------------------------------------------------------------------- |
| `InternalError`           | **ts-pptx**      | Always. The library says so itself in the message.                   |
| `PackageReadError`        | usually the file | Only if the file opens **cleanly in PowerPoint**. Then it's our gap. |
| `MediaError`              | usually the asset| Only if the image/font/AV loads fine elsewhere.                      |
| `UnsupportedFeatureError` | nobody's, yet    | If PowerPoint can express it, this is a feature request worth filing.|
| `InvalidOptionError`      | usually you      | Only if the deck it refused is one PowerPoint can express.           |

**The supported bar is "the output opens cleanly in Microsoft PowerPoint."** That is
the project's own stated standard, and it is the test to apply in both directions: a
file PowerPoint opens but ts-pptx rejects is our gap, and a file ts-pptx writes that
PowerPoint repairs is our bug.

Not every defect throws. These are ours too, and are worth reporting:

- Output PowerPoint opens with a **repair prompt** (`0x80070570` and friends), or
  renders differently than intended.
- A **round trip that loses a construct**: read a deck, write it back, something is
  gone or changed. If the loss surfaced as a conversion fidelity note, the note has
  already classified it for you — see below.
- The **read side cannot see what the write side authors** — a property ts-pptx emits
  but has no accessor to read back. This is the strongest case a gap can make, and it
  has its own form (see step 5).
- A `Diagnostic` that fires when it should not, does not fire when it should, or
  carries the wrong `code`.
- Types that make correct code fail to compile, or admit code that throws at runtime.
- Documented behaviour that does not match observed behaviour.

### If you are holding a fidelity note, it already told you

Anything that converts a deck through `@shbernal/ts-pptx/script` gets `FidelityNote`s
back, and `cause` is the library's own verdict on whose gap the loss is. Read it
before deciding whether to file — it answers the same question the error table above
answers for the throwing cases:

| `cause`       | what it means                                              | report it?                                      |
| ------------- | ---------------------------------------------------------- | ----------------------------------------------- |
| `unread`      | nothing on the read side can see it — a missing reader      | **Yes**, and `api-gap.yml` is its form.          |
| `unwritable`  | it is read fine, but no write option can author it back     | **Yes** — a missing option is a gap, not a limit.|
| `unsupported` | OOXML, or the tier you chose, cannot express it at all      | **No**. No amount of converter work closes it.   |

So a `dropped`/`unread` note is a report waiting to be written, and an `unsupported`
one is the format's answer rather than ours. Quote `construct` (`line.width`,
`table.rowAuto`) verbatim: it is a stable dotted key, the form asks for it, and it is
the fastest thing a maintainer can search on.

One more shape worth filing: a note that fires for a construct that **did** survive.
A declared loss that does not happen is a stale note, and it is a defect in the same
way a missing one is — it teaches every reader to discount the notes that are true.

Two things are **out of active maintenance scope** and will usually be closed — check
before spending effort on a reproduction:

- Live-DOM / browser-layout features, i.e. anything whose answer comes from a
  *rendered* page (real `offsetWidth` after layout, the resolved cascade). Converting
  an HTML `<table>` is *not* in this category.
- Third-party office-suite interop quirks that appear only after the file has been
  round-tripped through another application, when the package ts-pptx wrote is itself
  valid OOXML.

Issues and pull requests in those two areas are still welcome; just say so in the
report so nobody triages it as a regression.

Also: ts-pptx is **ESM, Node-first**. Reports about `require()`, a CJS build, or an
IIFE/global browser bundle will be closed — those are deliberately not supported.

## 2. Collect the facts

```bash
node -p "require('@shbernal/ts-pptx/package.json').version"   # exact installed version
node -v                                                       # runtime (>=24 required)
```

Capture, verbatim: `err.name`, `err.code`, the full message, `err.detail` if present,
and the stack. Do not paraphrase the message — the throw site is often identifiable
from its exact text, and `detail` is the structured context the site chose to attach.

If the failure is about how the file *renders* or whether PowerPoint repairs it, note
the PowerPoint version and platform too. For an API-shape report it is meaningless;
for a render report it is load-bearing.

## 3. Build a minimal reproduction

**Never attach or paste the user's deck.** A presentation in a real project carries
client names, unreleased strategy, pricing, and internal logos. Treat every `.pptx`
in this repo as confidential unless the user tells you otherwise, and never upload one
to a public tracker — including "just a screenshot of the slide". This is the same rule
the project already applies to its own filings: describe a downstream consumer's need
anonymously, never its name, its deck or client names, or a path from its tree.

Instead, write a self-contained script that *constructs* its input and fails:

```ts
import TsPptx from '@shbernal/ts-pptx'

// build the smallest deck that shows the problem
const pptx = new TsPptx()
const slide = pptx.addSlide()
slide.addText('x', { x: 1, y: 1, w: 8, h: 1 /* the option that triggers it */ })

const bytes = await pptx.write({ outputType: 'uint8array' })
// ...then whatever exposes the defect: read it back, inspect it, open it.
```

For a read-side or round-trip defect, generate the input with the write side first, so
the script still owns everything it touches:

```ts
import { Presentation } from '@shbernal/ts-pptx/read'

const pres = await Presentation.load(bytes) // bytes built above, not the user's file
// ...the accessor that returns the wrong thing, or throws
```

Then cut it down: remove slides, shapes, options, and styling one at a time, re-running
after each cut, until removing anything more makes the failure disappear. What is left
is the report.

If the failure only reproduces with a *specific deck* you cannot share, do not ask for
permission to share it — file without the file. Say exactly that in the report and
describe the structural feature you believe is responsible (a grouped shape, a chart
part, a custom table style, a particular namespace prefix, a layout inherited from a
master). A maintainer can usually author a fixture from that description in PowerPoint;
they can never unsee an attachment.

If you need a file attached, build one: reproduce the *structure* you suspect in a deck
you generate, with invented values. A synthesized file is always safe to attach. A
redacted one is not yours to judge — redaction fails quietly, and a public tracker is
permanent.

## 4. Check it is not already fixed, or already filed

**Fixed first.** The version in `node_modules` is whatever the project pinned, which is
not necessarily the current one:

```bash
npm view @shbernal/ts-pptx version   # latest published
```

If that is ahead of the version you collected in step 2, bump the pin and re-run the
reproduction from step 3 before writing anything. A report against a stale version costs
a maintainer the same triage as a real one and ends in a close; the bump costs you one
command, and you were going to need it anyway once the fix shipped.

Then the tracker:

```bash
gh issue list --repo shbernal/ts-pptx --state all --limit 20 --search "<distinctive phrase>"
```

Search the error `code` (`table/invalid-border`, `oxml/node-has-no-document`), the
fidelity note's `construct` key, or the distinctive part of the message — not your
description of it. Codes and construct keys are stable and searchable in a way prose is
not. If an open issue matches, add your reproduction as a comment instead of opening a
duplicate; if a closed one matches, reopen the conversation there with your version and
Node version — closed and unreleased is a real state, and so is closed and regressed.

## 5. Pick the form, then file it

**File it yourself. Do not stop to ask for permission.** A report whose reproduction
builds its own deck carries nothing of the user's into the tracker, so there is nothing
for them to weigh; interrupting them to approve a synthetic script is a question with
only one sensible answer. A duplicate or a thin report is a small cost to a maintainer.
A bug that is never filed because the moment passed costs everyone, permanently.

This rests entirely on step 3 holding: the reproduction constructs its own input, and no
file from this project is attached. When you cannot manage a self-contained reproduction,
the answer is a thinner report — the prose description from step 3 — not a question and
not an attachment.

Three forms; pick by what you are reporting:

| form               | for                                                                       |
| ------------------ | ------------------------------------------------------------------------- |
| `agent-report.yml` | you hit this while using the library in another project — start here       |
| `bug.yml`          | wrong output, a repair prompt, a regression, a fidelity limit              |
| `api-gap.yml`      | the OOXML carries it but no accessor reaches it                            |

If none fits, file a blank issue rather than bending one of them — blank issues are
enabled deliberately.

`gh` defaults to the *current* repository — which here is the consumer's, not ts-pptx's.
Always pass `--repo shbernal/ts-pptx` explicitly, or you will file the bug into the wrong
tracker.

```bash
gh issue create --repo shbernal/ts-pptx \
  --title "<InternalError|reads|writes|round-trip|types>: <one specific symptom>" \
  --label agent-reported \
  --body-file <a path your repo ignores>/ts-pptx-report.md
```

Write the body somewhere the consumer's own `.gitignore` already covers — its scratch or
temp directory, whatever that repo calls it. A report file committed by accident is a
second copy of the reproduction living in someone else's history.

The web form (`agent-report.yml`) is what the error message links to; `gh` does not apply
issue forms, so mirror its sections in the body file so both routes land the same shape:

````markdown
### ts-pptx version
<x.y.z>

### Node.js version
<vXX.Y.Z>

### Operating system
<Windows 11 x64 / macOS 15 arm64 / Ubuntu 24.04 x64>

### PowerPoint version
<only if the report depends on how PowerPoint renders or repairs the file>

### Error class and code
<InternalError / oxml/node-has-no-document>   (or: no error thrown — wrong output)

### What happened
<observed>

### What should have happened
<expected, and why you believe that — an ECMA-376 clause, what PowerPoint itself does
with the same input, or what the docs promise>

### Minimal reproduction
```js
<the script from step 3>
```

### Error output
```
<verbatim message, detail, and stack>
```

### Attached file
<none / synthesized — describe how it was generated>
````

Tell the user the issue number and URL once it exists — after the fact, as a result, not
as a request. They should be able to read what you filed and close it if they disagree.

## 6. Then, and only then, write the workaround

Filing does not unblock the user. Once the issue is open, implement the workaround in
this project and mark it, so that whoever bumps the pin later can find it and decide.

An issue number alone is not enough of a mark. `remove once fixed upstream` does not say
what *fixed* looks like, so at bump time it cannot be checked — it can only be
re-investigated, which means re-reading the issue and re-deriving the reproduction
someone already wrote. Write the comment so that verifying the fix is running one line:

```ts
// Workaround for ts-pptx#<N> — https://github.com/shbernal/ts-pptx/issues/<N>
//
// <what the library does instead, as an observable: the XML it emits, the value the
//  accessor returns, the option it ignores.>
//
// Remove when <the exact check — an accessor returning the right value, a written part
// containing the right element> holds, and write <the code this becomes> instead.
```

Keep `ts-pptx#` in that literal spelling wherever you mark one. It is the token that
makes `rg 'ts-pptx#'` list every stopgap in the project at once, and step 7 is exactly
the moment someone needs that list.

## 7. When the fix ships, delete the workaround

The other half of the cycle, and the half that quietly does not happen: a stopgap nobody
removes becomes indistinguishable from a design decision, and the next reader inherits it
as one. Do this in one unit of work, on the release that carries the fix.

```bash
npm view @shbernal/ts-pptx version                          # what is out
gh issue list --repo shbernal/ts-pptx --state closed --limit 30
rg 'ts-pptx#'                                               # every stopgap here
```

**A closed issue is not a released fix.** A fix can sit merged and unreleased for weeks,
so the published version is what to check, never the issue state — the repository's
[`CHANGELOG.md`](https://github.com/shbernal/ts-pptx/blob/master/CHANGELOG.md) and the
GitHub release notes name the issue numbers each version closes. Bump the pin, reinstall,
and refresh the installed skill in the same commit. Then, per stopgap:

1. **Run the check its comment names.** If it does not hold, the release does not carry
   that fix; say so on the issue rather than deleting anything.
2. **Delete the workaround** and write what the comment told you to write.
3. **Add the test the workaround was standing in for**, if the project has a layer that
   can hold it. The stopgap was the only thing keeping the defect from being visible; a
   deletion with nothing in its place means the same regression can return unnoticed.
4. **Replace the comment rather than only deleting it** where the code still looks odd
   without it. One sentence on why the shape it has is the shape it has — *this used to
   be X because of Y, which release Z fixed* — is the difference between a reader
   trusting the line and re-litigating it.
5. **Close the loop upstream**: comment on the issue with the check that now passes, in
   the same self-contained form as the original reproduction. A maintainer's own tests
   say the fix works; a consumer's say it works *where it was found*, which is the thing
   they cannot write themselves.

## If `gh` is unavailable

Print the assembled report and this URL, and ask the user to paste it in:

<https://github.com/shbernal/ts-pptx/issues/new?template=agent-report.yml>

## Keeping this skill current

This file ships inside the package, so the copy in `node_modules` always matches the
installed version — but the copy in an agent directory is a *copy*, and a version bump
does not move it. Refresh it in the same commit as the bump, which is step 7's commit:

```bash
npx skills update ts-pptx-upstream
```

If that reports nothing to do but the skill is not loading, the runtime link is missing
rather than the file — reinstall it with the command in the package README, which names
the runtimes explicitly. `skills experimental_install` and `experimental_sync` restore
the file from the lock and create no runtime links, so neither is the repair.
