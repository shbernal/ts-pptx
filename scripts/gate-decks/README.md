# Gate decks

Decks that exist to be **diffed**, not looked at.

`scripts/byte-identity.mjs` proves an emitter refactor changed no emitted byte by
generating a corpus of decks and comparing every part against a frozen baseline. Its
original corpus was the showcase decks (`demos/showcases/`), and AGENTS.md is explicit
about the limit that imposes:

> The corpus is only what those decks emit, so before trusting a PASS, confirm the part
> you touched is in it — an emitter no showcase reaches is unproven, not proven unchanged.

The showcases are presentation decks. They reach three chart types (bar, doughnut, line)
because those are the three a business-review deck wants, and no arrangement of a
*plausible* deck is going to reach a wireframe surface plot, a four-axis volume-open-high-
low-close stock chart, and a date axis with an inverted scale. Pushing them there would
make them worse at the job they exist for.

So the gate gets its own corpus. A gate deck is under no obligation to look like anything:
it is a fixture matrix that happens to be shaped like a `.pptx`, because going through
`addChart` end to end is the only way to exercise the emitters as they are actually called.
It is never built by `pnpm demos:build`, never shipped, and never opened by a human except
when a diff needs reading.

## Rules

- **A gate deck may only grow.** Removing a case silently shrinks the proof. If a case is
  genuinely obsolete, say so in the commit message — the same standard as lowering a
  ratchet number.
- **Deterministic or normalized.** `byte-identity.mjs` reseeds `Math.random` per deck and
  normalizes four patterns (see `NORMALIZERS` in `scripts/pptx-parts.mjs`). Anything else
  that varies between two runs of the same code makes the gate flap, which is worse than
  not having it.
- **No assertions.** A gate deck asserts nothing and cannot fail. Its only job is to reach
  code; the baseline diff is what fails. Structural contracts belong in
  `test/regression/chart/`, which is a different question ("is this markup right?") from
  the one this corpus answers ("did this markup change?").

## Adding one

Charts needed a deck first, but the coverage gap is not specific to charts: `byte-identity`
gates every `src/gen/` refactor, and its showcase corpus reaches whatever two presentation
decks happen to reach. There are three decks now, and each arrived the same way, from a PASS
that turned out to be a statement about nothing:

- `chart-matrix.mjs` for the six chart emitters the showcases never construct.
- `shape-matrix.mjs` for the slide objects a business deck has no reason to carry.
- `html-table.mjs` for `tableToSlides`, which is a whole public entry point rather than one
  emitter: no showcase and no other gate deck called it, so a change moving every HTML-table
  border width by a third reported PASS.

The next refactor that runs into "the showcases never emit this" should add a deck rather
than proceed on an unproven PASS. Registering one is a `gateDeck` export like the one at the
foot of `chart-matrix.mjs`; the harness picks it up from `index.mjs`.

Note that adding a deck changes the part count, so it needs a fresh
`pnpm run byte-identity:baseline` — take the baseline *before* the refactor it is meant to
gate, never after.
