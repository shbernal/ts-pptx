// Where the read-side fixture decks live, and nothing else.
//
// This is one line of `corpus.js`, split out for one reason: `corpus.js` enumerates the corpus
// under a top-level `await` and throws below `MIN_CORPUS`, so importing it for a path drags in
// that whole claim. `font-oracle.js` wants only `FIXTURES` (to locate the metrics sidecar), and
// two write-side suites under `test/regression/text/` import `font-oracle.js` for
// `resolveGenuineFontFile`, which has no fixture dependency at all -- so a short read corpus
// could fail the collection of a regression test that asserts nothing about it.
//
// `corpus.js` re-exports this, so it remains the one place a test reads the corpus from and the
// directory is still defined once.
//
// Not a test file (no `.test.` in the name) -- vitest's default glob skips it.

import path from 'node:path'
import { fileURLToPath } from 'node:url'

export const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures')
