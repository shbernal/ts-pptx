// Per-test reset of the process globals the library and the suite own.
//
// This file is what makes `isolate: false` (vitest.config.ts) safe rather than merely
// fast. With isolation on, every test *file* got a fresh module registry, so
// `src/diagnostics.ts`'s module-level `handler` was reborn at the default before each
// file and leaking it could only ever hurt the file that leaked it. With isolation off,
// one worker's registry is shared across every file it runs, and a handler left installed
// silently swallows — or throws on — diagnostics belonging to a test in a different
// directory, minutes later. That failure reads as a bug in the victim, which is the most
// expensive kind of failure this suite could grow.
//
// So the invariant is restored here rather than trusted: `captureDiagnostics` already
// restores in a `finally`, and 41 of the 81 direct `setDiagnosticHandler` calls in tests
// pass `null` themselves, but "every author remembers" is a convention, and a convention
// is exactly what isolation was previously covering for.
//
// `null` restores the console default (see `setDiagnosticHandler`'s contract), which is
// the state a freshly imported module would have been in.
//
// `resetDiagnosticState()` restores the *other* library global with the same lifetime: the
// `warnOnce` dedupe set. It is the subtler of the two, because leaking it makes a test go
// green rather than red — a `warnOnce` condition another file already tripped is silent
// here, so an assertion that the warning fires fails, while an assertion that it does not
// passes for the wrong reason. Confirmed live before this landed: two files each asserting
// `margin/legacy-points` fires, run in one worker (`VITEST_MAX_WORKERS=1`), and whichever
// went second failed. It had never bitten only because the pool is usually wide enough to
// put two such files in different workers, which is luck, not isolation.
//
// This does NOT reset `globalThis.fetch`, which test/regression/api/node-runtime-fetch.test.js
// swaps: that file installs and restores it in its own hooks, and a blanket reset here
// would have to capture the real `fetch` at import time and could just as easily clobber a
// deliberate stub mid-test. Keep such swaps local and paired; this file is for state the
// *library* owns.

import { afterEach } from 'vitest'
import { resetDiagnosticState, setDiagnosticHandler } from '../dist/node.js'

afterEach(() => {
	setDiagnosticHandler(null)
	resetDiagnosticState()
})
