import { describe, test } from 'vitest'
import { assert, assertEqual } from '../helpers.js'

// Guards against export-condition drift between the package's runtime entries.
//
// `package.json` resolves `.` per-condition (node → dist/node.js, browser →
// dist/browser.js, default → dist/index.js) while types always come from
// index.d.ts. When an entry used `export type *` from core-interfaces instead of
// `export *`, the value factory helpers `textRun`/`textRuns` were dropped from
// that entry's runtime surface — TypeScript stayed green (it reads index.d.ts)
// but Node threw "does not provide an export named 'textRun'" at import time.
// This locks every published runtime entry to the same value surface so the
// entries cannot silently diverge again.

// Every entry listed in package.json `exports` that ships a runtime module.
const ENTRIES = ['index', 'node', 'browser', 'standalone', 'core']

describe('Entry export surface', () => {
	for (const entry of ENTRIES) {
		test(`${entry}.js exports working textRun/textRuns value helpers`, async () => {
			// `.js` stays in the static part so vite's dynamic-import-vars is satisfied.
			const mod = await import(`../../dist/${entry}.js`)

			assertEqual(typeof mod.textRun, 'function', `${entry} must export textRun as a function`)
			assertEqual(typeof mod.textRuns, 'function', `${entry} must export textRuns as a function`)

			// Behavior, not just presence: the helpers must produce the documented shape.
			const run = mod.textRun('hi', { bold: true })
			assertEqual(run.text, 'hi', `${entry} textRun should carry text`)
			assertEqual(run.options?.bold, true, `${entry} textRun should carry options`)
			assert(mod.textRun('plain').options === undefined, `${entry} textRun should omit options when not given`)

			const runs = [mod.textRun('a'), mod.textRun('b')]
			assertEqual(mod.textRuns(runs), runs, `${entry} textRuns should pass the array through`)
		})
	}
})
