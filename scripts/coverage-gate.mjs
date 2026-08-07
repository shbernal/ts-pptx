#!/usr/bin/env node
/**
 * The merged coverage gate — and the point of slack, enforced instead of written down.
 *
 * Reads `coverage/merged/coverage-summary.json` (written by scripts/coverage-merge.mjs)
 * and checks it against `scripts/coverage-gates.json`. Two different failures, because
 * they call for two different fixes:
 *
 *   - **below the notch.** Coverage regressed past the gate. Cover the code, or explain
 *     what changed; never move the notch down to meet it.
 *   - **inside the point of slack.** The number is still above its notch but by less than
 *     a point. Nothing is red yet in the ordinary sense, and that is exactly why this
 *     exists.
 *
 * The second check is the one with a history. Every threshold in this repo is set a notch
 * below the measured number with at least a full point of headroom, and that rule lived
 * only in prose — in the comments on `vitest.config.ts`, and in a step's acceptance
 * criteria. Prose does not fail a build. So when dropping the `dist/browser.js` exclusion
 * put 13 unreachable-from-Node functions into the denominator, `functions` fell 98.33 ->
 * 97.35 and `lines` fell 96.00 -> 95.67: both still over their notches, both now inside
 * the point of slack, and the acceptance criterion ("the exclusion is gone and thresholds
 * still pass") was satisfied by a state the doctrine forbids. One of the two was noticed;
 * the other sat in a comment that still quoted the old number.
 *
 * That is not a mistake to be more careful about next time. It is a rule that was never
 * given a way to fail, and this is the way.
 *
 *   node scripts/coverage-gate.mjs
 *
 * ## Why this gates the merged report and not the Node one
 *
 * `vitest.config.ts` keeps its own thresholds and they are not touched here. They guard
 * the Node suite alone — the fast local loop, `pnpm run test:coverage`, no browser
 * required — and they are a regression floor: they may not be lowered, and they will
 * fail if the Node suite goes backwards.
 *
 * They are deliberately *not* where the point-of-slack rule applies, because the Node
 * report's denominator includes code the Node lane structurally cannot execute. `fetch`,
 * `FileReader` and a canvas are not missing tests; they are missing a runtime. Demanding a
 * point of slack there would demand covering the adapter from Node, which is impossible,
 * and the only ways to satisfy it would be to lower the notch or to re-hide the file —
 * the two things the doctrine exists to prevent.
 *
 * The merged report has a collector for every line it counts. That makes it the number
 * the doctrine can honestly be held against, and this is where it is held.
 */

import fs from 'node:fs'
import path from 'node:path'
import { ROOT } from './script-utils.mjs'

const SUMMARY = path.join(ROOT, 'coverage', 'merged', 'coverage-summary.json')
const GATES = path.join(ROOT, 'scripts', 'coverage-gates.json')

const AXES = ['statements', 'branches', 'functions', 'lines']

if (!fs.existsSync(SUMMARY)) {
	console.error('coverage-gate: no merged report. Run: pnpm run coverage:merge')
	process.exit(1)
}

const { total } = JSON.parse(fs.readFileSync(SUMMARY, 'utf8'))
const { minimumSlack, thresholds } = JSON.parse(fs.readFileSync(GATES, 'utf8'))

const failures = []
const available = []

console.log('')
console.log('  merged coverage gate')
console.log('')
console.log('  axis         measured      gate      slack')

for (const axis of AXES) {
	const measured = total[axis].pct
	const gate = thresholds[axis]
	const slack = measured - gate

	console.log(
		'  ' +
			axis.padEnd(12) +
			measured.toFixed(2).padStart(8) +
			String(gate).padStart(10) +
			slack.toFixed(2).padStart(11) +
			(measured < gate ? '  <- below the gate' : slack < minimumSlack ? '  <- inside the point of slack' : '')
	)

	if (measured < gate) {
		failures.push(
			`${axis}: ${measured.toFixed(2)} is below its gate of ${gate}. ` +
				`Cover the ${total[axis].total - total[axis].covered} uncovered ${axis}, or say what changed — ` +
				`the gate does not move down.`
		)
	} else if (slack < minimumSlack) {
		failures.push(
			`${axis}: ${measured.toFixed(2)} clears its gate of ${gate} by only ${slack.toFixed(2)}, ` +
				`under the ${minimumSlack.toFixed(2)} point every notch here is required to leave. ` +
				`Coverage has to come back up; the notch stays where it is.`
		)
	}

	// The highest notch that would still leave the required slack. Printed rather than
	// applied — ratcheting is a decision with a commit message, not a side effect.
	const ratchet = Math.floor(measured - minimumSlack)
	if (ratchet > gate) available.push(`  ${axis}: ${gate} -> ${ratchet} (measured ${measured.toFixed(2)})`)
}

console.log('')

if (failures.length) {
	for (const failure of failures) console.error('  FAIL  ' + failure)
	console.error('')
	process.exit(1)
}

if (available.length) {
	console.log('  a ratchet is available (scripts/coverage-gates.json):')
	for (const line of available) console.log(line)
	console.log('')
}
