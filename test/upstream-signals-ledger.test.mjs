import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { describe, expect, test } from 'vitest'
import {
	filterItems,
	parseLedgerText,
	removeLedgerItemText,
	runLedgerCommand,
	setLedgerItemStatusText,
	uniqueItemFieldValues,
	validateLedgerText,
} from '../scripts/upstream-signals-ledger.mjs'

const fixture = `schema: 1
source_repo: gitbrent/PptxGenJS
vocabulary:
  statuses:
    - unreviewed
    - needs-repro
    - target-candidate
    - accepted
    - interesting-with-tweaks
    - non-target
    - watch
    - implemented
    - superseded
  priorities:
    - p0
    - p1
    - p2
    - p3
    - none
  target_areas:
    - powerpoint-repair
    - chart-ooxml
    - package-boundary
  applies_to_current_project:
    - yes
    - partial
    - no
    - unknown
  non_target_reasons:
    - commonjs
    - insufficient-evidence
  evidence_kinds:
    - local-source-inspection
    - minimal-repro
    - local-test
items:
  - id: upstream-issue-1
    source: "gitbrent/PptxGenJS#1"
    type: issue
    first_seen: 2026-06-07
    last_reviewed: 2026-06-07
    status: needs-repro
    priority: p1
    target_area:
      - chart-ooxml
    applies_to_current_project: unknown
    non_target_reasons: []
    upstream_summary: Chart labels need a local reproduction.
    current_project_notes: >
      Keep this entry until the chart XML path is tested locally.
    evidence:
      kinds:
        - local-source-inspection
      local_files:
        - src/gen-charts.ts
      schema_fixture: null
      validator_result: null
      powerpoint_result: null
      spec_refs: []
    next_action: create-minimal-repro
  - id: upstream-pr-2
    source: "gitbrent/PptxGenJS#2"
    type: pull-request
    first_seen: 2026-06-07
    last_reviewed: 2026-06-07
    status: non-target
    priority: none
    target_area:
      - package-boundary
    applies_to_current_project: no
    non_target_reasons:
      - commonjs
    upstream_summary: Reintroduces CommonJS package support.
    current_project_notes: >
      This conflicts with the current ESM-only package target.
    evidence:
      kinds: []
      local_files: []
      schema_fixture: null
      validator_result: null
      powerpoint_result: null
      spec_refs: []
    next_action: none
`

describe('upstream signals ledger tooling', () => {
	test('filters parsed items by status, type, priority, and target area', () => {
		const parsed = parseLedgerText(fixture)
		expect(parsed.errors).toEqual([])

		const matches = filterItems(parsed.data.items, {
			status: ['needs-repro'],
			type: ['issue'],
			priority: ['p1'],
			targetArea: ['chart-ooxml'],
		})
		expect(matches.map((item) => item.id)).toEqual(['upstream-issue-1'])
	})

	test('validates duplicate ids and invalid vocabulary values', () => {
		const invalid = fixture
			.replace('id: upstream-pr-2', 'id: upstream-issue-1')
			.replace('status: non-target', 'status: invalid-status')
		const validation = validateLedgerText(invalid)

		expect(validation.errors).toContain('upstream-issue-1: duplicate id; first seen at item 0')
		expect(validation.errors).toContain('upstream-issue-1: unknown status: invalid-status')
	})

	test('removes one exact item block without rewriting the remaining item', () => {
		const updated = removeLedgerItemText(fixture, 'upstream-issue-1')
		const validation = validateLedgerText(updated)

		expect(validation.errors).toEqual([])
		expect(validation.data.items.map((item) => item.id)).toEqual(['upstream-pr-2'])
		expect(updated).toContain('Reintroduces CommonJS package support.')
		expect(updated).not.toContain('Chart labels need a local reproduction.')
	})

	test('updates status and last_reviewed for one exact item', () => {
		const updated = setLedgerItemStatusText(fixture, 'upstream-issue-1', 'implemented', '2026-06-08')
		const validation = validateLedgerText(updated)

		expect(validation.errors).toEqual([])
		expect(validation.data.items[0].status).toBe('implemented')
		expect(validation.data.items[0].last_reviewed).toBe('2026-06-08')
		expect(validation.data.items[1].status).toBe('non-target')
	})

	test('lists unique field values and counts', () => {
		const parsed = parseLedgerText(fixture)
		const values = uniqueItemFieldValues(parsed.data.items, 'status')

		expect(values).toEqual([
			{ value: 'needs-repro', count: 1 },
			{ value: 'non-target', count: 1 },
		])
	})

	test('prints JSON list output for command-line consumers', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'upstream-signals.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['list', '--ledger', ledger, '--status', 'needs-repro', '--json'], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})
			const report = JSON.parse(stdout[0])

			expect(code).toBe(0)
			expect(report.count).toBe(1)
			expect(report.items[0].id).toBe('upstream-issue-1')
			expect(report.items[0].upstream_summary).toBe('Chart labels need a local reproduction.')
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('prints JSON values output for command-line consumers', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'upstream-signals.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['values', 'status', '--ledger', ledger, '--json'], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})
			const report = JSON.parse(stdout[0])

			expect(code).toBe(0)
			expect(report.field).toBe('status')
			expect(report.values).toEqual([
				{ value: 'needs-repro', count: 1 },
				{ value: 'non-target', count: 1 },
			])
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})
})
