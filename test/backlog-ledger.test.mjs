import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { describe, expect, test } from 'vitest'
import {
	addLedgerItemText,
	filterItems,
	parseLedgerText,
	removeLedgerItemText,
	runLedgerCommand,
	setLedgerItemStatusText,
	uniqueItemFieldValues,
	validateLedgerText,
} from '../scripts/backlog-ledger.mjs'

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
  constructs:
    - gradient-fill
    - vertical-text
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
    summary: Chart labels need a local reproduction.
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
    summary: Reintroduces CommonJS package support.
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

describe('backlog ledger tooling', () => {
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
			const ledger = path.join(tmpDir, 'backlog.yml')
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
			expect(report.items[0].summary).toBe('Chart labels need a local reproduction.')
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('list --json emits full items, not a compact projection', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['list', '--ledger', ledger, '--status', 'non-target', '--json'], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})
			const report = JSON.parse(stdout[0])

			expect(code).toBe(0)
			expect(report.count).toBe(1)
			const item = report.items[0]
			expect(item.non_target_reasons).toEqual(['commonjs'])
			expect(item.current_project_notes).toContain('ESM-only package target')
			expect(item.evidence).toMatchObject({ kinds: [] })
			expect(item.first_seen).toBe('2026-06-07')
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('show text output includes non_target_reasons, first_seen, and evidence', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['show', 'upstream-pr-2', '--ledger', ledger], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})

			expect(code).toBe(0)
			expect(stdout[0]).toContain('non_target_reasons: commonjs')
			expect(stdout[0]).toContain('first_seen: 2026-06-07')
			expect(stdout[0]).toContain('evidence.kinds: []')
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('show accepts multiple ids', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['show', 'upstream-issue-1', 'upstream-pr-2', '--ledger', ledger, '--json'], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})
			const report = JSON.parse(stdout[0])

			expect(code).toBe(0)
			expect(report.count).toBe(2)
			expect(report.items.map((item) => item.id)).toEqual(['upstream-issue-1', 'upstream-pr-2'])
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('show selects by filter when no id is given', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
			await fs.writeFile(ledger, fixture)
			const stdout = []
			const code = await runLedgerCommand(['show', '--status', 'non-target', '--ledger', ledger, '--json'], {
				stdout: (message) => stdout.push(message),
				stderr: () => {},
			})
			const report = JSON.parse(stdout[0])

			expect(code).toBe(0)
			expect(report.count).toBe(1)
			expect(report.items[0].id).toBe('upstream-pr-2')
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('show without id or filter is an error', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
			await fs.writeFile(ledger, fixture)
			await expect(
				runLedgerCommand(['show', '--ledger', ledger], { stdout: () => {}, stderr: () => {} })
			).rejects.toThrow(/one or more item ids or a filter/)
		} finally {
			await fs.rm(tmpDir, { recursive: true, force: true })
		}
	})

	test('accepts a downstream-need item with a downstream source', () => {
		const updated = addLedgerItemText(
			fixture,
			{
				id: 'dn-text-direction',
				source: 'downstream:registry/components/quadrant-matrix.ts',
				type: 'downstream-need',
				summary: 'textDirection typed but not serialized',
				stopgap: 'registry/components/quadrant-matrix.ts',
			},
			'2026-06-18'
		)
		const validation = validateLedgerText(updated)

		expect(validation.errors).toEqual([])
		const added = validation.data.items.find((item) => item.id === 'dn-text-direction')
		expect(added.type).toBe('downstream-need')
		expect(added.status).toBe('target-candidate')
		expect(added.priority).toBe('p2')
		expect(added.stopgap).toBe('registry/components/quadrant-matrix.ts')
	})

	test('rejects an inconsistent source for the item type', () => {
		const issueWithDownstreamSource = fixture.replace(
			'source: "gitbrent/PptxGenJS#1"\n    type: issue',
			'source: "downstream"\n    type: issue'
		)
		expect(validateLedgerText(issueWithDownstreamSource).errors).toContain(
			'upstream-issue-1: issue source must be a GitHub reference (owner/repo#N or a github.com issues/pull URL)'
		)

		// addLedgerItemText validates after mutation, so a non-downstream
		// source on a downstream-need throws before returning.
		expect(() =>
			addLedgerItemText(
				fixture,
				{ id: 'dn-bad', source: 'gitbrent/PptxGenJS#9', type: 'downstream-need', summary: 'x' },
				'2026-06-18'
			)
		).toThrow(/downstream/)
	})

	test('accepts an optional constructs list and rejects unknown construct keys', () => {
		// Absent constructs is valid (the field is optional).
		expect(validateLedgerText(fixture).errors).toEqual([])

		// A known construct key validates clean...
		const tagged = addLedgerItemText(
			fixture,
			{
				id: 'dn-vert-text',
				source: 'downstream:registry/components/quadrant-matrix.ts',
				type: 'downstream-need',
				summary: 'vertical text gate',
				constructs: ['vertical-text'],
			},
			'2026-06-18'
		)
		const taggedValidation = validateLedgerText(tagged)
		expect(taggedValidation.errors).toEqual([])
		expect(taggedValidation.data.items.find((item) => item.id === 'dn-vert-text').constructs).toEqual(['vertical-text'])

		// ...an unmapped key is rejected against vocabulary.constructs. Target the
		// item's 6-space-indented entry, not the 4-space vocabulary declaration.
		const bogus = tagged.replace('      - vertical-text', '      - not-a-construct')
		expect(validateLedgerText(bogus).errors).toContain('dn-vert-text: unknown constructs value: not-a-construct')
	})

	test('refuses to add a duplicate id', () => {
		expect(() =>
			addLedgerItemText(
				fixture,
				{ id: 'upstream-issue-1', source: 'downstream', type: 'downstream-need', summary: 'dup' },
				'2026-06-18'
			)
		).toThrow(/already exists/)
	})

	test('prints JSON values output for command-line consumers', async () => {
		const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-ledger-'))
		try {
			const ledger = path.join(tmpDir, 'backlog.yml')
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
