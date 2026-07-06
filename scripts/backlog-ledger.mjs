#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseDocument, stringify } from 'yaml'
import { ROOT } from './script-utils.mjs'

export const DEFAULT_LEDGER = path.join(ROOT, 'docs', 'backlog.yml')

const COMMANDS = new Set(['list', 'show', 'values', 'validate', 'remove', 'set-status', 'add'])
const ITEM_TYPES = ['issue', 'pull-request', 'downstream-need', 'fork-internal-proposal']
const REQUIRED_ITEM_FIELDS = [
	'id',
	'source',
	'type',
	'first_seen',
	'last_reviewed',
	'status',
	'priority',
	'target_area',
	'applies_to_current_project',
	'non_target_reasons',
	'summary',
	'current_project_notes',
	'evidence',
	'next_action',
]
const REQUIRED_EVIDENCE_FIELDS = [
	'kinds',
	'local_files',
	'schema_fixture',
	'validator_result',
	'powerpoint_result',
	'spec_refs',
]

function usage() {
	return `Usage: pnpm run backlog -- <command> [options]

Reads and edits docs/backlog.yml.

Commands:
  list                         Print compact ledger rows (use --json for full items)
  show <id...>                 Print full ledger items by id(s), or by filter when no id is given
  values <field>               Print unique values and counts for an item field
  validate                     Validate ledger structure and vocabulary use
  add                          Append a new ledger item from flags, then validate
  remove <id>                  Remove one ledger item by exact id
  set-status <id> <status>     Update status and last_reviewed

Options:
  --ledger <path>                         Ledger path (default: docs/backlog.yml)
  --json                                  Print machine-readable JSON (full items for list/show)
  --print-limit <n>                       Max list rows to print; 0 prints all (default: 50)
  --status <value[,value...]>             Filter list by status; sets status on add
  --type <issue|pull-request|pr|downstream-need|fork-internal-proposal>  Filter list by type; sets type on add
  --priority <value[,value...]>           Filter list by priority; sets priority on add
  --target-area <value[,value...]>        Filter list by target area; sets target_area on add
  --applies <yes|partial|no|unknown>      Filter list by applies_to_current_project; sets it on add
  --search <text>                         Filter list by id, source, summary, notes, or next_action
  --review-date <YYYY-MM-DD>              Date written by set-status (default: today)
  --id <id>                               Item id (add)
  --source <ref>                          Item source: owner/repo#N, a github URL, downstream[:path], or fork[:path] (add)
  --summary <text>                        One-line summary (add)
  --notes <text>                          current_project_notes body (add)
  --stopgap <path>                        downstream file the gap forces a workaround in (add, optional)
  --constructs <key[,key...]>             OOXML construct keys this entry gates (add, optional; vocabulary.constructs)
  --next-action <text>                    next_action (add; default: none)
  --first-seen <YYYY-MM-DD>               first_seen date (add; default: today)
  --dry-run                               Validate mutation without writing
  --help                                  Show this help

Examples:
  pnpm run backlog -- list
  pnpm run backlog -- list --status needs-repro --type issue
  pnpm run backlog -- list --type downstream-need
  pnpm run backlog -- show upstream-issue-1440 upstream-pr-1302
  pnpm run backlog -- show --status non-target --json
  pnpm run backlog -- values status
  pnpm run backlog -- validate
  pnpm run backlog -- add --id dn-text-direction --type downstream-need \\
    --source downstream:registry/components/quadrant-matrix.ts \\
    --summary "textDirection typed but not serialized" --priority p2
  pnpm run backlog -- remove upstream-issue-1440
  pnpm run backlog -- set-status upstream-issue-1440 implemented
`
}

function readOptionValue(args, index, name) {
	const raw = args[index]
	const equalsIndex = raw.indexOf('=')
	if (equalsIndex >= 0) return { value: raw.slice(equalsIndex + 1), consumed: 1 }
	const value = args[index + 1]
	if (!value || value.startsWith('--')) throw new Error('missing value for ' + name)
	return { value, consumed: 2 }
}

function parseNonNegativeInteger(name, value) {
	if (!/^\d+$/.test(value)) throw new Error(name + ' must be a non-negative integer')
	const parsed = Number(value)
	if (!Number.isSafeInteger(parsed)) throw new Error(name + ' is too large')
	return parsed
}

function parseCsv(value) {
	return value
		.split(',')
		.map((entry) => entry.trim())
		.filter(Boolean)
}

function addFilter(filters, name, value) {
	filters[name] ??= []
	filters[name].push(...parseCsv(value))
}

function normalizeType(value) {
	if (value === 'pr') return 'pull-request'
	return value
}

function isIsoDate(value) {
	return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(Date.parse(value))
}

function todayIsoDate() {
	return new Date().toISOString().slice(0, 10)
}

export function parseArgs(argv) {
	/**
	 * @type {{
	 *   command: string | null,
	 *   args: string[],
	 *   ledger: string,
	 *   json: boolean,
	 *   printLimit: number,
	 *   filters: Record<string, any>,
	 *   fields: Record<string, any>,
	 *   reviewDate: string,
	 *   dryRun: boolean,
	 *   help: boolean,
	 * }}
	 */
	const options = {
		command: null,
		args: [],
		ledger: DEFAULT_LEDGER,
		json: false,
		printLimit: 50,
		filters: {},
		fields: {},
		reviewDate: todayIsoDate(),
		dryRun: false,
		help: false,
	}
	const positional = []

	for (let i = 0; i < argv.length; ) {
		const arg = argv[i]
		if (arg === '--') {
			i += 1
		} else if (arg === '--help' || arg === '-h') {
			options.help = true
			i += 1
		} else if (arg === '--json') {
			options.json = true
			i += 1
		} else if (arg === '--dry-run') {
			options.dryRun = true
			i += 1
		} else if (arg.startsWith('--ledger')) {
			const result = readOptionValue(argv, i, '--ledger')
			options.ledger = path.resolve(ROOT, result.value)
			i += result.consumed
		} else if (arg.startsWith('--print-limit')) {
			const result = readOptionValue(argv, i, '--print-limit')
			options.printLimit = parseNonNegativeInteger('--print-limit', result.value)
			i += result.consumed
		} else if (arg.startsWith('--status')) {
			const result = readOptionValue(argv, i, '--status')
			addFilter(options.filters, 'status', result.value)
			i += result.consumed
		} else if (arg.startsWith('--type')) {
			const result = readOptionValue(argv, i, '--type')
			addFilter(options.filters, 'type', result.value)
			i += result.consumed
		} else if (arg.startsWith('--priority')) {
			const result = readOptionValue(argv, i, '--priority')
			addFilter(options.filters, 'priority', result.value)
			i += result.consumed
		} else if (arg.startsWith('--target-area')) {
			const result = readOptionValue(argv, i, '--target-area')
			addFilter(options.filters, 'targetArea', result.value)
			i += result.consumed
		} else if (arg.startsWith('--applies')) {
			const result = readOptionValue(argv, i, '--applies')
			addFilter(options.filters, 'applies', result.value)
			i += result.consumed
		} else if (arg.startsWith('--search')) {
			const result = readOptionValue(argv, i, '--search')
			options.filters.search = result.value
			i += result.consumed
		} else if (arg.startsWith('--review-date')) {
			const result = readOptionValue(argv, i, '--review-date')
			if (!isIsoDate(result.value)) throw new Error('--review-date must be YYYY-MM-DD')
			options.reviewDate = result.value
			i += result.consumed
		} else if (arg.startsWith('--first-seen')) {
			const result = readOptionValue(argv, i, '--first-seen')
			if (!isIsoDate(result.value)) throw new Error('--first-seen must be YYYY-MM-DD')
			options.fields.first_seen = result.value
			i += result.consumed
		} else if (arg.startsWith('--id')) {
			const result = readOptionValue(argv, i, '--id')
			options.fields.id = result.value
			i += result.consumed
		} else if (arg.startsWith('--source')) {
			const result = readOptionValue(argv, i, '--source')
			options.fields.source = result.value
			i += result.consumed
		} else if (arg.startsWith('--summary')) {
			const result = readOptionValue(argv, i, '--summary')
			options.fields.summary = result.value
			i += result.consumed
		} else if (arg.startsWith('--notes')) {
			const result = readOptionValue(argv, i, '--notes')
			options.fields.notes = result.value
			i += result.consumed
		} else if (arg.startsWith('--stopgap')) {
			const result = readOptionValue(argv, i, '--stopgap')
			options.fields.stopgap = result.value
			i += result.consumed
		} else if (arg.startsWith('--constructs')) {
			const result = readOptionValue(argv, i, '--constructs')
			options.fields.constructs = parseCsv(result.value)
			i += result.consumed
		} else if (arg.startsWith('--next-action')) {
			const result = readOptionValue(argv, i, '--next-action')
			options.fields.next_action = result.value
			i += result.consumed
		} else if (arg.startsWith('--')) {
			throw new Error('unknown option: ' + arg)
		} else {
			positional.push(arg)
			i += 1
		}
	}

	const command = positional.shift() || 'list'
	options.command = command
	options.args = positional
	if (!COMMANDS.has(command) && !options.help) throw new Error('unknown command: ' + command)
	options.filters.type = options.filters.type?.map(normalizeType)
	return options
}

function yamlErrorMessage(error) {
	const line = error.linePos?.[0]?.line
	const column = error.linePos?.[0]?.col
	const location = line ? ':' + line + (column ? ':' + column : '') : ''
	return 'YAML' + location + ': ' + error.message
}

export function parseLedgerText(text) {
	const doc = parseDocument(text, { prettyErrors: false })
	const parseErrors = doc.errors.map(yamlErrorMessage)
	if (parseErrors.length > 0) return { data: null, errors: parseErrors }
	try {
		return { data: doc.toJS(), errors: [] }
	} catch (error) {
		return { data: null, errors: ['YAML: ' + (error instanceof Error ? error.message : String(error))] }
	}
}

function asSet(values) {
	return new Set(Array.isArray(values) ? values : [])
}

function hasOwn(object, field) {
	return Object.prototype.hasOwnProperty.call(object, field)
}

function itemLabel(item, index) {
	return typeof item?.id === 'string' ? item.id : 'items[' + index + ']'
}

function validateListField(errors, item, index, field, vocabulary) {
	const label = itemLabel(item, index)
	const value = item[field]
	if (!Array.isArray(value)) {
		errors.push(label + ': ' + field + ' must be a list')
		return
	}
	for (const entry of value) {
		if (typeof entry !== 'string') {
			errors.push(label + ': ' + field + ' entries must be strings')
		} else if (vocabulary && !vocabulary.has(entry)) {
			errors.push(label + ': unknown ' + field + ' value: ' + entry)
		}
	}
}

function sourceKind(source) {
	if (typeof source !== 'string') return null
	if (/github\.com\/[^/]+\/[^/]+\/issues\/\d+/.test(source)) return 'issue'
	if (/github\.com\/[^/]+\/[^/]+\/pull\/\d+/.test(source)) return 'pull-request'
	return null
}

function sourceNumber(source) {
	if (typeof source !== 'string') return null
	return source.match(/(?:#|\/(?:issues|pull)\/)(\d+)\b/)?.[1] || null
}

function idNumber(id) {
	if (typeof id !== 'string') return null
	return id.match(/-(\d+)$/)?.[1] || null
}

function idKind(id) {
	if (typeof id !== 'string') return null
	if (/^upstream-issue-\d+$/.test(id)) return 'issue'
	if (/^upstream-pr-\d+$/.test(id)) return 'pull-request'
	return null
}

function isGithubRef(source) {
	if (typeof source !== 'string') return false
	if (/^[\w.-]+\/[\w.-]+#\d+$/.test(source)) return true
	return /github\.com\/[^/]+\/[^/]+\/(issues|pull)\/\d+/.test(source)
}

function isDownstreamSource(source) {
	return typeof source === 'string' && /^downstream(:\S.*)?$/.test(source)
}

function isForkSource(source) {
	return typeof source === 'string' && /^fork(:\S.*)?$/.test(source)
}

export function validateLedgerData(data) {
	const errors = []
	if (!data || typeof data !== 'object' || Array.isArray(data)) return ['ledger root must be a mapping']

	const vocabulary = data.vocabulary && typeof data.vocabulary === 'object' ? data.vocabulary : {}
	const statuses = asSet(vocabulary.statuses)
	const priorities = asSet(vocabulary.priorities)
	const targetAreas = asSet(vocabulary.target_areas)
	const appliesValues = asSet(vocabulary.applies_to_current_project)
	const nonTargetReasons = asSet(vocabulary.non_target_reasons)
	const evidenceKinds = asSet(vocabulary.evidence_kinds)
	const constructs = asSet(vocabulary.constructs)

	if (!Array.isArray(data.items)) {
		errors.push('items must be a list')
		return errors
	}

	const ids = new Map()
	data.items.forEach((item, index) => {
		const label = itemLabel(item, index)
		if (!item || typeof item !== 'object' || Array.isArray(item)) {
			errors.push(label + ': item must be a mapping')
			return
		}

		for (const field of REQUIRED_ITEM_FIELDS) {
			if (!hasOwn(item, field)) errors.push(label + ': missing required field: ' + field)
		}

		if (typeof item.id === 'string') {
			if (ids.has(item.id)) errors.push(item.id + ': duplicate id; first seen at item ' + ids.get(item.id))
			else ids.set(item.id, index)
		} else {
			errors.push(label + ': id must be a string')
		}

		if (typeof item.source !== 'string') {
			errors.push(label + ': source must be a string')
		} else if (item.type === 'downstream-need') {
			if (!isDownstreamSource(item.source)) {
				errors.push(label + ': downstream-need source must be "downstream" or "downstream:<path>"')
			}
		} else if (item.type === 'fork-internal-proposal') {
			if (!isForkSource(item.source)) {
				errors.push(label + ': fork-internal-proposal source must be "fork" or "fork:<path>"')
			}
		} else if (item.type === 'issue' || item.type === 'pull-request') {
			if (!isGithubRef(item.source)) {
				errors.push(
					label + ': ' + item.type + ' source must be a GitHub reference (owner/repo#N or a github.com issues/pull URL)'
				)
			}
		}
		if (!ITEM_TYPES.includes(item.type)) errors.push(label + ': type must be one of ' + ITEM_TYPES.join(', '))
		if (!statuses.has(item.status)) errors.push(label + ': unknown status: ' + item.status)
		if (!priorities.has(item.priority)) errors.push(label + ': unknown priority: ' + item.priority)
		if (!appliesValues.has(item.applies_to_current_project)) {
			errors.push(label + ': unknown applies_to_current_project: ' + item.applies_to_current_project)
		}
		if (!isIsoDate(item.first_seen)) errors.push(label + ': first_seen must be YYYY-MM-DD')
		if (!isIsoDate(item.last_reviewed)) errors.push(label + ': last_reviewed must be YYYY-MM-DD')

		validateListField(errors, item, index, 'target_area', targetAreas)
		validateListField(errors, item, index, 'non_target_reasons', nonTargetReasons)

		if (typeof item.summary !== 'string') errors.push(label + ': summary must be a string')
		if (typeof item.current_project_notes !== 'string') errors.push(label + ': current_project_notes must be a string')
		if (typeof item.next_action !== 'string') errors.push(label + ': next_action must be a string')

		// constructs is optional: validate membership only when present.
		if (hasOwn(item, 'constructs')) validateListField(errors, item, index, 'constructs', constructs)

		const kindFromId = idKind(item.id)
		if (kindFromId && item.type && kindFromId !== item.type) {
			errors.push(label + ': id kind does not match type')
		}
		const kindFromSource = sourceKind(item.source)
		if (kindFromSource && item.type && kindFromSource !== item.type) {
			errors.push(label + ': source URL kind does not match type')
		}
		const idNo = idNumber(item.id)
		const sourceNo = sourceNumber(item.source)
		if (idNo && sourceNo && idNo !== sourceNo) errors.push(label + ': source number does not match id number')

		if (!item.evidence || typeof item.evidence !== 'object' || Array.isArray(item.evidence)) {
			errors.push(label + ': evidence must be a mapping')
			return
		}
		for (const field of REQUIRED_EVIDENCE_FIELDS) {
			if (!hasOwn(item.evidence, field)) errors.push(label + ': missing required field: evidence.' + field)
		}
		validateListField(errors, item.evidence, index, 'kinds', evidenceKinds)
		if (!Array.isArray(item.evidence.local_files)) errors.push(label + ': evidence.local_files must be a list')
		if (!Array.isArray(item.evidence.spec_refs)) errors.push(label + ': evidence.spec_refs must be a list')
	})

	return errors
}

export function validateLedgerText(text) {
	const parsed = parseLedgerText(text)
	if (parsed.errors.length > 0) return { data: parsed.data, errors: parsed.errors }
	return { data: parsed.data, errors: validateLedgerData(parsed.data) }
}

export async function loadLedgerFile(ledgerPath) {
	const text = await fs.readFile(ledgerPath, 'utf8')
	const parsed = parseLedgerText(text)
	if (parsed.errors.length > 0) throw new Error(parsed.errors.join('\n'))
	return { text, data: parsed.data }
}

function filterMatchesAny(values, candidate) {
	return !values?.length || values.includes(candidate)
}

function itemSearchText(item) {
	return [
		item.id,
		item.source,
		item.status,
		item.priority,
		item.applies_to_current_project,
		item.summary,
		item.current_project_notes,
		item.stopgap,
		item.next_action,
	]
		.filter(Boolean)
		.join('\n')
		.toLowerCase()
}

export function filterItems(items, filters = {}) {
	const search = filters.search?.toLowerCase()
	return items.filter((item) => {
		if (!filterMatchesAny(filters.status, item.status)) return false
		if (!filterMatchesAny(filters.type, item.type)) return false
		if (!filterMatchesAny(filters.priority, item.priority)) return false
		if (!filterMatchesAny(filters.applies, item.applies_to_current_project)) return false
		if (filters.targetArea?.length) {
			const areas = Array.isArray(item.target_area) ? item.target_area : []
			if (!filters.targetArea.some((area) => areas.includes(area))) return false
		}
		if (search && !itemSearchText(item).includes(search)) return false
		return true
	})
}

function formatList(value) {
	return Array.isArray(value) ? value.join(', ') : ''
}

function formatEvidence(evidence) {
	if (!evidence || typeof evidence !== 'object') return []
	return REQUIRED_EVIDENCE_FIELDS.map(
		(field) => 'evidence.' + field + ': ' + printableValue(evidence[field] === undefined ? null : evidence[field])
	)
}

function itemFieldValue(item, field) {
	return field.split('.').reduce((value, part) => {
		if (value === undefined || value === null || typeof value !== 'object') return undefined
		return value[part]
	}, item)
}

function printableValue(value) {
	if (value === null) return 'null'
	if (typeof value === 'string') return value
	if (typeof value === 'number' || typeof value === 'boolean') return String(value)
	return JSON.stringify(value)
}

export function uniqueItemFieldValues(items, field) {
	const counts = new Map()
	for (const item of items) {
		const value = itemFieldValue(item, field)
		const values = Array.isArray(value) ? value : [value]
		for (const entry of values) {
			if (entry === undefined) continue
			const key = printableValue(entry)
			counts.set(key, (counts.get(key) || 0) + 1)
		}
	}
	return [...counts.entries()].map(([value, count]) => ({ value, count }))
}

function printValues(field, values, options, ledgerPath) {
	if (options.json) {
		return JSON.stringify(
			{
				ledger: path.relative(ROOT, ledgerPath),
				field,
				count: values.length,
				values,
			},
			null,
			2
		)
	}

	const lines = [field + ' values: ' + values.length]
	for (const entry of values) lines.push(entry.value + ' (' + entry.count + ')')
	return lines.join('\n')
}

function printList(items, options, ledgerPath) {
	if (options.json) {
		return JSON.stringify(
			{
				ledger: path.relative(ROOT, ledgerPath),
				count: items.length,
				items,
			},
			null,
			2
		)
	}

	const limit = options.printLimit === 0 ? items.length : Math.min(options.printLimit, items.length)
	const lines = ['Ledger items: ' + items.length]
	for (const item of items.slice(0, limit)) {
		lines.push(
			item.id + ' [' + item.status + '/' + item.priority + '/' + item.applies_to_current_project + '] ' + item.summary
		)
	}
	const remaining = items.length - limit
	if (remaining > 0) lines.push('... ' + remaining + ' more; use --print-limit 0 to print all rows.')
	lines.push(
		'These are compact rows. For full detail (rationale, evidence, notes) use --json, or ' +
			'`show <id…>` / `show --status <value>` for full text.'
	)
	return lines.join('\n')
}

function printItem(item, options) {
	if (options.json) return JSON.stringify(item, null, 2)
	const lines = [
		'id: ' + item.id,
		'source: ' + item.source,
		'type: ' + item.type,
		'first_seen: ' + item.first_seen,
		'last_reviewed: ' + item.last_reviewed,
		'status: ' + item.status,
		'priority: ' + item.priority,
		'target_area: ' + formatList(item.target_area),
		'applies_to_current_project: ' + item.applies_to_current_project,
		'non_target_reasons: ' + formatList(item.non_target_reasons),
		'summary: ' + item.summary,
		'current_project_notes: ' + item.current_project_notes,
	]
	if (item.stopgap) lines.push('stopgap: ' + item.stopgap)
	if (Array.isArray(item.constructs) && item.constructs.length > 0)
		lines.push('constructs: ' + item.constructs.join(', '))
	lines.push(...formatEvidence(item.evidence))
	lines.push('next_action: ' + item.next_action)
	return lines.join('\n')
}

function printItems(items, options, ledgerPath) {
	if (options.json) {
		return JSON.stringify(
			{
				ledger: path.relative(ROOT, ledgerPath),
				count: items.length,
				items,
			},
			null,
			2
		)
	}
	return items.map((item) => printItem(item, options)).join('\n\n')
}

function splitLines(text) {
	const lines = text.split(/(?<=\n)/)
	if (lines.at(-1) === '') lines.pop()
	return lines
}

function lineContent(line) {
	return line.replace(/\r?\n$/, '')
}

function lineEnding(line) {
	return line.match(/\r?\n$/)?.[0] || ''
}

function setLineContent(line, content) {
	return content + lineEnding(line)
}

function parseItemBlock(blockText) {
	const parsed = parseLedgerText('items:\n' + blockText)
	if (parsed.errors.length > 0) throw new Error(parsed.errors.join('\n'))
	return parsed.data.items?.[0] || null
}

export function findItemBlocks(text) {
	const lines = splitLines(text)
	const itemsLineIndex = lines.findIndex((line) => /^items:\s*(?:\[\])?\s*(?:#.*)?$/.test(lineContent(line)))
	if (itemsLineIndex < 0) throw new Error('items section not found')
	if (/^items:\s*\[\]\s*(?:#.*)?$/.test(lineContent(lines[itemsLineIndex]))) {
		return { lines, itemsLineIndex, blocks: [] }
	}

	const starts = []
	for (let i = itemsLineIndex + 1; i < lines.length; i += 1) {
		if (/^ {2}-\s+/.test(lines[i])) starts.push(i)
	}

	const blocks = starts.map((startIndex, index) => {
		const endIndex = starts[index + 1] ?? lines.length
		const text = lines.slice(startIndex, endIndex).join('')
		return {
			startIndex,
			endIndex,
			startLine: startIndex + 1,
			endLine: endIndex,
			text,
			item: parseItemBlock(text),
		}
	})
	return { lines, itemsLineIndex, blocks }
}

function findUniqueBlock(blocks, id) {
	const matches = blocks.filter((block) => block.item?.id === id)
	if (matches.length === 0) throw new Error('ledger item not found: ' + id)
	if (matches.length > 1) throw new Error('ledger item id is duplicated: ' + id)
	return matches[0]
}

function assertValidAfterMutation(text) {
	const validation = validateLedgerText(text)
	if (validation.errors.length > 0) {
		throw new Error('mutation would leave invalid ledger:\n' + validation.errors.join('\n'))
	}
}

export function removeLedgerItemText(text, id) {
	const { lines, itemsLineIndex, blocks } = findItemBlocks(text)
	const block = findUniqueBlock(blocks, id)
	if (blocks.length === 1) lines[itemsLineIndex] = setLineContent(lines[itemsLineIndex], 'items: []')
	lines.splice(block.startIndex, block.endIndex - block.startIndex)
	const updated = lines.join('')
	assertValidAfterMutation(updated)
	return updated
}

function replaceFieldLine(lines, block, field, value) {
	for (let i = block.startIndex; i < block.endIndex; i += 1) {
		if (new RegExp('^    ' + field + ':\\s*').test(lines[i])) {
			lines[i] = setLineContent(lines[i], '    ' + field + ': ' + value)
			return true
		}
	}
	return false
}

export function setLedgerItemStatusText(text, id, status, reviewDate = todayIsoDate()) {
	const validation = validateLedgerText(text)
	if (validation.errors.length > 0) throw new Error('ledger is invalid:\n' + validation.errors.join('\n'))
	const statuses = validation.data.vocabulary.statuses || []
	if (!statuses.includes(status)) throw new Error('unknown status: ' + status)
	if (!isIsoDate(reviewDate)) throw new Error('review date must be YYYY-MM-DD')

	const { lines, blocks } = findItemBlocks(text)
	const block = findUniqueBlock(blocks, id)
	if (!replaceFieldLine(lines, block, 'status', status)) throw new Error(id + ': status field not found')
	if (!replaceFieldLine(lines, block, 'last_reviewed', reviewDate))
		throw new Error(id + ': last_reviewed field not found')
	const updated = lines.join('')
	assertValidAfterMutation(updated)
	return updated
}

const ADD_DEFAULTS_BY_TYPE = {
	'downstream-need': { status: 'target', priority: 'p2', applies: 'yes' },
	'fork-internal-proposal': { status: 'deferred', priority: 'p3', applies: 'no' },
}

function buildItemSkeleton(fields, reviewDate) {
	if (!fields.id) throw new Error('add requires --id')
	if (!fields.source) throw new Error('add requires --source')
	if (!fields.type) throw new Error('add requires --type')
	const typeDefaults = ADD_DEFAULTS_BY_TYPE[fields.type] || {
		status: 'target',
		priority: 'none',
		applies: 'unknown',
	}
	return {
		id: fields.id,
		source: fields.source,
		type: fields.type,
		first_seen: fields.first_seen || reviewDate,
		last_reviewed: reviewDate,
		status: fields.status || typeDefaults.status,
		priority: fields.priority || typeDefaults.priority,
		target_area: fields.target_area || [],
		applies_to_current_project: fields.applies || typeDefaults.applies,
		non_target_reasons: [],
		summary: fields.summary || '',
		current_project_notes: fields.notes || '',
		...(fields.stopgap ? { stopgap: fields.stopgap } : {}),
		...(fields.constructs?.length ? { constructs: fields.constructs } : {}),
		evidence: {
			kinds: [],
			local_files: [],
			schema_fixture: null,
			validator_result: null,
			powerpoint_result: null,
			spec_refs: [],
		},
		next_action: fields.next_action || 'none',
	}
}

export function addLedgerItemText(text, fields, reviewDate = todayIsoDate()) {
	const validation = validateLedgerText(text)
	if (validation.errors.length > 0) throw new Error('ledger is invalid:\n' + validation.errors.join('\n'))
	const item = buildItemSkeleton(fields, reviewDate)
	const { blocks } = findItemBlocks(text)
	if (blocks.some((block) => block.item?.id === item.id)) throw new Error('ledger item id already exists: ' + item.id)

	const itemBlock = stringify({ items: [item] }, { lineWidth: 0 }).replace(/^items:\r?\n/, '')
	const base = text.endsWith('\n') ? text : text + '\n'
	const updated =
		blocks.length === 0
			? base.replace(/^items:[^\S\r\n]*(?:\[\])?[^\S\r\n]*$/m, 'items:') + itemBlock
			: base + itemBlock
	assertValidAfterMutation(updated)
	return updated
}

function validationReport(validation, options, ledgerPath) {
	if (options.json) {
		return JSON.stringify(
			{
				ledger: path.relative(ROOT, ledgerPath),
				valid: validation.errors.length === 0,
				errorCount: validation.errors.length,
				errors: validation.errors,
			},
			null,
			2
		)
	}
	if (validation.errors.length === 0) return 'Ledger validation passed'
	return [
		'Ledger validation failed: ' + validation.errors.length + ' error(s)',
		...validation.errors.map((error) => '- ' + error),
	].join('\n')
}

function defaultIo() {
	return {
		stdout: (message) => console.log(message),
		stderr: (message) => console.error(message),
	}
}

export async function runLedgerCommand(argv, io = defaultIo()) {
	const options = parseArgs(argv)
	if (options.help) {
		io.stdout(usage())
		return 0
	}

	if (options.command === 'validate') {
		const text = await fs.readFile(options.ledger, 'utf8')
		const validation = validateLedgerText(text)
		io.stdout(validationReport(validation, options, options.ledger))
		return validation.errors.length === 0 ? 0 : 1
	}

	if (options.command === 'remove') {
		const id = options.args[0]
		if (!id || options.args.length !== 1) throw new Error('remove requires exactly one item id')
		const text = await fs.readFile(options.ledger, 'utf8')
		const updated = removeLedgerItemText(text, id)
		if (!options.dryRun) await fs.writeFile(options.ledger, updated)
		io.stdout((options.dryRun ? 'Would remove ' : 'Removed ') + id)
		return 0
	}

	if (options.command === 'add') {
		const fields = {
			...options.fields,
			type: options.filters.type?.[0] ?? options.fields.type,
			status: options.filters.status?.[0],
			priority: options.filters.priority?.[0],
			applies: options.filters.applies?.[0],
			target_area: options.filters.targetArea,
		}
		const text = await fs.readFile(options.ledger, 'utf8')
		const updated = addLedgerItemText(text, fields, options.reviewDate)
		if (!options.dryRun) await fs.writeFile(options.ledger, updated)
		io.stdout((options.dryRun ? 'Would add ' : 'Added ') + options.fields.id)
		return 0
	}

	if (options.command === 'set-status') {
		const [id, status] = options.args
		if (!id || !status || options.args.length !== 2) throw new Error('set-status requires an item id and status')
		const text = await fs.readFile(options.ledger, 'utf8')
		const updated = setLedgerItemStatusText(text, id, status, options.reviewDate)
		if (!options.dryRun) await fs.writeFile(options.ledger, updated)
		io.stdout(
			(options.dryRun ? 'Would update ' : 'Updated ') +
				id +
				' status -> ' +
				status +
				' (last_reviewed ' +
				options.reviewDate +
				')'
		)
		return 0
	}

	const { data } = await loadLedgerFile(options.ledger)
	const items = data.items || []

	if (options.command === 'show') {
		const ids = options.args
		const hasFilters = Object.values(options.filters).some((value) => value !== undefined)
		let selected
		if (ids.length > 0) {
			selected = ids.map((id) => {
				const matches = items.filter((item) => item.id === id)
				if (matches.length === 0) throw new Error('ledger item not found: ' + id)
				if (matches.length > 1) throw new Error('ledger item id is duplicated: ' + id)
				return matches[0]
			})
		} else if (hasFilters) {
			selected = filterItems(items, options.filters)
		} else {
			throw new Error('show requires one or more item ids or a filter (e.g. --status non-target)')
		}
		io.stdout(printItems(selected, options, options.ledger))
		return 0
	}

	if (options.command === 'values') {
		const field = options.args[0]
		if (!field || options.args.length !== 1) throw new Error('values requires exactly one item field')
		const filteredItems = filterItems(items, options.filters)
		const values = uniqueItemFieldValues(filteredItems, field)
		if (filteredItems.length > 0 && values.length === 0)
			throw new Error('field not found on any matching item: ' + field)
		io.stdout(printValues(field, values, options, options.ledger))
		return 0
	}

	if (options.command === 'list') {
		if (options.args.length > 0) throw new Error('list does not accept positional arguments')
		io.stdout(printList(filterItems(items, options.filters), options, options.ledger))
		return 0
	}

	throw new Error('unknown command: ' + options.command)
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)
if (isMain) {
	runLedgerCommand(process.argv.slice(2))
		.then((code) => {
			process.exitCode = code
		})
		.catch((error) => {
			console.error(error.message)
			process.exitCode = 1
		})
}
