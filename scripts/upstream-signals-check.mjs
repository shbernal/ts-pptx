#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, run } from './script-utils.mjs'

const DEFAULT_LEDGER = path.join(ROOT, 'docs', 'upstream-signals.yml')
const DEFAULT_REPO = 'gitbrent/PptxGenJS'
const GH_FIELDS = ['number', 'title', 'state', 'labels', 'createdAt', 'updatedAt', 'url'].join(',')

function usage() {
	return `Usage: pnpm run upstream:signals:check -- [options]

Checks upstream GitHub issue and PR metadata against docs/upstream-signals.yml.
The command is read-only: it never edits the ledger.

Options:
  --repo <owner/repo>          GitHub repo to query (default: ledger source_repo or ${DEFAULT_REPO})
  --ledger <path>             Signal ledger path (default: docs/upstream-signals.yml)
  --type <all|issue|pr>       Metadata kind to check (default: all)
  --state <all|open|closed>   GitHub state to fetch (default: all)
  --limit <n>                 Max items to fetch per kind via gh (default: 1000)
  --print-limit <n>           Max missing items to print per kind; 0 prints all (default: 50)
  --created-since <date>      Only report items created on or after date
  --updated-since <date>      Only report items updated on or after date
  --json                      Print machine-readable JSON
  --fail-on-missing           Exit with code 1 when untreated items are found
  --help                      Show this help

Examples:
  pnpm run upstream:signals:check
  pnpm run upstream:signals:check -- --state open --type issue
  pnpm run upstream:signals:check -- --created-since 2026-06-07
  pnpm run upstream:signals:check:json -- --updated-since 2026-06-07
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

function parsePositiveInteger(name, value) {
	if (!/^\d+$/.test(value)) throw new Error(name + ' must be a non-negative integer')
	const parsed = Number(value)
	if (!Number.isSafeInteger(parsed)) throw new Error(name + ' is too large')
	return parsed
}

function parseDateOption(name, value) {
	const timestamp = Date.parse(value)
	if (Number.isNaN(timestamp)) throw new Error(name + ' must be a parseable date')
	return new Date(timestamp)
}

function parseArgs(argv) {
	const options = {
		ledger: DEFAULT_LEDGER,
		repo: null,
		type: 'all',
		state: 'all',
		limit: 1000,
		printLimit: 50,
		createdSince: null,
		updatedSince: null,
		json: false,
		failOnMissing: false,
		help: false,
	}

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
		} else if (arg === '--fail-on-missing') {
			options.failOnMissing = true
			i += 1
		} else if (arg.startsWith('--repo')) {
			const result = readOptionValue(argv, i, '--repo')
			options.repo = result.value
			i += result.consumed
		} else if (arg.startsWith('--ledger')) {
			const result = readOptionValue(argv, i, '--ledger')
			options.ledger = path.resolve(ROOT, result.value)
			i += result.consumed
		} else if (arg.startsWith('--type')) {
			const result = readOptionValue(argv, i, '--type')
			options.type = result.value
			i += result.consumed
		} else if (arg.startsWith('--state')) {
			const result = readOptionValue(argv, i, '--state')
			options.state = result.value
			i += result.consumed
		} else if (arg.startsWith('--limit')) {
			const result = readOptionValue(argv, i, '--limit')
			options.limit = parsePositiveInteger('--limit', result.value)
			i += result.consumed
		} else if (arg.startsWith('--print-limit')) {
			const result = readOptionValue(argv, i, '--print-limit')
			options.printLimit = parsePositiveInteger('--print-limit', result.value)
			i += result.consumed
		} else if (arg.startsWith('--created-since')) {
			const result = readOptionValue(argv, i, '--created-since')
			options.createdSince = parseDateOption('--created-since', result.value)
			i += result.consumed
		} else if (arg.startsWith('--updated-since')) {
			const result = readOptionValue(argv, i, '--updated-since')
			options.updatedSince = parseDateOption('--updated-since', result.value)
			i += result.consumed
		} else {
			throw new Error('unknown option: ' + arg)
		}
	}

	if (!['all', 'issue', 'pr'].includes(options.type)) throw new Error('--type must be all, issue, or pr')
	if (!['all', 'open', 'closed'].includes(options.state)) throw new Error('--state must be all, open, or closed')
	if (options.limit < 1) throw new Error('--limit must be greater than 0')

	return options
}

function stripCommentOnlyLines(text) {
	return text
		.split(/\r?\n/)
		.filter((line) => !line.trimStart().startsWith('#'))
		.join('\n')
}

function parseSourceRepo(text) {
	const match = stripCommentOnlyLines(text).match(/^source_repo:\s*([^\s#]+)\s*$/m)
	return match?.[1] || DEFAULT_REPO
}

function escapeRegExp(value) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function parseLedgerItems(text) {
	const lines = stripCommentOnlyLines(text).split(/\r?\n/)
	const itemsIndex = lines.findIndex((line) => /^items:\s*(?:\[\])?\s*$/.test(line))
	if (itemsIndex < 0 || /\[\]\s*$/.test(lines[itemsIndex])) return []

	const blocks = []
	let current = []

	for (const line of lines.slice(itemsIndex + 1)) {
		if (/^\s*-\s+/.test(line)) {
			if (current.length > 0) blocks.push(current.join('\n'))
			current = [line]
		} else if (current.length > 0) {
			current.push(line)
		}
	}
	if (current.length > 0) blocks.push(current.join('\n'))
	return blocks
}

function addTreatedRef(treated, kind, number) {
	const numeric = Number(number)
	if (!Number.isSafeInteger(numeric)) return
	if (kind === 'issue') treated.issue.add(numeric)
	else if (kind === 'pr') treated.pr.add(numeric)
	else treated.generic.add(numeric)
}

function normalizeItemKind(block) {
	const typeMatch = block.match(/^\s*type:\s*([^\s#]+)\s*$/m)
	const value = typeMatch?.[1]
	if (value === 'issue') return 'issue'
	if (value === 'pr' || value === 'pull-request' || value === 'pull_request') return 'pr'
	return null
}

function parseTreatedRefs(text, repo) {
	const treated = {
		issue: new Set(),
		pr: new Set(),
		generic: new Set(),
	}
	const escapedRepo = escapeRegExp(repo)
	const githubUrlPattern = new RegExp('github\\.com/' + escapedRepo + '/(issues|pull)/(\\d+)', 'g')
	const shorthandPattern = new RegExp('\\b' + escapedRepo + '#(\\d+)\\b', 'g')

	for (const block of parseLedgerItems(text)) {
		const itemKind = normalizeItemKind(block)

		for (const match of block.matchAll(/\bupstream-(issue|pr)-(\d+)\b/g)) {
			addTreatedRef(treated, match[1] === 'issue' ? 'issue' : 'pr', match[2])
		}
		for (const match of block.matchAll(githubUrlPattern)) {
			addTreatedRef(treated, match[1] === 'issues' ? 'issue' : 'pr', match[2])
		}
		for (const match of block.matchAll(shorthandPattern)) {
			addTreatedRef(treated, itemKind, match[1])
		}
	}

	return treated
}

function isTreated(treated, kind, number) {
	return treated.generic.has(number) || treated[kind].has(number)
}

function kindsForType(type) {
	if (type === 'issue') return ['issue']
	if (type === 'pr') return ['pr']
	return ['issue', 'pr']
}

async function fetchGithubItems(kind, options) {
	const command = kind === 'issue' ? 'issue' : 'pr'
	const result = await run(
		'gh',
		[
			command,
			'list',
			'--repo',
			options.repo,
			'--state',
			options.state,
			'--limit',
			String(options.limit),
			'--json',
			GH_FIELDS,
		],
		{ capture: true }
	)
	return JSON.parse(result.stdout)
}

function timestamp(value) {
	const parsed = Date.parse(value)
	return Number.isNaN(parsed) ? null : parsed
}

function matchesDateFilters(item, options) {
	if (options.createdSince && timestamp(item.createdAt) < options.createdSince.getTime()) return false
	if (options.updatedSince && timestamp(item.updatedAt) < options.updatedSince.getTime()) return false
	return true
}

function labelNames(labels) {
	if (!Array.isArray(labels)) return []
	return labels.map((label) => label.name).filter(Boolean)
}

function normalizeItem(item) {
	return {
		number: item.number,
		title: item.title,
		state: item.state,
		labels: labelNames(item.labels),
		createdAt: item.createdAt,
		updatedAt: item.updatedAt,
		url: item.url,
	}
}

function summarizeTreated(treated) {
	return {
		issues: treated.issue.size,
		pullRequests: treated.pr.size,
		genericRefs: treated.generic.size,
	}
}

function printItems(title, items, printLimit) {
	console.log('')
	console.log(title + ': ' + items.length)
	if (items.length === 0) return

	const limit = printLimit === 0 ? items.length : Math.min(printLimit, items.length)
	for (const item of items.slice(0, limit)) {
		const labels = item.labels.length > 0 ? ' labels: ' + item.labels.join(', ') : ''
		console.log('#' + item.number + ' [' + item.state + '] ' + item.title)
		if (labels) console.log('  ' + labels.trim())
		console.log('  created: ' + item.createdAt + ' updated: ' + item.updatedAt)
		console.log('  ' + item.url)
	}

	const remaining = items.length - limit
	if (remaining > 0) {
		console.log('  ... ' + remaining + ' more; use --print-limit 0 to print all or --json for full output.')
	}
}

function printHumanReport(report) {
	console.log('Upstream signal check')
	console.log('Repo: ' + report.repo)
	console.log('Ledger: ' + path.relative(ROOT, report.ledger))
	console.log(
		'Treated refs: ' +
			report.treated.issues +
			' issues, ' +
			report.treated.pullRequests +
			' pull requests, ' +
			report.treated.genericRefs +
			' generic refs'
	)
	console.log(
		'Missing: ' + report.counts.missingIssues + ' issues, ' + report.counts.missingPullRequests + ' pull requests'
	)

	printItems('Issues', report.missing.issues, report.options.printLimit)
	printItems('Pull requests', report.missing.pullRequests, report.options.printLimit)
}

async function main() {
	const options = parseArgs(process.argv.slice(2))
	if (options.help) {
		console.log(usage())
		return
	}

	const ledgerText = await fs.readFile(options.ledger, 'utf8')
	options.repo ||= parseSourceRepo(ledgerText)

	const treated = parseTreatedRefs(ledgerText, options.repo)
	const missing = {
		issues: [],
		pullRequests: [],
	}

	const results = await Promise.all(
		kindsForType(options.type).map(async (kind) => {
			const items = await fetchGithubItems(kind, options)
			return { kind, items }
		})
	)

	for (const { kind, items } of results) {
		const target = kind === 'issue' ? missing.issues : missing.pullRequests
		for (const item of items) {
			if (isTreated(treated, kind, item.number)) continue
			if (!matchesDateFilters(item, options)) continue
			target.push(normalizeItem(item))
		}
	}

	const report = {
		repo: options.repo,
		ledger: options.ledger,
		options: {
			type: options.type,
			state: options.state,
			limit: options.limit,
			printLimit: options.printLimit,
			createdSince: options.createdSince?.toISOString() || null,
			updatedSince: options.updatedSince?.toISOString() || null,
		},
		treated: summarizeTreated(treated),
		counts: {
			missingIssues: missing.issues.length,
			missingPullRequests: missing.pullRequests.length,
			missingTotal: missing.issues.length + missing.pullRequests.length,
		},
		missing,
	}

	if (options.json) console.log(JSON.stringify(report, null, 2))
	else printHumanReport(report)

	if (options.failOnMissing && report.counts.missingTotal > 0) process.exitCode = 1
}

main().catch((error) => {
	console.error(error.message)
	process.exitCode = 1
})
