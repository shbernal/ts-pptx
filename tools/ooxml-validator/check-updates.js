#!/usr/bin/env node

// Compare the pinned OOXML-Validator version against the latest GitHub
// release. Designed to be run periodically (locally or in CI) to flag
// upstream drift. Exits 0 if up-to-date, 1 if a newer release is
// available, 2 on network/parse errors.

import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

const VERSION_FILE = path.join(__dirname, 'version.json')
const LATEST_URL = 'https://api.github.com/repos/mikeebowen/OOXML-Validator/releases/latest'

async function fetchJson(url) {
	const response = await fetch(url, {
		headers: {
			'User-Agent': 'pptxgenjs-tooling',
			Accept: 'application/vnd.github+json',
		},
		signal: AbortSignal.timeout(15_000),
	})
	if (!response.ok) throw new Error('HTTP ' + response.status + ' from ' + url)
	return await response.json()
}

function getErrorMessage(error) {
	return error instanceof Error ? error.message : String(error)
}

;(async () => {
	let pinned
	try {
		pinned = JSON.parse(await fs.readFile(VERSION_FILE, 'utf8'))
	} catch (e) {
		console.error('failed to read ' + VERSION_FILE + ': ' + getErrorMessage(e))
		process.exit(2)
	}
	const pinnedVersion = String(pinned.version || '').replace(/^v/, '')

	let latest
	try {
		latest = await fetchJson(LATEST_URL)
	} catch (e) {
		console.error('failed to query GitHub: ' + getErrorMessage(e))
		process.exit(2)
	}
	const latestTag = String(latest.tag_name || '').replace(/^v/, '')

	console.log('pinned: v' + pinnedVersion)
	console.log('latest: v' + latestTag + ' (' + (latest.published_at || '?') + ')')

	if (!latestTag) {
		console.error('could not determine latest release tag')
		process.exit(2)
	}
	if (latestTag === pinnedVersion) {
		console.log('up-to-date')
		process.exit(0)
	}
	console.log('')
	console.log('newer release available: ' + latest.html_url)
	console.log('to upgrade, edit tools/ooxml-validator/version.json and re-run install.sh')
	process.exit(1)
})()
