#!/usr/bin/env node
'use strict'

// Compare the pinned OOXML-Validator version against the latest GitHub
// release. Designed to be run periodically (locally or in CI) to flag
// upstream drift. Exits 0 if up-to-date, 1 if a newer release is
// available, 2 on network/parse errors.

import fs from 'node:fs'
import https from 'node:https'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

const VERSION_FILE = path.join(__dirname, 'version.json')
const LATEST_URL = 'https://api.github.com/repos/mikeebowen/OOXML-Validator/releases/latest'

function fetchJson (url) {
	return new Promise((resolve, reject) => {
		const req = https.get(
			url,
			{
				headers: {
					'User-Agent': 'pptxgenjs-tooling',
					'Accept': 'application/vnd.github+json'
				}
			},
			res => {
				if (res.statusCode !== 200) {
					reject(new Error('HTTP ' + res.statusCode + ' from ' + url))
					return
				}
				let body = ''
				res.setEncoding('utf8')
				res.on('data', c => { body += c })
				res.on('end', () => {
					try { resolve(JSON.parse(body)) }
					catch (e) { reject(new Error('Bad JSON from ' + url + ': ' + e.message)) }
				})
			}
		)
		req.on('error', reject)
		req.setTimeout(15_000, () => req.destroy(new Error('timeout')))
	})
}

;(async () => {
	let pinned
	try {
		pinned = JSON.parse(fs.readFileSync(VERSION_FILE, 'utf8'))
	} catch (e) {
		console.error('failed to read ' + VERSION_FILE + ': ' + e.message)
		process.exit(2)
	}
	const pinnedVersion = String(pinned.version || '').replace(/^v/, '')

	let latest
	try {
		latest = await fetchJson(LATEST_URL)
	} catch (e) {
		console.error('failed to query GitHub: ' + e.message)
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
