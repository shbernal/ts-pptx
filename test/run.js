'use strict'

const fs = require('fs')
const path = require('path')

const failures = []
const successes = []

async function loadAndRun() {
	const dir = __dirname
	const files = fs.readdirSync(dir).filter(f => /^bug-\d+\.test\.js$/.test(f)).sort()
	for (const f of files) {
		const full = path.join(dir, f)
		const cases = require(full)
		for (const c of cases) {
			try {
				await c.fn()
				successes.push(c.name)
				console.log('  ok ' + c.name)
			} catch (e) {
				failures.push({ name: c.name, error: e })
				console.log('  FAIL ' + c.name + ': ' + e.message)
			}
		}
	}
}

;(async () => {
	console.log('Running PptxGenJS regression tests')
	await loadAndRun()
	console.log('\nPassed: ' + successes.length + '  Failed: ' + failures.length)
	if (failures.length > 0) {
		failures.forEach(f => console.log(f.name + ' -- ' + (f.error.stack || f.error.message)))
		process.exit(1)
	}
})()
