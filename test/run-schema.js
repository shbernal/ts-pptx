'use strict'

// Runner for schema-validation fixtures (test/schema.test.js).
// Kept separate from test/run.js so a schema regression doesn't break
// the regular `npm test` while we baseline existing emitter output.

const fs = require('fs')
const path = require('path')

const { isInstalled, VALIDATOR } = require('./validator')

if (!isInstalled()) {
	console.error('OOXMLValidatorCLI not installed at: ' + VALIDATOR)
	console.error('Run: ./tools/ooxml-validator/install.sh')
	process.exit(2)
}

const failures = []
const successes = []

async function loadAndRun () {
	const file = path.join(__dirname, 'schema.test.js')
	if (!fs.existsSync(file)) {
		console.error('No schema fixtures found at ' + file)
		process.exit(2)
	}
	const cases = require(file)
	for (const c of cases) {
		try {
			await c.fn()
			successes.push(c.name)
			console.log('  ok ' + c.name)
		} catch (e) {
			failures.push({ name: c.name, error: e })
			console.log('  FAIL ' + c.name)
			const lines = String(e.message || e).split('\n')
			for (const ln of lines) console.log('       ' + ln)
		}
	}
}

;(async () => {
	console.log('Running PptxGenJS schema validation')
	await loadAndRun()
	console.log('\nPassed: ' + successes.length + '  Failed: ' + failures.length)
	if (failures.length > 0) process.exit(1)
})()
