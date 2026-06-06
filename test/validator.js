// Test-time helper: validate a `.pptx` Buffer against Microsoft's
// OpenXmlValidator (via the OOXML-Validator CLI binary installed under
// tools/ooxml-validator/bin/). Returns a Promise<Array> of validation
// errors (empty array on a clean file).

import { execFile } from 'node:child_process'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

const VALIDATOR = path.resolve(__dirname, '..', 'tools', 'ooxml-validator', 'bin', 'OOXMLValidatorCLI')

function isInstalled() {
	return fs.existsSync(VALIDATOR)
}

function runValidatorOnFile(filePath, fileFormat) {
	return new Promise((resolve, reject) => {
		const args = [filePath]
		const env = {
			...process.env,
			DOTNET_BUNDLE_EXTRACT_BASE_DIR: process.env.DOTNET_BUNDLE_EXTRACT_BASE_DIR || os.tmpdir(),
		}
		if (fileFormat) args.push(fileFormat)
		execFile(VALIDATOR, args, { env, maxBuffer: 32 * 1024 * 1024 }, (err, stdout) => {
			// The CLI prints a JSON array to stdout regardless of whether
			// errors were found; exit code is 0 in both cases.
			if (err && err.code !== 0) return reject(err)
			try {
				resolve(JSON.parse(stdout || '[]'))
			} catch {
				reject(new Error('failed to parse OOXMLValidatorCLI output: ' + String(stdout).slice(0, 500)))
			}
		})
	})
}

async function validateBuf(buf, fileFormat) {
	if (!isInstalled()) {
		throw new Error('OOXMLValidatorCLI not installed. Run ./tools/ooxml-validator/install.sh')
	}
	const tmp = path.join(os.tmpdir(), 'pptxgen-schema-' + process.pid + '-' + Date.now() + '.pptx')
	fs.writeFileSync(tmp, buf)
	try {
		return await runValidatorOnFile(tmp, fileFormat)
	} finally {
		try {
			fs.unlinkSync(tmp)
		} catch {
			/* ignore */
		}
	}
}

export { isInstalled, validateBuf, runValidatorOnFile, VALIDATOR }
