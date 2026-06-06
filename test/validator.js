// Test-time helper: validate a `.pptx` Buffer against Microsoft's
// OpenXmlValidator (via the OOXML-Validator CLI binary installed under
// tools/ooxml-validator/bin/). Returns a Promise<Array> of validation
// errors (empty array on a clean file).

import { execFile as execFileCallback } from 'node:child_process'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const execFile = promisify(execFileCallback)

const VALIDATOR = path.resolve(__dirname, '..', 'tools', 'ooxml-validator', 'bin', 'OOXMLValidatorCLI')

async function isInstalled() {
	try {
		await fs.access(VALIDATOR)
		return true
	} catch {
		return false
	}
}

async function runValidatorOnFile(filePath, fileFormat) {
	const args = [filePath]
	const env = {
		...process.env,
		DOTNET_BUNDLE_EXTRACT_BASE_DIR: process.env.DOTNET_BUNDLE_EXTRACT_BASE_DIR || os.tmpdir(),
	}
	if (fileFormat) args.push(fileFormat)

	// The CLI prints a JSON array to stdout regardless of whether
	// errors were found; exit code is 0 in both cases.
	const { stdout } = await execFile(VALIDATOR, args, { env, maxBuffer: 32 * 1024 * 1024 })
	try {
		return JSON.parse(stdout || '[]')
	} catch {
		throw new Error('failed to parse OOXMLValidatorCLI output: ' + String(stdout).slice(0, 500))
	}
}

async function validateBuf(buf, fileFormat) {
	if (!(await isInstalled())) {
		throw new Error('OOXMLValidatorCLI not installed. Run ./tools/ooxml-validator/install.sh')
	}
	const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgen-schema-'))
	const tmp = path.join(tmpDir, 'fixture.pptx')
	await fs.writeFile(tmp, buf)
	try {
		return await runValidatorOnFile(tmp, fileFormat)
	} finally {
		await fs.rm(tmpDir, { recursive: true, force: true })
	}
}

export { isInstalled, validateBuf, runValidatorOnFile, VALIDATOR }
