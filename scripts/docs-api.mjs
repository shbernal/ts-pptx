import { spawnSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const root = process.cwd()
const outDir = path.join(root, 'docs', 'reference', 'api')
const typedocBin = path.join(root, 'node_modules', '.bin', process.platform === 'win32' ? 'typedoc.cmd' : 'typedoc')

function walkMarkdown(dir) {
	const out = []
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const filePath = path.join(dir, entry.name)
		if (entry.isDirectory()) {
			out.push(...walkMarkdown(filePath))
		} else if (entry.isFile() && entry.name.endsWith('.md')) {
			out.push(filePath)
		}
	}
	return out.sort()
}

function titleFromMarkdown(markdown, filePath) {
	const heading = markdown.match(/^#\s+(.+)$/m)
	if (heading) return heading[1].replace(/\s+\|.*$/, '').trim()
	const basename = path.basename(filePath, '.md')
	return basename === 'index' ? 'Public API Reference' : basename
}

function frontmatterFor(filePath, markdown) {
	const rel = path.relative(outDir, filePath).split(path.sep).join('/')
	const title = rel === 'index.md' ? 'Public API Reference' : titleFromMarkdown(markdown, filePath)
	const summary =
		rel === 'index.md'
			? 'Generated TypeDoc reference for the public PptxGenJS package exports.'
			: `Generated TypeDoc reference for ${title}.`

	return [
		'---',
		'doc-schema-version: 1',
		`title: "${title.replaceAll('"', '\\"')}"`,
		`summary: "${summary.replaceAll('"', '\\"')}"`,
		'read_when:',
		'  - Looking up public PptxGenJS API details',
		'  - Verifying generated TypeScript API documentation',
		'doc_type: "reference"',
		'---',
		'',
	].join('\n')
}

function escapeVueUnsafeHtml(markdown) {
	let inFence = false
	return markdown
		.split('\n')
		.map((line) => {
			if (line.trimStart().startsWith('```')) {
				inFence = !inFence
				return line
			}
			if (inFence) return line

			return line
				.split('`')
				.map((segment, index) => {
					if (index % 2 === 1) return segment
					return segment
						.replaceAll(/<\/([A-Za-z][A-Za-z0-9:._-]*)>/g, '&lt;/$1&gt;')
						.replaceAll(/<([A-Za-z][A-Za-z0-9:._-]*)(\s[^>\n]*)?>/g, '&lt;$1$2&gt;')
				})
				.join('`')
		})
		.join('\n')
}

rmSync(outDir, { force: true, recursive: true })
mkdirSync(outDir, { recursive: true })

const typedoc = spawnSync(typedocBin, ['--options', 'typedoc.docs.json'], {
	cwd: root,
	stdio: 'inherit',
})

if (typedoc.status !== 0) {
	process.exit(typedoc.status ?? 1)
}

const readmePath = path.join(outDir, 'README.md')
const indexPath = path.join(outDir, 'index.md')
if (!existsSync(indexPath) && existsSync(readmePath)) {
	copyFileSync(readmePath, indexPath)
}

for (const filePath of walkMarkdown(outDir)) {
	const markdown = readFileSync(filePath, 'utf8')
	const body = markdown.startsWith('---\n') ? markdown.replace(/^---\n[\s\S]*?\n---\n+/, '') : markdown
	const safeBody = escapeVueUnsafeHtml(body.trimStart())
	writeFileSync(filePath, `${frontmatterFor(filePath, safeBody)}${safeBody}`, 'utf8')
}
