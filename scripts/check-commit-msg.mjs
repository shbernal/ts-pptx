#!/usr/bin/env node
// Reject a commit message whose subject line was mangled by a shell-quoting slip
// — a multi-line message passed through the wrong shell's heredoc/here-string
// (PowerShell `@'...'@` vs POSIX `<<'EOF'`). Wired as a lefthook `commit-msg`
// hook. Longer-than-one-line messages should be written to a file and committed
// with `git commit -F <file>`, which keeps the text off the shell command line.
import { readFile } from 'node:fs/promises'

const msgPath = process.argv[2]
if (!msgPath) {
	console.error('check-commit-msg: no commit-message file path was provided.')
	process.exit(2)
}

const raw = await readFile(msgPath, 'utf8')

// The subject is the first non-empty, non-comment line (git strips `#` lines).
const subject = raw
	.split('\n')
	.map((line) => line.trimEnd())
	.find((line) => line.length > 0 && !line.startsWith('#'))

if (!subject) {
	// Empty message — let git's own empty-message handling deal with it.
	process.exit(0)
}

// Signatures of a shell-quoting slip leaking into the message: a leading `@` or
// backtick (a here-string opener in the wrong shell), or a stray `@'` / `'@`
// here-string delimiter anywhere in the subject line.
const mangled = /^[@`]/.test(subject) || subject.includes("@'") || subject.includes("'@")

if (mangled) {
	console.error(`check-commit-msg: subject line looks shell-mangled:\n  ${subject}\n`)
	console.error('Write the message to a file and use `git commit -F <file>` instead of an inline heredoc/here-string.')
	process.exit(1)
}

process.exit(0)
