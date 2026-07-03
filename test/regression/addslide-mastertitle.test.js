import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// A master with a single title placeholder; a slide created against it should
// inherit the placeholder (rendered as a <p:sp> with <p:ph type="title" .../>).
function defineMaster(p, title) {
	p.defineSlideMaster({
		title,
		objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } }],
	})
}

// Capture warnings emitted while running `fn` (console.warn is the library's warning sink, see src/log.ts).
async function withCapturedWarnings(fn) {
	const original = console.warn
	const messages = []
	console.warn = (msg) => messages.push(String(msg))
	try {
		await fn()
	} finally {
		console.warn = original
	}
	return messages
}

defineRegressionSuite('addSlide masterTitle', [
	{
		name: 'canonical `masterTitle` applies the named master',
		fn: async () => {
			const { zip } = await build((p) => {
				defineMaster(p, 'MT_MASTER')
				p.addSlide({ masterTitle: 'MT_MASTER' }).addText('Title', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="title"/.test(xml), 'expected title placeholder from masterTitle; got: ' + xml)
		},
	},
	{
		name: 'deprecated `masterName` still applies the master and warns',
		fn: async () => {
			let zip
			const warnings = await withCapturedWarnings(async () => {
				;({ zip } = await build((p) => {
					defineMaster(p, 'MN_MASTER')
					p.addSlide({ masterName: 'MN_MASTER' }).addText('Title', { placeholder: 'title' })
				}))
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="title"/.test(xml), 'expected title placeholder from deprecated masterName; got: ' + xml)
			assert(
				warnings.some((m) => /masterName is deprecated/.test(m)),
				'expected a masterName deprecation warning; got: ' + JSON.stringify(warnings)
			)
		},
	},
])
