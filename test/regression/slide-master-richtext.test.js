import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// `defineSlideMaster` text objects previously wrapped `text.text` in a fresh
// one-item array unconditionally. A rich-text array (`text: [{ text, options }, ...]`)
// therefore arrived at addTextDefinition as `[{ text: TextProps[] }]` and the runs
// were lost or stringified instead of serialized (upstream issue #962).

async function findLayoutXmlContaining(zip, needle) {
	const layouts = listEntries(zip).filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p))
	for (const path of layouts) {
		const xml = await readEntry(zip, path)
		if (xml.includes(needle)) return xml
	}
	throw new Error(`no slideLayout XML contains "${needle}"; layouts: ${layouts.join(', ')}`)
}

defineRegressionSuite('Slide master rich-text arrays (#962)', [
	{
		name: 'master text object with TextProps[] emits one run per item with run options',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_962',
					objects: [
						{
							text: {
								text: [{ text: 'BoldRun962', options: { bold: true } }, { text: 'PlainRun962' }],
								options: { x: 0.5, y: 0.5, w: 9, h: 1 },
							},
						},
					],
				})
				p.addSlide({ masterName: 'TEST_MASTER_962' })
			})
			const xml = await findLayoutXmlContaining(zip, 'BoldRun962')
			const runs = xml.match(/<a:r>[\s\S]*?<\/a:r>/g) || []
			const boldRun = runs.find((r) => r.includes('BoldRun962'))
			const plainRun = runs.find((r) => r.includes('PlainRun962'))
			assert(boldRun, 'expected an <a:r> run containing "BoldRun962"; got runs: ' + runs.join('\n'))
			assert(plainRun, 'expected an <a:r> run containing "PlainRun962"; got runs: ' + runs.join('\n'))
			assert(/<a:rPr[^>]*\bb="1"/.test(boldRun), 'expected b="1" on the bold run rPr; got: ' + boldRun)
			assert(!/\bb="1"/.test(plainRun), 'expected no b="1" on the plain run; got: ' + plainRun)
			assert(!xml.includes('[object Object]'), 'rich-text array was stringified: ' + xml)
		},
	},
	{
		name: 'master text object with plain string still emits a single run',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_962_STR',
					objects: [{ text: { text: 'PlainString962', options: { x: 0.5, y: 0.5, w: 9, h: 1 } } }],
				})
				p.addSlide({ masterName: 'TEST_MASTER_962_STR' })
			})
			const xml = await findLayoutXmlContaining(zip, 'PlainString962')
			const runs = (xml.match(/<a:r>[\s\S]*?<\/a:r>/g) || []).filter((r) => r.includes('PlainString962'))
			assert(runs.length === 1, 'expected exactly one run with "PlainString962"; got ' + runs.length)
		},
	},
])
