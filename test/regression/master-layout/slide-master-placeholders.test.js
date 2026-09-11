import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	listEntries,
	selfClosingTags,
	xmlAttributes,
} from '../../helpers.js'

// A slide seeds every layout placeholder it leaves empty with a copy of that placeholder's options.
// The copy was shallow, so its nested `fill` and `hyperlink` objects were the layout's own:
// registering the slide's rels stamped the slide's ids onto them, and the layout then emitted an
// id from the slide's numbering. Seeding also ran after media encoding, so an image the seeded
// copy registered was never loaded and its part never written.
const PNG =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

/** A master with one body placeholder carrying `extra` options. */
function seededMaster(extra) {
	return {
		title: 'SEEDED',
		objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 6, h: 3, ...extra }, text: '' } }],
	}
}

/**
 * In every slide and layout part: rel ids are unique, every `r:embed` / `r:id` / `r:link` names a
 * rel the part declares, and every internal rel target is in the package.
 */
async function assertRelsResolve(zip) {
	const entries = listEntries(zip)
	for (const part of entries.filter((name) => /^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(name))) {
		const relsName = part.replace(/([^/]+)$/, '_rels/$1.rels')
		const rels = entries.includes(relsName)
			? selfClosingTags(await readEntry(zip, relsName), 'Relationship').map((tag) => xmlAttributes(tag))
			: []
		const ids = rels.map((rel) => rel.Id)
		assert(new Set(ids).size === ids.length, `${relsName} rel ids are unique: ${ids.join(',')}`)
		for (const [, id] of (await readEntry(zip, part)).matchAll(/r:(?:embed|id|link)="([^"]+)"/g)) {
			assert(ids.includes(id), `${part} references ${id}, which ${relsName} does not declare`)
		}
		for (const rel of rels.filter((candidate) => candidate.TargetMode !== 'External')) {
			const target = new URL(rel.Target, `file:///${part}`).pathname.slice(1)
			assert(entries.includes(target), `${relsName} ${rel.Id} targets ${rel.Target}, which is not in the package`)
		}
	}
}

defineRegressionSuite('Seeded layout placeholders keep their own rels', [
	{
		name: 'an image-fill placeholder a slide leaves empty, beside slide hyperlinks',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster(seededMaster({ fill: { image: { data: PNG } } }))
				const s = p.addSlide({ masterTitle: 'SEEDED' })
				s.addText('one', { x: 1, y: 5, w: 3, h: 0.5, hyperlink: { url: 'https://example.invalid/1' } })
				s.addText('two', { x: 4, y: 5, w: 3, h: 0.5, hyperlink: { url: 'https://example.invalid/2' } })
			})
			await assertRelsResolve(zip)
		},
	},
	{
		name: 'an image-fill placeholder loaded from a path, beside slide hyperlinks',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster(seededMaster({ fill: { image: { path: 'demos/common/images/cc_logo.jpg' } } }))
				const s = p.addSlide({ masterTitle: 'SEEDED' })
				s.addText('one', { x: 1, y: 5, w: 3, h: 0.5, hyperlink: { url: 'https://example.invalid/1' } })
				s.addText('two', { x: 4, y: 5, w: 3, h: 0.5, hyperlink: { url: 'https://example.invalid/2' } })
			})
			await assertRelsResolve(zip)
		},
	},
	{
		name: 'a hyperlinked placeholder a slide leaves empty, beside a slide image',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster(seededMaster({ hyperlink: { url: 'https://example.invalid/placeholder' } }))
				p.addSlide({ masterTitle: 'SEEDED' }).addImage({ data: PNG, x: 1, y: 5, w: 1, h: 1 })
			})
			await assertRelsResolve(zip)
		},
	},
])

defineRegressionSuite('Slide master placeholders [legacy bug-18]', [
	{
		name: 'master with title+body, only title populated → body stub emitted as placeholder (non-empty <a:lstStyle>)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } },
						{ placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 2, w: 9, h: 5 }, text: '' } },
					],
				})
				const s = p.addSlide({ masterTitle: 'TEST_MASTER_B18' })
				s.addText('Title Only', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// Two <p:sp> blocks must exist (title populated + body stub).
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(
				spBlocks.length === 2,
				'expected exactly 2 <p:sp> blocks (title + body stub); got ' + spBlocks.length + '\nxml: ' + xml
			)
			// Locate the body stub: the <p:sp> with <p:ph type="body".
			const bodyStub = spBlocks.find((sp) => /<p:ph[^>]*type="body"/.test(sp))
			assert(bodyStub, 'expected a <p:sp> with <p:ph type="body" .../> in the body stub; got: ' + xml)
			// Critical: body stub must NOT have a self-closing <a:lstStyle/>.
			// It must have the placeholder branch's <a:lstStyle>...</a:lstStyle> with paragraph properties.
			assert(
				bodyStub.indexOf('<a:lstStyle/>') === -1,
				'expected non-empty <a:lstStyle>...</a:lstStyle> in body stub (placeholder branch); got self-closing <a:lstStyle/>: ' +
					bodyStub
			)
			assert(
				/<a:lstStyle>[\s\S]+?<\/a:lstStyle>/.test(bodyStub),
				'expected <a:lstStyle>...</a:lstStyle> with content in body stub; got: ' + bodyStub
			)
		},
	},
	{
		name: 'master with single placeholder fully populated → exactly one <p:sp>, no empty stub',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18_SINGLE',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } },
					],
				})
				const s = p.addSlide({ masterTitle: 'TEST_MASTER_B18_SINGLE' })
				s.addText('Filled', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(
				spBlocks.length === 1,
				'expected exactly 1 <p:sp> block (no empty stub); got ' + spBlocks.length + '\nxml: ' + xml
			)
			assert(
				/<a:t>Filled<\/a:t>/.test(spBlocks[0]),
				'expected populated text run <a:t>Filled</a:t>; got: ' + spBlocks[0]
			)
		},
	},
	{
		name: 'master with two placeholders both populated → two <p:sp> with text runs, no empty stub',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18_BOTH',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } },
						{ placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 2, w: 9, h: 5 }, text: '' } },
					],
				})
				const s = p.addSlide({ masterTitle: 'TEST_MASTER_B18_BOTH' })
				s.addText('My Title', { placeholder: 'title' })
				s.addText('My Body', { placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(
				spBlocks.length === 2,
				'expected exactly 2 <p:sp> blocks (no extra stub); got ' + spBlocks.length + '\nxml: ' + xml
			)
			// Both blocks should contain <a:t>...</a:t> populated runs.
			const populated = spBlocks.filter((sp) => /<a:t>[^<]+<\/a:t>/.test(sp))
			assert(
				populated.length === 2,
				'expected both <p:sp> blocks to contain text runs <a:t>...</a:t>; got ' +
					populated.length +
					'\nblocks: ' +
					JSON.stringify(spBlocks)
			)
			assert(
				spBlocks.some((sp) => /<a:t>My Title<\/a:t>/.test(sp)),
				'expected <a:t>My Title</a:t> in one block; got: ' + xml
			)
			assert(
				spBlocks.some((sp) => /<a:t>My Body<\/a:t>/.test(sp)),
				'expected <a:t>My Body</a:t> in one block; got: ' + xml
			)
		},
	},
])
