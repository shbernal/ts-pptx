'use strict'

const { build, readEntry, assert } = require('./helpers')

module.exports = [
	{
		name: 'master with title+body, only title populated → body stub emitted as placeholder (non-empty <a:lstStyle>)',
		fn: async () => {
			const { zip } = await build(p => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } },
						{ placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 2, w: 9, h: 5 }, text: '' } }
					]
				})
				const s = p.addSlide({ masterName: 'TEST_MASTER_B18' })
				s.addText('Title Only', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// Two <p:sp> blocks must exist (title populated + body stub).
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(spBlocks.length === 2,
				'expected exactly 2 <p:sp> blocks (title + body stub); got ' + spBlocks.length + '\nxml: ' + xml)
			// Locate the body stub: the <p:sp> with <p:ph type="body".
			const bodyStub = spBlocks.find(sp => /<p:ph[^>]*type="body"/.test(sp))
			assert(bodyStub,
				'expected a <p:sp> with <p:ph type="body" .../> in the body stub; got: ' + xml)
			// Critical: body stub must NOT have a self-closing <a:lstStyle/>.
			// It must have the placeholder branch's <a:lstStyle>...</a:lstStyle> with paragraph properties.
			assert(bodyStub.indexOf('<a:lstStyle/>') === -1,
				'expected non-empty <a:lstStyle>...</a:lstStyle> in body stub (placeholder branch); got self-closing <a:lstStyle/>: ' + bodyStub)
			assert(/<a:lstStyle>[\s\S]+?<\/a:lstStyle>/.test(bodyStub),
				'expected <a:lstStyle>...</a:lstStyle> with content in body stub; got: ' + bodyStub)
		}
	},
	{
		name: 'master with single placeholder fully populated → exactly one <p:sp>, no empty stub',
		fn: async () => {
			const { zip } = await build(p => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18_SINGLE',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } }
					]
				})
				const s = p.addSlide({ masterName: 'TEST_MASTER_B18_SINGLE' })
				s.addText('Filled', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(spBlocks.length === 1,
				'expected exactly 1 <p:sp> block (no empty stub); got ' + spBlocks.length + '\nxml: ' + xml)
			assert(/<a:t>Filled<\/a:t>/.test(spBlocks[0]),
				'expected populated text run <a:t>Filled</a:t>; got: ' + spBlocks[0])
		}
	},
	{
		name: 'master with two placeholders both populated → two <p:sp> with text runs, no empty stub',
		fn: async () => {
			const { zip } = await build(p => {
				p.defineSlideMaster({
					title: 'TEST_MASTER_B18_BOTH',
					objects: [
						{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } },
						{ placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 2, w: 9, h: 5 }, text: '' } }
					]
				})
				const s = p.addSlide({ masterName: 'TEST_MASTER_B18_BOTH' })
				s.addText('My Title', { placeholder: 'title' })
				s.addText('My Body', { placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const spBlocks = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []
			assert(spBlocks.length === 2,
				'expected exactly 2 <p:sp> blocks (no extra stub); got ' + spBlocks.length + '\nxml: ' + xml)
			// Both blocks should contain <a:t>...</a:t> populated runs.
			const populated = spBlocks.filter(sp => /<a:t>[^<]+<\/a:t>/.test(sp))
			assert(populated.length === 2,
				'expected both <p:sp> blocks to contain text runs <a:t>...</a:t>; got ' + populated.length + '\nblocks: ' + JSON.stringify(spBlocks))
			assert(spBlocks.some(sp => /<a:t>My Title<\/a:t>/.test(sp)),
				'expected <a:t>My Title</a:t> in one block; got: ' + xml)
			assert(spBlocks.some(sp => /<a:t>My Body<\/a:t>/.test(sp)),
				'expected <a:t>My Body</a:t> in one block; got: ' + xml)
		}
	}
]
