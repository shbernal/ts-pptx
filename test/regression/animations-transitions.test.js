import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Write-side slide transitions and preset build animations
// (docs/animations-and-transitions.md, Phase 1). The emitters reproduce
// PowerPoint-authored XML verbatim, so these assert byte-equality against the
// PowerPoint oracles in test/read/fixtures (slide-transition / slide-animation-*).

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function loadOracle(name) {
	return JSON.parse(await readFile(path.join(__dirname, '..', 'read', 'fixtures', `${name}.oracle.json`), 'utf8'))
}

async function slideXml(buildFn, n = 1) {
	const { zip } = await build(buildFn)
	return readEntry(zip, `ppt/slides/slide${n}.xml`)
}

function timingOf(xml) {
	const m = xml.match(/<p:timing>[\s\S]*<\/p:timing>/)
	return m ? m[0] : null
}

function transitionOf(xml) {
	const m = xml.match(/<p:transition[\s\S]*?<\/p:transition>|<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/)
	return m ? m[0] : null
}

defineRegressionSuite('Slide transitions (write)', [
	{
		name: 'emits each PowerPoint transition form byte-for-byte (bare + mc:AlternateContent)',
		fn: async () => {
			const oracle = await loadOracle('slide-transition')
			const inputs = [
				{ type: 'fade' },
				{ type: 'push', durationMs: 1250, speed: 'slow', variant: { dir: 'd' } },
				{ type: 'wipe', speed: 'med', variant: { dir: 'u' } },
				{ type: 'cut' },
				{ type: 'dissolve', durationMs: 2000, speed: 'slow' },
				{ type: 'fade', speed: 'med', advanceOnClick: false, advanceAfterMs: 3000 },
			]
			const { zip } = await build((p) => {
				inputs.forEach((t) => {
					const s = p.addSlide()
					s.addText('x', { x: 1, y: 1, w: 1, h: 1 })
					s.transition = t
				})
			})
			for (let i = 0; i < inputs.length; i++) {
				const xml = await readEntry(zip, `ppt/slides/slide${i + 1}.xml`)
				assert(transitionOf(xml) === oracle.slides[i].transitionXml, `slide ${i + 1} transition matches oracle`)
				// Positioned between p:clrMapOvr and (absent) p:timing, inside p:sld.
				assert(
					/<\/p:clrMapOvr>(<p:transition|<mc:AlternateContent)/.test(xml),
					`slide ${i + 1} positioned after clrMapOvr`
				)
			}
		},
	},
	{
		name: 'derives a speed bucket from durationMs when speed is omitted',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().transition = { type: 'fade', durationMs: 1500 }
			})
			assert(/p14:dur="1500"/.test(xml), 'carries exact p14:dur')
			assert(/spd="slow"/.test(xml), 'derives slow bucket for 1500ms')
		},
	},
])

defineRegressionSuite('Preset build animations (write)', [
	{
		name: 'emits the rich multi-effect mainSeq byte-for-byte',
		fn: async () => {
			const oracle = await loadOracle('slide-animation-rich')
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				for (const nm of ['ent-fade-click', 'ent-fly-after', 'emph-grow-with', 'exit-fade-click'])
					s.addText(nm, { x: 1, y: 1, w: 3, h: 1, objectName: nm })
				s.addAnimation({ preset: 'fadeIn', shapeIndex: 0, trigger: 'onClick' })
				s.addAnimation({ preset: 'flyIn', shapeIndex: 1, trigger: 'afterPrevious' })
				s.addAnimation({ preset: 'grow', shapeIndex: 2, trigger: 'withPrevious' })
				s.addAnimation({ preset: 'fadeOut', shapeIndex: 3, trigger: 'onClick' })
			})
			assert(timingOf(xml) === oracle.timingXml, 'rich timing tree matches PowerPoint oracle')
		},
	},
	{
		name: 'emits the basic single fade-on-click mainSeq byte-for-byte',
		fn: async () => {
			const oracle = await loadOracle('slide-animation-basic')
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addText('fade-target', { x: 1, y: 1, w: 3, h: 1, objectName: 'fade-target' })
				s.addAnimation({ preset: 'fadeIn', shapeIndex: 0 })
			})
			assert(timingOf(xml) === oracle.timingXml, 'basic timing tree matches PowerPoint oracle')
		},
	},
	{
		name: 'resolves the target shape by objectName',
		fn: async () => {
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addText('a', { x: 1, y: 1, w: 1, h: 1, objectName: 'a' })
				s.addText('b', { x: 1, y: 2, w: 1, h: 1, objectName: 'b' })
				s.addAnimation({ preset: 'fadeIn', objectName: 'b' }) // spid = idx 1 + 2 = 3
			})
			assert(/<p:bldP spid="3" grpId="0"\/>/.test(xml), 'bldP targets resolved spid 3')
			assert(/<p:spTgt spid="3"\/>/.test(xml), 'effect targets resolved spid 3')
		},
	},
])
