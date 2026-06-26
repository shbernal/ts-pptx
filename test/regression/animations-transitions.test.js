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

function sndAcOf(xml) {
	const m = xml.match(/<p:sndAc>[\s\S]*?<\/p:sndAc>/)
	return m ? m[0] : null
}

// PptxGenJS numbers slide rels by its own deterministic scheme (media first), so an
// embedded sound's rId differs from PowerPoint's authored value; normalize for the
// structural comparison (internal consistency is asserted separately).
function normRid(s) {
	return s == null ? s : s.replace(/rId\d+/g, 'rId#')
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
		// Phase 2 capability B: the expanded preset set. One on-click effect per
		// preset across all three classes (entr/emph/exit), reproducing
		// slide-animation-presets.pptx (spids 2..9). Asserts the whole timing tree
		// byte-for-byte — reconfirms fadeIn/flyIn/grow/fadeOut and pins the new
		// appear/wipe/spin/flyOut templates against the PowerPoint oracle.
		name: 'emits every preset (incl. appear/wipe/spin/flyOut) byte-for-byte',
		fn: async () => {
			const oracle = await loadOracle('slide-animation-presets')
			const order = ['fadeIn', 'flyIn', 'appear', 'wipe', 'grow', 'spin', 'fadeOut', 'flyOut']
			const names = [
				'entr-fadeIn',
				'entr-flyIn',
				'entr-appear',
				'entr-wipe',
				'emph-grow',
				'emph-spin',
				'exit-fadeOut',
				'exit-flyOut',
			]
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				names.forEach((nm) => s.addText(nm, { x: 1, y: 1, w: 3, h: 1, objectName: nm }))
				order.forEach((preset, i) => s.addAnimation({ preset, shapeIndex: i, trigger: 'onClick' }))
			})
			assert(timingOf(xml) === oracle.timingXml, 'full preset timing tree matches PowerPoint oracle')
		},
	},
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

// Phase 2 capability C: transition sounds (p:sndAc). The writer reproduces
// PowerPoint's sndAc forms (embedded start / looped / stop-previous) plus the audio
// rel + media part + content-type graph, deduping identical sound bytes. rIds use
// PptxGenJS's own numbering, so the sndAc is compared rId-normalized to the oracle.
const SOUND_WAV = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAIA+AAACABAAZGF0YQAAAAA='

defineRegressionSuite('Transition sounds (write)', [
	{
		name: 'emits the sndAc start/loop/stop forms matching the PowerPoint oracle (rId-normalized)',
		fn: async () => {
			const oracle = await loadOracle('slide-transition-sound')
			const { zip } = await build((p) => {
				p.addSlide().transition = { type: 'fade', durationMs: 2000, sound: { data: SOUND_WAV, name: 'ding.wav' } }
				p.addSlide().transition = {
					type: 'fade',
					durationMs: 2000,
					sound: { data: SOUND_WAV, name: 'ding.wav', loop: true },
				}
				p.addSlide().transition = { type: 'fade', durationMs: 2000, sound: { stopPrevious: true } }
			})
			for (let i = 0; i < oracle.slides.length; i++) {
				const xml = await readEntry(zip, `ppt/slides/slide${i + 1}.xml`)
				assert(
					normRid(sndAcOf(xml)) === normRid(oracle.slides[i].soundRels.sndAcXml),
					`slide ${i + 1} sndAc matches oracle`
				)
				// sndAc sits inside the transition, after the type element.
				assert(/<p:fade\/><p:sndAc>/.test(xml), `slide ${i + 1} sndAc follows the transition type element`)
			}
		},
	},
	{
		name: 'wires each start sound to a real audio rel + embeds the WAV; stop-previous adds neither',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().transition = { type: 'fade', sound: { data: SOUND_WAV, name: 'ding.wav' } }
				p.addSlide().transition = { type: 'fade', sound: { stopPrevious: true } }
			})
			const xml1 = await readEntry(zip, 'ppt/slides/slide1.xml')
			const rels1 = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			const embedRid = xml1.match(/<p:snd r:embed="(rId\d+)"/)[1]
			// The r:embed points at a real ECMA audio relationship to an embedded media part.
			const relMatch = rels1.match(
				new RegExp(`<Relationship Id="${embedRid}"[^>]*relationships/audio[^>]*Target="([^"]+)"`)
			)
			assert(relMatch, 'start sound r:embed resolves to an audio relationship')
			const part = relMatch[1].replace('..', 'ppt')
			assert(zip.file(part) != null, `embedded WAV part ${part} is present`)
			// Stop-previous slide carries no sndAc rel and no media reference.
			const rels2 = await readEntry(zip, 'ppt/slides/_rels/slide2.xml.rels')
			assert(!/relationships\/audio/.test(rels2), 'stop-previous slide has no audio rel')
		},
	},
	{
		name: 'dedups identical sound bytes to one media part across slides + emits the wav content type',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().transition = { type: 'fade', sound: { data: SOUND_WAV, name: 'ding.wav' } }
				p.addSlide().transition = { type: 'fade', sound: { data: SOUND_WAV, name: 'ding.wav', loop: true } }
			})
			const wavParts = Object.keys(zip.files).filter((k) => k.startsWith('ppt/media/') && k.endsWith('.wav'))
			assert(wavParts.length === 1, 'identical sound bytes collapse to a single media part')
			const ct = await readEntry(zip, '[Content_Types].xml')
			assert(
				/<Default Extension="wav" ContentType="audio\/x-wav"\/>/.test(ct),
				'wav Default content type is audio/x-wav'
			)
		},
	},
])
