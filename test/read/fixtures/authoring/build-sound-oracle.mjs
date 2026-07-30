import { readFile, writeFile } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import JSZip from 'jszip'

// This script lives in test/read/fixtures/authoring/, so the fixtures dir is its parent.
const FIX = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const bytes = await readFile(resolve(FIX, 'slide-transition-sound.pptx'))
const sha256 = createHash('sha256').update(bytes).digest('hex')
const zip = await JSZip.loadAsync(bytes)

const sha = (b) => createHash('sha256').update(b).digest('hex')
const audioBytes = await zip.file('ppt/media/audio1.wav').async('nodebuffer')

const slides = []
for (let i = 1; i <= 3; i++) {
	const xml = await zip.file(`ppt/slides/slide${i}.xml`).async('string')
	const wrapped = xml.includes('<mc:AlternateContent')
	const transitionXml = xml.match(
		/<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>|<p:transition[\s\S]*?<\/p:transition>/
	)[0]
	const sndAcXml = (transitionXml.match(/<p:sndAc>[\s\S]*?<\/p:sndAc>/) || [null])[0]
	const relsXml = await zip.file(`ppt/slides/_rels/slide${i}.xml.rels`).async('string')
	// audio rel for this slide (if any)
	const relM = relsXml.match(/<Relationship Id="(rId\d+)" Type="([^"]*\/audio)" Target="([^"]*)"\/>/)
	const audioRel = relM ? { id: relM[1], type: relM[2], target: relM[3] } : null
	const sndM = sndAcXml ? sndAcXml.match(/<p:snd r:embed="(rId\d+)" name="([^"]*)"\/>/) : null
	const loop = sndAcXml ? /<p:stSnd loop="1">/.test(sndAcXml) : false
	const form = sndAcXml ? (sndAcXml.includes('<p:endSnd/>') ? 'endSnd' : 'stSnd') : null
	slides.push({
		slide: i,
		part: `slide${i}.xml`,
		wrapped,
		transitionXml,
		soundRels: {
			sndAcXml,
			form,
			loop,
			sndEmbedRid: sndM ? sndM[1] : null,
			sndName: sndM ? sndM[2] : null,
			audioRel,
		},
	})
}

const oracle = {
	deck: 'slide-transition-sound',
	schema: 'slide-transition-oracle@1',
	application: 'Microsoft Office PowerPoint',
	appVersion: '16.0000',
	sha256,
	notes:
		'PowerPoint-authored transition-sound oracle for docs/animations-and-transitions.md (Phase 2 capability C: transition sounds, p:sndAc). Three blank 16:9 slides, each a fade transition: slide 1 an embedded start sound (p:sndAc/p:stSnd/p:snd), slide 2 the same embedded sound looped until next (p:stSnd loop="1"), slide 3 the stop-previous-sound form (p:sndAc/p:endSnd, no rel, no media part). Structure (confirmed against CT_TransitionSoundAction): p:transition > p:sndAc (after the transition-type element, before extLst) > choice(p:stSnd[@loop] | p:endSnd); p:stSnd > p:snd (CT_EmbeddedWAVAudioFile) with required r:embed + optional name. The sound is referenced by an ECMA audio relationship (.../2006/relationships/audio) to an embedded WAV part (ppt/media/audioN.wav), content type via a single <Default Extension="wav" ContentType="audio/x-wav"/> (not an Override). PowerPoint dedups identical sound bytes: slides 1 and 2 share one media part (audio1.wav) via the same rId. All three transitions also carry PowerPoint\'s default exact fade duration, so each is emitted in the mc:AlternateContent (p14:dur="2000") form; the sndAc lives inside both the p14 Choice and the base Fallback transition, identical in each. BUILT-IN sounds: probed via SoundEffect.Name="Applause" — PowerPoint resolves the built-in to applause.wav and embeds it IDENTICALLY to a custom import (same audio rel, same audio/x-wav Default, same sndAc/stSnd/snd; only the embedded bytes + name differ), so the writer needs no separate built-in path. The committed fixture uses only a self-generated WAV to stay license-clean of Microsoft\'s bundled audio; the built-in equivalence is recorded here rather than embedded.',
	slides,
	contentTypes: { wav: 'audio/x-wav' },
	mediaParts: [
		{ part: 'ppt/media/audio1.wav', bytes: audioBytes.length, sha256: sha(audioBytes), sharedBySlides: [1, 2] },
	],
	soundProvenance:
		'Embedded WAV (ppt/media/audio1.wav) is a tiny self-generated 16-bit PCM mono 8kHz sine (authoring/assets/ding.wav), not a Microsoft asset — license-clean.',
}

await writeFile(resolve(FIX, 'slide-transition-sound.oracle.json'), JSON.stringify(oracle, null, '\t') + '\n')
console.log('sha256', sha256)
console.log('audio1.wav bytes', audioBytes.length, 'sha', sha(audioBytes))
for (const s of slides)
	console.log(
		`slide ${s.slide}: form=${s.soundRels.form} loop=${s.soundRels.loop} rel=${s.soundRels.audioRel?.target ?? '(none)'}`
	)
