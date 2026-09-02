import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.js'

// What a run inherits from its shape, and what it does not.
//
// `genXmlTextRun` is handed the run's own options bag and never the shape's, so anything the
// caller stated once on the shape has to be copied down onto each run to reach `<a:rPr>`. That
// copy used to be `Object.entries(opts)` guarded by `!textOptions[key]`, which answers "is the
// run's value falsy?" where the question is "did the run say anything?". An explicit `false`,
// `0` or `''` therefore counted as silence and was replaced by the value it was written to
// override — the same defect the two paragraph margins twenty lines above already carried a
// comment about.
//
// The same loop is why one run's font size leaked into the next: the `endParaRPr` bookkeeping
// wrote the first sized run's size back onto the SHAPE's bag, and the copy then handed it to
// every later run and every later paragraph.

/** The `<a:rPr>` of each run in the slide's first text body, in document order. */
function runProps(xml) {
	return xml.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/g) || []
}

/** Every `<a:endParaRPr>` in document order. */
function endParaProps(xml) {
	return xml.match(/<a:endParaRPr[^>]*\/>/g) || []
}

async function slideFor(text, opts) {
	const { zip } = await build((p) => {
		p.addSlide().addText(text, { x: 1, y: 1, w: 6, h: 3, ...opts })
	})
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

defineRegressionSuite('Run options inherited from the shape', [
	{
		name: "a run's explicit `bold: false` is not overwritten by the shape's `bold: true`",
		fn: async () => {
			const xml = await slideFor([{ text: 'a', options: { bold: false } }], { bold: true })
			const [rPr] = runProps(xml)
			assert(!/\bb="1"/.test(rPr), 'the run said not bold; got: ' + rPr)
		},
	},
	{
		name: "a run's explicit `transparency: 0` is not overwritten by the shape's",
		fn: async () => {
			const xml = await slideFor([{ text: 'a', options: { color: 'FF0000', transparency: 0 } }], { transparency: 50 })
			const [rPr] = runProps(xml)
			assert(!rPr.includes('<a:alpha'), 'a fully opaque run emits no alpha; got: ' + rPr)
		},
	},
	{
		name: 'the shape still fills in what a run leaves unstated',
		fn: async () => {
			// The inheritance itself is the point of the copy; only its guard changed.
			const xml = await slideFor([{ text: 'a' }], { bold: true, color: '112233', fontSize: 24 })
			const [rPr] = runProps(xml)
			assert(/\bb="1"/.test(rPr) && /\bsz="2400"/.test(rPr), 'bold and size come from the shape; got: ' + rPr)
			assert(rPr.includes('112233'), 'and so does the colour; got: ' + rPr)
		},
	},
	{
		name: "one run's fontSize does not leak onto the runs after it",
		fn: async () => {
			const xml = await slideFor([{ text: 'big', options: { fontSize: 40 } }, { text: 'normal' }], {})
			const [first, second] = runProps(xml)
			assert(/\bsz="4000"/.test(first), 'the sized run keeps its size; got: ' + first)
			assert(!/\bsz=/.test(second), 'the run after it states none; got: ' + second)
		},
	},
	{
		name: "one paragraph's fontSize does not leak into a later paragraph's endParaRPr",
		fn: async () => {
			const xml = await slideFor(
				[
					{ text: 'big', options: { fontSize: 40, breakLine: true } },
					{ text: 'small', options: { fontSize: 10 } },
				],
				{}
			)
			const ends = endParaProps(xml)
			assertEqual(ends.length, 2, `two paragraphs, two endParaRPr; got ${JSON.stringify(ends)}`)
			assert(/\bsz="4000"/.test(ends[0]), 'the first closes on 40pt; got: ' + ends[0])
			assert(/\bsz="1000"/.test(ends[1]), 'and the second on its own 10pt; got: ' + ends[1])
		},
	},
	{
		// `valign` reached `a:bodyPr/@anchor` three ways, and one of them (`gen/slide/object.ts`)
		// let any unrecognised string through verbatim into the attribute.
		// `ST_TextAnchoringType` is an enumeration, so that is a repair prompt.
		name: 'an unrecognised valign warns and takes the default anchor, not the caller-s string',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(() => slideFor('hi', { valign: 'sideways' }))
			const bodyPr = (/<a:bodyPr[^>]*>/.exec(xml) ?? [''])[0]
			assert(!bodyPr.includes('sideways'), 'the string never reaches the attribute; got: ' + bodyPr)
			assert(bodyPr.includes('anchor="ctr"'), 'the text definer-s own default applies; got: ' + bodyPr)
			assert(codes.includes('text/invalid-valign'), 'the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'the valign spellings the three definers accepted still resolve',
		fn: async () => {
			for (const [valign, anchor] of [
				['top', 't'],
				['middle', 'ctr'],
				['bottom', 'b'],
				['ctr', 'ctr'],
				['btm', 'b'],
				['MIDDLE', 'ctr'],
			]) {
				const xml = await slideFor('hi', { valign })
				assert(
					new RegExp(`<a:bodyPr[^>]*\\banchor="${anchor}"`).test(xml),
					`valign "${valign}" anchors ${anchor}; got: ` + (/<a:bodyPr[^>]*>/.exec(xml) ?? [''])[0]
				)
			}
		},
	},
])
