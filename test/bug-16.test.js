'use strict'

const { build, readEntry, assert } = require('./helpers')

// 1x1 PNG (red pixel)
const PNG_DATA = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='
// 1x1 JPEG
const JPG_DATA = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AL+AB//Z'

module.exports = [
	{
		name: 'B16: PNG-only deck emits png Default but not jpeg/jpg/svg/gif/m4v/mp4/vml/xlsx',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/Extension="xml"/.test(xml), 'expected xml Default; got: ' + xml)
			assert(/Extension="rels"/.test(xml), 'expected rels Default; got: ' + xml)
			assert(/Extension="png"/.test(xml), 'expected png Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="jpeg"/.test(xml), 'unexpected jpeg Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="jpg"/.test(xml), 'unexpected jpg Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="svg"/.test(xml), 'unexpected svg Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="gif"/.test(xml), 'unexpected gif Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="m4v"/.test(xml), 'unexpected m4v Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="mp4"/.test(xml), 'unexpected mp4 Default for PNG-only deck; got: ' + xml)
			assert(!/Extension="vml"/.test(xml), 'unexpected vml Default; got: ' + xml)
			assert(!/Extension="xlsx"/.test(xml), 'unexpected xlsx Default for chart-free deck; got: ' + xml)
		}
	},
	{
		name: 'B16: empty deck emits only xml + rels Defaults (no media defaults)',
		fn: async () => {
			const { zip } = await build(p => {
				p.addSlide()
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/Extension="xml"/.test(xml), 'expected xml Default; got: ' + xml)
			assert(/Extension="rels"/.test(xml), 'expected rels Default; got: ' + xml)
			assert(!/Extension="png"/.test(xml), 'unexpected png Default in empty deck; got: ' + xml)
			assert(!/Extension="jpeg"/.test(xml), 'unexpected jpeg Default in empty deck; got: ' + xml)
			assert(!/Extension="jpg"/.test(xml), 'unexpected jpg Default in empty deck; got: ' + xml)
			assert(!/Extension="svg"/.test(xml), 'unexpected svg Default in empty deck; got: ' + xml)
			assert(!/Extension="gif"/.test(xml), 'unexpected gif Default in empty deck; got: ' + xml)
			assert(!/Extension="m4v"/.test(xml), 'unexpected m4v Default in empty deck; got: ' + xml)
			assert(!/Extension="mp4"/.test(xml), 'unexpected mp4 Default in empty deck; got: ' + xml)
			assert(!/Extension="vml"/.test(xml), 'unexpected vml Default in empty deck; got: ' + xml)
			assert(!/Extension="xlsx"/.test(xml), 'unexpected xlsx Default in empty deck; got: ' + xml)
		}
	},
	{
		name: 'B16: PNG + JPEG deck emits both png and jpeg/jpg Defaults; gif/svg absent',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ data: JPG_DATA, x: 3, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/Extension="png"/.test(xml), 'expected png Default; got: ' + xml)
			// jpeg images are recorded with extn "jpg" and type "image/jpg" in this codebase
			assert(/Extension="jpg"/.test(xml) || /Extension="jpeg"/.test(xml),
				'expected jpg or jpeg Default for JPEG image; got: ' + xml)
			assert(!/Extension="gif"/.test(xml), 'unexpected gif Default; got: ' + xml)
			assert(!/Extension="svg"/.test(xml), 'unexpected svg Default; got: ' + xml)
			assert(!/Extension="m4v"/.test(xml), 'unexpected m4v Default; got: ' + xml)
			assert(!/Extension="mp4"/.test(xml), 'unexpected mp4 Default; got: ' + xml)
			assert(!/Extension="vml"/.test(xml), 'unexpected vml Default; got: ' + xml)
			assert(!/Extension="xlsx"/.test(xml), 'unexpected xlsx Default; got: ' + xml)
		}
	},
	{
		name: 'B16: chart deck emits xlsx Default',
		fn: async () => {
			const { pres, zip } = await build(p => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [
					{ name: 'series1', labels: ['a', 'b'], values: [1, 2] }
				], { x: 1, y: 1, w: 4, h: 3 })
			})
			void pres
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/Extension="xlsx"/.test(xml), 'expected xlsx Default for chart deck; got: ' + xml)
		}
	},
	{
		name: 'B16: regression - structural Override entries still emitted',
		fn: async () => {
			const { zip } = await build(p => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assert(/PartName="\/ppt\/presentation\.xml"/.test(xml), 'expected presentation.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/slideMasters\/slideMaster1\.xml"/.test(xml), 'expected slideMaster1.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/theme\/theme1\.xml"/.test(xml), 'expected theme1.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/theme\/theme2\.xml"/.test(xml), 'expected theme2.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/slides\/slide1\.xml"/.test(xml), 'expected slide1.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/slideLayouts\/slideLayout1\.xml"/.test(xml), 'expected slideLayout1.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/notesMasters\/notesMaster1\.xml"/.test(xml), 'expected notesMaster1.xml Override; got: ' + xml)
			assert(/PartName="\/ppt\/notesSlides\/notesSlide1\.xml"/.test(xml), 'expected notesSlide1.xml Override; got: ' + xml)
		}
	}
]
