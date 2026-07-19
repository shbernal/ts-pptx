import { describe, expect, test } from 'vitest'
import { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from '../../src/gen/slide/comments.ts'

// Characterization tests for comments XML that the byte-identity harness CANNOT see — the demo
// deck has no comments, so `<p:cmAuthor>`/`<p:cm>` carry ZERO baseline parts. schema-cases.js
// covers the happy path (author/idx numbering) through the public API; these pin the byte-level
// details the migration to el()/voidEl() must preserve: attribute order, the optional `dt`
// attribute, and metacharacter escaping (already correct pre-migration — characterized, not fixed).

const author = (over = {}) => ({ id: 0, name: 'Ada Lovelace', initials: 'AL', lastIdx: 1, clrIdx: 0, ...over })
const comment = (over = {}) => ({ author: 'Ada Lovelace', initials: 'AL', text: 'x', x: 1, y: 0.5, ...over })

describe('makeXmlCommentAuthors', () => {
	test('cmAuthor attribute order: id, name, initials, lastIdx, clrIdx', () => {
		const xml = makeXmlCommentAuthors([author()])
		expect(xml).toContain('<p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/>')
	})

	test('name and initials are escaped', () => {
		const xml = makeXmlCommentAuthors([author({ name: 'Q&A <Team>', initials: 'A&B' })])
		expect(xml).toContain('name="Q&amp;A &lt;Team&gt;"')
		expect(xml).toContain('initials="A&amp;B"')
	})

	test('empty author list emits an empty cmAuthorLst', () => {
		const xml = makeXmlCommentAuthors([])
		expect(xml).toContain('<p:cmAuthorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ')
		expect(xml).toContain('></p:cmAuthorLst>')
	})
})

describe('makeXmlComments', () => {
	test('cm attribute order: authorId, dt, idx (dt present)', () => {
		const c = comment({ date: '2026-06-24T10:00:00Z' })
		const meta = new Map([[c, { authorId: 0, idx: 1 }]])
		const xml = makeXmlComments({ _comments: [c] }, meta)
		expect(xml).toContain('<p:cm authorId="0" dt="2026-06-24T10:00:00Z" idx="1">')
	})

	test('dt attribute is omitted (not emitted empty) when date is unset', () => {
		const c = comment()
		const meta = new Map([[c, { authorId: 0, idx: 1 }]])
		const xml = makeXmlComments({ _comments: [c] }, meta)
		expect(xml).toContain('<p:cm authorId="0" idx="1">')
		expect(xml).not.toContain('dt=')
	})

	test('child order is p:pos then p:text; pos coords are rounded EMU', () => {
		const c = comment({ x: 1, y: 0.5, text: 'hello' })
		const meta = new Map([[c, { authorId: 0, idx: 1 }]])
		const xml = makeXmlComments({ _comments: [c] }, meta)
		expect(xml).toContain('<p:pos x="914400" y="457200"/><p:text>hello</p:text>')
	})

	test('text and dt are escaped', () => {
		const c = comment({ text: 'A & B < C', date: 'Q&A' })
		const meta = new Map([[c, { authorId: 0, idx: 1 }]])
		const xml = makeXmlComments({ _comments: [c] }, meta)
		expect(xml).toContain('<p:text>A &amp; B &lt; C</p:text>')
		expect(xml).toContain('dt="Q&amp;A"')
	})

	test('a comment absent from meta is silently skipped (defensive branch)', () => {
		const c = comment()
		const xml = makeXmlComments({ _comments: [c] }, new Map())
		expect(xml).not.toContain('<p:cm ')
		expect(xml).toContain('<p:cmLst')
	})

	test('no comments on the slide emits an empty cmLst', () => {
		const xml = makeXmlComments({ _comments: [] }, new Map())
		expect(xml).toContain('></p:cmLst>')
	})
})

describe('resolveCommentAuthors', () => {
	test('per-author idx numbering across slides in slide order', () => {
		const c1 = comment({ author: 'Ada Lovelace', initials: 'AL' })
		const c2 = comment({ author: 'Alan Turing', initials: 'AT' })
		const c3 = comment({ author: 'Ada Lovelace', initials: 'AL' })
		const slides = [{ _comments: [c1, c2] }, { _comments: [c3] }]
		const { authors, meta } = resolveCommentAuthors(slides)
		expect(authors.map((a) => [a.name, a.id, a.clrIdx, a.lastIdx])).toEqual([
			['Ada Lovelace', 0, 0, 2],
			['Alan Turing', 1, 1, 1],
		])
		expect(meta.get(c1)).toEqual({ authorId: 0, idx: 1 })
		expect(meta.get(c2)).toEqual({ authorId: 1, idx: 1 })
		expect(meta.get(c3)).toEqual({ authorId: 0, idx: 2 })
	})
})
