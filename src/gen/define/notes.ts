/**
 * ts-pptx: Notes Definition
 *
 * Normalizes plain or rich `addNotes()` input to a `TextProps[]` run list and pushes a
 * `notes`-type slide object; the notes-slide XML is emitted later by `gen/slide/notes.ts`.
 */
import { SlideObjectType } from '../../enums.js'
import type { NotesProps, TextProps } from '../../types/index.js'
import type { PresSlideInternal } from '../../types/internal.js'

/**
 * Adds Notes to a slide.
 * @param {PresSlideInternal} `target` slide object
 * @param {string | NotesProps | NotesProps[]} `notes` plain text, or rich runs (inline formatting / hyperlinks)
 */
export function addNotesDefinition(target: PresSlideInternal, notes: string | NotesProps | NotesProps[]): void {
	// Normalize all input forms to a TextProps[] run list so the notes-slide serializer
	// (which reuses the standard text-run generator) can handle plain and rich notes uniformly.
	const runs: TextProps[] =
		typeof notes === 'string'
			? [{ text: notes }]
			: (Array.isArray(notes) ? notes : [notes]).map((run) => ({ text: run.text, options: run.options }))

	target._slideObjects.push({
		_type: SlideObjectType.notes,
		text: runs,
	})
}
