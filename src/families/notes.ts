/**
 * ts-pptx: the speaker-notes construct family
 *
 * `addNotes` and the notes-slide parts it puts in the package. Cheap, and a deck without speaker
 * notes is a poor default, so every tier carries it -- which is a reason to compose it, not a
 * reason to wire it directly.
 */

import { addNotesDefinition } from '../gen/define/notes.js'
import type { ConstructFamily } from './shared.js'

export const notesFamily: ConstructFamily = {
	name: 'notes',
	authors: {
		addNotes(slide, notes) {
			addNotesDefinition(slide, notes)
		},
	},
}
