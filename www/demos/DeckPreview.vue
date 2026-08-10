<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import {
	buildDeckBytes,
	DECK,
	downloadDeck,
	failureMessage,
	previewDeck,
	slideList,
	summarizeNotes,
} from './deck-preview.ts'

// `preview` runs on mount; `download` runs on the button. Two states rather than one
// because either can fail on its own, and a failed render must not be reported as a
// failed build.
const preview = ref({ status: 'rendering', deck: null, error: '' })
const download = ref({ status: 'idle', error: '' })

// VitePress pre-renders this page, so the markup below — including an enabled button —
// exists in the served HTML before any of it is wired up. A click in that window does
// nothing at all and reports nothing, which is a bad half-second for a visitor and an
// unfalsifiable failure for the browser lane, whose first act is to click that button.
// Gating on mount makes "not ready yet" a state the page can show and a test can wait for.
const ready = ref(false)

const frame = ref(null)
const slide = ref(1)

const notes = computed(() => (preview.value.deck ? summarizeNotes(preview.value.deck.notes) : []))
const slideCount = computed(() => preview.value.deck?.slideCount ?? 0)

async function render() {
	preview.value = { status: 'rendering', deck: null, error: '' }
	try {
		const deck = await previewDeck(await buildDeckBytes())
		preview.value = { status: 'ready', deck, error: '' }
		slide.value = 1
	} catch (error) {
		preview.value = { status: 'failed', deck: null, error: failureMessage(error) }
	}
}

async function saveDeck() {
	download.value = { status: 'saving', error: '' }
	try {
		await downloadDeck()
		download.value = { status: 'saved', error: '' }
	} catch (error) {
		download.value = { status: 'failed', error: failureMessage(error) }
	}
}

/** The rendered document is same-origin (`srcdoc`), so its slides are reachable directly. */
function frameDocument() {
	return frame.value?.contentDocument ?? null
}

function goTo(number) {
	const target = Math.min(Math.max(number, 1), slideCount.value)
	slide.value = target

	const view = frame.value?.contentWindow
	const section = frameDocument()?.querySelector(`[data-pxh-slide="${target}"]`)
	if (!view || !section) return
	// Scroll the frame's own document, rather than `section.scrollIntoView()`. That call
	// would also scroll *this* page to bring the frame into view, which pulls these very
	// buttons off screen — pressing Next would move the control you pressed it with.
	view.scrollTo({ top: section.getBoundingClientRect().top + view.scrollY })
}

/**
 * Follow the iframe's own scrolling, so dragging its scrollbar moves the counter too.
 * Whichever slide's top edge is nearest the top of the frame is the one being read.
 */
function syncFromScroll() {
	const doc = frameDocument()
	if (!doc) return
	let nearest = slide.value
	let best = Infinity
	for (const section of doc.querySelectorAll('[data-pxh-slide]')) {
		const distance = Math.abs(section.getBoundingClientRect().top)
		if (distance < best) {
			best = distance
			nearest = Number(section.getAttribute('data-pxh-slide'))
		}
	}
	slide.value = nearest
}

let scrollTarget = null

function detachScroll() {
	scrollTarget?.removeEventListener('scroll', syncFromScroll)
	scrollTarget = null
}

function onFrameLoad() {
	detachScroll()
	scrollTarget = frame.value?.contentWindow ?? null
	scrollTarget?.addEventListener('scroll', syncFromScroll, { passive: true })
}

watch(() => preview.value.status, detachScroll)
onBeforeUnmount(detachScroll)

onMounted(() => {
	ready.value = true
	render()
})
</script>

<template>
	<section class="deck-preview">
		<header class="deck-preview__head">
			<div>
				<h2 class="deck-preview__title">{{ DECK.title }}</h2>
				<p class="deck-preview__blurb">{{ DECK.description }}</p>
			</div>
			<div role="group" aria-label="Download" class="deck-preview__download">
				<button type="button" :disabled="!ready || download.status === 'saving'" @click="saveDeck">
					{{ download.status === 'saving' ? 'Building…' : `Build ${DECK.fileName}` }}
				</button>
				<p v-if="download.status === 'saved'" role="status" class="deck-preview__note">
					Built <code>{{ DECK.fileName }}</code> — check your downloads.
				</p>
				<p v-else-if="download.status === 'failed'" role="alert" class="deck-preview__note deck-preview__note--bad">
					{{ download.error }}
				</p>
			</div>
		</header>

		<div class="deck-preview__bar">
			<template v-if="preview.status === 'ready'">
				<button type="button" :disabled="slide <= 1" aria-label="Previous slide" @click="goTo(slide - 1)">←</button>
				<span aria-live="polite">Slide {{ slide }} of {{ slideCount }}</span>
				<button type="button" :disabled="slide >= slideCount" aria-label="Next slide" @click="goTo(slide + 1)">→</button>
			</template>
			<span v-else-if="preview.status === 'rendering'">Building the deck and rendering it…</span>
			<button v-else type="button" @click="render">Try again</button>
		</div>

		<p v-if="preview.status === 'failed'" role="alert" class="deck-preview__failure">
			The preview could not be rendered: {{ preview.error }}
		</p>

		<iframe
			v-if="preview.status === 'ready'"
			ref="frame"
			class="deck-preview__frame"
			title="Rendered slides"
			:srcdoc="preview.deck.html"
			@load="onFrameLoad"
		/>
		<div v-else class="deck-preview__frame deck-preview__frame--empty" />

		<p v-if="preview.deck?.warnings.length" role="alert" class="deck-preview__failure">
			The renderer reported {{ preview.deck.warnings.length }} warning(s):
			{{ preview.deck.warnings.join('; ') }}
		</p>

		<details v-if="notes.length" class="deck-preview__ledger">
			<summary>All {{ notes.length }} declared difference(s), across the deck</summary>
			<p>
				The renderer prints these beside the slide they apply to; this is the same set
				gathered up, in its vocabulary rather than one invented here. They describe what a
				reader would and would not carry back out of this package — a construct that is
				<em>carried</em> untouched raises no note at all.
			</p>
			<dl>
				<template v-for="note in notes" :key="note.key">
					<dt>
						<code>{{ note.construct }}</code> — {{ note.disposition }} ({{ note.cause }})
						<span class="deck-preview__slides">slide{{ note.slides.length === 1 ? '' : 's' }} {{ slideList(note.slides) }}</span>
					</dt>
					<dd>{{ note.detail }}</dd>
				</template>
			</dl>
		</details>
	</section>
</template>
