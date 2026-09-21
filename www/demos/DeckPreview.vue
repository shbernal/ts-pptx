<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import {
	buildDeckBytes,
	counted,
	DECK,
	downloadDeck,
	failureMessage,
	previewDeck,
	slideList,
	summarizeNotes,
} from './deck-preview.ts'
import SlideFrame from './SlideFrame.vue'

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

const slide = ref(1)
const stage = ref(null)
const strip = ref(null)
const fullscreen = ref(false)

const deck = computed(() => preview.value.deck)
const slideCount = computed(() => deck.value?.slides.length ?? 0)
const current = computed(() => deck.value?.slides[slide.value - 1] ?? null)
const differences = computed(() => (deck.value ? summarizeNotes(deck.value.fidelity) : []))
const aspect = computed(() => ({ '--deck-aspect': String(deck.value?.aspectRatio ?? 16 / 9) }))

async function render() {
	preview.value = { status: 'rendering', deck: null, error: '' }
	try {
		const built = await previewDeck(await buildDeckBytes())
		slide.value = 1
		preview.value = { status: 'ready', deck: built, error: '' }
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

function goTo(number) {
	if (!slideCount.value) return
	slide.value = Math.min(Math.max(number, 1), slideCount.value)
}

// Keys anywhere in the player, the way a slide show takes them. Only unmodified keys, so
// Alt+Left still goes back a page. Up, Down and Space are left alone: they scroll the
// page, and a thumbnail that has focus still needs Space to press it.
function onKey(event) {
	if (event.altKey || event.ctrlKey || event.metaKey) return
	const target = {
		ArrowLeft: slide.value - 1,
		PageUp: slide.value - 1,
		ArrowRight: slide.value + 1,
		PageDown: slide.value + 1,
		Home: 1,
		End: slideCount.value,
	}[event.key]
	if (target !== undefined) {
		event.preventDefault()
		goTo(target)
	} else if (event.key === 'f') {
		event.preventDefault()
		toggleFullscreen()
	}
}

function toggleFullscreen() {
	if (document.fullscreenElement) document.exitFullscreen()
	else stage.value?.requestFullscreen?.()
}

function onFullscreenChange() {
	fullscreen.value = document.fullscreenElement === stage.value
	if (fullscreen.value) stage.value?.focus()
}

// Keep the current thumbnail in view. The strip is scrolled directly rather than through
// `scrollIntoView()`, which would also scroll the page and pull the stage away from the
// reader pressing the arrow keys.
watch(slide, async (number) => {
	await nextTick()
	const list = strip.value
	const thumb = list?.querySelector(`[data-slide="${number}"]`)
	if (!list || !thumb) return
	const left = thumb.offsetLeft - (list.clientWidth - thumb.offsetWidth) / 2
	list.scrollTo({ left, behavior: 'smooth' })
})

onMounted(() => {
	ready.value = true
	document.addEventListener('fullscreenchange', onFullscreenChange)
	render()
})
onBeforeUnmount(() => document.removeEventListener('fullscreenchange', onFullscreenChange))
</script>

<template>
	<section class="deck-viewer" :style="aspect">
		<header class="deck-viewer__head">
			<div class="deck-viewer__heading">
				<p class="deck-viewer__eyebrow">
					<span class="deck-viewer__live" aria-hidden="true" />
					Built in this tab
				</p>
				<h2 class="deck-viewer__title">{{ DECK.title }}</h2>
				<p class="deck-viewer__blurb">{{ DECK.description }}</p>
			</div>
			<div role="group" aria-label="Download" class="deck-viewer__download">
				<button type="button" class="deck-viewer__build" :disabled="!ready || download.status === 'saving'" @click="saveDeck">
					<svg viewBox="0 0 20 20" aria-hidden="true">
						<path d="M10 3v9m0 0-3.5-3.5M10 12l3.5-3.5M4 15.5h12" />
					</svg>
					{{ download.status === 'saving' ? 'Building…' : 'Build the .pptx' }}
				</button>
				<p v-if="download.status === 'saved'" role="status" class="deck-viewer__note">
					Built <code>{{ DECK.fileName }}</code>. Check your downloads.
				</p>
				<p v-else-if="download.status === 'failed'" role="alert" class="deck-viewer__note deck-viewer__note--bad">
					{{ download.error }}
				</p>
				<p v-else class="deck-viewer__note">{{ DECK.fileName }}</p>
			</div>
		</header>

		<div class="deck-viewer__player" @keydown="onKey">
			<div
				ref="stage"
				class="deck-viewer__stage"
				:class="{ 'is-fullscreen': fullscreen }"
				tabindex="0"
				role="region"
				aria-roledescription="slide viewer"
				:aria-label="deck ? `Slide ${slide} of ${slideCount}` : 'Slide viewer'"
			>
				<div class="deck-viewer__canvas">
					<SlideFrame v-if="current" :key="current.number" :markup="current.markup" :styles="deck.styles" />
					<div v-else-if="preview.status === 'rendering'" class="deck-viewer__placeholder">
						<span class="deck-viewer__spinner" aria-hidden="true" />
						<span>Building the deck and rendering it…</span>
					</div>
					<div v-else class="deck-viewer__placeholder" role="alert">
						<span>The preview could not be rendered: {{ preview.error }}</span>
						<button type="button" class="deck-viewer__retry" @click="render">Try again</button>
					</div>
				</div>
				<template v-if="deck">
					<button
						type="button"
						class="deck-viewer__arrow deck-viewer__arrow--prev"
						:disabled="slide <= 1"
						aria-label="Previous slide"
						@click="goTo(slide - 1)"
					>
						<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M12.5 4.5 7 10l5.5 5.5" /></svg>
					</button>
					<button
						type="button"
						class="deck-viewer__arrow deck-viewer__arrow--next"
						:disabled="slide >= slideCount"
						aria-label="Next slide"
						@click="goTo(slide + 1)"
					>
						<svg viewBox="0 0 20 20" aria-hidden="true"><path d="M7.5 4.5 13 10l-5.5 5.5" /></svg>
					</button>
				</template>
			</div>

			<div v-if="deck" class="deck-viewer__toolbar">
				<span class="deck-viewer__counter" aria-live="polite">
					<strong>{{ slide }}</strong> / {{ slideCount }}
				</span>
				<span class="deck-viewer__hint">Use ← → to move between slides</span>
				<button type="button" class="deck-viewer__tool" @click="toggleFullscreen">
					<svg viewBox="0 0 20 20" aria-hidden="true">
						<path v-if="fullscreen" d="M8 3v5H3m9-5v5h5M8 17v-5H3m9 5v-5h5" />
						<path v-else d="M3 8V3h5m4 0h5v5M3 12v5h5m4 0h5v-5" />
					</svg>
					{{ fullscreen ? 'Exit full screen' : 'Full screen' }}
				</button>
			</div>

			<ol v-if="deck" ref="strip" class="deck-viewer__strip" aria-label="Slides">
				<li v-for="item in deck.slides" :key="item.number">
					<button
						type="button"
						class="deck-viewer__thumb"
						:class="{ 'is-current': item.number === slide }"
						:data-slide="item.number"
						:aria-label="`Slide ${item.number}`"
						:aria-current="item.number === slide ? 'true' : undefined"
						@click="goTo(item.number)"
					>
						<SlideFrame :markup="item.markup" :styles="deck.styles" aria-hidden="true" />
						<span class="deck-viewer__thumb-number">{{ item.number }}</span>
					</button>
				</li>
			</ol>
		</div>

		<div v-if="current" class="deck-viewer__notes">
			<p class="deck-viewer__label">Speaker notes</p>
			<p v-for="(line, index) in current.notes" :key="index">{{ line }}</p>
			<p v-if="!current.notes.length" class="deck-viewer__muted">This slide has no notes.</p>
		</div>

		<p v-if="deck?.warnings.length" role="alert" class="deck-viewer__failure">
			The renderer reported {{ counted(deck.warnings.length, 'warning') }}:
			{{ deck.warnings.join('; ') }}
		</p>

		<details v-if="differences.length" class="deck-viewer__ledger">
			<summary>
				<span>What the renderer says it could not carry back</span>
				<span class="deck-viewer__count">{{ counted(differences.length, 'difference') }}</span>
			</summary>
			<p class="deck-viewer__muted">
				The renderer declares these per slide; this is the same set gathered up, in its
				vocabulary rather than one invented here. They describe what a reader would and would
				not carry back out of this package: a construct that is <em>carried</em> untouched
				raises no note at all.
			</p>
			<ul class="deck-viewer__differences">
				<li v-for="note in differences" :key="note.key">
					<div class="deck-viewer__difference-head">
						<code>{{ note.construct }}</code>
						<span class="deck-viewer__pill">{{ note.disposition }}</span>
						<span class="deck-viewer__pill deck-viewer__pill--quiet">{{ note.cause }}</span>
						<span class="deck-viewer__slides">
							slide{{ note.slides.length === 1 ? '' : 's' }} {{ slideList(note.slides) }}
						</span>
					</div>
					<p>{{ note.detail }}</p>
				</li>
			</ul>
		</details>
	</section>
</template>
