<script setup>
import { onMounted, ref, watch } from 'vue'
import { SLIDE_FRAME_CSS } from './deck-preview.ts'

// One slide in a shadow root. The root keeps the site's stylesheet out of the slide's
// text and scopes the slide's gradient ids to the slide; `splitDeck` says why both matter.
const props = defineProps({
	markup: { type: String, required: true },
	styles: { type: String, required: true },
})

const host = ref(null)
let root = null

function paint() {
	if (root) root.innerHTML = `<style>${props.styles}\n${SLIDE_FRAME_CSS}</style>${props.markup}`
}

onMounted(() => {
	root = host.value.attachShadow({ mode: 'open' })
	paint()
})
watch(() => [props.markup, props.styles], paint)
</script>

<template>
	<div ref="host" class="slide-frame" />
</template>
