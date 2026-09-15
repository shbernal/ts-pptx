<script setup>
// oxlint-disable-next-line import/named -- `vitepress` resolves to its Node entry here; the site build aliases it to the client, which exports `useData`.
import { useData } from 'vitepress'
import { onMounted, ref, watch } from 'vue'
import { decodeGraph } from './fence.ts'
import { renderMermaid } from './render.ts'

const props = defineProps({
	/** The graph source, URI-encoded by the fence rule. */
	graph: { type: String, required: true },
})

const { isDark } = useData()
const svg = ref('')
const error = ref('')

async function draw() {
	try {
		svg.value = await renderMermaid(decodeGraph(props.graph), isDark.value)
		error.value = ''
	} catch (caught) {
		svg.value = ''
		error.value = caught instanceof Error ? caught.message : String(caught)
	}
}

// Mermaid lays a diagram out against a live DOM, so pre-rendering leaves the container empty
// and the diagram is drawn on mount, then redrawn whenever the site's colour mode flips.
onMounted(() => {
	draw()
	watch(isDark, draw)
})
</script>

<template>
	<div class="mermaid-diagram">
		<div v-if="svg" v-html="svg" />
		<p v-else-if="error" role="alert" class="mermaid-diagram__failure">
			The diagram could not be drawn: {{ error }}
		</p>
	</div>
</template>
