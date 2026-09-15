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
const naturalWidth = ref('')

async function draw() {
	try {
		svg.value = await renderMermaid(decodeGraph(props.graph), isDark.value)
		// Mermaid sizes its SVG to `width="100%"`, so a wide graph shrinks to the column until
		// its labels cannot be read, and the container's horizontal scroll never engages. The
		// width the graph was laid out at gives the stylesheet a floor to scroll against.
		const viewBoxWidth = Number(svg.value.match(/viewBox="[-\d.]+ [-\d.]+ ([\d.]+)/)?.[1])
		naturalWidth.value = viewBoxWidth > 0 ? `${Math.round(viewBoxWidth)}px` : ''
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
	<div class="mermaid-diagram" :style="naturalWidth ? { '--mermaid-natural-width': naturalWidth } : undefined">
		<div v-if="svg" v-html="svg" />
		<p v-else-if="error" role="alert" class="mermaid-diagram__failure">
			The diagram could not be drawn: {{ error }}
		</p>
	</div>
</template>
