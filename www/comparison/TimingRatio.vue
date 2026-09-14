<script setup>
import { data } from './comparison.data.ts'
import { formatRatio, formatTick, ratioPosition, ratioSpan, ratioSummary } from './comparison.ts'

const at = (value) => `${ratioPosition(value, data.domain)}%`
const span = (row) => {
	const { left, width } = ratioSpan(row, data.domain)
	return { left: `${left}%`, width: `${width}%` }
}
</script>

<template>
	<figure class="cmp-figure">
		<ul class="cmp-legend" aria-label="Legend">
			<li><span class="cmp-dot cmp-dot--compressed cmp-dot--legend" aria-hidden="true" />Compressed</li>
			<li><span class="cmp-dot cmp-dot--stored cmp-dot--legend" aria-hidden="true" />Stored</li>
		</ul>
		<div class="cmp-ratio">
			<div class="cmp-ratio__row cmp-ratio__row--axis" aria-hidden="true">
				<span />
				<div class="cmp-ratio__scale">
					<span class="cmp-ratio__direction cmp-ratio__direction--faster">← ts-pptx faster</span>
					<span class="cmp-ratio__direction cmp-ratio__direction--slower">ts-pptx slower →</span>
					<span v-for="tick in data.ticks" :key="tick" class="cmp-ratio__tick" :style="{ left: at(tick) }">
						{{ formatTick(tick) }}
					</span>
				</div>
			</div>
			<div v-for="row in data.ratios" :key="row.id" class="cmp-ratio__row" role="img" :aria-label="ratioSummary(row)">
				<span class="cmp-ratio__name">{{ row.label }}</span>
				<div class="cmp-ratio__plot">
					<span
						v-for="tick in data.ticks"
						:key="tick"
						:class="['cmp-ratio__grid', { 'cmp-ratio__grid--one': tick === 1 }]"
						:style="{ left: at(tick) }"
					/>
					<span class="cmp-ratio__span" :style="span(row)" />
					<span
						class="cmp-dot cmp-dot--compressed"
						:style="{ left: at(row.compressed) }"
						:title="`Compressed: ${formatRatio(row.compressed)}`"
					/>
					<span class="cmp-dot cmp-dot--stored" :style="{ left: at(row.stored) }" :title="`Stored: ${formatRatio(row.stored)}`" />
				</div>
			</div>
		</div>
	</figure>
</template>
