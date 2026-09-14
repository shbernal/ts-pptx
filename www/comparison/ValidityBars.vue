<script setup>
import { data } from './comparison.data.ts'
import { validitySegments, validitySummary } from './comparison.ts'
</script>

<template>
	<figure class="cmp-figure">
		<ul class="cmp-legend" aria-label="Legend">
			<li><span class="cmp-swatch cmp-validity__seg--clean" aria-hidden="true" />Deck with no schema error</li>
			<li><span class="cmp-swatch cmp-validity__seg--errors" aria-hidden="true" />Deck with schema errors</li>
			<li><span class="cmp-swatch cmp-validity__seg--none" aria-hidden="true" />No deck to validate</li>
		</ul>
		<div v-for="bar in data.validity" :key="bar.subject" class="cmp-validity">
			<div class="cmp-validity__name">
				<span :class="['cmp-key', `cmp-subject--${bar.subject}`]" aria-hidden="true" />{{ bar.subject }}
			</div>
			<div :class="['cmp-validity__bar', `cmp-subject--${bar.subject}`]" role="img" :aria-label="validitySummary(bar)">
				<span
					v-for="segment in validitySegments(bar)"
					:key="segment.key"
					:class="['cmp-validity__seg', `cmp-validity__seg--${segment.key}`]"
					:style="{ flexGrow: segment.count }"
					:title="`${segment.count} ${segment.label}`"
				/>
			</div>
			<p class="cmp-validity__counts">
				{{ bar.clean }} with no error, {{ bar.withErrors }} with errors, {{ bar.notBuilt }} not built, of
				{{ bar.total }} intents
			</p>
		</div>
	</figure>
</template>
