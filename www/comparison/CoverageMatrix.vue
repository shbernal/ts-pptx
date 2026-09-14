<script setup>
import { computed, ref } from 'vue'
import { data } from './comparison.data.ts'
import { OUTCOMES, outcomeLabel, SUBJECTS } from './comparison.ts'

// The shared baseline is ten rows both libraries pass. Collapsed by default so the rows that
// differ are what the table opens on; the count beside it still says what the ten rows hold.
const open = ref(false)
const baseline = computed(() => data.groups.find((group) => group.id === 'shared'))
const others = computed(() => data.groups.filter((group) => group.id !== 'shared'))
</script>

<template>
	<figure class="cmp-figure">
		<ul class="cmp-legend" aria-label="Legend">
			<li v-for="outcome in OUTCOMES" :key="outcome">
				<span :class="['cmp-mark', `cmp-mark--${outcome}`]" aria-hidden="true" />{{ outcomeLabel(outcome) }}
			</li>
		</ul>
		<table class="cmp-matrix">
			<thead>
				<tr>
					<th scope="col">Intent</th>
					<th v-for="subject in SUBJECTS" :key="subject" scope="col" class="cmp-matrix__subject">
						<span :class="['cmp-key', `cmp-subject--${subject}`]" aria-hidden="true" />{{ subject }}
					</th>
				</tr>
			</thead>
			<tbody v-if="baseline">
				<tr class="cmp-matrix__group">
					<th colspan="3" scope="colgroup">{{ baseline.label }}</th>
				</tr>
				<tr>
					<th scope="row">
						<button
							type="button"
							class="cmp-matrix__toggle"
							:aria-expanded="String(open)"
							aria-controls="cmp-baseline-rows"
							@click="open = !open"
						>
							{{ baseline.rows.length }} everyday constructs
						</button>
					</th>
					<td v-for="subject in SUBJECTS" :key="subject" class="cmp-matrix__count">
						{{ baseline.emitted[subject] }} of {{ baseline.rows.length }}
					</td>
				</tr>
			</tbody>
			<tbody v-if="baseline" id="cmp-baseline-rows" :hidden="!open">
				<tr v-for="row in baseline.rows" :key="row.id" class="cmp-matrix__nested">
					<th scope="row">{{ row.label }}</th>
					<td v-for="subject in SUBJECTS" :key="subject">
						<span
							:class="['cmp-mark', `cmp-mark--${row.outcomes[subject]}`, `cmp-subject--${subject}`]"
							role="img"
							:aria-label="outcomeLabel(row.outcomes[subject])"
							:title="`${subject}: ${outcomeLabel(row.outcomes[subject])}`"
						/>
					</td>
				</tr>
			</tbody>
			<tbody v-for="group in others" :key="group.id">
				<tr class="cmp-matrix__group">
					<th colspan="3" scope="colgroup">{{ group.label }}</th>
				</tr>
				<tr v-for="row in group.rows" :key="row.id">
					<th scope="row">{{ row.label }}</th>
					<td v-for="subject in SUBJECTS" :key="subject">
						<span
							:class="['cmp-mark', `cmp-mark--${row.outcomes[subject]}`, `cmp-subject--${subject}`]"
							role="img"
							:aria-label="outcomeLabel(row.outcomes[subject])"
							:title="`${subject}: ${outcomeLabel(row.outcomes[subject])}`"
						/>
					</td>
				</tr>
			</tbody>
		</table>
	</figure>
</template>
