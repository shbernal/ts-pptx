/**
 * Build-time data for the comparison charts.
 *
 * VitePress runs `load` while it builds the site and inlines only what it returns, so the page
 * carries the few fields the charts draw and none of the recorded sources `snapshot.json` also
 * holds. The group labels come from the same `groups.json` the generated tables read.
 */
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { defineLoader } from 'vitepress'
import { type ComparisonData, shapeComparison, type SnapshotInput } from './comparison'

declare const data: ComparisonData
export { data }

export default defineLoader({
	watch: ['../../scripts/comparison/snapshot.json', '../../scripts/comparison/groups.json'],
	load(watchedFiles: string[]): ComparisonData {
		const read = (name: string): unknown => {
			const file = watchedFiles.find((candidate) => path.basename(candidate) === name)
			if (!file) throw new Error(`comparison charts: scripts/comparison/${name} was not found`)
			return JSON.parse(readFileSync(file, 'utf8'))
		}
		return shapeComparison(read('snapshot.json') as SnapshotInput, read('groups.json') as Record<string, string>)
	},
})
