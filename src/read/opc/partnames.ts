/**
 * OPC partname helpers (ECMA-376 Part 2 pack-URI grammar).
 *
 * A partname is an absolute, `/`-separated path inside the package, e.g.
 * `/ppt/slides/slide1.xml`. Zip entry paths omit the leading slash.
 */
import { PackageReadError } from '../../errors.js'
import type { OpcPackage } from './package.js'

/**
 * Resolve the single relationship of `type` owned by `partName` to its target partname, or
 * `null` when the part declares none. For the one-of-a-kind links in a deck's spine — a slide to
 * its layout, a layout to its master, a master to its theme — where a second relationship of the
 * same type would itself be malformed.
 *
 * The `OpcPackage` import is type-only, so this does not create a cycle with `package.ts` (which
 * imports this module at run time).
 */
export function resolveSingleRel(opc: OpcPackage, partName: string, type: string): string | null {
	const rels = opc.relationshipsFor(partName)
	const rel = rels.byType(type)[0]
	return rel ? rels.resolveTarget(rel.id) : null
}

export function zipPathToPartName(zipPath: string): string {
	return zipPath.startsWith('/') ? zipPath : `/${zipPath}`
}

export function partNameToZipPath(partName: string): string {
	return partName.startsWith('/') ? partName.slice(1) : partName
}

/**
 * Lowercased extension without the dot, or `''` when the part has none.
 * Per OPC, the extension is everything after the last dot — so the leading
 * dot of `/_rels/.rels` still yields `rels`.
 */
export function partNameExtension(partName: string): string {
	const lastSegment = partName.slice(partName.lastIndexOf('/') + 1)
	const dot = lastSegment.lastIndexOf('.')
	return dot < 0 ? '' : lastSegment.slice(dot + 1).toLowerCase()
}

/** Partname of the `.rels` part holding `sourcePartName`'s relationships ('/' = package). */
export function relsPartNameFor(sourcePartName: string): string {
	if (sourcePartName === '/') return '/_rels/.rels'
	const dir = sourcePartName.slice(0, sourcePartName.lastIndexOf('/'))
	const file = sourcePartName.slice(sourcePartName.lastIndexOf('/') + 1)
	return `${dir}/_rels/${file}.rels`
}

/**
 * Build a relationship `Target` for `targetPartName` relative to its source
 * part's directory — the inverse of {@link resolveRelativePartName}. Both names
 * are absolute partnames. E.g. source `/ppt/slides/slide1.xml`, target
 * `/ppt/media/image1.png` → `../media/image1.png`.
 */
export function relativePartName(sourcePartName: string, targetPartName: string): string {
	const from = sourcePartName.slice(1).split('/').slice(0, -1) // source directory segments
	const to = targetPartName.slice(1).split('/') // target segments incl. filename
	let common = 0
	while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++
	const up = from.slice(common).map(() => '..')
	const down = to.slice(common)
	return [...up, ...down].join('/')
}

/**
 * Resolve a relationship target against its source part, per OPC pack-URI
 * resolution: relative targets resolve against the source part's directory.
 *
 * @param sourcePartName partname of the part owning the relationship ('/' for package-level)
 * @param target relationship Target attribute (relative like `../media/image1.png`, or absolute)
 */
export function resolveRelativePartName(sourcePartName: string, target: string): string {
	const path = target.startsWith('/') ? target : sourcePartName.slice(0, sourcePartName.lastIndexOf('/') + 1) + target
	const segments: string[] = []
	for (const segment of path.split('/')) {
		if (segment === '' || segment === '.') continue
		if (segment === '..') {
			if (segments.length === 0)
				throw new PackageReadError(
					'package/relationship-target-escapes-root',
					`Relationship target ${target} escapes the package root (source ${sourcePartName})`
				)
			segments.pop()
		} else {
			segments.push(segment)
		}
	}
	return `/${segments.join('/')}`
}
