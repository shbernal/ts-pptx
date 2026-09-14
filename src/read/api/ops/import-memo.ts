/**
 * What a deck remembers between imports, and the one place that forgets it.
 *
 * Two memos outlive any one import call. The copy registry maps each source package's partnames to
 * the partnames their copies were given here, so a second import from the same source reuses the
 * layout, master, theme and media it already brought across. The rescale memo holds the parts whose
 * geometry a rescale has already rewritten, so a layout or master shared by several imports is
 * scaled once.
 *
 * Both hold destination partnames, and removing a slide frees names that `reservePartNameLike`
 * hands out again. An entry left behind then answers for a part that is no longer there: the
 * registry calls a deleted slide or image already copied, and the rescale memo calls a new slide
 * already scaled. So both live behind one object whose {@link ImportMemo.forget} is how a part
 * leaves the deck's import state, and a memo added here has to answer to it.
 */

import type { OpcPackage } from '../../opc/package.js'

export class ImportMemo {
	readonly #registries = new Map<OpcPackage, Map<string, string>>()
	readonly #rescaled = new Set<string>()

	/** The copy registry for imports out of `source`: source partname → destination partname. */
	registryFor(source: OpcPackage): Map<string, string> {
		let registry = this.#registries.get(source)
		if (!registry) {
			registry = new Map()
			this.#registries.set(source, registry)
		}
		return registry
	}

	/** Destination partnames whose geometry a rescale has already rewritten, as a live set the rescale adds to. */
	get rescaledParts(): Set<string> {
		return this.#rescaled
	}

	/** Drop every entry naming one of `partNames`, which have left the deck. */
	forget(partNames: Iterable<string>): void {
		const gone = new Set(partNames)
		for (const registry of this.#registries.values()) {
			for (const [sourcePartName, destPartName] of registry) {
				if (gone.has(destPartName)) registry.delete(sourcePartName)
			}
		}
		for (const partName of gone) this.#rescaled.delete(partName)
	}
}
