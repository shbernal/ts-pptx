/**
 * Asset and output paths for the showcase decks.
 *
 * Every path here is absolute, resolved from this file's own URL rather than from
 * `process.cwd()`. That is deliberate: `pnpm demos:build` runs the build from the
 * repository root, `pnpm --dir demos/showcases run build` runs it from this package,
 * and a reader poking at one deck will run `node quarterly-review/index.mjs` from a
 * third place. A relative `path:` would silently resolve against whichever of those
 * happened to be the cwd — `ts-pptx` hands the string straight to `fs.readFile`.
 */
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));

/** `demos/showcases/` */
export const SHOWCASES_DIR = path.resolve(HERE, "..");

/** `demos/common/` — the shared image and media library, used by every showcase. */
export const COMMON_DIR = path.resolve(SHOWCASES_DIR, "..", "common");

/** Where built decks land. Git-ignored; created on demand by the build. */
export const OUTPUT_DIR = path.join(SHOWCASES_DIR, "output");

/** Absolute path to a shared demo image, e.g. `image('chicago_bean_bohne.jpg')`. */
export function image(name) {
	return path.join(COMMON_DIR, "images", name);
}

/** Absolute path to a shared demo media file, e.g. `media('sample.mp4')`. */
export function media(name) {
	return path.join(COMMON_DIR, "media", name);
}

const MIME_BY_EXT = {
	".png": "image/png",
	".jpg": "image/jpeg",
	".jpeg": "image/jpeg",
	".gif": "image/gif",
};

/**
 * Read a shared image and return it as a `data:` URI.
 *
 * Needed for `addMedia`'s `cover`, which — unlike `addImage` — takes only a base64 string,
 * not a path. Everything else in these decks passes paths.
 */
export async function imageDataUri(name) {
	const file = image(name);
	const mime = MIME_BY_EXT[path.extname(file).toLowerCase()];
	if (!mime) throw new Error(`no MIME type known for ${name}`);
	return `data:${mime};base64,${(await readFile(file)).toString("base64")}`;
}
