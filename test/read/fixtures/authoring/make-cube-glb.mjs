/**
 * Writes `assets/cube.glb` — a 2×2×2 unit cube, one mesh, one material, no textures.
 *
 * Generated rather than downloaded so the asset is license-clean by construction (the same
 * reasoning as `make-assets.ps1` and `assets/ding.wav`). A representative 3D model would be
 * megabytes; this is ~1.4 KB, which is what keeps it committable as a fixture input.
 *
 * GLB container: 12-byte header (`glTF`, version, total length) then length-prefixed chunks —
 * JSON first, binary second, each padded to a 4-byte boundary (JSON with spaces, BIN with
 * zeros), per the glTF 2.0 specification §4.4.
 *
 *   node test/read/fixtures/authoring/make-cube-glb.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const HERE = dirname(fileURLToPath(import.meta.url))

/** The 6 faces as (normal, 4 corners) — 24 vertices, so each face gets its own flat normal. */
const FACES = [
	{ n: [0, 0, 1], v: [-1, -1, 1, 1, -1, 1, 1, 1, 1, -1, 1, 1] }, // +Z
	{ n: [0, 0, -1], v: [1, -1, -1, -1, -1, -1, -1, 1, -1, 1, 1, -1] }, // -Z
	{ n: [1, 0, 0], v: [1, -1, 1, 1, -1, -1, 1, 1, -1, 1, 1, 1] }, // +X
	{ n: [-1, 0, 0], v: [-1, -1, -1, -1, -1, 1, -1, 1, 1, -1, 1, -1] }, // -X
	{ n: [0, 1, 0], v: [-1, 1, 1, 1, 1, 1, 1, 1, -1, -1, 1, -1] }, // +Y
	{ n: [0, -1, 0], v: [-1, -1, -1, 1, -1, -1, 1, -1, 1, -1, -1, 1] }, // -Y
]

const positions = []
const normals = []
const indices = []
for (const [f, face] of FACES.entries()) {
	positions.push(...face.v)
	for (let k = 0; k < 4; k++) normals.push(...face.n)
	const base = f * 4
	indices.push(base, base + 1, base + 2, base, base + 2, base + 3)
}

/** Buffer layout: positions (f32), normals (f32), indices (u16) — each view 4-byte aligned. */
const posBytes = Buffer.alloc(positions.length * 4)
positions.forEach((value, i) => posBytes.writeFloatLE(value, i * 4))
const nrmBytes = Buffer.alloc(normals.length * 4)
normals.forEach((value, i) => nrmBytes.writeFloatLE(value, i * 4))
const idxBytes = Buffer.alloc(indices.length * 2)
indices.forEach((value, i) => idxBytes.writeUInt16LE(value, i * 2))
const bin = Buffer.concat([posBytes, nrmBytes, idxBytes])

const gltf = {
	asset: { version: '2.0', generator: 'ts-pptx fixture cube' },
	scene: 0,
	scenes: [{ nodes: [0] }],
	nodes: [{ mesh: 0, name: 'Cube' }],
	meshes: [{ name: 'Cube', primitives: [{ attributes: { POSITION: 0, NORMAL: 1 }, indices: 2, material: 0 }] }],
	materials: [
		{
			name: 'Gray',
			pbrMetallicRoughness: { baseColorFactor: [0.7, 0.7, 0.72, 1], metallicFactor: 0.1, roughnessFactor: 0.6 },
		},
	],
	accessors: [
		// POSITION requires min/max (glTF 2.0 §5.1); the others do not.
		{ bufferView: 0, componentType: 5126, count: 24, type: 'VEC3', min: [-1, -1, -1], max: [1, 1, 1] },
		{ bufferView: 1, componentType: 5126, count: 24, type: 'VEC3' },
		{ bufferView: 2, componentType: 5123, count: 36, type: 'SCALAR' },
	],
	bufferViews: [
		{ buffer: 0, byteOffset: 0, byteLength: posBytes.length, target: 34962 },
		{ buffer: 0, byteOffset: posBytes.length, byteLength: nrmBytes.length, target: 34962 },
		{ buffer: 0, byteOffset: posBytes.length + nrmBytes.length, byteLength: idxBytes.length, target: 34963 },
	],
	buffers: [{ byteLength: bin.length }],
}

const jsonRaw = Buffer.from(JSON.stringify(gltf), 'utf8')
const jsonPad = (4 - (jsonRaw.length % 4)) % 4
const jsonChunk = Buffer.concat([jsonRaw, Buffer.alloc(jsonPad, 0x20)]) // pad JSON with spaces
const binPad = (4 - (bin.length % 4)) % 4
const binChunk = Buffer.concat([bin, Buffer.alloc(binPad, 0)]) // pad BIN with zeros

const header = Buffer.alloc(12)
header.write('glTF', 0, 'ascii')
header.writeUInt32LE(2, 4)
header.writeUInt32LE(12 + 8 + jsonChunk.length + 8 + binChunk.length, 8)

function chunk(data, type) {
	const head = Buffer.alloc(8)
	head.writeUInt32LE(data.length, 0)
	head.write(type, 4, 'ascii')
	return Buffer.concat([head, data])
}

const glb = Buffer.concat([header, chunk(jsonChunk, 'JSON'), chunk(binChunk, 'BIN\0')])
mkdirSync(join(HERE, 'assets'), { recursive: true })
const out = join(HERE, 'assets', 'cube.glb')
writeFileSync(out, glb)
console.log(`wrote ${out} (${glb.length} bytes)`)
