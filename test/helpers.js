'use strict'

const JSZip = require('jszip')
const PptxGenJS = require('../src/bld/pptxgen.cjs.js')

async function build(buildFn) {
	const pres = new PptxGenJS()
	buildFn(pres)
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	return { pres, zip }
}

async function readEntry(zip, path) {
	const entry = zip.file(path)
	if (!entry) throw new Error('zip entry not found: ' + path)
	return entry.async('string')
}

function listEntries(zip) {
	return Object.keys(zip.files)
}

function assert(cond, msg) {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

function assertEqual(actual, expected, msg) {
	if (actual !== expected) throw new Error('assertion failed: ' + (msg || '') + ' expected ' + JSON.stringify(expected) + ' got ' + JSON.stringify(actual))
}

module.exports = { PptxGenJS, build, readEntry, listEntries, assert, assertEqual }
