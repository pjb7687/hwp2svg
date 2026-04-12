#!/usr/bin/env node

// Polyfill browser APIs for Node.js
import 'fake-indexeddb/auto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, basename, dirname } from 'node:path';
import { inflateRawSync } from 'node:zlib';

(globalThis as any).DOMParser = DOMParser;
(globalThis as any).XMLSerializer = XMLSerializer;
(globalThis as any).__decompressRawSync = (data: Uint8Array) => {
  return new Uint8Array(inflateRawSync(data));
};

// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — paths resolve correctly at runtime from dist/cli/
const { convertHwpToHwpx } = await import('../hwp/converter.js');
// @ts-ignore
const fs = await import('../fs/idb-fs.js');
import JSZip from 'jszip';

// Parse CLI args
const args = process.argv.slice(2);
let inputPath = '';
let outputPath = '';

for (let i = 0; i < args.length; i++) {
  if (args[i] === '-o' || args[i] === '--output') {
    outputPath = args[++i];
  } else if (!args[i].startsWith('-')) {
    inputPath = args[i];
  }
}

if (!inputPath) {
  console.error('Usage: hwp2hwpx <input.hwp> [-o <output.hwpx>]');
  console.error('  -o, --output   Output file (default: same name as input with .hwpx extension)');
  process.exit(1);
}

const resolvedInput = resolve(inputPath);
if (!outputPath) {
  const base = basename(resolvedInput, '.hwp');
  outputPath = resolve(dirname(resolvedInput), base + '.hwpx');
}

const inputData = readFileSync(resolvedInput);
const docId = `hwp2hwpx-${Date.now()}`;

console.log(`Converting ${basename(resolvedInput)} ...`);
await convertHwpToHwpx(inputData.buffer as ArrayBuffer, docId);

// Package all files from the virtual FS into a ZIP (.hwpx)
const zip = new JSZip();
const allFiles = await fs.readDir(docId, '');

for (const filePath of allFiles) {
  const data = await fs.readFile(docId, filePath);
  zip.file(filePath, data);
}

const zipData = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
writeFileSync(outputPath, zipData);
console.log(`Wrote ${outputPath}`);

await fs.deleteAll(docId);
