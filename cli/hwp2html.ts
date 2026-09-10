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
// @ts-ignore — path resolves correctly at runtime from dist/cli/
const { HwpxDocument } = await import('../index.js');

const args = process.argv.slice(2);
let inputPath = '';
let outputPath = '';
let fragmentOnly = false;
let title = '';

for (let i = 0; i < args.length; i++) {
  if (args[i] === '-o' || args[i] === '--output') {
    outputPath = args[++i];
  } else if (args[i] === '--fragment') {
    fragmentOnly = true;
  } else if (args[i] === '--title') {
    title = args[++i];
  } else if (!args[i].startsWith('-')) {
    inputPath = args[i];
  }
}

if (!inputPath) {
  console.error('Usage: hwp2html <input.hwp|hwpx> [-o <output.html>] [--fragment] [--title <title>]');
  console.error('  -o, --output   Output HTML file (default: same name as input with .html extension)');
  console.error('  --fragment     Emit body markup only, no <!doctype>/<html>/<head>');
  console.error('  --title        Title used in <title>');
  process.exit(1);
}

const resolvedInput = resolve(inputPath);
if (!outputPath) {
  const base = basename(resolvedInput).replace(/\.[^.]+$/, '');
  outputPath = resolve(dirname(resolvedInput), base + '.html');
}

const inputData = readFileSync(resolvedInput);
const doc = await HwpxDocument.open(inputData.buffer as ArrayBuffer);

const html = doc.renderHtml({
  full: !fragmentOnly,
  title: title || basename(resolvedInput).replace(/\.[^.]+$/, ''),
});

writeFileSync(outputPath, html);
console.log(`Wrote ${outputPath}`);

await doc.close();
