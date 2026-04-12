#!/usr/bin/env node

// Polyfill browser APIs for Node.js
import 'fake-indexeddb/auto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { resolve, basename } from 'node:path';
import { inflateRawSync } from 'node:zlib';

// Set up browser API polyfills on globalThis
(globalThis as any).DOMParser = DOMParser;
(globalThis as any).XMLSerializer = XMLSerializer;
(globalThis as any).__decompressRawSync = (data: Uint8Array) => {
  return new Uint8Array(inflateRawSync(data));
};

// Import library AFTER polyfills are set up
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — path resolves correctly at runtime from dist/cli/
const { HwpxDocument } = await import('../index.js');

// Fonts are installed system-wide (~/.local/share/fonts/ or system fonts dir).
// No embedding needed — SVG references fonts by name in font-family attributes.
// Install fonts with: cp fonts/*.TTF ~/.local/share/fonts/ && fc-cache -f

// Parse CLI args
const args = process.argv.slice(2);
let inputPath = '';
let outputDir = './output';
let pageRange = '';

for (let i = 0; i < args.length; i++) {
  if (args[i] === '-o' || args[i] === '--output') {
    outputDir = args[++i];
  } else if (args[i] === '-p' || args[i] === '--pages') {
    pageRange = args[++i];
  } else if (!args[i].startsWith('-')) {
    inputPath = args[i];
  }
}

if (!inputPath) {
  console.error('Usage: hwp2svg <input.hwp|hwpx> [-o <output-dir>] [-p <pages>]');
  console.error('  -o, --output   Output directory (default: ./output)');
  console.error('  -p, --pages    Page range (e.g. "0", "0-2", "1,3,5")');
  process.exit(1);
}

function parsePageRange(range: string, total: number): number[] {
  if (!range) return Array.from({ length: total }, (_, i) => i);
  const pages: number[] = [];
  for (const part of range.split(',')) {
    if (part.includes('-')) {
      const [a, b] = part.split('-').map(Number);
      for (let i = a; i <= b && i < total; i++) pages.push(i);
    } else {
      const n = Number(part);
      if (n < total) pages.push(n);
    }
  }
  return pages;
}

// Main
const inputData = readFileSync(resolve(inputPath));
const doc = await HwpxDocument.open(inputData.buffer as ArrayBuffer);
const allPages = doc.renderAllPages();

console.log(`${basename(inputPath)}: ${allPages.length} pages`);

const pages = parsePageRange(pageRange, allPages.length);

if (!existsSync(outputDir)) {
  mkdirSync(outputDir, { recursive: true });
}

for (const idx of pages) {
  const svg = allPages[idx];
  const outPath = resolve(outputDir, `page-${idx}.svg`);
  writeFileSync(outPath, svg);
  console.log(`  wrote ${outPath}`);
}

await doc.close();
