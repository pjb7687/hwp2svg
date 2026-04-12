/**
 * HWP 5.0 binary → HWPX XML converter.
 * Reads OLE2 compound documents using the `cfb` package and generates
 * HWPX XML files, writing them to the IndexedDB FS.
 */

import * as CFB from 'cfb';
import { dataView } from './record.js';
import * as TAG from './constants.js';
import * as fs from '../fs/idb-fs.js';
import { parseDocInfoData } from './hwp-docinfo.js';
import { parseSectionRecords } from './hwp-section.js';
import { generateHeaderXml, generateSectionXml, generateContentHpf, generateVersionXml } from './hwp-xml.js';
import type { DocHeader } from './hwp-types.js';

// ── Global decompression polyfill type ──

declare global {
  var __decompressRawSync: ((data: Uint8Array) => Uint8Array) | undefined;
}

// ── Decompression ──

async function decompressRaw(data: Uint8Array): Promise<Uint8Array> {
  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();

  writer.write(data as Uint8Array<ArrayBuffer>);
  writer.close();

  const chunks: Uint8Array[] = [];
  let done = false;
  while (!done) {
    const result = await reader.read();
    if (result.done) {
      done = true;
    } else {
      chunks.push(result.value);
    }
  }

  const totalLen = chunks.reduce((acc, c) => acc + c.length, 0);
  const out = new Uint8Array(totalLen);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

// ── Stream access ──

async function getStream(cfb: CFB.CFB$Container, path: string, decompress: boolean): Promise<Uint8Array> {
  const entry = CFB.find(cfb, `/${path}`);
  if (!entry) throw new Error(`Stream not found: ${path}`);
  let buf = new Uint8Array(entry.content as unknown as ArrayBuffer);
  if (decompress && buf.length > 0) {
    try {
      if (typeof globalThis.__decompressRawSync === 'function') {
        buf = globalThis.__decompressRawSync(buf) as Uint8Array<ArrayBuffer>;
      } else {
        buf = await decompressRaw(buf) as Uint8Array<ArrayBuffer>;
      }
    } catch {
      // Some streams might not actually be compressed
    }
  }
  return buf;
}

// ── FileHeader ──

function parseFileHeader(cfb: CFB.CFB$Container): DocHeader {
  const entry = CFB.find(cfb, '/FileHeader');
  if (!entry) throw new Error('FileHeader stream not found');
  const buf = new Uint8Array(entry.content as unknown as ArrayBuffer);
  const view = dataView(buf);

  let sig = '';
  for (let i = 0; i < 32 && buf[i] !== 0; i++) {
    sig += String.fromCharCode(buf[i]);
  }
  if (!sig.startsWith(TAG.HWP_SIGNATURE)) {
    throw new Error(`Invalid HWP signature: ${sig}`);
  }

  const ver = view.getUint32(32, true);
  const flags = view.getUint32(36, true);

  return {
    version: {
      major: (ver >> 24) & 0xFF,
      minor: (ver >> 16) & 0xFF,
      patch: (ver >> 8) & 0xFF,
      revision: ver & 0xFF,
    },
    flags,
  };
}

// ── BinData extraction ──

async function extractBinData(
  cfb: CFB.CFB$Container,
  docId: string,
  binDataItems: Array<{ type: string; binDataId?: number; extension?: string }>,
): Promise<void> {
  for (const item of binDataItems) {
    if (item.type === 'embedding' && item.binDataId !== undefined && item.extension) {
      const streamName = `BIN${item.binDataId.toString(16).toUpperCase().padStart(4, '0')}.${item.extension}`;
      const entry = CFB.find(cfb, `/BinData/${streamName}`);
      if (entry) {
        const data = new Uint8Array(entry.content as unknown as ArrayBuffer);
        await fs.writeFile(docId, `BinData/${streamName}`, data);
      }
    }
  }
}

// ── Main export ──

export async function convertHwpToHwpx(data: ArrayBuffer, docId: string): Promise<number> {
  const uint8 = new Uint8Array(data);
  const cfb = CFB.read(uint8, { type: 'array' });

  const header = parseFileHeader(cfb);
  const isCompressed = (header.flags & 1) !== 0;

  // Parse DocInfo
  const docInfoBuf = await getStream(cfb, 'DocInfo', isCompressed);
  const docInfo = parseDocInfoData(docInfoBuf);

  // Write header.xml
  const headerXml = generateHeaderXml(docInfo);
  await fs.writeFile(docId, 'Contents/header.xml', headerXml);

  // Parse and write each section
  const sectionCount = docInfo.sectionCount;
  for (let i = 0; i < sectionCount; i++) {
    const secBuf = await getStream(cfb, `BodyText/Section${i}`, isCompressed);
    const { pageDef, paragraphs } = parseSectionRecords(secBuf);
    const sectionXml = generateSectionXml(i, pageDef, paragraphs);
    await fs.writeFile(docId, `Contents/section${i}.xml`, sectionXml);
  }

  // Write manifest
  const contentHpf = generateContentHpf(sectionCount);
  await fs.writeFile(docId, 'Contents/content.hpf', contentHpf);

  // Write version.xml
  const versionXml = generateVersionXml(header);
  await fs.writeFile(docId, 'version.xml', versionXml);

  // Extract embedded binary data
  await extractBinData(cfb, docId, docInfo.binDataItems);

  return sectionCount;
}
