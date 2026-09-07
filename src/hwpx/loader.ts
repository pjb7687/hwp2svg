import JSZip from 'jszip';
import * as fs from '../fs/idb-fs.js';

export interface HwpxDom {
  docId: string;
  header: Document;
  sections: Document[];
  sectionCount: number;
  /** Embedded picture data: binDataId (from HWP DocInfo BinData order) → data URI. */
  binData: Map<number, string>;
}

/** Extract HWPX ZIP archive into IndexedDB FS. */
export async function extractHwpxZip(data: ArrayBuffer, docId: string): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  const entries = Object.entries(zip.files);

  for (const [path, file] of entries) {
    if (file.dir) continue;
    if (path.endsWith('.xml') || path.endsWith('.hpf') || path === 'mimetype') {
      const text = await file.async('text');
      await fs.writeFile(docId, path, text);
    } else {
      const buf = await file.async('uint8array');
      await fs.writeFile(docId, path, buf);
    }
  }
}

/** Load HWPX DOM from IndexedDB FS. */
export async function loadHwpxDom(docId: string): Promise<HwpxDom> {
  const parser = new DOMParser();

  const headerXml = await fs.readFileAsString(docId, 'Contents/header.xml');
  const header = parser.parseFromString(headerXml, 'text/xml');

  const sections: Document[] = [];
  let i = 0;
  while (true) {
    const path = `Contents/section${i}.xml`;
    if (!(await fs.exists(docId, path))) break;
    const xml = await fs.readFileAsString(docId, path);
    sections.push(parser.parseFromString(xml, 'text/xml'));
    i++;
  }

  if (sections.length === 0) {
    throw new Error('No section files found in HWPX document');
  }

  // Load embedded binary items (pictures) from BinData/BIN*.<ext>. The HWP
  // converter names each file after the binData id (hex, 4-digit uppercase).
  // We eager-load and base64-encode so the SVG renderer can synchronously
  // reference them as data URIs.
  const binData = new Map<number, string>();
  const binPaths = await fs.readDir(docId, 'BinData/');
  for (const path of binPaths) {
    const base = path.replace(/^BinData\//, '');
    const match = /^BIN([0-9A-Fa-f]{1,8})\.([a-zA-Z0-9]+)$/.exec(base);
    if (!match) continue;
    const id = parseInt(match[1], 16);
    const ext = match[2].toLowerCase();
    const mime = MIME_BY_EXT[ext];
    if (!mime) continue;
    const bytes = await fs.readFile(docId, path);
    if (typeof bytes === 'string') continue;
    binData.set(id, `data:${mime};base64,${bytesToBase64(bytes)}`);
  }

  return { docId, header, sections, sectionCount: sections.length, binData };
}

const MIME_BY_EXT: Record<string, string> = {
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
};

function bytesToBase64(bytes: Uint8Array): string {
  // btoa can't take a Uint8Array directly in Node (no `btoa`) or in browsers
  // (input must be a binary string). Chunk to keep String.fromCharCode from
  // blowing the arg-count limit on large images.
  let s = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    s += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  // In Node, Buffer.from(s, 'binary') gives us a way to reach base64. In the
  // browser, btoa is available on globalThis.
  if (typeof (globalThis as { Buffer?: { from: (s: string, enc: string) => { toString: (enc: string) => string } } }).Buffer !== 'undefined') {
    return (globalThis as { Buffer: { from: (s: string, enc: string) => { toString: (enc: string) => string } } }).Buffer.from(s, 'binary').toString('base64');
  }
  return (globalThis as { btoa: (s: string) => string }).btoa(s);
}
