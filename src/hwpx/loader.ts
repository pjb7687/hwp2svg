import JSZip from 'jszip';
import * as fs from '../fs/idb-fs.js';

export interface HwpxDom {
  docId: string;
  header: Document;
  sections: Document[];
  sectionCount: number;
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

  return { docId, header, sections, sectionCount: sections.length };
}
