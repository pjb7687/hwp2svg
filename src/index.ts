import { convertHwpToHwpx } from './hwp/converter.js';
import { extractHwpxZip, loadHwpxDom, type HwpxDom } from './hwpx/loader.js';
import { renderToSvg, renderPageToSvg, registerFont, clearFonts } from './renderer/svg-renderer.js';
export { registerFont, clearFonts } from './renderer/svg-renderer.js';
import * as fs from './fs/idb-fs.js';

export type { HwpxDom } from './hwpx/loader.js';

let docCounter = 0;
function generateDocId(): string {
  return `doc_${Date.now()}_${docCounter++}`;
}

function detectFormat(data: ArrayBuffer): 'hwp' | 'hwpx' {
  const bytes = new Uint8Array(data, 0, 8);
  if (bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) return 'hwp';
  if (bytes[0] === 0x50 && bytes[1] === 0x4B) return 'hwpx';
  throw new Error('Unknown file format: not HWP or HWPX');
}

export class HwpxDocument {
  private dom: HwpxDom;
  private cachedPages: string[] | null = null;

  private constructor(dom: HwpxDom) {
    this.dom = dom;
  }

  static async fromHwp(data: ArrayBuffer, docId?: string): Promise<HwpxDocument> {
    const id = docId ?? generateDocId();
    await convertHwpToHwpx(data, id);
    const dom = await loadHwpxDom(id);
    return new HwpxDocument(dom);
  }

  static async fromHwpx(data: ArrayBuffer, docId?: string): Promise<HwpxDocument> {
    const id = docId ?? generateDocId();
    await extractHwpxZip(data, id);
    const dom = await loadHwpxDom(id);
    return new HwpxDocument(dom);
  }

  static async open(data: ArrayBuffer, docId?: string): Promise<HwpxDocument> {
    const format = detectFormat(data);
    if (format === 'hwp') return HwpxDocument.fromHwp(data, docId);
    return HwpxDocument.fromHwpx(data, docId);
  }

  static async fromFs(docId: string): Promise<HwpxDocument> {
    const dom = await loadHwpxDom(docId);
    return new HwpxDocument(dom);
  }

  renderPage(pageIndex: number): string | null {
    return renderPageToSvg(this.dom, pageIndex);
  }

  renderAllPages(): string[] {
    if (!this.cachedPages) {
      this.cachedPages = renderToSvg(this.dom);
    }
    return this.cachedPages;
  }

  get pageCount(): number {
    return this.renderAllPages().length;
  }

  get docId(): string {
    return this.dom.docId;
  }

  getHeaderDom(): Document {
    return this.dom.header;
  }

  getSectionDom(index: number): Document {
    return this.dom.sections[index];
  }

  get sectionCount(): number {
    return this.dom.sectionCount;
  }

  invalidateCache(): void {
    this.cachedPages = null;
  }

  async save(): Promise<void> {
    const serializer = new XMLSerializer();
    await fs.writeFile(this.dom.docId, 'Contents/header.xml',
      serializer.serializeToString(this.dom.header));
    for (let i = 0; i < this.dom.sections.length; i++) {
      await fs.writeFile(this.dom.docId, `Contents/section${i}.xml`,
        serializer.serializeToString(this.dom.sections[i]));
    }
    this.invalidateCache();
  }

  async close(): Promise<void> {
    await fs.deleteAll(this.dom.docId);
  }
}
