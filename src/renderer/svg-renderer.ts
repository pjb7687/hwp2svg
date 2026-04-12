/**
 * SVG renderer: converts HWPX XML DOM directly to SVG strings, one per page.
 *
 * HWP uses HWPUNIT (1/7200 inch) internally.
 * We convert to mm for SVG output: 1 HWPUNIT = 25.4/7200 mm ≈ 0.003528 mm
 */

import type { HwpxDom } from '../hwpx/loader.js';
import type { Caches, LayoutItem, PageDims } from './svg-types.js';
import { hu2mm, intAttr, children, localName, buildFontFaceCSS } from './svg-utils.js';
import { buildCaches, getPageDims } from './svg-cache.js';
import { renderParagraphEl } from './svg-text.js';
import { renderTableEl, resetClipIdCounter } from './svg-table.js';

export { registerFont, clearFonts } from './svg-utils.js';
export type { CharShapeInfo, ParaShapeInfo, BorderFillInfo, Caches, PageDims, LayoutItem, TableResult } from './svg-types.js';

// ── Public API ──

export function renderToSvg(dom: HwpxDom): string[] {
  resetClipIdCounter();  // Reset clip ID counter for each document render
  const caches = buildCaches(dom.header);
  const pages: string[] = [];

  for (const sectionDoc of dom.sections) {
    const sectionPages = renderSection(sectionDoc, caches);
    pages.push(...sectionPages);
  }

  return pages;
}

export function renderPageToSvg(dom: HwpxDom, pageIndex: number): string | null {
  const all = renderToSvg(dom);
  return all[pageIndex] ?? null;
}

// ── Section rendering ──

function renderSection(sectionDoc: Document, caches: Caches): string[] {
  const root = sectionDoc.documentElement;
  const dims = getPageDims(root);

  const pages: LayoutItem[][] = [[]];
  let currentPage = 0;
  let yPos = dims.contentTop;
  // Track previous paragraph's START vertPos for page-break detection.
  // Using start (not end) avoids false breaks when a table's lineseg spans
  // its caption area, making lastVertEnd larger than where captions begin.
  let lastParaStartVP = 0;

  // Walk direct <hp:p> and <hp:tbl> children of section root
  for (const child of Array.from(root.children)) {
    const lname = localName(child);
    if (lname === 'p') {
      const linesegArray = children(child, 'linesegarray')[0];
      const firstSeg = linesegArray ? children(linesegArray, 'lineseg')[0] : null;
      if (firstSeg && lastParaStartVP > 0) {
        const vertpos = intAttr(firstSeg, 'vertpos', 0);
        if (vertpos < lastParaStartVP - 2000 && pages[currentPage].length > 0) {
          pages.push([]);
          currentPage = pages.length - 1;
          yPos = dims.contentTop;
        }
      }

      let useAbsolutePos = false;
      if (firstSeg) {
        const vertpos = intAttr(firstSeg, 'vertpos', 0);
        const absoluteY = dims.contentTop + hu2mm(vertpos);
        yPos = absoluteY;
        useAbsolutePos = true;
        lastParaStartVP = vertpos;
      }

      yPos = renderParagraphEl(child, caches, dims, pages, currentPage, yPos, useAbsolutePos);
      currentPage = pages.length - 1;
    } else if (lname === 'tbl') {
      const result = renderTableEl(child, caches, dims.contentLeft, yPos, dims.contentWidth);
      if (result) {
        const { svg, height } = result;
        if (yPos + height > dims.pageBottom && pages[currentPage].length > 0) {
          pages.push([]);
          currentPage++;
          yPos = dims.contentTop;
        }
        pages[currentPage].push({ svg, y: yPos, height });
        yPos += height;
      }
    }
  }

  return buildSvgPages(pages, dims);
}

function buildSvgPages(pages: LayoutItem[][], dims: PageDims): string[] {
  return pages.filter(p => p.length > 0).map(items => {
    const content = items.map(i => i.svg).join('\n');
    const fontCSS = buildFontFaceCSS();
    return [
      `<svg xmlns="http://www.w3.org/2000/svg" width="${dims.pageW.toFixed(2)}mm" height="${dims.pageH.toFixed(2)}mm" viewBox="0 0 ${dims.pageW.toFixed(2)} ${dims.pageH.toFixed(2)}">`,
      fontCSS,
      `<rect width="100%" height="100%" fill="white"/>`,
      content,
      `</svg>`,
    ].join('\n');
  });
}
