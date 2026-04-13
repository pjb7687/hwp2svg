/**
 * Paragraph and text rendering for the SVG renderer.
 */

import type { Caches, ParaShapeInfo, PageDims, LayoutItem } from './svg-types.js';
import {
  hu2mm, escapeXml, attr, intAttr, children, find,
  isCJK, fontForChar, spacingForChar, fontFamilyWithFallback, wrapText, estimateTextWidth,
} from './svg-utils.js';
import { findDirectTables, renderTableEl, renderTableWithPageBreaks } from './svg-table.js';
import { mapPuaStringToUnicode } from '../hwp/hwp-section.js';

export interface RunInfo {
  text: string;
  charPrId: string;
}

export function collectRuns(paraEl: Element): RunInfo[] {
  const runs: RunInfo[] = [];
  // Walk direct children: collect runs and handle lineBreak elements
  for (const child of Array.from(paraEl.children)) {
    const ln = localName(child);
    if (ln === 'run') {
      const charPrIdRef = attr(child, 'charPrIDRef', attr(child, 'charPrId', ''));
      let text = '';
      for (const tEl of children(child, 't')) {
        text += tEl.textContent || '';
      }
      if (text) runs.push({ text, charPrId: charPrIdRef });
    } else if (ln === 'lineBreak') {
      // Insert newline as a text run
      const lastId = runs.length > 0 ? runs[runs.length - 1].charPrId : '0';
      runs.push({ text: '\n', charPrId: lastId });
    }
  }
  return runs;
}

function localName(el: Element): string {
  return el.localName || el.nodeName.split(':').pop() || el.nodeName;
}

/** Compute letter-spacing attribute string from charShape spacing percentage and font size in mm.
 *  If sampleChar is given, uses script-specific spacing. */
export function letterSpacingAttr(cs: { height: number; spacing: number; spacingLatin: number } | undefined, sampleChar?: string): string {
  if (!cs) return '';
  const sp = sampleChar ? spacingForChar(cs, sampleChar) : cs.spacing;
  if (sp === 0) return '';
  const fontSize = hu2mm(cs.height);
  const spacingMm = (sp / 100) * fontSize;
  return ` letter-spacing="${spacingMm.toFixed(2)}"`;
}

export function renderRunsAsTspans(runs: RunInfo[], caches: Caches): string {
  const parts: string[] = [];
  for (const run of runs) {
    if (!run.text) continue;
    const cs = caches.charShapes.get(run.charPrId);
    // Split each run by script type for per-script font/spacing
    let i = 0;
    while (i < run.text.length) {
      const ch = run.text[i];
      const cjk = isCJK(ch);
      let end = i + 1;
      while (end < run.text.length) {
        if (run.text[end] === ' ') { end++; continue; }
        if (isCJK(run.text[end]) !== cjk) break;
        end++;
      }
      const segText = mapPuaStringToUnicode(run.text.substring(i, end));
      const sampleCh = segText.trim()[0] || segText[0] || '';
      const fs = cs ? hu2mm(cs.height) : 3.5;
      const ff = cs ? fontForChar(cs, sampleCh) : 'sans-serif';
      const c = cs?.textColor || '#000000';
      const fw = cs?.bold ? 'bold' : 'normal';
      const fi = cs?.italic ? 'italic' : 'normal';
      const ls = letterSpacingAttr(cs, sampleCh);
      parts.push(`<tspan font-size="${fs.toFixed(2)}" font-family="${escapeXml(fontFamilyWithFallback(ff))}" fill="${c}" font-weight="${fw}" font-style="${fi}"${ls}>${escapeXml(segText)}</tspan>`);
      i = end;
    }
  }
  return parts.join('');
}

export function renderLineSegEl(
  paraEl: Element,
  seg: Element,
  segIdx: number,
  allSegs: Element[],
  caches: Caches,
  contentLeft: number,
  contentTop: number,
  contentWidth: number,
  paraShape: ParaShapeInfo | undefined,
): string | null {
  const runs = collectRuns(paraEl);
  const fullText = runs.map(r => r.text).join('');

  // Determine the text slice for this lineseg using textpos
  const startPos = intAttr(seg, 'textpos', 0);
  const nextSeg = allSegs[segIdx + 1];
  const endPos = nextSeg ? intAttr(nextSeg, 'textpos', fullText.length) : fullText.length;
  const lineText = fullText.substring(startPos, endPos);

  if (!lineText.trim()) return null;

  // Determine which runs fall within this line's text range
  const lineRuns: RunInfo[] = [];
  let pos = 0;
  for (const run of runs) {
    const runStart = pos;
    const runEnd = pos + run.text.length;
    if (runEnd <= startPos || runStart >= endPos) {
      pos += run.text.length;
      continue;
    }
    const sliceStart = Math.max(runStart, startPos) - runStart;
    const sliceEnd = Math.min(runEnd, endPos) - runStart;
    lineRuns.push({ text: run.text.substring(sliceStart, sliceEnd), charPrId: run.charPrId });
    pos += run.text.length;
  }

  const firstRun = lineRuns[0] ?? runs[0];
  const cs = firstRun ? caches.charShapes.get(firstRun.charPrId) : undefined;
  const fontSize = cs ? hu2mm(cs.height) : 3.5;
  const fontFamily = cs?.fontName || 'sans-serif';
  const color = cs?.textColor || '#000000';
  const fontWeight = cs?.bold ? 'bold' : 'normal';
  const fontStyle = cs?.italic ? 'italic' : 'normal';

  // X position: use horzpos if available (encodes leftMargin from HWP layout)
  const horzpos = intAttr(seg, 'horzpos', 0);
  const horzsize = intAttr(seg, 'horzsize', 0);
  const isLastSeg = segIdx === allSegs.length - 1;

  let textAnchor = 'start';
  // horzpos encodes the line's actual start (includes leftMargin). If zero, fall back to leftMargin.
  let x = contentLeft;
  if (horzpos > 0) {
    x = contentLeft + hu2mm(horzpos);
  } else if (paraShape && paraShape.leftMargin > 0) {
    x = contentLeft + hu2mm(paraShape.leftMargin);
  }
  if (paraShape) {
    const align = paraShape.alignment;
    if (align === 3) { textAnchor = 'middle'; x = contentLeft + contentWidth / 2; }
    else if (align === 2) { textAnchor = 'end'; x = contentLeft + contentWidth; }
  }

  // Use vertpos for absolute Y positioning within the content area
  const vertpos = intAttr(seg, 'vertpos', 0);
  const baseline = intAttr(seg, 'baseline', 0);
  const baselineY = contentTop + hu2mm(vertpos) + hu2mm(baseline);

  // JUSTIFY: stretch non-last lines to fill horzsize using SVG textLength
  let textLengthAttr = '';
  const isJustify = !paraShape || paraShape.alignment === 0; // 0 = JUSTIFY (default)
  if (isJustify && !isLastSeg && horzsize > 0 && lineText.trim().length > 1) {
    const horzMm = hu2mm(horzsize);
    textLengthAttr = ` textLength="${horzMm.toFixed(2)}" lengthAdjust="spacing"`;
  }

  if (lineRuns.length > 1) {
    const tspans = renderRunsAsTspans(lineRuns, caches);
    if (tspans) {
      return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${baselineY.toFixed(2)}" text-anchor="${textAnchor}"${textLengthAttr}>${tspans}</text>`;
    }
  }

  const ls = letterSpacingAttr(cs);
  return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${baselineY.toFixed(2)}" font-size="${fontSize.toFixed(2)}" font-family="${escapeXml(fontFamilyWithFallback(fontFamily))}" fill="${color}" font-weight="${fontWeight}" font-style="${fontStyle}"${ls}${textLengthAttr} text-anchor="${textAnchor}">${escapeXml(mapPuaStringToUnicode(lineText))}</text>`;
}

export function renderSimpleParagraph(
  runs: RunInfo[],
  caches: Caches,
  contentLeft: number,
  yPos: number,
  contentWidth: number,
  paraShape: ParaShapeInfo | undefined,
): string | null {
  const text = runs.map(r => r.text).join('');
  if (!text.trim()) return null;

  const firstRun = runs[0];
  const cs = firstRun ? caches.charShapes.get(firstRun.charPrId) : undefined;
  const fontSize = cs ? hu2mm(cs.height) : 3.5;
  const fontFamily = cs?.fontName || 'sans-serif';
  const color = cs?.textColor || '#000000';
  const fontWeight = cs?.bold ? 'bold' : 'normal';
  const fontStyle = cs?.italic ? 'italic' : 'normal';

  let textAnchor = 'start';
  let x = contentLeft;
  if (paraShape) {
    const align = paraShape.alignment;
    if (align === 3) { textAnchor = 'middle'; x = contentLeft + contentWidth / 2; }
    else if (align === 2) { textAnchor = 'end'; x = contentLeft + contentWidth; }
  }

  const y = yPos + fontSize;

  if (runs.length > 1) {
    const tspans = renderRunsAsTspans(runs, caches);
    if (tspans) {
      return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${y.toFixed(2)}" text-anchor="${textAnchor}">${tspans}</text>`;
    }
  }

  const ls = letterSpacingAttr(cs);
  return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${fontSize.toFixed(2)}" font-family="${escapeXml(fontFamilyWithFallback(fontFamily))}" fill="${color}" font-weight="${fontWeight}" font-style="${fontStyle}"${ls} text-anchor="${textAnchor}">${escapeXml(mapPuaStringToUnicode(text))}</text>`;
}

export function renderParagraphEl(
  paraEl: Element,
  caches: Caches,
  dims: PageDims,
  pages: LayoutItem[][],
  currentPage: number,
  yPos: number,
  skipSpacing = false,
): number {
  const paraPrIdRef = attr(paraEl, 'paraPrIDRef', attr(paraEl, 'paraPrId', ''));
  const paraShape = paraPrIdRef ? caches.paraShapes.get(paraPrIdRef) : undefined;

  // Spacing before paragraph (skip when caller already set absolute position from vertpos)
  if (!skipSpacing && paraShape && paraShape.spacingBefore > 0) {
    yPos += hu2mm(paraShape.spacingBefore);
  }

  // Find tables that are direct children of this paragraph (inside <hp:ctrl> or <hp:run>),
  // but NOT tables nested inside other table cells (which would cause double rendering).
  const directTables = findDirectTables(paraEl);
  const hasTables = directTables.length > 0;

  // Get only the paragraph's own line segments (from direct linesegarray child),
  // NOT linesegs nested inside table cells.
  const linesegArray = children(paraEl, 'linesegarray')[0];
  const lineSegs = linesegArray ? children(linesegArray, 'lineseg') : [];

  if (hasTables) {
    // When a paragraph contains tables, render them with page-break support.
    // Skip linesegs for table paragraphs to avoid double-counting height.
    for (const tblEl of directTables) {
      const layoutResult = renderTableWithPageBreaks(tblEl, caches, dims, pages, currentPage, yPos);
      yPos = layoutResult.yPos;
      currentPage = layoutResult.currentPage;
    }
  } else if (lineSegs.length > 0) {
    for (let si = 0; si < lineSegs.length; si++) {
      const seg = lineSegs[si];
      const segH = hu2mm(intAttr(seg, 'vertsize', 0));
      if (segH === 0) continue;

      // Use absolute Y from vertpos so each line is correctly positioned
      const segAbsY = dims.contentTop + hu2mm(intAttr(seg, 'vertpos', 0));

      if (segAbsY + segH > dims.pageBottom && pages[currentPage].length > 0) {
        pages.push([]);
        currentPage = pages.length - 1;
      }

      const svg = renderLineSegEl(paraEl, seg, si, lineSegs, caches, dims.contentLeft, dims.contentTop, dims.contentWidth, paraShape);
      if (svg) {
        pages[currentPage].push({ svg, y: segAbsY, height: segH });
      }
      yPos = segAbsY + segH;
    }
  } else {
    // No line segs — render simple paragraph
    const runs = collectRuns(paraEl);
    const text = runs.map(r => r.text).join('');
    if (text.trim()) {
      const firstRun = runs[0];
      const cs = firstRun ? caches.charShapes.get(firstRun.charPrId) : undefined;
      const fontSize = cs ? hu2mm(cs.height) : 3.5;
      const h = fontSize * 1.6;

      if (yPos + h > dims.pageBottom && pages[currentPage].length > 0) {
        pages.push([]);
        currentPage = pages.length - 1;
        yPos = dims.contentTop;
      }

      const svg = renderSimpleParagraph(runs, caches, dims.contentLeft, yPos, dims.contentWidth, paraShape);
      if (svg) {
        pages[currentPage].push({ svg, y: yPos, height: h });
      }
      yPos += h;
    }
  }

  // Spacing after (skip when caller already set absolute position from vertpos)
  if (!skipSpacing && paraShape && paraShape.spacingAfter > 0) {
    yPos += hu2mm(paraShape.spacingAfter);
  }

  return yPos;
}
