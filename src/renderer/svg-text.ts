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

const UNDERLINE_STYLE_MAP: Record<string, string> = {
  DOUBLE: 'double',
  DOT:    'dotted',
  DASH:   'dashed',
  DASHED: 'dashed',
  WAVE:   'wavy',
};

/** Compute text-decoration attribute string ("underline", "line-through", or both).
 *  Emits an inline style="text-decoration: X Y" when a non-SOLID shape (DOUBLE,
 *  DOT, DASH, WAVE) is present — SVG 2's text-decoration attribute doesn't take
 *  a style modifier, but browsers honor the CSS shorthand. */
export function textDecorationAttr(cs: { underline?: boolean; strikeout?: boolean; underlineShape?: string } | undefined): string {
  if (!cs) return '';
  const lines: string[] = [];
  if (cs.underline) lines.push('underline');
  if (cs.strikeout) lines.push('line-through');
  if (lines.length === 0) return '';

  const styleKey = (cs.underlineShape || 'SOLID').toUpperCase();
  const cssStyle = UNDERLINE_STYLE_MAP[styleKey];
  if (cssStyle) {
    // CSS shorthand: line + style. Include as style="..." so browsers / SVG
    // renderers that recognize CSS shorthand pick it up.
    return ` style="text-decoration: ${lines.join(' ')} ${cssStyle};"`;
  }
  return ` text-decoration="${lines.join(' ')}"`;
}

/** Wrap a font name with the fallback chain matching its cached kind. */
function familyWithCachedKind(caches: Caches, name: string): string {
  const kind = caches.fontKinds?.get(name);
  return fontFamilyWithFallback(name, kind);
}

/** Render the bullet glyph for a bulleted paragraph, positioned at the start
 *  of its first line. Returns an SVG `<text>` string, or null if the paragraph
 *  isn't bulleted or the bullet char is unknown. */
export function renderBulletMark(
  paraEl: Element,
  firstSeg: Element,
  caches: Caches,
  dims: import('./svg-types.js').PageDims,
  paraShape: ParaShapeInfo,
): string | null {
  if (!paraShape.bulletId) return null;
  const rawChar = caches.bullets?.get(String(paraShape.bulletId));
  if (!rawChar) return null;
  const bulletText = mapPuaStringToUnicode(rawChar);
  if (!bulletText.trim()) return null;

  // Take font metrics from the first real run so the bullet visually matches.
  const firstRun = collectRuns(paraEl)[0];
  const cs = firstRun ? caches.charShapes.get(firstRun.charPrId) : undefined;
  const fontSize = cs ? hu2mm(cs.height) : 3.5;
  const color = cs?.textColor || '#000000';
  const family = cs ? fontForChar(cs, bulletText[0]) : 'sans-serif';

  // Baseline: same math as text lineseg rendering
  const vertpos = intAttr(firstSeg, 'vertpos', 0);
  const baseline = intAttr(firstSeg, 'baseline', 0);
  const y = dims.contentTop + hu2mm(vertpos) + hu2mm(baseline);

  // X: paragraph's left edge, before any first-line indent. HWP renders the
  // bullet in the reserved indent gap, so we place it at contentLeft +
  // leftMargin (roughly where the paragraph would start if the bullet weren't
  // pushing text over). The first lineseg's horzpos already includes the space
  // for the bullet, so we anchor the bullet to the left of that position.
  const horzpos = intAttr(firstSeg, 'horzpos', 0);
  const paraLeft = hu2mm(paraShape.leftMargin);
  const x = dims.contentLeft + Math.max(paraLeft, hu2mm(horzpos) - fontSize);

  return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${fontSize.toFixed(2)}" font-family="${escapeXml(familyWithCachedKind(caches, family))}" fill="${color}" text-anchor="start">${escapeXml(bulletText)}</text>`;
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
      const td = textDecorationAttr(cs);
      parts.push(`<tspan font-size="${fs.toFixed(2)}" font-family="${escapeXml(familyWithCachedKind(caches, ff))}" fill="${c}" font-weight="${fw}" font-style="${fi}"${ls}${td}>${escapeXml(segText)}</tspan>`);
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
  // Leading whitespace: browsers use the font's own space glyph width (often
  // 0.25em from Latin metrics), but Hancom lays leading spaces out at ~0.5em
  // in Korean paragraph contexts. Convert leading spaces to an x-offset and
  // strip them from the rendered text so the visual indent matches.
  let leadingSpaceOffset = 0;
  const leadMatch = /^[ \t]+/.exec(lineText);
  if (leadMatch) {
    const n = leadMatch[0].length;
    leadingSpaceOffset = n * fontSize * 0.5;
    // Rewrite the first run to drop the leading whitespace
    if (lineRuns.length > 0) {
      const first = lineRuns[0];
      const stripped = first.text.replace(/^[ \t]+/, '');
      lineRuns[0] = { text: stripped, charPrId: first.charPrId };
    }
  }
  const strippedLineText = leadingSpaceOffset > 0 ? lineText.replace(/^[ \t]+/, '') : lineText;
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
  // Apply hanging / first-line indent when horzpos didn't already encode it.
  // HWP paragraph indent is per-paragraph:
  //   indent > 0  → first-line indent (첫 줄 들여쓰기)
  //   indent < 0  → hanging indent (내어쓰기): body lines shift right by |indent|
  //                 so continuation lines align with the text AFTER the leading
  //                 marker (e.g. "① ", "1.", "가.") on the first line.
  if (paraShape && horzpos === 0) {
    const indentMm = hu2mm(paraShape.indent);
    if (segIdx === 0 && indentMm > 0) x += indentMm;      // first-line indent
    else if (segIdx > 0 && indentMm < 0) x += -indentMm;  // hanging body indent
  }
  if (paraShape) {
    const align = paraShape.alignment;
    if (align === 3) { textAnchor = 'middle'; x = contentLeft + contentWidth / 2; }
    else if (align === 2) { textAnchor = 'end'; x = contentLeft + contentWidth; }
  }
  // Shift x by leading-space width (only meaningful for start-anchored text)
  if (textAnchor === 'start' && leadingSpaceOffset > 0) {
    x += leadingSpaceOffset;
  }

  // Use vertpos for absolute Y positioning within the content area
  const vertpos = intAttr(seg, 'vertpos', 0);
  const baseline = intAttr(seg, 'baseline', 0);
  const baselineY = contentTop + hu2mm(vertpos) + hu2mm(baseline);

  // JUSTIFY: stretch non-last lines to fill horzsize using SVG textLength.
  // horzsize is the full content-column width; body lines of a hanging-indent
  // paragraph don't get an extra column — their right edge still lands on the
  // paragraph's right margin. So shrink textLength by the applied indent so
  // the stretched line ends at the same x as line 1 (rather than overshooting).
  let textLengthAttr = '';
  const isJustify = !paraShape || paraShape.alignment === 0; // 0 = JUSTIFY (default)
  if (isJustify && !isLastSeg && horzsize > 0 && lineText.trim().length > 1) {
    let horzMm = hu2mm(horzsize);
    if (paraShape && horzpos === 0) {
      const indentMm = hu2mm(paraShape.indent);
      if (segIdx === 0 && indentMm > 0) horzMm -= indentMm;         // first-line indent shrinks text width
      else if (segIdx > 0 && indentMm < 0) horzMm -= -indentMm;     // hanging body shifted right → shorter fill
    }
    // Leading spaces we lifted into x-offset also shorten the remaining text run
    horzMm -= leadingSpaceOffset;
    textLengthAttr = ` textLength="${horzMm.toFixed(2)}" lengthAdjust="spacing"`;
  }

  if (lineRuns.length > 1) {
    const tspans = renderRunsAsTspans(lineRuns, caches);
    if (tspans) {
      return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${baselineY.toFixed(2)}" text-anchor="${textAnchor}"${textLengthAttr}>${tspans}</text>`;
    }
  }

  const ls = letterSpacingAttr(cs);
  const td = textDecorationAttr(cs);
  return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${baselineY.toFixed(2)}" font-size="${fontSize.toFixed(2)}" font-family="${escapeXml(familyWithCachedKind(caches, fontFamily))}" fill="${color}" font-weight="${fontWeight}" font-style="${fontStyle}"${ls}${td}${textLengthAttr} text-anchor="${textAnchor}">${escapeXml(mapPuaStringToUnicode(strippedLineText))}</text>`;
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
  const td = textDecorationAttr(cs);
  return `<text xml:space="preserve" x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${fontSize.toFixed(2)}" font-family="${escapeXml(familyWithCachedKind(caches, fontFamily))}" fill="${color}" font-weight="${fontWeight}" font-style="${fontStyle}"${ls}${td} text-anchor="${textAnchor}">${escapeXml(mapPuaStringToUnicode(text))}</text>`;
}

let _gradIdCounter = 0;
function nextGradId(): string { return `hwp-grad-${_gradIdCounter++}`; }

/** Emit an SVG `<rect>` (plus optional linear-gradient `<defs>`) for the
 *  given hp:rect element, positioned at (originX, originY) with the size
 *  encoded on the element. Returns the SVG plus the reserved width/height. */
function renderRectShape(
  rect: Element,
  originX: number,
  originY: number,
): { svg: string; reservedH: number; reservedW: number } {
  const w = hu2mm(intAttr(rect, 'width', 0));
  const h = hu2mm(intAttr(rect, 'height', 0));
  if (w <= 0 || h <= 0) return { svg: '', reservedH: 0, reservedW: 0 };
  const xOff = hu2mm(intAttr(rect, 'xOffset', 0));
  const x = originX + xOff;
  const y = originY;
  const kind = (rect.getAttribute('fillKind') || '').toLowerCase();
  let fillAttr = 'none';
  let defs = '';
  if (kind === 'solid') {
    fillAttr = rect.getAttribute('fillColor') || 'none';
  } else if (kind === 'gradient') {
    const colors = (rect.getAttribute('gradientColors') || '').split(',').filter(Boolean);
    const angle = parseInt(rect.getAttribute('gradientAngle') || '0', 10) || 0;
    if (colors.length >= 1) {
      const gid = nextGradId();
      // HWP shear angle: 0° = top-to-bottom, 90° = left-to-right.
      const rad = (angle * Math.PI) / 180;
      const dx = Math.sin(rad), dy = Math.cos(rad);
      const stops = colors.map((c, i) => {
        const pct = colors.length === 1 ? 100 : (i * 100 / (colors.length - 1));
        return `<stop offset="${pct.toFixed(2)}%" stop-color="${escapeXml(c)}"/>`;
      }).join('');
      // Use gradientUnits="userSpaceOnUse" with a vector spanning the rect.
      defs = `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" x1="${x.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(x + w * dx).toFixed(2)}" y2="${(y + h * dy).toFixed(2)}">${stops}</linearGradient></defs>`;
      fillAttr = `url(#${gid})`;
    }
  }
  const svg = `${defs}<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${fillAttr}" stroke="none"/>`;
  return { svg, reservedH: h, reservedW: w };
}

/** Emit `<image>` elements for any `<hp:pic>` inside this paragraph.
 *  Pictures are inline anchors so they render at the caller's current x/y.
 *  Caller passes (originX, originY) — the top-left where images should land.
 *  Returns the SVG string (possibly empty) plus the total vertical space
 *  reserved for the pictures — callers need this so subsequent content
 *  doesn't overlap the image. */
export function renderPictures(
  paraEl: Element,
  caches: Caches,
  originX: number,
  originY: number,
): { svg: string; reservedH: number; reservedW: number } {
  const parts: string[] = [];
  let maxH = 0;
  let sumW = 0;
  void renderRectShape; // rect rendering parked — fill offset needs spec verification
  for (const run of children(paraEl, 'run')) {
    for (const pic of children(run, 'pic')) {
      const w = hu2mm(intAttr(pic, 'width', 0));
      const h = hu2mm(intAttr(pic, 'height', 0));
      const binId = intAttr(pic, 'binaryItemIDRef', 0);
      if (w <= 0 || h <= 0 || !binId) continue;
      const href = caches.binData?.get(binId);
      if (!href) continue;
      const xOff = hu2mm(intAttr(pic, 'xOffset', 0));
      const x = originX + xOff + sumW;
      const y = originY;
      // Default SVG image aspect-ratio (xMidYMid meet) preserves the source
      // aspect ratio and letterboxes inside the (w × h) box, matching HWP's
      // behavior when the anchored picture size doesn't match the image's
      // native pixel ratio. Using preserveAspectRatio="none" would stretch
      // the image incorrectly.
      parts.push(
        `<image xlink:href="${escapeXml(href)}" href="${escapeXml(href)}" x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}"/>`,
      );
      if (h > maxH) maxH = h;
      sumW += w;
    }
  }
  return { svg: parts.join('\n'), reservedH: maxH, reservedW: sumW };
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

  // Render any anchored pictures BEFORE text/tables and reserve their height.
  // Pictures live inside <hp:run> as <hp:pic>. We emit them at yPos so the
  // subsequent text starts below the image, matching HWP's "flow with text"
  // default. Overlap-with-text pictures aren't distinguished here yet.
  const pics = renderPictures(paraEl, caches, dims.contentLeft, yPos);
  if (pics.svg) {
    if (yPos + pics.reservedH > dims.pageBottom && pages[currentPage].length > 0) {
      pages.push([]);
      currentPage = pages.length - 1;
      yPos = dims.contentTop;
    }
    pages[currentPage].push({ svg: pics.svg, y: yPos, height: pics.reservedH });
    yPos += pics.reservedH;
  }

  // Find tables that are direct children of this paragraph (inside <hp:ctrl> or <hp:run>),
  // but NOT tables nested inside other table cells (which would cause double rendering).
  const directTables = findDirectTables(paraEl);
  const hasTables = directTables.length > 0;

  // Get only the paragraph's own line segments (from direct linesegarray child),
  // NOT linesegs nested inside table cells.
  const linesegArray = children(paraEl, 'linesegarray')[0];
  const lineSegs = linesegArray ? children(linesegArray, 'lineseg') : [];

  // Advance yPos to the lineseg that actually holds the inline table content.
  // The paragraph may start with invisible control chars (secPr, colPr,
  // pageNum, …) that occupy their own initial lineseg with vertsize=0/small
  // but no visible content. The table then sits on a later lineseg whose
  // vertpos reflects its true baseline. Without this shift, tables render
  // at contentTop even when HWP placed them lower in the paragraph flow.
  if (hasTables && lineSegs.length > 1) {
    const firstVp = intAttr(lineSegs[0], 'vertpos', 0);
    const lastVp = intAttr(lineSegs[lineSegs.length - 1], 'vertpos', firstVp);
    if (lastVp > firstVp) {
      yPos += hu2mm(lastVp - firstVp);
    }
  }

  if (hasTables) {
    // When a paragraph contains tables, render them with page-break support.
    // Skip linesegs for table paragraphs to avoid double-counting height.
    // Inline tables (treatAsChar=1) sit on the text baseline, so any text
    // runs BEFORE the table shift the table X by their width.
    let xOffset = 0;
    for (const child of Array.from(paraEl.children)) {
      const ln = child.localName || child.nodeName.split(':').pop();
      if (ln === 'run') {
        // Text before any tbl in this run
        let sawTable = false;
        for (const gchild of Array.from(child.children)) {
          const gln = gchild.localName || gchild.nodeName.split(':').pop();
          if (gln === 'tbl') { sawTable = true; break; }
          if (gln === 't') {
            const runCharPrIdRef = attr(child, 'charPrIDRef', attr(child, 'charPrId', ''));
            const runCs = runCharPrIdRef ? caches.charShapes.get(runCharPrIdRef) : undefined;
            const runFontSize = runCs ? hu2mm(runCs.height) : 3.5;
            xOffset += estimateTextWidth(gchild.textContent || '', runFontSize);
          }
        }
        if (sawTable) break;
      }
    }
    for (const tblEl of directTables) {
      const layoutResult = renderTableWithPageBreaks(tblEl, caches, dims, pages, currentPage, yPos, xOffset);
      yPos = layoutResult.yPos;
      currentPage = layoutResult.currentPage;
      xOffset = 0; // subsequent tables in same paragraph start fresh
    }
  } else if (lineSegs.length > 0) {
    // If the paragraph is bulleted, render the bullet glyph once at the
    // start of the first line. Hancom draws bullets in the reserved
    // left-margin area, using the first run's char shape for size/font.
    const bulletSvg = paraShape ? renderBulletMark(paraEl, lineSegs[0], caches, dims, paraShape) : null;
    if (bulletSvg) {
      const seg0 = lineSegs[0];
      const y0 = dims.contentTop + hu2mm(intAttr(seg0, 'vertpos', 0));
      const h0 = hu2mm(intAttr(seg0, 'vertsize', 0));
      pages[currentPage].push({ svg: bulletSvg, y: y0, height: h0 });
    }
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
