/**
 * Table rendering for the SVG renderer.
 */

import type { Caches, BorderFillInfo, PageDims, LayoutItem, TableResult } from './svg-types.js';
import {
  hu2mm, escapeXml, attr, intAttr, find, children, findChild,
  isCJK, fontForChar, spacingForChar, fontFamilyWithFallback, wrapText, estimateTextWidth,
} from './svg-utils.js';
import { collectRuns, letterSpacingAttr } from './svg-text.js';

// ── Helpers for avoiding double-rendering of nested tables ──

/**
 * Find tables that belong directly to a paragraph element.
 * Tables may be inside <hp:ctrl> or <hp:run> children, but we must NOT
 * descend into table cells (tc/subList) which would find inner tables.
 */
export function findDirectTables(paraEl: Element): Element[] {
  const results: Element[] = [];
  function walk(el: Element) {
    for (const child of Array.from(el.children)) {
      const ln = localName(child);
      if (ln === 'tbl') {
        results.push(child);
        // Don't descend into the table — inner tables belong to cell content
      } else if (ln === 'tc' || ln === 'subList') {
        // Don't descend into table cells
      } else {
        walk(child);
      }
    }
  }
  walk(paraEl);
  return results;
}

/**
 * Find direct child 'tc' elements of a table row, without descending
 * into nested tables inside cell content.
 */
export function findDirectCells(rowEl: Element): Element[] {
  const results: Element[] = [];
  function walk(el: Element) {
    for (const child of Array.from(el.children)) {
      const ln = localName(child);
      if (ln === 'tc') {
        results.push(child);
        // Don't descend further — nested tables inside cells are separate
      } else {
        walk(child);
      }
    }
  }
  walk(rowEl);
  return results;
}

function localName(el: Element): string {
  return el.localName || el.nodeName.split(':').pop() || el.nodeName;
}

// ── Cell content clip ID counter ──

let _clipIdCounter = 0;

export function resetClipIdCounter(): void {
  _clipIdCounter = 0;
}

// ── Table rendering helpers ──

/** Parse border width string like "0.3mm" to a number in mm */
function parseBorderWidth(s: string): number {
  const m = s.match(/([\d.]+)/);
  return m ? parseFloat(m[1]) : 0.1;
}

/** Check if a border type string indicates an invisible (NONE) border */
function isBorderNone(borderType: string): boolean {
  return borderType === 'NONE' || borderType === 'none';
}

/** HWP border types as strings from converter output. */
function borderStrokeDasharray(borderType: string): string {
  const t = borderType.toUpperCase();
  switch (t) {
    case 'DASHED':
    case 'DASH': return ' stroke-dasharray="1.5,0.5"';
    case 'DOT': return ' stroke-dasharray="0.3,0.3"';
    case 'DASH_DOT': return ' stroke-dasharray="1.5,0.5,0.3,0.5"';
    case 'DASH_DOT_DOT': return ' stroke-dasharray="1.5,0.3,0.3,0.3,0.3,0.3"';
    default: return '';  // SOLID or unknown
  }
}

/**
 * Render cell background fill and border lines.
 * Returns SVG elements for the cell background rect and 4 border lines.
 */
export function renderCellBorderAndFill(
  x: number, y: number, w: number, h: number,
  borderFill: BorderFillInfo | undefined,
): string[] {
  const parts: string[] = [];

  // Background fill
  if (borderFill?.fillColor) {
    parts.push(`<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${borderFill.fillColor}" stroke="none"/>`);
  }

  if (borderFill) {
    // Left border
    const lw = parseBorderWidth(borderFill.leftBorderWidth);
    if (lw > 0 && !isBorderNone(borderFill.leftBorderType)) {
      parts.push(`<line x1="${x.toFixed(2)}" y1="${y.toFixed(2)}" x2="${x.toFixed(2)}" y2="${(y + h).toFixed(2)}" stroke="${borderFill.leftBorderColor}" stroke-width="${lw.toFixed(2)}"${borderStrokeDasharray(borderFill.leftBorderType)}/>`);
    }
    // Right border
    const rw = parseBorderWidth(borderFill.rightBorderWidth);
    if (rw > 0 && !isBorderNone(borderFill.rightBorderType)) {
      parts.push(`<line x1="${(x + w).toFixed(2)}" y1="${y.toFixed(2)}" x2="${(x + w).toFixed(2)}" y2="${(y + h).toFixed(2)}" stroke="${borderFill.rightBorderColor}" stroke-width="${rw.toFixed(2)}"${borderStrokeDasharray(borderFill.rightBorderType)}/>`);
    }
    // Top border
    const tw = parseBorderWidth(borderFill.topBorderWidth);
    if (tw > 0 && !isBorderNone(borderFill.topBorderType)) {
      parts.push(`<line x1="${x.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(x + w).toFixed(2)}" y2="${y.toFixed(2)}" stroke="${borderFill.topBorderColor}" stroke-width="${tw.toFixed(2)}"${borderStrokeDasharray(borderFill.topBorderType)}/>`);
    }
    // Bottom border
    const bw = parseBorderWidth(borderFill.bottomBorderWidth);
    if (bw > 0 && !isBorderNone(borderFill.bottomBorderType)) {
      parts.push(`<line x1="${x.toFixed(2)}" y1="${(y + h).toFixed(2)}" x2="${(x + w).toFixed(2)}" y2="${(y + h).toFixed(2)}" stroke="${borderFill.bottomBorderColor}" stroke-width="${bw.toFixed(2)}"${borderStrokeDasharray(borderFill.bottomBorderType)}/>`);
    }
  } else {
    // Fallback: simple black border
    parts.push(`<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="none" stroke="#000" stroke-width="0.1"/>`);
  }

  return parts;
}

// ── Grid-line detection for table borders ──

/** A resolved border spec (non-NONE, positive width). */
interface BSpec { stroke: string; width: number; type: string }

/** Get one side's border spec from a borderFill, or null if NONE/absent. */
function bfSide(bf: BorderFillInfo | undefined, side: 'left' | 'right' | 'top' | 'bottom'): BSpec | null {
  if (!bf) return null;
  const type  = bf[`${side}BorderType`  as keyof BorderFillInfo] as string;
  const wStr  = bf[`${side}BorderWidth` as keyof BorderFillInfo] as string;
  const color = bf[`${side}BorderColor` as keyof BorderFillInfo] as string;
  const width = parseBorderWidth(wStr);
  if (width <= 0 || isBorderNone(type)) return null;
  return { stroke: color, width, type };
}

/** Return the stronger of two border specs (larger width wins; non-null beats null). */
function mergeSpec(a: BSpec | null, b: BSpec | null): BSpec | null {
  if (!a) return b;
  if (!b) return a;
  return a.width >= b.width ? a : b;
}

/** Round coordinate to 3 decimal places for map keys. */
function coord(n: number): string { return n.toFixed(3); }

/** Cell geometry + border fill, used for grid-line collection. */
interface CellGeo { x: number; y: number; w: number; h: number; bf: BorderFillInfo | undefined }

/**
 * Build H-line and V-line maps from all cell borders.
 *
 * Keys:
 *   H-lines → "y|x1|x2"   (y constant, segment from x1 to x2)
 *   V-lines → "x|y1|y2"   (x constant, segment from y1 to y2)
 *
 * For each position the strongest (widest) non-NONE spec wins.
 * The table's borderFillIDRef is the per-cell default, NOT an outer-boundary box,
 * so it is intentionally NOT applied as a fallback here.
 */
function buildGridLines(
  cells: CellGeo[],
  _tableBf: BorderFillInfo | undefined,
): { hLines: Map<string, BSpec | null>; vLines: Map<string, BSpec | null> } {
  const hLines = new Map<string, BSpec | null>();
  const vLines = new Map<string, BSpec | null>();

  function addH(y: number, x1: number, x2: number, spec: BSpec | null) {
    const key = `${coord(y)}|${coord(Math.min(x1, x2))}|${coord(Math.max(x1, x2))}`;
    hLines.set(key, mergeSpec(hLines.get(key) ?? null, spec));
  }
  function addV(x: number, y1: number, y2: number, spec: BSpec | null) {
    const key = `${coord(x)}|${coord(Math.min(y1, y2))}|${coord(Math.max(y1, y2))}`;
    vLines.set(key, mergeSpec(vLines.get(key) ?? null, spec));
  }

  for (const { x, y, w, h, bf } of cells) {
    addH(y,     x, x + w, bfSide(bf, 'top'));
    addH(y + h, x, x + w, bfSide(bf, 'bottom'));
    addV(x,     y, y + h, bfSide(bf, 'left'));
    addV(x + w, y, y + h, bfSide(bf, 'right'));
  }

  return { hLines, vLines };
}

/**
 * Emit SVG line element(s) for a border segment.
 *
 * DOUBLE borders produce two parallel thin lines in a 1:2:1 ratio:
 * each line is W/4 wide, gap is W/2, centers at ±3W/8 from nominal edge.
 * All other types emit a single <line> using borderStrokeDasharray for style.
 */
function emitLine(
  parts: string[],
  x1: number, y1: number, x2: number, y2: number,
  spec: BSpec,
): void {
  const isH = y1 === y2; // horizontal line (y constant)
  if (spec.type.toUpperCase() === 'DOUBLE') {
    // Each line W/3 thick, gap = W, centers at ±2W/3; total span = 5W/3
    const lineW  = spec.width / 3;
    const offset = spec.width * 2 / 3;
    if (isH) {
      parts.push(`<line x1="${x1.toFixed(2)}" y1="${(y1 - offset).toFixed(2)}" x2="${x2.toFixed(2)}" y2="${(y1 - offset).toFixed(2)}" stroke="${spec.stroke}" stroke-width="${lineW.toFixed(2)}"/>`);
      parts.push(`<line x1="${x1.toFixed(2)}" y1="${(y1 + offset).toFixed(2)}" x2="${x2.toFixed(2)}" y2="${(y1 + offset).toFixed(2)}" stroke="${spec.stroke}" stroke-width="${lineW.toFixed(2)}"/>`);
    } else {
      parts.push(`<line x1="${(x1 - offset).toFixed(2)}" y1="${y1.toFixed(2)}" x2="${(x1 - offset).toFixed(2)}" y2="${y2.toFixed(2)}" stroke="${spec.stroke}" stroke-width="${lineW.toFixed(2)}"/>`);
      parts.push(`<line x1="${(x1 + offset).toFixed(2)}" y1="${y1.toFixed(2)}" x2="${(x1 + offset).toFixed(2)}" y2="${y2.toFixed(2)}" stroke="${spec.stroke}" stroke-width="${lineW.toFixed(2)}"/>`);
    }
  } else {
    const da = borderStrokeDasharray(spec.type);
    parts.push(`<line x1="${x1.toFixed(2)}" y1="${y1.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y2.toFixed(2)}" stroke="${spec.stroke}" stroke-width="${spec.width.toFixed(2)}"${da}/>`);
  }
}

/** Render all non-null H/V line segments from buildGridLines. */
function renderGridLines(
  hLines: Map<string, BSpec | null>,
  vLines: Map<string, BSpec | null>,
): string[] {
  const parts: string[] = [];
  for (const [key, spec] of hLines) {
    if (!spec) continue;
    const [y, x1, x2] = key.split('|').map(Number);
    emitLine(parts, x1, y, x2, y, spec);
  }
  for (const [key, spec] of vLines) {
    if (!spec) continue;
    const [x, y1, y2] = key.split('|').map(Number);
    emitLine(parts, x, y1, x, y2, spec);
  }
  return parts;
}

/**
 * Build per-column widths by scanning ALL rows for cells with colSpan=1.
 * This avoids the bug where picking the "widest row" gives wrong widths
 * when that row contains merged (colSpan > 1) cells.
 */
export function buildColWidthsFromRows(rows: Element[]): Map<number, number> {
  const colWidths = new Map<number, number>(); // colAddr → width in HWPUNIT

  // First pass: collect widths from unmerged cells (colSpan === 1)
  for (const rowEl of rows) {
    for (const cellEl of findDirectCells(rowEl)) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const szEl = findChild(cellEl, 'cellSz');
      const wHu = szEl ? intAttr(szEl, 'width', 0) : intAttr(cellEl, 'width', 0);
      if (colSpan === 1 && wHu > 0) {
        colWidths.set(colAddr, wHu);
      }
    }
  }

  // Iterative derivation: repeat until no more columns can be inferred.
  // Each pass may unlock new derivations (e.g. col11 known → col10 derivable → col8 derivable).
  let progress = true;
  while (progress) {
    progress = false;
    for (const rowEl of rows) {
      for (const cellEl of findDirectCells(rowEl)) {
        const addrEl = findChild(cellEl, 'cellAddr');
        const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
        const spanEl = findChild(cellEl, 'cellSpan');
        const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
        if (colSpan <= 1) continue;
        const szEl = findChild(cellEl, 'cellSz');
        const totalW = szEl ? intAttr(szEl, 'width', 0) : intAttr(cellEl, 'width', 0);
        if (totalW <= 0) continue;

        // Count how many spanned columns are known vs unknown
        let knownSum = 0;
        let unknownCol = -1;
        let unknownCount = 0;
        for (let c = colAddr; c < colAddr + colSpan; c++) {
          if (colWidths.has(c)) {
            knownSum += colWidths.get(c)!;
          } else {
            unknownCol = c;
            unknownCount++;
          }
        }
        if (unknownCount === 1 && unknownCol >= 0) {
          const derived = totalW - knownSum;
          if (derived > 0 && !colWidths.has(unknownCol)) {
            colWidths.set(unknownCol, derived);
            progress = true;
          }
        }
      }
    }
  }

  return colWidths;
}

/**
 * Build column x-offsets (in mm) from colWidths (in HWPUNIT).
 * Returns offsets map and widths-in-mm map.
 */
export function buildColOffsetsFromWidths(colWidths: Map<number, number>): { offsets: Map<number, number>; widthsMm: Map<number, number> } {
  const sorted = [...colWidths.entries()].sort((a, b) => a[0] - b[0]);
  const offsets = new Map<number, number>();
  const widthsMm = new Map<number, number>();
  let x = 0;
  for (const [col, wHu] of sorted) {
    offsets.set(col, x);
    const wMm = hu2mm(wHu);
    widthsMm.set(col, wMm);
    x += wMm;
  }
  return { offsets, widthsMm };
}

/**
 * Distribute any shortfall between sum of row heights and declared tbl height.
 */
function distributeDeclaredTableHeight(
  rowHeights: Map<number, number>,
  declaredRowHeights: Map<number, number>,
  tblHeightHu: number,
): void {
  if (tblHeightHu <= 0 || rowHeights.size === 0) return;
  const tblHeightMm = hu2mm(tblHeightHu);
  let sum = 0;
  for (const h of rowHeights.values()) sum += h;
  if (sum >= tblHeightMm) return;
  const shortfall = tblHeightMm - sum;
  const SMALL_ROW_THRESHOLD = 5;
  let largeDeclaredSum = 0;
  for (const [k, declared] of declaredRowHeights) {
    if (declared >= SMALL_ROW_THRESHOLD && rowHeights.has(k)) {
      largeDeclaredSum += declared;
    }
  }
  if (largeDeclaredSum > 0) {
    for (const [k, v] of rowHeights) {
      const declared = declaredRowHeights.get(k) ?? 0;
      if (declared >= SMALL_ROW_THRESHOLD) {
        rowHeights.set(k, v + shortfall * (declared / largeDeclaredSum));
      }
    }
  } else {
    const scale = tblHeightMm / sum;
    for (const [k, v] of rowHeights) {
      rowHeights.set(k, v * scale);
    }
  }
}

export function adjustRowHeights(
  rows: Element[],
  rowHeights: Map<number, number>,
  colWidths: Map<number, number>,
): void {
  for (const rowEl of rows) {
    const cells = findDirectCells(rowEl);
    for (const cellEl of cells) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const rowAddr = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(cellEl, 'rowAddr', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);

      if (rowSpan !== 1) continue;

      const subList = find(cellEl, 'subList');
      const cellParas = subList ? children(subList, 'p') : children(cellEl, 'p');

      // Content height: bottom of the last lineseg (vertpos + vertsize).
      let linesegContentH = 0;
      for (const paraEl of cellParas) {
        for (const lsa of children(paraEl, 'linesegarray')) {
          for (const seg of children(lsa, 'lineseg')) {
            const bottom = intAttr(seg, 'vertpos', 0) + intAttr(seg, 'vertsize', 0);
            if (bottom > linesegContentH) linesegContentH = bottom;
          }
        }
      }
      if (linesegContentH === 0) continue;

      // Cell margins
      const marginEl = findChild(cellEl, 'cellMargin');
      const cellMTop = marginEl ? hu2mm(intAttr(marginEl, 'top', 141)) : hu2mm(intAttr(cellEl, 'marginTop', 141));
      const cellMBottom = marginEl ? hu2mm(intAttr(marginEl, 'bottom', 141)) : hu2mm(intAttr(cellEl, 'marginBottom', 141));

      const neededH = cellMTop + hu2mm(linesegContentH) + cellMBottom;
      const currentH = rowHeights.get(rowAddr) ?? 0;

      if (neededH > currentH) {
        rowHeights.set(rowAddr, neededH);
      }
    }
  }
}

// ── Cell content rendering ──

/**
 * Render the text content of a table cell. Returns SVG elements.
 * Handles text wrapping, vertical centering in merged cells, and clipping.
 */
export function renderCellContent(
  cellEl: Element,
  caches: Caches,
  x: number,
  y: number,
  cellW: number,
  cellH: number,
  _declaredCellH?: number,
): string[] {
  // Use declared height for centering when the row didn't grow for content
  const szElCenter = findChild(cellEl, 'cellSz');
  const rawDeclaredH = szElCenter ? hu2mm(intAttr(szElCenter, 'height', 0)) : hu2mm(intAttr(cellEl, 'height', 0));
  const spanElCenter = findChild(cellEl, 'cellSpan');
  const rowSpanCenter = spanElCenter ? intAttr(spanElCenter, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
  let centerCellH = cellH;
  if (rowSpanCenter === 1 && rawDeclaredH > 0 && cellH > rawDeclaredH * 1.05) {
    centerCellH = cellH;
  } else if (rowSpanCenter === 1 && rawDeclaredH > 0) {
    centerCellH = rawDeclaredH;
  }
  const parts: string[] = [];

  // Cell margins (child element or flat attributes)
  const marginEl = findChild(cellEl, 'cellMargin');
  const mLeft = marginEl ? hu2mm(intAttr(marginEl, 'left', 141)) : hu2mm(intAttr(cellEl, 'marginLeft', 141));
  const mRight = marginEl ? hu2mm(intAttr(marginEl, 'right', 141)) : hu2mm(intAttr(cellEl, 'marginRight', 141));
  const mTop = marginEl ? hu2mm(intAttr(marginEl, 'top', 141)) : hu2mm(intAttr(cellEl, 'marginTop', 141));

  const textX = x + mLeft;
  const textWidth = cellW - mLeft - mRight;
  if (textWidth <= 0) return parts;

  // Collect all paragraphs and vertical alignment
  const subList = find(cellEl, 'subList');
  const cellParas = subList ? children(subList, 'p') : children(cellEl, 'p');
  const vertAlignStr = subList ? attr(subList, 'vertAlign', 'TOP').toUpperCase() : 'TOP';
  const vertAlignMode = vertAlignStr === 'CENTER' ? 1 : vertAlignStr === 'BOTTOM' ? 2 : 0;

  // Cell bottom margin
  const mBottom = marginEl ? hu2mm(intAttr(marginEl, 'bottom', 141)) : hu2mm(intAttr(cellEl, 'marginBottom', 141));

  // First pass: compute total content height for vertical centering
  interface RunSegment {
    text: string;
    fontSize: number;
    fontFamily: string;
    color: string;
    fw: string;
    fi: string;
    ls: string; // letter-spacing attr string
  }
  interface LineInfo {
    segments: RunSegment[];
    fontSize: number;  // primary font size for line height computation
    lineSpacing: number; // line spacing ratio (e.g. 1.6 for 160%)
    anchor: string;
    tx: number;
    vertposMm?: number;  // cell-relative Y offset of line top (from lineseg)
    baselineMm?: number; // baseline offset within line (from lineseg)
    horzsizeMm?: number; // horizontal size from lineseg (for textLength compression)
    horzposMm?: number;  // horizontal offset from lineseg (indent/margin)
    vertsizeMm?: number; // vertical content height from lineseg (line box)
    spacingMm?: number;  // extra inter-line spacing (from lineseg)
    isLastLine?: boolean; // true if this is the last line of its paragraph
    justify?: boolean;   // paragraph uses JUSTIFY horizontal alignment
    distribute?: boolean; // paragraph uses DISTRIBUTE alignment (spread across horzsizeMm)
  }
  // Content items: either text lines or nested tables, in paragraph order
  type ContentItem = { kind: 'line'; line: LineInfo } | { kind: 'table'; tbl: Element; heightMm: number; outMarginTopMm: number; outMarginBottomMm: number; outMarginLeftMm: number; outMarginRightMm: number; vertposMm?: number };
  const contentItems: ContentItem[] = [];
  const allLines: LineInfo[] = [];
  let totalContentH = mTop + mBottom; // start with vertical margins
  let maxLineBottomMm = 0;
  let anyLinesegData = false;

  for (const paraEl of cellParas) {
    const paraPrIdRef = attr(paraEl, 'paraPrIDRef', attr(paraEl, 'paraPrId', ''));
    const paraShape = paraPrIdRef ? caches.paraShapes.get(paraPrIdRef) : undefined;
    const runs = collectRuns(paraEl);
    const fullText = runs.map(r => r.text).join('');

    const firstRun = runs[0];
    const cs = firstRun ? caches.charShapes.get(firstRun.charPrId) : undefined;
    const fontSize = cs ? hu2mm(cs.height) : 3.5;

    let anchor = 'start';
    let tx = textX;
    let bodyIndentMm = 0;
    let firstLineIndentMm = 0;
    if (paraShape) {
      const align = paraShape.alignment;
      if (align === 3) { anchor = 'middle'; tx = x + cellW / 2; }
      else if (align === 2) { anchor = 'end'; tx = x + cellW - mRight; }

      if (anchor === 'start') {
        tx += hu2mm(paraShape.leftMargin);
        if (paraShape.indent < 0) {
          // 내어쓰기 (hanging indent): body lines indented right by |indent|
          bodyIndentMm = hu2mm(-paraShape.indent);
        } else if (paraShape.indent > 0) {
          firstLineIndentMm = hu2mm(paraShape.indent);
        }
      }
    }

    const paraLineSpacing = paraShape ? paraShape.lineSpacing / 100 : 1.6;

    if (!fullText.trim()) {
      const linesegArray = children(paraEl, 'linesegarray')[0];
      const firstSeg = linesegArray ? children(linesegArray, 'lineseg')[0] : null;
      const segVertpos = firstSeg ? hu2mm(intAttr(firstSeg, 'vertpos', 0)) : 0;
      const segVertsize = firstSeg ? hu2mm(intAttr(firstSeg, 'vertsize', 0)) : 0;
      if (segVertsize > 0) {
        anyLinesegData = true;
        const bottom = segVertpos + segVertsize;
        if (bottom > maxLineBottomMm) maxLineBottomMm = bottom;
        totalContentH += segVertsize;
      } else {
        totalContentH += fontSize * paraLineSpacing;
      }
      const emptyLine: LineInfo = { segments: [], fontSize, lineSpacing: paraLineSpacing, anchor, tx, vertposMm: segVertpos, vertsizeMm: segVertsize > 0 ? segVertsize : undefined };
      allLines.push(emptyLine);
      contentItems.push({ kind: 'line', line: emptyLine });

      const paraNestedTables = findDirectTables(paraEl);
      for (const nestedTbl of paraNestedTables) {
        const szEl = find(nestedTbl, 'sz');
        const nestedH = szEl ? hu2mm(intAttr(szEl, 'height', 0)) : 0;
        const omTop = hu2mm(intAttr(nestedTbl, 'outMarginTop', 0));
        const omBottom = hu2mm(intAttr(nestedTbl, 'outMarginBottom', 0));
        const omLeft = hu2mm(intAttr(nestedTbl, 'outMarginLeft', 0));
        const omRight = hu2mm(intAttr(nestedTbl, 'outMarginRight', 0));
        contentItems.push({ kind: 'table', tbl: nestedTbl, heightMm: nestedH, outMarginTopMm: omTop, outMarginBottomMm: omBottom, outMarginLeftMm: omLeft, outMarginRightMm: omRight, vertposMm: segVertpos });
      }
      continue;
    }

    // Build a character-to-run mapping for per-character styling
    const charRunMap: number[] = [];
    for (let ri = 0; ri < runs.length; ri++) {
      for (let ci = 0; ci < runs[ri].text.length; ci++) {
        charRunMap.push(ri);
      }
    }

    // Split fullText into lines using lineseg or newlines
    const linesegArray = children(paraEl, 'linesegarray')[0];
    const lineSegs = linesegArray ? children(linesegArray, 'lineseg') : [];

    interface LineRange { start: number; end: number; vertposMm?: number; baselineMm?: number; horzsizeMm?: number; horzposMm?: number; vertsizeMm?: number; spacingMm?: number; }
    let lineRanges: LineRange[];
    if (lineSegs.length > 1) {
      const textPositions: number[] = lineSegs.map(seg => intAttr(seg, 'textpos', 0));
      lineRanges = [];
      for (let si = 0; si < textPositions.length; si++) {
        const start = textPositions[si];
        const end = si + 1 < textPositions.length ? textPositions[si + 1] : fullText.length;
        if (fullText.substring(start, end) || si === 0) {
          const vertpos = intAttr(lineSegs[si], 'vertpos', 0);
          const baseline = intAttr(lineSegs[si], 'baseline', 0);
          const horzsize = intAttr(lineSegs[si], 'horzsize', 0);
          const horzpos = intAttr(lineSegs[si], 'horzpos', 0);
          const vertsize = intAttr(lineSegs[si], 'vertsize', 0);
          const spacing = intAttr(lineSegs[si], 'spacing', 0);
          lineRanges.push({ start, end, vertposMm: hu2mm(vertpos), baselineMm: hu2mm(baseline), horzsizeMm: horzsize > 0 ? hu2mm(horzsize) : undefined, horzposMm: horzpos > 0 ? hu2mm(horzpos) : undefined, vertsizeMm: vertsize > 0 ? hu2mm(vertsize) : undefined, spacingMm: spacing > 0 ? hu2mm(spacing) : undefined });
        }
      }
    } else {
      const singleSegVertpos = lineSegs.length === 1 ? hu2mm(intAttr(lineSegs[0], 'vertpos', 0)) : undefined;
      const singleSegBaseline = lineSegs.length === 1 ? hu2mm(intAttr(lineSegs[0], 'baseline', 0)) : undefined;
      const singleSegHorzsize = lineSegs.length === 1 ? intAttr(lineSegs[0], 'horzsize', 0) : 0;
      const singleSegHorzsizeMm = singleSegHorzsize > 0 ? hu2mm(singleSegHorzsize) : undefined;
      const singleSegHorzpos = lineSegs.length === 1 ? intAttr(lineSegs[0], 'horzpos', 0) : 0;
      const singleSegHorzposMm = singleSegHorzpos > 0 ? hu2mm(singleSegHorzpos) : undefined;
      const singleSegVertsize = lineSegs.length === 1 ? intAttr(lineSegs[0], 'vertsize', 0) : 0;
      const singleSegVertsizeMm = singleSegVertsize > 0 ? hu2mm(singleSegVertsize) : undefined;

      lineRanges = [];
      let offset = 0;
      const rawLines = fullText.split('\n');
      for (let li = 0; li < rawLines.length; li++) {
        const rawLine = rawLines[li];
        let wrapped: string[];
        if (lineSegs.length === 1 && rawLines.length === 1) {
          wrapped = [rawLine];
        } else {
          wrapped = wrapText(rawLine, textWidth, fontSize);
        }
        for (let wi = 0; wi < wrapped.length; wi++) {
          const applyHorzsize = lineSegs.length === 1 && rawLines.length === 1 && wrapped.length === 1 && wi === 0;
          lineRanges.push({
            start: offset,
            end: offset + wrapped[wi].length,
            vertposMm: (li === 0 && wi === 0) ? singleSegVertpos : undefined,
            baselineMm: (li === 0 && wi === 0) ? singleSegBaseline : undefined,
            horzsizeMm: applyHorzsize ? singleSegHorzsizeMm : undefined,
            horzposMm: (li === 0 && wi === 0) ? singleSegHorzposMm : undefined,
            vertsizeMm: (li === 0 && wi === 0) ? singleSegVertsizeMm : undefined,
          });
          offset += wrapped[wi].length;
        }
        offset++; // skip the \n
      }
    }

    // For each line, build segments grouped by run styling AND script type
    for (const range of lineRanges) {
      const segments: RunSegment[] = [];
      let segStart = range.start;
      while (segStart < range.end) {
        const ri = charRunMap[segStart] ?? 0;
        let runEnd = segStart + 1;
        while (runEnd < range.end && (charRunMap[runEnd] ?? 0) === ri) {
          runEnd++;
        }
        const rcs = caches.charShapes.get(runs[ri]?.charPrId ?? '');
        let scriptStart = segStart;
        while (scriptStart < runEnd) {
          const ch = fullText[scriptStart];
          const isCjk = ch ? isCJK(ch) : false;
          let scriptEnd = scriptStart + 1;
          while (scriptEnd < runEnd) {
            const nextCh = fullText[scriptEnd];
            if (nextCh === ' ') { scriptEnd++; continue; }
            if (isCJK(nextCh) !== isCjk) break;
            scriptEnd++;
          }
          const segText = fullText.substring(scriptStart, scriptEnd);
          if (segText === '\n') { scriptStart = scriptEnd; continue; } // lineseg boundary marker, not display content
          const sampleCh = segText.trim()[0] || segText[0] || '';
          segments.push({
            text: segText,
            fontSize: rcs ? hu2mm(rcs.height) : 3.5,
            fontFamily: rcs ? fontForChar(rcs, sampleCh) : 'sans-serif',
            color: rcs?.textColor || '#000000',
            fw: rcs?.bold ? 'bold' : 'normal',
            fi: rcs?.italic ? 'italic' : 'normal',
            ls: letterSpacingAttr(rcs, sampleCh),
          });
          scriptStart = scriptEnd;
        }
        segStart = runEnd;
      }
      const lineFontSize = segments.length > 0 ? Math.max(...segments.map(s => s.fontSize)) : fontSize;
      const isFirstLine = range === lineRanges[0];
      let lineTx = tx;
      if (anchor === 'start') {
        if (isFirstLine) {
          lineTx += firstLineIndentMm;
        } else {
          lineTx += bodyIndentMm;
        }
      }
      const rangeVertsizeMm = range.vertsizeMm;
      const rangeSpacingMm = range.spacingMm;
      const isLast = range === lineRanges[lineRanges.length - 1];
      const isJustify = paraShape?.alignment === 0;
      const isDistribute = paraShape?.alignment === 4;
      const lineInfo: LineInfo = { segments, fontSize: lineFontSize, lineSpacing: paraLineSpacing, anchor, tx: lineTx, vertposMm: range.vertposMm, baselineMm: range.baselineMm, horzsizeMm: range.horzsizeMm, horzposMm: range.horzposMm, vertsizeMm: rangeVertsizeMm, spacingMm: rangeSpacingMm, isLastLine: isLast, justify: isJustify, distribute: isDistribute };
      allLines.push(lineInfo);
      contentItems.push({ kind: 'line', line: lineInfo });
      if (rangeVertsizeMm !== undefined && rangeVertsizeMm > 0) {
        totalContentH += rangeVertsizeMm + (rangeSpacingMm ?? 0);
        anyLinesegData = true;
        const bottom = (range.vertposMm ?? 0) + rangeVertsizeMm;
        if (bottom > maxLineBottomMm) maxLineBottomMm = bottom;
      } else {
        totalContentH += lineFontSize * paraLineSpacing;
      }
    }

    // After text lines for this paragraph, add any nested tables.
    const paraNestedTables = findDirectTables(paraEl);
    if (paraNestedTables.length > 0) {
      const linesegArrayNested = children(paraEl, 'linesegarray')[0];
      const firstSeg = linesegArrayNested ? children(linesegArrayNested, 'lineseg')[0] : null;
      const paraLineTopMm = firstSeg ? hu2mm(intAttr(firstSeg, 'vertpos', 0)) : undefined;
      for (const nestedTbl of paraNestedTables) {
        const szEl = find(nestedTbl, 'sz');
        const nestedH = szEl ? hu2mm(intAttr(szEl, 'height', 0)) : 0;
        const omTop = hu2mm(intAttr(nestedTbl, 'outMarginTop', 0));
        const omBottom = hu2mm(intAttr(nestedTbl, 'outMarginBottom', 0));
        const omLeft = hu2mm(intAttr(nestedTbl, 'outMarginLeft', 0));
        const omRight = hu2mm(intAttr(nestedTbl, 'outMarginRight', 0));
        contentItems.push({ kind: 'table', tbl: nestedTbl, heightMm: nestedH, outMarginTopMm: omTop, outMarginBottomMm: omBottom, outMarginLeftMm: omLeft, outMarginRightMm: omRight, vertposMm: paraLineTopMm });
        totalContentH += nestedH + omTop + omBottom;
      }
    }
  }

  // Collect all nested tables for counting (used in vert align heuristic)
  const nestedTables: Element[] = [];
  for (const paraEl of cellParas) {
    nestedTables.push(...findDirectTables(paraEl));
  }

  if (allLines.length === 0 && nestedTables.length === 0) return parts;

  // Add clip path — expand if lineseg content exceeds declared cell height
  const clipId = `cell-clip-${_clipIdCounter++}`;
  let clipH = cellH;
  if (totalContentH > cellH) {
    clipH = totalContentH + hu2mm(282); // add small padding
  }
  // Expand clip rect left by mLeft so right-aligned text that is slightly wider
  // than the content area isn't clipped — mirrors the mRight buffer on the right.
  const clipX = x - mLeft;
  const clipW = cellW + mLeft;
  parts.push(`<clipPath id="${clipId}"><rect x="${clipX.toFixed(2)}" y="${y.toFixed(2)}" width="${clipW.toFixed(2)}" height="${clipH.toFixed(2)}"/></clipPath>`);

  const textOnlyH = anyLinesegData ? maxLineBottomMm : totalContentH - mTop - mBottom;
  let effectiveVertAlign = vertAlignMode;
  if (effectiveVertAlign === 0 && textOnlyH > 0 && textOnlyH < centerCellH) {
    effectiveVertAlign = 1; // CENTER
  }

  let vertOffset = 0;
  if (effectiveVertAlign === 1) {
    vertOffset = Math.max(0, (centerCellH - mTop - mBottom - textOnlyH) / 2);
  } else if (effectiveVertAlign === 2) {
    vertOffset = Math.max(0, centerCellH - mTop - mBottom - textOnlyH);
  }

  let textY = y + mTop + vertOffset;

  // Second pass: render lines
  const cellContentY = y + mTop + vertOffset;
  let useLinesegPos = true;

  parts.push(`<g clip-path="url(#${clipId})">`);
  let isFirstRenderedLine = true;
  for (const item of contentItems) {
    if (item.kind === 'table') {
      let tableY = textY;
      if (item.vertposMm !== undefined) {
        tableY = cellContentY + item.vertposMm + item.outMarginTopMm;
      }
      const tableX = textX + item.outMarginLeftMm;
      const nestedResult = renderTableEl(item.tbl, caches, tableX, tableY, textWidth - item.outMarginLeftMm - item.outMarginRightMm);
      if (nestedResult) {
        textY = tableY;
        parts.push(nestedResult.svg);
        textY += nestedResult.height;
        isFirstRenderedLine = true; // reset: next text line starts a fresh block
      }
      continue;
    }
    const line = item.line;
    if (line.segments.length === 0) {
      if (!useLinesegPos && line.vertsizeMm !== undefined && line.vertsizeMm > 0) {
        textY += line.vertsizeMm;
      } else if (useLinesegPos && line.vertposMm !== undefined) {
        const linesegY = cellContentY + line.vertposMm + (line.baselineMm ?? line.fontSize);
        textY = Math.max(textY, linesegY);
      } else {
        textY += line.fontSize * line.lineSpacing;
      }
      continue;
    }
    if (useLinesegPos && line.vertposMm !== undefined && line.baselineMm !== undefined) {
      const linesegY = cellContentY + line.vertposMm + line.baselineMm;
      textY = Math.max(textY, linesegY);
    } else if (isFirstRenderedLine && line.baselineMm !== undefined && line.baselineMm > 0) {
      textY += line.baselineMm;
    } else if (line.vertsizeMm !== undefined && line.vertsizeMm > 0) {
      textY += line.vertsizeMm + (line.spacingMm ?? 0);
    } else {
      textY += line.fontSize;
    }
    isFirstRenderedLine = false;
    let textLengthAttr = '';
    if (line.horzsizeMm !== undefined && line.horzsizeMm > 0) {
      const lineText = line.segments.map(s => s.text).join('');
      const naturalWidth = estimateTextWidth(lineText, line.fontSize);
      // JUSTIFY: stretch non-last lines to fill horzsize (spacing only, glyphs stay natural)
      // DISTRIBUTE: always spread text across the full horzsize
      // Compression (naturalWidth > horzsize): scale both spacing and glyphs
      if ((line.justify && !line.isLastLine || line.distribute) && naturalWidth < line.horzsizeMm && lineText.trim().length > 1) {
        textLengthAttr = ` textLength="${line.horzsizeMm.toFixed(2)}" lengthAdjust="spacing"`;
      } else if (naturalWidth > line.horzsizeMm * 1.02) {
        textLengthAttr = ` textLength="${line.horzsizeMm.toFixed(2)}" lengthAdjust="spacingAndGlyphs"`;
      }
    }
    // When lineseg horzpos is set, use it directly as the text start (incorporates leftMargin).
    // Only apply horzpos offset for start-anchored text; center/end anchors are already
    // positioned correctly by tx (x+cellW/2 or x+cellW-mRight) and must not be shifted further.
    const lineTx = (line.horzposMm !== undefined && line.anchor === 'start')
      ? textX + line.horzposMm
      : line.tx;
    if (line.segments.length === 1) {
      const seg = line.segments[0];
      parts.push(`<text xml:space="preserve" x="${lineTx.toFixed(2)}" y="${textY.toFixed(2)}" font-size="${seg.fontSize.toFixed(2)}" font-family="${escapeXml(fontFamilyWithFallback(seg.fontFamily))}" fill="${seg.color}" font-weight="${seg.fw}" font-style="${seg.fi}"${seg.ls}${textLengthAttr} text-anchor="${line.anchor}">${escapeXml(seg.text)}</text>`);
    } else {
      const tspans = line.segments.map(seg =>
        `<tspan font-size="${seg.fontSize.toFixed(2)}" font-family="${escapeXml(fontFamilyWithFallback(seg.fontFamily))}" fill="${seg.color}" font-weight="${seg.fw}" font-style="${seg.fi}"${seg.ls}>${escapeXml(seg.text)}</tspan>`
      ).join('');
      parts.push(`<text xml:space="preserve" x="${lineTx.toFixed(2)}" y="${textY.toFixed(2)}"${textLengthAttr} text-anchor="${line.anchor}">${tspans}</text>`);
    }
    if (!line.vertsizeMm) {
      textY += line.fontSize * (line.lineSpacing - 1);
    }
  }

  parts.push('</g>');

  return parts;
}

/**
 * Render a table with page-break support.
 */
export function renderTableWithPageBreaks(
  tblEl: Element,
  caches: Caches,
  dims: PageDims,
  pages: LayoutItem[][],
  currentPage: number,
  yPos: number,
): { yPos: number; currentPage: number } {
  const rows = children(tblEl, 'tr');
  if (rows.length === 0) return { yPos, currentPage };

  const outMarginTop = hu2mm(intAttr(tblEl, 'outMarginTop', 0));
  const outMarginBottom = hu2mm(intAttr(tblEl, 'outMarginBottom', 0));
  yPos += outMarginTop;

  const innerML = hu2mm(intAttr(tblEl, 'innerMarginLeft', 0));
  const innerMR = hu2mm(intAttr(tblEl, 'innerMarginRight', 0));

  const rowMeta: Array<{ rowAddr: number; cells: Element[]; height: number }> = [];
  const rowHeights = new Map<number, number>();
  const renderedRows = new Set<number>();

  for (const rowEl of rows) {
    const cells = findDirectCells(rowEl);
    if (cells.length === 0) continue;
    const firstCell = cells[0];
    const addrEl = findChild(firstCell, 'cellAddr');
    const rowAddr = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(firstCell, 'rowAddr', 0);
    let maxH = 0;
    for (const cellEl of cells) {
      const szEl2 = findChild(cellEl, 'cellSz');
      const cellH = szEl2 ? intAttr(szEl2, 'height', 0) : intAttr(cellEl, 'height', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
      const perRow = rowSpan > 1 ? cellH / rowSpan : cellH;
      if (perRow > maxH) maxH = perRow;
    }
    const rowH = hu2mm(maxH);
    if (!rowHeights.has(rowAddr) || rowHeights.get(rowAddr)! < rowH) {
      rowHeights.set(rowAddr, rowH);
    }
    rowMeta.push({ rowAddr, cells, height: rowH });
  }

  const colWidthsHu = buildColWidthsFromRows(rows);
  const { offsets: colOffsets, widthsMm: colWidths } = buildColOffsetsFromWidths(colWidthsHu);

  adjustRowHeights(rows, rowHeights, colWidthsHu);

  for (const rm of rowMeta) {
    rm.height = rowHeights.get(rm.rowAddr) ?? rm.height;
  }

  const coveredCells = new Set<string>();
  for (const { cells } of rowMeta) {
    for (const cellEl of cells) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const rowAddrC = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(cellEl, 'rowAddr', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
      if (colSpan > 1 || rowSpan > 1) {
        for (let r = rowAddrC; r < rowAddrC + rowSpan; r++) {
          for (let c = colAddr; c < colAddr + colSpan; c++) {
            if (r !== rowAddrC || c !== colAddr) coveredCells.add(`${r},${c}`);
          }
        }
      }
    }
  }

  const tableBfIdRef = attr(tblEl, 'borderFillIDRef', '');
  const tableBf = tableBfIdRef ? caches.borderFills.get(tableBfIdRef) : undefined;

  // Per-page segment: accumulate cells, flush when page breaks or table ends
  interface PCell { x: number; y: number; w: number; h: number; bf: BorderFillInfo | undefined; el: Element; declaredH: number }
  let segmentCells: PCell[] = [];
  let segmentStartY = yPos;

  function flushSegment(parts: string[], segY: number): void {
    if (segmentCells.length === 0) return;
    // Phase 1: fills
    for (const pc of segmentCells) {
      if (pc.bf?.fillColor) {
        parts.push(`<rect x="${pc.x.toFixed(2)}" y="${pc.y.toFixed(2)}" width="${pc.w.toFixed(2)}" height="${pc.h.toFixed(2)}" fill="${pc.bf.fillColor}" stroke="none"/>`);
      }
    }
    // Phase 2: grid-detected border lines
    const { hLines, vLines } = buildGridLines(segmentCells.map(pc => ({ x: pc.x, y: pc.y, w: pc.w, h: pc.h, bf: pc.bf })), tableBf);
    parts.push(...renderGridLines(hLines, vLines));
    // Phase 3: cell content
    for (const pc of segmentCells) {
      parts.push(...renderCellContent(pc.el, caches, pc.x, pc.y, pc.w, pc.h, pc.declaredH));
    }
    void segY;
    segmentCells = [];
  }

  let pageParts: string[] = [`<g class="table">`];

  for (const { rowAddr, cells, height: rowH } of rowMeta) {
    if (renderedRows.has(rowAddr)) continue;
    renderedRows.add(rowAddr);

    const actualRowH = rowHeights.get(rowAddr) ?? rowH;

    if (yPos + actualRowH > dims.pageBottom && pages[currentPage].length > 0) {
      // Flush accumulated cells before page break
      flushSegment(pageParts, segmentStartY);
      pageParts.push('</g>');
      if (pageParts.length > 2) {
        pages[currentPage].push({ svg: pageParts.join('\n'), y: segmentStartY, height: 0 });
      }

      pages.push([]);
      currentPage = pages.length - 1;
      yPos = dims.contentTop;
      segmentStartY = yPos;
      pageParts = [`<g class="table">`];
    }

    for (const cellEl of cells) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const rowAddrCell = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(cellEl, 'rowAddr', 0);

      if (coveredCells.has(`${rowAddrCell},${colAddr}`)) continue;

      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);

      let cellW: number;
      if (colSpan > 1) {
        cellW = 0;
        for (let c = colAddr; c < colAddr + colSpan; c++) {
          cellW += colWidths.get(c) ?? 0;
        }
      } else {
        const szEl2 = findChild(cellEl, 'cellSz');
        cellW = szEl2 ? hu2mm(intAttr(szEl2, 'width', 0)) : hu2mm(intAttr(cellEl, 'width', 0));
      }

      let cellH = actualRowH;
      if (rowSpan > 1) {
        cellH = 0;
        for (let r = rowAddrCell; r < rowAddrCell + rowSpan; r++) {
          cellH += rowHeights.get(r) ?? 0;
        }
      }
      const szElCell = findChild(cellEl, 'cellSz');
      const declaredCellH = szElCell ? hu2mm(intAttr(szElCell, 'height', 0)) : hu2mm(intAttr(cellEl, 'height', 0));

      const x = dims.contentLeft + (colOffsets.get(colAddr) ?? 0);
      const bfIdRef = attr(cellEl, 'borderFillIDRef', '');
      const borderFill = bfIdRef ? caches.borderFills.get(bfIdRef) : undefined;
      segmentCells.push({ x, y: yPos, w: cellW, h: cellH, bf: borderFill, el: cellEl, declaredH: declaredCellH });
    }

    yPos += actualRowH;
  }

  // Flush final page segment
  flushSegment(pageParts, segmentStartY);
  pageParts.push('</g>');
  if (pageParts.length > 2) {
    pages[currentPage].push({ svg: pageParts.join('\n'), y: segmentStartY, height: yPos - segmentStartY });
  }

  yPos += outMarginBottom;

  return { yPos, currentPage };
}

export function renderTableEl(
  tblEl: Element,
  caches: Caches,
  startX: number,
  startY: number,
  contentWidth: number,
): TableResult | null {
  const szEl = find(tblEl, 'sz');
  const _tblWidth = szEl ? hu2mm(intAttr(szEl, 'width', 0)) : (intAttr(tblEl, 'width', 0) > 0 ? hu2mm(intAttr(tblEl, 'width', 0)) : contentWidth); void _tblWidth;
  const tblHeightHu = szEl ? intAttr(szEl, 'height', 0) : intAttr(tblEl, 'height', 0);

  const innerML = hu2mm(intAttr(tblEl, 'innerMarginLeft', 0));
  const innerMR = hu2mm(intAttr(tblEl, 'innerMarginRight', 0));
  startX += innerML;
  contentWidth -= innerML + innerMR;

  const rows = children(tblEl, 'tr');
  if (rows.length === 0) return null;

  const rowHeights = new Map<number, number>();
  for (const rowEl of rows) {
    const cells = findDirectCells(rowEl);
    if (cells.length === 0) continue;
    const firstCell = cells[0];
    const addrEl = findChild(firstCell, 'cellAddr');
    const rowAddr = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(firstCell, 'rowAddr', 0);
    let maxH = 0;
    for (const cellEl of cells) {
      const szEl2 = findChild(cellEl, 'cellSz');
      const cellH = szEl2 ? intAttr(szEl2, 'height', 0) : intAttr(cellEl, 'height', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
      const perRow = rowSpan > 1 ? cellH / rowSpan : cellH;
      if (perRow > maxH) maxH = perRow;
    }
    rowHeights.set(rowAddr, hu2mm(maxH));
  }

  const colWidthsHu = buildColWidthsFromRows(rows);
  const { offsets: colOffsets, widthsMm: colWidths } = buildColOffsetsFromWidths(colWidthsHu);

  const declaredRowHeights = new Map(rowHeights);
  adjustRowHeights(rows, rowHeights, colWidthsHu);
  void declaredRowHeights;
  void tblHeightHu;

  const coveredCells = new Set<string>();
  for (const rowEl of rows) {
    for (const cellEl of findDirectCells(rowEl)) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const rowAddrC = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(cellEl, 'rowAddr', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
      if (colSpan > 1 || rowSpan > 1) {
        for (let r = rowAddrC; r < rowAddrC + rowSpan; r++) {
          for (let c = colAddr; c < colAddr + colSpan; c++) {
            if (r !== rowAddrC || c !== colAddr) coveredCells.add(`${r},${c}`);
          }
        }
      }
    }
  }

  const tableBfIdRef = attr(tblEl, 'borderFillIDRef', '');
  const tableBf = tableBfIdRef ? caches.borderFills.get(tableBfIdRef) : undefined;

  // Collect all cell geometries + elements for two-phase rendering
  interface PCell { x: number; y: number; w: number; h: number; bf: BorderFillInfo | undefined; el: Element; declaredH: number }
  const pendingCells: PCell[] = [];

  let y = startY;
  const renderedRows = new Set<number>();

  for (const rowEl of rows) {
    const cells = findDirectCells(rowEl);
    if (cells.length === 0) continue;

    const firstCell = cells[0];
    const firstAddrEl = findChild(firstCell, 'cellAddr');
    const rowAddr = firstAddrEl ? intAttr(firstAddrEl, 'rowAddr', 0) : intAttr(firstCell, 'rowAddr', 0);
    if (renderedRows.has(rowAddr)) continue;
    renderedRows.add(rowAddr);

    const rowH = rowHeights.get(rowAddr) ?? 0;

    for (const cellEl of cells) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const rowAddrCell = addrEl ? intAttr(addrEl, 'rowAddr', 0) : intAttr(cellEl, 'rowAddr', 0);

      if (coveredCells.has(`${rowAddrCell},${colAddr}`)) continue;

      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);

      let cellW: number;
      if (colSpan > 1) {
        cellW = 0;
        for (let c = colAddr; c < colAddr + colSpan; c++) {
          cellW += colWidths.get(c) ?? 0;
        }
      } else {
        const szEl2 = findChild(cellEl, 'cellSz');
        cellW = szEl2 ? hu2mm(intAttr(szEl2, 'width', 0)) : hu2mm(intAttr(cellEl, 'width', 0));
      }

      let cellH = rowH;
      if (rowSpan > 1) {
        cellH = 0;
        for (let r = rowAddrCell; r < rowAddrCell + rowSpan; r++) {
          cellH += rowHeights.get(r) ?? 0;
        }
      }
      const szElCell2 = findChild(cellEl, 'cellSz');
      const declaredCellH = szElCell2 ? hu2mm(intAttr(szElCell2, 'height', 0)) : hu2mm(intAttr(cellEl, 'height', 0));

      const x = startX + (colOffsets.get(colAddr) ?? 0);
      const bfIdRef = attr(cellEl, 'borderFillIDRef', '');
      const borderFill = bfIdRef ? caches.borderFills.get(bfIdRef) : undefined;
      pendingCells.push({ x, y, w: cellW, h: cellH, bf: borderFill, el: cellEl, declaredH: declaredCellH });
    }

    y += rowH;
  }

  // Phase 1: all fills
  const parts: string[] = [`<g class="table">`];
  for (const pc of pendingCells) {
    if (pc.bf?.fillColor) {
      parts.push(`<rect x="${pc.x.toFixed(2)}" y="${pc.y.toFixed(2)}" width="${pc.w.toFixed(2)}" height="${pc.h.toFixed(2)}" fill="${pc.bf.fillColor}" stroke="none"/>`);
    }
  }
  // Phase 2: grid-detected border lines (each unique segment once, table outer border as fallback)
  const { hLines, vLines } = buildGridLines(pendingCells.map(pc => ({ x: pc.x, y: pc.y, w: pc.w, h: pc.h, bf: pc.bf })), tableBf);
  parts.push(...renderGridLines(hLines, vLines));
  // Phase 3: cell content (clips + text, over the borders)
  for (const pc of pendingCells) {
    parts.push(...renderCellContent(pc.el, caches, pc.x, pc.y, pc.w, pc.h, pc.declaredH));
  }

  parts.push('</g>');

  const computedHeight = y - startY;
  const declaredHeight = tblHeightHu > 0 ? hu2mm(tblHeightHu) : computedHeight;
  const totalHeight = Math.max(computedHeight, declaredHeight);
  return { svg: parts.join('\n'), height: totalHeight };
}
