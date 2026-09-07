/**
 * Table rendering for the SVG renderer.
 */

import type { Caches, BorderFillInfo, PageDims, LayoutItem, TableResult } from './svg-types.js';
import {
  hu2mm, escapeXml, attr, intAttr, find, children, findChild,
  isCJK, fontForChar, spacingForChar, fontFamilyWithFallback, wrapText, estimateTextWidth,
} from './svg-utils.js';
import { collectRuns, letterSpacingAttr, textDecorationAttr, renderPictures } from './svg-text.js';
import { mapPuaStringToUnicode } from '../hwp/hwp-section.js';

let _gradIdCounter = 0;

/** Cumulative visible height of a paragraph in mm. Only looks at the
 *  paragraph's OWN direct linesegarray (not any nested-table cell linesegs),
 *  and takes max with nested tables' declared heights so a treatAsChar=1
 *  table whose lineseg vertsize already counts its height isn't double-added. */
function paraTotalHeightMm(paraEl: Element): number {
  let linesegBottom = 0;
  const lsas: Element[] = [];
  for (const child of Array.from(paraEl.children)) {
    if ((child.localName || child.nodeName.split(':').pop()) === 'linesegarray') lsas.push(child);
  }
  for (const lsa of lsas) {
    for (const seg of children(lsa, 'lineseg')) {
      const bottom = intAttr(seg, 'vertpos', 0) + intAttr(seg, 'vertsize', 0);
      if (bottom > linesegBottom) linesegBottom = bottom;
    }
  }
  let nestedH = 0;
  for (const nestedTbl of findDirectTables(paraEl)) {
    const szEl = find(nestedTbl, 'sz');
    const h = szEl ? hu2mm(intAttr(szEl, 'height', 0)) : 0;
    nestedH += h;
  }
  return Math.max(hu2mm(linesegBottom), nestedH);
}

/** Emit fill SVG for a cell — solid `<rect fill="#..."/>` or, for gradient
 *  borderFills, `<defs><linearGradient>...` + a rect referencing it. */
function cellFillSvg(bf: BorderFillInfo | undefined, x: number, y: number, w: number, h: number, caches?: Caches): string {
  if (!bf) return '';
  if (bf.fillColor) {
    return `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${bf.fillColor}" stroke="none"/>`;
  }
  if (bf.imageBinDataId && caches) {
    const href = caches.binData.get(bf.imageBinDataId);
    if (href) {
      // HWP fill modes control how the image maps to the cell rect. RESIZE=
      // stretch to fill; CENTER=center at native size; TILE_ALL=repeat. We
      // approximate with preserveAspectRatio: RESIZE → "none" (stretch),
      // CENTER → "xMidYMid meet" (fit + center), tiling → same as RESIZE for
      // now (would need SVG <pattern> for true tiling).
      const mode = bf.imageFillMode ?? 'RESIZE';
      const aspect = (mode === 'RESIZE') ? 'none' : 'xMidYMid meet';
      return `<image preserveAspectRatio="${aspect}" xlink:href="${href}" href="${href}" x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}"/>`;
    }
  }
  if (bf.gradientColors && bf.gradientColors.length > 0) {
    if (bf.gradientColors.length === 1) {
      return `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${bf.gradientColors[0]}" stroke="none"/>`;
    }
    const gid = `hwp-cell-grad-${_gradIdCounter++}`;
    const angle = bf.gradientAngle ?? 0;
    // HWP shear angle: 0° = top-to-bottom, 90° = left-to-right (matches
    // Hancom Office behaviour). Vector components dx=sin, dy=cos give the
    // correct direction of gradient progression.
    const rad = (angle * Math.PI) / 180;
    const dx = Math.sin(rad), dy = Math.cos(rad);
    const stops = bf.gradientColors.map((c, i) => {
      const pct = (i * 100 / (bf.gradientColors!.length - 1));
      return `<stop offset="${pct.toFixed(2)}%" stop-color="${c}"/>`;
    }).join('');
    const defs = `<defs><linearGradient id="${gid}" gradientUnits="userSpaceOnUse" x1="${x.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(x + w * dx).toFixed(2)}" y2="${(y + h * dy).toFixed(2)}">${stops}</linearGradient></defs>`;
    return `${defs}<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="url(#${gid})" stroke="none"/>`;
  }
  return '';
}

// ── Helpers for avoiding double-rendering of nested tables ──

/**
 * Find tables that belong directly to a paragraph element.
 * Tables may be inside <hp:ctrl> or <hp:run> children, but we must NOT
 * descend into table cells (tc/subList) which would find inner tables.
 */
export function findDirectTables(paraEl: Element): Element[] {
  const results: Element[] = [];
  function walk(el: Element, insideRect: boolean) {
    for (const child of Array.from(el.children)) {
      const ln = localName(child);
      if (ln === 'tbl') {
        results.push(child);
        // Don't descend into the table — inner tables belong to cell content
      } else if (ln === 'tc') {
        // Don't descend into table cells
      } else if (ln === 'subList') {
        // subLists inside a rect wrap the shape's inner paragraphs;
        // descend so we can pick up tables inside shape text-boxes. subLists
        // inside a tc are cell content and handled by the cell renderer.
        if (insideRect) walk(child, false);
      } else if (ln === 'rect') {
        walk(child, true);
      } else {
        walk(child, insideRect);
      }
    }
  }
  walk(paraEl, false);
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
  const typeUpper = spec.type.toUpperCase();
  if (typeUpper === 'DOUBLE' || typeUpper === 'DOUBLE_SLIM') {
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

      // Cell margins — 0xFFFFFFFF (4294967295) is HWP "use default" sentinel → treat as 141 HU
      const HWP_DEFAULT_MARGIN = 141;
      const HWP_SENTINEL = 4294967295;
      const marginEl = findChild(cellEl, 'cellMargin');
      function readMarginVal(el: Element | null, fallbackEl: Element, attr1: string, attr2: string): number {
        const raw = el ? intAttr(el, attr1, HWP_SENTINEL) : intAttr(fallbackEl, attr2, HWP_SENTINEL);
        return raw === HWP_SENTINEL ? HWP_DEFAULT_MARGIN : raw;
      }
      const cellMTop = hu2mm(readMarginVal(marginEl, cellEl, 'top', 'marginTop'));
      const cellMBottom = hu2mm(readMarginVal(marginEl, cellEl, 'bottom', 'marginBottom'));

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
  // 0xFFFFFFFF (4294967295) is HWP "use default" sentinel → treat as 141 HU
  const HWP_MARGIN_SENTINEL = 4294967295;
  const HWP_DEFAULT_CELL_MARGIN = 141;
  const marginEl = findChild(cellEl, 'cellMargin');
  function readCellMarginVal(side: string): number {
    const raw = marginEl ? intAttr(marginEl, side, HWP_MARGIN_SENTINEL) : intAttr(cellEl, `margin${side.charAt(0).toUpperCase() + side.slice(1)}`, HWP_MARGIN_SENTINEL);
    return raw === HWP_MARGIN_SENTINEL ? HWP_DEFAULT_CELL_MARGIN : raw;
  }
  const mLeft = hu2mm(readCellMarginVal('left'));
  const mRight = hu2mm(readCellMarginVal('right'));
  const mTop = hu2mm(readCellMarginVal('top'));

  const textX = x + mLeft;
  const textWidth = cellW - mLeft - mRight;
  if (textWidth <= 0) return parts;

  // Collect all paragraphs and vertical alignment
  const subList = find(cellEl, 'subList');
  const cellParas = subList ? children(subList, 'p') : children(cellEl, 'p');
  const vertAlignStr = subList ? attr(subList, 'vertAlign', 'TOP').toUpperCase() : 'TOP';
  const vertAlignMode = vertAlignStr === 'CENTER' ? 1 : vertAlignStr === 'BOTTOM' ? 2 : 0;
  // 한 줄로 입력 at cell level: subList.lineWrap="Squeeze" means all paragraphs in this cell
  // must stay on one line, compressing character spacing to fit the cell width.
  const cellLineWrap = subList ? attr(subList, 'lineWrap', '').toUpperCase() : '';
  const cellSqueeze = cellLineWrap === 'SQUEEZE' || cellLineWrap === 'FIXED';

  // Cell bottom margin
  const mBottom = hu2mm(readCellMarginVal('bottom'));

  // First pass: compute total content height for vertical centering
  interface RunSegment {
    text: string;
    fontSize: number;
    fontFamily: string;
    color: string;
    fw: string;
    fi: string;
    ls: string; // letter-spacing attr string
    td: string; // text-decoration attr string
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
    flags?: number;      // raw lineseg flags
    isWordWrapped?: boolean; // line ended by word-wrap (auto), not by explicit break
    distribute?: boolean; // paragraph uses DISTRIBUTE alignment (spread across horzsizeMm)
    squeeze?: boolean;   // 한 줄로 입력: compress spacing to fit single line
  }
  // Content items: either text lines or nested tables, in paragraph order.
  // spacingBeforeMm is set on the FIRST item of each paragraph, spacingAfterMm
  // on the LAST item, so the render loop can add paragraph 앞/뒤 margins.
  type ContentItem = ({ kind: 'line'; line: LineInfo } | { kind: 'table'; tbl: Element; heightMm: number; outMarginTopMm: number; outMarginBottomMm: number; outMarginLeftMm: number; outMarginRightMm: number; vertposMm?: number }) & { spacingBeforeMm?: number; spacingAfterMm?: number };
  const contentItems: ContentItem[] = [];
  const allLines: LineInfo[] = [];
  let totalContentH = mTop + mBottom; // start with vertical margins
  let maxLineBottomMm = 0;
  let anyLinesegData = false;

  for (const paraEl of cellParas) {
    const paraPrIdRef = attr(paraEl, 'paraPrIDRef', attr(paraEl, 'paraPrId', ''));
    const paraShape = paraPrIdRef ? caches.paraShapes.get(paraPrIdRef) : undefined;
    // 한 줄로 입력: cell-level OR paragraph-level squeeze
    const paraLineWrap = paraShape?.lineWrap?.toUpperCase() ?? '';
    const isSqueeze = cellSqueeze || paraLineWrap === 'SQUEEZE' || paraLineWrap === 'FIXED';
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
        // paraShape.leftMargin / indent are HWPUNIT. hu2mm gives mm directly —
        // no scale factor needed (the old *0.5 / *0.53 heuristics were bugs
        // that under-indented hanging paragraphs).
        tx += hu2mm(paraShape.leftMargin);
        if (paraShape.indent < 0) {
          // 내어쓰기: body lines shift right by |indent| so continuation
          // lines align with the text after the first-line marker.
          bodyIndentMm = hu2mm(-paraShape.indent);
        } else if (paraShape.indent > 0) {
          firstLineIndentMm = hu2mm(paraShape.indent);
        }
      }
    }

    const paraLineSpacing = paraShape ? paraShape.lineSpacing / 100 : 1.6;
    const paraSpacingBeforeMm = paraShape ? hu2mm(paraShape.spacingBefore) : 0;
    const paraSpacingAfterMm = paraShape ? hu2mm(paraShape.spacingAfter) : 0;
    const itemsBefore = contentItems.length;

    if (!fullText.trim()) {
      const paraNestedTables = findDirectTables(paraEl);
      const linesegArray = children(paraEl, 'linesegarray')[0];
      const firstSeg = linesegArray ? children(linesegArray, 'lineseg')[0] : null;
      const segVertpos = firstSeg ? hu2mm(intAttr(firstSeg, 'vertpos', 0)) : 0;
      const segVertsize = firstSeg ? hu2mm(intAttr(firstSeg, 'vertsize', 0)) : 0;
      // An empty paragraph that ALSO carries a nested table (e.g. the outer
      // 별첨 wrapper cell where paragraphs stack purely to hold tables) should
      // not advance textY for its own text line — the table sits at
      // cellContentY + nested lineseg vertpos. If it has no nested table,
      // it's genuinely blank space that should consume vertical room.
      if (paraNestedTables.length === 0) {
        if (segVertsize > 0) {
          anyLinesegData = true;
          const bottom = segVertpos + segVertsize;
          if (bottom > maxLineBottomMm) maxLineBottomMm = bottom;
          totalContentH += segVertsize;
        } else {
          totalContentH += fontSize * paraLineSpacing;
        }
      }
      const emptyLine: LineInfo = { segments: [], fontSize, lineSpacing: paraLineSpacing, anchor, tx, vertposMm: segVertpos, vertsizeMm: (segVertsize > 0 && paraNestedTables.length === 0) ? segVertsize : undefined };
      if (paraNestedTables.length === 0) {
        allLines.push(emptyLine);
        contentItems.push({ kind: 'line', line: emptyLine });
      }

      for (const nestedTbl of paraNestedTables) {
        const szEl = find(nestedTbl, 'sz');
        const nestedH = szEl ? hu2mm(intAttr(szEl, 'height', 0)) : 0;
        const omEl = find(nestedTbl, 'outMargin');
        const omTop = hu2mm(omEl ? intAttr(omEl, 'top', 0) : intAttr(nestedTbl, 'outMarginTop', 0));
        const omBottom = hu2mm(omEl ? intAttr(omEl, 'bottom', 0) : intAttr(nestedTbl, 'outMarginBottom', 0));
        const omLeft = hu2mm(omEl ? intAttr(omEl, 'left', 0) : intAttr(nestedTbl, 'outMarginLeft', 0));
        const omRight = hu2mm(omEl ? intAttr(omEl, 'right', 0) : intAttr(nestedTbl, 'outMarginRight', 0));
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

    interface LineRange { start: number; end: number; vertposMm?: number; baselineMm?: number; horzsizeMm?: number; horzposMm?: number; vertsizeMm?: number; spacingMm?: number; flags?: number; }
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
          const flags = intAttr(lineSegs[si], 'flags', 0);
          lineRanges.push({ start, end, vertposMm: hu2mm(vertpos), baselineMm: hu2mm(baseline), horzsizeMm: horzsize > 0 ? hu2mm(horzsize) : undefined, horzposMm: horzpos > 0 ? hu2mm(horzpos) : undefined, vertsizeMm: vertsize > 0 ? hu2mm(vertsize) : undefined, spacingMm: spacing > 0 ? hu2mm(spacing) : undefined, flags });
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
      // In squeeze mode, ignore newlines — all text must fit on one line
      const rawLines = isSqueeze ? [fullText.replace(/\n/g, '')] : fullText.split('\n');
      for (let li = 0; li < rawLines.length; li++) {
        const rawLine = rawLines[li];
        let wrapped: string[];
        if (isSqueeze || (lineSegs.length === 1 && rawLines.length === 1)) {
          // Squeeze or single lineseg: no wrapping — force entire line to fit as-is
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
        offset++; // skip the \n (noop in squeeze mode since we removed newlines)
      }
    }

    // Bullet segment prepended to the first line of a bulleted paragraph.
    // We inherit the char shape of the first real run so the bullet's size
    // and color match the visible text.
    let bulletSeg: RunSegment | null = null;
    if (paraShape?.bulletId) {
      const rawBullet = caches.bullets?.get(String(paraShape.bulletId));
      if (rawBullet) {
        const bulletText = mapPuaStringToUnicode(rawBullet);
        if (bulletText.trim()) {
          const rcs = caches.charShapes.get(firstRun?.charPrId ?? '');
          const sampleCh = bulletText[0];
          bulletSeg = {
            text: bulletText + ' ',
            fontSize: rcs ? hu2mm(rcs.height) : 3.5,
            fontFamily: rcs ? fontForChar(rcs, sampleCh) : 'sans-serif',
            color: rcs?.textColor || '#000000',
            fw: rcs?.bold ? 'bold' : 'normal',
            fi: rcs?.italic ? 'italic' : 'normal',
            ls: '',
            td: '',
          };
        }
      }
    }

    // For each line, build segments grouped by run styling AND script type
    for (const range of lineRanges) {
      const segments: RunSegment[] = [];
      if (bulletSeg && range === lineRanges[0]) {
        segments.push(bulletSeg);
      }
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
            td: textDecorationAttr(rcs),
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
      // Detect whether this line ended by an explicit line break (hp:lineBreak
      // inserted as \n during collectRuns) vs auto word-wrap. Non-last lines
      // ending in \n should NOT be JUSTIFY-stretched — HWP renders them at
      // their natural width so the last text token stays where the author
      // intended.
      const lineText = fullText.substring(range.start, range.end);
      const endedByLineBreak = lineText.endsWith('\n') || fullText.charAt(range.end) === '\n';
      const lineInfo: LineInfo = { segments, fontSize: lineFontSize, lineSpacing: paraLineSpacing, anchor, tx: lineTx, vertposMm: range.vertposMm, baselineMm: range.baselineMm, horzsizeMm: range.horzsizeMm, horzposMm: range.horzposMm, vertsizeMm: rangeVertsizeMm, spacingMm: rangeSpacingMm, isLastLine: isLast || endedByLineBreak, justify: isJustify, distribute: isDistribute, squeeze: isSqueeze };
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
        const omEl = find(nestedTbl, 'outMargin');
        const omTop = hu2mm(omEl ? intAttr(omEl, 'top', 0) : intAttr(nestedTbl, 'outMarginTop', 0));
        const omBottom = hu2mm(omEl ? intAttr(omEl, 'bottom', 0) : intAttr(nestedTbl, 'outMarginBottom', 0));
        const omLeft = hu2mm(omEl ? intAttr(omEl, 'left', 0) : intAttr(nestedTbl, 'outMarginLeft', 0));
        const omRight = hu2mm(omEl ? intAttr(omEl, 'right', 0) : intAttr(nestedTbl, 'outMarginRight', 0));
        contentItems.push({ kind: 'table', tbl: nestedTbl, heightMm: nestedH, outMarginTopMm: omTop, outMarginBottomMm: omBottom, outMarginLeftMm: omLeft, outMarginRightMm: omRight, vertposMm: paraLineTopMm });
        totalContentH += nestedH + omTop + omBottom;
      }
    }
    // Attach the paragraph's spacingBefore to its FIRST new content item and
    // spacingAfter to its LAST so the render loop can push textY appropriately.
    if (contentItems.length > itemsBefore) {
      if (paraSpacingBeforeMm > 0) contentItems[itemsBefore].spacingBeforeMm = paraSpacingBeforeMm;
      if (paraSpacingAfterMm > 0) contentItems[contentItems.length - 1].spacingAfterMm = paraSpacingAfterMm;
    }
  }

  // Collect all nested tables for counting (used in vert align heuristic)
  const nestedTables: Element[] = [];
  for (const paraEl of cellParas) {
    nestedTables.push(...findDirectTables(paraEl));
  }

  // Emit any pictures anchored in this cell's paragraphs. HWP anchors a
  // "treatAsChar" picture inside the paragraph's text flow, so its horizontal
  // placement follows the paragraph's alignment (LEFT/CENTER/RIGHT).
  let picStripH = 0;
  const picYBase = y + mTop;
  for (const paraEl of cellParas) {
    const paraPrIdRefP = attr(paraEl, 'paraPrIDRef', attr(paraEl, 'paraPrId', ''));
    const paraShapeP = paraPrIdRefP ? caches.paraShapes.get(paraPrIdRefP) : undefined;
    // Peek at total pic width so alignment (center/right) can offset correctly.
    const probe = renderPictures(paraEl, caches, 0, 0);
    if (!probe.svg) continue;
    let picX = x + mLeft;
    const align = paraShapeP?.alignment;
    if (align === 3) picX = x + (cellW - probe.reservedW) / 2;          // CENTER
    else if (align === 2) picX = x + cellW - mRight - probe.reservedW;  // RIGHT
    const p = renderPictures(paraEl, caches, picX, picYBase + picStripH);
    if (p.svg) {
      parts.push(p.svg);
      picStripH += p.reservedH;
    }
  }

  if (allLines.length === 0 && nestedTables.length === 0 && picStripH === 0) return parts;

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
  // Honor subList.vertAlign strictly per spec: TOP=0 leaves content at the
  // top of the cell, CENTER=1 vertically centers, BOTTOM=2 pushes to bottom.
  let vertOffset = 0;
  if (vertAlignMode === 1) {
    vertOffset = Math.max(0, (centerCellH - mTop - mBottom - textOnlyH) / 2);
  } else if (vertAlignMode === 2) {
    vertOffset = Math.max(0, centerCellH - mTop - mBottom - textOnlyH);
  }

  let textY = y + mTop + vertOffset;

  // Second pass: render lines
  const cellContentY = y + mTop + vertOffset;
  let useLinesegPos = true;

  parts.push(`<g clip-path="url(#${clipId})">`);
  let isFirstRenderedLine = true;
  for (const item of contentItems) {
    // Advance textY for paragraph spacingBefore (문단 간격 위) attached to
    // the first item of each paragraph.
    if (item.spacingBeforeMm) textY += item.spacingBeforeMm;
    if (item.kind === 'table') {
      // Nested inline table's own outMargin.top (바깥 여백) pushes the box
      // down away from prior content.
      let tableY = textY + item.outMarginTopMm;
      if (item.vertposMm !== undefined) {
        // vertposMm is paragraph-local — the paragraph's own first lineseg
        // start. It resets to 0 for each paragraph in the cell, so never let
        // it pull tableY BACKWARDS above the running textY.
        const candidateY = cellContentY + item.vertposMm + item.outMarginTopMm;
        if (candidateY > tableY) tableY = candidateY;
      }
      // Inline (treatAsChar=1) nested tables sit at the RUN START of a HWP
      // hanging paragraph, which is leftMargin + |indent| (body-line
      // position). All inline tables in the same outer cell share the same
      // effective anchor, so use the FIRST cell paragraph's indent as the
      // representative value — this matches HWP's rendering that treats
      // stacked inline wrappers as sharing a common text baseline anchor.
      const firstPara = cellParas[0];
      let runOffsetMm = 0;
      if (firstPara) {
        const firstPrIdRef = attr(firstPara, 'paraPrIDRef', attr(firstPara, 'paraPrId', ''));
        const firstShape = firstPrIdRef ? caches.paraShapes.get(firstPrIdRef) : undefined;
        if (firstShape) {
          runOffsetMm += hu2mm(firstShape.leftMargin);
          if (firstShape.indent < 0) runOffsetMm += hu2mm(-firstShape.indent);
          else if (firstShape.indent > 0) runOffsetMm += hu2mm(firstShape.indent);
        }
      }
      const tableX = textX + runOffsetMm + item.outMarginLeftMm;
      const nestedResult = renderTableEl(item.tbl, caches, tableX, tableY, textWidth - item.outMarginLeftMm - item.outMarginRightMm);
      if (nestedResult) {
        textY = tableY;
        parts.push(nestedResult.svg);
        // Advance by the nested table's rendered height plus its own
        // outMargin.bottom, so following text sits clear of the box.
        textY += nestedResult.height + item.outMarginBottomMm;
        isFirstRenderedLine = true; // reset: next text line starts a fresh block
      }
      if (item.spacingAfterMm) textY += item.spacingAfterMm;
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
    if (isFirstRenderedLine && line.baselineMm !== undefined && line.baselineMm > 0) {
      // First line after a nested inline table (or start of content): position
      // baseline one ascent below the current textY so text doesn't overlap
      // the preceding block. The lineseg's baselineMm gives us the ascent
      // portion of the line box.
      textY += line.baselineMm;
    } else if (useLinesegPos && line.vertposMm !== undefined && line.baselineMm !== undefined) {
      const linesegY = cellContentY + line.vertposMm + line.baselineMm;
      if (linesegY >= textY) {
        textY = linesegY;
      } else if (line.vertsizeMm !== undefined && line.vertsizeMm > 0) {
        // linesegY is stale (from a base HWP encoded before layout push).
        // Advance by the line's own vertsize so the wrapped line sits below
        // the previous one instead of collapsing onto the same baseline.
        textY += line.vertsizeMm + (line.spacingMm ?? 0);
      } else {
        textY += line.fontSize * line.lineSpacing;
      }
    } else if (line.vertsizeMm !== undefined && line.vertsizeMm > 0) {
      textY += line.vertsizeMm + (line.spacingMm ?? 0);
    } else {
      textY += line.fontSize;
    }
    isFirstRenderedLine = false;
    let textLengthAttr = '';
    const lineText = line.segments.map(s => s.text).join('');
    if (line.horzsizeMm !== undefined && line.horzsizeMm > 0) {
      // HWP never scales glyph widths at line level — only adjusts inter-character spacing (자간).
      // horzsize is the pre-calculated target width from HWP's layout engine —
      // it's the FULL cell-content width; for hanging-indent wrap lines the
      // effective fill width is reduced by the body indent already applied
      // to line.tx, so cap by (textX + textWidth − line.tx).
      const spaceRight = Math.max(0, (textX + textWidth) - line.tx);
      if (((line.justify && !line.isLastLine) || line.distribute) && lineText.trim().length > 1) {
        const targetWidth = Math.min(line.horzsizeMm, spaceRight);
        textLengthAttr = ` textLength="${targetWidth.toFixed(2)}" lengthAdjust="spacing"`;
      } else if (line.anchor === 'end' && line.horzposMm !== undefined && lineText.trim().length > 1) {
        // RIGHT-aligned: force text to exactly fill horzsize so left edge lands at textX+horzpos.
        const targetWidth = Math.min(line.horzsizeMm, spaceRight);
        if (targetWidth > 0) {
          textLengthAttr = ` textLength="${targetWidth.toFixed(2)}" lengthAdjust="spacing"`;
        }
      }
    }
    if (!textLengthAttr && line.squeeze && lineText.trim().length > 1) {
      // 한 줄로 입력 (Squeeze): compress character spacing only when text would overflow.
      // Apply regardless of whether lineseg horzsize data is present.
      const est = estimateTextWidth(lineText, line.fontSize);
      if (est > textWidth) {
        textLengthAttr = ` textLength="${textWidth.toFixed(2)}" lengthAdjust="spacing"`;
      }
    }
    // horzpos = left edge of text segment (column-local, spec: "컬럼에서의 시작 위치").
    // horzsize = segment width. Together they define the text area regardless of alignment.
    //   start anchor: text left  = textX + horzpos
    //   end anchor:   text right = textX + horzpos + horzsize
    //   middle anchor: keep tx (x+cellW/2) — horzpos is not useful for centering
    const lineTx =
      (line.horzposMm !== undefined && line.anchor === 'start')
        ? textX + line.horzposMm
      : (line.horzposMm !== undefined && line.anchor === 'end' && line.horzsizeMm !== undefined)
        ? textX + line.horzposMm + line.horzsizeMm
      : line.tx;
    // Resolve font family with the kind-aware fallback chain from the cache.
    const family = (name: string) => fontFamilyWithFallback(name, caches.fontKinds?.get(name));
    if (line.segments.length === 1) {
      const seg = line.segments[0];
      // Suppress letter-spacing when textLength is applied with end-anchor:
      // SVG's negative letter-spacing adds trailing space AFTER the last glyph's advance,
      // which shifts the glyph visually past the anchor point. textLength already handles spacing.
      const lsAttr = (textLengthAttr && line.anchor === 'end') ? '' : seg.ls;
      // For right-aligned text, trailing spaces shift visible text left — trim them.
      const segText = line.anchor === 'end' ? seg.text.trimEnd() : seg.text;
      parts.push(`<text xml:space="preserve" x="${lineTx.toFixed(2)}" y="${textY.toFixed(2)}" font-size="${seg.fontSize.toFixed(2)}" font-family="${escapeXml(family(seg.fontFamily))}" fill="${seg.color}" font-weight="${seg.fw}" font-style="${seg.fi}"${lsAttr}${seg.td}${textLengthAttr} text-anchor="${line.anchor}">${escapeXml(segText)}</text>`);
    } else {
      // For right-aligned text, trailing spaces in the last segment shift visible text left — trim them.
      const segs = line.anchor === 'end'
        ? line.segments.map((seg, i) => i === line.segments.length - 1 ? { ...seg, text: seg.text.trimEnd() } : seg)
        : line.segments;
      const tspans = segs.map(seg =>
        `<tspan font-size="${seg.fontSize.toFixed(2)}" font-family="${escapeXml(family(seg.fontFamily))}" fill="${seg.color}" font-weight="${seg.fw}" font-style="${seg.fi}"${seg.ls}${seg.td}>${escapeXml(seg.text)}</tspan>`
      ).join('');
      parts.push(`<text xml:space="preserve" x="${lineTx.toFixed(2)}" y="${textY.toFixed(2)}"${textLengthAttr} text-anchor="${line.anchor}">${tspans}</text>`);
    }
    if (!line.vertsizeMm) {
      textY += line.fontSize * (line.lineSpacing - 1);
    }
    // Paragraph spacingAfter (문단 간격 아래) attached to the LAST item of a
    // paragraph — applied after that item renders.
    if (item.spacingAfterMm) textY += item.spacingAfterMm;
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
  xOffsetMm: number = 0,
): { yPos: number; currentPage: number } {
  const rows = children(tblEl, 'tr');
  if (rows.length === 0) return { yPos, currentPage };

  // Table declared height (hp:sz height). For single-row tables where the
  // cellSz height exceeds the table's own declared size, Hancom splits the
  // cell content across pages — we don't split, but we can use hp:sz as the
  // page-fit cap so the whole table doesn't get pushed to the next page.
  const tblSzEl = find(tblEl, 'sz');
  const tblDeclaredH = tblSzEl ? hu2mm(intAttr(tblSzEl, 'height', 0)) : 0;

  // HWPX stores outMargin (바깥 여백) as a child element <hp:outMargin ...>,
  // not as attributes on <hp:tbl>. Same for the anchored position offset
  // <hp:pos vertOffset/horzOffset>. Apply to the top and left edges so the
  // table starts at the (x, y) HWP intended.
  const outMarginEl = find(tblEl, 'outMargin');
  const outMarginTop = outMarginEl
    ? hu2mm(intAttr(outMarginEl, 'top', 0))
    : hu2mm(intAttr(tblEl, 'outMarginTop', 0));
  const outMarginBottom = outMarginEl
    ? hu2mm(intAttr(outMarginEl, 'bottom', 0))
    : hu2mm(intAttr(tblEl, 'outMarginBottom', 0));
  const outMarginLeft = outMarginEl
    ? hu2mm(intAttr(outMarginEl, 'left', 0))
    : hu2mm(intAttr(tblEl, 'outMarginLeft', 0));
  const posEl = find(tblEl, 'pos');
  const posVertOffset = posEl ? hu2mm(intAttr(posEl, 'vertOffset', 0)) : 0;
  const posHorzOffset = posEl ? hu2mm(intAttr(posEl, 'horzOffset', 0)) : 0;
  yPos += posVertOffset + outMarginTop;
  xOffsetMm += posHorzOffset + outMarginLeft;

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
      const s = cellFillSvg(pc.bf, pc.x, pc.y, pc.w, pc.h, caches);
      if (s) parts.push(s);
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

    // For single-row tables where content overflows, use the smaller of the
    // adjusted row H and the table's own declared height for page-fit checks
    // (Hancom splits inside the cell across pages; we render the whole cell
    // on whichever page it starts on).
    const fitH = (rows.length === 1 && tblDeclaredH > 0 && tblDeclaredH < actualRowH)
      ? tblDeclaredH
      : actualRowH;
    if (yPos + fitH > dims.pageBottom && pages[currentPage].length > 0) {
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

    // Single-row single-cell tables (typical 별첨 wrapper) can be taller than
    // one page can hold. If actualRowH > available space on the current page,
    // cap the CURRENT-page cell rect at what fits, then render a continuation
    // rect on the next page for the remainder. Content overflow is clipped by
    // renderCellContent; the visual approximation is that the cell renders as
    // a rectangle split across pages, mirroring Hancom's behaviour.
    const availOnPage = dims.pageBottom - yPos;
    const isSingleCellRow = rows.length === 1 && cells.length === 1;
    // "글자처럼 취급" (treatAsChar=1) tables are laid out as a single character in
    // the text flow and per HWP spec cannot be split across pages. Only allow
    // page-splitting for floating (treatAsChar=0) tables.
    const treatAsCharAttr = posEl ? attr(posEl, 'treatAsChar', '0') : '0';
    const isTreatAsChar = treatAsCharAttr === '1';
    const overflow = isSingleCellRow && !isTreatAsChar && actualRowH > availOnPage;
    const thisPageH = overflow ? availOnPage : actualRowH;

    if (overflow) {
      // Single-row single-cell overflow: render cell rect + content clipped to
      // this page, then continuation rect + shifted content on the next page.
      const cellEl = cells[0];
      const addrEl2 = findChild(cellEl, 'cellAddr');
      const colAddr2 = addrEl2 ? intAttr(addrEl2, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const szEl2 = findChild(cellEl, 'cellSz');
      const cellW2 = szEl2 ? hu2mm(intAttr(szEl2, 'width', 0)) : hu2mm(intAttr(cellEl, 'width', 0));
      const bfIdRef2 = attr(cellEl, 'borderFillIDRef', '');
      const borderFill2 = bfIdRef2 ? caches.borderFills.get(bfIdRef2) : undefined;
      const x2 = dims.contentLeft + xOffsetMm + (colOffsets.get(colAddr2) ?? 0);
      const rowY = yPos;
      // Snap the split point to a content boundary — no inner treatAsChar=1
      // (글자처럼 취급) nested table should straddle pageBottom (spec: those
      // are laid out as a single character). When a straddle is detected, cut
      // to the BOTTOM of the previous paragraph (its last lineseg) so the
      // preceding text's descenders remain within the clip.
      let consumedH = thisPageH;
      {
        const subList = find(cellEl, 'subList');
        const cellParas = subList ? children(subList, 'p') : children(cellEl, 'p');
        // HWP's per-paragraph lineseg vertpos is USUALLY cell-relative, but
        // some paragraphs (typically the last one containing a treatAsChar=1
        // table) reset to 0. Accumulate the max bottom seen so far as the
        // effective start for the next paragraph when its own vertpos is 0.
        let accBottomMm = 0;
        let prevParaBottomMm = 0;
        let prevFontSizeMm = 3.5;
        for (const paraEl of cellParas) {
          const lsas: Element[] = [];
          for (const ch of Array.from(paraEl.children)) {
            if ((ch.localName || ch.nodeName.split(':').pop()) === 'linesegarray') lsas.push(ch);
          }
          let paraStartHu = 0;
          let paraBottomHu = 0;
          for (const lsa of lsas) {
            const segs = children(lsa, 'lineseg');
            if (segs.length && paraStartHu === 0) paraStartHu = intAttr(segs[0], 'vertpos', 0);
            for (const seg of segs) {
              const b = intAttr(seg, 'vertpos', 0) + intAttr(seg, 'vertsize', 0);
              if (b > paraBottomHu) paraBottomHu = b;
            }
          }
          const paraStartMm = paraStartHu > 0 ? hu2mm(paraStartHu) : accBottomMm;
          const paraBottomMm = paraBottomHu > 0 ? hu2mm(paraBottomHu) : accBottomMm;
          const effectiveBottomMm = Math.max(paraBottomMm, paraStartMm);
          for (const nestedTbl of findDirectTables(paraEl)) {
            const posE = find(nestedTbl, 'pos');
            const tac = posE ? attr(posE, 'treatAsChar', '0') : '0';
            if (tac !== '1') continue;
            const szEl = find(nestedTbl, 'sz');
            const nestedH = szEl ? hu2mm(intAttr(szEl, 'height', 0)) : 0;
            const startInCell = paraStartMm;
            const endInCell = paraStartMm + nestedH;
            if (startInCell < consumedH && endInCell > consumedH) {
              // Snap to the previous paragraph's bottom + its font descent.
              // HWP's lineseg vertsize covers the em-height plus HWP's own
              // small descent allowance; our fallback font's descent is often
              // taller, so include a descent-worth pad to keep the last text
              // line's descenders on THIS page rather than bleeding to the
              // next.
              const descentPadMm = prevFontSizeMm * 0.25;
              consumedH = Math.max(0, prevParaBottomMm + descentPadMm);
            }
          }
          if (effectiveBottomMm > accBottomMm) accBottomMm = effectiveBottomMm;
          prevParaBottomMm = effectiveBottomMm;
          // Capture the paragraph's first-run font size as the "current" text
          // scale for descent-pad calculation on the next snap iteration.
          const paraFirstRun = collectRuns(paraEl)[0];
          const paraCs = paraFirstRun ? caches.charShapes.get(paraFirstRun.charPrId) : undefined;
          if (paraCs) prevFontSizeMm = hu2mm(paraCs.height);
        }
      }
      // Note: leftover content height = actualRowH − consumedH; used implicitly via contBoxH.

      // Render cell content ONCE — same string reused on both pages, only clip changes
      const contentSvg = renderCellContent(cellEl, caches, x2, rowY, cellW2, actualRowH, actualRowH).join('\n');

      // Page 2: draw the box rect at the FULL available height (Hancom's
      // wrapper extends to the page-bottom margin), while clipping content
      // to consumedH (the safe snap point above the next treatAsChar=1
      // nested table). remainH advances yPos by the leftover content, not by
      // the visible-box height.
      const boxH = thisPageH;   // visible rect fills the rest of the page
      flushSegment(pageParts, segmentStartY);       // flush any prior rows
      segmentCells.push({ x: x2, y: rowY, w: cellW2, h: boxH, bf: borderFill2, el: cellEl, declaredH: boxH });
      const fillSvg = cellFillSvg(borderFill2, x2, rowY, cellW2, boxH, caches);
      if (fillSvg) pageParts.push(fillSvg);
      const { hLines: h2, vLines: v2 } = buildGridLines([{ x: x2, y: rowY, w: cellW2, h: boxH, bf: borderFill2 }], tableBf);
      pageParts.push(...renderGridLines(h2, v2));
      segmentCells = [];

      // Clip content up to consumedH (bottom of the last paragraph that fits).
      const clipIdP2 = `page-split-clip-${_gradIdCounter++}`;
      pageParts.push(`<defs><clipPath id="${clipIdP2}"><rect x="${x2.toFixed(2)}" y="${rowY.toFixed(2)}" width="${cellW2.toFixed(2)}" height="${consumedH.toFixed(2)}"/></clipPath></defs>`);
      pageParts.push(`<g clip-path="url(#${clipIdP2})">${contentSvg}</g>`);

      pageParts.push('</g>');
      if (pageParts.length > 2) {
        pages[currentPage].push({ svg: pageParts.join('\n'), y: segmentStartY, height: 0 });
      }

      // Page 3: continuation box uses the leftover cell height (actualRowH − boxH).
      // Apply the outer table's outMargin.top on the new page too — each
      // fragment respects the outer margins independently.
      pages.push([]);
      currentPage = pages.length - 1;
      yPos = dims.contentTop + outMarginTop;
      segmentStartY = yPos;
      pageParts = [`<g class="table">`];

      const contBoxH = actualRowH - boxH;
      const fillSvg3 = cellFillSvg(borderFill2, x2, yPos, cellW2, contBoxH, caches);
      if (fillSvg3) pageParts.push(fillSvg3);
      const { hLines: h3, vLines: v3 } = buildGridLines([{ x: x2, y: yPos, w: cellW2, h: contBoxH, bf: borderFill2 }], tableBf);
      pageParts.push(...renderGridLines(h3, v3));

      // Same content shifted up so overflow appears at top of page-3 continuation.
      // The shift lines up the overflow content with the NEW cellContentY
      // (yPos + cellMargin.top) so the continuation fragment preserves the
      // cell's own top padding independently on the new page.
      const cellMarginEl = findChild(cellEl, 'cellMargin');
      const HWP_SENTINEL = 4294967295;
      const rawTop = cellMarginEl ? intAttr(cellMarginEl, 'top', HWP_SENTINEL) : intAttr(cellEl, 'marginTop', HWP_SENTINEL);
      const contMTopMm = hu2mm(rawTop === HWP_SENTINEL ? 141 : rawTop);
      const shift = (rowY + consumedH) - (yPos + contMTopMm);
      // Clip content to the cell-content area (below the box's top padding)
      // so duplicated content from page 2 (originally above the split point)
      // doesn't leak into page 3.
      const contentClipY = yPos + contMTopMm;
      const contentClipH = contBoxH - contMTopMm;
      const clipIdP3 = `page-split-clip-${_gradIdCounter++}`;
      pageParts.push(`<defs><clipPath id="${clipIdP3}"><rect x="${x2.toFixed(2)}" y="${contentClipY.toFixed(2)}" width="${cellW2.toFixed(2)}" height="${contentClipH.toFixed(2)}"/></clipPath></defs>`);
      pageParts.push(`<g clip-path="url(#${clipIdP3})"><g transform="translate(0, ${(-shift).toFixed(2)})">${contentSvg}</g></g>`);

      // Include the outer table's outMargin.bottom after the continuation
      // fragment so subsequent content on page 3 clears the wrapper.
      yPos += contBoxH + outMarginBottom;
    } else {
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
          for (let c = colAddr; c < colAddr + colSpan; c++) cellW += colWidths.get(c) ?? 0;
        } else {
          const szElW = findChild(cellEl, 'cellSz');
          cellW = szElW ? hu2mm(intAttr(szElW, 'width', 0)) : hu2mm(intAttr(cellEl, 'width', 0));
        }
        let cellH = actualRowH;
        if (rowSpan > 1) {
          cellH = 0;
          for (let r = rowAddrCell; r < rowAddrCell + rowSpan; r++) cellH += rowHeights.get(r) ?? 0;
        }
        const szElCell = findChild(cellEl, 'cellSz');
        const declaredCellH = szElCell ? hu2mm(intAttr(szElCell, 'height', 0)) : hu2mm(intAttr(cellEl, 'height', 0));
        const x = dims.contentLeft + xOffsetMm + (colOffsets.get(colAddr) ?? 0);
        const bfIdRef = attr(cellEl, 'borderFillIDRef', '');
        const borderFill = bfIdRef ? caches.borderFills.get(bfIdRef) : undefined;
        segmentCells.push({ x, y: yPos, w: cellW, h: cellH, bf: borderFill, el: cellEl, declaredH: declaredCellH });
      }
      yPos += actualRowH;
    }
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

  // Apply the anchored <hp:pos vertOffset/horzOffset> and outer margins so
  // the visible table lines up with HWP's own layout.
  const posEl = find(tblEl, 'pos');
  if (posEl) {
    startY += hu2mm(intAttr(posEl, 'vertOffset', 0));
    startX += hu2mm(intAttr(posEl, 'horzOffset', 0));
  }
  const outMarginEl = find(tblEl, 'outMargin');
  if (outMarginEl) startY += hu2mm(intAttr(outMarginEl, 'top', 0));

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
    const s = cellFillSvg(pc.bf, pc.x, pc.y, pc.w, pc.h, caches);
    if (s) parts.push(s);
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
