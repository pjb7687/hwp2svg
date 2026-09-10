/**
 * HTML renderer: converts HWPX XML DOM into a semantic HTML document.
 *
 * Unlike the SVG renderer, this emits flowing HTML (<p>, <table>, <span>,
 * <img>, <ul>) with inline styles derived from the same caches
 * (charShape / paraShape / borderFill) so the visual result matches HWP
 * as closely as possible while remaining editable / searchable HTML.
 */

import type { HwpxDom } from '../hwpx/loader.js';
import type { Caches, CharShapeInfo, ParaShapeInfo, BorderFillInfo } from './svg-types.js';
import {
  hu2mm, attr, intAttr, find, findChild, children, localName,
  fontFamilyWithFallback, isCJK,
} from './svg-utils.js';
import { buildCaches, getPageDims } from './svg-cache.js';
import { findDirectTables, findDirectCells } from './svg-table.js';
import { mapPuaStringToUnicode } from '../hwp/hwp-section.js';

export interface HtmlOptions {
  /** Emit a full <!doctype html>... wrapper (default true). */
  full?: boolean;
  /** Title used in <title> when `full` is true. */
  title?: string;
  /** Body class name (default "hwp-doc"). */
  className?: string;
}

// ── Public entry point ──

export function renderToHtml(dom: HwpxDom, options: HtmlOptions = {}): string {
  const caches = buildCaches(dom.header, dom.binData);
  const full = options.full !== false;
  const title = options.title ?? 'HWP Document';
  const className = options.className ?? 'hwp-doc';

  const bodyParts: string[] = [];
  // Page geometry captured from the first section; used both for the CSS
  // "paper" frame and for detecting page breaks inside a section.
  let paperCss: PageCss = {
    pageWmm: 210, pageHmm: 297,
    padTopMm: 15, padRightMm: 20, padBottomMm: 15, padLeftMm: 20,
  };

  for (let i = 0; i < dom.sections.length; i++) {
    const secDoc = dom.sections[i];
    const secRoot = secDoc.documentElement;
    const dims = getPageDims(secRoot);
    if (i === 0) {
      paperCss = {
        pageWmm: dims.pageW,
        pageHmm: dims.pageH,
        padTopMm: dims.contentTop,
        padRightMm: dims.pageW - (dims.contentLeft + dims.contentWidth),
        padBottomMm: dims.pageH - dims.pageBottom,
        padLeftMm: dims.contentLeft,
      };
    } else {
      // Each additional section starts a fresh page.
      bodyParts.push(pageBreakMarker());
    }
    bodyParts.push(...renderSection(secRoot, caches));
  }

  const bodyHtml = bodyParts.filter(Boolean).join('\n');

  if (!full) return bodyHtml;

  return [
    `<!doctype html>`,
    `<html lang="ko">`,
    `<head>`,
    `<meta charset="UTF-8">`,
    `<meta name="generator" content="hwp-to-svg (hwp2html)">`,
    `<title>${escapeHtml(title)}</title>`,
    `<style>${defaultCss(paperCss, className)}</style>`,
    `</head>`,
    `<body>`,
    `<article class="${className}">`,
    bodyHtml,
    `</article>`,
    `</body>`,
    `</html>`,
  ].join('\n');
}

interface PageCss {
  pageWmm: number;
  pageHmm: number;
  padTopMm: number;
  padRightMm: number;
  padBottomMm: number;
  padLeftMm: number;
}

/** Emit a hard page-break marker. On screen it renders as a subtle divider so
 *  users can see where a new page begins; in print it forces a real break. */
function pageBreakMarker(): string {
  return `<div class="hwp-pagebreak" style="break-before:page;page-break-before:always"></div>`;
}

// ── Container rendering (section body or table cell) ──

/** Walk direct <hp:p> and <hp:tbl> children of a container and emit HTML.
 *  Used for table-cell content — no page break detection here. */
function renderContainerChildren(container: Element, caches: Caches): string[] {
  const parts: string[] = [];
  for (const child of Array.from(container.children)) {
    const ln = localName(child);
    if (ln === 'p') {
      parts.push(...renderParagraph(child, caches));
    } else if (ln === 'tbl') {
      parts.push(renderTable(child, caches));
    }
    // secPr, ctrl, and other section-level metadata are skipped.
  }
  return parts;
}

/**
 * Walk a section's top-level content, inserting page-break markers to match
 * HWP's own pagination. Two signals drive page breaks:
 *
 *   1) The paragraph carries `pageBreak="1"` (author-forced break).
 *   2) The paragraph's first lineseg `vertpos` is smaller than the previous
 *      paragraph's — HWP resets vertpos when a new page begins, so a drop
 *      means the layout engine started a fresh page here.
 *
 * This mirrors the SVG renderer's page-splitting logic (see svg-renderer.ts).
 */
function renderSection(secRoot: Element, caches: Caches): string[] {
  const parts: string[] = [];
  let lastParaStartVP = -1;

  for (const child of Array.from(secRoot.children)) {
    const ln = localName(child);
    if (ln === 'p') {
      let forceBreak = false;
      const pbAttr = attr(child, 'pageBreak', '0');
      if (pbAttr === '1' || pbAttr.toUpperCase() === 'TRUE') forceBreak = true;

      // Peek at the paragraph's first lineseg vertpos.
      const linesegArray = children(child, 'linesegarray')[0];
      const firstSeg = linesegArray ? children(linesegArray, 'lineseg')[0] : null;
      let vertpos = -1;
      if (firstSeg) vertpos = intAttr(firstSeg, 'vertpos', -1);

      if (!forceBreak && vertpos >= 0 && lastParaStartVP >= 0 && vertpos < lastParaStartVP) {
        forceBreak = true;
      }
      if (forceBreak && parts.length > 0) parts.push(pageBreakMarker());
      if (vertpos >= 0) lastParaStartVP = vertpos;

      parts.push(...renderParagraph(child, caches));
    } else if (ln === 'tbl') {
      parts.push(renderTable(child, caches));
    }
  }
  return parts;
}

// ── Paragraph rendering ──

/**
 * Emit HTML for one paragraph. A single HWP paragraph may contain inline
 * tables and pictures. Inline tables are lifted out (HTML tables can't be
 * inside <p>), so we emit a `<p>` for text before the table, then the
 * `<table>`, then a fresh `<p>` for text after.
 */
function renderParagraph(paraEl: Element, caches: Caches): string[] {
  const parts: string[] = [];
  const paraPrIdRef = attr(paraEl, 'paraPrIDRef', attr(paraEl, 'paraPrId', ''));
  const paraShape = paraPrIdRef ? caches.paraShapes.get(paraPrIdRef) : undefined;

  const pStyle = paragraphStyle(paraShape);
  const bulletChar = getBulletChar(paraShape, caches);

  // Split the paragraph's children into runs of inline content, separated by
  // block-level items (nested tables) that can't sit inside <p>.
  type Chunk =
    | { kind: 'inline'; html: string }
    | { kind: 'table'; el: Element };

  const chunks: Chunk[] = [];
  let inlineBuf: string[] = [];
  const flushInline = () => {
    if (inlineBuf.length > 0) {
      chunks.push({ kind: 'inline', html: inlineBuf.join('') });
      inlineBuf = [];
    }
  };

  for (const child of Array.from(paraEl.children)) {
    const ln = localName(child);
    if (ln === 'run') {
      // Runs may contain text, line breaks, images, and inline tables. Emit
      // inline pieces to the buffer; break out any tables as their own chunk.
      for (const gchild of Array.from(child.children)) {
        const gln = localName(gchild);
        if (gln === 't') {
          const text = mapPuaStringToUnicode(gchild.textContent || '');
          if (text) {
            const charPrIdRef = attr(child, 'charPrIDRef', attr(child, 'charPrId', ''));
            const cs = charPrIdRef ? caches.charShapes.get(charPrIdRef) : undefined;
            inlineBuf.push(renderTextRun(text, cs, caches));
          }
        } else if (gln === 'lineBreak') {
          inlineBuf.push('<br/>');
        } else if (gln === 'tab') {
          inlineBuf.push('\t');
        } else if (gln === 'pic') {
          inlineBuf.push(renderPicture(gchild, caches));
        } else if (gln === 'tbl') {
          flushInline();
          chunks.push({ kind: 'table', el: gchild });
        } else if (gln === 'rect') {
          // Rectangle shapes may contain nested paragraphs (subList) — treat
          // as a table-like block for HTML output.
          const nestedTables = findDirectTables(gchild);
          for (const t of nestedTables) {
            flushInline();
            chunks.push({ kind: 'table', el: t });
          }
        }
        // ctrl/pageNum/secPr etc are skipped.
      }
    } else if (ln === 'lineBreak') {
      inlineBuf.push('<br/>');
    }
  }
  flushInline();

  // If the paragraph is entirely empty (no inline text and no tables), emit a
  // single empty <p> so vertical spacing still shows up.
  if (chunks.length === 0) {
    parts.push(`<p${pStyle}>&#8203;</p>`);
    return parts;
  }

  const bulletPrefix = bulletChar
    ? `<span class="hwp-bullet">${escapeHtml(bulletChar)}&nbsp;</span>`
    : '';
  let bulletUsed = false;

  for (const chunk of chunks) {
    if (chunk.kind === 'inline') {
      const html = chunk.html.trim() ? chunk.html : '&#8203;';
      const prefix = !bulletUsed ? bulletPrefix : '';
      bulletUsed = true;
      parts.push(`<p${pStyle}>${prefix}${html}</p>`);
    } else {
      parts.push(renderTable(chunk.el, caches));
    }
  }

  return parts;
}

/** Emit CSS style for a paragraph based on paraShape. */
function paragraphStyle(paraShape: ParaShapeInfo | undefined): string {
  const decl: string[] = [];
  decl.push('margin:0');
  if (!paraShape) {
    return decl.length ? ` style="${decl.join(';')}"` : '';
  }
  const align = paraShape.alignment;
  if (align === 1) decl.push('text-align:left');
  else if (align === 2) decl.push('text-align:right');
  else if (align === 3) decl.push('text-align:center');
  else if (align === 4) decl.push('text-align:justify;text-align-last:justify');
  else if (align === 0) decl.push('text-align:justify');
  if (paraShape.leftMargin) decl.push(`margin-left:${hu2mm(paraShape.leftMargin).toFixed(2)}mm`);
  if (paraShape.rightMargin) decl.push(`margin-right:${hu2mm(paraShape.rightMargin).toFixed(2)}mm`);
  if (paraShape.indent > 0) decl.push(`text-indent:${hu2mm(paraShape.indent).toFixed(2)}mm`);
  else if (paraShape.indent < 0) {
    // Hanging indent: shift body lines right, pull first-line back.
    const hang = hu2mm(-paraShape.indent);
    decl.push(`padding-left:${hang.toFixed(2)}mm`);
    decl.push(`text-indent:-${hang.toFixed(2)}mm`);
  }
  if (paraShape.spacingBefore) decl.push(`margin-top:${hu2mm(paraShape.spacingBefore).toFixed(2)}mm`);
  if (paraShape.spacingAfter) decl.push(`margin-bottom:${hu2mm(paraShape.spacingAfter).toFixed(2)}mm`);
  // lineSpacing is a percent-like value; 160 = 160% = 1.6.
  if (paraShape.lineSpacing && paraShape.lineSpacing !== 160) {
    decl.push(`line-height:${(paraShape.lineSpacing / 100).toFixed(2)}`);
  }
  decl.push('white-space:pre-wrap');
  return ` style="${decl.join(';')}"`;
}

/** Return the bullet character (already PUA-mapped) or empty when none. */
function getBulletChar(paraShape: ParaShapeInfo | undefined, caches: Caches): string {
  if (!paraShape?.bulletId) return '';
  const raw = caches.bullets?.get(String(paraShape.bulletId));
  if (!raw) return '';
  return mapPuaStringToUnicode(raw).trim();
}

// ── Text run styling ──

function renderTextRun(text: string, cs: CharShapeInfo | undefined, caches: Caches): string {
  // Text runs may span multiple scripts (Hangul/Latin) which use different
  // fonts. Rather than split every run per-character, use the primary
  // (Hangul) font for the span — browsers with the Korean font installed
  // will substitute Latin glyphs from that font automatically, matching HWP.
  const decl: string[] = [];
  if (cs) {
    const fontSize = hu2mm(cs.height);
    decl.push(`font-size:${fontSize.toFixed(2)}mm`);
    if (cs.fontName) {
      const kind = caches.fontKinds.get(cs.fontName);
      decl.push(`font-family:${cssFontFamily(cs.fontName, kind)}`);
    }
    if (cs.textColor && cs.textColor !== '#000000') decl.push(`color:${cs.textColor}`);
    if (cs.bold) decl.push('font-weight:bold');
    if (cs.italic) decl.push('font-style:italic');
    const decorations: string[] = [];
    if (cs.underline) decorations.push('underline');
    if (cs.strikeout) decorations.push('line-through');
    if (decorations.length) {
      const shape = (cs.underlineShape || 'SOLID').toUpperCase();
      const styleMap: Record<string, string> = {
        DOUBLE: 'double', DOT: 'dotted', DASH: 'dashed', DASHED: 'dashed', WAVE: 'wavy',
      };
      const style = styleMap[shape] ?? 'solid';
      decl.push(`text-decoration:${decorations.join(' ')} ${style}`);
    }
    if (cs.spacing) {
      const spacingMm = (cs.spacing / 100) * hu2mm(cs.height);
      decl.push(`letter-spacing:${spacingMm.toFixed(2)}mm`);
    }
  }
  const styleAttr = decl.length ? ` style="${decl.join(';')}"` : '';
  return `<span${styleAttr}>${escapeHtml(text)}</span>`;
}

function cssFontFamily(name: string, kind: ReturnType<Caches['fontKinds']['get']>): string {
  const chain = fontFamilyWithFallback(name, kind);
  // fontFamilyWithFallback returns a comma-separated list. Emit CSS with
  // single-quoted font names (the outer HTML attribute uses double quotes,
  // so we can't reuse " here without escaping every occurrence).
  return chain
    .split(',')
    .map(s => s.trim().replace(/^"|"$/g, ''))
    .map(s => {
      if (/^(serif|sans-serif|monospace|cursive|fantasy)$/.test(s)) return s;
      if (/[\s'"]/.test(s) || /[^\x20-\x7e]/.test(s)) return `'${s.replace(/'/g, "\\'")}'`;
      return s;
    })
    .join(', ');
}

// ── Images ──

function renderPicture(picEl: Element, caches: Caches): string {
  const w = hu2mm(intAttr(picEl, 'width', 0));
  const h = hu2mm(intAttr(picEl, 'height', 0));
  // In HWPX native format the image ref is on <hc:img binaryItemIDRef="...">,
  // not on <hp:pic> directly. Handle both.
  let ref = attr(picEl, 'binaryItemIDRef', '');
  if (!ref) {
    const imgEl = find(picEl, 'img');
    if (imgEl) ref = attr(imgEl, 'binaryItemIDRef', '');
  }
  if (!ref) return '';

  // The loader keys binData by numeric id parsed from "BINxxxx" filenames.
  // Modern HWPX uses string ids like "image1". Try numeric first, fall back
  // to the raw string via a linear match against the map (rare hot path).
  let href: string | undefined;
  const asNum = parseInt(ref, 10);
  if (!isNaN(asNum)) href = caches.binData?.get(asNum);
  if (!href) {
    // No numeric id match; the loader currently only understands BINxxxx —
    // if the file used string ids, the image just won't be present. Emit an
    // empty placeholder so the surrounding text still flows correctly.
    if (w > 0 && h > 0) {
      return `<span class="hwp-img-missing" style="display:inline-block;width:${w.toFixed(2)}mm;height:${h.toFixed(2)}mm;"></span>`;
    }
    return '';
  }
  const size = w > 0 && h > 0
    ? ` style="width:${w.toFixed(2)}mm;height:${h.toFixed(2)}mm;vertical-align:middle;"`
    : '';
  return `<img src="${escapeHtml(href)}" alt=""${size}/>`;
}

// ── Table rendering ──

function renderTable(tblEl: Element, caches: Caches): string {
  const rows = children(tblEl, 'tr');
  if (rows.length === 0) return '';

  const tableBorderFillId = attr(tblEl, 'borderFillIDRef', '');
  const tableBf = tableBorderFillId ? caches.borderFills.get(tableBorderFillId) : undefined;

  // Column widths (in HWPUNIT) — we output percentage or mm per column so the
  // table lays out at its declared size.
  const colWidths = buildColWidths(rows);
  const cols = [...colWidths.keys()].sort((a, b) => a - b);

  const tableDecl: string[] = [];
  tableDecl.push('border-collapse:collapse');
  // Anchored/inline table with an explicit width from <hp:sz>.
  const szEl = find(tblEl, 'sz');
  const declaredWidth = szEl ? hu2mm(intAttr(szEl, 'width', 0)) : 0;
  if (declaredWidth > 0) tableDecl.push(`width:${declaredWidth.toFixed(2)}mm`);
  tableDecl.push('table-layout:fixed');

  const parts: string[] = [`<table style="${tableDecl.join(';')}">`];

  // Emit <colgroup> so column widths line up even before any cell text pushes
  // things around. Uses the per-column widths derived above.
  if (cols.length > 0) {
    parts.push('<colgroup>');
    for (const c of cols) {
      const wHu = colWidths.get(c) ?? 0;
      const wMm = hu2mm(wHu);
      parts.push(`<col style="width:${wMm.toFixed(2)}mm"/>`);
    }
    parts.push('</colgroup>');
  }

  for (const rowEl of rows) {
    parts.push('<tr>');
    for (const cellEl of findDirectCells(rowEl)) {
      const cellHtml = renderCell(cellEl, caches, tableBf);
      if (cellHtml) parts.push(cellHtml);
    }
    parts.push('</tr>');
  }
  parts.push('</table>');
  return parts.join('');
}

/** Build a colAddr → widthHu map by scanning all rows, exactly like the SVG renderer. */
function buildColWidths(rows: Element[]): Map<number, number> {
  const colWidths = new Map<number, number>();
  for (const rowEl of rows) {
    for (const cellEl of findDirectCells(rowEl)) {
      const addrEl = findChild(cellEl, 'cellAddr');
      const colAddr = addrEl ? intAttr(addrEl, 'colAddr', 0) : intAttr(cellEl, 'colAddr', 0);
      const spanEl = findChild(cellEl, 'cellSpan');
      const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
      const szEl = findChild(cellEl, 'cellSz');
      const wHu = szEl ? intAttr(szEl, 'width', 0) : intAttr(cellEl, 'width', 0);
      if (colSpan === 1 && wHu > 0) colWidths.set(colAddr, wHu);
    }
  }
  // Derive unknown columns from merged cells iteratively.
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
        let known = 0, unknown = -1, unknownCount = 0;
        for (let c = colAddr; c < colAddr + colSpan; c++) {
          if (colWidths.has(c)) known += colWidths.get(c)!;
          else { unknown = c; unknownCount++; }
        }
        if (unknownCount === 1 && unknown >= 0) {
          const derived = totalW - known;
          if (derived > 0 && !colWidths.has(unknown)) {
            colWidths.set(unknown, derived);
            progress = true;
          }
        }
      }
    }
  }
  return colWidths;
}

function renderCell(cellEl: Element, caches: Caches, tableBf: BorderFillInfo | undefined): string {
  const addrEl = findChild(cellEl, 'cellAddr');
  void addrEl;
  const spanEl = findChild(cellEl, 'cellSpan');
  const colSpan = spanEl ? intAttr(spanEl, 'colSpan', 1) : intAttr(cellEl, 'colSpan', 1);
  const rowSpan = spanEl ? intAttr(spanEl, 'rowSpan', 1) : intAttr(cellEl, 'rowSpan', 1);
  const szEl = findChild(cellEl, 'cellSz');
  const cellHeightHu = szEl ? intAttr(szEl, 'height', 0) : intAttr(cellEl, 'height', 0);

  const borderFillId = attr(cellEl, 'borderFillIDRef', '');
  const bf = borderFillId ? caches.borderFills.get(borderFillId) : tableBf;

  const subList = find(cellEl, 'subList');
  const vertAlign = subList ? attr(subList, 'vertAlign', 'TOP').toUpperCase() : 'TOP';

  const marginEl = findChild(cellEl, 'cellMargin');
  const HWP_SENTINEL = 4294967295;
  const HWP_DEFAULT_MARGIN = 141;
  const readMargin = (side: string): number => {
    const raw = marginEl
      ? intAttr(marginEl, side, HWP_SENTINEL)
      : intAttr(cellEl, `margin${side.charAt(0).toUpperCase() + side.slice(1)}`, HWP_SENTINEL);
    return raw === HWP_SENTINEL ? HWP_DEFAULT_MARGIN : raw;
  };
  const mLeft = hu2mm(readMargin('left'));
  const mRight = hu2mm(readMargin('right'));
  const mTop = hu2mm(readMargin('top'));
  const mBottom = hu2mm(readMargin('bottom'));

  const decl: string[] = [];
  decl.push(`padding:${mTop.toFixed(2)}mm ${mRight.toFixed(2)}mm ${mBottom.toFixed(2)}mm ${mLeft.toFixed(2)}mm`);
  if (cellHeightHu > 0) decl.push(`height:${hu2mm(cellHeightHu).toFixed(2)}mm`);
  const va = vertAlign === 'CENTER' ? 'middle' : vertAlign === 'BOTTOM' ? 'bottom' : 'top';
  decl.push(`vertical-align:${va}`);
  const borders = cellBorderStyles(bf);
  decl.push(...borders);
  const fill = cellFillStyle(bf, caches);
  if (fill) decl.push(fill);

  const attrs: string[] = [];
  if (colSpan > 1) attrs.push(`colspan="${colSpan}"`);
  if (rowSpan > 1) attrs.push(`rowspan="${rowSpan}"`);
  attrs.push(`style="${decl.join(';')}"`);

  // Cell content: paragraphs (and possibly nested tables) live under <subList>.
  const cellChildren = subList ?? cellEl;
  const contentParts = renderContainerChildren(cellChildren, caches);
  const content = contentParts.filter(Boolean).join('');
  return `<td ${attrs.join(' ')}>${content}</td>`;
}

function cellBorderStyles(bf: BorderFillInfo | undefined): string[] {
  if (!bf) return [];
  const sides: Array<'left' | 'right' | 'top' | 'bottom'> = ['top', 'right', 'bottom', 'left'];
  const decl: string[] = [];
  for (const side of sides) {
    const type = bf[`${side}BorderType` as keyof BorderFillInfo] as string;
    const width = bf[`${side}BorderWidth` as keyof BorderFillInfo] as string;
    const color = bf[`${side}BorderColor` as keyof BorderFillInfo] as string;
    if (!type || type === 'NONE' || type === 'none') {
      decl.push(`border-${side}:none`);
      continue;
    }
    const cssStyle = mapBorderStyle(type);
    const w = normalizeWidth(width);
    decl.push(`border-${side}:${w} ${cssStyle} ${color}`);
  }
  return decl;
}

function mapBorderStyle(type: string): string {
  const t = type.toUpperCase();
  switch (t) {
    case 'DASHED':
    case 'DASH': return 'dashed';
    case 'DOT': return 'dotted';
    case 'DOUBLE':
    case 'DOUBLE_SLIM': return 'double';
    case 'DASH_DOT':
    case 'DASH_DOT_DOT': return 'dashed';
    default: return 'solid';
  }
}

function normalizeWidth(w: string): string {
  // Sample values: "0.12 mm", "0.5mm". Preserve as-is once whitespace is
  // squeezed out.
  const s = (w || '0.1mm').replace(/\s+/g, '');
  return s;
}

function cellFillStyle(bf: BorderFillInfo | undefined, caches: Caches): string {
  if (!bf) return '';
  if (bf.fillColor) return `background-color:${bf.fillColor}`;
  if (bf.imageBinDataId) {
    const href = caches.binData?.get(bf.imageBinDataId);
    if (href) {
      const size = (bf.imageFillMode ?? 'RESIZE') === 'RESIZE' ? 'cover' : 'contain';
      return `background:url(${href}) center/${size} no-repeat`;
    }
  }
  if (bf.gradientColors && bf.gradientColors.length > 0) {
    if (bf.gradientColors.length === 1) return `background-color:${bf.gradientColors[0]}`;
    const angle = bf.gradientAngle ?? 0;
    // HWP: 0° = top-to-bottom; CSS: 0deg = bottom-to-top, so shift by 180°.
    const cssAngle = (angle + 180) % 360;
    return `background:linear-gradient(${cssAngle}deg, ${bf.gradientColors.join(', ')})`;
  }
  return '';
}

// ── Utility ──

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function defaultCss(page: PageCss, className: string): string {
  const w = escapeHtml(className);
  const pageW = page.pageWmm.toFixed(2);
  const pageH = page.pageHmm.toFixed(2);
  const padT = page.padTopMm.toFixed(2);
  const padR = page.padRightMm.toFixed(2);
  const padB = page.padBottomMm.toFixed(2);
  const padL = page.padLeftMm.toFixed(2);
  return `
@page { size: ${pageW}mm ${pageH}mm; margin: ${padT}mm ${padR}mm ${padB}mm ${padL}mm; }
body { margin: 0; background: #d9d9d9; font-family: '함초롬돋움', 'HCR Dotum', '맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; }
.${w} { width: ${pageW}mm; margin: 8mm auto; padding: ${padT}mm ${padR}mm ${padB}mm ${padL}mm; background: #fff; box-shadow: 0 2px 10px rgba(0,0,0,0.15); box-sizing: border-box; color: #000; overflow-wrap: break-word; }
.${w} p { margin: 0; padding: 0; line-height: 1.6; word-break: keep-all; }
.${w} table { margin: 2mm 0; max-width: 100%; }
.${w} td { vertical-align: top; }
.${w} img { max-width: 100%; height: auto; }
.${w} .hwp-bullet { display: inline-block; }
.${w} .hwp-pagebreak { display: block; height: 0; margin: 12mm -${padL}mm; border-top: 1px dashed #bbb; }
@media print {
  body { background: #fff; }
  .${w} { width: auto; margin: 0; padding: 0; box-shadow: none; }
  .${w} .hwp-pagebreak { border: none; margin: 0; }
}
`;
}

// Silence unused-import warning for CharShapeInfo when TypeScript can't infer
// its use through the exported HtmlOptions type.
void isCJK;
