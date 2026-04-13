/**
 * Build caches (charShapes, paraShapes, borderFills) from header DOM.
 */

import type { CharShapeInfo, ParaShapeInfo, BorderFillInfo, Caches } from './svg-types.js';
import { attr, intAttr, find, findAll, children } from './svg-utils.js';

// ── Alignment string → int mapping ──

const ALIGN_MAP: Record<string, number> = {
  'JUSTIFY': 0,
  'LEFT': 1,
  'RIGHT': 2,
  'CENTER': 3,
  'DISTRIBUTE': 4,
  'DISTRIBUTE_SPACE': 4,
};

export function parseAlign(el: Element, attrName: string, def: number): number {
  const v = attr(el, attrName, '');
  if (!v) return def;
  // Try numeric first
  const n = parseInt(v, 10);
  if (!isNaN(n)) return n;
  // Map string value
  return ALIGN_MAP[v.toUpperCase()] ?? def;
}

// ── Build caches from header DOM ──

export function buildCaches(header: Document): Caches {
  const root = header.documentElement;

  // Build per-language font maps from HWPX native format:
  // <hh:fontface lang="HANGUL"><hh:font id="N" face="..."/></hh:fontface>
  // Also support legacy converter format: <hh:fontface id="N" name="..."/>
  const langFontMaps = new Map<string, Map<string, string>>();
  const legacyFontMap = new Map<string, string>();
  for (const fontfaceEl of findAll(root, 'fontface')) {
    const lang = attr(fontfaceEl, 'lang', '');
    const legacyId = attr(fontfaceEl, 'id', '');
    if (legacyId) {
      // Legacy: <hh:fontface id="N" name="..."/>
      const face = attr(fontfaceEl, 'name', '') || attr(fontfaceEl, 'face', '');
      if (face) legacyFontMap.set(legacyId, face);
    } else if (lang) {
      // HWPX native: per-language font list
      const langMap = new Map<string, string>();
      for (const fontEl of children(fontfaceEl, 'font')) {
        const fid = attr(fontEl, 'id', '');
        const face = attr(fontEl, 'face', '') || attr(fontEl, 'name', '');
        if (fid && face) langMap.set(fid, face);
      }
      langFontMaps.set(lang.toUpperCase(), langMap);
    }
  }
  // Helper: resolve font face by language and id
  function resolveFontFace(lang: string, id: string): string {
    const langMap = langFontMaps.get(lang.toUpperCase());
    if (langMap && id && langMap.has(id)) return langMap.get(id)!;
    if (legacyFontMap.has(id)) return legacyFontMap.get(id)!;
    return 'sans-serif';
  }

  // Build charShape cache
  const charShapes = new Map<string, CharShapeInfo>();
  for (const el of findAll(root, 'charPr')) {
    const id = attr(el, 'id', '');
    if (!id) continue;

    const height = intAttr(el, 'height', 1000);
    const textColor = attr(el, 'textColor', '#000000');
    // Bold/italic: may be attribute (legacy) or child element (HWPX native: <hh:bold/>, <hh:italic/>)
    const boldAttr = attr(el, 'bold', '');
    const bold = boldAttr === '1' || boldAttr === 'true' || find(el, 'bold') !== null;
    const italicAttr = attr(el, 'italic', '');
    const italic = italicAttr === '1' || italicAttr === 'true' || find(el, 'italic') !== null;

    // Font: look for fontRef child with per-language IDs
    let fontName = 'sans-serif';
    let fontNameLatin = 'sans-serif';
    let spacing = 0;
    let spacingLatin = 0;
    let ratio = 100;
    let ratioLatin = 100;

    const fontRef = find(el, 'fontRef');
    const ratioEl = find(el, 'ratio');
    const spacingEl = find(el, 'spacing');

    if (fontRef) {
      const hangulId = attr(fontRef, 'hangul', '');
      const latinId = attr(fontRef, 'latin', '');
      if (hangulId) fontName = resolveFontFace('HANGUL', hangulId);
      if (latinId) fontNameLatin = resolveFontFace('LATIN', latinId);
      else fontNameLatin = fontName;
    } else {
      const fontId = attr(el, 'fontId', '');
      if (fontId) {
        fontName = legacyFontMap.get(fontId) ?? 'sans-serif';
        fontNameLatin = fontName;
      }
    }

    if (ratioEl) {
      ratio = intAttr(ratioEl, 'hangul', 100);
      ratioLatin = intAttr(ratioEl, 'latin', 100);
    }
    if (spacingEl) {
      spacing = intAttr(spacingEl, 'hangul', 0);
      spacingLatin = intAttr(spacingEl, 'latin', 0);
    } else {
      spacing = intAttr(el, 'spacing', 0);
      spacingLatin = spacing;
    }

    charShapes.set(id, { height, textColor, bold, italic, fontName, fontNameLatin, spacing, spacingLatin, ratio, ratioLatin });
  }

  // Build paraShape cache
  const paraShapes = new Map<string, ParaShapeInfo>();
  for (const el of findAll(root, 'paraPr')) {
    const id = attr(el, 'id', '');
    if (!id) continue;

    // HWP converter: flat attributes like align="JUSTIFY", leftMargin="0"
    let alignment = parseAlign(el, 'align', 0);
    let leftMargin = intAttr(el, 'leftMargin', intAttr(el, 'marginLeft', 0));
    let rightMargin = intAttr(el, 'rightMargin', intAttr(el, 'marginRight', 0));
    let indent = intAttr(el, 'indent', 0);
    let spacingBefore = intAttr(el, 'spacingBefore', 0);
    let spacingAfter = intAttr(el, 'spacingAfter', 0);
    let lineSpacing = intAttr(el, 'lineSpacing', 160);

    // HWPX format: nested child elements <hh:align horizontal="...">,
    // <hh:margin> with <hc:left value="...">, <hh:lineSpacing value="...">
    const alignEl = find(el, 'align');
    if (alignEl) {
      alignment = parseAlign(alignEl, 'horizontal', alignment);
    }

    // Check for nested margin element (HWPX uses <hh:margin> with <hc:left>, <hc:right>, etc.)
    const marginEl = find(el, 'margin');
    if (marginEl) {
      // HWPX: child elements like <hc:left value="..."/>, <hc:right value="..."/>
      const leftEl = find(marginEl, 'left');
      const rightEl = find(marginEl, 'right');
      const intentEl = find(marginEl, 'intent');
      const prevEl = find(marginEl, 'prev');
      const nextEl = find(marginEl, 'next');
      if (leftEl) leftMargin = intAttr(leftEl, 'value', leftMargin);
      else leftMargin = intAttr(marginEl, 'left', leftMargin);
      if (rightEl) rightMargin = intAttr(rightEl, 'value', rightMargin);
      else rightMargin = intAttr(marginEl, 'right', rightMargin);
      if (intentEl) indent = intAttr(intentEl, 'value', indent);
      else indent = intAttr(marginEl, 'indent', indent);
      if (prevEl) spacingBefore = intAttr(prevEl, 'value', spacingBefore);
      if (nextEl) spacingAfter = intAttr(nextEl, 'value', spacingAfter);
    }

    // Also check older HWPX format with paraMargin/paraSpacing
    const paraMarginEl = find(el, 'paraMargin');
    if (paraMarginEl) {
      leftMargin = intAttr(paraMarginEl, 'left', leftMargin);
      rightMargin = intAttr(paraMarginEl, 'right', rightMargin);
      indent = intAttr(paraMarginEl, 'indent', indent);
    }
    const spacingEl = find(el, 'paraSpacing');
    if (spacingEl) {
      spacingBefore = intAttr(spacingEl, 'before', spacingBefore);
      spacingAfter = intAttr(spacingEl, 'after', spacingAfter);
      lineSpacing = intAttr(spacingEl, 'lineSpacing', lineSpacing);
    }

    // HWPX lineSpacing child element
    const lineSpacingEl = find(el, 'lineSpacing');
    if (lineSpacingEl) {
      const lsVal = intAttr(lineSpacingEl, 'value', 0);
      if (lsVal > 0) lineSpacing = lsVal;
    }

    // 한 줄로 입력 (LineWrap): 'Squeeze' = adjust spacing to fit single line; 'FIXED' (binary bit=1) also maps here
    const lineWrapRaw = attr(el, 'lineWrap', attr(el, 'LineWrap', ''));
    const lineWrap = lineWrapRaw ? lineWrapRaw.toUpperCase() : undefined;

    paraShapes.set(id, {
      alignment,
      leftMargin,
      rightMargin,
      indent,
      spacingBefore,
      spacingAfter,
      lineSpacing,
      lineWrap,
    });
  }

  // Build borderFill cache
  // HWP binary (via converter) stores border info as flat attributes on <borderFill>.
  // HWPX stores them as child elements: <leftBorder type="NONE" width="0.12 mm" color="#000000"/>
  // We handle both: child element takes precedence over flat attribute.
  const borderFills = new Map<string, BorderFillInfo>();
  for (const el of findAll(root, 'borderFill')) {
    const id = attr(el, 'id', '');
    if (!id) continue;

    // Legacy format: flat fillColor attribute
    // HWPX native: <hc:fillBrush><hc:winBrush faceColor="#RRGGBB" .../></hc:fillBrush>
    let fillColor: string | null = null;
    const fillColorAttr = attr(el, 'fillColor', '');
    if (fillColorAttr && fillColorAttr !== 'none') {
      fillColor = fillColorAttr;
    } else {
      const fillBrush = find(el, 'fillBrush');
      if (fillBrush) {
        const winBrush = find(fillBrush, 'winBrush');
        if (winBrush) {
          const fc = attr(winBrush, 'faceColor', 'none');
          if (fc && fc !== 'none') fillColor = fc;
        }
      }
    }

    // Helper: read border side from child element (HWPX) or flat attribute (HWP converter)
    function borderSide(side: string): { type: string; width: string; color: string } {
      const child = children(el, `${side}Border`)[0];
      if (child) {
        return {
          type: attr(child, 'type', '0'),
          width: attr(child, 'width', '0.1mm').replace(' ', ''),
          color: attr(child, 'color', '#000000'),
        };
      }
      return {
        type: attr(el, `${side}BorderType`, '0'),
        width: attr(el, `${side}BorderWidth`, '0.1mm'),
        color: attr(el, `${side}BorderColor`, '#000000'),
      };
    }

    const left = borderSide('left');
    const right = borderSide('right');
    const top = borderSide('top');
    const bottom = borderSide('bottom');

    borderFills.set(id, {
      fillColor,
      leftBorderType: left.type,
      leftBorderWidth: left.width,
      leftBorderColor: left.color,
      rightBorderType: right.type,
      rightBorderWidth: right.width,
      rightBorderColor: right.color,
      topBorderType: top.type,
      topBorderWidth: top.width,
      topBorderColor: top.color,
      bottomBorderType: bottom.type,
      bottomBorderWidth: bottom.width,
      bottomBorderColor: bottom.color,
    });
  }

  return { charShapes, paraShapes, borderFills };
}

// ── Page dimensions ──

export function getPageDims(sectionRoot: Element): import('./svg-types.js').PageDims {
  // Defaults for A4
  const DEF_W = 59528, DEF_H = 84188;
  const DEF_ML = 5669, DEF_MR = 5669, DEF_MT = 2834, DEF_MB = 2834;

  const pagePr = find(sectionRoot, 'pagePr');
  let width = DEF_W, height = DEF_H;
  let marginLeft = DEF_ML, marginRight = DEF_MR, marginTop = DEF_MT, marginBottom = DEF_MB;

  // Also check for pageDef (HWP converter uses <hs:pageDef>)
  const pageDef = find(sectionRoot, 'pageDef');
  const pageSource = pagePr || pageDef;

  let marginHeader = 0, marginFooter = 0;

  if (pageSource) {
    width = intAttr(pageSource, 'width', DEF_W);
    height = intAttr(pageSource, 'height', DEF_H);

    const marginEl = find(pageSource, 'margin');
    if (marginEl) {
      marginLeft = intAttr(marginEl, 'left', DEF_ML);
      marginRight = intAttr(marginEl, 'right', DEF_MR);
      marginTop = intAttr(marginEl, 'top', DEF_MT);
      marginBottom = intAttr(marginEl, 'bottom', DEF_MB);
      marginHeader = intAttr(marginEl, 'header', 0);
      marginFooter = intAttr(marginEl, 'footer', 0);
    } else {
      marginLeft = intAttr(pageSource, 'marginLeft', DEF_ML);
      marginRight = intAttr(pageSource, 'marginRight', DEF_MR);
      marginTop = intAttr(pageSource, 'marginTop', DEF_MT);
      marginBottom = intAttr(pageSource, 'marginBottom', DEF_MB);
      marginHeader = intAttr(pageSource, 'marginHeader', 0);
      marginFooter = intAttr(pageSource, 'marginFooter', 0);
    }
  }

  // Import hu2mm inline to avoid circular dep
  const HU_TO_MM = 25.4 / 7200;
  const h2m = (v: number) => v * HU_TO_MM;

  return {
    pageW: h2m(width),
    pageH: h2m(height),
    contentLeft: h2m(marginLeft),
    contentTop: h2m(marginTop + marginHeader),
    contentWidth: h2m(width - marginLeft - marginRight),
    pageBottom: h2m(height - marginBottom - marginFooter),
  };
}
