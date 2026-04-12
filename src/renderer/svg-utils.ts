/**
 * DOM helpers, numeric/text utilities, and font registry for the SVG renderer.
 */

export const HWPUNIT_TO_MM = 25.4 / 7200;

export function hu2mm(v: number): number { return v * HWPUNIT_TO_MM; }

export function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// ── Namespace-agnostic DOM helpers ──

/** Get local name (strip namespace prefix) */
export function localName(el: Element): string {
  return el.localName || el.nodeName.split(':').pop() || el.nodeName;
}

/** getAttribute by local attribute name (namespace-agnostic, case-insensitive) */
export function attr(el: Element, name: string, def = ''): string {
  // Try exact match first
  const v = el.getAttribute(name);
  if (v !== null) return v;
  // Try lowercase
  const lower = name.toLowerCase();
  const v2 = el.getAttribute(lower);
  if (v2 !== null) return v2;
  // Scan all attributes for case-insensitive match
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    const aLocal = (a.localName || a.name.split(':').pop() || '').toLowerCase();
    if (aLocal === lower) return a.value;
  }
  return def;
}

export function intAttr(el: Element, name: string, def = 0): number {
  const v = attr(el, name, '');
  if (!v) return def;
  const n = parseInt(v, 10);
  return isNaN(n) ? def : n;
}

/** First descendant element matching localName (depth-first) */
export function find(parent: Element, lName: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (localName(child) === lName) return child;
    const found = find(child, lName);
    if (found) return found;
  }
  return null;
}

/** All descendant elements matching localName */
export function findAll(parent: Element, lName: string): Element[] {
  const results: Element[] = [];
  function walk(el: Element) {
    for (const child of Array.from(el.children)) {
      if (localName(child) === lName) results.push(child);
      walk(child);
    }
  }
  walk(parent);
  return results;
}

/** Direct children matching localName */
export function children(parent: Element, lName: string): Element[] {
  return Array.from(parent.children).filter(c => localName(c) === lName);
}

/**
 * Direct child lookup — safe for tc elements that contain nested tables.
 * find() does DFS and will recurse into subList, finding inner cellSz/cellSpan/etc
 * before reaching the outer cell's own attributes. Use this for all cell attribute lookups.
 */
export function findChild(parent: Element, lName: string): Element | null {
  return children(parent, lName)[0] ?? null;
}

// ── CJK detection ──

/** Detect if a character is CJK (Korean/Chinese/Japanese). */
export function isCJK(ch: string): boolean {
  const code = ch.charCodeAt(0);
  return (
    (code >= 0xAC00 && code <= 0xD7AF) || // Hangul Syllables
    (code >= 0x1100 && code <= 0x11FF) || // Hangul Jamo
    (code >= 0x3130 && code <= 0x318F) || // Hangul Compatibility Jamo
    (code >= 0x3000 && code <= 0x9FFF) || // CJK Unified + symbols
    (code >= 0xF900 && code <= 0xFAFF) || // CJK Compatibility
    (code >= 0xFF00 && code <= 0xFFEF)    // Fullwidth Forms
  );
}

// ── Text measurement ──

/** Estimate text width in mm based on font size (in mm) and character types */
export function estimateTextWidth(text: string, fontSizeMm: number): number {
  let width = 0;
  for (const ch of text) {
    if (ch === ' ') {
      width += fontSizeMm * 0.4;  // Spaces are always Latin-width
    } else if (isCJK(ch)) {
      width += fontSizeMm;  // CJK chars are roughly square
    } else {
      width += fontSizeMm * 0.4;  // Latin chars in Korean fonts are ~40% width
    }
  }
  return width;
}

/** Word-wrap text into lines that fit within widthMm, using font size in mm */
export function wrapText(text: string, widthMm: number, fontSizeMm: number): string[] {
  if (!text) return [''];
  if (widthMm <= 0) return [text];

  // Quick check: does entire text fit?
  if (estimateTextWidth(text, fontSizeMm) <= widthMm) return [text];

  const lines: string[] = [];
  let lineStart = 0;
  let lineWidth = 0;
  let lastBreakable = -1;  // index of last space or CJK boundary

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const charW = ch === ' ' ? fontSizeMm * 0.4 : (isCJK(ch) ? fontSizeMm : fontSizeMm * 0.4);

    if (lineWidth + charW > widthMm && i > lineStart) {
      // Need to break
      if (lastBreakable > lineStart) {
        // Break at last breakable position
        const breakAt = text[lastBreakable] === ' ' ? lastBreakable + 1 : lastBreakable;
        lines.push(text.substring(lineStart, breakAt));
        lineStart = breakAt;
      } else {
        // No breakable position found, force break here
        lines.push(text.substring(lineStart, i));
        lineStart = i;
      }
      lineWidth = 0;
      // Recalculate width for carried-over characters
      for (let j = lineStart; j <= i; j++) {
        if (j < text.length) {
          const rch = text[j];
          lineWidth += rch === ' ' ? fontSizeMm * 0.4 : (isCJK(rch) ? fontSizeMm : fontSizeMm * 0.4);
        }
      }
      lastBreakable = -1;
    } else {
      lineWidth += charW;
    }

    if (ch === ' ') lastBreakable = i;
    else if (isCJK(ch)) lastBreakable = i;  // CJK chars can break at any boundary
  }

  // Remaining text
  if (lineStart < text.length) {
    lines.push(text.substring(lineStart));
  }

  return lines.length > 0 ? lines : [''];
}

// ── Font registry ──

/** Map of font-family name → base64-encoded font data (data URI) */
export const fontRegistry = new Map<string, string>();

/**
 * Register a font for embedding in SVG output.
 * @param familyName The CSS font-family name (e.g. "HY헤드라인M")
 * @param base64Data Base64-encoded font file content
 * @param format Font format: "truetype", "woff2", etc.
 */
export function registerFont(familyName: string, base64Data: string, format = 'truetype'): void {
  fontRegistry.set(familyName, `data:font/${format};base64,${base64Data}`);
}

/** Clear all registered fonts. */
export function clearFonts(): void {
  fontRegistry.clear();
}

/** Build @font-face CSS and global text styles for SVG. */
export function buildFontFaceCSS(): string {
  const rules: string[] = [];
  for (const [name, dataUri] of fontRegistry) {
    rules.push(`@font-face { font-family: "${name}"; src: url("${dataUri}"); }`);
  }
  // HWP renders spaces slightly wider than the SVG default font metrics.
  // An empirical 0.025em offset brings Korean body text widths in line with
  // the reference PDF rendering.
  // No global word-spacing: HWP uses font's natural space width (per-character spacing via charShape)
  return `<defs><style type="text/css">\n${rules.join('\n')}\n</style></defs>`;
}

/** Wrap font name with fallback chain for SVG font-family attribute.
 *  Primary fallback: 한컴바탕 (Haansoft Batang), then 함초롬바탕/돋움. */
export function fontFamilyWithFallback(name: string): string {
  return `${name}, 한컴바탕, Haansoft Batang, 함초롬바탕, HCR Batang, 함초롬돋움, HCR Dotum, Malgun Gothic, 맑은 고딕, Apple SD Gothic Neo, Noto Sans KR, sans-serif`;
}

/** Get font name for a character based on script detection. */
export function fontForChar(cs: { fontName: string; fontNameLatin: string }, ch: string): string {
  return isCJK(ch) ? cs.fontName : cs.fontNameLatin;
}

/** Get spacing for a character based on script detection. */
export function spacingForChar(cs: { spacing: number; spacingLatin: number }, ch: string): number {
  return isCJK(ch) ? cs.spacing : cs.spacingLatin;
}

/** Get ratio (width scale %) for a character based on script detection. */
export function ratioForChar(cs: { ratio: number; ratioLatin: number }, ch: string): number {
  return isCJK(ch) ? cs.ratio : cs.ratioLatin;
}
