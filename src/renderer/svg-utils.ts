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
  if (!ch) return false;
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

/** Estimate text width in mm based on font size (in mm) and character types.
 *  Korean font space glyph is roughly half-em (500/1000) which is wider than
 *  a typical Latin space (~400/1000). Using 0.5em matches what Hancom renders
 *  when the paragraph is set in a Korean font. */
export function estimateTextWidth(text: string, fontSizeMm: number): number {
  let width = 0;
  for (const ch of text) {
    if (ch === ' ') {
      width += fontSizeMm * 0.5;  // Korean-font space is ~half-em
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

/** Build @font-face CSS for any embedded fonts in the SVG. */
export function buildFontFaceCSS(): string {
  const rules: string[] = [];
  for (const [name, dataUri] of fontRegistry) {
    rules.push(`@font-face { font-family: "${name}"; src: url("${dataUri}"); }`);
  }
  return `<defs><style type="text/css">\n${rules.join('\n')}\n</style></defs>`;
}

import type { FontKind } from './svg-types.js';

/** HWP typeInfo familyType → CSS generic family bucket.
 *  These constants come from the HWP-XML typeInfo attribute (values like
 *  "FCAT_GOTHIC", "FCAT_MYUNGJO", ...). See Hancom's PANOSE-inspired taxonomy. */
const FAMILY_TYPE_MAP: Record<string, FontKind> = {
  FCAT_MYUNGJO:    'serif',       // 명조 – Korean serif
  FCAT_GOTHIC:     'sans-serif',  // 고딕 – Korean sans-serif
  FCAT_GRAPHIC:    'sans-serif',  // graphic – treat as sans
  FCAT_ROMAN:      'serif',       // Latin serif
  FCAT_SWISS:      'sans-serif',  // Latin sans
  FCAT_MODERN:     'monospace',   // typewriter / modern (monospaced)
  FCAT_SCRIPT:     'cursive',
  FCAT_DECORATIVE: 'fantasy',
  FCAT_SYMBOL:     'sans-serif',
  FCAT_ANY:        'sans-serif',
};

export function familyTypeToKind(familyType: string): FontKind | undefined {
  return FAMILY_TYPE_MAP[familyType.toUpperCase()];
}

/** Fallback classification by common Korean/Latin font naming conventions. */
export function classifyFontByName(name: string): FontKind {
  const n = name.toLowerCase();
  // Korean serif keywords
  if (/(바탕|명조|신명|궁서|batang|myungjo|myeongjo)/.test(name.toLowerCase())) return 'serif';
  // Korean sans-serif keywords
  if (/(돋움|고딕|굴림|dotum|gothic|gulim|malgun|맑은)/.test(name.toLowerCase())) return 'sans-serif';
  // Monospace
  if (/(mono|courier|consolas|d2coding|나눔고딕코딩|맑은고딕코딩)/.test(n)) return 'monospace';
  // Latin serif keywords
  if (/(serif|times|garamond|georgia|book antiqua|palatino|cambria)/.test(n)) return 'serif';
  // Latin sans keywords
  if (/(sans|arial|helvetica|verdana|tahoma|calibri|geneva|noto sans)/.test(n)) return 'sans-serif';
  // Script/decorative
  if (/(script|hand|brush|italic hand)/.test(n)) return 'cursive';
  // Default: sans-serif (matches SVG default and prior behavior)
  return 'sans-serif';
}

// Per-kind fallback chains. Each chain ends in the CSS generic family so
// browsers without any Korean font installed still pick a face of the right
// visual class (serif vs. sans-serif) instead of the default sans.
const FALLBACK_CHAINS: Record<FontKind, string> = {
  'serif':
    '한컴바탕, Haansoft Batang, 함초롬바탕, HCR Batang, 바탕, Batang, 나눔명조, NanumMyeongjo, 신명조, Noto Serif KR, Noto Serif CJK KR, "Times New Roman", Georgia, serif',
  'sans-serif':
    '함초롬돋움, HCR Dotum, 맑은 고딕, Malgun Gothic, 나눔고딕, NanumGothic, 돋움, Dotum, 굴림, Gulim, Apple SD Gothic Neo, Noto Sans KR, Noto Sans CJK KR, Arial, Helvetica, sans-serif',
  'monospace':
    'D2Coding, "Nanum Gothic Coding", Consolas, "Courier New", monospace',
  'cursive':
    '"Nanum Pen Script", "Nanum Brush Script", cursive',
  'fantasy':
    'Impact, fantasy',
};

/** Wrap font name with a fallback chain appropriate for its kind.
 *  If kind is omitted, classifies by name. */
export function fontFamilyWithFallback(name: string, kind?: FontKind): string {
  const k: FontKind = kind ?? classifyFontByName(name);
  // Deduplicate: if the primary name already appears in the chain, don't add
  // it twice — otherwise cascade behaves normally.
  const chain = FALLBACK_CHAINS[k];
  return `${name}, ${chain}`;
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
