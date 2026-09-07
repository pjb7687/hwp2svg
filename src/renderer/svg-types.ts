/**
 * Shared type definitions for the SVG renderer.
 */

export type FontKind = 'serif' | 'sans-serif' | 'monospace' | 'cursive' | 'fantasy';

export interface CharShapeInfo {
  height: number;      // HWPUNIT
  textColor: string;   // CSS color string
  bold: boolean;
  italic: boolean;
  underline: boolean;  // whether text has an underline decoration
  underlineColor: string; // color of the underline (defaults to textColor)
  underlineShape: string; // SOLID | DOUBLE | DOT | DASH | ... (for text-decoration-style)
  strikeout: boolean;  // whether text has a strikethrough decoration
  fontName: string;     // hangul font name
  fontNameLatin: string; // latin font name
  spacing: number;     // hangul spacing percentage
  spacingLatin: number; // latin spacing percentage
  ratio: number;       // hangul width ratio (100=normal)
  ratioLatin: number;  // latin width ratio
}

export interface ParaShapeInfo {
  alignment: number;       // 0=JUSTIFY, 1=LEFT, 2=RIGHT, 3=CENTER
  leftMargin: number;      // HWPUNIT
  rightMargin: number;
  indent: number;
  spacingBefore: number;
  spacingAfter: number;
  lineSpacing: number;
  lineWrap?: string;       // 'SQUEEZE' = 한 줄로 입력 (compress spacing to single line)
  /** id of a bullet definition (see Caches.bullets). 0 = no bullet. */
  bulletId: number;
}

export interface BorderFillInfo {
  fillColor: string | null;
  /** Gradient fill (HWP 표 28). Populated from <hc:gradation>. */
  gradientType?: string;
  gradientAngle?: number;
  gradientColors?: string[];
  /** Image fill (HWP 표 28, bit 1). Populated from <hc:imgBrush><hc:image binaryItemIDRef="N"/></hc:imgBrush>. */
  imageFillMode?: string;
  imageBinDataId?: number;
  leftBorderType: string;
  leftBorderWidth: string;   // e.g. "0.3mm"
  leftBorderColor: string;
  rightBorderType: string;
  rightBorderWidth: string;
  rightBorderColor: string;
  topBorderType: string;
  topBorderWidth: string;
  topBorderColor: string;
  bottomBorderType: string;
  bottomBorderWidth: string;
  bottomBorderColor: string;
}

export interface Caches {
  charShapes: Map<string, CharShapeInfo>;
  paraShapes: Map<string, ParaShapeInfo>;
  borderFills: Map<string, BorderFillInfo>;
  /** Kind of each known font face (name → serif|sans-serif|monospace|...).
   *  Populated from <hh:font typeInfo familyType="..."/> in the HWPX header
   *  and, as a fallback, from Korean/Latin font naming conventions. */
  fontKinds: Map<string, FontKind>;
  /** id → bullet glyph string. Populated from <hh:bullet id="N" char="..."/>.
   *  Characters are usually in Hancom PUA; render call sites should route
   *  through mapPuaStringToUnicode before display. */
  bullets: Map<string, string>;
  /** Embedded images keyed by BinData id → data URI (data:image/xxx;base64,...).
   *  Populated at load time by the HWPX loader. */
  binData: Map<number, string>;
}

export interface PageDims {
  pageW: number;   // mm
  pageH: number;   // mm
  contentLeft: number;   // mm
  contentTop: number;    // mm
  contentWidth: number;  // mm
  pageBottom: number;    // mm
}

export interface LayoutItem {
  svg: string;
  y: number;
  height: number;
}

export interface TableResult {
  svg: string;
  height: number;
}
