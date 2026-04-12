/**
 * Shared type definitions for the SVG renderer.
 */

export interface CharShapeInfo {
  height: number;      // HWPUNIT
  textColor: string;   // CSS color string
  bold: boolean;
  italic: boolean;
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
}

export interface BorderFillInfo {
  fillColor: string | null;
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
