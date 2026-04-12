/**
 * Type definitions for HWP binary parsing and XML generation.
 */

// ── DocInfo types ──

export interface FontFaceInfo {
  name: string;
}

export interface CharShapeInfo {
  fontIds: number[];
  height: number;
  textColor: number;
  bold: boolean;
  italic: boolean;
  underlineType: number;
  strikeout: number;
  superscript: boolean;
  subscript: boolean;
  spacing: number[];
  relSize: number[];
  offset: number[];
  ratio: number[];
}

export interface ParaShapeInfo {
  attrs1: number;
  leftMargin: number;
  rightMargin: number;
  indent: number;
  spacingBefore: number;
  spacingAfter: number;
  lineSpacing: number;
  alignment: number;
  tabDefId: number;
  borderFillId: number;
  lineWrap: number;  // from attrs2 bit 0~1: 0=Break, 1=Squeeze (한 줄로 입력), 2=Keep
}

export interface BorderFillInfo {
  id: number;
  leftBorderType: number;
  rightBorderType: number;
  topBorderType: number;
  bottomBorderType: number;
  leftBorderWidth: number;
  rightBorderWidth: number;
  topBorderWidth: number;
  bottomBorderWidth: number;
  leftBorderColor: number;
  rightBorderColor: number;
  topBorderColor: number;
  bottomBorderColor: number;
  fillColor: number | null;
}

export interface BinDataItemInfo {
  type: 'link' | 'embedding' | 'storage';
  binDataId?: number;
  extension?: string;
  absolutePath?: string;
  relativePath?: string;
}

export interface DocInfoData {
  sectionCount: number;
  fonts: FontFaceInfo[];
  charShapes: CharShapeInfo[];
  paraShapes: ParaShapeInfo[];
  borderFills: BorderFillInfo[];
  binDataItems: BinDataItemInfo[];
}

// ── Section types ──

export interface PageDefInfo {
  width: number;
  height: number;
  landscape: boolean;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  marginHeader: number;
  marginFooter: number;
  marginGutter: number;
}

export interface TextRunInfo {
  charPrId: number;
  text: string;
}

export interface LineSegInfo {
  textPos: number;
  vertPos: number;
  vertSize: number;
  textHeight: number;
  baseline: number;
  spacing: number;
  horzPos: number;
  horzSize: number;
  flags: number;
}

export interface ParaInfo {
  paraPrId: number;
  styleId: number;
  runs: TextRunInfo[];
  lineSegs: LineSegInfo[];
  controls: ControlInfo[];
}

export type ControlInfo = TableControlInfo | SectionDefInfo;

export interface TableControlInfo {
  type: 'table';
  rowCount: number;
  colCount: number;
  cellSpacing: number;
  borderFillId: number;
  cells: CellInfo[];
  rowSizes: number[];
  innerMarginLeft: number;
  innerMarginRight: number;
  innerMarginTop: number;
  innerMarginBottom: number;
  ctrlWidth: number;
  ctrlHeight: number;
  outMarginLeft: number;
  outMarginRight: number;
  outMarginTop: number;
  outMarginBottom: number;
  captionParas?: ParaInfo[];
  captionGap?: number;
  captionDir?: number;
}

export interface CellInfo {
  colAddr: number;
  rowAddr: number;
  colSpan: number;
  rowSpan: number;
  width: number;
  height: number;
  borderFillId: number;
  hasMargin: boolean;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  vertAlign: number; // 0=TOP, 1=CENTER, 2=BOTTOM
  lineWrap: number;  // 0=Break, 1=Squeeze (한 줄로 입력), 2=Keep
  paragraphs: ParaInfo[];
}

export interface SectionDefInfo {
  type: 'secd';
  pageDef: PageDefInfo;
}

// ── FileHeader type ──

export interface DocHeader {
  version: { major: number; minor: number; patch: number; revision: number };
  flags: number;
}
