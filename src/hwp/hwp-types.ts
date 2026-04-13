/**
 * Type definitions for HWP binary parsing and XML generation.
 */

// ── DocInfo types ──

export interface FontTypeInfo {
  familyType: number;
  serifType: number;
  weight: number;
  proportion: number;
  contrast: number;
  strokeVariation: number;
  armStyle: number;
  letterform: number;
  midline: number;
  xHeight: number;
}

export interface SubstFontInfo {
  type: number;  // 0=unknown, 1=TTF, 2=HFT
  name: string;
}

export interface FontFaceInfo {
  name: string;
  fontType: number;      // 0=unknown, 1=TTF, 2=HFT
  typeInfo?: FontTypeInfo;
  substFont?: SubstFontInfo;
  baseFontName?: string;
}

export interface CharShapeInfo {
  fontIds: number[];
  height: number;
  textColor: number;
  bold: boolean;
  italic: boolean;
  underlineType: number;
  underlineShape: number;
  underlineColor: number;
  shadeColor: number;
  shadowColor: number;
  shadowX: number;
  shadowY: number;
  outlineType: number;
  shadowType: number;
  symMark: number;
  useFontSpace: boolean;
  useKerning: boolean;
  strikeout: number;
  strikeoutShape: number;
  strikeoutColor: number;
  superscript: boolean;
  subscript: boolean;
  borderFillId: number;
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
  numberingId: number;
  borderFillId: number;
  borderLeft: number;
  borderRight: number;
  borderTop: number;
  borderBottom: number;
  lineWrap: number;  // 0=Break, 1=Squeeze, 2=Keep
  autoSpacingEng: boolean;
  autoSpacingNum: boolean;
  lineSpacingType: number;  // 0=PERCENT, 1=FIXED, 2=MINIMUM, 3=AT_LEAST
  // Derived from attrs1
  snapToGrid: boolean;
  condense: number;
  fontLineHeight: boolean;
  vertAlign: number;       // 0=BASELINE, 1=TOP, 2=CENTER, 3=BOTTOM
  breakLatinWord: number;  // 0=KEEP_WORD, 1=HYPHENATE, 2=BREAK_ALL
  breakNonLatinWord: boolean; // false=BREAK_WORD, true=KEEP_WORD
  widowOrphan: boolean;
  keepWithNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  headingType: number;     // 0=NONE, 1=OUTLINE, 2=NUMBERING, 3=BULLET
  headingLevel: number;    // 0-6
  borderConnect: boolean;
  ignoreMargin: boolean;
}

export interface BorderFillInfo {
  id: number;
  threeD: boolean;
  shadow: boolean;
  slashType: number;
  backSlashType: number;
  slashCrooked: number;
  backSlashCrooked: boolean;
  slashCounter: boolean;
  backSlashCounter: boolean;
  centerLine: boolean;
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
  diagonalType: number;
  diagonalWidth: number;
  diagonalColor: number;
  fillColor: number | null;
  fillBackColor: number | null;
  fillPatternType: number | null;
}

export interface TabItemInfo {
  pos: number;
  type: number;    // 0=LEFT, 1=RIGHT, 2=CENTER, 3=DECIMAL
  leader: number;  // 0=NONE, etc.
}

export interface TabDefInfo {
  autoTabLeft: boolean;
  autoTabRight: boolean;
  items: TabItemInfo[];
}

export interface StyleInfo {
  name: string;
  engName: string;
  type: number;        // 0=PARA, 1=CHAR
  nextStyleId: number;
  langId: number;
  paraPrId: number;
  charPrId: number;
}

export interface BulletInfo {
  id: number;
  char: string;
  useImage: boolean;
  level: number;
  align: number;
  useInstWidth: boolean;
  autoIndent: boolean;
  widthAdjust: number;
  textOffsetType: number;
  textOffset: number;
  numFormat: number;
  charPrIdRef: number;
  checkable: boolean;
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
  beginPage: number;
  beginFootnote: number;
  beginEndnote: number;
  beginPic: number;
  beginTbl: number;
  beginEquation: number;
  fontCounts: number[];  // [HANGUL, LATIN, HANJA, JAPANESE, OTHER, SYMBOL, USER]
  fonts: FontFaceInfo[];
  charShapes: CharShapeInfo[];
  paraShapes: ParaShapeInfo[];
  borderFills: BorderFillInfo[];
  tabDefs: TabDefInfo[];
  styles: StyleInfo[];
  bullets: BulletInfo[];
  binDataItems: BinDataItemInfo[];
}

// ── Section types ──

export interface PageDefInfo {
  width: number;
  height: number;
  landscape: number;   // 0=WIDELY, 1=LANDSCAPE
  gutterType: number;  // 0=LEFT_ONLY, 1=BOTH_SIDES, 2=TOP
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  marginHeader: number;
  marginFooter: number;
  marginGutter: number;
}

export interface FootnoteShapeInfo {
  numberType: number;
  placement: number;
  numbering: number;
  supscript: boolean;
  beneathText: boolean;
  userChar: number;
  prefixChar: number;
  suffixChar: number;
  startNumber: number;
  noteLineLength: number;
  noteLineTop: number;
  noteLineBottom: number;
  noteSpacing: number;
  lineType: number;
  lineWidth: number;
  lineColor: number;
}

export interface PageBorderFillInfo {
  textBorder: number;
  headerInside: boolean;
  footerInside: boolean;
  fillArea: number;
  leftGap: number;
  rightGap: number;
  topGap: number;
  bottomGap: number;
  borderFillId: number;
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
  paraId: number;
  pageBreak: boolean;
  columnBreak: boolean;
  merged: number;
  defaultCharPrId: number;
  paraBreakCharPrId: number;
  runs: TextRunInfo[];
  lineSegs: LineSegInfo[];
  controls: ControlInfo[];
  ctrlCharPrIds: number[];  // charPrId for each control in controls[], in same order
  ctrlStreamPositions: number[];  // logical stream positions for each control in controls[], same order
  textBeforeCtrl: boolean;  // true if text runs appear before ctrl chars in binary
}

export interface ColDefInfo {
  type: 'cold';
  colType: number;
  colCount: number;
  layout: number;
  sameSz: boolean;
  sameGap: number;
}

export interface PageNumInfo {
  type: 'pgnp';
  pos: number;
  formatType: number;
  sideChar: number;
}

export interface HeaderFooterInfo {
  type: 'head' | 'foot';
  id: number;
  applyPageType: number;
  textWidth: number;
  textHeight: number;
  paragraphs: ParaInfo[];
}

export interface FieldBeginControlInfo {
  type: 'fieldBegin';
  ctrlId: string;   // e.g. '%hlk' for hyperlink
  id: number;       // unique document-level id
  command: string;  // command string from CTRL_HEADER
  editable: boolean;
  dirty: boolean;
}

export interface FieldEndControlInfo {
  type: 'fieldEnd';
  ctrlId: string;
  beginId: number;  // id of the matching fieldBegin
}

export type ControlInfo = TableControlInfo | SectionDefInfo | ColDefInfo | PageNumInfo | HeaderFooterInfo | FieldBeginControlInfo | FieldEndControlInfo;

export interface TableControlInfo {
  type: 'table';
  instanceId: number;
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
  // From CTRL_HEADER attrs (offset 4)
  zOrder: number;
  textWrap: number;   // 0=TOP_AND_BOTTOM, 1=SQUARE, 2=TIGHT, 3=THROUGH, 4=NONE
  textFlow: number;   // 0=BOTH_SIDES, 1=LEFT_ONLY, 2=RIGHT_ONLY, 3=LARGER
  lock: boolean;
  treatAsChar: boolean;
  affectLSpacing: boolean;
  flowWithText: boolean;
  allowOverlap: boolean;
  holdAnchorAndSO: boolean;
  vertRelTo: number;    // 0=PAPER, 1=PAGE, 2=PARA, 3=LINE
  vertAlignPos: number; // 0=TOP, 1=CENTER, 2=BOTTOM
  horzRelTo: number;    // 0=PAPER, 1=PAGE, 2=COLUMN, 3=PARA
  horzAlignPos: number; // 0=LEFT, 1=CENTER, 2=RIGHT
  xOffset: number;
  yOffset: number;
  // From HWPTAG_TABLE attrs (offset 0)
  repeatHeader: boolean;
  noAdjust: boolean;
  pageBreakType: number; // 0=CELL, 1=PAGE, 2=COLUMN
  captionParas?: ParaInfo[];
  captionGap?: number;
  captionDir?: number;
  captionWidth?: number;
  captionLastWidth?: number;
  captionFullSz?: boolean;
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
  headerCell: boolean;
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
  attrs: number;
  spaceColumns: number;
  lineGrid: number;
  charGrid: number;
  tabStop: number;
  outlineShapeIDRef: number;
  pageNum: number;
  picNum: number;
  tblNum: number;
  eqNum: number;
  lineNumRestartType: number;
  lineNumCountBy: number;
  lineNumDistance: number;
  lineNumStartNumber: number;
  footnote: FootnoteShapeInfo;
  endnote: FootnoteShapeInfo;
  pageBorderFills: PageBorderFillInfo[];
}

// ── FileHeader type ──

export interface DocHeader {
  version: { major: number; minor: number; patch: number; revision: number };
  flags: number;
}
