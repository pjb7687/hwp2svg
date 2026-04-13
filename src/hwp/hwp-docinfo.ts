/**
 * HWP DocInfo stream parsing: fonts, char shapes, para shapes, border fills, bin data.
 */

import { parseRecords, readWString, readLPWString, dataView } from './record.js';
import * as TAG from './constants.js';
import type {
  FontFaceInfo, FontTypeInfo, SubstFontInfo,
  CharShapeInfo, ParaShapeInfo, BorderFillInfo,
  TabDefInfo, TabItemInfo, StyleInfo, BulletInfo,
  BinDataItemInfo, DocInfoData,
} from './hwp-types.js';

export const BORDER_WIDTHS = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];

export function parseFontFace(data: Uint8Array): FontFaceInfo {
  const dv = dataView(data);
  const attrs = data[0];
  const fontType = attrs & 0x07;           // 0=unknown, 1=TTF, 2=HFT
  const hasSubstFont = (attrs & 0x80) !== 0;
  const hasTypeInfo = (attrs & 0x40) !== 0;
  const hasBaseFont = (attrs & 0x20) !== 0;

  const nameLen = dv.getUint16(1, true);
  const name = readWString(data, 3, nameLen);
  let offset = 3 + nameLen * 2;

  const info: FontFaceInfo = { name, fontType };

  if (hasSubstFont && offset < data.length) {
    const substType = data[offset++];
    const substLen = data.length > offset + 1 ? dv.getUint16(offset, true) : 0;
    offset += 2;
    const substName = readWString(data, offset, substLen);
    offset += substLen * 2;
    info.substFont = { type: substType, name: substName };
  }

  if (hasTypeInfo && offset + 10 <= data.length) {
    info.typeInfo = {
      familyType: data[offset],
      serifType: data[offset + 1],
      weight: data[offset + 2],
      proportion: data[offset + 3],
      contrast: data[offset + 4],
      strokeVariation: data[offset + 5],
      armStyle: data[offset + 6],
      letterform: data[offset + 7],
      midline: data[offset + 8],
      xHeight: data[offset + 9],
    };
    offset += 10;
  }

  if (hasBaseFont && offset < data.length) {
    const baseLen = dv.getUint16(offset, true);
    offset += 2;
    info.baseFontName = readWString(data, offset, baseLen);
  }

  return info;
}

export function parseCharShape(data: Uint8Array): CharShapeInfo {
  const dv = dataView(data);
  const fontIds: number[] = [];
  for (let i = 0; i < 7; i++) fontIds.push(dv.getUint16(i * 2, true));

  const ratio: number[] = [];
  for (let i = 0; i < 7; i++) ratio.push(data[14 + i]);

  const spacing: number[] = [];
  for (let i = 0; i < 7; i++) spacing.push(dv.getInt8(21 + i));

  const relSize: number[] = [];
  for (let i = 0; i < 7; i++) relSize.push(data[28 + i]);

  const offsetArr: number[] = [];
  for (let i = 0; i < 7; i++) offsetArr.push(dv.getInt8(35 + i));

  const height = dv.getInt32(42, true);
  const attrs = dv.getUint32(46, true);

  const shadowX = data.length > 50 ? dv.getInt8(50) : 10;
  const shadowY = data.length > 51 ? dv.getInt8(51) : 10;
  const textColor = data.length >= 56 ? dv.getUint32(52, true) : 0;
  const underlineColor = data.length >= 60 ? dv.getUint32(56, true) : 0;
  const shadeColor = data.length >= 64 ? dv.getUint32(60, true) : 0xFFFFFFFF;
  const shadowColor = data.length >= 68 ? dv.getUint32(64, true) : 0x00C0C0C0;
  const borderFillId = data.length >= 70 ? dv.getUint16(68, true) : 0;
  const strikeoutColor = data.length >= 74 ? dv.getUint32(70, true) : 0;

  return {
    fontIds,
    height,
    textColor,
    bold: (attrs & 0x02) !== 0,
    italic: (attrs & 0x01) !== 0,
    underlineType: (attrs >> 2) & 0x03,
    underlineShape: (attrs >> 4) & 0x0F,
    underlineColor,
    shadeColor,
    shadowColor,
    shadowX,
    shadowY,
    outlineType: (attrs >> 8) & 0x07,
    shadowType: (attrs >> 11) & 0x03,
    symMark: (attrs >> 21) & 0x0F,
    useFontSpace: (attrs & (1 << 25)) !== 0,
    useKerning: (attrs & (1 << 30)) !== 0,
    strikeout: (attrs >> 18) & 0x07,
    strikeoutShape: (attrs >> 26) & 0x0F,
    strikeoutColor,
    superscript: (attrs & (1 << 15)) !== 0,
    subscript: (attrs & (1 << 16)) !== 0,
    borderFillId,
    spacing,
    relSize,
    offset: offsetArr,
    ratio,
  };
}

export function parseParaShape(data: Uint8Array): ParaShapeInfo {
  const dv = dataView(data);
  const attrs1 = dv.getUint32(0, true);
  const leftMargin = dv.getInt32(4, true);
  const rightMargin = dv.getInt32(8, true);
  const indent = dv.getInt32(12, true);
  const spacingBefore = dv.getInt32(16, true);
  const spacingAfter = dv.getInt32(20, true);
  let lineSpacing = dv.getInt32(24, true);
  const tabDefId = dv.getUint16(28, true);
  const numberingId = data.length >= 32 ? dv.getUint16(30, true) : 0;
  const borderFillId = data.length >= 34 ? dv.getUint16(32, true) : 0;
  const borderLeft = data.length >= 36 ? dv.getInt16(34, true) : 0;
  const borderRight = data.length >= 38 ? dv.getInt16(36, true) : 0;
  const borderTop = data.length >= 40 ? dv.getInt16(38, true) : 0;
  const borderBottom = data.length >= 42 ? dv.getInt16(40, true) : 0;

  const alignment = (attrs1 >> 2) & 0x07;

  let lineWrap = 0;
  let autoSpacingEng = false;
  let autoSpacingNum = false;
  if (data.length >= 46) {
    const attrs2 = dv.getUint32(42, true);
    lineWrap = attrs2 & 0x3;
    autoSpacingEng = (attrs2 & (1 << 4)) !== 0;
    autoSpacingNum = (attrs2 & (1 << 5)) !== 0;
  }

  let lineSpacingType = 0;
  if (data.length >= 50) {
    const attrs3 = dv.getUint32(46, true);
    lineSpacingType = attrs3 & 0x1F;
  }

  if (data.length >= 54) {
    const lineSpacing2 = dv.getUint32(50, true);
    if (lineSpacing2 > 0) lineSpacing = lineSpacing2;
  }

  // Derive from attrs1
  const snapToGrid = (attrs1 & (1 << 8)) !== 0;
  const condense = (attrs1 >> 9) & 0x7F;
  const widowOrphan = (attrs1 & (1 << 16)) !== 0;
  const keepWithNext = (attrs1 & (1 << 17)) !== 0;
  const keepLines = (attrs1 & (1 << 18)) !== 0;
  const pageBreakBefore = (attrs1 & (1 << 19)) !== 0;
  const vertAlign = (attrs1 >> 20) & 0x03;
  const fontLineHeight = (attrs1 & (1 << 22)) !== 0;
  const headingType = (attrs1 >> 23) & 0x03;
  const headingLevel = (attrs1 >> 25) & 0x07;
  const borderConnect = (attrs1 & (1 << 28)) !== 0;
  const ignoreMargin = (attrs1 & (1 << 29)) !== 0;
  const breakLatinWord = (attrs1 >> 5) & 0x03;
  const breakNonLatinWord = (attrs1 & (1 << 7)) !== 0;

  return {
    attrs1, leftMargin, rightMargin, indent, spacingBefore, spacingAfter, lineSpacing,
    alignment, tabDefId, numberingId, borderFillId,
    borderLeft, borderRight, borderTop, borderBottom,
    lineWrap, autoSpacingEng, autoSpacingNum, lineSpacingType,
    snapToGrid, condense, fontLineHeight, vertAlign,
    breakLatinWord, breakNonLatinWord,
    widowOrphan, keepWithNext, keepLines, pageBreakBefore,
    headingType, headingLevel, borderConnect, ignoreMargin,
  };
}

export function parseBorderFill(data: Uint8Array, id: number): BorderFillInfo {
  const dv = dataView(data);
  const attrs = data.length >= 2 ? dv.getUint16(0, true) : 0;

  // Border layout: interleaved per-border (type 1B, width 1B, color 4B)
  const readBorder = (off: number) => ({
    type: data.length > off ? data[off] : 0,
    width: data.length > off + 1 ? data[off + 1] : 0,
    color: data.length >= off + 6 ? dv.getUint32(off + 2, true) : 0,
  });

  const left = readBorder(2);
  const right = readBorder(8);
  const top = readBorder(14);
  const bottom = readBorder(20);

  const diagonalType = data.length > 26 ? data[26] : 0;
  const diagonalWidth = data.length > 27 ? data[27] : 0;
  const diagonalColor = data.length >= 32 ? dv.getUint32(28, true) : 0;

  const info: BorderFillInfo = {
    id,
    threeD: (attrs & 0x0001) !== 0,
    shadow: (attrs & 0x0002) !== 0,
    slashType: (attrs >> 2) & 0x07,
    backSlashType: (attrs >> 5) & 0x07,
    slashCrooked: (attrs >> 8) & 0x03,
    backSlashCrooked: (attrs & (1 << 10)) !== 0,
    slashCounter: (attrs & (1 << 11)) !== 0,
    backSlashCounter: (attrs & (1 << 12)) !== 0,
    centerLine: (attrs & (1 << 13)) !== 0,
    leftBorderType: left.type,
    rightBorderType: right.type,
    topBorderType: top.type,
    bottomBorderType: bottom.type,
    leftBorderWidth: left.width,
    rightBorderWidth: right.width,
    topBorderWidth: top.width,
    bottomBorderWidth: bottom.width,
    leftBorderColor: left.color,
    rightBorderColor: right.color,
    topBorderColor: top.color,
    bottomBorderColor: bottom.color,
    diagonalType,
    diagonalWidth,
    diagonalColor,
    fillColor: null,
    fillBackColor: null,
    fillPatternType: null,
  };

  if (data.length >= 36) {
    const fillOffset = 32;
    const fillType = dv.getUint32(fillOffset, true);
    if ((fillType & 1) !== 0 && fillOffset + 12 <= data.length) {
      info.fillColor = dv.getUint32(fillOffset + 4, true);
      info.fillBackColor = dv.getUint32(fillOffset + 8, true);
      if (fillOffset + 16 <= data.length) {
        info.fillPatternType = dv.getInt32(fillOffset + 12, true);
      }
    }
  }

  return info;
}

export function parseTabDef(data: Uint8Array): TabDefInfo {
  const dv = dataView(data);
  const attrs = dv.getUint32(0, true);
  const count = data.length >= 6 ? dv.getInt16(4, true) : 0;
  const items: TabItemInfo[] = [];
  for (let i = 0; i < count; i++) {
    const off = 8 + i * 8;
    if (off + 8 > data.length) break;
    items.push({
      pos: dv.getInt32(off, true),
      type: data[off + 4],
      leader: data[off + 5],
    });
  }
  return {
    autoTabLeft: (attrs & 1) !== 0,
    autoTabRight: (attrs & 2) !== 0,
    items,
  };
}

export function parseStyle(data: Uint8Array): StyleInfo {
  const dv = dataView(data);
  const len1 = dv.getUint16(0, true);
  const name = readWString(data, 2, len1);
  let off = 2 + len1 * 2;
  const len2 = dv.getUint16(off, true); off += 2;
  const engName = readWString(data, off, len2); off += len2 * 2;
  const styleAttrs = data[off++];
  const nextStyleId = data[off++];
  const langId = dv.getInt16(off, true); off += 2;
  const paraPrId = dv.getUint16(off, true); off += 2;
  const charPrId = dv.getUint16(off, true);
  return {
    name, engName,
    type: styleAttrs & 0x07,
    nextStyleId,
    langId,
    paraPrId,
    charPrId,
  };
}

export function parseBullet(data: Uint8Array, id: number): BulletInfo {
  const dv = dataView(data);
  // 8-byte paragraph head info (UINT32 attrs, HWPUNIT16 widthAdjust, HWPUNIT16 textOffset)
  const attrs = data.length >= 4 ? dv.getUint32(0, true) : 0;
  const align = attrs & 0x03;               // bits 0-1: 0=LEFT, 1=CENTER, 2=RIGHT
  const useInstWidth = (attrs & 0x04) !== 0; // bit 2
  const autoIndent = (attrs & 0x08) !== 0;   // bit 3
  const textOffsetType = (attrs >> 4) & 0x01; // bit 4: 0=PERCENT, 1=value
  const widthAdjust = data.length >= 6 ? dv.getUint16(4, true) : 0;
  const textOffset = data.length >= 8 ? dv.getUint16(6, true) : 50;
  // charPrIDRef at offset 8, bullet char at offset 12
  const charPrIdRef = data.length >= 12 ? dv.getUint32(8, true) : 0xFFFFFFFF;
  const charCode = data.length >= 14 ? dv.getUint16(12, true) : 0x2022;
  // image bullet flag at offset 14
  const imageFlag = data.length >= 18 ? dv.getInt32(14, true) : 0;
  const useImage = imageFlag !== 0;
  const char = String.fromCharCode(charCode);
  return {
    id, char, useImage, checkable: false,
    level: 0, align, useInstWidth, autoIndent, widthAdjust,
    textOffsetType, textOffset,
    numFormat: 0, charPrIdRef,
  };
}

export function parseBinDataItem(data: Uint8Array): BinDataItemInfo {
  const dv = dataView(data);
  const attrs = dv.getUint16(0, true);
  const type = attrs & 0x0F;
  if (type === 0) {
    let offset = 2;
    const [absPath, absBytes] = readLPWString(data, offset);
    offset += absBytes;
    const [relPath] = readLPWString(data, offset);
    return { type: 'link', absolutePath: absPath, relativePath: relPath };
  }
  if (type === 1) {
    const binDataId = dv.getUint16(2, true);
    const [ext] = readLPWString(data, 4);
    return { type: 'embedding', binDataId, extension: ext };
  }
  return { type: 'storage', binDataId: dv.getUint16(2, true) };
}

export function parseDocInfoData(buf: Uint8Array): DocInfoData {
  const records = parseRecords(buf);
  const info: DocInfoData = {
    sectionCount: 1,
    beginPage: 1,
    beginFootnote: 1,
    beginEndnote: 1,
    beginPic: 1,
    beginTbl: 1,
    beginEquation: 1,
    fontCounts: [0, 0, 0, 0, 0, 0, 0],
    fonts: [],
    charShapes: [],
    paraShapes: [],
    borderFills: [],
    tabDefs: [],
    styles: [],
    bullets: [],
    binDataItems: [],
  };

  let bulletIdCounter = 1;

  for (const rec of records) {
    switch (rec.tagId) {
      case TAG.HWPTAG_DOCUMENT_PROPERTIES: {
        const dv = dataView(rec.data);
        info.sectionCount = dv.getUint16(0, true);
        if (rec.data.length >= 14) {
          info.beginPage = dv.getUint16(2, true);
          info.beginFootnote = dv.getUint16(4, true);
          info.beginEndnote = dv.getUint16(6, true);
          info.beginPic = dv.getUint16(8, true);
          info.beginTbl = dv.getUint16(10, true);
          info.beginEquation = dv.getUint16(12, true);
        }
        break;
      }
      case TAG.HWPTAG_ID_MAPPINGS: {
        const dv = dataView(rec.data);
        // offset 0: binData count, offset 4-31: font counts per language (7 × INT32)
        for (let i = 0; i < 7; i++) {
          if (4 + i * 4 + 4 <= rec.data.length) {
            info.fontCounts[i] = dv.getInt32(4 + i * 4, true);
          }
        }
        break;
      }
      case TAG.HWPTAG_FACE_NAME:
        info.fonts.push(parseFontFace(rec.data));
        break;
      case TAG.HWPTAG_CHAR_SHAPE:
        info.charShapes.push(parseCharShape(rec.data));
        break;
      case TAG.HWPTAG_PARA_SHAPE:
        info.paraShapes.push(parseParaShape(rec.data));
        break;
      case TAG.HWPTAG_BORDER_FILL:
        info.borderFills.push(parseBorderFill(rec.data, info.borderFills.length + 1));
        break;
      case TAG.HWPTAG_TAB_DEF:
        info.tabDefs.push(parseTabDef(rec.data));
        break;
      case TAG.HWPTAG_STYLE:
        info.styles.push(parseStyle(rec.data));
        break;
      case TAG.HWPTAG_BULLET:
        info.bullets.push(parseBullet(rec.data, bulletIdCounter++));
        break;
      case TAG.HWPTAG_BIN_DATA:
        info.binDataItems.push(parseBinDataItem(rec.data));
        break;
    }
  }

  return info;
}
