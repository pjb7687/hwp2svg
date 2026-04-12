/**
 * HWP DocInfo stream parsing: fonts, char shapes, para shapes, border fills, bin data.
 */

import { parseRecords, readWString, readLPWString, dataView } from './record.js';
import * as TAG from './constants.js';
import type {
  FontFaceInfo, CharShapeInfo, ParaShapeInfo, BorderFillInfo,
  BinDataItemInfo, DocInfoData,
} from './hwp-types.js';

export const BORDER_WIDTHS = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];

export function parseFontFace(data: Uint8Array): FontFaceInfo {
  const nameLen = dataView(data).getUint16(1, true);
  const name = readWString(data, 3, nameLen);
  return { name };
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
  const textColor = data.length >= 56 ? dv.getUint32(52, true) : 0;

  return {
    fontIds,
    height,
    textColor,
    bold: (attrs & 0x02) !== 0,
    italic: (attrs & 0x01) !== 0,
    underlineType: (attrs >> 2) & 0x03,
    strikeout: (attrs >> 18) & 0x07,
    superscript: (attrs & (1 << 15)) !== 0,
    subscript: (attrs & (1 << 16)) !== 0,
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
  const borderFillId = data.length >= 34 ? dv.getUint16(32, true) : 0;
  const alignment = (attrs1 >> 2) & 0x07;

  if (data.length >= 54) {
    const lineSpacing2 = dv.getUint32(50, true);
    if (lineSpacing2 > 0) lineSpacing = lineSpacing2;
  }

  return { attrs1, leftMargin, rightMargin, indent, spacingBefore, spacingAfter, lineSpacing, alignment, tabDefId, borderFillId };
}

export function parseBorderFill(data: Uint8Array, id: number): BorderFillInfo {
  const dv = dataView(data);
  const readBorder = (off: number) => ({
    type: data.length > off ? data[off] : 0,
    width: data.length > off + 1 ? data[off + 1] : 0,
    color: data.length >= off + 6 ? dv.getUint32(off + 2, true) : 0,
  });

  const left = readBorder(2);
  const right = readBorder(8);
  const top = readBorder(14);
  const bottom = readBorder(20);

  const info: BorderFillInfo = {
    id,
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
    fillColor: null,
  };

  if (data.length >= 36) {
    const fillOffset = 32;
    const fillType = dv.getUint32(fillOffset, true);
    if ((fillType & 1) !== 0 && fillOffset + 8 <= data.length) {
      info.fillColor = dv.getUint32(fillOffset + 4, true);
    }
  }

  return info;
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
    fonts: [],
    charShapes: [],
    paraShapes: [],
    borderFills: [],
    binDataItems: [],
  };

  for (const rec of records) {
    switch (rec.tagId) {
      case TAG.HWPTAG_DOCUMENT_PROPERTIES:
        info.sectionCount = dataView(rec.data).getUint16(0, true);
        break;
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
      case TAG.HWPTAG_BIN_DATA:
        info.binDataItems.push(parseBinDataItem(rec.data));
        break;
    }
  }

  return info;
}
