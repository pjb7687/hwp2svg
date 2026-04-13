/**
 * HWP section stream parsing: paragraphs, text runs, line segments, tables.
 */

import { parseRecords, dataView, type HwpRecord } from './record.js';
import * as TAG from './constants.js';
import type {
  PageDefInfo, TextRunInfo, LineSegInfo, ParaInfo, ControlInfo,
  TableControlInfo, CellInfo, SectionDefInfo, FootnoteShapeInfo, PageBorderFillInfo,
  ColDefInfo, PageNumInfo, HeaderFooterInfo, FieldBeginControlInfo, FieldEndControlInfo,
} from './hwp-types.js';

/**
 * Map HWP Private Use Area (Wingdings) characters to standard Unicode.
 */
const PUA_TO_UNICODE: Record<number, number> = {
  0xF0A1: 0x2702, 0xF0A2: 0x2701, 0xF0A3: 0x2703, 0xF0A4: 0x2756,
  0xF0A7: 0x25C6, 0xF0A8: 0x25A1, 0xF0A9: 0x25CB, 0xF0AA: 0x2714,
  0xF0AB: 0x2718, 0xF0AC: 0x2721, 0xF0AD: 0x2606, 0xF0AE: 0x2605,
  0xF0B2: 0x25CF, 0xF0B7: 0x2022, 0xF0B9: 0x25A0, 0xF0D8: 0x25B2,
  0xF0DA: 0x25BA, 0xF0DB: 0x25BC, 0xF0DD: 0x25C4, 0xF0E0: 0x2709,
  0xF0E1: 0x270E, 0xF0E8: 0x2660, 0xF0E9: 0x2663, 0xF0EA: 0x2665,
  0xF0EB: 0x2666, 0xF0EF: 0x263A, 0xF0F0: 0x2639, 0xF0F1: 0x263C,
  0xF0F2: 0x2640, 0xF0F3: 0x2642, 0xF0F4: 0x2660, 0xF0F5: 0x2663,
  0xF0FC: 0x2713, 0xF020: 0x0020, 0xF021: 0x270F, 0xF022: 0x2702,
  0xF06C: 0x25CF, 0xF06D: 0x2B24, 0xF06E: 0x25A0, 0xF06F: 0x25B2,
  0xF070: 0x25C6, 0xF071: 0x2B25, 0xF072: 0x25CF, 0xF073: 0x2022,
  0xF074: 0x25AA, 0xF075: 0x25AB, 0xF076: 0x25A1, 0xF0FE: 0x2611,
  0xF0FD: 0x2612,
};

export function mapPuaToUnicode(ch: number): number {
  if (ch >= 0xE000 && ch <= 0xF8FF) {
    return PUA_TO_UNICODE[ch] ?? ch;
  }
  return ch;
}

export function mapPuaStringToUnicode(s: string): string {
  let result = '';
  for (let i = 0; i < s.length; i++) {
    result += String.fromCharCode(mapPuaToUnicode(s.charCodeAt(i)));
  }
  return result;
}

export function parsePageDef(data: Uint8Array): PageDefInfo {
  const dv = dataView(data);
  const attrs = data.length >= 40 ? dv.getUint32(36, true) : 0;
  return {
    width: dv.getUint32(0, true),
    height: dv.getUint32(4, true),
    marginLeft: dv.getUint32(8, true),
    marginRight: dv.getUint32(12, true),
    marginTop: dv.getUint32(16, true),
    marginBottom: dv.getUint32(20, true),
    marginHeader: dv.getUint32(24, true),
    marginFooter: dv.getUint32(28, true),
    marginGutter: dv.getUint32(32, true),
    landscape: attrs & 1,
    gutterType: (attrs >> 1) & 3,
  };
}

function parseFootnoteShape(data: Uint8Array): FootnoteShapeInfo {
  const dv = dataView(data);
  const attrs = data.length >= 4 ? dv.getUint32(0, true) : 0;
  return {
    numberType: attrs & 0xFF,
    placement: (attrs >> 8) & 3,
    numbering: (attrs >> 10) & 3,
    supscript: ((attrs >> 12) & 1) !== 0,
    beneathText: ((attrs >> 13) & 1) !== 0,
    userChar: data.length >= 6 ? dv.getUint16(4, true) : 0,
    prefixChar: data.length >= 8 ? dv.getUint16(6, true) : 0,
    suffixChar: data.length >= 10 ? dv.getUint16(8, true) : 0,
    startNumber: data.length >= 12 ? dv.getUint16(10, true) : 1,
    // noteLineLength is stored as HWPUNIT (4 bytes), interpreted as signed INT32
    noteLineLength: data.length >= 16 ? dv.getInt32(12, true) : 0,
    noteLineTop: data.length >= 18 ? dv.getUint16(16, true) : 0,
    noteLineBottom: data.length >= 20 ? dv.getUint16(18, true) : 0,
    noteSpacing: data.length >= 22 ? dv.getUint16(20, true) : 0,
    lineType: data.length >= 23 ? data[22] : 0,
    lineWidth: data.length >= 24 ? data[23] : 0,
    lineColor: data.length >= 28 ? dv.getUint32(24, true) : 0,
  };
}

function parsePageBorderFill(data: Uint8Array): PageBorderFillInfo {
  const dv = dataView(data);
  const attrs = data.length >= 4 ? dv.getUint32(0, true) : 0;
  return {
    textBorder: attrs & 1,
    headerInside: ((attrs >> 1) & 1) !== 0,
    footerInside: ((attrs >> 2) & 1) !== 0,
    fillArea: (attrs >> 3) & 3,
    leftGap: data.length >= 6 ? dv.getUint16(4, true) : 0,
    rightGap: data.length >= 8 ? dv.getUint16(6, true) : 0,
    topGap: data.length >= 10 ? dv.getUint16(8, true) : 0,
    bottomGap: data.length >= 12 ? dv.getUint16(10, true) : 0,
    borderFillId: data.length >= 14 ? dv.getUint16(12, true) : 0,
  };
}

export function parseParaText(data: Uint8Array): {
  runs: TextRunInfo[];
  totalCharCount: number;
  // Extended ctrl positions: positions of 16-byte ctrl chars that have corresponding CTRL_HEADER records.
  // Used to map ctrlCharIdx → stream position in parseParagraph.
  extendedCtrlPositions: number[];
  // All ctrl positions (including inline field-end): used for WCHAR position calculation in applyCharShapes.
  allCtrlPositions: number[];
  // Inline ctrl positions (field end): ctrl chars without CTRL_HEADER records.
  inlineCtrlPositions: number[];
  textBeforeCtrl: boolean;  // true if any text/tab/lf char appeared before the first 16-byte ctrl char
} {
  const runs: TextRunInfo[] = [];
  const extendedCtrlPositions: number[] = [];
  const allCtrlPositions: number[] = [];
  const inlineCtrlPositions: number[] = [];
  let currentText = '';
  let pos = 0;
  let charCount = 0;
  let firstCtrlPos = -1;
  let hasTextBeforeCtrl = false;
  const dv = dataView(data);

  while (pos + 2 <= data.length) {
    const ch = dv.getUint16(pos, true);

    if (ch === TAG.CTRL_PARA_BREAK) {
      if (currentText) {
        runs.push({ charPrId: 0, text: currentText });
        currentText = '';
      }
      charCount++;
      pos += 2;
      break;
    }

    if (ch < 32) {
      if (currentText) {
        runs.push({ charPrId: 0, text: currentText });
        currentText = '';
      }

      charCount++;
      if (ch === TAG.CTRL_TAB) {
        if (firstCtrlPos === -1) hasTextBeforeCtrl = true;
        runs.push({ charPrId: 0, text: '\t' });
        pos += 2;
      } else if (ch === TAG.CTRL_LINE_BREAK) {
        if (firstCtrlPos === -1) hasTextBeforeCtrl = true;
        runs.push({ charPrId: 0, text: '\n' });
        pos += 2;
      } else if (
        ch === TAG.CTRL_SECTION_COLUMN_DEF ||
        ch === TAG.CTRL_FIELD_BEGIN ||
        ch === TAG.CTRL_DRAWING_TABLE ||
        ch === TAG.CTRL_HIDDEN_COMMENT ||
        ch === TAG.CTRL_HEADER_FOOTER ||
        ch === TAG.CTRL_FOOTNOTE_ENDNOTE ||
        ch === TAG.CTRL_AUTO_NUMBER ||
        ch === TAG.CTRL_PAGE_CTRL ||
        ch === TAG.CTRL_BOOKMARK ||
        ch === TAG.CTRL_DUTMAL
      ) {
        // Extended ctrl: has a CTRL_HEADER record
        extendedCtrlPositions.push(charCount - 1);
        allCtrlPositions.push(charCount - 1);
        if (firstCtrlPos === -1) firstCtrlPos = charCount - 1;
        pos += 16;
      } else if (ch === TAG.CTRL_FIELD_END) {
        // Inline ctrl: no CTRL_HEADER record; tracked separately for pairing
        inlineCtrlPositions.push(charCount - 1);
        allCtrlPositions.push(charCount - 1);
        if (firstCtrlPos === -1) firstCtrlPos = charCount - 1;
        pos += 16;
      } else {
        pos += 2;
      }
    } else {
      charCount++;
      if (firstCtrlPos === -1) hasTextBeforeCtrl = true;
      currentText += String.fromCharCode(ch);
      pos += 2;
    }
  }

  if (currentText) {
    runs.push({ charPrId: 0, text: currentText });
  }

  return { runs, totalCharCount: charCount, extendedCtrlPositions, allCtrlPositions, inlineCtrlPositions, textBeforeCtrl: hasTextBeforeCtrl };
}

function getCharPrIdAt(pos: number, data: Uint8Array, count: number): number {
  const dv = dataView(data);
  let id = 0;
  for (let i = 0; i < count && i * 8 + 8 <= data.length; i++) {
    const shapePos = dv.getUint32(i * 8, true);
    if (shapePos <= pos) {
      id = dv.getUint32(i * 8 + 4, true);
    } else {
      break;
    }
  }
  return id;
}

export function applyCharShapes(runs: TextRunInfo[], data: Uint8Array, count: number, ctrlPositions?: number[]): TextRunInfo[] {
  const dv = dataView(data);
  const shapes: { pos: number; id: number }[] = [];
  for (let i = 0; i < count && i * 8 + 8 <= data.length; i++) {
    shapes.push({
      pos: dv.getUint32(i * 8, true),
      id: dv.getUint32(i * 8 + 4, true),
    });
  }

  if (shapes.length === 0 || runs.length === 0) return runs;

  // Build mapping: textIndex → WCHAR position (PARA_CHAR_SHAPE units).
  // Each text char occupies 1 WCHAR (2 bytes); each 16-byte ctrl char occupies 8 WCHARs.
  // ctrlPositions holds logical stream positions (each ctrl = 1 logical position).
  // Converting to WCHAR position: each ctrl adds 7 extra WCHARs (8 - 1 = 7).
  const sortedCtrlPositions = ctrlPositions ? [...ctrlPositions].sort((a, b) => a - b) : [];
  let ctrlIdx = 0;
  let logicalPos = 0;  // logical stream position
  let wcharPos = 0;    // WCHAR position used by PARA_CHAR_SHAPE

  const totalTextChars = runs.reduce((sum, r) => sum + r.text.length, 0);
  const textToStream = new Int32Array(totalTextChars);
  for (let ti = 0; ti < totalTextChars; ti++) {
    // Skip ctrl chars at current logical position; each ctrl occupies 8 WCHARs
    while (ctrlIdx < sortedCtrlPositions.length && sortedCtrlPositions[ctrlIdx] === logicalPos) {
      ctrlIdx++;
      logicalPos++;
      wcharPos += 8;
    }
    textToStream[ti] = wcharPos;
    logicalPos++;
    wcharPos++;
  }

  const result: TextRunInfo[] = [];
  let textIdx = 0;

  for (const run of runs) {
    const runEnd = textIdx + run.text.length;
    let segStart = 0;
    let currentShapeId = getShapeIdAt(textToStream[textIdx], shapes);

    for (let ci = 0; ci < run.text.length; ci++) {
      const thisSp = textToStream[textIdx + ci];
      const nextSp = textIdx + ci + 1 < totalTextChars ? textToStream[textIdx + ci + 1] : -1;
      const newShapeId = ci + 1 < run.text.length ? getShapeIdAt(nextSp, shapes) : -1;

      if (newShapeId !== currentShapeId || ci === run.text.length - 1) {
        const segText = run.text.substring(segStart, ci + 1);
        if (segText) {
          result.push({ charPrId: currentShapeId, text: segText });
        }
        segStart = ci + 1;
        currentShapeId = newShapeId;
      }
    }

    textIdx = runEnd;
  }

  return result;
}

function getShapeIdAt(pos: number, shapes: { pos: number; id: number }[]): number {
  let id = shapes[0].id;
  for (const s of shapes) {
    if (s.pos <= pos) id = s.id;
    else break;
  }
  return id;
}

export function parseLineSegs(data: Uint8Array, count: number): LineSegInfo[] {
  const dv = dataView(data);
  const segs: LineSegInfo[] = [];
  for (let i = 0; i < count && i * 36 + 36 <= data.length; i++) {
    const off = i * 36;
    segs.push({
      textPos: dv.getUint32(off, true),
      vertPos: dv.getInt32(off + 4, true),
      vertSize: dv.getInt32(off + 8, true),
      textHeight: dv.getInt32(off + 12, true),
      baseline: dv.getInt32(off + 16, true),
      spacing: dv.getInt32(off + 20, true),
      horzPos: dv.getInt32(off + 24, true),
      horzSize: dv.getInt32(off + 28, true),
      flags: dv.getUint32(off + 32, true),
    });
  }
  return segs;
}

export function parseParagraph(
  records: HwpRecord[],
  startIdx: number,
): [ParaInfo, number] {
  const headerRec = records[startIdx];
  const dv = dataView(headerRec.data);

  const nchars = dv.getUint32(0, true) & 0x7FFFFFFF;
  const paraShapeId = dv.getUint16(8, true);
  const styleId = headerRec.data[10];
  const charShapeCount = dv.getUint16(12, true);
  const lineSegCount = dv.getUint16(16, true);

  const breakTypeByte = headerRec.data.length > 11 ? headerRec.data[11] : 0;
  const pageBreak = (breakTypeByte & 0x04) !== 0;
  const columnBreak = (breakTypeByte & 0x08) !== 0;
  const paraId = headerRec.data.length >= 22 ? dv.getUint32(18, true) : 0;
  const merged = headerRec.data.length >= 24 ? dv.getUint16(22, true) : 0;

  const headerLevel = headerRec.level;
  const para: ParaInfo = {
    paraPrId: paraShapeId,
    styleId,
    paraId,
    pageBreak,
    columnBreak,
    merged,
    defaultCharPrId: 0,
    paraBreakCharPrId: 0,
    runs: [],
    lineSegs: [],
    controls: [],
    ctrlCharPrIds: [],
    ctrlStreamPositions: [],
    textBeforeCtrl: false,
  };

  let i = startIdx + 1;
  let totalCharCount = 1;  // default: at least 1 for para break
  let extendedCtrlPositions: number[] = [];
  let allCtrlPositions: number[] = [];
  let inlineCtrlPositions: number[] = [];
  let ctrlCharShapeData: Uint8Array | null = null;
  let ctrlCharIdx = 0;

  // Converts a logical stream position to the WCHAR position used by PARA_CHAR_SHAPE.
  // Each 16-byte ctrl char occupies 8 WCHARs (16/2=8), so each ctrl adds 7 extra WCHARs.
  function toWcharPos(logicalPos: number): number {
    let extras = 0;
    for (const p of allCtrlPositions) {
      if (p < logicalPos) extras++;
    }
    return logicalPos + extras * 7;
  }

  while (i < records.length && records[i].level > headerLevel) {
    const rec = records[i];

    switch (rec.tagId) {
      case TAG.HWPTAG_PARA_TEXT: {
        const { runs: rawRuns, totalCharCount: tc, extendedCtrlPositions: ecp, allCtrlPositions: acp, inlineCtrlPositions: icp, textBeforeCtrl: tbc } = parseParaText(rec.data);
        para.runs = rawRuns;
        totalCharCount = tc;
        extendedCtrlPositions = ecp;
        allCtrlPositions = acp;
        inlineCtrlPositions = icp;
        para.textBeforeCtrl = tbc;
        break;
      }

      case TAG.HWPTAG_PARA_CHAR_SHAPE:
        if (charShapeCount > 0 && rec.data.length >= 8) {
          const dv0 = dataView(rec.data);
          para.defaultCharPrId = dv0.getUint32(4, true);
          // Para-break charPrId uses WCHAR position of the last char
          const paraBreakWcharPos = toWcharPos(totalCharCount - 1);
          para.paraBreakCharPrId = getCharPrIdAt(paraBreakWcharPos, rec.data, charShapeCount);
          ctrlCharShapeData = rec.data;
        }
        para.runs = applyCharShapes(para.runs, rec.data, charShapeCount, allCtrlPositions);
        break;

      case TAG.HWPTAG_PARA_LINE_SEG:
        para.lineSegs = parseLineSegs(rec.data, lineSegCount);
        break;

      case TAG.HWPTAG_CTRL_HEADER: {
        const ctrlIdBuf = rec.data.subarray(0, 4);
        const ctrlId = String.fromCharCode(ctrlIdBuf[3], ctrlIdBuf[2], ctrlIdBuf[1], ctrlIdBuf[0]);
        const [ctrl, nextI] = parseControl(records, i, ctrlId);
        if (ctrl) {
          // Look up the charPrId for this ctrl char's position using WCHAR position
          const logicalPos = ctrlCharIdx < extendedCtrlPositions.length ? extendedCtrlPositions[ctrlCharIdx] : 0;
          const wcharPos = toWcharPos(logicalPos);
          const ctrlCharPrId = ctrlCharShapeData
            ? getCharPrIdAt(wcharPos, ctrlCharShapeData, charShapeCount)
            : para.defaultCharPrId;
          para.controls.push(ctrl);
          para.ctrlCharPrIds.push(ctrlCharPrId);
          para.ctrlStreamPositions.push(logicalPos);
        }
        ctrlCharIdx++;
        i = nextI;
        continue;
      }
    }

    i++;
  }

  // After all CTRL_HEADERs are processed, create FieldEndControlInfo for any inline
  // ctrl chars (field end) that don't have their own CTRL_HEADER record.
  if (inlineCtrlPositions.length > 0 && ctrlCharShapeData !== null) {
    // Find open field begins that need pairing with field ends.
    // Match them in order: first open fieldBegin pairs with first fieldEnd.
    const openFieldBegins: FieldBeginControlInfo[] = [];
    for (const ctrl of para.controls) {
      if (ctrl.type === 'fieldBegin') openFieldBegins.push(ctrl);
    }

    for (let fi = 0; fi < inlineCtrlPositions.length; fi++) {
      const logicalPos = inlineCtrlPositions[fi];
      const wcharPos = toWcharPos(logicalPos);
      const charPrId = getCharPrIdAt(wcharPos, ctrlCharShapeData, charShapeCount);
      const matchingBegin = openFieldBegins[fi];
      const fieldEnd: FieldEndControlInfo = {
        type: 'fieldEnd',
        ctrlId: matchingBegin?.ctrlId ?? '',
        beginId: matchingBegin?.id ?? 0,
      };
      para.controls.push(fieldEnd);
      para.ctrlCharPrIds.push(charPrId);
      para.ctrlStreamPositions.push(logicalPos);
    }
  }

  // suppress unused warning for nchars
  void nchars;

  return [para, i];
}

export function parseControl(
  records: HwpRecord[],
  ctrlHeaderIdx: number,
  ctrlId: string,
): [ControlInfo | null, number] {
  const ctrlLevel = records[ctrlHeaderIdx].level;
  let i = ctrlHeaderIdx + 1;

  if (ctrlId === 'tbl ') {
    const ctrlData = records[ctrlHeaderIdx].data;
    const ctrlDv = dataView(ctrlData);
    // GSOCommonProperties (개체 공통 속성) layout:
    //   offset  0: ctrlId UINT32
    //   offset  4: attrs UINT32 (bit fields)
    //   offset  8: yOffset INT32
    //   offset 12: xOffset INT32
    //   offset 16: width UINT32
    //   offset 20: height UINT32
    //   offset 24: z-order INT32
    //   offset 28-35: outer margins UINT16×4 (left, right, top, bottom)
    const ctrlAttrs = ctrlData.length >= 8 ? ctrlDv.getUint32(4, true) : 0;
    const treatAsChar = (ctrlAttrs & 1) !== 0;
    const affectLSpacing = ((ctrlAttrs >> 2) & 1) !== 0;
    const vertRelTo = (ctrlAttrs >> 3) & 0x3;
    const vertAlignPos = (ctrlAttrs >> 5) & 0x7;
    const horzRelTo = (ctrlAttrs >> 8) & 0x3;
    const horzAlignPos = (ctrlAttrs >> 10) & 0x7;
    const flowWithText = ((ctrlAttrs >> 13) & 1) !== 0;
    const allowOverlap = ((ctrlAttrs >> 14) & 1) !== 0;
    const textWrap = (ctrlAttrs >> 21) & 0x7;
    const textFlow = (ctrlAttrs >> 24) & 0x3;
    const yOffset = ctrlData.length >= 12 ? ctrlDv.getInt32(8, true) : 0;
    const xOffset = ctrlData.length >= 16 ? ctrlDv.getInt32(12, true) : 0;
    const ctrlWidth = ctrlData.length >= 20 ? ctrlDv.getUint32(16, true) : 0;
    const ctrlHeight = ctrlData.length >= 24 ? ctrlDv.getUint32(20, true) : 0;
    const zOrder = ctrlData.length >= 28 ? ctrlDv.getInt32(24, true) : 0;
    const outMarginLeft = ctrlData.length >= 30 ? ctrlDv.getUint16(28, true) : 0;
    const outMarginRight = ctrlData.length >= 32 ? ctrlDv.getUint16(30, true) : 0;
    const outMarginTop = ctrlData.length >= 34 ? ctrlDv.getUint16(32, true) : 0;
    const outMarginBottom = ctrlData.length >= 36 ? ctrlDv.getUint16(34, true) : 0;
    // Instance ID at offset 36 (4 bytes) in GSOCommonProperties
    const instanceId = ctrlData.length >= 40 ? ctrlDv.getUint32(36, true) : 0;

    // Parse caption LIST_H (and its paragraphs) that appear before HWPTAG_TABLE
    const captionParas: ReturnType<typeof parseParagraph>[0][] = [];
    let captionGap = 0;
    let captionDir = 3; // default: bottom
    let captionWidth = 0;
    let captionLastWidth = 0;
    let captionFullSz = false;
    while (i < records.length && records[i].level > ctrlLevel && records[i].tagId !== TAG.HWPTAG_TABLE) {
      if (records[i].tagId === TAG.HWPTAG_LIST_HEADER) {
        const listLevel = records[i].level;
        const lh = records[i];
        if (lh.data.length >= 22) {
          const ldv = dataView(lh.data);
          // Caption list header layout:
          //   offset 0-1: para count INT16
          //   offset 2-5: listAttrs UINT32
          //   offset 6-7: additional list flags UINT16
          //   offset 8-11: caption attrs UINT32 (bits 0-1=dir, bit 2=fullSz)
          //   offset 12-15: caption width HWPUNIT
          //   offset 16-17: gap HWPUNIT16
          //   offset 18-21: last width HWPUNIT
          const captionAttrs = ldv.getUint32(8, true);
          captionDir = captionAttrs & 3;
          captionFullSz = ((captionAttrs >> 2) & 1) !== 0;
          captionWidth = ldv.getUint32(12, true);
          captionGap = ldv.getUint16(16, true);
          captionLastWidth = ldv.getUint32(18, true);
        }
        i++;
        while (i < records.length && records[i].level >= listLevel && records[i].tagId !== TAG.HWPTAG_TABLE) {
          if (records[i].tagId === TAG.HWPTAG_PARA_HEADER) {
            const [para, nextI] = parseParagraph(records, i);
            captionParas.push(para);
            i = nextI;
          } else {
            i++;
          }
        }
      } else {
        i++;
      }
    }
    if (i < records.length && records[i].tagId === TAG.HWPTAG_TABLE) {
      const [table, nextI] = parseTableControl(records, i);
      table.instanceId = instanceId;
      table.ctrlWidth = ctrlWidth;
      table.ctrlHeight = ctrlHeight;
      table.zOrder = zOrder;
      table.treatAsChar = treatAsChar;
      table.affectLSpacing = affectLSpacing;
      table.vertRelTo = vertRelTo;
      table.vertAlignPos = vertAlignPos;
      table.horzRelTo = horzRelTo;
      table.horzAlignPos = horzAlignPos;
      table.flowWithText = flowWithText;
      table.allowOverlap = allowOverlap;
      table.textWrap = textWrap;
      table.textFlow = textFlow;
      table.xOffset = xOffset;
      table.yOffset = yOffset;
      table.outMarginLeft = outMarginLeft;
      table.outMarginRight = outMarginRight;
      table.outMarginTop = outMarginTop;
      table.outMarginBottom = outMarginBottom;
      if (captionParas.length > 0) {
        table.captionParas = captionParas;
        table.captionGap = captionGap;
        table.captionDir = captionDir;
        table.captionWidth = captionWidth;
        table.captionLastWidth = captionLastWidth;
        table.captionFullSz = captionFullSz;
      }
      return [table, nextI];
    }
  } else if (ctrlId === 'secd') {
    // Parse section definition from CTRL_HEADER body (starting at offset 4 after ctrlId)
    const secdData = records[ctrlHeaderIdx].data;
    const secdDv = dataView(secdData);
    const secdAttrs = secdData.length >= 8 ? secdDv.getUint32(4, true) : 0;
    const spaceColumns = secdData.length >= 10 ? secdDv.getUint16(8, true) : 0;
    const lineGrid = secdData.length >= 12 ? secdDv.getUint16(10, true) : 0;
    const charGrid = secdData.length >= 14 ? secdDv.getUint16(12, true) : 0;
    const tabStop = secdData.length >= 18 ? secdDv.getUint32(14, true) : 8000;
    const outlineShapeIDRef = secdData.length >= 20 ? secdDv.getUint16(18, true) : 0;
    const pageNum = secdData.length >= 22 ? secdDv.getUint16(20, true) : 0;
    const picNum = secdData.length >= 24 ? secdDv.getUint16(22, true) : 0;
    const tblNum = secdData.length >= 26 ? secdDv.getUint16(24, true) : 0;
    const eqNum = secdData.length >= 28 ? secdDv.getUint16(26, true) : 0;
    // Line number shape data at offset 30-37
    const lineNumRestartType = secdData.length >= 32 ? secdDv.getUint16(30, true) : 0;
    const lineNumCountBy = secdData.length >= 34 ? secdDv.getUint16(32, true) : 0;
    const lineNumDistance = secdData.length >= 36 ? secdDv.getUint16(34, true) : 0;
    const lineNumStartNumber = secdData.length >= 38 ? secdDv.getUint16(36, true) : 0;

    let pageDef: PageDefInfo | null = null;
    const footnoteShapes: FootnoteShapeInfo[] = [];
    const pageBorderFills: PageBorderFillInfo[] = [];

    while (i < records.length && records[i].level > ctrlLevel) {
      if (records[i].tagId === TAG.HWPTAG_PAGE_DEF) {
        pageDef = parsePageDef(records[i].data);
      } else if (records[i].tagId === TAG.HWPTAG_FOOTNOTE_SHAPE) {
        footnoteShapes.push(parseFootnoteShape(records[i].data));
      } else if (records[i].tagId === TAG.HWPTAG_PAGE_BORDER_FILL) {
        pageBorderFills.push(parsePageBorderFill(records[i].data));
      }
      i++;
    }
    if (pageDef) {
      const defaultNote: FootnoteShapeInfo = {
        numberType: 0, placement: 0, numbering: 0, supscript: false, beneathText: false,
        userChar: 0, prefixChar: 0, suffixChar: 0x29, startNumber: 1,
        noteLineLength: -1, noteLineTop: 0, noteLineBottom: 0, noteSpacing: 0,
        lineType: 0, lineWidth: 0, lineColor: 0,
      };
      return [{
        type: 'secd', pageDef,
        attrs: secdAttrs, spaceColumns, lineGrid, charGrid, tabStop,
        outlineShapeIDRef, pageNum, picNum, tblNum, eqNum,
        lineNumRestartType, lineNumCountBy, lineNumDistance, lineNumStartNumber,
        footnote: footnoteShapes[0] ?? defaultNote,
        endnote: footnoteShapes[1] ?? defaultNote,
        pageBorderFills,
      }, i];
    }
    return [null, i];
  }

  if (ctrlId === 'cold') {
    const coldData = records[ctrlHeaderIdx].data;
    const coldDv = dataView(coldData);
    const attrsLow = coldData.length >= 6 ? coldDv.getUint16(4, true) : 0;
    const sameGap = coldData.length >= 8 ? coldDv.getUint16(6, true) : 0;
    const cold: ColDefInfo = {
      type: 'cold',
      colType: attrsLow & 3,
      colCount: (attrsLow >> 2) & 0xFF,
      layout: (attrsLow >> 10) & 3,
      sameSz: ((attrsLow >> 12) & 1) !== 0,
      sameGap,
    };
    while (i < records.length && records[i].level > ctrlLevel) i++;
    return [cold, i];
  }

  if (ctrlId === 'pgnp') {
    const pgnpData = records[ctrlHeaderIdx].data;
    const pgnpDv = dataView(pgnpData);
    const attrs = pgnpData.length >= 8 ? pgnpDv.getUint32(4, true) : 0;
    const sideChar = pgnpData.length >= 16 ? pgnpDv.getUint16(14, true) : 0x2D;
    const pgnp: PageNumInfo = {
      type: 'pgnp',
      formatType: attrs & 0xFF,
      pos: (attrs >> 8) & 0xF,
      sideChar,
    };
    while (i < records.length && records[i].level > ctrlLevel) i++;
    return [pgnp, i];
  }

  if (ctrlId === 'head' || ctrlId === 'foot') {
    const hfData = records[ctrlHeaderIdx].data;
    const hfDv = dataView(hfData);
    const applyPageType = hfData.length >= 8 ? hfDv.getUint32(4, true) & 3 : 0;
    const id = hfData.length >= 12 ? hfDv.getUint32(8, true) : 0;
    let textWidth = 0;
    let textHeight = 0;
    const paragraphs: ReturnType<typeof parseParagraph>[0][] = [];

    while (i < records.length && records[i].level > ctrlLevel) {
      if (records[i].tagId === TAG.HWPTAG_LIST_HEADER) {
        const lhData = records[i].data;
        const lhDv = dataView(lhData);
        if (lhData.length >= 16) {
          textWidth = lhDv.getUint32(8, true);
          textHeight = lhDv.getUint32(12, true);
        }
        i++;
        // Parse paragraphs inside this list (paras are at the same level as the LIST_HEADER)
        const listLevel = records[i - 1].level;
        while (i < records.length && records[i].level >= listLevel && records[i].level > ctrlLevel) {
          if (records[i].tagId === TAG.HWPTAG_PARA_HEADER) {
            const [para, nextI] = parseParagraph(records, i);
            paragraphs.push(para);
            i = nextI;
          } else {
            i++;
          }
        }
      } else {
        i++;
      }
    }

    const hf: HeaderFooterInfo = {
      type: ctrlId as 'head' | 'foot',
      id,
      applyPageType,
      textWidth,
      textHeight,
      paragraphs,
    };
    return [hf, i];
  }

  // Field begin controls (e.g. '%hlk' hyperlink):
  // CTRL_HEADER data layout:
  //   offset 0-3: ctrlId UINT32
  //   offset 4-7: attrs UINT32 (bit 0=editable, bit 15=dirty)
  //   offset 8: extra attrs BYTE
  //   offset 9-10: command length len WORD
  //   offset 11..(11+2*len-1): command WCHAR array[len]
  //   offset (11+2*len)..(11+2*len+3): id UINT32 (unique document id)
  if (ctrlId.startsWith('%')) {
    const fData = records[ctrlHeaderIdx].data;
    const fDv = dataView(fData);
    const fAttrs = fData.length >= 8 ? fDv.getUint32(4, true) : 0;
    const editable = (fAttrs & 1) !== 0;
    const dirty = ((fAttrs >> 15) & 1) !== 0;
    let command = '';
    let fieldId = 0;
    if (fData.length >= 11) {
      const len = fDv.getUint16(9, true);
      const commandEnd = 11 + len * 2;
      for (let ci = 0; ci < len && 11 + ci * 2 + 2 <= fData.length; ci++) {
        command += String.fromCharCode(fDv.getUint16(11 + ci * 2, true));
      }
      fieldId = fData.length >= commandEnd + 4 ? fDv.getUint32(commandEnd, true) : 0;
    }
    const fieldBegin: FieldBeginControlInfo = {
      type: 'fieldBegin',
      ctrlId,
      id: fieldId,
      command,
      editable,
      dirty,
    };
    // Field begin has no meaningful child records in the section stream
    while (i < records.length && records[i].level > ctrlLevel) i++;
    return [fieldBegin, i];
  }

  // Skip all sub-records for unknown/unhandled controls
  while (i < records.length && records[i].level > ctrlLevel) {
    i++;
  }
  return [null, i];
}

export function parseTableControl(
  records: HwpRecord[],
  tableIdx: number,
): [TableControlInfo, number] {
  const rec = records[tableIdx];
  const dv = dataView(rec.data);

  // HWPTAG_TABLE data layout:
  //   offset  0: table attrs UINT32 (bit 0-1=pageBreak, bit 2=repeatHeader)
  //   offset  4: rowCount UINT16
  //   offset  6: colCount UINT16
  //   offset  8: cellSpacing UINT16
  //   offset 10-17: inner margins UINT16×4 (left, right, top, bottom)
  //   offset 18+: rowSizes UINT16×rowCount
  //   after rowSizes: borderFillId UINT16
  const tableAttrs = rec.data.length >= 4 ? dv.getUint32(0, true) : 0;
  const pageBreakType = tableAttrs & 0x3;
  const repeatHeader = ((tableAttrs >> 2) & 1) !== 0;
  const noAdjust = ((tableAttrs >> 3) & 1) !== 0;

  const rowCount = dv.getUint16(4, true);
  const colCount = dv.getUint16(6, true);
  const cellSpacing = dv.getUint16(8, true);

  const innerMarginLeft = rec.data.length >= 12 ? dv.getUint16(10, true) : 0;
  const innerMarginRight = rec.data.length >= 14 ? dv.getUint16(12, true) : 0;
  const innerMarginTop = rec.data.length >= 16 ? dv.getUint16(14, true) : 0;
  const innerMarginBottom = rec.data.length >= 18 ? dv.getUint16(16, true) : 0;

  const rowSizes: number[] = [];
  for (let ri = 0; ri < rowCount; ri++) {
    const off = 18 + ri * 2;
    if (off + 2 <= rec.data.length) {
      rowSizes.push(dv.getUint16(off, true));
    } else {
      rowSizes.push(0);
    }
  }
  const rowSizesEnd = 18 + rowCount * 2;
  const borderFillId = rowSizesEnd + 2 <= rec.data.length ? dv.getUint16(rowSizesEnd, true) : 0;

  const table: TableControlInfo = {
    type: 'table',
    instanceId: 0,
    rowCount,
    colCount,
    cellSpacing,
    borderFillId,
    rowSizes,
    innerMarginLeft,
    innerMarginRight,
    innerMarginTop,
    innerMarginBottom,
    cells: [],
    ctrlWidth: 0,
    ctrlHeight: 0,
    zOrder: 0,
    textWrap: 0,
    textFlow: 0,
    lock: false,
    treatAsChar: false,
    affectLSpacing: false,
    flowWithText: false,
    allowOverlap: false,
    holdAnchorAndSO: false,
    vertRelTo: 0,
    vertAlignPos: 0,
    horzRelTo: 0,
    horzAlignPos: 0,
    xOffset: 0,
    yOffset: 0,
    outMarginLeft: 0,
    outMarginRight: 0,
    outMarginTop: 0,
    outMarginBottom: 0,
    repeatHeader,
    noAdjust,
    pageBreakType,
  };

  const tableLevel = rec.level;
  let i = tableIdx + 1;

  while (i < records.length && records[i].level >= tableLevel) {
    if (records[i].tagId === TAG.HWPTAG_LIST_HEADER && records[i].level === tableLevel) {
      const [cell, nextI] = parseTableCell(records, i);
      table.cells.push(cell);
      i = nextI;
    } else {
      i++;
    }
  }

  return [table, i];
}

export function parseTableCell(
  records: HwpRecord[],
  listHeaderIdx: number,
): [CellInfo, number] {
  const listHeaderLevel = records[listHeaderIdx].level;
  const listData = records[listHeaderIdx].data;
  const dv = dataView(listData);

  const listAttrs = listData.length >= 6 ? dv.getUint32(2, true) : 0;
  // listAttrs bits 3-4 and 5-6 are within the third byte of the UINT32 (byte offset 4 from record start)
  // which corresponds to bits 19-20 (lineWrap) and 21-22 (vertAlign) of the full UINT32
  const lineWrap = (listAttrs >> 19) & 0x3;  // 0=Break, 1=Squeeze, 2=Keep
  const vertAlign = (listAttrs >> 21) & 0x3;

  const cellFlags = listData.length >= 8 ? dv.getUint16(6, true) : 0;
  const hasMargin = (cellFlags & 0x1) === 1;
  const headerCell = (cellFlags & 0x4) !== 0;

  const cell: CellInfo = {
    colAddr: 0,
    rowAddr: 0,
    colSpan: 1,
    rowSpan: 1,
    width: 0,
    height: 0,
    borderFillId: 0,
    hasMargin,
    headerCell,
    marginLeft: 141,
    marginRight: 141,
    marginTop: 141,
    marginBottom: 141,
    vertAlign,
    lineWrap,
    paragraphs: [],
  };

  if (listData.length >= 34) {
    cell.colAddr = dv.getUint16(8, true);
    cell.rowAddr = dv.getUint16(10, true);
    cell.colSpan = dv.getUint16(12, true);
    cell.rowSpan = dv.getUint16(14, true);
    cell.width = dv.getUint32(16, true);
    cell.height = dv.getUint32(20, true);
    cell.marginLeft = dv.getUint16(24, true);
    cell.marginRight = dv.getUint16(26, true);
    cell.marginTop = dv.getUint16(28, true);
    cell.marginBottom = dv.getUint16(30, true);
    cell.borderFillId = dv.getUint16(32, true);
  }

  let i = listHeaderIdx + 1;
  while (i < records.length) {
    const rec = records[i];
    if (rec.level < listHeaderLevel) break;
    if (rec.level === listHeaderLevel && rec.tagId === TAG.HWPTAG_LIST_HEADER) break;

    if (rec.tagId === TAG.HWPTAG_PARA_HEADER && rec.level === listHeaderLevel) {
      const [para, nextI] = parseParagraph(records, i);
      cell.paragraphs.push(para);
      i = nextI;
    } else {
      i++;
    }
  }

  return [cell, i];
}

export function parseSectionRecords(buf: Uint8Array): { pageDef: PageDefInfo; paragraphs: ParaInfo[] } {
  const records = parseRecords(buf);

  const defaultPageDef: PageDefInfo = {
    width: 59528, height: 84188, landscape: 0, gutterType: 0,
    marginLeft: 5669, marginRight: 5669,
    marginTop: 2834, marginBottom: 2834,
    marginHeader: 4251, marginFooter: 4251,
    marginGutter: 0,
  };

  let pageDef = defaultPageDef;
  const paragraphs: ParaInfo[] = [];

  let i = 0;
  while (i < records.length) {
    const rec = records[i];
    if (rec.tagId === TAG.HWPTAG_PARA_HEADER && rec.level === 0) {
      const [para, nextIdx] = parseParagraph(records, i);
      for (const ctrl of para.controls) {
        if (ctrl.type === 'secd') {
          pageDef = ctrl.pageDef;
        }
      }
      paragraphs.push(para);
      i = nextIdx;
    } else {
      i++;
    }
  }

  return { pageDef, paragraphs };
}
