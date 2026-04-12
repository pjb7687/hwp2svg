/**
 * HWP section stream parsing: paragraphs, text runs, line segments, tables.
 */

import { parseRecords, dataView, type HwpRecord } from './record.js';
import * as TAG from './constants.js';
import type {
  PageDefInfo, TextRunInfo, LineSegInfo, ParaInfo, ControlInfo,
  TableControlInfo, CellInfo, SectionDefInfo,
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

function mapPuaToUnicode(ch: number): number {
  if (ch >= 0xE000 && ch <= 0xF8FF) {
    return PUA_TO_UNICODE[ch] ?? ch;
  }
  return ch;
}

export function parsePageDef(data: Uint8Array): PageDefInfo {
  const dv = dataView(data);
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
    landscape: (dv.getUint32(36, true) & 1) !== 0,
  };
}

export function parseParaText(data: Uint8Array): TextRunInfo[] {
  const runs: TextRunInfo[] = [];
  let currentText = '';
  let pos = 0;
  const dv = dataView(data);

  while (pos + 2 <= data.length) {
    const ch = dv.getUint16(pos, true);

    if (ch === TAG.CTRL_PARA_BREAK) {
      if (currentText) {
        runs.push({ charPrId: 0, text: currentText });
        currentText = '';
      }
      pos += 2;
      break;
    }

    if (ch < 32) {
      if (currentText) {
        runs.push({ charPrId: 0, text: currentText });
        currentText = '';
      }

      if (ch === TAG.CTRL_TAB) {
        runs.push({ charPrId: 0, text: '\t' });
        pos += 2;
      } else if (ch === TAG.CTRL_LINE_BREAK) {
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
        pos += 16;
      } else if (ch === TAG.CTRL_FIELD_END) {
        pos += 16;
      } else {
        pos += 2;
      }
    } else {
      const mapped = mapPuaToUnicode(ch);
      currentText += String.fromCharCode(mapped);
      pos += 2;
    }
  }

  if (currentText) {
    runs.push({ charPrId: 0, text: currentText });
  }

  return runs;
}

export function applyCharShapes(runs: TextRunInfo[], data: Uint8Array, count: number): TextRunInfo[] {
  const dv = dataView(data);
  const shapes: { pos: number; id: number }[] = [];
  for (let i = 0; i < count && i * 8 + 8 <= data.length; i++) {
    shapes.push({
      pos: dv.getUint32(i * 8, true),
      id: dv.getUint32(i * 8 + 4, true),
    });
  }

  if (shapes.length === 0 || runs.length === 0) return runs;

  const result: TextRunInfo[] = [];
  let charPos = 0;

  for (const run of runs) {
    const runEnd = charPos + run.text.length;

    let segStart = 0;
    let currentShapeId = getShapeIdAt(charPos, shapes);

    for (let ci = 1; ci <= run.text.length; ci++) {
      const absPos = charPos + ci;
      const newShapeId = ci < run.text.length ? getShapeIdAt(absPos, shapes) : -1;

      if (newShapeId !== currentShapeId || ci === run.text.length) {
        const segText = run.text.substring(segStart, ci);
        if (segText) {
          result.push({ charPrId: currentShapeId, text: segText });
        }
        segStart = ci;
        currentShapeId = newShapeId;
      }
    }

    charPos = runEnd;
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

  const headerLevel = headerRec.level;
  const para: ParaInfo = {
    paraPrId: paraShapeId,
    styleId,
    runs: [],
    lineSegs: [],
    controls: [],
  };

  let i = startIdx + 1;

  while (i < records.length && records[i].level > headerLevel) {
    const rec = records[i];

    switch (rec.tagId) {
      case TAG.HWPTAG_PARA_TEXT:
        para.runs = parseParaText(rec.data);
        break;

      case TAG.HWPTAG_PARA_CHAR_SHAPE:
        para.runs = applyCharShapes(para.runs, rec.data, charShapeCount);
        break;

      case TAG.HWPTAG_PARA_LINE_SEG:
        para.lineSegs = parseLineSegs(rec.data, lineSegCount);
        break;

      case TAG.HWPTAG_CTRL_HEADER: {
        const ctrlIdBuf = rec.data.subarray(0, 4);
        const ctrlId = String.fromCharCode(ctrlIdBuf[3], ctrlIdBuf[2], ctrlIdBuf[1], ctrlIdBuf[0]);
        const [ctrl, nextI] = parseControl(records, i, ctrlId);
        if (ctrl) para.controls.push(ctrl);
        i = nextI;
        continue;
      }
    }

    i++;
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
    const ctrlWidth = ctrlData.length >= 20 ? ctrlDv.getUint32(16, true) : 0;
    const ctrlHeight = ctrlData.length >= 24 ? ctrlDv.getUint32(20, true) : 0;
    const outMarginLeft = ctrlData.length >= 30 ? ctrlDv.getUint16(28, true) : 0;
    const outMarginRight = ctrlData.length >= 32 ? ctrlDv.getUint16(30, true) : 0;
    const outMarginTop = ctrlData.length >= 34 ? ctrlDv.getUint16(32, true) : 0;
    const outMarginBottom = ctrlData.length >= 36 ? ctrlDv.getUint16(34, true) : 0;

    // Parse caption LIST_H (and its paragraphs) that appear before HWPTAG_TABLE
    const captionParas: ReturnType<typeof parseParagraph>[0][] = [];
    let captionGap = 0;
    let captionDir = 3; // default: bottom
    while (i < records.length && records[i].level > ctrlLevel && records[i].tagId !== TAG.HWPTAG_TABLE) {
      if (records[i].tagId === TAG.HWPTAG_LIST_HEADER) {
        const listLevel = records[i].level;
        const lh = records[i];
        if (lh.data.length >= 10) {
          const ldv = dataView(lh.data);
          captionDir = ldv.getUint32(0, true) & 3;
          captionGap = ldv.getUint16(8, true);
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
      table.ctrlWidth = ctrlWidth;
      table.ctrlHeight = ctrlHeight;
      table.outMarginLeft = outMarginLeft;
      table.outMarginRight = outMarginRight;
      table.outMarginTop = outMarginTop;
      table.outMarginBottom = outMarginBottom;
      if (captionParas.length > 0) {
        table.captionParas = captionParas;
        table.captionGap = captionGap;
        table.captionDir = captionDir;
      }
      return [table, nextI];
    }
  } else if (ctrlId === 'secd') {
    let pageDef: PageDefInfo | null = null;
    while (i < records.length && records[i].level > ctrlLevel) {
      if (records[i].tagId === TAG.HWPTAG_PAGE_DEF) {
        pageDef = parsePageDef(records[i].data);
      }
      i++;
    }
    if (pageDef) {
      return [{ type: 'secd', pageDef }, i];
    }
    return [null, i];
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
    outMarginLeft: 0,
    outMarginRight: 0,
    outMarginTop: 0,
    outMarginBottom: 0,
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
  const lineWrap = (listAttrs >> 3) & 0x3;  // 0=Break, 1=Squeeze, 2=Keep
  const vertAlign = (listAttrs >> 5) & 0x3;

  const cellFlags = listData.length >= 8 ? dv.getUint16(6, true) : 0;
  const hasMargin = (cellFlags & 0x1) === 1;

  const cell: CellInfo = {
    colAddr: 0,
    rowAddr: 0,
    colSpan: 1,
    rowSpan: 1,
    width: 0,
    height: 0,
    borderFillId: 0,
    hasMargin,
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
    width: 59528, height: 84188, landscape: false,
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
