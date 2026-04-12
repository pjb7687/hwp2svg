/**
 * XML generation: header.xml, section XML, manifest, version.xml.
 */

import type { DocInfoData, PageDefInfo, ParaInfo, TextRunInfo, TableControlInfo, CellInfo, DocHeader } from './hwp-types.js';
import { BORDER_WIDTHS } from './hwp-docinfo.js';

// ── XML Namespaces ──

const NS_HH = 'http://www.hancom.co.kr/hwpml/2011/head';
const NS_HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph';
const NS_HS = 'http://www.hancom.co.kr/hwpml/2011/section';
const NS_HC = 'http://www.hancom.co.kr/hwpml/2011/core';
const NS_OPF = 'http://www.idpf.org/2007/opf/';

// ── Utility ──

function colorrefToHex(colorref: number): string {
  if (((colorref >>> 24) & 0xFF) === 0xFF) {
    return '#ffffff';
  }
  const r = colorref & 0xFF;
  const g = (colorref >> 8) & 0xFF;
  const b = (colorref >> 16) & 0xFF;
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ── Alignment map ──

const ALIGN_MAP: Record<number, string> = {
  0: 'JUSTIFY',
  1: 'LEFT',
  2: 'RIGHT',
  3: 'CENTER',
  4: 'DISTRIBUTE',
  5: 'DISTRIBUTE_SPACE',
};

// ── DocInfo → header.xml ──

export function generateHeaderXml(info: DocInfoData): string {
  const lines: string[] = [];
  lines.push(`<?xml version="1.0" encoding="UTF-8"?>`);
  lines.push(`<hh:head xmlns:hh="${NS_HH}">`);

  // fontfaces
  lines.push(`  <hh:fontfaces>`);
  for (let i = 0; i < info.fonts.length; i++) {
    const f = info.fonts[i];
    lines.push(`    <hh:fontface id="${i}" name="${escapeXml(f.name)}"/>`);
  }
  lines.push(`  </hh:fontfaces>`);

  // charProperties
  const LANGS = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'];
  lines.push(`  <hh:charProperties>`);
  for (let i = 0; i < info.charShapes.length; i++) {
    const cs = info.charShapes[i];
    const attrs: string[] = [
      `id="${i}"`,
      `height="${cs.height}"`,
      `textColor="${colorrefToHex(cs.textColor)}"`,
      `bold="${cs.bold}"`,
      `italic="${cs.italic}"`,
      `underline="${cs.underlineType}"`,
      `strikeout="${cs.strikeout}"`,
      `superscript="${cs.superscript}"`,
      `subscript="${cs.subscript}"`,
    ];
    const fontRefAttrs = LANGS.map((l, j) => `${l}="${cs.fontIds[j] ?? 0}"`).join(' ');
    const ratioAttrs = LANGS.map((l, j) => `${l}="${cs.ratio[j] ?? 100}"`).join(' ');
    const spacingAttrs = LANGS.map((l, j) => `${l}="${cs.spacing[j] ?? 0}"`).join(' ');

    lines.push(`    <hh:charPr ${attrs.join(' ')}>`);
    lines.push(`      <hh:fontRef ${fontRefAttrs}/>`);
    lines.push(`      <hh:ratio ${ratioAttrs}/>`);
    lines.push(`      <hh:spacing ${spacingAttrs}/>`);
    lines.push(`    </hh:charPr>`);
  }
  lines.push(`  </hh:charProperties>`);

  // paraProperties
  lines.push(`  <hh:paraProperties>`);
  for (let i = 0; i < info.paraShapes.length; i++) {
    const ps = info.paraShapes[i];
    const align = ALIGN_MAP[ps.alignment] ?? 'JUSTIFY';
    const attrs: string[] = [
      `id="${i}"`,
      `align="${align}"`,
      `leftMargin="${ps.leftMargin}"`,
      `rightMargin="${ps.rightMargin}"`,
      `indent="${ps.indent}"`,
      `spacingBefore="${ps.spacingBefore}"`,
      `spacingAfter="${ps.spacingAfter}"`,
      `lineSpacing="${ps.lineSpacing}"`,
      `borderFillId="${ps.borderFillId}"`,
    ];
    lines.push(`    <hh:paraPr ${attrs.join(' ')}/>`);
  }
  lines.push(`  </hh:paraProperties>`);

  // borderFills
  lines.push(`  <hh:borderFills>`);
  for (const bf of info.borderFills) {
    const bw = (n: number) => `${BORDER_WIDTHS[n] ?? 0.1}mm`;
    const bt = (type: number, _width: number) => {
      if (type === 0) return 'NONE';
      if (type === 1) return 'SOLID';
      if (type === 2) return 'DASHED';
      if (type === 3) return 'DOTTED';
      if (type === 4) return 'DASH_DOT';
      if (type === 5) return 'DASH_DOT_DOT';
      if (type === 6) return 'LONG_DASH';
      if (type === 7) return 'LARGE_DOT';
      if (type === 8) return 'DOUBLE';
      return 'SOLID';
    };
    const attrs: string[] = [
      `id="${bf.id}"`,
      `leftBorderType="${bt(bf.leftBorderType, bf.leftBorderWidth)}"`,
      `leftBorderWidth="${bw(bf.leftBorderWidth)}"`,
      `leftBorderColor="${colorrefToHex(bf.leftBorderColor)}"`,
      `rightBorderType="${bt(bf.rightBorderType, bf.rightBorderWidth)}"`,
      `rightBorderWidth="${bw(bf.rightBorderWidth)}"`,
      `rightBorderColor="${colorrefToHex(bf.rightBorderColor)}"`,
      `topBorderType="${bt(bf.topBorderType, bf.topBorderWidth)}"`,
      `topBorderWidth="${bw(bf.topBorderWidth)}"`,
      `topBorderColor="${colorrefToHex(bf.topBorderColor)}"`,
      `bottomBorderType="${bt(bf.bottomBorderType, bf.bottomBorderWidth)}"`,
      `bottomBorderWidth="${bw(bf.bottomBorderWidth)}"`,
      `bottomBorderColor="${colorrefToHex(bf.bottomBorderColor)}"`,
    ];
    if (bf.fillColor !== null) {
      attrs.push(`fillColor="${colorrefToHex(bf.fillColor)}"`);
    }
    lines.push(`    <hh:borderFill ${attrs.join(' ')}/>`);
  }
  lines.push(`  </hh:borderFills>`);

  lines.push(`</hh:head>`);
  return lines.join('\n');
}

// ── Section XML ──

function generateRunXml(run: TextRunInfo, indent: string): string {
  const text = run.text.replace(/\t/g, ' ');
  if (!text || text === '\n') {
    if (run.text === '\n') return `${indent}<hp:lineBreak/>`;
    return '';
  }
  return `${indent}<hp:run charPrIDRef="${run.charPrId}"><hp:t>${escapeXml(text)}</hp:t></hp:run>`;
}

function generateParaXml(para: ParaInfo, indent: string): string {
  const lines: string[] = [];
  const ind2 = indent + '  ';
  const ind3 = ind2 + '  ';

  lines.push(`${indent}<hp:p paraPrIDRef="${para.paraPrId}" styleIDRef="${para.styleId}">`);

  if (para.lineSegs.length > 0) {
    lines.push(`${ind2}<hp:linesegarray>`);
    for (const seg of para.lineSegs) {
      lines.push(
        `${ind3}<hp:lineseg textpos="${seg.textPos}" vertpos="${seg.vertPos}" vertsize="${seg.vertSize}" ` +
        `textheight="${seg.textHeight}" baseline="${seg.baseline}" spacing="${seg.spacing}" ` +
        `horzpos="${seg.horzPos}" horzsize="${seg.horzSize}" flags="${seg.flags}"/>`
      );
    }
    lines.push(`${ind2}</hp:linesegarray>`);
  }

  for (const run of para.runs) {
    const runXml = generateRunXml(run, ind2);
    if (runXml) lines.push(runXml);
  }

  for (const ctrl of para.controls) {
    if (ctrl.type === 'table') {
      lines.push(...generateTableXml(ctrl, ind2));
    }
  }

  lines.push(`${indent}</hp:p>`);
  return lines.join('\n');
}

function generateTableXml(table: TableControlInfo, indent: string): string[] {
  const lines: string[] = [];
  const ind2 = indent + '  ';
  const ind3 = ind2 + '  ';
  const ind4 = ind3 + '  ';

  lines.push(
    `${indent}<hp:ctrl><hp:tbl rowCnt="${table.rowCount}" colCnt="${table.colCount}" ` +
    `cellSpacing="${table.cellSpacing}" borderFillIDRef="${table.borderFillId}" ` +
    `width="${table.ctrlWidth}" height="${table.ctrlHeight}" ` +
    `innerMarginLeft="${table.innerMarginLeft}" innerMarginRight="${table.innerMarginRight}" ` +
    `innerMarginTop="${table.innerMarginTop}" innerMarginBottom="${table.innerMarginBottom}" ` +
    `outMarginLeft="${table.outMarginLeft}" outMarginRight="${table.outMarginRight}" ` +
    `outMarginTop="${table.outMarginTop}" outMarginBottom="${table.outMarginBottom}">`
  );

  const rowMap = new Map<number, CellInfo[]>();
  for (const cell of table.cells) {
    const arr = rowMap.get(cell.rowAddr) ?? [];
    arr.push(cell);
    rowMap.set(cell.rowAddr, arr);
  }
  const sortedRows = [...rowMap.keys()].sort((a, b) => a - b);

  for (const rowAddr of sortedRows) {
    const cells = (rowMap.get(rowAddr) ?? []).sort((a, b) => a.colAddr - b.colAddr);
    const rowSize = table.rowSizes[rowAddr] ?? 0;
    lines.push(`${ind2}<hp:tr${rowSize > 0 ? ` rowSize="${rowSize}"` : ''}>`);
    for (const cell of cells) {
      const vertAlignStr = cell.vertAlign === 1 ? 'CENTER' : cell.vertAlign === 2 ? 'BOTTOM' : 'TOP';
      const mLeft = cell.hasMargin ? cell.marginLeft : table.innerMarginLeft;
      const mRight = cell.hasMargin ? cell.marginRight : table.innerMarginRight;
      const mTop = cell.hasMargin ? cell.marginTop : table.innerMarginTop;
      const mBottom = cell.hasMargin ? cell.marginBottom : table.innerMarginBottom;
      lines.push(
        `${ind3}<hp:tc colAddr="${cell.colAddr}" rowAddr="${cell.rowAddr}" ` +
        `colSpan="${cell.colSpan}" rowSpan="${cell.rowSpan}" ` +
        `width="${cell.width}" height="${cell.height}" ` +
        `borderFillIDRef="${cell.borderFillId}" ` +
        `marginLeft="${mLeft}" marginRight="${mRight}" ` +
        `marginTop="${mTop}" marginBottom="${mBottom}">`
      );
      const ind5 = ind4 + '  ';
      lines.push(`${ind4}<hp:subList vertAlign="${vertAlignStr}">`);
      for (const para of cell.paragraphs) {
        lines.push(generateParaXml(para, ind5));
      }
      lines.push(`${ind4}</hp:subList>`);
      lines.push(`${ind3}</hp:tc>`);
    }
    lines.push(`${ind2}</hp:tr>`);
  }

  lines.push(`${ind2}</hp:tbl></hp:ctrl>`);
  return lines;
}

export function generateSectionXml(
  sectionIndex: number,
  pageDef: PageDefInfo,
  paragraphs: ParaInfo[],
): string {
  const lines: string[] = [];
  lines.push(`<?xml version="1.0" encoding="UTF-8"?>`);
  lines.push(
    `<hs:sec xmlns:hs="${NS_HS}" xmlns:hp="${NS_HP}" ` +
    `id="${sectionIndex}">`
  );

  lines.push(
    `  <hs:pageDef width="${pageDef.width}" height="${pageDef.height}" ` +
    `landscape="${pageDef.landscape}" ` +
    `marginLeft="${pageDef.marginLeft}" marginRight="${pageDef.marginRight}" ` +
    `marginTop="${pageDef.marginTop}" marginBottom="${pageDef.marginBottom}" ` +
    `marginHeader="${pageDef.marginHeader}" marginFooter="${pageDef.marginFooter}" ` +
    `marginGutter="${pageDef.marginGutter}"/>`
  );

  for (const para of paragraphs) {
    lines.push(generateParaXml(para, '  '));

    // Emit caption paragraphs for any table control in this paragraph
    for (const ctrl of para.controls) {
      if (ctrl.type === 'table' && ctrl.captionParas && ctrl.captionParas.length > 0) {
        // Compute table paragraph's absolute vertpos (from its first lineseg)
        const tableParaVertpos = para.lineSegs.length > 0 ? para.lineSegs[0].vertPos : 0;
        const captionOffset = tableParaVertpos + ctrl.ctrlHeight + (ctrl.captionGap ?? 0);
        for (const capPara of ctrl.captionParas) {
          // Deep-copy lineSegs with adjusted vertPos
          const adjustedSegs = capPara.lineSegs.map(seg => ({
            ...seg,
            vertPos: captionOffset + seg.vertPos,
          }));
          const adjustedPara = { ...capPara, lineSegs: adjustedSegs };
          lines.push(generateParaXml(adjustedPara, '  '));
        }
      }
    }
  }

  lines.push(`</hs:sec>`);
  return lines.join('\n');
}

// ── Manifest / content.hpf ──

export function generateContentHpf(sectionCount: number): string {
  const lines: string[] = [];
  lines.push(`<?xml version="1.0" encoding="UTF-8"?>`);
  lines.push(`<opf:package xmlns:opf="${NS_OPF}" version="2.0">`);
  lines.push(`  <opf:manifest>`);
  lines.push(`    <opf:item id="header" href="header.xml" media-type="application/xml"/>`);
  for (let i = 0; i < sectionCount; i++) {
    lines.push(`    <opf:item id="section${i}" href="section${i}.xml" media-type="application/xml"/>`);
  }
  lines.push(`  </opf:manifest>`);
  lines.push(`  <opf:spine>`);
  for (let i = 0; i < sectionCount; i++) {
    lines.push(`    <opf:itemref idref="section${i}"/>`);
  }
  lines.push(`  </opf:spine>`);
  lines.push(`</opf:package>`);
  return lines.join('\n');
}

export function generateVersionXml(header: DocHeader): string {
  const { major, minor, patch, revision } = header.version;
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<hc:coreProperties xmlns:hc="${NS_HC}">\n` +
    `  <hc:version>${major}.${minor}.${patch}.${revision}</hc:version>\n` +
    `</hc:coreProperties>`
  );
}
