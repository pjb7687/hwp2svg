/**
 * XML generation: header.xml, section XML, manifest, version.xml.
 */

import type { DocInfoData, PageDefInfo, ParaInfo, TextRunInfo, TableControlInfo, CellInfo, DocHeader, SectionDefInfo, FootnoteShapeInfo, PageBorderFillInfo, ColDefInfo, PageNumInfo, HeaderFooterInfo, FieldBeginControlInfo, FieldEndControlInfo, BorderFillInfo } from './hwp-types.js';
import { BORDER_WIDTHS } from './hwp-docinfo.js';

// ── XML Namespaces ──

const NS_HH = 'http://www.hancom.co.kr/hwpml/2011/head';
const NS_HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph';
const NS_HS = 'http://www.hancom.co.kr/hwpml/2011/section';
const NS_HC = 'http://www.hancom.co.kr/hwpml/2011/core';
const NS_OPF = 'http://www.idpf.org/2007/opf/';

// Standard HWPX namespace block (all files must declare these)
const HWPX_NS =
  `xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" ` +
  `xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" ` +
  `xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" ` +
  `xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" ` +
  `xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" ` +
  `xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" ` +
  `xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" ` +
  `xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" ` +
  `xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
  `xmlns:opf="http://www.idpf.org/2007/opf/" ` +
  `xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" ` +
  `xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" ` +
  `xmlns:epub="http://www.idpf.org/2007/ops" ` +
  `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"`;

// ── Utility ──

function colorrefToHex(colorref: number): string {
  const r = colorref & 0xFF;
  const g = (colorref >> 8) & 0xFF;
  const b = (colorref >> 16) & 0xFF;
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${b.toString(16).padStart(2, '0').toUpperCase()}`;
}

function colorrefToHexOrNone(colorref: number): string {
  if (((colorref >>> 24) & 0xFF) === 0xFF) return 'none';
  return colorrefToHex(colorref);
}

function ctrlIdToMake4CHID(ctrlId: string): number {
  const a = ctrlId.charCodeAt(0) & 0xFF;
  const b = ctrlId.charCodeAt(1) & 0xFF;
  const c = ctrlId.charCodeAt(2) & 0xFF;
  const d = ctrlId.charCodeAt(3) & 0xFF;
  return (((a << 24) | (b << 16) | (c << 8) | d) >>> 0);
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

// ── Header XML helpers ──

const FONT_TYPE_MAP: Record<number, string> = { 0: 'UNKNOWN', 1: 'TTF', 2: 'HFT' };
const FAMILY_TYPE_MAP: Record<number, string> = {
  0: 'FCAT_GOTHIC', 1: 'FCAT_MYUNGJO', 2: 'FCAT_GOTHIC',
  3: 'FCAT_SCRIPT', 4: 'FCAT_OLDSTYLE', 5: 'FCAT_SLAB_SERIF',
  6: 'FCAT_FREEFORM', 7: 'FCAT_SANS_SERIF', 8: 'FCAT_ORNAMENTALS',
};
const LANG_NAMES = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
const LANG_ATTRS = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'];

function borderTypeStr(type: number): string {
  const map: Record<number, string> = {
    0: 'NONE', 1: 'SOLID', 2: 'DOTTED', 3: 'DASH',
    4: 'DASH_DOT', 5: 'DASH_DOT_DOT', 6: 'LONG_DASH', 7: 'LARGE_DOT',
    8: 'DOUBLE_SLIM', 9: 'DOUBLE', 10: 'DOUBLE_THICK',
    11: 'WAVE', 12: 'DOUBLE_WAVE', 13: 'THICK3D', 14: 'THICK3D_INVERT',
    15: 'SLIM3D', 16: 'SLIM3D_INVERT',
  };
  return map[type] ?? 'NONE';
}

function colorrefToHexWithAlpha(colorref: number): string {
  const alpha = (colorref >>> 24) & 0xFF;
  const r = colorref & 0xFF;
  const g = (colorref >> 8) & 0xFF;
  const b = (colorref >> 16) & 0xFF;
  const hex2 = (v: number) => v.toString(16).padStart(2, '0').toUpperCase();
  if (alpha !== 0) {
    return `#${hex2(alpha)}${hex2(r)}${hex2(g)}${hex2(b)}`;
  }
  return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
}

function borderWidthStr(n: number): string {
  return `${BORDER_WIDTHS[n] ?? 0.1} mm`;
}

function generateBorderFillXml(bf: BorderFillInfo, indent: string): string {
  const i2 = indent + '  ';
  const lines: string[] = [];
  const centerLine = bf.centerLine ? 'SOLID' : 'NONE';
  lines.push(`${indent}<hh:borderFill id="${bf.id}" threeD="${bf.threeD ? 1 : 0}" shadow="${bf.shadow ? 1 : 0}" centerLine="${centerLine}" breakCellSeparateLine="0">`);

  // slash/backSlash - derive type string from slashType bits
  const slashTypeStr = bf.slashType === 0 ? 'NONE' : bf.slashType === 2 ? 'SLASH' : 'SLASH';
  const backSlashTypeStr = bf.backSlashType === 0 ? 'NONE' : bf.backSlashType === 2 ? 'BACKSLASH' : 'BACKSLASH';
  lines.push(`${i2}<hh:slash type="${slashTypeStr}" Crooked="${bf.slashCrooked}" isCounter="${bf.slashCounter ? 1 : 0}"/>`);
  lines.push(`${i2}<hh:backSlash type="${backSlashTypeStr}" Crooked="${bf.backSlashCrooked ? 1 : 0}" isCounter="${bf.backSlashCounter ? 1 : 0}"/>`);

  lines.push(`${i2}<hh:leftBorder type="${borderTypeStr(bf.leftBorderType)}" width="${borderWidthStr(bf.leftBorderWidth)}" color="${colorrefToHex(bf.leftBorderColor)}"/>`);
  lines.push(`${i2}<hh:rightBorder type="${borderTypeStr(bf.rightBorderType)}" width="${borderWidthStr(bf.rightBorderWidth)}" color="${colorrefToHex(bf.rightBorderColor)}"/>`);
  lines.push(`${i2}<hh:topBorder type="${borderTypeStr(bf.topBorderType)}" width="${borderWidthStr(bf.topBorderWidth)}" color="${colorrefToHex(bf.topBorderColor)}"/>`);
  lines.push(`${i2}<hh:bottomBorder type="${borderTypeStr(bf.bottomBorderType)}" width="${borderWidthStr(bf.bottomBorderWidth)}" color="${colorrefToHex(bf.bottomBorderColor)}"/>`);
  if (bf.diagonalType !== 0) {
    lines.push(`${i2}<hh:diagonal type="${borderTypeStr(bf.diagonalType)}" width="${borderWidthStr(bf.diagonalWidth)}" color="${colorrefToHex(bf.diagonalColor)}"/>`);
  }

  if (bf.fillColor !== null) {
    const fc = colorrefToHexOrNone(bf.fillColor);
    const bc = bf.fillBackColor !== null ? colorrefToHexWithAlpha(bf.fillBackColor) : 'none';
    lines.push(`${i2}<hc:fillBrush>`);
    lines.push(`${i2}  <hc:winBrush faceColor="${fc}" hatchColor="${bc}" alpha="0"/>`);
    lines.push(`${i2}</hc:fillBrush>`);
  }

  lines.push(`${indent}</hh:borderFill>`);
  return lines.join('\n');
}

const UNDERLINE_TYPE_MAP: Record<number, string> = { 0: 'NONE', 1: 'BOTTOM', 2: 'CENTER', 3: 'TOP' };
const UNDERLINE_SHAPE_MAP: Record<number, string> = {
  0: 'SOLID', 1: 'DASHED', 2: 'DOTTED', 3: 'DASH_DOT',
  4: 'DASH_DOT_DOT', 5: 'LONG_DASH', 6: 'LARGE_DOT', 7: 'DOUBLE',
};
const OUTLINE_TYPE_MAP: Record<number, string> = {
  0: 'NONE', 1: 'SOLID', 2: 'DOTTED', 3: 'THICK',
  4: 'DASH', 5: 'DASH_DOT', 6: 'DASH_DOT_DOT',
};
const SHADOW_TYPE_MAP: Record<number, string> = { 0: 'NONE', 1: 'DISCONTINUOUS', 2: 'CONTINUOUS' };
const SYM_MARK_MAP: Record<number, string> = {
  0: 'NONE', 1: 'DOT_ABOVE', 2: 'RING_ABOVE', 3: 'CARON',
  4: 'TILDE', 5: 'KATAKANA_MIDDLE_DOT', 6: 'COLON',
};
const VERT_ALIGN_MAP: Record<number, string> = { 0: 'BASELINE', 1: 'TOP', 2: 'CENTER', 3: 'BOTTOM' };
const LINE_SPACING_TYPE_MAP: Record<number, string> = {
  0: 'PERCENT', 1: 'FIXED', 2: 'MINIMUM', 3: 'AT_LEAST',
};
const HEADING_TYPE_MAP: Record<number, string> = { 0: 'NONE', 1: 'OUTLINE', 2: 'NUMBERING', 3: 'BULLET' };
const BREAK_LATIN_MAP: Record<number, string> = { 0: 'KEEP_WORD', 1: 'HYPHENATE', 2: 'BREAK_ALL' };
const LINE_WRAP_MAP: Record<number, string> = { 0: 'BREAK', 1: 'SQUEEZE', 2: 'KEEP' };
const TAB_TYPE_MAP: Record<number, string> = { 0: 'LEFT', 1: 'RIGHT', 2: 'CENTER', 3: 'DECIMAL' };
const TAB_LEADER_MAP: Record<number, string> = { 0: 'NONE', 1: 'DOT', 2: 'HYPHEN', 3: 'DASH', 4: 'UNDERLINE', 5: 'EQUAL' };

// ── DocInfo → header.xml ──

export function generateHeaderXml(info: DocInfoData): string {
  const lines: string[] = [];
  lines.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`);
  lines.push(`<hh:head ${HWPX_NS} version="1.4" secCnt="${info.sectionCount}">`);

  // beginNum
  lines.push(`<hh:beginNum page="${info.beginPage}" footnote="${info.beginFootnote}" endnote="${info.beginEndnote}" pic="${info.beginPic}" tbl="${info.beginTbl}" equation="${info.beginEquation}"/>`);

  lines.push(`<hh:refList>`);

  // fontfaces — 7 per-language groups
  const fontCounts = info.fontCounts;
  const totalFontGroups = LANG_NAMES.length;
  lines.push(`<hh:fontfaces itemCnt="${totalFontGroups}">`);
  let fontOffset = 0;
  for (let lang = 0; lang < totalFontGroups; lang++) {
    const cnt = fontCounts[lang] ?? 0;
    lines.push(`<hh:fontface lang="${LANG_NAMES[lang]}" fontCnt="${cnt}">`);
    for (let fi = 0; fi < cnt; fi++) {
      const f = info.fonts[fontOffset + fi];
      if (!f) break;
      const typeStr = FONT_TYPE_MAP[f.fontType] ?? 'TTF';
      lines.push(`<hh:font id="${fi}" face="${escapeXml(f.name)}" type="${typeStr}" isEmbedded="0">`);
      if (f.substFont) {
        const stStr = FONT_TYPE_MAP[f.substFont.type] ?? 'TTF';
        lines.push(`<hh:substFont face="${escapeXml(f.substFont.name)}" type="${stStr}" isEmbedded="0" binaryItemIDRef=""/>`);
      }
      if (f.typeInfo) {
        const ti = f.typeInfo;
        const ft = FAMILY_TYPE_MAP[ti.familyType] ?? 'FCAT_GOTHIC';
        lines.push(`<hh:typeInfo familyType="${ft}" weight="${ti.weight}" proportion="${ti.proportion}" contrast="${ti.contrast}" strokeVariation="${ti.strokeVariation}" armStyle="${ti.armStyle}" letterform="${ti.letterform}" midline="${ti.midline}" xHeight="${ti.xHeight}"/>`);
      }
      lines.push(`</hh:font>`);
    }
    lines.push(`</hh:fontface>`);
    fontOffset += cnt;
  }
  lines.push(`</hh:fontfaces>`);

  // borderFills
  lines.push(`<hh:borderFills itemCnt="${info.borderFills.length}">`);
  for (const bf of info.borderFills) {
    lines.push(generateBorderFillXml(bf, ''));
  }
  lines.push(`</hh:borderFills>`);

  // charProperties
  lines.push(`<hh:charProperties itemCnt="${info.charShapes.length}">`);
  for (let i = 0; i < info.charShapes.length; i++) {
    const cs = info.charShapes[i];
    const shadeStr = colorrefToHexOrNone(cs.shadeColor);
    const bfRef = cs.borderFillId;
    const uFontSpace = cs.useFontSpace ? '1' : '0';
    const uKerning = cs.useKerning ? '1' : '0';
    const symMarkStr = SYM_MARK_MAP[cs.symMark] ?? 'NONE';
    lines.push(`<hh:charPr id="${i}" height="${cs.height}" textColor="${colorrefToHex(cs.textColor)}" shadeColor="${shadeStr}" useFontSpace="${uFontSpace}" useKerning="${uKerning}" symMark="${symMarkStr}" borderFillIDRef="${bfRef}">`);

    const fontRefAttrs = LANG_ATTRS.map((l, j) => `${l}="${cs.fontIds[j] ?? 0}"`).join(' ');
    const ratioAttrs = LANG_ATTRS.map((l, j) => `${l}="${cs.ratio[j] ?? 100}"`).join(' ');
    const spacingAttrs = LANG_ATTRS.map((l, j) => `${l}="${cs.spacing[j] ?? 0}"`).join(' ');
    const relSzAttrs = LANG_ATTRS.map((l, j) => `${l}="${cs.relSize[j] ?? 100}"`).join(' ');
    const offsetAttrs = LANG_ATTRS.map((l, j) => `${l}="${cs.offset[j] ?? 0}"`).join(' ');

    lines.push(`<hh:fontRef ${fontRefAttrs}/>`);
    lines.push(`<hh:ratio ${ratioAttrs}/>`);
    lines.push(`<hh:spacing ${spacingAttrs}/>`);
    lines.push(`<hh:relSz ${relSzAttrs}/>`);
    lines.push(`<hh:offset ${offsetAttrs}/>`);
    if (cs.bold) lines.push(`<hh:bold/>`);
    if (cs.italic) lines.push(`<hh:italic/>`);
    if (cs.superscript) lines.push(`<hh:superscript/>`);
    if (cs.subscript) lines.push(`<hh:subscript/>`);
    const ulType = UNDERLINE_TYPE_MAP[cs.underlineType] ?? 'NONE';
    const ulShape = UNDERLINE_SHAPE_MAP[cs.underlineShape] ?? 'SOLID';
    lines.push(`<hh:underline type="${ulType}" shape="${ulShape}" color="${colorrefToHex(cs.underlineColor)}"/>`);
    const stShape = cs.strikeout === 0 ? 'NONE' : (UNDERLINE_SHAPE_MAP[cs.strikeoutShape] ?? 'SOLID');
    lines.push(`<hh:strikeout shape="${stShape}" color="${colorrefToHex(cs.strikeoutColor)}"/>`);
    const olType = OUTLINE_TYPE_MAP[cs.outlineType] ?? 'NONE';
    lines.push(`<hh:outline type="${olType}"/>`);
    const shType = SHADOW_TYPE_MAP[cs.shadowType] ?? 'NONE';
    lines.push(`<hh:shadow type="${shType}" color="${colorrefToHex(cs.shadowColor)}" offsetX="${cs.shadowX}" offsetY="${cs.shadowY}"/>`);
    lines.push(`</hh:charPr>`);
  }
  lines.push(`</hh:charProperties>`);

  // tabProperties
  lines.push(`<hh:tabProperties itemCnt="${info.tabDefs.length}">`);
  for (let i = 0; i < info.tabDefs.length; i++) {
    const td = info.tabDefs[i];
    const autoL = td.autoTabLeft ? '1' : '0';
    const autoR = td.autoTabRight ? '1' : '0';
    if (td.items.length === 0) {
      lines.push(`<hh:tabPr id="${i}" autoTabLeft="${autoL}" autoTabRight="${autoR}"/>`);
    } else {
      lines.push(`<hh:tabPr id="${i}" autoTabLeft="${autoL}" autoTabRight="${autoR}">`);
      for (const item of td.items) {
        const typeStr = TAB_TYPE_MAP[item.type] ?? 'LEFT';
        const leaderStr = TAB_LEADER_MAP[item.leader] ?? 'NONE';
        lines.push(`<hp:switch>`);
        lines.push(`<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">`);
        lines.push(`<hh:tabItem pos="${Math.round(item.pos / 2)}" type="${typeStr}" leader="${leaderStr}" unit="HWPUNIT"/>`);
        lines.push(`</hp:case>`);
        lines.push(`<hp:default>`);
        lines.push(`<hh:tabItem pos="${item.pos}" type="${typeStr}" leader="${leaderStr}"/>`);
        lines.push(`</hp:default>`);
        lines.push(`</hp:switch>`);
      }
      lines.push(`</hh:tabPr>`);
    }
  }
  lines.push(`</hh:tabProperties>`);

  // bullets
  if (info.bullets.length > 0) {
    lines.push(`<hh:bullets itemCnt="${info.bullets.length}">`);
    for (const b of info.bullets) {
      lines.push(`<hh:bullet id="${b.id}" char="${escapeXml(b.char)}" useImage="${b.useImage ? 1 : 0}">`);
      const alignStr = b.align === 0 ? 'LEFT' : b.align === 1 ? 'CENTER' : 'RIGHT';
      const offsetTypeStr = b.textOffsetType === 0 ? 'PERCENT' : 'HWPUNIT';
      const numFmtStr = 'DIGIT';
      lines.push(`<hh:paraHead level="${b.level}" align="${alignStr}" useInstWidth="${b.useInstWidth ? 1 : 0}" autoIndent="${b.autoIndent ? 1 : 0}" widthAdjust="${b.widthAdjust}" textOffsetType="${offsetTypeStr}" textOffset="${b.textOffset}" numFormat="${numFmtStr}" charPrIDRef="${b.charPrIdRef >>> 0}" checkable="${b.checkable ? 1 : 0}"/>`);
      lines.push(`</hh:bullet>`);
    }
    lines.push(`</hh:bullets>`);
  }

  // paraProperties
  lines.push(`<hh:paraProperties itemCnt="${info.paraShapes.length}">`);
  for (let i = 0; i < info.paraShapes.length; i++) {
    const ps = info.paraShapes[i];
    const horAlign = ALIGN_MAP[ps.alignment] ?? 'JUSTIFY';
    const vertAlignStr = VERT_ALIGN_MAP[ps.vertAlign] ?? 'BASELINE';
    const headTypeStr = HEADING_TYPE_MAP[ps.headingType] ?? 'NONE';
    const breakLatinStr = BREAK_LATIN_MAP[ps.breakLatinWord] ?? 'KEEP_WORD';
    const breakNonLatinStr = ps.breakNonLatinWord ? 'KEEP_WORD' : 'BREAK_WORD';
    const lineWrapStr = LINE_WRAP_MAP[ps.lineWrap] ?? 'BREAK';
    const lineSpacingTypeStr = LINE_SPACING_TYPE_MAP[ps.lineSpacingType] ?? 'PERCENT';
    const snapGrid = ps.snapToGrid ? '1' : '0';
    const fontLH = ps.fontLineHeight ? '1' : '0';

    lines.push(`<hh:paraPr id="${i}" tabPrIDRef="${ps.tabDefId}" condense="${ps.condense}" fontLineHeight="${fontLH}" snapToGrid="${snapGrid}" suppressLineNumbers="0" checked="0">`);
    lines.push(`<hh:align horizontal="${horAlign}" vertical="${vertAlignStr}"/>`);
    lines.push(`<hh:heading type="${headTypeStr}" idRef="${ps.numberingId}" level="${ps.headingLevel}"/>`);
    lines.push(`<hh:breakSetting breakLatinWord="${breakLatinStr}" breakNonLatinWord="${breakNonLatinStr}" widowOrphan="${ps.widowOrphan ? 1 : 0}" keepWithNext="${ps.keepWithNext ? 1 : 0}" keepLines="${ps.keepLines ? 1 : 0}" pageBreakBefore="${ps.pageBreakBefore ? 1 : 0}" lineWrap="${lineWrapStr}"/>`);
    lines.push(`<hh:autoSpacing eAsianEng="${ps.autoSpacingEng ? 1 : 0}" eAsianNum="${ps.autoSpacingNum ? 1 : 0}"/>`);

    // hp:switch for margins and line spacing
    lines.push(`<hp:switch>`);
    lines.push(`<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">`);
    lines.push(`<hh:margin>`);
    lines.push(`<hc:intent value="${Math.round(ps.indent / 2)}" unit="HWPUNIT"/>`);
    lines.push(`<hc:left value="${Math.round(ps.leftMargin / 2)}" unit="HWPUNIT"/>`);
    lines.push(`<hc:right value="${Math.round(ps.rightMargin / 2)}" unit="HWPUNIT"/>`);
    lines.push(`<hc:prev value="${Math.round(ps.spacingBefore / 2)}" unit="HWPUNIT"/>`);
    lines.push(`<hc:next value="${Math.round(ps.spacingAfter / 2)}" unit="HWPUNIT"/>`);
    lines.push(`</hh:margin>`);
    if (lineSpacingTypeStr === 'PERCENT') {
      lines.push(`<hh:lineSpacing type="${lineSpacingTypeStr}" value="${ps.lineSpacing}" unit="HWPUNIT"/>`);
    } else {
      lines.push(`<hh:lineSpacing type="${lineSpacingTypeStr}" value="${Math.round(ps.lineSpacing / 2)}" unit="HWPUNIT"/>`);
    }
    lines.push(`</hp:case>`);
    lines.push(`<hp:default>`);
    lines.push(`<hh:margin>`);
    lines.push(`<hc:intent value="${ps.indent}" unit="HWPUNIT"/>`);
    lines.push(`<hc:left value="${ps.leftMargin}" unit="HWPUNIT"/>`);
    lines.push(`<hc:right value="${ps.rightMargin}" unit="HWPUNIT"/>`);
    lines.push(`<hc:prev value="${ps.spacingBefore}" unit="HWPUNIT"/>`);
    lines.push(`<hc:next value="${ps.spacingAfter}" unit="HWPUNIT"/>`);
    lines.push(`</hh:margin>`);
    lines.push(`<hh:lineSpacing type="${lineSpacingTypeStr}" value="${ps.lineSpacing}" unit="HWPUNIT"/>`);
    lines.push(`</hp:default>`);
    lines.push(`</hp:switch>`);

    const connect = ps.borderConnect ? '1' : '0';
    const ignore = ps.ignoreMargin ? '1' : '0';
    lines.push(`<hh:border borderFillIDRef="${ps.borderFillId}" offsetLeft="${ps.borderLeft}" offsetRight="${ps.borderRight}" offsetTop="${ps.borderTop}" offsetBottom="${ps.borderBottom}" connect="${connect}" ignoreMargin="${ignore}"/>`);
    lines.push(`</hh:paraPr>`);
  }
  lines.push(`</hh:paraProperties>`);

  // styles
  if (info.styles.length > 0) {
    lines.push(`<hh:styles itemCnt="${info.styles.length}">`);
    for (let i = 0; i < info.styles.length; i++) {
      const st = info.styles[i];
      const typeStr = st.type === 0 ? 'PARA' : 'CHAR';
      lines.push(`<hh:style id="${i}" type="${typeStr}" name="${escapeXml(st.name)}" engName="${escapeXml(st.engName)}" paraPrIDRef="${st.paraPrId}" charPrIDRef="${st.charPrId}" nextStyleIDRef="${st.nextStyleId}" langID="${st.langId}" lockForm="0"/>`);
    }
    lines.push(`</hh:styles>`);
  }

  lines.push(`</hh:refList>`);

  lines.push(`<hh:compatibleDocument targetProgram="HWP201X">`);
  lines.push(`<hh:layoutCompatibility/>`);
  lines.push(`</hh:compatibleDocument>`);
  lines.push(`<hh:docOption>`);
  lines.push(`<hh:linkinfo path="" pageInherit="1" footnoteInherit="0"/>`);
  lines.push(`</hh:docOption>`);
  lines.push(`<hh:trackchageConfig flags="56"/>`);

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
  const ind4 = ind3 + '  ';

  const pageBreakVal = para.pageBreak ? '1' : '0';
  const columnBreakVal = para.columnBreak ? '1' : '0';
  lines.push(`${indent}<hp:p id="${para.paraId}" paraPrIDRef="${para.paraPrId}" styleIDRef="${para.styleId}" pageBreak="${pageBreakVal}" columnBreak="${columnBreakVal}" merged="${para.merged}">`);

  const defCharPrId = para.defaultCharPrId;
  let lastRunCharPrId = -1;  // -1 means no run emitted yet

  // secd + cold go in the same run (no <hp:t/>)
  const secdCtrl = para.controls.find(c => c.type === 'secd') as (typeof para.controls[0] & { type: 'secd' }) | undefined;
  const coldCtrls = para.controls.filter(c => c.type === 'cold') as ColDefInfo[];
  if (secdCtrl) {
    lines.push(`${ind2}<hp:run charPrIDRef="${defCharPrId}">`);
    lines.push(generateSecPrXml(secdCtrl, ind3));
    for (const cold of coldCtrls) {
      lines.push(`${ind3}<hp:ctrl>`);
      lines.push(generateColPrXml(cold, ind4));
      lines.push(`${ind3}</hp:ctrl>`);
    }
    lines.push(`${ind2}</hp:run>`);
    lastRunCharPrId = defCharPrId;
  } else {
    // Standalone cold controls (not paired with secd)
    for (const cold of coldCtrls) {
      lines.push(`${ind2}<hp:run charPrIDRef="${defCharPrId}">`);
      lines.push(`${ind3}<hp:ctrl>`);
      lines.push(generateColPrXml(cold, ind4));
      lines.push(`${ind3}</hp:ctrl>`);
      lines.push(`${ind2}</hp:run>`);
      lastRunCharPrId = defCharPrId;
    }
  }

  // Determine which controls are "other" (non-secd/cold), with their charPrIds and stream positions.
  // Sort by stream position so they're emitted in document order.
  const otherCtrlsWithPrId: { ctrl: typeof para.controls[0]; charPrId: number; pos: number }[] = [];
  for (let ci = 0; ci < para.controls.length; ci++) {
    const ctrl = para.controls[ci];
    if (ctrl.type !== 'secd' && ctrl.type !== 'cold') {
      otherCtrlsWithPrId.push({
        ctrl,
        charPrId: para.ctrlCharPrIds[ci] ?? defCharPrId,
        pos: para.ctrlStreamPositions[ci] ?? 0,
      });
    }
  }
  otherCtrlsWithPrId.sort((a, b) => a.pos - b.pos);

  const hasFieldCtrls = otherCtrlsWithPrId.some(x => x.ctrl.type === 'fieldBegin' || x.ctrl.type === 'fieldEnd');

  if (hasFieldCtrls) {
    // Position-based interleaving for paragraphs with field controls.
    // textBeforeCtrl(i) = number of text chars before ctrl at position pos[i]
    //   = pos[i] - i  (since i ctrls before it each occupy 1 logical position)
    const totalTextChars = para.runs.reduce((s, r) => s + r.text.length, 0);
    // Build text segments between ctrls
    const segStarts: number[] = [];
    const segEnds: number[] = [];
    for (let oi = 0; oi <= otherCtrlsWithPrId.length; oi++) {
      const start = oi === 0 ? 0 : segEnds[oi - 1];
      const end = oi < otherCtrlsWithPrId.length
        ? otherCtrlsWithPrId[oi].pos - oi
        : totalTextChars;
      segStarts.push(start);
      segEnds.push(end);
    }

    for (let oi = 0; oi < otherCtrlsWithPrId.length; oi++) {
      // Emit text segment before this ctrl
      const textSeg = sliceRuns(para.runs, segStarts[oi], segEnds[oi]);
      for (const run of textSeg) {
        const runXml = generateRunXml(run, ind2);
        if (runXml) { lines.push(runXml); lastRunCharPrId = run.charPrId; }
      }

      const { ctrl, charPrId } = otherCtrlsWithPrId[oi];

      // For fieldEnd and table: check if following text segment can be embedded in this run
      let embeddedText = '';
      const nextSeg = sliceRuns(para.runs, segStarts[oi + 1], segEnds[oi + 1]);
      const nextText = nextSeg.map(r => r.text.replace(/\n/g, '')).join('');
      const nextCharPrId = nextSeg.length > 0 ? nextSeg[0].charPrId : -1;
      const canEmbed = nextText.length > 0 && nextCharPrId === charPrId &&
        (ctrl.type === 'fieldEnd' || ctrl.type === 'table');
      if (canEmbed) {
        embeddedText = nextText;
        // Mark the next segment as consumed by advancing its end to its start
        segEnds[oi + 1] = segStarts[oi + 1]; // consumed
      }

      lines.push(`${ind2}<hp:run charPrIDRef="${charPrId}">`);
      if (ctrl.type === 'fieldBegin') {
        lines.push(`${ind3}<hp:ctrl>`);
        lines.push(generateFieldBeginXml(ctrl as FieldBeginControlInfo, ind4));
        lines.push(`${ind3}</hp:ctrl>`);
      } else if (ctrl.type === 'fieldEnd') {
        lines.push(`${ind3}<hp:ctrl>`);
        lines.push(generateFieldEndXml(ctrl as FieldEndControlInfo, ind4));
        lines.push(`${ind3}</hp:ctrl>`);
        if (embeddedText) lines.push(`${ind3}<hp:t>${escapeXml(embeddedText)}</hp:t>`);
      } else if (ctrl.type === 'head' || ctrl.type === 'foot') {
        lines.push(`${ind3}<hp:ctrl>`);
        lines.push(...generateHeaderFooterXml(ctrl as HeaderFooterInfo, ind4));
        lines.push(`${ind3}</hp:ctrl>`);
      } else if (ctrl.type === 'pgnp') {
        lines.push(`${ind3}<hp:ctrl>`);
        lines.push(generatePageNumXml(ctrl as PageNumInfo, ind4));
        lines.push(`${ind3}</hp:ctrl>`);
      } else if (ctrl.type === 'table') {
        lines.push(...generateTableXml(ctrl as TableControlInfo, ind3));
        if (embeddedText) {
          lines.push(`${ind3}<hp:t>${escapeXml(embeddedText)}</hp:t>`);
        } else {
          lines.push(`${ind3}<hp:t/>`);
        }
      }
      lines.push(`${ind2}</hp:run>`);
      lastRunCharPrId = charPrId;
    }

    // Emit any remaining text after the last ctrl
    const lastSeg = sliceRuns(para.runs, segStarts[otherCtrlsWithPrId.length], segEnds[otherCtrlsWithPrId.length]);
    for (const run of lastSeg) {
      const runXml = generateRunXml(run, ind2);
      if (runXml) { lines.push(runXml); lastRunCharPrId = run.charPrId; }
    }
  } else {
    // Determine trailing text when ctrl chars precede text in binary.
    // This text gets embedded in the ctrl run's <hp:t> element (not a separate run).
    const trailingText = !para.textBeforeCtrl
      ? para.runs.map(r => r.text.replace(/\n/g, '')).join('')
      : '';

    // Helper to emit one combined run for otherCtrls, optionally embedding trailing text in table's <hp:t>
    function emitOtherCtrlsRun(charPrId: number, embedTrailingText: string): void {
      lines.push(`${ind2}<hp:run charPrIDRef="${charPrId}">`);
      for (let oi = 0; oi < otherCtrlsWithPrId.length; oi++) {
        const { ctrl } = otherCtrlsWithPrId[oi];
        const isLastTable = ctrl.type === 'table' && !otherCtrlsWithPrId.slice(oi + 1).some(x => x.ctrl.type === 'table');
        if (ctrl.type === 'head' || ctrl.type === 'foot') {
          lines.push(`${ind3}<hp:ctrl>`);
          lines.push(...generateHeaderFooterXml(ctrl as HeaderFooterInfo, ind4));
          lines.push(`${ind3}</hp:ctrl>`);
        } else if (ctrl.type === 'pgnp') {
          lines.push(`${ind3}<hp:ctrl>`);
          lines.push(generatePageNumXml(ctrl as PageNumInfo, ind4));
          lines.push(`${ind3}</hp:ctrl>`);
        } else if (ctrl.type === 'table') {
          lines.push(...generateTableXml(ctrl as TableControlInfo, ind3));
          if (isLastTable && embedTrailingText) {
            lines.push(`${ind3}<hp:t>${escapeXml(embedTrailingText)}</hp:t>`);
          } else {
            lines.push(`${ind3}<hp:t/>`);
          }
        }
      }
      lines.push(`${ind2}</hp:run>`);
      lastRunCharPrId = charPrId;
    }

    // If text appears before ctrl in binary, emit text runs first then other ctrl run
    // Otherwise emit other ctrl run (embedding any trailing text) then no separate text runs
    if (para.textBeforeCtrl && para.runs.length > 0 && otherCtrlsWithPrId.length > 0) {
      for (const run of para.runs) {
        const runXml = generateRunXml(run, ind2);
        if (runXml) { lines.push(runXml); lastRunCharPrId = run.charPrId; }
      }
      const runCharPrId = otherCtrlsWithPrId[0].charPrId;
      emitOtherCtrlsRun(runCharPrId, '');
    } else {
      if (otherCtrlsWithPrId.length > 0) {
        const runCharPrId = otherCtrlsWithPrId[0].charPrId;
        emitOtherCtrlsRun(runCharPrId, trailingText);
      } else {
        for (const run of para.runs) {
          const runXml = generateRunXml(run, ind2);
          if (runXml) { lines.push(runXml); lastRunCharPrId = run.charPrId; }
        }
      }
    }
  }

  // Para break empty run: emit if no run was emitted yet, or if para break charPrId differs from last run
  const paraBreakCharPrId = para.paraBreakCharPrId;
  if (lastRunCharPrId === -1 || paraBreakCharPrId !== lastRunCharPrId) {
    lines.push(`${ind2}<hp:run charPrIDRef="${paraBreakCharPrId}"/>`);
  }

  // linesegarray goes at the end
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

  lines.push(`${indent}</hp:p>`);
  return lines.join('\n');
}

// ── Field ctrl XML helpers ──

const FIELD_TYPE_MAP: Record<string, string> = {
  '%hlk': 'HYPERLINK', '%dte': 'DATE', '%ddt': 'DOCDATE', '%pat': 'PATH',
  '%bmk': 'BOOKMARK', '%mmg': 'MAILMERGE', '%xrf': 'CROSSREF', '%fmu': 'FORMULA',
  '%clk': 'CLICKHERE', '%smr': 'SUMMARY', '%usr': 'USERINFO', '%sig': 'REVISION_SIGN',
  '%spl': 'REVISION_SPLIT',
};

/** Slice a flat list of text runs to chars [start, end). */
function sliceRuns(runs: TextRunInfo[], start: number, end: number): TextRunInfo[] {
  if (start >= end) return [];
  const result: TextRunInfo[] = [];
  let pos = 0;
  for (const run of runs) {
    const runEnd = pos + run.text.length;
    if (runEnd <= start) { pos = runEnd; continue; }
    if (pos >= end) break;
    const s = Math.max(0, start - pos);
    const e = Math.min(run.text.length, end - pos);
    if (s < e) result.push({ charPrId: run.charPrId, text: run.text.substring(s, e) });
    pos = runEnd;
  }
  return result;
}

function generateFieldBeginXml(field: FieldBeginControlInfo, indent: string): string {
  const ind2 = indent + '  ';
  const ind3 = ind2 + '  ';
  const typeStr = FIELD_TYPE_MAP[field.ctrlId] ?? 'UNKNOWN';
  const editableVal = field.editable ? '1' : '0';
  const dirtyVal = field.dirty ? '1' : '0';
  const ctrlFieldId = ctrlIdToMake4CHID(field.ctrlId);
  const lines: string[] = [];

  // Derive parameters from command string for hyperlink
  if (field.ctrlId === '%hlk') {
    const parts = field.command.split(';');
    const path = parts[0] ?? '';
    let category = 'HWPHYPERLINK_TYPE_BASIC';
    if (path.startsWith('mailto:')) category = 'HWPHYPERLINK_TYPE_EMAIL';
    else if (path.startsWith('http://') || path.startsWith('https://')) category = 'HWPHYPERLINK_TYPE_WEB';
    lines.push(
      `${indent}<hp:fieldBegin id="${field.id}" type="${typeStr}" name="" editable="${editableVal}" ` +
      `dirty="${dirtyVal}" zorder="-1" fieldid="${ctrlFieldId}">`
    );
    lines.push(`${ind2}<hp:parameters cnt="6" name="">`);
    lines.push(`${ind3}<hp:integerParam name="Prop">0</hp:integerParam>`);
    lines.push(`${ind3}<hp:stringParam name="Command">${escapeXml(field.command)}</hp:stringParam>`);
    lines.push(`${ind3}<hp:stringParam name="Path">${escapeXml(path)}</hp:stringParam>`);
    lines.push(`${ind3}<hp:stringParam name="Category">${escapeXml(category)}</hp:stringParam>`);
    lines.push(`${ind3}<hp:stringParam name="TargetType">HWPHYPERLINK_TARGET_BOOKMARK</hp:stringParam>`);
    lines.push(`${ind3}<hp:stringParam name="DocOpenType">HWPHYPERLINK_JUMP_CURRENTTAB</hp:stringParam>`);
    lines.push(`${ind2}</hp:parameters>`);
    lines.push(`${indent}</hp:fieldBegin>`);
  } else {
    lines.push(
      `${indent}<hp:fieldBegin id="${field.id}" type="${typeStr}" name="" editable="${editableVal}" ` +
      `dirty="${dirtyVal}" zorder="-1" fieldid="${ctrlFieldId}"/>`
    );
  }
  return lines.join('\n');
}

function generateFieldEndXml(field: FieldEndControlInfo, indent: string): string {
  return `${indent}<hp:fieldEnd beginIDRef="${field.beginId}" fieldid="${ctrlIdToMake4CHID(field.ctrlId)}"/>`;
}

// ── Inline control XML helpers ──

const COL_TYPE_MAP: Record<number, string> = { 0: 'NEWSPAPER', 1: 'BALANCE', 2: 'PARALLEL' };
const COL_LAYOUT_MAP: Record<number, string> = { 0: 'LEFT', 1: 'RIGHT', 2: 'MIRROR' };

function generateColPrXml(cold: ColDefInfo, indent: string): string {
  const typeStr = COL_TYPE_MAP[cold.colType] ?? 'NEWSPAPER';
  const layoutStr = COL_LAYOUT_MAP[cold.layout] ?? 'LEFT';
  const sameSzVal = cold.sameSz ? '1' : '0';
  return `${indent}<hp:colPr id="" type="${typeStr}" layout="${layoutStr}" colCount="${cold.colCount}" sameSz="${sameSzVal}" sameGap="${cold.sameGap}"/>`;
}

const PAGE_NUM_POS_MAP: Record<number, string> = {
  0: 'NONE', 1: 'TOP_LEFT', 2: 'TOP_CENTER', 3: 'TOP_RIGHT',
  4: 'BOTTOM_LEFT', 5: 'BOTTOM_CENTER', 6: 'BOTTOM_RIGHT',
  7: 'OUTSIDE_TOP', 8: 'OUTSIDE_BOTTOM', 9: 'INSIDE_TOP', 10: 'INSIDE_BOTTOM',
};

const PAGE_NUM_FORMAT_MAP: Record<number, string> = {
  0: 'DIGIT', 1: 'CIRCLE_DIGIT', 2: 'ROMAN_CAPITAL', 3: 'ROMAN_SMALL',
  4: 'ALPHA_CAPITAL', 5: 'ALPHA_SMALL', 8: 'HANGUL', 9: 'CIRCLE_HANGUL',
};

function generatePageNumXml(pgnp: PageNumInfo, indent: string): string {
  const posStr = PAGE_NUM_POS_MAP[pgnp.pos] ?? 'NONE';
  const formatTypeStr = PAGE_NUM_FORMAT_MAP[pgnp.formatType] ?? 'DIGIT';
  const sideCharStr = pgnp.sideChar ? String.fromCharCode(pgnp.sideChar) : '-';
  return `${indent}<hp:pageNum pos="${posStr}" formatType="${formatTypeStr}" sideChar="${escapeXml(sideCharStr)}"/>`;
}

const APPLY_PAGE_TYPE_MAP: Record<number, string> = { 0: 'BOTH', 1: 'EVEN', 2: 'ODD' };
const TEXT_DIR_NAMES: Record<number, string> = { 0: 'HORIZONTAL', 1: 'VERTICAL' };
const LINE_WRAP_NAMES: Record<number, string> = { 0: 'BREAK', 1: 'SQUEEZE', 2: 'KEEP' };
const VERT_ALIGN_NAMES: Record<number, string> = { 0: 'TOP', 1: 'CENTER', 2: 'BOTTOM' };

function generateHeaderFooterXml(hf: HeaderFooterInfo, indent: string): string[] {
  const ind2 = indent + '  ';
  const ind3 = ind2 + '  ';
  const tag = hf.type === 'head' ? 'hp:header' : 'hp:footer';
  const applyPageTypeStr = APPLY_PAGE_TYPE_MAP[hf.applyPageType] ?? 'BOTH';
  const lines: string[] = [];
  lines.push(`${indent}<${tag} id="${hf.id}" applyPageType="${applyPageTypeStr}">`);
  // Determine subList attributes from the list header data (listAttrs=0 → default values)
  const textDirStr = TEXT_DIR_NAMES[0] ?? 'HORIZONTAL';
  const lineWrapStr = LINE_WRAP_NAMES[0] ?? 'BREAK';
  const vertAlignStr = VERT_ALIGN_NAMES[0] ?? 'TOP';
  lines.push(
    `${ind2}<hp:subList id="" textDirection="${textDirStr}" lineWrap="${lineWrapStr}" vertAlign="${vertAlignStr}" ` +
    `linkListIDRef="0" linkListNextIDRef="0" textWidth="${hf.textWidth}" textHeight="${hf.textHeight}" ` +
    `hasTextRef="0" hasNumRef="0">`
  );
  for (const para of hf.paragraphs) {
    lines.push(generateParaXml(para, ind3));
  }
  lines.push(`${ind2}</hp:subList>`);
  lines.push(`${indent}</${tag}>`);
  return lines;
}

// ── secPr XML generation ──

const NOTE_LINE_TYPE_MAP: Record<number, string> = {
  0: 'NONE', 1: 'SOLID', 2: 'DASHED', 3: 'DOTTED', 4: 'DASH_DOT',
  5: 'DASH_DOT_DOT', 6: 'LONG_DASH', 7: 'LARGE_DOT', 8: 'DOUBLE',
};
const NOTE_NUM_TYPE_MAP: Record<number, string> = {
  0: 'DIGIT', 1: 'CIRCLE_DIGIT', 2: 'ROMAN_CAPITAL', 3: 'ROMAN_SMALL',
  4: 'ALPHA_CAPITAL', 5: 'ALPHA_SMALL', 8: 'HANGUL', 9: 'CIRCLE_HANGUL',
  0x80: 'CUSTOM_4', 0x81: 'CUSTOM',
};
const NOTE_PLACE_FOOTNOTE_MAP: Record<number, string> = {
  0: 'EACH_COLUMN', 1: 'ALONGSIDE_TEXT', 2: 'RIGHTMOST_COLUMN',
};
const NOTE_PLACE_ENDNOTE_MAP: Record<number, string> = {
  0: 'END_OF_DOCUMENT', 1: 'END_OF_SECTION',
};
const NOTE_NUMBERING_MAP: Record<number, string> = {
  0: 'CONTINUOUS', 1: 'NEW_EACH_SECTION', 2: 'NEW_EACH_PAGE',
};

function generateNoteShapeXml(note: FootnoteShapeInfo, indent: string, isEndnote: boolean): string {
  const ind2 = indent + '  ';
  const tag = isEndnote ? 'hp:endNotePr' : 'hp:footNotePr';
  const placeMap = isEndnote ? NOTE_PLACE_ENDNOTE_MAP : NOTE_PLACE_FOOTNOTE_MAP;
  const numType = NOTE_NUM_TYPE_MAP[note.numberType] ?? 'DIGIT';
  const userCharStr = note.userChar ? String.fromCharCode(note.userChar) : '';
  const prefixCharStr = note.prefixChar ? String.fromCharCode(note.prefixChar) : '';
  const suffixCharStr = note.suffixChar ? String.fromCharCode(note.suffixChar) : '';
  const supscriptVal = note.supscript ? '1' : '0';
  const beneathVal = note.beneathText ? '1' : '0';
  const lineTypeStr = NOTE_LINE_TYPE_MAP[note.lineType] ?? 'NONE';
  const lineWidthStr = `${BORDER_WIDTHS[note.lineWidth] ?? 0.1} mm`;
  const lineColorStr = colorrefToHex(note.lineColor);
  const placeStr = placeMap[note.placement] ?? (isEndnote ? 'END_OF_DOCUMENT' : 'EACH_COLUMN');
  const numStr = NOTE_NUMBERING_MAP[note.numbering] ?? 'CONTINUOUS';
  const lines: string[] = [];
  lines.push(`${indent}<${tag}>`);
  lines.push(`${ind2}<hp:autoNumFormat type="${numType}" userChar="${escapeXml(userCharStr)}" prefixChar="${escapeXml(prefixCharStr)}" suffixChar="${escapeXml(suffixCharStr)}" supscript="${supscriptVal}"/>`);
  lines.push(`${ind2}<hp:noteLine length="${note.noteLineLength}" type="${lineTypeStr}" width="${lineWidthStr}" color="${lineColorStr}"/>`);
  lines.push(`${ind2}<hp:noteSpacing betweenNotes="${note.noteSpacing}" belowLine="${note.noteLineBottom}" aboveLine="${note.noteLineTop}"/>`);
  lines.push(`${ind2}<hp:numbering type="${numStr}" newNum="${note.startNumber}"/>`);
  lines.push(`${ind2}<hp:placement place="${placeStr}" beneathText="${beneathVal}"/>`);
  lines.push(`${indent}</${tag}>`);
  return lines.join('\n');
}

const PAGE_BORDER_TYPE_MAP: Record<number, string[]> = {
  0: ['BOTH',  'BOTH'],
  1: ['EVEN',  'EVEN'],
  2: ['ODD',   'ODD'],
};
const TEXT_BORDER_MAP: Record<number, string> = { 0: 'TEXT', 1: 'PAPER' };
const FILL_AREA_MAP: Record<number, string> = { 0: 'PAPER', 1: 'PAGE', 2: 'BORDER' };

function generatePageBorderFillXml(pbf: PageBorderFillInfo, pbfType: string, indent: string): string {
  const ind2 = indent + '  ';
  const textBorderStr = TEXT_BORDER_MAP[pbf.textBorder] ?? 'PAPER';
  const fillAreaStr = FILL_AREA_MAP[pbf.fillArea] ?? 'PAPER';
  const headerInsideVal = pbf.headerInside ? '1' : '0';
  const footerInsideVal = pbf.footerInside ? '1' : '0';
  return [
    `${indent}<hp:pageBorderFill type="${pbfType}" borderFillIDRef="${pbf.borderFillId}" textBorder="${textBorderStr}" headerInside="${headerInsideVal}" footerInside="${footerInsideVal}" fillArea="${fillAreaStr}">`,
    `${ind2}<hp:offset left="${pbf.leftGap}" right="${pbf.rightGap}" top="${pbf.topGap}" bottom="${pbf.bottomGap}"/>`,
    `${indent}</hp:pageBorderFill>`,
  ].join('\n');
}

const LANDSCAPE_MAP: Record<number, string> = { 0: 'WIDELY', 1: 'LANDSCAPE' };
const GUTTER_TYPE_MAP: Record<number, string> = { 0: 'LEFT_ONLY', 1: 'BOTH_SIDES', 2: 'TOP' };
const TEXT_DIR_MAP: Record<number, string> = { 0: 'HORIZONTAL', 1: 'VERTICAL' };
const PAGE_STARTS_MAP: Record<number, string> = { 0: 'BOTH', 1: 'ODD', 2: 'EVEN' };

function generateSecPrXml(secd: SectionDefInfo, indent: string): string {
  const ind2 = indent + '  ';
  const lines: string[] = [];
  const textDirStr = TEXT_DIR_MAP[(secd.attrs >> 16) & 7] ?? 'HORIZONTAL';
  const tabStopVal = Math.round(secd.tabStop / 2);
  lines.push(
    `${indent}<hp:secPr id="" textDirection="${textDirStr}" spaceColumns="${secd.spaceColumns}" ` +
    `tabStop="${secd.tabStop}" tabStopVal="${tabStopVal}" tabStopUnit="HWPUNIT" ` +
    `outlineShapeIDRef="${secd.outlineShapeIDRef}" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">`
  );
  lines.push(
    `${ind2}<hp:grid lineGrid="${secd.lineGrid}" charGrid="${secd.charGrid}" wonggojiFormat="${(secd.attrs >> 22) & 1}"/>`
  );
  const pageStartsStr = PAGE_STARTS_MAP[(secd.attrs >> 20) & 3] ?? 'BOTH';
  lines.push(
    `${ind2}<hp:startNum pageStartsOn="${pageStartsStr}" page="${secd.pageNum}" ` +
    `pic="${secd.picNum}" tbl="${secd.tblNum}" equation="${secd.eqNum}"/>`
  );
  const hideHeader = (secd.attrs & 1) ? '1' : '0';
  const hideFooter = (secd.attrs >> 1) & 1 ? '1' : '0';
  const hideMasterPage = (secd.attrs >> 2) & 1 ? '1' : '0';
  const hidePageNum = (secd.attrs >> 5) & 1 ? '1' : '0';
  const hideEmptyLine = (secd.attrs >> 19) & 1 ? '1' : '0';
  const borderBit = (secd.attrs >> 3) & 1;
  const borderFirstOnly = (secd.attrs >> 8) & 1;
  const borderStr = borderBit === 0 ? 'SHOW_ALL' : (borderFirstOnly ? 'SHOW_FIRST_ONLY' : 'HIDE_ALL');
  const fillBit = (secd.attrs >> 4) & 1;
  const fillFirstOnly = (secd.attrs >> 9) & 1;
  const fillStr = fillBit === 0 ? 'SHOW_ALL' : (fillFirstOnly ? 'SHOW_FIRST_ONLY' : 'HIDE_ALL');
  lines.push(
    `${ind2}<hp:visibility hideFirstHeader="${hideHeader}" hideFirstFooter="${hideFooter}" ` +
    `hideFirstMasterPage="${hideMasterPage}" border="${borderStr}" fill="${fillStr}" ` +
    `hideFirstPageNum="${hidePageNum}" hideFirstEmptyLine="${hideEmptyLine}" showLineNumber="0"/>`
  );
  lines.push(
    `${ind2}<hp:lineNumberShape restartType="${secd.lineNumRestartType}" countBy="${secd.lineNumCountBy}" ` +
    `distance="${secd.lineNumDistance}" startNumber="${secd.lineNumStartNumber}"/>`
  );
  // pagePr from pageDef
  const pd = secd.pageDef;
  const landscapeStr = LANDSCAPE_MAP[pd.landscape] ?? 'WIDELY';
  const gutterTypeStr = GUTTER_TYPE_MAP[pd.gutterType] ?? 'LEFT_ONLY';
  lines.push(`${ind2}<hp:pagePr landscape="${landscapeStr}" width="${pd.width}" height="${pd.height}" gutterType="${gutterTypeStr}">`);
  lines.push(
    `${ind2}  <hp:margin header="${pd.marginHeader}" footer="${pd.marginFooter}" gutter="${pd.marginGutter}" ` +
    `left="${pd.marginLeft}" right="${pd.marginRight}" top="${pd.marginTop}" bottom="${pd.marginBottom}"/>`
  );
  lines.push(`${ind2}</hp:pagePr>`);
  // footnote and endnote
  lines.push(generateNoteShapeXml(secd.footnote, ind2, false));
  lines.push(generateNoteShapeXml(secd.endnote, ind2, true));
  // page border/fill (BOTH, EVEN, ODD)
  const pbfTypes = ['BOTH', 'EVEN', 'ODD'];
  for (let bi = 0; bi < secd.pageBorderFills.length && bi < 3; bi++) {
    lines.push(generatePageBorderFillXml(secd.pageBorderFills[bi], pbfTypes[bi], ind2));
  }
  // presentation element (default values)
  lines.push(
    `${ind2}<hp:presentation effect="none" soundIDRef="" invertText="1" autoshow="0" showtime="0" applyto="WholeDoc">` +
    `<hc:fillBrush><hc:gradation type="LINEAR" angle="0" centerX="50" centerY="0" step="100" colorNum="2" stepCenter="50" alpha="0">` +
    `<hc:color value="#0000FF"/><hc:color value="#000000"/></hc:gradation></hc:fillBrush></hp:presentation>`
  );
  lines.push(`${indent}</hp:secPr>`);
  return lines.join('\n');
}

// textWrap binary value → HWPML string
// 0=SQUARE, 1=TOP_AND_BOTTOM, 2=TIGHT, 3=THROUGH, 4=BEHIND_TEXT, 5=IN_FRONT_OF_TEXT
const TEXT_WRAP_MAP: Record<number, string> = {
  0: 'SQUARE', 1: 'TOP_AND_BOTTOM', 2: 'TIGHT', 3: 'THROUGH',
  4: 'BEHIND_TEXT', 5: 'IN_FRONT_OF_TEXT',
};
// textFlow binary value → HWPML string
// 0=BOTH_SIDES, 1=LEFT_ONLY, 2=RIGHT_ONLY, 3=LARGEST_ONLY
const TEXT_FLOW_MAP: Record<number, string> = {
  0: 'BOTH_SIDES', 1: 'LEFT_ONLY', 2: 'RIGHT_ONLY', 3: 'LARGEST_ONLY',
};
// vertRelTo binary value → HWPML string (0=paper, 1=page, 2=para)
const VERT_REL_TO_MAP: Record<number, string> = { 0: 'PAPER', 1: 'PAGE', 2: 'PARA', 3: 'LINE' };
// horzRelTo binary value → HWPML string (0=page, 1=page, 2=column, 3=para)
const HORZ_REL_TO_MAP: Record<number, string> = { 0: 'PAGE', 1: 'PAGE', 2: 'COLUMN', 3: 'PARA' };
// cellVertAlign binary value → HWPML string (0=TOP, 1=CENTER, 2=BOTTOM)
const CELL_VERT_ALIGN_MAP: Record<number, string> = { 0: 'TOP', 1: 'CENTER', 2: 'BOTTOM' };
// horzAlign binary value → HWPML string (0=LEFT, 1=CENTER, 2=RIGHT)
const HORZ_ALIGN_MAP: Record<number, string> = { 0: 'LEFT', 1: 'CENTER', 2: 'RIGHT' };
// pageBreak binary value → HWPML string (0=NONE, 2=CELL)
const PAGE_BREAK_MAP: Record<number, string> = { 0: 'NONE', 1: 'PAGE', 2: 'CELL', 3: 'COLUMN' };

function generateTableXml(table: TableControlInfo, indent: string): string[] {
  const lines: string[] = [];
  const ind2 = indent + '  ';
  const ind3 = ind2 + '  ';
  const ind4 = ind3 + '  ';

  const textWrapStr = TEXT_WRAP_MAP[table.textWrap] ?? 'TOP_AND_BOTTOM';
  const textFlowStr = TEXT_FLOW_MAP[table.textFlow] ?? 'BOTH_SIDES';
  const pageBreakStr = PAGE_BREAK_MAP[table.pageBreakType] ?? 'NONE';
  const repeatHeaderVal = table.repeatHeader ? '1' : '0';
  const noAdjustVal = table.noAdjust ? '1' : '0';
  const lockVal = table.lock ? '1' : '0';

  // <hp:tbl> with required HWPML attributes
  lines.push(
    `${indent}<hp:tbl id="${table.instanceId}" zOrder="${table.zOrder}" numberingType="TABLE" textWrap="${textWrapStr}" ` +
    `textFlow="${textFlowStr}" lock="${lockVal}" dropcapstyle="None" pageBreak="${pageBreakStr}" repeatHeader="${repeatHeaderVal}" ` +
    `rowCnt="${table.rowCount}" colCnt="${table.colCount}" ` +
    `cellSpacing="${table.cellSpacing}" borderFillIDRef="${table.borderFillId}" noAdjust="${noAdjustVal}">`
  );
  // Size and position as child elements
  lines.push(
    `${ind2}<hp:sz width="${table.ctrlWidth}" widthRelTo="ABSOLUTE" ` +
    `height="${table.ctrlHeight}" heightRelTo="ABSOLUTE" protect="0"/>`
  );
  const treatAsCharVal = table.treatAsChar ? '1' : '0';
  const affectLSpacingVal = table.affectLSpacing ? '1' : '0';
  const flowWithTextVal = table.flowWithText ? '1' : '0';
  const allowOverlapVal = table.allowOverlap ? '1' : '0';
  const holdAnchorVal = table.holdAnchorAndSO ? '1' : '0';
  const vertRelToStr = VERT_REL_TO_MAP[table.vertRelTo] ?? 'PARA';
  const horzRelToStr = HORZ_REL_TO_MAP[table.horzRelTo] ?? 'PARA';
  const vertAlignStr = CELL_VERT_ALIGN_MAP[table.vertAlignPos] ?? 'TOP';
  const horzAlignStr = HORZ_ALIGN_MAP[table.horzAlignPos] ?? 'LEFT';
  lines.push(
    `${ind2}<hp:pos treatAsChar="${treatAsCharVal}" affectLSpacing="${affectLSpacingVal}" flowWithText="${flowWithTextVal}" allowOverlap="${allowOverlapVal}" ` +
    `holdAnchorAndSO="${holdAnchorVal}" vertRelTo="${vertRelToStr}" horzRelTo="${horzRelToStr}" ` +
    `vertAlign="${vertAlignStr}" horzAlign="${horzAlignStr}" vertOffset="${table.yOffset}" horzOffset="${table.xOffset}"/>`
  );
  lines.push(
    `${ind2}<hp:outMargin left="${table.outMarginLeft}" right="${table.outMarginRight}" ` +
    `top="${table.outMarginTop}" bottom="${table.outMarginBottom}"/>`
  );
  if (table.captionParas && table.captionParas.length > 0) {
    const CAPTION_SIDE_MAP: Record<number, string> = { 0: 'LEFT', 1: 'RIGHT', 2: 'TOP', 3: 'BOTTOM' };
    const side = CAPTION_SIDE_MAP[table.captionDir ?? 3] ?? 'BOTTOM';
    const fullSzVal = (table.captionFullSz ?? false) ? '1' : '0';
    lines.push(
      `${ind2}<hp:caption side="${side}" fullSz="${fullSzVal}" ` +
      `width="${table.captionWidth ?? 0}" gap="${table.captionGap ?? 0}" lastWidth="${table.captionLastWidth ?? 0}">`
    );
    lines.push(
      `${ind3}<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" ` +
      `linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`
    );
    for (const para of table.captionParas) {
      lines.push(generateParaXml(para, ind4));
    }
    lines.push(`${ind3}</hp:subList>`);
    lines.push(`${ind2}</hp:caption>`);
  }
  lines.push(
    `${ind2}<hp:inMargin left="${table.innerMarginLeft}" right="${table.innerMarginRight}" ` +
    `top="${table.innerMarginTop}" bottom="${table.innerMarginBottom}"/>`
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
    lines.push(`${ind2}<hp:tr>`);
    for (const cell of cells) {
      const vertAlignStr = cell.vertAlign === 1 ? 'CENTER' : cell.vertAlign === 2 ? 'BOTTOM' : 'TOP';
      const hasMarginVal = cell.hasMargin ? '1' : '0';
      const headerCellVal = cell.headerCell ? '1' : '0';
      lines.push(
        `${ind3}<hp:tc name="" header="${headerCellVal}" hasMargin="${hasMarginVal}" ` +
        `protect="0" editable="0" dirty="0" borderFillIDRef="${cell.borderFillId}">`
      );
      const ind5 = ind4 + '  ';
      const lineWrapStr = cell.lineWrap === 1 ? 'SQUEEZE' : cell.lineWrap === 2 ? 'KEEP' : 'BREAK';
      lines.push(
        `${ind4}<hp:subList id="" textDirection="HORIZONTAL" lineWrap="${lineWrapStr}" ` +
        `vertAlign="${vertAlignStr}" linkListIDRef="0" linkListNextIDRef="0" ` +
        `textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`
      );
      for (const para of cell.paragraphs) {
        lines.push(generateParaXml(para, ind5));
      }
      lines.push(`${ind4}</hp:subList>`);
      // Cell address/span/size/margin as child elements (after subList)
      lines.push(`${ind4}<hp:cellAddr colAddr="${cell.colAddr}" rowAddr="${cell.rowAddr}"/>`);
      lines.push(`${ind4}<hp:cellSpan colSpan="${cell.colSpan}" rowSpan="${cell.rowSpan}"/>`);
      lines.push(`${ind4}<hp:cellSz width="${cell.width}" height="${cell.height}"/>`);
      // HWPML uses UINT32 max (4294967295) for cells that inherit table margins (binary stores UINT16 max 0xFFFF)
      const cm = (v: number) => v === 0xFFFF ? 4294967295 : v;
      lines.push(
        `${ind4}<hp:cellMargin left="${cm(cell.marginLeft)}" right="${cm(cell.marginRight)}" ` +
        `top="${cm(cell.marginTop)}" bottom="${cm(cell.marginBottom)}"/>`
      );
      lines.push(`${ind3}</hp:tc>`);
    }
    lines.push(`${ind2}</hp:tr>`);
  }

  lines.push(`${ind2}</hp:tbl>`);
  return lines;
}

export function generateSectionXml(
  sectionIndex: number,
  pageDef: PageDefInfo,
  paragraphs: ParaInfo[],
): string {
  const lines: string[] = [];
  lines.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`);
  lines.push(`<hs:sec ${HWPX_NS}>`);

  for (const para of paragraphs) {
    lines.push(generateParaXml(para, '  '));
  }

  lines.push(`</hs:sec>`);
  return lines.join('\n');
}

// ── Manifest / content.hpf ──

export function generateContentHpf(sectionCount: number): string {
  const parts: string[] = [];
  parts.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`);
  parts.push(`<opf:package ${HWPX_NS} version="" unique-identifier="" id="">`);
  parts.push(`<opf:metadata>`);
  parts.push(`<opf:language>ko</opf:language>`);
  parts.push(`</opf:metadata>`);
  parts.push(`<opf:manifest>`);
  parts.push(`<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>`);
  for (let i = 0; i < sectionCount; i++) {
    parts.push(`<opf:item id="section${i}" href="Contents/section${i}.xml" media-type="application/xml"/>`);
  }
  parts.push(`<opf:item id="settings" href="settings.xml" media-type="application/xml"/>`);
  parts.push(`</opf:manifest>`);
  parts.push(`<opf:spine>`);
  parts.push(`<opf:itemref idref="header" linear="yes"/>`);
  for (let i = 0; i < sectionCount; i++) {
    parts.push(`<opf:itemref idref="section${i}" linear="yes"/>`);
  }
  parts.push(`</opf:spine>`);
  parts.push(`</opf:package>`);
  return parts.join('');
}

export function generateContainerXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">` +
    `<ocf:rootfiles>` +
    `<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>` +
    `<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>` +
    `<ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/>` +
    `</ocf:rootfiles></ocf:container>`
  );
}

export function generateContainerRdf(sectionCount: number): string {
  const NS = 'http://www.hancom.co.kr/hwpml/2016/meta/pkg#';
  const parts: string[] = [];
  parts.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`);
  parts.push(`<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">`);
  parts.push(
    `<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="${NS}" rdf:resource="Contents/header.xml"/></rdf:Description>` +
    `<rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="${NS}HeaderFile"/></rdf:Description>`
  );
  for (let i = 0; i < sectionCount; i++) {
    parts.push(
      `<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="${NS}" rdf:resource="Contents/section${i}.xml"/></rdf:Description>` +
      `<rdf:Description rdf:about="Contents/section${i}.xml"><rdf:type rdf:resource="${NS}SectionFile"/></rdf:Description>`
    );
  }
  parts.push(`<rdf:Description rdf:about=""><rdf:type rdf:resource="${NS}Document"/></rdf:Description>`);
  parts.push(`</rdf:RDF>`);
  return parts.join('');
}

export function generateManifestXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`
  );
}

export function generateSettingsXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" ` +
    `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">` +
    `<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>` +
    `</ha:HWPApplicationSetting>`
  );
}

export function generateVersionXml(header: DocHeader): string {
  const { major, minor, patch, revision } = header.version;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" ` +
    `tagetApplication="WORDPROCESSOR" ` +
    `major="${major}" minor="${minor}" micro="${patch}" buildNumber="${revision}" ` +
    `os="1" xmlVersion="1.4" application="hwp2svg"/>`
  );
}
