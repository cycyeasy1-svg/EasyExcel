/**
 * IStyleData → ExcelJS 样式反向映射（import.ts 的逆向）。
 * 供 apply.ts 在增量写回与整表重建时使用。
 */
import type * as ExcelJS from '@cweijan/exceljs';
import {
    BorderStyleTypes,
    HorizontalAlign,
    VerticalAlign,
    WrapStrategy,
    type IBorderData,
    type IStyleData,
} from '@univerjs/core';
import { hexToArgb } from '../excel_styles_core';

const BORDER_STYLE_REVERSE: Record<number, ExcelJS.BorderStyle> = {
    [BorderStyleTypes.THIN]: 'thin',
    [BorderStyleTypes.HAIR]: 'hair',
    [BorderStyleTypes.DOTTED]: 'dotted',
    [BorderStyleTypes.DASHED]: 'dashed',
    [BorderStyleTypes.DASH_DOT]: 'dashDot',
    [BorderStyleTypes.DASH_DOT_DOT]: 'dashDotDot',
    [BorderStyleTypes.DOUBLE]: 'double',
    [BorderStyleTypes.MEDIUM]: 'medium',
    [BorderStyleTypes.MEDIUM_DASHED]: 'mediumDashed',
    [BorderStyleTypes.MEDIUM_DASH_DOT]: 'mediumDashDot',
    [BorderStyleTypes.MEDIUM_DASH_DOT_DOT]: 'mediumDashDotDot',
    [BorderStyleTypes.SLANT_DASH_DOT]: 'slantDashDot',
    [BorderStyleTypes.THICK]: 'thick',
};

const HALIGN_REVERSE: Record<number, ExcelJS.Alignment['horizontal']> = {
    [HorizontalAlign.LEFT]: 'left',
    [HorizontalAlign.CENTER]: 'center',
    [HorizontalAlign.RIGHT]: 'right',
    [HorizontalAlign.JUSTIFIED]: 'justify',
};

const VALIGN_REVERSE: Record<number, ExcelJS.Alignment['vertical']> = {
    [VerticalAlign.TOP]: 'top',
    [VerticalAlign.MIDDLE]: 'middle',
    [VerticalAlign.BOTTOM]: 'bottom',
};

function colorToExcel(rgb?: string | null): { argb: string } | undefined {
    if (!rgb) return undefined;
    // Univer 可能给 rgb(r,g,b) 形态，仅处理 hex；其余忽略
    const argb = hexToArgb(rgb);
    return argb ? { argb } : undefined;
}

type BorderSideLike = { s: BorderStyleTypes; cl?: { rgb?: string | null } | null } | null | undefined | void;

function borderSideToExcel(side: BorderSideLike): Partial<ExcelJS.Border> | undefined {
    if (!side) return undefined;
    const style = BORDER_STYLE_REVERSE[side.s];
    if (!style) return undefined;
    return { style, color: colorToExcel(side.cl?.rgb) };
}

/** Univer 的 Nullable<T> 含 void，属性访问不友好；统一的宽松视图 */
interface LooseStyle {
    ff?: string | null;
    fs?: number;
    bl?: number;
    it?: number;
    st?: { s?: number } | null;
    ul?: { s?: number } | null;
    cl?: { rgb?: string | null } | null;
    bg?: { rgb?: string | null } | null;
    bd?: Record<string, BorderSideLike> | null;
    ht?: number;
    vt?: number;
    tb?: number;
    tr?: { a?: number; v?: number } | null;
    n?: { pattern?: string } | null;
}

function bordersToExcel(bdRaw?: IBorderData | null): Partial<ExcelJS.Borders> | undefined {
    if (!bdRaw) return undefined;
    const bd = bdRaw as Record<string, BorderSideLike>;
    const out: Partial<ExcelJS.Borders> = {};
    const t = borderSideToExcel(bd.t);
    const r = borderSideToExcel(bd.r);
    const b = borderSideToExcel(bd.b);
    const l = borderSideToExcel(bd.l);
    if (t) out.top = t;
    if (r) out.right = r;
    if (b) out.bottom = b;
    if (l) out.left = l;
    const down = borderSideToExcel(bd.tl_br);
    const up = borderSideToExcel(bd.bl_tr);
    if (down || up) {
        out.diagonal = { ...(down ?? up), up: !!up, down: !!down } as ExcelJS.Border;
    }
    return Object.keys(out).length ? out : undefined;
}

/** 把 IStyleData 应用到 ExcelJS Cell（覆盖式；调用方决定是否需要写） */
export function applyUniverStyleToCell(cell: ExcelJS.Cell, styleRaw: IStyleData | null | undefined) {
    if (!styleRaw) {
        cell.style = {};
        return;
    }
    const style = styleRaw as LooseStyle;

    const font: Partial<ExcelJS.Font> = {};
    if (style.ff) font.name = style.ff;
    if (style.fs) font.size = style.fs;
    if (style.bl != null) font.bold = style.bl === 1;
    if (style.it != null) font.italic = style.it === 1;
    if (style.st?.s === 1) font.strike = true;
    if (style.ul?.s === 1) font.underline = true;
    const fontColor = colorToExcel(style.cl?.rgb);
    if (fontColor) font.color = fontColor;
    if (Object.keys(font).length) cell.font = font;

    const alignment: Partial<ExcelJS.Alignment> = {};
    if (style.ht != null && HALIGN_REVERSE[style.ht]) alignment.horizontal = HALIGN_REVERSE[style.ht];
    if (style.vt != null && VALIGN_REVERSE[style.vt]) alignment.vertical = VALIGN_REVERSE[style.vt];
    if (style.tb === WrapStrategy.WRAP) alignment.wrapText = true;
    if (style.tr?.a) {
        alignment.textRotation = style.tr.v === 1 ? 'vertical' : (style.tr.a < 0 ? 90 - style.tr.a : style.tr.a) as ExcelJS.Alignment['textRotation'];
    }
    if (Object.keys(alignment).length) cell.alignment = alignment;

    const bg = colorToExcel(style.bg?.rgb);
    if (bg) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: bg };
    }

    const borders = bordersToExcel((styleRaw.bd ?? null) as IBorderData | null);
    if (borders) cell.border = borders;

    if (style.n?.pattern) cell.numFmt = style.n.pattern;
}

/** IStyleData 字体片段 → ExcelJS richText 字体（富文本 run 用） */
export function univerRunStyleToExcelFont(tsRaw?: IStyleData): Partial<ExcelJS.Font> | undefined {
    if (!tsRaw) return undefined;
    const ts = tsRaw as LooseStyle;
    const font: Partial<ExcelJS.Font> = {};
    if (ts.ff) font.name = ts.ff;
    if (ts.fs) font.size = ts.fs;
    if (ts.bl === 1) font.bold = true;
    if (ts.it === 1) font.italic = true;
    if (ts.st?.s === 1) font.strike = true;
    if (ts.ul?.s === 1) font.underline = true;
    const color = colorToExcel(ts.cl?.rgb);
    if (color) font.color = color;
    return Object.keys(font).length ? font : undefined;
}
