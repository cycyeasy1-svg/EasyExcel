/**
 * ExcelJS Workbook → Univer IWorkbookData 转换层（M1：读取方向）。
 *
 * 与旧 excel_reader 的差异：
 * - numfmt 原样直通（不再压缩为 11 种预设）
 * - 主题色/索引色/tint 经 theme_colors.ts 解析（不再丢弃）
 * - 富文本转 IDocumentData（不再压平）
 * - 数字/布尔/日期保留类型（日期转 Excel 序列值 + numfmt 渲染）
 * - 保留原始 workbook 引用，交给 M2 的增量导出
 */
import type * as ExcelJS from '@cweijan/exceljs';
import {
    BorderStyleTypes,
    HorizontalAlign,
    VerticalAlign,
    WrapStrategy,
    type IBorderData,
    type ICellData,
    type IStyleData,
    type ITextRun,
    type IWorkbookData,
    type IWorksheetData,
} from '@univerjs/core';
import { extractThemeXml, parseThemePalette, resolveExcelColor, type ExcelColorLike } from './theme_colors';
import { readWorksheetValidations, type SpreadsheetValidationItem } from '../excel_validation';
import { readWorksheetImages, type SheetImage } from '../excel_images';
import { isWorksheetProtected } from '../excel_protection';
import { CF_PLUGIN, excelCfToUniver, type UniverCfRule } from './cf';

const DEFAULT_COL_WIDTH_PX = 88;
const DEFAULT_ROW_HEIGHT_PX = 24;
const MIN_ROW_COUNT = 40;
const MIN_COL_COUNT = 26;

export interface UniverImportResult {
    workbookData: IWorkbookData;
    /** 原始 workbook，M2 增量导出的基底 */
    originalWorkbook: ExcelJS.Workbook;
    /** 超链接列表（导入后经 facade setHyperLink 应用） */
    hyperlinks: { sheetId: string; row: number; column: number; url: string; display?: string }[];
    /** Univer sheetId → ExcelJS worksheet.id（增量导出定位用） */
    sheetIdMap: Record<string, number>;
    /** 每 sheet 的附加特性（DV/图片/保护），加载后经 facade 应用 */
    sheetFeatures: Record<string, SheetFeatures>;
}

export interface SheetFeatures {
    validations: SpreadsheetValidationItem[];
    images: SheetImage[];
    protected: boolean;
}

const BORDER_STYLE_MAP: Record<string, BorderStyleTypes> = {
    thin: BorderStyleTypes.THIN,
    hair: BorderStyleTypes.HAIR,
    dotted: BorderStyleTypes.DOTTED,
    dashed: BorderStyleTypes.DASHED,
    dashDot: BorderStyleTypes.DASH_DOT,
    dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
    double: BorderStyleTypes.DOUBLE,
    medium: BorderStyleTypes.MEDIUM,
    mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
    mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
    mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
    slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
    thick: BorderStyleTypes.THICK,
};

const HALIGN_MAP: Record<string, HorizontalAlign> = {
    left: HorizontalAlign.LEFT,
    center: HorizontalAlign.CENTER,
    right: HorizontalAlign.RIGHT,
    justify: HorizontalAlign.JUSTIFIED,
};

const VALIGN_MAP: Record<string, VerticalAlign> = {
    top: VerticalAlign.TOP,
    middle: VerticalAlign.MIDDLE,
    bottom: VerticalAlign.BOTTOM,
};

/** Excel 字符宽 → px，与旧 reader 的 excelColWidthToPx 一致 */
const excelColWidthToPx = (width?: number) =>
    width == null ? undefined : Math.round(width * 7 + 5);

/** Excel 行高（磅）→ px */
const excelRowHeightToPx = (pt: number) => Math.round(pt * 96 / 72);

/** JS Date → Excel 1900 序列值（UTC，ExcelJS 解析日期即 UTC 存储） */
export const dateToExcelSerial = (d: Date): number =>
    (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;

interface FontLike {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    underline?: boolean | string;
    color?: ExcelColorLike;
}

function fontToStyle(font: FontLike | undefined, palette: string[]): IStyleData {
    const style: IStyleData = {};
    if (!font) return style;
    if (font.name) style.ff = font.name;
    if (font.size) style.fs = font.size;
    if (font.bold) style.bl = 1;
    if (font.italic) style.it = 1;
    if (font.strike) style.st = { s: 1 };
    if (font.underline && font.underline !== 'none') style.ul = { s: 1 };
    const cl = resolveExcelColor(font.color, palette);
    if (cl) style.cl = { rgb: cl };
    return style;
}

function excelJsCellToUniverStyle(cell: ExcelJS.Cell, palette: string[]): IStyleData | null {
    const style: IStyleData = fontToStyle(cell.font as FontLike, palette);
    let hasStyle = Object.keys(style).length > 0;

    const alignment = cell.alignment;
    if (alignment) {
        const ht = alignment.horizontal ? HALIGN_MAP[alignment.horizontal] : undefined;
        if (ht) {
            style.ht = ht;
            hasStyle = true;
        }
        const vt = alignment.vertical ? VALIGN_MAP[alignment.vertical] : undefined;
        if (vt) {
            style.vt = vt;
            hasStyle = true;
        }
        if (alignment.wrapText) {
            style.tb = WrapStrategy.WRAP;
            hasStyle = true;
        }
        if (alignment.textRotation != null && alignment.textRotation !== 0) {
            const a = alignment.textRotation === 'vertical' ? 0 : Number(alignment.textRotation);
            if (!Number.isNaN(a)) {
                // OOXML 的 90..180 表示向下 (90-角度)，Univer 用负角
                style.tr = { a: a > 90 ? 90 - a : a, ...(alignment.textRotation === 'vertical' ? { v: 1 as const } : {}) };
                hasStyle = true;
            }
        }
    }

    const fill = cell.fill;
    if (fill?.type === 'pattern') {
        const patternFill = fill as ExcelJS.FillPattern;
        const bg = resolveExcelColor(patternFill.fgColor as ExcelColorLike, palette)
            ?? resolveExcelColor(patternFill.bgColor as ExcelColorLike, palette);
        // 保留显式白色填充（旧 reader 丢弃）；'none' pattern 无填充
        if (bg && patternFill.pattern !== 'none') {
            style.bg = { rgb: bg };
            hasStyle = true;
        }
    } else if (fill?.type === 'gradient') {
        // 渐变降级：取第一个 stop（完整渐变列入远期）
        const stops = (fill as ExcelJS.FillGradientAngle).gradient ? (fill as never as { stops?: { color?: ExcelColorLike }[] }).stops : undefined;
        const bg = stops?.length ? resolveExcelColor(stops[0].color, palette) : undefined;
        if (bg) {
            style.bg = { rgb: bg };
            hasStyle = true;
        }
    }

    const border = bordersToUniver(cell.border, palette);
    if (border) {
        style.bd = border;
        hasStyle = true;
    }

    const numFmt = cell.numFmt;
    if (numFmt && numFmt !== 'General') {
        style.n = { pattern: numFmt };
        hasStyle = true;
    }

    return hasStyle ? style : null;
}

function bordersToUniver(borders: Partial<ExcelJS.Borders> | undefined, palette: string[]): IBorderData | undefined {
    if (!borders) return undefined;
    const out: IBorderData = {};
    const sides: [keyof IBorderData, Partial<ExcelJS.Border> | undefined][] = [
        ['t', borders.top],
        ['r', borders.right],
        ['b', borders.bottom],
        ['l', borders.left],
    ];
    let has = false;
    for (const [key, side] of sides) {
        if (!side?.style) continue;
        const s = BORDER_STYLE_MAP[side.style] ?? BorderStyleTypes.THIN;
        const rgb = resolveExcelColor(side.color as ExcelColorLike, palette) ?? '#000000';
        out[key] = { s, cl: { rgb } };
        has = true;
    }
    if (borders.diagonal?.style) {
        const s = BORDER_STYLE_MAP[borders.diagonal.style] ?? BorderStyleTypes.THIN;
        const rgb = resolveExcelColor(borders.diagonal.color as ExcelColorLike, palette) ?? '#000000';
        const diag = borders.diagonal as { up?: boolean; down?: boolean };
        if (diag.down) out.tl_br = { s, cl: { rgb } };
        if (diag.up) out.bl_tr = { s, cl: { rgb } };
        has = has || !!(diag.up || diag.down);
    }
    return has ? out : undefined;
}

interface RichTextRun {
    text: string;
    font?: FontLike;
}

/** 富文本 → 最小合法 IDocumentData（单段落，逐 run 的 textRuns） */
function richTextToDocumentData(runs: RichTextRun[], palette: string[]) {
    let text = '';
    const textRuns: ITextRun[] = [];
    for (const run of runs) {
        const st = text.length;
        text += run.text ?? '';
        const ed = text.length;
        if (ed <= st) continue;
        const ts = fontToStyle(run.font, palette);
        if (Object.keys(ts).length) {
            textRuns.push({ st, ed, ts });
        }
    }
    return {
        id: '__richText',
        body: {
            dataStream: `${text}\r\n`,
            textRuns,
            paragraphs: [{ startIndex: text.length }],
            sectionBreaks: [{ startIndex: text.length + 1 }],
        },
        documentStyle: {},
    };
}

/** 单元格值转换：保留数字/布尔类型；日期转序列值；公式走 f 字段 */
function convertCellValue(cell: ExcelJS.Cell, cellData: ICellData, palette: string[]) {
    const value = cell.value;
    if (value == null) return;

    if (cell.formula || (typeof value === 'object' && 'formula' in (value as object))) {
        const formula = cell.formula || (value as { formula?: string }).formula;
        if (formula) {
            cellData.f = `=${formula}`;
            const result = (value as { result?: unknown })?.result;
            if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') {
                cellData.v = result as never;
            }
            return;
        }
    }

    if (typeof value === 'object' && 'richText' in (value as object)) {
        const runs = (value as { richText: RichTextRun[] }).richText;
        cellData.p = richTextToDocumentData(runs, palette) as never;
        cellData.v = runs.map(r => r.text ?? '').join('');
        return;
    }

    if (typeof value === 'object' && 'hyperlink' in (value as object)) {
        const hv = value as ExcelJS.CellHyperlinkValue;
        cellData.v = hv.text || hv.hyperlink || '';
        return;
    }

    if (value instanceof Date) {
        cellData.v = dateToExcelSerial(value);
        return;
    }

    if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') {
        cellData.v = value;
        return;
    }

    if (typeof value === 'object' && 'error' in (value as object)) {
        cellData.v = String((value as { error: string }).error);
        return;
    }

    cellData.v = cell.text ?? String(value);
}

class UniverStyleRegistry {
    private styles: Record<string, IStyleData> = {};
    private index = new Map<string, string>();
    private seq = 0;

    add(style: IStyleData | null): string | undefined {
        if (!style) return undefined;
        const key = JSON.stringify(style);
        const existing = this.index.get(key);
        if (existing) return existing;
        this.seq += 1;
        const id = `s${this.seq}`;
        this.index.set(key, id);
        this.styles[id] = style;
        return id;
    }

    getStyles() {
        return this.styles;
    }
}

function readFreeze(worksheet: ExcelJS.Worksheet) {
    const views = worksheet.views;
    if (!views?.length) return undefined;
    for (const view of views) {
        if (view.state === 'frozen') {
            const xSplit = view.xSplit ?? 0;
            const ySplit = view.ySplit ?? 0;
            if (xSplit <= 0 && ySplit <= 0) return undefined;
            return { xSplit, ySplit, startRow: ySplit, startColumn: xSplit };
        }
    }
    return undefined;
}

function mergeToRange(master: ExcelJS.Cell, cell: ExcelJS.Cell) {
    return {
        startRow: Number(master.row) - 1,
        startColumn: Number(master.col) - 1,
        endRow: Number(cell.row) - 1,
        endColumn: Number(cell.col) - 1,
    };
}

function convertWorksheet(
    worksheet: ExcelJS.Worksheet,
    sheetId: string,
    palette: string[],
    styleRegistry: UniverStyleRegistry,
    hyperlinks: UniverImportResult['hyperlinks'],
): Partial<IWorksheetData> {
    const cellData: Record<number, Record<number, ICellData>> = {};
    const rowData: Record<number, { h?: number; ah?: number; ia?: 0 | 1; hd?: 0 | 1 }> = {};
    const columnData: Record<number, { w?: number; hd?: 0 | 1 }> = {};
    // merge 范围先收集（key 去重），master 单元格遍历时会多次回调
    const mergeMap = new Map<string, ReturnType<typeof mergeToRange>>();
    let maxRow = 0;
    let maxCol = 0;

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const ri = rowNumber - 1;
        if (row.height != null) {
            rowData[ri] = { h: excelRowHeightToPx(row.height), ia: 0 };
        }
        if (row.hidden) {
            rowData[ri] = { ...(rowData[ri] ?? {}), hd: 1 };
        }
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const ci = colNumber - 1;
            if (cell.isMerged) {
                const master = cell.master;
                const range = mergeMap.get(master.address) ?? mergeToRange(master, master);
                range.endRow = Math.max(range.endRow, ri);
                range.endColumn = Math.max(range.endColumn, ci);
                mergeMap.set(master.address, range);
                if (cell.address !== master.address) return;
            }

            const data: ICellData = {};
            convertCellValue(cell, data, palette);
            const styleId = styleRegistry.add(excelJsCellToUniverStyle(cell, palette));
            if (styleId) data.s = styleId;

            const value = cell.value;
            if (value && typeof value === 'object' && 'hyperlink' in (value as object)) {
                const hv = value as ExcelJS.CellHyperlinkValue;
                if (hv.hyperlink) {
                    hyperlinks.push({ sheetId, row: ri, column: ci, url: hv.hyperlink, display: hv.text });
                }
            }

            if (data.v == null && data.f == null && data.p == null && data.s == null) return;
            (cellData[ri] ??= {})[ci] = data;
            if (ri + 1 > maxRow) maxRow = ri + 1;
            if (ci + 1 > maxCol) maxCol = ci + 1;
        });
    });

    for (const range of mergeMap.values()) {
        if (range.endRow + 1 > maxRow) maxRow = range.endRow + 1;
        if (range.endColumn + 1 > maxCol) maxCol = range.endColumn + 1;
    }

    const colCount = Math.max(maxCol, worksheet.columnCount || 0);
    for (let i = 1; i <= colCount; i += 1) {
        const col = worksheet.getColumn(i);
        const w = excelColWidthToPx(col.width);
        if (w != null || col.hidden) {
            columnData[i - 1] = {
                ...(w != null ? { w } : {}),
                ...(col.hidden ? { hd: 1 as const } : {}),
            };
        }
    }

    const freeze = readFreeze(worksheet);
    const defaultRowHeightPt = (worksheet.properties as { defaultRowHeight?: number })?.defaultRowHeight;

    return {
        id: sheetId,
        name: worksheet.name,
        hidden: worksheet.state && worksheet.state !== 'visible' ? 1 : 0,
        rowCount: Math.max(maxRow, worksheet.rowCount || 0, MIN_ROW_COUNT),
        columnCount: Math.max(colCount, MIN_COL_COUNT),
        defaultColumnWidth: DEFAULT_COL_WIDTH_PX,
        defaultRowHeight: defaultRowHeightPt ? excelRowHeightToPx(defaultRowHeightPt) : DEFAULT_ROW_HEIGHT_PX,
        mergeData: [...mergeMap.values()],
        cellData,
        rowData,
        columnData,
        ...(freeze ? { freeze } : {}),
        ...(worksheet.properties?.tabColor
            ? { tabColor: resolveExcelColor(worksheet.properties.tabColor as ExcelColorLike, palette) }
            : {}),
    } as Partial<IWorksheetData>;
}

export function convertExcelJsToUniver(workbook: ExcelJS.Workbook, name: string): UniverImportResult {
    const themeXml = extractThemeXml((workbook as { model?: unknown }).model);
    const palette = parseThemePalette(themeXml);
    const styleRegistry = new UniverStyleRegistry();
    const hyperlinks: UniverImportResult['hyperlinks'] = [];

    const sheets: Record<string, Partial<IWorksheetData>> = {};
    const sheetOrder: string[] = [];
    const sheetIdMap: Record<string, number> = {};
    const sheetFeatures: Record<string, SheetFeatures> = {};
    const cfMap: Record<string, UniverCfRule[]> = {};
    workbook.worksheets.forEach((worksheet, i) => {
        const sheetId = `sheet-${i + 1}`;
        sheets[sheetId] = convertWorksheet(worksheet, sheetId, palette, styleRegistry, hyperlinks);
        sheetOrder.push(sheetId);
        sheetIdMap[sheetId] = worksheet.id;
        sheetFeatures[sheetId] = {
            validations: safeRead(() => readWorksheetValidations(worksheet), []),
            images: safeRead(() => readWorksheetImages(worksheet, workbook), []),
            protected: safeRead(() => isWorksheetProtected(worksheet), false),
        };
        const cfRules = safeRead(() => excelCfToUniver(worksheet, palette, sheetId), []);
        if (cfRules.length) cfMap[sheetId] = cfRules;
    });

    const resources: { name: string; data: string }[] = [];
    if (Object.keys(cfMap).length) {
        resources.push({ name: CF_PLUGIN, data: JSON.stringify(cfMap) });
    }

    const workbookData = {
        id: 'workbook-1',
        name,
        appVersion: '',
        locale: 'zhCN',
        styles: styleRegistry.getStyles(),
        sheetOrder,
        sheets,
        resources,
    } as unknown as IWorkbookData;

    return { workbookData, originalWorkbook: workbook, hyperlinks, sheetIdMap, sheetFeatures };
}

function safeRead<T>(fn: () => T, fallback: T): T {
    try {
        return fn();
    } catch {
        return fallback;
    }
}
