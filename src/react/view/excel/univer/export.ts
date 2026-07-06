/**
 * Univer 引擎的保存/另存出口。
 *
 * xlsx/xlsm：diff(baseline, current) → 增量应用到保留的原始 workbook →
 * writeBuffer（增量导出：未编辑的特性原样保留）。
 * xls/ods：snapshot → SheetJS；csv/tsv：snapshot → sheet_to_csv（复用
 * 现有编码/分隔符处理）。emit 协议与 excel_writer 完全一致。
 */
import ExcelJS from '@cweijan/exceljs';
import * as XLSX from 'xlsx';
import type { ICellData, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core';
import { handler } from '../../../util/vscode';
import { CsvEncoding, encodeCsvText } from '../csvEncoding';
import { diffWorkbook } from './diff';
import { applyDiffToWorkbook, excelSerialToDate, writeFullSheet } from './apply';
import type { UniverLoadResult } from './loader';

export interface UniverSaveContext {
    /** 当前文档扩展名（无点） */
    ext: string;
    loadResult: UniverLoadResult;
    baseline: IWorkbookData;
    structuralSheetIds: Set<string>;
    csvEncoding: CsvEncoding;
    csvDelimiter: string;
}

export interface UniverSaveOptions {
    saveAs?: boolean;
    saveAsExt?: string;
}

const emitBytes = (bytes: Uint8Array, options?: UniverSaveOptions) => {
    const content = [...bytes];
    if (options?.saveAs) {
        handler.emit('saveAs', { content, ext: options.saveAsExt ?? 'xlsx' });
    } else {
        handler.emit('save', content);
    }
};

const resolveStyle = (wb: IWorkbookData, s: ICellData['s']): IStyleData | null => {
    if (s == null) return null;
    if (typeof s === 'string') return (wb.styles?.[s] as IStyleData) ?? null;
    return s as IStyleData;
};

const isDateLikePattern = (pattern?: string | null): boolean => {
    if (!pattern) return false;
    const stripped = pattern.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '');
    return /[ymdhs]/i.test(stripped) && !/[#0?]/.test(stripped);
};

async function exportXlsxIncremental(
    current: IWorkbookData,
    context: UniverSaveContext,
    options?: UniverSaveOptions,
): Promise<Record<string, number>> {
    const { originalWorkbook, sheetIdMap } = context.loadResult;
    if (!originalWorkbook || !sheetIdMap) {
        // 非 xlsx 来源（csv/xls/ods 另存 xlsx）：从 snapshot 全量构建
        const workbook = new ExcelJS.Workbook();
        const newMap: Record<string, number> = {};
        for (const sheetId of current.sheetOrder) {
            const sheetData = current.sheets[sheetId];
            const ws = workbook.addWorksheet(sheetData.name || sheetId);
            writeFullSheet(ws, current, sheetData);
            newMap[sheetId] = ws.id;
        }
        const buffer = await workbook.xlsx.writeBuffer();
        emitBytes(new Uint8Array(buffer), options);
        return newMap;
    }

    const diff = diffWorkbook(context.baseline, current, context.structuralSheetIds);
    const newMap = applyDiffToWorkbook(originalWorkbook, current, diff, { sheetIdMap });
    const buffer = await originalWorkbook.xlsx.writeBuffer();
    emitBytes(new Uint8Array(buffer), options);
    return newMap;
}

/** snapshot → SheetJS worksheet（值/公式/合并/列宽，与 legacy 精度一致） */
function univerSheetToSheetJs(wb: IWorkbookData, sheet: Partial<IWorksheetData>): XLSX.WorkSheet {
    const aoa: (string | number | boolean | Date | null)[][] = [];
    for (const [rKey, row] of Object.entries(sheet.cellData ?? {})) {
        const r = Number(rKey);
        aoa[r] ??= [];
        for (const [cKey, cell] of Object.entries(row as Record<number, ICellData>)) {
            const c = Number(cKey);
            const data = cell as ICellData;
            let v: string | number | boolean | Date | null = (data.v as string | number | boolean | undefined) ?? null;
            const style = resolveStyle(wb, data.s);
            if (typeof v === 'number' && isDateLikePattern((style?.n as { pattern?: string } | null | undefined)?.pattern)) {
                v = excelSerialToDate(v);
            }
            aoa[r][c] = v;
        }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    const merges = (sheet.mergeData ?? []).map(m => ({
        s: { r: m.startRow, c: m.startColumn },
        e: { r: m.endRow, c: m.endColumn },
    }));
    if (merges.length) ws['!merges'] = merges;

    const colData = (sheet.columnData ?? {}) as Record<number, { w?: number }>;
    const colKeys = Object.keys(colData).map(Number);
    if (colKeys.length) {
        const cols: XLSX.ColInfo[] = [];
        for (const key of colKeys) {
            if (colData[key]?.w != null) cols[key] = { wpx: colData[key].w };
        }
        ws['!cols'] = cols;
    }
    return ws;
}

function exportSheetJs(current: IWorkbookData, bookType: XLSX.BookType, options?: UniverSaveOptions) {
    const workbook = XLSX.utils.book_new();
    current.sheetOrder.forEach((sheetId, i) => {
        const sheet = current.sheets[sheetId];
        XLSX.utils.book_append_sheet(workbook, univerSheetToSheetJs(current, sheet), sheet.name || `Sheet${i + 1}`);
    });
    const buffer = XLSX.write(workbook, { bookType, type: 'array' });
    emitBytes(new Uint8Array(buffer as ArrayBuffer), options);
}

function exportCsv(current: IWorkbookData, context: UniverSaveContext, fs: string, options?: UniverSaveOptions) {
    const first = current.sheets[current.sheetOrder[0]];
    const csvContent = XLSX.utils.sheet_to_csv(univerSheetToSheetJs(current, first), { FS: fs });
    const bytes = encodeCsvText(csvContent, context.csvEncoding);
    emitBytes(bytes, options);
}

/**
 * 保存当前 workbook。返回值供调用方更新会话状态：
 * baseline 应更新为本次的 current，结构日志应清空，sheetIdMap 应替换。
 */
export async function saveUniverWorkbook(
    current: IWorkbookData,
    context: UniverSaveContext,
    options?: UniverSaveOptions,
): Promise<{ newSheetIdMap?: Record<string, number> }> {
    const ext = (options?.saveAs ? options.saveAsExt ?? 'xlsx' : context.ext).replace(/^\./, '').toLowerCase();

    if (ext === 'xlsx' || ext === 'xlsm') {
        const newSheetIdMap = await exportXlsxIncremental(current, context, options);
        return { newSheetIdMap };
    }
    if (ext === 'xls' || ext === 'ods') {
        exportSheetJs(current, ext as XLSX.BookType, options);
        return {};
    }
    if (ext === 'csv' || ext === 'tsv') {
        exportCsv(current, context, ext === 'tsv' ? '\t' : context.csvDelimiter, options);
        return {};
    }
    // 未知扩展名兜底走 xlsx
    const newSheetIdMap = await exportXlsxIncremental(current, context, { ...options, saveAs: options?.saveAs, saveAsExt: 'xlsx' });
    return { newSheetIdMap };
}

export { hasFormattingChangedUniver } from './diff';
