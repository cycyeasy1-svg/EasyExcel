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
import { patchXlsxBytes, XmlPatchBlockedError } from './xml_patch';
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

interface XlsxExportOutcome {
    newSheetIdMap: Record<string, number>;
    richTextDowngraded?: boolean;
}

async function exportXlsxIncremental(
    current: IWorkbookData,
    context: UniverSaveContext,
    options?: UniverSaveOptions,
): Promise<XlsxExportOutcome> {
    const { originalWorkbook, sheetIdMap, originalBuffer } = context.loadResult;
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
        return { newSheetIdMap: newMap };
    }

    const diff = diffWorkbook(context.baseline, current, context.structuralSheetIds);

    // M5：保存回原文件走 XML 补丁 —— 在原始字节上只改被编辑的 cell，
    // 图形/图表/透视表/宏等 ExcelJS 未建模部件物理透传。另存为仍走 ExcelJS。
    const docExt = context.ext.replace(/^\./, '').toLowerCase();
    if (!options?.saveAs && originalBuffer && (docExt === 'xlsx' || docExt === 'xlsm')) {
        let patched: Awaited<ReturnType<typeof patchXlsxBytes>> | null = null;
        try {
            const sheetNames = Object.fromEntries(
                context.baseline.sheetOrder.map(id => [id, context.baseline.sheets[id]?.name ?? '']),
            );
            patched = await patchXlsxBytes(originalBuffer, diff, sheetNames);
        } catch (error) {
            // 结构性/DV/CF 变更：UI 已禁用，走到这里说明拦截失守 → 明确报错，
            // 绝不悄悄降级到有损路径
            if (error instanceof XmlPatchBlockedError) throw error;
            // 补丁自身失败（异常 XML 形态等）：回退 M2 增量导出保住保存能力，
            // 代价是 ExcelJS 未建模部件丢失 —— 大声记录
            console.error('EasyExcel: xml patch failed, falling back to ExcelJS rewrite (unmodeled parts may be lost)', error);
        }
        if (patched) {
            const { bytes, richTextDowngraded } = patched;
            // ExcelJS 模型同步推进，保证后续「另存为」包含本次编辑
            const newMap = applyDiffToWorkbook(originalWorkbook, current, diff, { sheetIdMap });
            context.loadResult.originalBuffer = bytes.buffer.slice(
                bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
            emitBytes(bytes, options);
            return { newSheetIdMap: newMap, richTextDowngraded };
        }
    }

    const newMap = applyDiffToWorkbook(originalWorkbook, current, diff, { sheetIdMap });
    const buffer = await originalWorkbook.xlsx.writeBuffer();
    emitBytes(new Uint8Array(buffer), options);
    return { newSheetIdMap: newMap };
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
 * richTextDowngraded 为 true 时应提示用户（富文本格已按纯文本保存）。
 * 不支持无损保存的变更（结构/DV/CF）抛 XmlPatchBlockedError。
 */
export async function saveUniverWorkbook(
    current: IWorkbookData,
    context: UniverSaveContext,
    options?: UniverSaveOptions,
): Promise<{ newSheetIdMap?: Record<string, number>; richTextDowngraded?: boolean }> {
    const ext = (options?.saveAs ? options.saveAsExt ?? 'xlsx' : context.ext).replace(/^\./, '').toLowerCase();

    if (ext === 'xlsx' || ext === 'xlsm') {
        return exportXlsxIncremental(current, context, options);
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
    return exportXlsxIncremental(current, context, { ...options, saveAs: options?.saveAs, saveAsExt: 'xlsx' });
}

export { hasFormattingChangedUniver } from './diff';
export { XmlPatchBlockedError } from './xml_patch';
