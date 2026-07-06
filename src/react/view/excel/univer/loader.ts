/**
 * Univer 引擎的文件加载入口：字节 + 扩展名 → IWorkbookData。
 * xlsx/xlsm 走 ExcelJS（先经 sanitize 预清洗）；xls/ods 走 SheetJS；
 * csv/tsv 走 udsv（复用现有 GBK 编码检测）。
 */
import ExcelJS from '@cweijan/exceljs';
import * as XLSX from 'xlsx';
import { inferSchema, initParser } from 'udsv';
import type { ICellData, IWorkbookData, IWorksheetData } from '@univerjs/core';
import { decodeCsvBuffer } from '../csvEncoding';
import { convertExcelJsToUniver, dateToExcelSerial, type UniverImportResult } from './import';
import { sanitizeXlsxForExcelJs, type SanitizeResult } from './sanitize';

export interface UniverLoadResult {
    workbookData: IWorkbookData;
    /** xlsx/xlsm 才有：原始 workbook（M2 增量导出基底）与超链接 */
    originalWorkbook?: UniverImportResult['originalWorkbook'];
    hyperlinks: UniverImportResult['hyperlinks'];
    /** xlsx/xlsm 才有：Univer sheetId → ExcelJS worksheet.id */
    sheetIdMap?: UniverImportResult['sheetIdMap'];
    /** xlsx/xlsm 才有：每 sheet 的 DV/图片/保护 */
    sheetFeatures?: UniverImportResult['sheetFeatures'];
    /** 打开时探测到的 ExcelJS 无法承载的特性（保存前警告用） */
    lossy?: SanitizeResult['lossy'];
    csvDelimiter?: string;
}

const isCsvExt = (ext: string) => /csv|tsv/.test(ext.toLowerCase());
const isSheetJsExt = (ext: string) => {
    const e = ext.toLowerCase().replace(/^\./, '');
    return e === 'xls' || e.includes('ods');
};

const baseWorkbookData = (name: string, sheets: Record<string, Partial<IWorksheetData>>, sheetOrder: string[]): IWorkbookData => ({
    id: 'workbook-1',
    name,
    appVersion: '',
    locale: 'zhCN',
    styles: {},
    sheetOrder,
    sheets,
    resources: [],
}) as unknown as IWorkbookData;

async function loadXlsx(buffer: ArrayBuffer, name: string): Promise<UniverLoadResult> {
    const sanitized = await sanitizeXlsxForExcelJs(buffer);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(sanitized.buffer);
    const result = convertExcelJsToUniver(workbook, name);
    return {
        workbookData: result.workbookData,
        originalWorkbook: result.originalWorkbook,
        hyperlinks: result.hyperlinks,
        sheetIdMap: result.sheetIdMap,
        sheetFeatures: result.sheetFeatures,
        lossy: sanitized.lossy,
    };
}

function sheetJsCellToUniver(cell: XLSX.CellObject): ICellData | null {
    if (cell.v == null) return null;
    const data: ICellData = {};
    if (cell.f) data.f = `=${cell.f}`;
    if (cell.v instanceof Date) {
        data.v = dateToExcelSerial(cell.v);
        data.s = { n: { pattern: 'yyyy-mm-dd' } } as never;
    } else if (typeof cell.v === 'number' || typeof cell.v === 'boolean') {
        data.v = cell.v;
    } else {
        data.v = String(cell.v);
    }
    return data;
}

function loadWithSheetJs(buffer: ArrayBuffer, name: string): UniverLoadResult {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheets: Record<string, Partial<IWorksheetData>> = {};
    const sheetOrder: string[] = [];

    workbook.SheetNames.forEach((sheetName, i) => {
        const ws = workbook.Sheets[sheetName];
        const sheetId = `sheet-${i + 1}`;
        const cellData: Record<number, Record<number, ICellData>> = {};
        const mergeData: { startRow: number; startColumn: number; endRow: number; endColumn: number }[] = [];
        let rowCount = 0;
        let colCount = 0;

        const ref = ws['!ref'];
        if (ref) {
            const range = XLSX.utils.decode_range(ref);
            for (let ri = range.s.r; ri <= range.e.r; ri += 1) {
                for (let ci = range.s.c; ci <= range.e.c; ci += 1) {
                    const cell = ws[XLSX.utils.encode_cell({ r: ri, c: ci })];
                    if (!cell) continue;
                    const data = sheetJsCellToUniver(cell);
                    if (!data) continue;
                    (cellData[ri] ??= {})[ci] = data;
                    if (ri + 1 > rowCount) rowCount = ri + 1;
                    if (ci + 1 > colCount) colCount = ci + 1;
                }
            }
        }

        for (const merge of ws['!merges'] ?? []) {
            mergeData.push({
                startRow: merge.s.r,
                startColumn: merge.s.c,
                endRow: merge.e.r,
                endColumn: merge.e.c,
            });
        }

        sheets[sheetId] = {
            id: sheetId,
            name: sheetName,
            rowCount: Math.max(rowCount, 40),
            columnCount: Math.max(colCount, 26),
            cellData,
            mergeData,
        };
        sheetOrder.push(sheetId);
    });

    return { workbookData: baseWorkbookData(name, sheets, sheetOrder), hyperlinks: [] };
}

function loadCsv(buffer: ArrayBuffer, name: string): UniverLoadResult {
    let csvStr = decodeCsvBuffer(buffer);
    const empty: UniverLoadResult = {
        workbookData: baseWorkbookData(name, { 'sheet-1': { id: 'sheet-1', name: 'Sheet1', rowCount: 40, columnCount: 26, cellData: {} } }, ['sheet-1']),
        hyperlinks: [],
    };
    if (!csvStr) return empty;

    if (!csvStr.includes('\n')) csvStr += '\n';
    const schema = inferSchema(csvStr, { header: () => [] });
    const rows: string[][] = initParser(schema).stringArrs(csvStr);

    const cellData: Record<number, Record<number, ICellData>> = {};
    let colCount = 0;
    rows.forEach((row, ri) => {
        row.forEach((text, ci) => {
            if (text == null || text === '') return;
            const num = Number(text);
            // 与旧 CSV 行为一致的宽松数字识别：纯数字文本按数字显示
            const v = text.trim() !== '' && !Number.isNaN(num) && /^-?\d+(\.\d+)?$/.test(text.trim()) ? num : String(text);
            (cellData[ri] ??= {})[ci] = { v };
            if (ci + 1 > colCount) colCount = ci + 1;
        });
    });

    return {
        workbookData: baseWorkbookData(name, {
            'sheet-1': {
                id: 'sheet-1',
                name: 'Sheet1',
                rowCount: Math.max(rows.length, 40),
                columnCount: Math.max(colCount, 26),
                cellData,
            },
        }, ['sheet-1']),
        hyperlinks: [],
        csvDelimiter: schema.col,
    };
}

export async function loadForUniver(buffer: ArrayBuffer, ext: string, fileName: string): Promise<UniverLoadResult> {
    if (isCsvExt(ext)) return loadCsv(buffer, fileName);
    if (isSheetJsExt(ext)) return loadWithSheetJs(buffer, fileName);
    return loadXlsx(buffer, fileName);
}
