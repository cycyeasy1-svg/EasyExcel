import type * as ExcelJS from '@cweijan/exceljs';

export type StoredSheetProtection = Record<string, unknown>;

export function readWorksheetProtection(worksheet: ExcelJS.Worksheet): StoredSheetProtection | undefined {
    const ws = worksheet as ExcelJS.Worksheet & { sheetProtection?: StoredSheetProtection };
    const sp = ws.sheetProtection ?? worksheet.model?.sheetProtection;
    if (!sp || typeof sp !== 'object') return undefined;
    return { ...sp };
}

export function isWorksheetProtected(worksheet: ExcelJS.Worksheet): boolean {
    return readWorksheetProtection(worksheet) != null;
}

/** 受保护工作表中：默认锁定；仅 protection.locked === false 的单元格可编辑 */
export function readCellEditableFromExcel(cell: ExcelJS.Cell, sheetProtected: boolean): boolean | undefined {
    if (!sheetProtected) return undefined;
    if (cell.protection?.locked === false) return true;
    return false;
}
