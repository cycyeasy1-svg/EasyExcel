/**
 * WorkbookDiff → 保留的原始 ExcelJS workbook 增量应用。
 *
 * - incremental sheet：只触碰 diff 出的 cell/merge/行列尺寸/冻结/名字，
 *   其余内容（批注、条件格式、DV、原始 numFmt、defined names……）不经手，
 *   ExcelJS 原样往返（P0 R1 已验证无图表文件零丢失）。
 * - added / rebuilt sheet：从当前 snapshot 整表写入（rebuilt 会丢失该
 *   sheet 上 Univer 未建模的特性——结构性编辑的已知代价，M4 优化）。
 */
import type * as ExcelJS from '@cweijan/exceljs';
import type { ICellData, IRange, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core';
import { applyUniverStyleToCell, univerRunStyleToExcelFont } from './export_styles';
import type { CellChange, SheetDiff, WorkbookDiff } from './diff';

const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function colName(index: number): string {
    let str = '';
    let i = index;
    while (i >= 0) {
        str = COL_LETTERS[i % 26] + str;
        i = Math.floor(i / 26) - 1;
    }
    return str;
}

const rangeToAddr = (r: IRange) =>
    `${colName(r.startColumn)}${r.startRow + 1}:${colName(r.endColumn)}${r.endRow + 1}`;

const pxToExcelColWidth = (px: number) => Math.max((px - 5) / 7, 0);
const pxToExcelRowHeightPt = (px: number) => px * 72 / 96;

/** 日期类 numfmt 判定（写回 Date 对象以保持单元格日期类型） */
const isDateLikePattern = (pattern?: string | null): boolean => {
    if (!pattern) return false;
    const stripped = pattern
        .replace(/\[[^\]]*\]/g, '')
        .replace(/"[^"]*"/g, '');
    return /[ymdhs]/i.test(stripped) && !/[#0?]/.test(stripped);
};

export const excelSerialToDate = (serial: number): Date =>
    new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));

function richTextToExcelValue(p: ICellData['p']): ExcelJS.CellRichTextValue | null {
    const body = (p as { body?: { dataStream?: string; textRuns?: { st: number; ed: number; ts?: IStyleData }[] } })?.body;
    if (!body?.dataStream) return null;
    const text = body.dataStream.replace(/\r?\n$/, '').replace(/\r\n$/, '');
    const runs = [...(body.textRuns ?? [])].sort((a, b) => a.st - b.st);
    const richText: ExcelJS.RichText[] = [];
    let pos = 0;
    for (const run of runs) {
        const st = Math.max(run.st, pos);
        const ed = Math.min(run.ed, text.length);
        if (st > pos) richText.push({ text: text.slice(pos, st) });
        if (ed > st) {
            const font = univerRunStyleToExcelFont(run.ts);
            richText.push({ text: text.slice(st, ed), ...(font ? { font } : {}) });
        }
        pos = Math.max(pos, ed);
    }
    if (pos < text.length) richText.push({ text: text.slice(pos) });
    if (!richText.length) return null;
    return { richText };
}

function writeCellValue(cell: ExcelJS.Cell, data: ICellData, style: IStyleData | null) {
    if (data.f) {
        const formula = data.f.replace(/^=/, '');
        const result = data.v;
        cell.value = {
            formula,
            ...(result != null ? { result } : {}),
        } as ExcelJS.CellFormulaValue;
        return;
    }
    if (data.p) {
        const rich = richTextToExcelValue(data.p);
        if (rich) {
            cell.value = rich;
            return;
        }
    }
    if (data.v == null) {
        cell.value = null;
        return;
    }
    if (typeof data.v === 'number' && isDateLikePattern((style?.n as { pattern?: string } | null | undefined)?.pattern)) {
        cell.value = excelSerialToDate(data.v);
        return;
    }
    cell.value = data.v as ExcelJS.CellValue;
}

function applyCellChange(worksheet: ExcelJS.Worksheet, change: CellChange) {
    const cell = worksheet.getCell(change.row + 1, change.col + 1);
    if (change.cell === null) {
        cell.value = null;
        cell.style = {};
        return;
    }
    if (change.valueChanged) {
        writeCellValue(cell, change.cell, change.style);
    }
    if (change.styleChanged) {
        applyUniverStyleToCell(cell, change.style);
    }
}

function applyFreeze(worksheet: ExcelJS.Worksheet, freeze: SheetDiff['freeze']) {
    if (freeze && (freeze.xSplit > 0 || freeze.ySplit > 0)) {
        worksheet.views = [{
            state: 'frozen',
            xSplit: freeze.xSplit,
            ySplit: freeze.ySplit,
            topLeftCell: `${colName(freeze.xSplit)}${freeze.ySplit + 1}`,
        }];
    } else {
        worksheet.views = (worksheet.views ?? []).filter(v => v.state !== 'frozen');
        if (!worksheet.views.length) worksheet.views = [{}] as ExcelJS.WorksheetView[];
    }
}

function applyIncremental(worksheet: ExcelJS.Worksheet, diff: SheetDiff) {
    if (diff.renamed) worksheet.name = diff.name;

    for (const merge of diff.mergesRemoved) {
        try {
            worksheet.unMergeCells(rangeToAddr(merge));
        } catch {
            // 原文件中不存在该合并（可能被 rebuild 或异常）：忽略
        }
    }
    for (const merge of diff.mergesAdded) {
        try {
            worksheet.mergeCells(rangeToAddr(merge));
        } catch {
            // 与既有合并冲突时忽略，避免整个保存失败
        }
    }

    for (const change of diff.cellChanges) {
        applyCellChange(worksheet, change);
    }

    for (const rc of diff.rowChanges) {
        const row = worksheet.getRow(rc.index + 1);
        if ('h' in rc) row.height = rc.h != null ? pxToExcelRowHeightPt(rc.h) : undefined;
        if ('hidden' in rc) row.hidden = !!rc.hidden;
    }
    for (const cc of diff.colChanges) {
        const col = worksheet.getColumn(cc.index + 1);
        if ('w' in cc) col.width = cc.w != null ? pxToExcelColWidth(cc.w) : undefined;
        if ('hidden' in cc) col.hidden = !!cc.hidden;
    }

    if (diff.freezeChanged) applyFreeze(worksheet, diff.freeze);
}

/** 整表写入（added / rebuilt sheet） */
export function writeFullSheet(
    worksheet: ExcelJS.Worksheet,
    workbookData: IWorkbookData,
    sheetData: Partial<IWorksheetData>,
) {
    const resolve = (s: ICellData['s']): IStyleData | null => {
        if (s == null) return null;
        if (typeof s === 'string') return (workbookData.styles?.[s] as IStyleData) ?? null;
        return s as IStyleData;
    };

    for (const [rKey, row] of Object.entries(sheetData.cellData ?? {})) {
        const r = Number(rKey);
        for (const [cKey, data] of Object.entries(row as Record<number, ICellData>)) {
            const c = Number(cKey);
            const style = resolve((data as ICellData).s);
            const cell = worksheet.getCell(r + 1, c + 1);
            writeCellValue(cell, data as ICellData, style);
            if (style) applyUniverStyleToCell(cell, style);
        }
    }

    for (const merge of sheetData.mergeData ?? []) {
        try {
            worksheet.mergeCells(rangeToAddr(merge as IRange));
        } catch {
            // 忽略非法/冲突合并
        }
    }

    for (const [key, rd] of Object.entries((sheetData.rowData ?? {}) as Record<number, { h?: number; hd?: 0 | 1 }>)) {
        const row = worksheet.getRow(Number(key) + 1);
        if (rd.h != null) row.height = pxToExcelRowHeightPt(rd.h);
        if (rd.hd === 1) row.hidden = true;
    }
    for (const [key, cd] of Object.entries((sheetData.columnData ?? {}) as Record<number, { w?: number; hd?: 0 | 1 }>)) {
        const col = worksheet.getColumn(Number(key) + 1);
        if (cd.w != null) col.width = pxToExcelColWidth(cd.w);
        if (cd.hd === 1) col.hidden = true;
    }

    if (sheetData.freeze && (sheetData.freeze.xSplit > 0 || sheetData.freeze.ySplit > 0)) {
        applyFreeze(worksheet, sheetData.freeze);
    }
}

export interface ApplyContext {
    /** 导入时建立的 Univer sheetId → ExcelJS worksheet.id 映射 */
    sheetIdMap: Record<string, number>;
}

/** 把 diff 应用到原始 workbook；返回更新后的 sheetIdMap（新增/重建的 sheet 换了 id） */
export function applyDiffToWorkbook(
    workbook: ExcelJS.Workbook,
    current: IWorkbookData,
    diff: WorkbookDiff,
    context: ApplyContext,
): Record<string, number> {
    const sheetIdMap = { ...context.sheetIdMap };

    for (const removedId of diff.removedSheetIds) {
        const excelId = sheetIdMap[removedId];
        if (excelId != null && workbook.getWorksheet(excelId)) {
            workbook.removeWorksheet(excelId);
        }
        delete sheetIdMap[removedId];
    }

    for (const sheetDiff of diff.sheets) {
        const sheetData = current.sheets[sheetDiff.sheetId];
        if (sheetDiff.status === 'added') {
            const ws = workbook.addWorksheet(sheetDiff.name || sheetDiff.sheetId);
            writeFullSheet(ws, current, sheetData);
            sheetIdMap[sheetDiff.sheetId] = ws.id;
            continue;
        }

        const excelId = sheetIdMap[sheetDiff.sheetId];
        const worksheet = excelId != null ? workbook.getWorksheet(excelId) : undefined;
        if (!worksheet) {
            // 映射丢失（异常情况）：按新增处理，保证不丢数据
            const ws = workbook.addWorksheet(sheetDiff.name || sheetDiff.sheetId);
            writeFullSheet(ws, current, sheetData);
            sheetIdMap[sheetDiff.sheetId] = ws.id;
            continue;
        }

        if (sheetDiff.status === 'rebuilt') {
            workbook.removeWorksheet(worksheet.id);
            const ws = workbook.addWorksheet(sheetDiff.name || sheetDiff.sheetId);
            writeFullSheet(ws, current, sheetData);
            sheetIdMap[sheetDiff.sheetId] = ws.id;
            continue;
        }

        applyIncremental(worksheet, sheetDiff);
    }

    // sheet 顺序：orderNo 决定写出顺序
    current.sheetOrder.forEach((sheetId, index) => {
        const excelId = sheetIdMap[sheetId];
        const ws = excelId != null ? workbook.getWorksheet(excelId) : undefined;
        if (ws) (ws as ExcelJS.Worksheet & { orderNo: number }).orderNo = index;
    });

    return sheetIdMap;
}
