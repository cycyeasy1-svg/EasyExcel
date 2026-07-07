/**
 * IWorkbookData 结构化 diff：baseline（导入完成时）vs current（保存时）。
 *
 * 增量导出的核心：只有 diff 出的变更会写回保留的原始 workbook，未触碰的
 * 单元格连同其未建模特性（批注、条件格式、原始 numFmt、DV……）原样保留。
 *
 * 行列插删（结构性变更）会使几乎所有 cell 位移，逐格 diff 无意义——
 * 由 adapter 的 mutation 日志标记 structuralSheetIds，这些 sheet 走整表
 * 重建（保真损失限定在被结构性编辑过的 sheet；对应计划 R4 兜底方案 B）。
 */
import type { ICellData, IFreeze, IRange, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core';

export interface CellChange {
    row: number;
    col: number;
    /** null = 单元格被清除 */
    cell: ICellData | null;
    /** 解析后的当前样式（cell.s 已解引用）；null 表示清除样式 */
    style: IStyleData | null;
    valueChanged: boolean;
    styleChanged: boolean;
}

export interface SheetDiff {
    sheetId: string;
    status: 'added' | 'rebuilt' | 'incremental';
    name: string;
    renamed: boolean;
    cellChanges: CellChange[];
    mergesAdded: IRange[];
    mergesRemoved: IRange[];
    rowChanges: { index: number; h?: number | null; hidden?: boolean }[];
    colChanges: { index: number; w?: number | null; hidden?: boolean }[];
    freezeChanged: boolean;
    freeze?: IFreeze | null;
}

/** Univer DV 规则（宽松形态；type/operator 字符串与 ExcelJS 同名） */
export interface UniverDvRule {
    uid?: string;
    type?: string;
    ranges?: IRange[];
    formula1?: string;
    formula2?: string;
    operator?: string;
    allowBlank?: boolean;
}

export interface WorkbookDiff {
    sheets: SheetDiff[];
    removedSheetIds: string[];
    orderChanged: boolean;
    isEmpty: boolean;
    /** 当前每 sheet 的 DV 规则（来自 snapshot resources） */
    dvRules: Record<string, UniverDvRule[]>;
    /** DV 相对 baseline 变更过的 sheetId */
    dvChangedSheetIds: string[];
    /** 当前每 sheet 的条件格式规则（来自 snapshot resources，宽松形态） */
    cfRules: Record<string, unknown[]>;
    /** 条件格式相对 baseline 变更过的 sheetId */
    cfChangedSheetIds: string[];
}

const DV_PLUGIN = 'SHEET_DATA_VALIDATION_PLUGIN';
const CF_PLUGIN = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

/** 解析 per-sheet 形态的插件资源（{ subUnitId: rules[] }） */
function parsePerSheetResource<T>(wb: IWorkbookData, pluginName: string): Record<string, T[]> {
    const resources = (wb as { resources?: { name: string; data: string }[] }).resources ?? [];
    const entry = resources.find(r => r.name === pluginName);
    if (!entry?.data) return {};
    try {
        const parsed = JSON.parse(entry.data) as Record<string, T[]>;
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

/** 对比 baseline / current 的 per-sheet 资源，返回变更过的 sheetId */
function diffPerSheetResource<T>(
    base: Record<string, T[]>,
    cur: Record<string, T[]>,
    validSheetIds: string[],
): string[] {
    const changed: string[] = [];
    for (const sheetId of new Set([...Object.keys(base), ...Object.keys(cur)])) {
        if (!validSheetIds.includes(sheetId)) continue;
        if (stableStringify(base[sheetId] ?? []) !== stableStringify(cur[sheetId] ?? [])) {
            changed.push(sheetId);
        }
    }
    return changed;
}

const stableStringify = (value: unknown): string => {
    if (value == null) return 'null';
    if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
    if (typeof value === 'object') {
        const keys = Object.keys(value as object).sort();
        return `{${keys
            .filter(k => (value as Record<string, unknown>)[k] !== undefined)
            .map(k => `${k}:${stableStringify((value as Record<string, unknown>)[k])}`)
            .join(',')}}`;
    }
    return JSON.stringify(value);
};

const resolveStyle = (wb: IWorkbookData, s: ICellData['s']): IStyleData | null => {
    if (s == null) return null;
    if (typeof s === 'string') return (wb.styles?.[s] as IStyleData) ?? null;
    return s as IStyleData;
};

/** 单元格值语义（忽略 t/si/custom 等派生字段） */
const cellValueKey = (cell: ICellData | undefined): string => {
    if (!cell) return '';
    return stableStringify({
        v: cell.v ?? null,
        f: cell.f ?? null,
        p: cell.p ? (cell.p as { body?: unknown }).body ?? null : null,
    });
};

const mergeKey = (r: IRange) => `${r.startRow},${r.startColumn},${r.endRow},${r.endColumn}`;

const isCellEmpty = (cell: ICellData | undefined, style: IStyleData | null): boolean =>
    !cell || (cell.v == null && cell.f == null && cell.p == null && style == null);

function diffCells(
    baselineWb: IWorkbookData,
    currentWb: IWorkbookData,
    baseline: Partial<IWorksheetData>,
    current: Partial<IWorksheetData>,
): CellChange[] {
    const changes: CellChange[] = [];
    const baseCells = baseline.cellData ?? {};
    const curCells = current.cellData ?? {};

    const rowKeys = new Set([...Object.keys(baseCells), ...Object.keys(curCells)]);
    for (const rKey of rowKeys) {
        const r = Number(rKey);
        const baseRow = (baseCells as Record<number, Record<number, ICellData>>)[r] ?? {};
        const curRow = (curCells as Record<number, Record<number, ICellData>>)[r] ?? {};
        const colKeys = new Set([...Object.keys(baseRow), ...Object.keys(curRow)]);
        for (const cKey of colKeys) {
            const c = Number(cKey);
            const baseCell = baseRow[c];
            const curCell = curRow[c];
            const baseStyle = resolveStyle(baselineWb, baseCell?.s);
            const curStyle = resolveStyle(currentWb, curCell?.s);

            const valueChanged = cellValueKey(baseCell) !== cellValueKey(curCell);
            const styleChanged = stableStringify(baseStyle) !== stableStringify(curStyle);
            if (!valueChanged && !styleChanged) continue;

            const nowEmpty = isCellEmpty(curCell, curStyle);
            changes.push({
                row: r,
                col: c,
                cell: nowEmpty ? null : (curCell ?? {}),
                style: curStyle,
                valueChanged,
                styleChanged,
            });
        }
    }
    return changes;
}

function diffSheet(
    baselineWb: IWorkbookData,
    currentWb: IWorkbookData,
    baseline: Partial<IWorksheetData>,
    current: Partial<IWorksheetData>,
): Omit<SheetDiff, 'sheetId' | 'status'> {
    const cellChanges = diffCells(baselineWb, currentWb, baseline, current);

    const baseMerges = new Map((baseline.mergeData ?? []).map(m => [mergeKey(m as IRange), m as IRange]));
    const curMerges = new Map((current.mergeData ?? []).map(m => [mergeKey(m as IRange), m as IRange]));
    const mergesAdded = [...curMerges.entries()].filter(([k]) => !baseMerges.has(k)).map(([, v]) => v);
    const mergesRemoved = [...baseMerges.entries()].filter(([k]) => !curMerges.has(k)).map(([, v]) => v);

    const rowChanges: SheetDiff['rowChanges'] = [];
    {
        const baseRows = (baseline.rowData ?? {}) as Record<number, { h?: number; hd?: 0 | 1 }>;
        const curRows = (current.rowData ?? {}) as Record<number, { h?: number; hd?: 0 | 1 }>;
        for (const key of new Set([...Object.keys(baseRows), ...Object.keys(curRows)])) {
            const i = Number(key);
            const b = baseRows[i];
            const c = curRows[i];
            const hChanged = (b?.h ?? null) !== (c?.h ?? null);
            const hdChanged = (b?.hd ?? 0) !== (c?.hd ?? 0);
            if (hChanged || hdChanged) {
                rowChanges.push({
                    index: i,
                    ...(hChanged ? { h: c?.h ?? null } : {}),
                    ...(hdChanged ? { hidden: c?.hd === 1 } : {}),
                });
            }
        }
    }

    const colChanges: SheetDiff['colChanges'] = [];
    {
        const baseCols = (baseline.columnData ?? {}) as Record<number, { w?: number; hd?: 0 | 1 }>;
        const curCols = (current.columnData ?? {}) as Record<number, { w?: number; hd?: 0 | 1 }>;
        for (const key of new Set([...Object.keys(baseCols), ...Object.keys(curCols)])) {
            const i = Number(key);
            const b = baseCols[i];
            const c = curCols[i];
            const wChanged = (b?.w ?? null) !== (c?.w ?? null);
            const hdChanged = (b?.hd ?? 0) !== (c?.hd ?? 0);
            if (wChanged || hdChanged) {
                colChanges.push({
                    index: i,
                    ...(wChanged ? { w: c?.w ?? null } : {}),
                    ...(hdChanged ? { hidden: c?.hd === 1 } : {}),
                });
            }
        }
    }

    const normFreeze = (f?: IFreeze | null) =>
        f && (f.xSplit > 0 || f.ySplit > 0) ? { xSplit: f.xSplit, ySplit: f.ySplit } : null;
    const freezeChanged = stableStringify(normFreeze(baseline.freeze)) !== stableStringify(normFreeze(current.freeze));

    return {
        name: current.name ?? '',
        renamed: (baseline.name ?? '') !== (current.name ?? ''),
        cellChanges,
        mergesAdded,
        mergesRemoved,
        rowChanges,
        colChanges,
        freezeChanged,
        freeze: current.freeze ?? null,
    };
}

export function diffWorkbook(
    baseline: IWorkbookData,
    current: IWorkbookData,
    structuralSheetIds: Set<string>,
): WorkbookDiff {
    const sheets: SheetDiff[] = [];
    const baseIds = new Set(baseline.sheetOrder);

    for (const sheetId of current.sheetOrder) {
        const curSheet = current.sheets[sheetId];
        if (!baseIds.has(sheetId)) {
            sheets.push({
                sheetId,
                status: 'added',
                name: curSheet.name ?? sheetId,
                renamed: false,
                cellChanges: [],
                mergesAdded: [],
                mergesRemoved: [],
                rowChanges: [],
                colChanges: [],
                freezeChanged: false,
                freeze: curSheet.freeze ?? null,
            });
            continue;
        }
        const baseSheet = baseline.sheets[sheetId];
        const d = diffSheet(baseline, current, baseSheet, curSheet);
        const structural = structuralSheetIds.has(sheetId);
        sheets.push({
            sheetId,
            status: structural ? 'rebuilt' : 'incremental',
            ...d,
        });
    }

    const removedSheetIds = baseline.sheetOrder.filter(id => !current.sheetOrder.includes(id));
    const remainingBaseOrder = baseline.sheetOrder.filter(id => current.sheetOrder.includes(id));
    const currentOldOrder = current.sheetOrder.filter(id => baseIds.has(id));
    const orderChanged = stableStringify(remainingBaseOrder) !== stableStringify(currentOldOrder)
        || current.sheetOrder.some((id, i) => !baseIds.has(id) && i !== current.sheetOrder.length - 1);

    const curDv = parsePerSheetResource<UniverDvRule>(current, DV_PLUGIN);
    const dvChangedSheetIds = diffPerSheetResource(
        parsePerSheetResource<UniverDvRule>(baseline, DV_PLUGIN), curDv, current.sheetOrder);

    const curCf = parsePerSheetResource<unknown>(current, CF_PLUGIN);
    const cfChangedSheetIds = diffPerSheetResource(
        parsePerSheetResource<unknown>(baseline, CF_PLUGIN), curCf, current.sheetOrder);

    const isEmpty = removedSheetIds.length === 0
        && !orderChanged
        && dvChangedSheetIds.length === 0
        && cfChangedSheetIds.length === 0
        && sheets.every(s => s.status === 'incremental'
            && !s.renamed
            && !s.cellChanges.length
            && !s.mergesAdded.length
            && !s.mergesRemoved.length
            && !s.rowChanges.length
            && !s.colChanges.length
            && !s.freezeChanged);

    return {
        sheets, removedSheetIds, orderChanged, isEmpty,
        dvRules: curDv, dvChangedSheetIds,
        cfRules: curCf, cfChangedSheetIds,
    };
}

/** 单工作簿的格式投影（样式/合并/冻结/行高列宽），用于「格式是否变更过」对比 */
function formattingProjection(wb: IWorkbookData): string {
    const proj = wb.sheetOrder.map(sheetId => {
        const sheet = wb.sheets[sheetId];
        const cellStyles: [number, number, string][] = [];
        for (const [rKey, row] of Object.entries(sheet.cellData ?? {})) {
            for (const [cKey, cell] of Object.entries(row as Record<number, ICellData>)) {
                const style = resolveStyle(wb, (cell as ICellData).s);
                if (style || (cell as ICellData).p) {
                    cellStyles.push([Number(rKey), Number(cKey), stableStringify({ s: style, p: !!(cell as ICellData).p })]);
                }
            }
        }
        cellStyles.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
        return {
            merges: (sheet.mergeData ?? []).map(m => mergeKey(m as IRange)).sort(),
            freeze: sheet.freeze && (sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0)
                ? { x: sheet.freeze.xSplit, y: sheet.freeze.ySplit }
                : null,
            rows: sheet.rowData ?? {},
            cols: sheet.columnData ?? {},
            cellStyles,
        };
    });
    return stableStringify(proj);
}

/** 相对 baseline 是否发生过格式性变更（非 xlsx 保存前确认对话框用，语义同 legacy hasFormattingChanged） */
export function hasFormattingChangedUniver(baseline: IWorkbookData, current: IWorkbookData): boolean {
    return formattingProjection(baseline) !== formattingProjection(current);
}
