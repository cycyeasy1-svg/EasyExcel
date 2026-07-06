/**
 * M3 功能对等层：数据验证 / 工作表保护 / 浮动图片的导入应用（facade 调用，
 * 在 adapter.loadWorkbook 内、取基线之前执行）。
 */
import type { SpreadsheetValidationItem } from '../excel_validation';
import type { SheetImage } from '../excel_images';

/** 'A1:B2' / 'C3' → 0-based range */
export function parseRef(ref: string): { startRow: number; startColumn: number; endRow: number; endColumn: number } | null {
    const m = ref.trim().toUpperCase().match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
    if (!m) return null;
    const colIndex = (letters: string) => {
        let n = 0;
        for (let i = 0; i < letters.length; i += 1) n = n * 26 + (letters.charCodeAt(i) - 64);
        return n - 1;
    };
    const startColumn = colIndex(m[1]);
    const startRow = Number(m[2]) - 1;
    const endColumn = m[3] ? colIndex(m[3]) : startColumn;
    const endRow = m[4] ? Number(m[4]) - 1 : startRow;
    return { startRow, startColumn, endRow, endColumn };
}

interface FWorksheetLike {
    getRange(row: number, col: number, numRows?: number, numCols?: number): {
        setDataValidation(rule: unknown): Promise<unknown>;
    };
    insertImage(url: string, column: number, row: number, offsetX: number, offsetY: number): Promise<boolean>;
    getWorksheetPermission(): {
        protect(options?: { name?: string }): Promise<string>;
        setReadOnly(): Promise<void>;
    };
}

interface DvBuilderLike {
    requireValueInList(values: string[], multiple?: boolean, showDropdown?: boolean): DvBuilderLike;
    requireNumberBetween(start: number, end: number): DvBuilderLike;
    requireNumberNotBetween(start: number, end: number): DvBuilderLike;
    requireNumberEqualTo(v: number): DvBuilderLike;
    requireNumberNotEqualTo(v: number): DvBuilderLike;
    requireNumberGreaterThan(v: number): DvBuilderLike;
    requireNumberGreaterThanOrEqualTo(v: number): DvBuilderLike;
    requireNumberLessThan(v: number): DvBuilderLike;
    requireNumberLessThanOrEqualTo(v: number): DvBuilderLike;
    requireFormulaSatisfied(formula: string): DvBuilderLike;
    setAllowInvalid(allow: boolean): DvBuilderLike;
    build(): unknown;
}

interface UniverAPILike {
    newDataValidation(): DvBuilderLike;
}

const num = (v: unknown): number => Number(String(v ?? '').trim());

function buildDvRule(api: UniverAPILike, item: SpreadsheetValidationItem): unknown | null {
    const b = api.newDataValidation();
    try {
        if (item.type === 'list') {
            const list = String(item.value ?? '').split(',').map(s => s.trim()).filter(Boolean);
            if (!list.length) return null;
            return b.requireValueInList(list).setAllowInvalid(!item.required).build();
        }
        if (item.type === 'number') {
            const op = item.operator ?? 'eq';
            const val = item.value;
            let builder: DvBuilderLike;
            if (op === 'be' || op === 'nbe') {
                const arr = Array.isArray(val) ? val : ['0', '0'];
                builder = op === 'be'
                    ? b.requireNumberBetween(num(arr[0]), num(arr[1]))
                    : b.requireNumberNotBetween(num(arr[0]), num(arr[1]));
            } else {
                const v = num(Array.isArray(val) ? val[0] : val);
                const map: Record<string, (x: number) => DvBuilderLike> = {
                    eq: x => b.requireNumberEqualTo(x),
                    neq: x => b.requireNumberNotEqualTo(x),
                    gt: x => b.requireNumberGreaterThan(x),
                    gte: x => b.requireNumberGreaterThanOrEqualTo(x),
                    lt: x => b.requireNumberLessThan(x),
                    lte: x => b.requireNumberLessThanOrEqualTo(x),
                };
                if (!map[op] || Number.isNaN(v)) return null;
                builder = map[op](v);
            }
            return builder.setAllowInvalid(!item.required).build();
        }
        if (item.type === 'phone' || item.type === 'email') {
            // 复用现有 custom 公式生成（refTopLeft 由调用方保证）
            const cell = item.refs[0]?.split(':')[0] ?? 'A1';
            const formula = item.type === 'phone'
                ? `=AND(LEN(${cell})=11,ISNUMBER(--${cell}),--LEFT(${cell},1)>=1)`
                : `=AND(NOT(ISERROR(FIND("@",${cell}))),NOT(ISERROR(FIND(".",${cell},FIND("@",${cell})+1))),LEN(${cell})>=5)`;
            return b.requireFormulaSatisfied(formula).setAllowInvalid(!item.required).build();
        }
        // date 等其余类型：M3 暂不映射（导入侧不应用，原文件未编辑时导出仍保留）
        return null;
    } catch {
        return null;
    }
}

export async function applyValidationsToSheet(
    api: UniverAPILike,
    fSheet: FWorksheetLike,
    items: SpreadsheetValidationItem[],
) {
    for (const item of items) {
        const rule = buildDvRule(api, item);
        if (!rule) continue;
        for (const ref of item.refs) {
            const range = parseRef(ref);
            if (!range) continue;
            try {
                await fSheet.getRange(
                    range.startRow,
                    range.startColumn,
                    range.endRow - range.startRow + 1,
                    range.endColumn - range.startColumn + 1,
                ).setDataValidation(rule);
            } catch (e) {
                console.warn('EasyExcel: failed to apply data validation', ref, e);
            }
        }
    }
}

export async function applyImagesToSheet(fSheet: FWorksheetLike, images: SheetImage[]) {
    for (const image of images) {
        try {
            const dataUrl = `data:image/${image.extension};base64,${image.base64}`;
            const col = Math.floor(image.anchor.col);
            const row = Math.floor(image.anchor.row);
            await fSheet.insertImage(dataUrl, col, row, 0, 0);
        } catch (e) {
            console.warn('EasyExcel: failed to insert image', e);
        }
    }
}

/** 受保护工作表 → Univer 只读（M3 取舍：unlocked 例外单元格暂不支持，见报告） */
export async function applyProtectionToSheet(fSheet: FWorksheetLike) {
    try {
        const permission = fSheet.getWorksheetPermission();
        // setReadOnly 前置要求保护规则已存在
        await permission.protect({ name: 'EasyExcel imported protection' });
        await permission.setReadOnly();
    } catch (e) {
        console.warn('EasyExcel: failed to protect sheet', e);
    }
}
