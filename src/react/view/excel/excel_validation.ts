import type * as ExcelJS from '@cweijan/exceljs';

export interface SpreadsheetValidationItem {
    refs: string[];
    mode: string;
    type: string;
    required?: boolean;
    operator?: string;
    value?: string | string[] | number;
}

const EXCEL_TO_SS_OPERATOR: Record<string, string> = {
    between: 'be',
    notBetween: 'nbe',
    equal: 'eq',
    notEqual: 'neq',
    lessThan: 'lt',
    lessThanOrEqual: 'lte',
    greaterThan: 'gt',
    greaterThanOrEqual: 'gte',
};

function parseListFormula(raw: unknown): string {
    if (raw == null) return '';
    let text = String(raw).trim();
    if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
        text = text.slice(1, -1);
    }
    return text;
}

function refTopLeft(ref: string): string {
    const idx = ref.indexOf(':');
    return idx === -1 ? ref : ref.slice(0, idx);
}

/** Excel custom validation: 11-digit mobile starting with 1–9 */
function buildPhoneFormula(cell: string): string {
    return `AND(LEN(${cell})=11,ISNUMBER(--${cell}),--LEFT(${cell},1)>=1)`;
}

/** Excel custom validation: basic email shape */
function buildEmailFormula(cell: string): string {
    return `AND(NOT(ISERROR(FIND("@",${cell}))),NOT(ISERROR(FIND(".",${cell},FIND("@",${cell})+1))),LEN(${cell})>=5)`;
}

function normalizeFormula(formula: string): string {
    return formula.trim().replace(/^=/, '').replace(/\$/g, '').toUpperCase();
}

function detectCustomValidationType(formula: string, ref: string): 'phone' | 'email' | null {
    const cell = refTopLeft(ref);
    const norm = normalizeFormula(formula);
    if (norm === normalizeFormula(buildPhoneFormula(cell))) return 'phone';
    if (norm === normalizeFormula(buildEmailFormula(cell))) return 'email';
    if (/LEN\([^)]+\)=11/.test(norm) && /ISNUMBER/.test(norm) && /--LEFT\([^)]+\)>=1/.test(norm)) {
        return 'phone';
    }
    if (/FIND\("@"/.test(norm) && /FIND\("\./.test(norm) && /LEN\([^)]+\)>=5/.test(norm)) {
        return 'email';
    }
    return null;
}

export function excelValidationToSpreadsheet(
    ref: string,
    dv: ExcelJS.DataValidation,
): SpreadsheetValidationItem | null {
    const mode = 'cell';
    if (dv.type === 'list') {
        return {
            refs: [ref],
            mode,
            type: 'list',
            required: !dv.allowBlank,
            value: parseListFormula(dv.formulae?.[0]),
        };
    }
    if (dv.type === 'whole' || dv.type === 'decimal') {
        const operator = dv.operator ? EXCEL_TO_SS_OPERATOR[dv.operator] : undefined;
        let value: string | string[] | number = '';
        if (dv.operator === 'between' || dv.operator === 'notBetween') {
            value = [String(dv.formulae?.[0] ?? ''), String(dv.formulae?.[1] ?? '')];
        } else {
            value = String(dv.formulae?.[0] ?? '');
        }
        return {
            refs: [ref],
            mode,
            type: 'number',
            required: !dv.allowBlank,
            operator,
            value,
        };
    }
    if (dv.type === 'date') {
        const operator = dv.operator ? EXCEL_TO_SS_OPERATOR[dv.operator] : undefined;
        let value: string | string[] = '';
        if (dv.operator === 'between' || dv.operator === 'notBetween') {
            value = [String(dv.formulae?.[0] ?? ''), String(dv.formulae?.[1] ?? '')];
        } else {
            value = String(dv.formulae?.[0] ?? '');
        }
        return {
            refs: [ref],
            mode,
            type: 'date',
            required: !dv.allowBlank,
            operator,
            value,
        };
    }
    if (dv.type === 'custom') {
        const customType = detectCustomValidationType(String(dv.formulae?.[0] ?? ''), ref);
        if (customType) {
            return {
                refs: [ref],
                mode,
                type: customType,
                required: !dv.allowBlank,
                value: '',
            };
        }
    }
    return null;
}

export function readWorksheetValidations(worksheet: ExcelJS.Worksheet): SpreadsheetValidationItem[] {
    const model = (worksheet as { dataValidations?: { model?: Record<string, ExcelJS.DataValidation> } })
        .dataValidations?.model ?? {};
    const items: SpreadsheetValidationItem[] = [];
    for (const ref of Object.keys(model)) {
        const dv = model[ref];
        if (!dv) continue;
        const converted = excelValidationToSpreadsheet(ref, dv);
        if (converted) items.push(converted);
    }
    return items;
}
