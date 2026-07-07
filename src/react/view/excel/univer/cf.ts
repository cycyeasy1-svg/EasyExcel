/**
 * 条件格式双向映射：ExcelJS conditionalFormattings ↔ Univer CF 资源规则。
 * 覆盖常见类型：cellIs(number)、containsText 系、expression、
 * duplicate/unique、colorScale、dataBar；其余类型跳过（未编辑时由
 * 增量导出保留原文件规则，不受影响）。
 */
import type * as ExcelJS from '@cweijan/exceljs';
import type { IRange, IStyleData } from '@univerjs/core';
import { resolveExcelColor, type ExcelColorLike } from './theme_colors';
import { parseRef } from './features';
import { hexToArgb } from '../excel_styles_core';

export const CF_PLUGIN = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

export interface UniverCfRule {
    cfId: string;
    ranges: IRange[];
    stopIfTrue: boolean;
    rule: Record<string, unknown> & { type: string };
}

interface ExcelCfEntry {
    ref: string;
    rules: ExcelCfRule[];
}

interface ExcelCfRule {
    type: string;
    priority?: number;
    operator?: string;
    formulae?: (string | number)[];
    text?: string;
    style?: {
        font?: { bold?: boolean; italic?: boolean; color?: ExcelColorLike };
        fill?: { bgColor?: ExcelColorLike; fgColor?: ExcelColorLike; type?: string; pattern?: string };
    };
    cfvo?: { type: string; value?: number | string }[];
    color?: ExcelColorLike[];
    stopIfTrue?: boolean;
}

const NUMBER_OPERATORS = new Set([
    'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual',
    'between', 'notBetween', 'equal', 'notEqual',
]);

const TEXT_TYPE_TO_OPERATOR: Record<string, string> = {
    containsText: 'containsText',
    notContainsText: 'notContainsText',
    beginsWith: 'beginsWith',
    endsWith: 'endsWith',
};

const CFVO_TYPE_MAP: Record<string, string> = {
    min: 'min',
    max: 'max',
    num: 'num',
    percent: 'percent',
    percentile: 'percentile',
    formula: 'formula',
    autoMin: 'min',
    autoMax: 'max',
};

function dxfToUniverStyle(style: ExcelCfRule['style'], palette: string[]): IStyleData {
    const out: IStyleData = {};
    if (!style) return out;
    if (style.font?.bold) out.bl = 1;
    if (style.font?.italic) out.it = 1;
    const cl = resolveExcelColor(style.font?.color, palette);
    if (cl) out.cl = { rgb: cl };
    const bg = resolveExcelColor(style.fill?.bgColor, palette) ?? resolveExcelColor(style.fill?.fgColor, palette);
    if (bg) out.bg = { rgb: bg };
    return out;
}

function cfvoToValueConfig(cfvo: { type: string; value?: number | string } | undefined, fallback: 'min' | 'max') {
    if (!cfvo) return { type: fallback };
    const type = CFVO_TYPE_MAP[cfvo.type] ?? fallback;
    if (type === 'min' || type === 'max') return { type };
    return { type, value: cfvo.value };
}

function excelRuleToUniver(entry: ExcelCfEntry, rule: ExcelCfRule, palette: string[], cfId: string): UniverCfRule | null {
    const ranges = entry.ref.split(/\s+/).map(parseRef).filter(Boolean) as IRange[];
    if (!ranges.length) return null;
    const base = { cfId, ranges, stopIfTrue: !!rule.stopIfTrue };

    if (rule.type === 'cellIs' && rule.operator && NUMBER_OPERATORS.has(rule.operator)) {
        const nums = (rule.formulae ?? []).map(f => Number(f));
        if (nums.some(Number.isNaN)) return null;
        const value = rule.operator === 'between' || rule.operator === 'notBetween'
            ? [nums[0] ?? 0, nums[1] ?? 0]
            : nums[0];
        return {
            ...base,
            rule: {
                type: 'highlightCell',
                subType: 'number',
                operator: rule.operator,
                value,
                style: dxfToUniverStyle(rule.style, palette),
            },
        };
    }

    if (TEXT_TYPE_TO_OPERATOR[rule.type]) {
        const text = rule.text ?? String(rule.formulae?.[0] ?? '');
        return {
            ...base,
            rule: {
                type: 'highlightCell',
                subType: 'text',
                operator: TEXT_TYPE_TO_OPERATOR[rule.type],
                value: text,
                style: dxfToUniverStyle(rule.style, palette),
            },
        };
    }

    if (rule.type === 'expression') {
        const formula = String(rule.formulae?.[0] ?? '');
        if (!formula) return null;
        return {
            ...base,
            rule: {
                type: 'highlightCell',
                subType: 'formula',
                value: formula.startsWith('=') ? formula : `=${formula}`,
                style: dxfToUniverStyle(rule.style, palette),
            },
        };
    }

    if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
        return {
            ...base,
            rule: {
                type: 'highlightCell',
                subType: rule.type === 'duplicateValues' ? 'duplicateValues' : 'uniqueValues',
                style: dxfToUniverStyle(rule.style, palette),
            },
        };
    }

    if (rule.type === 'colorScale' && rule.cfvo?.length && rule.color?.length) {
        const config = rule.cfvo.map((cfvo, i) => ({
            index: i,
            color: resolveExcelColor(rule.color?.[i], palette) ?? '#ffffff',
            value: cfvoToValueConfig(cfvo, i === 0 ? 'min' : 'max'),
        }));
        return { ...base, rule: { type: 'colorScale', config } };
    }

    if (rule.type === 'dataBar') {
        const color = resolveExcelColor(rule.color?.[0], palette) ?? '#638ec6';
        return {
            ...base,
            rule: {
                type: 'dataBar',
                isShowValue: true,
                config: {
                    min: cfvoToValueConfig(rule.cfvo?.[0], 'min'),
                    max: cfvoToValueConfig(rule.cfvo?.[1], 'max'),
                    isGradient: false,
                    positiveColor: color,
                    nativeColor: color,
                },
            },
        };
    }

    return null;
}

/** ExcelJS worksheet 的条件格式 → Univer CF 规则数组 */
export function excelCfToUniver(worksheet: ExcelJS.Worksheet, palette: string[], sheetId: string): UniverCfRule[] {
    const model = (worksheet as { conditionalFormattings?: ExcelCfEntry[] }).conditionalFormattings
        ?? (worksheet.model as { conditionalFormattings?: ExcelCfEntry[] })?.conditionalFormattings
        ?? [];
    const rules: UniverCfRule[] = [];
    let seq = 0;
    for (const entry of model) {
        for (const rule of entry.rules ?? []) {
            seq += 1;
            const converted = excelRuleToUniver(entry, rule, palette, `${sheetId}-cf-${seq}`);
            if (converted) rules.push(converted);
        }
    }
    return rules;
}

// ---------- 导出方向 ----------

const COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const colName = (index: number): string => {
    let str = '';
    let i = index;
    while (i >= 0) {
        str = COL_LETTERS[i % 26] + str;
        i = Math.floor(i / 26) - 1;
    }
    return str;
};
const rangeToAddr = (r: IRange) => `${colName(r.startColumn)}${r.startRow + 1}:${colName(r.endColumn)}${r.endRow + 1}`;

function univerStyleToDxf(style: unknown): ExcelCfRule['style'] {
    const s = style as { bl?: number; it?: number; cl?: { rgb?: string }; bg?: { rgb?: string } } | undefined;
    const out: ExcelCfRule['style'] = {};
    if (!s) return out;
    const font: NonNullable<ExcelCfRule['style']>['font'] = {};
    if (s.bl === 1) font.bold = true;
    if (s.it === 1) font.italic = true;
    if (s.cl?.rgb) font.color = { argb: hexToArgb(s.cl.rgb) };
    if (Object.keys(font).length) out.font = font;
    if (s.bg?.rgb) {
        out.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: hexToArgb(s.bg.rgb) } } as never;
    }
    return out;
}

function valueConfigToCfvo(vc: { type?: string; value?: unknown } | undefined, fallback: string) {
    const type = vc?.type ?? fallback;
    if (type === 'min' || type === 'max') return { type };
    return { type, value: vc?.value };
}

function univerRuleToExcel(rule: UniverCfRule): { ref: string; rules: Record<string, unknown>[] } | null {
    const ref = rule.ranges.map(rangeToAddr).join(' ');
    const r = rule.rule as Record<string, unknown>;

    if (r.type === 'highlightCell') {
        const style = univerStyleToDxf(r.style);
        if (r.subType === 'number' && r.operator) {
            const value = r.value;
            const formulae = Array.isArray(value) ? value : [value];
            return { ref, rules: [{ type: 'cellIs', operator: r.operator, formulae, style, priority: 1 }] };
        }
        if (r.subType === 'text' && r.operator) {
            const opToType: Record<string, string> = {
                containsText: 'containsText',
                notContainsText: 'notContainsText',
                beginsWith: 'beginsWith',
                endsWith: 'endsWith',
            };
            const type = opToType[String(r.operator)];
            if (!type) return null;
            return { ref, rules: [{ type, text: r.value ?? '', style, priority: 1 }] };
        }
        if (r.subType === 'formula' && r.value) {
            return { ref, rules: [{ type: 'expression', formulae: [String(r.value).replace(/^=/, '')], style, priority: 1 }] };
        }
        if (r.subType === 'duplicateValues' || r.subType === 'uniqueValues') {
            return { ref, rules: [{ type: r.subType, style, priority: 1 }] };
        }
        return null;
    }

    if (r.type === 'colorScale' && Array.isArray(r.config)) {
        const config = r.config as { color: string; value: { type?: string; value?: unknown } }[];
        return {
            ref,
            rules: [{
                type: 'colorScale',
                cfvo: config.map((c, i) => valueConfigToCfvo(c.value, i === 0 ? 'min' : 'max')),
                color: config.map(c => ({ argb: hexToArgb(c.color) })),
                priority: 1,
            }],
        };
    }

    if (r.type === 'dataBar') {
        const config = r.config as { min: { type?: string; value?: unknown }; max: { type?: string; value?: unknown }; positiveColor: string };
        return {
            ref,
            rules: [{
                type: 'dataBar',
                cfvo: [valueConfigToCfvo(config?.min, 'min'), valueConfigToCfvo(config?.max, 'max')],
                color: [{ argb: hexToArgb(config?.positiveColor ?? '#638ec6') }],
                priority: 1,
            }],
        };
    }

    return null;
}

/** 覆盖式重写 worksheet 的条件格式（CF 变更过的 sheet / 重建的 sheet） */
export function writeSheetConditionalFormattings(worksheet: ExcelJS.Worksheet, rules: UniverCfRule[]) {
    const ws = worksheet as ExcelJS.Worksheet & {
        conditionalFormattings?: unknown[];
        addConditionalFormatting?: (cf: { ref: string; rules: unknown[] }) => void;
    };
    if (Array.isArray(ws.conditionalFormattings)) {
        ws.conditionalFormattings.length = 0;
    }
    if (typeof ws.addConditionalFormatting !== 'function') return;
    for (const rule of rules) {
        const converted = univerRuleToExcel(rule);
        if (!converted) continue;
        try {
            ws.addConditionalFormatting(converted as never);
        } catch {
            // 非法规则忽略
        }
    }
}
