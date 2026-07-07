/**
 * M5：XML 级差分补丁 —— 字节级无损保存。
 *
 * 保存到原文件时不再经 ExcelJS 重新序列化，而是在原始 zip 字节上做
 * 字符串手术：只改被编辑 cell 所在 sheet XML 的对应 <c> 节点；样式变更
 * 对 styles.xml 只追加（原索引全部不动）；其余 zip 部件原样透传 ——
 * 图形/图表/透视表/宏/打印设置等 ExcelJS 未建模的部件物理上不被触碰。
 *
 * 关键取舍：
 * - 字符串写入用 inlineStr，不动 sharedStrings.xml（Excel 完全接受）。
 * - 结构性变更（插删行列/sheet、合并、行高列宽、冻结、DV/CF）不支持，
 *   由 assertPatchableDiff 防御性拒绝（UI 层已禁用对应操作）。
 * - 富文本格的值编辑降级为纯文本 inlineStr（结果标记 richTextDowngraded）。
 * - 有值变更且存在公式时删除 calcChain.xml（连同 content-types/rels 引用）
 *   并给 calcPr 补 fullCalcOnLoad="1"，让 Excel 打开时重算，避免陈旧缓存。
 * - 与 M2 相同的已知限制：仅有 si（共享公式引用）没有 f 的 cell 按计算值写入。
 */
import JSZip from 'jszip';
import {
    BorderStyleTypes,
    HorizontalAlign,
    VerticalAlign,
    WrapStrategy,
    type ICellData,
    type IStyleData,
} from '@univerjs/core';
import { hexToArgb } from '../excel_styles_core';
import type { CellChange, WorkbookDiff } from './diff';

/** 无法无损保存的变更种类（UI 拦截失守时的最后防线） */
export class XmlPatchBlockedError extends Error {
    constructor(readonly reason: 'structural' | 'dvcf') {
        super(`xml patch blocked: ${reason} change cannot be saved losslessly`);
        this.name = 'XmlPatchBlockedError';
    }
}

export interface XmlPatchResult {
    bytes: Uint8Array;
    /** 有富文本单元格的值被降级为纯文本写入 */
    richTextDowngraded: boolean;
    /** 本次实际改写的部件（审计用：除此之外全部逐字节透传） */
    changedParts: string[];
    /** 本次删除的部件（calcChain） */
    removedParts: string[];
}

/** diff 是否只含 XML 补丁能承载的变更；不能则抛 XmlPatchBlockedError */
export function assertPatchableDiff(diff: WorkbookDiff): void {
    const structural = diff.removedSheetIds.length > 0
        || diff.orderChanged
        || diff.sheets.some(s => s.status !== 'incremental'
            || s.renamed
            || s.mergesAdded.length > 0
            || s.mergesRemoved.length > 0
            || s.rowChanges.length > 0
            || s.colChanges.length > 0
            || s.freezeChanged);
    if (structural) throw new XmlPatchBlockedError('structural');
    if (diff.dvChangedSheetIds.length || diff.cfChangedSheetIds.length) {
        throw new XmlPatchBlockedError('dvcf');
    }
}

// ---------------------------------------------------------------- XML 基础

const escapeXmlText = (s: string): string => s
    // XML 1.0 非法控制字符（\t\n\r 除外）直接剔除
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const unescapeXml = (s: string): string => s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

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

/** "BC" → 54（0 起） */
function colIndex(letters: string): number {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
}

/** 文本编辑操作：先收集再按位置倒序应用，保持偏移量有效 */
interface TextEdit {
    start: number;
    end: number;
    text: string;
    /** 同一插入点的稳定次序（行号/列号升序） */
    order: number;
}

function applyEdits(source: string, edits: TextEdit[]): string {
    const sorted = [...edits].sort((a, b) => b.start - a.start || b.order - a.order);
    let out = source;
    for (const e of sorted) {
        out = out.slice(0, e.start) + e.text + out.slice(e.end);
    }
    return out;
}

// ---------------------------------------------------------- 样式 append-only

const HALIGN_XML: Record<number, string> = {
    [HorizontalAlign.LEFT]: 'left',
    [HorizontalAlign.CENTER]: 'center',
    [HorizontalAlign.RIGHT]: 'right',
    [HorizontalAlign.JUSTIFIED]: 'justify',
};

/** 注意：OOXML 的垂直居中是 "center"（非 ExcelJS 的 "middle"） */
const VALIGN_XML: Record<number, string> = {
    [VerticalAlign.TOP]: 'top',
    [VerticalAlign.MIDDLE]: 'center',
    [VerticalAlign.BOTTOM]: 'bottom',
};

const BORDER_XML: Record<number, string> = {
    [BorderStyleTypes.THIN]: 'thin',
    [BorderStyleTypes.HAIR]: 'hair',
    [BorderStyleTypes.DOTTED]: 'dotted',
    [BorderStyleTypes.DASHED]: 'dashed',
    [BorderStyleTypes.DASH_DOT]: 'dashDot',
    [BorderStyleTypes.DASH_DOT_DOT]: 'dashDotDot',
    [BorderStyleTypes.DOUBLE]: 'double',
    [BorderStyleTypes.MEDIUM]: 'medium',
    [BorderStyleTypes.MEDIUM_DASHED]: 'mediumDashed',
    [BorderStyleTypes.MEDIUM_DASH_DOT]: 'mediumDashDot',
    [BorderStyleTypes.MEDIUM_DASH_DOT_DOT]: 'mediumDashDotDot',
    [BorderStyleTypes.SLANT_DASH_DOT]: 'slantDashDot',
    [BorderStyleTypes.THICK]: 'thick',
};

/** 与 export_styles.LooseStyle 同形：Nullable<T> 属性访问的宽松视图 */
interface LooseStyle {
    ff?: string | null;
    fs?: number;
    bl?: number;
    it?: number;
    st?: { s?: number } | null;
    ul?: { s?: number } | null;
    cl?: { rgb?: string | null } | null;
    bg?: { rgb?: string | null } | null;
    bd?: Record<string, { s: BorderStyleTypes; cl?: { rgb?: string | null } | null } | null | undefined> | null;
    ht?: number;
    vt?: number;
    tb?: number;
    tr?: { a?: number; v?: number } | null;
    n?: { pattern?: string } | null;
}

const argbOf = (rgb?: string | null): string | undefined => hexToArgb(rgb ?? undefined);

function fontXml(style: LooseStyle): string | null {
    const parts: string[] = [];
    if (style.bl === 1) parts.push('<b/>');
    if (style.it === 1) parts.push('<i/>');
    if (style.st?.s === 1) parts.push('<strike/>');
    if (style.ul?.s === 1) parts.push('<u/>');
    if (style.fs) parts.push(`<sz val="${style.fs}"/>`);
    const color = argbOf(style.cl?.rgb);
    if (color) parts.push(`<color rgb="${color}"/>`);
    if (style.ff) parts.push(`<name val="${escapeXmlText(style.ff).replace(/"/g, '&quot;')}"/>`);
    return parts.length ? `<font>${parts.join('')}</font>` : null;
}

function fillXml(style: LooseStyle): string | null {
    const bg = argbOf(style.bg?.rgb);
    if (!bg) return null;
    return `<fill><patternFill patternType="solid"><fgColor rgb="${bg}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function borderXml(style: LooseStyle): string | null {
    const bd = style.bd;
    if (!bd) return null;
    const side = (key: string): string | null => {
        const b = bd[key];
        if (!b) return null;
        const token = BORDER_XML[b.s];
        if (!token) return null;
        const color = argbOf(b.cl?.rgb);
        return `${token}|${color ?? ''}`;
    };
    const render = (tag: string, spec: string | null): string => {
        if (!spec) return `<${tag}/>`;
        const [token, color] = spec.split('|');
        const colorEl = color ? `<color rgb="${color}"/>` : '';
        return `<${tag} style="${token}">${colorEl}</${tag}>`;
    };
    const l = side('l');
    const r = side('r');
    const t = side('t');
    const b = side('b');
    const down = side('tl_br');
    const up = side('bl_tr');
    if (!l && !r && !t && !b && !down && !up) return null;
    const diagSpec = down ?? up;
    const diagAttrs = `${up ? ' diagonalUp="1"' : ''}${down ? ' diagonalDown="1"' : ''}`;
    const diagonal = diagSpec
        ? render('diagonal', diagSpec)
        : '<diagonal/>';
    return `<border${diagAttrs}>${render('left', l)}${render('right', r)}${render('top', t)}${render('bottom', b)}${diagonal}</border>`;
}

function alignmentXml(style: LooseStyle): string | null {
    const attrs: string[] = [];
    const h = style.ht != null ? HALIGN_XML[style.ht] : undefined;
    const v = style.vt != null ? VALIGN_XML[style.vt] : undefined;
    if (h) attrs.push(`horizontal="${h}"`);
    if (v) attrs.push(`vertical="${v}"`);
    if (style.tb === WrapStrategy.WRAP) attrs.push('wrapText="1"');
    if (style.tr?.a || style.tr?.v) {
        const rotation = style.tr.v === 1 ? 255 : ((style.tr.a ?? 0) < 0 ? 90 - (style.tr.a ?? 0) : style.tr.a ?? 0);
        if (rotation) attrs.push(`textRotation="${rotation}"`);
    }
    return attrs.length ? `<alignment ${attrs.join(' ')}/>` : null;
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

/** 统计块内子元素个数（<font ...> 匹配、<fonts> 不匹配） */
function countChildren(block: string, tag: string): number {
    const re = new RegExp(`<${tag}[\\s/>]`, 'g');
    return (block.match(re) ?? []).length;
}

/**
 * styles.xml 的 append-only 修改器：对 styleChanged 的 cell 把「完整样式」
 * （Univer 语义，非增量）整体映射为新 font/fill/border/numFmt + 新 xf，
 * 追加到各集合尾部；原有索引与内容一个字节不动。
 */
class StylesPatcher {
    private readonly newFonts: string[] = [];
    private readonly newFills: string[] = [];
    private readonly newBorders: string[] = [];
    private readonly newNumFmts: string[] = [];
    private readonly newXfs: string[] = [];
    private readonly xfCache = new Map<string, number>();
    private readonly baseFontCount: number;
    private readonly baseFillCount: number;
    private readonly baseBorderCount: number;
    private readonly baseXfCount: number;
    private nextNumFmtId: number;

    constructor(private readonly xml: string) {
        const block = (tag: string): string => {
            const m = xml.match(new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>`));
            return m?.[0] ?? '';
        };
        this.baseFontCount = countChildren(block('fonts'), 'font');
        this.baseFillCount = countChildren(block('fills'), 'fill');
        this.baseBorderCount = countChildren(block('borders'), 'border');
        this.baseXfCount = countChildren(block('cellXfs'), 'xf');
        let maxId = 163;
        for (const m of xml.matchAll(/<numFmt\b[^>]*\bnumFmtId="(\d+)"/g)) {
            maxId = Math.max(maxId, Number(m[1]));
        }
        this.nextNumFmtId = maxId + 1;
    }

    /** 完整样式 → cellXfs 索引（同样式去重复用）；null 样式 → 0（默认 xf） */
    xfIndexFor(style: IStyleData | null): number {
        if (style == null) return 0;
        const key = stableStringify(style);
        const cached = this.xfCache.get(key);
        if (cached != null) return cached;

        const loose = style as LooseStyle;
        const font = fontXml(loose);
        const fill = fillXml(loose);
        const border = borderXml(loose);
        const alignment = alignmentXml(loose);

        let fontId = 0;
        if (font) {
            fontId = this.baseFontCount + this.newFonts.length;
            this.newFonts.push(font);
        }
        let fillId = 0;
        if (fill) {
            fillId = this.baseFillCount + this.newFills.length;
            this.newFills.push(fill);
        }
        let borderId = 0;
        if (border) {
            borderId = this.baseBorderCount + this.newBorders.length;
            this.newBorders.push(border);
        }
        let numFmtId = 0;
        if (loose.n?.pattern) {
            numFmtId = this.nextNumFmtId;
            this.nextNumFmtId += 1;
            const code = escapeXmlText(loose.n.pattern).replace(/"/g, '&quot;');
            this.newNumFmts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${code}"/>`);
        }

        const attrs = [
            `numFmtId="${numFmtId}"`,
            `fontId="${fontId}"`,
            `fillId="${fillId}"`,
            `borderId="${borderId}"`,
            'xfId="0"',
            ...(numFmtId ? ['applyNumberFormat="1"'] : []),
            ...(font ? ['applyFont="1"'] : []),
            ...(fill ? ['applyFill="1"'] : []),
            ...(border ? ['applyBorder="1"'] : []),
            ...(alignment ? ['applyAlignment="1"'] : []),
        ].join(' ');
        const xf = alignment ? `<xf ${attrs}>${alignment}</xf>` : `<xf ${attrs}/>`;
        const index = this.baseXfCount + this.newXfs.length;
        this.newXfs.push(xf);
        this.xfCache.set(key, index);
        return index;
    }

    get touched(): boolean {
        return this.newXfs.length > 0;
    }

    /** 应用追加；未追加过则原样返回（styles.xml 字节不动） */
    serialize(): string {
        if (!this.touched) return this.xml;
        let out = this.xml;

        const appendInto = (tag: string, items: string[], baseCount: number) => {
            if (!items.length) return;
            const close = `</${tag}>`;
            const idx = out.lastIndexOf(close);
            if (idx < 0) throw new Error(`styles.xml: missing </${tag}>`);
            out = out.slice(0, idx) + items.join('') + out.slice(idx);
            // count 属性同步（Excel 宽容 count 不符，但保持正确）
            out = out.replace(
                new RegExp(`(<${tag}\\b[^>]*\\bcount=")(\\d+)(")`),
                (_, pre, _n, post) => `${pre}${baseCount + items.length}${post}`,
            );
        };

        if (this.newNumFmts.length) {
            if (/<numFmts\b[^>]*>/.test(out)) {
                const base = countChildren(out.match(/<numFmts[^>]*>[\s\S]*?<\/numFmts>/)?.[0] ?? '', 'numFmt');
                appendInto('numFmts', this.newNumFmts, base);
            } else {
                const fontsIdx = out.indexOf('<fonts');
                if (fontsIdx < 0) throw new Error('styles.xml: missing <fonts>');
                out = out.slice(0, fontsIdx)
                    + `<numFmts count="${this.newNumFmts.length}">${this.newNumFmts.join('')}</numFmts>`
                    + out.slice(fontsIdx);
            }
        }
        appendInto('fonts', this.newFonts, this.baseFontCount);
        appendInto('fills', this.newFills, this.baseFillCount);
        appendInto('borders', this.newBorders, this.baseBorderCount);
        appendInto('cellXfs', this.newXfs, this.baseXfCount);
        return out;
    }
}

// ------------------------------------------------------------- cell 值写入

interface DocBodyLike {
    dataStream?: string;
}

/** 富文本降级：抽出纯文本（Univer doc 流固定以 \r\n 收尾） */
function plainTextOfRichCell(p: ICellData['p']): string {
    const stream = (p as { body?: DocBodyLike } | undefined)?.body?.dataStream ?? '';
    return stream.replace(/\r?\n$/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

const needsSpacePreserve = (text: string): boolean =>
    /^\s|\s$|[\n\t]/.test(text);

function inlineStrContent(text: string): { t: string; inner: string } {
    const space = needsSpacePreserve(text) ? ' xml:space="preserve"' : '';
    return { t: 'inlineStr', inner: `<is><t${space}>${escapeXmlText(text)}</t></is>` };
}

/**
 * 变更 cell 的类型化内容。返回 t 属性与节点内 XML；t 为空串表示无 t 属性。
 * Univer CellValueType：1=STRING 2=NUMBER 3=BOOLEAN 4=FORCE_STRING。
 */
function cellContent(cell: ICellData): { t: string; inner: string; richTextDowngraded: boolean } {
    if (cell.f) {
        const formula = escapeXmlText(String(cell.f).replace(/^=/, ''));
        const v = cell.v;
        if (typeof v === 'number' && Number.isFinite(v)) {
            return { t: '', inner: `<f>${formula}</f><v>${v}</v>`, richTextDowngraded: false };
        }
        if (typeof v === 'boolean' || cell.t === 3) {
            return { t: 'b', inner: `<f>${formula}</f><v>${v ? 1 : 0}</v>`, richTextDowngraded: false };
        }
        if (typeof v === 'string' && v !== '') {
            return { t: 'str', inner: `<f>${formula}</f><v>${escapeXmlText(v)}</v>`, richTextDowngraded: false };
        }
        return { t: '', inner: `<f>${formula}</f>`, richTextDowngraded: false };
    }
    if (cell.p) {
        const { t, inner } = inlineStrContent(plainTextOfRichCell(cell.p));
        return { t, inner, richTextDowngraded: true };
    }
    const v = cell.v;
    if (v == null) return { t: '', inner: '', richTextDowngraded: false };
    if (typeof v === 'boolean' || cell.t === 3) {
        return { t: 'b', inner: `<v>${v ? 1 : 0}</v>`, richTextDowngraded: false };
    }
    if (typeof v === 'number' && cell.t !== 4) {
        if (!Number.isFinite(v)) {
            const { t, inner } = inlineStrContent(String(v));
            return { t, inner, richTextDowngraded: false };
        }
        return { t: '', inner: `<v>${v}</v>`, richTextDowngraded: false };
    }
    const { t, inner } = inlineStrContent(String(v));
    return { t, inner, richTextDowngraded: false };
}

// ------------------------------------------------------------ sheet XML 补丁

interface CellChunk {
    start: number;
    end: number;
    col: number;
    openTag: string;
    inner: string;
    selfClosed: boolean;
}

interface RowChunk {
    start: number;
    end: number;
    r: number;
    chunk: string;
}

const ROW_RE = /<row\b[^>]*?\/>|<row\b[^>]*>[\s\S]*?<\/row>/g;
const CELL_RE = /<c\b[^>]*?\/>|<c\b[^>]*>[\s\S]*?<\/c>/g;

function parseRows(sheetDataInner: string, baseOffset: number): RowChunk[] {
    const rows: RowChunk[] = [];
    ROW_RE.lastIndex = 0;
    for (const m of sheetDataInner.matchAll(ROW_RE)) {
        const rAttr = m[0].match(/^<row\b[^>]*?\br="(\d+)"/);
        if (!rAttr) continue;
        rows.push({
            start: baseOffset + (m.index ?? 0),
            end: baseOffset + (m.index ?? 0) + m[0].length,
            r: Number(rAttr[1]),
            chunk: m[0],
        });
    }
    return rows;
}

function parseCells(rowChunk: string): CellChunk[] {
    const cells: CellChunk[] = [];
    for (const m of rowChunk.matchAll(CELL_RE)) {
        const open = m[0].match(/^<c\b[^>]*?>/)?.[0] ?? '';
        const refAttr = open.match(/\br="([A-Z]+)(\d+)"/);
        if (!refAttr) continue;
        const selfClosed = open.endsWith('/>');
        cells.push({
            start: m.index ?? 0,
            end: (m.index ?? 0) + m[0].length,
            col: colIndex(refAttr[1]),
            openTag: open,
            inner: selfClosed ? '' : m[0].slice(open.length, m[0].length - '</c>'.length),
            selfClosed,
        });
    }
    return cells;
}

/** 开标签里除 r/s/t 之外的属性（cm/vm/ph 等原样保留） */
function keptAttrs(openTag: string): string {
    const attrs: string[] = [];
    for (const m of openTag.matchAll(/([\w:]+)="([^"]*)"/g)) {
        if (m[1] === 'r' || m[1] === 's' || m[1] === 't') continue;
        attrs.push(` ${m[1]}="${m[2]}"`);
    }
    return attrs.join('');
}

const attrValue = (openTag: string, name: string): string | null =>
    openTag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? null;

interface BuiltCell {
    xml: string;
    richTextDowngraded: boolean;
    wroteFormula: boolean;
}

/** 由变更生成 <c> 节点；existing 为原节点（属性/内容按需保留） */
function buildCell(
    ref: string,
    change: CellChange,
    styles: StylesPatcher,
    existing: CellChunk | null,
): BuiltCell {
    const extra = existing ? keptAttrs(existing.openTag) : '';

    let sAttr = '';
    if (change.styleChanged) {
        if (change.style != null) {
            sAttr = ` s="${styles.xfIndexFor(change.style)}"`;
        }
    } else if (existing) {
        const s = attrValue(existing.openTag, 's');
        if (s != null) sAttr = ` s="${s}"`;
    }

    if (!change.valueChanged && existing) {
        // 只改样式：内容与 t 原样保留
        const t = attrValue(existing.openTag, 't');
        const tAttr = t != null ? ` t="${t}"` : '';
        const open = `<c r="${ref}"${sAttr}${tAttr}${extra}`;
        const xml = existing.selfClosed || existing.inner === ''
            ? `${open}/>`
            : `${open}>${existing.inner}</c>`;
        return { xml, richTextDowngraded: false, wroteFormula: false };
    }

    const content = cellContent(change.cell ?? {});
    const tAttr = content.t ? ` t="${content.t}"` : '';
    const open = `<c r="${ref}"${sAttr}${tAttr}${extra}`;
    const xml = content.inner ? `${open}>${content.inner}</c>` : `${open}/>`;
    return {
        xml,
        richTextDowngraded: content.richTextDowngraded,
        wroteFormula: !!change.cell?.f,
    };
}

interface SheetPatchOutcome {
    xml: string;
    richTextDowngraded: boolean;
    formulaTouched: boolean;
}

/** 对单个 sheet XML 应用 cell 变更（其余字节保持原样） */
export function patchSheetXml(
    sheetXml: string,
    changes: CellChange[],
    styles: StylesPatcher,
): SheetPatchOutcome {
    let richTextDowngraded = false;
    let formulaTouched = false;

    // sheetData 定位
    let dataStart: number;
    let dataEnd: number;
    let xml = sheetXml;
    const selfClosed = xml.match(/<sheetData\s*\/>/);
    if (selfClosed) {
        const at = selfClosed.index!;
        xml = `${xml.slice(0, at)}<sheetData></sheetData>${xml.slice(at + selfClosed[0].length)}`;
        dataStart = at + '<sheetData>'.length;
        dataEnd = dataStart;
    } else {
        const openMatch = xml.match(/<sheetData\b[^>]*>/);
        const closeIdx = xml.lastIndexOf('</sheetData>');
        if (!openMatch || closeIdx < 0) throw new Error('sheet xml: missing <sheetData>');
        dataStart = openMatch.index! + openMatch[0].length;
        dataEnd = closeIdx;
    }

    const rows = parseRows(xml.slice(dataStart, dataEnd), dataStart);

    // 变更按行分组（r 为 1 起行号）
    const byRow = new Map<number, CellChange[]>();
    for (const c of changes) {
        const list = byRow.get(c.row + 1) ?? [];
        list.push(c);
        byRow.set(c.row + 1, list);
    }

    const edits: TextEdit[] = [];
    for (const [rowNum, rowChanges] of [...byRow.entries()].sort((a, b) => a[0] - b[0])) {
        const existing = rows.find(row => row.r === rowNum);
        if (existing) {
            const patched = patchRowChunk(existing.chunk, rowNum, rowChanges, styles);
            richTextDowngraded ||= patched.richTextDowngraded;
            formulaTouched ||= patched.formulaTouched;
            if (patched.chunk !== existing.chunk) {
                edits.push({ start: existing.start, end: existing.end, text: patched.chunk, order: rowNum });
            }
        } else {
            const built = buildNewRow(rowNum, rowChanges, styles);
            if (!built) continue;
            richTextDowngraded ||= built.richTextDowngraded;
            formulaTouched ||= built.formulaTouched;
            const after = rows.find(row => row.r > rowNum);
            const pos = after ? after.start : dataEnd;
            edits.push({ start: pos, end: pos, text: built.chunk, order: rowNum });
        }
    }

    return { xml: applyEdits(xml, edits), richTextDowngraded, formulaTouched };
}

interface RowBuildOutcome {
    chunk: string;
    richTextDowngraded: boolean;
    formulaTouched: boolean;
}

function buildNewRow(
    rowNum: number,
    changes: CellChange[],
    styles: StylesPatcher,
): RowBuildOutcome | null {
    let richTextDowngraded = false;
    let formulaTouched = false;
    const cells: string[] = [];
    for (const change of [...changes].sort((a, b) => a.col - b.col)) {
        if (change.cell === null) continue; // 删除不存在的 cell：no-op
        const built = buildCell(`${colName(change.col)}${rowNum}`, change, styles, null);
        richTextDowngraded ||= built.richTextDowngraded;
        formulaTouched ||= built.wroteFormula;
        cells.push(built.xml);
    }
    if (!cells.length) return null;
    return { chunk: `<row r="${rowNum}">${cells.join('')}</row>`, richTextDowngraded, formulaTouched };
}

function patchRowChunk(
    rowChunk: string,
    rowNum: number,
    changes: CellChange[],
    styles: StylesPatcher,
): RowBuildOutcome {
    let chunk = rowChunk;
    let richTextDowngraded = false;
    let formulaTouched = false;

    // 自闭合行（只有行属性）需要撑开才能装 cell
    const selfClosed = /^<row\b[^>]*\/>$/.test(chunk);
    if (selfClosed && changes.some(c => c.cell !== null)) {
        chunk = `${chunk.slice(0, chunk.length - 2)}></row>`;
    }
    const bodyEnd = chunk.length - '</row>'.length;

    const cells = parseCells(chunk);
    const edits: TextEdit[] = [];
    for (const change of [...changes].sort((a, b) => a.col - b.col)) {
        const ref = `${colName(change.col)}${rowNum}`;
        const existing = cells.find(c => c.col === change.col) ?? null;
        if (change.cell === null) {
            if (existing) edits.push({ start: existing.start, end: existing.end, text: '', order: change.col });
            continue;
        }
        const built = buildCell(ref, change, styles, existing);
        richTextDowngraded ||= built.richTextDowngraded;
        formulaTouched ||= built.wroteFormula || !!existing?.inner.includes('<f');
        if (existing) {
            edits.push({ start: existing.start, end: existing.end, text: built.xml, order: change.col });
        } else {
            const after = cells.find(c => c.col > change.col);
            const pos = after ? after.start : bodyEnd;
            edits.push({ start: pos, end: pos, text: built.xml, order: change.col });
        }
    }

    return { chunk: applyEdits(chunk, edits), richTextDowngraded, formulaTouched };
}

// ------------------------------------------------------------ workbook 装配

/** workbook.xml + rels → sheet 名（已反转义）→ zip 内部件路径 */
function resolveSheetParts(workbookXml: string, relsXml: string): Map<string, string> {
    const relTargets = new Map<string, string>();
    for (const m of relsXml.matchAll(/<Relationship\b[^>]*/g)) {
        const id = m[0].match(/\bId="([^"]+)"/)?.[1];
        const target = m[0].match(/\bTarget="([^"]+)"/)?.[1];
        if (!id || !target) continue;
        const normalized = target.replace(/^\//, '').replace(/^\/?xl\//, '');
        relTargets.set(id, `xl/${normalized}`);
    }
    const map = new Map<string, string>();
    for (const m of workbookXml.matchAll(/<sheet\b[^>]*/g)) {
        const name = m[0].match(/\bname="([^"]*)"/)?.[1];
        const rid = m[0].match(/\br:id="([^"]+)"/)?.[1] ?? m[0].match(/\bd\d*:id="([^"]+)"/)?.[1];
        if (name == null || !rid) continue;
        const target = relTargets.get(rid);
        if (target) map.set(unescapeXml(name), target);
    }
    return map;
}

/** calcPr 补 fullCalcOnLoad="1"（无 calcPr 时不动 workbook.xml） */
function patchCalcPr(workbookXml: string): string {
    const m = workbookXml.match(/<calcPr\b[^>]*\/?>/);
    if (!m) return workbookXml;
    let tag = m[0];
    if (/\bfullCalcOnLoad="/.test(tag)) {
        tag = tag.replace(/\bfullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
    } else {
        tag = tag.replace(/\/?>$/, suffix => ` fullCalcOnLoad="1"${suffix}`);
    }
    return tag === m[0] ? workbookXml : workbookXml.replace(m[0], tag);
}

/**
 * 在原始 xlsx/xlsm 字节上应用 diff。除被改 sheet 的 XML、（必要时的）
 * styles.xml 追加与 calcChain 删除之外，所有部件内容逐字节保持原样。
 *
 * @param sheetNames baseline 的 sheetId → sheet 名（重命名被禁用，与原文件一致）
 */
export async function patchXlsxBytes(
    originalBuffer: ArrayBuffer,
    diff: WorkbookDiff,
    sheetNames: Record<string, string>,
): Promise<XmlPatchResult> {
    assertPatchableDiff(diff);

    const zip = await JSZip.loadAsync(originalBuffer);
    const readPart = async (name: string): Promise<string> => {
        const file = zip.file(name);
        if (!file) throw new Error(`xlsx part missing: ${name}`);
        return file.async('string');
    };

    const workbookXml = await readPart('xl/workbook.xml');
    const relsXml = await readPart('xl/_rels/workbook.xml.rels');
    const sheetParts = resolveSheetParts(workbookXml, relsXml);

    const stylesPatcher = new StylesPatcher(await readPart('xl/styles.xml'));

    let richTextDowngraded = false;
    let formulaTouched = false;
    let anyValueChanged = false;
    const changedParts: string[] = [];
    const removedParts: string[] = [];
    const writePart = (name: string, content: string) => {
        // createFolders:false —— 不新增原文件没有的目录条目（审计干净）
        zip.file(name, content, { createFolders: false });
        changedParts.push(name);
    };

    for (const sheetDiff of diff.sheets) {
        if (!sheetDiff.cellChanges.length) continue;
        const name = sheetNames[sheetDiff.sheetId];
        const partPath = name != null ? sheetParts.get(name) : undefined;
        if (!partPath) {
            throw new Error(`cannot resolve sheet part for "${name ?? sheetDiff.sheetId}"`);
        }
        const sheetXml = await readPart(partPath);
        const outcome = patchSheetXml(sheetXml, sheetDiff.cellChanges, stylesPatcher);
        richTextDowngraded ||= outcome.richTextDowngraded;
        formulaTouched ||= outcome.formulaTouched || sheetXml.includes('<f>') || sheetXml.includes('<f ');
        anyValueChanged ||= sheetDiff.cellChanges.some(c => c.valueChanged);
        if (outcome.xml !== sheetXml) writePart(partPath, outcome.xml);
    }

    if (stylesPatcher.touched) {
        writePart('xl/styles.xml', stylesPatcher.serialize());
    }

    // 公式缓存可能陈旧：删 calcChain（Excel 自动重建）+ fullCalcOnLoad 强制重算
    const hasCalcChain = !!zip.file('xl/calcChain.xml');
    if (anyValueChanged && (hasCalcChain || formulaTouched)) {
        if (hasCalcChain) {
            zip.remove('xl/calcChain.xml');
            removedParts.push('xl/calcChain.xml');
            const contentTypes = await readPart('[Content_Types].xml');
            const cleanedTypes = contentTypes.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
            if (cleanedTypes !== contentTypes) writePart('[Content_Types].xml', cleanedTypes);
            const cleanedRels = relsXml.replace(/<Relationship\b[^>]*Target="calcChain\.xml"[^>]*\/>/, '');
            if (cleanedRels !== relsXml) writePart('xl/_rels/workbook.xml.rels', cleanedRels);
        }
        const patchedWorkbook = patchCalcPr(workbookXml);
        if (patchedWorkbook !== workbookXml) writePart('xl/workbook.xml', patchedWorkbook);
    }

    const bytes = await zip.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
    });
    return { bytes, richTextDowngraded, changedParts: changedParts.sort(), removedParts };
}
