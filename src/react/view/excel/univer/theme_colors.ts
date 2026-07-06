/**
 * OOXML 主题色 / 索引色 / tint 解析。
 *
 * xlsx 中颜色有四种形态：argb 字面色、theme 索引（可带 tint）、indexed
 * 传统索引、auto。现有 excel_styles.colorToHex 只认 argb，其余全部丢弃——
 * 这是渲染保真的最大缺口，本模块补齐后由 univer/import.ts 统一调用。
 */

/**
 * Office 默认主题调色板，按 ECMA-376 color.theme 属性的索引序排列。
 * 注意 0/1 与 2/3 相对 theme1.xml 文档顺序（dk1,lt1,dk2,lt2）是互换的：
 * theme="0" 是 lt1(白)、theme="1" 是 dk1(黑) —— Excel 默认黑色文字即 theme 1。
 */
export const OFFICE_DEFAULT_THEME: string[] = [
    'FFFFFF', // 0 lt1
    '000000', // 1 dk1
    'E7E6E6', // 2 lt2
    '44546A', // 3 dk2
    '4472C4', // 4 accent1
    'ED7D31', // 5 accent2
    'A5A5A5', // 6 accent3
    'FFC000', // 7 accent4
    '5B9BD5', // 8 accent5
    '70AD47', // 9 accent6
    '0563C1', // 10 hlink
    '954F72', // 11 folHlink
];

/** theme1.xml clrScheme 子元素的文档顺序 → theme 索引 */
const CLR_SCHEME_ORDER_TO_INDEX = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];
const CLR_SCHEME_TAGS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** 传统 64 色索引调色板（ECMA-376 18.8.27），64/65 为系统前景/背景色 */
export const INDEXED_COLORS: string[] = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
    '000000', 'FFFFFF',
];

/** 解析 theme1.xml 的 clrScheme，失败时回退 Office 默认主题 */
export function parseThemePalette(themeXml?: string): string[] {
    if (!themeXml) return OFFICE_DEFAULT_THEME;
    try {
        const doc = new DOMParser().parseFromString(themeXml, 'application/xml');
        // DOMParser 对命名空间的容忍度不一，用 localName 匹配
        const all = doc.getElementsByTagName('*');
        let scheme: Element | null = null;
        for (let i = 0; i < all.length; i += 1) {
            if (all[i].localName === 'clrScheme') {
                scheme = all[i];
                break;
            }
        }
        if (!scheme) return OFFICE_DEFAULT_THEME;
        const palette = [...OFFICE_DEFAULT_THEME];
        for (let t = 0; t < CLR_SCHEME_TAGS.length; t += 1) {
            const tag = CLR_SCHEME_TAGS[t];
            let node: Element | null = null;
            for (let i = 0; i < scheme.children.length; i += 1) {
                if (scheme.children[i].localName === tag) {
                    node = scheme.children[i];
                    break;
                }
            }
            if (!node) continue;
            const hex = readSchemeColor(node);
            if (hex) palette[CLR_SCHEME_ORDER_TO_INDEX[t]] = hex;
        }
        return palette;
    } catch {
        return OFFICE_DEFAULT_THEME;
    }
}

function readSchemeColor(node: Element): string | undefined {
    for (let i = 0; i < node.children.length; i += 1) {
        const child = node.children[i];
        if (child.localName === 'srgbClr') {
            const val = child.getAttribute('val');
            if (val && /^[0-9a-fA-F]{6}$/.test(val)) return val.toUpperCase();
        }
        if (child.localName === 'sysClr') {
            const val = child.getAttribute('lastClr');
            if (val && /^[0-9a-fA-F]{6}$/.test(val)) return val.toUpperCase();
        }
    }
    return undefined;
}

/**
 * OOXML tint 算法（ECMA-376 18.3.1.15）：在 HSL 亮度通道上运算。
 * tint < 0 加深：L' = L * (1 + tint)；tint > 0 提亮：L' = L * (1 - tint) + tint。
 */
export function applyTint(hex6: string, tint: number): string {
    if (!tint) return hex6.toUpperCase();
    const r = parseInt(hex6.slice(0, 2), 16) / 255;
    const g = parseInt(hex6.slice(2, 4), 16) / 255;
    const b = parseInt(hex6.slice(4, 6), 16) / 255;
    const { h, s, l } = rgbToHsl(r, g, b);
    const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const { r: r2, g: g2, b: b2 } = hslToRgb(h, s, Math.min(1, Math.max(0, l2)));
    return [r2, g2, b2]
        .map(v => Math.round(v * 255).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number) {
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return { h: 0, s: 0, l };
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h: number;
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    return { h: h / 6, s, l };
}

function hslToRgb(h: number, s: number, l: number) {
    if (s === 0) return { r: l, g: l, b: l };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return {
        r: hueToRgb(p, q, h + 1 / 3),
        g: hueToRgb(p, q, h),
        b: hueToRgb(p, q, h - 1 / 3),
    };
}

function hueToRgb(p: number, q: number, t0: number) {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
}

export interface ExcelColorLike {
    argb?: string;
    theme?: number;
    tint?: number;
    indexed?: number;
}

/**
 * 统一解析 ExcelJS 颜色对象 → `#rrggbb`。
 * 依次尝试 argb → theme(+tint) → indexed；解析不出返回 undefined。
 */
export function resolveExcelColor(color: ExcelColorLike | undefined, palette: string[]): string | undefined {
    if (!color) return undefined;
    if (color.argb) {
        const argb = color.argb.replace(/^#/, '');
        if (argb.length === 8) return `#${argb.slice(2).toLowerCase()}`;
        if (argb.length === 6) return `#${argb.toLowerCase()}`;
        return undefined;
    }
    if (color.theme != null) {
        const base = palette[color.theme] ?? OFFICE_DEFAULT_THEME[color.theme];
        if (!base) return undefined;
        return `#${applyTint(base, color.tint ?? 0).toLowerCase()}`;
    }
    if (color.indexed != null) {
        const hex = INDEXED_COLORS[color.indexed];
        return hex ? `#${hex.toLowerCase()}` : undefined;
    }
    return undefined;
}

/** 从 ExcelJS workbook.model.themes 提取 theme1.xml 字符串（形态因 fork 而异，做防御） */
export function extractThemeXml(workbookModel: unknown): string | undefined {
    const themes = (workbookModel as { themes?: unknown })?.themes;
    if (!themes) return undefined;
    if (typeof themes === 'string') return themes;
    const theme1 = (themes as Record<string, unknown>).theme1;
    return typeof theme1 === 'string' ? theme1 : undefined;
}
