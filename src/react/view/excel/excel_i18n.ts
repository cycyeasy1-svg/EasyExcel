/**
 * 独立 i18n（外壳 UI 文案）。表格引擎内部文案由 Univer 自带 locale 提供，
 * 这里只覆盖 Excel.tsx 的对话框/横幅/控件字符串。
 * de/nl 等其余语言回退 en（引擎 UI 仍是对应语言的 Univer locale）。
 */
import { getConfigs } from '../../util/vscodeConfig.ts';

type Messages = Record<string, string>;

const en: Messages = {
    'button.cancel': 'Cancel',
    'button.save': 'Save',
    'button.saveAs': 'Save As',
    'viewer.readonlyBanner': 'Read-only — browse and find; use Save As to save changes',
    'viewer.formatCannotPreserveTitle': 'Formatting cannot be preserved',
    'viewer.formatCannotPreserveContent': '{} cannot preserve styles, merges, formulas, etc. Save as xlsx instead?',
    'viewer.saveAsXlsx': 'Save as xlsx',
    'viewer.saveAsOriginal': 'Save as original',
    'viewer.chooseExportFormat': 'Choose export format',
    'viewer.exportXlsxLabel': 'Excel Workbook (.xlsx)',
    'viewer.exportXlsxDesc': 'Supports styles, formulas, merged cells',
    'viewer.exportCsvLabel': 'CSV (.csv)',
    'viewer.exportCsvDesc': 'Plain text, first sheet only',
    'viewer.exportXlsLabel': 'Excel 97-2003 (.xls)',
    'viewer.exportXlsDesc': 'Legacy Excel format',
    'viewer.exportOdsLabel': 'OpenDocument (.ods)',
    'viewer.exportOdsDesc': 'Compatible with LibreOffice / WPS',
    'viewer.switchToLightMode': 'Switch to light mode',
    'viewer.switchToDarkMode': 'Switch to dark mode',
    'viewer.zoom': 'Zoom',
    'viewer.zoomOut': 'Zoom out',
    'viewer.zoomIn': 'Zoom in',
    'viewer.lossyTitle': 'Some content cannot be preserved',
    'viewer.lossyContent': 'This file contains charts, pivot tables or macros. Exporting a copy will drop them (the original file is untouched).',
    'viewer.lossySaveAnyway': 'Save anyway',
    'viewer.saveFailed': 'Save failed, see developer console',
    'viewer.structuralBlocked': 'Please do this in Excel — structural changes (rows/columns/sheets, merges, sizes, freeze, validation, conditional formatting) are disabled so saving never alters untouched content',
    'viewer.patchBlockedSave': 'This edit cannot be saved without altering untouched content. Undo it and retry, or make the change in Excel',
    'viewer.richTextDowngraded': 'Rich-text cells were saved as plain text',
};

const zhCn: Messages = {
    'button.cancel': '取消',
    'button.save': '保存',
    'button.saveAs': '另存为',
    'viewer.readonlyBanner': '只读模式 — 可浏览与查找，请使用「另存为」保存更改',
    'viewer.formatCannotPreserveTitle': '格式将无法保存',
    'viewer.formatCannotPreserveContent': '{} 格式不支持保存样式、合并单元格、公式等富格式内容。是否另存为 xlsx？',
    'viewer.saveAsXlsx': '另存为 xlsx',
    'viewer.saveAsOriginal': '仍保存原格式',
    'viewer.chooseExportFormat': '选择导出格式',
    'viewer.exportXlsxLabel': 'Excel 工作簿 (.xlsx)',
    'viewer.exportXlsxDesc': '支持样式、公式、合并单元格',
    'viewer.exportCsvLabel': 'CSV (.csv)',
    'viewer.exportCsvDesc': '纯文本，仅首个工作表',
    'viewer.exportXlsLabel': 'Excel 97-2003 (.xls)',
    'viewer.exportXlsDesc': '旧版 Excel 格式',
    'viewer.exportOdsLabel': 'OpenDocument (.ods)',
    'viewer.exportOdsDesc': '兼容 LibreOffice / WPS',
    'viewer.switchToLightMode': '切换到浅色模式',
    'viewer.switchToDarkMode': '切换到深色模式',
    'viewer.zoom': '缩放',
    'viewer.zoomOut': '缩小',
    'viewer.zoomIn': '放大',
    'viewer.lossyTitle': '部分内容无法保留',
    'viewer.lossyContent': '此文件包含图表、透视表或宏。导出的副本将丢失这些内容（原文件不受影响）。',
    'viewer.lossySaveAnyway': '仍要保存',
    'viewer.saveFailed': '保存失败，详见开发者工具控制台',
    'viewer.structuralBlocked': '此操作请在 Excel 中进行 —— 为保证保存不改动未编辑的内容，插删行列/工作表、合并、行高列宽、冻结、数据验证、条件格式已禁用',
    'viewer.patchBlockedSave': '本次编辑包含无法无损保存的修改，请撤销后重试，或在 Excel 中进行该修改',
    'viewer.richTextDowngraded': '富文本单元格已按纯文本保存',
};

const zhTw: Messages = {
    'button.cancel': '取消',
    'button.save': '儲存',
    'button.saveAs': '另存新檔',
    'viewer.readonlyBanner': '唯讀模式 — 可瀏覽與尋找，請使用「另存新檔」儲存變更',
    'viewer.formatCannotPreserveTitle': '格式將無法保留',
    'viewer.formatCannotPreserveContent': '{} 格式不支援儲存樣式、合併儲存格、公式等豐富格式內容。是否另存為 xlsx？',
    'viewer.saveAsXlsx': '另存為 xlsx',
    'viewer.saveAsOriginal': '仍以原格式儲存',
    'viewer.chooseExportFormat': '選擇匯出格式',
    'viewer.exportXlsxLabel': 'Excel 活頁簿 (.xlsx)',
    'viewer.exportXlsxDesc': '支援樣式、公式、合併儲存格',
    'viewer.exportCsvLabel': 'CSV (.csv)',
    'viewer.exportCsvDesc': '純文字，僅第一個工作表',
    'viewer.exportXlsLabel': 'Excel 97-2003 (.xls)',
    'viewer.exportXlsDesc': '舊版 Excel 格式',
    'viewer.exportOdsLabel': 'OpenDocument (.ods)',
    'viewer.exportOdsDesc': '相容 LibreOffice / WPS',
    'viewer.switchToLightMode': '切換至淺色模式',
    'viewer.switchToDarkMode': '切換至深色模式',
    'viewer.zoom': '縮放',
    'viewer.zoomOut': '縮小',
    'viewer.zoomIn': '放大',
    'viewer.lossyTitle': '部分內容無法保留',
    'viewer.lossyContent': '此檔案包含圖表、樞紐分析表或巨集。匯出的副本將遺失這些內容（原始檔案不受影響）。',
    'viewer.lossySaveAnyway': '仍要儲存',
    'viewer.saveFailed': '儲存失敗，詳見開發者工具主控台',
    'viewer.structuralBlocked': '此操作請在 Excel 中進行 —— 為確保儲存不更動未編輯的內容，插刪列欄/工作表、合併、列高欄寬、凍結、資料驗證、設定格式化的條件已停用',
    'viewer.patchBlockedSave': '本次編輯包含無法無損儲存的修改，請復原後重試，或在 Excel 中進行該修改',
    'viewer.richTextDowngraded': '富文字儲存格已以純文字儲存',
};

const TRADITIONAL_ZH = new Set(['zh-tw', 'zh-hk', 'zh-mo', 'zh-hant']);
const SIMPLIFIED_ZH = new Set(['zh-cn', 'zh-sg', 'zh-hans', 'zh']);

function resolveMessages(vscodeLang: string): Messages {
    const lower = (vscodeLang || 'en').toLowerCase();
    if (SIMPLIFIED_ZH.has(lower)) return zhCn;
    if (TRADITIONAL_ZH.has(lower)) return zhTw;
    return en;
}

let messages: Messages = en;
let initialized = false;

export function initExcelLocale(): void {
    if (initialized) return;
    messages = resolveMessages(getConfigs()?.language ?? 'en');
    initialized = true;
}

/** Translate a locale key; `{}` placeholders are filled in order. */
export function t(key: string, ...args: (string | number)[]): string {
    if (!initialized) initExcelLocale();
    let text = messages[key] ?? en[key] ?? key;
    for (const arg of args) {
        text = text.replace('{}', String(arg));
    }
    return text;
}
