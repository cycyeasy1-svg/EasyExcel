/**
 * 独立 i18n（外壳 UI 文案）。表格引擎内部文案由 Univer 自带 locale 提供，
 * 这里只覆盖 Excel.tsx 的对话框/横幅/控件字符串。
 * 语言由用户在界面右上角选择并持久化（默认简体中文），不跟随 VSCode 显示语言。
 */
import { loadLanguage } from '../../util/vscode.ts';

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
    'viewer.saved': 'Saved',
    'viewer.structuralBlocked': 'Please do this in Excel — structural changes (rows/columns/sheets, merges, sizes, freeze, validation, conditional formatting) are disabled so saving never alters untouched content',
    'viewer.patchBlockedSave': 'This edit cannot be saved without altering untouched content. Undo it and retry, or make the change in Excel',
    'viewer.richTextDowngraded': 'Rich-text cells were saved as plain text',
    'viewer.language': 'Language',
    'viewer.langSwitchTitle': 'Switch language',
    'viewer.langSwitchContent': 'Switching language reloads the spreadsheet. Unsaved changes will be lost.',
    'viewer.langSwitchAnyway': 'Switch anyway',
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
    'viewer.saved': '已保存',
    'viewer.structuralBlocked': '此操作请在 Excel 中进行 —— 为保证保存不改动未编辑的内容，插删行列/工作表、合并、行高列宽、冻结、数据验证、条件格式已禁用',
    'viewer.patchBlockedSave': '本次编辑包含无法无损保存的修改，请撤销后重试，或在 Excel 中进行该修改',
    'viewer.richTextDowngraded': '富文本单元格已按纯文本保存',
    'viewer.language': '语言',
    'viewer.langSwitchTitle': '切换语言',
    'viewer.langSwitchContent': '切换语言将重新加载表格，未保存的更改会丢失。',
    'viewer.langSwitchAnyway': '仍要切换',
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
    'viewer.saved': '已儲存',
    'viewer.structuralBlocked': '此操作請在 Excel 中進行 —— 為確保儲存不更動未編輯的內容，插刪列欄/工作表、合併、列高欄寬、凍結、資料驗證、設定格式化的條件已停用',
    'viewer.patchBlockedSave': '本次編輯包含無法無損儲存的修改，請復原後重試，或在 Excel 中進行該修改',
    'viewer.richTextDowngraded': '富文字儲存格已以純文字儲存',
    'viewer.language': '語言',
    'viewer.langSwitchTitle': '切換語言',
    'viewer.langSwitchContent': '切換語言將重新載入表格，未儲存的變更會遺失。',
    'viewer.langSwitchAnyway': '仍要切換',
};

const ja: Messages = {
    'button.cancel': 'キャンセル',
    'button.save': '保存',
    'button.saveAs': '名前を付けて保存',
    'viewer.readonlyBanner': '読み取り専用モード — 閲覧と検索が可能です。変更を保存するには「名前を付けて保存」をご利用ください',
    'viewer.formatCannotPreserveTitle': '書式を保存できません',
    'viewer.formatCannotPreserveContent': '{} 形式ではスタイル・結合セル・数式などの書式を保存できません。xlsx として保存しますか？',
    'viewer.saveAsXlsx': 'xlsx で保存',
    'viewer.saveAsOriginal': '元の形式で保存',
    'viewer.chooseExportFormat': 'エクスポート形式を選択',
    'viewer.exportXlsxLabel': 'Excel ブック (.xlsx)',
    'viewer.exportXlsxDesc': 'スタイル・数式・結合セルに対応',
    'viewer.exportCsvLabel': 'CSV (.csv)',
    'viewer.exportCsvDesc': 'プレーンテキスト、先頭シートのみ',
    'viewer.exportXlsLabel': 'Excel 97-2003 (.xls)',
    'viewer.exportXlsDesc': '旧形式の Excel',
    'viewer.exportOdsLabel': 'OpenDocument (.ods)',
    'viewer.exportOdsDesc': 'LibreOffice / WPS と互換',
    'viewer.switchToLightMode': 'ライトモードに切り替え',
    'viewer.switchToDarkMode': 'ダークモードに切り替え',
    'viewer.zoom': 'ズーム',
    'viewer.zoomOut': '縮小',
    'viewer.zoomIn': '拡大',
    'viewer.lossyTitle': '一部の内容を保持できません',
    'viewer.lossyContent': 'このファイルにはグラフ、ピボットテーブル、またはマクロが含まれています。コピーをエクスポートするとこれらは失われます（元のファイルは変更されません）。',
    'viewer.lossySaveAnyway': 'このまま保存',
    'viewer.saveFailed': '保存に失敗しました。開発者コンソールをご確認ください',
    'viewer.saved': '保存しました',
    'viewer.structuralBlocked': 'この操作は Excel で行ってください —— 保存時に未編集の内容が変わらないよう、行・列やシートの挿入/削除、結合、行高・列幅、固定、データの入力規則、条件付き書式は無効化されています',
    'viewer.patchBlockedSave': 'この編集は無損失で保存できません。元に戻してからやり直すか、Excel で編集してください',
    'viewer.richTextDowngraded': 'リッチテキストのセルはプレーンテキストとして保存されました',
    'viewer.language': '言語',
    'viewer.langSwitchTitle': '言語を切り替え',
    'viewer.langSwitchContent': '言語を切り替えると表が再読み込みされ、未保存の変更は失われます。',
    'viewer.langSwitchAnyway': 'それでも切り替える',
};

const TRADITIONAL_ZH = new Set(['zh-tw', 'zh-hk', 'zh-mo', 'zh-hant']);
const SIMPLIFIED_ZH = new Set(['zh-cn', 'zh-sg', 'zh-hans', 'zh']);

function resolveMessages(vscodeLang: string): Messages {
    const lower = (vscodeLang || 'en').toLowerCase();
    if (SIMPLIFIED_ZH.has(lower)) return zhCn;
    if (TRADITIONAL_ZH.has(lower)) return zhTw;
    if (lower.startsWith('ja')) return ja;
    return en;
}

let messages: Messages = en;
let initialized = false;

export function initExcelLocale(): void {
    if (initialized) return;
    messages = resolveMessages(loadLanguage());
    initialized = true;
}

/** 用户在界面上切换语言时调用（外壳文案即时生效，引擎 locale 需重建 Univer） */
export function setExcelLocale(language: string): void {
    messages = resolveMessages(language);
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
