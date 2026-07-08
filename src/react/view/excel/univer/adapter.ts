/**
 * Univer 生命周期封装：preset 组装、创建/销毁、只读、暗色、缩放、
 * 视图状态（activeSheet + 选区）持久化与恢复、超链接应用、dirty 事件桥。
 * Excel.tsx 与 UniverSpike 共用此模块；所有 Univer API 经此收口（R2 缓解）。
 */
import { createUniver, LocaleType, mergeLocales, type IWorkbookData } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import { UniverSheetsFilterPreset } from '@univerjs/presets/preset-sheets-filter';
import { UniverSheetsSortPreset } from '@univerjs/presets/preset-sheets-sort';
import { UniverSheetsFindReplacePreset } from '@univerjs/presets/preset-sheets-find-replace';
import { UniverSheetsDataValidationPreset } from '@univerjs/presets/preset-sheets-data-validation';
import { UniverSheetsHyperLinkPreset } from '@univerjs/presets/preset-sheets-hyper-link';
import { UniverSheetsDrawingPreset } from '@univerjs/presets/preset-sheets-drawing';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/presets/preset-sheets-conditional-formatting';
import CoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';
import CoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import CoreJaJP from '@univerjs/preset-sheets-core/locales/ja-JP';
import FilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN';
import FilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import FilterJaJP from '@univerjs/preset-sheets-filter/locales/ja-JP';
import SortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN';
import SortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import SortJaJP from '@univerjs/preset-sheets-sort/locales/ja-JP';
import FindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN';
import FindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import FindReplaceJaJP from '@univerjs/preset-sheets-find-replace/locales/ja-JP';
import DataValidationZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN';
import DataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import DataValidationJaJP from '@univerjs/preset-sheets-data-validation/locales/ja-JP';
import HyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN';
import HyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import HyperLinkJaJP from '@univerjs/preset-sheets-hyper-link/locales/ja-JP';
import DrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN';
import DrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import DrawingJaJP from '@univerjs/preset-sheets-drawing/locales/ja-JP';
import CfZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN';
import CfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';
import CfJaJP from '@univerjs/preset-sheets-conditional-formatting/locales/ja-JP';

import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/presets/lib/styles/preset-sheets-filter.css';
import '@univerjs/presets/lib/styles/preset-sheets-sort.css';
import '@univerjs/presets/lib/styles/preset-sheets-find-replace.css';
import '@univerjs/presets/lib/styles/preset-sheets-data-validation.css';
import '@univerjs/presets/lib/styles/preset-sheets-hyper-link.css';
import '@univerjs/presets/lib/styles/preset-sheets-drawing.css';
import '@univerjs/presets/lib/styles/preset-sheets-conditional-formatting.css';

import type { UniverLoadResult } from './loader';

export interface UniverViewState {
    sheetIndex: number;
    ri: number;
    ci: number;
}

export interface UniverAdapterOptions {
    darkMode?: boolean;
    readOnly?: boolean;
    language?: string;
    /** 外链打开回调（webview 中 window.open 不可用，转宿主打开） */
    onOpenExternal?: (url: string) => void;
}

export interface UniverEditSession {
    /** 发生过结构性变更（行列插删/移动）的 sheetId 集合 → 这些 sheet 保存时整表重建 */
    structuralSheetIds: Set<string>;
    /** 保存成功后调用：清空结构日志（baseline 已更新） */
    reset(): void;
    stop(): void;
}

/** 行列插删/移动类 mutation：使该 sheet 的逐格 diff 失效（注意 Univer 的
 *  命名不对称：insert-row/remove-rows、remove-col） */
const STRUCTURAL_MUTATION = /^sheet\.mutation\.(insert-(row|col)|remove-(rows|col)|move-(rows|columns|range)|reorder-range)$/;

/**
 * M5：xlsx 原文件走 XML 补丁保存，以下操作无法无损写回 → 命令级拦截
 * （BeforeCommandExecute 里 cancel，Univer 抛 CanceledError 中止命令与
 * 其全部 mutation）。粘贴等直接派发 mutation 的路径拦不住，保存时由
 * assertPatchableDiff 兜底拒绝。
 *
 * 覆盖：行列插删/移动/排序、范围插删平移、分列、sheet 增删/复制/重命名/
 * 排序，以及初版一并禁用的行高列宽/隐藏、合并、冻结、DV、条件格式。
 */
const RESTRICTED_COMMAND = new RegExp(`^(?:${[
    // 行列插删/移动/排序/分列（含 -confirm/-ctx/-before/-after 等变体）
    'sheet\\.command\\.(?:insert|remove)-(?:row|col|multi-rows|multi-cols)[a-z-]*',
    'sheet\\.command\\.append-row',
    'sheet\\.command\\.move-(?:rows|cols|range)',
    'sheet\\.command\\.reorder-range',
    'sheet\\.command\\.sort-range[a-z-]*',
    'sheet\\.command\\.split-text-to-columns',
    // 范围插删（cell 平移）
    'sheet\\.command\\.(?:insert|delete)-range-move-[a-z-]*',
    // sheet 级结构
    'sheet\\.command\\.(?:insert-sheet|remove-sheet|copy-sheet|set-worksheet-name|set-worksheet-order)[a-z-]*',
    // 行高列宽/行列隐藏（二期再开）
    'sheet\\.command\\.(?:set|delta)-(?:row-height|column-width)[a-z-]*',
    'sheet\\.command\\.set-worksheet-(?:row-height|col-width)[a-z-]*',
    'sheet\\.command\\.set-(?:row-is-auto-height|col-auto-width)',
    'sheet\\.command\\.set-[a-z-]*(?:hidden|visible)',
    'sheet\\.command\\.hide-(?:row|col)[a-z-]*',
    // 合并/冻结（二期再开）
    'sheet\\.command\\.(?:add|remove)-worksheet-merge[a-z-]*',
    'sheet\\.command\\.[a-z-]*frozen[a-z-]*',
    // 数据验证 / 条件格式（资源级重写只保留在另存为路径）
    'sheet\\.command\\.addDataValidation',
    '[a-z]+\\.command\\.[a-z-]*data-validation[a-z-]*',
    'sheet\\.command\\.[a-z-]*conditional-rule[a-z-]*',
].join('|')})$`);

/** 应用超链接的上限，防止极端文件逐链接跑命令拖慢加载 */
const MAX_APPLIED_HYPERLINKS = 2000;

type FUniverLike = ReturnType<typeof createUniver>['univerAPI'];

const pickLocale = (language?: string) => {
    const lang = language ?? '';
    if (/^zh/i.test(lang)) return LocaleType.ZH_CN;
    if (/^ja/i.test(lang)) return LocaleType.JA_JP;
    return LocaleType.EN_US;
};

export class UniverAdapter {
    private disposed = false;

    private constructor(
        readonly univer: ReturnType<typeof createUniver>['univer'],
        readonly univerAPI: FUniverLike,
    ) { }

    static create(container: HTMLElement, opts: UniverAdapterOptions): UniverAdapter {
        const { univer, univerAPI } = createUniver({
            locale: pickLocale(opts.language),
            darkMode: opts.darkMode,
            locales: {
                [LocaleType.ZH_CN]: mergeLocales(
                    CoreZhCN, FilterZhCN, SortZhCN, FindReplaceZhCN,
                    DataValidationZhCN, HyperLinkZhCN, DrawingZhCN, CfZhCN,
                ),
                [LocaleType.EN_US]: mergeLocales(
                    CoreEnUS, FilterEnUS, SortEnUS, FindReplaceEnUS,
                    DataValidationEnUS, HyperLinkEnUS, DrawingEnUS, CfEnUS,
                ),
                [LocaleType.JA_JP]: mergeLocales(
                    CoreJaJP, FilterJaJP, SortJaJP, FindReplaceJaJP,
                    DataValidationJaJP, HyperLinkJaJP, DrawingJaJP, CfJaJP,
                ),
            },
            presets: [
                UniverSheetsCorePreset({ container }),
                UniverSheetsFilterPreset(),
                UniverSheetsSortPreset(),
                UniverSheetsFindReplacePreset(),
                UniverSheetsDataValidationPreset(),
                UniverSheetsHyperLinkPreset(),
                UniverSheetsDrawingPreset(),
                UniverSheetsConditionalFormattingPreset(),
            ],
        });
        return new UniverAdapter(univer, univerAPI);
    }

    async loadWorkbook(result: UniverLoadResult, opts: UniverAdapterOptions) {
        if (opts.onOpenExternal) {
            this.patchWindowOpen(opts.onOpenExternal);
        }
        const fWorkbook = this.univerAPI.createWorkbook(result.workbookData as IWorkbookData);
        await this.applyHyperlinks(result);
        await this.applySheetFeatures(result);
        if (opts.readOnly) {
            fWorkbook.setEditable(false);
        }
        return fWorkbook;
    }

    /** hyperlink-ui 用 window.open 打开外链；webview 中不可用，转宿主 */
    private patchWindowOpen(onOpenExternal: (url: string) => void) {
        const original = window.open.bind(window);
        window.open = ((url?: string | URL, target?: string, features?: string) => {
            const href = url == null ? '' : String(url);
            if (/^(https?:|mailto:)/i.test(href)) {
                onOpenExternal(href);
                return null;
            }
            return original(url as never, target, features);
        }) as typeof window.open;
    }

    private async applySheetFeatures(result: UniverLoadResult) {
        const features = result.sheetFeatures;
        if (!features) return;
        const fWorkbook = this.univerAPI.getActiveWorkbook();
        if (!fWorkbook) return;
        const { applyValidationsToSheet, applyImagesToSheet, applyProtectionToSheet } = await import('./features');
        for (const [sheetId, f] of Object.entries(features)) {
            const sheet = fWorkbook.getSheetBySheetId(sheetId);
            if (!sheet) continue;
            const sheetLike = sheet as never;
            if (f.validations.length) {
                await applyValidationsToSheet(this.univerAPI as never, sheetLike, f.validations);
            }
            if (f.images.length) {
                await applyImagesToSheet(sheetLike, f.images);
            }
            if (f.protected) {
                await applyProtectionToSheet(sheetLike);
            }
        }
    }

    private async applyHyperlinks(result: UniverLoadResult) {
        const links = result.hyperlinks.slice(0, MAX_APPLIED_HYPERLINKS);
        if (result.hyperlinks.length > links.length) {
            console.warn(`EasyExcel: ${result.hyperlinks.length - links.length} hyperlinks skipped (limit ${MAX_APPLIED_HYPERLINKS})`);
        }
        const fWorkbook = this.univerAPI.getActiveWorkbook();
        if (!fWorkbook) return;
        for (const link of links) {
            try {
                const sheet = fWorkbook.getSheetBySheetId(link.sheetId);
                if (!sheet) continue;
                const range = sheet.getRange(link.row, link.column);
                await (range as never as { setHyperLink(url: string, label?: string): Promise<boolean> })
                    .setHyperLink(link.url, link.display);
            } catch (e) {
                console.warn('EasyExcel: failed to apply hyperlink', link, e);
            }
        }
    }

    setDarkMode(dark: boolean) {
        if (this.disposed) return;
        this.univerAPI.toggleDarkMode(dark);
    }

    setZoom(ratio: number) {
        if (this.disposed) return;
        const sheet = this.univerAPI.getActiveWorkbook()?.getActiveSheet();
        (sheet as never as { zoom(r: number): void })?.zoom(ratio);
    }

    getViewState(): UniverViewState | null {
        if (this.disposed) return null;
        const fWorkbook = this.univerAPI.getActiveWorkbook();
        if (!fWorkbook) return null;
        const sheet = fWorkbook.getActiveSheet();
        const snapshot = fWorkbook.getSnapshot();
        const sheetIndex = snapshot.sheetOrder.indexOf(sheet.getSheetId());
        const selection = fWorkbook.getActiveRange();
        return {
            sheetIndex: Math.max(0, sheetIndex),
            ri: selection?.getRow() ?? 0,
            ci: selection?.getColumn() ?? 0,
        };
    }

    restoreViewState(state: UniverViewState) {
        if (this.disposed) return;
        const fWorkbook = this.univerAPI.getActiveWorkbook();
        if (!fWorkbook) return;
        const snapshot = fWorkbook.getSnapshot();
        const sheetId = snapshot.sheetOrder[Math.min(Math.max(0, state.sheetIndex), snapshot.sheetOrder.length - 1)];
        if (!sheetId) return;
        const sheet = fWorkbook.getSheetBySheetId(sheetId);
        if (!sheet) return;
        fWorkbook.setActiveSheet(sheet);
        const maxRow = Math.max(0, (sheet.getMaxRows?.() ?? 1) - 1);
        const maxCol = Math.max(0, (sheet.getMaxColumns?.() ?? 1) - 1);
        const ri = Math.min(Math.max(0, state.ri), maxRow);
        const ci = Math.min(Math.max(0, state.ci), maxCol);
        try {
            (sheet as never as { scrollToCell(r: number, c: number): void }).scrollToCell(ri, ci);
            sheet.getRange(ri, ci).activate();
        } catch {
            // 滚动恢复失败不阻断加载
        }
    }

    /** 当前 workbook 数据的深拷贝（diff 基线/当前态用） */
    getWorkbookDataCopy(): unknown | null {
        if (this.disposed) return null;
        const fWorkbook = this.univerAPI.getActiveWorkbook();
        if (!fWorkbook) return null;
        return JSON.parse(JSON.stringify(fWorkbook.save()));
    }

    /**
     * 开始编辑会话：监听 mutation 驱动 dirty 回调（首次变更触发一次），
     * 并记录结构性变更的 sheet。必须在 loadWorkbook（含超链接应用）完成后调用。
     */
    startEditSession(onDirty: () => void): UniverEditSession {
        const api = this.univerAPI as never as {
            addEvent(name: string, cb: (p: { id?: string; params?: { subUnitId?: string } }) => void): { dispose(): void };
            Event: Record<string, string>;
        };
        const structuralSheetIds = new Set<string>();
        let dirtyNotified = false;
        const disposable = api.addEvent(api.Event.CommandExecuted, (e) => {
            const id = e?.id ?? '';
            if (!id.startsWith('sheet.mutation.')) return;
            if (STRUCTURAL_MUTATION.test(id)) {
                const subUnitId = e.params?.subUnitId;
                if (subUnitId) structuralSheetIds.add(subUnitId);
            }
            if (!dirtyNotified) {
                dirtyNotified = true;
                onDirty();
            }
        });
        return {
            structuralSheetIds,
            reset() {
                structuralSheetIds.clear();
                dirtyNotified = false;
            },
            stop() {
                disposable.dispose();
            },
        };
    }

    /**
     * M5：禁用无法无损保存的编辑（结构性操作 + DV/CF），命中时 cancel 命令
     * 并回调 onBlocked（UI 弹「此操作请在 Excel 中进行」）。返回解绑函数。
     * 仅对走 XML 补丁保存的文件（xlsx/xlsm 原文件编辑）调用。
     */
    restrictLossyEdits(onBlocked: (commandId: string) => void): () => void {
        const api = this.univerAPI as never as {
            addEvent(name: string, cb: (p: { id?: string; cancel?: boolean }) => void): { dispose(): void };
            Event: Record<string, string>;
        };
        const disposable = api.addEvent(api.Event.BeforeCommandExecute, (e) => {
            const id = e?.id ?? '';
            if (!RESTRICTED_COMMAND.test(id)) return;
            e.cancel = true;
            onBlocked(id);
        });
        return () => disposable.dispose();
    }

    /** 选区/activeSheet 变化时回调（视图状态持久化用），返回解绑函数 */
    onViewStateChange(cb: () => void): () => void {
        const api = this.univerAPI as never as {
            addEvent(name: string, cb: (p: unknown) => void): { dispose(): void };
            Event: Record<string, string>;
        };
        const disposables: { dispose(): void }[] = [];
        for (const name of ['SelectionChanged', 'ActiveSheetChanged']) {
            const eventId = api.Event?.[name];
            if (!eventId) continue;
            try {
                disposables.push(api.addEvent(eventId, () => cb()));
            } catch {
                // 事件不可用时静默跳过
            }
        }
        return () => disposables.forEach(d => d.dispose());
    }

    dispose() {
        if (this.disposed) return;
        this.disposed = true;
        this.univer.dispose();
    }
}
