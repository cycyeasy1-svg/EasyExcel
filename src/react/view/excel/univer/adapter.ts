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
import FilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN';
import FilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import SortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN';
import SortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import FindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN';
import FindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import DataValidationZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN';
import DataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import HyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN';
import HyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import DrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN';
import DrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import CfZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN';
import CfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';

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
}

export interface UniverEditSession {
    /** 发生过结构性变更（行列插删/移动）的 sheetId 集合 → 这些 sheet 保存时整表重建 */
    structuralSheetIds: Set<string>;
    /** 保存成功后调用：清空结构日志（baseline 已更新） */
    reset(): void;
    stop(): void;
}

/** 行列插删/移动类 mutation：使该 sheet 的逐格 diff 失效（注意 Univer 的
 *  命名不对称：insert-row/remove-rows、insert-col/remove-col） */
const STRUCTURAL_MUTATION = /^sheet\.mutation\.(insert-(row|col)|remove-(rows|col)|move-(rows|columns|range)|reorder-range)$/;

/** 应用超链接的上限，防止极端文件逐链接跑命令拖慢加载 */
const MAX_APPLIED_HYPERLINKS = 2000;

type FUniverLike = ReturnType<typeof createUniver>['univerAPI'];

const pickLocale = (language?: string) =>
    /^zh/i.test(language ?? '') ? LocaleType.ZH_CN : LocaleType.EN_US;

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
        const fWorkbook = this.univerAPI.createWorkbook(result.workbookData as IWorkbookData);
        await this.applyHyperlinks(result);
        if (opts.readOnly) {
            fWorkbook.setEditable(false);
        }
        return fWorkbook;
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
