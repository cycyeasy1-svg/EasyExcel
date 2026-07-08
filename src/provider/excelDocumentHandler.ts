import { join, parse } from 'path';
import * as vscode from 'vscode';
import { Handler } from '@/common/handler';
import { isUriReadOnly } from '@/common/fileReadOnly';
import { emitFileExcelOpen, emitVirtualExcelOpen, isVirtualUri } from './excelDocumentContent';

const EXCEL_VIEW_TYPE = 'easychen.easyExcel';
const fileSaveTimes: Record<string, number> = {};

export function shouldSkipFileChange(uri: vscode.Uri): boolean {
    const lastSaveTime = fileSaveTimes[uri.toString()];
    return !!(lastSaveTime && Date.now() - lastSaveTime < 100);
}

/** 宿主侧对单个编辑器的控制句柄，交给 CustomEditorProvider 驱动原生保存/撤销 */
export interface ExcelEditorController {
    /** VSCode 触发保存（Ctrl+S / 关闭确认「保存」/ 文件菜单）时调用，
     *  转 webview 走既有保存流程，写盘成功后 resolve、被取消则 reject */
    requestSave(): Promise<void>;
    /** VSCode 触发「还原」时调用：从磁盘重新加载 */
    reload(): void;
}

export interface ExcelEditorHooks {
    /** webview 报告发生真实变更时调用：驱动 VSCode 标记文档为 dirty */
    markDirty(): void;
}

function buildSaveAsUri(uri: vscode.Uri, ext: string): vscode.Uri {
    const { dir, name } = parse(uri.fsPath);
    const defaultFileName = `${name}.${ext}`;
    if (uri.scheme === 'file') {
        return vscode.Uri.file(join(dir, defaultFileName));
    }
    return vscode.Uri.joinPath(uri, '..', defaultFileName);
}

export function handleExcelDocumentEvents(
    uri: vscode.Uri,
    handler: Handler,
    hooks: ExcelEditorHooks,
): ExcelEditorController {
    let readOnly = false;
    let pinned = false;
    // requestSave 的待决 promise：写盘成功→resolve，取消/无法保存→reject
    let pendingSave: { resolve: () => void; reject: (reason?: unknown) => void } | null = null;

    const resolvePending = () => {
        const p = pendingSave;
        pendingSave = null;
        p?.resolve();
    };
    const rejectPending = (reason?: unknown) => {
        const p = pendingSave;
        pendingSave = null;
        p?.reject(reason ?? new Error('Save cancelled'));
    };

    const sendOpenPayload = async () => {
        if (shouldSkipFileChange(uri)) {
            return;
        }
        readOnly = await isUriReadOnly(uri);
        if (isVirtualUri(uri)) {
            await emitVirtualExcelOpen(handler, uri);
            return;
        }
        await emitFileExcelOpen(handler, uri, handler.panel.webview);
    };

    handler
        .on('init', () => {
            void sendOpenPayload();
        })
        .on('fileChange', () => {
            void sendOpenPayload();
        })
        .on('change', () => {
            // 首次变更把编辑器从「预览」态固定，避免打开其他文件时被替换而丢失编辑
            if (!pinned) {
                pinned = true;
                void vscode.commands.executeCommand('workbench.action.keepEditor');
            }
            hooks.markDirty();
        })
        // webview 的保存按钮 / Ctrl+S 统一转由 VSCode 发起保存，
        // 以便原生 dirty 状态在写盘后正确清除
        .on('hostSave', () => {
            void vscode.commands.executeCommand('workbench.action.files.save');
        })
        .on('save', async (content) => {
            const bytes = Array.isArray(content) ? new Uint8Array(content) : new TextEncoder().encode(content);
            if (readOnly) {
                handler.emit('saveAs', { content: [...bytes] });
                return;
            }
            try {
                await vscode.workspace.fs.writeFile(uri, bytes);
                fileSaveTimes[uri.toString()] = Date.now();
                handler.emit('saveDone');
                resolvePending();
            } catch (e) {
                rejectPending(e);
                throw e;
            }
        })
        .on('saveAs', async (payload: { content: number[]; ext?: string }) => {
            const bytes = new Uint8Array(payload.content);
            const ext = (payload.ext ?? 'xlsx').toLowerCase();
            const filterMap: Record<string, { label: string; exts: string[] }> = {
                xlsx: { label: 'Excel Workbook', exts: ['xlsx'] },
                xlsm: { label: 'Excel Macro-Enabled Workbook', exts: ['xlsm'] },
                xls: { label: 'Excel 97-2003 Workbook', exts: ['xls'] },
                ods: { label: 'OpenDocument Spreadsheet', exts: ['ods'] },
                csv: { label: 'CSV (Comma delimited)', exts: ['csv'] },
                tsv: { label: 'TSV (Tab delimited)', exts: ['tsv'] },
            };
            const info = filterMap[ext] ?? { label: ext.toUpperCase(), exts: [ext] };
            const target = await vscode.window.showSaveDialog({
                defaultUri: buildSaveAsUri(uri, ext),
                filters: { [info.label]: info.exts },
            });
            if (!target) {
                // 取消另存：若这是一次被请求的保存（如非 xlsx 改存 xlsx），需回滚 dirty
                rejectPending();
                return;
            }
            await vscode.workspace.fs.writeFile(target, bytes);
            fileSaveTimes[target.toString()] = Date.now();
            handler.emit('saveDone');
            resolvePending();
            await vscode.commands.executeCommand('vscode.openWith', target, EXCEL_VIEW_TYPE);
        })
        // webview 侧保存流程结束但未写盘（用户取消对话框 / 无法无损保存）：
        // 回滚待决保存，保持 dirty，从而中止关闭
        .on('saveSettled', (payload: { cancelled?: boolean } | undefined) => {
            if (payload?.cancelled) rejectPending();
        })
        .on('developerTool', () => vscode.commands.executeCommand('workbench.action.toggleDevTools'))
        .on('openExternal', (url: string) => {
            if (url) {
                void vscode.env.openExternal(vscode.Uri.parse(url));
            }
        })
        .on('dispose', () => {
            rejectPending(new Error('Editor closed'));
            delete fileSaveTimes[uri.toString()];
        });

    return {
        requestSave() {
            // 取消上一个未决保存（理论上不会并发），避免悬挂
            rejectPending(new Error('Superseded by a newer save'));
            return new Promise<void>((resolve, reject) => {
                pendingSave = { resolve, reject };
                handler.emit('requestSave');
            });
        },
        reload() {
            void sendOpenPayload();
        },
    };
}
