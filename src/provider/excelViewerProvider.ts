import * as vscode from 'vscode';
import { ReactApp } from '@/common/reactApp';
import { getFileSuffix } from '@/common/fileSuffix';
import { Handler } from '@/common/handler';
import { getExtensionResourceRoots } from '@/common/extensionResource';
import { handleExcelDocumentEvents, type ExcelEditorController } from './excelDocumentHandler';

export const EXCEL_VIEW_TYPE = 'easychen.easyExcel';

const SUPPORTED_EXCEL_SUFFIXES = new Set(['.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.ods']);

class ExcelDocument implements vscode.CustomDocument {
    constructor(readonly uri: vscode.Uri) { }
    dispose(): void { /* no resources held on the document itself */ }
}

/**
 * 可编辑的自定义编辑器：以 IWorkbookData 的实际 diff 驱动 VSCode 原生 dirty 状态，
 * 关闭未保存文档时触发原生「是否保存」确认。保存交由 webview 既有流程完成
 * （字节级增量补丁 / 另存对话框），宿主经 requestSave 桥接并在写盘后清除 dirty。
 */
export class ExcelViewerProvider implements vscode.CustomEditorProvider<ExcelDocument> {
    private readonly _onDidChangeCustomDocument =
        new vscode.EventEmitter<vscode.CustomDocumentContentChangeEvent<ExcelDocument>>();
    readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

    /** uri.toString() → 当前编辑器控制句柄（不支持同一文档多编辑器，键唯一） */
    private readonly controllers = new Map<string, ExcelEditorController>();

    constructor(private context: vscode.ExtensionContext) { }

    bindCustomEditor(viewOption: { webviewOptions: vscode.WebviewPanelOptions }) {
        return vscode.window.registerCustomEditorProvider(EXCEL_VIEW_TYPE, this, viewOption);
    }

    openCustomDocument(uri: vscode.Uri): ExcelDocument {
        return new ExcelDocument(uri);
    }

    resolveCustomEditor(document: ExcelDocument, webviewPanel: vscode.WebviewPanel): void | Thenable<void> {
        const uri = document.uri;
        const suffix = getFileSuffix(uri.fsPath);
        if (!SUPPORTED_EXCEL_SUFFIXES.has(suffix)) {
            void vscode.commands.executeCommand('vscode.openWith', uri, 'default');
            return;
        }

        const webview = webviewPanel.webview;
        const folderPath = vscode.Uri.joinPath(uri, '..');
        webview.options = {
            enableScripts: true,
            localResourceRoots: [...getExtensionResourceRoots(this.context), folderPath],
        };

        const handler = Handler.bind(webviewPanel, uri);
        const controller = handleExcelDocumentEvents(uri, handler, {
            markDirty: () => this._onDidChangeCustomDocument.fire({ document }),
        });
        const key = uri.toString();
        this.controllers.set(key, controller);
        webviewPanel.onDidDispose(() => {
            if (this.controllers.get(key) === controller) {
                this.controllers.delete(key);
            }
        });
        return ReactApp.view(webview, { route: 'excel' });
    }

    saveCustomDocument(document: ExcelDocument): Thenable<void> {
        return this.controllers.get(document.uri.toString())?.requestSave() ?? Promise.resolve();
    }

    // 「文件 > 名前を付けて保存」等 VSCode 原生另存：复用 webview 保存流程写回文档。
    // 面向用户的多格式「另存为」由 webview 工具栏按钮独立提供。
    saveCustomDocumentAs(document: ExcelDocument): Thenable<void> {
        return this.controllers.get(document.uri.toString())?.requestSave() ?? Promise.resolve();
    }

    revertCustomDocument(document: ExcelDocument): Thenable<void> {
        this.controllers.get(document.uri.toString())?.reload();
        return Promise.resolve();
    }

    // 未保存内容存活于 webview 内，无法在宿主侧序列化备份 → 不参与热退出恢复；
    // 关闭标签页时仍会弹出原生保存确认（本次改善的核心目标）。
    backupCustomDocument(
        document: ExcelDocument,
        _context: vscode.CustomDocumentBackupContext,
    ): Thenable<vscode.CustomDocumentBackup> {
        return Promise.resolve({ id: document.uri.toString(), delete: () => { /* no-op */ } });
    }
}
