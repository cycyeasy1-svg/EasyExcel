import * as vscode from 'vscode';
import { ReactApp } from '@/common/reactApp';
import { getFileSuffix } from '@/common/fileSuffix';
import { Handler } from '@/common/handler';
import { getExtensionResourceRoots } from '@/common/extensionResource';
import { handleExcelDocumentEvents } from './excelDocumentHandler';

export const EXCEL_VIEW_TYPE = 'easychen.easyExcel';

const SUPPORTED_EXCEL_SUFFIXES = new Set(['.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.ods']);

export class ExcelViewerProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private context: vscode.ExtensionContext) { }

    bindCustomEditor(viewOption: { webviewOptions: vscode.WebviewPanelOptions }) {
        return vscode.window.registerCustomEditorProvider(EXCEL_VIEW_TYPE, this, viewOption);
    }

    openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
        return { uri, dispose: (): void => { } };
    }

    resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel): void | Thenable<void> {
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
        handleExcelDocumentEvents(uri, handler);
        const engine = vscode.workspace.getConfiguration('easyExcel').get<string>('engine', 'legacy');
        return ReactApp.view(webview, { route: 'excel', engine });
    }
}
