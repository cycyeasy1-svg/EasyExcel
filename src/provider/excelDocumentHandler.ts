import { basename, join, parse } from 'path';
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

function setDirty(handler: Handler, uri: vscode.Uri, dirty: boolean) {
    const fileName = basename(uri.fsPath);
    handler.panel.title = dirty ? `* ${fileName}` : fileName;
    if (dirty) {
        void vscode.commands.executeCommand('workbench.action.keepEditor');
    }
}

function buildSaveAsUri(uri: vscode.Uri, ext: string): vscode.Uri {
    const { dir, name } = parse(uri.fsPath);
    const defaultFileName = `${name}.${ext}`;
    if (uri.scheme === 'file') {
        return vscode.Uri.file(join(dir, defaultFileName));
    }
    return vscode.Uri.joinPath(uri, '..', defaultFileName);
}

export function handleExcelDocumentEvents(uri: vscode.Uri, handler: Handler) {
    let readOnly = false;

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
            setDirty(handler, uri, true);
        })
        .on('save', async (content) => {
            const bytes = Array.isArray(content) ? new Uint8Array(content) : new TextEncoder().encode(content);
            if (readOnly) {
                handler.emit('saveAs', { content: [...bytes] });
                return;
            }
            await vscode.workspace.fs.writeFile(uri, bytes);
            fileSaveTimes[uri.toString()] = Date.now();
            setDirty(handler, uri, false);
            handler.emit('saveDone');
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
            if (!target) return;
            await vscode.workspace.fs.writeFile(target, bytes);
            fileSaveTimes[target.toString()] = Date.now();
            setDirty(handler, uri, false);
            handler.emit('saveDone');
            await vscode.commands.executeCommand('vscode.openWith', target, EXCEL_VIEW_TYPE);
        })
        .on('developerTool', () => vscode.commands.executeCommand('workbench.action.toggleDevTools'))
        .on('openExternal', (url: string) => {
            if (url) {
                void vscode.env.openExternal(vscode.Uri.parse(url));
            }
        })
        .on('dispose', () => {
            delete fileSaveTimes[uri.toString()];
        });
}
