import { basename, extname } from 'path';
import { Uri, workspace, type Webview } from 'vscode';
import { isUriReadOnly } from '@/common/fileReadOnly';
import { Handler } from '@/common/handler';

export function isVirtualUri(uri: Uri): boolean {
    return uri.scheme !== 'file';
}

export function buildDocumentCacheId(uri: Uri): string {
    return `${uri.scheme}:${uri.toString()}`;
}

export async function readUriBytes(uri: Uri): Promise<Uint8Array> {
    return workspace.fs.readFile(uri);
}

export function bytesToPayloadBuffer(data: Uint8Array): number[] {
    const buffer: number[] = new Array(data.length);
    for (let i = 0; i < data.length; i++) {
        buffer[i] = data[i];
    }
    return buffer;
}

export async function emitVirtualExcelOpen(handler: Handler, uri: Uri): Promise<void> {
    const ext = extname(uri.fsPath);
    const readOnly = await isUriReadOnly(uri);
    try {
        const data = await readUriBytes(uri);
        handler.emit('open', {
            ext,
            path: uri.toString(),
            fileName: basename(uri.fsPath),
            scheme: uri.scheme,
            documentCacheId: buildDocumentCacheId(uri),
            readOnly,
            buffer: bytesToPayloadBuffer(data),
        });
    } catch (error) {
        handler.emit('open', {
            ext,
            path: uri.toString(),
            fileName: basename(uri.fsPath),
            scheme: uri.scheme,
            documentCacheId: buildDocumentCacheId(uri),
            readOnly,
            error: error instanceof Error ? error.message : 'Failed to read file',
        });
    }
}

export async function emitFileExcelOpen(handler: Handler, uri: Uri, webview: Webview): Promise<void> {
    handler.emit('open', {
        ext: extname(uri.fsPath),
        path: webview.asWebviewUri(uri).toString(),
        fileName: basename(uri.fsPath),
        documentCacheId: buildDocumentCacheId(uri),
        readOnly: await isUriReadOnly(uri),
    });
}
