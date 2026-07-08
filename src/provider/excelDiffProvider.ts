import { basename, dirname, extname, relative } from 'path';
import { spawn } from 'child_process';
import * as vscode from 'vscode';
import { ReactApp } from '@/common/reactApp';
import { getExtensionResourceRoots } from '@/common/extensionResource';
import { bytesToPayloadBuffer } from './excelDocumentContent';

const EXCEL_DIFF_VIEW_TYPE = 'easychen.easyExcelDiff';
const OPEN_EXCEL_DIFF_COMMAND = 'easyexcel.openDiff';
const SUPPORTED_EXCEL_SUFFIXES = new Set(['.xlsx', '.xlsm', '.xls', '.csv', '.tsv', '.ods']);
const MAX_GIT_BUFFER_BYTES = 80 * 1024 * 1024;

type DiffSide = {
    label: string;
    buffer?: number[];
    missing?: boolean;
};

type DiffMode = 'staged' | 'unstaged';

interface ExcelDiffPayload {
    ext: string;
    fileName: string;
    sourceLabel: string;
    left: DiffSide;
    right: DiffSide;
}

function git(repoRoot: string, args: string[], binary = false): Promise<Uint8Array | string> {
    return new Promise((resolve, reject) => {
        const child = spawn('git', ['-C', repoRoot, ...args], {
            stdio: ['ignore', 'pipe', 'pipe'],
            windowsHide: true,
        });
        const stdout: Buffer[] = [];
        const stderr: Buffer[] = [];
        let stdoutLength = 0;

        child.stdout.on('data', (chunk: Buffer) => {
            stdoutLength += chunk.length;
            if (stdoutLength > MAX_GIT_BUFFER_BYTES) {
                child.kill();
                reject(new Error('Git output is too large to diff in EasyExcel.'));
                return;
            }
            stdout.push(chunk);
        });
        child.stderr.on('data', (chunk: Buffer) => stderr.push(chunk));
        child.on('error', reject);
        child.on('close', code => {
            const out = Buffer.concat(stdout);
            if (code === 0) {
                resolve(binary ? new Uint8Array(out) : out.toString('utf8'));
                return;
            }
            const message = Buffer.concat(stderr).toString('utf8').trim();
            reject(new Error(message || `git ${args.join(' ')} failed with exit code ${code}`));
        });
    });
}

function extractUri(input: unknown): vscode.Uri | undefined {
    if (input instanceof vscode.Uri) return input;
    const candidate = input as { resourceUri?: vscode.Uri; uri?: vscode.Uri; fsPath?: string } | undefined;
    if (candidate?.resourceUri instanceof vscode.Uri) return candidate.resourceUri;
    if (candidate?.uri instanceof vscode.Uri) return candidate.uri;
    if (candidate?.fsPath) return vscode.Uri.file(candidate.fsPath);
    return vscode.window.activeTextEditor?.document.uri;
}

function toGitPath(repoRoot: string, uri: vscode.Uri): string {
    return relative(repoRoot, uri.fsPath).replace(/\\/g, '/');
}

async function findGitRoot(uri: vscode.Uri): Promise<string> {
    const folder = vscode.workspace.getWorkspaceFolder(uri);
    const cwd = folder?.uri.fsPath ?? dirname(uri.fsPath);
    const root = String(await git(cwd, ['rev-parse', '--show-toplevel'])).trim();
    if (!root) {
        throw new Error('This file is not inside a Git repository.');
    }
    return root;
}

async function hasGitOutput(repoRoot: string, args: string[]): Promise<boolean> {
    const out = String(await git(repoRoot, args)).trim();
    return out.length > 0;
}

async function readGitBlob(repoRoot: string, spec: string): Promise<Uint8Array | undefined> {
    try {
        return await git(repoRoot, ['show', spec], true) as Uint8Array;
    } catch {
        return undefined;
    }
}

async function readWorktreeFile(uri: vscode.Uri): Promise<Uint8Array | undefined> {
    try {
        return await vscode.workspace.fs.readFile(uri);
    } catch {
        return undefined;
    }
}

async function pickDiffMode(repoRoot: string, gitPath: string): Promise<DiffMode> {
    const [hasStaged, hasUnstaged, isUntracked] = await Promise.all([
        hasGitOutput(repoRoot, ['diff', '--cached', '--name-only', '--', gitPath]),
        hasGitOutput(repoRoot, ['diff', '--name-only', '--', gitPath]),
        hasGitOutput(repoRoot, ['ls-files', '--others', '--exclude-standard', '--', gitPath]),
    ]);

    if (hasStaged && (hasUnstaged || isUntracked)) {
        const picked = await vscode.window.showQuickPick([
            { label: 'Staged changes', mode: 'staged' as const, description: 'HEAD <-> index' },
            { label: 'Working tree changes', mode: 'unstaged' as const, description: 'index <-> working tree' },
        ], { placeHolder: 'Select which Excel changes to compare' });
        if (!picked) {
            throw new Error('Excel diff cancelled.');
        }
        return picked.mode;
    }
    if (hasStaged) return 'staged';
    return 'unstaged';
}

async function buildDiffPayload(uri: vscode.Uri): Promise<ExcelDiffPayload> {
    const ext = extname(uri.fsPath).toLowerCase();
    if (!SUPPORTED_EXCEL_SUFFIXES.has(ext)) {
        throw new Error('EasyExcel Diff supports .xls, .xlsx, .xlsm, .csv, .tsv, and .ods files.');
    }

    const repoRoot = await findGitRoot(uri);
    const gitPath = toGitPath(repoRoot, uri);
    const mode = await pickDiffMode(repoRoot, gitPath);

    const [headBytes, indexBytes, worktreeBytes] = await Promise.all([
        mode === 'staged' ? readGitBlob(repoRoot, `HEAD:${gitPath}`) : Promise.resolve(undefined),
        readGitBlob(repoRoot, `:${gitPath}`),
        mode === 'unstaged' ? readWorktreeFile(uri) : Promise.resolve(undefined),
    ]);

    const leftBytes = mode === 'staged' ? headBytes : indexBytes;
    const rightBytes = mode === 'staged' ? indexBytes : worktreeBytes;
    if (!leftBytes && !rightBytes) {
        throw new Error('No Git version could be read for this Excel file.');
    }

    const leftLabel = mode === 'staged' ? 'HEAD' : 'Index';
    const rightLabel = mode === 'staged' ? 'Staged' : 'Working Tree';
    return {
        ext,
        fileName: basename(uri.fsPath),
        sourceLabel: `${leftLabel} <-> ${rightLabel}`,
        left: leftBytes
            ? { label: leftLabel, buffer: bytesToPayloadBuffer(leftBytes) }
            : { label: leftLabel, missing: true },
        right: rightBytes
            ? { label: rightLabel, buffer: bytesToPayloadBuffer(rightBytes) }
            : { label: rightLabel, missing: true },
    };
}

async function openDiffPanel(context: vscode.ExtensionContext, payload: ExcelDiffPayload) {
    const panel = vscode.window.createWebviewPanel(
        EXCEL_DIFF_VIEW_TYPE,
        `Excel Diff: ${payload.fileName}`,
        vscode.ViewColumn.Active,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
            localResourceRoots: getExtensionResourceRoots(context),
        }
    );

    panel.webview.onDidReceiveMessage(message => {
        if (message.type === 'init') {
            void panel.webview.postMessage({ type: 'diffOpen', content: payload });
            return;
        }
        if (message.type === 'developerTool') {
            void vscode.commands.executeCommand('workbench.action.toggleDevTools');
        }
    });

    await ReactApp.view(panel.webview, { route: 'excel-diff' });
}

export function registerExcelDiffCommand(context: vscode.ExtensionContext): vscode.Disposable {
    return vscode.commands.registerCommand(OPEN_EXCEL_DIFF_COMMAND, async (input?: unknown) => {
        try {
            const uri = extractUri(input);
            if (!uri) {
                throw new Error('Select an Excel file to compare.');
            }
            const payload = await buildDiffPayload(uri);
            await openDiffPanel(context, payload);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (message !== 'Excel diff cancelled.') {
                void vscode.window.showErrorMessage(message);
            }
        }
    });
}
