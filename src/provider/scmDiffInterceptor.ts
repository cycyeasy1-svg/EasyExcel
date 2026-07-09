import { basename, extname } from 'path';
import * as vscode from 'vscode';
import { Output } from '@/common/Output';
import { runExcelDiff, SUPPORTED_EXCEL_SUFFIXES } from './excelDiffProvider';

// VS Code has no custom-diff-editor API (microsoft/vscode#97169). When the Git
// extension runs `git.openChange` on a file claimed by our custom editor, the
// resulting tab renders two custom editors side by side and the tabs API reports
// `tab.input === undefined` — no URIs, no view type. The only stable handle we
// get is `tab.label`, which is `"<basename> (<localized suffix>)"`. The suffix is
// translated ("(Working Tree)" / "(工作树)" / "(作業ツリー)"), so we match on the
// basename prefix alone and recover the URI from the Git extension's own state.

// Mirrors vscode/extensions/git/src/api/git.d.ts — the URI lives on `uri`, not
// `resourceUri` (that one belongs to SourceControlResourceState, a different type).
interface GitChange {
    uri: vscode.Uri;
}

interface GitRepository {
    state: {
        indexChanges: GitChange[];
        workingTreeChanges: GitChange[];
        mergeChanges: GitChange[];
        untrackedChanges: GitChange[];
    };
}

interface GitApi {
    repositories: GitRepository[];
}

function isSupported(uri: vscode.Uri): boolean {
    return SUPPORTED_EXCEL_SUFFIXES.has(extname(uri.fsPath).toLowerCase());
}

function interceptEnabled(): boolean {
    return vscode.workspace.getConfiguration('easyexcel').get<boolean>('interceptScmDiff', true);
}

function getGitApi(): GitApi | undefined {
    const extension = vscode.extensions.getExtension('vscode.git');
    if (!extension?.isActive) return undefined;
    try {
        return extension.exports?.getAPI?.(1) as GitApi | undefined;
    } catch {
        return undefined;
    }
}

/** Every Excel-ish file the Git extension currently reports as changed. */
function changedExcelUris(): vscode.Uri[] {
    const api = getGitApi();
    if (!api) return [];

    const seen = new Map<string, vscode.Uri>();
    for (const repo of api.repositories) {
        const groups = [
            repo.state.indexChanges,
            repo.state.workingTreeChanges,
            repo.state.mergeChanges,
            repo.state.untrackedChanges,
        ];
        for (const group of groups) {
            for (const change of group ?? []) {
                const uri = change?.uri;
                if (uri && isSupported(uri)) seen.set(uri.toString(), uri);
            }
        }
    }
    return [...seen.values()];
}

/**
 * Resolve the file a diff tab is showing.
 *
 * `TabInputTextDiff` (plain text editors, or a future VS Code that models custom
 * diffs properly) hands us the URIs directly. The custom-editor diff does not, so
 * we fall back to matching `tab.label` against the basenames of changed files.
 * Ambiguous basenames are skipped rather than guessed at.
 */
function resolveDiffTarget(tab: vscode.Tab): vscode.Uri | undefined {
    if (tab.input instanceof vscode.TabInputTextDiff) {
        const modified = tab.input.modified;
        return isSupported(modified) ? modified : undefined;
    }

    // Custom-editor diff: input is undefined. Anything else with a real input is
    // a normal editor and must not be touched.
    if (tab.input !== undefined) return undefined;

    const matches = changedExcelUris().filter(uri => tab.label.startsWith(`${basename(uri.fsPath)} (`));
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) {
        Output.debug(`[easyexcel] ambiguous diff tab "${tab.label}" matched ${matches.length} files; leaving it to VS Code`);
    }
    return undefined;
}

export function registerScmDiffInterceptor(context: vscode.ExtensionContext): vscode.Disposable {
    Output.debug('[easyexcel] SCM diff interceptor armed');

    return vscode.window.tabGroups.onDidChangeTabs(async event => {
        // Nothing here may throw: an unhandled rejection inside a tab listener is
        // swallowed by VS Code, so a bug would silently degrade to the native diff.
        try {
            if (!interceptEnabled()) return;

            for (const tab of event.opened) {
                const uri = resolveDiffTarget(tab);
                if (!uri) continue;

                Output.debug(`[easyexcel] intercepting diff tab "${tab.label}" -> ${uri.fsPath}`);
                await vscode.window.tabGroups.close(tab, true);
                await runExcelDiff(context, uri);
            }
        } catch (error) {
            Output.debug(`[easyexcel] SCM diff interception failed: ${error instanceof Error ? error.stack : error}`);
        }
    });
}
