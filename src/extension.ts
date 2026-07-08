import * as vscode from 'vscode';
import { ReactApp } from './common/reactApp';
import { registerExcelDiffCommand } from './provider/excelDiffProvider';
import { ExcelViewerProvider } from './provider/excelViewerProvider';

export async function activate(context: vscode.ExtensionContext) {
	const viewOption = { webviewOptions: { retainContextWhenHidden: true } };
	ReactApp.init(context);
	const viewerProvider = new ExcelViewerProvider(context);
	context.subscriptions.push(
		viewerProvider.bindCustomEditor(viewOption),
		registerExcelDiffCommand(context),
	);
}

export function deactivate() { }
