import * as vscode from 'vscode';
import { ReactApp } from './common/reactApp';
import { registerExcelDiffCommand } from './provider/excelDiffProvider';
import { ExcelViewerProvider } from './provider/excelViewerProvider';
import { registerScmDiffInterceptor } from './provider/scmDiffInterceptor';

export async function activate(context: vscode.ExtensionContext) {
	const viewOption = { webviewOptions: { retainContextWhenHidden: true } };
	ReactApp.init(context);
	const viewerProvider = new ExcelViewerProvider(context);
	context.subscriptions.push(
		viewerProvider.bindCustomEditor(viewOption),
		registerExcelDiffCommand(context),
		registerScmDiffInterceptor(context),
	);
}

export function deactivate() { }
