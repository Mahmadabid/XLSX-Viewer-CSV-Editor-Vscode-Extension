import * as vscode from 'vscode';
import { XLSXEditorProvider } from './xlsxEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    const provider = new XLSXEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('xlsxViewer.xlsx', provider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );
}

export function deactivate() { }