import * as vscode from 'vscode';
import { XLSXEditorProvider } from './xlsxEditorProvider';
import { CSVEditorProvider } from './csvEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    const xlsxProvider = new XLSXEditorProvider(context);
    const csvProvider = new CSVEditorProvider(context);

    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('xlsxViewer.xlsx', xlsxProvider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        }),
        vscode.window.registerCustomEditorProvider('xlsxViewer.csv', csvProvider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.goBackToTableView', async () => {
            const activeEditor = vscode.window.activeTextEditor;
            if (activeEditor && activeEditor.document.uri.fsPath.toLowerCase().endsWith('.csv')) {
                const uri = activeEditor.document.uri;
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, 'xlsxViewer.csv');
            }
        })
    );
}

export function deactivate() { }