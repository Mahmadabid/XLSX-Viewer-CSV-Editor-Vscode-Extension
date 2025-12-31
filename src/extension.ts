import * as vscode from 'vscode';
import { XLSXEditorProvider } from './xlsxEditorProvider';
import { CSVEditorProvider } from './csvEditorProvider';
import { TSVEditorProvider } from './tsvEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    const xlsxProvider = new XLSXEditorProvider(context);
    const csvProvider = new CSVEditorProvider(context);
    const tsvProvider = new TSVEditorProvider(context);

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
        }),
        vscode.window.registerCustomEditorProvider('xlsxViewer.tsv', tsvProvider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.goBackToTableView', async () => {
            const activeEditor = vscode.window.activeTextEditor;
            if (!activeEditor) return;
            const path = activeEditor.document.uri.fsPath.toLowerCase();
            const uri = activeEditor.document.uri;

            if (path.endsWith('.csv')) {
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, 'xlsxViewer.csv');
            } else if (path.endsWith('.tsv')) {
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, 'xlsxViewer.tsv');
            }
        })
    );
}

export function deactivate() { }