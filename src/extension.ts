import * as vscode from 'vscode';
import { XLSXEditorProvider } from './xlsxEditorProvider';
import { CSVEditorProvider } from './csvEditorProvider';
import { TSVEditorProvider } from './tsvEditorProvider';
import { MDEditorProvider } from './mdEditorProvider';

export function activate(context: vscode.ExtensionContext) {
    const xlsxProvider = new XLSXEditorProvider(context);
    const csvProvider = new CSVEditorProvider(context);
    const tsvProvider = new TSVEditorProvider(context);
    const mdProvider = new MDEditorProvider(context);

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
        }),
        vscode.window.registerCustomEditorProvider('xlsxViewer.md', mdProvider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.goBackToTableView', async (uri?: vscode.Uri) => {
            if (uri instanceof vscode.Uri) {
                const path = uri.fsPath.toLowerCase();
                const viewType = path.endsWith('.tsv') ? 'xlsxViewer.tsv' : 'xlsxViewer.csv';
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, viewType);
                return;
            }

            const activeEditor = vscode.window.activeTextEditor;
            if (activeEditor) {
                const docUri = activeEditor.document.uri;
                const path = docUri.fsPath.toLowerCase();
                if (path.endsWith('.csv')) {
                    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                    await vscode.commands.executeCommand('vscode.openWith', docUri, 'xlsxViewer.csv');
                } else if (path.endsWith('.tsv')) {
                    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                    await vscode.commands.executeCommand('vscode.openWith', docUri, 'xlsxViewer.tsv');
                }
            } else {
                // Try to get URI from active tab if not a text editor
                const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
                if (activeTab?.input instanceof (vscode as any).TabInputCustom || activeTab?.input instanceof (vscode as any).TabInputText) {
                    const tabUri = (activeTab.input as any).uri;
                    if (tabUri) {
                        const path = tabUri.fsPath.toLowerCase();
                        if (path.endsWith('.csv')) {
                            await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                            await vscode.commands.executeCommand('vscode.openWith', tabUri, 'xlsxViewer.csv');
                        } else if (path.endsWith('.tsv')) {
                            await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                            await vscode.commands.executeCommand('vscode.openWith', tabUri, 'xlsxViewer.tsv');
                        }
                    }
                }
            }
        }),

        vscode.commands.registerCommand('xlsx-viewer.goBackToXlsxView', async (uri?: vscode.Uri) => {
            if (uri instanceof vscode.Uri) {
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, 'xlsxViewer.xlsx');
                return;
            }

            const activeEditor = vscode.window.activeTextEditor;
            if (activeEditor) {
                const docUri = activeEditor.document.uri;
                if (docUri.fsPath.toLowerCase().endsWith('.xlsx')) {
                    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                    await vscode.commands.executeCommand('vscode.openWith', docUri, 'xlsxViewer.xlsx');
                }
            } else {
                // Try to get URI from active tab if not a text editor
                const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
                if (activeTab?.input instanceof (vscode as any).TabInputCustom || activeTab?.input instanceof (vscode as any).TabInputText) {
                    const tabUri = (activeTab.input as any).uri;
                    if (tabUri && tabUri.fsPath.toLowerCase().endsWith('.xlsx')) {
                        await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                        await vscode.commands.executeCommand('vscode.openWith', tabUri, 'xlsxViewer.xlsx');
                    }
                }
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.goBackToMdPreview', async (uri?: vscode.Uri) => {
            if (uri instanceof vscode.Uri) {
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', uri, 'xlsxViewer.md');
                return;
            }

            const activeEditor = vscode.window.activeTextEditor;
            if (activeEditor) {
                const docUri = activeEditor.document.uri;
                await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                await vscode.commands.executeCommand('vscode.openWith', docUri, 'xlsxViewer.md');
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.toggleAssociation', async (params: { type: 'xlsx' | 'csv' | 'tsv' | 'md', enable: boolean }) => {
            try {
                const { type, enable } = params;
                const patternMap = {
                    'md': '*.md',
                    'xlsx': '*.xlsx',
                    'csv': '*.csv',
                    'tsv': '*.tsv'
                };
                const viewTypeMap = {
                    'md': 'xlsxViewer.md',
                    'xlsx': 'xlsxViewer.xlsx',
                    'csv': 'xlsxViewer.csv',
                    'tsv': 'xlsxViewer.tsv'
                };
                const labelMap = {
                    'md': 'Markdown',
                    'xlsx': 'XLSX',
                    'csv': 'CSV',
                    'tsv': 'TSV'
                };

                const pattern = patternMap[type];
                const viewType = viewTypeMap[type];
                const label = labelMap[type];

                const cfg = vscode.workspace.getConfiguration();
                const associations: any = cfg.get('workbench.editorAssociations') || {};
                let newAssociations: any;

                if (enable) {
                    if (Array.isArray(associations)) {
                        newAssociations = associations.filter(a => a.filenamePattern !== pattern && a.filenamePattern !== `**/${pattern}`);
                        newAssociations.push({ viewType: viewType, filenamePattern: pattern });
                    } else {
                        newAssociations = { ...associations };
                        newAssociations[pattern] = viewType;
                    }
                    await cfg.update('workbench.editorAssociations', newAssociations, vscode.ConfigurationTarget.Global);
                    vscode.window.showInformationMessage(`XLSX Viewer is now set as the default editor for ${label} files.`);
                } else {
                    if (Array.isArray(associations)) {
                        newAssociations = associations.filter(a => a.viewType !== viewType && a.filenamePattern !== pattern && a.filenamePattern !== `**/${pattern}`);
                    } else {
                        newAssociations = { ...associations };
                        delete (newAssociations as any)[pattern];
                        delete (newAssociations as any)[`**/${pattern}`];
                    }
                    await cfg.update('workbench.editorAssociations', newAssociations, vscode.ConfigurationTarget.Global);
                    vscode.window.showInformationMessage(`${label} association has been removed from settings.`);
                }
            } catch (err) {
                vscode.window.showErrorMessage(`Error updating association: ${err}`);
            }
        })
    );

    // Keep the old command for backward compatibility if needed, but point it to the new one
    context.subscriptions.push(
        vscode.commands.registerCommand('xlsx-viewer.toggleMdAssociation', async (enable: boolean) => {
            await vscode.commands.executeCommand('xlsx-viewer.toggleAssociation', { type: 'md', enable });
        })
    );
}

export function deactivate() { }