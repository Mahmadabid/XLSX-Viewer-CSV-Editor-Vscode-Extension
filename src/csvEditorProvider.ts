import * as vscode from 'vscode';
import * as fs from 'fs';

export class CSVEditorProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private readonly context: vscode.ExtensionContext) { }

    async openCustomDocument(
        uri: vscode.Uri,
        openContext: vscode.CustomDocumentOpenContext,
        token: vscode.CancellationToken
    ): Promise<vscode.CustomDocument> {
        return { uri, dispose: () => { } };
    }

    async resolveCustomEditor(
        document: vscode.CustomDocument,
        webviewPanel: vscode.WebviewPanel,
        token: vscode.CancellationToken
    ): Promise<void> {
        try {
            const csvContent = fs.readFileSync(document.uri.fsPath, 'utf-8');
            const rows = csvContent.split('\n').map(row => row.split(','));

            const generateTableHtml = () => {
                let tableHtml = '<table border="1" cellspacing="0" cellpadding="5">';
                rows.forEach(row => {
                    tableHtml += '<tr>';
                    row.forEach(cell => {
                        const cellContent = cell.trim();
                        const isEmpty = cellContent === '';
                        const dataAttrs = [
                            'data-default-bg="true"',
                            'data-default-color="true"',
                            isEmpty ? 'data-empty="true"' : ''
                        ].filter(Boolean).join(' ');

                        tableHtml += `<td ${dataAttrs}>${isEmpty ? '&nbsp;' : cellContent}</td>`;
                    });
                    tableHtml += '</tr>';
                });
                tableHtml += '</table>';
                return tableHtml;
            };

            webviewPanel.webview.options = { enableScripts: true };
            webviewPanel.webview.html = this.getWebviewContent(generateTableHtml(), webviewPanel);

            webviewPanel.webview.onDidReceiveMessage(async message => {
                if (message.command === 'toggleView') {
                    const isTableView = message.isTableView;
                    if (isTableView) {
                        webviewPanel.webview.html = this.getWebviewContent(generateTableHtml(), webviewPanel);
                    } else {
                        // Open in default editor
                        await vscode.commands.executeCommand('vscode.openWith', document.uri, 'default');
                        // Hide the current webview panel
                        webviewPanel.dispose();
                    }
                } else if (message.command === 'toggleBackground') {
                    webviewPanel.webview.postMessage({ command: 'toggleBackground' });
                }
            });
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading CSV file: ${error}`);
        }
    }

    private getWebviewContent(tableHtml: string, webviewPanel: vscode.WebviewPanel): string {
        const imgUri = webviewPanel.webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'view.png'));
        const svgUri = webviewPanel.webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table.svg'));

        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>CSV Viewer</title>
            <style>
                body { 
                    font-family: sans-serif; 
                    padding: 10px; 
                    background-color: rgb(255, 255, 255);
                    margin: 0;
                    overflow-x: auto;
                }
                .table-container {
                    width: 100%;
                    overflow-x: auto;
                }
                table { 
                    border-collapse: collapse; 
                    width: auto;
                    min-width: 100%;
                    table-layout: fixed;
                }
                th, td { 
                    border: 1px solid #ccc; 
                    padding: 8px; 
                    text-align: left;
                    white-space: nowrap;
                }
                .button-container {
                    margin-bottom: 10px;
                    display: flex;
                    gap: 10px;
                    position: sticky;
                    top: 0;
                    background-color: inherit;
                    z-index: 1;
                }
                .toggle-button {
                    max-height: 42px;
                    padding: 8px 16px;
                    font-size: 14px;
                    font-weight: bold;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    background-color: #2196f3;
                    color: white;
                    transition: all 0.2s ease;
                }
                .toggle-button:hover {
                    background-color: #1976d2;
                }
                .toggle-button:active {
                    background-color: #1565c0;
                }
                .toggle-button svg {
                    width: 16px;
                    height: 16px;
                    stroke: white;
                }
                /* Default background */
                td { 
                    background-color: rgb(255, 255, 255);
                    color: rgb(0, 0, 0);
                }
                /* Alternate background when toggled */
                body.alt-bg { background-color: rgb(0, 0, 0); }
                body.alt-bg td { 
                    background-color: rgb(0, 0, 0);
                    color: rgb(255, 255, 255);
                }
                /* Empty cell border styles */
                td:empty { border-color: rgba(204, 204, 204, 1); }

                .tooltip {
                    position: relative;
                    display: inline-block;
                }

                .tooltip .tooltiptext {
                    visibility: hidden;
                    width: 250px;
                    background-color: rgb(231, 248, 255);
                    color: rgb(0, 0, 0);
                    text-align: center;
                    border-radius: 6px;
                    padding: 10px;
                    position: absolute;
                    z-index: 1;
                    top: 130%;
                    left: 50%;
                    margin-left: -125px;
                    opacity: 0;
                    transition: opacity 0.3s;
                    white-space: normal;
                    line-height: 1.5;
                }

                .tooltip:hover .tooltiptext {
                    visibility: visible;
                    opacity: 1;
                }

                .tooltip .tooltiptext span {
                    color: #ff6600;
                }

                .tooltip .tooltiptext .warning {
                    color: red;
                    font-weight: bold;
                    margin-bottom: 5px;
                }
                .tooltip .tooltiptext .instruction {
                    color:rgb(2, 105, 190);
                    font-weight: normal;
                }

            </style>
        </head>
        <body>
            <div class="button-container">
                <button id="toggleViewButton" class="toggle-button">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 20h9" />
                        <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z" />
                    </svg>
                    Edit File
                </button>
                <button id="toggleBackgroundButton" class="toggle-button">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M12 2v20M2 12h20M4.93 4.93l14.14 14.14M19.07 4.93L4.93 19.07"/>
                    </svg>
                    Toggle Background
                </button>
                <div class = "tooltip">
                    <img src="${imgUri}" alt="Change to table view"  style="width: auto; height: 32px; margin-left: auto;" />
                    <span class="tooltiptext">
                        <span class="warning">Important:</span> Click the blue table icon <img src="${svgUri}" alt="Table Icon" style="width: 16px; vertical-align: middle; height: 16px;" />
                         to switch to table view from edit file mode. <br>
                        <span class = "instruction">The table icon will only work on edit file mode and is located on the top right corner in the editor toolbar as shown in the image.</span>
                    </span>
                </div>
            </div>
            <div id="content">${tableHtml}</div>
            <script>
                const vscode = acquireVsCodeApi();
                let isTableView = true;

                document.getElementById('toggleViewButton').addEventListener('click', () => {
                    isTableView = !isTableView;
                    vscode.postMessage({ command: 'toggleView', isTableView });
                });

                document.getElementById('toggleBackgroundButton').addEventListener('click', () => {
                    document.body.classList.toggle('alt-bg');

                    const defaultBgCells = document.querySelectorAll('td[data-default-bg="true"]');
                    defaultBgCells.forEach(cell => {
                        if (document.body.classList.contains('alt-bg')) {
                            cell.style.backgroundColor = "rgb(0, 0, 0)";
                        } else {
                            cell.style.backgroundColor = "rgb(255, 255, 255)";
                        }
                    });

                    const defaultBothCells = document.querySelectorAll('td[data-default-bg="true"][data-default-color="true"]');
                    defaultBothCells.forEach(cell => {
                        if (document.body.classList.contains('alt-bg')) {
                            cell.style.color = "rgb(255, 255, 255)";
                        } else {
                            cell.style.color = "rgb(0, 0, 0)";
                        }
                    });
                });
            </script>
        </body>
        </html>`;
    }
}
