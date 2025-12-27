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
            const BATCH_SIZE = 1000;
            const filePath = document.uri.fsPath;
            let leftover = '';
            let rows: string[][] = [];
            let rowCount = 0;
            let columnCount = 0;
            let isFirstBatch = true;
            let streamStarted = false;

            // Helper to generate table HTML for a batch
            function generateTableRowsHtml(batchRows: string[][], startIndex: number): string {
                let html = '';
                batchRows.forEach((row, rowIndex) => {
                    html += `<tr><th class="row-header" data-row="${startIndex + rowIndex}">${startIndex + rowIndex + 1}</th>`;
                    row.forEach((cell, colIndex) => {
                        const cellContent = cell.trim();
                        const isEmpty = cellContent === '';
                        const dataAttrs = [
                            'data-default-bg="true"',
                            'data-default-color="true"',
                            isEmpty ? 'data-empty="true"' : '',
                            `data-row="${startIndex + rowIndex}"`,
                            `data-col="${colIndex}"`
                        ].filter(Boolean).join(' ');
                        html += `<td ${dataAttrs}><span class="cell-content">${isEmpty ? '&nbsp;' : cellContent}</span></td>`;
                    });
                    html += '</tr>';
                });
                return html;
            }

            // Helper to generate table header
            function generateTableHeaderHtml(colCount: number): string {
                let html = '<thead><tr><th class="row-header">&nbsp;</th>';
                for (let colNumber = 1; colNumber <= colCount; colNumber++) {
                    const colLabel = String.fromCharCode(64 + colNumber);
                    html += `<th class="col-header" data-col="${colNumber - 1}">${colLabel}</th>`;
                }
                html += '</tr></thead>';
                return html;
            }

            // Set up webview
            webviewPanel.webview.options = {
                enableScripts: true,
                localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'resources')]
            };
            webviewPanel.webview.html = this.getWebviewContent(
                `<table id="csv-table">
                    <colgroup></colgroup>
                    <thead></thead><tbody></tbody>
                </table>`,
                webviewPanel
            );

            const startStreaming = () => {
                if (streamStarted) return;
                streamStarted = true;

                const fileStream = fs.createReadStream(filePath, { encoding: 'utf-8' });
                webviewPanel.onDidDispose(() => {
                    try { fileStream.destroy(); } catch { }
                });

                fileStream.on('data', chunk => {
                    let data = leftover + chunk;
                    let lines = data.split('\n');
                    leftover = lines.pop() || '';
                    for (let line of lines) {
                        rows.push(line.split(','));
                        if (rowCount === 0) {
                            columnCount = rows[0].length;
                        }
                        rowCount++;
                        if (rowCount % BATCH_SIZE === 0) {
                            if (isFirstBatch) {
                                webviewPanel.webview.postMessage({
                                    command: 'initTable',
                                    headerHtml: generateTableHeaderHtml(columnCount),
                                    rowsHtml: generateTableRowsHtml(rows, 0)
                                });
                                isFirstBatch = false;
                            } else {
                                webviewPanel.webview.postMessage({
                                    command: 'appendRows',
                                    rowsHtml: generateTableRowsHtml(rows, rowCount - rows.length)
                                });
                            }
                            rows = [];
                        }
                    }
                });

                fileStream.on('end', () => {
                    if (leftover) {
                        rows.push(leftover.split(','));
                        if (rowCount === 0) {
                            columnCount = rows[0].length;
                        }
                        rowCount++;
                    }
                    if (rows.length > 0) {
                        if (isFirstBatch) {
                            webviewPanel.webview.postMessage({
                                command: 'initTable',
                                headerHtml: generateTableHeaderHtml(columnCount),
                                rowsHtml: generateTableRowsHtml(rows, 0)
                            });
                        } else {
                            webviewPanel.webview.postMessage({
                                command: 'appendRows',
                                rowsHtml: generateTableRowsHtml(rows, rowCount - rows.length)
                            });
                        }
                    }
                });

                fileStream.on('error', err => {
                    vscode.window.showErrorMessage(`Error reading CSV file: ${err}`);
                });
            };

            // Listen for messages
            webviewPanel.webview.onDidReceiveMessage(async message => {
                if (message.command === 'webviewReady') {
                    startStreaming();
                    return;
                }

                if (message.command === 'toggleView') {
                    if (!message.isTableView) {
                        await vscode.commands.executeCommand('vscode.openWith', document.uri, 'default');
                        webviewPanel.dispose();
                    }
                    return;
                }

                if (message.command === 'saveCsv') {
                    try {
                        const text = typeof message.text === 'string' ? message.text : '';
                        await vscode.workspace.fs.writeFile(document.uri, Buffer.from(text, 'utf8'));
                        webviewPanel.webview.postMessage({ command: 'saveResult', ok: true });
                    } catch (err) {
                        webviewPanel.webview.postMessage({ command: 'saveResult', ok: false, error: String(err) });
                    }
                }
            });

            // Fallback: don't block forever if the webview never sends webviewReady
            setTimeout(() => startStreaming(), 500);
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading CSV file: ${error}`);
        }
    }

    private getWebviewContent(tableHtml: string, webviewPanel: vscode.WebviewPanel): string {
        const webview = webviewPanel.webview;
        const imgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'view.png'));
        const svgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table.svg'));
        const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csvWebview.js'));
        const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csvWebview.css'));
        const cspSource = webview.cspSource;

        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource};">
            <meta name="viewport" width="device-width, initial-scale=1.0">
            <title>CSV Viewer</title>
            <link href="${styleUri}" rel="stylesheet" />
        </head>
        <body>
            <div class="header-background"></div>
            <div class="button-container">
                <button id="toggleViewButton" class="toggle-button" title="Edit File in Vscode Default Editor">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 20h9" />
                        <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z" />
                    </svg>
                    Edit File
                </button>
                <button id="toggleTableEditButton" class="toggle-button" title="Edit CSV directly in the table">
                    Edit Table
                </button>
                <button id="saveTableEditsButton" class="toggle-button hidden" title="Save table edits">
                    Save
                </button>
                <button id="cancelTableEditsButton" class="toggle-button hidden" title="Cancel table edits">
                    Cancel
                </button>
                <button id="toggleExpandButton" class="toggle-button" title="Toggle Column Widths (Default / Expand All)">
                    <svg id="expandIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: block;">
                        <path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>
                    </svg>
                    <svg id="collapseIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: none;">
                        <path d="M4 14h6v6M20 10h-6V4M14 10l7-7M10 14l-7 7"/>
                    </svg>
                    <span id="expandButtonText">Expand</span>
                </button>
                <button id="toggleBackgroundButton" class="toggle-button" title="Toggle Light/Dark Mode">
                    <svg id="lightIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: none;">
                        <circle cx="12" cy="12" r="5"/>
                        <line x1="12" y1="1" x2="12" y2="3"/>
                        <line x1="12" y1="21" x2="12" y2="23"/>
                        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/>
                        <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
                        <line x1="1" y1="12" x2="3" y2="12"/>
                        <line x1="21" y1="12" x2="23" y2="12"/>
                        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/>
                        <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
                    </svg>
                    <svg id="darkIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>
                    </svg>
                </button>
                <div class="tooltip">
                    <img src="${imgUri}" alt="Change to table view"  style="width: auto; height: 32px; margin-left: auto; margin-top: 2px;" />
                    <span class="tooltiptext">
                        <span class="warning">Important:</span> Click the blue table icon <img src="${svgUri}" alt="Table Icon" style="width: 16px; vertical-align: middle; height: 16px;" />
                         to switch to table view from edit file mode. <br>
                        <span class="instruction">The table icon will only work on edit file mode and is located on the top right corner in the editor toolbar as shown in the image.</span>
                    </span>
                </div>
            </div>
            <div id="content">${tableHtml}</div>
            <div class="selection-info" id="selectionInfo"></div>
            <noscript>
                <div style="padding: 8px; margin-top: 10px; background: #fff3cd; border: 1px solid #ffeeba;">
                    JavaScript is disabled in this webview, so the CSV table cannot load.
                </div>
            </noscript>
            <script src="${scriptUri}"></script>
        </body>
        </html>`;
    }
}