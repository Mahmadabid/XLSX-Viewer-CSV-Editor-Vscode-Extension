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
            function excelColumnLabel(n: number): string {
                // Convert 1 -> A, 26 -> Z, 27 -> AA, etc.
                let label = '';
                while (n > 0) {
                    const rem = (n - 1) % 26;
                    label = String.fromCharCode(65 + rem) + label;
                    n = Math.floor((n - 1) / 26);
                }
                return label;
            }

            function generateTableHeaderHtml(colCount: number): string {
                let html = '<thead><tr><th class="row-header">&nbsp;</th>';
                for (let colNumber = 1; colNumber <= colCount; colNumber++) {
                    const colLabel = excelColumnLabel(colNumber);
                    html += `<th class="col-header" data-col="${colNumber - 1}">${colLabel}</th>`;
                }
                html += '</tr></thead>';
                return html;
            }

            // Robust CSV parsing helpers (handles quoted fields, embedded commas and newlines)
            function parseRowString(rowStr: string): string[] {
                if (rowStr.endsWith('\r')) rowStr = rowStr.slice(0, -1);
                const fields: string[] = [];
                let field = '';
                let inQuotes = false;
                for (let i = 0; i < rowStr.length; i++) {
                    const ch = rowStr[i];
                    if (ch === '"') {
                        if (inQuotes && i + 1 < rowStr.length && rowStr[i + 1] === '"') {
                            field += '"';
                            i++;
                        } else {
                            inQuotes = !inQuotes;
                        }
                        continue;
                    }
                    if (ch === ',' && !inQuotes) {
                        fields.push(field);
                        field = '';
                        continue;
                    }
                    field += ch;
                }
                fields.push(field);
                return fields;
            }

            function parseCsvChunk(data: string): { rows: string[][]; leftover: string } {
                const rows: string[][] = [];
                let inQuotes = false;
                let lastRowEnd = 0;
                for (let i = 0; i < data.length; i++) {
                    const ch = data[i];
                    if (ch === '"') {
                        if (inQuotes && i + 1 < data.length && data[i + 1] === '"') {
                            i++; // skip escaped quote
                        } else {
                            inQuotes = !inQuotes;
                        }
                    } else if (!inQuotes && ch === '\n') {
                        const rowStr = data.slice(lastRowEnd, i);
                        rows.push(parseRowString(rowStr));
                        lastRowEnd = i + 1;
                    }
                }
                const leftover = data.slice(lastRowEnd);
                return { rows, leftover };
            }

            // Set up webview
            webviewPanel.webview.options = {
                enableScripts: true,
                localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'resources')]
            };
            webviewPanel.webview.html = this.getWebviewContent(
                `<div class="table-scroll"><table id="csv-table">
                    <colgroup></colgroup>
                    <thead></thead><tbody></tbody>
                </table></div>`,
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
                    const parsed = parseCsvChunk(leftover + chunk);
                    leftover = parsed.leftover;
                    for (let parsedRow of parsed.rows) {
                        rows.push(parsedRow);
                        if (rowCount === 0 && rows[0]) {
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
                    // Process any remaining complete rows
                    const parsed = parseCsvChunk(leftover);
                    leftover = parsed.leftover;
                    for (let parsedRow of parsed.rows) {
                        rows.push(parsedRow);
                        if (rowCount === 0 && rows[0]) {
                            columnCount = rows[0].length;
                        }
                        rowCount++;
                    }
                    // If leftover still contains data (final partial row), push it as last row
                    if (leftover && leftover.length > 0) {
                        rows.push(parseRowString(leftover));
                        if (rowCount === 0 && rows[0]) {
                            columnCount = rows[0].length;
                        }
                        rowCount++;
                        leftover = '';
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
                    // Send current persisted settings to the webview so it can apply them
                    startStreaming();
                    const cfg = vscode.workspace.getConfiguration('xlsxViewer');

                    // Send current theme info to webview
                    try {
                        webviewPanel.webview.postMessage({ type: 'setTheme', kind: vscode.window.activeColorTheme.kind });
                    } catch { /* ignore */ }                    
                    return;                
                    const settings = {
                        firstRowIsHeader: cfg.get('csv.firstRowIsHeader', false),
                        stickyHeader: cfg.get('csv.stickyHeader', false),
                        stickyToolbar: cfg.get('csv.stickyToolbar', true)
                    };
                    webviewPanel.webview.postMessage({ command: 'initSettings', settings });
                    return;
                }

                if (message.command === 'updateSettings') {
                    try {
                        const s = message.settings || {};
                        const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                        // Persist globally (user scope) so setting applies to all CSV files
                        await cfg.update('csv.firstRowIsHeader', !!s.firstRowIsHeader, vscode.ConfigurationTarget.Global);
                        await cfg.update('csv.stickyHeader', !!s.stickyHeader, vscode.ConfigurationTarget.Global);
                        await cfg.update('csv.stickyToolbar', !!s.stickyToolbar, vscode.ConfigurationTarget.Global);
                    } catch (err) {
                        console.error('Failed to persist settings:', err);
                    }
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

            // Forward settings changes made outside the webview to the webview
            const configChangeDisposable = vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration('xlsxViewer.csv') || e.affectsConfiguration('xlsxViewer')) {
                    const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                    const settings = {
                        firstRowIsHeader: cfg.get('csv.firstRowIsHeader', false),
                        stickyHeader: cfg.get('csv.stickyHeader', false),
                        stickyToolbar: cfg.get('csv.stickyToolbar', true)
                    };
                    try { webviewPanel.webview.postMessage({ command: 'settingsUpdated', settings }); } catch { }
                }
            });

            // Sync VS Code theme changes to the webview
            const themeChangeDisposable = vscode.window.onDidChangeActiveColorTheme(() => {
                try { webviewPanel.webview.postMessage({ type: 'setTheme', kind: vscode.window.activeColorTheme.kind }); } catch { }
            });

            webviewPanel.onDidDispose(() => { configChangeDisposable.dispose(); themeChangeDisposable.dispose(); });

            // Fallback: don't block forever if the webview never sends webviewReady
            setTimeout(() => startStreaming(), 500);
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading CSV file: ${error}`);
        }
    }

    private getWebviewContent(tableHtml: string, webviewPanel: vscode.WebviewPanel): string {
        const webview = webviewPanel.webview;
        const imgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csv', 'view.png'));
        const svgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csv', 'table.svg'));
        const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csv', 'csvWebview.js'));
        const themeScriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'theme', 'themeManager.js'));
        const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'csv', 'csvWebview.css'));
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
            <div class="toolbar">
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
                <!-- Settings button (appears after Expand, before Theme) -->
                <button id="openSettingsButton" class="toggle-button icon-only" title="CSV Settings">
                   <svg fill="#ffffff" version="1.1" id="Capa_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 389.663 389.663" xml:space="preserve"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <g> <path d="M194.832,132.997c-34.1,0-61.842,27.74-61.842,61.838c0,34.1,27.742,61.841,61.842,61.841 c34.099,0,61.841-27.741,61.841-61.841C256.674,160.737,228.932,132.997,194.832,132.997z M194.832,226.444 c-17.429,0-31.608-14.182-31.608-31.61c0-17.428,14.18-31.605,31.608-31.605c17.429,0,31.607,14.178,31.607,31.605 C226.439,212.264,212.262,226.444,194.832,226.444z"></path> <path d="M385.23,150.784c-2.816-2.812-6.714-4.427-10.688-4.427l-49.715,0.015l-3.799-9.194l35.149-35.155 c5.892-5.894,5.892-15.483,0-21.377l-47.166-47.162c-2.688-2.691-6.586-4.235-10.688-4.235c-4.103,0-7.996,1.544-10.687,4.235 L252.48,68.639l-9.188-3.797V15.116C243.292,6.781,236.511,0,228.177,0h-66.694c-8.335,0-15.116,6.78-15.116,15.115v49.716 l-9.194,3.801l-35.151-35.135c-2.855-2.854-6.65-4.426-10.686-4.426c-4.036,0-7.832,1.572-10.688,4.427L33.476,80.67 c-2.813,2.814-4.427,6.711-4.427,10.688c0,3.984,1.613,7.882,4.427,10.693l35.151,35.127l-3.811,9.188l-49.697,0.005 C6.781,146.372,0,153.153,0,161.488v66.708c0,4.035,1.573,7.832,4.431,10.689c2.817,2.815,6.713,4.432,10.688,4.432l49.708-0.021 l3.799,9.195l-35.133,35.149c-5.894,5.896-5.894,15.484,0,21.378l47.161,47.172c2.692,2.69,6.591,4.233,10.693,4.233 c4.105,0,8.002-1.543,10.69-4.233l35.136-35.162l9.186,3.815l0.008,49.691c0,8.338,6.781,15.121,15.116,15.121l66.708,0.006h0.162 c8.336,0,15.116-6.781,15.116-15.117c0-0.721-0.049-1.444-0.147-2.151l-0.015-0.207l-0.013-47.355l9.195-3.801l35.149,35.139 c2.855,2.857,6.65,4.432,10.688,4.432c4.035,0,7.83-1.573,10.686-4.432l47.172-47.166c2.855-2.854,4.429-6.649,4.429-10.688 c0-4.045-1.572-7.847-4.429-10.699l-35.157-35.125l3.809-9.195h49.707c8.336,0,15.119-6.78,15.119-15.114v-66.708 C389.662,157.438,388.088,153.641,385.23,150.784z M359.428,213.063h-44.696c-6.134,0-11.615,3.662-13.966,9.328l-11.534,27.865 c-2.351,5.672-1.062,12.141,3.274,16.482l31.609,31.58l-25.789,25.789l-31.605-31.603c-2.854-2.853-6.649-4.422-10.69-4.422 c-1.992,0-3.938,0.388-5.785,1.147l-27.854,11.537c-5.666,2.349-9.327,7.832-9.327,13.972l0.008,44.688l-36.468-0.01 l-0.008-44.686c0-6.136-3.661-11.615-9.328-13.966l-27.856-11.536c-1.854-0.768-3.806-1.155-5.802-1.155 c-4.036,0-7.829,1.571-10.677,4.43l-31.586,31.615L65.559,298.33l31.592-31.604c4.339-4.343,5.625-10.81,3.275-16.478 L88.89,222.393c-2.352-5.666-7.833-9.328-13.965-9.328l-44.688,0.01v-36.466l44.688-0.01c6.134,0,11.615-3.662,13.965-9.328 l11.536-27.854c2.349-5.676,1.063-12.146-3.275-16.482L65.548,91.359l25.79-25.796l31.599,31.582 c2.856,2.857,6.658,4.43,10.704,4.43c1.988,0,3.928-0.385,5.764-1.144l27.861-11.524c5.671-2.351,9.336-7.834,9.336-13.97V30.231 h36.459v44.705c0,6.137,3.662,11.618,9.328,13.965l27.855,11.534c1.848,0.766,3.795,1.153,5.789,1.153 c4.039,0,7.832-1.572,10.684-4.429l31.607-31.617l25.789,25.789l-31.609,31.607c-4.336,4.339-5.621,10.806-3.274,16.478 l11.534,27.858c2.351,5.669,7.832,9.332,13.966,9.332l44.696-0.01L359.428,213.063L359.428,213.063z"></path> </g> </g> </g></svg>
                </button>

                <div id="settingsPanel" class="settings-panel hidden" role="dialog" aria-hidden="true">
                    <div class="settings-group">
                        <label class="setting-item"><input type="checkbox" id="chkHeaderRow"/> <span>Header Row</span></label>
                        <label class="setting-item"><input type="checkbox" id="chkStickyHeader"/> <span>Sticky Header</span></label>
                        <label class="setting-item"><input type="checkbox" id="chkStickyToolbar"/> <span>Sticky Toolbar</span></label>
                    </div>
                    
                    <button id="settingsCancelButton" class="toggle-button" title="Close">Close</button>
                </div>

                <button id="toggleBackgroundButton" class="toggle-button" title="Toggle Theme (Light / Dark / VS Code)">
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
                    <svg id="vscodeIcon" width="24px" height="24px" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><rect x="0" fill="none" width="24" height="24"/><g><path fill="currentColor" d="M4 6c-1.105 0-2 .895-2 2v12c0 1.1.9 2 2 2h12c1.105 0 2-.895 2-2H4V6zm16-4H8c-1.105 0-2 .895-2 2v12c0 1.105.895 2 2 2h12c1.105 0 2-.895 2-2V4c0-1.105-.895-2-2-2zm-5 14H8V9h7v7zm5 0h-3V9h3v7zm0-9H8V4h12v3z"/></g></svg>        
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
            <script src="${themeScriptUri}"></script>
            <script src="${scriptUri}"></script>
        </body>
        </html>`;
    }
}