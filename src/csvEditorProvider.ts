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
            const filePath = document.uri.fsPath;
            
            // Storage for parsed CSV data
            let allRows: string[][] = [];
            let columnCount = 0;
            let parseComplete = false;

            // CSV parsing helpers
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
                            i++;
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

            function excelColumnLabel(n: number): string {
                let label = '';
                while (n > 0) {
                    const rem = (n - 1) % 26;
                    label = String.fromCharCode(65 + rem) + label;
                    n = Math.floor((n - 1) / 26);
                }
                return label;
            }

            // Set up webview
            webviewPanel.webview.options = {
                enableScripts: true,
                localResourceRoots: [
                    vscode.Uri.joinPath(this.context.extensionUri, 'resources'),
                    vscode.Uri.joinPath(this.context.extensionUri, 'dist')
                ]
            };
            webviewPanel.webview.html = this.getWebviewContent(webviewPanel);

            // Parse the entire CSV file
            const parseCSV = (): Promise<void> => {
                return new Promise((resolve, reject) => {
                    let leftover = '';
                    const fileStream = fs.createReadStream(filePath, { encoding: 'utf-8' });

                    webviewPanel.onDidDispose(() => {
                        try { fileStream.destroy(); } catch { }
                    });

                    fileStream.on('data', (chunk: string) => {
                        const parsed = parseCsvChunk(leftover + chunk);
                        leftover = parsed.leftover;
                        for (const row of parsed.rows) {
                            if (allRows.length === 0 && row.length > 0) {
                                columnCount = row.length;
                            }
                            allRows.push(row);
                        }
                    });

                    fileStream.on('end', () => {
                        // Handle final partial row
                        if (leftover && leftover.length > 0) {
                            const row = parseRowString(leftover);
                            if (allRows.length === 0 && row.length > 0) {
                                columnCount = row.length;
                            }
                            allRows.push(row);
                        }
                        parseComplete = true;
                        resolve();
                    });

                    fileStream.on('error', (err) => {
                        reject(err);
                    });
                });
            };

            // Handle messages from webview
            webviewPanel.webview.onDidReceiveMessage(async message => {
                switch (message.command) {
                    case 'webviewReady':
                        try {
                            // Parse the CSV first
                            await parseCSV();

                            // Generate header HTML
                            let headerHtml = '<tr><th class="row-header">&nbsp;</th>';
                            for (let i = 1; i <= columnCount; i++) {
                                headerHtml += `<th class="col-header" data-col="${i - 1}">${excelColumnLabel(i)}</th>`;
                            }
                            headerHtml += '</tr>';

                            // Send initial metadata to webview
                            webviewPanel.webview.postMessage({
                                command: 'initVirtualTable',
                                headerHtml,
                                totalRows: allRows.length,
                                columnCount,
                                format: 'csv'
                            });

                            // Send settings
                            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                            const settings = {
                                firstRowIsHeader: cfg.get('csv.firstRowIsHeader', false),
                                stickyHeader: cfg.get('csv.stickyHeader', false),
                                stickyToolbar: cfg.get('csv.stickyToolbar', true)
                            };
                            webviewPanel.webview.postMessage({ command: 'initSettings', settings });

                            // Send theme
                            webviewPanel.webview.postMessage({ 
                                type: 'setTheme', 
                                kind: vscode.window.activeColorTheme.kind 
                            });
                        } catch (err) {
                            vscode.window.showErrorMessage(`Error parsing CSV: ${err}`);
                        }
                        break;

                    case 'getRows':
                        if (parseComplete) {
                            const { start, end, requestId } = message;
                            const clampedStart = Math.max(0, start);
                            const clampedEnd = Math.min(allRows.length, end);
                            const rows = allRows.slice(clampedStart, clampedEnd);
                            
                            webviewPanel.webview.postMessage({
                                command: 'rowsData',
                                rows,
                                start: clampedStart,
                                end: clampedEnd,
                                requestId
                            });
                        }
                        break;

                    case 'getRowCount':
                        webviewPanel.webview.postMessage({
                            command: 'rowCount',
                            totalRows: allRows.length,
                            requestId: message.requestId
                        });
                        break;

                    case 'updateSettings':
                        try {
                            const s = message.settings || {};
                            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                            await cfg.update('csv.firstRowIsHeader', !!s.firstRowIsHeader, vscode.ConfigurationTarget.Global);
                            await cfg.update('csv.stickyHeader', !!s.stickyHeader, vscode.ConfigurationTarget.Global);
                            await cfg.update('csv.stickyToolbar', !!s.stickyToolbar, vscode.ConfigurationTarget.Global);
                        } catch (err) {
                            console.error('Failed to persist settings:', err);
                        }
                        break;

                    case 'toggleView':
                        if (!message.isTableView) {
                            await vscode.commands.executeCommand('vscode.openWith', document.uri, 'default');
                            webviewPanel.dispose();
                        }
                        break;

                    case 'saveCsv':
                        try {
                            const text = typeof message.text === 'string' ? message.text : '';
                            // Update in-memory data after save
                            const lines: string[] = text.split('\n').filter((l: string) => l.length > 0 || text.endsWith('\n'));
                            allRows = [];
                            for (const line of lines) {
                                if (line.trim() || lines.indexOf(line) < lines.length - 1) {
                                    allRows.push(parseRowString(line));
                                }
                            }
                            await vscode.workspace.fs.writeFile(document.uri, Buffer.from(text, 'utf8'));
                            webviewPanel.webview.postMessage({ command: 'saveResult', ok: true });
                        } catch (err) {
                            webviewPanel.webview.postMessage({ command: 'saveResult', ok: false, error: String(err) });
                        }
                        break;

                    case 'updateRow':
                        // Update a single row in memory (for edit mode)
                        if (message.rowIndex !== undefined && message.rowData) {
                            allRows[message.rowIndex] = message.rowData;
                        }
                        break;

                    case 'openExternal':
                        try {
                            const url = typeof message.url === 'string' ? message.url : '';
                            if (url) {
                                await vscode.env.openExternal(vscode.Uri.parse(url));
                            }
                        } catch {
                            // ignore
                        }
                        break;
                }
            });

            // Forward settings changes
            const configChangeDisposable = vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration('xlsxViewer.csv') || e.affectsConfiguration('xlsxViewer')) {
                    const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                    const settings = {
                        firstRowIsHeader: cfg.get('csv.firstRowIsHeader', false),
                        stickyHeader: cfg.get('csv.stickyHeader', false),
                        stickyToolbar: cfg.get('csv.stickyToolbar', true)
                    };
                    try { 
                        webviewPanel.webview.postMessage({ command: 'settingsUpdated', settings }); 
                    } catch { }
                }
            });

            // Theme change listener
            const themeChangeDisposable = vscode.window.onDidChangeActiveColorTheme(() => {
                try { 
                    webviewPanel.webview.postMessage({ 
                        type: 'setTheme', 
                        kind: vscode.window.activeColorTheme.kind 
                    }); 
                } catch { }
            });

            webviewPanel.onDidDispose(() => { 
                configChangeDisposable.dispose(); 
                themeChangeDisposable.dispose(); 
            });

        } catch (error) {
            vscode.window.showErrorMessage(`Error reading CSV file: ${error}`);
        }
    }

    private getWebviewContent(webviewPanel: vscode.WebviewPanel): string {
        const webview = webviewPanel.webview;
        const imgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table', 'view.png'));
        const svgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table', 'table.svg'));
        const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'table', 'tableWebview.js'));
        const themeStyleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'shared', 'theme.css'));
        const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table', 'tableWebview.css'));
        const cspSource = webview.cspSource;

        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource} 'unsafe-inline';">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>CSV Viewer</title>
            <link href="${themeStyleUri}" rel="stylesheet" />
            <link href="${styleUri}" rel="stylesheet" />
            <script>
                window.viewImgUri = "${imgUri}";
                window.logoSvgUri = "${svgUri}";
            </script>
        </head>
        <body>
            <div class="header-background"></div>
            <div class="toolbar" id="toolbar"></div>
            <div id="content">
                <div id="loadingIndicator" class="loading-indicator">Loading CSV...</div>
                <div class="table-scroll" id="tableContainer">
                    <table id="csv-table">
                        <colgroup></colgroup>
                        <thead></thead>
                        <tbody></tbody>
                    </table>
                </div>
            </div>
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