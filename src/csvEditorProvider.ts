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
                `<table id="csv-table" border="1" cellspacing="0" cellpadding="5">
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
        const cspSource = webview.cspSource;

        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource};">
            <meta name="viewport" width="device-width, initial-scale=1.0">
            <title>CSV Viewer</title>
            <style>
                body { 
                    font-family: sans-serif; 
                    padding: 10px; 
                    background-color: rgb(255, 255, 255);
                    margin: 0;
                    overflow-x: auto;
                }
                table { 
                    border-collapse: collapse; 
                    width: auto;
                    min-width: 100%;
                    table-layout: fixed;
                    user-select: none;
                    -webkit-user-select: none;
                    -moz-user-select: none;
                    -ms-user-select: none;
                }
                th, td { 
                    border: 1px solid #ccc; 
                    padding: 8px; 
                    text-align: left;
                    white-space: nowrap;
                    position: relative;
                    cursor: cell;
                    min-width: 80px;
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
                
                th {
                    background-color: rgb(247, 247, 247);
                    font-weight: bold;
                }
                
                td:nth-child(1), th:nth-child(1) {
                    width: 50px !important;
                    min-width: 50px !important;
                    background-color: rgb(247, 247, 247);
                    text-align: center;
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
                
                .toggle-button svg {
                    width: 20px;
                    height: 20px;
                    stroke: white;
                }
                
                td { 
                    background-color: rgb(255, 255, 255);
                    color: rgb(0, 0, 0);
                    padding: 4px 8px;
                    border: 1px solid #e2e3e3;
                }
                
                td span.cell-content {
                    display: block;
                    user-select: none;
                    pointer-events: none;
                }
                
                body.alt-bg { 
                    background-color: rgb(33, 33, 33); 
                }
                
                body.alt-bg td { 
                    background-color: rgb(33, 33, 33) !important;
                    color: rgb(255, 255, 255);
                }
                
                body.alt-bg th.col-header, 
                body.alt-bg th.row-header { 
                    background-color: rgb(69, 69, 69); 
                    color: #fff; 
                }

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

                /* Cell selection styles */
                td.selected {
                    border: 2px solid rgb(26, 115, 232) !important;
                    background-color: rgba(26, 115, 232, 0.1) !important;
                    z-index: 2;
                }

                body.alt-bg td.selected {
                    background-color: rgba(138, 180, 248, 0.24) !important;
                    border: 2px solid rgb(138, 180, 248) !important;                
                }
                
                td.active-cell {
                    border: 2px solid rgb(26, 115, 232) !important;
                    background-color: white !important;
                    z-index: 3;
                }

                body.alt-bg td.active-cell {
                    background-color: rgb(33, 33, 33) !important;
                    border: 2px solid rgb(138, 180, 248) !important;
                }

                td.active-cell::after {
                    content: '';
                    position: absolute;
                    right: -2px;
                    bottom: -2px;
                    width: 6px;
                    height: 6px;
                    background: #1a73e8;
                    border: 1px solid white;
                }

                /* Row and Column selection styles */
                td.column-selected, th.column-selected {
                    background-color: rgba(26, 115, 232, 0.1) !important;
                    border-left: 2px solid rgb(26, 115, 232) !important;
                    border-right: 2px solid rgb(26, 115, 232) !important;
                }

                td.row-selected, th.row-selected {
                    background-color: rgba(26, 115, 232, 0.1) !important;
                    border-top: 2px solid rgb(26, 115, 232) !important;
                    border-bottom: 2px solid rgb(26, 115, 232) !important;
                }

                body.alt-bg td.column-selected,
                body.alt-bg th.column-selected,
                body.alt-bg td.row-selected,
                body.alt-bg th.row-selected {
                    background-color: rgba(138, 180, 248, 0.24) !important;
                }

                th.col-header, th.row-header {
                    cursor: pointer;
                    user-select: none;
                }

                th.col-header:hover, th.row-header:hover {
                    background-color: rgba(26, 115, 232, 0.2);
                }

                /* Copy animation */
                td.copying {
                    animation: copyFlash 0.2s ease-in-out;
                }

                @keyframes copyFlash {
                    0% { background-color: inherit; }
                    50% { background-color: rgba(26, 115, 232, 0.3) !important; }
                    100% { background-color: inherit; }
                }

                body.alt-bg td.copying {
                    animation: copyFlashDark 0.2s ease-in-out;
                }

                @keyframes copyFlashDark {
                    0% { background-color: inherit; }
                    50% { background-color: rgba(138, 180, 248, 0.3) !important; }
                    100% { background-color: inherit; }
                }

                ::selection {
                    background-color: transparent;
                }
                
                ::-moz-selection {
                    background-color: transparent;
                }

                td:hover {
                    background-color: rgba(0, 0, 0, 0.05) !important;
                }

                body.alt-bg td:hover {
                    background-color: rgba(255, 255, 255, 0.1) !important;
                }

                .selection-info {
                    position: fixed;
                    bottom: 10px;
                    right: 10px;
                    background: rgba(0, 0, 0, 0.8);
                    color: white;
                    padding: 5px 10px;
                    border-radius: 4px;
                    font-size: 12px;
                    display: none;
                    z-index: 1001;
                }

                body.alt-bg .selection-info {
                    background: rgba(255, 255, 255, 0.8);
                    color: black;
                                    }

                .hidden { display: none !important; }

                body.edit-mode td {
                    user-select: text;
                    cursor: text;
                }
            </style>
        </head>
        <body>
            <div class="button-container">
                <button id="toggleViewButton" class="toggle-button" title="Edit File in Vscode Default Editor">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
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
            <script>
                const vscode = acquireVsCodeApi();
                let isTableView = true;
                let isSelecting = false;
                let startCell = null;
                let endCell = null;
                let selectedCells = new Set();
                let activeCell = null;
                let selectedRows = new Set();
                let selectedColumns = new Set();
                let lastSelectedRow = null;
                let lastSelectedColumn = null;

                window.addEventListener('message', event => {
                    const message = event.data;
                    if (message.command === 'initTable') {
                        const table = document.getElementById('csv-table');
                        table.querySelector('thead').innerHTML = message.headerHtml;
                        table.querySelector('tbody').innerHTML = message.rowsHtml;
                        initializeSelection();
                    } else if (message.command === 'appendRows') {
                        const table = document.getElementById('csv-table');
                        table.querySelector('tbody').insertAdjacentHTML('beforeend', message.rowsHtml);
                        initializeSelection();
                    }
                });

                document.getElementById('toggleViewButton').addEventListener('click', () => {
                    isTableView = !isTableView;
                    vscode.postMessage({ command: 'toggleView', isTableView });
                });

                document.getElementById('toggleBackgroundButton').addEventListener('click', () => {
                    document.body.classList.toggle('alt-bg');
                    const isDarkMode = document.body.classList.contains('alt-bg');

                    if (isDarkMode) {
                        document.getElementById('lightIcon').style.display = 'block';
                        document.getElementById('darkIcon').style.display = 'none';
                    } else {
                        document.getElementById('lightIcon').style.display = 'none';
                        document.getElementById('darkIcon').style.display = 'block';
                    }

                    const defaultBgCells = document.querySelectorAll('td[data-default-bg="true"]');
                    defaultBgCells.forEach(cell => {
                        if (isDarkMode) {
                            cell.style.backgroundColor = "rgb(33, 33, 33)";
                        } else {
                            cell.style.backgroundColor = "rgb(255, 255, 255)";
                        }
                    });

                    const defaultBothCells = document.querySelectorAll('td[data-default-bg="true"][data-default-color="true"]');
                    defaultBothCells.forEach(cell => {
                        if (isDarkMode) {
                            cell.style.color = "rgb(255, 255, 255)";
                        } else {
                            cell.style.color = "rgb(0, 0, 0)";
                        }
                    });
                });

                function getCellCoordinates(cell) {
                    if (!cell || !cell.dataset) return null;
                    return {
                        row: parseInt(cell.dataset.row),
                        col: parseInt(cell.dataset.col)
                    };
                }

                function clearSelection() {
                    document.querySelectorAll('td.selected, td.active-cell, td.column-selected, td.row-selected, th.column-selected, th.row-selected').forEach(el => {
                        el.classList.remove('selected', 'active-cell', 'column-selected', 'row-selected', 'copying');
                    });
                    selectedCells.clear();
                    selectedRows.clear();
                    selectedColumns.clear();
                    activeCell = null;
                    lastSelectedRow = null;
                    lastSelectedColumn = null;
                    document.getElementById('selectionInfo').style.display = 'none';
                }
                        } else {
                            selectedColumns.add(columnIndex);
                            const cells = document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]');
                            cells.forEach(cell => {
                                cell.classList.add('column-selected');
                                if (cell.tagName === 'TD') selectedCells.add(cell);
                            });
                        }
                        lastSelectedColumn = columnIndex;
                    }
                    updateSelectionInfo();
                }

                function selectRow(rowIndex, ctrlKey, shiftKey) {
                    if (!ctrlKey && !shiftKey) {
                        clearSelection();
                    }

                    if (shiftKey && lastSelectedRow !== null) {
                        if (!ctrlKey) {
                            clearSelection();
                        }
                        const minRow = Math.min(lastSelectedRow, rowIndex);
                        const maxRow = Math.max(lastSelectedRow, rowIndex);
                        for (let row = minRow; row <= maxRow; row++) {
                            selectedRows.add(row);
                            const rowHeader = document.querySelector('th[data-row="' + row + '"]');
                            if (rowHeader) {
                                const rowElement = rowHeader.parentElement;
                                const cells = rowElement.querySelectorAll('td, th');
                                cells.forEach(cell => {
                                    cell.classList.add('row-selected');
                                    if (cell.tagName === 'TD') selectedCells.add(cell);
                                });
                            }
                        }
                    } else {
                        if (ctrlKey && selectedRows.has(rowIndex)) {
                            selectedRows.delete(rowIndex);
                            const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
                            if (rowHeader) {
                                const rowElement = rowHeader.parentElement;
                                const cells = rowElement.querySelectorAll('td, th');
                                cells.forEach(cell => {
                                    cell.classList.remove('row-selected');
                                    if (cell.tagName === 'TD') selectedCells.delete(cell);
                                });
                            }
                        } else {
                            selectedRows.add(rowIndex);
                            const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
                            if (rowHeader) {
                                const rowElement = rowHeader.parentElement;
                                const cells = rowElement.querySelectorAll('td, th');
                                cells.forEach(cell => {
                                    cell.classList.add('row-selected');
                                    if (cell.tagName === 'TD') selectedCells.add(cell);
                                });
                            }
                        }
                        lastSelectedRow = rowIndex;
                    }
                    updateSelectionInfo();
                }

                function updateSelectionInfo() {
                    if (selectedCells.size > 1) {
                        const cellsArray = Array.from(selectedCells);
                        const rows = new Set(cellsArray.map(cell => parseInt(cell.dataset.row)));
                        const cols = new Set(cellsArray.map(cell => parseInt(cell.dataset.col)));
                        document.getElementById('selectionInfo').textContent = rows.size + 'R × ' + cols.size + 'C';
                        document.getElementById('selectionInfo').style.display = 'block';
                    } else {
                        document.getElementById('selectionInfo').style.display = 'none';
                    }
                }

                function copySelectionToClipboard() {
                    if (selectedCells.size === 0) return;

                    // Convert selected cells to array and sort by position
                    const cellsArray = Array.from(selectedCells);
                    const cellData = cellsArray.map(cell => ({
                        row: parseInt(cell.dataset.row),
                        col: parseInt(cell.dataset.col),
                        text: cell.textContent.trim()
                    }));

                    // Sort by row then column
                    cellData.sort((a, b) => a.row - b.row || a.col - b.col);

                    // Find bounds
                    const minRow = Math.min(...cellData.map(c => c.row));
                    const maxRow = Math.max(...cellData.map(c => c.row));
                    const minCol = Math.min(...cellData.map(c => c.col));
                    const maxCol = Math.max(...cellData.map(c => c.col));

                    // Create a 2D array to represent the selection
                    const grid = [];
                    for (let r = minRow; r <= maxRow; r++) {
                        const row = [];
                        for (let c = minCol; c <= maxCol; c++) {
                            const cellData = cellsArray.find(cell => 
                                parseInt(cell.dataset.row) === r && parseInt(cell.dataset.col) === c
                            );
                            row.push(cellData ? cellData.textContent.trim() : '');
                        }
                        grid.push(row);
                    }

                    // Convert to tab-separated format
                    const clipboardText = grid.map(row => row.join('\\t')).join('\\n');
                    
                    // Copy to clipboard
                    navigator.clipboard.writeText(clipboardText).then(() => {
                        // Visual feedback
                        selectedCells.forEach(cell => {
                            cell.classList.add('copying');
                            setTimeout(() => {
                                cell.classList.remove('copying');
                            }, 200);
                        });
                    });
                }

                function initializeSelection() {
    const table = document.getElementById('csv-table');
    if (!table) return;

    // Remove any existing event listeners first
    const newTable = table.cloneNode(true);
    table.parentNode.replaceChild(newTable, table);
    const tableElement = document.getElementById('csv-table');

    tableElement.addEventListener('selectstart', (e) => {
        e.preventDefault();
        return false;
    });

    tableElement.addEventListener('mousedown', (e) => {
        const target = e.target.closest('td, th');
        if (!target) return;

        e.preventDefault(); // Prevent text selection

        // Handle column header clicks
        if (target.classList.contains('col-header')) {
            const columnIndex = parseInt(target.dataset.col);
            if (isNaN(columnIndex)) return;
            
            // Always update last selected column
            if (!e.shiftKey) {
                lastSelectedColumn = columnIndex;
            }
            
            selectColumn(columnIndex, e.ctrlKey || e.metaKey, e.shiftKey);
            return;
        }

        // Handle row header clicks
        if (target.classList.contains('row-header')) {
            const rowIndex = parseInt(target.dataset.row);
            if (isNaN(rowIndex)) return;
            
            // Always update last selected row
            if (!e.shiftKey) {
                lastSelectedRow = rowIndex;
            }
            
            selectRow(rowIndex, e.ctrlKey || e.metaKey, e.shiftKey);
            return;
        }

        // Handle regular cell clicks
        if (target.tagName === 'TD') {
            const coords = getCellCoordinates(target);
            if (!coords) return;

            // Ctrl/Cmd + Click: Toggle individual cell
            if (e.ctrlKey || e.metaKey) {
                e.stopPropagation();
                
                // Don't clear other selections
                if (target.classList.contains('selected')) {
                    // Remove from selection
                    target.classList.remove('selected');
                    selectedCells.delete(target);
                    
                    if (target === activeCell) {
                        target.classList.remove('active-cell');
                        activeCell = null;
                        
                        // Find another selected cell to make active
                        const remainingSelected = document.querySelector('td.selected');
                        if (remainingSelected) {
                            remainingSelected.classList.add('active-cell');
                            activeCell = remainingSelected;
                            startCell = getCellCoordinates(remainingSelected);
                        }
                    }
                } else {
                    // Add to selection
                    target.classList.add('selected');
                    selectedCells.add(target);
                    
                    // Update active cell
                    if (activeCell) {
                        activeCell.classList.remove('active-cell');
                    }
                    target.classList.add('active-cell');
                    activeCell = target;
                    startCell = coords;
                }
                
                updateSelectionInfo();
                return;
            }

            // Shift + Click: Range selection
            if (e.shiftKey && startCell) {
                e.stopPropagation();
                selectCellsInRange(startCell, coords);
                updateSelectionInfo();
                return;
            }

            // Normal click: Start new selection
            clearSelection();
            isSelecting = true;
            startCell = coords;
            endCell = coords;
            target.classList.add('selected');
            target.classList.add('active-cell');
            selectedCells.add(target);
            activeCell = target;
            updateSelectionInfo();
        }
    });

    tableElement.addEventListener('mousemove', (e) => {
        if (!isSelecting || !startCell) return;

        const target = e.target.closest('td');
        if (!target) return;

        const coords = getCellCoordinates(target);
        if (!coords) return;

        if (!endCell || coords.row !== endCell.row || coords.col !== endCell.col) {
            endCell = coords;
            selectCellsInRange(startCell, endCell);
        }
    });

    document.addEventListener('mouseup', () => {
        isSelecting = false;
    });

    // Rest of the event listeners...
    document.addEventListener('keydown', (e) => {
        // Copy functionality
        if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
            e.preventDefault();
            copySelectionToClipboard();
        }

        // Select all
        if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
            e.preventDefault();
            const allCells = document.querySelectorAll('td[data-row][data-col]');
            if (allCells.length > 0) {
                clearSelection();
                allCells.forEach(cell => {
                    cell.classList.add('selected');
                    selectedCells.add(cell);
                });
                
                const firstCell = allCells[0];
                if (firstCell) {
                    firstCell.classList.add('active-cell');
                    activeCell = firstCell;
                    startCell = getCellCoordinates(firstCell);
                }
                
                updateSelectionInfo();
            }
        }

        // Arrow key navigation (keep existing code)
        if (!activeCell) return;
        const coords = getCellCoordinates(activeCell);
        if (!coords) return;

        let newCoords = { ...coords };
        let moved = false;

        switch(e.key) {
            case 'ArrowUp':
                if (coords.row > 0) {
                    newCoords.row--;
                    moved = true;
                }
                break;
            case 'ArrowDown':
                newCoords.row++;
                moved = true;
                break;
            case 'ArrowLeft':
                if (coords.col > 0) {
                    newCoords.col--;
                    moved = true;
                }
                break;
            case 'ArrowRight':
                newCoords.col++;
                moved = true;
                break;
        }

        if (moved) {
            e.preventDefault();
            const newCell = document.querySelector('td[data-row="' + newCoords.row + '"][data-col="' + newCoords.col + '"]');
            if (newCell) {
                if (e.shiftKey) {
                    selectCellsInRange(startCell || coords, newCoords);
                } else {
                    clearSelection();
                    newCell.classList.add('selected');
                    newCell.classList.add('active-cell');
                    selectedCells.add(newCell);
                    activeCell = newCell;
                    startCell = newCoords;
                    endCell = newCoords;
                }
                updateSelectionInfo();
            }
        }
    });

    document.addEventListener('click', (e) => {
        if (!e.target.closest('#csv-table')) {
            clearSelection();
        }
    });

    tableElement.addEventListener('dblclick', (e) => {
        const target = e.target.closest('td');
        if (target) {
            const value = target.textContent.trim();
            console.log('Cell value:', value);
        }
    });
}

// Updated selectCellsInRange to not clear Ctrl+clicked cells
function selectCellsInRange(start, end) {
    if (!start || !end) return;
    
    // Store ctrl+clicked cells
    const ctrlSelectedCells = new Set();
    selectedCells.forEach(cell => {
        const coords = getCellCoordinates(cell);
        if (coords) {
            const inRange = coords.row >= Math.min(start.row, end.row) && 
                           coords.row <= Math.max(start.row, end.row) &&
                           coords.col >= Math.min(start.col, end.col) && 
                           coords.col <= Math.max(start.col, end.col);
            if (!inRange) {
                ctrlSelectedCells.add(cell);
            }
        }
    });
    
    // Clear only cells in the range
    document.querySelectorAll('td.selected').forEach(el => {
        if (!ctrlSelectedCells.has(el)) {
            el.classList.remove('selected');
            selectedCells.delete(el);
        }
    });
    
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    const cells = document.querySelectorAll('td[data-row][data-col]');

    cells.forEach(cell => {
        const coords = getCellCoordinates(cell);
        if (coords && coords.row >= minRow && coords.row <= maxRow && 
            coords.col >= minCol && coords.col <= maxCol) {
            cell.classList.add('selected');
            selectedCells.add(cell);
        }
    });

    // Ensure active cell
    const startCellElement = document.querySelector('td[data-row="' + start.row + '"][data-col="' + start.col + '"]');
    if (startCellElement) {
        document.querySelectorAll('td.active-cell').forEach(el => el.classList.remove('active-cell'));
        startCellElement.classList.add('active-cell');
        activeCell = startCellElement;
    }

    updateSelectionInfo();
}

// Updated column selection
function selectColumn(columnIndex, ctrlKey, shiftKey) {
    if (!ctrlKey && !shiftKey) {
        clearSelection();
        lastSelectedColumn = columnIndex;
    }

    if (shiftKey && lastSelectedColumn !== null && lastSelectedColumn !== columnIndex) {
        // Clear if not Ctrl+Shift
        if (!ctrlKey) {
            clearSelection();
        }
        
        const minCol = Math.min(lastSelectedColumn, columnIndex);
        const maxCol = Math.max(lastSelectedColumn, columnIndex);
        
        for (let col = minCol; col <= maxCol; col++) {
            if (!selectedColumns.has(col)) {
                selectedColumns.add(col);
                const cells = document.querySelectorAll('td[data-col="' + col + '"], th[data-col="' + col + '"]');
                cells.forEach(cell => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
        }
    } else if (ctrlKey) {
        // Toggle selection
        if (selectedColumns.has(columnIndex)) {
            selectedColumns.delete(columnIndex);
            const cells = document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]');
            cells.forEach(cell => {
                cell.classList.remove('column-selected');
                if (cell.tagName === 'TD') selectedCells.delete(cell);
            });
        } else {
            selectedColumns.add(columnIndex);
            const cells = document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]');
            cells.forEach(cell => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell);
            });
        }
    } else {
        // Single selection
        selectedColumns.add(columnIndex);
        const cells = document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]');
        cells.forEach(cell => {
            cell.classList.add('column-selected');
            if (cell.tagName === 'TD') selectedCells.add(cell);
        });
    }
    
    updateSelectionInfo();
}

// Updated row selection
function selectRow(rowIndex, ctrlKey, shiftKey) {
    if (!ctrlKey && !shiftKey) {
        clearSelection();
        lastSelectedRow = rowIndex;
    }

    if (shiftKey && lastSelectedRow !== null && lastSelectedRow !== rowIndex) {
        // Clear if not Ctrl+Shift
        if (!ctrlKey) {
            clearSelection();
        }
        
        const minRow = Math.min(lastSelectedRow, rowIndex);
        const maxRow = Math.max(lastSelectedRow, rowIndex);
        
        for (let row = minRow; row <= maxRow; row++) {
            if (!selectedRows.has(row)) {
                selectedRows.add(row);
                const rowHeader = document.querySelector('th[data-row="' + row + '"]');
                if (rowHeader && rowHeader.parentElement) {
                    const cells = rowHeader.parentElement.querySelectorAll('td, th');
                    cells.forEach(cell => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') selectedCells.add(cell);
                    });
                }
            }
        }
    } else if (ctrlKey) {
        // Toggle selection
        if (selectedRows.has(rowIndex)) {
            selectedRows.delete(rowIndex);
            const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
            if (rowHeader && rowHeader.parentElement) {
                const cells = rowHeader.parentElement.querySelectorAll('td, th');
                cells.forEach(cell => {
                    cell.classList.remove('row-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell);
                });
            }
        } else {
            selectedRows.add(rowIndex);
            const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
            if (rowHeader && rowHeader.parentElement) {
                const cells = rowHeader.parentElement.querySelectorAll('td, th');
                cells.forEach(cell => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
        }
    } else {
        // Single selection
        selectedRows.add(rowIndex);
        const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
        if (rowHeader && rowHeader.parentElement) {
            const cells = rowHeader.parentElement.querySelectorAll('td, th');
            cells.forEach(cell => {
                cell.classList.add('row-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell);
            });
        }
    }
    
    updateSelectionInfo();
}
    
                document.addEventListener('DOMContentLoaded', initializeSelection);
            </script>
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