import * as vscode from 'vscode';
import * as Excel from 'exceljs';
import { convertARGBToRGBA, isShadeOfBlack, isShadeOfWhite } from './xlsx/xlsxUtilities';

export class XLSXEditorProvider implements vscode.CustomReadonlyEditorProvider {
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
        const webview = webviewPanel.webview;

        webview.options = {
            enableScripts: true,
            localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'resources')]
        };

        // Set the webview shell immediately; it will show its own loading overlay.
        webview.html = this.getWebviewContent(webviewPanel);

        let isWebviewReady = false;
        let worksheetsPayload: any[] | null = null;
        let rowHeaderWidth = 60;

        const getPersistedSettings = () => {
            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
            return {
                firstRowIsHeader: cfg.get('xlsx.firstRowIsHeader', false),
                stickyToolbar: cfg.get('xlsx.stickyToolbar', true),
                stickyHeader: cfg.get('xlsx.stickyHeader', false),
                hyperlinkPreview: cfg.get('xlsx.hyperlinkPreview', true)
            };
        };

        const trySendSettings = () => {
            if (!isWebviewReady) return;
            try {
                webview.postMessage({
                    command: 'initSettings',
                    settings: getPersistedSettings()
                });
            } catch {
                // ignore
            }
        };

        const trySendInit = () => {
            if (!isWebviewReady || !worksheetsPayload) return;
            try {
                webview.postMessage({
                    command: 'init',
                    worksheets: worksheetsPayload,
                    rowHeaderWidth
                });
            } catch {
                // ignore
            }
        };

        const loadWorkbookPayload = async () => {
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile(document.uri.fsPath);

            const worksheets = workbook.worksheets.map((worksheet, index) => {
                const data = this.extractWorksheetData(worksheet);
                return {
                    name: worksheet.name,
                    index,
                    data
                };
            });

            const maxRows = worksheets.length ? Math.max(...worksheets.map(ws => ws.data.maxRow)) : 0;
            rowHeaderWidth = Math.max(60, Math.ceil(Math.log10(maxRows + 1)) * 12 + 20);
            worksheetsPayload = worksheets;
        };

        // Listen for messages
        webview.onDidReceiveMessage(async message => {
            if (message?.command === 'webviewReady') {
                isWebviewReady = true;
                trySendSettings();
                trySendInit();
                return;
            }

            if (message?.command === 'updateSettings') {
                try {
                    const s = message.settings || {};
                    const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                    await cfg.update('xlsx.firstRowIsHeader', !!s.firstRowIsHeader, vscode.ConfigurationTarget.Global);
                    await cfg.update('xlsx.stickyToolbar', !!s.stickyToolbar, vscode.ConfigurationTarget.Global);
                    await cfg.update('xlsx.stickyHeader', !!s.stickyHeader, vscode.ConfigurationTarget.Global);
                    await cfg.update('xlsx.hyperlinkPreview', !!s.hyperlinkPreview, vscode.ConfigurationTarget.Global);
                } catch (err) {
                    console.error('Failed to persist XLSX settings:', err);
                }
                return;
            }

            if (message?.command === 'toggleView') {
                if (!message.isTableView) {
                    await vscode.commands.executeCommand('vscode.openWith', document.uri, 'default');
                    webviewPanel.dispose();
                }
                return;
            }

            if (message?.command === 'openExternal') {
                try {
                    const url = typeof message.url === 'string' ? message.url : '';
                    if (url) {
                        await vscode.env.openExternal(vscode.Uri.parse(url));
                    }
                } catch {
                    // ignore
                }
                return;
            }

            if (message?.command === 'saveXlsxEdits') {
                try {
                    const edits = Array.isArray(message.edits) ? message.edits : [];
                    const sheetIndex = typeof message.sheetIndex === 'number' ? message.sheetIndex : 0;

                    if (!edits.length) {
                        try { webview.postMessage({ command: 'saveResult', ok: true }); } catch { }
                        return;
                    }

                    const workbook = new Excel.Workbook();
                    await workbook.xlsx.readFile(document.uri.fsPath);
                    const ws = workbook.worksheets[sheetIndex];
                    if (!ws) {
                        throw new Error('Worksheet not found');
                    }

                    for (const edit of edits) {
                        const row = typeof edit.row === 'number' ? edit.row : undefined;
                        const col = typeof edit.col === 'number' ? edit.col : undefined;
                        if (!row || !col) continue;

                        const newText = typeof edit.value === 'string' ? edit.value : '';
                        const cell = ws.getRow(row).getCell(col);

                        if (cell.type === Excel.ValueType.Hyperlink) {
                            const hyperlinkValue = cell.value as Excel.CellHyperlinkValue;
                            cell.value = {
                                text: newText,
                                hyperlink: hyperlinkValue.hyperlink
                            } as Excel.CellHyperlinkValue;
                        } else {
                            cell.value = newText;
                        }
                    }

                    await workbook.xlsx.writeFile(document.uri.fsPath);

                    // Refresh the rendered payload after saving
                    await loadWorkbookPayload();
                    trySendInit();
                    try { webview.postMessage({ command: 'saveResult', ok: true }); } catch { }
                } catch (err) {
                    try { webview.postMessage({ command: 'saveResult', ok: false, error: String(err) }); } catch { }
                }
            }
        });

        // Forward settings changes made outside the webview
        const configChangeDisposable = vscode.workspace.onDidChangeConfiguration(e => {
            if (e.affectsConfiguration('xlsxViewer.xlsx') || e.affectsConfiguration('xlsxViewer')) {
                try {
                    webview.postMessage({ command: 'settingsUpdated', settings: getPersistedSettings() });
                } catch {
                    // ignore
                }
            }
        });
        webviewPanel.onDidDispose(() => configChangeDisposable.dispose());

        try {
            await loadWorkbookPayload();
            trySendSettings();
            trySendInit();
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading XLSX file: ${error}`);
        }
    }

    private getLoadingContent(): string {
        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Loading XLSX File</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #ffffff;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            flex-direction: column;
        }
        
        .loading-container {
            text-align: center;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #2196f3;
            border-radius: 50%;
            width: 60px;
            height: 60px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .loading-text {
            font-size: 18px;
            color: #333;
            margin-bottom: 10px;
        }
        
        .loading-subtext {
            font-size: 14px;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="loading-container">
        <div class="spinner"></div>
        <div class="loading-text">Loading XLSX File...</div>
        <div class="loading-subtext">Please wait while we process your spreadsheet</div>
    </div>
</body>
</html>`;
    }

    private extractWorksheetData(worksheet: Excel.Worksheet): any {
        const data: any = {
            rows: [],
            maxRow: 0,
            maxCol: 0,
            mergedCells: []
        };

        // Extract merged cell ranges
        try {
            // Check different ways ExcelJS might store merged cells
            let merges: any[] = [];

            // Method 1: Check worksheet.model.merges
            if ((worksheet as any).model && (worksheet as any).model.merges) {
                merges = (worksheet as any).model.merges;
            }

            // Method 2: Check worksheet._merges (fallback)
            if (merges.length === 0 && (worksheet as any)._merges) {
                merges = (worksheet as any)._merges;
            }

            // Method 3: Check worksheet.merged (another possible location)
            if (merges.length === 0 && (worksheet as any).merged) {
                merges = (worksheet as any).merged;
            }

            merges.forEach((merge: any) => {
                // Handle different merge formats
                let startRow, startCol, endRow, endCol;

                if (merge.top !== undefined) {
                    // Format: {top, left, bottom, right}
                    startRow = merge.top;
                    startCol = merge.left;
                    endRow = merge.bottom;
                    endCol = merge.right;
                } else if (merge.start && merge.end) {
                    // Format: {start: {row, col}, end: {row, col}}
                    startRow = merge.start.row;
                    startCol = merge.start.col;
                    endRow = merge.end.row;
                    endCol = merge.end.col;
                } else if (typeof merge === 'string') {
                    // Format: "A1:B2" - parse range string
                    const range = this.parseRange(merge);
                    if (range) {
                        startRow = range.startRow;
                        startCol = range.startCol;
                        endRow = range.endRow;
                        endCol = range.endCol;
                    }
                }

                if (startRow && startCol && endRow && endCol) {
                    data.mergedCells.push({
                        startRow,
                        startCol,
                        endRow,
                        endCol
                    });
                }
            });
        } catch {
            // Silently continue without merged cells if there's an error
        }

        // Find actual data bounds
        let maxRow = 0;
        let maxCol = 0;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            maxRow = Math.max(maxRow, rowNumber);
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                maxCol = Math.max(maxCol, colNumber);
            });
        });

        // Include at least some empty rows/cols for better display
        maxRow = Math.max(maxRow, 20);
        maxCol = Math.max(maxCol, 10);

        data.maxRow = maxRow;
        data.maxCol = maxCol;

        // Create a 2D grid to track merged cells
        const cellGrid: any[][] = [];
        for (let r = 0; r <= maxRow; r++) {
            cellGrid[r] = [];
        }

        // Mark merged cells in the grid
        data.mergedCells.forEach((range: any) => {
            for (let r = range.startRow; r <= range.endRow; r++) {
                for (let c = range.startCol; c <= range.endCol; c++) {
                    cellGrid[r][c] = {
                        isMerged: true,
                        isMaster: r === range.startRow && c === range.startCol,
                        rowspan: range.endRow - range.startRow + 1,
                        colspan: range.endCol - range.startCol + 1,
                        masterRow: range.startRow,
                        masterCol: range.startCol
                    };
                }
            }
        });

        // Extract all cell data
        for (let r = 1; r <= maxRow; r++) {
            const row = worksheet.getRow(r);
            const rowData: any = {
                rowNumber: r,
                cells: [],
                height: row.height || 15 // Default row height
            };

            for (let c = 1; c <= maxCol; c++) {
                const mergeInfo = cellGrid[r] && cellGrid[r][c];

                // Skip cells that are part of a merged range but not the master
                if (mergeInfo && mergeInfo.isMerged && !mergeInfo.isMaster) {
                    continue; // Don't add to cells array
                }

                const cell = worksheet.getRow(r).getCell(c);
                const cellStyle = this.getCellStyle(cell);
                let cellValue = this.getCellValue(cell);
                const hyperlinkUrl = this.getCellHyperlink(cell);

                // For merged master cells, ensure we get the value
                if (mergeInfo && mergeInfo.isMaster && !cellValue) {
                    // Try to get value from any cell in the merged range
                    for (let mr = mergeInfo.masterRow; mr <= mergeInfo.masterRow + mergeInfo.rowspan - 1; mr++) {
                        for (let mc = mergeInfo.masterCol; mc <= mergeInfo.masterCol + mergeInfo.colspan - 1; mc++) {
                            const testCell = worksheet.getRow(mr).getCell(mc);
                            const testValue = this.getCellValue(testCell);
                            if (testValue) {
                                cellValue = testValue;
                                break;
                            }
                        }
                        if (cellValue) break;
                    }
                }

                const cellData: any = {
                    value: cellValue,
                    hyperlink: hyperlinkUrl,
                    style: cellStyle,
                    colNumber: c,
                    rowNumber: r,
                    // Add data attributes for proper color handling
                    isDefaultColor: cellStyle._isDefaultColor || false,
                    // True when cell had no explicit background defined in the file
                    hasDefaultBg: !cellStyle.backgroundColor,
                    // True when the cell had an explicit white (or near-white) background
                    hasWhiteBackground: cellStyle._hasWhiteBackground || false,
                    hasBlackBorder: cellStyle._hasBlackBorder || false,
                    hasBlackBackground: cellStyle._hasBlackBackground || false,
                    originalColor: cellStyle.color || 'rgb(0, 0, 0)',
                    isEmpty: !cell || (cell.value === null && !cellStyle.backgroundColor),
                    // Merged cell info
                    rowspan: mergeInfo ? mergeInfo.rowspan : 1,
                    colspan: mergeInfo ? mergeInfo.colspan : 1,
                    isMerged: !!(mergeInfo && mergeInfo.isMerged)
                };

                rowData.cells.push(cellData);
            }

            data.rows.push(rowData);
        }

        // Column widths
        data.columnWidths = [];
        for (let c = 1; c <= maxCol; c++) {
            const col = worksheet.getColumn(c);
            data.columnWidths.push(col.width || 10);
        }

        return data;
    }

    private parseRange(rangeStr: string): any {
        // Parse range like "A1:B2" to coordinates
        try {
            const [start, end] = rangeStr.split(':');
            const startCoord = this.parseCell(start);
            const endCoord = this.parseCell(end);

            if (startCoord && endCoord) {
                return {
                    startRow: startCoord.row,
                    startCol: startCoord.col,
                    endRow: endCoord.row,
                    endCol: endCoord.col
                };
            }
        } catch {
            // Silently continue if parsing fails
        }
        return null;
    }

    private parseCell(cellStr: string): any {
        // Parse cell like "A1" to {row: 1, col: 1}
        try {
            const match = cellStr.match(/^([A-Z]+)(\d+)$/);
            if (match) {
                const colStr = match[1];
                const rowStr = match[2];

                let col = 0;
                for (let i = 0; i < colStr.length; i++) {
                    col = col * 26 + (colStr.charCodeAt(i) - 64);
                }

                return {
                    row: parseInt(rowStr, 10),
                    col
                };
            }
        } catch {
            // Silently continue if parsing fails
        }
        return null;
    }

    private getCellValue(cell: Excel.Cell): string {
        if (!cell || !cell.value) return '';

        // Some ExcelJS hyperlink cells expose the URL via cell.hyperlink even when cell.type isn't Hyperlink.
        // In those cases, keep showing the displayed text/value.
        const anyCell = cell as any;
        if (typeof anyCell.hyperlink === 'string' && anyCell.hyperlink) {
            const v = cell.value as any;
            if (typeof v === 'string') return v;
            if (v && typeof v === 'object' && typeof v.text === 'string') return v.text;
        }

        // Handle different value types with proper type checking
        if (cell.type === Excel.ValueType.Hyperlink) {
            const hyperlinkValue = cell.value as Excel.CellHyperlinkValue;
            return hyperlinkValue.text || '';
        } else if (cell.type === Excel.ValueType.Formula) {
            return cell.result?.toString() || '';
        } else if (cell.type === Excel.ValueType.RichText) {
            const richTextValue = cell.value as Excel.CellRichTextValue;
            return richTextValue.richText.map((rt: any) => rt.text).join('');
        } else if (cell.type === Excel.ValueType.Date) {
            const dateValue = cell.value as Date;
            return dateValue.toLocaleDateString();
        } else if (cell.value instanceof Date) {
            // Additional check for Date objects
            return cell.value.toLocaleDateString();
        } else {
            return cell.value.toString();
        }
    }

    private getCellHyperlink(cell: Excel.Cell): string {
        try {
            if (!cell) return '';

            const anyCell = cell as any;
            if (typeof anyCell.hyperlink === 'string' && anyCell.hyperlink) {
                return anyCell.hyperlink;
            }

            if (cell.type === Excel.ValueType.Hyperlink) {
                const hyperlinkValue = cell.value as Excel.CellHyperlinkValue;
                return hyperlinkValue.hyperlink || '';
            }

            const v = cell.value as any;
            if (v && typeof v === 'object' && typeof v.hyperlink === 'string') {
                return v.hyperlink;
            }
        } catch {
            // ignore
        }
        return '';
    }

    private getCellStyle(cell: Excel.Cell): any {
        const style: any = {};
        let isDefaultColor = false;
        let hasBlackBorder = false;
        let hasBlackBackground = false;
        let hasWhiteBackground = false;

        // Background color
        if (cell.fill && cell.fill.type === 'pattern' && (cell.fill as any).fgColor) {
            const color = (cell.fill as any).fgColor;
            if (color.argb) {
                const bgColor = convertARGBToRGBA(color.argb);
                style.backgroundColor = bgColor;
                // Check if background is black or shade of black - be very strict
                hasBlackBackground = isShadeOfBlack(bgColor);
                // Check if background is white or shade of white
                hasWhiteBackground = isShadeOfWhite(bgColor);
            }
        }

        // Font
        if (cell.font) {
            if (cell.font.color && cell.font.color.argb) {
                const fontColor = convertARGBToRGBA(cell.font.color.argb);
                style.color = fontColor;
                // If it's a shade of black, we can treat it as default color for theme switching
                if (isShadeOfBlack(fontColor)) {
                    isDefaultColor = true;
                }
            } else {
                // No custom font color set, defaults to black
                style.color = 'rgb(0, 0, 0)';
                isDefaultColor = true;
            }
            if (cell.font.bold) style.fontWeight = 'bold';
            if (cell.font.italic) style.fontStyle = 'italic';
            if (cell.font.underline) style.textDecoration = 'underline';
            if (cell.font.strike) style.textDecoration = (style.textDecoration || '') + ' line-through';
            if (cell.font.size) style.fontSize = `${cell.font.size}pt`;
            if (cell.font.name) style.fontFamily = cell.font.name;
        } else {
            // No font styling at all, defaults to black
            style.color = 'rgb(0, 0, 0)';
            isDefaultColor = true;
        }

        // Alignment
        if (cell.alignment) {
            if (cell.alignment.horizontal) {
                switch (cell.alignment.horizontal) {
                    case 'left':
                        style.textAlign = 'left';
                        break;
                    case 'center':
                        style.textAlign = 'center';
                        break;
                    case 'right':
                        style.textAlign = 'right';
                        break;
                    case 'justify':
                        style.textAlign = 'justify';
                        break;
                    default:
                        style.textAlign = cell.alignment.horizontal;
                }
            }
            if (cell.alignment.vertical) {
                switch (cell.alignment.vertical) {
                    case 'top':
                        style.verticalAlign = 'top';
                        break;
                    case 'middle':
                        style.verticalAlign = 'middle';
                        break;
                    case 'bottom':
                        style.verticalAlign = 'bottom';
                        break;
                    default:
                        style.verticalAlign = cell.alignment.vertical;
                }
            }
            if (cell.alignment.wrapText) {
                style.whiteSpace = 'pre-wrap';
                style.wordWrap = 'break-word';
            }
            if (cell.alignment.indent) {
                style.paddingLeft = `${cell.alignment.indent * 8}px`;
            }
        }

        // Borders
        if (cell.border) {
            style.border = {};
            ['top', 'right', 'bottom', 'left'].forEach(side => {
                const border = (cell.border as any)[side];
                if (border && border.style) {
                    const originalColor = border.color && border.color.argb
                        ? convertARGBToRGBA(border.color.argb)
                        : 'rgba(0, 0, 0, 1)';

                    // Only mark as black border if it's actually black or a shade of black
                    const isBlack = isShadeOfBlack(originalColor);
                    if (isBlack) {
                        hasBlackBorder = true;
                    }

                    let width = '1px';
                    let styleStr = 'solid';

                    switch (border.style) {
                        case 'thin': width = '1px'; break;
                        case 'medium': width = '2px'; break;
                        case 'thick': width = '3px'; break;
                        case 'dotted': styleStr = 'dotted'; break;
                        case 'dashed': styleStr = 'dashed'; break;
                        case 'double': styleStr = 'double'; width = '3px'; break;
                    }

                    style.border[side] = `${width} ${styleStr} ${originalColor}`;
                }
            });
        }

        // Add tracking properties for dark mode handling
        style._isDefaultColor = isDefaultColor;
        style._hasBlackBorder = hasBlackBorder;
        style._hasBlackBackground = hasBlackBackground;
        style._hasWhiteBackground = hasWhiteBackground;

        return style;
    }

    private getExcelColumnLabel(n: number): string {
        let label = '';
        while (n > 0) {
            let rem = (n - 1) % 26;
            label = String.fromCharCode(65 + rem) + label;
            n = Math.floor((n - 1) / 26);
        }
        return label;
    }

    private getWebviewContent(webviewPanel: vscode.WebviewPanel): string {
        const webview = webviewPanel.webview;

        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'xlsx', 'xlsxWebview.js')
        );
        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'xlsx', 'xlsxWebview.css')
        );
        const cspSource = webview.cspSource;

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource};">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX Viewer</title>
    <link href="${styleUri}" rel="stylesheet" />
</head>
<body>
    <div class="header-background"></div>

    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="loading-text">Rendering worksheet...</div>
    </div>

    <div class="toolbar">
        <select id="sheetSelector" class="sheet-selector" title="Select sheet"></select>

        <button id="toggleTableEditButton" class="toggle-button" title="Edit XLSX directly in the table (text only)">
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

        <button id="openSettingsButton" class="toggle-button icon-only" title="XLSX Settings">
           <svg fill="#ffffff" version="1.1" id="Capa_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 389.663 389.663" xml:space="preserve"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <g> <path d="M194.832,132.997c-34.1,0-61.842,27.74-61.842,61.838c0,34.1,27.742,61.841,61.842,61.841 c34.099,0,61.841-27.741,61.841-61.841C256.674,160.737,228.932,132.997,194.832,132.997z M194.832,226.444 c-17.429,0-31.608-14.182-31.608-31.61c0-17.428,14.18-31.605,31.608-31.605c17.429,0,31.607,14.178,31.607,31.605 C226.439,212.264,212.262,226.444,194.832,226.444z"></path> <path d="M385.23,150.784c-2.816-2.812-6.714-4.427-10.688-4.427l-49.715,0.015l-3.799-9.194l35.149-35.155 c5.892-5.894,5.892-15.483,0-21.377l-47.166-47.162c-2.688-2.691-6.586-4.235-10.688-4.235c-4.103,0-7.996,1.544-10.687,4.235 L252.48,68.639l-9.188-3.797V15.116C243.292,6.781,236.511,0,228.177,0h-66.694c-8.335,0-15.116,6.78-15.116,15.115v49.716 l-9.194,3.801l-35.151-35.135c-2.855-2.854-6.65-4.426-10.686-4.426c-4.036,0-7.832,1.572-10.688,4.427L33.476,80.67 c-2.813,2.814-4.427,6.711-4.427,10.688c0,3.984,1.613,7.882,4.427,10.693l35.151,35.127l-3.811,9.188l-49.697,0.005 C6.781,146.372,0,153.153,0,161.488v66.708c0,4.035,1.573,7.832,4.431,10.689c2.817,2.815,6.713,4.432,10.688,4.432l49.708-0.021 l3.799,9.195l-35.133,35.149c-5.894,5.896-5.894,15.484,0,21.378l47.161,47.172c2.692,2.69,6.591,4.233,10.693,4.233 c4.105,0,8.002-1.543,10.69-4.233l35.136-35.162l9.186,3.815l0.008,49.691c0,8.338,6.781,15.121,15.116,15.121l66.708,0.006h0.162 c8.336,0,15.116-6.781,15.116-15.117c0-0.721-0.049-1.444-0.147-2.151l-0.015-0.207l-0.013-47.355l9.195-3.801l35.149,35.139 c2.855,2.857,6.65,4.432,10.688,4.432c4.035,0,7.83-1.573,10.686-4.432l47.172-47.166c2.855-2.854,4.429-6.649,4.429-10.688 c0-4.045-1.572-7.847-4.429-10.699l-35.157-35.125l3.809-9.195h49.707c8.336,0,15.119-6.78,15.119-15.114v-66.708 C389.662,157.438,388.088,153.641,385.23,150.784z M359.428,213.063h-44.696c-6.134,0-11.615,3.662-13.966,9.328l-11.534,27.865 c-2.351,5.672-1.062,12.141,3.274,16.482l31.609,31.58l-25.789,25.789l-31.605-31.603c-2.854-2.853-6.649-4.422-10.69-4.422 c-1.992,0-3.938,0.388-5.785,1.147l-27.854,11.537c-5.666,2.349-9.327,7.832-9.327,13.972l0.008,44.688l-36.468-0.01 l-0.008-44.686c0-6.136-3.661-11.615-9.328-13.966l-27.856-11.536c-1.854-0.768-3.806-1.155-5.802-1.155 c-4.036,0-7.829,1.571-10.677,4.43l-31.586,31.615L65.559,298.33l31.592-31.604c4.339-4.343,5.625-10.81,3.275-16.478 L88.89,222.393c-2.352-5.666-7.833-9.328-13.965-9.328l-44.688,0.01v-36.466l44.688-0.01c6.134,0,11.615-3.662,13.965-9.328 l11.536-27.854c2.349-5.676,1.063-12.146-3.275-16.482L65.548,91.359l25.79-25.796l31.599,31.582 c2.856,2.857,6.658,4.43,10.704,4.43c1.988,0,3.928-0.385,5.764-1.144l27.861-11.524c5.671-2.351,9.336-7.834,9.336-13.97V30.231 h36.459v44.705c0,6.137,3.662,11.618,9.328,13.965l27.855,11.534c1.848,0.766,3.795,1.153,5.789,1.153 c4.039,0,7.832-1.572,10.684-4.429l31.607-31.617l25.789,25.789l-31.609,31.607c-4.336,4.339-5.621,10.806-3.274,16.478 l11.534,27.858c2.351,5.669,7.832,9.332,13.966,9.332l44.696-0.01L359.428,213.063L359.428,213.063z"></path> </g> </g> </g></svg>
        </button>

        <div id="settingsPanel" class="settings-panel hidden" role="dialog" aria-hidden="true">
            <div class="settings-group">
                <label class="setting-item"><input type="checkbox" id="chkHeaderRow"/> <span>Header Row</span></label>
                <label class="setting-item"><input type="checkbox" id="chkStickyHeader"/> <span>Sticky Header</span></label>
                <label class="setting-item"><input type="checkbox" id="chkStickyToolbar"/> <span>Sticky Toolbar</span></label>
                <label class="setting-item"><input type="checkbox" id="chkHyperlinkPreview"/> <span>Hyperlink Preview</span></label>
            </div>

            <button id="settingsCancelButton" class="toggle-button" title="Close">Close</button>
        </div>

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
    </div>

    <div id="content">
        <div id="tableContainer"></div>
    </div>
    <div class="selection-info" id="selectionInfo"></div>
    <div class="resize-indicator" id="resizeIndicator"></div>

    <noscript>
        <div style="padding: 8px; margin-top: 10px; background: #fff3cd; border: 1px solid #ffeeba;">
            JavaScript is disabled in this webview, so the XLSX table cannot load.
        </div>
    </noscript>

    <script src="${scriptUri}"></script>
</body>
</html>`;
    }
}
