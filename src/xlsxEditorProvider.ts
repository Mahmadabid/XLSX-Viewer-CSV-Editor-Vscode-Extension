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
            localResourceRoots: [
                vscode.Uri.joinPath(this.context.extensionUri, 'resources'),
                vscode.Uri.joinPath(this.context.extensionUri, 'dist')
            ]
        };

        // Set the webview shell immediately; it will show its own loading overlay.
        webview.html = this.getWebviewContent(webviewPanel);

        let isWebviewReady = false;
        // Store parsed worksheet data for virtualization
        let worksheetsData: any[] = [];
        let rowHeaderWidth = 60;

        const getPersistedSettings = () => {
            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
            const globalCfg = vscode.workspace.getConfiguration('workbench');
            const associations: any = globalCfg.get('editorAssociations');
            let isDefault = false;
            
            if (associations) {
                if (Array.isArray(associations)) {
                    isDefault = associations.some(a => 
                        a.viewType === 'xlsxViewer.xlsx' && 
                        (a.filenamePattern === '*.xlsx' || a.filenamePattern === '**/*.xlsx')
                    );
                } else {
                    isDefault = associations["*.xlsx"] === 'xlsxViewer.xlsx' || associations["**/*.xlsx"] === 'xlsxViewer.xlsx';
                }
            }

            return {
                firstRowIsHeader: cfg.get('xlsx.firstRowIsHeader', false),
                stickyToolbar: cfg.get('xlsx.stickyToolbar', true),
                stickyHeader: cfg.get('xlsx.stickyHeader', false),
                hyperlinkPreview: cfg.get('xlsx.hyperlinkPreview', true),
                isDefaultEditor: isDefault
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

        // Sync VS Code theme changes to the webview
        const themeChangeDisposable = vscode.window.onDidChangeActiveColorTheme(() => {
            try { webview.postMessage({ type: 'setTheme', kind: vscode.window.activeColorTheme.kind }); } catch { }
        });

        webviewPanel.onDidDispose(() => themeChangeDisposable.dispose());

        const trySendInit = () => {
            if (!isWebviewReady || !worksheetsData.length) return;
            try {
                // Send metadata for virtual scrolling instead of full data
                // Include row heights for stable scroll calculations
                const worksheetsMeta = worksheetsData.map((ws, index) => ({
                    name: ws.name,
                    index,
                    totalRows: ws.data.maxRow,
                    columnCount: ws.data.maxCol,
                    columnWidths: ws.data.columnWidths,
                    mergedCells: ws.data.mergedCells,
                    rowHeights: ws.data.rows.map((row: any) => row.height || 28)
                }));
                webview.postMessage({
                    command: 'initVirtualTable',
                    worksheets: worksheetsMeta,
                    rowHeaderWidth
                });
            } catch {
                // ignore
            }
        };

        const loadWorkbookPayload = async () => {
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile(document.uri.fsPath);

            worksheetsData = workbook.worksheets.map((worksheet, index) => {
                const data = this.extractWorksheetData(worksheet);
                return {
                    name: worksheet.name,
                    index,
                    data
                };
            });

            const maxRows = worksheetsData.length ? Math.max(...worksheetsData.map(ws => ws.data.maxRow)) : 0;
            rowHeaderWidth = Math.max(60, Math.ceil(Math.log10(maxRows + 1)) * 12 + 20);
        };

        // Listen for messages
        webview.onDidReceiveMessage(async message => {
            if (message?.command === 'webviewReady') {
                isWebviewReady = true;
                trySendSettings();
                trySendInit();
                // Send current theme info to webview
                try {
                    webview.postMessage({ type: 'setTheme', kind: vscode.window.activeColorTheme.kind });
                } catch { }
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

            if (message?.command === 'enableDefaultEditor' || message?.command === 'enableAsDefault') {
                await vscode.commands.executeCommand('xlsx-viewer.toggleAssociation', { type: 'xlsx', enable: true });
                trySendSettings();
                return;
            }

            if (message?.command === 'disableDefaultEditor') {
                try {
                    const result = await vscode.window.showWarningMessage(
                        "Are you sure you want to disable XLSX Viewer for all .xlsx files? You will be prompted to select a new default editor.",
                        "Yes, Disable",
                        "Cancel"
                    );

                    if (result === "Yes, Disable") {
                        await vscode.commands.executeCommand('xlsx-viewer.toggleAssociation', { type: 'xlsx', enable: false });
                        await vscode.commands.executeCommand('workbench.action.reopenWithEditor');
                    }
                } catch (err) {
                    vscode.window.showErrorMessage(`Error disabling editor: ${err}`);
                }
                return;
            }

            // Handle getRows request for virtual scrolling
            if (message?.command === 'getRows') {
                const { start, end, requestId, sheetIndex } = message;
                const wsIndex = typeof sheetIndex === 'number' ? sheetIndex : 0;
                const ws = worksheetsData[wsIndex];
                if (!ws) {
                    webview.postMessage({
                        command: 'rowsData',
                        rows: [],
                        start,
                        end,
                        requestId
                    });
                    return;
                }

                const clampedStart = Math.max(0, start);
                const clampedEnd = Math.min(ws.data.rows.length, end);
                const rows = ws.data.rows.slice(clampedStart, clampedEnd);

                webview.postMessage({
                    command: 'rowsData',
                    rows,
                    start: clampedStart,
                    end: clampedEnd,
                    requestId
                });
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
            if (e.affectsConfiguration('xlsxViewer.xlsx') || e.affectsConfiguration('xlsxViewer') || e.affectsConfiguration('workbench.editorAssociations')) {
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
                    hasWhiteBorder: cellStyle._hasWhiteBorder || false,
                    hasBlackBackground: cellStyle._hasBlackBackground || false,
                    // True when cell has no explicit border (should use theme default)
                    hasDefaultBorder: cellStyle._hasDefaultBorder || false,
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
        let hasWhiteBorder = false;
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
                    
                    // Check for white borders
                    const isWhite = isShadeOfWhite(originalColor);
                    if (isWhite) {
                        hasWhiteBorder = true;
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

        // Track if cell has no explicit border (should use theme default)
        const hasExplicitBorder = cell.border && (
            cell.border.top || cell.border.right || cell.border.bottom || cell.border.left
        );

        // Add tracking properties for dark mode handling
        style._isDefaultColor = isDefaultColor;
        style._hasBlackBorder = hasBlackBorder;
        style._hasWhiteBorder = hasWhiteBorder;
        style._hasBlackBackground = hasBlackBackground;
        style._hasWhiteBackground = hasWhiteBackground;
        style._hasDefaultBorder = !hasExplicitBorder;

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
            vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'xlsx', 'xlsxWebview.js')
        );
        const themeStyleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'shared', 'theme.css')
        );
        const styleUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'xlsx', 'xlsxWebview.css')
        );
        const imgUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'table', 'view.png')
        );
        const svgUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'xlsx', 'table.svg')
        );
        const cspSource = webview.cspSource;

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource} 'unsafe-inline';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX Viewer</title>
    <link href="${themeStyleUri}" rel="stylesheet" />
    <link href="${styleUri}" rel="stylesheet" />
    <script>
        window.viewImgUri = "${imgUri}";
        window.logoSvgUri = "${svgUri}";
    </script>
</head>
<body>
    <div class="header-background"></div>

    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="loading-text">Rendering worksheet...</div>
    </div>

    <div class="toolbar" id="toolbar"></div>
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