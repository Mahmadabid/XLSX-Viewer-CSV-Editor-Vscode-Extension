import * as vscode from 'vscode';
import * as Excel from 'exceljs';
import { convertARGBToRGBA, isShadeOfBlack } from './utilities';

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
        try {
            const workbook = new Excel.Workbook();
            
            // Show loading first
            webviewPanel.webview.options = { enableScripts: true };
            webviewPanel.webview.html = this.getLoadingContent();
            
            // Load the workbook
            await workbook.xlsx.readFile(document.uri.fsPath);

            // Generate worksheet data
            const worksheets = workbook.worksheets.map((worksheet, index) => {
                const data = this.extractWorksheetData(worksheet);
                return {
                    name: worksheet.name,
                    index: index,
                    data: data
                };
            });

            // Update with actual content
            webviewPanel.webview.html = this.getWebviewContent(worksheets);

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
                        startRow: startRow,
                        startCol: startCol,
                        endRow: endRow,
                        endCol: endCol
                    });
                }
            });
        } catch (error) {
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
                    style: cellStyle,
                    colNumber: c,
                    rowNumber: r,
                    // Add data attributes for proper color handling
                    isDefaultColor: cellStyle._isDefaultColor || false,
                    hasDefaultBg: !cellStyle.backgroundColor,
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
        } catch (error) {
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
                    row: parseInt(rowStr),
                    col: col
                };
            }
        } catch (error) {
            // Silently continue if parsing fails
        }
        return null;
    }

    private getCellValue(cell: Excel.Cell): string {
        if (!cell || !cell.value) return '';

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

    private getCellStyle(cell: Excel.Cell): any {
        const style: any = {};
        let isDefaultColor = false;
        let hasBlackBorder = false;
        let hasBlackBackground = false;

        // Background color
        if (cell.fill && cell.fill.type === 'pattern' && (cell.fill as any).fgColor) {
            const color = (cell.fill as any).fgColor;
            if (color.argb) {
                const bgColor = convertARGBToRGBA(color.argb);
                style.backgroundColor = bgColor;
                // Check if background is black or shade of black - be very strict
                hasBlackBackground = isShadeOfBlack(bgColor);
            }
        }

        // Font
        if (cell.font) {
            if (cell.font.color && cell.font.color.argb) {
                const fontColor = convertARGBToRGBA(cell.font.color.argb);
                style.color = fontColor;
                // Don't set isDefaultColor for custom colors
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

    private getWebviewContent(worksheets: any[]): string {
        // Calculate the width needed for row headers based on max row number
        const maxRows = Math.max(...worksheets.map(ws => ws.data.maxRow));
        const rowHeaderWidth = Math.max(60, Math.ceil(Math.log10(maxRows + 1)) * 12 + 20);

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX Viewer</title>
    <style>
        * {
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 10px;
            background-color: #ffffff;
            overflow-x: auto;
        }

        body.dark-mode {
            background-color: rgb(33, 33, 33);
        }

        .button-container {
            margin-bottom: 10px;
            display: flex;
            gap: 10px;
            position: sticky;
            top: 0;
            background-color: inherit;
            z-index: 1000;
            padding: 5px 0;
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

        .sheet-selector {
            padding: 8px 16px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            background-color: white;
            cursor: pointer;
        }

        .sheet-selector:focus {
            outline: none;
            border-color: #2196f3 !important;
        }

        .sheet-selector:hover {
            border-color: #1976d2;
        }

        body.dark-mode .sheet-selector {
            background-color: #3c3c3c;
            color: #cccccc;
            border-color: #555;
        }

        .table-container {
            width: 100%;
            overflow: auto;
            background-color: white;
            border: 1px solid #d0d0d0;
            position: relative;
        }

        body.dark-mode .table-container {
            background-color: #252526;
            border-color: #464647;
        }

        table {
            border-collapse: separate;
            border-spacing: 0;
            font-size: 11pt;
            position: relative;
            background-color: white;
            user-select: none;
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
        }

        body.dark-mode table {
            background-color: #1e1e1e;
        }

        th, td {
            border: 1px solid #d0d0d0;
            padding: 2px 4px;
            position: relative;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: cell;
            border-right: 1px solid #d0d0d0;
            border-bottom: 1px solid #d0d0d0;
        }

        /* First row and column get top and left borders */
        thead th {
            border-top: 1px solid #d0d0d0;
        }

        tbody tr th:first-child,
        tbody tr td:first-child {
            border-left: 1px solid #d0d0d0;
        }

        thead th:first-child {
            border-left: 1px solid #d0d0d0;
        }

        body.dark-mode th,
        body.dark-mode td {
            border-color: #464647;
            border-right: 1px solid #464647;
            border-bottom: 1px solid #464647;
        }

        body.dark-mode thead th {
            border-top: 1px solid #464647;
        }

        body.dark-mode tbody tr th:first-child,
        body.dark-mode tbody tr td:first-child {
            border-left: 1px solid #464647;
        }

        body.dark-mode thead th:first-child {
            border-left: 1px solid #464647;
        }

        /* Excel-like headers */
        th.row-header,
        th.col-header {
            background-color: #f0f0f0;
            color: #333;
            font-weight: normal;
            font-size: 11pt;
            text-align: center;
            user-select: none;
            cursor: pointer;
            min-width: 25px;
            position: relative;
            z-index: 1;
        }

        body.dark-mode th.row-header,
        body.dark-mode th.col-header {
            background-color: #2d2d30;
            color: #cccccc;
        }

        th.row-header {
            width: ${rowHeaderWidth}px;
            min-width: ${rowHeaderWidth}px;
            max-width: ${rowHeaderWidth}px;
            text-align: center;
            white-space: nowrap;
            overflow: visible;
        }

        th.col-header {
            height: 20px;
            min-height: 20px;
            min-width: 80px;
        }

        th.row-header:hover,
        th.col-header:hover {
            background-color: rgba(26, 115, 232, 0.2);
        }

        body.dark-mode th.row-header:hover,
        body.dark-mode th.col-header:hover {
            background-color: #3e3e42;
        }

        /* Resize handles */
        .col-resize-handle {
            position: absolute;
            top: 0;
            right: -2px;
            width: 4px;
            height: 100%;
            cursor: col-resize;
            background: transparent;
            z-index: 10;
        }

        .col-resize-handle:hover {
            background: rgba(26, 115, 232, 0.3);
        }

        .row-resize-handle {
            position: absolute;
            bottom: -2px;
            left: 0;
            width: 100%;
            height: 4px;
            cursor: row-resize;
            background: transparent;
            z-index: 10;
        }

        .row-resize-handle:hover {
            background: rgba(26, 115, 232, 0.3);
        }

        /* Cell default styles */
        td {
            background-color: white;
            color: black;
            min-height: 20px;
            white-space: nowrap;
            min-width: 80px;
            vertical-align: top;
            position: relative;
        }

        /* Custom border cells get higher z-index to show over default borders */
        td[style*="border"] {
            z-index: 2;
        }

        body.dark-mode td[data-default-bg="true"] {
            background-color: #1e1e1e !important;
        }

        body.dark-mode td[data-default-color="true"][data-default-bg="true"]:not([data-black-bg="true"]) {
            color: #cccccc !important;
        }

        /* Merged cells styling - FIXED VERTICAL ALIGNMENT */
        td.merged-cell {
            position: relative;
            padding: 0;
            z-index: 3;
        }

        td.merged-cell .cell-content {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            padding: 2px 4px;
            width: 100%;
            height: 100%;
            box-sizing: border-box;
        }

        /* Override vertical alignment for merged cells */
        td.merged-cell[style*="vertical-align: top"] .cell-content {
            align-items: flex-start;
        }

        td.merged-cell[style*="vertical-align: bottom"] .cell-content {
            align-items: flex-end;
        }

        td.merged-cell[style*="vertical-align: middle"] .cell-content {
            align-items: center;
        }

        /* Override text alignment for merged cells */
        td.merged-cell[style*="text-align: left"] .cell-content {
            justify-content: flex-start;
            text-align: left;
        }

        td.merged-cell[style*="text-align: right"] .cell-content {
            justify-content: flex-end;
            text-align: right;
        }

        td.merged-cell[style*="text-align: center"] .cell-content {
            justify-content: center;
            text-align: center;
        }

        /* Selection styles - Excel-like */
        td.selected {
            background-color: rgba(26, 115, 232, 0.1) !important;
            border: 2px solid rgb(26, 115, 232) !important;
            z-index: 4;
        }

        body.dark-mode td.selected {
            background-color: rgba(138, 180, 248, 0.24) !important;
            border: 2px solid rgb(138, 180, 248) !important;
        }

        td.active-cell {
            border: 2px solid rgb(26, 115, 232) !important;
            background-color: white !important;
            z-index: 5;
        }

        body.dark-mode td.active-cell {
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
            z-index: 6;
        }

        body.dark-mode td.active-cell::after {
            background-color: #1ba1e2;
            border-color: #1e1e1e;
        }

        /* Row/Column selection */
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

        body.dark-mode td.column-selected,
        body.dark-mode th.column-selected,
        body.dark-mode td.row-selected,
        body.dark-mode th.row-selected {
            background-color: rgba(138, 180, 248, 0.24) !important;
        }

        /* Copy animation */
        @keyframes copyFlash {
            0% { background-color: inherit; }
            50% { background-color: rgba(26, 115, 232, 0.3) !important; }
            100% { background-color: inherit; }
        }

        body.dark-mode td.copying {
            animation: copyFlashDark 0.2s ease-in-out;
        }

        @keyframes copyFlashDark {
            0% { background-color: inherit; }
            50% { background-color: rgba(138, 180, 248, 0.3) !important; }
            100% { background-color: inherit; }
        }

        td.copying {
            animation: copyFlash 0.2s ease-in-out;
        }

        /* Selection info */
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

        body.dark-mode .selection-info {
            background: rgba(255, 255, 255, 0.8);
            color: black;
        }

        /* Corner cell */
        th.corner-cell {
            background-color: #f0f0f0;
            cursor: pointer;
            width: ${rowHeaderWidth}px;
            min-width: ${rowHeaderWidth}px;
            max-width: ${rowHeaderWidth}px;
        }

        body.dark-mode th.corner-cell {
            background-color: #2d2d30;
        }

        /* Cell content wrapper */
        .cell-content {
            display: block;
            width: 100%;
            height: 100%;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            user-select: none;
            pointer-events: none;
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

        body.dark-mode td:hover {
            background-color: rgba(255, 255, 255, 0.1) !important;
        }

        /* Loading overlay */
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.9);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            flex-direction: column;
        }

        body.dark-mode .loading-overlay {
            background-color: rgba(33, 33, 33, 0.9);
        }

        .loading-overlay.hidden {
            display: none;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #2196f3;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin-bottom: 15px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .loading-text {
            font-size: 16px;
            color: #333;
        }

        body.dark-mode .loading-text {
            color: #cccccc;
        }

        /* Tooltip styles */
        .tooltip {
            position: relative;
            display: inline-block;
        }

        .tooltip .tooltiptext {
            visibility: hidden;
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
            width: 250px;
        }

        .tooltip .tooltiptext span {
            color: #0066cc;
        }

        .tooltip:hover .tooltiptext {
            opacity: 1;
            visibility: visible;
        }

        /* Resize indicator */
        .resize-indicator {
            position: fixed;
            background: rgba(26, 115, 232, 0.8);
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 10px;
            z-index: 1002;
            pointer-events: none;
            display: none;
        }
    </style>
</head>
<body>
    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="loading-text">Rendering worksheet...</div>
    </div>

    <div class="button-container">
        <select id="sheetSelector" class="sheet-selector">
            ${worksheets.map((ws, i) => `<option value="${i}">${ws.name}</option>`).join('')}
        </select>
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
        <button id="toggleMinWidthButton" class="toggle-button tooltip" title="Toggle table width">
            <svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff">
                <g id="SVGRepo_bgCarrier" stroke-width="0"></g>
                <g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g>
                <g id="SVGRepo_iconCarrier">
                    <path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path>
                </g>
            </svg>
            &nbsp; Default
            <span class="tooltiptext">Toggle table width between 50%, 100%, and default. <br><span>It will only work for tables which have size less than editor screen width.</span></span>
        </button>
        <button id="autoFitButton" class="toggle-button" title="Auto-fit columns">
            <svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                <path d="M3 7h18v2H3V7zm0 4h18v2H3v-2zm0 4h18v2H3v-2z"/>
            </svg>
            Auto-fit
        </button>
    </div>
    
    <div class="table-container" id="tableContainer">
        <!-- Table will be inserted here -->
    </div>
    
    <div class="selection-info" id="selectionInfo"></div>
    <div class="resize-indicator" id="resizeIndicator"></div>

    <script>
        // Global state
        const worksheetsData = ${JSON.stringify(worksheets)};
        let currentWorksheet = 0;
        let selectedCells = new Set();
        let activeCell = null;
        let isSelecting = false;
        let selectionStart = null;
        let selectionEnd = null;
        let selectedRows = new Set();
        let selectedColumns = new Set();
        let lastSelectedRow = null;
        let lastSelectedColumn = null;
        
        // Resize state
        let isResizing = false;
        let resizeType = null; // 'column' or 'row'
        let resizeIndex = -1;
        let resizeStartPos = 0;
        let resizeStartSize = 0;
        
        // Helper functions
        function getExcelColumnLabel(n) {
            let label = '';
            while (n > 0) {
                let rem = (n - 1) % 26;
                label = String.fromCharCode(65 + rem) + label;
                n = Math.floor((n - 1) / 26);
            }
            return label;
        }

        function createTable(worksheetData) {
            const data = worksheetData.data;
            
            let html = '<table>';
            
            // Header row
            html += '<thead><tr>';
            html += '<th class="corner-cell"></th>';
            for (let c = 1; c <= data.maxCol; c++) {
                const width = data.columnWidths[c-1] || 80;
                html += '<th class="col-header" data-col="' + (c-1) + '" style="width: ' + width + 'px; min-width: ' + width + 'px;">';
                html += getExcelColumnLabel(c);
                html += '<div class="col-resize-handle" data-col="' + (c-1) + '"></div>';
                html += '</th>';
            }
            html += '</tr></thead><tbody>';
            
            // Data rows
            data.rows.forEach((row, rowIndex) => {
                const height = row.height || 20;
                html += '<tr style="height: ' + height + 'px;">';
                html += '<th class="row-header" data-row="' + rowIndex + '" style="height: ' + height + 'px;">';
                html += row.rowNumber;
                html += '<div class="row-resize-handle" data-row="' + rowIndex + '"></div>';
                html += '</th>';
                
                // Create a virtual column index to account for merged cells
                let virtualColIndex = 0;
                
                for (let actualCol = 1; actualCol <= data.maxCol; actualCol++) {
                    // Find the cell data for this actual column
                    const cellData = row.cells.find(cell => cell.colNumber === actualCol);
                    
                    if (cellData) {
                        const styleStr = formatCellStyle(cellData.style);
                        const cellHeight = height * cellData.rowspan;
                        const cellWidth = data.columnWidths.slice(actualCol-1, actualCol-1+cellData.colspan).reduce((sum, w) => sum + (w || 80), 0);
                        
                        html += '<td';
                        html += ' data-row="' + rowIndex + '"';
                        html += ' data-col="' + virtualColIndex + '"';
                        if (cellData.hasDefaultBg) html += ' data-default-bg="true"';
                        if (cellData.isDefaultColor) html += ' data-default-color="true"';
                        if (cellData.hasBlackBorder) html += ' data-black-border="true"';
                        if (cellData.hasBlackBackground) html += ' data-black-bg="true"';
                        if (cellData.isEmpty) html += ' data-empty="true"';
                        html += ' data-original-color="' + cellData.originalColor + '"';
                        
                        // Add rowspan and colspan for merged cells
                        if (cellData.rowspan > 1) html += ' rowspan="' + cellData.rowspan + '"';
                        if (cellData.colspan > 1) html += ' colspan="' + cellData.colspan + '"';
                        if (cellData.isMerged) html += ' class="merged-cell"';
                        
                        // Set explicit height and width for merged cells
                        let cellStyleStr = styleStr;
                        if (cellData.isMerged) {
                            cellStyleStr += 'height: ' + cellHeight + 'px; width: ' + cellWidth + 'px;';
                        } else {
                            cellStyleStr += 'height: ' + height + 'px;';
                        }
                        
                        html += ' style="' + cellStyleStr + '"';
                        html += '>';
                        html += '<span class="cell-content">' + (cellData.value || '&nbsp;') + '</span>';
                        html += '</td>';
                    }
                    
                    virtualColIndex++;
                }
                
                html += '</tr>';
            });
            
            html += '</tbody></table>';
            return html;
        }

        function formatCellStyle(style) {
            let css = '';
            
            if (style.backgroundColor) css += 'background-color: ' + style.backgroundColor + ';';
            if (style.color) css += 'color: ' + style.color + ';';
            if (style.fontWeight) css += 'font-weight: ' + style.fontWeight + ';';
            if (style.fontStyle) css += 'font-style: ' + style.fontStyle + ';';
            if (style.textDecoration) css += 'text-decoration: ' + style.textDecoration + ';';
            if (style.fontSize) css += 'font-size: ' + style.fontSize + ';';
            if (style.fontFamily) css += 'font-family: ' + style.fontFamily + ';';
            if (style.textAlign) css += 'text-align: ' + style.textAlign + ';';
            if (style.verticalAlign) css += 'vertical-align: ' + style.verticalAlign + ';';
            if (style.whiteSpace) css += 'white-space: ' + style.whiteSpace + ';';
            if (style.wordWrap) css += 'word-wrap: ' + style.wordWrap + ';';
            if (style.paddingLeft) css += 'padding-left: ' + style.paddingLeft + ';';
            
            // Borders
            if (style.border) {
                if (style.border.top) css += 'border-top: ' + style.border.top + ';';
                if (style.border.right) css += 'border-right: ' + style.border.right + ';';
                if (style.border.bottom) css += 'border-bottom: ' + style.border.bottom + ';';
                if (style.border.left) css += 'border-left: ' + style.border.left + ';';
            }
            
            return css;
        }

        function showLoading() {
            document.getElementById('loadingOverlay').classList.remove('hidden');
        }

        function hideLoading() {
            document.getElementById('loadingOverlay').classList.add('hidden');
        }

        function renderWorksheet(index) {
            showLoading();
            
            // Use setTimeout to allow loading overlay to show
            setTimeout(() => {
                const container = document.getElementById('tableContainer');
                container.innerHTML = createTable(worksheetsData[index]);
                initializeSelection();
                initializeResize();
                hideLoading();
            }, 100);
        }

        function initializeResize() {
            const table = document.querySelector('table');
            if (!table) return;

            // Column resize handles
            table.addEventListener('mousedown', (e) => {
                if (e.target.classList.contains('col-resize-handle')) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    isResizing = true;
                    resizeType = 'column';
                    resizeIndex = parseInt(e.target.dataset.col);
                    resizeStartPos = e.clientX;
                    
                    const header = e.target.parentElement;
                    resizeStartSize = header.offsetWidth;
                    
                    document.body.style.cursor = 'col-resize';
                    document.getElementById('resizeIndicator').style.display = 'block';
                    
                    return false;
                }
                
                if (e.target.classList.contains('row-resize-handle')) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    isResizing = true;
                    resizeType = 'row';
                    resizeIndex = parseInt(e.target.dataset.row);
                    resizeStartPos = e.clientY;
                    
                    const header = e.target.parentElement;
                    resizeStartSize = header.offsetHeight;
                    
                    document.body.style.cursor = 'row-resize';
                    document.getElementById('resizeIndicator').style.display = 'block';
                    
                    return false;
                }
            });

            // Mouse move for resizing
            document.addEventListener('mousemove', (e) => {
                if (!isResizing) return;
                
                const indicator = document.getElementById('resizeIndicator');
                
                if (resizeType === 'column') {
                    const delta = e.clientX - resizeStartPos;
                    const newSize = Math.max(20, resizeStartSize + delta);
                    
                    // Update all cells in this column
                    const headers = table.querySelectorAll('th.col-header[data-col="' + resizeIndex + '"]');
                    const cells = table.querySelectorAll('td[data-col="' + resizeIndex + '"]');
                    
                    headers.forEach(header => {
                        header.style.width = newSize + 'px';
                        header.style.minWidth = newSize + 'px';
                    });
                    
                    cells.forEach(cell => {
                        if (!cell.getAttribute('colspan') || cell.getAttribute('colspan') === '1') {
                            cell.style.width = newSize + 'px';
                            cell.style.minWidth = newSize + 'px';
                        }
                    });
                    
                    indicator.style.left = e.clientX + 'px';
                    indicator.style.top = e.clientY + 'px';
                    indicator.textContent = newSize + 'px';
                    
                } else if (resizeType === 'row') {
                    const delta = e.clientY - resizeStartPos;
                    const newSize = Math.max(15, resizeStartSize + delta);
                    
                    // Update the row
                    const headers = table.querySelectorAll('th.row-header[data-row="' + resizeIndex + '"]');
                    const row = table.querySelectorAll('tr')[resizeIndex + 1]; // +1 for header row
                    
                    headers.forEach(header => {
                        header.style.height = newSize + 'px';
                    });
                    
                    if (row) {
                        row.style.height = newSize + 'px';
                        const cells = row.querySelectorAll('td');
                        cells.forEach(cell => {
                            if (!cell.getAttribute('rowspan') || cell.getAttribute('rowspan') === '1') {
                                cell.style.height = newSize + 'px';
                            }
                        });
                    }
                    
                    indicator.style.left = e.clientX + 'px';
                    indicator.style.top = e.clientY + 'px';
                    indicator.textContent = newSize + 'px';
                }
            });

            // Mouse up to end resizing
            document.addEventListener('mouseup', () => {
                if (isResizing) {
                    isResizing = false;
                    resizeType = null;
                    resizeIndex = -1;
                    document.body.style.cursor = '';
                    document.getElementById('resizeIndicator').style.display = 'none';
                }
            });

            // Double-click to auto-fit
            table.addEventListener('dblclick', (e) => {
                if (e.target.classList.contains('col-resize-handle')) {
                    e.preventDefault();
                    autoFitColumn(parseInt(e.target.dataset.col));
                } else if (e.target.classList.contains('row-resize-handle')) {
                    e.preventDefault();
                    autoFitRow(parseInt(e.target.dataset.row));
                }
            });
        }

        function autoFitColumn(colIndex) {
            const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
            let maxWidth = 50;
            
            cells.forEach(cell => {
                const content = cell.textContent.trim();
                const tempSpan = document.createElement('span');
                tempSpan.style.visibility = 'hidden';
                tempSpan.style.position = 'absolute';
                tempSpan.style.whiteSpace = 'nowrap';
                tempSpan.style.font = window.getComputedStyle(cell).font;
                tempSpan.textContent = content;
                document.body.appendChild(tempSpan);
                
                const contentWidth = tempSpan.offsetWidth + 10; // padding
                maxWidth = Math.max(maxWidth, contentWidth);
                
                document.body.removeChild(tempSpan);
            });
            
            maxWidth = Math.min(maxWidth, 300); // Cap at 300px
            
            cells.forEach(cell => {
                cell.style.width = maxWidth + 'px';
                cell.style.minWidth = maxWidth + 'px';
            });
        }

        function autoFitRow(rowIndex) {
            const row = document.querySelectorAll('tr')[rowIndex + 1]; // +1 for header row
            if (!row) return;
            
            const cells = row.querySelectorAll('td');
            let maxHeight = 20;
            
            cells.forEach(cell => {
                const content = cell.textContent.trim();
                if (content.length > 50) { // If content is long, might need more height
                    maxHeight = Math.max(maxHeight, 40);
                }
            });
            
            row.style.height = maxHeight + 'px';
            const headers = document.querySelectorAll('th.row-header[data-row="' + rowIndex + '"]');
            headers.forEach(header => {
                header.style.height = maxHeight + 'px';
            });
            
            cells.forEach(cell => {
                if (!cell.getAttribute('rowspan') || cell.getAttribute('rowspan') === '1') {
                    cell.style.height = maxHeight + 'px';
                }
            });
        }

        function autoFitAllColumns() {
            const data = worksheetsData[currentWorksheet].data;
            for (let c = 0; c < data.maxCol; c++) {
                autoFitColumn(c);
            }
        }

        function clearSelection() {
            document.querySelectorAll('.selected, .active-cell, .row-selected, .column-selected').forEach(el => {
                el.classList.remove('selected', 'active-cell', 'row-selected', 'column-selected');
            });
            selectedCells.clear();
            selectedRows.clear();
            selectedColumns.clear();
            activeCell = null;
            lastSelectedRow = null;
            lastSelectedColumn = null;
            document.getElementById('selectionInfo').style.display = 'none';
        }

        function selectCell(cell, isMulti = false) {
            if (!isMulti) {
                clearSelection();
            }
            
            cell.classList.add('selected');
            cell.classList.add('active-cell');
            selectedCells.add(cell);
            activeCell = cell;
            updateSelectionInfo();
        }

        function selectRange(startRow, startCol, endRow, endCol) {
            clearSelection();
            
            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);
            
            const cells = document.querySelectorAll('td');
            cells.forEach(cell => {
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                
                if (row >= minRow && row <= maxRow && col >= minCol && col <= maxCol) {
                    cell.classList.add('selected');
                    selectedCells.add(cell);
                }
            });
            
            // Set active cell
            const startCell = document.querySelector('td[data-row="' + startRow + '"][data-col="' + startCol + '"]');
            if (startCell) {
                startCell.classList.add('active-cell');
                activeCell = startCell;
            }
            
            updateSelectionInfo();
        }

        function selectRow(rowIndex, ctrlKey, shiftKey) {
            if (!ctrlKey && !shiftKey) {
                clearSelection();
                lastSelectedRow = rowIndex;
            }

            if (shiftKey && lastSelectedRow !== null && lastSelectedRow !== rowIndex) {
                if (!ctrlKey) {
                    clearSelection();
                }
                
                const minRow = Math.min(lastSelectedRow, rowIndex);
                const maxRow = Math.max(lastSelectedRow, rowIndex);
                
                for (let row = minRow; row <= maxRow; row++) {
                    if (!selectedRows.has(row)) {
                        selectedRows.add(row);
                        const cells = document.querySelectorAll('td[data-row="' + row + '"], th[data-row="' + row + '"]');
                        cells.forEach(cell => {
                            cell.classList.add('row-selected');
                            if (cell.tagName === 'TD') {
                                selectedCells.add(cell);
                            }
                        });
                    }
                }
            } else if (ctrlKey) {
                if (selectedRows.has(rowIndex)) {
                    selectedRows.delete(rowIndex);
                    const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
                    cells.forEach(cell => {
                        cell.classList.remove('row-selected');
                        if (cell.tagName === 'TD') selectedCells.delete(cell);
                    });
                } else {
                    selectedRows.add(rowIndex);
                    const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
                    cells.forEach(cell => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') {
                            selectedCells.add(cell);
                        }
                    });
                }
            } else {
                selectedRows.add(rowIndex);
                const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') {
                        selectedCells.add(cell);
                    }
                });
            }
            
            updateSelectionInfo();
        }

        function selectColumn(colIndex, ctrlKey, shiftKey) {
            if (!ctrlKey && !shiftKey) {
                clearSelection();
                lastSelectedColumn = colIndex;
            }

            if (shiftKey && lastSelectedColumn !== null && lastSelectedColumn !== colIndex) {
                if (!ctrlKey) {
                    clearSelection();
                }
                
                const minCol = Math.min(lastSelectedColumn, colIndex);
                const maxCol = Math.max(lastSelectedColumn, colIndex);
                
                for (let col = minCol; col <= maxCol; col++) {
                    if (!selectedColumns.has(col)) {
                        selectedColumns.add(col);
                        const cells = document.querySelectorAll('td[data-col="' + col + '"], th[data-col="' + col + '"]');
                        cells.forEach(cell => {
                            cell.classList.add('column-selected');
                            if (cell.tagName === 'TD') {
                                selectedCells.add(cell);
                            }
                        });
                    }
                }
            } else if (ctrlKey) {
                if (selectedColumns.has(colIndex)) {
                    selectedColumns.delete(colIndex);
                    const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
                    cells.forEach(cell => {
                        cell.classList.remove('column-selected');
                        if (cell.tagName === 'TD') selectedCells.delete(cell);
                    });
                } else {
                    selectedColumns.add(colIndex);
                    const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
                    cells.forEach(cell => {
                        cell.classList.add('column-selected');
                        if (cell.tagName === 'TD') {
                            selectedCells.add(cell);
                        }
                    });
                }
            } else {
                selectedColumns.add(colIndex);
                const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') {
                        selectedCells.add(cell);
                    }
                });
            }
            
            updateSelectionInfo();
        }

        function updateSelectionInfo() {
            const info = document.getElementById('selectionInfo');
            if (selectedCells.size > 1) {
                const rows = new Set();
                const cols = new Set();
                selectedCells.forEach(cell => {
                    rows.add(cell.dataset.row);
                    cols.add(cell.dataset.col);
                });
                info.textContent = rows.size + 'R × ' + cols.size + 'C';
                info.style.display = 'block';
            } else {
                info.style.display = 'none';
            }
        }

        function copySelection() {
            if (selectedCells.size === 0) return;
            
            // Get all selected cells
            const cellsArray = Array.from(selectedCells);
            const cellData = cellsArray.map(cell => ({
                row: parseInt(cell.dataset.row),
                col: parseInt(cell.dataset.col),
                text: cell.textContent.trim()
            }));
            
            // Sort by row then column
            cellData.sort((a, b) => a.row - b.row || a.col - b.col);
            
            // Find bounds
            const rows = [...new Set(cellData.map(c => c.row))].sort((a, b) => a - b);
            const cols = [...new Set(cellData.map(c => c.col))].sort((a, b) => a - b);
            const minRow = Math.min(...rows);
            const maxRow = Math.max(...rows);
            const minCol = Math.min(...cols);
            const maxCol = Math.max(...cols);
            
            // Build 2D array
            const grid = [];
            for (let r = minRow; r <= maxRow; r++) {
                const row = [];
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = cellData.find(cd => cd.row === r && cd.col === c);
                    row.push(cell ? cell.text : '');
                }
                grid.push(row);
            }
            
            // Convert to TSV format (tab-separated values)
            const text = grid.map(row => row.join('\\t')).join('\\n');
            
            // Copy to clipboard
            navigator.clipboard.writeText(text).then(() => {
                // Visual feedback
                selectedCells.forEach(cell => {
                    cell.classList.add('copying');
                    setTimeout(() => cell.classList.remove('copying'), 300);
                });
            }).catch(err => {
                // Silently handle copy errors
            });
        }

        function invertColor(color) {
            const match = color.match(/rgb\KATEX_INLINE_OPEN(\\d+),\\s*(\\d+),\\s*(\\d+)\KATEX_INLINE_CLOSE/);
            if (!match) return color;
            
            const r = 255 - parseInt(match[1]);
            const g = 255 - parseInt(match[2]);
            const b = 255 - parseInt(match[3]);
            return \`rgb(\${r}, \${g}, \${b})\`;
        }

        function initializeSelection() {
            const tableContainer = document.getElementById('tableContainer');
            const table = tableContainer.querySelector('table');
            
            // Prevent text selection
            table.addEventListener('selectstart', (e) => {
                e.preventDefault();
                return false;
            });
            
            // Cell selection
            table.addEventListener('mousedown', (e) => {
                // Skip if this is a resize handle
                if (e.target.classList.contains('col-resize-handle') || 
                    e.target.classList.contains('row-resize-handle')) {
                    return;
                }
                
                const target = e.target.closest('td, th');
                if (!target) return;
                
                e.preventDefault(); // Prevent text selection
                
                // Column header click
                if (target.classList.contains('col-header')) {
                    const colIndex = parseInt(target.dataset.col);
                    
                    // Always update last selected column
                    if (!e.shiftKey) {
                        lastSelectedColumn = colIndex;
                    }
                    
                    selectColumn(colIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                    return;
                }
                
                // Row header click
                if (target.classList.contains('row-header')) {
                    const rowIndex = parseInt(target.dataset.row);
                    
                    // Always update last selected row
                    if (!e.shiftKey) {
                        lastSelectedRow = rowIndex;
                    }
                    
                    selectRow(rowIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                    return;
                }
                
                // Corner cell - select all
                if (target.classList.contains('corner-cell')) {
                    clearSelection();
                    const allCells = table.querySelectorAll('td');
                    allCells.forEach(cell => {
                        cell.classList.add('selected');
                        selectedCells.add(cell);
                    });
                    if (allCells.length > 0) {
                        allCells[0].classList.add('active-cell');
                        activeCell = allCells[0];
                    }
                    updateSelectionInfo();
                    return;
                }
                
                // Regular cell click
                if (target.tagName === 'TD') {
                    const row = parseInt(target.dataset.row);
                    const col = parseInt(target.dataset.col);
                    
                    // Get the actual cell bounds for merged cells
                    const rowspan = parseInt(target.getAttribute('rowspan')) || 1;
                    const colspan = parseInt(target.getAttribute('colspan')) || 1;
                    
                    if (e.ctrlKey || e.metaKey) {
                        // Multi-selection
                        if (target.classList.contains('selected')) {
                            target.classList.remove('selected');
                            selectedCells.delete(target);
                            if (target === activeCell) {
                                target.classList.remove('active-cell');
                                activeCell = null;
                                
                                const remainingSelected = document.querySelector('td.selected');
                                if (remainingSelected) {
                                    remainingSelected.classList.add('active-cell');
                                    activeCell = remainingSelected;
                                }
                            }
                        } else {
                            target.classList.add('selected');
                            selectedCells.add(target);
                            if (activeCell) {
                                activeCell.classList.remove('active-cell');
                            }
                            target.classList.add('active-cell');
                            activeCell = target;
                        }
                        updateSelectionInfo();
                    } else if (e.shiftKey && activeCell) {
                        // Range selection with merged cells
                        const startRow = parseInt(activeCell.dataset.row);
                        const startCol = parseInt(activeCell.dataset.col);
                        
                        selectRange(startRow, startCol, row, col);
                    } else {
                        // Single selection
                        isSelecting = true;
                        selectionStart = { row, col };
                        selectCell(target);
                    }
                }
            });
            
            // Drag selection
            table.addEventListener('mousemove', (e) => {
                if (!isSelecting || !selectionStart) return;
                
                const target = e.target.closest('td');
                if (!target) return;
                
                const row = parseInt(target.dataset.row);
                const col = parseInt(target.dataset.col);
                
                if (!selectionEnd || selectionEnd.row !== row || selectionEnd.col !== col) {
                    selectionEnd = { row, col };
                    selectRange(selectionStart.row, selectionStart.col, row, col);
                }
            });
            
            // End selection
            document.addEventListener('mouseup', () => {
                isSelecting = false;
                selectionStart = null;
                selectionEnd = null;
            });
            
            // Keyboard shortcuts
            document.addEventListener('keydown', (e) => {
                // Copy
                if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                    e.preventDefault();
                    copySelection();
                }
                
                // Select all
                if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
                    e.preventDefault();
                    const allCells = table.querySelectorAll('td');
                    clearSelection();
                    allCells.forEach(cell => {
                        cell.classList.add('selected');
                        selectedCells.add(cell);
                    });
                    if (allCells.length > 0) {
                        allCells[0].classList.add('active-cell');
                        activeCell = allCells[0];
                    }
                    updateSelectionInfo();
                }
            });
            
            // Click outside to clear selection
            document.addEventListener('click', (e) => {
                if (!e.target.closest('table') && !e.target.closest('.button-container')) {
                    clearSelection();
                }
            });
        }

        // Initialize
        document.addEventListener('DOMContentLoaded', () => {
            // Render first worksheet
            renderWorksheet(0);
            
            // Sheet selector
            document.getElementById('sheetSelector').addEventListener('change', (e) => {
                currentWorksheet = parseInt(e.target.value);
                clearSelection();
                renderWorksheet(currentWorksheet);
            });
            
            // Auto-fit button
            document.getElementById('autoFitButton').addEventListener('click', () => {
                autoFitAllColumns();
            });
            
            // Dark mode toggle with fixed color inversion logic
            document.getElementById('toggleBackgroundButton').addEventListener('click', () => {
                document.body.classList.toggle('dark-mode');
                const isDarkMode = document.body.classList.contains('dark-mode');

                if (isDarkMode) {
                    document.getElementById('lightIcon').style.display = 'block';
                    document.getElementById('darkIcon').style.display = 'none';
                } else {
                    document.getElementById('lightIcon').style.display = 'none';
                    document.getElementById('darkIcon').style.display = 'block';
                }

                // Handle default background cells (cells with no custom background)
                const defaultBgCells = document.querySelectorAll('td[data-default-bg="true"]');
                defaultBgCells.forEach(cell => {
                    if (isDarkMode) {
                        cell.style.backgroundColor = "rgb(33, 33, 33)";
                    } else {
                        cell.style.backgroundColor = "rgb(255, 255, 255)";
                    }
                });

                // Handle default color cells - ONLY those with default backgrounds AND NOT black backgrounds
                const defaultColorCells = document.querySelectorAll('td[data-default-color="true"][data-default-bg="true"]:not([data-black-bg="true"])');
                defaultColorCells.forEach(cell => {
                    if (isDarkMode) {
                        cell.style.color = "rgb(255, 255, 255)";
                    } else {
                        cell.style.color = "rgb(0, 0, 0)";
                    }
                });

                // Handle text inversion ONLY for cells with CONFIRMED black backgrounds
                const blackBgCells = document.querySelectorAll('td[data-black-bg="true"]');
                blackBgCells.forEach(cell => {
                    const originalColor = cell.getAttribute('data-original-color');
                    if (isDarkMode) {
                        // Invert the text color for better readability on black backgrounds
                        cell.style.color = invertColor(originalColor);
                    } else {
                        // Restore original color
                        cell.style.color = originalColor;
                    }
                });

                // Handle black borders
                const blackBorderCells = document.querySelectorAll('td[data-black-border="true"]');
                blackBorderCells.forEach(cell => {
                    if (isDarkMode) {
                        cell.style.borderColor = 'rgb(255, 255, 255)';
                    } else {
                        cell.style.borderColor = 'rgb(0, 0, 0)';
                    }
                });
            });
            
            // Width toggle button
            let minWidthState = 2; // 0: 50%, 1: 100%, 2: default
            const minWidthValues = ['50%', '100%', ''];
            const buttonLabels = [
                '<svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path></g></svg>&nbsp; 50%',
                '<svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path></g></svg>&nbsp; 100%',
                '<svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path></g></svg>&nbsp; Default'
            ];

            document.getElementById('toggleMinWidthButton').addEventListener('click', (e) => {
                const table = document.querySelector('table');
                if (table) {
                    minWidthState = (minWidthState + 1) % 3;
                    table.style.minWidth = minWidthValues[minWidthState];
                    
                    const btn = e.target.closest('button');
                    const tooltip = btn.querySelector('.tooltiptext');
                    const tooltipContent = tooltip.innerHTML;
                    
                    // Hide tooltip temporarily
                    tooltip.style.opacity = '0';
                    tooltip.style.visibility = 'hidden';
                    
                    // Update button content while preserving tooltip
                    btn.innerHTML = buttonLabels[minWidthState];
                    
                    // Recreate and reattach the tooltip
                    const newTooltip = document.createElement('span');
                    newTooltip.className = 'tooltiptext';
                    newTooltip.innerHTML = tooltipContent;
                    btn.appendChild(newTooltip);
                }
            });
        });
    </script>
</body>
</html>`;
    }
}