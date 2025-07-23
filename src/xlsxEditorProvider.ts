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
            maxCol: 0
        };

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

        // Extract all cell data
        for (let r = 1; r <= maxRow; r++) {
            const rowData: any = {
                rowNumber: r,
                cells: [],
                height: worksheet.getRow(r).height
            };

            for (let c = 1; c <= maxCol; c++) {
                const cell = worksheet.getRow(r).getCell(c);
                const cellData: any = {
                    value: this.getCellValue(cell),
                    style: this.getCellStyle(cell),
                    colNumber: c
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

    private getCellValue(cell: Excel.Cell): string {
        if (!cell.value) return '';

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

        // Background color
        if (cell.fill && cell.fill.type === 'pattern' && (cell.fill as any).fgColor) {
            const color = (cell.fill as any).fgColor;
            if (color.argb) {
                style.backgroundColor = convertARGBToRGBA(color.argb);
            }
        }

        // Font
        if (cell.font) {
            if (cell.font.color && cell.font.color.argb) {
                style.color = convertARGBToRGBA(cell.font.color.argb);
            }
            if (cell.font.bold) style.fontWeight = 'bold';
            if (cell.font.italic) style.fontStyle = 'italic';
            if (cell.font.underline) style.textDecoration = 'underline';
            if (cell.font.strike) style.textDecoration = (style.textDecoration || '') + ' line-through';
            if (cell.font.size) style.fontSize = `${cell.font.size}pt`;
            if (cell.font.name) style.fontFamily = cell.font.name;
        }

        // Alignment
        if (cell.alignment) {
            if (cell.alignment.horizontal) {
                style.textAlign = cell.alignment.horizontal;
            }
            if (cell.alignment.vertical) {
                style.verticalAlign = cell.alignment.vertical === 'middle' ? 'middle' : cell.alignment.vertical;
            }
            if (cell.alignment.wrapText) {
                style.whiteSpace = 'pre-wrap';
                style.wordWrap = 'break-word';
            }
        }

        // Borders
        if (cell.border) {
            style.border = {};
            ['top', 'right', 'bottom', 'left'].forEach(side => {
                const border = (cell.border as any)[side];
                if (border && border.style) {
                    const color = border.color && border.color.argb 
                        ? convertARGBToRGBA(border.color.argb) 
                        : '#000000';
                    
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
                    
                    style.border[side] = `${width} ${styleStr} ${color}`;
                }
            });
        }

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
        const rowHeaderWidth = Math.max(60, Math.ceil(Math.log10(maxRows + 1)) * 12 + 20); // Dynamic width based on number of digits

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
        }

        body.dark-mode .table-container {
            background-color: #252526;
            border-color: #464647;
        }

        table {
            border-collapse: collapse;
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
        }

        body.dark-mode th,
        body.dark-mode td {
            border-color: #464647;
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

        /* Cell default styles */
        td {
            background-color: white;
            color: black;
            min-height: 20px;
            height: 20px;
            white-space: nowrap;
            min-width: 80px;
        }

        body.dark-mode td[data-default-bg="true"] {
            background-color: #1e1e1e !important;
        }

        body.dark-mode td[data-default-color="true"] {
            color: #cccccc !important;
        }

        /* Selection styles - Excel-like */
        td.selected {
            background-color: rgba(26, 115, 232, 0.1) !important;
            border: 2px solid rgb(26, 115, 232) !important;
            z-index: 2;
        }

        body.dark-mode td.selected {
            background-color: rgba(138, 180, 248, 0.24) !important;
            border: 2px solid rgb(138, 180, 248) !important;
        }

        td.active-cell {
            border: 2px solid rgb(26, 115, 232) !important;
            background-color: white !important;
            z-index: 3;
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
        <button id="toggleMinWidthButton" class="toggle-button" title="Toggle table width">
            <svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff">
                <g id="SVGRepo_bgCarrier" stroke-width="0"></g>
                <g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g>
                <g id="SVGRepo_iconCarrier">
                    <path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path>
                </g>
            </svg>
            &nbsp; Default
        </button>
    </div>
    
    <div class="table-container" id="tableContainer">
        <!-- Table will be inserted here -->
    </div>
    
    <div class="selection-info" id="selectionInfo"></div>

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
                html += '<th class="col-header" data-col="' + (c-1) + '">' + getExcelColumnLabel(c) + '</th>';
            }
            html += '</tr></thead><tbody>';
            
            // Data rows
            data.rows.forEach((row, rowIndex) => {
                html += '<tr>';
                html += '<th class="row-header" data-row="' + rowIndex + '">' + row.rowNumber + '</th>';
                
                row.cells.forEach((cell, colIndex) => {
                    const styleStr = formatCellStyle(cell.style);
                    const isDefaultBg = !cell.style.backgroundColor;
                    const isDefaultColor = !cell.style.color;
                    
                    html += '<td';
                    html += ' data-row="' + rowIndex + '"';
                    html += ' data-col="' + colIndex + '"';
                    if (isDefaultBg) html += ' data-default-bg="true"';
                    if (isDefaultColor) html += ' data-default-color="true"';
                    html += ' style="' + styleStr + '"';
                    html += '>';
                    html += '<span class="cell-content">' + (cell.value || '') + '</span>';
                    html += '</td>';
                });
                
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
                hideLoading();
            }, 100);
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
                console.error('Failed to copy:', err);
            });
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
                        // Range selection
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
            
            // Dark mode toggle
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
            
            // Width toggle button
            let minWidthState = 2; // 0: 50%, 1: 100%, 2: default
            const minWidthValues = ['50%', '100%', ''];
            const buttonLabels = ['50%', '100%', 'Default'];

            document.getElementById('toggleMinWidthButton').addEventListener('click', () => {
                const table = document.querySelector('table');
                if (table) {
                    minWidthState = (minWidthState + 1) % 3;
                    table.style.minWidth = minWidthValues[minWidthState];
                    
                    const btn = document.getElementById('toggleMinWidthButton');
                    const svg = btn.querySelector('svg');
                    btn.innerHTML = '';
                    btn.appendChild(svg);
                    btn.innerHTML += '&nbsp; ' + buttonLabels[minWidthState];
                }
            });
        });
    </script>
</body>
</html>`;
    }
}