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
            console.error('Error in XLSX viewer:', error);
            vscode.window.showErrorMessage(`Error reading XLSX file: ${error}`);
            webviewPanel.webview.html = this.getErrorContent(error);
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

    private getErrorContent(error: any): string {
        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Error Loading XLSX</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .error-container {
            max-width: 600px;
            margin: 50px auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #d73a49;
            margin-top: 0;
        }
        pre {
            background: #f6f8fa;
            padding: 16px;
            border-radius: 6px;
            overflow-x: auto;
        }
    </style>
</head>
<body>
    <div class="error-container">
        <h1>Error Loading Spreadsheet</h1>
        <p>We encountered an error while loading your spreadsheet file.</p>
        <pre>${error?.message || 'Unknown error'}</pre>
        <p>Please check that the file is a valid Excel file and try again.</p>
    </div>
</body>
</html>`;
    }

    private extractWorksheetData(worksheet: Excel.Worksheet): any {
        const data: any = {
            rows: [],
            maxRow: 0,
            maxCol: 0,
            mergedCells: [],
            hiddenCells: []
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

        // Extract merged cells using the model property
        try {
            const merges = (worksheet.model as any).merges || {};
            for (const merge of Object.values(merges)) {
                const mergeRange = merge as string;
                const match = mergeRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
                if (match) {
                    const startCol = this.columnLetterToNumber(match[1]);
                    const startRow = parseInt(match[2]);
                    const endCol = this.columnLetterToNumber(match[3]);
                    const endRow = parseInt(match[4]);
                    
                    data.mergedCells.push({
                        top: startRow - 1,
                        left: startCol - 1,
                        bottom: endRow - 1,
                        right: endCol - 1,
                        rowSpan: endRow - startRow + 1,
                        colSpan: endCol - startCol + 1
                    });

                    // Mark hidden cells
                    for (let r = startRow; r <= endRow; r++) {
                        for (let c = startCol; c <= endCol; c++) {
                            if (r !== startRow || c !== startCol) {
                                data.hiddenCells.push(`${r - 1}_${c - 1}`);
                            }
                        }
                    }
                }
            }
        } catch (e) {
            console.warn('Could not extract merged cells:', e);
        }

        // Extract all cell data
        for (let r = 1; r <= maxRow; r++) {
            const excelRow = worksheet.getRow(r);
            const rowData: any = {
                rowNumber: r,
                cells: [],
                height: excelRow.height || 21
            };

            for (let c = 1; c <= maxCol; c++) {
                const cell = excelRow.getCell(c);
                const cellData: any = {
                    value: this.getCellValue(cell),
                    style: this.getCellStyle(cell),
                    colNumber: c,
                    isMerged: cell.isMerged || false,
                };

                // Find merge info for this cell
                const mergeInfo = data.mergedCells.find((m: any) => 
                    m.top === r - 1 && m.left === c - 1
                );
                if (mergeInfo) {
                    cellData.rowSpan = mergeInfo.rowSpan;
                    cellData.colSpan = mergeInfo.colSpan;
                }

                rowData.cells.push(cellData);
            }

            data.rows.push(rowData);
        }

        // Column widths
        data.columnWidths = [];
        for (let c = 1; c <= maxCol; c++) {
            const col = worksheet.getColumn(c);
            const width = col.width ? col.width * 7.5 : 64;
            data.columnWidths.push(width);
        }

        return data;
    }

    private columnLetterToNumber(letters: string): number {
        let result = 0;
        for (let i = 0; i < letters.length; i++) {
            result = result * 26 + (letters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return result;
    }

    private getCellValue(cell: Excel.Cell): string {
        if (!cell.value) return '';

        try {
            if (cell.type === Excel.ValueType.Hyperlink) {
                const hyperlinkValue = cell.value as Excel.CellHyperlinkValue;
                return hyperlinkValue.text || '';
            } else if (cell.type === Excel.ValueType.Formula) {
                return cell.result?.toString() || '';
            } else if (cell.type === Excel.ValueType.RichText) {
                const richTextValue = cell.value as Excel.CellRichTextValue;
                return richTextValue.richText.map((rt: Excel.CellRichTextValue['richText'][0]) => rt.text).join('');
            } else if (cell.type === Excel.ValueType.Date) {
                const dateValue = cell.value as Date;
                return dateValue.toLocaleDateString();
            } else if (cell.value instanceof Date) {
                return cell.value.toLocaleDateString();
            } else {
                return cell.value.toString();
            }
        } catch (e) {
            console.warn('Error getting cell value:', e);
            return String(cell.value || '');
        }
    }

    private getCellStyle(cell: Excel.Cell): any {
        const style: any = {};

        try {
            // Background color
            let hasBackground = false;
            let isLightBg = false;
            
            if (cell.fill && cell.fill.type === 'pattern' && 'fgColor' in cell.fill && cell.fill.fgColor) {
                const color = cell.fill.fgColor;
                if ('argb' in color && color.argb) {
                    const convertedColor = convertARGBToRGBA(color.argb);
                    style.backgroundColor = convertedColor;
                    hasBackground = true;
                    
                    // Calculate if background is light
                    const match = convertedColor.match(/rgba?KATEX_INLINE_OPEN(\d+),\s*(\d+),\s*(\d+)/);
                    if (match) {
                        const r = parseInt(match[1]);
                        const g = parseInt(match[2]);
                        const b = parseInt(match[3]);
                        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
                        isLightBg = luminance > 0.5;
                    }
                    
                    console.log(`Cell bg: ${convertedColor}, luminance: ${isLightBg ? 'light' : 'dark'}`);
                }
            }
            
            style.hasBackground = hasBackground;
            style.isLightBackground = isLightBg;

            // Font
            let hasCustomTextColor = false;
            if (cell.font) {
                if (cell.font.color && 'argb' in cell.font.color && cell.font.color.argb) {
                    style.color = convertARGBToRGBA(cell.font.color.argb);
                    hasCustomTextColor = true;
                    console.log(`Cell text color: ${style.color}`);
                }
                if (cell.font.bold) style.fontWeight = 'bold';
                if (cell.font.italic) style.fontStyle = 'italic';
                if (cell.font.underline) style.textDecoration = 'underline';
                if (cell.font.strike) {
                    style.textDecoration = (style.textDecoration || '') + ' line-through';
                }
                if (cell.font.size) style.fontSize = cell.font.size;
                if (cell.font.name) style.fontFamily = cell.font.name;
            }
            
            style.hasCustomTextColor = hasCustomTextColor;

            // Alignment
            if (cell.alignment) {
                style.alignment = {
                    horizontal: cell.alignment.horizontal || 'left',
                    vertical: cell.alignment.vertical || 'middle',
                    wrapText: cell.alignment.wrapText || false,
                    indent: cell.alignment.indent || 0,
                    textRotation: cell.alignment.textRotation || 0
                };
            }

            // Borders - DETAILED DEBUGGING
            style.borderData = {};
            let hasBorderData = false;
            
            if (cell.border) {
                console.log(`Cell has border object:`, cell.border);
                
                const sides: ('top' | 'left' | 'bottom' | 'right')[] = ['top', 'left', 'bottom', 'right'];
                sides.forEach(side => {
                    const border = cell.border?.[side];
                    if (border && border.style) {
                        let color = '#000000';
                        if (border.color && 'argb' in border.color && border.color.argb) {
                            color = convertARGBToRGBA(border.color.argb);
                        }
                        
                        style.borderData[side] = {
                            style: border.style,
                            color: color
                        };
                        hasBorderData = true;
                        
                        console.log(`Found border on ${side}: ${border.style} ${color}`);
                    }
                });
            }
            
            style.hasBorderData = hasBorderData;
            
            if (hasBorderData) {
                console.log(`Cell has borders:`, style.borderData);
            }

        } catch (e) {
            console.warn('Error extracting cell style:', e);
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
        const maxRows = Math.max(...worksheets.map(ws => ws.data.maxRow));
        const rowHeaderWidth = Math.max(40, Math.ceil(Math.log10(maxRows + 1)) * 12 + 20);

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX Viewer</title>
    <style>
        :root {
            --header-bg: #f8f9fa;
            --header-border: #e0e0e0;
            --header-text: #5f6368;
            --cell-border: #e0e0e0;
            --selection-border: #1a73e8;
            --selection-bg: rgba(26, 115, 232, 0.1);
            --hover-bg: rgba(0, 0, 0, 0.04);
            --resize-hover: #1a73e8;
            --default-font-size: 10pt;
            --default-font-family: Arial, sans-serif;
            --default-bg-color: #ffffff;
            --default-text-color: #000000;
        }

        * {
            box-sizing: border-box;
        }
        
        body {
            font-family: var(--default-font-family);
            margin: 0;
            padding: 0;
            background-color: #ffffff;
            overflow: hidden;
            height: 100vh;
            display: flex;
            flex-direction: column;
        }

        body.dark-mode {
            background-color: #1e1e1e;
            --header-bg: #252526;
            --header-border: #3c3c3c;
            --header-text: #cccccc;
            --cell-border: #3c3c3c;
            --selection-border: #0e639c;
            --selection-bg: rgba(14, 99, 156, 0.25);
            --hover-bg: rgba(90, 93, 94, 0.31);
            --resize-hover: #007acc;
            --default-bg-color: #1e1e1e;
            --default-text-color: #d4d4d4;
        }

        .button-container {
            margin: 10px;
            display: flex;
            gap: 10px;
            flex-shrink: 0;
        }

        .toggle-button {
            max-height: 42px;
            padding: 8px 16px;
            font-size: 14px;
            font-weight: 500;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            background-color: #0e639c;
            color: white;
            transition: all 0.2s ease;
        }

        .toggle-button:hover {
            background-color: #1177bb;
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
            border-color: #0e639c !important;
        }

        .sheet-selector:hover {
            border-color: #007acc;
        }

        body.dark-mode .sheet-selector {
            background-color: #3c3c3c;
            color: #cccccc;
            border-color: #555;
        }

        .spreadsheet-container {
            flex: 1;
            overflow: hidden;
            position: relative;
            background-color: white;
            border: 1px solid var(--header-border);
            margin: 0 10px 10px 10px;
            margin-bottom: 32px;
        }

        body.dark-mode .spreadsheet-container {
            background-color: #1e1e1e;
        }

        .table-wrapper {
            width: 100%;
            height: 100%;
            overflow: auto;
            position: relative;
        }

        body.dark-mode .table-wrapper::-webkit-scrollbar {
            width: 14px;
            height: 14px;
        }

        body.dark-mode .table-wrapper::-webkit-scrollbar-track {
            background: #1e1e1e;
        }

        body.dark-mode .table-wrapper::-webkit-scrollbar-thumb {
            background: #424242;
            border: 3px solid #1e1e1e;
        }

        body.dark-mode .table-wrapper::-webkit-scrollbar-thumb:hover {
            background: #4f4f4f;
        }

        table {
            border-collapse: separate;
            border-spacing: 0;
            font-size: var(--default-font-size);
            font-family: var(--default-font-family);
            position: relative;
            background-color: var(--default-bg-color);
            user-select: none;
        }

        th, td {
            padding: 2px 4px;
            position: relative;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: cell;
        }

        /* DEFAULT BORDERS - Lower specificity */
        td {
            border-top: 1px solid var(--cell-border);
            border-left: 1px solid var(--cell-border);
        }

        td:last-child {
            border-right: 1px solid var(--cell-border);
        }

        tr:last-child td {
            border-bottom: 1px solid var(--cell-border);
        }

        th.row-header,
        th.col-header {
            background-color: var(--header-bg);
            color: var(--header-text);
            font-weight: normal;
            font-size: 11px;
            text-align: center;
            user-select: none;
            cursor: pointer;
            position: relative;
            border: 1px solid var(--header-border);
        }

        th.row-header {
            width: ${rowHeaderWidth}px;
            min-width: ${rowHeaderWidth}px;
            max-width: ${rowHeaderWidth}px;
            text-align: center;
            white-space: nowrap;
            height: 21px;
            padding: 0 4px;
        }

        th.col-header {
            height: 21px;
            min-height: 21px;
            padding: 0 4px;
        }

        .col-resize {
            position: absolute;
            right: -3px;
            top: 0;
            width: 6px;
            height: 100%;
            cursor: col-resize;
            z-index: 10;
        }

        .row-resize {
            position: absolute;
            bottom: -3px;
            left: 0;
            height: 6px;
            width: 100%;
            cursor: row-resize;
            z-index: 10;
        }

        .col-resize:hover,
        .row-resize:hover,
        .col-resize.resizing,
        .row-resize.resizing {
            background-color: var(--resize-hover);
        }

        td {
            background-color: var(--default-bg-color);
            color: var(--default-text-color);
            min-height: 21px;
            height: 21px;
            white-space: nowrap;
            font-size: var(--default-font-size);
            font-family: var(--default-font-family);
            padding: 1px 3px;
        }

        /* SIMPLIFIED DARK MODE TEXT RULE */
        body.dark-mode td {
            color: var(--default-text-color);
        }

        /* Keep custom text colors in dark mode */
        body.dark-mode td[data-has-custom-text="true"] {
            /* Custom color will be applied via inline style */
        }

        /* Keep dark text on light backgrounds in dark mode */
        body.dark-mode td[data-light-bg="true"] {
            color: #000000 !important;
        }

        td.hidden-cell {
            display: none;
        }

        td[colspan], td[rowspan] {
            z-index: 1;
        }

        td[data-align-h="center"] .cell-content { 
            text-align: center; 
            display: block;
        }
        td[data-align-h="right"] .cell-content { 
            text-align: right; 
            display: block;
        }
        td[data-align-h="justify"] .cell-content { 
            text-align: justify; 
            display: block;
        }

        td[data-align-v="top"] {
            vertical-align: top;
        }

        td[data-align-v="middle"] {
            vertical-align: middle;
        }

        td[data-align-v="bottom"] {
            vertical-align: bottom;
        }

        td.has-vertical-align {
            height: 100%;
            padding: 0;
        }

        td.has-vertical-align .cell-content-wrapper {
            display: flex;
            height: 100%;
            padding: 1px 3px;
        }

        td[data-align-v="top"] .cell-content-wrapper {
            align-items: flex-start;
        }

        td[data-align-v="middle"] .cell-content-wrapper {
            align-items: center;
        }

        td[data-align-v="bottom"] .cell-content-wrapper {
            align-items: flex-end;
        }

        td[data-wrap="true"] .cell-content {
            white-space: pre-wrap;
            word-wrap: break-word;
        }

        td.selected {
            background-color: var(--selection-bg) !important;
            outline: 1px solid var(--selection-border) !important;
            outline-offset: -1px;
        }

        td.active-cell {
            outline: 2px solid var(--selection-border) !important;
            outline-offset: -2px;
            z-index: 3;
        }

        td.column-selected, th.column-selected {
            background-color: var(--selection-bg) !important;
        }

        td.row-selected, th.row-selected {
            background-color: var(--selection-bg) !important;
        }

        th.row-header:hover,
        th.col-header:hover {
            background-color: var(--hover-bg);
        }

        th.corner-cell {
            background-color: var(--header-bg);
            cursor: pointer;
            width: ${rowHeaderWidth}px;
            min-width: ${rowHeaderWidth}px;
            max-width: ${rowHeaderWidth}px;
            border: 1px solid var(--header-border);
        }

        th.corner-cell:hover {
            background-color: var(--hover-bg);
        }

        .cell-content {
            display: inline-block;
            width: 100%;
            overflow: hidden;
            text-overflow: ellipsis;
            user-select: none;
            pointer-events: none;
        }

        .cell-content-wrapper {
            width: 100%;
        }

        @keyframes copyFlash {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }

        td.copying {
            animation: copyFlash 0.3s ease-in-out;
        }

        .selection-info {
            position: fixed;
            bottom: 30px;
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
            background: rgba(0, 0, 0, 0.9);
            border: 1px solid #3c3c3c;
        }

        table:focus {
            outline: 2px solid var(--selection-border);
            outline-offset: -2px;
        }

        td:focus {
            outline: 2px solid var(--selection-border);
            outline-offset: -2px;
        }

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
            background-color: rgba(30, 30, 30, 0.9);
        }

        .loading-overlay.hidden {
            display: none;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #0e639c;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin-bottom: 15px;
        }

        body.dark-mode .spinner {
            border-color: #3c3c3c;
            border-top-color: #007acc;
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

        .resize-line {
            position: absolute;
            background-color: var(--resize-hover);
            z-index: 1000;
            display: none;
        }

        .resize-line.vertical {
            width: 2px;
            height: 100%;
            top: 0;
        }

        .resize-line.horizontal {
            height: 2px;
            width: 100%;
            left: 0;
        }

        .skip-link {
            position: absolute;
            top: -40px;
            left: 0;
            background: #000;
            color: #fff;
            padding: 8px;
            text-decoration: none;
            border-radius: 0 0 4px 0;
        }

        .skip-link:focus {
            top: 0;
        }

        .status-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            height: 22px;
            background-color: var(--header-bg);
            border-top: 1px solid var(--header-border);
            display: flex;
            align-items: center;
            padding: 0 10px;
            font-size: 11px;
            color: var(--header-text);
            z-index: 100;
        }

        .status-bar-item {
            margin-right: 20px;
        }

        /* DEBUG: Simple test border to see if CSS is working */
        .test-border {
            border: 3px solid red !important;
        }
    </style>
</head>
<body>
    <a href="#spreadsheet" class="skip-link">Skip to spreadsheet content</a>
    
    <div class="loading-overlay" id="loadingOverlay">
        <div class="spinner"></div>
        <div class="loading-text" role="status" aria-live="polite">Rendering worksheet...</div>
    </div>

    <div class="button-container">
        <select id="sheetSelector" class="sheet-selector" aria-label="Select worksheet">
            ${worksheets.map((ws, i) => `<option value="${i}">${ws.name}</option>`).join('')}
        </select>
        <button id="toggleBackgroundButton" class="toggle-button" title="Toggle Light/Dark Mode" aria-label="Toggle between light and dark mode">
            <svg id="lightIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: none;" aria-hidden="true">
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
            <svg id="darkIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                <path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>
            </svg>
        </button>
        <button id="toggleMinWidthButton" class="toggle-button" title="Toggle table width" aria-label="Toggle table width">
            <svg fill="#ffffff" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" stroke="#ffffff" aria-hidden="true">
                <g id="SVGRepo_bgCarrier" stroke-width="0"></g>
                <g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g>
                <g id="SVGRepo_iconCarrier">
                    <path d="M19.5,21 C20.3284271,21 21,20.3284271 21,19.5 L21,11.5 C21,10.6715729 20.3284271,10 19.5,10 L11.5,10 C10.6715729,10 10,10.6715729 10,11.5 L10,19.5 C10,20.3284271 10.6715729,21 11.5,21 L19.5,21 Z M5,20.2928932 L6.14644661,19.1464466 C6.34170876,18.9511845 6.65829124,18.9511845 6.85355339,19.1464466 C7.04881554,19.3417088 7.04881554,19.6582912 6.85355339,19.8535534 L4.85355339,21.8535534 C4.65829124,22.0488155 4.34170876,22.0488155 4.14644661,21.8535534 L2.14644661,19.8535534 C1.95118446,19.6582912 1.95118446,19.3417088 2.14644661,19.1464466 C2.34170876,18.9511845 2.65829124,18.9511845 2.85355339,19.1464466 L4,20.2928932 L4,7.5 C4,7.22385763 4.22385763,7 4.5,7 C4.77614237,7 5,7.22385763 5,7.5 L5,20.2928932 L5,20.2928932 Z M20.2928932,4 L19.1464466,2.85355339 C18.9511845,2.65829124 18.9511845,2.34170876 19.1464466,2.14644661 C19.3417088,1.95118446 19.6582912,1.95118446 19.8535534,2.14644661 L21.8535534,4.14644661 C22.0488155,4.34170876 22.0488155,4.65829124 21.8535534,4.85355339 L19.8535534,6.85355339 C19.6582912,7.04881554 19.3417088,7.04881554 19.1464466,6.85355339 C18.9511845,6.65829124 18.9511845,6.34170876 19.1464466,6.14644661 L20.2928932,5 L7.5,5 C7.22385763,5 7,4.77614237 7,4.5 C7,4.22385763 7.22385763,4 7.5,4 L20.2928932,4 Z M19.5,22 L11.5,22 C10.1192881,22 9,20.8807119 9,19.5 L9,11.5 C9,10.1192881 10.1192881,9 11.5,9 L19.5,9 C20.8807119,9 22,10.1192881 22,11.5 L22,19.5 C22,20.8807119 20.8807119,22 19.5,22 Z"></path>
                </g>
            </svg>
            &nbsp; Default
        </button>
        <button id="testBorderButton" class="toggle-button" title="Test Border">Test Border</button>
    </div>
    
    <div class="spreadsheet-container" id="spreadsheet">
        <div class="table-wrapper" id="tableWrapper">
            <!-- Table will be inserted here -->
        </div>
        <div class="resize-line vertical" id="colResizeLine"></div>
        <div class="resize-line horizontal" id="rowResizeLine"></div>
    </div>
    
    <div class="selection-info" id="selectionInfo" role="status" aria-live="polite"></div>
    
    <div class="status-bar">
        <div class="status-bar-item" id="cellPosition" aria-live="polite"></div>
        <div class="status-bar-item" id="selectionStatus" aria-live="polite"></div>
    </div>

    <script>
        console.log('Starting spreadsheet viewer...');
        
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
        let isResizing = false;
        let resizeTarget = null;
        let resizeStartPos = 0;
        let resizeStartSize = 0;
        
        console.log('Worksheets data loaded:', worksheetsData);
        
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
            try {
                console.log('Creating table with data:', worksheetData);
                
                const data = worksheetData.data;
                const hiddenCells = new Set(data.hiddenCells || []);
                let html = '<table role="grid" aria-label="Spreadsheet data">';
                
                // Header row
                html += '<thead><tr role="row">';
                html += '<th class="corner-cell" role="columnheader" aria-label="Select all cells"></th>';
                for (let c = 1; c <= data.maxCol; c++) {
                    const width = data.columnWidths[c-1] || 64;
                    html += '<th class="col-header" data-col="' + (c-1) + '" role="columnheader" style="width: ' + width + 'px; min-width: ' + width + 'px; max-width: ' + width + 'px;" aria-label="Column ' + getExcelColumnLabel(c) + '">';
                    html += getExcelColumnLabel(c);
                    html += '<div class="col-resize" data-col="' + (c-1) + '" aria-hidden="true"></div>';
                    html += '</th>';
                }
                html += '</tr></thead><tbody>';
                
                // Data rows
                data.rows.forEach((row, rowIndex) => {
                    const height = row.height || 21;
                    html += '<tr role="row" style="height: ' + height + 'px;">';
                    html += '<th class="row-header" data-row="' + rowIndex + '" role="rowheader" aria-label="Row ' + row.rowNumber + '">';
                    html += row.rowNumber;
                    html += '<div class="row-resize" data-row="' + rowIndex + '" aria-hidden="true"></div>';
                    html += '</th>';
                    
                    row.cells.forEach((cell, colIndex) => {
                        const cellKey = rowIndex + '_' + colIndex;
                        if (hiddenCells.has(cellKey)) {
                            return;
                        }
                        
                        const styleStr = formatCellStyle(cell.style, cell.value);
                        const hasCustomText = cell.style.hasCustomTextColor === true;
                        const isLightBg = cell.style.isLightBackground === true;
                        const colLabel = getExcelColumnLabel(colIndex + 1);
                        const hasVerticalAlign = cell.style.alignment && cell.style.alignment.vertical && cell.style.alignment.vertical !== 'middle';
                        
                        console.log('Cell (' + rowIndex + ',' + colIndex + ') - hasCustomText:', hasCustomText, 'isLightBg:', isLightBg, 'style:', cell.style);
                        
                        html += '<td role="gridcell"';
                        html += ' data-row="' + rowIndex + '"';
                        html += ' data-col="' + colIndex + '"';
                        html += ' aria-label="Cell ' + colLabel + row.rowNumber + (cell.value ? ', ' + cell.value : ', empty') + '"';
                        html += ' tabindex="-1"';
                        
                        if (hasCustomText) html += ' data-has-custom-text="true"';
                        if (isLightBg) html += ' data-light-bg="true"';
                        
                        if (cell.style.alignment) {
                            if (cell.style.alignment.horizontal && cell.style.alignment.horizontal !== 'left') {
                                html += ' data-align-h="' + cell.style.alignment.horizontal + '"';
                            }
                            if (cell.style.alignment.vertical && cell.style.alignment.vertical !== 'middle') {
                                html += ' data-align-v="' + cell.style.alignment.vertical + '"';
                            }
                            if (cell.style.alignment.wrapText) {
                                html += ' data-wrap="true"';
                            }
                        }
                        
                        if (cell.rowSpan && cell.rowSpan > 1) {
                            html += ' rowspan="' + cell.rowSpan + '"';
                        }
                        if (cell.colSpan && cell.colSpan > 1) {
                            html += ' colspan="' + cell.colSpan + '"';
                        }
                        
                        if (hasVerticalAlign) {
                            html += ' class="has-vertical-align"';
                        }
                        
                        html += ' style="' + styleStr + '"';
                        html += '>';
                        
                        if (hasVerticalAlign) {
                            html += '<div class="cell-content-wrapper">';
                        }
                        html += '<span class="cell-content">' + (cell.value || '') + '</span>';
                        if (hasVerticalAlign) {
                            html += '</div>';
                        }
                        
                        html += '</td>';
                    });
                    
                    html += '</tr>';
                });
                
                html += '</tbody></table>';
                console.log('Table HTML created successfully');
                return html;
            } catch (error) {
                console.error('Error creating table:', error);
                return '<div style="padding: 20px; color: red;">Error creating table: ' + error.message + '</div>';
            }
        }

        function formatCellStyle(style, value) {
            let css = '';
            
            try {
                if (style.backgroundColor) {
                    css += 'background-color: ' + style.backgroundColor + ' !important;';
                }
                
                if (style.color) css += 'color: ' + style.color + ' !important;';
                if (style.fontWeight) css += 'font-weight: ' + style.fontWeight + ';';
                if (style.fontStyle) css += 'font-style: ' + style.fontStyle + ';';
                if (style.textDecoration) css += 'text-decoration: ' + style.textDecoration + ';';
                if (style.fontSize) {
                    const size = typeof style.fontSize === 'number' ? style.fontSize : parseInt(style.fontSize);
                    css += 'font-size: ' + size + 'pt;';
                }
                if (style.fontFamily) css += 'font-family: "' + style.fontFamily + '", var(--default-font-family);';
                
                if (style.alignment) {
                    if (style.alignment.indent && style.alignment.indent > 0) {
                        css += 'padding-left: ' + (3 + style.alignment.indent * 10) + 'px;';
                    }
                }
                
                // DETAILED BORDER DEBUGGING
                if (style.hasBorderData) {
                    console.log('PROCESSING CELL WITH BORDERS:', style.borderData);
                    
                    // Remove ALL default borders
                    css += 'border: none !important;';
                    
                    const sides = ['top', 'right', 'bottom', 'left'];
                    sides.forEach(side => {
                        if (style.borderData[side]) {
                            const border = style.borderData[side];
                            let width = '1px';
                            let borderStyle = 'solid';
                            
                            switch (border.style) {
                                case 'thin': width = '1px'; break;
                                case 'medium': width = '2px'; break;
                                case 'thick': width = '3px'; break;
                                case 'hair': width = '0.5px'; break;
                                case 'dotted': borderStyle = 'dotted'; break;
                                case 'dashed': borderStyle = 'dashed'; break;
                                case 'double': borderStyle = 'double'; width = '3px'; break;
                                case 'dashDot': borderStyle = 'dashed'; break;
                                case 'dashDotDot': borderStyle = 'dashed'; break;
                                case 'slantDashDot': borderStyle = 'dashed'; break;
                                case 'mediumDashed': width = '2px'; borderStyle = 'dashed'; break;
                                case 'mediumDashDot': width = '2px'; borderStyle = 'dashed'; break;
                                case 'mediumDashDotDot': width = '2px'; borderStyle = 'dashed'; break;
                                case 'mediumDotted': width = '2px'; borderStyle = 'dotted'; break;
                                case 'thickDashed': width = '3px'; borderStyle = 'dashed'; break;
                                case 'thickDotted': width = '3px'; borderStyle = 'dotted'; break;
                                default: borderStyle = 'solid'; break;
                            }
                            
                            const borderCSS = width + ' ' + borderStyle + ' ' + border.color;
                            css += 'border-' + side + ': ' + borderCSS + ' !important;';
                            console.log('APPLIED border-' + side + ':', borderCSS);
                        }
                    });
                    
                    console.log('FINAL CSS WITH BORDERS:', css);
                }
            } catch (error) {
                console.error('Error formatting cell style:', error);
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
            console.log('Rendering worksheet:', index);
            showLoading();
            
            setTimeout(() => {
                try {
                    const wrapper = document.getElementById('tableWrapper');
                    wrapper.innerHTML = createTable(worksheetsData[index]);
                    initializeSelection();
                    initializeResize();
                    hideLoading();
                    
                    const firstCell = wrapper.querySelector('td');
                    if (firstCell) {
                        firstCell.tabIndex = 0;
                    }
                    
                    console.log('Worksheet rendered successfully');
                } catch (error) {
                    console.error('Error rendering worksheet:', error);
                    hideLoading();
                    document.getElementById('tableWrapper').innerHTML = 
                        '<div style="padding: 20px; color: red;">Error rendering worksheet: ' + error.message + '</div>';
                }
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
            updateStatusBar();
        }

        function selectCell(cell, isMulti = false) {
            if (!isMulti) {
                clearSelection();
            }
            
            if (activeCell) {
                activeCell.tabIndex = -1;
            }
            cell.tabIndex = 0;
            
            cell.classList.add('selected');
            cell.classList.add('active-cell');
            selectedCells.add(cell);
            activeCell = cell;
            updateSelectionInfo();
            updateStatusBar();
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
            
            const startCell = document.querySelector('td[data-row="' + startRow + '"][data-col="' + startCol + '"]');
            if (startCell) {
                if (activeCell) {
                    activeCell.tabIndex = -1;
                }
                startCell.tabIndex = 0;
                startCell.classList.add('active-cell');
                activeCell = startCell;
            }
            
            updateSelectionInfo();
            updateStatusBar();
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
            updateStatusBar();
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
            updateStatusBar();
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

        function updateStatusBar() {
            const cellPos = document.getElementById('cellPosition');
            const selStatus = document.getElementById('selectionStatus');
            
            if (activeCell) {
                const row = parseInt(activeCell.dataset.row) + 1;
                const col = parseInt(activeCell.dataset.col) + 1;
                const colLabel = getExcelColumnLabel(col);
                cellPos.textContent = colLabel + row;
            } else {
                cellPos.textContent = '';
            }
            
            if (selectedCells.size > 1) {
                selStatus.textContent = 'Selected: ' + selectedCells.size + ' cells';
            } else {
                selStatus.textContent = '';
            }
        }

        function copySelection() {
            if (selectedCells.size === 0) return;
            
            const cellsArray = Array.from(selectedCells);
            const cellData = cellsArray.map(cell => ({
                row: parseInt(cell.dataset.row),
                col: parseInt(cell.dataset.col),
                text: cell.textContent.trim()
            }));
            
            cellData.sort((a, b) => a.row - b.row || a.col - b.col);
            
            const rows = [...new Set(cellData.map(c => c.row))].sort((a, b) => a - b);
            const cols = [...new Set(cellData.map(c => c.col))].sort((a, b) => a - b);
            const minRow = Math.min(...rows);
            const maxRow = Math.max(...rows);
            const minCol = Math.min(...cols);
            const maxCol = Math.max(...cols);
            
            const grid = [];
            for (let r = minRow; r <= maxRow; r++) {
                const row = [];
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = cellData.find(cd => cd.row === r && cd.col === c);
                    row.push(cell ? cell.text : '');
                }
                grid.push(row);
            }
            
            const text = grid.map(row => row.join('\\t')).join('\\n');
            
            navigator.clipboard.writeText(text).then(() => {
                selectedCells.forEach(cell => {
                    cell.classList.add('copying');
                    setTimeout(() => cell.classList.remove('copying'), 300);
                });
            }).catch(err => {
                console.error('Failed to copy:', err);
            });
        }

        function initializeResize() {
            const colResizeLine = document.getElementById('colResizeLine');
            const rowResizeLine = document.getElementById('rowResizeLine');
            
            document.querySelectorAll('.col-resize').forEach(handle => {
                handle.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    isResizing = true;
                    resizeTarget = handle.parentElement;
                    resizeStartPos = e.pageX;
                    resizeStartSize = resizeTarget.offsetWidth;
                    
                    handle.classList.add('resizing');
                    colResizeLine.style.display = 'block';
                    colResizeLine.style.left = e.pageX + 'px';
                    
                    document.body.style.cursor = 'col-resize';
                });
            });
            
            document.querySelectorAll('.row-resize').forEach(handle => {
                handle.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    isResizing = true;
                    resizeTarget = handle.parentElement.parentElement;
                    resizeStartPos = e.pageY;
                    resizeStartSize = resizeTarget.offsetHeight;
                    
                    handle.classList.add('resizing');
                    rowResizeLine.style.display = 'block';
                    rowResizeLine.style.top = e.pageY + 'px';
                    
                    document.body.style.cursor = 'row-resize';
                });
            });
            
            document.addEventListener('mousemove', (e) => {
                if (!isResizing || !resizeTarget) return;
                
                if (colResizeLine.style.display === 'block') {
                    colResizeLine.style.left = e.pageX + 'px';
                } else if (rowResizeLine.style.display === 'block') {
                    rowResizeLine.style.top = e.pageY + 'px';
                }
            });
            
            document.addEventListener('mouseup', (e) => {
                if (!isResizing || !resizeTarget) return;
                
                if (colResizeLine.style.display === 'block') {
                    const diff = e.pageX - resizeStartPos;
                    const newWidth = Math.max(30, resizeStartSize + diff);
                    const colIndex = resizeTarget.dataset.col;
                    
                    document.querySelectorAll('th[data-col="' + colIndex + '"], td[data-col="' + colIndex + '"]').forEach(cell => {
                        cell.style.width = newWidth + 'px';
                        cell.style.minWidth = newWidth + 'px';
                        cell.style.maxWidth = newWidth + 'px';
                    });
                    
                    colResizeLine.style.display = 'none';
                } else if (rowResizeLine.style.display === 'block') {
                    const diff = e.pageY - resizeStartPos;
                    const newHeight = Math.max(21, resizeStartSize + diff);
                    
                    resizeTarget.style.height = newHeight + 'px';
                    rowResizeLine.style.display = 'none';
                }
                
                document.querySelectorAll('.resizing').forEach(el => el.classList.remove('resizing'));
                document.body.style.cursor = '';
                isResizing = false;
                resizeTarget = null;
            });
        }

        function initializeSelection() {
            const tableWrapper = document.getElementById('tableWrapper');
            const table = tableWrapper.querySelector('table');
            
            if (!table) {
                console.error('Table not found');
                return;
            }
            
            table.addEventListener('selectstart', (e) => {
                if (!isResizing) {
                    e.preventDefault();
                    return false;
                }
            });
            
            table.addEventListener('mousedown', (e) => {
                if (isResizing) return;
                
                const target = e.target.closest('td, th');
                if (!target) return;
                
                e.preventDefault();
                
                if (target.classList.contains('col-header')) {
                    const colIndex = parseInt(target.dataset.col);
                    
                    if (!e.shiftKey) {
                        lastSelectedColumn = colIndex;
                    }
                    
                    selectColumn(colIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                    return;
                }
                
                if (target.classList.contains('row-header')) {
                    const rowIndex = parseInt(target.dataset.row);
                    
                    if (!e.shiftKey) {
                        lastSelectedRow = rowIndex;
                    }
                    
                    selectRow(rowIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                    return;
                }
                
                if (target.classList.contains('corner-cell')) {
                    clearSelection();
                    const allCells = table.querySelectorAll('td');
                    allCells.forEach(cell => {
                        cell.classList.add('selected');
                        selectedCells.add(cell);
                    });
                    if (allCells.length > 0) {
                        if (activeCell) {
                            activeCell.tabIndex = -1;
                        }
                        allCells[0].tabIndex = 0;
                        allCells[0].classList.add('active-cell');
                        activeCell = allCells[0];
                    }
                    updateSelectionInfo();
                    updateStatusBar();
                    return;
                }
                
                if (target.tagName === 'TD') {
                    const row = parseInt(target.dataset.row);
                    const col = parseInt(target.dataset.col);
                    
                    if (e.ctrlKey || e.metaKey) {
                        if (target.classList.contains('selected')) {
                            target.classList.remove('selected');
                            selectedCells.delete(target);
                            if (target === activeCell) {
                                target.classList.remove('active-cell');
                                target.tabIndex = -1;
                                activeCell = null;
                                
                                const remainingSelected = document.querySelector('td.selected');
                                if (remainingSelected) {
                                    remainingSelected.classList.add('active-cell');
                                    remainingSelected.tabIndex = 0;
                                    activeCell = remainingSelected;
                                }
                            }
                        } else {
                            target.classList.add('selected');
                            selectedCells.add(target);
                            if (activeCell) {
                                activeCell.classList.remove('active-cell');
                                activeCell.tabIndex = -1;
                            }
                            target.classList.add('active-cell');
                            target.tabIndex = 0;
                            activeCell = target;
                        }
                        updateSelectionInfo();
                        updateStatusBar();
                    } else if (e.shiftKey && activeCell) {
                        const startRow = parseInt(activeCell.dataset.row);
                        const startCol = parseInt(activeCell.dataset.col);
                        selectRange(startRow, startCol, row, col);
                    } else {
                        isSelecting = true;
                        selectionStart = { row, col };
                        selectCell(target);
                    }
                }
            });
            
            table.addEventListener('mousemove', (e) => {
                if (!isSelecting || !selectionStart || isResizing) return;
                
                const target = e.target.closest('td');
                if (!target) return;
                
                const row = parseInt(target.dataset.row);
                const col = parseInt(target.dataset.col);
                
                if (!selectionEnd || selectionEnd.row !== row || selectionEnd.col !== col) {
                    selectionEnd = { row, col };
                    selectRange(selectionStart.row, selectionStart.col, row, col);
                }
            });
            
            document.addEventListener('mouseup', () => {
                if (!isResizing) {
                    isSelecting = false;
                    selectionStart = null;
                    selectionEnd = null;
                }
            });
            
            table.addEventListener('keydown', (e) => {
                if (!activeCell) return;
                
                const row = parseInt(activeCell.dataset.row);
                const col = parseInt(activeCell.dataset.col);
                let newRow = row;
                let newCol = col;
                
                switch (e.key) {
                    case 'ArrowUp':
                        e.preventDefault();
                        newRow = Math.max(0, row - 1);
                        break;
                    case 'ArrowDown':
                        e.preventDefault();
                        newRow = row + 1;
                        break;
                    case 'ArrowLeft':
                        e.preventDefault();
                        newCol = Math.max(0, col - 1);
                        break;
                    case 'ArrowRight':
                        e.preventDefault();
                        newCol = col + 1;
                        break;
                    case 'Tab':
                        e.preventDefault();
                        if (e.shiftKey) {
                            newCol = Math.max(0, col - 1);
                        } else {
                            newCol = col + 1;
                        }
                        break;
                    case 'Enter':
                        e.preventDefault();
                        if (e.shiftKey) {
                            newRow = Math.max(0, row - 1);
                        } else {
                            newRow = row + 1;
                        }
                        break;
                    case 'Home':
                        e.preventDefault();
                        if (e.ctrlKey) {
                            newRow = 0;
                            newCol = 0;
                        } else {
                            newCol = 0;
                        }
                        break;
                    case 'End':
                        e.preventDefault();
                        if (e.ctrlKey) {
                            const lastRow = document.querySelectorAll('tr').length - 2;
                            const lastCol = document.querySelectorAll('th.col-header').length - 1;
                            newRow = lastRow;
                            newCol = lastCol;
                        } else {
                            const lastCol = document.querySelectorAll('th.col-header').length - 1;
                            newCol = lastCol;
                        }
                        break;
                }
                
                const newCell = document.querySelector('td[data-row="' + newRow + '"][data-col="' + newCol + '"]');
                if (newCell) {
                    if (e.shiftKey && (e.key.startsWith('Arrow') || e.key === 'Tab')) {
                        const startRow = parseInt(selectionStart?.row ?? row);
                        const startCol = parseInt(selectionStart?.col ?? col);
                        selectRange(startRow, startCol, newRow, newCol);
                    } else {
                        selectCell(newCell);
                        const rect = newCell.getBoundingClientRect();
                        const wrapperRect = tableWrapper.getBoundingClientRect();
                        
                        if (rect.top < wrapperRect.top || rect.bottom > wrapperRect.bottom ||
                            rect.left < wrapperRect.left || rect.right > wrapperRect.right) {
                            newCell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                        }
                    }
                }
            });
            
            document.addEventListener('keydown', (e) => {
                if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                    e.preventDefault();
                    copySelection();
                }
                
                if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
                    e.preventDefault();
                    const allCells = table.querySelectorAll('td');
                    clearSelection();
                    allCells.forEach(cell => {
                        cell.classList.add('selected');
                        selectedCells.add(cell);
                    });
                    if (allCells.length > 0) {
                        if (activeCell) {
                            activeCell.tabIndex = -1;
                        }
                        allCells[0].tabIndex = 0;
                        allCells[0].classList.add('active-cell');
                        activeCell = allCells[0];
                    }
                    updateSelectionInfo();
                    updateStatusBar();
                }
            });
            
            document.addEventListener('click', (e) => {
                if (!e.target.closest('table') && !e.target.closest('.button-container')) {
                    clearSelection();
                }
            });
        }

        document.addEventListener('DOMContentLoaded', () => {
            console.log('DOM loaded, initializing...');
            
            renderWorksheet(0);
            
            document.getElementById('sheetSelector').addEventListener('change', (e) => {
                currentWorksheet = parseInt(e.target.value);
                clearSelection();
                renderWorksheet(currentWorksheet);
            });
            
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
            });
            
            let minWidthState = 2;
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
            
            // DEBUG: Test border button
            document.getElementById('testBorderButton').addEventListener('click', () => {
                const firstCell = document.querySelector('td');
                if (firstCell) {
                    firstCell.classList.toggle('test-border');
                    console.log('Toggled test border on first cell');
                }
            });
        });
    </script>
</body>
</html>`;
    }
}