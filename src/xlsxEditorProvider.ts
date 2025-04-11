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
            await workbook.xlsx.readFile(document.uri.fsPath);

            let sheetOptionsHtml = '';
            workbook.worksheets.forEach((sheet, index) => {
                sheetOptionsHtml += `<option value="${index}">${sheet.name}</option>`;
            });

            const getExcelColLabel = (n: number): string => {
                let label = '';
                while (n > 0) {
                    let rem = (n - 1) % 26;
                    label = String.fromCharCode(65 + rem) + label;
                    n = Math.floor((n - 1) / 26);
                }
                return label;
            };

            const getBorderStyle = (border?: Partial<Excel.Border>): { styleString: string; isBlackOrShade: boolean } => {
                const defaultBorderStyle = '1px solid #c4c4c4';
                const defaultResult = { styleString: defaultBorderStyle, isBlackOrShade: false };

                if (!border || !border.style) return defaultResult;

                const originalColor = border.color?.argb
                    ? convertARGBToRGBA(border.color.argb)
                    : 'rgba(0, 0, 0, 1)';
                const isBlack = isShadeOfBlack(originalColor);
                const displayColor = originalColor;

                const borderStyles: Record<string, string> = {
                    thin: `1px solid ${displayColor}`,
                    medium: `2px solid ${displayColor}`,
                    thick: `3px solid ${displayColor}`,
                    dotted: `1px dotted ${displayColor}`,
                    dashed: `1px dashed ${displayColor}`,
                    double: `3px double ${displayColor}`
                };

                const stylePart = borderStyles[border.style] || `1px solid ${displayColor}`;
                return { styleString: stylePart, isBlackOrShade: isBlack };
            };

            const generateTableHtml = (worksheet: Excel.Worksheet): string => {
                const isCellEmpty = (cell: Excel.Cell): boolean => {
                    if (cell.value !== null && cell.value !== undefined && cell.value !== '') return false;
                    if (cell.fill && (cell.fill as any).fgColor && (cell.fill as any).fgColor.argb) return false;
                    if (cell.border) {
                        const { top, bottom, left, right } = cell.border;
                        if (top || bottom || left || right) return false;
                    }
                    return true;
                };

                const isRowEmpty = (row: Excel.Row): boolean => {
                    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                        if (!isCellEmpty(row.getCell(colNumber))) return false;
                    }
                    return true;
                };

                let lastNonEmptyRow = worksheet.rowCount;
                while (lastNonEmptyRow > 0 && isRowEmpty(worksheet.getRow(lastNonEmptyRow))) {
                    lastNonEmptyRow--;
                }
                if (lastNonEmptyRow === 0) lastNonEmptyRow = 1;

                let tableHtml = '<table id="xlsx-table" border="1" cellspacing="0" cellpadding="5"><thead><tr><th class="row-header">&nbsp;</th>';
                for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                    const colLabel = getExcelColLabel(colNumber);
                    tableHtml += `<th class="col-header">${colLabel}</th>`;
                }
                tableHtml += '</tr></thead><tbody>';

                for (let rowNumber = 1; rowNumber <= lastNonEmptyRow; rowNumber++) {
                    const row = worksheet.getRow(rowNumber);
                    tableHtml += `<tr><th class="row-header">${rowNumber}</th>`;

                    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                        let cellValue = '&nbsp;';
                        let style = '';
                        let isDefaultBlack = false;
                        let hasCustomBackground = false;
                        let cellHasBlackBorder = false;

                        const cell = row.getCell(colNumber);

                        if (cell?.fill && (cell.fill as any).fgColor && (cell.fill as any).fgColor.argb) {
                            style += `background-color:${convertARGBToRGBA((cell.fill as any).fgColor.argb)};`;
                            hasCustomBackground = true;
                        }

                        if (cell && cell.value !== null && cell.value !== undefined) {
                            cellValue = cell.value.toString();

                            if (cell.font) {
                                style += cell.font.bold ? 'font-weight:bold;' : '';
                                style += cell.font.italic ? 'font-style:italic;' : '';
                                style += cell.font.strike ? 'text-decoration:line-through;' : '';
                                if (cell.font.size) style += `font-size:${cell.font.size + 2}px;`;
                                if (cell.font.name) style += `font-family:${cell.font.name};`;
                                if (cell.font.color && typeof cell.font.color.argb === 'string') {
                                    style += `color: ${convertARGBToRGBA(cell.font.color.argb)};`;
                                } else {
                                    style += 'color: rgb(0, 0, 0);';
                                    isDefaultBlack = true;
                                }
                            } else {
                                style += 'color: rgb(0, 0, 0);';
                                isDefaultBlack = true;
                            }
                        } else {
                            style += 'color: rgb(0, 0, 0);';
                            isDefaultBlack = true;
                        }

                        const topBorder = getBorderStyle(cell?.border?.top);
                        const leftBorder = getBorderStyle(cell?.border?.left);
                        const bottomBorder = getBorderStyle(cell?.border?.bottom);
                        const rightBorder = getBorderStyle(cell?.border?.right);

                        style += `border-top:${topBorder.styleString};`;
                        style += `border-left:${leftBorder.styleString};`;
                        style += `border-bottom:${bottomBorder.styleString};`;
                        style += `border-right:${rightBorder.styleString};`;

                        if (topBorder.isBlackOrShade || leftBorder.isBlackOrShade || bottomBorder.isBlackOrShade || rightBorder.isBlackOrShade) {
                            cellHasBlackBorder = true;
                        }

                        const dataAttrs = [];
                        if (isDefaultBlack) dataAttrs.push('data-default-color="true"');
                        if (!hasCustomBackground) dataAttrs.push('data-default-bg="true"');
                        if (cellHasBlackBorder) dataAttrs.push('data-black-border="true"');
                        if (!cell || (cell.value === null && !hasCustomBackground)) dataAttrs.push('data-empty="true"');

                        tableHtml += `<td ${dataAttrs.join(' ')} style="${style}">${cellValue}</td>`;
                    }
                    tableHtml += '</tr>';
                }
                tableHtml += '</tbody></table>';
                return tableHtml;
            };

            const workbookState = {
                worksheets: workbook.worksheets.map(ws => ({
                    name: ws.name,
                    tableHtml: generateTableHtml(ws)
                }))
            };

            webviewPanel.webview.options = { enableScripts: true };
            webviewPanel.webview.html = this.getWebviewContent(sheetOptionsHtml, workbookState);
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading XLSX file: ${error}`);
        }
    }

    private getWebviewContent(sheetOptionsHtml: string, workbookState: any): string {
        return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>XLSX Viewer</title>
<style>
    body { font-family: sans-serif; padding: 10px; background-color: white; margin: 0; overflow-x: auto; }
    .table-container { width: 100%; overflow-x: auto; }
    table { 
        border-collapse: collapse; 
        width: auto; 
        table-layout: fixed;
    }
    th, td { 
        border: none; 
        padding: 8px; 
        text-align: left; 
        white-space: nowrap; 
    }
    td:nth-child(1), th:nth-child(1) {
        width: 20px !important;
    }
    td { background-color: white; }
    body.alt-bg { background-color: black; }
    .alt-bg td[data-default-bg="true"] { background-color: black; }
    .button-container { margin-bottom: 10px; display: flex; gap: 10px; position: sticky; top: 0; background-color: inherit; z-index: 1; }
    .toggle-button { max-height: 42px; padding: 8px; font-size: 14px; border: none; border-radius: 4px; cursor: pointer; display: flex; align-items: center; justify-content: center; background-color: #2196f3; color: white; transition: all 0.2s ease; }
    .toggle-button:hover { background-color: #1976d2; }
    .toggle-button:active { background-color: #1565c0; }
    .toggle-button svg { width: 20px; height: 20px; stroke: white; }
    .sheet-selector { padding: 8px 16px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; background-color: white; cursor: pointer; }
    .sheet-selector:focus { outline: none; border-color: #2196f3 !important; }
    .sheet-selector:hover { border-color: #1976d2; }
    th.col-header, th.row-header { font-weight: bold; text-align: center; background-color: #f1f1f1; color: #000; border: 1px solid #ccc; }
    body.alt-bg th.col-header, body.alt-bg th.row-header { background-color: #222; color: #fff; border-color: #444; }
</style>
</head>
<body>
<div class="button-container">
<select id="sheetSelector" class="sheet-selector">${sheetOptionsHtml}</select>
<button id="toggleButton" class="toggle-button">
<svg id="lightIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: none;">
<circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
</svg>
<svg id="darkIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
<path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>
</svg>
</button>
</div>
<div class="table-container"><div id="table-content"></div></div>
<script>
const workbookState = ${JSON.stringify(workbookState)};
const sheetSelector = document.getElementById('sheetSelector');
const toggleButton = document.getElementById('toggleButton');
const body = document.body;
const tableContent = document.getElementById('table-content');
const lightIcon = document.getElementById('lightIcon');
const darkIcon = document.getElementById('darkIcon');

const updateTable = (sheetIndex) => {
    tableContent.innerHTML = workbookState.worksheets[sheetIndex].tableHtml;
};
updateTable(0);
sheetSelector.addEventListener('change', (e) => {
    updateTable(parseInt(e.target.value));
});
toggleButton.addEventListener('click', () => {
    body.classList.toggle('alt-bg');
    const isDarkMode = body.classList.contains('alt-bg');
    lightIcon.style.display = isDarkMode ? 'block' : 'none';
    darkIcon.style.display = isDarkMode ? 'none' : 'block';
    const whiteBorderColor = 'rgb(255, 255, 255)';
    const blackBorderColor = 'rgb(0, 0, 0)';
    const defaultBgCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"]');
    defaultBgCells.forEach(cell => {
        cell.style.backgroundColor = isDarkMode ? "rgb(0, 0, 0)" : "rgb(255, 255, 255)";
    });
    const defaultBothCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"][data-default-color="true"]');
    defaultBothCells.forEach(cell => {
        cell.style.color = isDarkMode ? "rgb(255, 255, 255)" : "rgb(0, 0, 0)";
    });
    const blackBorderCells = document.querySelectorAll('#xlsx-table td[data-black-border="true"]');
    blackBorderCells.forEach(cell => {
        cell.style.borderColor = isDarkMode ? whiteBorderColor : blackBorderColor;
    });
});
</script>
</body>
</html>`;
    }
}
