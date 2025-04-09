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

            // Create sheet selector options HTML (this is just for the dropdown)
            let sheetOptionsHtml = '';
            workbook.worksheets.forEach((sheet, index) => {
                sheetOptionsHtml += `<option value="${index}">${sheet.name}</option>`;
            });

            // Helper function to map Excel border styles to CSS styles.
            // It returns both the CSS style string and a flag to denote if the original border color is black or near-black.
            const getBorderStyle = (border?: Partial<Excel.Border>): { styleString: string; isBlackOrShade: boolean } => {
                const defaultBorderStyle = '1px solid #c4c4c4'; // Fallback border style
                const defaultResult = { styleString: defaultBorderStyle, isBlackOrShade: false };

                if (!border || !border.style) {
                    return defaultResult;
                }

                const originalColor = border.color?.argb
                    ? convertARGBToRGBA(border.color.argb)
                    : 'rgba(0, 0, 0, 1)'; // Default to black if color is missing

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

            // Generate table HTML for a worksheet.
            const generateTableHtml = (worksheet: Excel.Worksheet): string => {
                // Helper: Determine if a cell is effectively empty.
                // A cell is considered empty only when:
                //   - It has no value (null, undefined, or empty string), AND
                //   - It does NOT have a fill with an fgColor argb (custom background color), AND
                //   - It does NOT have any border defined on any side.
                const isCellEmpty = (cell: Excel.Cell): boolean => {
                    // Check if the cell has a text or numeric value.
                    if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
                        return false;
                    }
                    // Check for custom fill: if a fill exists, and it has a defined foreground color.
                    if (cell.fill && (cell.fill as any).fgColor && (cell.fill as any).fgColor.argb) {
                        return false;
                    }
                    // Check for borders: if any side is defined, consider the cell non-empty.
                    if (cell.border) {
                        const { top, bottom, left, right } = cell.border;
                        if (top || bottom || left || right) {
                            return false;
                        }
                    }
                    return true;
                };

                // Helper: Determine if an entire row is empty (all cells empty).
                const isRowEmpty = (row: Excel.Row): boolean => {
                    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                        const cell = row.getCell(colNumber);
                        if (!isCellEmpty(cell)) {
                            return false;
                        }
                    }
                    return true;
                };

                // Use a backward scan from the total rowCount.
                let lastNonEmptyRow = worksheet.rowCount;
                while (lastNonEmptyRow > 0 && isRowEmpty(worksheet.getRow(lastNonEmptyRow))) {
                    lastNonEmptyRow--;
                }
                // Ensure at least one row is rendered if the worksheet is completely empty.
                if (lastNonEmptyRow === 0) {
                    lastNonEmptyRow = 1;
                }

                // Build the HTML table from row 1 up to lastNonEmptyRow.
                let tableHtml = '<table id="xlsx-table" border="1" cellspacing="0" cellpadding="5">';
                for (let rowNumber = 1; rowNumber <= lastNonEmptyRow; rowNumber++) {
                    tableHtml += '<tr>';
                    const row = worksheet.getRow(rowNumber);

                    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
                        let cellValue = '&nbsp;';
                        let style = '';
                        let isDefaultBlack = false;
                        let hasCustomBackground = false;
                        let cellHasBlackBorder = false;

                        const cell = row.getCell(colNumber);

                        // Process custom background color.
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
                                if (cell.font.size) {
                                    style += `font-size:${cell.font.size + 2}px;`;
                                }
                                if (cell.font.name) {
                                    style += `font-family:${cell.font.name};`;
                                }
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

                        // Process the border styles for all four sides.
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

                        // Add data attributes for potential further usage.
                        const dataAttrs = [];
                        if (isDefaultBlack) {
                            dataAttrs.push('data-default-color="true"');
                        }
                        if (!hasCustomBackground) {
                            dataAttrs.push('data-default-bg="true"');
                        }
                        if (cellHasBlackBorder) {
                            dataAttrs.push('data-black-border="true"');
                        }
                        if (!cell || (cell.value === null && !hasCustomBackground)) {
                            dataAttrs.push('data-empty="true"');
                        }
                        const dataAttrStr = dataAttrs.join(' ');

                        tableHtml += `<td ${dataAttrStr} style="${style}">${cellValue}</td>`;
                    }
                    tableHtml += '</tr>';
                }
                tableHtml += '</table>';
                return tableHtml;
            };

            // Prepare state for the webview with the worksheets and table HTML.
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
        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>XLSX Viewer</title>
            <style>
                body { 
                    font-family: sans-serif; 
                    padding: 10px; 
                    background-color: rgb(255, 255, 255);
                    margin: 0;
                    overflow-x: auto;
                }
                .table-container {
                    width: 100%;
                    overflow-x: auto;
                }
                table { 
                    border-collapse: collapse; 
                    width: auto;
                    min-width: 100%;
                    table-layout: fixed;
                }
                th, td {
                    border: none; 
                    padding: 8px;
                    text-align: left;
                    white-space: nowrap;
                }
                /* Default cell background */
                td { background-color: rgb(255, 255, 255); }
                /* Alternate background used in dark mode */
                body.alt-bg { background-color: rgb(0, 0, 0); }
                .alt-bg td[data-default-bg="true"] { background-color: rgb(0, 0, 0); }
                .button-container {
                    margin-bottom: 10px;
                    display: flex;
                    gap: 10px;
                    position: sticky;
                    top: 0;
                    background-color: inherit;
                    z-index: 1;
                }
                .toggle-button {
                    max-height: 42px;
                    padding: 8px;
                    font-size: 14px;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background-color: #2196f3;
                    color: white;
                    transition: all 0.2s ease;
                }
                .toggle-button:hover {
                    background-color: #1976d2;
                }
                .toggle-button:active {
                    background-color: #1565c0;
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
            </style>
        </head>
        <body>
            <div class="button-container">
                <select id="sheetSelector" class="sheet-selector">
                    ${sheetOptionsHtml}
                </select>
                <button id="toggleButton" class="toggle-button">
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
            <div class="table-container">
                <div id="table-content"></div>
            </div>
            <script>
                const workbookState = ${JSON.stringify(workbookState)};
                const sheetSelector = document.getElementById('sheetSelector');
                const toggleButton = document.getElementById('toggleButton');
                const body = document.body;
                const tableContent = document.getElementById('table-content');
                const lightIcon = document.getElementById('lightIcon');
                const darkIcon = document.getElementById('darkIcon');

                // Function to update table content based on selected sheet.
                const updateTable = (sheetIndex) => {
                    tableContent.innerHTML = workbookState.worksheets[sheetIndex].tableHtml;
                };

                // Load the first sheet initially.
                updateTable(0);

                sheetSelector.addEventListener('change', (e) => {
                    updateTable(parseInt(e.target.value));
                });

                toggleButton.addEventListener('click', () => {
                    body.classList.toggle('alt-bg');
                    const isDarkMode = body.classList.contains('alt-bg');

                    // Toggle icons
                    if (isDarkMode) {
                        lightIcon.style.display = 'block';
                        darkIcon.style.display = 'none';
                    } else {
                        lightIcon.style.display = 'none';
                        darkIcon.style.display = 'block';
                    }

                    const whiteBorderColor = 'rgb(255, 255, 255)'; // White for dark mode borders
                    const blackBorderColor = 'rgb(0, 0, 0)'; // Black for light mode borders

                    // Toggle the default cell background colors.
                    const defaultBgCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"]');
                    defaultBgCells.forEach(cell => {
                        cell.style.backgroundColor = isDarkMode ? "rgb(0, 0, 0)" : "rgb(255, 255, 255)";
                    });

                    // Toggle the default text color for cells that are both default background and default text color.
                    const defaultBothCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"][data-default-color="true"]');
                    defaultBothCells.forEach(cell => {
                        cell.style.color = isDarkMode ? "rgb(255, 255, 255)" : "rgb(0, 0, 0)";
                    });

                    // Toggle border colors for cells that originally had black or near-black borders.
                    const blackBorderCells = document.querySelectorAll('#xlsx-table td[data-black-border="true"]');
                    blackBorderCells.forEach(cell => {
                        cell.style.borderColor = isDarkMode ? whiteBorderColor : blackBorderColor;
                    });
                });
            </script>
        </body>
        </html>
        `;
    }
}