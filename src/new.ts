import * as vscode from 'vscode';
import * as Excel from 'exceljs';
import { convertARGBToRGBA } from './utilities';

// Helper function to check if an RGBA color is black or a shade of black
const isShadeOfBlack = (rgbaColor: string): boolean => {
    const match = rgbaColor.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d.]+)?\)/);
    if (!match) return false; // Should not happen with convertARGBToRGBA output
    const r = parseInt(match[1], 10);
    const g = parseInt(match[2], 10);
    const b = parseInt(match[3], 10);
    const threshold = 30; // Define shade threshold
    return r <= threshold && g <= threshold && b <= threshold;
};

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
            
            // Create sheet selector options HTML (not the full select element)
            let sheetOptionsHtml = '';
            workbook.worksheets.forEach((sheet, index) => {
                sheetOptionsHtml += `<option value="${index}">${sheet.name}</option>`;
            });

            // Helper function to map Excel border styles to CSS
            // Returns the style string and whether the original color was black/shade
            const getBorderStyle = (border?: Partial<Excel.Border>): { styleString: string; isBlackOrShade: boolean } => {
                const defaultBorderStyle = '1px solid #c4c4c4'; // Default light gray border
                const defaultResult = { styleString: defaultBorderStyle, isBlackOrShade: false };

                if (!border || !border.style) {
                    return defaultResult; // Apply default if style is missing for this edge
                }

                const originalColor = border.color && border.color.argb ? convertARGBToRGBA(border.color.argb) : 'rgba(0, 0, 0, 1)'; // Default to black RGBA
                const isBlack = isShadeOfBlack(originalColor);
                // Always use the original color for initial display
                const displayColor = originalColor; 

                let stylePart = '';
                switch (border.style) {
                    case 'thin': stylePart = `1px solid ${displayColor}`; break;
                    case 'medium': stylePart = `2px solid ${displayColor}`; break;
                    case 'thick': stylePart = `3px solid ${displayColor}`; break;
                    case 'dotted': stylePart = `1px dotted ${displayColor}`; break;
                    case 'dashed': stylePart = `1px dashed ${displayColor}`; break;
                    case 'double': stylePart = `3px double ${displayColor}`; break;
                    default: stylePart = `1px solid ${displayColor}`; break; // Fallback
                }
                // Return the style string based on original color, and the flag
                return { styleString: stylePart, isBlackOrShade: isBlack };
            };


            // Function to generate table HTML for a worksheet
            const generateTableHtml = (worksheet: Excel.Worksheet) => {
                const rowCount = worksheet.rowCount;
                const columnCount = worksheet.columnCount;

                let tableHtml = '<table id="xlsx-table" border="1" cellspacing="0" cellpadding="5">';

                // Iterate through all rows, including empty ones
                for (let rowNumber = 1; rowNumber <= rowCount; rowNumber++) {
                    tableHtml += '<tr>';
                    const row = worksheet.getRow(rowNumber);
                    
                    // Iterate through all columns
                    for (let colNumber = 1; colNumber <= columnCount; colNumber++) {
                        let cellValue = '&nbsp;';
                        let style = '';
                        let isDefaultBlack = false;
                        let hasCustomBackground = false;
                        let cellHasBlackBorder = false; // Track if any border was originally black/shade

                        const cell = row.getCell(colNumber);

                        // Check for background color first, even in empty cells
                        if (cell && cell.fill && (cell.fill as any).fgColor && (cell.fill as any).fgColor.argb) {
                            style += `background-color:${convertARGBToRGBA((cell.fill as any).fgColor.argb)};`;
                            hasCustomBackground = true;
                        }

                        if (cell && cell.value !== null && cell.value !== undefined) {
                            cellValue = cell.value.toString();
                            
                            // Font styles
                            if (cell.font) {
                                style += cell.font.bold ? 'font-weight:bold;' : '';
                                style += cell.font.italic ? 'font-style:italic;' : '';
                                style += cell.font.strike ? 'text-decoration:line-through;' : '';
                                if (cell.font.size) {
                                    style += `font-size:${cell.font.size+2}px;`;
                                }
                                if (cell.font.name) {
                                    style += `font-family:${cell.font.name};`;
                                }

                                if (cell.font.color && typeof cell.font.color.argb === "string") {
                                    style += `color: ${convertARGBToRGBA(cell.font.color.argb)};`;
                                } else {
                                    style += `color: rgb(0, 0, 0);`;  // Default black
                                    isDefaultBlack = true;  // Mark for toggling
                                }
                            } else {
                                style += `color: rgb(0, 0, 0);`;  // Default black
                                isDefaultBlack = true;  // Mark for toggling
                            }
                        } else {
                            style += `color: rgb(0, 0, 0);`;  // Default black
                            isDefaultBlack = true;  // Mark for toggling
                        }

                        // Border styles
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


                        // Add data attributes for default colors and borders
                        const dataAttrs = [];
                        if (isDefaultBlack) {
                            dataAttrs.push('data-default-color="true"');
                        }
                        if (!hasCustomBackground) {
                            dataAttrs.push('data-default-bg="true"');
                        }
                        if (cellHasBlackBorder) { // Add the new attribute
                            dataAttrs.push('data-black-border="true"');
                        }
                        // Add empty cell attribute if cell has no content and no custom background
                        if (!cell || (!cell.value && !hasCustomBackground)) {
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

            // Store the workbook in the webview state
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
                    /* border: 1px solid #ccc; */ /* Removed default border */
                    border: none; /* Explicitly set to none initially */
                    padding: 8px;
                    text-align: left;
                    white-space: nowrap;
                }
                /* Default background */
                td { background-color: rgb(255, 255, 255); }
                /* Alternate background when toggled */
                body.alt-bg { background-color: rgb(0, 0, 0); }
                .alt-bg td[data-default-bg="true"] { background-color: rgb(0, 0, 0); }
                /* Removed fade-empty-borders styles */
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
                .toggle-button:active {
                    background-color: #1565c0;
                }
                .toggle-button svg {
                    width: 16px;
                    height: 16px;
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
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M12 2v20M2 12h20M4.93 4.93l14.14 14.14M19.07 4.93L4.93 19.07"/>
                    </svg>
                    Toggle Background
                </button>
                <!-- Removed Fade Empty Cells button -->
            </div>
            <div class="table-container">
                <div id="table-content"></div>
            </div>
            <script>
                const workbookState = ${JSON.stringify(workbookState)};
                const sheetSelector = document.getElementById('sheetSelector');
                const toggleButton = document.getElementById('toggleButton');
                // Removed toggleBordersButton reference
                const body = document.body;
                const tableContent = document.getElementById('table-content');

                // Function to update the table content
                const updateTable = (sheetIndex) => {
                    tableContent.innerHTML = workbookState.worksheets[sheetIndex].tableHtml;
                };

                // Initial table load
                updateTable(0);

                // Sheet selection handler
                sheetSelector.addEventListener('change', (e) => {
                    updateTable(parseInt(e.target.value));
                });

                toggleButton.addEventListener('click', () => {
                    body.classList.toggle('alt-bg');
                    const isDarkMode = body.classList.contains('alt-bg');
                    const whiteBorderColor = 'rgb(255, 255, 255)'; // White for dark mode borders
                    const blackBorderColor = 'rgb(0, 0, 0)'; // Black for light mode borders (reverting)

                    // Change background color of cells with default background
                    const defaultBgCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"]');
                    defaultBgCells.forEach(cell => {
                        cell.style.backgroundColor = isDarkMode ? "rgb(0, 0, 0)" : "rgb(255, 255, 255)";
                    });

                    // Change text color only for cells with both default background and default text color
                    const defaultBothCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"][data-default-color="true"]');
                    defaultBothCells.forEach(cell => {
                        cell.style.color = isDarkMode ? "rgb(255, 255, 255)" : "rgb(0, 0, 0)";
                    });

                    // Toggle border colors for cells that originally had black/shade borders
                    const blackBorderCells = document.querySelectorAll('#xlsx-table td[data-black-border="true"]');
                    blackBorderCells.forEach(cell => {
                        // Set borderColor which applies to all four sides
                        // In dark mode, change black borders to white. In light mode, change them back to black.
                        cell.style.borderColor = isDarkMode ? whiteBorderColor : blackBorderColor;
                    });
                });

                // Removed toggleBordersButton event listener
            </script>
        </body>
        </html>
        `;
    }
}
