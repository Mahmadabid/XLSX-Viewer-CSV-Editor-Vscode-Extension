import * as vscode from 'vscode';
import * as Excel from 'exceljs';
import { convertARGBToRGBA } from './utilities';

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
            const worksheet = workbook.worksheets[0];
            
            // Get the actual row and column counts
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
                            if (cell.font.size) {
                                style += `font-size:${cell.font.size}px;`;
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

                    // Add data attributes for default colors
                    const dataAttrs = [];
                    if (isDefaultBlack) {
                        dataAttrs.push('data-default-color="true"');
                    }
                    if (!hasCustomBackground) {
                        dataAttrs.push('data-default-bg="true"');
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

            webviewPanel.webview.options = { enableScripts: true };
            webviewPanel.webview.html = this.getWebviewContent(tableHtml);
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading XLSX file: ${error}`);
        }
    }

    private getWebviewContent(content: string): string {
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
                    border: 1px solid #ccc; 
                    padding: 8px; 
                    text-align: left;
                    white-space: nowrap;
                }
                /* Default background */
                td { background-color: rgb(255, 255, 255); }
                /* Alternate background when toggled */
                body.alt-bg { background-color: rgb(0, 0, 0); }
                .alt-bg td[data-default-bg="true"] { background-color: rgb(0, 0, 0); }
                /* Empty cell border styles */
                td[data-empty="true"] { border-color: rgba(204, 204, 204, 1); }
                .fade-empty-borders td[data-empty="true"] { border-color: rgba(204, 204, 204, 0.3); }
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
            </style>
        </head>
        <body>
            <div class="button-container">
                <button id="toggleButton" class="toggle-button">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M12 2v20M2 12h20M4.93 4.93l14.14 14.14M19.07 4.93L4.93 19.07"/>
                    </svg>
                    Toggle Background
                </button>
                <button id="toggleBordersButton" class="toggle-button">
                    <svg id="borderIcon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
                        <line x1="9" y1="3" x2="9" y2="21"/>
                        <line x1="15" y1="3" x2="15" y2="21"/>
                        <line x1="3" y1="9" x2="21" y2="9"/>
                        <line x1="3" y1="15" x2="21" y2="15"/>
                    </svg>
                    Fade Empty Cells
                </button>
            </div>
            <div class="table-container">
                ${content}
            </div>
            <script>
                const toggleButton = document.getElementById('toggleButton');
                const toggleBordersButton = document.getElementById('toggleBordersButton');
                const body = document.body;
                const table = document.getElementById('xlsx-table');

                toggleButton.addEventListener('click', () => {
                    body.classList.toggle('alt-bg');

                    // Change background color of cells with default background
                    const defaultBgCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"]');
                    defaultBgCells.forEach(cell => {
                        if (body.classList.contains('alt-bg')) {
                            cell.style.backgroundColor = "rgb(0, 0, 0)";
                        } else {
                            cell.style.backgroundColor = "rgb(255, 255, 255)";
                        }
                    });

                    // Change text color only for cells with both default background and default text color
                    const defaultBothCells = document.querySelectorAll('#xlsx-table td[data-default-bg="true"][data-default-color="true"]');
                    defaultBothCells.forEach(cell => {
                        if (body.classList.contains('alt-bg')) {
                            cell.style.color = "rgb(255, 255, 255)";
                        } else {
                            cell.style.color = "rgb(0, 0, 0)";
                        }
                    });
                });

                toggleBordersButton.addEventListener('click', () => {
                    table.classList.toggle('fade-empty-borders');
                });
            </script>
        </body>
        </html>
        `;
    }
} 