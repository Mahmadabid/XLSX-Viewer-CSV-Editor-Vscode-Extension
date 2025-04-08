import * as vscode from 'vscode';
import * as Excel from 'exceljs';

export function activate(context: vscode.ExtensionContext) {
    const provider = new XLSXEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('xlsxViewer.xlsx', provider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );
}

class XLSXEditorProvider implements vscode.CustomReadonlyEditorProvider {
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
                body { font-family: sans-serif; padding: 10px; background-color: rgb(255, 255, 255); }
                table { border-collapse: collapse; width: 100%; }
                th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                /* Default background */
                td { background-color: rgb(255, 255, 255); }
                /* Alternate background when toggled */
                body.alt-bg { background-color: rgb(0, 0, 0); }
                .alt-bg td[data-default-bg="true"] { background-color: rgb(0, 0, 0); }
                #toggleButton {
                    margin-bottom: 10px;
                    padding: 5px 10px;
                    font-size: 14px;
                }
            </style>
        </head>
        <body>
            <button id="toggleButton">Toggle Background Color</button>
            <div class="table-container">
                ${content}
            </div>
            <script>
                const toggleButton = document.getElementById('toggleButton');
                const body = document.body;

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
            </script>
        </body>
        </html>
        `;
    }
}

export function deactivate() { }

/**
 * Converts an Excel ARGB color string ("AARRGGBB") to a CSS rgba() string.
 */
function convertARGBToRGBA(argb: string): string {
    if (argb.length !== 8) {
        return `#${argb}`;
    }
    const alpha = parseInt(argb.substring(0, 2), 16) / 255;
    const red = parseInt(argb.substring(2, 4), 16);
    const green = parseInt(argb.substring(4, 6), 16);
    const blue = parseInt(argb.substring(6, 8), 16);
    return `rgba(${red}, ${green}, ${blue}, ${alpha.toFixed(2)})`;
}