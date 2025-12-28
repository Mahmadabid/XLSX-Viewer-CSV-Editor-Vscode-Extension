/* global acquireVsCodeApi */

(function () {
    const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : { postMessage: () => { } };

    // Data injected from the extension via postMessage
    let worksheetsData = [];
    let currentWorksheet = 0;

    // Selection state
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

    // Auto-scroll while dragging selection
    let autoScrollRequest = null;
    let lastMousePos = null; // { x, y }
    const AUTO_SCROLL_THRESHOLD = 40; // px
    const AUTO_SCROLL_STEP = 20; // px per frame

    let handlersAttached = false;

    // Settings (persisted by extension)
    let currentSettings = {
        firstRowIsHeader: false,
        stickyToolbar: true,
        stickyHeader: false,
        hyperlinkPreview: true
    };

    // Table edit mode (text-only)
    let isEditMode = false;

    // Save state (CSV-parity)
    let isSaving = false;
    let exitAfterSave = false;

    // Hyperlink hover tooltip
    let linkTooltip = null;
    let linkTooltipHideTimer = null;

    // Toast
    let toastEl = null;

    // Copy state (CSV-parity: avoid concurrent copies)
    let isCopying = false;

    function setButtonsEnabled(enabled) {
        const saveBtn = document.getElementById('saveTableEditsButton');
        const cancelBtn = document.getElementById('cancelTableEditsButton');
        if (saveBtn) saveBtn.disabled = !enabled;
        if (cancelBtn) cancelBtn.disabled = !enabled;
    }

    function normalizeCellText(text) {
        if (!text) return '';
        return String(text).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trimEnd();
    }

    function yieldToMain() {
        return new Promise(resolve => {
            requestAnimationFrame(() => {
                setTimeout(resolve, 0);
            });
        });
    }

    async function writeToClipboardAsync(text) {
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            try {
                await navigator.clipboard.writeText(text);
                return;
            } catch {
                // fall through to execCommand
            }
        }

        await new Promise((resolve, reject) => {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.setAttribute('readonly', '');
            textarea.style.cssText = `
                position: fixed;
                left: -9999px;
                top: 0;
                width: 2px;
                height: 2px;
                padding: 0;
                border: none;
                outline: none;
                opacity: 0;
                pointer-events: none;
            `;

            document.body.appendChild(textarea);

            try {
                textarea.focus();
                textarea.select();
                textarea.setSelectionRange(0, text.length);
                const ok = document.execCommand('copy');
                document.body.removeChild(textarea);
                if (ok) resolve();
                else reject(new Error('execCommand("copy") returned false'));
            } catch (err) {
                document.body.removeChild(textarea);
                reject(err);
            }
        });
    }

    function getExcelColumnLabel(n) {
        let label = '';
        while (n > 0) {
            const rem = (n - 1) % 26;
            label = String.fromCharCode(65 + rem) + label;
            n = Math.floor((n - 1) / 26);
        }
        return label;
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

    function createTable(worksheetData) {
        const data = worksheetData.data;

        let html = '<div class="table-scroll"><table>';

        // Header row
        html += '<thead><tr>';
        html += '<th class="corner-cell"></th>';
        for (let c = 1; c <= data.maxCol; c++) {
            const width = data.columnWidths[c - 1] || 80;
            html += '<th class="col-header" data-col="' + (c - 1) + '" style="width: ' + width + 'px; min-width: ' + width + 'px;">';
            html += getExcelColumnLabel(c);
            html += '<div class="col-resize-handle" data-col="' + (c - 1) + '"></div>';
            html += '</th>';
        }
        html += '</tr></thead><tbody>';

        // Data rows
        data.rows.forEach((row, rowIndex) => {
            const height = row.height || 20;
            const isHeaderRow = rowIndex === 0;
            html += '<tr style="height: ' + height + 'px;"' + (isHeaderRow ? ' class="header-row"' : '') + '>';
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
                    const cellWidth = data.columnWidths
                        .slice(actualCol - 1, actualCol - 1 + cellData.colspan)
                        .reduce((sum, w) => sum + (w || 80), 0);

                    html += '<td';
                    html += ' data-row="' + rowIndex + '"';
                    html += ' data-col="' + virtualColIndex + '"';
                    html += ' data-rownum="' + cellData.rowNumber + '"';
                    html += ' data-colnum="' + cellData.colNumber + '"';
                    if (cellData.hasDefaultBg) html += ' data-default-bg="true"';
                    if (cellData.isDefaultColor) html += ' data-default-color="true"';
                    if (cellData.hasBlackBorder) html += ' data-black-border="true"';
                    if (cellData.hasBlackBackground) html += ' data-black-bg="true"';
                    if (cellData.isEmpty) html += ' data-empty="true"';
                    if (cellData.hyperlink) html += ' data-hyperlink="' + String(cellData.hyperlink).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;') + '"';
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

        html += '</tbody></table></div>';
        return html;
    }

    function ensureToast() {
        if (toastEl) return toastEl;
        toastEl = document.createElement('div');
        toastEl.id = 'saveToast';
        toastEl.className = 'toast-notification';
        toastEl.innerHTML = `
            <div class="toast-icon-wrapper">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="20 6 9 17 4 12"></polyline>
              </svg>
            </div>
            <span class="toast-text"></span>
        `;
        document.body.appendChild(toastEl);
        return toastEl;
    }

    function showToast(message) {
        const toast = ensureToast();
        const textEl = toast.querySelector('.toast-text');
        if (textEl) textEl.textContent = String(message || '');
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 2000);
    }

    function setLoadingText(text) {
        const el = document.querySelector('.loading-text');
        if (el) el.textContent = text;
    }

    function showLoading() {
        const el = document.getElementById('loadingOverlay');
        if (el) el.classList.remove('hidden');
    }

    function hideLoading() {
        const el = document.getElementById('loadingOverlay');
        if (el) el.classList.add('hidden');
    }

    function renderWorksheet(index) {
        if (!worksheetsData || !worksheetsData.length) return;

        showLoading();

        // Allow the overlay to render
        setTimeout(() => {
            const container = document.getElementById('tableContainer');
            if (!container) return;

            container.innerHTML = createTable(worksheetsData[index]);
            initializeSelection();
            initializeResize();
            initializeHyperlinkHover();
            hideLoading();
        }, 100);
    }

    function initializeResize() {
        const table = document.querySelector('table');
        if (!table) return;

        // Column/row resize handles
        table.addEventListener('mousedown', (e) => {
            if (isEditMode) return;
            const target = e.target;
            if (target && target.classList && target.classList.contains('col-resize-handle')) {
                e.preventDefault();
                e.stopPropagation();

                isResizing = true;
                resizeType = 'column';
                resizeIndex = parseInt(target.dataset.col, 10);
                resizeStartPos = e.clientX;

                const header = target.parentElement;
                resizeStartSize = header ? header.offsetWidth : 0;

                document.body.style.cursor = 'col-resize';
                const indicator = document.getElementById('resizeIndicator');
                if (indicator) indicator.style.display = 'block';
                return false;
            }

            if (target && target.classList && target.classList.contains('row-resize-handle')) {
                e.preventDefault();
                e.stopPropagation();

                isResizing = true;
                resizeType = 'row';
                resizeIndex = parseInt(target.dataset.row, 10);
                resizeStartPos = e.clientY;

                const header = target.parentElement;
                resizeStartSize = header ? header.offsetHeight : 0;

                document.body.style.cursor = 'row-resize';
                const indicator = document.getElementById('resizeIndicator');
                if (indicator) indicator.style.display = 'block';
                return false;
            }
        });

        document.addEventListener('mousemove', (e) => {
            if (isEditMode) return;
            if (!isResizing) return;

            const tableEl = document.querySelector('table');
            if (!tableEl) return;

            const indicator = document.getElementById('resizeIndicator');

            if (resizeType === 'column') {
                const delta = e.clientX - resizeStartPos;
                const newSize = Math.max(20, resizeStartSize + delta);

                const headers = tableEl.querySelectorAll('th.col-header[data-col="' + resizeIndex + '"]');
                const cells = tableEl.querySelectorAll('td[data-col="' + resizeIndex + '"]');

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

                if (indicator) {
                    indicator.style.left = e.clientX + 'px';
                    indicator.style.top = e.clientY + 'px';
                    indicator.textContent = newSize + 'px';
                }
            } else if (resizeType === 'row') {
                const delta = e.clientY - resizeStartPos;
                const newSize = Math.max(15, resizeStartSize + delta);

                const headers = tableEl.querySelectorAll('th.row-header[data-row="' + resizeIndex + '"]');
                const row = tableEl.querySelectorAll('tr')[resizeIndex + 1]; // +1 for header row

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

                if (indicator) {
                    indicator.style.left = e.clientX + 'px';
                    indicator.style.top = e.clientY + 'px';
                    indicator.textContent = newSize + 'px';
                }
            }
        });

        document.addEventListener('mouseup', () => {
            if (isResizing) {
                isResizing = false;
                resizeType = null;
                resizeIndex = -1;
                document.body.style.cursor = '';
                const indicator = document.getElementById('resizeIndicator');
                if (indicator) indicator.style.display = 'none';
            }
        });

        // Double-click to auto-fit
        table.addEventListener('dblclick', (e) => {
            if (isEditMode) return;
            const target = e.target;
            if (target && target.classList && target.classList.contains('col-resize-handle')) {
                e.preventDefault();
                autoFitColumn(parseInt(target.dataset.col, 10));
            } else if (target && target.classList && target.classList.contains('row-resize-handle')) {
                e.preventDefault();
                autoFitRow(parseInt(target.dataset.row, 10));
            }
        });
    }

    function autoFitColumn(colIndex) {
        const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
        let maxWidth = 50;

        cells.forEach(cell => {
            const content = (cell.textContent || '').trim();
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
            const content = (cell.textContent || '').trim();
            if (content.length > 50) {
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
        if (!worksheetsData || !worksheetsData.length) return;
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
        const info = document.getElementById('selectionInfo');
        if (info) info.style.display = 'none';
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
            const row = parseInt(cell.dataset.row, 10);
            const col = parseInt(cell.dataset.col, 10);

            if (row >= minRow && row <= maxRow && col >= minCol && col <= maxCol) {
                cell.classList.add('selected');
                selectedCells.add(cell);
            }
        });

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
        if (!info) return;

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
        copySelectionToClipboard();
    }

    async function copySelectionToClipboard() {
        if (!selectedCells || selectedCells.size === 0) return;
        if (isCopying) return;

        isCopying = true;
        const CHUNK_SIZE = 2000;

        try {
            showToast('Copying...');
            await yieldToMain();

            const cellsArray = Array.from(selectedCells);
            const totalCells = cellsArray.length;
            const rowSet = new Set();
            const colSet = new Set();

            for (let i = 0; i < totalCells; i++) {
                const td = cellsArray[i];
                const r = parseInt(td.dataset.row, 10);
                const c = parseInt(td.dataset.col, 10);
                if (!isNaN(r) && !isNaN(c)) {
                    rowSet.add(r);
                    colSet.add(c);
                }
                if ((i + 1) % CHUNK_SIZE === 0) {
                    await yieldToMain();
                }
            }

            await yieldToMain();

            const sortedRows = Array.from(rowSet).sort((a, b) => a - b);
            const sortedCols = Array.from(colSet).sort((a, b) => a - b);
            const outputLines = new Array(sortedRows.length);

            for (let i = 0; i < sortedRows.length; i++) {
                const r = sortedRows[i];
                const lineParts = new Array(sortedCols.length);
                for (let j = 0; j < sortedCols.length; j++) {
                    const c = sortedCols[j];
                    const cell = document.querySelector('td[data-row="' + r + '"][data-col="' + c + '"]');
                    lineParts[j] = normalizeCellText(cell ? (cell.textContent || '') : '');
                }
                outputLines[i] = lineParts.join('\t');
                if ((i + 1) % CHUNK_SIZE === 0) {
                    await yieldToMain();
                }
            }

            await yieldToMain();

            const tsv = outputLines.join('\n');
            await writeToClipboardAsync(tsv);

            selectedCells.forEach(cell => cell.classList.add('copying'));
            setTimeout(() => selectedCells.forEach(cell => cell.classList.remove('copying')), 300);

            showToast('Copied ' + totalCells + ' cells');
        } catch (err) {
            console.error('Copy operation failed:', err);
            showToast('Copy failed');
        } finally {
            isCopying = false;
        }
    }

    function invertColor(color) {
        const match = String(color || '').match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
        if (!match) return color;

        const r = 255 - parseInt(match[1], 10);
        const g = 255 - parseInt(match[2], 10);
        const b = 255 - parseInt(match[3], 10);
        return `rgb(${r}, ${g}, ${b})`;
    }

    function initializeSelection() {
        const tableContainer = document.getElementById('tableContainer');
        const table = tableContainer ? tableContainer.querySelector('table') : null;
        if (!table) return;

        table.addEventListener('selectstart', (e) => {
            if (isEditMode) return;
            e.preventDefault();
            return false;
        });

        table.addEventListener('mousedown', (e) => {
            if (isEditMode) return;
            if (e.target && e.target.classList && (e.target.classList.contains('col-resize-handle') || e.target.classList.contains('row-resize-handle'))) {
                return;
            }

            const target = e.target.closest('td, th');
            if (!target) return;

            e.preventDefault();

            if (target.classList.contains('col-header')) {
                const colIndex = parseInt(target.dataset.col, 10);
                if (!e.shiftKey) {
                    lastSelectedColumn = colIndex;
                }
                selectColumn(colIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }

            if (target.classList.contains('row-header')) {
                const rowIndex = parseInt(target.dataset.row, 10);
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
                    allCells[0].classList.add('active-cell');
                    activeCell = allCells[0];
                }
                updateSelectionInfo();
                return;
            }

            if (target.tagName === 'TD') {
                const row = parseInt(target.dataset.row, 10);
                const col = parseInt(target.dataset.col, 10);

                if (e.ctrlKey || e.metaKey) {
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
                    const startRow = parseInt(activeCell.dataset.row, 10);
                    const startCol = parseInt(activeCell.dataset.col, 10);
                    selectRange(startRow, startCol, row, col);
                } else {
                    isSelecting = true;
                    selectionStart = { row, col };
                    selectCell(target);
                }
            }
        });

        table.addEventListener('mousemove', (e) => {
            if (isEditMode) return;
            if (!isSelecting || !selectionStart) return;

            // Track last mouse position for auto-scroll
            lastMousePos = { x: e.clientX, y: e.clientY };

            const target = e.target.closest('td');
            if (!target) return;

            const row = parseInt(target.dataset.row, 10);
            const col = parseInt(target.dataset.col, 10);

            if (!selectionEnd || selectionEnd.row !== row || selectionEnd.col !== col) {
                selectionEnd = { row, col };
                selectRange(selectionStart.row, selectionStart.col, row, col);
            }

            // Start auto-scroll loop if needed
            startAutoScroll();
        });

        function startAutoScroll() {
            if (autoScrollRequest) return;
            autoScrollLoop();
        }

        function stopAutoScroll() {
            if (autoScrollRequest) {
                cancelAnimationFrame(autoScrollRequest);
                autoScrollRequest = null;
            }
        }

        function autoScrollLoop() {
            autoScrollRequest = requestAnimationFrame(() => {
                if (!isSelecting || !lastMousePos) {
                    stopAutoScroll();
                    return;
                }

                const tableContainer = document.getElementById('tableContainer');
                const scrollArea = tableContainer ? tableContainer.querySelector('.table-scroll') : null;
                if (!scrollArea) {
                    stopAutoScroll();
                    return;
                }

                const rect = scrollArea.getBoundingClientRect();
                let dx = 0;
                let dy = 0;

                if (lastMousePos.x < rect.left + AUTO_SCROLL_THRESHOLD) dx = -AUTO_SCROLL_STEP;
                else if (lastMousePos.x > rect.right - AUTO_SCROLL_THRESHOLD) dx = AUTO_SCROLL_STEP;

                if (lastMousePos.y < rect.top + AUTO_SCROLL_THRESHOLD) dy = -AUTO_SCROLL_STEP;
                else if (lastMousePos.y > rect.bottom - AUTO_SCROLL_THRESHOLD) dy = AUTO_SCROLL_STEP;

                if (dx !== 0 || dy !== 0) {
                    scrollArea.scrollBy({ left: dx, top: dy, behavior: 'auto' });

                    // After scrolling, determine the element under the pointer and update selection
                    const el = document.elementFromPoint(lastMousePos.x, lastMousePos.y);
                    const nearestCell = el ? el.closest && el.closest('td') : null;
                    if (nearestCell) {
                        const r = parseInt(nearestCell.dataset.row, 10);
                        const c = parseInt(nearestCell.dataset.col, 10);
                        if (!selectionEnd || selectionEnd.row !== r || selectionEnd.col !== c) {
                            selectionEnd = { row: r, col: c };
                            selectRange(selectionStart.row, selectionStart.col, r, c);
                        }
                    }
                }

                // Continue loop
                autoScrollLoop();
            });
        }

        document.addEventListener('mouseup', () => {
            isSelecting = false;
            selectionStart = null;
            selectionEnd = null;
            lastMousePos = null;
            stopAutoScroll();
        });

        document.addEventListener('keydown', (e) => {
            const isCmdOrCtrl = e.ctrlKey || e.metaKey;

            if (isCmdOrCtrl && e.key.toLowerCase() === 's') {
                e.preventDefault();
                saveEdits(false);
                return;
            }

            if (isEditMode) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    const active = document.activeElement;
                    if (active && active.tagName === 'TD') {
                        const r = parseInt(active.getAttribute('data-row') || '0', 10);
                        const c = parseInt(active.getAttribute('data-col') || '0', 10);
                        const next = document.querySelector('td[data-row="' + (r + 1) + '"][data-col="' + c + '"]');
                        if (next) {
                            next.focus();
                            const range = document.createRange();
                            const sel = window.getSelection();
                            range.selectNodeContents(next);
                            range.collapse(false);
                            sel.removeAllRanges();
                            sel.addRange(range);
                        }
                    }
                }
                return;
            }

            if (isCmdOrCtrl && e.key.toLowerCase() === 'c') {
                e.preventDefault();
                copySelectionToClipboard();
                return;
            }

            if (isCmdOrCtrl && e.key.toLowerCase() === 'a') {
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

        document.addEventListener('click', (e) => {
            if (isEditMode) return;
            if (!e.target.closest('table') && !e.target.closest('.toolbar')) {
                clearSelection();
            }
        });
    }

    function ensureLinkTooltip() {
        if (linkTooltip) return linkTooltip;
        linkTooltip = document.createElement('div');
        linkTooltip.id = 'linkTooltip';
        linkTooltip.className = 'link-tooltip hidden';
        linkTooltip.innerHTML = `
            <div class="link-tooltip-url" id="linkTooltipUrl"></div>
            <div class="link-tooltip-actions">
                <button type="button" id="linkTooltipOpen" class="toggle-button">Open in Browser</button>
                <button type="button" id="linkTooltipCopy" class="toggle-button">Copy Link</button>
            </div>
        `;

        linkTooltip.addEventListener('mouseenter', () => {
            if (linkTooltipHideTimer) {
                clearTimeout(linkTooltipHideTimer);
                linkTooltipHideTimer = null;
            }
        });

        linkTooltip.addEventListener('mouseleave', () => {
            scheduleHideLinkTooltip();
        });

        document.body.appendChild(linkTooltip);
        return linkTooltip;
    }

    function showLinkTooltipForCell(cellEl) {
        if (!currentSettings.hyperlinkPreview) return;
        if (!cellEl) return;
        const url = cellEl.getAttribute('data-hyperlink') || '';
        if (!url) return;

        const tt = ensureLinkTooltip();
        const urlEl = tt.querySelector('#linkTooltipUrl');
        if (urlEl) urlEl.textContent = url;

        const openBtn = tt.querySelector('#linkTooltipOpen');
        const copyBtn = tt.querySelector('#linkTooltipCopy');

        if (openBtn) {
            openBtn.onclick = (e) => {
                e.preventDefault();
                e.stopPropagation();
                vscode.postMessage({ command: 'openExternal', url });
                hideLinkTooltip();
            };
        }
        if (copyBtn) {
            copyBtn.onclick = async (e) => {
                e.preventDefault();
                e.stopPropagation();
                try {
                    await writeToClipboardAsync(url);
                    showToast('Copied URL');
                    hideLinkTooltip();
                } catch {
                    // ignore
                }
            };
        }

        tt.classList.remove('hidden');

        const rect = cellEl.getBoundingClientRect();
        // Measure after showing
        const ttRect = tt.getBoundingClientRect();
        const left = Math.min(Math.max(8, rect.left), window.innerWidth - ttRect.width - 8);
        const top = Math.min(rect.bottom, window.innerHeight - ttRect.height - 2);
        tt.style.left = left + 'px';
        tt.style.top = top + 'px';
    }

    function hideLinkTooltip() {
        if (!linkTooltip) return;
        linkTooltip.classList.add('hidden');
        linkTooltip.style.left = '';
        linkTooltip.style.top = '';
    }

    function scheduleHideLinkTooltip() {
        if (linkTooltipHideTimer) clearTimeout(linkTooltipHideTimer);
        linkTooltipHideTimer = setTimeout(() => {
            hideLinkTooltip();
            linkTooltipHideTimer = null;
        }, 120);
    }

    function initializeHyperlinkHover() {
        const table = document.querySelector('table');
        if (!table) return;

        table.addEventListener('mouseover', (e) => {
            if (isEditMode) return;
            const t = e && e.target;
            const el = t && t.nodeType === 3 ? t.parentElement : t;
            const cell = el && el.closest ? el.closest('td[data-hyperlink]') : null;
            if (!cell) return;
            if (linkTooltipHideTimer) {
                clearTimeout(linkTooltipHideTimer);
                linkTooltipHideTimer = null;
            }
            showLinkTooltipForCell(cell);
        });

        table.addEventListener('mouseout', (e) => {
            if (isEditMode) return;
            const toEl = e.relatedTarget;
            if (!toEl) {
                scheduleHideLinkTooltip();
                return;
            }

            // If we are moving to an element inside the same cell, don't hide
            const fromCell = e.target.closest('td[data-hyperlink]');
            const toCell = toEl.closest ? toEl.closest('td[data-hyperlink]') : null;
            if (fromCell && toCell === fromCell) {
                return;
            }

            // If we are moving to the tooltip itself, don't hide
            if (linkTooltip && linkTooltip.contains(toEl)) {
                return;
            }

            scheduleHideLinkTooltip();
        });
    }

    function applySettings(settings) {
        currentSettings = {
            firstRowIsHeader: settings && typeof settings.firstRowIsHeader === 'boolean' ? settings.firstRowIsHeader : currentSettings.firstRowIsHeader,
            stickyToolbar: settings && typeof settings.stickyToolbar === 'boolean' ? settings.stickyToolbar : currentSettings.stickyToolbar,
            stickyHeader: settings && typeof settings.stickyHeader === 'boolean' ? settings.stickyHeader : currentSettings.stickyHeader,
            hyperlinkPreview: settings && typeof settings.hyperlinkPreview === 'boolean' ? settings.hyperlinkPreview : currentSettings.hyperlinkPreview
        };

        // Sticky toolbar behavior (CSV parity): when disabled, move toolbar into the scrollable content.
        const toolbar = document.querySelector('.toolbar');
        const headerBg = document.querySelector('.header-background');
        const content = document.getElementById('content');
        const scrollArea = document.querySelector('.table-scroll');

        document.body.classList.toggle('sticky-toolbar-enabled', !!currentSettings.stickyToolbar);
        document.body.classList.toggle('sticky-header-enabled', !!currentSettings.stickyHeader);
        document.body.classList.toggle('first-row-as-header', !!currentSettings.firstRowIsHeader);
        // keep legacy class during transition
        document.body.classList.toggle('sticky-toolbar-disabled', !currentSettings.stickyToolbar);

        if (currentSettings.stickyToolbar) {
            if (toolbar && toolbar.parentElement !== document.body) {
                document.body.insertBefore(toolbar, document.body.firstChild);
            }
            if (toolbar) toolbar.classList.remove('not-sticky');
            if (headerBg) headerBg.style.display = '';
        } else {
            const target = scrollArea || content;
            if (toolbar && target && toolbar.parentNode !== target) {
                target.insertBefore(toolbar, target.firstChild);
            }
            if (toolbar) toolbar.classList.add('not-sticky');
            if (headerBg) headerBg.style.display = 'none';
        }

        const chkHeaderRow = document.getElementById('chkHeaderRow');
        if (chkHeaderRow) chkHeaderRow.checked = !!currentSettings.firstRowIsHeader;
        const chkStickyToolbar = document.getElementById('chkStickyToolbar');
        if (chkStickyToolbar) chkStickyToolbar.checked = !!currentSettings.stickyToolbar;
        const chkStickyHeader = document.getElementById('chkStickyHeader');
        if (chkStickyHeader) {
            chkStickyHeader.checked = !!currentSettings.stickyHeader;
            chkStickyHeader.disabled = !currentSettings.firstRowIsHeader;
            if (chkStickyHeader.disabled) {
                chkStickyHeader.checked = false;
                currentSettings.stickyHeader = false;
                document.body.classList.remove('sticky-header-enabled');
            }
        }
        const chkHyperlinkPreview = document.getElementById('chkHyperlinkPreview');
        if (chkHyperlinkPreview) chkHyperlinkPreview.checked = !!currentSettings.hyperlinkPreview;

        if (!currentSettings.hyperlinkPreview) hideLinkTooltip();
    }

    function postSettings() {
        vscode.postMessage({ command: 'updateSettings', settings: currentSettings });
    }

    function setEditMode(enabled) {
        isEditMode = !!enabled;
        document.body.classList.toggle('edit-mode', isEditMode);

        const sheetSelector = document.getElementById('sheetSelector');
        const toggleExpandButton = document.getElementById('toggleExpandButton');
        const openSettingsButton = document.getElementById('openSettingsButton');
        const toggleBackgroundButton = document.getElementById('toggleBackgroundButton');

        const toggleTableEditButton = document.getElementById('toggleTableEditButton');
        const saveTableEditsButton = document.getElementById('saveTableEditsButton');
        const cancelTableEditsButton = document.getElementById('cancelTableEditsButton');

        if (toggleTableEditButton) toggleTableEditButton.classList.toggle('hidden', isEditMode);
        if (saveTableEditsButton) saveTableEditsButton.classList.toggle('hidden', !isEditMode);
        if (cancelTableEditsButton) cancelTableEditsButton.classList.toggle('hidden', !isEditMode);

        if (sheetSelector) sheetSelector.classList.toggle('hidden', isEditMode);
        if (toggleExpandButton) toggleExpandButton.classList.toggle('hidden', isEditMode);
        if (openSettingsButton) openSettingsButton.classList.toggle('hidden', isEditMode);
        if (toggleBackgroundButton) toggleBackgroundButton.classList.toggle('hidden', isEditMode);

        if (!isEditMode) {
            hideLinkTooltip();
            clearSelection();
            return;
        }

        // Enable contenteditable for table cells
        const table = document.querySelector('#tableContainer table');
        if (!table) return;
        table.querySelectorAll('td').forEach(td => {
            td.setAttribute('contenteditable', 'true');
            td.setAttribute('spellcheck', 'false');
            td.classList.add('editable-cell');
            const currentText = normalizeCellText(td.textContent || '');
            td.textContent = currentText;
            td.dataset.originalText = currentText;
        });
    }

    function captureOriginalCellValues() {
        const table = document.querySelector('#tableContainer table');
        if (!table) return;
        table.querySelectorAll('td[contenteditable="true"]').forEach(td => {
            const currentText = normalizeCellText(td.textContent || '');
            td.textContent = currentText;
            td.dataset.originalText = currentText;
        });
    }

    function saveEdits(shouldExit = false) {
        if (isSaving || !isEditMode) return;
        const table = document.querySelector('#tableContainer table');
        if (!table) return;

        isSaving = true;
        exitAfterSave = !!shouldExit;
        setButtonsEnabled(false);

        if (document.activeElement && document.activeElement.tagName === 'TD') {
            document.activeElement.blur();
        }
        clearSelection();
        if (window.getSelection) {
            window.getSelection().removeAllRanges();
        }

        const edits = [];
        table.querySelectorAll('td[contenteditable="true"]').forEach(td => {
            const row = parseInt(td.getAttribute('data-rownum') || '0', 10);
            const col = parseInt(td.getAttribute('data-colnum') || '0', 10);
            if (!row || !col) return;

            const original = (td.dataset.originalText || '').replace(/\u00a0/g, '');
            const current = (td.textContent || '').replace(/\u00a0/g, '');

            if (current !== original) {
                edits.push({ row, col, value: current });
            }
        });

        setLoadingText('Saving worksheet...');
        showLoading();
        vscode.postMessage({ command: 'saveXlsxEdits', sheetIndex: currentWorksheet, edits });
    }

    function setExpandedMode(isExpanded) {
        document.body.classList.toggle('expanded-mode', !!isExpanded);

        const expandIcon = document.getElementById('expandIcon');
        const collapseIcon = document.getElementById('collapseIcon');
        const text = document.getElementById('expandButtonText');

        if (expandIcon) expandIcon.style.display = isExpanded ? 'none' : 'block';
        if (collapseIcon) collapseIcon.style.display = isExpanded ? 'block' : 'none';
        if (text) text.textContent = isExpanded ? 'Default' : 'Expand';
    }

    function wireSettingsUI() {
        const openBtn = document.getElementById('openSettingsButton');
        const panel = document.getElementById('settingsPanel');
        const cancelBtn = document.getElementById('settingsCancelButton');
        if (!openBtn || !panel || !cancelBtn) return;

        let repositionHandlers = null;

        function repositionPanel() {
            const container = document.querySelector('.toolbar');
            if (!container) return;
            const rect = container.getBoundingClientRect();
            panel.style.position = 'fixed';
            panel.style.left = Math.max(8, rect.left) + 'px';
            panel.style.top = rect.bottom + 'px';
            const maxWidth = Math.min(window.innerWidth - 16, rect.width);
            panel.style.width = Math.max(280, maxWidth) + 'px';
            panel.style.zIndex = '200001';
        }

        function openPanel() {
            // If the panel is inside the toolbar, move it to document.body so it's not clipped by toolbar overflow/stacking context
            const container = document.querySelector('.toolbar');
            if (container && panel.parentNode !== document.body) {
                document.body.appendChild(panel);
            }

            panel.classList.remove('hidden');
            panel.classList.add('floating');
            panel.setAttribute('aria-hidden', 'false');
            document.body.classList.add('settings-open');
            if (container) {
                container.classList.add('settings-open');
                container.classList.add('expanded-toolbar');
            }

            repositionPanel();
            repositionHandlers = () => repositionPanel();
            window.addEventListener('resize', repositionHandlers);
            window.addEventListener('scroll', repositionHandlers, true);
        }

        function closePanel() {
            panel.classList.add('hidden');
            panel.classList.remove('floating');
            panel.setAttribute('aria-hidden', 'true');
            document.body.classList.remove('settings-open');
            const container = document.querySelector('.toolbar');
            if (container) {
                container.classList.remove('settings-open');
                const chkStickyToolbar = document.getElementById('chkStickyToolbar');
                const cfgSticky = chkStickyToolbar ? chkStickyToolbar.checked : true;
                if (!cfgSticky) container.classList.remove('expanded-toolbar');

                // Move the panel back into the toolbar (keep DOM tidy)
                if (panel.parentNode === document.body) {
                    container.insertBefore(panel, container.firstChild);
                }
            }

            // remove inline positioning
            panel.style.position = '';
            panel.style.left = '';
            panel.style.top = '';
            panel.style.width = '';
            panel.style.zIndex = '';

            if (repositionHandlers) {
                window.removeEventListener('resize', repositionHandlers);
                window.removeEventListener('scroll', repositionHandlers, true);
                repositionHandlers = null;
            }
        }

        openBtn.addEventListener('click', () => {
            if (panel.classList.contains('hidden')) openPanel();
            else closePanel();
        });

        cancelBtn.addEventListener('click', () => closePanel());

        // Settings checkboxes
        const chkHeaderRow = document.getElementById('chkHeaderRow');
        if (chkHeaderRow) {
            chkHeaderRow.addEventListener('change', () => {
                currentSettings.firstRowIsHeader = !!chkHeaderRow.checked;
                applySettings(currentSettings);
                postSettings();
            });
        }

        const chkStickyToolbar = document.getElementById('chkStickyToolbar');
        if (chkStickyToolbar) {
            chkStickyToolbar.addEventListener('change', () => {
                currentSettings.stickyToolbar = !!chkStickyToolbar.checked;
                applySettings(currentSettings);
                postSettings();
            });
        }

        const chkStickyHeader = document.getElementById('chkStickyHeader');
        if (chkStickyHeader) {
            chkStickyHeader.addEventListener('change', () => {
                currentSettings.stickyHeader = !!chkStickyHeader.checked;
                applySettings(currentSettings);
                postSettings();
            });
        }

        const chkHyperlinkPreview = document.getElementById('chkHyperlinkPreview');
        if (chkHyperlinkPreview) {
            chkHyperlinkPreview.addEventListener('change', () => {
                currentSettings.hyperlinkPreview = !!chkHyperlinkPreview.checked;
                applySettings(currentSettings);
                postSettings();
            });
        }

        document.addEventListener('click', (e) => {
            if (!panel.classList.contains('hidden')) {
                if (!e.target.closest('.settings-panel') && !e.target.closest('#openSettingsButton')) {
                    closePanel();
                }
            }
        });
    }

    function attachHandlersOnce() {
        if (handlersAttached) return;
        handlersAttached = true;

        // Sheet selector
        const sheetSelector = document.getElementById('sheetSelector');
        if (sheetSelector) {
            sheetSelector.addEventListener('change', (e) => {
                if (isEditMode) return;
                currentWorksheet = parseInt(e.target.value, 10);
                clearSelection();
                renderWorksheet(currentWorksheet);
            });
        }

        // Edit Table
        const toggleTableEditButton = document.getElementById('toggleTableEditButton');
        const saveTableEditsButton = document.getElementById('saveTableEditsButton');
        const cancelTableEditsButton = document.getElementById('cancelTableEditsButton');

        if (toggleTableEditButton) {
            toggleTableEditButton.addEventListener('click', () => {
                setEditMode(true);
            });
        }
        if (saveTableEditsButton) {
            saveTableEditsButton.addEventListener('click', () => {
                saveEdits(true);
            });
        }
        if (cancelTableEditsButton) {
            cancelTableEditsButton.addEventListener('click', () => {
                setEditMode(false);
                renderWorksheet(currentWorksheet);
            });
        }

        // Dark mode toggle
        const toggleBackgroundButton = document.getElementById('toggleBackgroundButton');
        if (toggleBackgroundButton) {
            toggleBackgroundButton.addEventListener('click', () => {
                if (isEditMode) return;
                document.body.classList.toggle('alt-bg');
                const isDarkMode = document.body.classList.contains('alt-bg');

                const lightIcon = document.getElementById('lightIcon');
                const darkIcon = document.getElementById('darkIcon');

                if (isDarkMode) {
                    if (lightIcon) lightIcon.style.display = 'block';
                    if (darkIcon) darkIcon.style.display = 'none';
                } else {
                    if (lightIcon) lightIcon.style.display = 'none';
                    if (darkIcon) darkIcon.style.display = 'block';
                }

                // Handle text inversion for black backgrounds
                const blackBgCells = document.querySelectorAll('td[data-black-bg="true"]');
                blackBgCells.forEach(cell => {
                    const originalColor = cell.getAttribute('data-original-color');
                    if (isDarkMode) {
                        cell.style.color = invertColor(originalColor);
                    } else {
                        cell.style.color = originalColor;
                    }
                });
            });
        }

        // Expand toggle
        const toggleExpandButton = document.getElementById('toggleExpandButton');
        if (toggleExpandButton) {
            toggleExpandButton.addEventListener('click', () => {
                if (isEditMode) return;
                const state = toggleExpandButton.getAttribute('data-state') || 'default';
                if (state === 'default') {
                    toggleExpandButton.setAttribute('data-state', 'expanded');
                    setExpandedMode(true);
                } else {
                    toggleExpandButton.setAttribute('data-state', 'default');
                    setExpandedMode(false);
                }
            });
        }

        wireSettingsUI();
    }

    function populateSheetSelector() {
        const selector = document.getElementById('sheetSelector');
        if (!selector) return;

        selector.innerHTML = '';
        worksheetsData.forEach((ws, i) => {
            const opt = document.createElement('option');
            opt.value = String(i);
            opt.textContent = ws.name;
            selector.appendChild(opt);
        });
        selector.value = '0';
    }

    window.addEventListener('message', (event) => {
        const message = event.data;
        if (!message || typeof message !== 'object') return;

        if (message.command === 'initSettings') {
            applySettings(message.settings || {});
            return;
        }

        if (message.command === 'settingsUpdated') {
            applySettings(message.settings || {});
            return;
        }

        if (message.command === 'saveResult') {
            hideLoading();
            setLoadingText('Rendering worksheet...');
            isSaving = false;
            setButtonsEnabled(true);
            if (message.ok) {
                showToast('Saved');
                if (exitAfterSave) {
                    setEditMode(false);
                } else {
                    captureOriginalCellValues();
                }
            } else {
                showToast('Error saving');
            }
            return;
        }

        if (message.command === 'init') {
            worksheetsData = Array.isArray(message.worksheets) ? message.worksheets : [];
            currentWorksheet = 0;

            const rowHeaderWidth = typeof message.rowHeaderWidth === 'number' ? message.rowHeaderWidth : 60;
            document.documentElement.style.setProperty('--row-header-width', rowHeaderWidth + 'px');

            populateSheetSelector();
            attachHandlersOnce();
            const expandBtn = document.getElementById('toggleExpandButton');
            if (expandBtn) expandBtn.setAttribute('data-state', 'default');
            setExpandedMode(false);
            renderWorksheet(0);
        }
    });

    document.addEventListener('DOMContentLoaded', () => {
        vscode.postMessage({ command: 'webviewReady' });
    });
})();
