/* global acquireVsCodeApi */

(function () {
    const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : { postMessage: () => { } };

    // ===== Configuration =====
    const ROW_HEIGHT = 28;
    const BUFFER_ROWS = 20;
    const CHUNK_SIZE = 100;

    // ===== State =====
    let isTableView = true;
    let isEditMode = false;
    let isSaving = false;
    let exitAfterSave = false;

    // Virtual scrolling state
    let totalRows = 0;
    let columnCount = 0;
    let rowCache = new Map();
    let pendingRequests = new Map();
    let currentVisibleStart = 0;
    let currentVisibleEnd = 0;
    let isRequestingRows = false;
    let scrollDebounceTimer = null;

    // Selection state
    let isSelecting = false;
    let startCell = null;
    let endCell = null;
    const selectedCells = new Set();
    let activeCell = null;
    const selectedRows = new Set();
    const selectedColumns = new Set();
    let lastSelectedRow = null;
    let lastSelectedColumn = null;

    // Track selected row/column indices for full copy
    let selectedRowIndices = new Set();
    let selectedColumnIndices = new Set();

    // Undo/Redo for edit mode
    let undoStack = [];
    let redoStack = [];
    const MAX_HISTORY = 50;

    // Settings
    let currentSettings = {
        firstRowIsHeader: false,
        stickyToolbar: true,
        stickyHeader: false
    };

    // ===== Utilities =====
    function $(id) {
        return document.getElementById(id);
    }

    function normalizeCellText(text) {
        if (!text) return '';
        return String(text).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trimEnd();
    }

    function escapeCsvCell(value) {
        const v = value ?? '';
        const needsQuotes = /[",\n\r]/.test(v);
        if (!needsQuotes) return v;
        return '"' + v.replace(/"/g, '""') + '"';
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function setButtonsEnabled(enabled) {
        const ids = ['toggleViewButton', 'toggleTableEditButton', 'saveTableEditsButton',
            'cancelTableEditsButton', 'toggleBackgroundButton', 'toggleExpandButton'];
        ids.forEach((id) => {
            const el = $(id);
            if (el) el.disabled = !enabled;
        });
    }

    // ===== Virtual Scrolling Core =====

    function getTableContainer() {
        return $('tableContainer');
    }

    function requestRows(start, end, timeout = 10000) {
        return new Promise((resolve) => {
            const requestId = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
            pendingRequests.set(requestId, { resolve, start, end });

            vscode.postMessage({
                command: 'getRows',
                start,
                end,
                requestId
            });

            setTimeout(() => {
                if (pendingRequests.has(requestId)) {
                    console.warn(`Request ${requestId} timed out for rows ${start}-${end}`);
                    pendingRequests.delete(requestId);
                    resolve([]);
                }
            }, timeout);
        });
    }

    function requestAllRows() {
        // Use longer timeout for full data fetch (30 seconds)
        return requestRows(0, totalRows, 30000);
    }

    function createRowHtml(rowData, rowIndex) {
        let html = `<tr data-virtual-row="${rowIndex}">`;
        html += `<th class="row-header" data-row="${rowIndex}">${rowIndex + 1}</th>`;

        for (let colIndex = 0; colIndex < columnCount; colIndex++) {
            const cellContent = (rowData && rowData[colIndex]) ? rowData[colIndex].trim() : '';
            const isEmpty = cellContent === '';
            const displayContent = isEmpty ? '&nbsp;' : escapeHtml(cellContent);

            html += `<td data-row="${rowIndex}" data-col="${colIndex}" `;
            html += `data-default-bg="true" data-default-color="true"`;
            if (isEmpty) html += ` data-empty="true"`;
            if (isEditMode) html += ` contenteditable="true" spellcheck="false"`;
            html += `>${displayContent}</td>`;
        }

        html += '</tr>';
        return html;
    }

    function renderVirtualRows(startIndex, endIndex, rowsData) {
        const tbody = document.querySelector('#csv-table tbody');
        if (!tbody) return;

        rowsData.forEach((row, i) => {
            rowCache.set(startIndex + i, row);
        });

        const topSpacerHeight = startIndex * ROW_HEIGHT;
        const bottomSpacerHeight = Math.max(0, (totalRows - endIndex) * ROW_HEIGHT);

        let html = '';

        if (topSpacerHeight > 0) {
            html += `<tr class="virtual-spacer top-spacer"><td colspan="${columnCount + 1}" style="height: ${topSpacerHeight}px; padding: 0; border: none;"></td></tr>`;
        }

        for (let i = startIndex; i < endIndex; i++) {
            const rowData = rowCache.get(i) || [];
            html += createRowHtml(rowData, i);
        }

        if (bottomSpacerHeight > 0) {
            html += `<tr class="virtual-spacer bottom-spacer"><td colspan="${columnCount + 1}" style="height: ${bottomSpacerHeight}px; padding: 0; border: none;"></td></tr>`;
        }

        tbody.innerHTML = html;

        if (isEditMode) {
            applyEditModeToCells();
        }

        reapplySelection();
    }

    async function updateVisibleRows() {
        const container = getTableContainer();
        if (!container || totalRows === 0) return;

        const scrollTop = container.scrollTop;
        const clientHeight = container.clientHeight;

        const firstVisibleRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - BUFFER_ROWS);
        const lastVisibleRow = Math.min(
            totalRows,
            Math.ceil((scrollTop + clientHeight) / ROW_HEIGHT) + BUFFER_ROWS
        );

        const chunkStart = Math.floor(firstVisibleRow / CHUNK_SIZE) * CHUNK_SIZE;
        const chunkEnd = Math.min(totalRows, Math.ceil(lastVisibleRow / CHUNK_SIZE) * CHUNK_SIZE);

        if (chunkStart === currentVisibleStart && chunkEnd === currentVisibleEnd) {
            return;
        }

        let needsFetch = false;
        for (let i = chunkStart; i < chunkEnd; i++) {
            if (!rowCache.has(i)) {
                needsFetch = true;
                break;
            }
        }

        if (needsFetch && !isRequestingRows) {
            isRequestingRows = true;

            try {
                const rows = await requestRows(chunkStart, chunkEnd);

                if (rows && rows.length > 0) {
                    currentVisibleStart = chunkStart;
                    currentVisibleEnd = chunkStart + rows.length;
                    renderVirtualRows(chunkStart, chunkStart + rows.length, rows);
                }
            } finally {
                isRequestingRows = false;
            }
        } else if (!needsFetch) {
            currentVisibleStart = chunkStart;
            currentVisibleEnd = chunkEnd;

            const cachedRows = [];
            for (let i = chunkStart; i < chunkEnd; i++) {
                cachedRows.push(rowCache.get(i) || []);
            }
            renderVirtualRows(chunkStart, chunkEnd, cachedRows);
        }
    }

    function onScroll() {
        if (scrollDebounceTimer) {
            clearTimeout(scrollDebounceTimer);
        }
        scrollDebounceTimer = setTimeout(() => {
            updateVisibleRows();
        }, 16);
    }

    function initializeVirtualScrolling() {
        const container = getTableContainer();
        if (!container) return;

        container.addEventListener('scroll', onScroll, { passive: true });
        updateVisibleRows();
    }

    // ===== Selection =====

    function clearSelection() {
        document.querySelectorAll(
            'td.selected, td.active-cell, td.column-selected, td.row-selected, th.column-selected, th.row-selected'
        ).forEach((el) => {
            el.classList.remove('selected', 'active-cell', 'column-selected', 'row-selected', 'copying');
        });
        selectedCells.clear();
        selectedRows.clear();
        selectedColumns.clear();
        selectedRowIndices.clear();
        selectedColumnIndices.clear();
        activeCell = null;
        lastSelectedRow = null;
        lastSelectedColumn = null;
        const selectionInfo = $('selectionInfo');
        if (selectionInfo) selectionInfo.style.display = 'none';
    }

    function reapplySelection() {
        // Re-apply column selection
        selectedColumnIndices.forEach(colIdx => {
            document.querySelectorAll(`td[data-col="${colIdx}"], th[data-col="${colIdx}"]`).forEach((cell) => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell);
            });
        });

        // Re-apply row selection
        selectedRowIndices.forEach(rowIdx => {
            const rowHeader = document.querySelector(`th[data-row="${rowIdx}"]`);
            if (rowHeader && rowHeader.parentElement) {
                rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
        });

        // Re-apply active cell
        if (activeCell) {
            const row = activeCell.dataset?.row;
            const col = activeCell.dataset?.col;
            if (row !== undefined && col !== undefined) {
                const newCell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
                if (newCell) {
                    newCell.classList.add('active-cell');
                    activeCell = newCell;
                }
            }
        }
    }

    function getCellCoordinates(cell) {
        if (!cell || !cell.dataset) return null;
        return {
            row: parseInt(cell.dataset.row, 10),
            col: parseInt(cell.dataset.col, 10),
        };
    }

    function updateSelectionInfo() {
        const selectionInfo = $('selectionInfo');
        if (!selectionInfo) return;

        // For full column/row selection, show total counts
        if (selectedColumnIndices.size > 0 || selectedRowIndices.size > 0) {
            let rowCount = selectedRowIndices.size > 0 ? selectedRowIndices.size : totalRows;
            let colCount = selectedColumnIndices.size > 0 ? selectedColumnIndices.size : columnCount;

            if (selectedRowIndices.size > 0 && selectedColumnIndices.size === 0) {
                colCount = columnCount;
            }
            if (selectedColumnIndices.size > 0 && selectedRowIndices.size === 0) {
                rowCount = totalRows;
            }

            selectionInfo.textContent = rowCount + 'R × ' + colCount + 'C';
            selectionInfo.style.display = 'block';
        } else if (selectedCells.size > 1) {
            const cellsArray = Array.from(selectedCells);
            const rows = new Set(cellsArray.map((cell) => parseInt(cell.dataset.row, 10)));
            const cols = new Set(cellsArray.map((cell) => parseInt(cell.dataset.col, 10)));
            selectionInfo.textContent = rows.size + 'R × ' + cols.size + 'C';
            selectionInfo.style.display = 'block';
        } else {
            selectionInfo.style.display = 'none';
        }
    }

    function selectCellsInRange(start, end) {
        if (!start || !end) return;

        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);

        document.querySelectorAll('td.selected, td.active-cell').forEach((el) => {
            el.classList.remove('selected', 'active-cell');
        });
        selectedCells.clear();

        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            const coords = getCellCoordinates(cell);
            if (!coords) return;
            if (coords.row >= minRow && coords.row <= maxRow &&
                coords.col >= minCol && coords.col <= maxCol) {
                cell.classList.add('selected');
                selectedCells.add(cell);
            }
        });

        const startCellElement = document.querySelector(
            `td[data-row="${start.row}"][data-col="${start.col}"]`
        );
        if (startCellElement) {
            startCellElement.classList.add('active-cell');
            activeCell = startCellElement;
        }

        updateSelectionInfo();
    }

    function selectColumn(columnIndex, ctrlKey, shiftKey) {
        if (!ctrlKey && !shiftKey) clearSelection();

        if (shiftKey && lastSelectedColumn !== null) {
            if (!ctrlKey) clearSelection();
            const minCol = Math.min(lastSelectedColumn, columnIndex);
            const maxCol = Math.max(lastSelectedColumn, columnIndex);
            for (let col = minCol; col <= maxCol; col++) {
                selectedColumns.add(col);
                selectedColumnIndices.add(col);
                document.querySelectorAll(`td[data-col="${col}"], th[data-col="${col}"]`).forEach((cell) => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
        } else if (ctrlKey) {
            if (selectedColumns.has(columnIndex)) {
                selectedColumns.delete(columnIndex);
                selectedColumnIndices.delete(columnIndex);
                document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                    cell.classList.remove('column-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell);
                });
            } else {
                selectedColumns.add(columnIndex);
                selectedColumnIndices.add(columnIndex);
                document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
            lastSelectedColumn = columnIndex;
        } else {
            selectedColumns.add(columnIndex);
            selectedColumnIndices.add(columnIndex);
            document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell);
            });
            lastSelectedColumn = columnIndex;
        }
        updateSelectionInfo();
    }

    function selectRow(rowIndex, ctrlKey, shiftKey) {
        if (!ctrlKey && !shiftKey) clearSelection();

        if (shiftKey && lastSelectedRow !== null) {
            if (!ctrlKey) clearSelection();
            const minRow = Math.min(lastSelectedRow, rowIndex);
            const maxRow = Math.max(lastSelectedRow, rowIndex);
            for (let row = minRow; row <= maxRow; row++) {
                selectedRows.add(row);
                selectedRowIndices.add(row);
                const rowHeader = document.querySelector(`th[data-row="${row}"]`);
                if (rowHeader && rowHeader.parentElement) {
                    rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') selectedCells.add(cell);
                    });
                }
            }
        } else if (ctrlKey) {
            if (selectedRows.has(rowIndex)) {
                selectedRows.delete(rowIndex);
                selectedRowIndices.delete(rowIndex);
                const rowHeader = document.querySelector(`th[data-row="${rowIndex}"]`);
                if (rowHeader && rowHeader.parentElement) {
                    rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                        cell.classList.remove('row-selected');
                        if (cell.tagName === 'TD') selectedCells.delete(cell);
                    });
                }
            } else {
                selectedRows.add(rowIndex);
                selectedRowIndices.add(rowIndex);
                const rowHeader = document.querySelector(`th[data-row="${rowIndex}"]`);
                if (rowHeader && rowHeader.parentElement) {
                    rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') selectedCells.add(cell);
                    });
                }
            }
            lastSelectedRow = rowIndex;
        } else {
            selectedRows.add(rowIndex);
            selectedRowIndices.add(rowIndex);
            const rowHeader = document.querySelector(`th[data-row="${rowIndex}"]`);
            if (rowHeader && rowHeader.parentElement) {
                rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
            lastSelectedRow = rowIndex;
        }
        updateSelectionInfo();
    }

    // ===== Edit Mode =====

    function getTableData() {
        const data = [];
        for (let i = 0; i < totalRows; i++) {
            const cached = rowCache.get(i);
            if (cached) {
                data.push([...cached]);
            } else {
                data.push([]);
            }
        }
        return data;
    }

    function pushToUndo() {
        const currentData = getTableData();
        if (undoStack.length > 0) {
            if (JSON.stringify(undoStack[undoStack.length - 1]) === JSON.stringify(currentData)) {
                return;
            }
        }
        undoStack.push(currentData);
        if (undoStack.length > MAX_HISTORY) undoStack.shift();
        redoStack = [];
    }

    function undo() {
        if (undoStack.length <= 1) return;
        const current = undoStack.pop();
        redoStack.push(current);
        const previous = undoStack[undoStack.length - 1];

        previous.forEach((row, i) => {
            rowCache.set(i, row);
        });

        updateVisibleRows();
    }

    function redo() {
        if (redoStack.length === 0) return;
        const data = redoStack.pop();
        undoStack.push(data);

        data.forEach((row, i) => {
            rowCache.set(i, row);
        });

        updateVisibleRows();
    }

    function captureOriginalCellValues() {
        // Store original data from cache for cancel functionality
        window._originalCacheSnapshot = new Map();
        rowCache.forEach((value, key) => {
            window._originalCacheSnapshot.set(key, [...value]);
        });
    }

    function restoreOriginalCellValues() {
        if (window._originalCacheSnapshot) {
            rowCache = new Map(window._originalCacheSnapshot);
            window._originalCacheSnapshot = null;
            updateVisibleRows();
        }
    }

    function applyEditModeToCells() {
        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            if (isEditMode) {
                const current = normalizeCellText(cell.textContent || '');
                cell.textContent = current;
                cell.setAttribute('contenteditable', 'true');
                cell.setAttribute('spellcheck', 'false');
            } else {
                cell.removeAttribute('contenteditable');
                cell.removeAttribute('spellcheck');
                const current = normalizeCellText(cell.textContent || '');
                cell.textContent = current === '' ? '\u00a0' : current;
            }
        });
    }

    function setEditMode(enabled) {
        isEditMode = !!enabled;
        document.body.classList.toggle('edit-mode', isEditMode);

        const saveBtn = $('saveTableEditsButton');
        const cancelBtn = $('cancelTableEditsButton');
        const toggleViewBtn = $('toggleViewButton');
        const toggleTableEditBtn = $('toggleTableEditButton');
        const toggleExpandBtn = $('toggleExpandButton');

        if (toggleViewBtn) toggleViewBtn.classList.toggle('hidden', isEditMode);
        if (toggleTableEditBtn) toggleTableEditBtn.classList.toggle('hidden', isEditMode);
        if (toggleExpandBtn) toggleExpandBtn.classList.toggle('hidden', isEditMode);
        if (saveBtn) saveBtn.classList.toggle('hidden', !isEditMode);
        if (cancelBtn) cancelBtn.classList.toggle('hidden', !isEditMode);

        if (isEditMode) {
            clearSelection();
            captureOriginalCellValues();
            undoStack = [getTableData()];
            redoStack = [];
        } else {
            undoStack = [];
            redoStack = [];
        }

        applyEditModeToCells();
    }

    function serializeTableToCsv() {
        const rows = [];

        for (let i = 0; i < totalRows; i++) {
            const rowData = rowCache.get(i) || [];
            const row = [];
            for (let j = 0; j < columnCount; j++) {
                const value = normalizeCellText(rowData[j] || '');
                row.push(escapeCsvCell(value));
            }
            rows.push(row.join(','));
        }

        return rows.join('\n') + (rows.length ? '\n' : '');
    }

    function performSave(shouldExit = false) {
        if (isSaving || !isEditMode) return;
        isSaving = true;
        exitAfterSave = shouldExit;
        setButtonsEnabled(false);

        // Commit current edits to cache
        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            const row = parseInt(cell.dataset.row, 10);
            const col = parseInt(cell.dataset.col, 10);
            const value = normalizeCellText(cell.textContent || '');

            let rowData = rowCache.get(row);
            if (!rowData) {
                rowData = [];
                rowCache.set(row, rowData);
            }
            rowData[col] = value;
        });

        if (document.activeElement && document.activeElement.tagName === 'TD') {
            document.activeElement.blur();
        }

        clearSelection();

        if (window.getSelection) {
            window.getSelection().removeAllRanges();
        }

        const csvText = serializeTableToCsv();
        vscode.postMessage({ command: 'saveCsv', text: csvText });
    }

    // ===== Copy =====

    let isCopying = false;
    let copyOperationTimeout = null;

    function resetCopyState() {
        isCopying = false;
        if (copyOperationTimeout) {
            clearTimeout(copyOperationTimeout);
            copyOperationTimeout = null;
        }
    }

    async function copySelectionToClipboard() {
        // Prevent duplicate operations but with a safety check
        if (isCopying) {
            console.warn('Copy operation already in progress');
            return;
        }

        const hasFullColumnSelection = selectedColumnIndices.size > 0;
        const hasFullRowSelection = selectedRowIndices.size > 0;

        if (!hasFullColumnSelection && !hasFullRowSelection && selectedCells.size === 0) {
            return;
        }

        isCopying = true;

        // Safety timeout - reset state after 60 seconds max to prevent permanent lock
        copyOperationTimeout = setTimeout(() => {
            if (isCopying) {
                console.warn('Copy operation timed out, resetting state');
                resetCopyState();
                showToast('Copy timed out');
            }
        }, 60000);

        try {
            showToast('Copying...');

            let outputLines = [];

            if (hasFullColumnSelection || hasFullRowSelection) {
                // Need to fetch all rows for complete copy
                const allRows = await requestAllRows();

                // Validate we got data back - don't corrupt cache with empty data
                if (!allRows || allRows.length === 0) {
                    showToast('Failed to fetch data');
                    return;
                }

                // Only cache if we got the expected amount of data
                if (allRows.length >= totalRows * 0.9) { // Allow some tolerance
                    allRows.forEach((row, i) => {
                        rowCache.set(i, row);
                    });
                }

                const rowCount = allRows.length;

                if (hasFullColumnSelection && !hasFullRowSelection) {
                    // Copy entire columns
                    const sortedCols = Array.from(selectedColumnIndices).sort((a, b) => a - b);

                    for (let r = 0; r < rowCount; r++) {
                        const rowData = allRows[r] || [];
                        const lineParts = sortedCols.map(c => rowData[c] || '');
                        outputLines.push(lineParts.join('\t'));
                    }
                } else if (hasFullRowSelection && !hasFullColumnSelection) {
                    // Copy entire rows
                    const sortedRows = Array.from(selectedRowIndices).sort((a, b) => a - b);

                    for (const r of sortedRows) {
                        if (r < rowCount) {
                            const rowData = allRows[r] || [];
                            const lineParts = [];
                            for (let c = 0; c < columnCount; c++) {
                                lineParts.push(rowData[c] || '');
                            }
                            outputLines.push(lineParts.join('\t'));
                        }
                    }
                } else {
                    // Both rows and columns selected - intersection
                    const sortedRows = Array.from(selectedRowIndices).sort((a, b) => a - b);
                    const sortedCols = Array.from(selectedColumnIndices).sort((a, b) => a - b);

                    for (const r of sortedRows) {
                        if (r < rowCount) {
                            const rowData = allRows[r] || [];
                            const lineParts = sortedCols.map(c => rowData[c] || '');
                            outputLines.push(lineParts.join('\t'));
                        }
                    }
                }

                const cellCount = hasFullColumnSelection ?
                    rowCount * selectedColumnIndices.size :
                    (hasFullRowSelection ? selectedRowIndices.size * columnCount : 0);

                const tsv = outputLines.join('\n');
                
                const writeSuccess = await writeToClipboardAsync(tsv);
                if (!writeSuccess) {
                    showToast('Copy failed');
                    return;
                }

                // Flash visible selected cells
                selectedCells.forEach(cell => cell.classList.add('copying'));
                setTimeout(() => {
                    selectedCells.forEach(cell => cell.classList.remove('copying'));
                }, 300);

                showToast('Copied ' + cellCount + ' cells');
            } else {
                // Regular cell selection - use cached data
                const cellsArray = Array.from(selectedCells);
                const rowSet = new Set();
                const colSet = new Set();

                cellsArray.forEach(td => {
                    const r = parseInt(td.dataset.row, 10);
                    const c = parseInt(td.dataset.col, 10);
                    if (!isNaN(r) && !isNaN(c)) {
                        rowSet.add(r);
                        colSet.add(c);
                    }
                });

                const sortedRows = Array.from(rowSet).sort((a, b) => a - b);
                const sortedCols = Array.from(colSet).sort((a, b) => a - b);

                for (const r of sortedRows) {
                    const rowData = rowCache.get(r) || [];
                    const lineParts = sortedCols.map(c => rowData[c] || '');
                    outputLines.push(lineParts.join('\t'));
                }

                const tsv = outputLines.join('\n');
                
                const writeSuccess = await writeToClipboardAsync(tsv);
                if (!writeSuccess) {
                    showToast('Copy failed');
                    return;
                }

                selectedCells.forEach(cell => cell.classList.add('copying'));
                setTimeout(() => {
                    selectedCells.forEach(cell => cell.classList.remove('copying'));
                }, 300);

                showToast('Copied ' + cellsArray.length + ' cells');
            }
        } catch (err) {
            console.error('Copy operation failed:', err);
            showToast('Copy failed');
        } finally {
            resetCopyState();
        }
    }

    async function writeToClipboardAsync(text) {
        // Add size check for very large copies
        if (text.length > 10 * 1024 * 1024) { // 10MB warning
            console.warn('Large clipboard operation:', (text.length / 1024 / 1024).toFixed(2), 'MB');
        }

        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            try {
                await navigator.clipboard.writeText(text);
                return true;
            } catch (e) {
                console.warn('Clipboard API failed, trying fallback:', e.message);
            }
        }
        
        try {
            await execCommandFallback(text);
            return true;
        } catch (e) {
            console.error('All clipboard methods failed:', e);
            return false;
        }
    }

    function execCommandFallback(text) {
        return new Promise((resolve, reject) => {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.setAttribute('readonly', '');
            textarea.style.cssText = `
                position: fixed; left: -9999px; top: 0;
                width: 2px; height: 2px; padding: 0;
                border: none; outline: none; opacity: 0;
            `;
            document.body.appendChild(textarea);
            
            // Use setTimeout to ensure DOM is ready
            setTimeout(() => {
                try {
                    textarea.focus();
                    textarea.select();
                    textarea.setSelectionRange(0, text.length);
                    const successful = document.execCommand('copy');
                    document.body.removeChild(textarea);
                    successful ? resolve() : reject(new Error('execCommand failed'));
                } catch (err) {
                    try { document.body.removeChild(textarea); } catch {}
                    reject(err);
                }
            }, 0);
        });
    }

    // ===== UI Helpers =====

    function showToast(message) {
        let toast = $('saveToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'saveToast';
            toast.className = 'toast-notification';
            toast.innerHTML = `
                <div class="toast-icon-wrapper">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
                        <polyline points="20 6 9 17 4 12"></polyline>
                    </svg>
                </div>
                <span class="toast-text"></span>
            `;
            document.body.appendChild(toast);
        }
        toast.querySelector('.toast-text').textContent = message;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 2000);
    }

    function adjustColumnWidths(mode) {
        try {
            const table = $('csv-table');
            if (!table) return;
            const colGroup = table.querySelector('colgroup');
            if (!colGroup) return;

            const headerCells = table.querySelectorAll('th.col-header');
            if (headerCells.length === 0) return;

            table.style.tableLayout = 'auto';
            const ctx = document.createElement('canvas').getContext('2d');
            ctx.font = '13px "Segoe UI", Roboto, Helvetica, Arial, sans-serif';

            const visibleRows = table.querySelectorAll('tbody tr:not(.virtual-spacer)');
            const limit = Math.min(visibleRows.length, 50);

            // Measure first (row header) column so it can grow to fit content (min 30px)
            let firstColMax = 30;
            for (let r = 0; r < limit; r++) {
                const row = visibleRows[r];
                const cell = row && row.children && row.children[0];
                if (cell) {
                    const width = ctx.measureText(cell.textContent.trim()).width + 24; // include padding
                    if (width > firstColMax) firstColMax = width;
                }
            }

            let colGroupHtml = `<col style="width: ${firstColMax}px;">`;
            headerCells.forEach((th, index) => {
                let maxWidth = ctx.measureText(th.textContent.trim()).width + 32;
                for (let r = 0; r < limit; r++) {
                    const row = visibleRows[r];
                    const cell = row.children[index + 1];
                    if (cell) {
                        const width = ctx.measureText(cell.textContent.trim()).width + 32;
                        if (width > maxWidth) maxWidth = width;
                    }
                }
                const finalWidth = mode === 'expand' ? maxWidth : Math.min(maxWidth, 180);
                colGroupHtml += `<col style="width: ${finalWidth}px; max-width: ${mode === 'expand' ? 'none' : '180px'};">`;
            });

            colGroup.innerHTML = colGroupHtml;
            table.style.tableLayout = 'fixed';
            /* Keep table intrinsic so it doesn't expand to fill the viewport */
            table.style.width = 'max-content';
        } catch (e) {
            console.error('Error adjusting columns:', e);
        }
    }

    // ===== Floating Tooltips =====

    function setupFloatingTooltips() {
        let activeTip = null;
        let activeTrigger = null;

        function positionTip(trigger, tip) {
            if (!trigger || !tip) return;

            tip.dataset.origPosition = tip.style.position || '';
            tip.dataset.origLeft = tip.style.left || '';
            tip.dataset.origTop = tip.style.top || '';
            tip.dataset.origTransform = tip.style.transform || '';

            tip.style.position = 'fixed';
            tip.style.visibility = 'visible';
            tip.style.opacity = '0';
            tip.style.pointerEvents = 'none';

            requestAnimationFrame(() => {
                const r = trigger.getBoundingClientRect();
                const tr = tip.getBoundingClientRect();
                let left = r.left + r.width / 2 - tr.width / 2;
                left = Math.max(8, Math.min(left, window.innerWidth - tr.width - 8));
                let top = r.bottom + 8;
                if (top + tr.height > window.innerHeight - 8) {
                    top = r.top - tr.height - 8;
                }
                tip.style.left = left + 'px';
                tip.style.top = top + 'px';
                tip.style.transform = 'translateX(0) translateY(0)';
                tip.style.opacity = '1';
                tip.style.pointerEvents = 'auto';
                activeTip = tip;
                activeTrigger = trigger;
            });

            function onTipEnter() { }
            function onTipLeave() {
                hideTip(trigger);
            }
            tip.addEventListener('mouseenter', onTipEnter, { once: true });
            tip.addEventListener('mouseleave', onTipLeave, { once: true });
        }

        function hideTip(trigger) {
            const tip = trigger.querySelector('.tooltiptext');
            if (!tip) return;
            tip.style.opacity = '0';
            tip.style.pointerEvents = 'none';
            tip.style.visibility = '';
            tip.style.left = tip.dataset.origLeft || '';
            tip.style.top = tip.dataset.origTop || '';
            tip.style.position = tip.dataset.origPosition || '';
            tip.style.transform = tip.dataset.origTransform || '';
            activeTip = null;
            activeTrigger = null;
        }

        document.addEventListener('mouseover', (e) => {
            const t = e.target.closest('.tooltip');
            if (!t) return;
            const tip = t.querySelector('.tooltiptext');
            if (!tip) return;
            positionTip(t, tip);
        });

        document.addEventListener('mouseout', (e) => {
            const t = e.target.closest('.tooltip');
            if (!t) return;
            const rel = e.relatedTarget;
            if (rel) {
                if (t.contains(rel)) return;
                if (activeTip && activeTip.contains(rel)) return;
            }
            hideTip(t);
        });

        window.addEventListener('resize', () => {
            if (activeTip && activeTrigger) positionTip(activeTrigger, activeTip);
        });
        window.addEventListener('resize', updateHeaderHeight);
        window.addEventListener('scroll', () => {
            if (activeTip && activeTrigger) positionTip(activeTrigger, activeTip);
        }, true);
    }

    // ===== Toolbar Scroll Sync =====

    function syncToolbarScroll() {
        const area = $('buttonScrollArea');
        const bar = $('buttonScrollbar');
        const inner = $('scrollInner');
        if (!area || !bar || !inner) return;

        inner.style.width = area.scrollWidth + 'px';
        bar.scrollLeft = area.scrollLeft;

        if (!area._scrollWire) {
            let syncing = false;
            area.addEventListener('scroll', () => {
                if (syncing) return;
                syncing = true;
                bar.scrollLeft = area.scrollLeft;
                setTimeout(() => (syncing = false), 20);
            });
            bar.addEventListener('scroll', () => {
                if (syncing) return;
                syncing = true;
                area.scrollLeft = bar.scrollLeft;
                setTimeout(() => (syncing = false), 20);
            });
            area._scrollWire = true;
        }
    }

    // ===== Settings =====

    function applySettings(settings, saveLocal = false) {
        currentSettings = settings || {};
        if (!settings) return;

        document.body.classList.toggle('first-row-as-header', !!settings.firstRowIsHeader);
        document.body.classList.toggle('sticky-header-enabled', !!settings.stickyHeader);
        document.body.classList.toggle('sticky-toolbar-enabled', !!settings.stickyToolbar);

        const chkHeader = $('chkHeaderRow');
        const chkSticky = $('chkStickyHeader');
        const chkToolbar = $('chkStickyToolbar');

        if (chkHeader) chkHeader.checked = !!settings.firstRowIsHeader;
        if (chkSticky) chkSticky.checked = !!settings.stickyHeader;
        if (chkToolbar) chkToolbar.checked = !!settings.stickyToolbar;

        // Bold first row when firstRowIsHeader is enabled
        const table = $('csv-table');
        if (table) {
            const firstRow = table.querySelector('tbody tr:not(.virtual-spacer)');
            if (firstRow) {
                if (settings.firstRowIsHeader) {
                    firstRow.classList.add('header-row');
                } else {
                    firstRow.classList.remove('header-row');
                }
            }
        }

        // Update toolbar stickiness
        const container = document.querySelector('.toolbar');
        const content = $('content');
        const scrollArea = document.querySelector('.table-scroll');
        const headerBg = document.querySelector('.header-background');

        if (container) {
            if (settings.stickyToolbar) {
                if (container.parentNode !== document.body) {
                    if (content && content.parentNode) document.body.insertBefore(container, content);
                    else document.body.appendChild(container);
                }
                container.classList.remove('not-sticky');
                container.classList.add('expanded-toolbar');
                if (headerBg) headerBg.style.display = '';
            } else {
                const target = scrollArea || content;
                if (target && container.parentNode !== target) {
                    target.insertBefore(container, target.firstChild);
                }
                container.classList.add('not-sticky');
                container.classList.remove('expanded-toolbar');
                if (headerBg) headerBg.style.display = 'none';
            }
        }
        setTimeout(updateHeaderHeight, 0);

        if (chkSticky) chkSticky.disabled = !chkHeader?.checked;

        if (saveLocal) {
            vscode.postMessage({ command: 'updateSettings', settings });
        }
    }

    function wireSettingsUI() {
        const openBtn = $('openSettingsButton');
        const panel = $('settingsPanel');
        const chkH = $('chkHeaderRow');
        const chkSH = $('chkStickyHeader');
        const chkST = $('chkStickyToolbar');
        const cancelBtn = $('settingsCancelButton');

        if (!openBtn || !panel) return;

        let snapshot = null;
        let repositionHandlers = null;
        let panelOriginalParent = null;
        let panelOriginalNext = null;

        function repositionPanel() {
            const container = document.querySelector('.toolbar');
            if (!container) return;
            const rect = container.getBoundingClientRect();
            // Use fixed positioning anchored to toolbar; explicitly override CSS right to allow custom width
            panel.style.position = 'fixed';
            panel.style.left = Math.max(8, rect.left) + 'px';
            panel.style.top = rect.bottom + 'px';
            panel.style.right = 'auto';
            const maxWidth = Math.min(window.innerWidth - 16, rect.width);
            panel.style.width = Math.max(280, maxWidth) + 'px';
            panel.style.zIndex = '10001';
        }

        function openPanel() {
            snapshot = {
                firstRowIsHeader: chkH?.checked,
                stickyHeader: chkSH?.checked,
                stickyToolbar: chkST?.checked
            };
            // Save original parent so we can restore later. Move panel to document.body so it is not affected
            // by toolbar descendant CSS rules when toolbar is sticky.
            if (!panelOriginalParent) {
                panelOriginalParent = panel.parentNode;
                panelOriginalNext = panel.nextSibling;
            }
            if (panel.parentNode !== document.body) {
                document.body.appendChild(panel);
            }

            panel.classList.remove('hidden');
            panel.classList.add('floating');
            panel.setAttribute('aria-hidden', 'false');
            document.body.classList.add('settings-open');

            const container = document.querySelector('.toolbar');
            if (container) {
                container.classList.add('settings-open');
                // only expand the toolbar vertically when the toolbar is configured to be sticky
                if (document.body.classList.contains('sticky-toolbar-enabled')) {
                    container.classList.add('expanded-toolbar');
                }
            }

            repositionPanel();
            updateHeaderHeight();
            repositionHandlers = () => {
                repositionPanel();
                updateHeaderHeight();
            };
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
                const cfgSticky = chkST && chkST.checked;
                if (!cfgSticky) container.classList.remove('expanded-toolbar');
            }

            panel.style.position = '';
            panel.style.left = '';
            panel.style.top = '';
            panel.style.width = '';
            panel.style.right = '';
            panel.style.zIndex = '';

            // Restore original parent/position so stylesheet rules apply again
            if (panelOriginalParent && panelOriginalParent !== panel.parentNode) {
                try {
                    panelOriginalParent.insertBefore(panel, panelOriginalNext);
                } catch (e) {
                    // fallback: append
                    panelOriginalParent.appendChild(panel);
                }
            }

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

        function onChange() {
            const s = {
                firstRowIsHeader: !!chkH?.checked,
                stickyHeader: !!chkSH?.checked,
                stickyToolbar: !!chkST?.checked
            };
            if (!s.firstRowIsHeader) s.stickyHeader = false;
            applySettings(s, true);
        }

        if (chkH) chkH.addEventListener('change', () => {
            if (chkSH) chkSH.disabled = !chkH.checked;
            if (!chkH.checked && chkSH) chkSH.checked = false;
            onChange();
        });
        if (chkSH) chkSH.addEventListener('change', onChange);
        if (chkST) chkST.addEventListener('change', onChange);

        if (cancelBtn) cancelBtn.addEventListener('click', closePanel);

        document.addEventListener('click', (e) => {
            if (!panel.classList.contains('hidden')) {
                if (!e.target.closest('.settings-panel') && !e.target.closest('#openSettingsButton')) {
                    closePanel();
                }
            }
        });
    }

    function updateHeaderHeight() {
        const container = document.querySelector('.toolbar');
        if (!container) return;
        // Respect non-sticky mode (stylesheet sets --header-height for non-sticky)
        if (!document.body.classList.contains('sticky-toolbar-enabled')) {
            document.documentElement.style.removeProperty('--header-height');
            const headerBg = document.querySelector('.header-background');
            if (headerBg) headerBg.style.height = '';
            return;
        }
        let h = Math.max(6, Math.ceil(container.getBoundingClientRect().height));
        // Cap header height to avoid large gaps when toolbar wraps or settings are open
        const maxH = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--header-height-max')) || 96;
        h = Math.min(h, maxH);
        document.documentElement.style.setProperty('--header-height', h + 'px');
        const headerBg = document.querySelector('.header-background');
        if (headerBg) headerBg.style.height = h + 'px';
    }

    // ===== Event Handlers =====

    function initializeSelection() {
        const table = $('csv-table');
        if (!table || table.dataset.listenersAdded === 'true') return;
        table.dataset.listenersAdded = 'true';

        table.addEventListener('focusout', (e) => {
            if (isEditMode && e.target.tagName === 'TD') {
                const row = parseInt(e.target.dataset.row, 10);
                const col = parseInt(e.target.dataset.col, 10);
                const value = normalizeCellText(e.target.textContent || '');

                let rowData = rowCache.get(row);
                if (!rowData) {
                    rowData = [];
                    rowCache.set(row, rowData);
                }
                rowData[col] = value;

                pushToUndo();
            }
        });

        table.addEventListener('mousedown', (e) => {
            if (isEditMode) return;
            const target = e.target.closest('td, th');
            if (!target) return;
            e.preventDefault();

            if (target.classList.contains('col-header')) {
                const colIdx = parseInt(target.dataset.col, 10);
                if (!e.shiftKey) lastSelectedColumn = colIdx;
                selectColumn(colIdx, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }
            if (target.classList.contains('row-header')) {
                const rowIdx = parseInt(target.dataset.row, 10);
                if (!e.shiftKey) lastSelectedRow = rowIdx;
                selectRow(rowIdx, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }
            if (target.tagName === 'TD') {
                const coords = getCellCoordinates(target);
                if (!coords) return;

                // Clear column/row selection when selecting individual cells
                selectedColumnIndices.clear();
                selectedRowIndices.clear();

                if (e.ctrlKey || e.metaKey) {
                    e.stopPropagation();
                    if (target.classList.contains('selected')) {
                        target.classList.remove('selected');
                        selectedCells.delete(target);
                        if (target === activeCell) activeCell = null;
                    } else {
                        target.classList.add('selected');
                        selectedCells.add(target);
                        if (activeCell) activeCell.classList.remove('active-cell');
                        target.classList.add('active-cell');
                        activeCell = target;
                        startCell = coords;
                    }
                } else if (e.shiftKey && startCell) {
                    e.stopPropagation();
                    selectCellsInRange(startCell, coords);
                } else {
                    clearSelection();
                    isSelecting = true;
                    startCell = coords;
                    endCell = coords;
                    target.classList.add('selected', 'active-cell');
                    selectedCells.add(target);
                    activeCell = target;
                }
                updateSelectionInfo();
            }
        });

        table.addEventListener('mousemove', (e) => {
            if (isEditMode || !isSelecting || !startCell) return;
            const target = e.target.closest('td');
            if (!target) return;
            const coords = getCellCoordinates(target);
            if (!coords || (endCell && coords.row === endCell.row && coords.col === endCell.col)) return;
            endCell = coords;
            selectCellsInRange(startCell, endCell);
        });

        document.addEventListener('mouseup', () => { isSelecting = false; });

        document.addEventListener('keydown', (e) => {
            const isCmdOrCtrl = e.ctrlKey || e.metaKey;

            if (isEditMode && isCmdOrCtrl) {
                if (e.key.toLowerCase() === 'z') {
                    e.preventDefault();
                    if (e.shiftKey) redo();
                    else undo();
                    return;
                }
                if (e.key.toLowerCase() === 'y') {
                    e.preventDefault();
                    redo();
                    return;
                }
            }

            if (isCmdOrCtrl && e.key.toLowerCase() === 's') {
                e.preventDefault();
                performSave(false);
                return;
            }

            if (isEditMode) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    const active = document.activeElement;
                    const coords = getCellCoordinates(active);
                    if (coords) {
                        const nextCell = document.querySelector(
                            `td[data-row="${coords.row + 1}"][data-col="${coords.col}"]`
                        );
                        if (nextCell) {
                            nextCell.focus();
                            const range = document.createRange();
                            const sel = window.getSelection();
                            range.selectNodeContents(nextCell);
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
            } else if (isCmdOrCtrl && e.key.toLowerCase() === 'a') {
                e.preventDefault();
                clearSelection();

                // Select all - mark all columns as selected
                for (let c = 0; c < columnCount; c++) {
                    selectedColumnIndices.add(c);
                }

                const all = document.querySelectorAll('td[data-row][data-col]');
                all.forEach(c => {
                    c.classList.add('selected');
                    selectedCells.add(c);
                });
                if (all[0]) {
                    all[0].classList.add('active-cell');
                    activeCell = all[0];
                    startCell = getCellCoordinates(all[0]);
                }
                updateSelectionInfo();
            } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && activeCell) {
                const coords = getCellCoordinates(activeCell);
                let nr = coords.row, nc = coords.col;
                if (e.key === 'ArrowUp' && nr > 0) nr--;
                else if (e.key === 'ArrowDown') nr++;
                else if (e.key === 'ArrowLeft' && nc > 0) nc--;
                else if (e.key === 'ArrowRight') nc++;

                const next = document.querySelector(`td[data-row="${nr}"][data-col="${nc}"]`);
                if (next) {
                    e.preventDefault();
                    if (e.shiftKey) {
                        selectCellsInRange(startCell || coords, { row: nr, col: nc });
                    } else {
                        clearSelection();
                        next.classList.add('selected', 'active-cell');
                        selectedCells.add(next);
                        activeCell = next;
                        startCell = { row: nr, col: nc };
                    }
                    updateSelectionInfo();
                }
            }
        });

        document.addEventListener('click', (e) => {
            if (!e.target.closest('#csv-table') && !e.target.closest('.toolbar')) {
                clearSelection();
            }
        });
    }

    // ===== Message Handler =====

    window.addEventListener('message', (event) => {
        const m = event.data;

        switch (m.command) {
            case 'initVirtualTable':
                const loading = $('loadingIndicator');
                if (loading) loading.style.display = 'none';

                totalRows = m.totalRows || 0;
                columnCount = m.columnCount || 0;

                const thead = document.querySelector('#csv-table thead');
                if (thead) thead.innerHTML = m.headerHtml || '';

                const table = $('csv-table');
                if (table) {
                    const colgroup = table.querySelector('colgroup');
                    if (colgroup) {
                        let colHtml = '<col style="width: 30px;">';
                        for (let i = 0; i < columnCount; i++) {
                            colHtml += '<col style="width: 150px;">';
                        }
                        colgroup.innerHTML = colHtml;
                    }
                }

                initializeVirtualScrolling();
                initializeSelection();

                setTimeout(() => {
                    adjustColumnWidths('default');
                    syncToolbarScroll();
                    applySettings(currentSettings, false);
                }, 200);
                break;

            case 'rowsData':
                if (m.requestId && pendingRequests.has(m.requestId)) {
                    const { resolve } = pendingRequests.get(m.requestId);
                    pendingRequests.delete(m.requestId);
                    resolve(m.rows || []);
                }
                break;

            case 'rowCount':
                totalRows = m.totalRows || 0;
                if (m.requestId && pendingRequests.has(m.requestId)) {
                    const { resolve } = pendingRequests.get(m.requestId);
                    pendingRequests.delete(m.requestId);
                    resolve(totalRows);
                }
                break;

            case 'initSettings':
            case 'settingsUpdated':
                applySettings(m.settings, false);
                setTimeout(syncToolbarScroll, 20);
                break;

            case 'saveResult':
                isSaving = false;
                setButtonsEnabled(true);
                if (m.ok) {
                    showToast('Saved');
                    window._originalCacheSnapshot = null;
                    if (exitAfterSave) {
                        setEditMode(false);
                    } else {
                        captureOriginalCellValues();
                    }
                } else {
                    showToast('Error saving');
                }
                break;
        }

        // Handle theme messages
        if (m.type === 'setTheme') {
            // Theme manager will handle this
        }
    });

    // ===== Button Handlers =====

    function wireButtons() {
        const btnMap = {
            toggleViewButton: () => {
                isTableView = !isTableView;
                vscode.postMessage({ command: 'toggleView', isTableView });
            },
            toggleTableEditButton: () => setEditMode(!isEditMode),
            saveTableEditsButton: () => performSave(true),
            cancelTableEditsButton: () => {
                restoreOriginalCellValues();
                setEditMode(false);
            },
            toggleExpandButton: () => {
                const btn = $('toggleExpandButton');
                const state = btn?.getAttribute('data-state') || 'default';
                if (state === 'default') {
                    btn?.setAttribute('data-state', 'expanded');
                    document.body.classList.add('expanded-mode');
                    $('expandIcon').style.display = 'none';
                    $('collapseIcon').style.display = 'block';
                    $('expandButtonText').textContent = 'Default';
                    adjustColumnWidths('expand');
                } else {
                    btn?.setAttribute('data-state', 'default');
                    document.body.classList.remove('expanded-mode');
                    $('expandIcon').style.display = 'block';
                    $('collapseIcon').style.display = 'none';
                    $('expandButtonText').textContent = 'Expand';
                    adjustColumnWidths('default');
                }
            }
        };

        // Theme manager
        if (typeof ThemeManager !== 'undefined') {
            new ThemeManager('toggleBackgroundButton', {
                onBeforeCycle: () => !isEditMode
            }, vscode);
        }

        Object.entries(btnMap).forEach(([id, handler]) => {
            const el = $(id);
            if (el) el.addEventListener('click', handler);
        });
    }

    // ===== Initialize =====

    setupFloatingTooltips();
    wireButtons();
    wireSettingsUI();
    updateHeaderHeight();
    vscode.postMessage({ command: 'webviewReady' });
})();