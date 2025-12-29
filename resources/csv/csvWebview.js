/* global acquireVsCodeApi */

(function () {
    const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : { postMessage: () => { } };

    let isTableView = true;
    let isEditMode = false;
    let isSaving = false;

    let isSelecting = false;
    let startCell = null;
    let endCell = null;
    const selectedCells = new Set();
    let activeCell = null;
    const selectedRows = new Set();
    const selectedColumns = new Set();
    let lastSelectedRow = null;
    let lastSelectedColumn = null;

    // Undo/Redo Stacks
    let undoStack = [];
    let redoStack = [];
    const MAX_HISTORY = 50;

    function $(id) {
        return document.getElementById(id);
    }

    function normalizeCellText(text) {
        if (!text) return '';
        return String(text).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trimEnd();
    }

    function setButtonsEnabled(enabled) {
        const ids = ['toggleViewButton', 'toggleTableEditButton', 'saveTableEditsButton', 'cancelTableEditsButton', 'toggleBackgroundButton', 'toggleExpandButton'];
        ids.forEach((id) => {
            const el = $(id);
            if (el) el.disabled = !enabled;
        });
    }

    function clearSelection() {
        document
            .querySelectorAll(
                'td.selected, td.active-cell, td.column-selected, td.row-selected, th.column-selected, th.row-selected'
            )
            .forEach((el) => {
                el.classList.remove('selected', 'active-cell', 'column-selected', 'row-selected', 'copying');
            });
        selectedCells.clear();
        selectedRows.clear();
        selectedColumns.clear();
        activeCell = null;
        lastSelectedRow = null;
        lastSelectedColumn = null;
        const selectionInfo = $('selectionInfo');
        if (selectionInfo) selectionInfo.style.display = 'none';
    }

    // --- History Management ---

    function getTableData() {
        const data = [];
        const trList = document.querySelectorAll('#csv-table tbody tr');
        trList.forEach(tr => {
            const rowData = [];
            tr.querySelectorAll('td[data-row][data-col]').forEach(td => {
                rowData.push(td.textContent);
            });
            data.push(rowData);
        });
        return data;
    }

    function setTableData(data) {
        if (!data) return;
        const trList = document.querySelectorAll('#csv-table tbody tr');
        data.forEach((rowData, rIdx) => {
            const tr = trList[rIdx];
            if (tr) {
                const cells = tr.querySelectorAll('td[data-row][data-col]');
                rowData.forEach((val, cIdx) => {
                    if (cells[cIdx]) {
                        cells[cIdx].textContent = val;
                    }
                });
            }
        });
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
        setTableData(previous);
    }

    function redo() {
        if (redoStack.length === 0) return;
        const data = redoStack.pop();
        undoStack.push(data);
        setTableData(data);
    }

    // ---------------------------

    function captureOriginalCellValues() {
        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            if (cell.dataset.original !== undefined) return;
            cell.dataset.original = normalizeCellText(cell.textContent || '');
        });
    }

    function clearOriginalCellValues() {
        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            delete cell.dataset.original;
        });
    }

    function restoreOriginalCellValues() {
        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            if (cell.dataset.original === undefined) return;
            const value = cell.dataset.original;
            cell.textContent = value === '' ? '\u00a0' : value;
        });
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
        setTimeout(() => {
            toast.classList.remove('show');
        }, 2000);
    }

    function escapeCsvCell(value) {
        const v = value ?? '';
        const needsQuotes = /[",\n\r]/.test(v);
        if (!needsQuotes) return v;
        return '"' + v.replace(/"/g, '""') + '"';
    }

    function serializeTableToCsv() {
        const headerCols = document.querySelectorAll('th.col-header');
        const colCount = headerCols.length;
        const rows = [];
        const trList = document.querySelectorAll('#csv-table tbody tr');

        trList.forEach((tr) => {
            const row = [];
            for (let c = 0; c < colCount; c++) {
                const td = tr.querySelector('td[data-col="' + c + '"]');
                const value = normalizeCellText(td ? td.textContent || '' : '');
                row.push(escapeCsvCell(value));
            }
            rows.push(row.join(','));
        });

        return rows.join('\n') + (rows.length ? '\n' : '');
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

        if (selectedCells.size > 1) {
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

        const ctrlSelectedCells = new Set();
        selectedCells.forEach(cell => {
            const coords = getCellCoordinates(cell);
            if (coords) {
                const inRange = coords.row >= Math.min(start.row, end.row) &&
                    coords.row <= Math.max(start.row, end.row) &&
                    coords.col >= Math.min(start.col, end.col) &&
                    coords.col <= Math.max(start.col, end.col);
                if (!inRange) ctrlSelectedCells.add(cell);
            }
        });

        document.querySelectorAll('td.selected, td.active-cell').forEach((el) => {
            if (!ctrlSelectedCells.has(el)) {
                el.classList.remove('selected', 'active-cell');
                selectedCells.delete(el);
            }
        });

        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);

        document.querySelectorAll('td[data-row][data-col]').forEach((cell) => {
            const coords = getCellCoordinates(cell);
            if (!coords) return;
            if (coords.row >= minRow && coords.row <= maxRow && coords.col >= minCol && coords.col <= maxCol) {
                cell.classList.add('selected');
                selectedCells.add(cell);
            }
        });

        const startCellElement = document.querySelector('td[data-row="' + start.row + '"][data-col="' + start.col + '"]');
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
                document.querySelectorAll('td[data-col="' + col + '"], th[data-col="' + col + '"]').forEach((cell) => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
        } else if (ctrlKey) {
            if (selectedColumns.has(columnIndex)) {
                selectedColumns.delete(columnIndex);
                document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]').forEach((cell) => {
                    cell.classList.remove('column-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell);
                });
            } else {
                selectedColumns.add(columnIndex);
                document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]').forEach((cell) => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell);
                });
            }
            lastSelectedColumn = columnIndex;
        } else {
            selectedColumns.add(columnIndex);
            document.querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]').forEach((cell) => {
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
                const rowHeader = document.querySelector('th[data-row="' + row + '"]');
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
                const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
                if (rowHeader && rowHeader.parentElement) {
                    rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                        cell.classList.remove('row-selected');
                        if (cell.tagName === 'TD') selectedCells.delete(cell);
                    });
                }
            } else {
                selectedRows.add(rowIndex);
                const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
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
            const rowHeader = document.querySelector('th[data-row="' + rowIndex + '"]');
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

            const rowsToCheck = table.querySelectorAll('tbody tr');
            const limit = Math.min(rowsToCheck.length, 50);

            let colGroupHtml = '<col style="width: 30px;">';
            headerCells.forEach((th, index) => {
                let maxWidth = ctx.measureText(th.textContent.trim()).width + 32;
                for (let r = 0; r < limit; r++) {
                    const row = rowsToCheck[r];
                    const cell = row.children[index + 1];
                    if (cell) {
                        const width = ctx.measureText(cell.textContent.trim()).width + 32;
                        if (width > maxWidth) maxWidth = width;
                    }
                }
                const finalWidth = mode === 'expand' ? maxWidth : Math.min(maxWidth, 180);
                colGroupHtml += '<col style="width: ' + finalWidth + 'px; max-width: ' + (mode === 'expand' ? 'none' : '180px') + ';">';
            });

            colGroup.innerHTML = colGroupHtml;
            table.style.tableLayout = 'fixed';
            table.style.width = 'auto';
        } catch (e) {
            console.error('Error adjusting columns:', e);
        }
    }

    /**
     * Yields control back to the main thread to prevent UI freezing.
     * Uses scheduler.yield() if available, otherwise setTimeout.
     */
    function yieldToMain() {
        return new Promise(resolve => {
            // Double RAF + setTimeout ensures we truly yield to the browser
            requestAnimationFrame(() => {
                setTimeout(resolve, 0);
            });
        });
    }

    /**
     * Production-ready async chunked copy function.
     * Processes large selections without freezing the UI.
     */
    let isCopying = false;

    async function copySelectionToClipboard() {
        if (!selectedCells || selectedCells.size === 0) return;
        if (isCopying) return; // Prevent concurrent copies

        isCopying = true;
        const CHUNK_SIZE = 2000;

        try {
            // Immediate feedback
            showToast('Copying...');
            await yieldToMain();

            // Phase 1: Extract coordinates in chunks
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
                // Yield every CHUNK_SIZE cells
                if ((i + 1) % CHUNK_SIZE === 0) {
                    await yieldToMain();
                }
            }

            await yieldToMain();

            // Phase 2: Sort rows and columns
            const sortedRows = Array.from(rowSet).sort((a, b) => a - b);
            const sortedCols = Array.from(colSet).sort((a, b) => a - b);
            const numRows = sortedRows.length;
            const numCols = sortedCols.length;

            await yieldToMain();

            // Phase 3: Build TSV in chunks
            const tableData = getTableData();
            const outputLines = new Array(numRows);

            for (let i = 0; i < numRows; i++) {
                const r = sortedRows[i];
                const rowData = tableData[r];
                const lineParts = new Array(numCols);

                for (let j = 0; j < numCols; j++) {
                    const c = sortedCols[j];
                    lineParts[j] = rowData ? (rowData[c] ?? '') : '';
                }

                outputLines[i] = lineParts.join('\t');

                // Yield every CHUNK_SIZE rows
                if ((i + 1) % CHUNK_SIZE === 0) {
                    await yieldToMain();
                }
            }

            await yieldToMain();

            // Phase 4: Join all lines
            const tsv = outputLines.join('\n');

            await yieldToMain();

            // Phase 5: Write to clipboard with robust fallback
            await writeToClipboardAsync(tsv);

            // Flash animation on copied cells
            selectedCells.forEach(cell => {
                cell.classList.add('copying');
            });
            setTimeout(() => {
                selectedCells.forEach(cell => {
                    cell.classList.remove('copying');
                });
            }, 300);

            showToast('Copied ' + totalCells + ' cells');

        } catch (err) {
            console.error('Copy operation failed:', err);
            showToast('Copy failed');
        } finally {
            isCopying = false;
        }
    }

    /**
     * Async clipboard write with automatic fallback.
     */
    async function writeToClipboardAsync(text) {
        // Attempt 1: Modern Async Clipboard API
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            try {
                await navigator.clipboard.writeText(text);
                return; // Success
            } catch (e) {
                console.warn('Clipboard API failed, trying fallback:', e.message);
            }
        }

        // Attempt 2: execCommand fallback (wrapped in Promise)
        await execCommandFallback(text);
    }

    /**
     * execCommand('copy') fallback wrapped in a Promise.
     */
    function execCommandFallback(text) {
        return new Promise((resolve, reject) => {
            const textarea = document.createElement('textarea');
            textarea.value = text;

            // Critical: Make it invisible but still focusable
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

                const successful = document.execCommand('copy');

                document.body.removeChild(textarea);

                if (successful) {
                    resolve();
                } else {
                    reject(new Error('execCommand("copy") returned false'));
                }
            } catch (err) {
                document.body.removeChild(textarea);
                reject(err);
            }
        });
    }
    let exitAfterSave = false;

    function performSave(shouldExit = false) {
        if (isSaving || !isEditMode) return;
        isSaving = true;
        exitAfterSave = shouldExit;
        setButtonsEnabled(false);

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

    function initializeSelection() {
        const table = $('csv-table');
        if (!table || table.dataset.listenersAdded === 'true') return;
        table.dataset.listenersAdded = 'true';

        table.addEventListener('focusout', (e) => {
            if (isEditMode && e.target.tagName === 'TD') {
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
                        const nextCell = document.querySelector(`td[data-row="${coords.row + 1}"][data-col="${coords.col}"]`);
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
                const all = document.querySelectorAll('td[data-row][data-col]');
                all.forEach(c => { c.classList.add('selected'); selectedCells.add(c); });
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

                const next = document.querySelector('td[data-row="' + nr + '"][data-col="' + nc + '"]');
                if (next) {
                    e.preventDefault();
                    if (e.shiftKey) selectCellsInRange(startCell || coords, { row: nr, col: nc });
                    else {
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
            if (!e.target.closest('#csv-table') && !e.target.closest('.toolbar')) clearSelection();
        });
    }

    window.addEventListener('message', (event) => {
        const m = event.data;
        if (m.command === 'initTable' || m.command === 'appendRows') {
            const table = $('csv-table');
            if (!table) return;
            if (m.command === 'initTable') {
                table.querySelector('thead').innerHTML = m.headerHtml || '';
                table.querySelector('tbody').innerHTML = m.rowsHtml || '';
                adjustColumnWidths('default');
                setTimeout(() => adjustColumnWidths('default'), 100);
            } else {
                table.querySelector('tbody').insertAdjacentHTML('beforeend', m.rowsHtml || '');
            }
            applyEditModeToCells();
            initializeSelection();
            // Ensure toolbar scrollbar matches size after DOM changes
            setTimeout(syncToolbarScroll, 50);

            // Re-apply cached settings after rows are added so header/sticky behavior takes effect
            try { applySettings(currentSettings, false); } catch (e) { /* ignore */ }
        } else if (m.command === 'initSettings' || m.command === 'settingsUpdated') {
            // Apply incoming settings from the extension (persisted settings)
            const settings = (m && m.settings) ? m.settings : {};
            applySettings(settings, false);
            setTimeout(syncToolbarScroll, 20);
        } else if (m.command === 'saveResult') {
            isSaving = false;
            setButtonsEnabled(true);
            if (m.ok) {
                showToast('Saved');
                clearOriginalCellValues();
                if (exitAfterSave) {
                    setEditMode(false);
                } else {
                    captureOriginalCellValues();
                }
            } else {
                showToast('Error saving');
            }
        }
    });

    function wireButtons() {
        const btnMap = {
            toggleViewButton: () => vscode.postMessage({ command: 'toggleView', isTableView: (isTableView = !isTableView) }),
            toggleTableEditButton: () => setEditMode(!isEditMode),
            saveTableEditsButton: () => performSave(true),
            cancelTableEditsButton: () => {
                restoreOriginalCellValues();
                setEditMode(false);
            },
            toggleExpandButton: () => {
                const btn = $('toggleExpandButton');
                const state = btn.getAttribute('data-state') || 'default';
                if (state === 'default') {
                    btn.setAttribute('data-state', 'expanded');
                    document.body.classList.add('expanded-mode');
                    $('expandIcon').style.display = 'none';
                    $('collapseIcon').style.display = 'block';
                    $('expandButtonText').textContent = 'Default';
                    adjustColumnWidths('expand');
                } else {
                    btn.setAttribute('data-state', 'default');
                    document.body.classList.remove('expanded-mode');
                    $('expandIcon').style.display = 'block';
                    $('collapseIcon').style.display = 'none';
                    $('expandButtonText').textContent = 'Expand';
                    adjustColumnWidths('default');
                }
            }
        };

        // Theme manager component
        // Uses shared ThemeManager (resources/themeManager.js)
        const themeManager = new ThemeManager('toggleBackgroundButton', { 
            onBeforeCycle: () => !isEditMode
        }, vscode);

        Object.entries(btnMap).forEach(([id, handler]) => {
            const el = $(id);
            if (el) el.addEventListener('click', handler);
        });
    }

    // Setup floating tooltips so they are visible even when the toolbar is scrollable
    function setupFloatingTooltips() {
        let activeTip = null;
        let activeTrigger = null;

        function positionTip(trigger, tip) {
            if (!trigger || !tip) return;
            // ensure tip is fixed so it escapes any scrollable container
            tip.dataset.origPosition = tip.style.position || '';
            tip.dataset.origLeft = tip.style.left || '';
            tip.dataset.origTop = tip.style.top || '';
            tip.dataset.origTransform = tip.style.transform || '';

            tip.style.position = 'fixed';
            tip.style.visibility = 'visible';
            tip.style.opacity = '0';
            tip.style.pointerEvents = 'none';

            // measure and place on next frame
            requestAnimationFrame(() => {
                const r = trigger.getBoundingClientRect();
                const tr = tip.getBoundingClientRect();
                let left = r.left + r.width / 2 - tr.width / 2;
                left = Math.max(8, Math.min(left, window.innerWidth - tr.width - 8));
                let top = r.bottom + 8;
                // flip up if not enough space below
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

            // keep tooltip visible if mouse enters the floating tip
            function onTipEnter() {
                // noop, keep visible
            }
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
            // If mouse moved into the floating tip, don't hide
            const rel = e.relatedTarget;
            if (rel) {
                if (t.contains(rel)) return;
                if (activeTip && activeTip.contains(rel)) return;
            }
            hideTip(t);
        });

        // Reposition on resize/scroll while tooltip is visible
        window.addEventListener('resize', () => {
            if (activeTip && activeTrigger) positionTip(activeTrigger, activeTip);
        });
        window.addEventListener('scroll', (ev) => {
            if (activeTip && activeTrigger) positionTip(activeTrigger, activeTip);
        }, true);
    }

    setupFloatingTooltips();

    // Sync toolbar hidden scroll area with visible scrollbar
    function syncToolbarScroll() {
        const area = $('buttonScrollArea');
        const bar = $('buttonScrollbar');
        const inner = $('scrollInner');
        if (!area || !bar || !inner) return;
        // Set inner width to match the scrollable width of the button area
        inner.style.width = area.scrollWidth + 'px';
        // Keep the visible scrollbar at the same position
        bar.scrollLeft = area.scrollLeft;

        // Wire listeners once
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

    // --- Settings UI and behavior ---
    // Keep the currently-applied settings so they can be re-applied after rows are inserted
    let currentSettings = {
        firstRowIsHeader: false,
        stickyToolbar: true,
        stickyHeader: false
    };

    function applySettings(settings, saveLocal = false) {
        // cache latest settings
        currentSettings = settings || {};
        if (!settings) return;
        // Toggle classes for visual changes
        document.body.classList.toggle('first-row-as-header', !!settings.firstRowIsHeader);
        document.body.classList.toggle('sticky-header-enabled', !!settings.stickyHeader);
        document.body.classList.toggle('sticky-toolbar-enabled', !!settings.stickyToolbar);

        // Update settings panel UI
        const chkHeader = $('chkHeaderRow');
        const chkSticky = $('chkStickyHeader');
        const chkToolbar = $('chkStickyToolbar');
        if (chkHeader) chkHeader.checked = !!settings.firstRowIsHeader;
        if (chkSticky) chkSticky.checked = !!settings.stickyHeader;
        if (chkToolbar) chkToolbar.checked = !!settings.stickyToolbar;

        // Update header / sticky behavior
        // Bold first row when firstRowIsHeader is enabled
        const table = $('csv-table');
        if (table) {
            const firstRow = table.querySelector('tbody tr');
            if (firstRow) {
                if (settings.firstRowIsHeader) {
                    firstRow.classList.add('header-row');
                } else {
                    firstRow.classList.remove('header-row');
                }
            }
        }

        // Update toolbar stickiness and expansion
        const container = document.querySelector('.toolbar');
        const content = document.getElementById('content');
        const scrollArea = document.querySelector('.table-scroll');
        const headerBg = document.querySelector('.header-background');
        if (container) {
            if (settings.stickyToolbar) {
                // Ensure toolbar is a top-level element so it can be fixed to the viewport
                if (container.parentNode !== document.body) {
                    if (content && content.parentNode) document.body.insertBefore(container, content);
                    else document.body.appendChild(container);
                }
                container.classList.remove('not-sticky');
                container.classList.add('expanded-toolbar');
                if (headerBg) headerBg.style.display = '';
            } else {
                // Move toolbar into the scrollable table area so it scrolls with content
                const target = scrollArea || content;
                if (target && container.parentNode !== target) {
                    target.insertBefore(container, target.firstChild);
                }
                container.classList.add('not-sticky');
                container.classList.remove('expanded-toolbar');
                if (headerBg) headerBg.style.display = 'none';
            }
        }

        // update checkboxes if present
        const chkH = $('chkHeaderRow');
        const chkSH = $('chkStickyHeader');
        const chkST = $('chkStickyToolbar');
        if (chkH) chkH.checked = !!settings.firstRowIsHeader;
        if (chkSH) chkSH.checked = !!settings.stickyHeader;
        if (chkST) chkST.checked = !!settings.stickyToolbar;

        // Sticky header only meaningful when first row is header
        if (chkSH) chkSH.disabled = !chkH.checked;

        // If asked, optionally persist the settings via message
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

        if (!openBtn || !panel || !chkH || !chkSH || !chkST || !cancelBtn) return;

        // Keep a snapshot to allow cancel to revert
        let snapshot = null;

        let repositionHandlers = null;

        function repositionPanel() {
            const container = document.querySelector('.toolbar');
            if (!container) return;
            const rect = container.getBoundingClientRect();
            // Position the panel immediately below the toolbar, centered within viewport with some margin
            panel.style.position = 'fixed';
            panel.style.left = Math.max(8, rect.left) + 'px';
            panel.style.top = rect.bottom + 'px';
            // Keep the panel width no larger than container and leave small margins on the sides
            const maxWidth = Math.min(window.innerWidth - 16, rect.width);
            panel.style.width = Math.max(280, maxWidth) + 'px';
            panel.style.zIndex = '10001';
        }

        function openPanel() {
            snapshot = { firstRowIsHeader: chkH.checked, stickyHeader: chkSH.checked, stickyToolbar: chkST.checked };
            panel.classList.remove('hidden');
            panel.classList.add('floating');
            panel.setAttribute('aria-hidden', 'false');
            document.body.classList.add('settings-open');
            // Expand toolbar area visually but do not affect document flow
            const container = document.querySelector('.toolbar');
            if (container) {
                container.classList.add('settings-open');
                container.classList.add('expanded-toolbar');
            }

            // Position panel as an overlay beneath the toolbar and keep it in place while scrolling/resizing
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
                // If stickyToolbar setting is off, also remove expanded-toolbar
                const cfgSticky = chkST && chkST.checked;
                if (!cfgSticky) container.classList.remove('expanded-toolbar');
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

        // When a checkbox changes, apply immediately and persist
        function onChange() {
            const s = { firstRowIsHeader: !!chkH.checked, stickyHeader: !!chkSH.checked, stickyToolbar: !!chkST.checked };
            // When firstRowIsHeader is unchecked, ensure sticky header is also disabled
            if (!s.firstRowIsHeader) s.stickyHeader = false;
            applySettings(s, true);
        }

        chkH.addEventListener('change', () => {
            chkSH.disabled = !chkH.checked;
            if (!chkH.checked) chkSH.checked = false;
            onChange();
        });
        chkSH.addEventListener('change', onChange);
        chkST.addEventListener('change', onChange);

        cancelBtn.addEventListener('click', () => {
            // Close panel without reverting changes
            closePanel();
        });

        // Close panel if click outside
        document.addEventListener('click', (e) => {
            if (!panel.classList.contains('hidden')) {
                if (!e.target.closest('.settings-panel') && !e.target.closest('#openSettingsButton')) {
                    closePanel();
                }
            }
        });
    }

    // --- End settings UI ---

    setupFloatingTooltips();

    wireButtons();
    wireSettingsUI();
    vscode.postMessage({ command: 'webviewReady' });
})();
