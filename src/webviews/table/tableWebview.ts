/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-non-null-assertion */

import { ThemeManager } from '../shared/themeManager';
import { SettingsManager } from '../shared/settingsManager';
import { ToolbarManager } from '../shared/toolbarManager';
import { Utils } from '../shared/utils';
import { Icons } from '../shared/icons';
import { vscode, VirtualScrollConfig, debounce } from '../shared/common';
import { VirtualLoader } from '../shared/virtualLoader';
import { InfoTooltip } from '../shared/infoTooltip';

(function () {
    // ===== Configuration =====
    const { ROW_HEIGHT, BUFFER_ROWS, CHUNK_SIZE } = VirtualScrollConfig;

    // ===== State =====
    let isTableView = true;
    let isEditMode = false;
    let isSaving = false;
    let exitAfterSave = false;

    // Virtual scrolling state
    let totalRows = 0;
    let columnCount = 0;
    let rowCache = new Map<number, string[]>();
    const virtualLoader = new VirtualLoader<string[]>('getRows');
    let currentVisibleStart = 0;
    let currentVisibleEnd = 0;
    let isRequestingRows = false;

    // Selection state
    let isSelecting = false;
    let startCell: { row: number, col: number } | null = null;
    let endCell: { row: number, col: number } | null = null;
    const selectedCells = new Set<HTMLElement>();
    let activeCell: HTMLElement | null = null;
    const selectedRows = new Set<number>();
    const selectedColumns = new Set<number>();
    let lastSelectedRow: number | null = null;
    let lastSelectedColumn: number | null = null;

    // Track selected row/column indices for full copy
    const selectedRowIndices = new Set<number>();
    const selectedColumnIndices = new Set<number>();

    // Undo/Redo for edit mode
    let undoStack: string[][][] = [];
    let redoStack: string[][][] = [];
    const MAX_HISTORY = 50;

    // Settings
    interface Settings {
        firstRowIsHeader: boolean;
        stickyToolbar: boolean;
        stickyHeader: boolean;
        isDefaultEditor?: boolean;
    }

    let currentSettings: Settings = {
        firstRowIsHeader: false,
        stickyToolbar: true,
        stickyHeader: false,
        isDefaultEditor: true
    };

    // Current file format: 'csv' or 'tsv'
    let fileFormat = 'csv';

    // Toolbar manager (global for settings access)
    let toolbarManager: ToolbarManager | null = null;

    // ===== Utilities =====
    const $ = Utils.$;
    const normalizeCellText = Utils.normalizeCellText;
    const escapeHtml = Utils.escapeHtml;
    const showToast = Utils.showToast;
    const writeToClipboardAsync = Utils.writeToClipboardAsync;

    function escapeCsvCell(value: string): string {
        const v = value ?? '';
        // For CSV/TSV we need to escape quotes and newlines; also treat tab as special when serializing TSV
        const needsQuotes = /["\t,\n\r]/.test(v);
        if (!needsQuotes) return v;
        return '"' + v.replace(/"/g, '""') + '"';
    }

    function setButtonsEnabled(enabled: boolean) {
        const ids = ['toggleViewButton', 'toggleTableEditButton', 'saveTableEditsButton',
            'cancelTableEditsButton', 'toggleBackgroundButton', 'toggleExpandButton'];
        ids.forEach((id) => {
            const el = $(id) as HTMLButtonElement;
            if (el) el.disabled = !enabled;
        });
    }

    // ===== Virtual Scrolling Core =====

    function getTableContainer(): HTMLElement | null {
        return $('tableContainer');
    }

    function requestRows(start: number, end: number, timeout = 10000): Promise<string[][]> {
        return virtualLoader.requestRows(start, end, timeout);
    }

    function requestAllRows(): Promise<string[][]> {
        // Use longer timeout for full data fetch (30 seconds)
        return requestRows(0, totalRows, 30000);
    }

    function createRowHtml(rowData: string[], rowIndex: number): string {
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

    function renderVirtualRows(startIndex: number, endIndex: number, rowsData: string[][]) {
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

            const cachedRows: string[][] = [];
            for (let i = chunkStart; i < chunkEnd; i++) {
                cachedRows.push(rowCache.get(i) || []);
            }
            renderVirtualRows(chunkStart, chunkEnd, cachedRows);
        }
    }

    const onScroll = debounce(() => {
        updateVisibleRows();
    }, 16);

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
                if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
            });
        });

        // Re-apply row selection
        selectedRowIndices.forEach(rowIdx => {
            const rowHeader = document.querySelector(`th[data-row="${rowIdx}"]`);
            if (rowHeader && rowHeader.parentElement) {
                rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
                });
            }
        });

        // Re-apply active cell
        if (activeCell) {
            const row = activeCell.dataset?.row;
            const col = activeCell.dataset?.col;
            if (row !== undefined && col !== undefined) {
                const newCell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`) as HTMLElement;
                if (newCell) {
                    newCell.classList.add('active-cell');
                    activeCell = newCell;
                }
            }
        }
    }

    function getCellCoordinates(cell: HTMLElement | null): { row: number, col: number } | null {
        if (!cell || !cell.dataset) return null;
        return {
            row: parseInt(cell.dataset.row!, 10),
            col: parseInt(cell.dataset.col!, 10),
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
            const rows = new Set(cellsArray.map((cell) => parseInt(cell.dataset.row!, 10)));
            const cols = new Set(cellsArray.map((cell) => parseInt(cell.dataset.col!, 10)));
            selectionInfo.textContent = rows.size + 'R × ' + cols.size + 'C';
            selectionInfo.style.display = 'block';
        } else {
            selectionInfo.style.display = 'none';
        }
    }

    function selectCellsInRange(start: { row: number, col: number }, end: { row: number, col: number }) {
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
            const htmlCell = cell as HTMLElement;
            const coords = getCellCoordinates(htmlCell);
            if (!coords) return;
            if (coords.row >= minRow && coords.row <= maxRow &&
                coords.col >= minCol && coords.col <= maxCol) {
                htmlCell.classList.add('selected');
                selectedCells.add(htmlCell);
            }
        });

        const startCellElement = document.querySelector(
            `td[data-row="${start.row}"][data-col="${start.col}"]`
        ) as HTMLElement;
        if (startCellElement) {
            startCellElement.classList.add('active-cell');
            activeCell = startCellElement;
        }

        updateSelectionInfo();
    }

    function selectColumn(columnIndex: number, ctrlKey: boolean, shiftKey: boolean) {
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
                    if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
                });
            }
        } else if (ctrlKey) {
            if (selectedColumns.has(columnIndex)) {
                selectedColumns.delete(columnIndex);
                selectedColumnIndices.delete(columnIndex);
                document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                    cell.classList.remove('column-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell as HTMLElement);
                });
            } else {
                selectedColumns.add(columnIndex);
                selectedColumnIndices.add(columnIndex);
                document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
                });
            }
            lastSelectedColumn = columnIndex;
        } else {
            selectedColumns.add(columnIndex);
            selectedColumnIndices.add(columnIndex);
            document.querySelectorAll(`td[data-col="${columnIndex}"], th[data-col="${columnIndex}"]`).forEach((cell) => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
            });
            lastSelectedColumn = columnIndex;
        }
        updateSelectionInfo();
    }

    function selectRow(rowIndex: number, ctrlKey: boolean, shiftKey: boolean) {
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
                        if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
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
                        if (cell.tagName === 'TD') selectedCells.delete(cell as HTMLElement);
                    });
                }
            } else {
                selectedRows.add(rowIndex);
                selectedRowIndices.add(rowIndex);
                const rowHeader = document.querySelector(`th[data-row="${rowIndex}"]`);
                if (rowHeader && rowHeader.parentElement) {
                    rowHeader.parentElement.querySelectorAll('td, th').forEach((cell) => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
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
                    if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
                });
            }
            lastSelectedRow = rowIndex;
        }
        updateSelectionInfo();
    }

    // ===== Edit Mode =====

    function getTableData(): string[][] {
        const data: string[][] = [];
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
        if (current) redoStack.push(current);
        const previous = undoStack[undoStack.length - 1];

        previous.forEach((row, i) => {
            rowCache.set(i, row);
        });

        updateVisibleRows();
    }

    function redo() {
        if (redoStack.length === 0) return;
        const data = redoStack.pop();
        if (data) {
            undoStack.push(data);

            data.forEach((row, i) => {
                rowCache.set(i, row);
            });

            updateVisibleRows();
        }
    }

    function captureOriginalCellValues() {
        // Store original data from cache for cancel functionality
        (window as any)._originalCacheSnapshot = new Map();
        rowCache.forEach((value, key) => {
            (window as any)._originalCacheSnapshot.set(key, [...value]);
        });
    }

    function restoreOriginalCellValues() {
        if ((window as any)._originalCacheSnapshot) {
            rowCache = new Map((window as any)._originalCacheSnapshot);
            (window as any)._originalCacheSnapshot = null;
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

    function setEditMode(enabled: boolean) {
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

    function serializeTableToCsv(): string {
        const rows: string[] = [];
        const delimiter = fileFormat === 'tsv' ? '\t' : ',';

        for (let i = 0; i < totalRows; i++) {
            const rowData = rowCache.get(i) || [];
            const row: string[] = [];
            for (let j = 0; j < columnCount; j++) {
                const value = normalizeCellText(rowData[j] || '');
                row.push(escapeCsvCell(value));
            }
            rows.push(row.join(delimiter));
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
            const htmlCell = cell as HTMLElement;
            const row = parseInt(htmlCell.dataset.row!, 10);
            const col = parseInt(htmlCell.dataset.col!, 10);
            const value = normalizeCellText(htmlCell.textContent || '');

            let rowData = rowCache.get(row);
            if (!rowData) {
                rowData = [];
                rowCache.set(row, rowData);
            }
            rowData[col] = value;
        });

        if (document.activeElement && document.activeElement.tagName === 'TD') {
            (document.activeElement as HTMLElement).blur();
        }

        clearSelection();

        if (window.getSelection) {
            window.getSelection()!.removeAllRanges();
        }

        const csvText = serializeTableToCsv();
        vscode.postMessage({ command: 'saveCsv', text: csvText });
    }

    // ===== Copy =====

    let isCopying = false;
    let copyOperationTimeout: any = null;

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

            let outputLines: string[] = [];

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
                            const lineParts: string[] = [];
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
                const rowSet = new Set<number>();
                const colSet = new Set<number>();

                cellsArray.forEach(td => {
                    const r = parseInt(td.dataset.row!, 10);
                    const c = parseInt(td.dataset.col!, 10);
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



    // ===== UI Helpers =====



    function adjustColumnWidths(mode: 'expand' | 'default') {
        try {
            const table = $('csv-table');
            if (!table) return;
            const colGroup = table.querySelector('colgroup');
            if (!colGroup) return;

            const headerCells = table.querySelectorAll('th.col-header');
            if (headerCells.length === 0) return;

            table.style.tableLayout = 'auto';
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            if (!ctx) return;
            ctx.font = '13px "Segoe UI", Roboto, Helvetica, Arial, sans-serif';

            const visibleRows = table.querySelectorAll('tbody tr:not(.virtual-spacer)');
            const limit = Math.min(visibleRows.length, 50);

            // Measure first (row header) column so it can grow to fit content (min 30px)
            let firstColMax = 30;
            for (let r = 0; r < limit; r++) {
                const row = visibleRows[r];
                const cell = row && row.children && row.children[0];
                if (cell) {
                    const width = ctx.measureText(cell.textContent!.trim()).width + 24; // include padding
                    if (width > firstColMax) firstColMax = width;
                }
            }

            let colGroupHtml = `<col style="width: ${firstColMax}px;">`;
            headerCells.forEach((th, index) => {
                let maxWidth = ctx.measureText(th.textContent!.trim()).width + 32;
                for (let r = 0; r < limit; r++) {
                    const row = visibleRows[r];
                    const cell = row.children[index + 1];
                    if (cell) {
                        const width = ctx.measureText(cell.textContent!.trim()).width + 32;
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

    // ===== Toolbar Scroll Sync =====

    function syncToolbarScroll() {
        const area = $('buttonScrollArea');
        const bar = $('buttonScrollbar');
        const inner = $('scrollInner');
        if (!area || !bar || !inner) return;

        inner.style.width = area.scrollWidth + 'px';
        bar.scrollLeft = area.scrollLeft;

        if (!(area as any)._scrollWire) {
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
            (area as any)._scrollWire = true;
        }
    }

    // ===== Settings =====

    function applySettings(settings: Settings | null, saveLocal = false) {
        currentSettings = settings || {} as Settings;
        if (!settings) return;

        document.body.classList.toggle('first-row-as-header', !!settings.firstRowIsHeader);
        document.body.classList.toggle('sticky-header-enabled', !!settings.stickyHeader);
        document.body.classList.toggle('sticky-toolbar-enabled', !!settings.stickyToolbar);

        const chkHeader = $('chkHeaderRow') as HTMLInputElement;
        const chkSticky = $('chkStickyHeader') as HTMLInputElement;
        const chkToolbar = $('chkStickyToolbar') as HTMLInputElement;

        if (chkHeader) chkHeader.checked = !!settings.firstRowIsHeader;
        if (chkSticky) chkSticky.checked = !!settings.stickyHeader;
        if (chkToolbar) chkToolbar.checked = !!settings.stickyToolbar;

        // Show/hide enable button based on whether this is the default editor
        if (toolbarManager) {
            toolbarManager.setButtonVisibility('enableAsDefaultButton', settings.isDefaultEditor === false);
        }

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
        const headerBg = document.querySelector('.header-background') as HTMLElement;

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
        const settings = [
            {
                id: 'chkHeaderRow',
                label: 'Header Row',
                onChange: (val: boolean) => {
                    const chkSticky = document.getElementById('chkStickyHeader') as HTMLInputElement;
                    if (chkSticky) {
                        chkSticky.disabled = !val;
                        if (!val) {
                            chkSticky.checked = false;
                            currentSettings.stickyHeader = false;
                        }
                    }
                    currentSettings.firstRowIsHeader = val;
                    applySettings(currentSettings, true);
                },
                defaultValue: currentSettings.firstRowIsHeader
            },
            {
                id: 'chkStickyHeader',
                label: 'Sticky Header',
                onChange: (val: boolean) => {
                    currentSettings.stickyHeader = val;
                    applySettings(currentSettings, true);
                },
                defaultValue: currentSettings.stickyHeader
            },
            {
                id: 'chkStickyToolbar',
                label: 'Sticky Toolbar',
                onChange: (val: boolean) => {
                    currentSettings.stickyToolbar = val;
                    applySettings(currentSettings, true);
                },
                defaultValue: currentSettings.stickyToolbar
            }
        ];

        SettingsManager.renderPanel(document.getElementById('toolbar')!, 'settingsPanel', 'settingsCancelButton', settings);

        new SettingsManager('openSettingsButton', 'settingsPanel', 'settingsCancelButton', settings, () => {
            updateHeaderHeight();
        });
    }

    function updateHeaderHeight() {
        const container = document.querySelector('.toolbar');
        if (!container) return;
        // Respect non-sticky mode (stylesheet sets --header-height for non-sticky)
        if (!document.body.classList.contains('sticky-toolbar-enabled')) {
            document.documentElement.style.removeProperty('--header-height');
            const headerBg = document.querySelector('.header-background') as HTMLElement;
            if (headerBg) headerBg.style.height = '';
            return;
        }
        let h = Math.max(6, Math.ceil(container.getBoundingClientRect().height));
        // Cap header height to avoid large gaps when toolbar wraps or settings are open
        const maxH = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--header-height-max')) || 96;
        h = Math.min(h, maxH);
        document.documentElement.style.setProperty('--header-height', h + 'px');
        const headerBg = document.querySelector('.header-background') as HTMLElement;
        if (headerBg) headerBg.style.height = h + 'px';
    }

    // ===== Event Handlers =====

    function initializeSelection() {
        const table = $('csv-table');
        if (!table || table.dataset.listenersAdded === 'true') return;
        table.dataset.listenersAdded = 'true';

        table.addEventListener('focusout', (e) => {
            if (isEditMode && (e.target as HTMLElement).tagName === 'TD') {
                const target = e.target as HTMLElement;
                const row = parseInt(target.dataset.row!, 10);
                const col = parseInt(target.dataset.col!, 10);
                const value = normalizeCellText(target.textContent || '');

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
            const target = (e.target as HTMLElement).closest('td, th') as HTMLElement;
            if (!target) return;
            e.preventDefault();

            if (target.classList.contains('col-header')) {
                const colIdx = parseInt(target.dataset.col!, 10);
                if (!e.shiftKey) lastSelectedColumn = colIdx;
                selectColumn(colIdx, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }
            if (target.classList.contains('row-header')) {
                const rowIdx = parseInt(target.dataset.row!, 10);
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
            const target = (e.target as HTMLElement).closest('td') as HTMLElement;
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
                    const active = document.activeElement as HTMLElement;
                    const coords = getCellCoordinates(active);
                    if (coords) {
                        const nextCell = document.querySelector(
                            `td[data-row="${coords.row + 1}"][data-col="${coords.col}"]`
                        ) as HTMLElement;
                        if (nextCell) {
                            nextCell.focus();
                            const range = document.createRange();
                            const sel = window.getSelection();
                            range.selectNodeContents(nextCell);
                            range.collapse(false);
                            sel!.removeAllRanges();
                            sel!.addRange(range);
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
                    selectedCells.add(c as HTMLElement);
                });
                if (all[0]) {
                    all[0].classList.add('active-cell');
                    activeCell = all[0] as HTMLElement;
                    startCell = getCellCoordinates(all[0] as HTMLElement);
                }
                updateSelectionInfo();
            } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && activeCell) {
                const coords = getCellCoordinates(activeCell);
                if (!coords) return;
                let nr = coords.row, nc = coords.col;
                if (e.key === 'ArrowUp' && nr > 0) nr--;
                else if (e.key === 'ArrowDown') nr++;
                else if (e.key === 'ArrowLeft' && nc > 0) nc--;
                else if (e.key === 'ArrowRight') nc++;

                const next = document.querySelector(`td[data-row="${nr}"][data-col="${nc}"]`) as HTMLElement;
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
            if (!(e.target as HTMLElement).closest('#csv-table') && !(e.target as HTMLElement).closest('.toolbar')) {
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

                // Set current format (csv or tsv)
                fileFormat = m.format || 'csv';

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
                virtualLoader.resolveRequest(m.requestId, m.rows || []);
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
                    (window as any)._originalCacheSnapshot = null;
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
        toolbarManager = new ToolbarManager('toolbar');
        
        toolbarManager.setButtons([
            {
                id: 'toggleViewButton',
                icon: Icons.EditFile,
                label: 'Edit File',
                tooltip: 'Edit File in Vscode Default Editor',
                onClick: () => {
                    isTableView = !isTableView;
                    vscode.postMessage({ command: 'toggleView', isTableView });
                }
            },
            {
                id: 'toggleTableEditButton',
                icon: '',
                label: 'Edit Table',
                tooltip: 'Edit CSV directly in the table',
                onClick: () => setEditMode(!isEditMode)
            },
            {
                id: 'saveTableEditsButton',
                icon: '',
                label: 'Save',
                tooltip: 'Save table edits',
                hidden: true,
                onClick: () => performSave(true)
            },
            {
                id: 'cancelTableEditsButton',
                icon: '',
                label: 'Cancel',
                tooltip: 'Cancel table edits',
                hidden: true,
                onClick: () => {
                    restoreOriginalCellValues();
                    setEditMode(false);
                }
            },
            {
                id: 'toggleExpandButton',
                icon: Icons.Expand,
                label: 'Expand',
                tooltip: 'Toggle Column Widths (Default / Expand All)',
                cls: 'edit-mode-hide',
                onClick: () => {
                    const btn = $('toggleExpandButton');
                    const state = btn?.getAttribute('data-state') || 'default';
                    if (state === 'default') {
                        btn?.setAttribute('data-state', 'expanded');
                        document.body.classList.add('expanded-mode');
                        if(btn) btn.innerHTML = Icons.Collapse + ' <span class="btn-label">Default</span>';
                        adjustColumnWidths('expand');
                    } else {
                        btn?.setAttribute('data-state', 'default');
                        document.body.classList.remove('expanded-mode');
                        if(btn) btn.innerHTML = Icons.Expand + ' <span class="btn-label">Expand</span>';
                        adjustColumnWidths('default');
                    }
                }
            },
            {
                id: 'openSettingsButton',
                icon: Icons.Settings,
                tooltip: 'CSV Settings',
                cls: 'icon-only',
                onClick: () => {}
            },
            {
                id: 'toggleBackgroundButton',
                icon: Icons.ThemeLight + Icons.ThemeDark + Icons.ThemeVSCode,
                tooltip: 'Toggle Theme',
                cls: 'edit-mode-hide',
                onClick: () => {}
            },
            {
                id: 'helpButton',
                icon: Icons.Help,
                tooltip: 'Help & Feedback',
                cls: 'icon-only',
                onClick: () => {
                    vscode.postMessage({
                        command: 'openExternal',
                        url: 'https://docs.google.com/forms/d/e/1FAIpQLSe5AqE_f1-WqUlQmvuPn1as3Mkn4oLjA0EDhNssetzt63ONzA/viewform'
                    });
                }
            },
            {
                id: 'enableAsDefaultButton',
                icon: Icons.Zap,
                label: 'Set as Default',
                tooltip: `Make XLSX Viewer the default editor for ${fileFormat.toUpperCase()} files`,
                cls: 'edit-mode-hide',
                hidden: true,
                onClick: () => {
                    vscode.postMessage({ command: 'enableAsDefault' });
                }
            }
        ]);

        // Inject tooltip if variables are present
        InfoTooltip.inject('toolbar', (window as any).viewImgUri, (window as any).logoSvgUri, 'table view');

        // Theme manager
        if (typeof ThemeManager !== 'undefined') {
            new ThemeManager('toggleBackgroundButton', {
                onBeforeCycle: () => !isEditMode
            }, vscode);
        }
    }

    // ===== Initialize =====

    wireButtons();
    wireSettingsUI();
    updateHeaderHeight();
    window.addEventListener('resize', updateHeaderHeight);
    vscode.postMessage({ command: 'webviewReady' });
})();
