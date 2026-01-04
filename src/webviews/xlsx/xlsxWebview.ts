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
    // ===== Virtual Scrolling Configuration =====
    const { ROW_HEIGHT, BUFFER_ROWS, CHUNK_SIZE } = VirtualScrollConfig;

    // Data injected from the extension via postMessage
    let worksheetsMeta: any[] = [];
    let currentWorksheet = 0;

    // Virtual scrolling state
    let totalRows = 0;
    let columnCount = 0;
    let columnWidths: number[] = [];
    let mergedCells: any[] = [];
    let allRowHeights: number[] = []; // Pre-loaded row heights from extension
    let totalContentHeight = 0; // Pre-calculated total height
    let rowCache = new Map<number, any>();
    const virtualLoader = new VirtualLoader<any[]>('getRows');
    let currentVisibleStart = 0;
    let currentVisibleEnd = 0;
    let isRequestingRows = false;
    let isRendering = false; // Prevent re-render during render

    // Selection state
    const selectedCells = new Set<HTMLElement>();
    let activeCell: HTMLElement | null = null;
    let isSelecting = false;
    let selectionStart: { row: number, col: number } | null = null;
    let selectionEnd: { row: number, col: number } | null = null;
    const selectedRows = new Set<number>();
    const selectedColumns = new Set<number>();
    let lastSelectedRow: number | null = null;
    let lastSelectedColumn: number | null = null;

    // Track selected row/column indices for full copy (virtualization support)
    const selectedRowIndices = new Set<number>();
    const selectedColumnIndices = new Set<number>();

    // Resize state
    let isResizing = false;
    let resizeType: 'column' | 'row' | null = null; // 'column' or 'row'
    let resizeIndex = -1;
    let resizeStartPos = 0;
    let resizeStartSize = 0;

    // Auto-scroll while dragging selection
    let autoScrollRequest: any = null;
    let lastMousePos: { x: number, y: number } | null = null; // { x, y }
    const AUTO_SCROLL_THRESHOLD = 40; // px
    const AUTO_SCROLL_STEP = 20; // px per frame

    let handlersAttached = false;

    // Settings (persisted by extension)
    interface Settings {
        firstRowIsHeader: boolean;
        stickyToolbar: boolean;
        stickyHeader: boolean;
        hyperlinkPreview: boolean;
    }

    let currentSettings: Settings = {
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

    // Plain view mode (removes all XLSX styling)
    let isPlainView = false;

    // Hyperlink hover tooltip
    let linkTooltip: HTMLElement | null = null;
    let linkTooltipHideTimer: any = null;

    // Toast
    let toastEl: HTMLElement | null = null;

    // Copy state (CSV-parity: avoid concurrent copies)
    let isCopying = false;

    function setButtonsEnabled(enabled: boolean) {
        const saveBtn = document.getElementById('saveTableEditsButton') as HTMLButtonElement;
        const cancelBtn = document.getElementById('cancelTableEditsButton') as HTMLButtonElement;
        if (saveBtn) saveBtn.disabled = !enabled;
        if (cancelBtn) cancelBtn.disabled = !enabled;
    }

    function normalizeCellText(text: string | null | undefined): string {
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

    async function writeToClipboardAsync(text: string) {
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            try {
                await navigator.clipboard.writeText(text);
                return;
            } catch {
                // fall through to execCommand
            }
        }

        await new Promise<void>((resolve, reject) => {
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

    function getExcelColumnLabel(n: number): string {
        let label = '';
        while (n > 0) {
            const rem = (n - 1) % 26;
            label = String.fromCharCode(65 + rem) + label;
            n = Math.floor((n - 1) / 26);
        }
        return label;
    }

    function formatCellStyle(style: any): string {
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

    // ===== Virtual Scrolling Core =====

    function getTableContainer(): HTMLElement | null {
        return document.querySelector('.table-scroll');
    }

    function requestRows(start: number, end: number, timeout = 10000): Promise<any[]> {
        return virtualLoader.requestRows(start, end, timeout, { sheetIndex: currentWorksheet });
    }

    function requestAllRows(): Promise<any[]> {
        return requestRows(0, totalRows, 30000);
    }

    function createRowHtml(rowData: any, rowIndex: number): string {
        const height = rowData.height || ROW_HEIGHT;
        const isHeaderRow = rowIndex === 0;

        let html = '<tr data-virtual-row="' + rowIndex + '" style="height: ' + height + 'px;"' + (isHeaderRow ? ' class="header-row"' : '') + '>';
        html += '<th class="row-header" data-row="' + rowIndex + '" style="height: ' + height + 'px;">';
        html += rowData.rowNumber || (rowIndex + 1);
        html += '<div class="row-resize-handle" data-row="' + rowIndex + '"></div>';
        html += '</th>';

        let virtualColIndex = 0;
        for (let actualCol = 1; actualCol <= columnCount; actualCol++) {
            const cellData = rowData.cells ? rowData.cells.find((cell: any) => cell.colNumber === actualCol) : null;

            if (cellData) {
                // In plain view mode, skip all styling
                const styleStr = isPlainView ? '' : formatCellStyle(cellData.style || {});
                const cellHeight = height * (cellData.rowspan || 1);
                const cellWidth = columnWidths
                    .slice(actualCol - 1, actualCol - 1 + (cellData.colspan || 1))
                    .reduce((sum, w) => sum + (w || 80), 0);

                html += '<td';
                html += ' data-row="' + rowIndex + '"';
                html += ' data-col="' + virtualColIndex + '"';
                html += ' data-rownum="' + cellData.rowNumber + '"';
                html += ' data-colnum="' + cellData.colNumber + '"';
                
                // Only add styling data attributes if not in plain view
                if (!isPlainView) {
                    if (cellData.hasDefaultBg) html += ' data-default-bg="true"';
                    if (cellData.hasWhiteBackground) html += ' data-white-bg="true"';
                    if (cellData.isDefaultColor) html += ' data-default-color="true"';
                    if (cellData.hasBlackBorder) html += ' data-black-border="true"';
                    if (cellData.hasWhiteBorder) html += ' data-white-border="true"';
                    if (cellData.hasBlackBackground) html += ' data-black-bg="true"';
                    if (cellData.hasDefaultBorder) html += ' data-default-border="true"';
                }
                if (cellData.isEmpty) html += ' data-empty="true"';
                if (cellData.hyperlink) html += ' data-hyperlink="' + String(cellData.hyperlink).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;') + '"';
                html += ' data-original-color="' + (cellData.originalColor || 'rgb(0, 0, 0)') + '"';

                // Skip rowspan/colspan in plain view for simpler display
                if (!isPlainView) {
                    if (cellData.rowspan > 1) html += ' rowspan="' + cellData.rowspan + '"';
                    if (cellData.colspan > 1) html += ' colspan="' + cellData.colspan + '"';
                    if (cellData.isMerged) html += ' class="merged-cell"';
                }

                let cellStyleStr = styleStr;
                if (!isPlainView && cellData.isMerged) {
                    cellStyleStr += 'height: ' + cellHeight + 'px; width: ' + cellWidth + 'px;';
                } else {
                    cellStyleStr += 'height: ' + height + 'px;';
                }

                if (isEditMode) {
                    html += ' contenteditable="true" spellcheck="false"';
                }

                if (cellStyleStr) {
                    html += ' style="' + cellStyleStr + '"';
                }
                html += '>';
                html += '<span class="cell-content">' + (cellData.value || '&nbsp;') + '</span>';
                html += '</td>';
            } else {
                // Empty cell - include all data attributes for proper theme styling
                html += '<td data-row="' + rowIndex + '" data-col="' + virtualColIndex + '"';
                html += ' data-rownum="' + (rowIndex + 1) + '"';
                html += ' data-colnum="' + actualCol + '"';
                if (!isPlainView) {
                    html += ' data-default-bg="true" data-default-color="true" data-default-border="true"';
                }
                html += ' data-empty="true"';
                html += ' data-original-color="rgb(0, 0, 0)"';
                if (isEditMode) {
                    html += ' contenteditable="true" spellcheck="false"';
                }
                html += ' style="height: ' + height + 'px;">';
                html += '<span class="cell-content">&nbsp;</span>';
                html += '</td>';
            }
            virtualColIndex++;
        }

        html += '</tr>';
        return html;
    }

    function renderVirtualRows(startIndex: number, endIndex: number, rowsData: any[]) {
        if (isRendering) return;
        isRendering = true;

        const tbody = document.querySelector('#xlsxTable tbody');
        if (!tbody) {
            isRendering = false;
            return;
        }

        // Cache rows
        rowsData.forEach((row, i) => {
            rowCache.set(startIndex + i, row);
        });

        // Calculate spacer heights using pre-loaded heights (stable)
        let topSpacerHeight = 0;
        for (let i = 0; i < startIndex; i++) {
            topSpacerHeight += allRowHeights[i] || ROW_HEIGHT;
        }

        let bottomSpacerHeight = 0;
        for (let i = endIndex; i < totalRows; i++) {
            bottomSpacerHeight += allRowHeights[i] || ROW_HEIGHT;
        }

        let html = '';

        if (topSpacerHeight > 0) {
            html += '<tr class="virtual-spacer top-spacer"><td colspan="' + (columnCount + 1) + '" style="height: ' + topSpacerHeight + 'px; padding: 0; border: none;"></td></tr>';
        }

        for (let i = startIndex; i < endIndex; i++) {
            const rowData = rowCache.get(i) || { cells: [], rowNumber: i + 1 };
            html += createRowHtml(rowData, i);
        }

        if (bottomSpacerHeight > 0) {
            html += '<tr class="virtual-spacer bottom-spacer"><td colspan="' + (columnCount + 1) + '" style="height: ' + bottomSpacerHeight + 'px; padding: 0; border: none;"></td></tr>';
        }

        tbody.innerHTML = html;
        reapplySelection();
        isRendering = false;
    }

    async function updateVisibleRows() {
        if (isRendering) return;
        
        const container = getTableContainer();
        if (!container || totalRows === 0) return;

        const scrollTop = container.scrollTop;
        const clientHeight = container.clientHeight;

        // Calculate which rows are visible using pre-loaded heights
        let accumulatedHeight = 0;
        let firstVisibleRow = 0;
        
        for (let i = 0; i < totalRows; i++) {
            const rowHeight = allRowHeights[i] || ROW_HEIGHT;
            if (accumulatedHeight + rowHeight > scrollTop) {
                firstVisibleRow = i;
                break;
            }
            accumulatedHeight += rowHeight;
            if (i === totalRows - 1) {
                firstVisibleRow = totalRows - 1;
            }
        }
        
        // Find last visible row
        let lastVisibleRow = firstVisibleRow;
        let visibleHeight = 0;
        for (let i = firstVisibleRow; i < totalRows; i++) {
            const rowHeight = allRowHeights[i] || ROW_HEIGHT;
            visibleHeight += rowHeight;
            lastVisibleRow = i + 1;
            if (visibleHeight >= clientHeight) {
                break;
            }
        }

        // Add buffer
        const bufferedStart = Math.max(0, firstVisibleRow - BUFFER_ROWS);
        const bufferedEnd = Math.min(totalRows, lastVisibleRow + BUFFER_ROWS);

        // Align to chunk boundaries
        let chunkStart = Math.floor(bufferedStart / CHUNK_SIZE) * CHUNK_SIZE;
        let chunkEnd = Math.ceil(bufferedEnd / CHUNK_SIZE) * CHUNK_SIZE;
        
        // Clamp to totalRows
        chunkEnd = Math.min(totalRows, chunkEnd);

        // CRITICAL FIX: If we're within 2 chunks of the end, just render to the end
        // This prevents fluctuation at boundaries like 2224 rows (22.24 chunks)
        const remainingRows = totalRows - chunkEnd;
        if (remainingRows > 0 && remainingRows < CHUNK_SIZE * 2) {
            chunkEnd = totalRows;
        }

        // Skip if we're already showing these rows (with some tolerance)
        if (chunkStart === currentVisibleStart && chunkEnd === currentVisibleEnd) {
            return;
        }

        // Check if current range still covers what we need
        if (currentVisibleStart <= bufferedStart && currentVisibleEnd >= bufferedEnd) {
            return; // Current render still covers visible area
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

            const cachedRows: any[] = [];
            for (let i = chunkStart; i < chunkEnd; i++) {
                cachedRows.push(rowCache.get(i) || { cells: [], rowNumber: i + 1 });
            }
            renderVirtualRows(chunkStart, chunkEnd, cachedRows);
        }
    }

    const onScroll = debounce(() => {
        updateVisibleRows();
    }, 50); // Increased debounce to reduce fluctuation

    function initializeVirtualScrolling() {
        const container = getTableContainer();
        if (!container) return;

        // Remove any existing listener first
        container.removeEventListener('scroll', onScroll);
        container.addEventListener('scroll', onScroll, { passive: true });
        updateVisibleRows();
    }

    function reapplySelection() {
        // Re-apply column selection
        selectedColumnIndices.forEach(colIdx => {
            document.querySelectorAll('td[data-col="' + colIdx + '"], th[data-col="' + colIdx + '"]').forEach((cell) => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
            });
        });

        // Re-apply row selection
        selectedRowIndices.forEach(rowIdx => {
            document.querySelectorAll('td[data-row="' + rowIdx + '"], th[data-row="' + rowIdx + '"]').forEach((cell) => {
                cell.classList.add('row-selected');
                if (cell.tagName === 'TD') selectedCells.add(cell as HTMLElement);
            });
        });

        // Re-apply active cell
        if (activeCell) {
            const row = activeCell.dataset?.row;
            const col = activeCell.dataset?.col;
            if (row !== undefined && col !== undefined) {
                const newCell = document.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]') as HTMLElement;
                if (newCell) {
                    newCell.classList.add('active-cell');
                    activeCell = newCell;
                }
            }
        }
    }

    function createTableShell(): string {
        let html = '<div class="table-scroll"><table id="xlsxTable">';

        // Header row
        html += '<thead><tr>';
        html += '<th class="corner-cell"></th>';
        for (let c = 1; c <= columnCount; c++) {
            const width = columnWidths[c - 1] || 80;
            html += '<th class="col-header" data-col="' + (c - 1) + '" style="width: ' + width + 'px; min-width: ' + width + 'px;">';
            html += getExcelColumnLabel(c);
            html += '<div class="col-resize-handle" data-col="' + (c - 1) + '"></div>';
            html += '</th>';
        }
        html += '</tr></thead><tbody></tbody></table></div>';
        return html;
    }

    const showToast = Utils.showToast;

    function setLoadingText(text: string) {
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

    function renderWorksheet(index: number) {
        if (!worksheetsMeta || !worksheetsMeta.length) return;

        showLoading();

        // Reset virtual scrolling state for new worksheet
        rowCache.clear();
        currentVisibleStart = 0;
        currentVisibleEnd = 0;
        virtualLoader.clear();
        isRendering = false;

        const wsMeta = worksheetsMeta[index];
        totalRows = wsMeta.totalRows || 0;
        columnCount = wsMeta.columnCount || 0;
        columnWidths = wsMeta.columnWidths || [];
        mergedCells = wsMeta.mergedCells || [];
        allRowHeights = wsMeta.rowHeights || [];
        
        // Pre-calculate total content height for stable scrolling
        totalContentHeight = 0;
        for (let i = 0; i < totalRows; i++) {
            totalContentHeight += allRowHeights[i] || ROW_HEIGHT;
        }

        // Allow the overlay to render
        setTimeout(() => {
            const container = document.getElementById('tableContainer');
            if (!container) return;

            container.innerHTML = createTableShell();
            initializeSelection();
            initializeResize();
            initializeHyperlinkHover();
            initializeVirtualScrolling();
            hideLoading();
        }, 100);
    }

    function initializeResize() {
        const table = document.querySelector('table');
        if (!table) return;

        // Column/row resize handles
        table.addEventListener('mousedown', (e) => {
            if (isEditMode) return;
            const target = e.target as HTMLElement;
            if (target && target.classList && target.classList.contains('col-resize-handle')) {
                e.preventDefault();
                e.stopPropagation();

                isResizing = true;
                resizeType = 'column';
                resizeIndex = parseInt(target.dataset.col!, 10);
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
                resizeIndex = parseInt(target.dataset.row!, 10);
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

                const headers = tableEl.querySelectorAll('th.col-header[data-col="' + resizeIndex + '"]') as NodeListOf<HTMLElement>;
                const cells = tableEl.querySelectorAll('td[data-col="' + resizeIndex + '"]') as NodeListOf<HTMLElement>;

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

                const headers = tableEl.querySelectorAll('th.row-header[data-row="' + resizeIndex + '"]') as NodeListOf<HTMLElement>;
                const row = tableEl.querySelectorAll('tr')[resizeIndex + 1] as HTMLElement; // +1 for header row

                headers.forEach(header => {
                    header.style.height = newSize + 'px';
                });

                if (row) {
                    row.style.height = newSize + 'px';
                    const cells = row.querySelectorAll('td') as NodeListOf<HTMLElement>;
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
            const target = e.target as HTMLElement;
            if (target && target.classList && target.classList.contains('col-resize-handle')) {
                e.preventDefault();
                autoFitColumn(parseInt(target.dataset.col!, 10));
            } else if (target && target.classList && target.classList.contains('row-resize-handle')) {
                e.preventDefault();
                autoFitRow(parseInt(target.dataset.row!, 10));
            }
        });
    }

    function autoFitColumn(colIndex: number) {
        const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]') as NodeListOf<HTMLElement>;
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

    function autoFitRow(rowIndex: number) {
        const row = document.querySelectorAll('tr')[rowIndex + 1] as HTMLElement; // +1 for header row
        if (!row) return;

        const cells = row.querySelectorAll('td') as NodeListOf<HTMLElement>;
        let maxHeight = 20;

        cells.forEach(cell => {
            const content = (cell.textContent || '').trim();
            if (content.length > 50) {
                maxHeight = Math.max(maxHeight, 40);
            }
        });

        row.style.height = maxHeight + 'px';
        const headers = document.querySelectorAll('th.row-header[data-row="' + rowIndex + '"]') as NodeListOf<HTMLElement>;
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
        if (!worksheetsMeta || !worksheetsMeta.length) return;
        // Note: worksheetsData was not defined in original JS, assuming it meant worksheetsMeta or similar
        // But autoFitAllColumns was not called anywhere in the original JS.
        // Keeping it but commenting out usage if any.
        /*
        const data = worksheetsMeta[currentWorksheet].data;
        for (let c = 0; c < data.maxCol; c++) {
            autoFitColumn(c);
        }
        */
    }

    function clearSelection() {
        document.querySelectorAll('.selected, .active-cell, .row-selected, .column-selected').forEach(el => {
            el.classList.remove('selected', 'active-cell', 'row-selected', 'column-selected');
        });
        selectedCells.clear();
        selectedRows.clear();
        selectedColumns.clear();
        selectedRowIndices.clear();
        selectedColumnIndices.clear();
        activeCell = null;
        lastSelectedRow = null;
        lastSelectedColumn = null;
        const info = document.getElementById('selectionInfo');
        if (info) info.style.display = 'none';
    }

    function selectCell(cell: HTMLElement, isMulti = false) {
        if (!isMulti) {
            clearSelection();
        }

        cell.classList.add('selected');
        cell.classList.add('active-cell');
        selectedCells.add(cell);
        activeCell = cell;
        updateSelectionInfo();
    }

    function selectRange(startRow: number, startCol: number, endRow: number, endCol: number) {
        clearSelection();

        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);

        const cells = document.querySelectorAll('td') as NodeListOf<HTMLElement>;
        cells.forEach(cell => {
            const row = parseInt(cell.dataset.row!, 10);
            const col = parseInt(cell.dataset.col!, 10);

            if (row >= minRow && row <= maxRow && col >= minCol && col <= maxCol) {
                cell.classList.add('selected');
                selectedCells.add(cell);
            }
        });

        const startCell = document.querySelector('td[data-row="' + startRow + '"][data-col="' + startCol + '"]') as HTMLElement;
        if (startCell) {
            startCell.classList.add('active-cell');
            activeCell = startCell;
        }

        updateSelectionInfo();
    }

    function selectRow(rowIndex: number, ctrlKey: boolean, shiftKey: boolean) {
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
                    selectedRowIndices.add(row);
                    const cells = document.querySelectorAll('td[data-row="' + row + '"], th[data-row="' + row + '"]');
                    cells.forEach(cell => {
                        cell.classList.add('row-selected');
                        if (cell.tagName === 'TD') {
                            selectedCells.add(cell as HTMLElement);
                        }
                    });
                }
            }
        } else if (ctrlKey) {
            if (selectedRows.has(rowIndex)) {
                selectedRows.delete(rowIndex);
                selectedRowIndices.delete(rowIndex);
                const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.remove('row-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell as HTMLElement);
                });
            } else {
                selectedRows.add(rowIndex);
                selectedRowIndices.add(rowIndex);
                const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.add('row-selected');
                    if (cell.tagName === 'TD') {
                        selectedCells.add(cell as HTMLElement);
                    }
                });
            }
        } else {
            selectedRows.add(rowIndex);
            selectedRowIndices.add(rowIndex);
            const cells = document.querySelectorAll('td[data-row="' + rowIndex + '"], th[data-row="' + rowIndex + '"]');
            cells.forEach(cell => {
                cell.classList.add('row-selected');
                if (cell.tagName === 'TD') {
                    selectedCells.add(cell as HTMLElement);
                }
            });
        }

        updateSelectionInfo();
    }

    function selectColumn(colIndex: number, ctrlKey: boolean, shiftKey: boolean) {
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
                    selectedColumnIndices.add(col);
                    const cells = document.querySelectorAll('td[data-col="' + col + '"], th[data-col="' + col + '"]');
                    cells.forEach(cell => {
                        cell.classList.add('column-selected');
                        if (cell.tagName === 'TD') {
                            selectedCells.add(cell as HTMLElement);
                        }
                    });
                }
            }
        } else if (ctrlKey) {
            if (selectedColumns.has(colIndex)) {
                selectedColumns.delete(colIndex);
                selectedColumnIndices.delete(colIndex);
                const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.remove('column-selected');
                    if (cell.tagName === 'TD') selectedCells.delete(cell as HTMLElement);
                });
            } else {
                selectedColumns.add(colIndex);
                selectedColumnIndices.add(colIndex);
                const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
                cells.forEach(cell => {
                    cell.classList.add('column-selected');
                    if (cell.tagName === 'TD') {
                        selectedCells.add(cell as HTMLElement);
                    }
                });
            }
        } else {
            selectedColumns.add(colIndex);
            selectedColumnIndices.add(colIndex);
            const cells = document.querySelectorAll('td[data-col="' + colIndex + '"], th[data-col="' + colIndex + '"]');
            cells.forEach(cell => {
                cell.classList.add('column-selected');
                if (cell.tagName === 'TD') {
                    selectedCells.add(cell as HTMLElement);
                }
            });
        }

        updateSelectionInfo();
    }

    function updateSelectionInfo() {
        const info = document.getElementById('selectionInfo');
        if (!info) return;

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

            info.textContent = rowCount + 'R × ' + colCount + 'C';
            info.style.display = 'block';
        } else if (selectedCells.size > 1) {
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
        const hasFullColumnSelection = selectedColumnIndices.size > 0;
        const hasFullRowSelection = selectedRowIndices.size > 0;

        if (!hasFullColumnSelection && !hasFullRowSelection && selectedCells.size === 0) return;
        if (isCopying) return;

        isCopying = true;

        try {
            showToast('Copying...');
            await yieldToMain();

            let outputLines: string[] = [];

            if (hasFullColumnSelection || hasFullRowSelection) {
                // Need to fetch all rows for complete copy
                const allRows = await requestAllRows();

                if (!allRows || allRows.length === 0) {
                    showToast('Failed to fetch data');
                    isCopying = false;
                    return;
                }

                // Cache the fetched rows
                if (allRows.length >= totalRows * 0.9) {
                    allRows.forEach((row, i) => {
                        rowCache.set(i, row);
                    });
                }

                const rowCount = allRows.length;

                if (hasFullColumnSelection && !hasFullRowSelection) {
                    // Copy entire columns
                    const sortedCols = Array.from(selectedColumnIndices).sort((a, b) => a - b);

                    for (let r = 0; r < rowCount; r++) {
                        const rowData = allRows[r] || { cells: [] };
                        const lineParts = sortedCols.map(c => {
                            const cellData = rowData.cells ? rowData.cells.find((cell: any) => cell.colNumber === c + 1) : null;
                            return cellData ? normalizeCellText(cellData.value || '') : '';
                        });
                        outputLines.push(lineParts.join('\t'));
                    }
                } else if (hasFullRowSelection && !hasFullColumnSelection) {
                    // Copy entire rows
                    const sortedRows = Array.from(selectedRowIndices).sort((a, b) => a - b);

                    for (const r of sortedRows) {
                        if (r < rowCount) {
                            const rowData = allRows[r] || { cells: [] };
                            const lineParts: string[] = [];
                            for (let c = 0; c < columnCount; c++) {
                                const cellData = rowData.cells ? rowData.cells.find((cell: any) => cell.colNumber === c + 1) : null;
                                lineParts.push(cellData ? normalizeCellText(cellData.value || '') : '');
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
                            const rowData = allRows[r] || { cells: [] };
                            const lineParts = sortedCols.map(c => {
                                const cellData = rowData.cells ? rowData.cells.find((cell: any) => cell.colNumber === c + 1) : null;
                                return cellData ? normalizeCellText(cellData.value || '') : '';
                            });
                            outputLines.push(lineParts.join('\t'));
                        }
                    }
                }

                const cellCount = hasFullColumnSelection ?
                    rowCount * selectedColumnIndices.size :
                    (hasFullRowSelection ? selectedRowIndices.size * columnCount : 0);

                const tsv = outputLines.join('\n');
                await writeToClipboardAsync(tsv);

                selectedCells.forEach(cell => cell.classList.add('copying'));
                setTimeout(() => selectedCells.forEach(cell => cell.classList.remove('copying')), 300);

                showToast('Copied ' + cellCount + ' cells');
            } else {
                // Regular cell selection - use DOM/cache
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
                    const lineParts = sortedCols.map(c => {
                        const cell = document.querySelector('td[data-row="' + r + '"][data-col="' + c + '"]');
                        return normalizeCellText(cell ? (cell.textContent || '') : '');
                    });
                    outputLines.push(lineParts.join('\t'));
                }

                const tsv = outputLines.join('\n');
                await writeToClipboardAsync(tsv);

                selectedCells.forEach(cell => cell.classList.add('copying'));
                setTimeout(() => selectedCells.forEach(cell => cell.classList.remove('copying')), 300);

                showToast('Copied ' + cellsArray.length + ' cells');
            }
        } catch (err) {
            console.error('Copy operation failed:', err);
            showToast('Copy failed');
        } finally {
            isCopying = false;
        }
    }

    function invertColor(color: string) {
        const match = String(color || '').match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
        if (!match) return color;

        const r = 255 - parseInt(match[1], 10);
        const g = 255 - parseInt(match[2], 10);
        const b = 255 - parseInt(match[3], 10);
        const a = match[4] ? match[4] : '1';
        return `rgba(${r}, ${g}, ${b}, ${a})`;
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
            const target = e.target as HTMLElement;
            if (target && target.classList && (target.classList.contains('col-resize-handle') || target.classList.contains('row-resize-handle'))) {
                return;
            }

            const cellTarget = target.closest('td, th') as HTMLElement;
            if (!cellTarget) return;

            e.preventDefault();

            if (cellTarget.classList.contains('col-header')) {
                const colIndex = parseInt(cellTarget.dataset.col!, 10);
                if (!e.shiftKey) {
                    lastSelectedColumn = colIndex;
                }
                selectColumn(colIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }

            if (cellTarget.classList.contains('row-header')) {
                const rowIndex = parseInt(cellTarget.dataset.row!, 10);
                if (!e.shiftKey) {
                    lastSelectedRow = rowIndex;
                }
                selectRow(rowIndex, e.ctrlKey || e.metaKey, e.shiftKey);
                return;
            }

            if (cellTarget.classList.contains('corner-cell')) {
                clearSelection();
                const allCells = table.querySelectorAll('td') as NodeListOf<HTMLElement>;
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

            if (cellTarget.tagName === 'TD') {
                const row = parseInt(cellTarget.dataset.row!, 10);
                const col = parseInt(cellTarget.dataset.col!, 10);

                if (e.ctrlKey || e.metaKey) {
                    if (cellTarget.classList.contains('selected')) {
                        cellTarget.classList.remove('selected');
                        selectedCells.delete(cellTarget);
                        if (cellTarget === activeCell) {
                            cellTarget.classList.remove('active-cell');
                            activeCell = null;

                            const remainingSelected = document.querySelector('td.selected') as HTMLElement;
                            if (remainingSelected) {
                                remainingSelected.classList.add('active-cell');
                                activeCell = remainingSelected;
                            }
                        }
                    } else {
                        cellTarget.classList.add('selected');
                        selectedCells.add(cellTarget);
                        if (activeCell) {
                            activeCell.classList.remove('active-cell');
                        }
                        cellTarget.classList.add('active-cell');
                        activeCell = cellTarget;
                    }
                    updateSelectionInfo();
                } else if (e.shiftKey && activeCell) {
                    const startRow = parseInt(activeCell.dataset.row!, 10);
                    const startCol = parseInt(activeCell.dataset.col!, 10);
                    selectRange(startRow, startCol, row, col);
                } else {
                    isSelecting = true;
                    selectionStart = { row, col };
                    selectCell(cellTarget);
                }
            }
        });

        table.addEventListener('mousemove', (e) => {
            if (isEditMode) return;
            if (!isSelecting || !selectionStart) return;

            // Track last mouse position for auto-scroll
            lastMousePos = { x: e.clientX, y: e.clientY };

            const target = (e.target as HTMLElement).closest('td') as HTMLElement;
            if (!target) return;

            const row = parseInt(target.dataset.row!, 10);
            const col = parseInt(target.dataset.col!, 10);

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
                        const htmlCell = nearestCell as HTMLElement;
                        const r = parseInt(htmlCell.dataset.row!, 10);
                        const c = parseInt(htmlCell.dataset.col!, 10);
                        if (!selectionEnd || selectionEnd.row !== r || selectionEnd.col !== c) {
                            selectionEnd = { row: r, col: c };
                            selectRange(selectionStart!.row, selectionStart!.col, r, c);
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
                    const active = document.activeElement as HTMLElement;
                    if (active && active.tagName === 'TD') {
                        const r = parseInt(active.getAttribute('data-row') || '0', 10);
                        const c = parseInt(active.getAttribute('data-col') || '0', 10);
                        const next = document.querySelector('td[data-row="' + (r + 1) + '"][data-col="' + c + '"]') as HTMLElement;
                        if (next) {
                            next.focus();
                            const range = document.createRange();
                            const sel = window.getSelection();
                            range.selectNodeContents(next);
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
                return;
            }

            if (isCmdOrCtrl && e.key.toLowerCase() === 'a') {
                e.preventDefault();
                const allCells = table.querySelectorAll('td') as NodeListOf<HTMLElement>;
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
            if (!(e.target as HTMLElement).closest('table') && !(e.target as HTMLElement).closest('.toolbar')) {
                clearSelection();
            }
        });
    }

    function ensureLinkTooltip(): HTMLElement {
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

    function showLinkTooltipForCell(cellEl: HTMLElement | null) {
        if (!currentSettings.hyperlinkPreview) return;
        if (!cellEl) return;
        const url = cellEl.getAttribute('data-hyperlink') || '';
        if (!url) return;

        const tt = ensureLinkTooltip();
        const urlEl = tt.querySelector('#linkTooltipUrl');
        if (urlEl) urlEl.textContent = url;

        const openBtn = tt.querySelector('#linkTooltipOpen') as HTMLElement;
        const copyBtn = tt.querySelector('#linkTooltipCopy') as HTMLElement;

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
            const t = e && (e.target as HTMLElement);
            const el = t && t.nodeType === 3 ? t.parentElement : t;
            const cell = el && el.closest ? el.closest('td[data-hyperlink]') : null;
            if (!cell) return;
            if (linkTooltipHideTimer) {
                clearTimeout(linkTooltipHideTimer);
                linkTooltipHideTimer = null;
            }
            showLinkTooltipForCell(cell as HTMLElement);
        });

        table.addEventListener('mouseout', (e) => {
            if (isEditMode) return;
            const toEl = e.relatedTarget as HTMLElement;
            if (!toEl) {
                scheduleHideLinkTooltip();
                return;
            }

            // If we are moving to an element inside the same cell, don't hide
            const fromCell = (e.target as HTMLElement).closest('td[data-hyperlink]');
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

    function applySettings(settings: any) {
        currentSettings = {
            firstRowIsHeader: settings && typeof settings.firstRowIsHeader === 'boolean' ? settings.firstRowIsHeader : currentSettings.firstRowIsHeader,
            stickyToolbar: settings && typeof settings.stickyToolbar === 'boolean' ? settings.stickyToolbar : currentSettings.stickyToolbar,
            stickyHeader: settings && typeof settings.stickyHeader === 'boolean' ? settings.stickyHeader : currentSettings.stickyHeader,
            hyperlinkPreview: settings && typeof settings.hyperlinkPreview === 'boolean' ? settings.hyperlinkPreview : currentSettings.hyperlinkPreview
        };

        // Update settings panel UI
        const chkHeader = document.getElementById('chkHeaderRow') as HTMLInputElement;
        const chkSticky = document.getElementById('chkStickyHeader') as HTMLInputElement;
        const chkToolbar = document.getElementById('chkStickyToolbar') as HTMLInputElement;
        const chkHyperlink = document.getElementById('chkHyperlinkPreview') as HTMLInputElement;
        if (chkHeader) chkHeader.checked = !!currentSettings.firstRowIsHeader;
        if (chkSticky) chkSticky.checked = !!currentSettings.stickyHeader;
        if (chkToolbar) chkToolbar.checked = !!currentSettings.stickyToolbar;
        if (chkHyperlink) chkHyperlink.checked = !!currentSettings.hyperlinkPreview;

        // Sticky toolbar behavior (CSV parity): when disabled, move toolbar into the scrollable content.
        const toolbar = document.querySelector('.toolbar');
        const headerBg = document.querySelector('.header-background') as HTMLElement;
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

        const chkHeaderRow = document.getElementById('chkHeaderRow') as HTMLInputElement;
        if (chkHeaderRow) chkHeaderRow.checked = !!currentSettings.firstRowIsHeader;
        const chkStickyToolbar = document.getElementById('chkStickyToolbar') as HTMLInputElement;
        if (chkStickyToolbar) chkStickyToolbar.checked = !!currentSettings.stickyToolbar;
        const chkStickyHeader = document.getElementById('chkStickyHeader') as HTMLInputElement;
        if (chkStickyHeader) {
            chkStickyHeader.checked = !!currentSettings.stickyHeader;
            chkStickyHeader.disabled = !currentSettings.firstRowIsHeader;
            if (chkStickyHeader.disabled) {
                chkStickyHeader.checked = false;
                currentSettings.stickyHeader = false;
                document.body.classList.remove('sticky-header-enabled');
            }
        }
        const chkHyperlinkPreview = document.getElementById('chkHyperlinkPreview') as HTMLInputElement;
        if (chkHyperlinkPreview) chkHyperlinkPreview.checked = !!currentSettings.hyperlinkPreview;

        if (!currentSettings.hyperlinkPreview) hideLinkTooltip();
    }

    function postSettings() {
        vscode.postMessage({ command: 'updateSettings', settings: currentSettings });
    }

    function setEditMode(enabled: boolean) {
        isEditMode = !!enabled;
        document.body.classList.toggle('edit-mode', isEditMode);

        const sheetSelector = document.getElementById('sheetSelector');
        const toggleExpandButton = document.getElementById('toggleExpandButton');
        const togglePlainViewButton = document.getElementById('togglePlainViewButton');
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
        if (togglePlainViewButton) togglePlainViewButton.classList.toggle('hidden', isEditMode);
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
            const htmlTd = td as HTMLElement;
            const currentText = normalizeCellText(htmlTd.textContent || '');
            htmlTd.textContent = currentText;
            htmlTd.dataset.originalText = currentText;
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
            (document.activeElement as HTMLElement).blur();
        }
        clearSelection();
        if (window.getSelection) {
            window.getSelection()!.removeAllRanges();
        }

        const edits: any[] = [];
        table.querySelectorAll('td[contenteditable="true"]').forEach(td => {
            const htmlTd = td as HTMLElement;
            const row = parseInt(htmlTd.getAttribute('data-rownum') || '0', 10);
            const col = parseInt(htmlTd.getAttribute('data-colnum') || '0', 10);
            if (!row || !col) return;

            const original = (htmlTd.dataset.originalText || '').replace(/\u00a0/g, '');
            const current = (htmlTd.textContent || '').replace(/\u00a0/g, '');

            if (current !== original) {
                edits.push({ row, col, value: current });
            }
        });

        setLoadingText('Saving worksheet...');
        showLoading();
        vscode.postMessage({ command: 'saveXlsxEdits', sheetIndex: currentWorksheet, edits });
    }

    function setExpandedMode(isExpanded: boolean) {
        document.body.classList.toggle('expanded-mode', !!isExpanded);

        const expandIcon = document.getElementById('expandIcon');
        const collapseIcon = document.getElementById('collapseIcon');
        const text = document.getElementById('expandButtonText');

        if (expandIcon) expandIcon.style.display = isExpanded ? 'none' : 'block';
        if (collapseIcon) collapseIcon.style.display = isExpanded ? 'block' : 'none';
        if (text) text.textContent = isExpanded ? 'Default' : 'Expand';
    }

    function wireSettingsUI() {
        const settings = [
            {
                id: 'chkHeaderRow',
                label: 'Header Row',
                onChange: (val: boolean) => {
                    currentSettings.firstRowIsHeader = val;
                    applySettings(currentSettings);
                    postSettings();
                },
                defaultValue: currentSettings.firstRowIsHeader
            },
            {
                id: 'chkStickyHeader',
                label: 'Sticky Header',
                onChange: (val: boolean) => {
                    currentSettings.stickyHeader = val;
                    applySettings(currentSettings);
                    postSettings();
                },
                defaultValue: currentSettings.stickyHeader
            },
            {
                id: 'chkStickyToolbar',
                label: 'Sticky Toolbar',
                onChange: (val: boolean) => {
                    currentSettings.stickyToolbar = val;
                    applySettings(currentSettings);
                    postSettings();
                },
                defaultValue: currentSettings.stickyToolbar
            },
            {
                id: 'chkHyperlinkPreview',
                label: 'Hyperlink Preview',
                onChange: (val: boolean) => {
                    currentSettings.hyperlinkPreview = val;
                    applySettings(currentSettings);
                    postSettings();
                },
                defaultValue: currentSettings.hyperlinkPreview
            }
        ];

        SettingsManager.renderPanel(document.getElementById('toolbar')!, 'settingsPanel', 'settingsCancelButton', settings);

        new SettingsManager('openSettingsButton', 'settingsPanel', 'settingsCancelButton', settings);
    }

    function attachHandlersOnce() {
        if (handlersAttached) return;
        handlersAttached = true;

        const toolbar = new ToolbarManager('toolbar');

        // Sheet Selector
        const sheetSelector = document.createElement('select');
        sheetSelector.id = 'sheetSelector';
        sheetSelector.className = 'sheet-selector';
        sheetSelector.title = 'Select sheet';
        sheetSelector.addEventListener('change', (e) => {
            if (isEditMode) return;
            currentWorksheet = parseInt((e.target as HTMLSelectElement).value, 10);
            clearSelection();
            renderWorksheet(currentWorksheet);
        });
        
        toolbar.prependElement(sheetSelector);

        // Inject tooltip if variables are present
        InfoTooltip.inject('toolbar', (window as any).viewImgUri, (window as any).logoSvgUri, 'table view');

        toolbar.setButtons([
            {
                id: 'toggleTableEditButton',
                icon: '',
                label: 'Edit Table',
                tooltip: 'Edit XLSX directly in the table (text only)',
                onClick: () => setEditMode(true)
            },
            {
                id: 'saveTableEditsButton',
                icon: '',
                label: 'Save',
                tooltip: 'Save table edits',
                hidden: true,
                onClick: () => saveEdits(true)
            },
            {
                id: 'cancelTableEditsButton',
                icon: '',
                label: 'Cancel',
                tooltip: 'Cancel table edits',
                hidden: true,
                onClick: () => {
                    setEditMode(false);
                    renderWorksheet(currentWorksheet);
                }
            },
            {
                id: 'toggleExpandButton',
                icon: Icons.Expand,
                label: 'Expand',
                tooltip: 'Toggle Column Widths (Default / Expand All)',
                onClick: () => {
                    if (isEditMode) return;
                    const btn = document.getElementById('toggleExpandButton');
                    const state = btn?.getAttribute('data-state') || 'default';
                    if (state === 'default') {
                        btn?.setAttribute('data-state', 'expanded');
                        if(btn) btn.innerHTML = Icons.Collapse + ' <span class="btn-label">Default</span>';
                        setExpandedMode(true);
                    } else {
                        btn?.setAttribute('data-state', 'default');
                        if(btn) btn.innerHTML = Icons.Expand + ' <span class="btn-label">Expand</span>';
                        setExpandedMode(false);
                    }
                }
            },
            {
                id: 'togglePlainViewButton',
                icon: Icons.Table,
                label: 'Plain',
                tooltip: 'Toggle Plain View (removes all styling)',
                onClick: () => {
                    if (isEditMode) return;
                    isPlainView = !isPlainView;
                    document.body.classList.toggle('plain-view', isPlainView);
                    
                    const btn = document.getElementById('togglePlainViewButton');
                    if (btn) {
                        const labelSpan = btn.querySelector('.btn-label');
                        if (labelSpan) labelSpan.textContent = isPlainView ? 'Styled' : 'Plain';
                    }
                    
                    rowCache.clear();
                    currentVisibleStart = 0;
                    currentVisibleEnd = 0;
                    renderWorksheet(currentWorksheet);
                }
            },
            {
                id: 'openSettingsButton',
                icon: Icons.Settings,
                tooltip: 'XLSX Settings',
                cls: 'icon-only',
                onClick: () => {}
            },
            {
                id: 'toggleBackgroundButton',
                icon: Icons.ThemeLight + Icons.ThemeDark + Icons.ThemeVSCode,
                tooltip: 'Toggle Theme',
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
            }
        ]);

        // Ensure the "Plain/Styled" toggle shows the correct label on initial render
        const togglePlainViewBtn = document.getElementById('togglePlainViewButton');
        if (togglePlainViewBtn) {
            const labelSpan = togglePlainViewBtn.querySelector('.btn-label');
            if (labelSpan) labelSpan.textContent = isPlainView ? 'Styled' : 'Plain';
        }

        if (typeof ThemeManager !== 'undefined') {
            new ThemeManager('toggleBackgroundButton', {
                onBeforeCycle: () => !isEditMode
            }, vscode);
        }

        wireSettingsUI();
    }

    function populateSheetSelector() {
        const selector = document.getElementById('sheetSelector') as HTMLSelectElement;
        if (!selector) return;

        selector.innerHTML = '';
        worksheetsMeta.forEach((ws, i) => {
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

        // Handle rowsData response for virtual scrolling
        if (message.command === 'rowsData') {
            virtualLoader.resolveRequest(message.requestId, message.rows || []);
            return;
        }

        // Handle initVirtualTable for virtual scrolling
        if (message.command === 'initVirtualTable') {
            worksheetsMeta = Array.isArray(message.worksheets) ? message.worksheets : [];
            currentWorksheet = 0;

            const rowHeaderWidth = typeof message.rowHeaderWidth === 'number' ? message.rowHeaderWidth : 60;
            document.documentElement.style.setProperty('--row-header-width', rowHeaderWidth + 'px');

            populateSheetSelector();
            attachHandlersOnce();
            const expandBtn = document.getElementById('toggleExpandButton');
            if (expandBtn) expandBtn.setAttribute('data-state', 'default');
            setExpandedMode(false);
            renderWorksheet(0);
            return;
        }

        // Legacy init handler (for backwards compatibility)
        if (message.command === 'init') {
            // Convert old format to new format
            const worksheets = Array.isArray(message.worksheets) ? message.worksheets : [];
            worksheetsMeta = worksheets.map((ws: any, index: number) => ({
                name: ws.name,
                index,
                totalRows: ws.data ? ws.data.maxRow : 0,
                columnCount: ws.data ? ws.data.maxCol : 0,
                columnWidths: ws.data ? ws.data.columnWidths : [],
                mergedCells: ws.data ? ws.data.mergedCells : []
            }));
            // Also cache all rows since they were sent
            worksheets.forEach((ws: any, wsIndex: number) => {
                if (ws.data && ws.data.rows) {
                    ws.data.rows.forEach((row: any, rowIndex: number) => {
                        if (wsIndex === 0) {
                            rowCache.set(rowIndex, row);
                        }
                    });
                }
            });
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
