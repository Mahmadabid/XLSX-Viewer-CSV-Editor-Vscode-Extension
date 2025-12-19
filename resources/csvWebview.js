/* global acquireVsCodeApi */

(function () {
  const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : { postMessage: () => {} };

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

  function $(id) {
    return document.getElementById(id);
  }

  function normalizeCellText(text) {
    if (!text) return '';
    return String(text).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trimEnd();
  }

  function setButtonsEnabled(enabled) {
    const ids = ['toggleViewButton', 'toggleTableEditButton', 'saveTableEditsButton', 'cancelTableEditsButton', 'toggleBackgroundButton'];
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

    const toggleViewButton = $('toggleViewButton');
    const toggleTableEditButton = $('toggleTableEditButton');

    // UX requirement: while editing, hide "Edit File" and "Edit Table".
    if (toggleViewButton) toggleViewButton.classList.toggle('hidden', isEditMode);
    if (toggleTableEditButton) toggleTableEditButton.classList.toggle('hidden', isEditMode);
    if (saveBtn) saveBtn.classList.toggle('hidden', !isEditMode);
    if (cancelBtn) cancelBtn.classList.toggle('hidden', !isEditMode);

    if (isEditMode) {
      clearSelection();
      captureOriginalCellValues();
    }

    applyEditModeToCells();
    refreshInteractions();
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

  function refreshInteractions() {
    if (isEditMode) return;
    initializeSelection();
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
    document.querySelectorAll('td.selected, td.active-cell').forEach((el) => {
      el.classList.remove('selected', 'active-cell');
    });
    selectedCells.clear();

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
    if (!ctrlKey && !shiftKey) {
      clearSelection();
    }

    if (shiftKey && lastSelectedColumn !== null) {
      if (!ctrlKey) {
        clearSelection();
      }
      const minCol = Math.min(lastSelectedColumn, columnIndex);
      const maxCol = Math.max(lastSelectedColumn, columnIndex);
      for (let col = minCol; col <= maxCol; col++) {
        selectedColumns.add(col);
        document.querySelectorAll('td[data-col="' + col + '"], th[data-col="' + col + '"]').forEach((cell) => {
          cell.classList.add('column-selected');
          if (cell.tagName === 'TD') selectedCells.add(cell);
        });
      }
    } else {
      if (ctrlKey && selectedColumns.has(columnIndex)) {
        selectedColumns.delete(columnIndex);
        document
          .querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]')
          .forEach((cell) => {
            cell.classList.remove('column-selected');
            if (cell.tagName === 'TD') selectedCells.delete(cell);
          });
      } else {
        selectedColumns.add(columnIndex);
        document
          .querySelectorAll('td[data-col="' + columnIndex + '"], th[data-col="' + columnIndex + '"]')
          .forEach((cell) => {
            cell.classList.add('column-selected');
            if (cell.tagName === 'TD') selectedCells.add(cell);
          });
      }
      lastSelectedColumn = columnIndex;
    }

    updateSelectionInfo();
  }

  function selectRow(rowIndex, ctrlKey, shiftKey) {
    if (!ctrlKey && !shiftKey) {
      clearSelection();
    }

    if (shiftKey && lastSelectedRow !== null) {
      if (!ctrlKey) {
        clearSelection();
      }
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
    } else {
      if (ctrlKey && selectedRows.has(rowIndex)) {
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
    }

    updateSelectionInfo();
  }

  function copySelectionToClipboard() {
    if (selectedCells.size === 0) return;

    const cellsArray = Array.from(selectedCells);
    const cellData = cellsArray
      .map((cell) => ({
        row: parseInt(cell.dataset.row, 10),
        col: parseInt(cell.dataset.col, 10),
        text: normalizeCellText(cell.textContent || ''),
      }))
      .sort((a, b) => a.row - b.row || a.col - b.col);

    const minRow = Math.min(...cellData.map((c) => c.row));
    const maxRow = Math.max(...cellData.map((c) => c.row));
    const minCol = Math.min(...cellData.map((c) => c.col));
    const maxCol = Math.max(...cellData.map((c) => c.col));

    const grid = [];
    for (let r = minRow; r <= maxRow; r++) {
      const row = [];
      for (let c = minCol; c <= maxCol; c++) {
        const found = cellsArray.find(
          (cell) => parseInt(cell.dataset.row, 10) === r && parseInt(cell.dataset.col, 10) === c
        );
        row.push(found ? normalizeCellText(found.textContent || '') : '');
      }
      grid.push(row);
    }

    const clipboardText = grid.map((row) => row.join('\t')).join('\n');

    navigator.clipboard.writeText(clipboardText).then(() => {
      selectedCells.forEach((cell) => {
        cell.classList.add('copying');
        setTimeout(() => cell.classList.remove('copying'), 200);
      });
    });
  }

  function initializeSelection() {
    const table = $('csv-table');
    if (!table) return;

    if (table.dataset.listenersAdded === 'true') return;
    table.dataset.listenersAdded = 'true';

    table.addEventListener('selectstart', (e) => {
      if (isEditMode) return;
      e.preventDefault();
      return false;
    });

    table.addEventListener('mousedown', (e) => {
      if (isEditMode) return;

      const target = e.target.closest('td, th');
      if (!target) return;

      e.preventDefault();

      if (target.classList.contains('col-header')) {
        const columnIndex = parseInt(target.dataset.col, 10);
        if (Number.isNaN(columnIndex)) return;
        if (!e.shiftKey) lastSelectedColumn = columnIndex;
        selectColumn(columnIndex, e.ctrlKey || e.metaKey, e.shiftKey);
        return;
      }

      if (target.classList.contains('row-header')) {
        const rowIndex = parseInt(target.dataset.row, 10);
        if (Number.isNaN(rowIndex)) return;
        if (!e.shiftKey) lastSelectedRow = rowIndex;
        selectRow(rowIndex, e.ctrlKey || e.metaKey, e.shiftKey);
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
            if (target === activeCell) {
              target.classList.remove('active-cell');
              activeCell = null;
            }
          } else {
            target.classList.add('selected');
            selectedCells.add(target);
            if (activeCell) activeCell.classList.remove('active-cell');
            target.classList.add('active-cell');
            activeCell = target;
            startCell = coords;
          }
          updateSelectionInfo();
          return;
        }

        if (e.shiftKey && startCell) {
          e.stopPropagation();
          selectCellsInRange(startCell, coords);
          updateSelectionInfo();
          return;
        }

        clearSelection();
        isSelecting = true;
        startCell = coords;
        endCell = coords;
        target.classList.add('selected');
        target.classList.add('active-cell');
        selectedCells.add(target);
        activeCell = target;
        updateSelectionInfo();
      }
    });

    table.addEventListener('mousemove', (e) => {
      if (isEditMode) return;
      if (!isSelecting || !startCell) return;
      const target = e.target.closest('td');
      if (!target) return;
      const coords = getCellCoordinates(target);
      if (!coords) return;
      if (!endCell || coords.row !== endCell.row || coords.col !== endCell.col) {
        endCell = coords;
        selectCellsInRange(startCell, endCell);
      }
    });

    document.addEventListener('mouseup', () => {
      isSelecting = false;
    });

    if (!window.__csvViewerDocKeyListenerAdded) {
      window.__csvViewerDocKeyListenerAdded = true;

      document.addEventListener('keydown', (e) => {
        if (isEditMode) return;

        if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
          e.preventDefault();
          copySelectionToClipboard();
        }

        if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
          e.preventDefault();
          const allCells = document.querySelectorAll('td[data-row][data-col]');
          if (allCells.length > 0) {
            clearSelection();
            allCells.forEach((cell) => {
              cell.classList.add('selected');
              selectedCells.add(cell);
            });
            const firstCell = allCells[0];
            if (firstCell) {
              firstCell.classList.add('active-cell');
              activeCell = firstCell;
              startCell = getCellCoordinates(firstCell);
            }
            updateSelectionInfo();
          }
        }
      });

      document.addEventListener('click', (e) => {
        if (!e.target.closest('#csv-table')) {
          clearSelection();
        }
      });
    }
  }

  // Message handling
  window.addEventListener('message', (event) => {
    const message = event.data;

    if (message.command === 'initTable') {
      const table = $('csv-table');
      if (!table) return;
      const thead = table.querySelector('thead');
      const tbody = table.querySelector('tbody');
      if (!thead || !tbody) return;
      thead.innerHTML = message.headerHtml || '';
      tbody.innerHTML = message.rowsHtml || '';
      applyEditModeToCells();
      refreshInteractions();
      return;
    }

    if (message.command === 'appendRows') {
      const table = $('csv-table');
      if (!table) return;
      const tbody = table.querySelector('tbody');
      if (!tbody) return;
      tbody.insertAdjacentHTML('beforeend', message.rowsHtml || '');
      applyEditModeToCells();
      refreshInteractions();
      return;
    }

    if (message.command === 'saveResult') {
      isSaving = false;
      setButtonsEnabled(true);
      if (message.ok) {
        clearOriginalCellValues();
        setEditMode(false);
      }
    }
  });

  // Button wiring
  function wireButtons() {
    const toggleViewButton = $('toggleViewButton');
    const toggleTableEditButton = $('toggleTableEditButton');
    const saveTableEditsButton = $('saveTableEditsButton');
    const cancelTableEditsButton = $('cancelTableEditsButton');
    const toggleBackgroundButton = $('toggleBackgroundButton');

    if (toggleViewButton) {
      toggleViewButton.addEventListener('click', () => {
        isTableView = !isTableView;
        vscode.postMessage({ command: 'toggleView', isTableView });
      });
    }

    if (toggleTableEditButton) {
      toggleTableEditButton.addEventListener('click', () => {
        setEditMode(!isEditMode);
      });
    }

    if (saveTableEditsButton) {
      saveTableEditsButton.addEventListener('click', () => {
        if (isSaving) return;
        isSaving = true;
        setButtonsEnabled(false);
        const csvText = serializeTableToCsv();
        vscode.postMessage({ command: 'saveCsv', text: csvText });
      });
    }

    if (cancelTableEditsButton) {
      cancelTableEditsButton.addEventListener('click', () => {
        if (!isEditMode) return;
        restoreOriginalCellValues();
        setEditMode(false);
      });
    }

    if (toggleBackgroundButton) {
      toggleBackgroundButton.addEventListener('click', () => {
        document.body.classList.toggle('alt-bg');
        const isDarkMode = document.body.classList.contains('alt-bg');

        const lightIcon = $('lightIcon');
        const darkIcon = $('darkIcon');
        if (lightIcon && darkIcon) {
          if (isDarkMode) {
            lightIcon.style.display = 'block';
            darkIcon.style.display = 'none';
          } else {
            lightIcon.style.display = 'none';
            darkIcon.style.display = 'block';
          }
        }

        document.querySelectorAll('td[data-default-bg="true"]').forEach((cell) => {
          cell.style.backgroundColor = isDarkMode ? 'rgb(33, 33, 33)' : 'rgb(255, 255, 255)';
        });

        document.querySelectorAll('td[data-default-bg="true"][data-default-color="true"]').forEach((cell) => {
          cell.style.color = isDarkMode ? 'rgb(255, 255, 255)' : 'rgb(0, 0, 0)';
        });
      });
    }
  }

  // Kickoff
  try {
    wireButtons();
    vscode.postMessage({ command: 'webviewReady' });
  } catch {
    // ignore
  }
})();
