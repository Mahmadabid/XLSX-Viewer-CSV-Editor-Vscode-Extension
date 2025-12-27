# Changelog

## v1.5.2 - Premium UX Refinements & Bug Fixes
- **Premium CSV Editor UX**:
  - Added **Undo (Ctrl+Z)** and **Redo (Ctrl+Y)** functionality in table edit mode.
  - Improved Keyboard Navigation: **Enter** key now moves to the cell below instead of adding a newline.
  - Refined **Save Behavior**: Ctrl+S now saves changes, clears selection, and blurs active cell without exiting edit mode.
  - Added visual **Save Confirmation** (premium horizontal toast with green tick).
  - Added **Edit Mode Indicator**: Sharp outer border and active cell highlighting.
  - Fixed horizontal scrolling and text truncation issues in edit mode.
  - Added subtle hover highlights for table cells.

## v1.5.1 - CSV Table Editing
- Added in-table **Edit Table** mode for CSV files with **Save** and **Cancel** actions.
- While editing, the **Edit File** and **Edit Table** buttons are hidden to reduce accidental mode switching.
- Improved webview reliability by waiting for the webview to be ready before streaming table rows.

## v1.5.0 - Merged Cells & Resizing Support

### **Merged Cell Support:**
- Full support for both horizontal and vertical merged cells from Excel files
- Proper content alignment and positioning within merged cells
- Maintains original Excel formatting and alignment

### **Interactive Resizing:**
- Drag column borders to resize column widths
- Drag row borders to resize row heights
- Visual resize handles on headers with hover effects
- Real-time size indicators during resizing

### **Auto-Fit Functionality:**
- Auto-fit button to automatically resize all columns based on content
- Double-click column borders to auto-fit individual columns
- Double-click row borders to auto-fit individual rows
- Smart content-based sizing with maximum width limits

## v1.4.0 - Excel-like Multi-Selection & Copy
- **Multi-Selection for Rows/Columns:**
  - Hold <kbd>Ctrl</kbd> and click multiple row or column headers to select/deselect multiple rows or columns.
  - Hold <kbd>Shift</kbd> and click to select a range of rows or columns.
- **Excel/Google Sheets Compatible Copy:**
  - Pasting into Excel or Google Sheets will place data in the correct cells, not a single cell.
- **Improved Selection Management:**
  - Visual feedback for multi-row and multi-column selection.
  - Selection info box shows the size of the current selection, Displayed at bottom right corner.

## v1.3.0 - Enhanced Selection Features
- **Text Selection**: Added text selection for copying with ease.
- **Cell Selection**: Improved cell, row, and column selection functionality
- **Dark Mode Support**: Enhanced text selection visibility in both light and dark modes
- **UI Improvements**: Better visual feedback for selections and copying

## v1.2.0 - Enhanced Toggle Background
- **Improved Toggle Background**: Updated toggle button functionality for light and dark modes with alternating icons.
- **UI Enhancements**: Adjusted icon sizes and improved visual consistency.

## v1.1.0 - XLSX Viewer & CSV Editor (New Name)
- **New Name**: Previously known as `XLSX Viewer`.
- **Features**: Added CSV file editing capabilities in a structured table view.
- **Bug Fixes**: Improved performance and UI enhancements.

## v1.0.0 - XLSX Viewer
- Initial release with basic functionality for viewing Excel files.