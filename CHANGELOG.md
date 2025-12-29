# Changelog

## v1.5.6 - VS Code Theme Support
- **VS Code Theme Support**:
  - Added **VS Code** theme option that mirrors the editor's native theme (Light / Dark / High Contrast).
  - New `ThemeManager` component centralizes theme logic and persistence.
  - **Persistent Theme**: The extension now automatically remembers your last used theme and applies it to new files.
  - Interactive tooltip on the theme button with quick-switch action and accessibility labels.

## v1.5.5 - Dark Mode Fixes
- **Dark Mode Fixes**:
  - Corrected text color in dark mode for XLSX views to ensure readability.
  - Updated CSS rules to maintain consistent appearance across different themes.
  - Ensured that default cell colors adapt properly in dark mode without losing visibility.

## v1.5.4 - XLSX Editing & UI Improvements
- **New Name & Description**:
  - Extension renamed from `XLSX Viewer & CSV Editor` to `XLSX & CSV Editor` to better reflect its dual functionality.
  - Updated extension description to highlight both XLSX viewing/editing and CSV editing capabilities.
- **XLSX Editing & Toolbar (New approach)**:
  - Introduced **in-webview table editing** for XLSX files: toggle **Edit** to make changes, then **Save** to persist changes back to the `.xlsx` file or **Cancel** to discard.
  - **Implementation detail**: edits are applied in the webview and written to disk using ExcelJS; the extension attempts to preserve formatting and merged cells where possible.
  - Added **Undo/Redo** support and keyboard shortcuts for edit mode (Ctrl+S / Ctrl+Z / Ctrl+Y / Enter).
  - **Toolbar & Settings parity**: toolbar controls and the Settings panel (header toggle, sticky header, sticky toolbar, hyperlink preview) were added to XLSX views to match CSV editor UX.
  - **UX improvements**: refined toolbar responsiveness, consistent sticky headers, and polished visual styles across XLSX and CSV editors.



## v1.5.3 - UI Polish & Settings UX
- **UI Polish & Settings UX**:
  - Redesigned **Settings panel** with backdrop blur, smoother rounded corners, grouped checkboxes, and responsive Cancel button that wraps on small screens.
  - Settings panel features:
    - **Header Row**: toggle the first row to be treated as header (bold first row).
    - **Sticky Header**: keep the first row sticky when header is enabled.
    - **Sticky Toolbar**: keep the toolbar fixed at the top of the editor.

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