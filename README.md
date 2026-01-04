# XLSX, CSV, TSV & Markdown Editor - VS Code Extension

This is an open-source project that allows you to view and edit XLSX files with styles, fonts, and colors from Excel files, supporting multiple sheets. Additionally, it provides a table view and editing capabilities for CSV, TSV, and Github Flavored Markdown (.md) files directly in VS Code.

## 📌 Overview

XLSX, CSV, TSV & Markdown Editor is a powerful **Visual Studio Code** extension that allows users to **open, view, and edit Excel files (.xlsx)**, **CSV**, **TSV**, and **Markdown (.md) files directly within VS Code**. It provides a seamless, lightweight experience without requiring external software like Microsoft Excel or Google Sheets.

**🚨 Important Note**: This extension was previously named **`XLSX Viewer & CSV Editor`**. It has since been renamed to **`XLSX, CSV, TSV & Markdown Editor`** to reflect added support for TSV and Markdown (.md) files.

## 🚀 Features

### XLSX Viewing
✅ **Fast & Lightweight** - View Excel spreadsheets quickly within VS Code\
✅ **Retains Formatting** - Keeps cell styles, colors, and text formatting\
✅ **Multiple Sheet Support** - View all sheets in your Excel workbook\
✅ **Toggle Background Mode** - Easily switch between Light, Dark, and **VS Code** (mirrors editor theme) backgrounds\
✅ **Persistent Theme** - Remembers your last used theme across files and sessions\
✅ **Interactive Table View** - Display spreadsheet data in a structured HTML table\
✅ **Color Detection & Conversion** - Converts ARGB Excel colors into CSS-compatible formats\
✅ **Plain View Mode** - Strip all Excel styling to view data like CSV/TSV for cleaner inspection

### XLSX Editing
✅ **In-Table Editing** - Edit XLSX sheets directly in the webview table with **Edit**, **Save**, and **Cancel** actions.
✅ **Undo/Redo & Shortcuts** - Support for <kbd>Ctrl+Z</kbd>/<kbd>Ctrl+Y</kbd>, <kbd>Ctrl+S</kbd> to save, and <kbd>Enter</kbd> to navigate cells while editing.
✅ **Toolbar & Settings Panel** - Toolbar and settings (sticky toolbar, header options, hyperlink preview, toggle header) are available in the XLSX view for quick access and parity with the CSV editor.
✅ **How it works** - Changes are made in the webview and persisted back to the original `.xlsx` file when you click **Save**. The extension uses ExcelJS to write workbook changes while attempting to preserve styles and formatting where possible. **Cancel** discards any unsaved changes.

### Settings for XLSX Editing
- `xlsxViewer.xlsx.firstRowIsHeader` — Treat the first row as a header (renders bold).
- `xlsxViewer.xlsx.stickyHeader` — Make the header row sticky when enabled.
- `xlsxViewer.xlsx.stickyToolbar` — Keep the workbook toolbar fixed at the top of the editor.
- `xlsxViewer.xlsx.hyperlinkPreview` — Show hover previews for hyperlinks with Open/Copy actions.

> ⚠️ Notes & Limitations: Formulas are not automatically recalculated by the webview; very large workbooks or complex charts may be slow to edit. The editor aims to preserve styles and merged-cell behavior where possible.


### 🆕 Merged Cells & Resizing Support
✅ **Merged Cell Support** - Full support for both horizontal and vertical merged cells from Excel files, with proper content alignment and original Excel formatting\
✅ **Interactive Resizing** - Drag column/row borders to resize, with visual resize handles, hover effects, and real-time size indicators\
✅ **Auto-Fit Functionality** - Auto-fit button and double-click to auto-fit columns/rows based on content, with smart content-based sizing and width limits

### 🆕 Excel-like Multi-Selection & Copy
✅ **Multi-Selection for Rows/Columns** - Hold <kbd>Ctrl</kbd> to select/deselect multiple rows or columns, <kbd>Shift</kbd> to select a range\
✅ **Excel/Google Sheets Compatible Copy** - Copying and pasting preserves cell structure in Excel/Google Sheets\
✅ **Improved Selection Management** - Visual feedback for multi-row/column selection and selection size info box in the bottom right corner.

### CSV & TSV Editing
✅ **Table View** - View CSV and TSV files in a structured table format\
✅ **Virtualized Rendering for Large Files** - Virtualization (windowed rendering) improves performance and reduces memory usage by rendering only visible rows for both CSV and TSV files\
✅ **Edit Table Mode** - Edit directly in the table with **Save**, **Cancel**, and **Undo/Redo** support\
✅ **Excel-like Shortcuts** - <kbd>Ctrl+S</kbd> to save, <kbd>Enter</kbd> to move down, <kbd>Ctrl+Z/Y</kbd> for undo/redo\
✅ **Premium UI** - Smooth animations, sticky headers, and Apple-like visual feedback\
✅ **Edit File** - Open the CSV/TSV in VS Code’s default text editor when needed\

**TSV notes:** Copying from the table and pasting into other apps will use tab-delimited data for TSV files; serialization and escaping behave like CSV but use `\t` as the delimiter for TSV.

### Markdown Viewer & Editor
✅ **Preview & Edit Modes** - Toggle between a rendered Markdown preview and an inline editor with **Edit Preview**, **Save**, and **Cancel**.
✅ **Toolbar & Settings** - Word Wrap, Line Numbers, Sticky Toolbar and theme toggle available in the editor toolbar.
✅ **Local Rendering** - Fast, lightweight rendering for common Markdown syntax (headers, lists, code blocks, tables, task lists, images, links). **Uses GitHub-Flavored Markdown (GFM) for rendering (tables, task lists, fenced code blocks, etc.).**
✅ **Theme Support** - Respects Light / Dark / VS Code theme using the shared `ThemeManager`.
✅ **Open with Preview** - Open `.md` files with the extension via the editor toolbar or the command palette.

## 🛠️ Installation

1. Open **VS Code**
2. Go to the **Extensions Marketplace** (`Ctrl+Shift+X`)
3. Search for `XLSX, CSV, TSV & Markdown Editor`
4. Click **Install**
5. Open any `.xlsx`, `.csv`, `.tsv` or `.md` file to start viewing or editing!

Alternatively, you can install it manually using:

```
code --install-extension muhammad-ahmad.xlsx-viewer
```
## 📖 Usage

### For XLSX Files
1. Open VS Code
2. Open an **.xlsx file**
3. View and analyze your data in an HTML table format
4. Use the **toggle button** to switch the background color
5. Navigate between sheets using the sheet selector

### For CSV Files
1. Open any **.csv file**
2. Click **Edit Table** to enable in-table editing
3. Click **Save** to write changes to the CSV file, or **Cancel** to discard edits
4. To edit raw text, click **Edit File** (opens VS Code’s default editor)
5. To return to the table view, use **Open in Table View** (table icon) in the editor toolbar

## 🛠️ Contributing

We welcome contributions! Feel free to **submit issues, feature requests, or pull requests** in the GitHub repository.

## 📜 License

This project is licensed under the **MIT License** - feel free to use and modify it.

---

### 💬 Help & Feedback
If you have any questions, encounter issues, or want to provide feedback, please fill out our [Help & Feedback Form](https://docs.google.com/forms/d/e/1FAIpQLSe5AqE_f1-WqUlQmvuPn1as3Mkn4oLjA0EDhNssetzt63ONzA/viewform).

## ⭐ Support

If you find this extension helpful, please **rate it on the VS Code Marketplace** and share it with others!

---

📢 **Follow us for updates!**\
🔗 GitHub: [XLSX, CSV, TSV & Markdown Editor Github Link](https://github.com/Mahmadabid/XLSX-CSV-TSV-MARKDOWN-Editor-Vscode-Extension)\
🔗 Marketplace: [VS Code Extension Link](https://marketplace.visualstudio.com/items?itemName=muhammad-ahmad.xlsx-viewer)\
🔗 Open VSX: [Open VSX Link](https://open-vsx.org/extension/muhammad-ahmad/xlsx-viewer)
