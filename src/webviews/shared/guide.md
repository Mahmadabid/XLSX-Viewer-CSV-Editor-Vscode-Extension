# Shared Webview Components Guide

This directory contains shared components and utilities used across the different webviews (Markdown, Table, XLSX).

## Components

### `common.ts`
Contains common utilities and constants.
- `vscode`: Wrapper for the VS Code API.
- `VirtualScrollConfig`: Constants for virtual scrolling (ROW_HEIGHT, BUFFER_ROWS, CHUNK_SIZE).
- `debounce`: Utility function for debouncing events.

### `virtualLoader.ts`
Handles data fetching for virtual scrolling.
- `VirtualLoader`: Class that manages requests to the extension host for row data.

### `themeManager.ts`
Manages theme switching (Light, Dark, VS Code High Contrast).

### `settingsManager.ts`
Manages the settings panel and persistence.

### `toolbarManager.ts`
Manages the toolbar buttons and their states.

### `utils.ts`
General utilities.
- `$`: DOM element selector.
- `showToast`: Displays a toast notification.
- `writeToClipboardAsync`: Handles clipboard operations.
- `escapeHtml`: Escapes HTML strings.
- `normalizeCellText`: Normalizes text for cell display.

### `icons.ts`
Contains SVG icon strings used in the toolbar.

### `infoTooltip.ts`
Injects an informational tooltip into the toolbar, typically used to explain how to switch views.
