import MarkdownIt from 'markdown-it';
// @ts-ignore
import taskLists from 'markdown-it-task-lists';
// @ts-ignore
import container from 'markdown-it-container';
// @ts-ignore
import deflist from 'markdown-it-deflist';
// @ts-ignore
import footnote from 'markdown-it-footnote';
// @ts-ignore
import sub from 'markdown-it-sub';
// @ts-ignore
import sup from 'markdown-it-sup';
// @ts-ignore
import ins from 'markdown-it-ins';
// @ts-ignore
import mark from 'markdown-it-mark';
// @ts-ignore
import abbr from 'markdown-it-abbr';

import hljs from 'highlight.js';
import { ThemeManager } from '../shared/themeManager';
import { SettingsManager } from '../shared/settingsManager';
import { ToolbarManager } from '../shared/toolbarManager';
import { Utils } from '../shared/utils';
import { Icons } from '../shared/icons';
import { vscode, debounce } from '../shared/common';
import { InfoTooltip } from '../shared/infoTooltip';

// ===== State =====
let isPreviewView = true;
let isEditMode = false;
let isSaving = false;
let shouldExitEditMode = false;
let originalContent = '';
let currentContent = '';
let toolbarManager: ToolbarManager | null = null;

// Settings
let currentSettings = {
    stickyToolbar: true,
    wordWrap: true,
    syncScroll: true,
    previewPosition: 'right',
    isMdEnabled: true
};

// ===== Utilities =====
const $ = Utils.$;

function setButtonsEnabled(enabled: boolean) {
    const ids = ['toggleViewButton', 'toggleEditModeButton', 'saveEditsButton',
        'cancelEditsButton', 'toggleBackgroundButton', 'openSettingsButton', 'disableMdEditorButton'];
    ids.forEach((id) => {
        const el = $(id) as HTMLButtonElement;
        if (el) el.disabled = !enabled;
    });
}

// ===== Markdown-it Setup =====
const md = new MarkdownIt({
    html: true,
    linkify: true,
    typographer: true,
    breaks: true, // GFM style line breaks
    highlight: function (str, lang) {
        if (lang && hljs.getLanguage(lang)) {
            try {
                return hljs.highlight(str, { language: lang }).value;
            } catch (__) {}
        }
        return ''; // use external default escaping
    }
});
md.use(taskLists, { enabled: false, label: true, labelAfter: true });
// eslint-disable-next-line @typescript-eslint/no-explicit-any
md.use(container as any, 'warning');
// eslint-disable-next-line @typescript-eslint/no-explicit-any
md.use(container as any, 'info');
// eslint-disable-next-line @typescript-eslint/no-explicit-any
md.use(container as any, 'error');
// eslint-disable-next-line @typescript-eslint/no-explicit-any
md.use(container as any, 'success');

md.use(deflist);
md.use(footnote);
md.use(sub);
md.use(sup);
md.use(ins);
md.use(mark);
md.use(abbr);

// Inject line numbers for sync scroll
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function injectLineNumbers(tokens: any, idx: number, options: any, env: any, self: any) {
    const token = tokens[idx];
    if (token.map && token.level === 0) {
        token.attrSet('data-line', String(token.map[0]));
    }
    return self.renderToken(tokens, idx, options, env, self);
}

// Apply to block-level elements
md.renderer.rules.paragraph_open = injectLineNumbers;
md.renderer.rules.heading_open = injectLineNumbers;
md.renderer.rules.bullet_list_open = injectLineNumbers;
md.renderer.rules.ordered_list_open = injectLineNumbers;
md.renderer.rules.blockquote_open = injectLineNumbers;
md.renderer.rules.hr = injectLineNumbers;

md.renderer.rules.table_open = function(tokens: any, idx: number, options: any, env: any, self: any) {
    tokens[idx].attrJoin('class', 'md-table');
    return injectLineNumbers(tokens, idx, options, env, self);
};

// Fence (code blocks) needs special handling as it's a self-closing block token in terms of rendering
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const defaultFence = md.renderer.rules.fence || function(tokens: any, idx: number, options: any, env: any, self: any) {
    return self.renderToken(tokens, idx, options);
};
// eslint-disable-next-line @typescript-eslint/no-explicit-any
md.renderer.rules.fence = function (tokens: any, idx: number, options: any, env: any, self: any) {
    if (tokens[idx].map && tokens[idx].level === 0) {
        tokens[idx].attrSet('data-line', String(tokens[idx].map[0]));
    }
    tokens[idx].attrJoin('class', 'code-block');
    return defaultFence(tokens, idx, options, env, self);
};

function parseGFM(text: string) {
    if (!text) return '';
    return md.render(text);
}

// ===== Rendering =====
function renderMarkdown(content: string) {
    const preview = $('markdownPreview');
    if (preview) {
        preview.innerHTML = parseGFM(content);
    }
}

// ===== Edit Mode (Split View) =====
function setEditMode(enabled: boolean) {
    isEditMode = enabled;
    document.body.classList.toggle('edit-mode', enabled);

    const editBtn = $('toggleEditModeButton');
    const saveBtn = $('saveEditsButton');
    const cancelBtn = $('cancelEditsButton');
    const container = $('markdownContainer');
    const editor = $('markdownEditor') as HTMLTextAreaElement;
    const preview = $('markdownPreview');

    if (saveBtn) saveBtn.classList.toggle('hidden', !enabled);
    if (cancelBtn) cancelBtn.classList.toggle('hidden', !enabled);
    if (editBtn) editBtn.classList.toggle('hidden', enabled);

    if (enabled) {
        // Enter split-view edit mode
        originalContent = currentContent;
        container?.classList.add('split-view');
        
        // Apply preview position (left or right)
        if (currentSettings.previewPosition === 'left') {
            container?.classList.add('preview-left');
        } else {
            container?.classList.remove('preview-left');
        }
        
        if (editor) editor.value = currentContent;
        
        // IMPORTANT: Scroll both editor and preview to TOP
        requestAnimationFrame(() => {
            if (editor) {
                editor.scrollTop = 0;
                editor.focus();
                editor.setSelectionRange(0, 0);
            }
            if (preview) preview.scrollTop = 0;
            
            setTimeout(() => {
                if (editor) editor.scrollTop = 0;
                if (preview) preview.scrollTop = 0;
            }, 50);
        });
    } else {
        // Exit edit mode
        container?.classList.remove('split-view');
        container?.classList.remove('preview-left');
        renderMarkdown(currentContent);
    }

    updateStatusInfo();
}

function performSave(exitAfterSave = false) {
    if (isSaving || !isEditMode) return;
    isSaving = true;
    shouldExitEditMode = exitAfterSave;
    setButtonsEnabled(false);

    const editor = $('markdownEditor') as HTMLTextAreaElement;
    if (editor) {
        currentContent = editor.value;
    }

    vscode.postMessage({ command: 'saveMarkdown', text: currentContent });
}

function cancelEdit() {
    currentContent = originalContent;
    const editor = $('markdownEditor') as HTMLTextAreaElement;
    if (editor) {
        editor.value = originalContent;
    }
    renderMarkdown(originalContent);
    setEditMode(false);
}

// ===== Live Preview =====
const debouncedRender = debounce((content: string) => {
    renderMarkdown(content);
}, 150);

function onEditorInput() {
    const editor = $('markdownEditor') as HTMLTextAreaElement;
    if (!editor) return;

    currentContent = editor.value;

    // Debounced live preview
    debouncedRender(currentContent);

    updateStatusInfo();
}

// ===== Sync Scroll (improved accuracy using line-based mapping) =====
let activeScrollSource: string | null = null; // 'editor' or 'preview' or null
let scrollTimeout: any = null;

function syncEditorToPreview() {
    if (!currentSettings.syncScroll) return;
    if (activeScrollSource === 'preview') return;

    activeScrollSource = 'editor';
    if (scrollTimeout) clearTimeout(scrollTimeout);

    const editor = $('markdownEditor') as HTMLTextAreaElement;
    const preview = $('markdownPreview');
    if (!editor || !preview) return;

    requestAnimationFrame(() => {
        // Calculate approximate line number
        const lineHeight = 21; 
        const scrollTop = editor.scrollTop;
        const lineNo = Math.floor(scrollTop / lineHeight);
        
        // Find element with data-line closest to lineNo
        const elements = Array.from(preview.querySelectorAll('[data-line]'));
        let target: Element | null = null;
        
        for (const el of elements) {
            const l = parseInt(el.getAttribute('data-line') || '0');
            if (l >= lineNo) {
                target = el;
                break;
            }
        }
        
        if (target) {
            preview.scrollTop = (target as HTMLElement).offsetTop;
        } else if (elements.length > 0 && lineNo > parseInt(elements[elements.length-1].getAttribute('data-line') || '0')) {
            // Scroll to bottom if past last element
            preview.scrollTop = preview.scrollHeight;
        }
    });

    scrollTimeout = setTimeout(() => {
        activeScrollSource = null;
    }, 100);
}

function syncPreviewToEditor() {
    if (!currentSettings.syncScroll) return;
    if (activeScrollSource === 'editor') return;

    activeScrollSource = 'preview';
    if (scrollTimeout) clearTimeout(scrollTimeout);

    const editor = $('markdownEditor') as HTMLTextAreaElement;
    const preview = $('markdownPreview');
    if (!editor || !preview) return;

    requestAnimationFrame(() => {
        // Find the first visible element in preview
        const elements = Array.from(preview.querySelectorAll('[data-line]'));
        const scrollTop = preview.scrollTop;
        
        let target: Element | null = null;
        for (const el of elements) {
            if ((el as HTMLElement).offsetTop >= scrollTop) {
                target = el;
                break;
            }
        }
        
        if (target) {
            const lineNo = parseInt(target.getAttribute('data-line') || '0');
            const lineHeight = 21; // Match the editor line height
            editor.scrollTop = lineNo * lineHeight;
        }
    });

    scrollTimeout = setTimeout(() => {
        activeScrollSource = null;
    }, 100);
}

// ===== UI Helpers =====
function showToast(message: string) {
    let toast = $('toastNotification');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toastNotification';
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
    if (toast) {
        const toastText = toast.querySelector('.toast-text') || $('toastText');
        if (toastText) toastText.textContent = message;
        toast.classList.add('show');
        setTimeout(() => toast!.classList.remove('show'), 2000);
    }
}

function updateStatusInfo() {
    const statusInfo = $('statusInfo');
    if (!statusInfo) return;

    const lines = currentContent.split('\n').length;
    const chars = currentContent.length;
    const words = currentContent.trim().split(/\s+/).filter(w => w).length;
    statusInfo.textContent = `${lines} lines, ${words} words, ${chars} chars`;
    statusInfo.style.display = 'block';
}

// ===== Settings =====
function applySettings(settings: any, persist = false) {
    if (!settings) return;
    currentSettings = { ...currentSettings, ...settings };

    const container = $('markdownContainer');
    const editor = $('markdownEditor');

    // Word wrap
    if (container) {
        container.classList.toggle('word-wrap', currentSettings.wordWrap);
    }
    if (editor) {
        editor.style.whiteSpace = currentSettings.wordWrap ? 'pre-wrap' : 'pre';
    }

    // Sticky toolbar
    document.body.classList.toggle('sticky-toolbar-enabled', currentSettings.stickyToolbar);

    // Preview position (left or right)
    if (container && isEditMode) {
        if (currentSettings.previewPosition === 'left') {
            container.classList.add('preview-left');
        } else {
            container.classList.remove('preview-left');
        }
    }

    // Update checkbox UI
    const chkWordWrap = $('chkWordWrap') as HTMLInputElement;
    const chkStickyToolbar = $('chkStickyToolbar') as HTMLInputElement;
    const chkSyncScroll = $('chkSyncScroll') as HTMLInputElement;
    const chkPreviewLeft = $('chkPreviewLeft') as HTMLInputElement;
    const chkMdEnabled = $('chkMdEnabled') as HTMLInputElement;

    if (chkWordWrap) chkWordWrap.checked = currentSettings.wordWrap;
    if (chkStickyToolbar) chkStickyToolbar.checked = currentSettings.stickyToolbar;
    if (chkSyncScroll) chkSyncScroll.checked = currentSettings.syncScroll;
    if (chkPreviewLeft) chkPreviewLeft.checked = currentSettings.previewPosition === 'left';
    if (chkMdEnabled) chkMdEnabled.checked = currentSettings.isMdEnabled;

    if (toolbarManager) {
        toolbarManager.setButtonVisibility('disableMdEditorButton', !!currentSettings.isMdEnabled);
        toolbarManager.setButtonVisibility('enableMdEditorButton', !currentSettings.isMdEnabled);
    }

    if (persist) {
        vscode.postMessage({ command: 'updateSettings', settings: currentSettings });
    }
}

function initializeSettings() {
    const settingsDefs = [
        {
            id: 'chkWordWrap',
            label: 'Word Wrap',
            defaultValue: currentSettings.wordWrap,
            onChange: (val: boolean) => {
                currentSettings.wordWrap = val;
                applySettings(currentSettings, true);
            }
        },
        {
            id: 'chkStickyToolbar',
            label: 'Sticky Toolbar',
            defaultValue: currentSettings.stickyToolbar,
            onChange: (val: boolean) => {
                currentSettings.stickyToolbar = val;
                applySettings(currentSettings, true);
            }
        },
        {
            id: 'chkSyncScroll',
            label: 'Sync Scrolling',
            defaultValue: currentSettings.syncScroll,
            onChange: (val: boolean) => {
                currentSettings.syncScroll = val;
                applySettings(currentSettings, true);
            }
        },
        {
            id: 'chkPreviewLeft',
            label: 'Preview on Left',
            defaultValue: currentSettings.previewPosition === 'left',
            onChange: (val: boolean) => {
                currentSettings.previewPosition = val ? 'left' : 'right';
                applySettings(currentSettings, true);
            }
        },
        {
            id: 'chkMdEnabled',
            label: 'Enable XLSX Viewer for .md files',
            defaultValue: currentSettings.isMdEnabled,
            onChange: (val: boolean) => {
                currentSettings.isMdEnabled = val;
                vscode.postMessage({ command: 'toggleMdAssociation', enable: val });
            }
        }
    ];

    // Render panel
    SettingsManager.renderPanel(document.body, 'settingsPanel', 'settingsCancelButton', settingsDefs);

    // Initialize manager
    new SettingsManager('openSettingsButton', 'settingsPanel', 'settingsCancelButton', settingsDefs);
}

// ===== Header Height =====
function updateHeaderHeight() {
    const toolbar = document.querySelector('.toolbar') as HTMLElement;
    if (toolbar) {
        const height = toolbar.offsetHeight;
        document.documentElement.style.setProperty('--header-height', height + 'px');
    }
}

// ===== Message Handler =====
window.addEventListener('message', (event) => {
    const m = event.data;

    switch (m.command) {
        case 'initMarkdown':
            const loading = $('loadingIndicator');
            if (loading) loading.style.display = 'none';

            currentContent = m.content || '';
            originalContent = currentContent;
            renderMarkdown(currentContent);
            updateStatusInfo();
            break;

        case 'initSettings':
        case 'settingsUpdated':
            applySettings(m.settings, false);
            break;

        case 'saveResult':
            isSaving = false;
            setButtonsEnabled(true);
            if (m.ok) {
                showToast('Saved');
                originalContent = currentContent;
                if (shouldExitEditMode) {
                    setEditMode(false);
                }
                shouldExitEditMode = false;
            } else {
                showToast('Error saving');
                shouldExitEditMode = false;
            }
            break;
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
                isPreviewView = !isPreviewView;
                vscode.postMessage({ command: 'toggleView', isPreviewView });
            }
        },
        {
            id: 'toggleEditModeButton',
            icon: Icons.SplitEdit,
            label: 'Split Edit',
            tooltip: 'Edit Markdown side-by-side',
            onClick: () => setEditMode(true)
        },
        {
            id: 'saveEditsButton',
            icon: '',
            label: 'Save',
            tooltip: 'Save Changes (Ctrl+S)',
            hidden: true,
            onClick: () => performSave(true)
        },
        {
            id: 'cancelEditsButton',
            icon: '',
            label: 'Cancel',
            tooltip: 'Cancel Changes (Esc)',
            hidden: true,
            onClick: () => cancelEdit()
        },
        {
            id: 'openSettingsButton',
            icon: Icons.Settings,
            tooltip: 'Settings',
            cls: 'icon-only',
            onClick: () => { /* Handled by wireSettingsUI */ }
        },
        {
            id: 'toggleBackgroundButton',
            icon: Icons.ThemeLight + Icons.ThemeDark + Icons.ThemeVSCode,
            tooltip: 'Toggle Theme',
            cls: 'edit-mode-hide',
            onClick: () => { /* Handled by ThemeManager */ }
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
            id: 'disableMdEditorButton',
            icon: Icons.ZapOff,
            label: 'Disable MD',
            tooltip: 'Disable XLSX Viewer for all Markdown files',
            cls: 'edit-mode-hide',
            onClick: () => {
                vscode.postMessage({ command: 'disableMdEditor' });
            }
        },
        {
            id: 'enableMdEditorButton',
            icon: Icons.Zap,
            label: 'Enable MD',
            tooltip: 'Enable XLSX Viewer for all Markdown files (Make Default)',
            cls: 'edit-mode-hide',
            hidden: true,
            onClick: () => {
                vscode.postMessage({ command: 'enableMdEditor' });
            }
        }
    ]);

    // Inject tooltip if variables are present
    InfoTooltip.inject('toolbar', (window as any).viewImgUri, (window as any).logoSvgUri, 'GitHub Flavored Markdown');

    // Theme manager
    new ThemeManager('toggleBackgroundButton', {
        onBeforeCycle: () => true
    }, vscode);
}

// ===== Keyboard Shortcuts =====
document.addEventListener('keydown', (e) => {
    const isCmdOrCtrl = e.ctrlKey || e.metaKey;

    if (isCmdOrCtrl && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (isEditMode) {
            performSave(false);
        }
        return;
    }

    if (e.key === 'Escape' && isEditMode) {
        e.preventDefault();
        cancelEdit();
        return;
    }
});

// ===== Editor Events =====
function wireEditor() {
    const editor = $('markdownEditor') as HTMLTextAreaElement;
    const preview = $('markdownPreview');
    if (!editor) return;

    editor.addEventListener('input', onEditorInput);

    editor.addEventListener('scroll', () => {
        syncEditorToPreview();
    });

    if (preview) {
        preview.addEventListener('scroll', () => {
            syncPreviewToEditor();
        });
    }

    editor.addEventListener('keydown', (e) => {
        if (e.key === 'Tab') {
            e.preventDefault();
            const start = editor.selectionStart;
            const end = editor.selectionEnd;
            const value = editor.value;

            if (e.shiftKey) {
                const lineStart = value.lastIndexOf('\n', start - 1) + 1;
                const lineContent = value.substring(lineStart, start);
                if (lineContent.startsWith('    ')) {
                    editor.value = value.substring(0, lineStart) + value.substring(lineStart + 4);
                    editor.selectionStart = editor.selectionEnd = start - 4;
                } else if (lineContent.startsWith('\t')) {
                    editor.value = value.substring(0, lineStart) + value.substring(lineStart + 1);
                    editor.selectionStart = editor.selectionEnd = start - 1;
                }
            } else {
                editor.value = value.substring(0, start) + '    ' + value.substring(end);
                editor.selectionStart = editor.selectionEnd = start + 4;
            }
            onEditorInput();
        }

        const pairs: {[key: string]: string} = { '(': ')', '[': ']', '{': '}', '"': '"', "'": "'", '`': '`' };
        if (pairs[e.key]) {
            const start = editor.selectionStart;
            const end = editor.selectionEnd;
            const selected = editor.value.substring(start, end);

            if (selected) {
                e.preventDefault();
                editor.value = editor.value.substring(0, start) + e.key + selected + pairs[e.key] + editor.value.substring(end);
                editor.selectionStart = start + 1;
                editor.selectionEnd = end + 1;
                onEditorInput();
            }
        }
    });
}

// ===== Hover Tooltip =====
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let hoverHideTimer: any = null;

function wireHoverTooltip() {
    const trigger = $('hoverPicTrigger');
    const tooltip = $('hoverTooltip');
    if (!trigger || !tooltip) return;

    function showTooltip() {
        if (hoverHideTimer) {
            clearTimeout(hoverHideTimer);
            hoverHideTimer = null;
        }
        const rect = trigger!.getBoundingClientRect();
        const tooltipWidth = tooltip!.offsetWidth || 300;
        const left = Math.max(8, Math.min(window.innerWidth - tooltipWidth - 8, rect.left - 100));
        const top = rect.bottom + 8;
        tooltip!.style.top = top + 'px';
        tooltip!.style.left = left + 'px';
        tooltip!.classList.remove('hidden');
        tooltip!.classList.add('visible');
    }

    function hideTooltip() {
        if (hoverHideTimer) {
            clearTimeout(hoverHideTimer);
        }
        tooltip!.classList.remove('visible');
        tooltip!.classList.add('hidden');
    }

    function hideTooltipDelayed() {
        if (hoverHideTimer) {
            clearTimeout(hoverHideTimer);
        }
        hoverHideTimer = setTimeout(() => hideTooltip(), 250);
    }

    trigger.addEventListener('mouseenter', showTooltip);
    trigger.addEventListener('mouseleave', hideTooltipDelayed);
    trigger.addEventListener('focus', showTooltip);
    trigger.addEventListener('blur', hideTooltip);

    tooltip!.addEventListener('mouseenter', () => {
        if (hoverHideTimer) {
            clearTimeout(hoverHideTimer);
            hoverHideTimer = null;
        }
    });
    tooltip!.addEventListener('mouseleave', hideTooltipDelayed);
}

// ===== Initialize =====
wireButtons();
initializeSettings();
wireEditor();
wireHoverTooltip();
updateHeaderHeight();

// Ensure settings are applied once toolbar is ready
if (currentSettings) {
    applySettings(currentSettings);
}

vscode.postMessage({ command: 'webviewReady' });
