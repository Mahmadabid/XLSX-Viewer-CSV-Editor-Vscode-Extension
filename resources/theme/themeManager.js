/* global acquireVsCodeApi */

/**
 * ThemeManager handles cycling between Light, Dark, and VS Code themes.
 * It manages the UI button icons, body classes, and persistence.
 */
class ThemeManager {
    constructor(buttonId, options = {}, vscodeApi = null) {
        this.button = document.getElementById(buttonId);
        this.vscodeApi = vscodeApi;
        this.options = Object.assign({
            persistKey: 'last_used_theme'
        }, options);

        this.themes = ['light', 'dark', 'vscode'];
        this.init();
    }

    init() {
        if (!this.button) return;

        // Listen for theme messages from VS Code
        window.addEventListener('message', event => {
            const msg = event.data;
            if (msg && msg.type === 'setTheme') {
                this.handleVsCodeThemeChange(msg.kind);
            }
        });

        this.button.addEventListener('click', () => {
            this.cycleTheme();
        });
        
        // Initial apply
        this.applyTheme(this.getStoredTheme(), false);
    }

    getStoredTheme() {
        // 1. Try to get from vscode state if available (per-file persistence for current session)
        if (this.vscodeApi && typeof this.vscodeApi.getState === 'function') {
            const state = this.vscodeApi.getState();
            if (state && state.theme) return state.theme;
        }

        // 2. Fallback to global localStorage (last used theme)
        try {
            const lastUsed = localStorage.getItem(this.options.persistKey);
            if (lastUsed && this.themes.includes(lastUsed)) return lastUsed;
        } catch (e) { /* ignore */ }

        return 'vscode';
    }

    setStoredTheme(theme) {
        // Always update vscode state for per-file consistency if available
        if (this.vscodeApi && typeof this.vscodeApi.getState === 'function') {
            const state = this.vscodeApi.getState() || {};
            state.theme = theme;
            this.vscodeApi.setState(state);
        }

        // Update global localStorage (last used theme)
        try {
            localStorage.setItem(this.options.persistKey, theme);
        } catch (e) { /* ignore */ }
    }

    applyTheme(theme, save = true) {
        document.body.classList.remove('dark-mode', 'vscode-theme');
        
        if (theme === 'dark') {
            document.body.classList.add('dark-mode');
        } else if (theme === 'vscode') {
            document.body.classList.add('vscode-theme');
        }

        if (save) {
            this.setStoredTheme(theme);
        }

        this.updateIcons(theme);
    }

    cycleTheme() {
        if (typeof this.options.onBeforeCycle === 'function') {
            if (this.options.onBeforeCycle() === false) return;
        }
        const current = this.getStoredTheme();
        const currentIndex = this.themes.indexOf(current);
        const nextIndex = (currentIndex + 1) % this.themes.length;
        const nextTheme = this.themes[nextIndex];
        this.applyTheme(nextTheme);
    }

    updateIcons(currentTheme) {
        const nextIndex = (this.themes.indexOf(currentTheme) + 1) % this.themes.length;
        const nextTheme = this.themes[nextIndex];

        const icons = {
            light: document.getElementById('lightIcon'),
            dark: document.getElementById('darkIcon'),
            vscode: document.getElementById('vscodeIcon')
        };

        Object.keys(icons).forEach(key => {
            if (icons[key]) {
                icons[key].style.display = (key === nextTheme) ? 'block' : 'none';
            }
        });

        // Update button tooltip / aria-label to indicate the next action
        const label = (t) => {
            if (!t) return '';
            if (t === 'vscode') return 'VS Code';
            return t.charAt(0).toUpperCase() + t.slice(1);
        };

        // Use a concise title for the native tooltip (avoid duplicating the 'Switch' button text)
        const titleTextShort = `${label(nextTheme)} mode (from ${label(currentTheme)})`;
        // Keep a full descriptive text for the interactive tooltip and accessibility
        const tooltipTextFull = `Switch to ${label(nextTheme)} mode from ${label(currentTheme)}`;

        if (this.button) {
            this.button.title = titleTextShort;
            // ARIA should be descriptive for screen readers
            this.button.setAttribute('aria-label', tooltipTextFull);
            // expose current/next for testing/debugging
            this.button.dataset.nextTheme = nextTheme;
            this.button.dataset.currentTheme = currentTheme;
        }

        // Interactive tooltip handling (create on demand)
        this._ensureTooltip();
        this.tooltipText = tooltipTextFull;
        if (this.tooltipSpan) this.tooltipSpan.innerHTML = `Switch to <strong>${label(nextTheme)}</strong> mode from ${label(currentTheme)}`;
        // Ensure tooltip element has accessible label for screen readers
        if (this.tooltipEl) this.tooltipEl.setAttribute('aria-label', this.tooltipText);
    }

    // Creates and wires an interactive tooltip (shown on hover/focus, contains a "Switch" button)
    createTooltip() {
        if (!this.button || this.tooltipEl) return;

        this.tooltipEl = document.createElement('div');
        this.tooltipEl.className = 'theme-tooltip';
        this.tooltipEl.setAttribute('role', 'dialog');
        this.tooltipEl.setAttribute('aria-hidden', 'true');

        this.tooltipSpan = document.createElement('span');
        this.tooltipSpan.className = 'theme-tooltip-text';
        this.tooltipSpan.style.whiteSpace = 'nowrap';

        // Append span and make the tooltip element clickable to perform the switch
        this.tooltipEl.appendChild(this.tooltipSpan);
        document.body.appendChild(this.tooltipEl);

        // show/hide wiring
        this.button.addEventListener('mouseenter', () => this.showTooltip());
        this.button.addEventListener('mouseleave', () => this.hideTooltipDelayed());
        this.button.addEventListener('focus', () => this.showTooltip());
        this.button.addEventListener('blur', () => this.hideTooltip());

        // Make tooltip actionable (click or key activate on the tooltip itself)
        this.tooltipEl.addEventListener('click', (e) => {
            e.stopPropagation();
            this.cycleTheme();
            this.hideTooltip();
        });
        this.tooltipEl.tabIndex = 0; // make focusable
        this.tooltipEl.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                this.cycleTheme();
                this.hideTooltip();
            }
        });

        this.tooltipEl.addEventListener('mouseenter', () => this.clearHideTimer());
        this.tooltipEl.addEventListener('mouseleave', () => this.hideTooltipDelayed());
    }

    _ensureTooltip() {
        if (!this.tooltipEl) {
            try {
                this.createTooltip();
            } catch (err) {
                // ignore DOM errors
            }
        }
    }

    showTooltip() {
        if (!this.tooltipEl || !this.button) return;
        this.clearHideTimer();
        const rect = this.button.getBoundingClientRect();
        // Position centered above the button
        const el = this.tooltipEl;
        el.style.position = 'fixed';
        requestAnimationFrame(() => {
            const elW = el.offsetWidth || 160;
            const left = Math.max(8, Math.min(window.innerWidth - elW - 8, rect.left + (rect.width - elW) / 2));
            const top = rect.bottom + 8;
            el.style.left = `${left}px`;
            el.style.top = `${top}px`;
            el.classList.add('visible');
            el.setAttribute('aria-hidden', 'false');
        });
    }

    hideTooltip() {
        if (!this.tooltipEl) return;
        this.clearHideTimer();
        this.tooltipEl.classList.remove('visible');
        this.tooltipEl.setAttribute('aria-hidden', 'true');
    }

    hideTooltipDelayed() {
        this.clearHideTimer();
        this._hideTimer = setTimeout(() => this.hideTooltip(), 250);
    }

    clearHideTimer() {
        if (this._hideTimer) {
            clearTimeout(this._hideTimer);
            this._hideTimer = null;
        }
    }

    handleVsCodeThemeChange(kind) {
        // kind: 1=Light, 2=Dark, 3=HighContrast
        document.body.classList.toggle('vscode-dark', kind === 2);
        document.body.classList.toggle('vscode-light', kind === 1);
        document.body.classList.toggle('vscode-highcontrast', kind === 3);

        // If current preference is vscode, re-apply to ensure classes are fresh
        if (this.getStoredTheme() === 'vscode') {
            this.applyTheme('vscode', false);
        }
    }
}

// Export for use in other scripts
window.ThemeManager = ThemeManager;
