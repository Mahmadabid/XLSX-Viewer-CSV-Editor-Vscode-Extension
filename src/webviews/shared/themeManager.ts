/* eslint-disable @typescript-eslint/no-explicit-any */

export interface ThemeManagerOptions {
    persistKey?: string;
    onBeforeCycle?: () => boolean | void;
}

export class ThemeManager {
    private button: HTMLElement | null;
    private vscodeApi: any;
    private options: ThemeManagerOptions;
    private themes: string[];

    constructor(buttonId: string, options: ThemeManagerOptions = {}, vscodeApi: any = null) {
        this.button = document.getElementById(buttonId);
        this.vscodeApi = vscodeApi;
        this.options = {
            persistKey: 'last_used_theme',
            ...options
        };

        this.themes = ['light', 'dark', 'vscode'];
        this.init();
    }

    private init() {
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

    private getStoredTheme(): string {
        // 1. Try to get from vscode state if available (per-file persistence for current session)
        if (this.vscodeApi && typeof this.vscodeApi.getState === 'function') {
            const state = this.vscodeApi.getState();
            if (state && state.theme) return state.theme;
        }

        // 2. Fallback to global localStorage (last used theme)
        try {
            const lastUsed = localStorage.getItem(this.options.persistKey!);
            if (lastUsed && this.themes.includes(lastUsed)) return lastUsed;
        } catch (e) { /* ignore */ }

        return 'vscode';
    }

    private setStoredTheme(theme: string) {
        // Always update vscode state for per-file consistency if available
        if (this.vscodeApi && typeof this.vscodeApi.getState === 'function') {
            const state = this.vscodeApi.getState() || {};
            state.theme = theme;
            this.vscodeApi.setState(state);
        }

        // Update global localStorage (last used theme)
        try {
            localStorage.setItem(this.options.persistKey!, theme);
        } catch (e) { /* ignore */ }
    }

    private applyTheme(theme: string, save = true) {
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

    private cycleTheme() {
        if (typeof this.options.onBeforeCycle === 'function') {
            if (this.options.onBeforeCycle() === false) return;
        }
        const current = this.getStoredTheme();
        const currentIndex = this.themes.indexOf(current);
        const nextIndex = (currentIndex + 1) % this.themes.length;
        const nextTheme = this.themes[nextIndex];
        this.applyTheme(nextTheme);
    }

    private updateIcons(currentTheme: string) {
        const nextIndex = (this.themes.indexOf(currentTheme) + 1) % this.themes.length;
        const nextTheme = this.themes[nextIndex];

        const icons = {
            light: document.getElementById('lightIcon'),
            dark: document.getElementById('darkIcon'),
            vscode: document.getElementById('vscodeIcon')
        };

        // Hide all first
        Object.values(icons).forEach(icon => {
            if (icon) icon.style.display = 'none';
        });

        // Show the icon representing the *next* state (standard toggle behavior)
        // or show the current state? 
        // The original code showed the icon for the *next* theme to indicate what clicking does.
        // Let's check the original code logic.
        // Original: const nextTheme = this.themes[nextIndex]; ... if (icons[nextTheme]) icons[nextTheme].style.display = 'block';
        
        if (icons[nextTheme as keyof typeof icons]) {
            (icons[nextTheme as keyof typeof icons] as HTMLElement).style.display = 'block';
        }
        
        // Update title/tooltip
        const currentName = currentTheme.charAt(0).toUpperCase() + currentTheme.slice(1);
        const nextName = nextTheme.charAt(0).toUpperCase() + nextTheme.slice(1);
        const tooltipText = `Switch to <b>${nextName} theme</b> from ${currentName} theme`;
        
        const wrapper = this.button!.closest('.tooltip');
        if (wrapper) {
            const tip = wrapper.querySelector('.tooltiptext');
            if (tip) tip.innerHTML = tooltipText;
        } else {
            this.button!.title = tooltipText;
        }
    }

    private handleVsCodeThemeChange(kind: number) {
        // kind: 1 = Light, 2 = Dark, 3 = High Contrast
        // Only relevant if we are in 'vscode' mode
        if (this.getStoredTheme() === 'vscode') {
            // We don't need to do anything because CSS handles .vscode-theme
            // But we might want to update internal state if we were tracking it
        }
    }
}
