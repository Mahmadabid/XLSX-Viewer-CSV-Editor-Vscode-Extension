/* eslint-disable @typescript-eslint/no-explicit-any */

export interface SettingDefinition {
    id: string;
    label: string;
    onChange: (value: boolean) => void;
    defaultValue?: boolean;
}

export class SettingsManager {
    private openBtn: HTMLElement | null;
    private panel: HTMLElement | null;
    private cancelBtn: HTMLElement | null;
    private settings: SettingDefinition[];
    private repositionHandlers: any = null;
    private panelOriginalParent: Node | null = null;
    private panelOriginalNext: Node | null = null;
    private onReposition?: () => void;

    constructor(buttonId: string, panelId: string, cancelId: string, settings: SettingDefinition[], onReposition?: () => void) {
        this.openBtn = document.getElementById(buttonId);
        this.panel = document.getElementById(panelId);
        this.cancelBtn = document.getElementById(cancelId);
        this.settings = settings;
        this.onReposition = onReposition;

        this.init();
    }

    static renderPanel(container: HTMLElement, panelId: string, cancelId: string, settings: SettingDefinition[]) {
        const panel = document.createElement('div');
        panel.id = panelId;
        panel.className = 'settings-panel hidden';
        panel.setAttribute('role', 'dialog');
        panel.setAttribute('aria-hidden', 'true');
        
        let html = '<div class="settings-group">';
        settings.forEach(s => {
            html += `<label class="setting-item"><input type="checkbox" id="${s.id}"/> <span>${s.label}</span></label>`;
        });
        html += '</div>';
        html += `<button id="${cancelId}" class="toggle-button" title="Close">Close</button>`;
        
        panel.innerHTML = html;
        container.appendChild(panel);
    }

    private init() {
        if (!this.openBtn || !this.panel) return;

        this.openBtn.addEventListener('click', () => {
            if (this.panel!.classList.contains('hidden')) {
                this.openPanel();
            } else {
                this.closePanel();
            }
        });

        if (this.cancelBtn) {
            this.cancelBtn.addEventListener('click', () => this.closePanel());
        }

        // Wire up settings
        this.settings.forEach(setting => {
            const el = document.getElementById(setting.id) as HTMLInputElement;
            if (el) {
                if (setting.defaultValue !== undefined) {
                    el.checked = setting.defaultValue;
                }
                el.addEventListener('change', () => {
                    setting.onChange(el.checked);
                });
            }
        });

        // Close on click outside
        document.addEventListener('click', (e) => {
            if (!this.panel!.classList.contains('hidden')) {
                if (!(e.target as HTMLElement).closest('.settings-panel') && 
                    !(e.target as HTMLElement).closest('#' + this.openBtn!.id)) {
                    this.closePanel();
                }
            }
        });
    }

    private repositionPanel() {
        const container = document.querySelector('.toolbar');
        if (!container) return;
        const rect = container.getBoundingClientRect();
        
        this.panel!.style.position = 'fixed';
        this.panel!.style.left = Math.max(8, rect.left) + 'px';
        this.panel!.style.top = rect.bottom + 'px';
        this.panel!.style.right = 'auto';
        const maxWidth = Math.min(window.innerWidth - 16, rect.width);
        this.panel!.style.width = Math.max(280, maxWidth) + 'px';
        this.panel!.style.zIndex = '10001';
        
        if (this.onReposition) this.onReposition();
    }

    private openPanel() {
        // Save original parent so we can restore later
        if (!this.panelOriginalParent) {
            this.panelOriginalParent = this.panel!.parentNode;
            this.panelOriginalNext = this.panel!.nextSibling;
        }
        if (this.panel!.parentNode !== document.body) {
            document.body.appendChild(this.panel!);
        }

        this.panel!.classList.remove('hidden');
        this.panel!.classList.add('floating');
        this.panel!.setAttribute('aria-hidden', 'false');
        document.body.classList.add('settings-open');

        const container = document.querySelector('.toolbar');
        if (container) {
            container.classList.add('settings-open');
            // only expand the toolbar vertically when the toolbar is configured to be sticky
            // We check the body class for this state as it's usually managed by the consumer
            if (document.body.classList.contains('sticky-toolbar-enabled')) {
                container.classList.add('expanded-toolbar');
            }
        }

        this.repositionPanel();
        this.repositionHandlers = () => {
            this.repositionPanel();
        };
        window.addEventListener('resize', this.repositionHandlers);
        window.addEventListener('scroll', this.repositionHandlers, true);
    }

    private closePanel() {
        this.panel!.classList.add('hidden');
        this.panel!.classList.remove('floating');
        this.panel!.setAttribute('aria-hidden', 'true');
        document.body.classList.remove('settings-open');

        const container = document.querySelector('.toolbar');
        if (container) {
            container.classList.remove('settings-open');
            // We can't easily know if sticky toolbar is disabled here without querying the checkbox
            // But usually the consumer handles the class removal on the container if needed
            // For safety, we remove expanded-toolbar if the body doesn't say it's enabled
            if (!document.body.classList.contains('sticky-toolbar-enabled')) {
                container.classList.remove('expanded-toolbar');
            }
        }

        this.panel!.style.position = '';
        this.panel!.style.left = '';
        this.panel!.style.top = '';
        this.panel!.style.width = '';
        this.panel!.style.right = '';
        this.panel!.style.zIndex = '';

        // Restore original parent/position
        if (this.panelOriginalParent && this.panelOriginalParent !== this.panel!.parentNode) {
            try {
                this.panelOriginalParent.insertBefore(this.panel!, this.panelOriginalNext);
            } catch (e) {
                this.panelOriginalParent.appendChild(this.panel!);
            }
        }

        if (this.repositionHandlers) {
            window.removeEventListener('resize', this.repositionHandlers);
            window.removeEventListener('scroll', this.repositionHandlers, true);
            this.repositionHandlers = null;
        }
    }
}
