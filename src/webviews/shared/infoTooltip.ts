export class InfoTooltip {
    private static isSetup = false;
    private static globalTip: HTMLElement | null = null;
    private static activeTrigger: HTMLElement | null = null;

    static inject(toolbarId: string, viewImgUri: string, logoSvgUri: string, viewName: string = 'table view') {
        const toolbarEl = document.getElementById(toolbarId);
        if (toolbarEl && viewImgUri && logoSvgUri) {
            const tooltipDiv = document.createElement('div');
            tooltipDiv.className = 'tooltip edit-mode-hide';
            tooltipDiv.innerHTML = `
                <img src="${viewImgUri}" alt="Change to ${viewName}" style="width: auto; height: 32px; margin-left: auto; margin-top: 2px;" />
                <span class="tooltiptext hidden">
                    <span class="warning">Important:</span> Click the blue table icon <img src="${logoSvgUri}" alt="Table Icon" style="width: 16px; vertical-align: middle; height: 16px;" />
                     to switch to ${viewName} from edit file mode.
                    <span class="instruction">The table icon will only work on edit file mode and is located on the top right corner in the editor toolbar.</span>
                </span>
            `;
            toolbarEl.appendChild(tooltipDiv);

            if (!this.isSetup) {
                this.setupFloatingTooltips();
                this.isSetup = true;
            }
        }
    }

    private static ensureGlobalTip(): HTMLElement {
        if (this.globalTip) return this.globalTip;
        this.globalTip = document.createElement('div');
        this.globalTip.className = 'tooltiptext global-tooltip';
        this.globalTip.style.position = 'fixed';
        this.globalTip.style.zIndex = '2000000';
        this.globalTip.style.pointerEvents = 'none';
        this.globalTip.style.opacity = '0';
        this.globalTip.style.visibility = 'hidden';
        this.globalTip.style.display = 'block';
        this.globalTip.style.transition = 'opacity 0.15s ease, transform 0.15s ease';
        document.body.appendChild(this.globalTip);
        return this.globalTip;
    }

    private static setupFloatingTooltips() {
        const positionTip = (trigger: HTMLElement, content: string | HTMLElement) => {
            const tip = this.ensureGlobalTip();
            
            if (typeof content === 'string') {
                tip.innerHTML = content;
            } else {
                tip.innerHTML = content.innerHTML;
            }

            tip.style.visibility = 'visible';
            tip.style.opacity = '0';
            tip.style.transform = 'translateY(-6px)';

            requestAnimationFrame(() => {
                const r = trigger.getBoundingClientRect();
                const tr = tip.getBoundingClientRect();
                let left = r.left + r.width / 2 - tr.width / 2;
                left = Math.max(8, Math.min(left, window.innerWidth - tr.width - 8));
                let top = r.bottom + 8;
                if (top + tr.height > window.innerHeight - 8) {
                    top = r.top - tr.height - 8;
                }
                tip.style.left = left + 'px';
                tip.style.top = top + 'px';
                tip.style.transform = 'translateY(0)';
                tip.style.opacity = '1';
                this.activeTrigger = trigger;
            });
        };

        const hideTip = () => {
            if (!this.globalTip) return;
            this.globalTip.style.opacity = '0';
            this.globalTip.style.transform = 'translateY(-6px)';
            this.globalTip.style.visibility = 'hidden';
            this.activeTrigger = null;
        };

        document.addEventListener('mouseover', (e) => {
            const t = (e.target as HTMLElement).closest('.tooltip') as HTMLElement;
            if (!t) return;
            
            // If it's a button with a title but no .tooltiptext, we can still show it
            const tipContent = t.querySelector('.tooltiptext') as HTMLElement;
            if (tipContent) {
                positionTip(t, tipContent);
            } else if (t.title) {
                // Store original title to prevent native tooltip
                t.dataset.title = t.title;
                t.title = '';
                positionTip(t, t.dataset.title!);
            } else if (t.dataset.title) {
                positionTip(t, t.dataset.title!);
            }
        });

        document.addEventListener('mouseout', (e) => {
            const t = (e.target as HTMLElement).closest('.tooltip') as HTMLElement;
            if (!t) return;
            
            const rel = e.relatedTarget as HTMLElement;
            if (rel && (t.contains(rel) || (this.globalTip && this.globalTip.contains(rel)))) {
                return;
            }
            hideTip();
        });

        window.addEventListener('resize', () => {
            if (this.activeTrigger) {
                const tipContent = this.activeTrigger.querySelector('.tooltiptext') as HTMLElement;
                if (tipContent) positionTip(this.activeTrigger, tipContent);
                else if (this.activeTrigger.dataset.title) positionTip(this.activeTrigger, this.activeTrigger.dataset.title);
            }
        });
    }
}
