
export class Utils {
    static $(id: string): HTMLElement | null {
        return document.getElementById(id);
    }

    static normalizeCellText(text: string | null | undefined): string {
        if (!text) return '';
        return String(text).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trimEnd();
    }

    static escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    static showToast(message: string) {
        let toast = document.getElementById('saveToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'saveToast';
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
        
        const textSpan = toast.querySelector('.toast-text');
        if (textSpan) textSpan.textContent = message;
        
        toast.classList.add('show');
        setTimeout(() => {
            toast?.classList.remove('show');
        }, 3000);
    }

    static async writeToClipboardAsync(text: string): Promise<boolean> {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            try {
                await navigator.clipboard.writeText(text);
                return true;
            } catch (err) {
                console.warn('Clipboard API failed, trying fallback');
            }
        }
        
        // Fallback
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.style.position = 'fixed';
        textarea.style.left = '-9999px';
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
            document.body.removeChild(textarea);
            return true;
        } catch (err) {
            document.body.removeChild(textarea);
            return false;
        }
    }
}
