import { vscode } from './common';

export class VirtualLoader<T> {
    private pendingRequests = new Map<string, { resolve: (value: T[] | PromiseLike<T[]>) => void, start: number, end: number }>();

    constructor(private command: string) {}

    requestRows(start: number, end: number, timeout = 10000, extraData?: any): Promise<T[]> {
        return new Promise((resolve) => {
            const id = `${start}-${end}-${Date.now()}`;
            this.pendingRequests.set(id, { resolve, start, end });

            vscode.postMessage({
                command: this.command,
                start,
                end,
                requestId: id,
                ...extraData
            });

            setTimeout(() => {
                if (this.pendingRequests.has(id)) {
                    this.pendingRequests.delete(id);
                    resolve([]);
                }
            }, timeout);
        });
    }

    resolveRequest(requestId: string, data: T[]) {
        const req = this.pendingRequests.get(requestId);
        if (req) {
            req.resolve(data);
            this.pendingRequests.delete(requestId);
        }
    }
    
    clear() {
        this.pendingRequests.clear();
    }
}
