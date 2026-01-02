/* eslint-disable @typescript-eslint/no-explicit-any */

declare function acquireVsCodeApi(): any;

export const vscode = (function () {
    if (typeof acquireVsCodeApi === 'function') {
        return acquireVsCodeApi();
    }
    return { postMessage: () => { } };
})();

export const VirtualScrollConfig = {
    ROW_HEIGHT: 28,
    BUFFER_ROWS: 20,
    CHUNK_SIZE: 100
};

export function debounce(func: Function, wait: number) {
    let timeout: any;
    return function (this: any, ...args: any[]) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), wait);
    };
}
