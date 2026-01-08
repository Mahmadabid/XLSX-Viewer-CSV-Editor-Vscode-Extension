import * as vscode from 'vscode';
import * as fs from 'fs';

export class MDEditorProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private readonly context: vscode.ExtensionContext) { }

    async openCustomDocument(
        uri: vscode.Uri,
        openContext: vscode.CustomDocumentOpenContext,
        token: vscode.CancellationToken
    ): Promise<vscode.CustomDocument> {
        return { uri, dispose: () => { } };
    }

    async resolveCustomEditor(
        document: vscode.CustomDocument,
        webviewPanel: vscode.WebviewPanel,
        token: vscode.CancellationToken
    ): Promise<void> {
        try {
            const filePath = document.uri.fsPath;

            // Set up webview
            webviewPanel.webview.options = {
                enableScripts: true,
                localResourceRoots: [
                    vscode.Uri.joinPath(this.context.extensionUri, 'resources'),
                    vscode.Uri.joinPath(this.context.extensionUri, 'dist')
                ]
            };
            webviewPanel.webview.html = this.getWebviewContent(webviewPanel);

            // Handle messages from webview
            webviewPanel.webview.onDidReceiveMessage(async message => {
                switch (message.command) {
                    case 'webviewReady':
                        try {
                            // Read the markdown file
                            const content = await fs.promises.readFile(filePath, 'utf-8');

                            // Send content to webview
                            webviewPanel.webview.postMessage({
                                command: 'initMarkdown',
                                content,
                                fileName: vscode.workspace.asRelativePath(document.uri)
                            });

                            // Calculate if MD is enabled as default
                            const globalCfg = vscode.workspace.getConfiguration('workbench');
                            const associations: any = globalCfg.get('editorAssociations');
                            let isMdEnabled = false;
                            
                            if (associations) {
                                if (Array.isArray(associations)) {
                                    isMdEnabled = associations.some(a => a.viewType === 'xlsxViewer.md' && (a.filenamePattern === '*.md' || a.filenamePattern === '**/*.md'));
                                } else {
                                    isMdEnabled = associations["*.md"] === 'xlsxViewer.md' || associations["**/*.md"] === 'xlsxViewer.md';
                                }
                            }

                            // Send settings
                            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                            const settings = {
                                stickyToolbar: cfg.get('md.stickyToolbar', true),
                                wordWrap: cfg.get('md.wordWrap', true),
                                syncScroll: cfg.get('md.syncScroll', true),
                                previewPosition: cfg.get('md.previewPosition', 'right'),
                                isMdEnabled: isMdEnabled
                            };
                            webviewPanel.webview.postMessage({ command: 'initSettings', settings });

                            // Send theme
                            webviewPanel.webview.postMessage({
                                type: 'setTheme',
                                kind: vscode.window.activeColorTheme.kind
                            });
                        } catch (err) {
                            vscode.window.showErrorMessage(`Error reading Markdown file: ${err}`);
                        }
                        break;

                    case 'updateSettings':
                        try {
                            const s = message.settings || {};
                            const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                            await cfg.update('md.stickyToolbar', !!s.stickyToolbar, vscode.ConfigurationTarget.Global);
                            await cfg.update('md.wordWrap', !!s.wordWrap, vscode.ConfigurationTarget.Global);
                            await cfg.update('md.syncScroll', !!s.syncScroll, vscode.ConfigurationTarget.Global);
                            await cfg.update('md.previewPosition', s.previewPosition || 'right', vscode.ConfigurationTarget.Global);
                        } catch (err) {
                            console.error('Failed to persist settings:', err);
                        }
                        break;

                    case 'toggleView':
                        if (!message.isPreviewView) {
                            await vscode.commands.executeCommand('vscode.openWith', document.uri, 'default');
                            webviewPanel.dispose();
                        }
                        break;

                    case 'saveMarkdown':
                        try {
                            const text = typeof message.text === 'string' ? message.text : '';
                            await vscode.workspace.fs.writeFile(document.uri, Buffer.from(text, 'utf8'));
                            webviewPanel.webview.postMessage({ command: 'saveResult', ok: true });
                        } catch (err) {
                            webviewPanel.webview.postMessage({ command: 'saveResult', ok: false, error: String(err) });
                        }
                        break;

                    case 'openExternal':
                        try {
                            const url = typeof message.url === 'string' ? message.url : '';
                            if (url) {
                                await vscode.env.openExternal(vscode.Uri.parse(url));
                            }
                        } catch {
                            // ignore
                        }
                        break;

                    case 'disableMdEditor':
                        try {
                            const result = await vscode.window.showWarningMessage(
                                "Are you sure you want to disable XLSX Viewer for all Markdown files? You will be prompted to select a new default editor.",
                                "Yes, Disable",
                                "Cancel"
                            );

                            if (result === "Yes, Disable") {
                                // 1. First remove our association so it's not the default anymore
                                await vscode.commands.executeCommand('xlsx-viewer.toggleMdAssociation', false);
                                
                                // 2. Trigger the "Reopen With..." picker which allows selecting a new default
                                // We use this command as it is more widely available than changeDefaultViewType
                                await vscode.commands.executeCommand('workbench.action.reopenWithEditor');
                            }
                        } catch (err) {
                            vscode.window.showErrorMessage(`Error disabling MD editor: ${err}`);
                        }
                        break;
                    
                    case 'enableMdEditor':
                        try {
                            await vscode.commands.executeCommand('xlsx-viewer.toggleMdAssociation', true);
                            
                             // Send updated settings
                             const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                             const settings = {
                                 stickyToolbar: cfg.get('md.stickyToolbar', true),
                                 wordWrap: cfg.get('md.wordWrap', true),
                                 syncScroll: cfg.get('md.syncScroll', true),
                                 previewPosition: cfg.get('md.previewPosition', 'right'),
                                 isMdEnabled: true
                             };
                             webviewPanel.webview.postMessage({ command: 'initSettings', settings });

                        } catch (err) {
                            vscode.window.showErrorMessage(`Error enabling MD editor: ${err}`);
                        }
                        break;
                    
                    case 'toggleMdAssociation':
                        await vscode.commands.executeCommand('xlsx-viewer.toggleMdAssociation', !!message.enable);
                        break;
                }
            });

            // Forward settings changes
            const configChangeDisposable = vscode.workspace.onDidChangeConfiguration(e => {
                if (e.affectsConfiguration('xlsxViewer.md') || e.affectsConfiguration('xlsxViewer') || e.affectsConfiguration('workbench.editorAssociations')) {
                    const cfg = vscode.workspace.getConfiguration('xlsxViewer');
                    const globalCfg = vscode.workspace.getConfiguration('workbench');
                    const associations: any = globalCfg.get('editorAssociations');
                    let isMdEnabled = false;
                    
                    if (associations) {
                        if (Array.isArray(associations)) {
                            isMdEnabled = associations.some(a => a.viewType === 'xlsxViewer.md' && (a.filenamePattern === '*.md' || a.filenamePattern === '**/*.md'));
                        } else {
                            isMdEnabled = associations["*.md"] === 'xlsxViewer.md' || associations["**/*.md"] === 'xlsxViewer.md';
                        }
                    }

                    const settings = {
                        stickyToolbar: cfg.get('md.stickyToolbar', true),
                        wordWrap: cfg.get('md.wordWrap', true),
                        syncScroll: cfg.get('md.syncScroll', true),
                        previewPosition: cfg.get('md.previewPosition', 'right'),
                        isMdEnabled: isMdEnabled
                    };
                    try {
                        webviewPanel.webview.postMessage({ command: 'settingsUpdated', settings });
                    } catch { }
                }
            });

            // Theme change listener
            const themeChangeDisposable = vscode.window.onDidChangeActiveColorTheme(() => {
                try {
                    webviewPanel.webview.postMessage({
                        type: 'setTheme',
                        kind: vscode.window.activeColorTheme.kind
                    });
                } catch { }
            });

            webviewPanel.onDidDispose(() => {
                configChangeDisposable.dispose();
                themeChangeDisposable.dispose();
            });

        } catch (error) {
            vscode.window.showErrorMessage(`Error reading Markdown file: ${error}`);
        }
    }

    private getWebviewContent(webviewPanel: vscode.WebviewPanel): string {
        const webview = webviewPanel.webview;
        const imgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'md', 'view.png'));
        const svgUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'md', 'logo.svg'));
        const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'md', 'mdWebview.js'));
        const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'md', 'mdWebview.css'));
        const themeUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'shared', 'theme.css'));
        const highlightUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'resources', 'md', 'highlight.css'));
        const cspSource = webview.cspSource;

        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${cspSource} https: data:; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource} 'unsafe-inline';">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Markdown Viewer</title>
            <link href="${themeUri}" rel="stylesheet" />
            <link href="${styleUri}" rel="stylesheet" />
            <link href="${highlightUri}" rel="stylesheet" />
            <script>
                window.viewImgUri = "${imgUri}";
                window.logoSvgUri = "${svgUri}";
            </script>
        </head>
        <body>
            <div class="header-background"></div>
            <div class="toolbar" id="toolbar"></div>

            <div id="content">
                <div id="loadingIndicator" class="loading-indicator">Loading Markdown...</div>
                <div class="markdown-container" id="markdownContainer">
                    <div class="editor-wrapper">
                        <textarea id="markdownEditor" class="markdown-editor" spellcheck="false"></textarea>
                    </div>
                    <div id="markdownPreview" class="markdown-preview"></div>
                </div>
            </div>

            <div class="status-info" id="statusInfo"></div>

            <noscript>
                <div style="padding: 8px; margin-top: 10px; background: #fff3cd; border: 1px solid #ffeeba;">
                    JavaScript is disabled in this webview, so the Markdown preview cannot load.
                </div>
            </noscript>
            <script src="${scriptUri}"></script>
        </body>
        </html>`;
    }
}