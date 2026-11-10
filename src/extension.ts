import * as vscode from 'vscode';

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(PptxEditorProvider.register(context));
}

export function deactivate() {}

class PptxEditorProvider implements vscode.CustomReadonlyEditorProvider {
  static readonly viewType = 'pptxViewer.preview';

  static register(context: vscode.ExtensionContext): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      PptxEditorProvider.viewType,
      new PptxEditorProvider(context),
      {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: true,
      },
    );
  }

  constructor(private readonly context: vscode.ExtensionContext) {}

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose: () => {} };
  }

  resolveCustomEditor(document: vscode.CustomDocument, panel: vscode.WebviewPanel): void {
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, localResourceRoots: [mediaRoot] };
    panel.webview.html = this.getHtml(panel.webview, mediaRoot);

    const fileName = document.uri.path.split('/').pop() ?? 'presentation.pptx';
    const sendFile = async () => {
      try {
        const bytes = await vscode.workspace.fs.readFile(document.uri);
        // Send base64: typed arrays/Buffers don't survive postMessage serialization reliably.
        const data = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64');
        await panel.webview.postMessage({ type: 'load', data, fileName });
      } catch (err) {
        await panel.webview.postMessage({ type: 'error', message: String(err) });
      }
    };

    // Reload the preview when the file changes on disk (e.g. re-saved from PowerPoint).
    let reloadTimer: NodeJS.Timeout | undefined;
    const scheduleReload = () => {
      clearTimeout(reloadTimer);
      reloadTimer = setTimeout(sendFile, 300);
    };
    const dir = vscode.Uri.joinPath(document.uri, '..');
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(dir, fileName));
    watcher.onDidChange(scheduleReload);
    watcher.onDidCreate(scheduleReload);

    const messages = panel.webview.onDidReceiveMessage((msg) => {
      if (msg?.type === 'ready') {
        void sendFile();
      }
    });

    panel.onDidDispose(() => {
      clearTimeout(reloadTimer);
      watcher.dispose();
      messages.dispose();
    });
  }

  private getHtml(webview: vscode.Webview, mediaRoot: vscode.Uri): string {
    const uri = (...p: string[]) => webview.asWebviewUri(vscode.Uri.joinPath(mediaRoot, ...p));
    const nonce = makeNonce();
    const csp = [
      `default-src 'none'`,
      `img-src ${webview.cspSource} blob: data:`,
      `style-src ${webview.cspSource} 'unsafe-inline'`,
      `font-src ${webview.cspSource}`,
      `script-src 'nonce-${nonce}'`,
    ].join('; ');

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="${csp}">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="${uri('viewer.css')}">
  <title>PPTX Viewer</title>
</head>
<body>
  <div id="app">
    <aside id="sidebar" aria-label="Slides"></aside>
    <main id="main">
      <div id="toolbar">
        <button id="prev" title="Previous slide (←)">&#x2039;</button>
        <span id="counter">–</span>
        <button id="next" title="Next slide (→)">&#x203A;</button>
        <span class="spacer"></span>
        <button id="toggle-notes" title="Toggle speaker notes (N)">Notes</button>
      </div>
      <div id="stage"><div id="status">Loading…</div></div>
      <section id="notes" aria-label="Speaker notes"></section>
    </main>
  </div>
  <script nonce="${nonce}" src="${uri('vendor', 'jszip.min.js')}"></script>
  <script nonce="${nonce}" src="${uri('pptx-parser.js')}"></script>
  <script nonce="${nonce}" src="${uri('viewer.js')}"></script>
</body>
</html>`;
  }
}

function makeNonce(): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let out = '';
  for (let i = 0; i < 32; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}
