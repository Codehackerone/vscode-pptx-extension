import * as vscode from 'vscode';
import { applyNotesEdits } from './notes';

export function activate(context: vscode.ExtensionContext) {
  context.subscriptions.push(PptxEditorProvider.register(context));
}

export function deactivate() {}

/**
 * A .pptx file plus unsaved speaker-notes changes. `bytes` is the last saved/loaded
 * content; `notes` holds the full notes text for every slide edited since then.
 */
class PptxDocument implements vscode.CustomDocument {
  readonly notes = new Map<number, string>();
  lastSaved = 0;
  /** Bumped whenever `bytes` is replaced from disk, so webviews know to re-render. */
  version = 0;

  private readonly _onDidChange = new vscode.EventEmitter<void>();
  /** Fires when the notes change (edit, undo, redo, revert) so webviews can refresh. */
  readonly onDidChange = this._onDidChange.event;

  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  readonly onDidDispose = this._onDidDispose.event;

  constructor(readonly uri: vscode.Uri, public bytes: Uint8Array) {}

  setNotes(index: number, text: string) {
    this.notes.set(index, text);
    this._onDidChange.fire();
  }

  /** Bytes of the file with the unsaved notes applied. */
  async render(): Promise<Uint8Array> {
    return applyNotesEdits(this.bytes, this.notes);
  }

  async reload(): Promise<void> {
    this.bytes = await vscode.workspace.fs.readFile(this.uri);
    this.notes.clear();
    this.version++;
    this._onDidChange.fire();
  }

  dispose() {
    this._onDidDispose.fire();
    this._onDidChange.dispose();
    this._onDidDispose.dispose();
  }
}

class PptxEditorProvider implements vscode.CustomEditorProvider<PptxDocument> {
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

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<PptxDocument>>();
  readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): Promise<PptxDocument> {
    // Restoring from a hot-exit backup: the backup is a complete .pptx with the edits applied.
    const source = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    return new PptxDocument(uri, await vscode.workspace.fs.readFile(source));
  }

  resolveCustomEditor(document: PptxDocument, panel: vscode.WebviewPanel): void {
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, localResourceRoots: [mediaRoot] };
    panel.webview.html = this.getHtml(panel.webview, mediaRoot);

    const fileName = document.uri.path.split('/').pop() ?? 'presentation.pptx';
    const post = (msg: unknown) => void panel.webview.postMessage(msg);
    const notesObject = () => Object.fromEntries(document.notes);

    const sendFile = () => {
      // Send base64: typed arrays/Buffers don't survive postMessage serialization reliably.
      const b = document.bytes;
      const data = Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('base64');
      post({ type: 'load', data, fileName, notes: notesObject() });
    };

    // Reload when the file changes on disk (e.g. re-saved from PowerPoint), unless the
    // change is our own save or there are unsaved edits we'd throw away.
    let reloadTimer: NodeJS.Timeout | undefined;
    const onDiskChange = () => {
      clearTimeout(reloadTimer);
      reloadTimer = setTimeout(async () => {
        if (Date.now() - document.lastSaved < 2000 || document.notes.size > 0) return;
        try {
          await document.reload();
        } catch (err) {
          post({ type: 'error', message: String(err) });
        }
      }, 300);
    };
    const dir = vscode.Uri.joinPath(document.uri, '..');
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(dir, fileName));
    watcher.onDidChange(onDiskChange);
    watcher.onDidCreate(onDiskChange);

    let loadedVersion = document.version;
    const changes = document.onDidChange(() => {
      if (document.version !== loadedVersion) {
        loadedVersion = document.version;
        sendFile();
      } else {
        post({ type: 'notes', notes: notesObject() });
      }
    });

    const messages = panel.webview.onDidReceiveMessage((msg) => {
      switch (msg?.type) {
        case 'ready':
          sendFile();
          break;
        case 'editNotes':
          this.editNotes(document, Number(msg.index), String(msg.before ?? ''), String(msg.after ?? ''));
          break;
      }
    });

    panel.onDidDispose(() => {
      clearTimeout(reloadTimer);
      watcher.dispose();
      changes.dispose();
      messages.dispose();
    });
  }

  private editNotes(document: PptxDocument, index: number, before: string, after: string) {
    if (!Number.isInteger(index) || index < 0 || before === after) return;
    document.setNotes(index, after);
    this._onDidChangeCustomDocument.fire({
      document,
      label: 'Edit Speaker Notes',
      undo: () => document.setNotes(index, before),
      redo: () => document.setNotes(index, after),
    });
  }

  async saveCustomDocument(document: PptxDocument): Promise<void> {
    await this.saveCustomDocumentAs(document, document.uri);
  }

  async saveCustomDocumentAs(document: PptxDocument, destination: vscode.Uri): Promise<void> {
    const saved = new Map(document.notes);
    const bytes = await applyNotesEdits(document.bytes, saved);
    const inPlace = destination.toString() === document.uri.toString();
    if (inPlace) document.lastSaved = Date.now(); // Ignore the watcher event our write causes.
    await vscode.workspace.fs.writeFile(destination, bytes);
    if (inPlace) {
      // The file now holds these edits. Drop them from the pending map unless they changed
      // during the save. Undo/redo set absolute text, so they still work afterwards.
      // The webview already shows this content, so it isn't reloaded.
      document.bytes = bytes;
      document.lastSaved = Date.now();
      for (const [i, text] of saved) if (document.notes.get(i) === text) document.notes.delete(i);
    }
  }

  async revertCustomDocument(document: PptxDocument): Promise<void> {
    await document.reload();
  }

  async backupCustomDocument(
    document: PptxDocument,
    context: vscode.CustomDocumentBackupContext,
  ): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, await document.render());
    return {
      id: context.destination.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(context.destination);
        } catch {
          // Already gone.
        }
      },
    };
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
      <section id="notes" aria-label="Speaker notes">
        <label class="notes-title" for="notes-input">Speaker notes</label>
        <textarea id="notes-input" spellcheck="true" placeholder="Click to add notes"></textarea>
      </section>
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
