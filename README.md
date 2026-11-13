# PPTX Viewer

Preview PowerPoint files (`.pptx`, `.pptm`, `.ppsx`, `.potx`) right inside VS Code, and edit their speaker notes.

- **Slide thumbnails** sidebar: click to jump to a slide. Hidden slides are dimmed.
- **Speaker notes** panel under the current slide. Toggle it with the **Notes** button or `N`.
- **Edit speaker notes**: type in the notes panel. The tab shows unsaved changes, `Cmd/Ctrl+S` saves, and Undo, Redo, Revert File and hot exit all work. Slides with no notes get a notes page when you add some. If the deck has no notes master, one is created too.
- **Keyboard navigation**: `←` `→` `↑` `↓`, `PageUp`/`PageDown`, `Space`, `Home`, `End`. While you're typing notes, those keys stay in the notes box. Press `Esc` to leave it.
- **Live reload**: when the file changes on disk (for example, after you save it in PowerPoint), the preview refreshes. It doesn't refresh while you have unsaved notes edits.
- Saving changes only the notes parts of the file. Slide content is never rewritten.

## What renders

The renderer is pure JavaScript, with no LibreOffice or Office install needed. It covers:

- Text: fonts, sizes, bold/italic/underline/strike, colors, alignment, bullets and numbering, line spacing, autofit scale
- Placeholder inheritance from slide layouts and masters, and theme colors and fonts
- Shapes: common preset geometries, custom geometry, solid and gradient fills, outlines, dashes and arrowheads
- Images, including cropping, and image fills
- Tables, including table styles, merged cells and borders
- Groups, plus SmartArt through its pre-rendered drawing
- Slide, layout and master backgrounds

Known gaps: charts show as a labelled placeholder. EMF/WMF images aren't supported. Shadows, 3D and other effects are ignored. Text can wrap a little differently from PowerPoint when a slide's fonts aren't installed.

## Build & install

```sh
npm install
npm run compile      # copies JSZip into media/vendor and compiles TypeScript
npm run package      # produces pptx-viewer.vsix
code --install-extension pptx-viewer.vsix
```

Then open any `.pptx` file. If another editor opens it, right-click the file and choose **Open With… → PPTX Viewer**.

## Develop

Open this folder in VS Code and press `F5`, choosing **VS Code Extension Development**. A new window opens with the extension loaded.

```sh
npm test                          # generates a fixture deck, checks the parser and the notes writer
node test/parse.test.js deck.pptx # prints a per-slide summary for any deck
```

## Layout

| Path | Role |
| --- | --- |
| `src/extension.ts` | Custom editor and document model: undo/redo, save, revert, backups, file watching |
| `src/notes.ts` | Writes notes edits into the .pptx, adding notes slides and a notes master when needed |
| `media/pptx-parser.js` | Turns OOXML into a render model (positions in px, resolved styles). Runs in the webview and in Node tests |
| `media/viewer.js` | Draws the render model as HTML/SVG and handles the thumbnails, notes and navigation |
| `media/viewer.css` | Viewer styles, which follow your VS Code theme |
