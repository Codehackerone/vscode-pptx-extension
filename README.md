# PPTX Viewer

Preview PowerPoint files (`.pptx`, `.pptm`, `.ppsx`, `.potx`) right inside VS Code.

- **Slide thumbnails** sidebar: click to jump to a slide. Hidden slides are dimmed.
- **Speaker notes** panel under the current slide. Toggle it with the **Notes** button or `N`.
- **Keyboard navigation**: `←` `→` `↑` `↓`, `PageUp`/`PageDown`, `Space`, `Home`, `End`.
- **Live reload**: when the file changes on disk (for example, after you save it in PowerPoint), the preview refreshes.
- Read-only. The file is never modified.

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
npm test                          # generates a fixture deck and checks the parser output
node test/parse.test.js deck.pptx # prints a per-slide summary for any deck
```

## Layout

| Path | Role |
| --- | --- |
| `src/extension.ts` | Registers the custom editor, sends file bytes to the webview, and watches for changes |
| `media/pptx-parser.js` | Turns OOXML into a render model (positions in px, resolved styles). Runs in the webview and in Node tests |
| `media/viewer.js` | Draws the render model as HTML/SVG and handles the thumbnails, notes and navigation |
| `media/viewer.css` | Viewer styles, which follow your VS Code theme |
