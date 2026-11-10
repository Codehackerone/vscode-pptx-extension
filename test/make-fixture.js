// Generates test/fixtures/sample.pptx covering text, bullets, shapes, an image, a table and notes.
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// 2x2 red PNG
const PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg==';

async function main() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';

  const s1 = pptx.addSlide();
  s1.background = { color: '1F2A44' };
  s1.addText('PPTX Viewer', { x: 0.5, y: 2.2, w: 12.3, h: 1.4, fontSize: 54, bold: true, color: 'FFFFFF', align: 'center' });
  s1.addText('Fixture deck', { x: 0.5, y: 3.6, w: 12.3, h: 0.8, fontSize: 24, color: 'C8D3F5', align: 'center' });
  s1.addNotes('Welcome everyone.\nThis is the title slide.');

  const s2 = pptx.addSlide();
  s2.addText('Bullets & shapes', { x: 0.5, y: 0.3, w: 12, h: 1, fontSize: 36, bold: true, color: '1F2A44' });
  s2.addText(
    [
      { text: 'First point', options: { bullet: true } },
      { text: 'Second point, in italics', options: { bullet: true, italic: true } },
      { text: 'Nested point', options: { bullet: { indent: 20 }, indentLevel: 1 } },
      { text: 'Numbered one', options: { bullet: { type: 'number' } } },
      { text: 'Numbered two', options: { bullet: { type: 'number' } } },
    ],
    { x: 0.5, y: 1.4, w: 6, h: 3.5, fontSize: 20, color: '333333' },
  );
  s2.addShape(pptx.ShapeType.roundRect, { x: 7, y: 1.5, w: 2.5, h: 1.5, fill: { color: '4472C4' }, rectRadius: 0.2 });
  s2.addShape(pptx.ShapeType.ellipse, { x: 10, y: 1.5, w: 2, h: 2, fill: { color: 'ED7D31' }, line: { color: '000000', width: 2 } });
  s2.addShape(pptx.ShapeType.rightArrow, { x: 7, y: 3.6, w: 3, h: 1, fill: { color: '70AD47' } });
  s2.addShape(pptx.ShapeType.line, { x: 7, y: 5.2, w: 5, h: 0, line: { color: 'FF0000', width: 3, dashType: 'dash', endArrowType: 'triangle' } });
  s2.addImage({ data: PNG, x: 0.5, y: 5.2, w: 1.5, h: 1.5 });
  s2.addNotes('Talk about bullets.');

  const s3 = pptx.addSlide();
  s3.addText('A table', { x: 0.5, y: 0.3, w: 12, h: 1, fontSize: 36, bold: true });
  s3.addTable(
    [
      [{ text: 'Name', options: { bold: true, fill: { color: '4472C4' }, color: 'FFFFFF' } }, { text: 'Value', options: { bold: true, fill: { color: '4472C4' }, color: 'FFFFFF' } }],
      ['Alpha', '1'],
      ['Beta', '2'],
    ],
    { x: 0.5, y: 1.5, w: 8, colW: [5, 3], fontSize: 18, border: { type: 'solid', pt: 1, color: '999999' } },
  );

  const outDir = path.join(__dirname, 'fixtures');
  fs.mkdirSync(outDir, { recursive: true });
  const buf = await pptx.write({ outputType: 'nodebuffer' });
  fs.writeFileSync(path.join(outDir, 'sample.pptx'), buf);
  console.log('wrote test/fixtures/sample.pptx');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
