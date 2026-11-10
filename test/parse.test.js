// Parses .pptx files with the webview parser under Node and checks the render model.
// Usage: node test/parse.test.js [file.pptx ...]   (defaults to the generated fixture)
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const JSZip = require('jszip');
const { DOMParser } = require('@xmldom/xmldom');
const { parsePptx } = require('../media/pptx-parser.js');

async function parseFile(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  return parsePptx(zip, { DOMParser, makeImageUrl: (bytes, mime, p) => `blob:${p}` });
}

const allText = (slide) =>
  slide.elements
    .flatMap((e) => (e.text ? e.text.paragraphs : e.rows ? e.rows.flatMap((r) => r.cells.flatMap((c) => c.paragraphs || [])) : []))
    .flatMap((p) => p.runs.map((r) => r.text || ''))
    .join(' ');

async function checkFixture(file) {
  const m = await parseFile(file);
  assert.strictEqual(m.slides.length, 3, 'slide count');
  assert.ok(Math.abs(m.width - 1280) < 1 && Math.abs(m.height - 720) < 1, `16:9 size, got ${m.width}x${m.height}`);

  const [s1, s2, s3] = m.slides;
  assert.deepStrictEqual(s1.background, { type: 'solid', color: 'rgb(31, 42, 68)' });
  assert.match(allText(s1), /PPTX Viewer/);
  assert.strictEqual(s1.notes, 'Welcome everyone.\nThis is the title slide.');
  const title = s1.elements.find((e) => e.text && /PPTX Viewer/.test(allText({ elements: [e] })));
  const run = title.text.paragraphs[0].runs[0];
  assert.ok(run.bold && Math.abs(run.size - 72) < 0.01 && run.color === 'rgb(255, 255, 255)', 'title run style');

  const bullets = s2.elements.find((e) => e.text && /First point/.test(allText({ elements: [e] })));
  const paras = bullets.text.paragraphs;
  assert.strictEqual(paras[0].bullet.text, '•');
  assert.ok(paras[1].runs[0].italic, 'italic run');
  assert.strictEqual(paras[3].bullet.text, '1.');
  assert.strictEqual(paras[4].bullet.text, '2.');
  const geoms = s2.elements.filter((e) => e.type === 'shape').map((e) => e.geom);
  for (const g of ['roundRect', 'ellipse', 'rightArrow', 'line']) assert.ok(geoms.includes(g), `has ${g}`);
  const line = s2.elements.find((e) => e.geom === 'line');
  assert.ok(line.line && line.line.dash === 'dash' && line.line.tail, 'dashed arrow line');
  const img = s2.elements.find((e) => e.type === 'image');
  assert.ok(img && img.url && img.url.startsWith('blob:ppt/media/'), 'image url');
  assert.strictEqual(s2.notes, 'Talk about bullets.');

  const table = s3.elements.find((e) => e.type === 'table');
  assert.ok(table, 'table present');
  assert.strictEqual(table.rows.length, 3);
  assert.strictEqual(table.cols.length, 2);
  assert.strictEqual(table.rows[0].cells[0].fill.color, 'rgb(68, 114, 196)');
  assert.strictEqual(s3.notes, '');
  console.log(`ok  ${path.relative(process.cwd(), file)} (fixture assertions passed)`);
}

function summarize(file, m) {
  console.log(`\n${file}: ${m.slides.length} slides, ${Math.round(m.width)}x${Math.round(m.height)}px`);
  for (const s of m.slides) {
    const counts = {};
    for (const e of s.elements) counts[e.type] = (counts[e.type] || 0) + 1;
    const text = allText(s).replace(/\s+/g, ' ').trim().slice(0, 60);
    console.log(`  #${s.number}${s.hidden ? ' (hidden)' : ''} ${JSON.stringify(counts)} ${s.error ? 'ERROR ' + s.error : ''} "${text}"${s.notes ? ' [notes]' : ''}`);
  }
}

(async () => {
  const args = process.argv.slice(2);
  if (!args.length) {
    await checkFixture(path.join(__dirname, 'fixtures', 'sample.pptx'));
    return;
  }
  for (const f of args) summarize(f, await parseFile(f));
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
