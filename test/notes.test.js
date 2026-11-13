// Checks that speaker-notes edits are written into the .pptx and read back correctly.
// Requires `npm run compile` (uses out/notes.js) and the generated fixture.
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const JSZip = require('jszip');
const { DOMParser } = require('@xmldom/xmldom');
const { parsePptx } = require('../media/pptx-parser.js');
const { applyNotesEdits } = require('../out/notes.js');

const fixture = fs.readFileSync(path.join(__dirname, 'fixtures', 'sample.pptx'));

async function notesOf(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const model = await parsePptx(zip, { DOMParser, makeImageUrl: () => null });
  return model.slides.map((s) => s.notes);
}

/** Every XML part must parse without errors or warnings. */
async function assertWellFormed(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  for (const name of Object.keys(zip.files)) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    const problems = [];
    const parser = new DOMParser({ onError: (level, msg) => problems.push(`${level}: ${msg}`) });
    parser.parseFromString(await zip.file(name).async('string'), 'application/xml');
    assert.deepStrictEqual(problems, [], `${name} is not well-formed`);
  }
}

/** Removes every notes slide and the notes master, like a deck that never had notes. */
async function stripNotes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/notes(Slides|Masters)\//.test(name)) zip.remove(name);
  }
  const dropRel = (xml, type) => xml.replace(new RegExp(`<Relationship [^>]*Type="[^"]*/${type}"[^>]*/>`, 'g'), '');
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/(slides\/_rels\/.*|_rels\/presentation\.xml)\.rels$/.test(name)) {
      let xml = await zip.file(name).async('string');
      xml = dropRel(dropRel(xml, 'notesSlide'), 'notesMaster');
      zip.file(name, xml);
    }
  }
  let pres = await zip.file('ppt/presentation.xml').async('string');
  pres = pres.replace(/<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/, '');
  zip.file('ppt/presentation.xml', pres);
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace(/<Override PartName="\/ppt\/notes(Slides|Masters)\/[^"]*"[^>]*\/>/g, '');
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array' });
}

(async () => {
  // 1. Edit existing notes, including multi-line text and XML special characters.
  const tricky = 'Line one <with> & "quotes"\n\nLine three';
  const edited = await applyNotesEdits(fixture, new Map([[0, tricky], [2, 'Notes for the table slide']]));
  await assertWellFormed(edited);
  assert.deepStrictEqual(await notesOf(edited), [tricky, 'Talk about bullets.', 'Notes for the table slide']);

  // Clearing notes leaves an empty body.
  const cleared = await applyNotesEdits(edited, new Map([[0, '']]));
  assert.strictEqual((await notesOf(cleared))[0], '');

  // 2. Deck without any notes: adding notes creates a notes slide, a notes master and a theme.
  const bare = await stripNotes(fixture);
  assert.deepStrictEqual(await notesOf(bare), ['', '', '']);
  const added = await applyNotesEdits(bare, new Map([[1, 'Brand new notes'], [2, 'More']]));
  await assertWellFormed(added);
  assert.deepStrictEqual(await notesOf(added), ['', 'Brand new notes', 'More']);

  const zip = await JSZip.loadAsync(added);
  const pres = await zip.file('ppt/presentation.xml').async('string');
  assert.match(pres, /<\/p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst><p:sldIdLst>/, 'notesMasterIdLst in schema order');
  const ct = await zip.file('[Content_Types].xml').async('string');
  for (const part of ['/ppt/notesMasters/notesMaster1.xml', '/ppt/notesSlides/notesSlide1.xml', '/ppt/notesSlides/notesSlide2.xml']) {
    assert.ok(ct.includes(`PartName="${part}"`), `content type for ${part}`);
  }
  const masterRels = await zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels').async('string');
  const themeTarget = /Target="\.\.\/theme\/(theme\d+\.xml)"/.exec(masterRels)[1];
  assert.ok(zip.file(`ppt/theme/${themeTarget}`), 'notes master theme exists');
  const slideRels = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string');
  assert.match(slideRels, /notesSlide" Target="\.\.\/notesSlides\/notesSlide1\.xml"/);

  // A second round of edits reuses the master instead of adding another.
  const again = await applyNotesEdits(added, new Map([[0, 'First slide too']]));
  const zip2 = await JSZip.loadAsync(again);
  assert.ok(!zip2.file('ppt/notesMasters/notesMaster2.xml'), 'no duplicate notes master');
  assert.deepStrictEqual(await notesOf(again), ['First slide too', 'Brand new notes', 'More']);

  // 3. Empty notes on a slide that has none changes nothing.
  const untouched = await applyNotesEdits(bare, new Map([[0, '  ']]));
  assert.ok(!(await JSZip.loadAsync(untouched)).file('ppt/notesMasters/notesMaster1.xml'));

  console.log('ok  notes editing (edit, clear, create notes + master, reuse master)');
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
