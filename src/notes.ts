/*
 * Writes speaker-notes edits back into a .pptx.
 *
 * Existing notes slides get their body placeholder's paragraphs replaced. Slides without
 * notes get a new notes slide, and a notes master (with its own theme) is created first
 * when the deck has none.
 */
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

const REL_TYPE = {
  notesSlide: `${NS.r}/notesSlide`,
  notesMaster: `${NS.r}/notesMaster`,
  slide: `${NS.r}/slide`,
  theme: `${NS.r}/theme`,
};

const CONTENT_TYPE = {
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
};

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Map of slide index (0-based, presentation order) to the full notes text for that slide. */
export type NotesEdits = ReadonlyMap<number, string>;

export async function applyNotesEdits(pptx: Uint8Array, edits: NotesEdits): Promise<Uint8Array> {
  if (edits.size === 0) return pptx;
  const pkg = await Package.open(pptx);
  const slides = await pkg.slidePaths();
  for (const [index, text] of edits) {
    const slidePath = slides[index];
    if (!slidePath) throw new Error(`Slide ${index + 1} does not exist.`);
    await pkg.setNotes(slidePath, text);
  }
  return pkg.save();
}

// ---------------------------------------------------------------------------

type Doc = ReturnType<DOMParser['parseFromString']>;
type El = NonNullable<Doc['documentElement']>;

function children(el: El | null, localName?: string): El[] {
  const out: El[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (!localName || (n as El).localName === localName)) out.push(n as El);
  }
  return out;
}

function child(el: El | null, localName: string): El | null {
  return children(el, localName)[0] ?? null;
}

function dirOf(p: string): string {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i);
}

function relsPathFor(partPath: string): string {
  const dir = dirOf(partPath);
  return `${dir}/_rels/${partPath.slice(dir.length + 1)}.rels`;
}

function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = baseDir ? baseDir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

/** Relative target from one part's directory to another part, as used in .rels files. */
function relativeTarget(fromPart: string, toPart: string): string {
  const from = dirOf(fromPart).split('/').filter(Boolean);
  const to = toPart.split('/');
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...from.slice(i).map(() => '..'), ...to.slice(i)].join('/');
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // Strip characters XML 1.0 can't carry.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

class Rels {
  constructor(readonly partPath: string, readonly doc: Doc) {}

  get root(): El {
    return this.doc.documentElement as El;
  }

  all(): { id: string; type: string; target: string }[] {
    return children(this.root, 'Relationship').map((r) => ({
      id: r.getAttribute('Id') ?? '',
      type: r.getAttribute('Type') ?? '',
      target: r.getAttribute('TargetMode') === 'External'
        ? r.getAttribute('Target') ?? ''
        : resolvePath(dirOf(this.partPath), r.getAttribute('Target') ?? ''),
    }));
  }

  byType(type: string) {
    return this.all().find((r) => r.type === type) ?? null;
  }

  add(type: string, targetPart: string): string {
    const used = new Set(this.all().map((r) => r.id));
    let n = 1;
    while (used.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    const el = this.doc.createElementNS(NS.rel, 'Relationship');
    el.setAttribute('Id', id);
    el.setAttribute('Type', type);
    el.setAttribute('Target', relativeTarget(this.partPath, targetPart));
    this.root.appendChild(el);
    return id;
  }
}

class Package {
  private docs = new Map<string, Doc>();
  private dirty = new Set<string>();

  private constructor(private readonly zip: JSZip) {}

  static async open(bytes: Uint8Array): Promise<Package> {
    return new Package(await JSZip.loadAsync(bytes));
  }

  async xml(path: string): Promise<Doc | null> {
    const cached = this.docs.get(path);
    if (cached) return cached;
    const file = this.zip.file(path);
    if (!file) return null;
    const doc = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    this.docs.set(path, doc);
    return doc;
  }

  private setXml(path: string, doc: Doc) {
    this.docs.set(path, doc);
    this.dirty.add(path);
  }

  private touch(path: string) {
    this.dirty.add(path);
  }

  async rels(partPath: string): Promise<Rels> {
    const path = relsPathFor(partPath);
    let doc = await this.xml(path);
    if (!doc) {
      doc = new DOMParser().parseFromString(`${XML_DECL}<Relationships xmlns="${NS.rel}"/>`, 'application/xml');
      this.setXml(path, doc);
    }
    return new Rels(partPath, doc);
  }

  private async addRel(partPath: string, type: string, target: string): Promise<string> {
    const rels = await this.rels(partPath);
    const id = rels.add(type, target);
    this.touch(relsPathFor(partPath));
    return id;
  }

  /** True if the part is in the zip or was created in this session (not yet written). */
  private exists(path: string): boolean {
    return !!this.zip.file(path) || this.docs.has(path);
  }

  /** First unused path like ppt/notesSlides/notesSlide7.xml. */
  private freePartName(dir: string, base: string): string {
    let n = 1;
    while (this.exists(`${dir}/${base}${n}.xml`)) n++;
    return `${dir}/${base}${n}.xml`;
  }

  private async addContentType(partPath: string, contentType: string) {
    const doc = await this.xml('[Content_Types].xml');
    if (!doc) throw new Error('[Content_Types].xml is missing.');
    const root = doc.documentElement as El;
    const partName = `/${partPath}`;
    if (children(root, 'Override').some((o) => o.getAttribute('PartName') === partName)) return;
    const el = doc.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', partName);
    el.setAttribute('ContentType', contentType);
    root.appendChild(el);
    this.touch('[Content_Types].xml');
  }

  async slidePaths(): Promise<string[]> {
    const pres = await this.xml('ppt/presentation.xml');
    if (!pres) throw new Error('ppt/presentation.xml is missing.');
    const rels = await this.rels('ppt/presentation.xml');
    const byId = new Map(rels.all().map((r) => [r.id, r.target]));
    return children(child(pres.documentElement as El, 'sldIdLst'), 'sldId')
      .map((s) => byId.get(s.getAttributeNS(NS.r, 'id') ?? s.getAttribute('r:id') ?? ''))
      .filter((p): p is string => !!p);
  }

  async setNotes(slidePath: string, text: string) {
    const slideRels = await this.rels(slidePath);
    const existing = slideRels.byType(REL_TYPE.notesSlide);
    if (existing && this.exists(existing.target)) {
      await this.rewriteNotesBody(existing.target, text);
      return;
    }
    if (text.trim() === '') return; // Nothing to add.
    const masterPath = await this.ensureNotesMaster();
    const notesPath = this.freePartName('ppt/notesSlides', 'notesSlide');
    const doc = new DOMParser().parseFromString(notesSlideXml(paragraphsXml(text, null)), 'application/xml');
    this.setXml(notesPath, doc);
    await this.addContentType(notesPath, CONTENT_TYPE.notesSlide);
    await this.addRel(notesPath, REL_TYPE.notesMaster, masterPath);
    await this.addRel(notesPath, REL_TYPE.slide, slidePath);
    await this.addRel(slidePath, REL_TYPE.notesSlide, notesPath);
  }

  private async rewriteNotesBody(notesPath: string, text: string) {
    const doc = await this.xml(notesPath);
    if (!doc) return;
    const root = doc.documentElement as El;
    const spTree = child(child(root, 'cSld'), 'spTree');
    if (!spTree) throw new Error(`${notesPath} has no shape tree.`);

    let body = children(spTree, 'sp').find((sp) => {
      const ph = child(child(child(sp, 'nvSpPr'), 'nvPr'), 'ph');
      return ph?.getAttribute('type') === 'body';
    });
    if (!body) {
      if (text.trim() === '') return;
      const ids = Array.from(doc.getElementsByTagNameNS(NS.p, 'cNvPr')).map((e) => Number(e.getAttribute('id')) || 0);
      const frag = parseFragment(doc, bodyShapeXml(Math.max(1, ...ids) + 1, ''));
      body = frag[0];
      spTree.appendChild(body);
    }

    let txBody = child(body, 'txBody');
    if (!txBody) {
      txBody = parseFragment(doc, `<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>`)[0];
      body.appendChild(txBody);
    }

    // Keep the language of the first existing run so spell-check settings survive.
    const lang = txBody.getElementsByTagNameNS(NS.a, 'rPr')[0]?.getAttribute('lang') ?? null;

    for (const p of children(txBody, 'p')) txBody.removeChild(p);
    for (const p of parseFragment(doc, paragraphsXml(text, lang))) txBody.appendChild(p);
    this.touch(notesPath);
  }

  private async ensureNotesMaster(): Promise<string> {
    const presRels = await this.rels('ppt/presentation.xml');
    const existing = presRels.byType(REL_TYPE.notesMaster);
    if (existing && this.exists(existing.target)) return existing.target;

    // The notes master needs a theme of its own; copy the first slide master's theme.
    const pres = (await this.xml('ppt/presentation.xml'))!;
    const presRoot = pres.documentElement as El;
    const masterId = child(child(presRoot, 'sldMasterIdLst'), 'sldMasterId');
    const masterRelId = masterId?.getAttributeNS(NS.r, 'id') ?? masterId?.getAttribute('r:id') ?? '';
    const slideMaster = presRels.all().find((r) => r.id === masterRelId);
    const masterTheme = slideMaster ? (await this.rels(slideMaster.target)).byType(REL_TYPE.theme) : null;
    const themeFile = masterTheme && this.zip.file(masterTheme.target);
    if (!themeFile) throw new Error('Cannot add notes: the deck has no theme to base a notes master on.');
    const themePath = this.freePartName('ppt/theme', 'theme');
    this.zip.file(themePath, await themeFile.async('uint8array'));
    await this.addContentType(themePath, CONTENT_TYPE.theme);

    const masterPath = this.freePartName('ppt/notesMasters', 'notesMaster');
    this.setXml(masterPath, new DOMParser().parseFromString(NOTES_MASTER_XML, 'application/xml'));
    await this.addContentType(masterPath, CONTENT_TYPE.notesMaster);
    await this.addRel(masterPath, REL_TYPE.theme, themePath);

    const relId = await this.addRel('ppt/presentation.xml', REL_TYPE.notesMaster, masterPath);
    const lst = parseFragment(pres, `<p:notesMasterIdLst><p:notesMasterId r:id="${relId}"/></p:notesMasterIdLst>`)[0];
    // Schema order: sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst, ...
    const old = child(presRoot, 'notesMasterIdLst');
    if (old) presRoot.removeChild(old);
    const after = child(presRoot, 'sldMasterIdLst');
    presRoot.insertBefore(lst, after ? after.nextSibling : presRoot.firstChild);
    if (!child(presRoot, 'notesSz')) {
      const sz = parseFragment(pres, `<p:notesSz cx="6858000" cy="9144000"/>`)[0];
      const sldSz = child(presRoot, 'sldSz');
      presRoot.insertBefore(sz, sldSz ? sldSz.nextSibling : null);
    }
    this.touch('ppt/presentation.xml');
    return masterPath;
  }

  async save(): Promise<Uint8Array> {
    const serializer = new XMLSerializer();
    for (const path of this.dirty) {
      const doc = this.docs.get(path);
      if (!doc) continue;
      let xml = serializer.serializeToString(doc);
      if (!xml.startsWith('<?xml')) xml = XML_DECL + xml;
      this.zip.file(path, xml);
    }
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }
}

/** Parses XML using the a:/p:/r: prefixes and imports the top-level elements into doc. */
function parseFragment(doc: Doc, xml: string): El[] {
  const wrapper = new DOMParser().parseFromString(
    `<w xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}">${xml}</w>`,
    'application/xml',
  );
  return children(wrapper.documentElement as El).map((e) => doc.importNode(e, true) as El);
}

function paragraphsXml(text: string, lang: string | null): string {
  const langAttr = ` lang="${escapeXml(lang || 'en-US')}" dirty="0"`;
  return text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) =>
      line === ''
        ? `<a:p><a:endParaRPr${langAttr}/></a:p>`
        : `<a:p><a:r><a:rPr${langAttr}/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
    )
    .join('');
}

function bodyShapeXml(id: number, paragraphs: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Notes Placeholder ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs || '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'}</p:txBody></p:sp>`;
}

function notesSlideXml(paragraphs: string): string {
  return `${XML_DECL}<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
    bodyShapeXml(3, paragraphs) +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

const NOTES_LEVELS = [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600]
  .map((marL, i) =>
    `<a:lvl${i + 1}pPr marL="${marL}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
    `<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
    `<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`)
  .join('');

// A minimal portrait notes master: slide image on top, notes body below.
const NOTES_MASTER_XML = `${XML_DECL}<p:notesMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld>` +
  `<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
  `<p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/>` +
  `<a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>` +
  `</p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `<p:notesStyle>${NOTES_LEVELS}</p:notesStyle></p:notesMaster>`;
