/*
 * PPTX parser: turns the OOXML parts of a .pptx (opened with JSZip) into a simple
 * render model of absolutely positioned elements, measured in CSS pixels (96 dpi).
 *
 * Runs in the webview (browser DOMParser) and in Node tests (@xmldom/xmldom).
 * All element lookups go through localName so namespace prefixes don't matter.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.PptxParser = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const EMU_PER_PX = 9525;
  const PT_TO_PX = 96 / 72;
  const DEFAULT_FONT_PT = 18;

  // ---------------------------------------------------------------------------
  // XML helpers
  // ---------------------------------------------------------------------------

  function kids(el, name) {
    const out = [];
    if (!el) return out;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && (!name || n.localName === name)) out.push(n);
    }
    return out;
  }

  function kid(el, name) {
    if (!el) return null;
    for (let n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && n.localName === name) return n;
    }
    return null;
  }

  function path(el, ...names) {
    for (const name of names) {
      el = kid(el, name);
      if (!el) return null;
    }
    return el;
  }

  function attr(el, name) {
    if (!el) return null;
    const v = el.getAttribute(name);
    return v === '' || v === undefined ? null : v;
  }

  function num(el, name, fallback) {
    const v = attr(el, name);
    if (v == null) return fallback;
    const n = Number(v);
    return Number.isFinite(n) ? n : fallback;
  }

  function bool(el, name) {
    const v = attr(el, name);
    if (v == null) return undefined;
    return v === '1' || v === 'true' || v === 'on';
  }

  function textOf(el) {
    return el ? el.textContent || '' : '';
  }

  /** Expands mc:AlternateContent into its Fallback (or Choice) children. */
  function flatKids(el) {
    const out = [];
    for (const c of kids(el)) {
      if (c.localName === 'AlternateContent') {
        const branch = kid(c, 'Fallback') || kid(c, 'Choice');
        if (branch) out.push(...flatKids(branch));
      } else {
        out.push(c);
      }
    }
    return out;
  }

  const px = (emu) => emu / EMU_PER_PX;

  // ---------------------------------------------------------------------------
  // Paths and relationships
  // ---------------------------------------------------------------------------

  function dirOf(p) {
    const i = p.lastIndexOf('/');
    return i < 0 ? '' : p.slice(0, i);
  }

  function resolvePath(baseDir, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = baseDir ? baseDir.split('/') : [];
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  }

  function relsPathFor(partPath) {
    const dir = dirOf(partPath);
    const base = partPath.slice(dir.length + (dir ? 1 : 0));
    return (dir ? dir + '/' : '') + '_rels/' + base + '.rels';
  }

  const MIME = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff',
    tiff: 'image/tiff', ico: 'image/x-icon',
  };

  // ---------------------------------------------------------------------------
  // Colors
  // ---------------------------------------------------------------------------

  const PRESET_COLORS = {
    black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
    yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080',
    darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080',
    navy: '000080', teal: '008080', maroon: '800000', olive: '808000', silver: 'C0C0C0',
    lime: '00FF00', darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400',
  };

  const SYS_COLORS = { windowText: '000000', window: 'FFFFFF', btnFace: 'F0F0F0', btnText: '000000' };

  function hexToRgb(hex) {
    const h = (hex || '000000').replace('#', '');
    return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16), a: 1 };
  }

  function rgbToHsl({ r, g, b }) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    return { h, s, l };
  }

  function hslToRgb({ h, s, l }) {
    if (s === 0) {
      const v = l * 255;
      return { r: v, g: v, b: v };
    }
    const hue = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return { r: hue(p, q, h + 1 / 3) * 255, g: hue(p, q, h) * 255, b: hue(p, q, h - 1 / 3) * 255 };
  }

  const clamp01 = (v) => Math.max(0, Math.min(1, v));

  function applyColorMods(c, colorEl) {
    let { r, g, b, a } = c;
    for (const m of kids(colorEl)) {
      const v = num(m, 'val', 0) / 100000;
      switch (m.localName) {
        case 'alpha': a = v; break;
        case 'alphaMod': a *= v; break;
        case 'tint': r += (255 - r) * (1 - v); g += (255 - g) * (1 - v); b += (255 - b) * (1 - v); break;
        case 'shade': r *= v; g *= v; b *= v; break;
        case 'lumMod': case 'lumOff': case 'satMod': {
          const hsl = rgbToHsl({ r, g, b });
          if (m.localName === 'lumMod') hsl.l = clamp01(hsl.l * v);
          else if (m.localName === 'lumOff') hsl.l = clamp01(hsl.l + v);
          else hsl.s = clamp01(hsl.s * v);
          ({ r, g, b } = hslToRgb(hsl));
          break;
        }
        default: break;
      }
    }
    return { r, g, b, a };
  }

  function cssColor(c) {
    if (!c) return null;
    const f = (v) => Math.round(Math.max(0, Math.min(255, v)));
    return c.a >= 1 ? `rgb(${f(c.r)}, ${f(c.g)}, ${f(c.b)})` : `rgba(${f(c.r)}, ${f(c.g)}, ${f(c.b)}, ${+c.a.toFixed(3)})`;
  }

  const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);

  /** Finds the first color element (srgbClr, schemeClr, ...) under el. */
  function colorChild(el) {
    for (const c of kids(el)) if (COLOR_TAGS.has(c.localName)) return c;
    return null;
  }

  // ---------------------------------------------------------------------------
  // Preset geometry -> SVG path
  // ---------------------------------------------------------------------------

  function adjValues(geomEl) {
    const out = {};
    for (const gd of kids(kid(geomEl, 'avLst'), 'gd')) {
      const m = /val\s+(-?\d+)/.exec(attr(gd, 'fmla') || '');
      if (m) out[attr(gd, 'name')] = Number(m[1]);
    }
    return out;
  }

  const LINE_PRESETS = /^(line|straightConnector1|bentConnector\d|curvedConnector\d|lineInv)$/;

  function poly(pts) {
    return 'M' + pts.map((p) => `${r2(p[0])} ${r2(p[1])}`).join(' L') + ' Z';
  }
  const r2 = (n) => Math.round(n * 100) / 100;

  function ellipsePath(w, h) {
    const rx = w / 2, ry = h / 2;
    return `M0 ${r2(ry)} A${r2(rx)} ${r2(ry)} 0 1 1 ${r2(w)} ${r2(ry)} A${r2(rx)} ${r2(ry)} 0 1 1 0 ${r2(ry)} Z`;
  }

  function roundRectPath(w, h, rad) {
    const r = Math.min(rad, w / 2, h / 2);
    if (r <= 0) return poly([[0, 0], [w, 0], [w, h], [0, h]]);
    return `M${r2(r)} 0 H${r2(w - r)} A${r2(r)} ${r2(r)} 0 0 1 ${r2(w)} ${r2(r)} V${r2(h - r)} ` +
      `A${r2(r)} ${r2(r)} 0 0 1 ${r2(w - r)} ${r2(h)} H${r2(r)} A${r2(r)} ${r2(r)} 0 0 1 0 ${r2(h - r)} ` +
      `V${r2(r)} A${r2(r)} ${r2(r)} 0 0 1 ${r2(r)} 0 Z`;
  }

  function starPath(w, h, points, innerRatio) {
    const pts = [];
    for (let i = 0; i < points * 2; i++) {
      const ang = -Math.PI / 2 + (i * Math.PI) / points;
      const rr = i % 2 === 0 ? 1 : innerRatio;
      pts.push([w / 2 + (w / 2) * rr * Math.cos(ang), h / 2 + (h / 2) * rr * Math.sin(ang)]);
    }
    return poly(pts);
  }

  function presetPath(name, w, h, adj) {
    const ss = Math.min(w, h);
    const a = (key, def) => (adj[key] != null ? adj[key] : def) / 100000;
    switch (name) {
      case 'ellipse': case 'flowChartConnector': case 'cloud': case 'flowChartOr': case 'flowChartSummingJunction':
        return ellipsePath(w, h);
      case 'roundRect': case 'round2SameRect': case 'round1Rect': case 'flowChartAlternateProcess':
        return roundRectPath(w, h, ss * a('adj', 16667));
      case 'flowChartTerminator':
        return roundRectPath(w, h, ss / 2);
      case 'triangle': case 'flowChartExtract': {
        const x = w * a('adj', 50000);
        return poly([[x, 0], [w, h], [0, h]]);
      }
      case 'flowChartMerge':
        return poly([[0, 0], [w, 0], [w / 2, h]]);
      case 'rtTriangle':
        return poly([[0, 0], [w, h], [0, h]]);
      case 'diamond': case 'flowChartDecision':
        return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
      case 'parallelogram': case 'flowChartInputOutput': {
        const o = name === 'parallelogram' ? ss * a('adj', 25000) : w / 5;
        return poly([[o, 0], [w, 0], [w - o, h], [0, h]]);
      }
      case 'trapezoid': {
        const o = ss * a('adj', 25000);
        return poly([[0, h], [o, 0], [w - o, 0], [w, h]]);
      }
      case 'hexagon': case 'flowChartPreparation': {
        const o = ss * a('adj', 25000);
        return poly([[0, h / 2], [o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h]]);
      }
      case 'octagon': {
        const o = ss * a('adj', 29289);
        return poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]);
      }
      case 'pentagon':
        return poly([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]]);
      case 'homePlate': {
        const o = ss * a('adj', 50000);
        return poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]);
      }
      case 'chevron': {
        const o = ss * a('adj', 50000);
        return poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]);
      }
      case 'plus': case 'mathPlus': {
        const o = ss * a('adj', 25000);
        return poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h],
          [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]);
      }
      case 'rightArrow': case 'leftArrow': {
        const sh = h * a('adj1', 50000), hl = ss * a('adj2', 50000), y1 = (h - sh) / 2;
        let pts = [[0, y1], [w - hl, y1], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, y1 + sh], [0, y1 + sh]];
        if (name === 'leftArrow') pts = pts.map(([x, y]) => [w - x, y]);
        return poly(pts);
      }
      case 'downArrow': case 'upArrow': {
        const sw = w * a('adj1', 50000), hl = ss * a('adj2', 50000), x1 = (w - sw) / 2;
        let pts = [[x1, 0], [x1 + sw, 0], [x1 + sw, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [x1, h - hl]];
        if (name === 'upArrow') pts = pts.map(([x, y]) => [x, h - y]);
        return poly(pts);
      }
      case 'star5': return starPath(w, h, 5, 0.382);
      case 'star4': return starPath(w, h, 4, 0.25);
      case 'star6': return starPath(w, h, 6, 0.577);
      case 'star8': return starPath(w, h, 8, 0.7);
      case 'snip1Rect': {
        const o = ss * a('adj', 16667);
        return poly([[0, 0], [w - o, 0], [w, o], [w, h], [0, h]]);
      }
      default:
        if (LINE_PRESETS.test(name)) return `M0 0 L${r2(w)} ${r2(h)}`;
        return poly([[0, 0], [w, 0], [w, h], [0, h]]);
    }
  }

  function customPath(custGeom, w, h) {
    let d = '';
    for (const p of kids(kid(custGeom, 'pathLst'), 'path')) {
      const pw = num(p, 'w', 0) || w * EMU_PER_PX;
      const ph = num(p, 'h', 0) || h * EMU_PER_PX;
      const sx = pw ? w / pw : 0, sy = ph ? h / ph : 0;
      let cx = 0, cy = 0;
      const pt = (el) => [num(el, 'x', 0) * sx, num(el, 'y', 0) * sy];
      for (const cmd of kids(p)) {
        const pts = kids(cmd, 'pt').map(pt);
        switch (cmd.localName) {
          case 'moveTo': [cx, cy] = pts[0]; d += `M${r2(cx)} ${r2(cy)} `; break;
          case 'lnTo': [cx, cy] = pts[0]; d += `L${r2(cx)} ${r2(cy)} `; break;
          case 'cubicBezTo':
            d += 'C' + pts.map((q) => `${r2(q[0])} ${r2(q[1])}`).join(' ') + ' ';
            [cx, cy] = pts[2];
            break;
          case 'quadBezTo':
            d += 'Q' + pts.map((q) => `${r2(q[0])} ${r2(q[1])}`).join(' ') + ' ';
            [cx, cy] = pts[1];
            break;
          case 'arcTo': {
            const wr = num(cmd, 'wR', 0) * sx, hr = num(cmd, 'hR', 0) * sy;
            const st = (num(cmd, 'stAng', 0) / 60000) * (Math.PI / 180);
            const sw = (num(cmd, 'swAng', 0) / 60000) * (Math.PI / 180);
            const ox = cx - wr * Math.cos(st), oy = cy - hr * Math.sin(st);
            const ex = ox + wr * Math.cos(st + sw), ey = oy + hr * Math.sin(st + sw);
            d += `A${r2(wr)} ${r2(hr)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${r2(ex)} ${r2(ey)} `;
            cx = ex; cy = ey;
            break;
          }
          case 'close': d += 'Z '; break;
          default: break;
        }
      }
    }
    return d.trim() || poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }

  // ---------------------------------------------------------------------------
  // Numbering for auto-numbered bullets
  // ---------------------------------------------------------------------------

  function toRoman(n) {
    const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'],
      [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let out = '';
    for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
    return out;
  }

  function toAlpha(n) {
    let out = '';
    while (n > 0) { n--; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); }
    return out;
  }

  function formatAutoNum(type, n) {
    const t = type || 'arabicPeriod';
    let core = String(n);
    if (t.startsWith('romanLc')) core = toRoman(n);
    else if (t.startsWith('romanUc')) core = toRoman(n).toUpperCase();
    else if (t.startsWith('alphaLc')) core = toAlpha(n);
    else if (t.startsWith('alphaUc')) core = toAlpha(n).toUpperCase();
    if (t.endsWith('ParenBoth')) return `(${core})`;
    if (t.endsWith('ParenR')) return `${core})`;
    if (t.endsWith('Plain')) return core;
    return `${core}.`;
  }

  // ---------------------------------------------------------------------------
  // Parser
  // ---------------------------------------------------------------------------

  class Parser {
    constructor(zip, opts) {
      this.zip = zip;
      this.DOMParser = opts.DOMParser;
      this.makeImageUrl = opts.makeImageUrl || (() => null);
      this.xmlCache = new Map();
      this.relsCache = new Map();
      this.imageCache = new Map();
      this.masterCache = new Map();
      this.layoutCache = new Map();
      this.tableStyles = null;
    }

    async xml(p) {
      if (this.xmlCache.has(p)) return this.xmlCache.get(p);
      const f = this.zip.file(p);
      let rootEl = null;
      if (f) {
        const text = await f.async('string');
        const doc = new this.DOMParser().parseFromString(text, 'application/xml');
        rootEl = doc.documentElement;
      }
      this.xmlCache.set(p, rootEl);
      return rootEl;
    }

    async rels(partPath) {
      if (this.relsCache.has(partPath)) return this.relsCache.get(partPath);
      const map = new Map();
      const rootEl = await this.xml(relsPathFor(partPath));
      for (const r of kids(rootEl, 'Relationship')) {
        const external = attr(r, 'TargetMode') === 'External';
        const target = attr(r, 'Target') || '';
        map.set(attr(r, 'Id'), {
          type: attr(r, 'Type') || '',
          target: external ? target : resolvePath(dirOf(partPath), target),
          external,
        });
      }
      this.relsCache.set(partPath, map);
      return map;
    }

    relByType(rels, suffix) {
      for (const r of rels.values()) if (r.type.endsWith(suffix)) return r;
      return null;
    }

    async image(partPath) {
      if (this.imageCache.has(partPath)) return this.imageCache.get(partPath);
      const f = this.zip.file(partPath);
      let url = null;
      const ext = (partPath.split('.').pop() || '').toLowerCase();
      const mime = MIME[ext];
      if (f && mime) {
        const bytes = await f.async('uint8array');
        url = this.makeImageUrl(bytes, mime, partPath);
      }
      this.imageCache.set(partPath, url);
      return url;
    }

    // ---- top level -----------------------------------------------------------

    async parse() {
      const pres = await this.xml('ppt/presentation.xml');
      if (!pres) throw new Error('Not a PowerPoint file: ppt/presentation.xml is missing.');
      const presRels = await this.rels('ppt/presentation.xml');

      const sldSz = kid(pres, 'sldSz');
      this.slideW = num(sldSz, 'cx', 9144000);
      this.slideH = num(sldSz, 'cy', 6858000);
      this.defaultTextStyle = kid(pres, 'defaultTextStyle');

      const slidePaths = [];
      for (const s of kids(kid(pres, 'sldIdLst'), 'sldId')) {
        const rel = presRels.get(attr(s, 'r:id'));
        if (rel) slidePaths.push(rel.target);
      }

      const slides = [];
      for (let i = 0; i < slidePaths.length; i++) {
        try {
          slides.push(await this.parseSlide(slidePaths[i], i + 1));
        } catch (err) {
          slides.push({ number: i + 1, hidden: false, background: null, elements: [], notes: '', error: String(err && err.message || err) });
        }
      }
      return { width: px(this.slideW), height: px(this.slideH), slides };
    }

    // ---- masters, layouts, themes ------------------------------------------

    async loadTheme(p) {
      const rootEl = await this.xml(p);
      const theme = { colors: {}, majorFont: null, minorFont: null, bgFills: [], fills: [], lines: [] };
      const elements = kid(rootEl, 'themeElements');
      for (const c of kids(kid(elements, 'clrScheme'))) {
        const v = colorChild(c);
        if (!v) continue;
        if (v.localName === 'srgbClr') theme.colors[c.localName] = attr(v, 'val');
        else if (v.localName === 'sysClr') theme.colors[c.localName] = attr(v, 'lastClr') || SYS_COLORS[attr(v, 'val')] || '000000';
      }
      const fonts = kid(elements, 'fontScheme');
      theme.majorFont = attr(path(fonts, 'majorFont', 'latin'), 'typeface');
      theme.minorFont = attr(path(fonts, 'minorFont', 'latin'), 'typeface');
      const fmt = kid(elements, 'fmtScheme');
      theme.fills = kids(kid(fmt, 'fillStyleLst'));
      theme.lines = kids(kid(fmt, 'lnStyleLst'));
      theme.bgFills = kids(kid(fmt, 'bgFillStyleLst'));
      return theme;
    }

    collectPlaceholders(spTree) {
      const list = [];
      const walk = (tree) => {
        for (const el of flatKids(tree)) {
          if (el.localName === 'grpSp') { walk(el); continue; }
          const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvGraphicFramePr');
          const ph = path(nv, 'nvPr', 'ph');
          if (ph) list.push({ type: attr(ph, 'type') || 'obj', idx: attr(ph, 'idx'), el });
        }
      };
      walk(spTree);
      return list;
    }

    async loadMaster(p) {
      if (this.masterCache.has(p)) return this.masterCache.get(p);
      const rootEl = await this.xml(p);
      const rels = await this.rels(p);
      const themeRel = this.relByType(rels, '/theme');
      const theme = themeRel ? await this.loadTheme(themeRel.target) : { colors: {}, fills: [], lines: [], bgFills: [] };
      const clrMap = {};
      const cm = kid(rootEl, 'clrMap');
      if (cm) for (const a of Array.from(cm.attributes)) clrMap[a.localName || a.name] = a.value;
      const txStyles = kid(rootEl, 'txStyles');
      const master = {
        path: p, root: rootEl, rels, theme, clrMap,
        titleStyle: kid(txStyles, 'titleStyle'),
        bodyStyle: kid(txStyles, 'bodyStyle'),
        otherStyle: kid(txStyles, 'otherStyle'),
        spTree: path(rootEl, 'cSld', 'spTree'),
      };
      master.placeholders = this.collectPlaceholders(master.spTree);
      this.masterCache.set(p, master);
      return master;
    }

    async loadLayout(p) {
      if (this.layoutCache.has(p)) return this.layoutCache.get(p);
      const rootEl = await this.xml(p);
      const rels = await this.rels(p);
      const masterRel = this.relByType(rels, '/slideMaster');
      const master = masterRel ? await this.loadMaster(masterRel.target) : null;
      const layout = {
        path: p, root: rootEl, rels, master,
        spTree: path(rootEl, 'cSld', 'spTree'),
        showMasterSp: bool(rootEl, 'showMasterSp') !== false,
      };
      layout.placeholders = this.collectPlaceholders(layout.spTree);
      this.layoutCache.set(p, layout);
      return layout;
    }

    makeClrMap(master, ...parts) {
      let map = Object.assign({}, master ? master.clrMap : {});
      for (const part of parts) {
        const ovr = path(part, 'clrMapOvr', 'overrideClrMapping');
        if (ovr) {
          map = {};
          for (const a of Array.from(ovr.attributes)) map[a.localName || a.name] = a.value;
        }
      }
      return map;
    }

    // ---- colors & fills -----------------------------------------------------

    /** Resolves a color element to {r,g,b,a}. ctx carries theme + clrMap; phClr substitutes placeholder color. */
    color(colorEl, ctx, phClr) {
      if (!colorEl) return null;
      let base;
      switch (colorEl.localName) {
        case 'srgbClr': base = hexToRgb(attr(colorEl, 'val')); break;
        case 'sysClr': base = hexToRgb(attr(colorEl, 'lastClr') || SYS_COLORS[attr(colorEl, 'val')]); break;
        case 'prstClr': base = hexToRgb(PRESET_COLORS[attr(colorEl, 'val')] || '000000'); break;
        case 'scrgbClr':
          base = { r: num(colorEl, 'r', 0) / 100000 * 255, g: num(colorEl, 'g', 0) / 100000 * 255, b: num(colorEl, 'b', 0) / 100000 * 255, a: 1 };
          break;
        case 'hslClr': {
          const rgb = hslToRgb({ h: num(colorEl, 'hue', 0) / 21600000, s: num(colorEl, 'sat', 0) / 100000, l: num(colorEl, 'lum', 0) / 100000 });
          base = Object.assign(rgb, { a: 1 });
          break;
        }
        case 'schemeClr': {
          const val = attr(colorEl, 'val');
          if (val === 'phClr') {
            base = phClr ? Object.assign({}, phClr) : hexToRgb('000000');
          } else {
            const mapped = (ctx.clrMap && ctx.clrMap[val]) || val;
            const hex = ctx.theme.colors[mapped] || ctx.theme.colors[val];
            base = hexToRgb(hex || '000000');
          }
          break;
        }
        default: return null;
      }
      return applyColorMods(base, colorEl);
    }

    themeColor(name, ctx) {
      const mapped = (ctx.clrMap && ctx.clrMap[name]) || name;
      return hexToRgb(ctx.theme.colors[mapped] || '000000');
    }

    /**
     * Parses the fill declared directly under `el` (spPr, bgPr, tcPr, ...).
     * Returns undefined when no fill is declared, so callers can fall back.
     */
    async fill(el, ctx, phClr) {
      if (!el) return undefined;
      for (const c of kids(el)) {
        const f = await this.fillFrom(c, ctx, phClr);
        if (f) return f;
      }
      return undefined;
    }

    /** Parses a single fill element (solidFill, gradFill, ...); undefined if c isn't one. */
    async fillFrom(c, ctx, phClr) {
      {
        switch (c.localName) {
          case 'noFill': return { type: 'none' };
          case 'solidFill': {
            const col = this.color(colorChild(c), ctx, phClr);
            return col ? { type: 'solid', color: cssColor(col) } : { type: 'none' };
          }
          case 'gradFill': return this.gradient(c, ctx, phClr);
          case 'blipFill': {
            const url = await this.blipUrl(c, ctx);
            if (!url) return { type: 'none' };
            const tile = kid(c, 'tile');
            return { type: 'image', url, tile: !!tile };
          }
          case 'pattFill': {
            const fg = this.color(colorChild(kid(c, 'fgClr')), ctx, phClr);
            return fg ? { type: 'solid', color: cssColor(fg) } : { type: 'none' };
          }
          case 'grpFill': return ctx.groupFill || undefined;
          default: break;
        }
      }
      return undefined;
    }

    gradient(gradEl, ctx, phClr) {
      const stops = kids(kid(gradEl, 'gsLst'), 'gs')
        .map((gs) => ({ pos: num(gs, 'pos', 0) / 1000, color: cssColor(this.color(colorChild(gs), ctx, phClr)) }))
        .filter((s) => s.color)
        .sort((x, y) => x.pos - y.pos);
      if (!stops.length) return { type: 'none' };
      const list = stops.map((s) => `${s.color} ${r2(s.pos)}%`).join(', ');
      const pathEl = kid(gradEl, 'path');
      if (pathEl) return { type: 'gradient', css: `radial-gradient(ellipse at center, ${list})`, color: stops[0].color };
      const ang = num(kid(gradEl, 'lin'), 'ang', 0) / 60000;
      return { type: 'gradient', css: `linear-gradient(${r2(ang + 90)}deg, ${list})`, color: stops[0].color };
    }

    async blipUrl(blipFillEl, ctx) {
      const blip = kid(blipFillEl, 'blip');
      const id = attr(blip, 'r:embed');
      if (!id) return null;
      const rel = ctx.rels.get(id);
      if (!rel || rel.external) return null;
      return this.image(rel.target);
    }

    /** Fill from a theme style reference (p:style/a:fillRef or bgRef). */
    async styleRefFill(refEl, ctx, list, offset) {
      if (!refEl) return undefined;
      const idx = num(refEl, 'idx', 0);
      if (idx === 0) return { type: 'none' };
      const phClr = this.color(colorChild(refEl), ctx);
      const styleEl = list[idx - offset];
      if (styleEl) {
        const f = await this.fillFrom(styleEl, ctx, phClr);
        if (f) return f;
      }
      return phClr ? { type: 'solid', color: cssColor(phClr) } : undefined;
    }

    line(lnEl, ctx, styleLnRef) {
      let color = null, width = null, none = false;
      if (styleLnRef && num(styleLnRef, 'idx', 0) > 0) {
        color = this.color(colorChild(styleLnRef), ctx);
        const themeLn = ctx.theme.lines[num(styleLnRef, 'idx', 1) - 1];
        width = themeLn ? num(themeLn, 'w', 9525) : 9525;
      }
      let dash = null, head = null, tail = null;
      if (lnEl) {
        if (kid(lnEl, 'noFill')) none = true;
        const solid = kid(lnEl, 'solidFill');
        if (solid) { color = this.color(colorChild(solid), ctx); none = false; }
        const grad = kid(lnEl, 'gradFill');
        if (grad) {
          const first = kid(kid(grad, 'gsLst'), 'gs');
          if (first) { color = this.color(colorChild(first), ctx); none = false; }
        }
        if (attr(lnEl, 'w') != null) width = num(lnEl, 'w', 12700);
        const pd = attr(kid(lnEl, 'prstDash'), 'val');
        if (pd && pd !== 'solid') dash = pd;
        const he = attr(kid(lnEl, 'headEnd'), 'type');
        const te = attr(kid(lnEl, 'tailEnd'), 'type');
        if (he && he !== 'none') head = he;
        if (te && te !== 'none') tail = te;
      }
      if (none || !color) return null;
      if (width == null) width = 12700;
      return { color: cssColor(color), width: Math.max(px(width), 0.75), dash, head, tail };
    }

    // ---- slides --------------------------------------------------------------

    async parseSlide(slidePath, number) {
      const rootEl = await this.xml(slidePath);
      const rels = await this.rels(slidePath);
      const layoutRel = this.relByType(rels, '/slideLayout');
      const layout = layoutRel ? await this.loadLayout(layoutRel.target) : null;
      const master = layout ? layout.master : null;
      const theme = master ? master.theme : { colors: {}, fills: [], lines: [], bgFills: [] };
      const clrMap = this.makeClrMap(master, layout && layout.root, rootEl);

      const mkCtx = (part) => ({ theme, clrMap, rels: part.rels, master, layout, slideNumber: number });
      const slideCtx = Object.assign(mkCtx({ rels }), { kind: 'slide', root: rootEl, path: slidePath });
      const layoutCtx = layout && Object.assign(mkCtx(layout), { kind: 'layout', root: layout.root });
      const masterCtx = master && Object.assign(mkCtx(master), { kind: 'master', root: master.root });

      const background =
        (await this.background(rootEl, slideCtx)) ||
        (layout && (await this.background(layout.root, layoutCtx))) ||
        (master && (await this.background(master.root, masterCtx))) ||
        { type: 'solid', color: '#ffffff' };

      const elements = [];
      const showSlideMasterSp = bool(rootEl, 'showMasterSp') !== false;
      if (master && showSlideMasterSp && (!layout || layout.showMasterSp)) {
        await this.walkTree(master.spTree, masterCtx, null, elements);
      }
      if (layout && showSlideMasterSp) {
        await this.walkTree(layout.spTree, layoutCtx, null, elements);
      }
      await this.walkTree(path(rootEl, 'cSld', 'spTree'), slideCtx, null, elements);

      const notes = await this.notes(rels);
      return { number, hidden: bool(rootEl, 'show') === false, background, elements, notes };
    }

    async background(partRoot, ctx) {
      const bg = path(partRoot, 'cSld', 'bg');
      if (!bg) return null;
      const bgPr = kid(bg, 'bgPr');
      if (bgPr) {
        const f = await this.fill(bgPr, ctx);
        if (f && f.type !== 'none') return f;
        return null;
      }
      const bgRef = kid(bg, 'bgRef');
      if (bgRef) {
        const idx = num(bgRef, 'idx', 0);
        if (idx >= 1001) return this.styleRefFill(bgRef, ctx, ctx.theme.bgFills, 1001);
        if (idx > 0) return this.styleRefFill(bgRef, ctx, ctx.theme.fills, 1);
      }
      return null;
    }

    async notes(slideRels) {
      const rel = this.relByType(slideRels, '/notesSlide');
      if (!rel) return '';
      const rootEl = await this.xml(rel.target);
      const lines = [];
      const walk = (tree) => {
        for (const el of flatKids(tree)) {
          if (el.localName === 'grpSp') { walk(el); continue; }
          if (el.localName !== 'sp') continue;
          const ph = path(el, 'nvSpPr', 'nvPr', 'ph');
          if (!ph || attr(ph, 'type') !== 'body') continue;
          for (const p of kids(kid(el, 'txBody'), 'p')) {
            let line = '';
            for (const r of kids(p)) {
              if (r.localName === 'r' || r.localName === 'fld') line += textOf(kid(r, 't'));
              else if (r.localName === 'br') line += '\n';
            }
            lines.push(line);
          }
        }
      };
      walk(path(rootEl, 'cSld', 'spTree'));
      return lines.join('\n').trim();
    }

    // ---- shape tree ----------------------------------------------------------

    /** xf maps a child rect in EMU to slide space (for groups). null = identity. */
    async walkTree(tree, ctx, xf, out) {
      if (!tree) return;
      for (const el of flatKids(tree)) {
        try {
          switch (el.localName) {
            case 'sp': case 'cxnSp': {
              const shape = await this.shape(el, ctx, xf);
              if (shape) out.push(shape);
              break;
            }
            case 'pic': {
              const pic = await this.picture(el, ctx, xf);
              if (pic) out.push(pic);
              break;
            }
            case 'grpSp': await this.group(el, ctx, xf, out); break;
            case 'graphicFrame': await this.graphicFrame(el, ctx, xf, out); break;
            default: break;
          }
        } catch (err) {
          // Skip shapes we fail to understand rather than losing the whole slide.
          if (typeof console !== 'undefined') console.warn('pptx-viewer: skipped element', err);
        }
      }
    }

    async group(el, ctx, xf, out) {
      const xfrm = path(el, 'grpSpPr', 'xfrm');
      if (!xfrm) return this.walkTree(el, ctx, xf, out);
      const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
      const chOff = kid(xfrm, 'chOff'), chExt = kid(xfrm, 'chExt');
      const ox = num(off, 'x', 0), oy = num(off, 'y', 0);
      const ew = num(ext, 'cx', 0), eh = num(ext, 'cy', 0);
      const cox = num(chOff, 'x', ox), coy = num(chOff, 'y', oy);
      const cw = num(chExt, 'cx', ew) || 1, ch = num(chExt, 'cy', eh) || 1;
      const sx = ew / cw, sy = eh / ch;
      const local = (r) => ({ x: ox + (r.x - cox) * sx, y: oy + (r.y - coy) * sy, w: r.w * sx, h: r.h * sy });
      const childXf = xf ? (r) => xf(local(r)) : local;
      const groupFill = await this.fill(kid(el, 'grpSpPr'), ctx);
      await this.walkTree(el, Object.assign({}, ctx, { groupFill }), childXf, out);
    }

    readXfrm(xfrmEl) {
      if (!xfrmEl) return null;
      const off = kid(xfrmEl, 'off'), ext = kid(xfrmEl, 'ext');
      if (!off && !ext) return null;
      return {
        x: num(off, 'x', 0), y: num(off, 'y', 0), w: num(ext, 'cx', 0), h: num(ext, 'cy', 0),
        rot: num(xfrmEl, 'rot', 0) / 60000,
        flipH: bool(xfrmEl, 'flipH') === true,
        flipV: bool(xfrmEl, 'flipV') === true,
      };
    }

    placeholderOf(el) {
      const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvGraphicFramePr') || kid(el, 'nvCxnSpPr');
      const ph = path(nv, 'nvPr', 'ph');
      if (!ph) return null;
      return { type: attr(ph, 'type') || 'obj', idx: attr(ph, 'idx') };
    }

    findPlaceholder(list, ph, byIdx) {
      if (!list || !ph) return null;
      if (byIdx && ph.idx != null) {
        const m = list.find((p) => p.idx === ph.idx);
        if (m) return m.el;
      }
      const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
      const exact = list.find((p) => p.type === ph.type);
      if (exact) return exact.el;
      const m = list.find((p) => norm(p.type) === norm(ph.type));
      return m ? m.el : null;
    }

    /** Returns [shape, layoutPlaceholder, masterPlaceholder] (missing entries omitted). */
    inheritanceChain(el, ctx) {
      const chain = [el];
      const ph = this.placeholderOf(el);
      if (!ph) return { chain, ph };
      if (ctx.kind === 'slide' && ctx.layout) {
        const l = this.findPlaceholder(ctx.layout.placeholders, ph, true);
        if (l) chain.push(l);
        const m = this.findPlaceholder(ctx.master && ctx.master.placeholders, ph, false);
        if (m) chain.push(m);
      } else if (ctx.kind === 'layout' && ctx.master) {
        const m = this.findPlaceholder(ctx.master.placeholders, ph, false);
        if (m) chain.push(m);
      }
      return { chain, ph };
    }

    toBox(rect, xf) {
      const r = xf ? xf(rect) : rect;
      return { x: px(r.x), y: px(r.y), w: px(r.w), h: px(r.h) };
    }

    async shape(el, ctx, xf) {
      const { chain, ph } = this.inheritanceChain(el, ctx);
      // Placeholders on layouts/masters are prompts ("Click to add title"), not content.
      if (ph && ctx.kind !== 'slide') return null;

      let xfrm = null;
      for (const c of chain) {
        xfrm = this.readXfrm(path(c, 'spPr', 'xfrm'));
        if (xfrm) break;
      }
      if (!xfrm) return null;
      const box = this.toBox(xfrm, xf);

      const style = kid(el, 'style');
      let geomName = 'rect', d;
      let geomEl = null;
      for (const c of chain) {
        geomEl = kid(kid(c, 'spPr'), 'prstGeom') || kid(kid(c, 'spPr'), 'custGeom');
        if (geomEl) break;
      }
      if (geomEl && geomEl.localName === 'custGeom') {
        geomName = 'custom';
        d = customPath(geomEl, box.w, box.h);
      } else {
        geomName = attr(geomEl, 'prst') || 'rect';
        d = presetPath(geomName, box.w, box.h, geomEl ? adjValues(geomEl) : {});
      }
      const isLine = el.localName === 'cxnSp' || LINE_PRESETS.test(geomName);

      let fill;
      for (const c of chain) {
        fill = await this.fill(kid(c, 'spPr'), ctx);
        if (fill) break;
      }
      if (!fill && style) fill = await this.styleRefFill(kid(style, 'fillRef'), ctx, ctx.theme.fills, 1);
      if (isLine) fill = { type: 'none' };

      let lnEl = null;
      for (const c of chain) {
        lnEl = kid(kid(c, 'spPr'), 'ln');
        if (lnEl) break;
      }
      const line = this.line(lnEl, ctx, style && kid(style, 'lnRef'));

      const text = await this.textBody(el, chain, ph, ctx);

      if ((!fill || fill.type === 'none') && !line && !text) return null;
      return {
        type: 'shape',
        ...box,
        rot: xfrm.rot, flipH: xfrm.flipH, flipV: xfrm.flipV,
        geom: geomName, path: d, isLine,
        fill: fill && fill.type !== 'none' ? fill : null,
        line,
        text,
      };
    }

    async picture(el, ctx, xf) {
      const { chain, ph } = this.inheritanceChain(el, ctx);
      if (ph && ctx.kind !== 'slide') return null;
      let xfrm = null;
      for (const c of chain) {
        xfrm = this.readXfrm(path(c, 'spPr', 'xfrm'));
        if (xfrm) break;
      }
      if (!xfrm) return null;
      const blipFill = kid(el, 'blipFill');
      const url = await this.blipUrl(blipFill, ctx);
      const blip = kid(blipFill, 'blip');
      const rel = blip && ctx.rels.get(attr(blip, 'r:embed'));
      const src = kid(blipFill, 'srcRect');
      const crop = src ? {
        l: num(src, 'l', 0) / 100000, t: num(src, 't', 0) / 100000,
        r: num(src, 'r', 0) / 100000, b: num(src, 'b', 0) / 100000,
      } : null;
      const geomEl = kid(kid(el, 'spPr'), 'prstGeom');
      const box = this.toBox(xfrm, xf);
      const geomName = attr(geomEl, 'prst') || 'rect';
      return {
        type: 'image',
        ...box,
        rot: xfrm.rot, flipH: xfrm.flipH, flipV: xfrm.flipV,
        url,
        unsupported: url ? null : (rel ? rel.target.split('/').pop() : 'image'),
        crop,
        clipPath: geomName !== 'rect' ? presetPath(geomName, box.w, box.h, adjValues(geomEl)) : null,
        line: this.line(kid(kid(el, 'spPr'), 'ln'), ctx, null),
      };
    }

    async graphicFrame(el, ctx, xf, out) {
      if (this.placeholderOf(el) && ctx.kind !== 'slide') return;
      const xfrm = this.readXfrm(kid(el, 'xfrm'));
      if (!xfrm) return;
      const box = this.toBox(xfrm, xf);
      const gd = path(el, 'graphic', 'graphicData');
      const uri = attr(gd, 'uri') || '';
      const tbl = kid(gd, 'tbl');
      if (tbl) {
        out.push(await this.table(tbl, ctx, box));
        return;
      }
      if (uri.endsWith('/diagram')) {
        const drawn = await this.smartArt(gd, ctx, xfrm, xf, out);
        if (drawn) return;
        out.push({ type: 'placeholder', ...box, label: 'SmartArt' });
        return;
      }
      if (uri.endsWith('/chart')) {
        out.push({ type: 'placeholder', ...box, label: await this.chartLabel(gd, ctx) });
        return;
      }
      out.push({ type: 'placeholder', ...box, label: 'Embedded object' });
    }

    async chartLabel(gd, ctx) {
      const id = attr(kid(gd, 'chart'), 'r:id');
      const rel = id && ctx.rels.get(id);
      if (!rel) return 'Chart';
      const rootEl = await this.xml(rel.target);
      const chart = kid(rootEl, 'chart');
      const titleText = [];
      const collect = (n) => {
        for (const c of kids(n)) {
          if (c.localName === 't') titleText.push(textOf(c));
          else collect(c);
        }
      };
      collect(kid(chart, 'title'));
      const plot = kid(chart, 'plotArea');
      const kind = kids(plot).map((c) => c.localName).find((n) => /Chart$/.test(n));
      const pretty = kind ? kind.replace(/3D/, ' 3D').replace(/Chart$/, '') : '';
      const title = titleText.join('').trim();
      return `Chart${pretty ? ` (${pretty})` : ''}${title ? `: ${title}` : ''}`;
    }

    async smartArt(gd, ctx, frameXfrm, xf, out) {
      // SmartArt ships a pre-rendered drawing part; find it via the data part's dataModelExt.
      const relIds = kid(gd, 'relIds');
      const dmRel = relIds && ctx.rels.get(attr(relIds, 'r:dm'));
      let drawingPath = null;
      if (dmRel) {
        const data = await this.xml(dmRel.target);
        let relId = null;
        const find = (n) => {
          for (const c of kids(n)) {
            if (relId) return;
            if (c.localName === 'dataModelExt') relId = attr(c, 'relId');
            else find(c);
          }
        };
        find(data);
        const r = relId && ctx.rels.get(relId);
        if (r) drawingPath = r.target;
      }
      if (!drawingPath) {
        const drawings = Array.from(ctx.rels.values()).filter((r) => r.type.endsWith('/diagramDrawing'));
        if (drawings.length === 1) drawingPath = drawings[0].target;
      }
      if (!drawingPath) return false;
      const drawing = await this.xml(drawingPath);
      const tree = kid(drawing, 'spTree');
      if (!tree) return false;
      // Drawing coordinates are relative to the graphic frame.
      const local = (r) => ({ x: r.x + frameXfrm.x, y: r.y + frameXfrm.y, w: r.w, h: r.h });
      const drawingXf = xf ? (r) => xf(local(r)) : local;
      const drawingCtx = Object.assign({}, ctx, { rels: await this.rels(drawingPath), kind: 'drawing' });
      const before = out.length;
      await this.walkTree(tree, drawingCtx, drawingXf, out);
      return out.length > before;
    }

    // ---- text ------------------------------------------------------------------

    masterStyleFor(ph, ctx) {
      const m = ctx.master;
      if (!m) return null;
      if (!ph) return m.otherStyle;
      if (ph.type === 'title' || ph.type === 'ctrTitle') return m.titleStyle;
      if (ph.type === 'body' || ph.type === 'subTitle' || ph.type === 'obj') return m.bodyStyle;
      return m.otherStyle;
    }

    fontFace(typeface, ctx) {
      if (!typeface) return null;
      if (typeface === '+mj-lt' || typeface === '+mj-ea' || typeface === '+mj-cs') return ctx.theme.majorFont;
      if (typeface === '+mn-lt' || typeface === '+mn-ea' || typeface === '+mn-cs') return ctx.theme.minorFont;
      return typeface;
    }

    mergeRPr(t, rPr, ctx) {
      if (!rPr) return;
      if (attr(rPr, 'sz') != null) t.size = num(rPr, 'sz', 1800) / 100;
      const b = bool(rPr, 'b'); if (b !== undefined) t.bold = b;
      const i = bool(rPr, 'i'); if (i !== undefined) t.italic = i;
      const u = attr(rPr, 'u'); if (u != null) t.underline = u !== 'none';
      const s = attr(rPr, 'strike'); if (s != null) t.strike = s !== 'noStrike';
      if (attr(rPr, 'baseline') != null) t.baseline = num(rPr, 'baseline', 0) / 1000;
      const cap = attr(rPr, 'cap'); if (cap != null) t.cap = cap;
      if (attr(rPr, 'spc') != null) t.spacing = num(rPr, 'spc', 0) / 100;
      const solid = kid(rPr, 'solidFill');
      if (solid) t.color = this.color(colorChild(solid), ctx);
      const grad = kid(rPr, 'gradFill');
      if (grad) {
        const first = kid(kid(grad, 'gsLst'), 'gs');
        if (first) t.color = this.color(colorChild(first), ctx);
      }
      if (kid(rPr, 'noFill')) t.color = { r: 0, g: 0, b: 0, a: 0 };
      const latin = attr(kid(rPr, 'latin'), 'typeface');
      if (latin) t.font = this.fontFace(latin, ctx);
      const hl = kid(rPr, 'highlight');
      if (hl) t.highlight = this.color(colorChild(hl), ctx);
      if (kid(rPr, 'hlinkClick')) { t.link = true; }
    }

    mergePPr(t, pPr, ctx) {
      if (!pPr) return;
      const algn = attr(pPr, 'algn'); if (algn) t.align = algn;
      if (attr(pPr, 'marL') != null) t.marL = num(pPr, 'marL', 0);
      if (attr(pPr, 'indent') != null) t.indent = num(pPr, 'indent', 0);
      for (const c of kids(pPr)) {
        switch (c.localName) {
          case 'buNone': t.bullet = { type: 'none' }; break;
          case 'buChar': t.bullet = { type: 'char', char: attr(c, 'char') || '•' }; break;
          case 'buAutoNum': t.bullet = { type: 'num', scheme: attr(c, 'type'), start: num(c, 'startAt', 1) }; break;
          case 'buBlip': t.bullet = { type: 'char', char: '•' }; break;
          case 'buClr': t.buColor = this.color(colorChild(c), ctx); break;
          case 'buClrTx': t.buColor = null; break;
          case 'buSzPct': t.buSize = num(c, 'val', 100000) / 100000; break;
          case 'buFont': t.buFont = this.fontFace(attr(c, 'typeface'), ctx); break;
          case 'lnSpc': {
            const pct = kid(c, 'spcPct'), pts = kid(c, 'spcPts');
            if (pct) t.lnSpc = { pct: num(pct, 'val', 100000) / 100000 };
            else if (pts) t.lnSpc = { pt: num(pts, 'val', 0) / 100 };
            break;
          }
          case 'spcBef': case 'spcAft': {
            const pct = kid(c, 'spcPct'), pts = kid(c, 'spcPts');
            const v = pct ? { pct: num(pct, 'val', 0) / 100000 } : pts ? { pt: num(pts, 'val', 0) / 100 } : null;
            if (c.localName === 'spcBef') t.spcBef = v; else t.spcAft = v;
            break;
          }
          case 'defRPr': this.mergeRPr(t.rPr, c, ctx); break;
          default: break;
        }
      }
    }

    async textBody(el, chain, ph, ctx) {
      const txBody = kid(el, 'txBody');
      if (!txBody) return null;
      const paragraphsEl = kids(txBody, 'p');
      const hasText = paragraphsEl.some((p) => kids(p).some((r) => (r.localName === 'r' || r.localName === 'fld') && textOf(kid(r, 't')).length));
      if (!hasText) return null;

      // Body properties: master -> layout -> shape (later wins).
      const body = { anchor: 't', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, wrap: true, vert: null, fontScale: 1, lnReduce: 0, anchorCtr: false };
      for (const c of chain.slice().reverse()) {
        const bp = path(c, 'txBody', 'bodyPr');
        if (!bp) continue;
        const anchor = attr(bp, 'anchor'); if (anchor) body.anchor = anchor;
        for (const k of ['lIns', 'tIns', 'rIns', 'bIns']) if (attr(bp, k) != null) body[k] = num(bp, k, body[k]);
        if (attr(bp, 'wrap') != null) body.wrap = attr(bp, 'wrap') !== 'none';
        if (attr(bp, 'vert') != null) body.vert = attr(bp, 'vert') === 'horz' ? null : attr(bp, 'vert');
        if (attr(bp, 'anchorCtr') != null) body.anchorCtr = bool(bp, 'anchorCtr') === true;
        const auto = kid(bp, 'normAutofit');
        if (auto) {
          body.fontScale = num(auto, 'fontScale', 100000) / 100000;
          body.lnReduce = num(auto, 'lnSpcReduction', 0) / 100000;
        }
      }

      // Text style layers, lowest priority first.
      const isTitle = ph && (ph.type === 'title' || ph.type === 'ctrTitle');
      const layers = [this.defaultTextStyle, this.masterStyleFor(ph, ctx)];
      for (const c of chain.slice(1).reverse()) layers.push(path(c, 'txBody', 'lstStyle'));
      const ownLst = kid(txBody, 'lstStyle');

      const style = kid(el, 'style');
      const fontRef = style && kid(style, 'fontRef');
      const fontRefColor = fontRef ? this.color(colorChild(fontRef), ctx) : null;
      const fontRefFace = fontRef ? (attr(fontRef, 'idx') === 'major' ? ctx.theme.majorFont : attr(fontRef, 'idx') === 'minor' ? ctx.theme.minorFont : null) : null;

      const counters = [];
      const paragraphs = [];
      for (const p of paragraphsEl) {
        const pPr = kid(p, 'pPr');
        const lvl = Math.min(num(pPr, 'lvl', 0), 8);
        const lvlName = `lvl${lvl + 1}pPr`;
        const t = { align: 'l', marL: 0, indent: 0, bullet: null, rPr: { size: DEFAULT_FONT_PT, color: null, font: null } };
        for (const layer of layers) this.mergePPr(t, kid(layer, lvlName), ctx);
        if (fontRefColor) t.rPr.color = fontRefColor;
        if (fontRefFace) t.rPr.font = fontRefFace;
        this.mergePPr(t, kid(ownLst, lvlName), ctx);
        this.mergePPr(t, pPr, ctx);

        if (!t.rPr.color) t.rPr.color = this.themeColor('tx1', ctx);
        if (!t.rPr.font) t.rPr.font = isTitle ? ctx.theme.majorFont : ctx.theme.minorFont;

        const runs = [];
        for (const r of kids(p)) {
          if (r.localName === 'r' || r.localName === 'fld') {
            const rs = Object.assign({}, t.rPr);
            this.mergeRPr(rs, kid(r, 'rPr'), ctx);
            let txt = textOf(kid(r, 't'));
            if (r.localName === 'fld' && attr(r, 'type') === 'slidenum' && ctx.slideNumber) txt = String(ctx.slideNumber);
            if (rs.link && !(kid(kid(r, 'rPr'), 'solidFill'))) {
              const hl = ctx.theme.colors.hlink;
              if (hl) rs.color = hexToRgb(hl);
              rs.underline = true;
            }
            runs.push(this.runOut(txt, rs, body));
          } else if (r.localName === 'br') {
            runs.push({ br: true });
          }
        }
        const endRPr = Object.assign({}, t.rPr);
        this.mergeRPr(endRPr, kid(p, 'endParaRPr'), ctx);
        const firstRun = runs.find((r) => !r.br);
        const hasRunText = runs.some((r) => !r.br && r.text.length);

        // Bullets
        let bullet = null;
        if (t.bullet && t.bullet.type !== 'none' && hasRunText) {
          let text = t.bullet.char;
          if (t.bullet.type === 'num') {
            counters.length = lvl + 1;
            counters[lvl] = (counters[lvl] || t.bullet.start - 1) + 1;
            text = formatAutoNum(t.bullet.scheme, counters[lvl]);
          }
          const baseSize = firstRun ? firstRun.size : endRPr.size * PT_TO_PX * body.fontScale;
          bullet = {
            text,
            color: cssColor(t.buColor || (firstRun ? null : endRPr.color)) || (firstRun ? firstRun.color : null),
            font: t.buFont || null,
            size: baseSize * (t.buSize || 1),
          };
        } else if (t.bullet && t.bullet.type !== 'num') {
          counters.length = lvl;
        }

        const sizeRef = (firstRun ? firstRun.size : endRPr.size * PT_TO_PX * body.fontScale);
        const spc = (v) => (!v ? 0 : v.pt != null ? v.pt * PT_TO_PX : v.pct * sizeRef * 1.2);
        let lineHeight = null;
        if (t.lnSpc && t.lnSpc.pct != null) lineHeight = { factor: Math.max(t.lnSpc.pct * (1 - body.lnReduce), 0.5) * 1.2 };
        else if (t.lnSpc && t.lnSpc.pt != null) lineHeight = { px: t.lnSpc.pt * PT_TO_PX };
        else if (body.lnReduce) lineHeight = { factor: (1 - body.lnReduce) * 1.2 };

        paragraphs.push({
          align: t.align,
          marL: px(t.marL),
          indent: px(t.indent),
          spaceBefore: spc(t.spcBef),
          spaceAfter: spc(t.spcAft),
          lineHeight,
          bullet,
          runs,
          endSize: endRPr.size * PT_TO_PX * body.fontScale,
        });
      }

      return {
        anchor: body.anchor,
        anchorCtr: body.anchorCtr,
        insets: { l: px(body.lIns), t: px(body.tIns), r: px(body.rIns), b: px(body.bIns) },
        wrap: body.wrap,
        vert: body.vert,
        paragraphs,
      };
    }

    runOut(text, rs, body) {
      let t = text;
      if (rs.cap === 'all') t = t.toUpperCase();
      return {
        text: t,
        size: rs.size * PT_TO_PX * (body ? body.fontScale : 1),
        bold: !!rs.bold,
        italic: !!rs.italic,
        underline: !!rs.underline,
        strike: !!rs.strike,
        smallCaps: rs.cap === 'small',
        baseline: rs.baseline || 0,
        spacing: rs.spacing ? rs.spacing * PT_TO_PX : 0,
        color: cssColor(rs.color),
        highlight: rs.highlight ? cssColor(rs.highlight) : null,
        font: rs.font || null,
      };
    }

    // ---- tables ------------------------------------------------------------------

    async loadTableStyle(id) {
      if (!this.tableStyles) {
        this.tableStyles = new Map();
        const rootEl = await this.xml('ppt/tableStyles.xml');
        for (const s of kids(rootEl, 'tblStyle')) this.tableStyles.set(attr(s, 'styleId'), s);
      }
      return id ? this.tableStyles.get(id) || null : null;
    }

    async tableStylePart(partEl, ctx) {
      if (!partEl) return null;
      const out = {};
      const tx = kid(partEl, 'tcTxStyle');
      if (tx) {
        const b = attr(tx, 'b');
        if (b) out.bold = b === 'on';
        const i = attr(tx, 'i');
        if (i) out.italic = i === 'on';
        const c = colorChild(tx) || colorChild(kid(tx, 'fontRef'));
        if (c) out.color = this.color(c, ctx);
      }
      const tc = kid(partEl, 'tcStyle');
      if (tc) {
        const f = await this.fill(kid(tc, 'fill'), ctx);
        if (f) out.fill = f;
        else if (kid(tc, 'fillRef')) out.fill = await this.styleRefFill(kid(tc, 'fillRef'), ctx, ctx.theme.fills, 1);
        const bdr = kid(tc, 'tcBdr');
        if (bdr) {
          out.borders = {};
          for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
            const s = kid(bdr, side);
            if (!s) continue;
            const ln = kid(s, 'ln');
            out.borders[side] = ln ? this.line(ln, ctx, null) : this.line(null, ctx, kid(s, 'lnRef'));
          }
        }
      }
      return out;
    }

    async table(tbl, ctx, box) {
      const tblPr = kid(tbl, 'tblPr');
      const flags = {
        firstRow: bool(tblPr, 'firstRow') === true, lastRow: bool(tblPr, 'lastRow') === true,
        firstCol: bool(tblPr, 'firstCol') === true, lastCol: bool(tblPr, 'lastCol') === true,
        bandRow: bool(tblPr, 'bandRow') === true, bandCol: bool(tblPr, 'bandCol') === true,
      };
      const styleId = textOf(kid(tblPr, 'tableStyleId')).trim();
      const styleEl = await this.loadTableStyle(styleId);
      const sp = {};
      if (styleEl) {
        for (const part of ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'firstRow', 'lastRow', 'firstCol', 'lastCol']) {
          sp[part] = await this.tableStylePart(kid(styleEl, part), ctx);
        }
      } else if (styleId || flags.firstRow || flags.bandRow) {
        // Approximate PowerPoint's default "Medium Style 2 - Accent 1".
        const accent = hexToRgb(ctx.theme.colors.accent1 || '4472C4');
        const tint = (v) => ({ r: accent.r + (255 - accent.r) * v, g: accent.g + (255 - accent.g) * v, b: accent.b + (255 - accent.b) * v, a: 1 });
        const white = { color: 'rgb(255, 255, 255)', width: 1 };
        sp.wholeTbl = { fill: { type: 'solid', color: cssColor(tint(0.8)) }, borders: { left: white, right: white, top: white, bottom: white, insideH: white, insideV: white } };
        sp.band1H = { fill: { type: 'solid', color: cssColor(tint(0.6)) } };
        sp.firstRow = { fill: { type: 'solid', color: cssColor(accent) }, bold: true, color: { r: 255, g: 255, b: 255, a: 1 } };
        sp.lastRow = { fill: { type: 'solid', color: cssColor(accent) }, bold: true, color: { r: 255, g: 255, b: 255, a: 1 } };
        sp.firstCol = { bold: true };
        sp.lastCol = { bold: true };
      }

      const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => px(num(g, 'w', 0)));
      const rowEls = kids(tbl, 'tr');
      const nRows = rowEls.length, nCols = cols.length;
      const rows = [];
      for (let ri = 0; ri < nRows; ri++) {
        const tr = rowEls[ri];
        const cells = [];
        const tcs = kids(tr, 'tc');
        for (let ci = 0; ci < tcs.length; ci++) {
          const tc = tcs[ci];
          if (bool(tc, 'hMerge') || bool(tc, 'vMerge')) { cells.push({ skip: true }); continue; }

          // Which style parts apply to this cell (later = higher priority)
          const parts = ['wholeTbl'];
          const dataRow = ri - (flags.firstRow ? 1 : 0);
          if (flags.bandRow && dataRow >= 0) parts.push(dataRow % 2 === 0 ? 'band1H' : 'band2H');
          if (flags.bandCol) parts.push(ci % 2 === 0 ? 'band1V' : 'band2V');
          if (flags.firstCol && ci === 0) parts.push('firstCol');
          if (flags.lastCol && ci === nCols - 1) parts.push('lastCol');
          if (flags.lastRow && ri === nRows - 1) parts.push('lastRow');
          if (flags.firstRow && ri === 0) parts.push('firstRow');
          const cs = { fill: null, bold: false, italic: false, color: null, borders: {} };
          for (const name of parts) {
            const part = sp[name];
            if (!part) continue;
            if (part.fill) cs.fill = part.fill;
            if (part.bold !== undefined) cs.bold = part.bold;
            if (part.italic !== undefined) cs.italic = part.italic;
            if (part.color) cs.color = part.color;
            if (part.borders) Object.assign(cs.borders, part.borders);
          }
          const pick = (outer, inner, isEdge) => (isEdge ? cs.borders[outer] : cs.borders[inner] || null) || null;
          const span = num(tc, 'gridSpan', 1), rspan = num(tc, 'rowSpan', 1);
          const borders = {
            l: pick('left', 'insideV', ci === 0),
            r: pick('right', 'insideV', ci + span >= nCols),
            t: pick('top', 'insideH', ri === 0),
            b: pick('bottom', 'insideH', ri + rspan >= nRows),
          };

          const tcPr = kid(tc, 'tcPr');
          const ownFill = await this.fill(tcPr, ctx);
          if (ownFill) cs.fill = ownFill;
          for (const [k, tag] of [['l', 'lnL'], ['r', 'lnR'], ['t', 'lnT'], ['b', 'lnB']]) {
            const ln = kid(tcPr, tag);
            if (ln) borders[k] = this.line(ln, ctx, null);
          }

          const text = await this.cellText(tc, ctx, cs);
          cells.push({
            colspan: span, rowspan: rspan,
            fill: cs.fill && cs.fill.type !== 'none' ? cs.fill : null,
            borders,
            anchor: attr(tcPr, 'anchor') || 't',
            padding: {
              l: px(num(tcPr, 'marL', 91440)), r: px(num(tcPr, 'marR', 91440)),
              t: px(num(tcPr, 'marT', 45720)), b: px(num(tcPr, 'marB', 45720)),
            },
            paragraphs: text,
          });
        }
        rows.push({ h: px(num(tr, 'h', 0)), cells });
      }
      return { type: 'table', ...box, cols, rows };
    }

    async cellText(tc, ctx, cs) {
      const txBody = kid(tc, 'txBody');
      const out = [];
      for (const p of kids(txBody, 'p')) {
        const t = { align: 'l', marL: 0, indent: 0, bullet: null, rPr: { size: DEFAULT_FONT_PT, color: null, font: null } };
        this.mergePPr(t, kid(this.defaultTextStyle, 'lvl1pPr'), ctx);
        if (cs.color) t.rPr.color = cs.color;
        if (cs.bold) t.rPr.bold = true;
        if (cs.italic) t.rPr.italic = true;
        this.mergePPr(t, kid(kid(txBody, 'lstStyle'), 'lvl1pPr'), ctx);
        this.mergePPr(t, kid(p, 'pPr'), ctx);
        if (!t.rPr.color) t.rPr.color = this.themeColor('tx1', ctx);
        if (!t.rPr.font) t.rPr.font = ctx.theme.minorFont;
        const runs = [];
        for (const r of kids(p)) {
          if (r.localName === 'r' || r.localName === 'fld') {
            const rs = Object.assign({}, t.rPr);
            this.mergeRPr(rs, kid(r, 'rPr'), ctx);
            runs.push(this.runOut(textOf(kid(r, 't')), rs, null));
          } else if (r.localName === 'br') {
            runs.push({ br: true });
          }
        }
        const endRPr = Object.assign({}, t.rPr);
        this.mergeRPr(endRPr, kid(p, 'endParaRPr'), ctx);
        out.push({
          align: t.align, marL: px(t.marL), indent: px(t.indent), spaceBefore: 0, spaceAfter: 0,
          lineHeight: null, bullet: null, runs, endSize: endRPr.size * PT_TO_PX,
        });
      }
      return out;
    }
  }

  async function parsePptx(zip, opts) {
    return new Parser(zip, opts || {}).parse();
  }

  return { parsePptx };
});
