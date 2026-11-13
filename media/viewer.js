/* global acquireVsCodeApi, JSZip, PptxParser */
(function () {
  'use strict';

  const vscode = acquireVsCodeApi();
  const $ = (id) => document.getElementById(id);
  const sidebar = $('sidebar');
  const stage = $('stage');
  const notesEl = $('notes');
  const notesInput = $('notes-input');
  const counter = $('counter');

  const saved = vscode.getState() || {};
  const state = {
    model: null,
    index: saved.index || 0,
    showNotes: saved.showNotes !== false,
    blobUrls: [],
  };

  // ---------------------------------------------------------------------------
  // Loading
  // ---------------------------------------------------------------------------

  window.addEventListener('message', async (event) => {
    const msg = event.data;
    if (msg.type === 'load') await load(msg.data, msg.notes);
    else if (msg.type === 'notes') applyNotes(msg.notes);
    else if (msg.type === 'error') showStatus(`Could not read file: ${msg.message}`, true);
  });

  async function load(data, notes) {
    try {
      if (!state.model) showStatus('Loading…');
      const zip = await JSZip.loadAsync(data, { base64: typeof data === 'string' });
      const urls = [];
      const model = await PptxParser.parsePptx(zip, {
        DOMParser,
        makeImageUrl: (bytes, mime) => {
          const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
          urls.push(url);
          return url;
        },
      });
      state.blobUrls.forEach((u) => URL.revokeObjectURL(u));
      state.blobUrls = urls;
      state.model = model;
      cancelPendingNotes();
      applyNotes(notes, true);
      if (!model.slides.length) {
        showStatus('This presentation has no slides.');
        sidebar.replaceChildren();
        return;
      }
      state.index = Math.min(state.index, model.slides.length - 1);
      buildSidebar();
      show(state.index);
    } catch (err) {
      console.error(err);
      showStatus(`Could not open presentation: ${err && err.message ? err.message : err}`, true);
    }
  }

  function showStatus(text, isError) {
    const div = document.createElement('div');
    div.id = 'status';
    div.textContent = text;
    if (isError) div.classList.add('error');
    stage.replaceChildren(div);
  }

  // ---------------------------------------------------------------------------
  // Navigation & layout
  // ---------------------------------------------------------------------------

  function show(i) {
    const { model } = state;
    if (!model) return;
    flushNotes();
    state.index = Math.max(0, Math.min(i, model.slides.length - 1));
    vscode.setState({ index: state.index, showNotes: state.showNotes });

    const slide = model.slides[state.index];
    const frame = document.createElement('div');
    frame.className = 'slide-frame';
    frame.appendChild(renderSlide(slide, model));
    stage.replaceChildren(frame);
    fitMain();

    counter.textContent = `${state.index + 1} / ${model.slides.length}${slide.hidden ? ' (hidden)' : ''}`;
    $('prev').disabled = state.index === 0;
    $('next').disabled = state.index === model.slides.length - 1;

    for (const t of sidebar.children) t.classList.toggle('active', Number(t.dataset.index) === state.index);
    const active = sidebar.children[state.index];
    if (active) active.scrollIntoView({ block: 'nearest' });

    renderNotes(slide);
  }

  function fitMain() {
    const { model } = state;
    const frame = stage.querySelector('.slide-frame');
    if (!model || !frame) return;
    const pad = 24;
    const availW = Math.max(stage.clientWidth - pad * 2, 50);
    const availH = Math.max(stage.clientHeight - pad * 2, 50);
    const scale = Math.min(availW / model.width, availH / model.height);
    frame.style.width = `${model.width * scale}px`;
    frame.style.height = `${model.height * scale}px`;
    frame.firstChild.style.transform = `scale(${scale})`;
  }

  function renderNotes(slide) {
    notesEl.classList.toggle('hidden', !state.showNotes);
    $('toggle-notes').classList.toggle('on', state.showNotes);
    if (notesInput.value !== slide.notes) notesInput.value = slide.notes;
    requestAnimationFrame(fitMain);
  }

  // ---------------------------------------------------------------------------
  // Notes editing. Edits are batched while typing and sent to the extension, which
  // owns the undo stack and saving. slide.notes always holds the last sent text.
  // ---------------------------------------------------------------------------

  let notesTimer = null;
  let pendingIndex = -1;

  notesInput.addEventListener('input', () => {
    pendingIndex = state.index;
    clearTimeout(notesTimer);
    notesTimer = setTimeout(flushNotes, 500);
  });
  notesInput.addEventListener('blur', flushNotes);

  function flushNotes() {
    clearTimeout(notesTimer);
    notesTimer = null;
    const { model } = state;
    if (pendingIndex < 0 || !model) return;
    const slide = model.slides[pendingIndex];
    pendingIndex = -1;
    if (!slide) return;
    const after = notesInput.value;
    if (after === slide.notes) return;
    vscode.postMessage({ type: 'editNotes', index: slide.number - 1, before: slide.notes, after });
    slide.notes = after;
  }

  function cancelPendingNotes() {
    clearTimeout(notesTimer);
    notesTimer = null;
    pendingIndex = -1;
  }

  /** Applies notes text pushed by the extension (unsaved edits, undo/redo). */
  function applyNotes(notes, silent) {
    const { model } = state;
    if (!model || !notes) return;
    for (const [key, text] of Object.entries(notes)) {
      const slide = model.slides[Number(key)];
      if (!slide) continue;
      slide.notes = text;
      if (!silent && Number(key) === state.index) {
        cancelPendingNotes();
        if (notesInput.value !== text) notesInput.value = text;
      }
    }
  }

  function buildSidebar() {
    const { model } = state;
    const thumbW = 168;
    const scale = thumbW / model.width;
    const observer = new IntersectionObserver((entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue;
        const box = e.target.querySelector('.thumb-box');
        if (box.childElementCount) continue;
        const slideEl = renderSlide(model.slides[Number(e.target.dataset.index)], model);
        slideEl.style.transform = `scale(${scale})`;
        box.appendChild(slideEl);
        observer.unobserve(e.target);
      }
    }, { root: sidebar, rootMargin: '400px 0px' });

    const items = model.slides.map((slide, i) => {
      const item = document.createElement('button');
      item.className = 'thumb' + (slide.hidden ? ' hidden-slide' : '');
      item.dataset.index = String(i);
      item.title = `Slide ${i + 1}${slide.hidden ? ' (hidden)' : ''}`;
      const numEl = document.createElement('span');
      numEl.className = 'thumb-num';
      numEl.textContent = String(i + 1);
      const box = document.createElement('div');
      box.className = 'thumb-box';
      box.style.width = `${thumbW}px`;
      box.style.height = `${model.height * scale}px`;
      item.append(numEl, box);
      item.addEventListener('click', () => show(i));
      observer.observe(item);
      return item;
    });
    sidebar.replaceChildren(...items);
  }

  $('prev').addEventListener('click', () => show(state.index - 1));
  $('next').addEventListener('click', () => show(state.index + 1));
  $('toggle-notes').addEventListener('click', toggleNotes);

  function toggleNotes() {
    state.showNotes = !state.showNotes;
    if (state.model) show(state.index);
  }

  window.addEventListener('keydown', (e) => {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    const tag = e.target && e.target.tagName;
    if (tag === 'TEXTAREA' || tag === 'INPUT') {
      if (e.key === 'Escape') e.target.blur();
      return;
    }
    switch (e.key) {
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ':
        show(state.index + 1); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp':
        show(state.index - 1); break;
      case 'Home': show(0); break;
      case 'End': if (state.model) show(state.model.slides.length - 1); break;
      case 'n': case 'N': toggleNotes(); break;
      default: return;
    }
    e.preventDefault();
  });

  new ResizeObserver(fitMain).observe(stage);

  // ---------------------------------------------------------------------------
  // Slide rendering
  // ---------------------------------------------------------------------------

  let markerSeq = 0;

  function renderSlide(slide, model) {
    const el = document.createElement('div');
    el.className = 'slide';
    el.style.width = `${model.width}px`;
    el.style.height = `${model.height}px`;
    applyFill(el, slide.background);
    for (const item of slide.elements) {
      let node = null;
      if (item.type === 'shape') node = renderShape(item);
      else if (item.type === 'image') node = renderImage(item);
      else if (item.type === 'table') node = renderTable(item);
      else if (item.type === 'placeholder') node = renderPlaceholder(item);
      if (node) el.appendChild(node);
    }
    if (slide.error) {
      const err = document.createElement('div');
      err.className = 'slide-error';
      err.textContent = `This slide could not be fully rendered: ${slide.error}`;
      el.appendChild(err);
    }
    return el;
  }

  function positioned(item, className) {
    const el = document.createElement('div');
    el.className = `el ${className}`;
    el.style.left = `${item.x}px`;
    el.style.top = `${item.y}px`;
    el.style.width = `${item.w}px`;
    el.style.height = `${item.h}px`;
    if (item.rot) el.style.transform = `rotate(${item.rot}deg)`;
    return el;
  }

  function flipTransform(item) {
    if (!item.flipH && !item.flipV) return '';
    return `scale(${item.flipH ? -1 : 1}, ${item.flipV ? -1 : 1})`;
  }

  function applyFill(el, fill) {
    if (!fill) return;
    if (fill.type === 'solid') el.style.background = fill.color;
    else if (fill.type === 'gradient') el.style.background = fill.css;
    else if (fill.type === 'image') {
      el.style.backgroundImage = `url("${fill.url}")`;
      el.style.backgroundSize = fill.tile ? 'auto' : '100% 100%';
      el.style.backgroundRepeat = fill.tile ? 'repeat' : 'no-repeat';
    }
  }

  const SVG_NS = 'http://www.w3.org/2000/svg';

  const DASHES = {
    dash: [4, 3], sysDash: [3, 1], dot: [1, 2], sysDot: [1, 1], dashDot: [4, 3, 1, 3],
    sysDashDot: [3, 1, 1, 1], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
  };

  function renderShape(item) {
    const el = positioned(item, 'shape');
    const flip = flipTransform(item);
    const isRect = item.geom === 'rect';

    if (item.fill) {
      const fillEl = document.createElement('div');
      fillEl.className = 'fill';
      applyFill(fillEl, item.fill);
      if (!isRect) fillEl.style.clipPath = `path('${item.path}')`;
      if (flip) fillEl.style.transform = flip;
      el.appendChild(fillEl);
    }

    if (item.line) {
      const svg = document.createElementNS(SVG_NS, 'svg');
      svg.setAttribute('class', 'outline');
      svg.setAttribute('width', String(Math.max(item.w, 1)));
      svg.setAttribute('height', String(Math.max(item.h, 1)));
      if (flip) svg.style.transform = flip;
      const p = document.createElementNS(SVG_NS, 'path');
      p.setAttribute('d', item.path);
      p.setAttribute('fill', 'none');
      p.setAttribute('stroke', item.line.color);
      p.setAttribute('stroke-width', String(item.line.width));
      p.setAttribute('stroke-linejoin', 'round');
      if (item.line.dash && DASHES[item.line.dash]) {
        p.setAttribute('stroke-dasharray', DASHES[item.line.dash].map((v) => v * item.line.width).join(' '));
      }
      if (item.line.head || item.line.tail) {
        const defs = document.createElementNS(SVG_NS, 'defs');
        if (item.line.tail) p.setAttribute('marker-end', `url(#${addMarker(defs, item.line.color, 'auto')})`);
        if (item.line.head) p.setAttribute('marker-start', `url(#${addMarker(defs, item.line.color, 'auto-start-reverse')})`);
        svg.appendChild(defs);
      }
      svg.appendChild(p);
      el.appendChild(svg);
    }

    if (item.text) el.appendChild(renderTextBody(item.text));
    return el;
  }

  function addMarker(defs, color, orient) {
    const id = `arrow-${++markerSeq}`;
    const m = document.createElementNS(SVG_NS, 'marker');
    m.setAttribute('id', id);
    m.setAttribute('viewBox', '0 0 10 10');
    m.setAttribute('refX', '8');
    m.setAttribute('refY', '5');
    m.setAttribute('markerWidth', '4');
    m.setAttribute('markerHeight', '4');
    m.setAttribute('orient', orient);
    const tri = document.createElementNS(SVG_NS, 'path');
    tri.setAttribute('d', 'M0 0 L10 5 L0 10 Z');
    tri.setAttribute('fill', color);
    m.appendChild(tri);
    defs.appendChild(m);
    return id;
  }

  function renderImage(item) {
    const el = positioned(item, 'image');
    if (item.clipPath) el.style.clipPath = `path('${item.clipPath}')`;
    const inner = document.createElement('div');
    inner.className = 'image-inner';
    const flip = flipTransform(item);
    if (flip) inner.style.transform = flip;
    if (item.url) {
      const img = document.createElement('img');
      img.src = item.url;
      img.alt = '';
      img.draggable = false;
      const c = item.crop;
      if (c) {
        const fw = Math.max(1 - c.l - c.r, 0.01), fh = Math.max(1 - c.t - c.b, 0.01);
        img.style.width = `${item.w / fw}px`;
        img.style.height = `${item.h / fh}px`;
        img.style.left = `${(-c.l / fw) * item.w}px`;
        img.style.top = `${(-c.t / fh) * item.h}px`;
      } else {
        img.style.width = '100%';
        img.style.height = '100%';
      }
      inner.appendChild(img);
    } else {
      inner.classList.add('unsupported');
      inner.textContent = `Unsupported image (${item.unsupported})`;
    }
    el.appendChild(inner);
    if (item.line) {
      el.style.outline = `${item.line.width}px solid ${item.line.color}`;
    }
    return el;
  }

  function renderPlaceholder(item) {
    const el = positioned(item, 'placeholder');
    const label = document.createElement('span');
    label.textContent = item.label;
    el.appendChild(label);
    return el;
  }

  function borderCss(line) {
    return line ? `${line.width}px ${line.dash ? 'dashed' : 'solid'} ${line.color}` : 'none';
  }

  function renderTable(item) {
    const el = positioned(item, 'table');
    const table = document.createElement('table');
    const colgroup = document.createElement('colgroup');
    for (const w of item.cols) {
      const col = document.createElement('col');
      col.style.width = `${w}px`;
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);
    table.style.width = `${item.cols.reduce((a, b) => a + b, 0)}px`;
    for (const row of item.rows) {
      const tr = document.createElement('tr');
      tr.style.height = `${row.h}px`;
      for (const cell of row.cells) {
        if (cell.skip) continue;
        const td = document.createElement('td');
        if (cell.colspan > 1) td.colSpan = cell.colspan;
        if (cell.rowspan > 1) td.rowSpan = cell.rowspan;
        applyFill(td, cell.fill);
        td.style.borderLeft = borderCss(cell.borders.l);
        td.style.borderRight = borderCss(cell.borders.r);
        td.style.borderTop = borderCss(cell.borders.t);
        td.style.borderBottom = borderCss(cell.borders.b);
        td.style.padding = `${cell.padding.t}px ${cell.padding.r}px ${cell.padding.b}px ${cell.padding.l}px`;
        td.style.verticalAlign = cell.anchor === 'ctr' ? 'middle' : cell.anchor === 'b' ? 'bottom' : 'top';
        for (const p of cell.paragraphs) td.appendChild(renderParagraph(p, true));
        tr.appendChild(td);
      }
      table.appendChild(tr);
    }
    el.appendChild(table);
    return el;
  }

  // ---- text ----------------------------------------------------------------------

  const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify' };
  const ANCHOR = { t: 'flex-start', ctr: 'center', b: 'flex-end', just: 'center', dist: 'center' };

  function fontStack(name) {
    const fallback = 'Calibri, Carlito, "Segoe UI", "Helvetica Neue", Arial, sans-serif';
    return name ? `"${name.replace(/"/g, '')}", ${fallback}` : fallback;
  }

  function renderTextBody(text) {
    const box = document.createElement('div');
    box.className = 'text';
    box.style.padding = `${text.insets.t}px ${text.insets.r}px ${text.insets.b}px ${text.insets.l}px`;
    box.style.justifyContent = ANCHOR[text.anchor] || 'flex-start';
    if (text.anchorCtr) box.style.alignItems = 'center';
    if (!text.wrap) box.style.whiteSpace = 'pre';
    if (text.vert === 'vert' || text.vert === 'eaVert' || text.vert === 'wordArtVertRtl') {
      box.style.writingMode = 'vertical-rl';
    } else if (text.vert === 'vert270') {
      box.style.writingMode = 'vertical-rl';
      box.style.transform = 'rotate(180deg)';
    }
    for (const p of text.paragraphs) box.appendChild(renderParagraph(p, text.wrap));
    return box;
  }

  function renderParagraph(p) {
    const div = document.createElement('div');
    div.className = 'para';
    div.style.textAlign = ALIGN[p.align] || 'left';
    if (p.marL) div.style.paddingLeft = `${p.marL}px`;
    if (p.indent) div.style.textIndent = `${p.indent}px`;
    if (p.spaceBefore) div.style.marginTop = `${p.spaceBefore}px`;
    if (p.spaceAfter) div.style.marginBottom = `${p.spaceAfter}px`;
    if (p.lineHeight) div.style.lineHeight = p.lineHeight.px != null ? `${p.lineHeight.px}px` : String(p.lineHeight.factor);

    if (p.bullet) {
      const bu = document.createElement('span');
      bu.className = 'bullet';
      bu.textContent = p.bullet.text;
      bu.style.fontSize = `${p.bullet.size}px`;
      if (p.bullet.color) bu.style.color = p.bullet.color;
      if (p.bullet.font) bu.style.fontFamily = fontStack(p.bullet.font);
      bu.style.minWidth = `${Math.max(-p.indent, 0)}px`;
      div.appendChild(bu);
    }

    let any = false;
    for (const r of p.runs) {
      if (r.br) {
        div.appendChild(document.createElement('br'));
        continue;
      }
      const span = document.createElement('span');
      span.textContent = r.text;
      span.style.fontSize = `${r.size}px`;
      span.style.fontFamily = fontStack(r.font);
      if (r.color) span.style.color = r.color;
      if (r.bold) span.style.fontWeight = '700';
      if (r.italic) span.style.fontStyle = 'italic';
      const deco = [r.underline && 'underline', r.strike && 'line-through'].filter(Boolean).join(' ');
      if (deco) span.style.textDecoration = deco;
      if (r.smallCaps) span.style.fontVariant = 'small-caps';
      if (r.spacing) span.style.letterSpacing = `${r.spacing}px`;
      if (r.highlight) span.style.backgroundColor = r.highlight;
      if (r.baseline) {
        span.style.verticalAlign = `${r.baseline > 0 ? 'super' : 'sub'}`;
        span.style.fontSize = `${r.size * 0.7}px`;
      }
      div.appendChild(span);
      if (r.text) any = true;
    }
    if (!any) {
      // Keep the height of empty paragraphs.
      const span = document.createElement('span');
      span.textContent = '​';
      span.style.fontSize = `${p.endSize}px`;
      div.appendChild(span);
    }
    return div;
  }

  vscode.postMessage({ type: 'ready' });
})();
