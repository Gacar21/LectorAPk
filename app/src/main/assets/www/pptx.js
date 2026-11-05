// Renderizador de PowerPoint (.pptx) a HTML.
// Lee el XML de cada diapositiva y dibuja formas, texto, imágenes, tablas y SmartArt.
// No es una réplica exacta de PowerPoint, pero conserva posiciones, colores, tamaños y fondos.

const BASE_WIDTH = 960; // ancho interno de cada diapositiva en px (luego se escala a la pantalla)
const EMU_PER_PT = 12700;

/* ---------- utilidades XML (por nombre local, sin depender del prefijo) ---------- */

const parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');

const children = (el, name) => {
  if (!el) return [];
  const out = [];
  for (const c of el.children) if (!name || c.localName === name) out.push(c);
  return out;
};
const child = (el, name) => {
  if (!el) return null;
  for (const c of el.children) if (c.localName === name) return c;
  return null;
};
const path = (el, ...names) => names.reduce((node, n) => child(node, n), el);
const attr = (el, name) => (el ? el.getAttribute(name) : null);
const num = (el, name, fallback = null) => {
  const v = attr(el, name);
  return v === null || v === '' ? fallback : Number(v);
};
const relId = (el, name) => {
  if (!el) return null;
  for (const a of el.attributes) if (a.localName === name && a.prefix === 'r') return a.value;
  return el.getAttribute('r:' + name);
};

const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
const resolvePath = (base, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out = [];
  for (const part of parts) {
    if (part === '..') out.pop();
    else if (part !== '.' && part !== '') out.push(part);
  }
  return out.join('/');
};
const relsPathOf = (part) => dirOf(part) + '_rels/' + part.slice(part.lastIndexOf('/') + 1) + '.rels';

/* ---------- colores ---------- */

const PRESET_COLORS = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
  yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', darkBlue: '00008B',
  darkRed: '8B0000', darkGreen: '006400', lightGray: 'D3D3D3', navy: '000080',
};

const hexToRgb = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));

const rgbToHsl = ([r, g, b]) => {
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
  return [h, s, l];
};

const hslToRgb = ([h, s, l]) => {
  if (s === 0) return [l * 255, l * 255, l * 255];
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
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
};

const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

/* ---------- clase principal ---------- */

export class PptxRenderer {
  constructor(zip) {
    this.zip = zip;
    this.xmlCache = new Map();
    this.relsCache = new Map();
    this.imageCache = new Map();
    this.objectUrls = [];
  }

  static async load(arrayBuffer) {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const r = new PptxRenderer(zip);
    await r.init();
    return r;
  }

  async xml(part) {
    if (!this.xmlCache.has(part)) {
      const f = this.zip.file(part);
      this.xmlCache.set(part, f ? parseXml(await f.async('string')) : null);
    }
    return this.xmlCache.get(part);
  }

  async rels(part) {
    if (!this.relsCache.has(part)) {
      const map = new Map();
      const doc = await this.xml(relsPathOf(part));
      if (doc) {
        for (const r of children(doc.documentElement, 'Relationship')) {
          const external = attr(r, 'TargetMode') === 'External';
          map.set(attr(r, 'Id'), {
            type: attr(r, 'Type') || '',
            target: external ? attr(r, 'Target') : resolvePath(part, attr(r, 'Target')),
            external,
          });
        }
      }
      this.relsCache.set(part, map);
    }
    return this.relsCache.get(part);
  }

  async relOfType(part, suffix) {
    for (const r of (await this.rels(part)).values()) if (r.type.endsWith(suffix)) return r.target;
    return null;
  }

  async init() {
    const pres = await this.xml('ppt/presentation.xml');
    if (!pres) throw new Error('El archivo no parece una presentación de PowerPoint (.pptx)');
    const root = pres.documentElement;
    const sz = child(root, 'sldSz');
    this.cx = num(sz, 'cx', 12192000);
    this.cy = num(sz, 'cy', 6858000);
    this.k = BASE_WIDTH / this.cx; // px por EMU
    this.width = BASE_WIDTH;
    this.height = Math.round(this.cy * this.k);

    const presRels = await this.rels('ppt/presentation.xml');
    this.defaultTextStyle = child(root, 'defaultTextStyle');
    this.slidePaths = children(child(root, 'sldIdLst'), 'sldId')
      .map((s) => presRels.get(relId(s, 'id')))
      .filter(Boolean)
      .map((r) => r.target);

    // Tema y fuentes del primer patrón
    const masterRel = children(child(root, 'sldMasterIdLst'), 'sldMasterId')[0];
    const masterPath = masterRel ? presRels.get(relId(masterRel, 'id'))?.target : null;
    this.theme = { colors: {}, major: 'sans-serif', minor: 'sans-serif' };
    if (masterPath) await this.loadTheme(masterPath);
  }

  async loadTheme(masterPath) {
    const themePath = await this.relOfType(masterPath, '/theme');
    const doc = themePath && (await this.xml(themePath));
    if (!doc) return;
    const elements = doc.documentElement.getElementsByTagName('*');
    for (const el of elements) {
      if (el.localName === 'clrScheme') {
        for (const c of el.children) {
          const v = child(c, 'srgbClr') ? attr(child(c, 'srgbClr'), 'val') : attr(child(c, 'sysClr'), 'lastClr');
          if (v) this.theme.colors[c.localName] = v;
        }
      } else if (el.localName === 'majorFont') {
        this.theme.major = attr(child(el, 'latin'), 'typeface') || this.theme.major;
      } else if (el.localName === 'minorFont') {
        this.theme.minor = attr(child(el, 'latin'), 'typeface') || this.theme.minor;
      }
    }
  }

  get slideCount() {
    return this.slidePaths.length;
  }

  /* ---------- color ---------- */

  colorFrom(container, ctx, inherited = null) {
    if (!container) return null;
    let el = null;
    for (const c of container.children) {
      if (['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'].includes(c.localName)) {
        el = c;
        break;
      }
    }
    if (!el) return null;
    let hex = null;
    switch (el.localName) {
      case 'srgbClr': hex = attr(el, 'val'); break;
      case 'sysClr': hex = attr(el, 'lastClr') || (attr(el, 'val') === 'window' ? 'FFFFFF' : '000000'); break;
      case 'prstClr': hex = PRESET_COLORS[attr(el, 'val')] || '000000'; break;
      case 'scrgbClr':
        hex = ['r', 'g', 'b'].map((k) => Math.round(clamp(num(el, k, 0) / 100000, 0, 1) * 255).toString(16).padStart(2, '0')).join('');
        break;
      case 'schemeClr': {
        let key = attr(el, 'val');
        if (key === 'phClr') {
          if (!inherited) return null;
          hex = inherited;
          break;
        }
        key = (ctx?.clrMap && ctx.clrMap[key]) || { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }[key] || key;
        hex = this.theme.colors[key] || '000000';
        break;
      }
      default: hex = '000000';
    }
    if (!hex || hex.length !== 6) return null;
    let rgb = hexToRgb(hex);
    let alpha = 1;
    for (const m of el.children) {
      const v = num(m, 'val', 100000) / 100000;
      if (m.localName === 'alpha') alpha = v;
      else if (m.localName === 'tint') rgb = rgb.map((c) => c + (255 - c) * (1 - v));
      else if (m.localName === 'shade') rgb = rgb.map((c) => c * v);
      else if (m.localName === 'lumMod' || m.localName === 'lumOff') {
        const hsl = rgbToHsl(rgb);
        hsl[2] = m.localName === 'lumMod' ? hsl[2] * v : hsl[2] + v;
        hsl[2] = clamp(hsl[2], 0, 1);
        rgb = hslToRgb(hsl);
      } else if (m.localName === 'satMod') {
        const hsl = rgbToHsl(rgb);
        hsl[1] = clamp(hsl[1] * v, 0, 1);
        rgb = hslToRgb(hsl);
      }
    }
    rgb = rgb.map((c) => Math.round(clamp(c, 0, 255)));
    return alpha < 1 ? `rgba(${rgb.join(',')},${alpha.toFixed(3)})` : `rgb(${rgb.join(',')})`;
  }

  schemeColor(key, ctx) {
    const mapped = (ctx?.clrMap && ctx.clrMap[key]) || { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }[key] || key;
    const hex = this.theme.colors[mapped];
    return hex ? '#' + hex : '#000';
  }

  /* ---------- imágenes ---------- */

  async imageUrl(partPath, rId) {
    const rel = (await this.rels(partPath)).get(rId);
    if (!rel || rel.external) return null;
    const target = rel.target;
    if (this.imageCache.has(target)) return this.imageCache.get(target);
    const f = this.zip.file(target);
    let url = null;
    const ext = target.split('.').pop().toLowerCase();
    const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' }[ext];
    if (f && mime) {
      const blob = new Blob([await f.async('uint8array')], { type: mime });
      url = URL.createObjectURL(blob);
      this.objectUrls.push(url);
    }
    this.imageCache.set(target, url);
    return url;
  }

  dispose() {
    this.objectUrls.forEach((u) => URL.revokeObjectURL(u));
    this.objectUrls = [];
  }

  /* ---------- contexto de cada diapositiva (diseño + patrón) ---------- */

  async partContext(slidePath) {
    const slide = await this.xml(slidePath);
    const layoutPath = await this.relOfType(slidePath, '/slideLayout');
    const layout = layoutPath ? await this.xml(layoutPath) : null;
    const masterPath = layoutPath ? await this.relOfType(layoutPath, '/slideMaster') : null;
    const master = masterPath ? await this.xml(masterPath) : null;

    const clrMap = {};
    const mapEl = master && child(master.documentElement, 'clrMap');
    if (mapEl) for (const a of mapEl.attributes) clrMap[a.name] = a.value;
    const override = slide && path(slide.documentElement, 'clrMapOvr', 'overrideClrMapping');
    if (override) for (const a of override.attributes) clrMap[a.name] = a.value;

    return {
      slidePath, layoutPath, masterPath,
      slide: slide?.documentElement,
      layout: layout?.documentElement,
      master: master?.documentElement,
      clrMap,
      txStyles: master ? child(master.documentElement, 'txStyles') : null,
    };
  }

  findPlaceholder(root, ph) {
    if (!root || !ph) return null;
    const type = attr(ph, 'type') || 'body';
    const idx = attr(ph, 'idx');
    const tree = path(root, 'cSld', 'spTree');
    const shapes = tree ? [...tree.getElementsByTagName('*')].filter((e) => e.localName === 'sp') : [];
    const phOf = (sp) => path(sp, 'nvSpPr', 'nvPr', 'ph');
    const same = (a, b) => a === b || (['title', 'ctrTitle'].includes(a) && ['title', 'ctrTitle'].includes(b));
    if (idx !== null) {
      const byIdx = shapes.find((sp) => attr(phOf(sp), 'idx') === idx);
      if (byIdx) return byIdx;
    }
    return shapes.find((sp) => phOf(sp) && same(attr(phOf(sp), 'type') || 'body', type)) || null;
  }

  /* ---------- render de una diapositiva ---------- */

  async renderSlide(index) {
    const ctx = await this.partContext(this.slidePaths[index]);
    ctx.slideNumber = index + 1;

    const el = document.createElement('div');
    el.className = 'pptx-slide';
    el.style.width = this.width + 'px';
    el.style.height = this.height + 'px';
    el.style.background = (await this.background(ctx)) || '#fff';

    const showMasterShapes = attr(ctx.slide, 'showMasterSp') !== '0';
    if (showMasterShapes && ctx.master && attr(ctx.layout, 'showMasterSp') !== '0') {
      await this.renderTree(path(ctx.master, 'cSld', 'spTree'), el, { ...ctx, part: ctx.masterPath, level: 'master' });
    }
    if (showMasterShapes && ctx.layout) {
      await this.renderTree(path(ctx.layout, 'cSld', 'spTree'), el, { ...ctx, part: ctx.layoutPath, level: 'layout' });
    }
    await this.renderTree(path(ctx.slide, 'cSld', 'spTree'), el, { ...ctx, part: ctx.slidePath, level: 'slide' });
    return el;
  }

  async background(ctx) {
    const sources = [
      [ctx.slide, ctx.slidePath],
      [ctx.layout, ctx.layoutPath],
      [ctx.master, ctx.masterPath],
    ];
    for (const [root, part] of sources) {
      const bg = path(root, 'cSld', 'bg');
      if (!bg) continue;
      const bgPr = child(bg, 'bgPr');
      if (bgPr) return this.fillCss(bgPr, ctx, part);
      const ref = child(bg, 'bgRef');
      if (ref) return this.colorFrom(ref, ctx);
    }
    return null;
  }

  async fillCss(container, ctx, part, inherited = null) {
    if (!container) return null;
    if (child(container, 'noFill')) return 'transparent';
    const solid = child(container, 'solidFill');
    if (solid) return this.colorFrom(solid, ctx, inherited);
    const grad = child(container, 'gradFill');
    if (grad) {
      const stops = children(child(grad, 'gsLst'), 'gs')
        .map((gs) => `${this.colorFrom(gs, ctx, inherited) || '#ccc'} ${(num(gs, 'pos', 0) / 1000).toFixed(1)}%`);
      if (!stops.length) return null;
      const lin = child(grad, 'lin');
      if (child(grad, 'path')) return `radial-gradient(circle, ${stops.join(', ')})`;
      const angle = lin ? num(lin, 'ang', 0) / 60000 : 90;
      return `linear-gradient(${(angle + 90).toFixed(1)}deg, ${stops.join(', ')})`;
    }
    const blip = child(container, 'blipFill');
    if (blip) {
      const url = await this.imageUrl(part, relId(child(blip, 'blip'), 'embed'));
      if (url) return `center / cover no-repeat url("${url}")`;
    }
    const patt = child(container, 'pattFill');
    if (patt) return this.colorFrom(child(patt, 'fgClr'), ctx) || null;
    return null;
  }

  /* ---------- árbol de formas ---------- */

  identity() {
    return (x, y, w, h) => ({ x, y, w, h });
  }

  async renderTree(tree, parentEl, ctx, transform = this.identity()) {
    if (!tree) return;
    for (let node of tree.children) {
      if (node.localName === 'AlternateContent') {
        node = child(node, 'Fallback') || child(node, 'Choice');
        if (!node) continue;
        for (const inner of node.children) await this.renderNode(inner, parentEl, ctx, transform);
        continue;
      }
      await this.renderNode(node, parentEl, ctx, transform);
    }
  }

  async renderNode(node, parentEl, ctx, transform) {
    try {
      switch (node.localName) {
        case 'sp': return await this.renderShape(node, parentEl, ctx, transform);
        case 'cxnSp': return await this.renderShape(node, parentEl, ctx, transform, true);
        case 'pic': return await this.renderPicture(node, parentEl, ctx, transform);
        case 'grpSp': return await this.renderGroup(node, parentEl, ctx, transform);
        case 'graphicFrame': return await this.renderFrame(node, parentEl, ctx, transform);
        default: return null;
      }
    } catch (err) {
      console.warn('Forma omitida', err);
      return null;
    }
  }

  xfrmBox(xfrm, transform) {
    if (!xfrm) return null;
    const off = child(xfrm, 'off');
    const ext = child(xfrm, 'ext');
    if (!off || !ext) return null;
    const box = transform(num(off, 'x', 0), num(off, 'y', 0), num(ext, 'cx', 0), num(ext, 'cy', 0));
    return {
      ...box,
      rot: num(xfrm, 'rot', 0) / 60000,
      flipH: attr(xfrm, 'flipH') === '1',
      flipV: attr(xfrm, 'flipV') === '1',
    };
  }

  place(el, box) {
    const k = this.k;
    el.style.left = box.x * k + 'px';
    el.style.top = box.y * k + 'px';
    el.style.width = Math.max(box.w * k, 0.5) + 'px';
    el.style.height = Math.max(box.h * k, 0.5) + 'px';
    const t = [];
    if (box.rot) t.push(`rotate(${box.rot}deg)`);
    if (box.flipH) t.push('scaleX(-1)');
    if (box.flipV) t.push('scaleY(-1)');
    if (t.length) el.style.transform = t.join(' ');
  }

  async renderGroup(node, parentEl, ctx, transform) {
    const xfrm = path(node, 'grpSpPr', 'xfrm');
    let inner = transform;
    if (xfrm) {
      const off = child(xfrm, 'off'), ext = child(xfrm, 'ext');
      const chOff = child(xfrm, 'chOff'), chExt = child(xfrm, 'chExt');
      if (off && ext && chOff && chExt) {
        const ox = num(off, 'x', 0), oy = num(off, 'y', 0);
        const sx = num(chExt, 'cx', 1) ? num(ext, 'cx', 0) / num(chExt, 'cx', 1) : 1;
        const sy = num(chExt, 'cy', 1) ? num(ext, 'cy', 0) / num(chExt, 'cy', 1) : 1;
        const cx = num(chOff, 'x', 0), cy = num(chOff, 'y', 0);
        inner = (x, y, w, h) => transform(ox + (x - cx) * sx, oy + (y - cy) * sy, w * sx, h * sy);
      }
    }
    await this.renderTree(node, parentEl, ctx, inner);
  }

  async renderShape(sp, parentEl, ctx, transform, isConnector = false) {
    const ph = path(sp, 'nvSpPr', 'nvPr', 'ph');
    // En diseño y patrón solo se dibujan formas decorativas, no los marcadores vacíos
    if (ph && ctx.level !== 'slide') return;

    const layoutPh = ph ? this.findPlaceholder(ctx.layout, ph) : null;
    const masterPh = ph ? this.findPlaceholder(ctx.master, layoutPh ? path(layoutPh, 'nvSpPr', 'nvPr', 'ph') || ph : ph) : null;
    const chain = [sp, layoutPh, masterPh].filter(Boolean);

    const spPr = child(sp, 'spPr');
    let box = null;
    for (const s of chain) {
      box = this.xfrmBox(path(s, 'spPr', 'xfrm'), s === sp ? transform : this.identity());
      if (box) break;
    }
    if (!box) return;

    const el = document.createElement('div');
    el.className = 'pptx-shape';
    this.place(el, box);

    const style = child(sp, 'style');
    const fillRef = child(style, 'fillRef');
    const styleColor = fillRef && num(fillRef, 'idx', 0) > 0 ? this.colorFrom(fillRef, ctx) : null;

    let fill = await this.fillCss(spPr, ctx, ctx.part, null);
    if (!fill && layoutPh) fill = await this.fillCss(child(layoutPh, 'spPr'), ctx, ctx.layoutPath);
    if (!fill && !isConnector) fill = styleColor;

    const geom = child(spPr, 'prstGeom');
    const prst = attr(geom, 'prst') || (child(spPr, 'custGeom') ? 'custom' : 'rect');
    const line = this.lineCss(child(spPr, 'ln'), child(style, 'lnRef'), ctx);

    if (isConnector || prst === 'line' || prst.startsWith('straightConnector') || prst.startsWith('bentConnector') || prst.startsWith('curvedConnector')) {
      this.drawLine(el, box, line);
      parentEl.appendChild(el);
      return;
    }

    const custGeom = child(spPr, 'custGeom');
    if (custGeom) {
      el.appendChild(this.customGeometry(custGeom, box, fill, line));
    } else {
      if (fill && fill !== 'transparent') el.style.background = fill;
      if (line) el.style.border = `${line.width}px ${line.dash} ${line.color}`;
      this.applyPreset(el, prst, box);
    }

    const txBody = child(sp, 'txBody');
    if (txBody) {
      const fontRef = child(style, 'fontRef');
      const text = this.renderText(txBody, ctx, {
        phType: ph ? attr(ph, 'type') || 'body' : null,
        layoutPh,
        masterPh,
        defaultColor: fontRef ? this.colorFrom(fontRef, ctx) : null,
        fontRefIdx: attr(fontRef, 'idx'),
      });
      if (text) {
        // El texto no debe heredar el giro espejo de la forma
        if (box.flipH || box.flipV) text.style.transform = `${box.flipH ? 'scaleX(-1)' : ''} ${box.flipV ? 'scaleY(-1)' : ''}`;
        el.appendChild(text);
      }
    }
    parentEl.appendChild(el);
  }

  lineCss(ln, lnRef, ctx) {
    if (ln && child(ln, 'noFill')) return null;
    let color = ln ? this.colorFrom(child(ln, 'solidFill'), ctx) : null;
    if (!color && lnRef && num(lnRef, 'idx', 0) > 0) color = this.colorFrom(lnRef, ctx);
    if (!color) return null;
    const width = Math.max(((ln && num(ln, 'w', 12700)) || 12700) * this.k, 0.75);
    const dashVal = attr(child(ln, 'prstDash'), 'val');
    const dash = !dashVal || dashVal === 'solid' ? 'solid' : dashVal.toLowerCase().includes('dot') && !dashVal.toLowerCase().includes('dash') ? 'dotted' : 'dashed';
    return { color, width, dash };
  }

  drawLine(el, box, line) {
    if (!line) return;
    const w = box.w * this.k, h = box.h * this.k;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('width', Math.max(w, 1));
    svg.setAttribute('height', Math.max(h, 1));
    svg.style.overflow = 'visible';
    const l = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    l.setAttribute('x1', 0); l.setAttribute('y1', 0);
    l.setAttribute('x2', w); l.setAttribute('y2', h);
    l.setAttribute('stroke', line.color);
    l.setAttribute('stroke-width', line.width);
    if (line.dash !== 'solid') l.setAttribute('stroke-dasharray', line.dash === 'dotted' ? `${line.width} ${line.width * 2}` : `${line.width * 4} ${line.width * 3}`);
    svg.appendChild(l);
    el.appendChild(svg);
  }

  applyPreset(el, prst, box) {
    const minSide = Math.min(box.w, box.h) * this.k;
    const clips = {
      triangle: 'polygon(50% 0, 100% 100%, 0 100%)',
      rtTriangle: 'polygon(0 0, 100% 100%, 0 100%)',
      diamond: 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)',
      parallelogram: 'polygon(25% 0, 100% 0, 75% 100%, 0 100%)',
      trapezoid: 'polygon(25% 0, 75% 0, 100% 100%, 0 100%)',
      pentagon: 'polygon(50% 0, 100% 38%, 82% 100%, 18% 100%, 0 38%)',
      hexagon: 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)',
      octagon: 'polygon(30% 0, 70% 0, 100% 30%, 100% 70%, 70% 100%, 30% 100%, 0 70%, 0 30%)',
      homePlate: 'polygon(0 0, 80% 0, 100% 50%, 80% 100%, 0 100%)',
      chevron: 'polygon(0 0, 80% 0, 100% 50%, 80% 100%, 0 100%, 20% 50%)',
      rightArrow: 'polygon(0 25%, 65% 25%, 65% 0, 100% 50%, 65% 100%, 65% 75%, 0 75%)',
      leftArrow: 'polygon(100% 25%, 35% 25%, 35% 0, 0 50%, 35% 100%, 35% 75%, 100% 75%)',
      upArrow: 'polygon(25% 100%, 25% 35%, 0 35%, 50% 0, 100% 35%, 75% 35%, 75% 100%)',
      downArrow: 'polygon(25% 0, 75% 0, 75% 65%, 100% 65%, 50% 100%, 0 65%, 25% 65%)',
      star5: 'polygon(50% 0, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)',
      plus: 'polygon(35% 0, 65% 0, 65% 35%, 100% 35%, 100% 65%, 65% 65%, 65% 100%, 35% 100%, 35% 65%, 0 65%, 0 35%, 35% 35%)',
    };
    if (prst === 'ellipse' || prst === 'donut' || prst === 'pie' || prst === 'flowChartConnector') el.style.borderRadius = '50%';
    else if (prst === 'roundRect' || prst === 'flowChartAlternateProcess') el.style.borderRadius = minSide * 0.1667 + 'px';
    else if (prst === 'snipRoundRect' || prst === 'round2SameRect' || prst === 'round1Rect') el.style.borderRadius = minSide * 0.1667 + 'px ' + minSide * 0.1667 + 'px 0 0';
    else if (clips[prst]) el.style.clipPath = clips[prst];
  }

  customGeometry(custGeom, box, fill, line) {
    const w = box.w * this.k, h = box.h * this.k;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('width', w);
    svg.setAttribute('height', h);
    svg.style.position = 'absolute';
    svg.style.inset = '0';
    svg.style.overflow = 'visible';
    for (const p of children(child(custGeom, 'pathLst'), 'path')) {
      const pw = num(p, 'w', box.w) || box.w, phh = num(p, 'h', box.h) || box.h;
      const sx = w / pw, sy = h / phh;
      const pt = (ptEl) => `${num(ptEl, 'x', 0) * sx},${num(ptEl, 'y', 0) * sy}`;
      let d = '';
      for (const cmd of p.children) {
        const pts = children(cmd, 'pt');
        if (cmd.localName === 'moveTo') d += `M${pt(pts[0])} `;
        else if (cmd.localName === 'lnTo') d += `L${pt(pts[0])} `;
        else if (cmd.localName === 'cubicBezTo' && pts.length === 3) d += `C${pts.map(pt).join(' ')} `;
        else if (cmd.localName === 'quadBezTo' && pts.length === 2) d += `Q${pts.map(pt).join(' ')} `;
        else if (cmd.localName === 'close') d += 'Z ';
      }
      const pathEl = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      pathEl.setAttribute('d', d.trim());
      const isFill = fill && fill !== 'transparent' && attr(p, 'fill') !== 'none' && !fill.includes('gradient') && !fill.includes('url(');
      pathEl.setAttribute('fill', isFill ? fill : 'none');
      if (line && attr(p, 'stroke') !== '0') {
        pathEl.setAttribute('stroke', line.color);
        pathEl.setAttribute('stroke-width', line.width);
      }
      svg.appendChild(pathEl);
    }
    return svg;
  }

  async renderPicture(pic, parentEl, ctx, transform) {
    const box = this.xfrmBox(path(pic, 'spPr', 'xfrm'), transform);
    if (!box) return;
    const blipFill = child(pic, 'blipFill');
    const url = await this.imageUrl(ctx.part, relId(child(blipFill, 'blip'), 'embed'));
    if (!url) return;
    const el = document.createElement('div');
    el.className = 'pptx-shape pptx-pic';
    this.place(el, box);
    const img = document.createElement('img');
    img.src = url;
    img.alt = '';
    const crop = child(blipFill, 'srcRect');
    if (crop) {
      const l = num(crop, 'l', 0) / 100000, t = num(crop, 't', 0) / 100000;
      const r = num(crop, 'r', 0) / 100000, b = num(crop, 'b', 0) / 100000;
      const fw = 1 - l - r, fh = 1 - t - b;
      if (fw > 0 && fh > 0) {
        img.style.width = (100 / fw).toFixed(3) + '%';
        img.style.height = (100 / fh).toFixed(3) + '%';
        img.style.left = (-l / fw * 100).toFixed(3) + '%';
        img.style.top = (-t / fh * 100).toFixed(3) + '%';
      }
    }
    if (path(pic, 'spPr', 'prstGeom') && attr(path(pic, 'spPr', 'prstGeom'), 'prst') === 'ellipse') el.style.borderRadius = '50%';
    const line = this.lineCss(path(pic, 'spPr', 'ln'), null, ctx);
    if (line) el.style.outline = `${line.width}px solid ${line.color}`;
    el.appendChild(img);
    parentEl.appendChild(el);
  }

  async renderFrame(frame, parentEl, ctx, transform) {
    const box = this.xfrmBox(child(frame, 'xfrm'), transform);
    if (!box) return;
    const data = path(frame, 'graphic', 'graphicData');
    const uri = attr(data, 'uri') || '';
    const el = document.createElement('div');
    el.className = 'pptx-shape';
    this.place(el, box);

    if (uri.endsWith('/table')) {
      el.appendChild(this.renderTable(child(data, 'tbl'), ctx));
      el.style.height = 'auto';
    } else if (uri.endsWith('/diagram')) {
      const rendered = await this.renderSmartArt(data, el, ctx, box);
      if (!rendered) el.appendChild(this.unsupported('Diagrama SmartArt'));
    } else if (uri.endsWith('/chart')) {
      el.appendChild(this.unsupported('Gráfico'));
    } else if (uri.includes('ole')) {
      // Objetos incrustados: suelen traer una imagen de vista previa
      const pic = [...data.getElementsByTagName('*')].find((e) => e.localName === 'pic');
      if (pic) {
        await this.renderPicture(pic, parentEl, ctx, transform);
        return;
      }
      el.appendChild(this.unsupported('Objeto incrustado'));
    } else {
      return;
    }
    parentEl.appendChild(el);
  }

  unsupported(label) {
    const d = document.createElement('div');
    d.className = 'pptx-unsupported';
    d.textContent = label;
    return d;
  }

  async renderSmartArt(data, el, ctx, box) {
    const relIds = child(data, 'relIds');
    const dmRel = relIds && (await this.rels(ctx.part)).get(relId(relIds, 'dm'));
    let drawingPath = null;
    if (dmRel) {
      const dm = await this.xml(dmRel.target);
      const ext = dm && [...dm.getElementsByTagName('*')].find((e) => e.localName === 'dataModelExt');
      const id = ext && attr(ext, 'relId');
      if (id) drawingPath = (await this.rels(ctx.part)).get(id)?.target;
    }
    if (!drawingPath) drawingPath = await this.relOfType(ctx.part, '/diagramDrawing');
    const drawing = drawingPath && (await this.xml(drawingPath));
    const tree = drawing && [...drawing.getElementsByTagName('*')].find((e) => e.localName === 'spTree');
    if (!tree) return false;
    // Según la versión de Office, las coordenadas del dibujo vienen relativas al marco
    // o absolutas en la diapositiva. Si la forma más a la izquierda empieza antes que
    // el marco, son relativas; si no, se restan las coordenadas del marco.
    const offs = [...tree.getElementsByTagName('*')].filter((e) => e.localName === 'off');
    const minX = offs.length ? Math.min(...offs.map((o) => num(o, 'x', 0))) : 0;
    const minY = offs.length ? Math.min(...offs.map((o) => num(o, 'y', 0))) : 0;
    const absolute = box.x > 0 && minX >= box.x * 0.98 && minY >= box.y * 0.98;
    const local = absolute
      ? (x, y, w, h) => ({ x: x - box.x, y: y - box.y, w, h })
      : (x, y, w, h) => ({ x, y, w, h });
    await this.renderTree(tree, el, { ...ctx, part: drawingPath, level: 'slide' }, local);
    return true;
  }

  renderTable(tbl, ctx) {
    const table = document.createElement('table');
    table.className = 'pptx-table';
    const tblPr = child(tbl, 'tblPr');
    const firstRow = attr(tblPr, 'firstRow') === '1';
    const bandRow = attr(tblPr, 'bandRow') === '1';
    const accentHex = this.theme.colors.accent1 || '4472C4';

    const colgroup = document.createElement('colgroup');
    for (const gc of children(child(tbl, 'tblGrid'), 'gridCol')) {
      const col = document.createElement('col');
      col.style.width = num(gc, 'w', 0) * this.k + 'px';
      colgroup.appendChild(col);
    }
    table.appendChild(colgroup);

    children(tbl, 'tr').forEach((tr, r) => {
      const row = document.createElement('tr');
      row.style.height = num(tr, 'h', 0) * this.k + 'px';
      for (const tc of children(tr, 'tc')) {
        if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') continue;
        const cell = document.createElement('td');
        const gs = num(tc, 'gridSpan', 1), rs = num(tc, 'rowSpan', 1);
        if (gs > 1) cell.colSpan = gs;
        if (rs > 1) cell.rowSpan = rs;
        const tcPr = child(tc, 'tcPr');
        let bg = tcPr ? this.colorFrom(child(tcPr, 'solidFill'), ctx) : null;
        let defaultColor = null;
        if (!bg && !(tcPr && child(tcPr, 'noFill'))) {
          if (firstRow && r === 0) {
            bg = '#' + accentHex;
            defaultColor = '#fff';
          } else if (bandRow) {
            bg = (r % 2 === (firstRow ? 1 : 0)) ? `rgba(${hexToRgb(accentHex).join(',')},0.22)` : `rgba(${hexToRgb(accentHex).join(',')},0.10)`;
          }
        }
        if (bg) cell.style.background = bg;
        const anchor = attr(tcPr, 'anchor');
        cell.style.verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';
        const txBody = child(tc, 'txBody');
        if (txBody) {
          const t = this.renderText(txBody, ctx, { defaultColor, inTable: true });
          if (t) cell.appendChild(t);
        }
        row.appendChild(cell);
      }
      table.appendChild(row);
    });
    return table;
  }

  /* ---------- texto ---------- */

  levelStyles(lvl, info, txBody, ctx) {
    const key = `lvl${lvl + 1}pPr`;
    const sources = [];
    const push = (lst) => {
      const s = lst && child(lst, key);
      if (s) sources.push(s);
    };
    push(child(txBody, 'lstStyle'));
    if (info.layoutPh) push(path(info.layoutPh, 'txBody', 'lstStyle'));
    if (info.masterPh) push(path(info.masterPh, 'txBody', 'lstStyle'));
    const tx = ctx.txStyles;
    if (tx && !info.inTable) {
      const type = info.phType;
      if (type === 'title' || type === 'ctrTitle') push(child(tx, 'titleStyle'));
      else if (type && !['dt', 'ftr', 'sldNum'].includes(type)) push(child(tx, 'bodyStyle'));
      else push(child(tx, 'otherStyle'));
    }
    // Formas sueltas (no marcadores) heredan del estilo por defecto de la presentación
    if (!info.phType && !info.inTable && this.defaultTextStyle) {
      sources.length = Math.min(sources.length, child(txBody, 'lstStyle') && child(child(txBody, 'lstStyle'), key) ? 1 : 0);
      push(this.defaultTextStyle);
    }
    return sources;
  }

  bodyProps(txBody, info) {
    const chain = [child(txBody, 'bodyPr'), info.layoutPh && path(info.layoutPh, 'txBody', 'bodyPr'), info.masterPh && path(info.masterPh, 'txBody', 'bodyPr')].filter(Boolean);
    const get = (name) => {
      for (const b of chain) if (b.hasAttribute(name)) return b.getAttribute(name);
      return null;
    };
    const auto = chain.map((b) => child(b, 'normAutofit')).find(Boolean);
    return {
      anchor: get('anchor') || (info.phType === 'ctrTitle' ? 'b' : info.phType === 'title' ? 'ctr' : 't'),
      lIns: Number(get('lIns') ?? 91440), rIns: Number(get('rIns') ?? 91440),
      tIns: Number(get('tIns') ?? 45720), bIns: Number(get('bIns') ?? 45720),
      wrap: get('wrap') !== 'none',
      anchorCtr: get('anchorCtr') === '1',
      vert: get('vert'),
      fontScale: auto ? num(auto, 'fontScale', 100000) / 100000 : 1,
      lnReduction: auto ? num(auto, 'lnSpcReduction', 0) / 100000 : 0,
    };
  }

  renderText(txBody, ctx, info) {
    const paragraphs = children(txBody, 'p');
    const hasText = paragraphs.some((p) => p.textContent.trim().length);
    if (!hasText) return null;

    const bp = this.bodyProps(txBody, info);
    const k = this.k;
    const box = document.createElement('div');
    box.className = 'pptx-text';
    box.style.padding = `${bp.tIns * k}px ${bp.rIns * k}px ${bp.bIns * k}px ${bp.lIns * k}px`;
    box.style.justifyContent = bp.anchor === 'ctr' ? 'center' : bp.anchor === 'b' ? 'flex-end' : 'flex-start';
    if (!bp.wrap) box.style.whiteSpace = 'nowrap';
    if (bp.vert === 'vert' || bp.vert === 'eaVert') box.style.writingMode = 'vertical-rl';
    if (bp.vert === 'vert270') {
      box.style.writingMode = 'vertical-rl';
      box.style.transform = 'rotate(180deg)';
    }
    if (info.inTable) box.style.position = 'static';

    const counters = [];
    for (const p of paragraphs) {
      const pPr = child(p, 'pPr');
      const lvl = num(pPr, 'lvl', 0);
      const styles = this.levelStyles(lvl, info, txBody, ctx);
      const pChain = [pPr, ...styles].filter(Boolean);
      const pGet = (name) => {
        for (const s of pChain) if (s.hasAttribute(name)) return s.getAttribute(name);
        return null;
      };
      const pChild = (name) => {
        for (const s of pChain) {
          const c = child(s, name);
          if (c) return c;
        }
        return null;
      };
      const defRPrs = pChain.map((s) => child(s, 'defRPr')).filter(Boolean);

      const para = document.createElement('div');
      para.className = 'pptx-p';
      const algn = pGet('algn');
      para.style.textAlign = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'left';
      const marL = Number(pGet('marL') ?? 0) * k;
      const indent = Number(pGet('indent') ?? 0) * k;
      para.style.paddingLeft = marL + 'px';
      para.style.textIndent = indent + 'px';

      const lnSpc = pChild('lnSpc');
      if (lnSpc && child(lnSpc, 'spcPct')) {
        const pct = num(child(lnSpc, 'spcPct'), 'val', 100000) / 100000 - bp.lnReduction;
        para.style.lineHeight = (1.2 * Math.max(pct, 0.5)).toFixed(3);
      } else if (bp.lnReduction) {
        para.style.lineHeight = (1.2 * (1 - bp.lnReduction)).toFixed(3);
      }
      const spcBef = pChild('spcBef');
      if (spcBef && child(spcBef, 'spcPts')) para.style.marginTop = num(child(spcBef, 'spcPts'), 'val', 0) / 100 * EMU_PER_PT * k + 'px';
      const spcAft = pChild('spcAft');
      if (spcAft && child(spcAft, 'spcPts')) para.style.marginBottom = num(child(spcAft, 'spcPts'), 'val', 0) / 100 * EMU_PER_PT * k + 'px';

      // tamaño base del párrafo (para viñetas y líneas vacías)
      const runs = children(p).filter((c) => ['r', 'br', 'fld'].includes(c.localName));
      const firstRPr = runs.length ? child(runs[0], 'rPr') : child(p, 'endParaRPr');
      const baseRun = this.runStyle([firstRPr, ...defRPrs], ctx, info, bp);

      // viñeta
      const noBullet = pChild('buNone') && !(pPr && (child(pPr, 'buChar') || child(pPr, 'buAutoNum')));
      const buChar = !noBullet ? pChild('buChar') : null;
      const buAuto = !noBullet ? pChild('buAutoNum') : null;
      const firstNonNone = pChain.find((s) => child(s, 'buNone') || child(s, 'buChar') || child(s, 'buAutoNum'));
      const bulletOn = firstNonNone && !child(firstNonNone, 'buNone') && runs.some((r) => r.textContent.trim());

      for (let i = lvl + 1; i < counters.length; i++) counters[i] = 0;
      if (bulletOn && (buChar || buAuto)) {
        const b = document.createElement('span');
        b.className = 'pptx-bullet';
        if (buAuto && child(firstNonNone, 'buAutoNum')) {
          const autoEl = child(firstNonNone, 'buAutoNum');
          const start = num(autoEl, 'startAt', 1);
          counters[lvl] = (counters[lvl] || start - 1) + 1;
          b.textContent = this.autoNumber(attr(autoEl, 'type') || 'arabicPeriod', counters[lvl]);
        } else {
          counters[lvl] = 0;
          const ch = attr(child(firstNonNone, 'buChar'), 'char') || '•';
          b.textContent = { '§': '■', 'Ø': '➢', 'ü': '✓', 'q': '❑', 'n': '■', 'v': '❖', 'Ü': '➔', 'à': '→' }[ch] || ch;
        }
        const buClr = pChild('buClr');
        b.style.color = (buClr && this.colorFrom(buClr, ctx)) || baseRun.color;
        const buSz = pChild('buSzPct');
        b.style.fontSize = baseRun.size * (buSz ? num(buSz, 'val', 100000) / 100000 : 1) + 'px';
        b.style.fontFamily = baseRun.font;
        b.style.width = Math.max(-indent, 0) + 'px';
        if (indent < 0) {
          b.style.marginLeft = indent + 'px';
          para.style.textIndent = '0';
        } else b.style.marginRight = '0.4em';
        para.appendChild(b);
      } else {
        counters[lvl] = 0;
      }

      if (!runs.length || !runs.some((r) => r.textContent.length)) {
        const empty = document.createElement('span');
        empty.style.fontSize = baseRun.size + 'px';
        empty.textContent = '​';
        para.appendChild(empty);
      }

      for (const r of runs) {
        if (r.localName === 'br') {
          para.appendChild(document.createElement('br'));
          continue;
        }
        let text = child(r, 't')?.textContent ?? '';
        if (r.localName === 'fld' && (attr(r, 'type') || '') === 'slidenum') text = String(ctx.slideNumber);
        const st = this.runStyle([child(r, 'rPr'), ...defRPrs], ctx, info, bp);
        const span = document.createElement('span');
        span.textContent = text;
        span.style.fontSize = st.size + 'px';
        span.style.color = st.color;
        span.style.fontFamily = st.font;
        if (st.bold) span.style.fontWeight = '700';
        if (st.italic) span.style.fontStyle = 'italic';
        const deco = [st.underline && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ');
        if (deco) span.style.textDecoration = deco;
        if (st.caps) span.style.textTransform = 'uppercase';
        if (st.spacing) span.style.letterSpacing = st.spacing + 'px';
        if (st.baseline) {
          span.style.verticalAlign = st.baseline > 0 ? 'super' : 'sub';
          span.style.fontSize = st.size * 0.65 + 'px';
        }
        if (st.highlight) span.style.background = st.highlight;
        para.appendChild(span);
      }
      box.appendChild(para);
    }
    return box;
  }

  runStyle(chain, ctx, info, bp) {
    chain = chain.filter(Boolean);
    const get = (name) => {
      for (const s of chain) if (s.hasAttribute(name)) return s.getAttribute(name);
      return null;
    };
    const getChild = (name) => {
      for (const s of chain) {
        const c = child(s, name);
        if (c) return c;
      }
      return null;
    };
    const defaultSize = info.phType === 'title' || info.phType === 'ctrTitle' ? 44 : info.inTable ? 14 : 18;
    const sizePt = (Number(get('sz')) || defaultSize * 100) / 100 * bp.fontScale;

    // Prioridad del color: el run, luego el color de fuente del estilo de la forma, luego la herencia
    let color = null;
    const ownFill = child(chain[0], 'solidFill');
    if (ownFill && chain[0].localName === 'rPr') color = this.colorFrom(ownFill, ctx);
    if (!color && info.defaultColor) color = info.defaultColor;
    if (!color) {
      const fillEl = getChild('solidFill');
      if (fillEl) color = this.colorFrom(fillEl, ctx);
    }
    if (!color) color = this.schemeColor('tx1', ctx);

    let face = attr(getChild('latin'), 'typeface') || attr(getChild('ea'), 'typeface');
    if (!face || face === '+mn-lt') face = info.phType === 'title' || info.phType === 'ctrTitle' ? this.theme.major : this.theme.minor;
    if (face === '+mj-lt') face = this.theme.major;
    if (face === '+mn-lt') face = this.theme.minor;

    const highlight = getChild('highlight');
    return {
      size: sizePt * EMU_PER_PT * this.k,
      color,
      font: `"${face}", "Roboto", "Noto Sans", sans-serif`,
      bold: get('b') === '1' || get('b') === 'true',
      italic: get('i') === '1' || get('i') === 'true',
      underline: get('u') && get('u') !== 'none',
      strike: get('strike') && get('strike') !== 'noStrike',
      caps: get('cap') === 'all',
      spacing: get('spc') ? Number(get('spc')) / 100 * EMU_PER_PT * this.k : 0,
      baseline: Number(get('baseline') || 0),
      highlight: highlight ? this.colorFrom(highlight, ctx) : null,
    };
  }

  autoNumber(type, n) {
    const roman = (v) => {
      const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
      let out = '';
      for (const [val, s] of map) while (v >= val) { out += s; v -= val; }
      return out;
    };
    const alpha = (v) => String.fromCharCode(96 + (((v - 1) % 26) + 1));
    let core = String(n);
    if (type.startsWith('alphaLc')) core = alpha(n);
    else if (type.startsWith('alphaUc')) core = alpha(n).toUpperCase();
    else if (type.startsWith('romanLc')) core = roman(n);
    else if (type.startsWith('romanUc')) core = roman(n).toUpperCase();
    if (type.endsWith('ParenBoth')) return `(${core})`;
    if (type.endsWith('ParenR')) return `${core})`;
    if (type.endsWith('Plain')) return core;
    return `${core}.`;
  }
}
