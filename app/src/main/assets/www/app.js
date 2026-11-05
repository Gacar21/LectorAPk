import * as pdfjsLib from './lib/pdf.min.mjs';
import { PptxRenderer } from './pptx.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('lib/pdf.worker.min.mjs', location.href).href;

const native = window.Android || null;
const $ = (id) => document.getElementById(id);

const ZOOM_MIN = 0.5;
const ZOOM_MAX = 4;
const ZOOM_STEP = 1.25;
const RECENT_KEY = 'lectordocs.recientes';
const RECENT_MAX = 20;
const KIND_LABEL = { pdf: 'PDF', docx: 'WORD', pptx: 'POWERPOINT' };

const scroller = $('scroller');
const content = $('content');

/** Estado del documento abierto */
let current = null; // { kind, name, uri, pages: HTMLElement[], layout(), dispose() }
let zoom = 1;
let openToken = 0;

/* =================== utilidades de interfaz =================== */

let toastTimer;
const toast = (msg) => {
  const t = $('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), 2600);
};

let dialogResolve = null;
const askText = (title, placeholder = '', type = 'text') =>
  new Promise((resolve) => {
    $('dialogTitle').textContent = title;
    const input = $('dialogInput');
    input.type = type;
    input.inputMode = type === 'number' ? 'numeric' : 'text';
    input.placeholder = placeholder;
    input.value = '';
    $('dialog').hidden = false;
    dialogResolve = resolve;
    setTimeout(() => input.focus(), 50);
  });
const closeDialog = (value) => {
  $('dialog').hidden = true;
  const r = dialogResolve;
  dialogResolve = null;
  r && r(value);
};
$('dialogForm').addEventListener('submit', (e) => {
  e.preventDefault();
  closeDialog($('dialogInput').value);
});
$('dialogCancel').addEventListener('click', () => closeDialog(null));

const setLoading = (text) => {
  $('loading').hidden = text === null;
  if (text !== null) $('loadingText').textContent = text;
};

const showError = (msg) => {
  setLoading(null);
  $('errorText').textContent = msg;
  $('error').hidden = false;
};

const formatSize = (bytes) => {
  if (!bytes || bytes < 0) return '';
  if (bytes < 1024 * 1024) return Math.max(1, Math.round(bytes / 1024)) + ' KB';
  return (bytes / 1024 / 1024).toFixed(1).replace('.', ',') + ' MB';
};

const timeAgo = (ts) => {
  const min = Math.round((Date.now() - ts) / 60000);
  if (min < 1) return 'hace un momento';
  if (min < 60) return `hace ${min} min`;
  const h = Math.round(min / 60);
  if (h < 24) return `hace ${h} h`;
  const d = Math.round(h / 24);
  if (d === 1) return 'ayer';
  if (d < 30) return `hace ${d} días`;
  return new Date(ts).toLocaleDateString('es');
};

/* =================== recientes =================== */

const readRecent = () => {
  try {
    return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
  } catch {
    return [];
  }
};
const writeRecent = (list) => {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, RECENT_MAX)));
  } catch { /* almacenamiento no disponible */ }
};
const addRecent = (item) => {
  if (!item.uri) return;
  writeRecent([{ ...item, at: Date.now() }, ...readRecent().filter((r) => r.uri !== item.uri)]);
};
const removeRecent = (uri) => {
  writeRecent(readRecent().filter((r) => r.uri !== uri));
  renderRecent();
};

const renderRecent = () => {
  const list = readRecent();
  const ul = $('recentList');
  ul.replaceChildren();
  $('recentEmpty').hidden = list.length > 0;
  $('btnClearRecent').hidden = list.length === 0;
  for (const r of list) {
    const li = document.createElement('li');
    li.className = 'recent-item';

    const open = document.createElement('button');
    open.className = 'recent-open';
    open.type = 'button';
    const badge = document.createElement('span');
    badge.className = `badge ${r.kind}`;
    badge.textContent = r.kind === 'pptx' ? 'PPT' : r.kind === 'docx' ? 'DOC' : 'PDF';
    const text = document.createElement('span');
    text.className = 'recent-text';
    const name = document.createElement('span');
    name.className = 'recent-name';
    name.textContent = r.name;
    const meta = document.createElement('span');
    meta.className = 'recent-meta';
    meta.textContent = [timeAgo(r.at), formatSize(r.size)].filter(Boolean).join(' · ');
    text.append(name, meta);
    open.append(badge, text);
    open.addEventListener('click', () => {
      if (!native || !native.openUri(r.uri)) {
        toast('Ya no hay acceso a ese archivo. Ábrelo de nuevo desde el botón.');
        removeRecent(r.uri);
      }
    });

    const del = document.createElement('button');
    del.className = 'recent-remove';
    del.type = 'button';
    del.setAttribute('aria-label', 'Quitar de recientes');
    del.textContent = '×';
    del.addEventListener('click', () => removeRecent(r.uri));

    li.append(open, del);
    ul.appendChild(li);
  }
};
$('btnClearRecent').addEventListener('click', () => {
  writeRecent([]);
  renderRecent();
});

/* =================== abrir archivos =================== */

const pickFile = () => {
  if (native) native.pickFile();
  else $('fileInput').click();
};
$('btnOpen').addEventListener('click', pickFile);
document.querySelectorAll('.fmt').forEach((b) => b.addEventListener('click', pickFile));

// Modo navegador (para pruebas fuera de Android)
$('fileInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  openDocument({ name: file.name, size: file.size, mime: file.type, uri: null }, () => file.arrayBuffer());
  e.target.value = '';
});

// Llamado desde Java cuando hay un documento listo en /doc/actual
window.onNativeFile = (info) => {
  openDocument(info, async () => {
    const res = await fetch('/doc/actual?t=' + Date.now(), { cache: 'no-store' });
    if (!res.ok) throw new Error('No se pudo leer el archivo');
    return res.arrayBuffer();
  });
};

const kindFromName = (name = '', mime = '') => {
  const ext = name.toLowerCase().split('.').pop();
  if (ext === 'pdf' || mime === 'application/pdf') return 'pdf';
  if (ext === 'docx' || mime.includes('wordprocessingml')) return 'docx';
  if (ext === 'pptx' || ext === 'ppsx' || mime.includes('presentationml')) return 'pptx';
  return null;
};

const sniffKind = async (buffer) => {
  const head = new Uint8Array(buffer.slice(0, 8));
  if (head[0] === 0x25 && head[1] === 0x50 && head[2] === 0x44 && head[3] === 0x46) return 'pdf';
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) return 'legacy';
  if (head[0] === 0x50 && head[1] === 0x4b) {
    const zip = await window.JSZip.loadAsync(buffer);
    if (zip.file('word/document.xml')) return 'docx';
    if (zip.file('ppt/presentation.xml')) return 'pptx';
  }
  return null;
};

async function openDocument(info, loadBuffer) {
  const token = ++openToken;
  closeCurrent();
  zoom = 1;
  updateZoomLabel();
  showViewer(info.name, kindFromName(info.name, info.mime || ''));
  setLoading('Abriendo…');

  try {
    const buffer = await loadBuffer();
    if (token !== openToken) return;
    let kind = await sniffKind(buffer).catch(() => null);
    if (kind === 'legacy') {
      throw new Error('Es un archivo de Office antiguo (.doc o .ppt). Guárdalo como .docx o .pptx para abrirlo aquí.');
    }
    kind = kind || kindFromName(info.name, info.mime || '');
    if (!kind) throw new Error('Este tipo de archivo no es compatible. Puedes abrir PDF, Word (.docx) y PowerPoint (.pptx).');
    setKind(kind);

    const opener = { pdf: openPdf, docx: openDocx, pptx: openPptx }[kind];
    current = await opener(buffer, token);
    if (!current || token !== openToken) return;
    current.kind = kind;
    current.name = info.name;
    current.uri = info.uri;
    updatePageInfo();
    addRecent({ uri: info.uri, name: info.name, kind, size: info.size || buffer.byteLength });
  } catch (err) {
    if (token !== openToken) return;
    console.error(err);
    showError(err && err.message ? err.message : 'El archivo parece dañado.');
  }
}

function closeCurrent() {
  if (current && current.dispose) current.dispose();
  current = null;
  content.replaceChildren();
  content.className = 'content';
  scroller.scrollTo(0, 0);
  $('error').hidden = true;
}

function showViewer(name, kind) {
  $('home').hidden = true;
  $('viewer').hidden = false;
  $('docTitle').textContent = name || 'Documento';
  setKind(kind);
  $('pageInfo').textContent = '–';
}

function setKind(kind) {
  const k = $('docKind');
  k.className = 'kind ' + (kind || '');
  k.textContent = KIND_LABEL[kind] || '';
}

function showHome() {
  openToken++;
  closeCurrent();
  setLoading(null);
  $('viewer').hidden = true;
  $('home').hidden = false;
  renderRecent();
}

window.handleBack = () => {
  if (!$('dialog').hidden) {
    closeDialog(null);
    return true;
  }
  if (!$('viewer').hidden) {
    showHome();
    return true;
  }
  return false;
};
$('btnBack').addEventListener('click', showHome);
$('btnErrorBack').addEventListener('click', showHome);

const baseWidth = () => Math.max(scroller.clientWidth - 16, 200);

/* =================== PDF =================== */

async function openPdf(buffer, token) {
  const task = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    standardFontDataUrl: new URL('lib/standard_fonts/', location.href).href,
    isEvalSupported: false,
  });
  task.onPassword = async (update, reason) => {
    setLoading(null);
    const pass = await askText(reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD ? 'Contraseña incorrecta. Intenta de nuevo.' : 'Este PDF está protegido con contraseña.', 'Contraseña', 'password');
    if (pass === null) {
      task.destroy();
      showHome();
      return;
    }
    setLoading('Abriendo…');
    update(pass);
  };

  let doc;
  try {
    doc = await task.promise;
  } catch (err) {
    if (err && err.name === 'PasswordException') return null;
    throw new Error('No se pudo leer el PDF. Puede estar dañado.');
  }
  if (token !== openToken) {
    doc.destroy();
    return null;
  }

  const first = await doc.getPage(1);
  const vp = first.getViewport({ scale: 1 });
  const defaultAspect = vp.height / vp.width;

  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const el = document.createElement('div');
    el.className = 'pdf-page';
    el.dataset.page = i;
    const label = document.createElement('span');
    label.className = 'num';
    label.textContent = i;
    el.appendChild(label);
    pages.push({ el, n: i, aspect: defaultAspect, renderedAt: 0, task: null, canvas: null });
    content.appendChild(el);
  }

  const layout = () => {
    const w = Math.round(baseWidth() * zoom);
    for (const p of pages) {
      p.el.style.width = w + 'px';
      p.el.style.height = Math.round(w * p.aspect) + 'px';
    }
    pages.forEach((p) => p.visible && renderPage(p));
  };

  const MAX_PIXELS = 16_000_000;
  const renderPage = async (p) => {
    const targetZoom = zoom;
    if (p.renderedAt === targetZoom && p.canvas) return;
    if (p.task) p.task.cancel();
    try {
      const page = await doc.getPage(p.n);
      const base = page.getViewport({ scale: 1 });
      const aspect = base.height / base.width;
      if (Math.abs(aspect - p.aspect) > 0.001) {
        p.aspect = aspect;
        p.el.style.height = Math.round(parseFloat(p.el.style.width) * aspect) + 'px';
      }
      const cssWidth = parseFloat(p.el.style.width);
      const dpr = Math.min(window.devicePixelRatio || 1, 3);
      let scale = (cssWidth / base.width) * dpr;
      if (base.width * base.height * scale * scale > MAX_PIXELS) scale = Math.sqrt(MAX_PIXELS / (base.width * base.height));
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      p.task = page.render({ canvasContext: canvas.getContext('2d', { alpha: false }), viewport });
      await p.task.promise;
      p.task = null;
      if (!p.visible || zoom !== targetZoom) {
        canvas.width = canvas.height = 0;
        return;
      }
      if (p.canvas) p.canvas.remove();
      p.canvas = canvas;
      p.renderedAt = targetZoom;
      p.el.appendChild(canvas);
    } catch (err) {
      if (err && err.name !== 'RenderingCancelledException') console.warn('Página', p.n, err);
    }
  };

  const freePage = (p) => {
    if (p.task) {
      p.task.cancel();
      p.task = null;
    }
    if (p.canvas) {
      p.canvas.width = p.canvas.height = 0;
      p.canvas.remove();
      p.canvas = null;
    }
    p.renderedAt = 0;
  };

  const byEl = new Map(pages.map((p) => [p.el, p]));
  const observer = new IntersectionObserver((entries) => {
    for (const e of entries) {
      const p = byEl.get(e.target);
      p.visible = e.isIntersecting;
      if (p.visible) renderPage(p);
      else freePage(p);
    }
  }, { root: scroller, rootMargin: '1200px 0px' });

  layout();
  pages.forEach((p) => observer.observe(p.el));
  setLoading(null);

  return {
    pages: pages.map((p) => p.el),
    layout,
    dispose() {
      observer.disconnect();
      pages.forEach(freePage);
      doc.destroy();
    },
  };
}

/* =================== Word =================== */

async function openDocx(buffer, token) {
  const host = document.createElement('div');
  host.className = 'docx-host';
  const styleHost = document.createElement('div');
  content.append(styleHost, host);

  await window.docx.renderAsync(buffer, host, styleHost, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    experimental: true,
    useBase64URL: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
  });
  if (token !== openToken) return null;
  fixSymbolBullets(styleHost);

  const wrapper = host.querySelector('.docx-wrapper') || host;
  const sections = [...wrapper.querySelectorAll(':scope > section')];
  wrapper.style.zoom = '1';
  const natural = Math.max(...sections.map((s) => s.offsetWidth), 300);

  const layout = () => {
    wrapper.style.zoom = String((baseWidth() / natural) * zoom);
  };
  layout();
  setLoading(null);

  // Los enlaces del documento se abren fuera de la app
  host.addEventListener('click', (e) => {
    const a = e.target.closest('a[href]');
    if (a && a.getAttribute('href').startsWith('#')) {
      e.preventDefault();
      const target = host.querySelector(`[id="${CSS.escape(a.getAttribute('href').slice(1))}"]`);
      target && target.scrollIntoView();
    }
  });

  return { pages: sections, layout, dispose() {} };
}

// Word guarda las viñetas como caracteres de las fuentes Symbol/Wingdings,
// que Android no trae. Se cambian por sus equivalentes Unicode.
const SYMBOL_BULLETS = {
  '': '•', '': '▪', '': '□', '': '■', '': '●',
  '': '➢', '': '✓', '': '❖', '': '➔', '': '➔',
  '': '❑', '': '▫', '': '–',
};
function fixSymbolBullets(styleHost) {
  for (const style of styleHost.querySelectorAll('style')) {
    let css = style.textContent;
    if (!/[-]|Symbol|Wingdings/.test(css)) continue;
    css = css.replace(/[-]/g, (ch) => SYMBOL_BULLETS[ch] || '•');
    css = css.replace(/font-family:\s*["']?(Symbol|Wingdings[^;"']*)["']?;?/g, '');
    style.textContent = css;
  }
}

/* =================== PowerPoint =================== */

async function openPptx(buffer, token) {
  const renderer = await PptxRenderer.load(buffer);
  if (token !== openToken) {
    renderer.dispose();
    return null;
  }
  const total = renderer.slideCount;
  if (!total) throw new Error('La presentación no tiene diapositivas.');

  const frames = [];
  const doc = {
    pages: frames,
    layout() {
      const w = Math.round(baseWidth() * zoom);
      const scale = w / renderer.width;
      for (const f of frames) {
        f.style.width = w + 'px';
        f.style.height = Math.round(renderer.height * scale) + 'px';
        if (f.firstChild) f.firstChild.style.transform = `scale(${scale})`;
      }
    },
    dispose() {
      renderer.dispose();
    },
  };

  // Marcos vacíos para todas las diapositivas; se llenan en orden
  for (let i = 0; i < total; i++) {
    const frame = document.createElement('div');
    frame.className = 'slide-frame';
    const label = document.createElement('p');
    label.className = 'slide-label';
    label.textContent = `${i + 1} / ${total}`;
    content.append(frame, label);
    frames.push(frame);
  }
  doc.layout();

  for (let i = 0; i < total; i++) {
    if (token !== openToken) return null;
    let slide;
    try {
      slide = await renderer.renderSlide(i);
    } catch (err) {
      console.warn('Diapositiva', i + 1, err);
      slide = document.createElement('div');
      slide.className = 'pptx-slide';
      slide.style.width = renderer.width + 'px';
      slide.style.height = renderer.height + 'px';
      slide.appendChild(Object.assign(document.createElement('div'), { className: 'pptx-unsupported', textContent: 'No se pudo mostrar esta diapositiva' }));
    }
    frames[i].replaceChildren(slide);
    if (i === 0) setLoading(null);
    doc.layout();
    // Cede el hilo para que la interfaz siga respondiendo
    await new Promise((r) => setTimeout(r, 0));
  }
  return doc;
}

/* =================== zoom, página, modo nocturno =================== */

function updateZoomLabel() {
  $('zoomInfo').textContent = Math.round(zoom * 100) + '%';
}

function setZoom(next, anchorX = scroller.clientWidth / 2, anchorY = scroller.clientHeight / 2) {
  next = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, next));
  if (!current || Math.abs(next - zoom) < 0.001) return;
  const ratio = next / zoom;
  const px = scroller.scrollLeft + anchorX;
  const py = scroller.scrollTop + anchorY;
  zoom = next;
  current.layout();
  scroller.scrollLeft = px * ratio - anchorX;
  scroller.scrollTop = py * ratio - anchorY;
  updateZoomLabel();
}

$('btnZoomIn').addEventListener('click', () => setZoom(zoom * ZOOM_STEP));
$('btnZoomOut').addEventListener('click', () => setZoom(zoom / ZOOM_STEP));

// Pellizcar para hacer zoom
let pinch = null;
const dist = (t) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
scroller.addEventListener('touchstart', (e) => {
  if (e.touches.length !== 2 || !current) return;
  const rect = scroller.getBoundingClientRect();
  const cx = (e.touches[0].clientX + e.touches[1].clientX) / 2 - rect.left;
  const cy = (e.touches[0].clientY + e.touches[1].clientY) / 2 - rect.top;
  pinch = { start: dist(e.touches), ratio: 1, cx, cy };
  content.style.transformOrigin = `${scroller.scrollLeft + cx}px ${scroller.scrollTop + cy}px`;
  content.classList.add('pinching');
}, { passive: true });
scroller.addEventListener('touchmove', (e) => {
  if (!pinch || e.touches.length !== 2) return;
  e.preventDefault();
  const ratio = dist(e.touches) / pinch.start;
  pinch.ratio = Math.min(ZOOM_MAX / zoom, Math.max(ZOOM_MIN / zoom, ratio));
  content.style.transform = `scale(${pinch.ratio})`;
}, { passive: false });
const endPinch = () => {
  if (!pinch) return;
  const { ratio, cx, cy } = pinch;
  pinch = null;
  content.style.transform = '';
  content.classList.remove('pinching');
  setZoom(zoom * ratio, cx, cy);
};
scroller.addEventListener('touchend', (e) => e.touches.length < 2 && endPinch());
scroller.addEventListener('touchcancel', endPinch);

// Doble toque: alterna entre ajustar al ancho y 200 %
let lastTap = 0;
scroller.addEventListener('touchend', (e) => {
  if (e.changedTouches.length !== 1 || e.touches.length) return;
  const now = Date.now();
  if (now - lastTap < 280) {
    const rect = scroller.getBoundingClientRect();
    const t = e.changedTouches[0];
    setZoom(zoom > 1.05 ? 1 : 2, t.clientX - rect.left, t.clientY - rect.top);
    lastTap = 0;
  } else lastTap = now;
});

function currentPageIndex() {
  const pages = current?.pages || [];
  if (!pages.length) return -1;
  const top = scroller.getBoundingClientRect().top + scroller.clientHeight * 0.35;
  let lo = 0, hi = pages.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (pages[mid].getBoundingClientRect().top <= top) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

function updatePageInfo() {
  const pages = current?.pages || [];
  $('pageInfo').textContent = pages.length ? `${currentPageIndex() + 1} / ${pages.length}` : '–';
}

let scrollQueued = false;
scroller.addEventListener('scroll', () => {
  if (scrollQueued) return;
  scrollQueued = true;
  requestAnimationFrame(() => {
    scrollQueued = false;
    updatePageInfo();
  });
}, { passive: true });

$('btnPage').addEventListener('click', async () => {
  const pages = current?.pages || [];
  if (pages.length < 2) return;
  const label = current.kind === 'pptx' ? 'diapositiva' : 'página';
  const v = await askText(`Ir a la ${label} (1 – ${pages.length})`, 'Número', 'number');
  const n = parseInt(v, 10);
  if (!n || n < 1 || n > pages.length) return;
  const target = pages[n - 1];
  scroller.scrollTop += target.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 8;
});

let night = false;
try { night = localStorage.getItem('lectordocs.noche') === '1'; } catch { /* sin almacenamiento */ }
const applyNight = () => {
  document.documentElement.dataset.theme = night ? 'dark' : 'light';
  scroller.classList.toggle('night', night);
  $('btnNight').classList.toggle('on', night);
};
$('btnNight').addEventListener('click', () => {
  night = !night;
  try { localStorage.setItem('lectordocs.noche', night ? '1' : '0'); } catch { /* sin almacenamiento */ }
  applyNight();
});

$('btnShare').addEventListener('click', () => {
  if (native && current?.uri) native.share(current.uri);
  else toast('Compartir solo está disponible en la app de Android');
});

// Al girar el teléfono se reajusta al nuevo ancho
let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => current && current.layout(), 150);
});

applyNight();
renderRecent();
