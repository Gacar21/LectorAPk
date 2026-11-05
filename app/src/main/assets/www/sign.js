// Firmar PDFs: firmas guardadas, dibujar con el dedo o usar un PNG,
// colocarlas sobre la página y guardar un PDF nuevo con pdf-lib.

import { $, toast, askConfirm, saveBytes, baseName } from './ui.js';

const STORE_KEY = 'lectordocs.firmas';
const MAX_SAVED = 6;
const INK_COLORS = { negro: '#111111', azul: '#1A3D8F' };

let session = null; // { doc, stamps: [] }

/* ================= firmas guardadas ================= */

const readSaved = () => {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY) || '[]');
  } catch {
    return [];
  }
};
const writeSaved = (list) => {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(list.slice(0, MAX_SAVED)));
  } catch {
    toast('No hay espacio para guardar más firmas');
  }
};
const rememberSignature = (dataUrl) => writeSaved([dataUrl, ...readSaved().filter((s) => s !== dataUrl)]);

/* ================= procesar imágenes ================= */

const loadImage = (src) =>
  new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('No se pudo leer la imagen'));
    img.src = src;
  });

/** Recorta los bordes vacíos y devuelve un PNG (dataURL). */
function trimCanvas(canvas) {
  const ctx = canvas.getContext('2d');
  const { width, height } = canvas;
  const data = ctx.getImageData(0, 0, width, height).data;
  let top = height, left = width, right = -1, bottom = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 10) {
        if (x < left) left = x;
        if (x > right) right = x;
        if (y < top) top = y;
        if (y > bottom) bottom = y;
      }
    }
  }
  if (right < 0) return null;
  const pad = Math.round(Math.max(width, height) * 0.01);
  left = Math.max(0, left - pad);
  top = Math.max(0, top - pad);
  right = Math.min(width - 1, right + pad);
  bottom = Math.min(height - 1, bottom + pad);
  const out = document.createElement('canvas');
  out.width = right - left + 1;
  out.height = bottom - top + 1;
  out.getContext('2d').drawImage(canvas, left, top, out.width, out.height, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}

/**
 * Prepara una imagen elegida por la persona: la reduce, y si no tiene
 * transparencia (foto de una firma en papel) vuelve transparente el fondo claro.
 */
async function prepareImage(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url);
    const scale = Math.min(1, 1000 / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(img.naturalWidth * scale));
    c.height = Math.max(1, Math.round(img.naturalHeight * scale));
    const ctx = c.getContext('2d');
    ctx.drawImage(img, 0, 0, c.width, c.height);
    const imgData = ctx.getImageData(0, 0, c.width, c.height);
    const d = imgData.data;
    let opaque = true;
    for (let i = 3; i < d.length; i += 4) {
      if (d[i] < 250) {
        opaque = false;
        break;
      }
    }
    if (opaque) {
      for (let i = 0; i < d.length; i += 4) {
        const light = Math.min(d[i], d[i + 1], d[i + 2]);
        if (light > 215) d[i + 3] = 0;
        else if (light > 160) d[i + 3] = Math.round(((215 - light) / 55) * 255);
      }
      ctx.putImageData(imgData, 0, 0);
    }
    return trimCanvas(c) || c.toDataURL('image/png');
  } finally {
    URL.revokeObjectURL(url);
  }
}

/* ================= panel para elegir firma ================= */

function renderSavedList() {
  const list = readSaved();
  const box = $('sigList');
  box.replaceChildren();
  $('sigEmpty').hidden = list.length > 0;
  for (const dataUrl of list) {
    const item = document.createElement('div');
    item.className = 'sig-item';
    const use = document.createElement('button');
    use.type = 'button';
    use.className = 'sig-use';
    use.setAttribute('aria-label', 'Usar esta firma');
    const img = document.createElement('img');
    img.src = dataUrl;
    img.alt = '';
    use.appendChild(img);
    use.addEventListener('click', () => {
      closeSheet();
      placeStamp(dataUrl);
    });
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'sig-forget';
    del.setAttribute('aria-label', 'Borrar firma guardada');
    del.textContent = '×';
    del.addEventListener('click', async () => {
      if (await askConfirm('¿Borrar esta firma guardada?', 'Borrar')) {
        writeSaved(readSaved().filter((s) => s !== dataUrl));
        renderSavedList();
      }
    });
    item.append(use, del);
    box.appendChild(item);
  }
}

const openSheet = () => {
  renderSavedList();
  $('signSheet').hidden = false;
};
const closeSheet = () => {
  $('signSheet').hidden = true;
};

$('btnSigCancel').addEventListener('click', closeSheet);
$('signSheet').addEventListener('click', (e) => {
  if (e.target === $('signSheet')) closeSheet();
});
$('btnSigDraw').addEventListener('click', () => {
  closeSheet();
  openPad();
});
$('btnSigImage').addEventListener('click', () => $('sigFile').click());
$('sigFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const dataUrl = await prepareImage(file);
    rememberSignature(dataUrl);
    closeSheet();
    placeStamp(dataUrl);
  } catch (err) {
    toast(err.message || 'No se pudo usar esa imagen');
  }
});

/* ================= dibujar con el dedo ================= */

let pad = null;

function openPad() {
  $('padModal').hidden = false;
  const canvas = $('padCanvas');
  const rect = canvas.getBoundingClientRect();
  const dpr = Math.min(window.devicePixelRatio || 1, 3);
  canvas.width = Math.round(rect.width * dpr);
  canvas.height = Math.round(rect.height * dpr);
  const ctx = canvas.getContext('2d');
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  pad = { canvas, ctx, dpr, drawing: false, last: null, mid: null, empty: true, color: pad?.color || INK_COLORS.negro };
  document.querySelectorAll('.ink').forEach((b) => b.classList.toggle('on', INK_COLORS[b.dataset.ink] === pad.color));
}

const closePad = () => {
  $('padModal').hidden = true;
  pad = null;
};

const padPoint = (e) => {
  const r = pad.canvas.getBoundingClientRect();
  return { x: (e.clientX - r.left) * pad.dpr, y: (e.clientY - r.top) * pad.dpr, p: e.pressure || 0.5 };
};

$('padCanvas').addEventListener('pointerdown', (e) => {
  if (!pad) return;
  e.preventDefault();
  pad.canvas.setPointerCapture(e.pointerId);
  pad.drawing = true;
  const pt = padPoint(e);
  pad.last = pt;
  pad.mid = pt;
  pad.ctx.fillStyle = pad.color;
  pad.ctx.beginPath();
  pad.ctx.arc(pt.x, pt.y, 1.4 * pad.dpr, 0, Math.PI * 2);
  pad.ctx.fill();
  pad.empty = false;
  $('padHint').hidden = true;
});
$('padCanvas').addEventListener('pointermove', (e) => {
  if (!pad || !pad.drawing) return;
  e.preventDefault();
  const pt = padPoint(e);
  const mid = { x: (pad.last.x + pt.x) / 2, y: (pad.last.y + pt.y) / 2 };
  const { ctx } = pad;
  ctx.strokeStyle = pad.color;
  ctx.lineWidth = (2.2 + (e.pointerType === 'pen' ? pt.p * 2 : 0.8)) * pad.dpr;
  ctx.beginPath();
  ctx.moveTo(pad.mid.x, pad.mid.y);
  ctx.quadraticCurveTo(pad.last.x, pad.last.y, mid.x, mid.y);
  ctx.stroke();
  pad.last = pt;
  pad.mid = mid;
});
const stopDraw = () => {
  if (pad) pad.drawing = false;
};
$('padCanvas').addEventListener('pointerup', stopDraw);
$('padCanvas').addEventListener('pointercancel', stopDraw);

document.querySelectorAll('.ink').forEach((b) =>
  b.addEventListener('click', () => {
    if (!pad) return;
    pad.color = INK_COLORS[b.dataset.ink];
    document.querySelectorAll('.ink').forEach((x) => x.classList.toggle('on', x === b));
  }),
);
$('btnPadClear').addEventListener('click', () => {
  if (!pad) return;
  pad.ctx.clearRect(0, 0, pad.canvas.width, pad.canvas.height);
  pad.empty = true;
  $('padHint').hidden = false;
});
$('btnPadCancel').addEventListener('click', closePad);
$('btnPadUse').addEventListener('click', () => {
  if (!pad || pad.empty) {
    toast('Dibuja tu firma primero');
    return;
  }
  const dataUrl = trimCanvas(pad.canvas);
  closePad();
  if (!dataUrl) return;
  rememberSignature(dataUrl);
  placeStamp(dataUrl);
});

/* ================= colocar firmas sobre la página ================= */

function visiblePageIndex() {
  const pages = session.doc.pages;
  const box = $('scroller').getBoundingClientRect();
  const mid = box.top + box.height / 2;
  let best = 0, bestDist = Infinity;
  pages.forEach((el, i) => {
    const r = el.getBoundingClientRect();
    const d = mid < r.top ? r.top - mid : mid > r.bottom ? mid - r.bottom : 0;
    if (d < bestDist) {
      bestDist = d;
      best = i;
    }
  });
  return best;
}

async function placeStamp(dataUrl) {
  if (!session) return;
  const img = await loadImage(dataUrl);
  const pageIndex = visiblePageIndex();
  const pageEl = session.doc.pages[pageIndex];
  const pr = pageEl.getBoundingClientRect();
  const box = $('scroller').getBoundingClientRect();
  const pageAspect = pr.height / pr.width;
  const w = 0.32;
  const h = (w * img.naturalHeight) / img.naturalWidth / pageAspect;
  // centro de la parte visible de la página
  const visTop = Math.max(pr.top, box.top), visBottom = Math.min(pr.bottom, box.bottom);
  const cy = ((visTop + visBottom) / 2 - pr.top) / pr.height;
  const stamp = {
    page: pageIndex,
    dataUrl,
    ratio: img.naturalWidth / img.naturalHeight,
    x: 0.5 - w / 2,
    y: Math.min(Math.max(cy - h / 2, 0), 1 - h),
    w,
    h,
    el: null,
  };
  stamp.el = buildStampElement(stamp, pageEl);
  session.stamps.push(stamp);
  updateBar();
}

function buildStampElement(stamp, pageEl) {
  const el = document.createElement('div');
  el.className = 'sig-stamp';
  const img = document.createElement('img');
  img.src = stamp.dataUrl;
  img.alt = 'Firma';
  img.draggable = false;
  const del = document.createElement('button');
  del.type = 'button';
  del.className = 'sig-del';
  del.setAttribute('aria-label', 'Quitar firma');
  del.textContent = '×';
  const handle = document.createElement('span');
  handle.className = 'sig-resize';
  el.append(img, del, handle);
  pageEl.appendChild(el);

  const apply = () => {
    el.style.left = stamp.x * 100 + '%';
    el.style.top = stamp.y * 100 + '%';
    el.style.width = stamp.w * 100 + '%';
    el.style.height = stamp.h * 100 + '%';
  };
  apply();

  del.addEventListener('pointerdown', (e) => e.stopPropagation());
  del.addEventListener('click', () => {
    el.remove();
    session.stamps = session.stamps.filter((s) => s !== stamp);
    updateBar();
  });

  let drag = null;
  const start = (mode) => (e) => {
    e.preventDefault();
    e.stopPropagation();
    el.setPointerCapture(e.pointerId);
    const r = pageEl.getBoundingClientRect();
    drag = { mode, sx: e.clientX, sy: e.clientY, x: stamp.x, y: stamp.y, w: stamp.w, h: stamp.h, pw: r.width, ph: r.height };
    el.classList.add('active');
  };
  el.addEventListener('pointerdown', start('move'));
  handle.addEventListener('pointerdown', start('resize'));
  el.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const dx = (e.clientX - drag.sx) / drag.pw;
    const dy = (e.clientY - drag.sy) / drag.ph;
    if (drag.mode === 'move') {
      stamp.x = Math.min(Math.max(drag.x + dx, 0), 1 - stamp.w);
      stamp.y = Math.min(Math.max(drag.y + dy, 0), 1 - stamp.h);
    } else {
      const aspect = drag.ph / drag.pw;
      stamp.w = Math.min(Math.max(drag.w + dx, 0.06), 1 - stamp.x);
      stamp.h = (stamp.w / stamp.ratio) / aspect;
      if (stamp.y + stamp.h > 1) {
        stamp.h = 1 - stamp.y;
        stamp.w = stamp.h * aspect * stamp.ratio;
      }
    }
    apply();
  });
  const end = () => {
    drag = null;
    el.classList.remove('active');
  };
  el.addEventListener('pointerup', end);
  el.addEventListener('pointercancel', end);
  return el;
}

function updateBar() {
  const active = !!session;
  $('signBar').hidden = !active;
  $('dock').hidden = active;
  if (active) $('btnSignSave').disabled = session.stamps.length === 0;
  $('signCount').textContent = active && session.stamps.length
    ? `${session.stamps.length} firma${session.stamps.length > 1 ? 's' : ''}`
    : 'Toca “Agregar firma”';
}

/* ================= guardar el PDF firmado ================= */

async function saveSigned() {
  if (!session || !session.stamps.length) return;
  const { PDFDocument, degrees } = window.PDFLib;
  const btn = $('btnSignSave');
  btn.disabled = true;
  btn.textContent = 'Guardando…';
  try {
    let pdf;
    try {
      pdf = await PDFDocument.load(session.doc.originalBytes.slice(), { ignoreEncryption: true });
    } catch {
      throw new Error('Este PDF no se puede modificar (puede estar protegido).');
    }
    const embedded = new Map();
    for (const s of session.stamps) {
      if (!embedded.has(s.dataUrl)) embedded.set(s.dataUrl, await pdf.embedPng(s.dataUrl));
      const image = embedded.get(s.dataUrl);
      const page = pdf.getPage(s.page);
      const rot = ((page.getRotation().angle % 360) + 360) % 360;
      const crop = page.getCropBox();
      const w = crop.width, h = crop.height;
      const quarter = rot === 90 || rot === 270;
      const dispW = quarter ? h : w;
      const dispH = quarter ? w : h;
      const dw = s.w * dispW, dh = s.h * dispH;
      // esquina inferior izquierda de la firma en coordenadas de pantalla
      const ax = s.x * dispW, ay = s.y * dispH + dh;
      let px, py;
      if (rot === 90) { px = ay; py = ax; }
      else if (rot === 180) { px = w - ax; py = ay; }
      else if (rot === 270) { px = w - ay; py = h - ax; }
      else { px = ax; py = h - ay; }
      page.drawImage(image, { x: crop.x + px, y: crop.y + py, width: dw, height: dh, rotate: degrees(rot) });
    }
    const bytes = await pdf.save();
    const name = `${baseName(session.doc.name)}_firmado.pdf`;
    const result = await saveBytes(bytes, name, 'application/pdf', true);
    if (result.ok) {
      toast(`Guardado: ${result.name}`);
      endSigning();
    } else if (!result.cancelled) {
      toast('No se pudo guardar el archivo');
    }
  } catch (err) {
    console.error(err);
    toast(err.message || 'No se pudo firmar el PDF');
  } finally {
    btn.textContent = 'Guardar firmado';
    if (session) btn.disabled = session.stamps.length === 0;
  }
}

$('btnSignAdd').addEventListener('click', openSheet);
$('btnSignSave').addEventListener('click', saveSigned);
$('btnSignCancel').addEventListener('click', async () => {
  if (session && session.stamps.length && !(await askConfirm('¿Descartar las firmas colocadas?', 'Descartar'))) return;
  endSigning();
});

/* ================= API para app.js ================= */

export function startSigning(doc) {
  if (!doc || !doc.originalBytes) return;
  if (!session) session = { doc, stamps: [] };
  updateBar();
  openSheet();
}

export function endSigning() {
  if (session) session.stamps.forEach((s) => s.el.remove());
  session = null;
  closeSheet();
  closePad();
  updateBar();
}

export const isSigning = () => !!session;

/** Botón atrás de Android: cierra lo que esté abierto. Devuelve true si consumió el evento. */
export function signBack() {
  if (!$('padModal').hidden) {
    closePad();
    return true;
  }
  if (!$('signSheet').hidden) {
    closeSheet();
    return true;
  }
  if (session) {
    endSigning();
    return true;
  }
  return false;
}
