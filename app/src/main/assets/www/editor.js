// Editor de Word: abre un .docx (o uno nuevo), permite dar formato y guarda .docx o PDF.

import { $, toast, askText, askConfirm, saveBytes, printNode, baseName } from './ui.js';
import { htmlToDocx, pageFromSection, LETTER_PAGE } from './docx-export.js';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const COLORS = ['#000000', '#5F6368', '#C8402F', '#D06A22', '#E0A800', '#2E8B57', '#2B6CB0', '#4B3FA8'];
const HIGHLIGHTS = ['transparent', '#FFF176', '#A5F3A1', '#9FD8FF', '#FFB3C1'];

const sheet = $('edSheet');
let state = null; // { name, page, dirty, onClose, onSaved }
let savedRange = null;

/* ---------- selección (se pierde al tocar la barra en el teléfono) ---------- */

document.addEventListener('selectionchange', () => {
  if (!state) return;
  const sel = window.getSelection();
  if (sel.rangeCount && sheet.contains(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange();
});

const restoreSelection = () => {
  sheet.focus({ preventScroll: true });
  if (!savedRange) return;
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(savedRange);
};

const exec = (cmd, value = null) => {
  restoreSelection();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand(cmd, false, value);
  markDirty();
  refreshToolbar();
};

const insertHtml = (html) => {
  restoreSelection();
  document.execCommand('insertHTML', false, html);
  markDirty();
};

const markDirty = () => {
  if (state) state.dirty = true;
};

/* ---------- barra de herramientas ---------- */

// Evita que los botones roben el foco del texto
$('edTools').addEventListener('pointerdown', (e) => {
  if (e.target.closest('button')) e.preventDefault();
});

$('edTools').addEventListener('click', async (e) => {
  const b = e.target.closest('button[data-cmd]');
  if (!b) return;
  const cmd = b.dataset.cmd;
  switch (cmd) {
    case 'color':
    case 'highlight':
      openPalette(cmd, b);
      break;
    case 'table':
      await insertTable();
      break;
    case 'image':
      $('edImageFile').click();
      break;
    case 'pagebreak':
      insertHtml('<hr class="page-break"><p><br></p>');
      break;
    default:
      exec(cmd, b.dataset.value || null);
  }
});

$('edBlock').addEventListener('change', (e) => {
  exec('formatBlock', e.target.value);
  e.target.blur();
});

$('edSize').addEventListener('change', (e) => {
  const pt = e.target.value;
  e.target.value = '';
  if (!pt) return;
  restoreSelection();
  document.execCommand('styleWithCSS', false, false);
  document.execCommand('fontSize', false, '7');
  // el navegador marca el tamaño 7; lo cambiamos por el tamaño real en pt
  sheet.querySelectorAll('font[size="7"]').forEach((f) => {
    const span = document.createElement('span');
    span.style.fontSize = pt + 'pt';
    span.append(...f.childNodes);
    span.querySelectorAll('[style*="font-size"]').forEach((n) => (n.style.fontSize = ''));
    f.replaceWith(span);
  });
  sheet.querySelectorAll('span[style*="xxx-large"]').forEach((s) => (s.style.fontSize = pt + 'pt'));
  markDirty();
});

function openPalette(kind, anchor) {
  const pop = $('edPalette');
  pop.replaceChildren();
  const list = kind === 'color' ? COLORS : HIGHLIGHTS;
  for (const c of list) {
    const sw = document.createElement('button');
    sw.type = 'button';
    sw.className = 'swatch' + (c === 'transparent' ? ' none' : '');
    sw.style.setProperty('--sw', c);
    sw.setAttribute('aria-label', c === 'transparent' ? 'Sin resaltado' : c);
    sw.addEventListener('pointerdown', (e) => e.preventDefault());
    sw.addEventListener('click', () => {
      pop.hidden = true;
      exec(kind === 'color' ? 'foreColor' : 'hiliteColor', c);
    });
    pop.appendChild(sw);
  }
  const r = anchor.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 230)) + 'px';
  pop.style.top = r.bottom + 6 + 'px';
  pop.hidden = false;
}

document.addEventListener('pointerdown', (e) => {
  const pop = $('edPalette');
  if (!pop.hidden && !pop.contains(e.target) && !e.target.closest('[data-cmd="color"],[data-cmd="highlight"]')) pop.hidden = true;
});

async function insertTable() {
  const v = await askText('Tamaño de la tabla (filas × columnas)', 'Ej: 3x3', 'text', '3x3');
  if (!v) return;
  const m = v.match(/(\d+)\s*[x×*,]\s*(\d+)/i);
  if (!m) {
    toast('Escribe filas x columnas, por ejemplo 3x2');
    return;
  }
  const rows = Math.min(Number(m[1]), 30), cols = Math.min(Number(m[2]), 10);
  const row = '<tr>' + '<td><br></td>'.repeat(cols) + '</tr>';
  insertHtml(`<table class="ed-table"><tbody>${row.repeat(rows)}</tbody></table><p><br></p>`);
}

$('edImageFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = () => rej(new Error('No se pudo leer la imagen'));
      i.src = url;
    });
    const scale = Math.min(1, 1400 / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * scale);
    c.height = Math.round(img.naturalHeight * scale);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    const data = c.toDataURL(file.type === 'image/png' ? 'image/png' : 'image/jpeg', 0.88);
    const w = Math.min(c.width, 480);
    insertHtml(`<img src="${data}" style="width:${w}px">`);
  } catch (err) {
    toast(err.message);
  } finally {
    URL.revokeObjectURL(url);
  }
});

/** Marca en la barra el formato del texto donde está el cursor. */
function refreshToolbar() {
  for (const b of document.querySelectorAll('#edTools [data-state]')) {
    let on = false;
    try {
      on = document.queryCommandState(b.dataset.cmd);
    } catch { /* comando sin estado */ }
    b.classList.toggle('on', on);
  }
}
document.addEventListener('selectionchange', () => state && refreshToolbar());

sheet.addEventListener('input', markDirty);

/* ---------- abrir / guardar / cerrar ---------- */

export async function openEditor({ buffer = null, name = 'Documento nuevo.docx', onClose, onSaved }) {
  state = { name: name.toLowerCase().endsWith('.docx') ? name : baseName(name) + '.docx', page: LETTER_PAGE, dirty: false, onClose, onSaved };
  $('edTitle').value = baseName(state.name);
  $('edStyles').replaceChildren();
  sheet.replaceChildren();
  $('home').hidden = true;
  $('viewer').hidden = true;
  $('editor').hidden = false;
  $('edLoading').hidden = !buffer;

  if (buffer) {
    try {
      await window.docx.renderAsync(buffer, sheet, $('edStyles'), {
        className: 'docx',
        inWrapper: false,
        breakPages: false,
        ignoreLastRenderedPageBreak: true,
        renderHeaders: false,
        renderFooters: false,
        renderFootnotes: false,
        renderEndnotes: false,
        renderComments: false,
        useBase64URL: true,
        experimental: true,
      });
      window.fixSymbolBullets && window.fixSymbolBullets($('edStyles'));
      state.page = pageFromSection(sheet.querySelector('section'));
    } catch (err) {
      console.error(err);
      toast('No se pudo abrir el documento para editar');
      closeEditor(true);
      return;
    } finally {
      $('edLoading').hidden = true;
    }
  } else {
    sheet.innerHTML = '<h1><br></h1><p><br></p>';
  }
  savedRange = null;
  setTimeout(() => {
    sheet.focus();
    // cursor al inicio del documento
    const r = document.createRange();
    r.setStart(sheet, 0);
    r.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }, 60);
}

async function closeEditor(force = false) {
  if (!state) return;
  if (!force && state.dirty && !(await askConfirm('Tienes cambios sin guardar. ¿Salir de todas formas?', 'Salir sin guardar', 'Seguir editando'))) return;
  const cb = state.onClose;
  state = null;
  $('editor').hidden = true;
  $('edPalette').hidden = true;
  sheet.replaceChildren();
  $('edStyles').replaceChildren();
  cb && cb();
}

const currentName = () => {
  const t = $('edTitle').value.trim().replace(/[\\/:*?"<>|]+/g, '-') || 'Documento';
  return t;
};

async function saveDocx() {
  if (!state) return;
  const btn = $('btnEdSave');
  btn.disabled = true;
  btn.textContent = 'Guardando…';
  try {
    sheet.blur();
    const bytes = await htmlToDocx(sheet, state.page);
    const result = await saveBytes(bytes, currentName() + '.docx', DOCX_MIME, false);
    if (result.ok) {
      state.dirty = false;
      toast(`Guardado: ${result.name}`);
      state.onSaved && state.onSaved({ name: result.name, uri: result.uri, size: bytes.length });
    } else if (!result.cancelled) {
      toast('No se pudo guardar el archivo');
    }
  } catch (err) {
    console.error(err);
    toast('No se pudo crear el Word: ' + (err.message || err));
  } finally {
    btn.disabled = false;
    btn.textContent = 'Guardar';
  }
}

function exportPdf() {
  if (!state) return;
  const p = state.page;
  const sheetClone = sheet.cloneNode(true);
  sheetClone.removeAttribute('contenteditable');
  sheetClone.removeAttribute('id');
  sheetClone.className = 'ed-sheet print';
  const twToIn = (v) => (v / 1440).toFixed(3) + 'in';
  printNode(sheetClone, {
    jobName: currentName(),
    page: `size: ${twToIn(p.widthTwips)} ${twToIn(p.heightTwips)}; margin: ${twToIn(p.marginTwips.top)} ${twToIn(p.marginTwips.right)} ${twToIn(p.marginTwips.bottom)} ${twToIn(p.marginTwips.left)}`,
    className: 'print-editor',
  });
}

$('btnEdBack').addEventListener('click', () => closeEditor());
$('btnEdSave').addEventListener('click', saveDocx);
$('btnEdPdf').addEventListener('click', exportPdf);
$('edTitle').addEventListener('input', markDirty);

export const isEditing = () => !!state;

export function editorBack() {
  if (!state) return false;
  if (!$('edPalette').hidden) {
    $('edPalette').hidden = true;
    return true;
  }
  closeEditor();
  return true;
}
