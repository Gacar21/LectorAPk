// Utilidades compartidas: avisos, diálogos, guardar archivos e imprimir a PDF.

export const native = window.Android || null;
export const $ = (id) => document.getElementById(id);

let toastTimer;
export const toast = (msg) => {
  const t = $('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (t.hidden = true), 2800);
};

/* ---------- diálogo (texto o confirmación) ---------- */

let dialogResolve = null;
let dialogMode = 'text';

const openDialog = ({ title, placeholder = '', type = 'text', value = '', ok = 'Aceptar', cancel = 'Cancelar', confirm = false }) =>
  new Promise((resolve) => {
    dialogMode = confirm ? 'confirm' : 'text';
    $('dialogTitle').textContent = title;
    const input = $('dialogInput');
    input.hidden = confirm;
    input.type = type;
    input.inputMode = type === 'number' ? 'numeric' : 'text';
    input.placeholder = placeholder;
    input.value = value;
    $('dialogOk').textContent = ok;
    $('dialogCancel').textContent = cancel;
    $('dialog').hidden = false;
    dialogResolve = resolve;
    if (!confirm) setTimeout(() => input.focus(), 50);
  });

export const askText = (title, placeholder = '', type = 'text', value = '') => openDialog({ title, placeholder, type, value });
export const askConfirm = (title, ok = 'Aceptar', cancel = 'Cancelar') => openDialog({ title, ok, cancel, confirm: true });

export const isDialogOpen = () => !$('dialog').hidden;
export const closeDialog = (value) => {
  $('dialog').hidden = true;
  const r = dialogResolve;
  dialogResolve = null;
  if (r) r(dialogMode === 'confirm' ? value !== null : value);
};

$('dialogForm').addEventListener('submit', (e) => {
  e.preventDefault();
  closeDialog($('dialogInput').value);
});
$('dialogCancel').addEventListener('click', () => closeDialog(null));

/* ---------- guardar archivos en el teléfono ---------- */

const toBase64 = (bytes) => {
  let bin = '';
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
  return btoa(bin);
};

/**
 * Abre el selector "Guardar como" de Android con el nombre sugerido.
 * Devuelve { ok, cancelled, name, uri }.
 */
export async function saveBytes(bytes, name, mime, openAfter = false) {
  if (native && native.beginSave) {
    native.beginSave(name);
    const CHUNK = 384 * 1024;
    for (let i = 0; i < bytes.length; i += CHUNK) native.appendSave(toBase64(bytes.subarray(i, i + CHUNK)));
    return new Promise((resolve) => {
      window.onNativeSaved = (info) => {
        window.onNativeSaved = null;
        resolve(info);
      };
      native.finishSave(mime, openAfter);
    });
  }
  // Modo navegador: descarga directa
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  return { ok: true, name, uri: '' };
}

/* ---------- imprimir / guardar como PDF ---------- */

/**
 * Pone `node` en el área de impresión y abre el diálogo de Android
 * (ahí se elige "Guardar como PDF"). `page` es la regla CSS @page.
 */
export function printNode(node, { page = 'size: letter; margin: 0', jobName = 'Documento', className = '' } = {}) {
  let area = $('printArea');
  if (!area) {
    area = document.createElement('div');
    area.id = 'printArea';
    document.body.appendChild(area);
  }
  area.className = className;
  area.replaceChildren(node);
  let pageStyle = $('printPageStyle');
  if (!pageStyle) {
    pageStyle = document.createElement('style');
    pageStyle.id = 'printPageStyle';
    document.head.appendChild(pageStyle);
  }
  pageStyle.textContent = `@page { ${page} }`;
  // Espera a que las imágenes del área estén listas
  const imgs = [...area.querySelectorAll('img')].filter((i) => !i.complete);
  Promise.all(imgs.map((i) => new Promise((r) => { i.onload = i.onerror = r; }))).then(() => {
    if (native && native.printPage) native.printPage(jobName);
    else window.print();
  });
}

export const baseName = (name = 'documento') => name.replace(/\.[^.]+$/, '');
