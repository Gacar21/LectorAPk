// Convierte el contenido del editor (HTML) a un archivo .docx con la librería "docx".
// Lee los estilos ya calculados por el navegador (negrita, color, tamaño, alineación…),
// así funciona igual para documentos nuevos y para los abiertos desde un .docx.

let libPromise = null;
const loadLib = () => {
  if (window.DocxLib) return Promise.resolve(window.DocxLib);
  if (!libPromise) {
    libPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'lib/docxlib.js';
      s.onload = () => resolve(window.DocxLib);
      s.onerror = () => reject(new Error('No se pudo cargar el generador de Word'));
      document.head.appendChild(s);
    });
  }
  return libPromise;
};

const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'LI', 'UL', 'OL', 'TABLE', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'HR', 'FIGURE']);
const PX_TO_TWIP = 15; // 1px = 0.75pt = 15 twips
const MAX_IMAGE_PX = 624; // ancho útil de una página carta con márgenes de 1"

const rgbToHex = (css) => {
  const m = css && css.match(/rgba?\(([^)]+)\)/);
  if (!m) return null;
  const [r, g, b, a] = m[1].split(',').map((v) => parseFloat(v));
  if (a === 0) return null;
  return [r, g, b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
};

const isBlock = (node) => node.nodeType === 1 && BLOCK_TAGS.has(node.tagName);

export async function htmlToDocx(root, page) {
  const D = await loadLib();
  const blockRoot = root;

  /* ---------- texto ---------- */

  const decorations = (el) => {
    let underline = false, strike = false, background = null;
    for (let n = el; n && n !== blockRoot && n.nodeType === 1; n = n.parentElement) {
      const cs = getComputedStyle(n);
      const line = cs.textDecorationLine || cs.textDecoration || '';
      if (line.includes('underline')) underline = true;
      if (line.includes('line-through')) strike = true;
      if (!background) {
        const bg = rgbToHex(cs.backgroundColor);
        if (bg && !['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'TD', 'TH', 'DIV'].includes(n.tagName)) background = bg;
      }
      if (isBlock(n) || n.tagName === 'TD' || n.tagName === 'TH') break;
    }
    return { underline, strike, background };
  };

  const runOptions = (el) => {
    const cs = getComputedStyle(el);
    const deco = decorations(el);
    const opts = {
      bold: parseInt(cs.fontWeight, 10) >= 600 || undefined,
      italics: cs.fontStyle === 'italic' || undefined,
      size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2),
      font: (cs.fontFamily.split(',')[0] || '').replace(/["']/g, '').trim() || undefined,
    };
    const color = rgbToHex(cs.color);
    if (color && color !== '000000') opts.color = color;
    if (deco.underline) opts.underline = {};
    if (deco.strike) opts.strike = true;
    if (deco.background) opts.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: deco.background };
    const va = cs.verticalAlign;
    if (va === 'super') opts.superScript = true;
    if (va === 'sub') opts.subScript = true;
    return opts;
  };

  const imageRun = async (img) => {
    try {
      const res = await fetch(img.src);
      let blob = await res.blob();
      let type = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp' }[blob.type];
      if (!type) {
        // otros formatos (svg, webp) se convierten a PNG
        const c = document.createElement('canvas');
        c.width = img.naturalWidth || 300;
        c.height = img.naturalHeight || 200;
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        blob = await new Promise((r) => c.toBlob(r, 'image/png'));
        type = 'png';
      }
      const data = new Uint8Array(await blob.arrayBuffer());
      const styleW = parseFloat(img.style.width);
      let width = img.style.width.endsWith('pt') ? styleW / 0.75 : img.style.width.endsWith('px') ? styleW : img.naturalWidth;
      if (!width) width = img.getBoundingClientRect().width || 200;
      const ratio = (img.naturalHeight || 1) / (img.naturalWidth || 1);
      width = Math.min(width, MAX_IMAGE_PX);
      return new D.ImageRun({ type, data, transformation: { width: Math.round(width), height: Math.round(width * ratio) } });
    } catch (err) {
      console.warn('Imagen omitida', err);
      return null;
    }
  };

  const inlineRuns = async (node, out) => {
    for (const child of node.childNodes) {
      if (child.nodeType === 3) {
        const text = child.textContent.replace(/ /g, ' ').replace(/[\r\n]+/g, ' ');
        if (text) out.push(new D.TextRun({ text, ...runOptions(child.parentElement) }));
      } else if (child.nodeType === 1) {
        const tag = child.tagName;
        if (tag === 'BR') out.push(new D.TextRun({ break: 1 }));
        else if (tag === 'IMG') {
          const r = await imageRun(child);
          if (r) out.push(r);
        } else if (tag === 'A' && /^https?:|^mailto:/.test(child.getAttribute('href') || '')) {
          const inner = [];
          await inlineRuns(child, inner);
          out.push(new D.ExternalHyperlink({ link: child.getAttribute('href'), children: inner.filter((r) => r instanceof D.TextRun) }));
        } else if (tag === 'STYLE' || tag === 'SCRIPT' || tag === 'UL' || tag === 'OL' || tag === 'TABLE') {
          // las sublistas y tablas se procesan como bloques aparte
          continue;
        } else {
          await inlineRuns(child, out);
        }
      }
    }
    return out;
  };

  /* ---------- párrafos ---------- */

  const headingOf = (el) => {
    const m = el.tagName.match(/^H([1-6])$/);
    let level = m ? Number(m[1]) : null;
    if (!level) {
      const cls = el.className || '';
      const c = cls.match(/(?:heading|t[ií]?tulo)\s*_?(\d)/i);
      if (c) level = Number(c[1]);
      else if (/(^|[\s_])(title|t[ií]?tulo)($|\s)/i.test(cls)) return D.HeadingLevel.TITLE;
    }
    return level ? D.HeadingLevel['HEADING_' + Math.min(level, 6)] : undefined;
  };

  const alignmentOf = (cs) => ({ center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, end: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED })[cs.textAlign];

  const paragraph = async (el, numbering) => {
    const cs = getComputedStyle(el);
    let runs = await inlineRuns(el, []);
    // un párrafo que solo tiene un salto de línea es un párrafo vacío
    if (runs.length === 1 && runs[0] instanceof D.TextRun && el.textContent.trim() === '' && !el.querySelector('img')) runs = [];
    const opts = {
      children: runs,
      heading: headingOf(el),
      alignment: alignmentOf(cs),
      spacing: {
        before: Math.round((parseFloat(cs.marginTop) || 0) * PX_TO_TWIP),
        after: Math.round((parseFloat(cs.marginBottom) || 0) * PX_TO_TWIP),
      },
    };
    // listas de documentos abiertos desde Word (docx-preview usa clases docx-num-ID-NIVEL)
    const num = (el.className || '').match(/docx-num-(\d+)-(\d+)/);
    if (num && !numbering) {
      const before = getComputedStyle(el, '::before').content || '';
      numbering = { reference: before.includes('counter') ? 'numeros' : 'vinetas', level: Math.min(Number(num[2]), 4) };
    }
    if (numbering) opts.numbering = numbering;
    else {
      const indent = (parseFloat(cs.marginLeft) || 0) + (parseFloat(cs.paddingLeft) || 0);
      if (indent > 4) opts.indent = { left: Math.round(indent * PX_TO_TWIP) };
    }
    if (opts.heading) delete opts.spacing;
    return new D.Paragraph(opts);
  };

  /* ---------- tablas ---------- */

  const table = async (tbl) => {
    const rows = [...tbl.rows];
    const tableWidth = tbl.getBoundingClientRect().width || 1;
    const contentTwips = page.contentWidthTwips;
    const outRows = [];
    for (const tr of rows) {
      const cells = [];
      for (const td of tr.cells) {
        const children = await blocks(td, []);
        const cs = getComputedStyle(td);
        const fill = rgbToHex(cs.backgroundColor);
        const width = Math.round((td.getBoundingClientRect().width / tableWidth) * contentTwips);
        cells.push(new D.TableCell({
          children: children.length ? children : [new D.Paragraph({})],
          columnSpan: td.colSpan > 1 ? td.colSpan : undefined,
          rowSpan: td.rowSpan > 1 ? td.rowSpan : undefined,
          shading: fill && fill !== 'FFFFFF' ? { type: D.ShadingType.CLEAR, color: 'auto', fill } : undefined,
          width: { size: width, type: D.WidthType.DXA },
          verticalAlign: cs.verticalAlign === 'middle' ? D.VerticalAlign.CENTER : cs.verticalAlign === 'bottom' ? D.VerticalAlign.BOTTOM : undefined,
        }));
      }
      if (cells.length) outRows.push(new D.TableRow({ children: cells }));
    }
    return outRows.length ? new D.Table({ rows: outRows, width: { size: 100, type: D.WidthType.PERCENTAGE } }) : null;
  };

  /* ---------- recorrido de bloques ---------- */

  const list = async (listEl, out, level) => {
    const reference = listEl.tagName === 'OL' ? 'numeros' : 'vinetas';
    for (const li of listEl.children) {
      if (li.tagName !== 'LI') continue;
      const subs = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
      out.push(await paragraph(li, { reference, level: Math.min(level, 4) }));
      for (const s of subs) await list(s, out, level + 1);
    }
  };

  const blocks = async (container, out) => {
    let pending = [];
    const flush = async () => {
      if (!pending.length) return;
      const nodes = pending;
      pending = [];
      const hasContent = nodes.some((n) => n.textContent.trim() || (n.nodeType === 1 && (n.tagName === 'IMG' || n.querySelector?.('img'))));
      if (!hasContent) return;
      // se leen estilos sobre el contenedor real para no perder herencias
      const runs = [];
      for (const n of nodes) {
        if (n.nodeType === 3) {
          const text = n.textContent.replace(/ /g, ' ');
          if (text.trim() || runs.length) runs.push(new D.TextRun({ text, ...runOptions(container) }));
        } else {
          const fake = { childNodes: [n] };
          await inlineRuns(fake, runs);
        }
      }
      if (runs.length) out.push(new D.Paragraph({ children: runs, alignment: alignmentOf(getComputedStyle(container)) }));
    };

    for (const child of container.childNodes) {
      if (!isBlock(child)) {
        if (child.nodeType === 1 && (child.tagName === 'STYLE' || child.tagName === 'SCRIPT')) continue;
        pending.push(child);
        continue;
      }
      await flush();
      const tag = child.tagName;
      if (getComputedStyle(child).display === 'none') continue;
      if (tag === 'TABLE') {
        const t = await table(child);
        if (t) out.push(t);
      } else if (tag === 'UL' || tag === 'OL') {
        await list(child, out, 0);
      } else if (tag === 'HR') {
        out.push(new D.Paragraph({ children: child.classList.contains('page-break') ? [new D.PageBreak()] : [] }));
      } else if ([...child.childNodes].some(isBlock)) {
        // el editor del navegador a veces anida listas o bloques dentro de un párrafo
        await blocks(child, out);
      } else {
        out.push(await paragraph(child));
      }
    }
    await flush();
    return out;
  };

  const children = await blocks(root, []);
  if (!children.length) children.push(new D.Paragraph({}));

  const listLevels = (format, texts) =>
    texts.map((text, level) => ({
      level,
      format,
      text,
      alignment: D.AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
    }));

  const doc = new D.Document({
    creator: 'Lector Docs · DECOM',
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 } } },
    },
    numbering: {
      config: [
        { reference: 'vinetas', levels: listLevels(D.LevelFormat.BULLET, ['•', '◦', '▪', '•', '◦']) },
        { reference: 'numeros', levels: listLevels(D.LevelFormat.DECIMAL, ['%1.', '%2.', '%3.', '%4.', '%5.']).map((l, i) => (i === 1 ? { ...l, format: D.LevelFormat.LOWER_LETTER, text: '%2)' } : l)) },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: page.widthTwips, height: page.heightTwips },
          margin: { top: page.marginTwips.top, right: page.marginTwips.right, bottom: page.marginTwips.bottom, left: page.marginTwips.left },
        },
      },
      children,
    }],
  });

  const blob = await D.Packer.toBlob(doc);
  return new Uint8Array(await blob.arrayBuffer());
}

/** Tamaño carta con márgenes de 2,54 cm (por defecto para documentos nuevos). */
export const LETTER_PAGE = {
  widthTwips: 12240,
  heightTwips: 15840,
  marginTwips: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
  contentWidthTwips: 9360,
};

/** Lee el tamaño de página que docx-preview dejó en la sección. */
export function pageFromSection(section) {
  if (!section) return LETTER_PAGE;
  const pt = (v) => {
    const n = parseFloat(v);
    if (!n) return null;
    if (String(v).endsWith('pt')) return Math.round(n * 20);
    if (String(v).endsWith('px')) return Math.round(n * PX_TO_TWIP);
    if (String(v).endsWith('cm')) return Math.round(n * 567);
    if (String(v).endsWith('in')) return Math.round(n * 1440);
    return null;
  };
  const s = section.style;
  const widthTwips = pt(s.width) || LETTER_PAGE.widthTwips;
  const heightTwips = pt(s.minHeight) || pt(s.height) || Math.round(widthTwips * 1.294);
  const marginTwips = {
    top: pt(s.paddingTop) ?? 1440,
    right: pt(s.paddingRight) ?? 1440,
    bottom: pt(s.paddingBottom) ?? 1440,
    left: pt(s.paddingLeft) ?? 1440,
  };
  return { widthTwips, heightTwips, marginTwips, contentWidthTwips: widthTwips - marginTwips.left - marginTwips.right };
}
