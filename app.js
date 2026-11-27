/* PPT Cavalieri — genera una presentazione con una slide per relatore partendo da un template .pptx.
   Tutto in locale (JSZip + PptxGenJS inclusi in /lib). */
'use strict';

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const CT_PRES = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CT_POTX = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml';
const EMU_PT = 12700;

const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const log = (m, cls) => { $('log').textContent = m; $('log').style.color = cls === 'err' ? 'var(--err)' : cls === 'ok' ? 'var(--ok)' : 'var(--mut)'; };

/* ------------------------------------------------------------------ stato */
const state = { tplBuf: null, tplName: '', info: null, thumbUrl: null, bgData: null };

/* ------------------------------------------------------------------ helpers XML */
function parseXml(s) {
  const d = new DOMParser().parseFromString(s, 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw new Error('XML non valido nel file PowerPoint');
  return d;
}
const ser = d => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(d.documentElement);
const kids = (el, ns, name) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name && (!ns || n.namespaceURI === ns));
const firstKid = (el, ns, name) => kids(el, ns, name)[0] || null;
function normPath(t) { return t.startsWith('/') ? t.slice(1) : 'ppt/' + t.replace(/^\.\//, ''); }
function relsPathOf(path) { const i = path.lastIndexOf('/'); return path.slice(0, i) + '/_rels/' + path.slice(i + 1) + '.rels'; }

/* ------------------------------------------------------------------ lettura struttura del pptx */
async function readDeck(zip) {
  const presFile = zip.file('ppt/presentation.xml');
  if (!presFile) throw new Error('Non sembra un file PowerPoint valido (manca ppt/presentation.xml)');
  const presDoc = parseXml(await presFile.async('string'));
  const relsDoc = parseXml(await zip.file('ppt/_rels/presentation.xml.rels').async('string'));
  const relMap = {};
  for (const r of Array.from(relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship'))) relMap[r.getAttribute('Id')] = r;
  const lst = presDoc.getElementsByTagNameNS(NS_P, 'sldIdLst')[0];
  const slides = [];
  if (lst) for (const s of kids(lst, NS_P, 'sldId')) {
    const rid = s.getAttributeNS(NS_R, 'id');
    const rel = relMap[rid];
    if (rel) slides.push({ rid, id: +s.getAttribute('id'), path: normPath(rel.getAttribute('Target')) });
  }
  const sz = presDoc.getElementsByTagNameNS(NS_P, 'sldSz')[0];
  const cx = sz ? +sz.getAttribute('cx') : 12192000, cy = sz ? +sz.getAttribute('cy') : 6858000;
  return { presDoc, relsDoc, slides, cx, cy };
}

const PH_TEST = /\{\{\s*(NOME|NAME|RUOLO|ROLE)\s*\}\}/i;
const PH_ALL = /\{\{\s*(NOME|NAME|RUOLO|ROLE)\s*\}\}/gi;

async function inspectTemplate(buf) {
  const zip = await JSZip.loadAsync(buf);
  const deck = await readDeck(zip);
  const info = { slides: deck.slides.length, cx: deck.cx, cy: deck.cy, hasPh: [], hasTitle: [], thumb: null };
  for (const s of deck.slides) {
    const x = await zip.file(s.path).async('string');
    info.hasPh.push(PH_TEST.test(x));
    info.hasTitle.push(/<p:ph[^>]*type="(title|ctrTitle)"/.test(x));
  }
  const th = zip.file('docProps/thumbnail.jpeg') || zip.file('docProps/thumbnail.jpg') || zip.file('docProps/thumbnail.png');
  if (th) info.thumb = await th.async('blob');
  return info;
}

/* ------------------------------------------------------------------ misura testo (per ridurre i nomi lunghi) */
const _cv = document.createElement('canvas').getContext('2d');
function textWidthPt(text, font, sizePt, bold, italic) {
  _cv.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePt}px "${font}", Arial, sans-serif`;
  return Math.max(...String(text).split('\n').map(l => _cv.measureText(l).width));
}
function fitSize(text, font, sizePt, bold, italic, boxWidthPt) {
  if (!boxWidthPt) return sizePt;
  let s = sizePt; const min = sizePt * 0.45;
  while (s > min && textWidthPt(text, font, s, bold, italic) > boxWidthPt) s -= Math.max(1, s * 0.03);
  return Math.round(Math.max(s, min));
}

/* ------------------------------------------------------------------ testo */
function applyCase(t, mode) {
  if (mode === 'upper') return t.toLocaleUpperCase('it-IT');
  if (mode === 'title') return t.toLocaleLowerCase('it-IT').replace(/(^|[\s'’\-])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('it-IT'));
  return t;
}
function splitLines(name, twoLines) {
  if (name.includes('/')) return name.split('/').map(s => s.trim()).filter(Boolean).join('\n'); // a capo manuale con "/"
  const w = name.split(/\s+/);
  return twoLines && w.length > 1 ? w.slice(0, -1).join(' ') + '\n' + w[w.length - 1] : name;
}
function parseNames(txt, caseMode, twoLines) {
  return txt.split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(l => {
    let parts = l.split(/\t|\|/).map(x => x.trim());
    return { name: splitLines(applyCase(parts[0], caseMode), twoLines), role: parts[1] ? applyCase(parts[1], caseMode === 'upper' ? 'upper' : 'none') : '' };
  });
}

function styleRun(doc, r, o) {
  let rPr = firstKid(r, NS_A, 'rPr');
  if (!rPr) { rPr = doc.createElementNS(NS_A, 'a:rPr'); r.insertBefore(rPr, r.firstChild); }
  rPr.setAttribute('lang', 'it-IT');
  rPr.setAttribute('sz', String(Math.round(o.size * 100)));
  rPr.setAttribute('b', o.bold ? '1' : '0');
  rPr.setAttribute('i', o.italic ? '1' : '0');
  for (const n of ['solidFill', 'gradFill', 'noFill', 'latin', 'ea', 'cs']) kids(rPr, NS_A, n).forEach(k => rPr.removeChild(k));
  const fill = doc.createElementNS(NS_A, 'a:solidFill');
  const clr = doc.createElementNS(NS_A, 'a:srgbClr'); clr.setAttribute('val', o.color.replace('#', '').toUpperCase());
  fill.appendChild(clr);
  const ln = firstKid(rPr, NS_A, 'ln');
  rPr.insertBefore(fill, ln ? ln.nextSibling : rPr.firstChild);
  const tail = ['sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].map(n => firstKid(rPr, NS_A, n)).find(Boolean) || null;
  for (const n of ['latin', 'ea', 'cs']) {
    const f = doc.createElementNS(NS_A, 'a:' + n); f.setAttribute('typeface', o.font);
    rPr.insertBefore(f, tail);
  }
}

function setParagraphText(doc, p, text, styleOpts) {
  const runs = kids(p, NS_A, 'r');
  let first = runs[0];
  if (!first) {
    first = doc.createElementNS(NS_A, 'a:r');
    first.appendChild(doc.createElementNS(NS_A, 'a:t'));
    const end = firstKid(p, NS_A, 'endParaRPr');
    p.insertBefore(first, end);
  }
  runs.slice(1).forEach(r => p.removeChild(r));
  kids(p, NS_A, 'fld').forEach(r => p.removeChild(r));
  kids(p, NS_A, 'br').forEach(r => p.removeChild(r));
  let t = firstKid(first, NS_A, 't');
  if (!t) { t = doc.createElementNS(NS_A, 'a:t'); first.appendChild(t); }
  const lines = String(text).split('\n');
  t.textContent = lines[0];
  if (styleOpts) styleRun(doc, first, styleOpts);
  let prev = first;
  for (let k = 1; k < lines.length; k++) {
    const br = doc.createElementNS(NS_A, 'a:br');
    const rp = firstKid(first, NS_A, 'rPr'); if (rp) br.appendChild(rp.cloneNode(true));
    p.insertBefore(br, prev.nextSibling);
    const r2 = first.cloneNode(true); firstKid(r2, NS_A, 't').textContent = lines[k];
    p.insertBefore(r2, br.nextSibling);
    prev = r2;
  }
}

function ancestorSp(el) { while (el && !(el.localName === 'sp' && el.namespaceURI === NS_P)) el = el.parentNode; return el; }
function shapeWidthPt(sp) {
  if (!sp) return null;
  const x = sp.getElementsByTagNameNS(NS_A, 'xfrm')[0];
  const ext = x && firstKid(x, NS_A, 'ext');
  const pad = 2 * 7.2; // inset di default
  return ext ? ext.getAttribute('cx') / EMU_PT - pad : null;
}

function maxId(doc) {
  let m = 1;
  for (const e of Array.from(doc.getElementsByTagNameNS(NS_P, 'cNvPr'))) m = Math.max(m, +e.getAttribute('id') || 0);
  return m;
}

function textboxXml(id, label, o, xEmu, yEmu, wEmu, hEmu, paras) {
  const pXml = paras.map(p => {
    const pr = `<a:rPr lang="it-IT" sz="${Math.round(p.size * 100)}" b="${p.bold ? 1 : 0}" i="${p.italic ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${p.color.replace('#', '').toUpperCase()}"/></a:solidFill><a:latin typeface="${esc(o.font)}"/><a:ea typeface="${esc(o.font)}"/><a:cs typeface="${esc(o.font)}"/></a:rPr>`;
    const runs = String(p.text).split('\n').map(l => `<a:r>${pr}<a:t>${esc(l)}</a:t></a:r>`).join(`<a:br>${pr}</a:br>`);
    return `<a:p><a:pPr algn="${o.align}"/>${runs}</a:p>`;
  }).join('');
  return `<p:sp xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:nvSpPr><p:cNvPr id="${id}" name="${esc(label)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(xEmu)}" y="${Math.round(yEmu)}"/><a:ext cx="${Math.round(wEmu)}" cy="${Math.round(hEmu)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="${o.anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${pXml}</p:txBody></p:sp>`;
}

/* ------------------------------------------------------------------ banner */
const BN_PRESETS = {
  full: o => ({ x: 0, y: 100 - o.h, w: 100 }),
  inset: o => ({ x: 5, y: 100 - o.h - 6, w: 90 }),
  top: o => ({ x: 0, y: 0, w: 100 }),
  middle: o => ({ x: 0, y: (100 - o.h) / 2, w: 100 })
};
function bannerRect(o) {
  const w = Math.min(Math.max(o.bnW, 5), 100), h = Math.min(Math.max(o.bnH, 3), 100);
  return { x: Math.min(Math.max(o.bnX, 0), 100 - w), y: Math.min(Math.max(o.bnY, 0), 100 - h), w, h };
}
function rectXml(id, label, xE, yE, wE, hE, color, alpha) {
  const a = alpha < 100 ? `<a:alpha val="${Math.round(alpha * 1000)}"/>` : '';
  return `<p:sp xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:nvSpPr><p:cNvPr id="${id}" name="${esc(label)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(xE)}" y="${Math.round(yE)}"/><a:ext cx="${Math.round(wE)}" cy="${Math.round(hE)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}">${a}</a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
}
function addBanner(doc, o, W, H) {
  const spTree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0];
  const r = bannerRect(o);
  const x = W * r.x / 100, y = H * r.y / 100, w = W * r.w / 100, h = H * r.h / 100;
  let id = maxId(doc) + 1;
  spTree.appendChild(doc.importNode(parseXml(rectXml(id++, 'Banner', x, y, w, h, o.bnColor, o.bnOpacity)).documentElement, true));
  if (o.bnLine) {
    const lh = Math.max(H * 0.008, 25400);
    const ly = o.bnLinePos === 'below' ? y + h : y - lh;
    spTree.appendChild(doc.importNode(parseXml(rectXml(id++, 'Banner linea', x, ly, w, lh, o.bnLineColor, 100)).documentElement, true));
  }
}

/* ------------------------------------------------------------------ transizioni */
function buildTransition(doc, o) {
  const tr = doc.createElementNS(NS_P, 'p:transition');
  tr.setAttribute('spd', o.trSpd);
  tr.setAttribute('advClick', '1');
  if (o.advSec > 0) tr.setAttribute('advTm', String(Math.round(o.advSec * 1000)));
  const mk = (n, attrs) => { const e = doc.createElementNS(NS_P, 'p:' + n); for (const k in (attrs || {})) e.setAttribute(k, attrs[k]); tr.appendChild(e); };
  const d = o.trDir;
  const hv = (d === 'l' || d === 'r') ? 'horz' : 'vert';
  switch (o.trType) {
    case 'fade': mk('fade'); break;
    case 'dissolve': mk('dissolve'); break;
    case 'push': mk('push', { dir: d }); break;
    case 'wipe': mk('wipe', { dir: d }); break;
    case 'cover': mk('cover', { dir: d }); break;
    case 'pull': mk('pull', { dir: d }); break;
    case 'split': mk('split', { orient: hv, dir: 'out' }); break;
    case 'zoom': mk('zoom', { dir: 'in' }); break;
    case 'blinds': mk('blinds', { dir: hv }); break;
    case 'checker': mk('checker', { dir: hv }); break;
    case 'circle': mk('circle'); break;
    case 'diamond': mk('diamond'); break;
    case 'plus': mk('plus'); break;
    case 'wedge': mk('wedge'); break;
    case 'wheel': mk('wheel', { spokes: '4' }); break;
    case 'randomBar': mk('randomBar', { dir: hv }); break;
    case 'comb': mk('comb', { dir: hv }); break;
    case 'newsflash': mk('newsflash'); break;
  }
  return tr;
}
function applyTransition(doc, o) {
  if (o.trType === 'keep' && !(o.advSec > 0)) return;
  const root = doc.documentElement;
  // elimina transizioni esistenti (anche dentro mc:AlternateContent)
  for (const t of Array.from(root.getElementsByTagNameNS(NS_P, 'transition'))) {
    if (!root.contains(t)) continue; // già rimossa insieme al suo mc:AlternateContent
    let n = t;
    while (n.parentNode !== root) n = n.parentNode;
    root.removeChild(n);
  }
  if (o.trType === 'none' && !(o.advSec > 0)) return;
  const tr = o.trType === 'none' || o.trType === 'keep'
    ? (() => { const e = doc.createElementNS(NS_P, 'p:transition'); e.setAttribute('advClick', '1'); e.setAttribute('advTm', String(Math.round(o.advSec * 1000))); return e; })()
    : buildTransition(doc, o);
  const before = ['timing', 'extLst'].map(n => firstKid(root, NS_P, n)).find(Boolean) || null;
  root.insertBefore(tr, before);
}

/* ------------------------------------------------------------------ trasformazione di una slide */
function transformSlide(xml, person, o, slideW, slideH) {
  const doc = parseXml(xml);
  const stats = { ph: 0, title: 0, box: 0 };
  const nameOpts = { font: o.font, size: o.size, bold: o.bold, italic: o.italic, color: o.color };
  const roleOpts = { font: o.font, size: Math.round(o.size * o.roleRatio / 100), bold: false, italic: o.italic, color: o.roleColor };

  const nameSize = (box) => o.fit ? fitSize(person.name, o.font, o.size, o.bold, o.italic, box) : o.size;
  const roleSize = (box) => o.fit && person.role ? fitSize(person.role, o.font, roleOpts.size, false, o.italic, box) : roleOpts.size;

  if (o.stripText) {
    const keepTitle = !o.forceNew;
    for (const sp of Array.from(doc.getElementsByTagNameNS(NS_P, 'sp'))) {
      const txt = Array.from(sp.getElementsByTagNameNS(NS_A, 't')).map(t => t.textContent).join('').trim();
      if (!txt || PH_TEST.test(txt)) continue;
      const ph = sp.getElementsByTagNameNS(NS_P, 'ph')[0];
      if (keepTitle && ph && /^(title|ctrTitle)$/.test(ph.getAttribute('type') || '')) continue;
      sp.parentNode.removeChild(sp);
    }
  }

  if (o.bnOn) addBanner(doc, o, slideW, slideH);

  if (!o.forceNew) {
    // 1) segnaposto {{NOME}} / {{RUOLO}}
    for (const p of Array.from(doc.getElementsByTagNameNS(NS_A, 'p'))) {
      const full = kids(p, NS_A, 'r').concat(kids(p, NS_A, 'fld')).map(r => (firstKid(r, NS_A, 't') || {}).textContent || '').join('');
      if (!PH_TEST.test(full)) continue;
      const isRoleOnly = /^\s*\{\{\s*(RUOLO|ROLE)\s*\}\}\s*$/i.test(full);
      const out = full.replace(PH_ALL, (m, k) => /^(NOME|NAME)$/i.test(k) ? person.name : person.role);
      const w = shapeWidthPt(ancestorSp(p));
      let st = null;
      if (!o.keepStyle) {
        st = Object.assign({}, isRoleOnly ? roleOpts : nameOpts);
        st.size = isRoleOnly ? roleSize(w) : nameSize(w);
      }
      setParagraphText(doc, p, out, st);
      stats.ph++;
    }
  }

  if (!stats.ph && !o.forceNew) {
    // 2) segnaposto titolo (+ sottotitolo per il ruolo)
    const sps = Array.from(doc.getElementsByTagNameNS(NS_P, 'sp'));
    const phType = sp => { const ph = sp.getElementsByTagNameNS(NS_P, 'ph')[0]; return ph ? (ph.getAttribute('type') || 'body') + ':' + (ph.getAttribute('idx') || '0') : null; };
    const title = sps.find(sp => /^(title|ctrTitle):/.test(phType(sp) || ''));
    const sub = sps.find(sp => /^(subTitle|body):/.test(phType(sp) || ''));
    if (title) {
      const tb = title.getElementsByTagNameNS(NS_P, 'txBody')[0];
      const paras = kids(tb, NS_A, 'p');
      const w = shapeWidthPt(title);
      paras.slice(1).forEach(p => tb.removeChild(p));
      setParagraphText(doc, paras[0], person.name, o.keepStyle ? null : Object.assign({}, nameOpts, { size: nameSize(w) }));
      stats.title++;
      if (person.role) {
        if (sub) {
          const sb = sub.getElementsByTagNameNS(NS_P, 'txBody')[0];
          const sp2 = kids(sb, NS_A, 'p');
          sp2.slice(1).forEach(p => sb.removeChild(p));
          setParagraphText(doc, sp2[0], person.role, o.keepStyle ? null : Object.assign({}, roleOpts, { size: roleSize(shapeWidthPt(sub)) }));
        } else {
          const p2 = paras[0].cloneNode(true);
          tb.insertBefore(p2, null);
          setParagraphText(doc, p2, person.role, o.keepStyle ? null : Object.assign({}, roleOpts, { size: roleSize(w) }));
        }
      } else if (sub) {
        const sb = sub.getElementsByTagNameNS(NS_P, 'txBody')[0];
        const sp2 = kids(sb, NS_A, 'p');
        sp2.slice(1).forEach(p => sb.removeChild(p));
        setParagraphText(doc, sp2[0], '', null);
      }
    }
  }

  if (!stats.ph && !stats.title) {
    // 3) nuova casella di testo
    const spTree = doc.getElementsByTagNameNS(NS_P, 'spTree')[0];
    const x = slideW * o.px / 100, y = slideH * o.py / 100, w = slideW * o.pw / 100, h = slideH * o.ph / 100;
    const wPt = w / EMU_PT;
    const paras = [Object.assign({ text: person.name }, nameOpts, { size: nameSize(wPt) })];
    if (person.role) paras.push(Object.assign({ text: person.role }, roleOpts, { size: roleSize(wPt) }));
    const id = maxId(doc) + 1;
    const frag = parseXml(textboxXml(id, 'Relatore', o, x, y, w, h, paras)).documentElement;
    spTree.appendChild(doc.importNode(frag, true));
    stats.box++;
  }

  applyTransition(doc, o);
  return { xml: ser(doc), stats };
}

/* ------------------------------------------------------------------ rimozione di una slide dal pacchetto */
async function dropSlide(zip, ctDoc, relsDoc, slide) {
  const relsPath = relsPathOf(slide.path);
  const rf = zip.file(relsPath);
  if (rf) {
    const rd = parseXml(await rf.async('string'));
    for (const r of Array.from(rd.getElementsByTagNameNS(NS_REL, 'Relationship'))) {
      if (/\/notesSlide$/.test(r.getAttribute('Type'))) {
        const real = r.getAttribute('Target').startsWith('/') ? r.getAttribute('Target').slice(1) : 'ppt/' + r.getAttribute('Target').replace(/^\.\.\//, '');
        zip.remove(real); zip.remove(relsPathOf(real));
        for (const o of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override'))) if (o.getAttribute('PartName') === '/' + real) o.parentNode.removeChild(o);
      }
    }
  }
  zip.remove(slide.path); zip.remove(relsPath);
  for (const o of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override'))) if (o.getAttribute('PartName') === '/' + slide.path) o.parentNode.removeChild(o);
  const rel = Array.from(relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship')).find(r => r.getAttribute('Id') === slide.rid);
  if (rel) rel.parentNode.removeChild(rel);
}

/* ------------------------------------------------------------------ generazione del deck */
async function buildDeck(buf, people, o) {
  const zip = await JSZip.loadAsync(buf);
  const deck = await readDeck(zip);
  if (!deck.slides.length) throw new Error('Il template non contiene slide. Aggiungi almeno una slide modello e riprova.');
  const idx = Math.min(Math.max(o.slideIdx, 1), deck.slides.length) - 1;
  const tpl = deck.slides[idx];
  const tplXml = await zip.file(tpl.path).async('string');
  const tplRelsFile = zip.file(relsPathOf(tpl.path));
  const tplRels = tplRelsFile ? await tplRelsFile.async('string') : null;

  const ctDoc = parseXml(await zip.file('[Content_Types].xml').async('string'));
  const { presDoc, relsDoc } = deck;

  // numerazione nuova
  let maxN = 0;
  zip.forEach(p => { const m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/); if (m) maxN = Math.max(maxN, +m[1]); });
  let maxId = 255;
  deck.slides.forEach(s => { maxId = Math.max(maxId, s.id); });
  let ridN = 1; const usedRid = new Set(Array.from(relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship')).map(r => r.getAttribute('Id')));
  const newRid = () => { let r; do { r = 'rIdCav' + (ridN++); } while (usedRid.has(r)); usedRid.add(r); return r; };

  // pulisci le relazioni della slide modello (note e commenti non vanno duplicati)
  let cleanRels = tplRels;
  if (tplRels) {
    const rd = parseXml(tplRels);
    for (const r of Array.from(rd.getElementsByTagNameNS(NS_REL, 'Relationship'))) {
      if (/notesSlide$|comments/i.test(r.getAttribute('Type'))) r.parentNode.removeChild(r);
    }
    cleanRels = ser(rd);
  }

  const created = [];
  const totals = { ph: 0, title: 0, box: 0 };
  for (let i = 0; i < people.length; i++) {
    const n = maxN + 1 + i;
    const path = `ppt/slides/slide${n}.xml`;
    const { xml, stats } = transformSlide(tplXml, people[i], o, deck.cx, deck.cy);
    totals.ph += stats.ph; totals.title += stats.title; totals.box += stats.box;
    zip.file(path, xml);
    if (cleanRels) zip.file(relsPathOf(path), cleanRels);
    const ov = ctDoc.createElementNS(NS_CT, 'Override');
    ov.setAttribute('PartName', '/' + path); ov.setAttribute('ContentType', CT_SLIDE);
    ctDoc.documentElement.appendChild(ov);
    const rid = newRid();
    const rel = relsDoc.createElementNS(NS_REL, 'Relationship');
    rel.setAttribute('Id', rid); rel.setAttribute('Type', REL_SLIDE); rel.setAttribute('Target', `slides/slide${n}.xml`);
    relsDoc.documentElement.appendChild(rel);
    created.push({ rid, id: ++maxId });
  }

  // nuovo elenco slide: le nuove prendono il posto del modello
  const lst = presDoc.getElementsByTagNameNS(NS_P, 'sldIdLst')[0];
  const oldEntries = kids(lst, NS_P, 'sldId');
  const mkEntry = c => { const e = presDoc.createElementNS(NS_P, 'p:sldId'); e.setAttribute('id', String(c.id)); e.setAttributeNS(NS_R, 'r:id', c.rid); return e; };
  const anchor = oldEntries[idx];
  for (const c of created) lst.insertBefore(mkEntry(c), anchor);
  const toDrop = o.keepOthers ? [tpl] : deck.slides;
  for (const s of toDrop) {
    const entry = oldEntries.find(e => e.getAttributeNS(NS_R, 'id') === s.rid);
    if (entry) lst.removeChild(entry);
    await dropSlide(zip, ctDoc, relsDoc, s);
  }

  // sezioni e custom show puntano a id di slide ormai spariti: li tolgo
  for (const sec of Array.from(presDoc.getElementsByTagName('*')).filter(e => e.localName === 'sectionLst')) {
    let n = sec; while (n.parentNode && n.localName !== 'ext') n = n.parentNode;
    const target = n.localName === 'ext' ? n : sec; target.parentNode.removeChild(target);
  }
  for (const c of Array.from(presDoc.getElementsByTagNameNS(NS_P, 'custShowLst'))) c.parentNode.removeChild(c);

  // .potx -> .pptx
  for (const ov of Array.from(ctDoc.getElementsByTagNameNS(NS_CT, 'Override'))) {
    if (ov.getAttribute('ContentType') === CT_POTX) ov.setAttribute('ContentType', CT_PRES);
  }

  zip.file('ppt/presentation.xml', ser(presDoc));
  zip.file('ppt/_rels/presentation.xml.rels', ser(relsDoc));
  zip.file('[Content_Types].xml', ser(ctDoc));
  const blob = await zip.generateAsync({
    type: 'blob', compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  });
  return { blob, totals };
}

/* ------------------------------------------------------------------ template base / di esempio (PptxGenJS) */
async function makeBaseTemplate(withPlaceholders) {
  const pptx = new PptxGenJS();
  const r = $('ratio').value;
  const W = r === '43' ? 10 : r === '916' ? 5.625 : 13.333, H = r === '916' ? 10 : 7.5;
  pptx.defineLayout({ name: 'CAV', width: W, height: H });
  pptx.layout = 'CAV';
  const s = pptx.addSlide();
  s.background = { color: $('bgColor').value.replace('#', '') };
  if (state.bgData) s.addImage({ data: state.bgData, x: 0, y: 0, w: W, h: H });
  if (withPlaceholders) {
    s.addText('{{NOME}}', { x: W * .06, y: H * .68, w: W * .88, h: H * .14, fontFace: 'Arial', fontSize: 54, bold: true, color: 'FFFFFF', margin: 0 });
    s.addText('{{RUOLO}}', { x: W * .06, y: H * .82, w: W * .88, h: H * .09, fontFace: 'Arial', fontSize: 28, color: '9EC8FF', margin: 0 });
  }
  return await pptx.write({ outputType: 'arraybuffer' });
}

/* ------------------------------------------------------------------ UI */
function readOpts() {
  const n = (id, d) => { const v = parseFloat($(id).value); return isNaN(v) ? d : v; };
  const o = {
    slideIdx: Math.round(n('slideIdx', 1)),
    font: $('font').value.trim() || 'Arial',
    size: n('size', 54), roleRatio: n('roleRatio', 55),
    color: $('color').value, roleColor: $('roleColor').value,
    align: $('align').value, anchor: $('anchor').value,
    bold: $('bold').checked, italic: $('italic').checked, fit: $('fit').checked,
    keepStyle: $('keepStyle').checked, forceNew: $('forceNew').checked,
    px: n('px', 6), py: n('py', 68), pw: n('pw', 88), ph: n('ph', 24),
    trType: $('trType').value, trDir: $('trDir').value, trSpd: $('trSpd').value, advSec: n('advSec', 0),
    twoLines: $('lines').value === '2', keepOthers: $('keepOthers').checked, stripText: $('stripText').checked,
    bnOn: $('bnOn').checked, bnColor: $('bnColor').value, bnOpacity: n('bnOpacity', 90), bnX: n('bnX', 0), bnY: n('bnY', 78), bnW: n('bnW', 100), bnH: n('bnH', 20), bnLinePos: $('bnLinePos').value, bnLine: $('bnLine').checked, bnLineColor: $('bnLineColor').value, caseMode: $('caseMode').value
  };
  if (o.bnOn) { // il testo va dentro il banner
    const r = bannerRect(o);
    Object.assign(o, { forceNew: true, px: r.x + 3, pw: r.w - 6, py: r.y, ph: r.h, anchor: 'ctr' });
  }
  return o;
}

const DIR_TYPES = ['push', 'wipe', 'cover', 'pull', 'split', 'blinds', 'checker', 'randomBar', 'comb'];
function refreshUi() {
  const o = readOpts();
  const people = parseNames($('names').value, o.caseMode, o.twoLines);
  $('cnt').textContent = people.length + (people.length === 1 ? ' relatore' : ' relatori');
  $('trDir').disabled = !DIR_TYPES.includes(o.trType);
  $('baseBox').style.display = state.tplBuf ? 'none' : '';
  $('keepOthers').disabled = !state.tplBuf;

  // info template
  const box = $('tplInfo');
  box.innerHTML = '';
  const badge = (t, c) => { const s = document.createElement('span'); s.className = 'badge ' + (c || ''); s.textContent = t; box.appendChild(s); };
  let willUsePh = false, willUseTitle = false;
  if (!state.tplBuf) badge('Nessun template: userò uno sfondo semplice');
  else if (state.info) {
    const i = Math.min(Math.max(o.slideIdx, 1), state.info.slides) - 1;
    badge(`${state.tplName} · ${state.info.slides} slide`, 'ok');
    willUsePh = state.info.hasPh[i] && !o.forceNew;
    willUseTitle = !willUsePh && state.info.hasTitle[i] && !o.forceNew;
    if (willUsePh) badge('Trovato {{NOME}} nella slide ' + (i + 1), 'ok');
    else if (willUseTitle) badge('Nessun {{NOME}}: scrivo nel titolo della slide ' + (i + 1), 'warn');
    else badge('Nessun segnaposto né titolo: aggiungo una casella di testo', 'warn');
  }

  // anteprima
  const stage = $('stage'), ovl = $('ovl'), note = $('stageNote');
  const rr = $('ratio').value;
  const asp = state.info ? state.info.cx / state.info.cy : (rr === '43' ? 4 / 3 : rr === '916' ? 9 / 16 : 16 / 9);
  stage.style.aspectRatio = String(asp);
  stage.style.maxWidth = asp < 1 ? '300px' : 'none'; stage.style.marginLeft = stage.style.marginRight = 'auto';
  if (state.thumbUrl && state.tplBuf) { stage.style.backgroundImage = `url(${state.thumbUrl})`; stage.style.backgroundColor = '#2a2f39'; }
  else stage.style.backgroundImage = state.bgData && !state.tplBuf ? `url(data:${state.bgData})` : 'none';
  stage.style.backgroundColor = state.tplBuf ? '#2a2f39' : $('bgColor').value;
  const bn = $('bnr'), bl = $('bnrLine');
  if (o.bnOn) {
    const r = bannerRect(o);
    Object.assign(bn.style, { display: 'block', left: r.x + '%', top: r.y + '%', width: r.w + '%', height: r.h + '%', background: o.bnColor, opacity: o.bnOpacity / 100 });
    Object.assign(bl.style, { display: o.bnLine ? 'block' : 'none', left: r.x + '%', width: r.w + '%', height: '0.8%', background: o.bnLineColor, top: (o.bnLinePos === 'below' ? r.y + r.h : r.y - 0.8) + '%' });
  } else { bn.style.display = 'none'; bl.style.display = 'none'; }
  const sample = people[0] || { name: 'Mario Rossi', role: 'Amministratore Delegato' };
  const slideWpt = state.info ? state.info.cx / EMU_PT : (rr === '43' ? 720 : rr === '916' ? 405 : 960);
  const k = stage.clientWidth / slideWpt;
  $('pvName').style.whiteSpace = 'pre';
  if (willUsePh || willUseTitle) {
    ovl.classList.add('off');
    note.style.display = 'block';
    note.textContent = willUsePh ? 'Il testo andrà al posto di {{NOME}} / {{RUOLO}} nel template.' : 'Il testo andrà nel titolo della slide modello.';
  } else {
    note.style.display = 'none'; ovl.classList.remove('off');
    Object.assign(ovl.style, {
      left: o.px + '%', top: o.py + '%', width: o.pw + '%', height: o.ph + '%',
      justifyContent: { t: 'flex-start', ctr: 'center', b: 'flex-end' }[o.anchor],
      textAlign: { l: 'left', ctr: 'center', r: 'right' }[o.align]
    });
    const n = $('pvName'), r = $('pvRole');
    n.textContent = sample.name; r.textContent = sample.role; r.style.display = sample.role ? '' : 'none';
    const nsz = o.fit ? fitSize(sample.name, o.font, o.size, o.bold, o.italic, stage.clientWidth / k * o.pw / 100) : o.size;
    const rsz0 = o.size * o.roleRatio / 100;
    const rsz = o.fit && sample.role ? fitSize(sample.role, o.font, rsz0, false, o.italic, stage.clientWidth / k * o.pw / 100) : rsz0;
    Object.assign(n.style, { fontFamily: `"${o.font}", Arial, sans-serif`, fontSize: nsz * k + 'px', fontWeight: o.bold ? 700 : 400, fontStyle: o.italic ? 'italic' : 'normal', color: o.color, lineHeight: 1.2 });
    Object.assign(r.style, { fontFamily: `"${o.font}", Arial, sans-serif`, fontSize: rsz * k + 'px', fontWeight: 400, fontStyle: o.italic ? 'italic' : 'normal', color: o.roleColor, lineHeight: 1.2 });
  }
  save();
}

const SAVE_IDS = ['lines', 'slideIdx', 'bgColor', 'ratio', 'names', 'caseMode', 'font', 'size', 'roleRatio', 'color', 'roleColor', 'align', 'anchor', 'px', 'py', 'pw', 'ph', 'bnColor', 'bnOpacity', 'bnX', 'bnY', 'bnW', 'bnLinePos', 'bnH', 'bnLineColor', 'trType', 'trDir', 'trSpd', 'advSec', 'outName'];
const SAVE_CHK = ['bold', 'italic', 'fit', 'keepStyle', 'forceNew', 'keepOthers', 'stripText', 'bnOn', 'bnLine'];
function save() {
  try {
    const d = {}; SAVE_IDS.forEach(i => d[i] = $(i).value); SAVE_CHK.forEach(i => d[i] = $(i).checked);
    localStorage.setItem('pptcavalieri', JSON.stringify(d));
  } catch (e) { /* storage non disponibile: pazienza */ }
}
function load() {
  try {
    const d = JSON.parse(localStorage.getItem('pptcavalieri') || 'null'); if (!d) return;
    SAVE_IDS.forEach(i => { if (d[i] !== undefined) $(i).value = d[i]; });
    SAVE_CHK.forEach(i => { if (d[i] !== undefined) $(i).checked = d[i]; });
  } catch (e) { /* ignora */ }
}

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

async function onTemplate(file) {
  if (state.thumbUrl) { URL.revokeObjectURL(state.thumbUrl); state.thumbUrl = null; }
  if (!file) { state.tplBuf = null; state.info = null; refreshUi(); return; }
  try {
    const buf = await file.arrayBuffer();
    const info = await inspectTemplate(buf);
    state.tplBuf = buf; state.tplName = file.name; state.info = info;
    if (info.thumb) state.thumbUrl = URL.createObjectURL(info.thumb);
    $('slideIdx').max = info.slides;
    log('Template caricato.', 'ok');
  } catch (e) {
    state.tplBuf = null; state.info = null; $('tplFile').value = '';
    log('Impossibile leggere il template: ' + e.message, 'err');
  }
  refreshUi();
}

async function onGo() {
  const btn = $('go'); const o = readOpts();
  const people = parseNames($('names').value, o.caseMode, o.twoLines);
  if (!people.length) { log('Inserisci almeno un nome.', 'err'); return; }
  btn.disabled = true; log('Creo la presentazione…');
  try {
    const buf = state.tplBuf || await makeBaseTemplate(false);
    const { blob, totals } = await buildDeck(buf, people, Object.assign({}, o, { slideIdx: state.tplBuf ? o.slideIdx : 1 }));
    const fname = ($('outName').value.trim().replace(/[\\/:*?"<>|]+/g, '_') || 'cavalieri') + '.pptx';
    download(blob, fname);
    const how = totals.ph ? 'segnaposto {{NOME}}' : totals.title ? 'titolo della slide' : 'nuova casella di testo';
    log(`Fatto: ${people.length} slide in ${fname} (testo inserito tramite ${how}).`, 'ok');
  } catch (e) {
    console.error(e); log('Errore: ' + e.message, 'err');
  }
  btn.disabled = false;
}

document.addEventListener('DOMContentLoaded', () => {
  load();
  $('tplFile').addEventListener('change', e => onTemplate(e.target.files[0]));
  $('bgImg').addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) { state.bgData = null; refreshUi(); return; }
    const r = new FileReader(); r.onload = () => { state.bgData = String(r.result).replace(/^data:/, ''); refreshUi(); }; r.readAsDataURL(f);
  });
  $('btnSample').addEventListener('click', async () => {
    try { download(new Blob([await makeBaseTemplate(true)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), 'template_esempio.pptx'); }
    catch (e) { log('Errore: ' + e.message, 'err'); }
  });
  $('go').addEventListener('click', onGo);
  document.querySelectorAll('input,select,textarea').forEach(el => { if (el.type !== 'file') { el.addEventListener('input', refreshUi); el.addEventListener('change', refreshUi); } });
  window.addEventListener('resize', refreshUi);
  refreshUi();
});

/* ------------------------------------------------------------------ programma -> nomi */
const prog = { entries: [] };
function renderProg() {
  const keepTitle = $('progTitles').checked, withRole = $('progRoles').checked;
  const list = $('progList'); list.innerHTML = '';
  prog.entries.forEach((e, i) => {
    const row = document.createElement('label'); row.className = 'chk'; row.style.margin = '3px 0';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = e.on !== false; cb.addEventListener('change', () => { e.on = cb.checked; });
    const tx = document.createElement('input'); tx.type = 'text'; tx.value = e.edited !== undefined ? e.edited : window.__prog.entryToLine(e, keepTitle, withRole);
    tx.addEventListener('input', () => { e.edited = tx.value; });
    row.appendChild(cb); row.appendChild(tx); list.appendChild(row);
  });
  $('progTitle').textContent = prog.entries.length + (prog.entries.length === 1 ? ' nome trovato' : ' nomi trovati');
}
function progLines() {
  const keepTitle = $('progTitles').checked, withRole = $('progRoles').checked;
  return prog.entries.filter(e => e.on !== false).map(e => (e.edited !== undefined ? e.edited : window.__prog.entryToLine(e, keepTitle, withRole)).trim()).filter(Boolean);
}
async function onProgram(file) {
  if (!file) return;
  try {
    log('Leggo il programma…');
    const text = await window.__prog.textFromFile(file);
    prog.entries = window.__prog.parseProgram(text);
    if (!prog.entries.length) { $('progBox').style.display = 'none'; log('Non ho trovato nomi in questo file. Prova con un altro formato o incolla i nomi a mano.', 'err'); return; }
    $('progBox').style.display = 'block'; renderProg(); log(`Programma letto: ${prog.entries.length} nomi trovati, controllali e aggiungili alla lista.`, 'ok');
  } catch (e) { console.error(e); log('Impossibile leggere il programma: ' + e.message, 'err'); }
}
function applyProg(replace) {
  const lines = progLines(); if (!lines.length) return;
  const cur = $('names').value.replace(/\s+$/, '');
  $('names').value = replace || !cur ? lines.join('\n') : cur + '\n' + lines.join('\n');
  $('progBox').style.display = 'none'; $('progFile').value = '';
  refreshUi(); log(`${lines.length} nomi inseriti nella lista.`, 'ok');
}
/* ------------------------------------------------------------------ banner: preset e trascinamento */
document.addEventListener('DOMContentLoaded', () => {
  $('bnPreset').addEventListener('change', () => {
    const f = BN_PRESETS[$('bnPreset').value]; if (!f) return;
    const h = parseFloat($('bnH').value) || 20, r = f({ h });
    $('bnX').value = r.x; $('bnY').value = Math.round(r.y * 10) / 10; $('bnW').value = r.w; refreshUi();
  });
  const bn = $('bnr'), stage = $('stage');
  let drag = null;
  bn.addEventListener('pointerdown', e => {
    const o = readOpts(); const r = bannerRect(o);
    drag = { x0: e.clientX, y0: e.clientY, bx: r.x, by: r.y, w: r.w, h: r.h };
    bn.setPointerCapture(e.pointerId); e.preventDefault();
  });
  bn.addEventListener('pointermove', e => {
    if (!drag) return;
    const sr = stage.getBoundingClientRect();
    const nx = drag.bx + (e.clientX - drag.x0) / sr.width * 100, ny = drag.by + (e.clientY - drag.y0) / sr.height * 100;
    $('bnX').value = Math.round(Math.min(Math.max(nx, 0), 100 - drag.w) * 10) / 10;
    $('bnY').value = Math.round(Math.min(Math.max(ny, 0), 100 - drag.h) * 10) / 10;
    $('bnPreset').value = 'custom'; refreshUi();
  });
  const end = () => { drag = null; };
  bn.addEventListener('pointerup', end); bn.addEventListener('pointercancel', end);
  ['bnX', 'bnY', 'bnW', 'bnH'].forEach(id => $(id).addEventListener('input', () => { $('bnPreset').value = 'custom'; }));
});

document.addEventListener('DOMContentLoaded', () => {
  $('progFile').addEventListener('change', e => onProgram(e.target.files[0]));
  $('progClose').addEventListener('click', () => { $('progBox').style.display = 'none'; $('progFile').value = ''; });
  $('progAdd').addEventListener('click', () => applyProg(false));
  $('progReplace').addEventListener('click', () => applyProg(true));
  ['progTitles', 'progRoles'].forEach(id => $(id).addEventListener('change', () => { prog.entries.forEach(e => delete e.edited); renderProg(); }));
});

// per i test automatici
window.__cav = { buildDeck, parseNames, makeBaseTemplate, inspectTemplate };
