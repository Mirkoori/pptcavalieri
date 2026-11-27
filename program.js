/* Estrazione dei nomi dei relatori da un programma (PDF, Word .docx, testo). Tutto in locale. */
'use strict';

/* ------------------------------------------------------------ lettura file -> testo */
async function textFromPdf(buf) {
  if (typeof pdfjsLib === 'undefined') throw new Error('Libreria PDF non caricata');
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false }).promise;
  const out = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const tc = await page.getTextContent();
    const rows = [];
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const y = it.transform[5], x = it.transform[4];
      let row = rows.find(r => Math.abs(r.y - y) < 3);
      if (!row) { row = { y, items: [] }; rows.push(row); }
      row.items.push({ x, w: it.width || 0, s: it.str });
    }
    rows.sort((a, b) => b.y - a.y);
    for (const r of rows) {
      r.items.sort((a, b) => a.x - b.x);
      let line = '', prevEnd = null;
      for (const it of r.items) {
        if (prevEnd !== null && it.x - prevEnd > 0.6 && !/\s$/.test(line) && !/^\s/.test(it.s)) line += ' ';
        line += it.s; prevEnd = it.x + it.w;
      }
      out.push(line.replace(/\s+/g, ' ').trim());
    }
    out.push('');
  }
  return out.join('\n');
}

async function textFromDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file('word/document.xml');
  if (!f) throw new Error('File Word non valido');
  const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return Array.from(doc.getElementsByTagNameNS(W, 'p')).map(p => {
    let s = '';
    const walk = n => {
      for (const c of Array.from(n.childNodes)) {
        if (c.nodeType !== 1) continue;
        if (c.localName === 't') s += c.textContent;
        else if (c.localName === 'tab') s += ' ';
        else if (c.localName === 'br') s += ' ';
        else walk(c);
      }
    };
    walk(p);
    return s.replace(/\s+/g, ' ').trim();
  }).join('\n');
}

async function textFromFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.pdf')) return textFromPdf(await file.arrayBuffer());
  if (name.endsWith('.docx')) return textFromDocx(await file.arrayBuffer());
  if (name.endsWith('.doc')) throw new Error('Il formato .doc non è supportato: salvalo come .docx o PDF');
  return await file.text();
}

/* ------------------------------------------------------------ riconoscimento dei nomi */
const STOP = new Set(('societa società universita università azienda ospedale ospedaliera policlinico palazzo piazza viale via corso largo srl spa snc onlus fondazione istituto ' +
  'dipartimento centro congresso convegno programma informazioni segreteria organizzativa sede sabato domenica lunedi lunedì martedi martedì mercoledi mercoledì giovedi giovedì venerdi venerdì ' +
  'gennaio febbraio marzo aprile maggio giugno luglio agosto settembre ottobre novembre dicembre ' +
  'apertura chiusura lavori saluto saluti introduzione conclusione conclusioni pausa coffee break lunch pranzo registrazione iscrizione iscrizioni discussione sessione tavola rotonda ' +
  'presidente direttore direttrice rettore magnifico generale scientifico comitato faculty provider ecm crediti accreditamento patrocinio patrocini con il la lo gli le un una').split(/\s+/));
const PARTICLE = /^(de|di|del|della|dei|degli|delle|da|dal|dalla|dalle|van|von|der|den|la|le|lo|el|al|bin|ibn|mac|d['’]|dell['’]\p{L}+)$/iu;
const CAP = /^\p{Lu}[\p{L}'’\-]*$/u;
const INITIAL = /^\p{Lu}\.$/u;
const TITLE_RE = /^((?:prof(?:\.|essor[ea]?)?(?:\s*ssa)?|dott(?:\.|ore)?(?:\s*ssa|essa)?|dr(?:\.)?(?:\s*ssa)?|ing(?:\.)?|avv(?:\.)?|on(?:\.)?|sen(?:\.)?|arch(?:\.)?|mons(?:\.)?|s\.e\.|sig(?:\.|nora|nor)?(?:\s*ra)?)\s*)+/i;
const LEAD_RE = /^(moderator[ei]?|moderatrice|relator[ei]?|relatrice|chairs?|chairman|coordinator[ei]?|coordinatrice|interviene|intervengono|intervento di|a cura di|faculty|docent[ei]|speaker[s]?|presiede|presidenti?)\s*[:\-–]\s*/i;
const TIME_RE = /^\s*(\d{1,2}[:.]\d{2})(\s*[-–]\s*\d{1,2}[:.]\d{2})?\s*/;

function isNameTokens(tokens) {
  if (tokens.length < 2 || tokens.length > 5) return false;
  let caps = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (STOP.has(t.toLowerCase())) return false;
    if (CAP.test(t) || INITIAL.test(t)) { caps++; if (/\d/.test(t)) return false; }
    else if (PARTICLE.test(t) && i > 0 && i < tokens.length - 1) continue;
    else return false;
  }
  return caps >= 2 && !PARTICLE.test(tokens[0]) && !PARTICLE.test(tokens[tokens.length - 1]);
}
function asName(s) {
  s = s.trim().replace(/[.;:]+$/, '');
  if (!s || /[\d@()\/]/.test(s)) return null;
  const tokens = s.split(/\s+/);
  return isNameTokens(tokens) ? tokens.join(' ') : null;
}

/* restituisce [{ name, role, title }] in ordine di apparizione, senza doppioni */
function parseProgram(text, opts) {
  opts = opts || {};
  const lines = text.split(/\r?\n/).map(l => l.replace(/\s+/g, ' ').trim());
  const found = [], seen = new Set();
  const add = (name, role, title) => {
    const key = name.toLowerCase().replace(/[^\p{L}]+/gu, ' ').trim();
    if (seen.has(key)) return;
    seen.add(key);
    found.push({ name, role: role || '', title: title || '' });
  };

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];
    if (!line || /@|www\.|https?:/i.test(line)) continue;
    line = line.replace(TIME_RE, '').replace(LEAD_RE, '').trim();
    if (!line) continue;

    // 1) riga con titolo: "Prof. Mario Rossi, Presidente della Società ..."
    const tm = line.match(TITLE_RE);
    if (tm) {
      const rest = line.slice(tm[0].length);
      const comma = rest.search(/,|\s[–-]\s/);
      const namePart = comma >= 0 ? rest.slice(0, comma) : rest;
      const nm = asName(namePart);
      if (nm) {
        let role = comma >= 0 ? rest.slice(comma).replace(/^[,\s–-]+/, '').trim() : '';
        while (/[–-]$/.test(role) && lines[i + 1] && !TIME_RE.test(lines[i + 1])) { role = role.replace(/[–-]$/, '').trim() + ' ' + lines[++i]; }
        add(nm, role.replace(/[.;]+$/, ''), tm[0].trim());
        continue;
      }
    }

    // 2) riga di soli nomi separati da virgole / " e ": "Fabio Cirignotta, Giuseppe Loddo, Carolina Lombardi"
    const parts = line.split(/\s*,\s*|\s+e\s+|\s*;\s*/).filter(Boolean);
    const names = parts.map(p => { const t = p.match(TITLE_RE); const n = asName(t ? p.slice(t[0].length) : p); return n ? { n, t: t ? t[0].trim() : '' } : null; });
    if (parts.length && names.every(Boolean)) { names.forEach(x => add(x.n, '', x.t)); continue; }

    // 3) "Nome Cognome, Affiliazione" (solo se il primo pezzo è un nome e non c'è nulla di strano)
    if (parts.length >= 2 && names[0] && line.length < 110) {
      const role = line.slice(line.indexOf(',') + 1).trim();
      if (role && !/\d/.test(role)) add(names[0].n, role, names[0].t);
    }
  }
  return found;
}

function entryToLine(e, keepTitle, withRole) {
  const n = (keepTitle && e.title ? e.title + ' ' : '') + e.name;
  return withRole && e.role ? n + ' | ' + e.role : n;
}

window.__prog = { textFromFile, parseProgram, entryToLine };
