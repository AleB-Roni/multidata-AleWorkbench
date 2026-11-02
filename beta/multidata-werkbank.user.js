// ==UserScript==
// @name         MultiData Werkbank (Beta)
// @namespace    https://github.com/lukasebner1998-ship-it/multidata-werkbank
// @version      1.6.1.6
// @description  Ticket-Werkbank für MultiData: Suche, Tickets & Kommentare bearbeiten, Zeiten buchen, Mails senden, Anhänge, Ticket anlegen, QuickClose
// @author       Lukas Ebner
// @match        https://csp.synaforce.com/mdweb/*
// @homepageURL  https://github.com/lukasebner1998-ship-it/multidata-werkbank
// @updateURL    https://raw.githubusercontent.com/lukasebner1998-ship-it/multidata-werkbank/main/beta/multidata-werkbank.user.js
// @downloadURL  https://raw.githubusercontent.com/lukasebner1998-ship-it/multidata-werkbank/main/beta/multidata-werkbank.user.js
// @require      https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js#sha256-rMfkFFWoB2W1/Zx+4bgHim0WC7vKRVrq6FTeZclH1Z4=
// @require      https://cdn.jsdelivr.net/npm/docx-preview@0.4.1/dist/docx-preview.min.js#sha256-xKEzxloRJ5njWxQ8Vy3/+ekjo785tf07B6BqoPj1MMI=
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js#sha256-zAFRMKqFIefwiPiImOupSczcv7ON8L0Sm0S3Jzw6b0E=
// @grant        none
// @run-at       document-start
// ==/UserScript==

/*
 * Aufbau
 *  1. Konfiguration       – feste Werte; kundenspezifisches steht NICHT im Code, sondern in der Team-Konfiguration (⚙)
 *  2. Hilfsfunktionen     – Zeit, HTML, Speicher (localStorage = Browser, sessionStorage = einzelner Tab), Protokoll
 *  3. Synchronisation     – Ereignisse zwischen mehreren Werkbank-Tabs (BroadcastChannel + storage-Events)
 *  4. API-Schicht         – interne /mdapi der MultiData-Weboberfläche mit der eigenen Sitzung
 *  5. Fachfunktionen      – Tickets, Kommentare, Anhänge, Inbox, Einstellungen, Team-Konfiguration
 *  6. Oberfläche          – Tabs: Inbox | Suche | Ticket | Zeiten | Neues Ticket | Entwürfe | Protokoll | ⚙
 *
 * Alle Schreibaktionen prüfen nach dem Speichern, ob MultiData die Werte übernommen hat, und melden
 * die Änderung an alle offenen Werkbank-Tabs. Destruktive Aktionen brauchen einen zweiten Klick.
 */
(function () {
  'use strict';
  // Bibliotheken für die Datei-Vorschau (per @require geladen, mit Prüfsumme): sofort übernehmen und aus dem
  // Seitenkontext entfernen, damit MultiData eigene Versionen (z. B. JSZip für den Excel-Export) ungestört nutzt.
  /* global XLSX, docx */
  const LIBS = {
    XLSX: typeof XLSX !== 'undefined' ? XLSX : window.XLSX,
    docx: typeof docx !== 'undefined' ? docx : window.docx,
  };
  for (const k of ['JSZip', 'docx', 'XLSX']) { try { delete window[k]; } catch { /* nicht löschbar */ } }
  if (window.__mdqtLoaded || location.protocol === 'about:') return;
  window.__mdqtLoaded = true;
  // Läuft schon beim Seitenstart: Adresse sichern (#werkbank/… bzw. #mdopen/…), bevor MultiData sie umschreibt.
  const START_HASH = location.hash;
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', main, { once: true }); else main();

  function main() {

  // ===========================================================================
  // 1. Konfiguration
  // ===========================================================================
  const CFG = {
    tokenKey: 'csp.synaforce.com/mdweb#system/accessToken',
    // Alle Speicher-Schlüssel an einer Stelle (localStorage, gilt für alle Tabs dieses Browsers)
    store: {
      recent: 'mdqt.recentCustomers', filters: 'mdqt.filters', log: 'mdqt.log', drafts: 'mdqt.drafts', size: 'mdqt.size',
      settings: 'mdqt.settings', inbox: 'mdqt.inbox', sections: 'mdqt.sections', comments: 'mdqt.comments', personio: 'mdqt.personio',
      timesView: 'mdqt.timesView', templates: 'mdqt.searchTemplates', team: 'mdqt.team', facetSort: 'mdqt.facetSort',
    },
    legacyKeys: ['mdqt.myCustomers', 'mdqt.lastText', 'mdqt.graph', 'mdqt.authResult'], // Reste älterer Versionen, werden beim Start entfernt
    pollTicketSec: 45,           // offenes Ticket auf Änderungen von außen prüfen
    pollSearchSec: 120,          // offene Suchergebnisse auffrischen
    inboxIntervalMin: 5,        // Inbox automatisch alle n Minuten aktualisieren
    inboxDays: 14,              // Zeitraum für Kundenantworten / Kommentare in der Inbox
    appHash: '#werkbank',        // App-Modus: …/mdweb/rwi/HOME/0#werkbank[/ticket/<Nr> | /search | /times …]
    maxUploadMB: 25,
    fonts: ['Arial', 'Consolas', 'Calibri', 'Segoe UI', 'Verdana', 'Tahoma', 'Times New Roman', 'Courier New'],
    defaultFont: 'Arial',
    fontSizes: [8, 9, 10, 11, 12, 14, 16, 18, 24],
    defaultFontSize: 10,
    roundTo: 15,               // Buchung in 15-Minuten-Schritten, immer aufrunden
    maxHours: 12,              // Plausibilitätsgrenze pro Buchung
    defaultStartIfPast: 9 * 60, // Startzeit (Minuten) bei Buchung auf vergangene Tage
    defaultKontaktart: 'SUPPORT',
    searchLimit: 300,
    // Priorität (MultiData-Feld PRIORITY; 3 = normal)
    priorities: [[1, 'Kritisch', '#b71c1c'], [2, 'Hoch', '#ef6c00'], [3, 'Normal', '#fbc02d'], [4, 'Niedrig', '#4fc3f7'], [5, 'Sehr niedrig', '#90a4ae']],
    // Farbcodierung Arbeitstage im Zeiten-Tab
    dayColors: { low: 6 },       // 0 h = rot, < 6 h = gelb, ≥ 6 h = grün
    personioRoundMin: 7,         // Personio: ab 7 Restminuten auf die nächste Viertelstunde aufrunden
    statuses: [
      ['N', 'Neu', '#e53935'], ['NR', 'Wieder neu', '#e53935'], ['NK', 'Kontaktaufnahme', '#f9a825'],
      ['W', 'In Arbeit', '#1e3fff'], ['WE', 'Warten auf Antwort', '#455a64'], ['WF', 'Warten auf Freigabe', '#757575'],
      ['WG', 'Gelöst - Warten', '#558b2f'], ['WI', 'Warten intern', '#7b1fa2'], ['WS', 'Spätere Bearbeitung', '#6d4c41'],
      ['WU', 'Bearbeitung unterbrochen', '#9e9e9e'], ['Z', 'Abgeschlossen', '#2e7d32'],
    ],
  };
  const TICKET_COLS = ['ROW_ID', 'SID', 'TITEL', 'STATUS', 'PRIORITY', 'KND_ID', 'MATCHCODE', 'ANLAGE_DATUM', 'BEARBEITER_USER_ID',
    'SUM_STD', 'MEMO_SUM_ABRSTD', 'SUPPRESS_NOTIFY', 'DATE_MODIFIED'];
  const TICKET_DETAIL_COLS = [...TICKET_COLS, 'ANLAGE_USER_ID', 'EMAIL_TO', 'EMAIL_CC', 'ANLAGE_NOTIZ', 'MDERRDA_SID', 'ERRQ_ID', 'MDERRQ_FK__NAME_KB', 'TERMIN'];
  // MEMO_IA = '1' → eingehende Mail (Absender steht dann in EMAIL_TO)
  const COMMENT_COLS = ['ROW_ID', 'SID', 'MDERRDA_SID', 'STATUS', 'BEARBEITER_USER_ID', 'LST_TITEL', 'LST_ID', 'LST_DATUM',
    'LFZ_VON', 'LFZ_BIS', 'LFZ_STD', 'INABR_STD', 'LFZ_STD_MAN', 'NOTIZ', 'MEMO_MIN', 'EMAIL_TO', 'EMAIL_CC', 'EMAIL_BCC', 'EMAIL_DATUM', 'MEMO_IA'];
  const DOC_COLS = ['ROW_ID', 'IND_ARCHIV_DATUM', 'CREATED_BY', 'SID', 'MDERRDA_SID', 'DATUM', 'DOCTY_ID', 'IS_IDX', 'TITEL', 'DOCUMENT', 'DOCUMENT_NAME', 'ATT_KZ', 'MEMO_DMS_URL'];

  const statusInfo = id => CFG.statuses.find(s => s[0] === id) || [id, id || '–', '#999'];
  const prioInfo = p => CFG.priorities.find(x => x[0] === Number(p)) || [p, p == null ? '–' : String(p), '#bbb'];
  const prioHtml = p => { const [, n, c] = prioInfo(p); return `<span class="flag" style="background:${c}"></span>${esc(n)}`; };

  // ===========================================================================
  // 2. Hilfsfunktionen
  // ===========================================================================
  const pad = n => String(n).padStart(2, '0');
  const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const addDays = (s, n) => { const d = new Date(s + 'T12:00:00'); d.setDate(d.getDate() + n); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const fmtDate = s => s ? `${s.slice(8, 10)}.${s.slice(5, 7)}.${s.slice(0, 4)}` : '';
  const fmtDateTime = s => s ? `${fmtDate(s)} ${s.slice(11, 16)}` : '';
  const fmtH = h => (h == null || h === '') ? '–' : Number(h).toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const toMin = hhmm => { const [h, m] = String(hhmm).split(':').map(Number); return h * 60 + (m || 0); };
  const toHHMM = min => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
  const dbDT = (date, min) => `${date}T${toHHMM(min)}:00`;          // Format der DB-Felder (lokal, ohne Zone)
  const apiDT = (date, min) => `${date}T${toHHMM(min)}:00.000Z`;    // Format, das die Oberfläche beim Anlegen sendet

  // Dauer → Minuten. Standard-Einheit ist STUNDEN: "0,25" = 15 min, "1,5" / "1,5h" = 90 min,
  // "1:10" = 70 min, Minuten nur mit Einheit: "20m" / "20 min".
  function parseMinutes(s) {
    s = String(s || '').trim().toLowerCase().replace(',', '.');
    let m;
    if (!s) return NaN;
    if ((m = s.match(/^(\d+):(\d{1,2})$/))) return +m[1] * 60 + +m[2];
    if ((m = s.match(/^(\d+(?:\.\d+)?)\s*(m|min|minuten)$/))) return Math.round(parseFloat(m[1]));
    if ((m = s.match(/^(\d*\.?\d+)\s*(h|std|stunden)?$/))) return Math.round(parseFloat(m[1]) * 60);
    return NaN;
  }
  const roundUp = min => min <= 0 ? 0 : Math.ceil(min / CFG.roundTo) * CFG.roundTo;
  function parseHours(s, allowEmpty = true) {
    s = String(s ?? '').trim().replace(',', '.');
    if (!s) { if (allowEmpty) return null; throw new Error('Stunden fehlen.'); }
    const v = Number(s);
    if (!isFinite(v) || v < 0 || v > CFG.maxHours) throw new Error(`Ungültige Stunden: ${s}`);
    return Math.round(v * 100) / 100;
  }

  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function textToHtml(t) {
    let html = '', inList = false;
    for (const raw of String(t || '').split(/\r?\n/)) {
      const line = raw.trimEnd(), li = line.match(/^\s*[-*•]\s+(.*)$/);
      if (li) { if (!inList) { html += '<ul>'; inList = true; } html += `<li>${esc(li[1])}</li>`; continue; }
      if (inList) { html += '</ul>'; inList = false; }
      if (line.trim()) html += `<p>${esc(line)}</p>`;
    }
    return html + (inList ? '</ul>' : '') || '<p></p>';
  }
  // HTML aus MultiData zum Anzeigen/Bearbeiten entschärfen (keine Skripte, keine Event-Handler)
  function sanitize(html) {
    const doc = new DOMParser().parseFromString(`<div>${html || ''}</div>`, 'text/html');
    doc.querySelectorAll('script,style,iframe,object,embed,link,meta,form,input,button').forEach(n => n.remove());
    doc.querySelectorAll('*').forEach(n => [...n.attributes].forEach(a => {
      if (/^on/i.test(a.name) || (/^(href|src)$/i.test(a.name) && /^\s*javascript:/i.test(a.value))) n.removeAttribute(a.name);
    }));
    return doc.body.firstChild.innerHTML;
  }
  // Blöcke/Zeilenumbrüche als Leerzeichen, damit Wörter nicht zusammenkleben („WerkbankTest“)
  const htmlToText = html => {
    const d = document.createElement('div');
    d.innerHTML = sanitize(String(html || '').replace(/<(br|\/p|\/div|\/li|\/tr|\/td|\/th|\/h\d)[^>]*>/gi, ' $&'));
    return d.textContent.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
  };

  const store = {
    get(k, def) { try { return JSON.parse(localStorage.getItem(k)) ?? def; } catch { return def; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* ignore */ } },
  };
  // Pro Browser-Tab (bleibt bei F5 erhalten, wird beim Öffnen eines neuen Tabs per window.open kopiert)
  const tabStore = {
    get(k, def) { try { return JSON.parse(sessionStorage.getItem('mdqt.tab.' + k)) ?? def; } catch { return def; } },
    set(k, v) { try { sessionStorage.setItem('mdqt.tab.' + k, JSON.stringify(v)); } catch { /* ignore */ } },
    del(k) { try { sessionStorage.removeItem('mdqt.tab.' + k); } catch { /* ignore */ } },
  };
  // Datei im Browser herunterladen (CSV-Export, Team-Konfiguration)
  function download(name, content, type = 'text/plain;charset=utf-8') {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const a = document.createElement('a'); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 5000);
  }
  CFG.legacyKeys.forEach(k => { try { localStorage.removeItem(k); } catch { /* ignore */ } });

  // Team-Konfiguration (kundenspezifisch, nicht im Code): wird unter ⚙ gepflegt bzw. als Datei importiert
  //   nonBillable: Kunden mit Servicevertrag (Zeit standardmäßig nicht abrechenbar)
  //   intern:      eigene Firmen (Zeiten-Tab „Intern“, nie abrechenbar)
  //   othersBillable: Vorschlag für alle übrigen Kunden
  const team = () => ({ nonBillable: [], intern: [], othersBillable: true, ...store.get(CFG.store.team, {}) });
  const teamConfigured = () => { const t = store.get(CFG.store.team, null); return !!t && ('nonBillable' in t || 'intern' in t); };
  const normName = s => String(s || '').replace(/^\d+\s*-\s*/, '').toUpperCase().replace(/[^A-ZÄÖÜ0-9]+/g, ' ').trim();
  const inList = (list, kndId, name) => kndId ? list.some(c => String(c.id) === String(kndId))
    : !!name && list.some(c => { const a = normName(c.name), b = normName(name); return a && b && (a.includes(b) || b.includes(a)); });
  function billableDefault({ kndId, name }) {
    const tc = team();
    if (inList(tc.nonBillable, kndId, name) || inList(tc.intern, kndId, name)) return false;
    return tc.othersBillable !== false;
  }
  const isIntern = kndId => team().intern.some(c => String(c.id) === String(kndId));
  function log(action, detail) {
    const l = store.get(CFG.store.log, []);
    l.unshift({ t: new Date().toISOString(), action, detail });
    store.set(CFG.store.log, l.slice(0, 300));
    if (ui.tab === 'log') renderLog();
  }

  // ===========================================================================
  // 3. Synchronisation zwischen Werkbank-Tabs
  // ===========================================================================
  // bus.emit(typ, daten) erreicht diesen und alle anderen offenen Werkbank-Tabs (BroadcastChannel).
  // Änderungen an gemeinsam gespeicherten Daten (Entwürfe, Inbox, Einstellungen …) kommen als 'store'-Ereignis.
  const TAB_ID = Math.random().toString(36).slice(2, 10);
  const bus = (() => {
    const handlers = {};
    const bc = 'BroadcastChannel' in window ? new BroadcastChannel('mdqt') : null;
    const fire = (type, data, remote) => (handlers[type] || []).forEach(h => { try { h(data, remote); } catch (e) { console.warn('[Werkbank]', e); } });
    if (bc) bc.onmessage = ev => { const m = ev.data || {}; if (m.from !== TAB_ID) fire(m.type, m.data, true); };
    window.addEventListener('storage', ev => { if (ev.key && ev.key.startsWith('mdqt.')) fire('store', ev.key, true); });
    return {
      on(type, h) { (handlers[type] ||= []).push(h); },
      emit(type, data) { fire(type, data, false); try { bc?.postMessage({ type, data, from: TAB_ID }); } catch { /* ignore */ } },
    };
  })();
  // Nach jeder Schreibaktion: betroffene Tickets melden (Suche, offenes Ticket und Zeiten aktualisieren sich)
  const ticketsChanged = (...sids) => bus.emit('tickets', sids.filter(Boolean).map(Number));

  // ===========================================================================
  // 4. API-Schicht
  // ===========================================================================
  function token() {
    const raw = localStorage.getItem(CFG.tokenKey);
    if (!raw) throw new Error('Kein Sitzungs-Token – bitte neu anmelden.');
    try { return JSON.parse(raw); } catch { return raw; }
  }
  // Token läuft nach wenigen Minuten ab; MultiData erneuert es nur bei eigener Aktivität → Aktualisierung anstoßen.
  async function renewToken() {
    const old = localStorage.getItem(CFG.tokenKey);
    const btn = [...document.querySelectorAll('[aria-label="refresh"]')].find(b => b.offsetParent !== null && !b.closest('#mdqt-panel'));
    if (btn) btn.click();
    for (let i = 0; i < 40 && localStorage.getItem(CFG.tokenKey) === old; i++) await new Promise(r => setTimeout(r, 250));
    return localStorage.getItem(CFG.tokenKey) !== old;
  }
  async function api(path, body, method = 'POST', retry = true) {
    const res = await fetch('/mdapi/' + path, {
      method,
      headers: { 'Accept': 'application/json, text/plain, */*', 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token() },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    if (res.status === 401) {
      if (retry && await renewToken()) return api(path, body, method, false);
      throw new Error('Sitzung abgelaufen – in MultiData einmal „Aktualisieren“ klicken oder F5.');
    }
    if (!res.ok) {
      let msg = text.slice(0, 300);
      try { msg = JSON.parse(text).message || msg; } catch { /* ignore */ }
      throw new Error(`MultiData-Fehler (${res.status}): ${msg}`);
    }
    try { return text ? JSON.parse(text) : {}; } catch { return text; }
  }
  const load = (tableName, opts) => api('data/load/MDERRDA', { tableName, ...opts }).then(r => r.data || []);

  // Datei-Upload (wie der Kommentardialog): multipart "files[]" → Liste von Datei-IDs für Mt_Document
  async function uploadFiles(files, retry = true) {
    if (!files.length) return [];
    const fd = new FormData();
    for (const f of files) {
      if (f.size > CFG.maxUploadMB * 1024 * 1024) throw new Error(`${f.name} ist größer als ${CFG.maxUploadMB} MB.`);
      fd.append('files[]', f, f.name);
    }
    const res = await fetch('/mdapi/data/uploadFiles', { method: 'POST', headers: { 'Accept': 'application/json', 'Authorization': 'Bearer ' + token() }, body: fd });
    if (res.status === 401 && retry && await renewToken()) return uploadFiles(files, false);
    const text = await res.text();
    if (!res.ok) throw new Error(`Upload fehlgeschlagen (${res.status}): ${text.slice(0, 200)}`);
    const ids = JSON.parse(text);
    if (!Array.isArray(ids) || ids.length !== files.length) throw new Error('Upload: unerwartete Antwort ' + text.slice(0, 200));
    return ids;
  }

  // Anhang öffnen/herunterladen: 1) Download-Token holen, 2) Datei-URL im (vorab geöffneten) Fenster laden
  async function docUrl(doc, download, { pdf = true } = {}, retry = true) {
    const qs = new URLSearchParams({
      rexId: 'MDERRDA', tableName: 'MDERRDA_MDDOC', blobToPdf: pdf ? 'true' : 'false',
      filter: JSON.stringify([[['ROW_ID', '=', doc.ROW_ID]], 'and', [['MDERRDA_SID', '=', Number(doc.MDERRDA_SID)]]]),
      sqlColumns: JSON.stringify(DOC_COLS), options: JSON.stringify([{ column: 'MDDOC.DOCUMENT' }]), // muss eine Liste sein
    });
    const res = await fetch('/mdapi/data/downloadFile?' + qs, { headers: { 'Accept': 'application/json', 'Authorization': 'Bearer ' + token() } });
    if (res.status === 401 && retry && await renewToken()) return docUrl(doc, download, { pdf }, false);
    if (!res.ok) throw new Error(`Anhang „${doc.DOCUMENT_NAME}“ nicht abrufbar (${res.status}).`);
    const fi = (await res.json()).fileInfo?.[0];
    if (!fi?.actionToken) throw new Error('Kein Download-Token erhalten.');
    return { url: `/mdapi/action/DownloadFile/${encodeURIComponent(fi.filename)}?tok=${fi.actionToken}&download=${download ? 'true' : 'false'}`, fi };
  }
  async function openDocument(doc, download, win) {
    const { url } = await docUrl(doc, download);
    if (win && !win.closed) win.location.href = url; else window.open(url, '_blank');
  }
  // Anhang als Datei laden (Original, nicht als PDF) – z. B. zum Weiterleiten
  async function docAsFile(doc) {
    const { url, fi } = await docUrl(doc, true, { pdf: false });
    const r = await fetch(url);
    if (!r.ok) throw new Error(`Anhang „${doc.DOCUMENT_NAME}“ konnte nicht geladen werden (${r.status}).`);
    const blob = await r.blob();
    return new File([blob], doc.DOCUMENT_NAME || fi.filename, { type: blob.type || fi.mimeType || 'application/octet-stream' });
  }
  // --- Office-Dateien lokal öffnen ---
  // Die Originaldatei wird heruntergeladen; der Browser öffnet sie im installierten Office (Word, Excel, PowerPoint, Visio).
  // Ohne Rückfrage geht das, wenn im Browser „Dateien dieses Typs immer öffnen“ gesetzt ist (einmal je Dateityp).
  const OFFICE_EXT = new Set(['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'odt', 'xls', 'xlsx', 'xlsm', 'xlsb', 'ods', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'pot', 'potx', 'odp', 'vsd', 'vsdx']);
  const fileExt = n => (String(n || '').toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1] || '';
  async function openLocal(doc) {
    const { url, fi } = await docUrl(doc, true, { pdf: false });
    const a = document.createElement('a'); a.href = url; a.download = doc.DOCUMENT_NAME || fi.filename || 'Datei'; a.rel = 'noopener';
    document.body.appendChild(a); a.click(); a.remove();
  }
  const OFFICE_ACK = 'mdqt.officeAck'; // Dateitypen, für die „immer öffnen“ eingerichtet wurde (bzw. Hinweis aus)
  const officeAck = () => store.get(OFFICE_ACK, { all: false, ext: [] });
  const officeHintNeeded = ext => !officeAck().all && !officeAck().ext.includes(ext);
  const isEdge = () => /\bEdg\//.test(navigator.userAgent);
  const officeHowTo = ext => isEdge()
    ? `Edge: oben rechts erscheint der Download. Dort bei der .${ext}-Datei auf „…“ klicken → „Dateien dieses Typs immer öffnen“.`
    : `Chrome: oben rechts auf das Download-Symbol, dann mit der rechten Maustaste auf die .${ext}-Datei → „Dateien dieses Typs immer öffnen“.`;
  // Fenster-Hilfen für das vorab geöffnete Fenster (Popup-Blocker)
  function winMessage(win, text) {
    try { win.document.open(); win.document.write(`<!doctype html><meta charset="utf-8"><title>Werkbank</title><body style="font:15px Segoe UI,Arial,sans-serif;padding:40px;color:#333"><h2 style="color:#00796b">⚡ Werkbank</h2><p>${esc(text)}</p></body>`); win.document.close(); } catch { /* fremde Seite */ }
  }

  // --- .eml-Vorschau (lokal im Browser, nichts wird hochgeladen) ---
  const bin2bytes = s => { const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 255; return b; };
  const bytes2bin = b => { let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode.apply(null, b.subarray(i, i + 32768)); return s; };
  const decodeCharset = (bytes, cs) => { try { return new TextDecoder((cs || 'utf-8').toLowerCase().replace(/^(us-ascii|ascii)$/, 'utf-8')).decode(bytes); } catch { return new TextDecoder('windows-1252').decode(bytes); } };
  const qpDecode = s => bin2bytes(s.replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))));
  const b64Decode = s => { try { return bin2bytes(atob(s.replace(/[^A-Za-z0-9+/=]/g, ''))); } catch { return new Uint8Array(0); } };
  // RFC 2047: =?utf-8?B?…?= / =?iso-8859-1?Q?…?=
  const decodeWords = v => String(v || '').replace(/(=\?[^?]+\?[bBqQ]\?[^?]*\?=)\s+(?==\?)/g, '$1').replace(/=\?([^?*]+)(?:\*[^?]*)?\?([bBqQ])\?([^?]*)\?=/g, (_, cs, enc, txt) =>
    decodeCharset(enc.toUpperCase() === 'B' ? b64Decode(txt) : qpDecode(txt.replace(/_/g, ' ')), cs));
  function headerParams(v) {
    const parts = String(v || '').split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/), out = { value: parts.shift().trim().toLowerCase() };
    const ext = {};
    for (const p of parts) {
      const m = p.match(/^\s*([^=\s]+)\s*=\s*(.*)$/s); if (!m) continue;
      let [, k, val] = m; k = k.toLowerCase(); val = val.trim().replace(/^"(.*)"$/s, '$1');
      const star = k.match(/^([^*]+)\*(\d+)?(\*)?$/);
      if (star) { (ext[star[1]] ||= []).push({ n: +(star[2] || 0), enc: !!star[3] || (!star[2] && k.endsWith('*')), val }); continue; }
      out[k] = decodeWords(val);
    }
    for (const [k, segs] of Object.entries(ext)) { // RFC 2231
      segs.sort((a, b) => a.n - b.n);
      let cs = 'utf-8', raw = '';
      segs.forEach((sg, i) => { let v = sg.val; if (sg.enc && i === 0) { const mm = v.match(/^([^']*)'[^']*'(.*)$/s); if (mm) { cs = mm[1] || cs; v = mm[2]; } } raw += sg.enc ? v.replace(/%([0-9A-Fa-f]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))) : v; });
      out[k] = decodeCharset(bin2bytes(raw), cs);
    }
    return out;
  }
  function parseMime(bin) {
    const sep = bin.search(/\r?\n\r?\n/);
    const head = sep < 0 ? bin : bin.slice(0, sep), body = sep < 0 ? '' : bin.slice(sep).replace(/^\r?\n\r?\n/, '');
    const headers = {};
    for (const line of head.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/)) { const i = line.indexOf(':'); if (i > 0) { const k = line.slice(0, i).trim().toLowerCase(); if (!(k in headers)) headers[k] = line.slice(i + 1).trim(); } }
    const ct = headerParams(headers['content-type'] || 'text/plain'), cd = headerParams(headers['content-disposition'] || '');
    const part = { headers, type: ct.value, ct, cd, cid: (headers['content-id'] || '').replace(/[<>]/g, '').trim(), name: cd.filename || ct.name || '' };
    if (part.type.startsWith('multipart/') && ct.boundary) {
      const b = '--' + ct.boundary, chunks = body.split(new RegExp('\\r?\\n?' + b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:--)?[ \\t]*(?:\\r?\\n|$)'));
      part.parts = chunks.slice(1).filter(c => c.trim()).map(parseMime);
    } else {
      const te = (headers['content-transfer-encoding'] || '').toLowerCase();
      part.bytes = te === 'base64' ? b64Decode(body) : te === 'quoted-printable' ? qpDecode(body) : bin2bytes(body);
    }
    return part;
  }
  function mailContent(root) {
    let html = null, text = null; const atts = [], cids = {};
    const walk = p => {
      if (p.parts) {
        if (p.type === 'multipart/alternative') {
          const pick = [...p.parts].reverse().find(x => x.type === 'text/html' || x.parts) || p.parts[p.parts.length - 1];
          p.parts.forEach(x => { if (x !== pick && !x.parts && x.cd.value === 'attachment') atts.push(x); });
          return walk(pick);
        }
        return p.parts.forEach(walk);
      }
      const isAtt = p.cd.value === 'attachment' || (!!p.name && !/^text\/(html|plain)$/.test(p.type)) || p.type === 'message/rfc822';
      if (!isAtt && p.type === 'text/html' && html === null) html = decodeCharset(p.bytes, p.ct.charset);
      else if (!isAtt && p.type === 'text/plain' && text === null && html === null) text = decodeCharset(p.bytes, p.ct.charset);
      else if (p.cid && p.type.startsWith('image/') && p.cd.value !== 'attachment') cids[p.cid] = p;
      else if (p.bytes && (isAtt || !/^text\//.test(p.type))) atts.push(p);
    };
    walk(root);
    // Bilder, die nicht im Text verwendet werden, als Anhang zeigen
    for (const [cid, p] of Object.entries(cids)) if (!html || !html.includes('cid:' + cid)) atts.push(p);
    return { html, text, atts, cids };
  }
  // .eml im vorab geöffneten Fenster anzeigen. Mailinhalt läuft in einer abgeschotteten Ansicht ohne Skripte;
  // externe Bilder werden erst auf Klick geladen.
  function renderEml(win, bytes, fallbackName) {
    const root = parseMime(bytes2bin(bytes)), h = root.headers, c = mailContent(root);
    const toData = p => `data:${p.type};base64,${btoa(bytes2bin(p.bytes))}`;
    let body = c.html ?? `<pre style="white-space:pre-wrap;font:13px Consolas,monospace">${esc(c.text || '')}</pre>`;
    body = body.replace(/(["'(])cid:([^"')\s>]+)/gi, (m, q, id) => { const p = c.cids[decodeURIComponent(id)] || c.cids[id]; return p ? q + toData(p) : m; });
    const remote = /<img[^>]+src\s*=\s*["']?https?:/i.test(body) || /url\(\s*["']?https?:/i.test(body);
    const frame = allow => `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">
      <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: ${allow ? 'https: http:' : ''}; style-src 'unsafe-inline' ${allow ? 'https:' : ''}; font-src data: ${allow ? 'https:' : ''}">
      <style>body{font:14px Calibri,Segoe UI,Arial,sans-serif;margin:12px;color:#222;background:#fff}img{max-width:100%;height:auto}</style></head><body>${body}</body></html>`;
    const subj = decodeWords(h.subject || '') || fallbackName || '(ohne Betreff)';
    const fmtD = v => { const d = new Date(v); return isNaN(d) ? (v || '') : d.toLocaleString('de-DE'); };
    const rows = [['Von', h.from], ['An', h.to], ['Cc', h.cc], ['Datum', h.date && fmtD(h.date)], ['Betreff', h.subject]].filter(r => r[1]).map(([k, v]) => `<tr><th>${k}</th><td>${esc(decodeWords(v))}</td></tr>`).join('');
    const d = win.document;
    d.open(); d.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(subj)}</title><style>
      body{margin:0;font:13px Segoe UI,Arial,sans-serif;color:#222;background:#f4f6f6;display:flex;flex-direction:column;height:100vh}
      header{background:#fff;border-bottom:1px solid #ddd;padding:12px 16px}h1{font-size:18px;margin:0 0 8px;color:#00695c}
      table{border-collapse:collapse}th{text-align:left;color:#666;font-weight:600;padding:2px 12px 2px 0;vertical-align:top;white-space:nowrap}td{padding:2px 0}
      .att{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}.att a{border:1px solid #00897b;border-radius:12px;padding:2px 10px;color:#00695c;text-decoration:none;background:#fff}
      .bar{margin-top:8px;color:#666}.bar button{border:1px solid #00897b;background:#e0f2f1;color:#00695c;border-radius:4px;padding:3px 10px;cursor:pointer}
      iframe{flex:1;border:0;background:#fff;width:100%}</style></head><body>
      <header><h1></h1><table>${rows}</table><div class="att"></div><div class="bar"></div></header><iframe sandbox="allow-popups allow-popups-to-escape-sandbox"></iframe></body></html>`);
    d.close();
    d.querySelector('h1').textContent = '✉ ' + subj;
    const ifr = d.querySelector('iframe'); ifr.srcdoc = frame(false);
    if (remote) {
      const b = d.createElement('button'); b.textContent = 'Externe Bilder laden';
      const note = d.createElement('span'); note.textContent = ' Externe Bilder wurden zum Schutz der Privatsphäre blockiert. ';
      b.addEventListener('click', () => { ifr.srcdoc = frame(true); d.querySelector('.bar').remove(); });
      d.querySelector('.bar').append(note, b);
    }
    const ab = d.querySelector('.att');
    for (const p of c.atts) {
      const name = p.name || (p.type === 'message/rfc822' ? 'Nachricht.eml' : 'Anhang');
      const a = d.createElement('a'); a.textContent = '📎 ' + name + ` (${Math.max(1, Math.round(p.bytes.length / 1024))} KB)`;
      const blob = new Blob([p.bytes], { type: p.type || 'application/octet-stream' });
      a.href = win.URL.createObjectURL(blob);
      const kind = previewKind(fileExt(name)) || (/^image\//.test(p.type) ? 'image' : /^video\//.test(p.type) ? 'video' : null);
      if (/^application\/pdf/.test(p.type) || /\.pdf$/i.test(name)) { a.target = '_blank'; a.title = 'Klick: im Browser öffnen'; } else a.download = name;
      if (kind) {
        a.title = 'Klick: Vorschau · Rechtsklick: speichern';
        a.addEventListener('click', ev => { ev.preventDefault(); const w2 = win.open('about:blank', '_blank'); if (w2) renderPreview(w2, p.bytes, name, kind).catch(e => winMessage(w2, '✖ ' + (e.message || e))); });
      }
      if (p.type === 'message/rfc822' || fileExt(name) === 'eml') {
        a.title = 'Klick: Vorschau · Rechtsklick: speichern';
        a.addEventListener('click', ev => { ev.preventDefault(); const w2 = win.open('about:blank', '_blank'); if (w2) renderEml(w2, p.bytes, name); });
      }
      ab.appendChild(a);
    }
    if (!c.atts.length) ab.remove();
  }
  async function openEml(doc, win) {
    winMessage(win, `„${doc.DOCUMENT_NAME}“ wird geladen …`);
    const f = await docAsFile(doc);
    renderEml(win, new Uint8Array(await f.arrayBuffer()), doc.DOCUMENT_NAME);
  }

  // --- Vorschau im Browser: Bilder, Video, Audio, Text, Word (.docx), Excel (.xlsx/.xls/.ods/.csv) ---
  // Die Datei wird nur in den Arbeitsspeicher geladen und im Fenster angezeigt – kein Download, kein Upload.
  // Dokumentinhalte laufen wie bei .eml in einer abgeschotteten Ansicht ohne Skripte.
  const PREVIEW = {
    image: ['png', 'jpg', 'jpeg', 'jfif', 'gif', 'webp', 'bmp', 'svg', 'avif', 'ico'],
    video: ['mp4', 'm4v', 'webm', 'mov', 'ogv', 'mkv'],
    audio: ['mp3', 'wav', 'm4a', 'aac', 'ogg', 'oga', 'opus', 'flac'],
    text: ['txt', 'log', 'json', 'xml', 'ini', 'cfg', 'conf', 'md', 'ps1', 'psm1', 'bat', 'cmd', 'sh', 'yml', 'yaml', 'reg'],
    word: ['docx', 'docm', 'dotx'],
    sheet: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv'],
  };
  const previewKind = ext => Object.keys(PREVIEW).find(k => PREVIEW[k].includes(ext)) || null;
  const MEDIA_MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jfif: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp',
    svg: 'image/svg+xml', avif: 'image/avif', ico: 'image/x-icon', mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg',
    mkv: 'video/webm', mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', aac: 'audio/aac', ogg: 'audio/ogg', oga: 'audio/ogg', opus: 'audio/ogg', flac: 'audio/flac' };
  const PREVIEW_MAX_ROWS = 2000; // Excel: je Blatt höchstens so viele Zeilen anzeigen
  const PREVIEW_MAX_TEXT = 2 * 1024 * 1024;

  // Grundgerüst des Vorschaufensters: Kopfzeile mit Name, Größe, Herunterladen; darunter der Inhalt
  function previewShell(win, name, bytes, icon) {
    const d = win.document;
    d.open(); d.write(`<!doctype html><html><head><meta charset="utf-8"><title></title><style>
      body{margin:0;font:13px Segoe UI,Arial,sans-serif;color:#222;background:#f4f6f6;display:flex;flex-direction:column;height:100vh}
      header{background:#fff;border-bottom:1px solid #ddd;padding:10px 16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
      h1{font-size:17px;margin:0;color:#00695c;flex:1;min-width:200px;word-break:break-all}.sz{color:#777}
      header button,.tabs button{border:1px solid #00897b;background:#e0f2f1;color:#00695c;border-radius:4px;padding:4px 10px;cursor:pointer;font:inherit}
      .tabs{display:flex;gap:4px;flex-wrap:wrap;padding:6px 16px 0;background:#fff;border-bottom:1px solid #ddd}
      .tabs button{border-radius:4px 4px 0 0;border-bottom:none;background:#f4f6f6}.tabs button.on{background:#00897b;color:#fff}
      .note{padding:6px 16px;color:#8a5a00;background:#fff8e1;border-bottom:1px solid #ffe0a0}
      main{flex:1;min-height:0;display:flex;overflow:auto}
      .stage{margin:auto;padding:16px;text-align:center}.stage img{max-width:calc(100vw - 32px);max-height:calc(100vh - 110px);cursor:zoom-in;background:#fff;
        background-image:linear-gradient(45deg,#eee 25%,transparent 25%,transparent 75%,#eee 75%),linear-gradient(45deg,#eee 25%,transparent 25%,transparent 75%,#eee 75%);
        background-size:16px 16px;background-position:0 0,8px 8px;box-shadow:0 1px 6px #0002}
      .stage img.full{max-width:none;max-height:none;cursor:zoom-out}
      .stage video{max-width:calc(100vw - 32px);max-height:calc(100vh - 110px);background:#000}.stage audio{width:min(640px,calc(100vw - 32px))}
      pre{margin:0;padding:16px;white-space:pre-wrap;word-break:break-word;font:12.5px Consolas,monospace;background:#fff;flex:1}
      iframe{flex:1;border:0;background:#fff;width:100%}
      .err{margin:auto;padding:24px;max-width:560px;text-align:center;color:#555}</style></head><body>
      <header><h1></h1><span class="sz"></span><button data-a="dl">⬇ Herunterladen</button></header><main></main></body></html>`);
    d.close();
    d.title = name;
    d.querySelector('h1').textContent = `${icon} ${name}`;
    d.querySelector('.sz').textContent = bytes.length >= 1048576 ? (bytes.length / 1048576).toFixed(1).replace('.', ',') + ' MB' : Math.max(1, Math.round(bytes.length / 1024)) + ' KB';
    d.querySelector('[data-a=dl]').addEventListener('click', () => {
      const a = d.createElement('a'); a.href = win.URL.createObjectURL(new win.Blob([bytes], { type: 'application/octet-stream' })); a.download = name;
      d.body.appendChild(a); a.click(); a.remove();
    });
    return { d, main: d.querySelector('main'), header: d.querySelector('header') };
  }
  const previewFail = (main, text) => { const e = main.ownerDocument.createElement('div'); e.className = 'err'; e.textContent = text; main.replaceChildren(e); };
  // Abgeschottete Ansicht für Dokumentinhalte (keine Skripte, nur eingebettete Bilder/Schriften, Links in neuem Tab)
  const sandboxDoc = (css, body) => `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src data:">
    <style>${css}</style></head><body>${body}</body></html>`;
  function sandboxFrame(d, html) {
    const f = d.createElement('iframe');
    f.setAttribute('sandbox', 'allow-popups allow-popups-to-escape-sandbox');
    f.srcdoc = html; return f;
  }

  async function renderPreview(win, bytes, name, kind = previewKind(fileExt(name))) {
    const ext = fileExt(name);
    const icon = { image: '🖼', video: '🎬', audio: '🎵', text: '📄', word: '📝', sheet: '📊' }[kind] || '📄';
    const { d, main, header } = previewShell(win, name, bytes, icon);
    const blobUrl = type => win.URL.createObjectURL(new win.Blob([bytes], { type }));
    if (kind === 'image') {
      const st = d.createElement('div'); st.className = 'stage';
      const img = d.createElement('img'); img.alt = name; img.src = blobUrl(MEDIA_MIME[ext] || 'image/*');
      img.title = 'Klick: Originalgröße / einpassen';
      img.addEventListener('click', () => img.classList.toggle('full'));
      img.addEventListener('error', () => previewFail(main, 'Dieses Bild kann der Browser nicht anzeigen. Bitte herunterladen.'));
      st.appendChild(img); main.appendChild(st);
    } else if (kind === 'video' || kind === 'audio') {
      const st = d.createElement('div'); st.className = 'stage';
      const m = d.createElement(kind); m.controls = true; m.preload = 'metadata'; m.src = blobUrl(MEDIA_MIME[ext] || kind + '/*');
      m.addEventListener('error', () => previewFail(main, `Dieses ${kind === 'video' ? 'Video' : 'Audio'}format (.${ext}) kann der Browser nicht abspielen. Bitte herunterladen.`));
      st.appendChild(m); main.appendChild(st);
    } else if (kind === 'text') {
      let t = decodeCharset(bytes.subarray(0, PREVIEW_MAX_TEXT), 'utf-8');
      if (t.includes('�')) t = decodeCharset(bytes.subarray(0, PREVIEW_MAX_TEXT), 'windows-1252');
      const pre = d.createElement('pre'); pre.textContent = t + (bytes.length > PREVIEW_MAX_TEXT ? '\n\n… (gekürzt – vollständig über „Herunterladen“)' : '');
      main.appendChild(pre);
    } else if (kind === 'word') {
      if (!LIBS.docx?.renderAsync) return previewFail(main, 'Word-Vorschau nicht verfügbar (Bibliothek nicht geladen). Bitte „In Office öffnen“.');
      previewFail(main, 'Dokument wird aufbereitet …');
      const body = document.createElement('div'), style = document.createElement('div');
      try {
        await LIBS.docx.renderAsync(new Blob([bytes]), body, style, { className: 'docx', inWrapper: true, useBase64URL: true, breakPages: true,
          ignoreLastRenderedPageBreak: true, renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true, experimental: false });
      } catch (e) { return previewFail(main, 'Dieses Word-Dokument lässt sich nicht als Vorschau anzeigen (' + (e.message || e) + '). Bitte „In Office öffnen“.'); }
      const css = 'body{margin:0;background:#f4f6f6}.docx-wrapper{background:#f4f6f6 !important;padding:16px !important}.docx-wrapper>section.docx{box-shadow:0 1px 6px #0002 !important;margin-bottom:16px !important}';
      main.replaceChildren(sandboxFrame(d, sandboxDoc(css, style.innerHTML + body.innerHTML)));
    } else if (kind === 'sheet') {
      const X = LIBS.XLSX;
      if (!X?.read) return previewFail(main, 'Excel-Vorschau nicht verfügbar (Bibliothek nicht geladen). Bitte „In Office öffnen“.');
      let wb;
      const opts = { sheetRows: PREVIEW_MAX_ROWS + 1, cellDates: true, cellStyles: false, cellFormula: false, dense: true };
      try {
        if (ext === 'csv') { // Zeichensatz selbst erkennen (UTF-8, sonst Windows-1252 wie aus Excel)
          let txt = decodeCharset(bytes, 'utf-8'); if (txt.includes('\uFFFD')) txt = decodeCharset(bytes, 'windows-1252');
          wb = X.read(txt.replace(/^\uFEFF/, ''), { ...opts, type: 'string' });
        } else wb = X.read(bytes, { ...opts, type: 'array' });
      }
      catch (e) { return previewFail(main, 'Diese Tabelle lässt sich nicht als Vorschau anzeigen (' + (e.message || e) + '). Bitte „In Office öffnen“.'); }
      const names = (wb.SheetNames || []).filter(n => !wb.Workbook?.Sheets?.find(s => s.name === n && s.Hidden));
      if (!names.length) return previewFail(main, 'Die Tabelle enthält keine sichtbaren Blätter.');
      const css = `body{margin:0;font:12.5px Calibri,Segoe UI,Arial,sans-serif;color:#222}table{border-collapse:collapse}
        td,th{border:1px solid #d4d4d4;padding:2px 6px;white-space:pre-wrap;vertical-align:top;max-width:480px}
        tr:first-child td{background:#f3f6f6;font-weight:600;position:sticky;top:0}`;
      const tabs = d.createElement('div'); tabs.className = 'tabs';
      const note = d.createElement('div'); note.className = 'note'; note.style.display = 'none';
      const show = n => {
        const ws = wb.Sheets[n];
        tabs.querySelectorAll('button').forEach(b => b.classList.toggle('on', b.dataset.n === n));
        const full = ws['!fullref'] && ws['!ref'] && ws['!fullref'] !== ws['!ref'];
        const total = full ? X.utils.decode_range(ws['!fullref']).e.r + 1 : 0;
        note.textContent = full ? `Vorschau zeigt die ersten ${PREVIEW_MAX_ROWS.toLocaleString('de-DE')} von ${total.toLocaleString('de-DE')} Zeilen. Vollständig: „In Office öffnen“.` : '';
        note.style.display = full ? '' : 'none';
        const html = ws['!ref'] ? X.utils.sheet_to_html(ws, { header: '', footer: '' }) : '<p style="padding:16px;color:#777">(leeres Blatt)</p>';
        main.replaceChildren(sandboxFrame(d, sandboxDoc(css, html)));
      };
      for (const n of names) {
        const b = d.createElement('button'); b.textContent = n; b.dataset.n = n;
        b.addEventListener('click', () => show(n)); tabs.appendChild(b);
      }
      if (names.length > 1 || ext !== 'csv') header.after(tabs);
      (names.length > 1 || ext !== 'csv' ? tabs : header).after(note);
      show(names[0]);
    } else {
      previewFail(main, `Für .${ext}-Dateien gibt es keine Vorschau. Bitte herunterladen bzw. in Office öffnen.`);
    }
  }
  async function openPreview(doc, win) {
    winMessage(win, `„${doc.DOCUMENT_NAME}“ wird geladen …`);
    const f = await docAsFile(doc);
    await renderPreview(win, new Uint8Array(await f.arrayBuffer()), doc.DOCUMENT_NAME || f.name);
  }

  const exec = (workItem, commandName, action, data, extra = {}) =>
    api(`data/execCommand/${workItem}`, { workItemName: workItem, commandName, rexId: workItem, action, data, ...extra });

  // Generisches Update (wie die Oberfläche): nur geänderte Felder + Originalwerte; prüft danach das Ergebnis.
  async function updateRow(tableName, row, changes, cols) {
    const values = {}, originalValues = {};
    for (const [k, v] of Object.entries(changes)) {
      if (row[k] === v || (row[k] == null && v == null)) continue;
      values[k] = v; originalValues[k] = row[k] ?? null;
    }
    if (!Object.keys(values).length) return null; // nichts geändert
    const r = await api('data/MDERRDA', { tableName, key: { ROW_ID: row.ROW_ID }, values, originalValues, dbParameters: [], sqlColumns: cols }, 'PUT');
    const saved = r.data && r.data[0];
    if (!saved) throw new Error('Keine Antwort nach dem Speichern – bitte neu laden und prüfen.');
    const bad = Object.keys(values).filter(k => !['NOTIZ', 'LST_DATUM'].includes(k) && String(saved[k] ?? '').replace(/\.000Z$/, '') !== String(values[k] ?? '').replace(/\.000Z$/, '') && Number(saved[k]) !== Number(values[k]));
    if (bad.length) throw new Error('Nicht übernommen: ' + bad.join(', ') + ' – MultiData hat die Werte verändert oder abgelehnt.');
    return saved;
  }
  const deleteRows = (tableName, rowIds) => api('data/MDERRDA', { tableName, key: null, rowIds, values: {} }, 'DELETE');

  // ===========================================================================
  // 5. Fachfunktionen
  // ===========================================================================
  let ME = null;
  async function me() {
    if (ME) return ME;
    const d = await api('data/InitNewRow/DLG_TICKET_CREATE', { tableName: 'DLG_TICKET_CREATE', values: { MEMO_BATCH: 'true' }, copy: false });
    ME = d.BEARBEITER_USER_ID;
    return ME;
  }

  // --- Tickets ---
  const ticketBy = async filter => (await load('MDERRDA', { filter, sqlColumns: TICKET_DETAIL_COLS }))[0] || null;
  const ticketBySid = sid => ticketBy([['SID', '=', Number(sid)]]);
  const ticketByRowId = id => ticketBy([['ROW_ID', '=', id]]);
  const orFilter = (field, vals) => vals.map(v => [field, '=', v]).flatMap((f, i) => i ? ['or', f] : [f]);
  const andAll = parts => parts.filter(Boolean).flatMap((p, i) => i ? ['and', p] : [p]);
  // „leer“ als Filterwert (z. B. Tickets ohne Bearbeiter / ohne Queue)
  const NONE = '__NONE__';
  const isEmptyVal = v => v === null || v === undefined || v === '' || v === 0;
  function valueFilter(field, vals, { numeric = false } = {}) {
    const real = vals.filter(v => v !== NONE).map(v => numeric ? Number(v) : v);
    const conds = real.map(v => [field, '=', v]);
    if (vals.includes(NONE)) { conds.push([field, '=', null]); if (!numeric) conds.push([field, '=', '']); }
    return conds.flatMap((f, i) => i ? ['or', f] : [f]);
  }
  // Sicherheitsnetz: Ergebnis zusätzlich im Browser prüfen (falls MultiData „= null“ anders auslegt)
  const matchesVals = (v, vals, numeric) => vals.some(x => x === NONE ? isEmptyVal(v) : (numeric ? Number(x) === Number(v) : String(x) === String(v)));

  async function currentPageTicket() {
    const m = decodeURIComponent(location.pathname).match(/\/rwi\/MDERRDA\/(\{.*\})/);
    if (!m) return null;
    try { const k = JSON.parse(m[1]); return k.SID ? ticketBySid(k.SID) : k.ROW_ID ? ticketByRowId(k.ROW_ID) : null; } catch { return null; }
  }

  // Suche. f = { users[] ('__ME__' = ich, leer = alle), kndIds[], queues[], statuses[], dateMode: 'created'|'booked'|'none', from, to, text }
  // opts.cols/take erlauben schlanke Abfragen (z. B. für die Kunden-Übersicht)
  async function searchTickets(f, opts = {}) {
    const cols = opts.cols || TICKET_COLS, take = opts.take || CFG.searchLimit;
    const meId = await me();
    const users = [...new Set((f.users || []).map(u => u === '__ME__' ? meId : u))];
    const queues = f.queues || [];
    const userFilter = users.length ? valueFilter('BEARBEITER_USER_ID', users) : null;
    const parts = [];
    let bookedHours = null;
    if (f.statuses.length) parts.push(orFilter('STATUS', f.statuses));
    if (queues.length) parts.push(valueFilter('ERRQ_ID', queues, { numeric: true }));
    if (f.kndIds.length) parts.push(orFilter('KND_ID', f.kndIds));
    if (f.text) parts.push(/^\d{5,}$/.test(f.text) ? ['SID', '=', Number(f.text)] : ['TITEL', 'contains', f.text]);

    if (f.dateMode === 'booked') {
      // Tickets, auf die im Zeitraum Zeit gebucht wurde (von mir bzw. von allen)
      const realUsers = users.filter(u => u !== NONE);
      if (users.includes(NONE)) parts.push(valueFilter('BEARBEITER_USER_ID', [NONE])); // Ticket ohne Bearbeiter
      const cf = andAll([
        realUsers.length ? orFilter('BEARBEITER_USER_ID', realUsers) : null,
        f.from ? ['LST_DATUM', '>=', f.from + 'T00:00:00'] : null,
        f.to ? ['LST_DATUM', '<', addDays(f.to, 1) + 'T00:00:00'] : null,
      ]);
      const comments = await load('MDERRDA_MDERRDAC', { take: 5000, filter: cf.length ? cf : undefined, sqlColumns: ['SID', 'MDERRDA_SID', 'LFZ_STD', 'INABR_STD'] });
      bookedHours = {};
      for (const c of comments) {
        const b = bookedHours[c.MDERRDA_SID] ||= { std: 0, abr: 0 };
        b.std += c.LFZ_STD || 0; b.abr += c.INABR_STD || 0;
      }
      const sids = Object.keys(bookedHours).map(Number);
      if (!sids.length) return { rows: [], bookedHours };
      const rows = [];
      for (let i = 0; i < sids.length; i += 60) {
        const chunk = sids.slice(i, i + 60);
        rows.push(...await load('MDERRDA', { take: 1000, filter: andAll([orFilter('SID', chunk), ...parts]), sqlColumns: cols }));
      }
      rows.sort((a, b) => b.SID - a.SID);
      const ok = rows.filter(r => !queues.length || !('ERRQ_ID' in r) || matchesVals(r.ERRQ_ID, queues, true));
      return { rows: ok.slice(0, take), total: ok.length, bookedHours };
    }

    if (userFilter) parts.push(userFilter);
    if (f.dateMode === 'created') {
      if (f.from) parts.push(['ANLAGE_DATUM', '>=', f.from + 'T00:00:00']);
      if (f.to) parts.push(['ANLAGE_DATUM', '<', addDays(f.to, 1) + 'T00:00:00']);
    }
    const r = await api('data/load/MDERRDA', {
      tableName: 'MDERRDA', take, requireTotalCount: true,
      filter: parts.length ? andAll(parts) : undefined, sort: [{ selector: 'SID', desc: true }], sqlColumns: cols,
    });
    let rows = r.data || [];
    const n0 = rows.length;
    if (users.includes(NONE) && rows.length && 'BEARBEITER_USER_ID' in rows[0]) rows = rows.filter(x => matchesVals(x.BEARBEITER_USER_ID, users, false));
    if (queues.includes(NONE) && rows.length && 'ERRQ_ID' in rows[0]) rows = rows.filter(x => matchesVals(x.ERRQ_ID, queues, true));
    return { rows, total: rows.length === n0 ? r.totalCount : rows.length, bookedHours: null };
  }

  // Kunden-Übersicht: welche Kunden kommen in den Tickets der aktuellen Filter (ohne Kundenfilter) vor?
  async function customerFacets(f) {
    const res = await searchTickets({ ...f, kndIds: [] }, { cols: ['SID', 'KND_ID', 'MATCHCODE'], take: 5000 });
    const map = {};
    for (const r of res.rows) if (r.KND_ID) (map[r.KND_ID] ||= { id: r.KND_ID, name: r.MATCHCODE || r.KND_ID, n: 0 }).n++;
    return { list: Object.values(map).sort((a, b) => b.n - a.n), truncated: (res.total || 0) > res.rows.length };
  }
  // Kundensuche über alle Tickets (Kurzname/Nummer enthält q)
  async function searchCustomersByTickets(q) {
    const rows = await load('MDERRDA', { take: 3000, filter: [/^\d+$/.test(q) ? ['KND_ID', '=', q] : ['MATCHCODE', 'contains', q]], sort: [{ selector: 'SID', desc: true }], sqlColumns: ['SID', 'KND_ID', 'MATCHCODE'] });
    const map = {};
    for (const r of rows) if (r.KND_ID) (map[r.KND_ID] ||= { id: r.KND_ID, name: r.MATCHCODE || r.KND_ID, n: 0 }).n++;
    return Object.values(map).sort((a, b) => b.n - a.n);
  }

  async function searchAddresses(q) {
    const r = await api('data/load/DLG_TICKET_CREATE', {
      tableName: 'DLG_TICKET_CREATE', lookupColumnName: 'MEMO_ADRSEARCH', skip: 0, take: 20, searchValue: q, searchExpr: ['MATCH'], options: [], userData: {},
    });
    return r.data || [];
  }

  // Kundennummer zu einer Adresse (Firmensuche „Neues Ticket“) – aus bisherigen Tickets dieser Firma ermittelt.
  // Nur eindeutige Treffer zählen; sonst entscheidet billableDefault über den Namen.
  const kndCache = {};
  async function kndForAddress(c) {
    if (c.KND_ID) return c.KND_ID;
    if (c.ROW_ID in kndCache) return kndCache[c.ROW_ID];
    let id = null;
    try {
      const name = String(c.MATCH || '').trim();
      if (name.length >= 3) {
        const rows = await load('MDERRDA', { take: 200, filter: [['MATCHCODE', 'contains', name.slice(0, 40)]], sort: [{ selector: 'SID', desc: true }], sqlColumns: ['SID', 'KND_ID', 'MATCHCODE'] });
        const ids = [...new Set(rows.filter(r => r.KND_ID && normName(r.MATCHCODE) === normName(name)).map(r => String(r.KND_ID)))];
        if (ids.length === 1) id = ids[0];
      }
    } catch { /* Fallback: Name */ }
    return (kndCache[c.ROW_ID] = id);
  }

  // Bearbeiter-Liste (MultiData-Benutzer) – einmal pro Sitzung geladen
  let USERS = null;
  async function loadUsers() {
    if (USERS) return USERS;
    const r = await load('MDERRDA', { lookupColumnName: 'BEARBEITER_USER_ID', skip: 0, take: 2000, options: [] });
    USERS = r.filter(u => u.USER_ID).map(u => ({ id: u.USER_ID, name: `${u.VORNAME || ''} ${u.ZUNAME || ''}`.trim(), abt: u.ABT_ID || '', email: u.EMAIL || '' }))
      .sort((a, b) => a.name.localeCompare(b.name, 'de'));
    return USERS;
  }
  const userLabel = id => { const u = (USERS || []).find(x => x.id === id); return u ? `${u.name} (${u.id})` : (id || '–'); };

  // Ticket in Zielticket zusammenführen (wie „Zusätze → Tickets zusammenführen…“)
  async function mergeTickets(src, target) {
    if (Number(src.SID) === Number(target.SID)) throw new Error('Quelle und Ziel sind dasselbe Ticket.');
    await exec('MDERRDA', 'Cmd.Ticket.Merge', 'Md_Crm.TicketMerge', { destSID: Number(target.SID), retType: 'JSON', cmdName: 'Cmd.Ticket.Merge' }, { rowIds: [src.ROW_ID] });
    log('Tickets zusammengeführt', `#${src.SID} „${src.TITEL}“ → #${target.SID} „${target.TITEL}“`);
    ticketsChanged(src.SID, target.SID);
  }
  // Child-Ticket anlegen (wie „Zusätze → Child-Ticket erstellen“) – liefert die neue Ticket-ID
  async function createChildTicket(parent) {
    const r = await exec('MDERRDA', 'Cmd.Ticket.CreateSlave', 'Md_Crm.CreateSlaveTicket', { sid: Number(parent.SID) });
    if (!r || !r.ret) throw new Error('Child-Ticket wurde evtl. nicht angelegt.');
    log('Child-Ticket angelegt', `#${r.ret} unter #${parent.SID} „${parent.TITEL}“`);
    ticketsChanged(parent.SID, r.ret);
    return r.ret;
  }

  let QUEUES = null;
  async function loadQueues() {
    if (QUEUES) return QUEUES;
    QUEUES = (await load('MDERRDA', { lookupColumnName: 'ERRQ_ID', skip: 0, take: 200, options: [] })).map(q => ({ id: q.ERRQ_ID, name: q.MEMO_FK_DISPLAY || String(q.ERRQ_ID) }));
    return QUEUES;
  }
  const loadChildren = sid => load('MDERRDA', { take: 200, filter: [['MDERRDA_SID', '=', Number(sid)]], sort: [{ selector: 'SID', desc: true }], sqlColumns: TICKET_COLS });

  async function updateTicket(t, changes) {
    const saved = await updateRow('MDERRDA', t, changes, TICKET_DETAIL_COLS);
    if (saved) { log('Ticket geändert', `#${t.SID}: ` + Object.entries(changes).map(([k, v]) => `${k}=${k === 'STATUS' ? statusInfo(v)[1] : v}`).join(', ')); ticketsChanged(t.SID); }
    return saved;
  }
  async function changeFirm(t, adr) {
    await exec('MDERRDA', 'Cmd.Ticket.ChangeAdr', 'Md_WebUser.ChangeTicketAdr', { SID: Number(t.SID), rowId: adr.ROW_ID });
    log('Firma geändert', `#${t.SID}: ${t.MATCHCODE} → ${adr.MATCH}`);
    ticketsChanged(t.SID);
  }
  async function quickClose(t) {
    await exec('MDERRDA', 'IND_TICKET.QuickClose', 'IND_TICKET.QuickClose', { cmdName: 'IND_TICKET.QuickClose' }, { rowIds: [t.ROW_ID] });
    log('QuickClose', `#${t.SID} ${t.TITEL}`);
    ticketsChanged(t.SID);
  }
  async function createTicket({ adrRowId, title, html, suppressNotify }) {
    const d = await api('data/InitNewRow/DLG_TICKET_CREATE', { tableName: 'DLG_TICKET_CREATE', values: { MEMO_BATCH: 'true' }, copy: false });
    const r = await exec('DLG_TICKET_CREATE', 'Cmd.CreateTicket', 'Md_WebUser.CreateTicket', {
      adrRowId, title, theNotiz: html, priority: d.PRIORITY ?? 3, status: d.STATUS ?? 'N', errId: d.ERR_ID ?? 'REQUEST',
      errqId: d.ERRQ_ID ?? 1, userId: d.BEARBEITER_USER_ID, abtId: null, termin: null, vDate: null,
      supNotify: !!suppressNotify, noSla: false, ticketIdExt: null, prjId: null, Mt_Document: null,
    });
    const rowId = JSON.parse(r.selectInfos || '[]')[0]?.rowIds?.[0];
    if (!rowId) throw new Error('Ticket angelegt, aber keine ID erhalten.');
    const t = await ticketByRowId(rowId);
    log('Ticket angelegt', `#${t.SID} ${title}`);
    ticketsChanged(t.SID);
    return t;
  }

  // --- Kommentare ---
  const loadComments = sid => load('MDERRDA_MDERRDAC', { take: 1000, filter: [['MDERRDA_SID', '=', Number(sid)]], sort: [{ selector: 'LST_DATUM', desc: true }], sqlColumns: COMMENT_COLS });
  const commentBySid = async sid => (await load('MDERRDA_MDERRDAC', { filter: [['SID', '=', Number(sid)]], sqlColumns: COMMENT_COLS }))[0];

  const loadDocs = sid => load('MDERRDA_MDDOC', { take: 500, filter: [['MDERRDA_SID', '=', Number(sid)]], sort: [{ selector: 'DATUM', desc: true }], sqlColumns: DOC_COLS });

  // Empfänger-Vorschläge (Kontakte der Firma) – wie das Feld „An“ in MultiData
  async function emailSuggestions(kndId) {
    try {
      const r = await load('MDERRDA', { lookupColumnName: 'EMAIL_TO', skip: 0, take: 300,
        options: [{ parameters: [{ name: 'pKndId', value: null, dbType: 28 }, { name: 'kndId', value: String(kndId), dbType: 28 }] }] });
      return r.filter(x => x.EMAIL).map(x => ({ email: x.EMAIL, label: (x.BEZEICHNUNG || x.EMAIL).trim() }));
    } catch { return []; }
  }
  const splitMails = s => String(s || '').split(/[;,\s]+/).map(x => x.trim()).filter(Boolean);
  const validMail = m => /^[^@\s;,<>]+@[^@\s;,<>]+\.[a-z]{2,}$/i.test(m);

  // Neuer Kommentar inkl. Zeit. minutes = 0 → reiner Kommentar ohne Zeit.
  // mail = { to[], cc[], bcc[] } → wird als Mail versendet (emailKz). files = File[] → Anhänge.
  async function addComment(t, { date, minutes, html, kontaktart, billable, abrHours, mail = null, files = [] }) {
    if (mail) {
      const fresh = await load('MDERRDA', { filter: [['SID', '=', Number(t.SID)]], sqlColumns: ['ROW_ID', 'SID', 'SUPPRESS_NOTIFY'] }).then(r => r[0]);
      if (!fresh) throw new Error(`Ticket #${t.SID} nicht gefunden – Mail nicht gesendet.`);
      t.SUPPRESS_NOTIFY = fresh.SUPPRESS_NOTIFY;
      if (fresh.SUPPRESS_NOTIFY == 1) {
        const e = new Error('Am Ticket ist (inzwischen) „Keine Benachrichtigungen“ gesetzt – Mail wurde NICHT gesendet. Haken entfernen und erneut senden.');
        e.code = 'SUPPRESS'; throw e;
      }
      const all = [...mail.to, ...mail.cc, ...mail.bcc];
      if (!mail.to.length) throw new Error('Mindestens ein Empfänger in „An“ nötig.');
      const bad = all.filter(m => !validMail(m));
      if (bad.length) throw new Error('Ungültige Adresse(n): ' + bad.join(', '));
    }
    const docIds = await uploadFiles(files);
    let start;
    if (date === todayStr()) { const n = new Date(); start = Math.max(0, Math.floor((n.getHours() * 60 + n.getMinutes()) / CFG.roundTo) * CFG.roundTo - minutes); }
    else start = CFG.defaultStartIfPast;
    const end = Math.min(start + minutes, 23 * 60 + 59);
    await api('data/InitNewRow/DLG_MDERRDAC', { tableName: 'DLG_MDERRDAC', values: { ROW_ID: crypto.randomUUID() }, copy: false });
    const r = await exec('DLG_MDERRDAC', 'CMD.CreateComment', 'Md_WebUser.CreateTicketComment', {
      sid: Number(t.SID), title: t.TITEL, lstId: kontaktart || (mail ? 'MAIL' : CFG.defaultKontaktart), status: 'Z',
      emailKz: !!mail, // nur bei ausdrücklichem Mailversand
      notiz: html, Mt_Document: docIds.length ? docIds : null, startTime: apiDT(date, start), endTime: apiDT(date, end), timeText: null,
      isComment: false,
      mailTo: mail ? mail.to.join(';') : null, mailCc: mail && mail.cc.length ? mail.cc.join(';') : null, mailBcc: mail && mail.bcc.length ? mail.bcc.join(';') : null,
      CSID: null, parentCSID: null, createThread: false,
    });
    const cid = r && r.ret;
    if (!cid) throw new Error('Kommentar wurde evtl. nicht angelegt.');
    await exec('DLG_MDERRDAC', 'CMD.UpdateInAbrStd', 'IND_WEB_RESETABRSTD', { commentSid: cid, setAbrStd: !!billable });
    let c = null;
    if (billable && minutes > 0) {
      // Beim Anlegen übernimmt MultiData „abrechenbar“ i. d. R. direkt; abweichende Stunden bzw. Fallback über „Manuell“
      c = await commentBySid(cid);
      const want = abrHours ?? c.LFZ_STD;
      if (Number(c.INABR_STD) !== Number(want)) await updateRow('MDERRDA_MDERRDAC', c, { INABR_STD: want, LFZ_STD_MAN: 1 }, COMMENT_COLS);
    }
    if (mail) {
      c = await commentBySid(cid);
      log('Mail gesendet', `#${t.SID}: An ${mail.to.join(', ')}${mail.cc.length ? ' · Cc ' + mail.cc.join(', ') : ''}${mail.bcc.length ? ' · Bcc ' + mail.bcc.join(', ') : ''}` +
        `${files.length ? ' · ' + files.length + ' Anhang/Anhänge' : ''} · ${c?.EMAIL_DATUM ? 'versendet ' + fmtDateTime(c.EMAIL_DATUM) : 'Versandzeit noch offen'}`);
    }
    log('Kommentar angelegt', `#${t.SID}: ${fmtDate(date)} ${minutes} min${billable ? ', abrechenbar ' + fmtH(abrHours ?? minutes / 60) : ''}${files.length ? ', ' + files.length + ' Anhang/Anhänge' : ''}`);
    ticketsChanged(t.SID);
    return { cid, emailDate: c?.EMAIL_DATUM || null };
  }

  // Kommentar bearbeiten. changes: { html?, date?, from?, to?, manual?, std?, abr?, kontaktart? }
  async function editComment(c, ch) {
    const v = {};
    if (ch.html != null) v.NOTIZ = ch.html;
    if (ch.kontaktart) v.LST_ID = ch.kontaktart;
    if (ch.date && ch.from != null && ch.to != null) {
      v.LFZ_VON = dbDT(ch.date, ch.from); v.LFZ_BIS = dbDT(ch.date, ch.to); v.LST_DATUM = dbDT(ch.date, ch.from);
    }
    if (ch.manual) { v.LFZ_STD_MAN = 1; if (ch.std != null) v.LFZ_STD = ch.std; if (ch.abr != null) v.INABR_STD = ch.abr; }
    else if (ch.manual === false) { v.LFZ_STD_MAN = 0; }
    const saved = await updateRow('MDERRDA_MDERRDAC', c, v, COMMENT_COLS);
    if (!saved) return null;
    log('Kommentar geändert', `#${c.MDERRDA_SID} / K${c.SID}: ` + Object.keys(v).map(k => k === 'NOTIZ' ? 'Text' : `${k}=${v[k]}`).join(', '));
    ticketsChanged(c.MDERRDA_SID);
    return saved;
  }
  async function deleteComment(c) {
    await deleteRows('MDERRDA_MDERRDAC', [c.ROW_ID]);
    log('Kommentar gelöscht', `#${c.MDERRDA_SID} / K${c.SID} (${fmtDateTime(c.LST_DATUM)}, ${fmtH(c.LFZ_STD)} h): ${htmlToText(c.NOTIZ).slice(0, 120)}`);
    ticketsChanged(c.MDERRDA_SID);
  }

  let KONTAKTARTEN = null;
  async function kontaktarten() {
    if (KONTAKTARTEN) return KONTAKTARTEN;
    try {
      const r = await load('MDERRDA_MDERRDAC', { lookupColumnName: 'LST_ID', skip: 0, take: 100, options: [] });
      KONTAKTARTEN = r.map(x => [x.LST_ID, x.BEZEICHNUNG1 || x.LST_ID]);
    } catch { KONTAKTARTEN = [['SUPPORT', 'Support'], ['VORORT', 'vor Ort'], ['TELEFON', 'Telefonisch'], ['WARTUNG', 'Wartung'], ['INFO', 'interne Notiz']]; }
    return KONTAKTARTEN;
  }

  // --- Einstellungen (Signatur, eigene Textbausteine) – nur in diesem Browser ---
  const settings = () => ({ signature: '', autoSignature: true, blocks: [], showMdBlocks: true, ...store.get(CFG.store.settings, {}) });
  const saveSettings = s => store.set(CFG.store.settings, s);

  // Textbausteine aus MultiData (nur HTML-Bausteine; RTF-Bausteine sind für Berichte gedacht)
  let MD_BLOCKS = null;
  async function mdTextBlocks() {
    if (MD_BLOCKS) return MD_BLOCKS;
    try {
      const l = await api('rex/TextBlocks/SYTXTBR_HTML', undefined, 'GET');
      MD_BLOCKS = (Array.isArray(l) ? l : []).filter(b => b.text && !/^\{\\rtf/.test(b.text) && /<\w/.test(b.text))
        .map(b => ({ id: b.id, name: (b.title && b.title !== 'x' ? b.title : b.id), html: b.text, group: b.grouping || '' }));
    } catch { MD_BLOCKS = []; }
    return MD_BLOCKS;
  }

  // --- Entwürfe (automatisch gespeichert, pro Ticket bzw. für „Neues Ticket“) ---
  const drafts = {
    all: () => store.get(CFG.store.drafts, {}),
    get: key => drafts.all()[key] || null,
    set(key, d) { const a = drafts.all(); a[key] = { ...d, key, tab: TAB_ID, savedAt: new Date().toISOString() }; store.set(CFG.store.drafts, a); updateBadges(); },
    del(key) { const a = drafts.all(); delete a[key]; store.set(CFG.store.drafts, a); updateBadges(); },
  };

  // --- Inbox: Benachrichtigungen aus meinen Tickets ---
  // Typen: assign (neu zugewiesen), reply (Kundenantwort), comment (Kommentar von Kolleg:in), reopen (Wieder neu), due (Termin fällig), md (MultiData)
  const INBOX_TYPES = { assign: '🆕 Neu zugewiesen', reply: '✉ Kundenantwort', comment: '💬 Kommentar Kolleg:in', reopen: '🔁 Wieder neu', due: '⏰ Termin fällig', md: '🔔 MultiData' };
  const inboxState = () => ({ items: {}, known: null, last: null, ...store.get(CFG.store.inbox, {}) });
  let inboxBusy = null;
  function refreshInbox() { return inboxBusy ||= doRefreshInbox().finally(() => { inboxBusy = null; }); }
  async function doRefreshInbox() {
    const meId = await me();
    const st = inboxState();
    const firstRun = !st.known;
    const now = new Date().toISOString(), today = todayStr();
    const since = addDays(today, -CFG.inboxDays) + 'T00:00:00';
    const add = (key, item, readOnFirstRun = true) => {
      if (st.items[key]) return;
      st.items[key] = { ...item, key, found: now, read: firstRun && readOnFirstRun && (item.date || '') < addDays(today, -1) };
    };
    const mine = await load('MDERRDA', {
      take: 3000, filter: andAll([['BEARBEITER_USER_ID', '=', meId], [['STATUS', '<>', 'Z'], 'or', ['DATE_MODIFIED', '>=', since]]]),
      sqlColumns: [...TICKET_COLS, 'TERMIN'],
    });
    const bySid = Object.fromEntries(mine.map(t => [t.SID, t]));
    const open = mine.filter(t => t.STATUS !== 'Z');
    const base = t => ({ sid: t.SID, title: t.TITEL, cust: t.MATCHCODE, kndId: t.KND_ID });
    // 1) Neu zugewiesen: offene Tickets, die beim letzten Abruf noch nicht meine waren
    const known = new Set(st.known || []);
    for (const t of open) {
      if (firstRun ? (t.ANLAGE_DATUM || '') >= today : !known.has(t.SID))
        add('assign:' + t.SID + ':' + (t.DATE_MODIFIED || ''), { type: 'assign', date: t.DATE_MODIFIED || t.ANLAGE_DATUM, ...base(t), text: `Status: ${statusInfo(t.STATUS)[1]}` }, false);
    }
    st.known = open.map(t => t.SID);
    // 2) Wieder neu  3) Termin fällig
    for (const t of open) {
      if (t.STATUS === 'NR') add('reopen:' + t.SID + ':' + (t.DATE_MODIFIED || ''), { type: 'reopen', date: t.DATE_MODIFIED, ...base(t), text: 'Ticket wurde wieder geöffnet.' });
      if (t.TERMIN && t.TERMIN.slice(0, 10) <= today) add('due:' + t.SID + ':' + t.TERMIN.slice(0, 10), { type: 'due', date: t.TERMIN, ...base(t), text: `Termin ${fmtDateTime(t.TERMIN)}` }, false);
    }
    // 4) Kundenantworten und Kommentare anderer in meinen Tickets
    const sids = Object.keys(bySid).map(Number);
    for (let i = 0; i < sids.length; i += 60) {
      const cs = await load('MDERRDA_MDERRDAC', {
        take: 2000, filter: andAll([orFilter('MDERRDA_SID', sids.slice(i, i + 60)), ['LST_DATUM', '>=', since], ['BEARBEITER_USER_ID', '<>', meId]]),
        sqlColumns: ['SID', 'MDERRDA_SID', 'BEARBEITER_USER_ID', 'LST_DATUM', 'LST_ID', 'MEMO_IA', 'EMAIL_TO', 'NOTIZ'],
      });
      for (const c of cs) {
        const t = bySid[c.MDERRDA_SID]; if (!t) continue;
        const snip = htmlToText(c.NOTIZ).slice(0, 220);
        if (c.MEMO_IA === '1' || c.MEMO_IA === 1) add('reply:' + c.SID, { type: 'reply', date: c.LST_DATUM, ...base(t), from: c.EMAIL_TO, text: snip });
        else if (c.BEARBEITER_USER_ID && c.BEARBEITER_USER_ID !== 'SUPPORT') add('comment:' + c.SID, { type: 'comment', date: c.LST_DATUM, ...base(t), from: c.BEARBEITER_USER_ID, text: snip });
      }
    }
    // 5) Benachrichtigungen von MultiData selbst
    try {
      const n = await api('notification/load', {});
      for (const x of (n.notifications || [])) {
        const id = x.id ?? x.ID ?? x.SID ?? JSON.stringify(x).slice(0, 80);
        add('md:' + id, { type: 'md', date: x.date || x.created || x.DATE_CREATED || now, sid: x.sid || x.SID || null, title: x.title || x.subject || x.TITEL || 'MultiData', cust: '', text: x.text || x.message || x.body || '' }, false);
      }
    } catch { /* optional */ }
    // Alte Einträge (> 30 Tage) aufräumen
    const cut = addDays(today, -30);
    for (const [k, v] of Object.entries(st.items)) if ((v.found || '') < cut) delete st.items[k];
    st.last = now;
    store.set(CFG.store.inbox, st);
    return st;
  }
  function setInboxRead(keys, read = true) {
    const st = inboxState();
    for (const k of keys) if (st.items[k]) st.items[k].read = read;
    store.set(CFG.store.inbox, st);
  }

  // ===========================================================================
  // 6. Oberfläche
  // ===========================================================================
  const css = `
  #mdqt-fab{position:fixed;right:18px;bottom:18px;z-index:99998;background:#00897b;color:#fff;border:none;border-radius:24px;
    padding:10px 16px;font:600 13px Roboto,Arial,sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);cursor:pointer}
  #mdqt-panel{position:fixed;top:56px;right:12px;bottom:66px;z-index:99999;width:min(1000px,calc(100vw - 24px));display:none;
    flex-direction:column;background:var(--mq-bg);color:var(--mq-fg);border-radius:8px;box-shadow:0 6px 28px rgba(0,0,0,.35);
    font:13px/1.35 Roboto,Arial,sans-serif;color-scheme:var(--mq-scheme);overflow:hidden}
  #mdqt-panel.open{display:flex}
  #mdqt-panel *{box-sizing:border-box}
  #mdqt-panel .hdr{display:flex;align-items:center;background:#00897b;color:#fff}
  #mdqt-panel .hdr button.tab{background:none;border:none;color:#d7f0ed;padding:11px 14px;cursor:pointer;font:600 13px inherit;font-family:inherit;border-bottom:3px solid transparent}
  #mdqt-panel .hdr button.tab.active{color:#fff;border-bottom-color:#fff}
  #mdqt-panel .hdr .sp{flex:1}
  #mdqt-panel .hdr .x{background:none;border:none;color:#fff;font-size:20px;padding:6px 14px;cursor:pointer}
  #mdqt-panel .body{flex:1;overflow:auto;padding:12px 14px}
  #mdqt-panel section{display:none}#mdqt-panel section.active{display:block}
  #mdqt-panel label.l{display:block;margin:8px 0 3px;color:var(--mq-muted2);font-size:12px}
  #mdqt-panel input[type=text],#mdqt-panel input[type=date],#mdqt-panel input[type=time],#mdqt-panel select,#mdqt-panel textarea,#mdqt-panel .rte{
    width:100%;border:1px solid var(--mq-border);border-radius:4px;padding:5px 7px;font:13px inherit;font-family:inherit;background:var(--mq-input) !important;color:var(--mq-fg) !important}
  #mdqt-panel textarea{min-height:64px;resize:vertical}
  #mdqt-panel .rte{min-height:70px;max-height:260px;overflow:auto}
  #mdqt-panel .rte:focus{outline:2px solid #80cbc4}
  #mdqt-panel input[type=checkbox]{appearance:auto !important;-webkit-appearance:checkbox !important;width:15px;height:15px;margin:0;accent-color:#00897b;flex:none}
  #mdqt-panel .row{display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end}
  #mdqt-panel .row>.c{flex:1;min-width:120px}
  #mdqt-panel .row>.c.s{flex:0 0 110px;min-width:0}
  #mdqt-panel .chips{display:flex;flex-wrap:wrap;gap:4px;margin-top:4px}
  #mdqt-panel .chip.more{border-style:dashed;font-weight:600}
  #mdqt-panel .chip{border:1px solid #00897b;color:var(--mq-accent-fg);background:var(--mq-bg);border-radius:12px;padding:2px 9px;cursor:pointer;font-size:12px}
  #mdqt-panel .chip.sel{background:#00897b;color:#fff}
  #mdqt-panel .chip .n{opacity:.7;margin-left:4px;font-size:11px}
  #mdqt-panel .mp{position:relative;display:flex;flex-wrap:wrap;gap:4px;align-items:center;border:1px solid var(--mq-border);border-radius:4px;padding:3px 5px;background:var(--mq-bg);min-height:32px;cursor:text}
  #mdqt-panel .mp .pill{background:#00897b;color:#fff;border-radius:12px;padding:1px 4px 1px 9px;font-size:12px;display:inline-flex;align-items:center;gap:3px;max-width:100%}
  #mdqt-panel .mp .pill span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  #mdqt-panel .mp .pill b{cursor:pointer;padding:0 4px;font-weight:700}
  #mdqt-panel .mp input[type=text]{border:none !important;flex:1;min-width:120px;padding:3px !important;outline:none}
  #mdqt-panel .mp .results{position:absolute;left:0;right:0;top:100%;z-index:6}
  #mdqt-panel .chk{display:inline-flex;gap:6px;align-items:center;margin:6px 12px 0 0;cursor:pointer}
  #mdqt-panel .btn{padding:7px 12px;border:none;border-radius:4px;background:#00897b;color:#fff;font:600 13px inherit;font-family:inherit;cursor:pointer}
  #mdqt-panel .btn.sec{background:var(--mq-accent-soft);color:var(--mq-accent-fg)}
  #mdqt-panel .btn.warn{background:#c62828}
  #mdqt-panel .btn.sm{padding:3px 8px;font-size:12px}
  #mdqt-panel .btn:disabled{opacity:.5;cursor:default}
  #mdqt-panel .msg{margin:8px 0;font-size:12px;white-space:pre-wrap}
  #mdqt-panel .msg.err{color:#c62828}#mdqt-panel .msg.ok{color:#2e7d32}
  #mdqt-panel .toasts{position:absolute;left:50%;bottom:18px;transform:translateX(-50%);z-index:60;display:flex;flex-direction:column;align-items:center;gap:6px;
    width:max-content;max-width:min(560px,calc(100% - 32px));pointer-events:none}
  #mdqt-panel .toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:9px 12px;border-radius:6px;font-size:13px;line-height:1.35;
    white-space:pre-wrap;color:#fff;background:#c62828;box-shadow:0 4px 16px rgba(0,0,0,.35);cursor:pointer;animation:mqToastIn .15s ease-out}
  #mdqt-panel .toast.warn{background:#e65100}
  #mdqt-panel .toast b{font-size:16px;line-height:1;opacity:.85}
  @keyframes mqToastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
  #mdqt-panel .hint{color:var(--mq-muted);font-size:12px}
  #mdqt-panel table{width:100%;border-collapse:collapse;margin-top:8px}
  #mdqt-panel th{text-align:left;font-size:12px;color:var(--mq-muted2);border-bottom:2px solid var(--mq-border);padding:4px 6px;cursor:pointer;white-space:nowrap;position:sticky;top:-12px;background:var(--mq-bg)}
  #mdqt-panel td{border-bottom:1px solid var(--mq-line);padding:4px 6px;vertical-align:top}
  #mdqt-panel tr.r:hover td{background:var(--mq-accent-soft);cursor:pointer}
  #mdqt-panel td.n{text-align:right;white-space:nowrap}
  #mdqt-panel .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px;vertical-align:middle}
  #mdqt-panel .flag{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px;vertical-align:middle}
  #mdqt-panel tr.ghead td{background:var(--mq-soft);font-weight:700;padding-top:8px}
  #mdqt-panel tr.gsum td{color:var(--mq-muted);border-bottom:2px solid var(--mq-border)}
  #mdqt-panel tr.tsum td{border-top:2px solid var(--mq-border)}
  #mdqt-panel .card{border:1px solid var(--mq-border);border-radius:6px;padding:8px 10px;margin-top:8px}
  #mdqt-panel .card.edit{border-color:#00897b;background:var(--mq-edit)}
  #mdqt-panel .card .meta{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--mq-muted2);align-items:center}
  #mdqt-panel .card .meta b{color:var(--mq-fg)}
  #mdqt-panel .card .txt{margin-top:4px}
  #mdqt-panel .card .txt p{margin:2px 0}#mdqt-panel .card .txt ul{margin:2px 0 2px 18px;padding:0}
  #mdqt-panel .head{background:var(--mq-soft);border-radius:6px;padding:10px}
  #mdqt-panel .head h3{margin:0 0 6px;font-size:15px}
  #mdqt-panel .sums{font-size:12px;color:var(--mq-muted2)}
  #mdqt-panel .results{border:1px solid var(--mq-border);border-radius:4px;max-height:180px;overflow:auto;margin-top:2px;background:var(--mq-bg)}
  #mdqt-panel .results div{padding:5px 8px;cursor:pointer;border-bottom:1px solid var(--mq-line)}
  #mdqt-panel .results div:hover{background:var(--mq-accent-soft)}
  #mdqt-panel .tag{display:inline-block;background:var(--mq-tag);border-radius:3px;padding:0 5px;font-size:11px;color:var(--mq-muted2)}
  #mdqt-panel .tag.abr{background:var(--mq-bg)3e0;color:#e65100}
  #mdqt-panel details summary{cursor:pointer;color:var(--mq-accent-fg);font-weight:600;margin-top:6px}
  #mdqt-panel a.ic{text-decoration:none;display:inline-block}
  #mdqt-panel tr.r td.l{padding:0}
  #mdqt-panel a.rl{display:block;padding:4px 6px;color:inherit;text-decoration:none;min-height:1em}
  #mdqt-panel a.ibx{color:inherit;text-decoration:none}
  #mdqt-panel a.btn{text-decoration:none;display:inline-block}
  #mdqt-panel .hdr .ic{background:none;border:none;color:#fff;font-size:16px;padding:6px 8px;cursor:pointer;opacity:.9}
  #mdqt-panel .rz{position:absolute;z-index:2}
  #mdqt-panel .rz-l{left:0;top:0;bottom:0;width:6px;cursor:ew-resize}
  #mdqt-panel .rz-b{left:0;right:0;bottom:0;height:6px;cursor:ns-resize}
  #mdqt-panel .rz-bl{left:0;bottom:0;width:14px;height:14px;cursor:nesw-resize}
  #mdqt-panel .rz-r{right:0;top:0;bottom:0;width:6px;cursor:ew-resize}
  #mdqt-panel .rz-br{right:0;bottom:0;width:14px;height:14px;cursor:nwse-resize}
  #mdqt-panel.full,#mdqt-panel.app{top:0 !important;right:0 !important;left:0 !important;bottom:0 !important;width:auto !important;height:auto !important;border-radius:0}
  #mdqt-panel.full .rz,#mdqt-panel.app .rz{display:none}
  #mdqt-panel.app{z-index:2147483000}
  #mdqt-panel.app .hdr [data-ref=close],#mdqt-panel.app .hdr [data-ref=full]{display:none}
  #mdqt-panel .tb{display:flex;gap:2px;margin-bottom:2px}
  #mdqt-panel .tb{flex-wrap:wrap}
  #mdqt-panel .tb button{border:1px solid var(--mq-border);background:var(--mq-btn);border-radius:3px;padding:1px 7px;cursor:pointer;font-size:12px;color:var(--mq-fg)}
  #mdqt-panel .tb select,#mdqt-panel .tb input[type=number]{width:auto !important;padding:1px 4px !important;font-size:12px !important}
  #mdqt-panel .rte table td,#mdqt-panel .rte table th{border:1px solid #999}
  #mdqt-panel .rte img{max-width:100%}
  #mdqt-panel .txt img{max-width:100%}
  #mdqt-panel .drop{border:2px dashed #b2dfdb;border-radius:6px;padding:8px;text-align:center;color:#00796b;margin-top:6px;cursor:pointer}
  #mdqt-panel .drop.over{background:var(--mq-accent-soft)}
  #mdqt-panel .files{display:flex;flex-wrap:wrap;gap:4px;margin-top:4px}
  #mdqt-panel .file{background:var(--mq-tag);border-radius:3px;padding:1px 6px;font-size:12px}
  #mdqt-panel .file b{cursor:pointer;color:#c62828;margin-left:4px}
  #mdqt-panel .mailbox{border-left:3px solid #1e88e5;padding:4px 0 4px 10px;margin-top:8px;background:var(--mq-mail-bg)}
  #mdqt-panel .warnbox{background:var(--mq-bg)3e0;border:1px solid #ffb74d;border-radius:4px;padding:6px 8px;margin-top:6px;font-size:12px}
  #mdqt-panel .doc{display:flex;gap:8px;align-items:center;padding:4px 0;border-bottom:1px solid var(--mq-line)}
  #mdqt-panel .doc .nm{flex:1;word-break:break-all}
  #mdqt-panel .mailinfo{font-size:12px;color:var(--mq-link2);margin-top:3px}
  #mdqt-panel .kpis{display:flex;flex-wrap:wrap;gap:8px;margin:8px 0}
  #mdqt-panel .kpi{flex:1 1 140px;background:var(--mq-soft);border-radius:6px;padding:8px 10px}
  #mdqt-panel .kpi .kl{font-size:11px;color:var(--mq-muted)}#mdqt-panel .kpi .kv{font-size:17px;font-weight:700}
  #mdqt-panel .kpi .ks{font-size:11px;color:var(--mq-muted);margin-top:2px}
  #mdqt-panel .ztbl tr.zday td{border-top:2px solid var(--mq-border);padding-top:7px;background:var(--mq-soft)}
  #mdqt-panel .ztbl tr.z-red td{background:#fde2e2}#mdqt-panel .ztbl tr.z-yellow td{background:#fff4cc}#mdqt-panel .ztbl tr.z-green td{background:#dff3e3}
  #mdqt-panel .ztbl tr.z-free td{background:var(--mq-tag)}
  #mdqt-panel.dark .ztbl tr.z-red td{background:#4a2326}#mdqt-panel.dark .ztbl tr.z-yellow td{background:#4a4120}#mdqt-panel.dark .ztbl tr.z-green td{background:#203d29}
  #mdqt-panel .ztbl tr.z-we td{opacity:.85}
  #mdqt-panel .ztbl tr.zgrp td{font-size:12px;color:var(--mq-muted2);font-weight:600}
  #mdqt-panel .ztbl tr.zi td{font-size:12px;border-bottom:none}
  #mdqt-panel td.neg{color:#c62828;font-weight:700}#mdqt-panel td.pos{color:#ef6c00;font-weight:700}
  #mdqt-panel.dark td.neg{color:#ef9a9a}#mdqt-panel.dark td.pos{color:#ffcc80}
  #mdqt-panel .badge{display:inline-block;min-width:16px;padding:0 4px;border-radius:8px;background:#ffb300;color:var(--mq-fg);font-size:11px;line-height:16px;text-align:center}
  #mdqt-panel .badge:empty{display:none}
  #mdqt-fab .badge{display:inline-block;margin-left:6px;min-width:18px;padding:0 5px;border-radius:9px;background:#ffb300;color:var(--mq-fg);font-size:11px;line-height:18px}
  #mdqt-fab .badge:empty{display:none}
  #mdqt-panel .ibx{display:flex;gap:10px;padding:7px 6px;border-bottom:1px solid var(--mq-line);cursor:pointer;align-items:flex-start}
  #mdqt-panel .ibx:hover{background:var(--mq-hover)}
  #mdqt-panel .ibx.unread{background:var(--mq-bg)df4}
  #mdqt-panel .ibx.unread .tt{font-weight:700}
  #mdqt-panel .ibx .ty{flex:0 0 150px;font-size:12px;color:var(--mq-muted2)}
  #mdqt-panel .ibx .bd{flex:1;min-width:0}
  #mdqt-panel .ibx .sn{color:var(--mq-muted);font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  #mdqt-panel .ibx .dt{flex:0 0 auto;font-size:12px;color:var(--mq-muted);white-space:nowrap}
  #mdqt-panel .card .meta[data-ref=cHead]{cursor:pointer}
  #mdqt-panel .card .tog::before{content:'▾';color:#00897b;font-size:11px}
  #mdqt-panel .card.collapsed .tog::before{content:'▸'}
  #mdqt-panel .card .prev{display:none;color:var(--mq-muted);font-style:italic;flex:1 1 0;min-width:0;max-width:520px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding-right:4px}
  #mdqt-panel .card.collapsed .prev{display:block}
  #mdqt-panel .card.collapsed .txt,#mdqt-panel .card.collapsed .mailinfo{display:none !important}

  #mdqt-panel{--mq-bg:#fff;--mq-fg:#222;--mq-input:#fff;--mq-soft:#f5f5f5;--mq-border:#d0d0d0;--mq-line:#eee;--mq-muted:#777;--mq-muted2:#555;
    --mq-accent-soft:#e0f2f1;--mq-accent-fg:#00695c;--mq-edit:#f4fbfa;--mq-btn:#fafafa;--mq-warn-bg:#fff3e0;--mq-mail-bg:#f5f9ff;--mq-unread:#fffdf4;
    --mq-hover:#f1f8f7;--mq-tag:#eceff1;--mq-link2:#1565c0;--mq-link:#00796b;--mq-scheme:light}
  #mdqt-panel.dark{--mq-bg:#1f2124;--mq-fg:#e4e6e8;--mq-input:#2a2d31;--mq-soft:#2a2c30;--mq-border:#4a4e54;--mq-line:#33363b;--mq-muted:#9aa0a6;--mq-muted2:#b6bcc2;
    --mq-accent-soft:#1d3a37;--mq-accent-fg:#80cbc4;--mq-edit:#1c2b2a;--mq-btn:#2d3034;--mq-warn-bg:#3a2f1d;--mq-mail-bg:#1b2533;--mq-unread:#2a2920;
    --mq-hover:#25302f;--mq-tag:#353a40;--mq-link2:#90caf9;--mq-link:#80cbc4;--mq-scheme:dark}
  #mdqt-panel a{color:var(--mq-link)}
  #mdqt-panel table{color:var(--mq-fg);font-size:inherit}
  /* Mail-/Kommentarinhalte im Dark Mode auf hellem „Papier“, damit fremde Textfarben lesbar bleiben */
  #mdqt-panel.dark .rte,#mdqt-panel.dark .card .txt{background:#f4f4f4 !important;color:#222 !important;border-radius:4px}
  #mdqt-panel.dark .card .txt{padding:6px 8px}
  #mdqt-panel.dark .card .txt a,#mdqt-panel.dark .rte a{color:#00695c}
  #mdqt-panel.dark .card .txt .hint{color:#666}
  #mdqt-panel.dark .tb button,#mdqt-panel.dark .tb select{color:var(--mq-fg)}
  #mdqt-panel.dark .chip.sel{color:#fff}
  #mdqt-panel.dark input[type=checkbox]{accent-color:#26a69a}
  #mdqt-panel.dark .flag,#mdqt-panel.dark .dot{box-shadow:0 0 0 1px #0006}
  `;
  const styleEl = document.createElement('style'); styleEl.textContent = css; document.head.appendChild(styleEl);

  const fab = document.createElement('button');
  fab.id = 'mdqt-fab'; fab.innerHTML = '⚡ Werkbank<span class="badge"></span>'; fab.title = 'MultiData Werkbank (Alt+Q)';
  document.body.appendChild(fab);

  const panel = document.createElement('div');
  panel.id = 'mdqt-panel';
  panel.innerHTML = `
    <div class="hdr">
      <button class="tab" data-tab="inbox">📥 Inbox <span class="badge" data-ref="bInbox"></span></button>
      <button class="tab active" data-tab="search">🔎 Suche</button>
      <button class="tab" data-tab="ticket">🎫 Ticket</button>
      <button class="tab" data-tab="times">⏱ Zeiten</button>
      <button class="tab" data-tab="new">＋ Neues Ticket</button>
      <button class="tab" data-tab="drafts">📝 Entwürfe <span class="badge" data-ref="bDrafts"></span></button>
      <button class="tab" data-tab="log">📜 Protokoll</button>
      <button class="tab" data-tab="settings" title="Einstellungen: Signatur, Textbausteine">⚙</button>
      <span class="sp"></span>
      <button class="ic" data-ref="theme" title="Hell / Dunkel umschalten">🌙</button>
      <a class="ic" data-ref="app" href="#" title="Werkbank als eigene Seite in neuem Tab öffnen">↗</a>
      <button class="ic" data-ref="full" title="Vollbild an/aus">⤢</button>
      <button class="x" data-ref="close" title="Schließen (Alt+Q)">×</button>
    </div>
    <div class="rz rz-l" data-rz="l"></div><div class="rz rz-r" data-rz="r"></div><div class="rz rz-b" data-rz="b"></div><div class="rz rz-bl" data-rz="bl"></div><div class="rz rz-br" data-rz="br"></div>
    <div class="body">
      <div class="warnbox" data-ref="teamHint" style="display:none;margin-bottom:8px">Noch keine Team-Konfiguration: Bitte unter ⚙ festlegen, welche Kunden nicht abrechenbar bzw. intern sind – oder eine Datei von Kolleg:innen importieren.
        <button class="btn sm" data-ref="teamHintGo">Zu den Einstellungen</button> <button class="btn sm sec" data-ref="teamHintNo">Später</button></div>
      <!-- SUCHE -->
      <section data-tab="search" class="active">
        <div class="row" style="align-items:center">
          <span class="hint">Vorlagen:</span><div class="chips" data-ref="sTpl" style="flex:1;margin:0"></div>
          <span data-ref="sTplNew" style="display:none"><input type="text" data-ref="sTplName" placeholder="Name der Vorlage" style="width:200px"> <button class="btn sm" data-ref="sTplOk">Speichern</button></span>
          <button class="btn sm sec" data-ref="sTplAdd" title="Aktuelle Filter (Bearbeiter, Kunden, Queues, Status, Zeitraum, Freitext) als Vorlage speichern">＋ Als Vorlage</button>
        </div>
        <div class="row">
          <div class="c"><label class="l">Freitext (Betreff oder Ticket-ID)</label><input type="text" data-ref="sText" placeholder="z. B. Exchange oder Ticket-Nr."></div>
        </div>
        <div class="row" style="align-items:flex-start">
          <div class="c"><label class="l">Bearbeiter <span class="hint">(leer = alle · „ohne“ = nicht zugewiesen)</span></label><div data-ref="sUsers"></div></div>
          <div class="c"><label class="l">Kunden <span class="hint">(suchen oder unten wählen – leer = alle)</span></label><div data-ref="sCusts"></div></div>
          <div class="c"><label class="l">Queues <span class="hint">(leer = alle · „ohne“ = keine Queue)</span></label><div data-ref="sQueues"></div></div>
        </div>
        <label class="l" style="display:flex;align-items:center;gap:8px">Kunden in den aktuellen Filtern <span class="hint" data-ref="sFacetHint" style="flex:1"></span>
          <select data-ref="sFacetSort" style="width:auto;font-size:12px;padding:2px 4px" title="Reihenfolge der Kunden">
            <option value="n-desc">Anzahl Tickets ↓</option><option value="n-asc">Anzahl Tickets ↑</option>
            <option value="name-asc">Name A–Z</option><option value="name-desc">Name Z–A</option></select></label>
        <div class="chips" data-ref="sFacets"><span class="hint">…</span></div>
        <label class="l">Status <span class="hint">(nichts gewählt = alle)</span>
          <button class="btn sm sec" data-ref="sOpen">nur offene</button> <button class="btn sm sec" data-ref="sAll">alle</button></label>
        <div class="chips" data-ref="sStatus"></div>
        <div class="row">
          <div class="c s" style="flex-basis:190px"><label class="l">Zeitraum bezieht sich auf</label>
            <select data-ref="sDateMode"><option value="none">– kein Zeitraum –</option><option value="created">Anlagedatum</option><option value="booked">Gebuchte Zeit</option></select></div>
          <div class="c s" style="flex-basis:150px"><label class="l">von</label><input type="date" data-ref="sFrom"></div>
          <div class="c s" style="flex-basis:150px"><label class="l">bis</label><input type="date" data-ref="sTo"></div>
          <div class="c"><div class="chips" data-ref="sRanges"></div></div>
          <div><button class="btn sec" data-ref="sReset" title="Bearbeiter: Ich · nur offene · alle Kunden · Freitext und Zeitraum leeren">↺ Zurücksetzen</button>
            <button class="btn" data-ref="sGo">Suchen</button></div>
        </div>
        <label class="chk"><input type="checkbox" data-ref="sGroup"> Nach Priorität gruppieren <span class="hint">(Sortierung gilt innerhalb der Gruppen)</span></label>
        <div class="msg" data-ref="sMsg"></div>
        <div data-ref="sResults"></div>
      </section>

      <!-- TICKET -->
      <section data-tab="ticket">
        <div class="row" style="margin-bottom:8px">
          <div class="c s" style="flex-basis:160px"><input type="text" data-ref="tSid" placeholder="Ticket-ID"></div>
          <button class="btn sec" data-ref="tLoad">Laden</button>
          <button class="btn sec" data-ref="tPage">Aktuelles MultiData-Ticket</button>
          <span style="flex:1"></span>
          <button class="btn sec" data-ref="tReload">↻ Neu laden</button>
        </div>
        <div class="msg" data-ref="tMsg"></div>
        <div class="warnbox" data-ref="tBanner" style="display:none"><span></span> <button class="btn sm" data-ref="tBannerReload">Neu laden</button>
          <span class="hint">(ungespeicherte Eingaben im Formular bleiben als Entwurf erhalten)</span></div>
        <div data-ref="tView"><div class="hint">Ticket über die Suche öffnen oder ID eingeben.</div></div>
      </section>

      <!-- NEUES TICKET -->
      <section data-tab="new">
        <label class="l">Firma *</label>
        <input type="text" data-ref="nSearch" placeholder="Suchen (mind. 3 Zeichen) …">
        <div class="results" data-ref="nResults" style="display:none"></div>
        <div class="chips" data-ref="nRecent"></div>
        <div class="msg" data-ref="nFirm"></div>
        <label class="l">Betreff *</label><input type="text" data-ref="nTitle" spellcheck="true" lang="de">
        <label class="l">Beschreibung</label><textarea data-ref="nDesc" spellcheck="true" lang="de"></textarea>
        <label class="chk"><input type="checkbox" data-ref="nNoNotify"> Keine Benachrichtigungen (dann gehen zu diesem Ticket keine Mails raus)</label>
        <label class="chk"><input type="checkbox" data-ref="nWithTime"> Direkt Zeit buchen</label>
        <div data-ref="nTimeBox" style="display:none"></div>
        <label class="chk"><input type="checkbox" data-ref="nClose"> Danach QuickClose</label>
        <div style="margin-top:10px"><button class="btn" data-ref="nGo">Ticket anlegen</button></div>
        <div class="msg" data-ref="nMsg"></div>
      </section>

      <!-- ZEITEN -->
      <section data-tab="times">
        <div class="row" style="align-items:flex-start">
          <div class="c" style="max-width:340px"><label class="l">Bearbeiter</label><div data-ref="zUsers"></div></div>
          <div class="c s" style="flex-basis:150px"><label class="l">von</label><input type="date" data-ref="zFrom"></div>
          <div class="c s" style="flex-basis:150px"><label class="l">bis</label><input type="date" data-ref="zTo"></div>
          <div class="c"><label class="l">&nbsp;</label><div class="chips" data-ref="zRanges"></div></div>
          <div><label class="l">&nbsp;</label><button class="btn" data-ref="zGo">Laden</button></div>
        </div>
        <div class="row" style="align-items:center;margin-top:4px">
          <label class="chk"><input type="checkbox" data-ref="zColor"> Arbeitstage farbig</label>
          <label class="chk">grün ab <input type="text" data-ref="zLow" style="width:44px"> h</label>
          <label class="chk"><input type="checkbox" data-ref="zWeekend"> Wochenenden ohne Buchung ausblenden</label>
          <label class="chk"><input type="checkbox" data-ref="zDetails"> Einzelbuchungen anzeigen</label>
          <span style="flex:1"></span>
          <button class="btn sm sec" data-ref="zCsv">⬇ CSV-Export</button>
        </div>
        <details class="card" data-ref="zPDet"><summary>⚖ Personio-Abgleich <span class="hint" data-ref="zPInfo"></span></summary>
          <div class="drop" data-ref="zDrop">Personio-CSV („Arbeitszeit pro Tag“, Minuten) hierher ziehen oder <u>auswählen</u><input type="file" accept=".csv,text/csv" data-ref="zFile" style="display:none"></div>
          <div class="row" style="margin-top:6px;align-items:center">
            <label class="chk">Fehlzeiten ausgleichen auf Ticket <input type="text" data-ref="zTicket" style="width:110px" placeholder="Ticket-Nr."></label>
            <span class="hint" style="flex:1">Gerundet wird pro Tag ab ${CFG.personioRoundMin} Restminuten auf die nächste Viertelstunde. Der Abgleich gilt nur, wenn nur „Ich“ gewählt ist.</span>
            <button class="btn sm sec" data-ref="zPClear">Personio-Daten entfernen</button>
          </div>
        </details>
        <div class="msg" data-ref="zMsg"></div>
        <div data-ref="zView"></div>
      </section>

      <!-- INBOX -->
      <section data-tab="inbox">
        <div class="row">
          <div class="c"><div class="chips" data-ref="iTypes"></div></div>
          <button class="btn sm sec" data-ref="iRefresh">↻ Aktualisieren</button>
          <button class="btn sm sec" data-ref="iAllRead">✓ Alle als gelesen</button>
        </div>
        <div class="row" style="margin-top:6px">
          <div class="c s" style="flex-basis:220px"><select data-ref="iCust"><option value="">Alle Kunden</option></select></div>
          <div class="c s" style="flex-basis:160px"><select data-ref="iRange"><option value="1">Heute</option><option value="7" selected>Letzte 7 Tage</option><option value="14">Letzte 14 Tage</option><option value="30">Letzte 30 Tage</option></select></div>
          <label class="chk"><input type="checkbox" data-ref="iUnread" checked> nur ungelesene</label>
          <span class="hint" data-ref="iLast" style="flex:1;text-align:right"></span>
        </div>
        <div class="msg" data-ref="iMsg"></div>
        <div data-ref="iList"></div>
      </section>

      <!-- ENTWÜRFE -->
      <section data-tab="drafts">
        <div class="hint">Angefangene Kommentare, Mails und neue Tickets werden automatisch gespeichert (nur in diesem Browser; Dateianhänge nicht).</div>
        <div data-ref="dView"></div>
      </section>

      <!-- EINSTELLUNGEN -->
      <section data-tab="settings">
        <h3 style="margin:0 0 6px">Darstellung</h3>
        <div class="row"><div class="c s" style="flex-basis:260px"><select data-ref="stTheme"><option value="auto">Automatisch (wie Windows/Browser)</option><option value="light">Hell</option><option value="dark">Dunkel</option></select></div></div>
        <h3 style="margin:14px 0 6px">Signatur</h3>
        <div class="hint">Wird bei „Als Mail senden“ automatisch angehängt (abschaltbar) und steht im Editor unter „✍ Signatur“ bereit. Einfügen aus Outlook inkl. Logo per Strg+V möglich.</div>
        <div class="rte" contenteditable="true" spellcheck="true" lang="de" data-ref="stSig" style="margin-top:6px;font-family:Arial;font-size:10pt"></div>
        <label class="chk"><input type="checkbox" data-ref="stAuto"> Signatur bei neuen Mails automatisch anhängen</label>
        <h3 style="margin:14px 0 6px">Eigene Textbausteine</h3>
        <div data-ref="stBlocks"></div>
        <button class="btn sm sec" data-ref="stAdd" style="margin-top:6px">＋ Baustein hinzufügen</button>
        <label class="chk" style="display:flex;margin-top:10px"><input type="checkbox" data-ref="stMd"> Textbausteine aus MultiData zusätzlich anbieten</label>
        <div style="margin-top:12px"><button class="btn" data-ref="stSave">Einstellungen speichern</button></div>
        <div class="msg" data-ref="stMsg"></div>

        <h3 style="margin:18px 0 6px">Team-Konfiguration (Abrechenbarkeit)</h3>
        <div class="hint">Steuert nur den Vorschlag für die Checkbox „abrechenbar“ und die Gruppierung im Zeiten-Tab. Änderungen werden sofort gespeichert.</div>
        <div style="margin-top:8px"><b>Nicht abrechenbar</b> <span class="hint">– Kunden mit Servicevertrag/Pauschale</span></div>
        <div data-ref="stNb"></div>
        <div style="margin-top:8px"><b>Intern</b> <span class="hint">– eigene Firmen; nie abrechenbar, im Zeiten-Tab als „Intern“</span></div>
        <div data-ref="stIn"></div>
        <label class="chk" style="display:flex;margin-top:8px"><input type="checkbox" data-ref="stOthers"> Alle anderen Kunden standardmäßig abrechenbar</label>
        <div class="row" style="margin-top:8px;align-items:center">
          <button class="btn sm sec" data-ref="stExport">⤓ Exportieren</button>
          <button class="btn sm sec" data-ref="stImport">⤒ Importieren …</button>
          <label class="chk" style="margin:0"><input type="checkbox" data-ref="stExpTpl"> Suchvorlagen mit exportieren</label>
          <input type="file" data-ref="stImportFile" accept=".json,application/json" style="display:none">
        </div>
        <h3 style="margin:14px 0 6px">Office-Dateien (Word, Excel, PowerPoint, Visio)</h3>
        <div class="hint">„🖥 In Office öffnen“ lädt die Originaldatei herunter, der Browser öffnet sie im installierten Office. Damit das ohne weiteren Klick passiert, einmal je Dateityp „Dateien dieses Typs immer öffnen“ wählen – die Werkbank zeigt dazu beim ersten Öffnen eine kurze Anleitung. Änderungen werden nicht nach MultiData zurückgeschrieben.</div>
        <div class="row" style="margin-top:6px;align-items:center"><span class="hint" data-ref="oStatus" style="flex:1"></span>
          <button class="btn sm sec" data-ref="oReset">Hinweise wieder anzeigen</button></div>
        <div class="hint" style="margin-top:8px">Die exportierte Team-Datei (JSON) enthält die Kundenlisten und kann an Kolleg:innen weitergegeben werden.</div>
        <div class="msg" data-ref="stTeamMsg"></div>

        <h3 style="margin:18px 0 6px">Lokale Daten</h3>
        <div class="hint">Signatur, Bausteine, Entwürfe, Filter, Personio-Zeiten, Posteingang und Protokoll liegen nur in diesem Browser (localStorage) – nichts davon wird in MultiData oder anderswo gespeichert.</div>
        <button class="btn sm warn" data-ref="stWipe" style="margin-top:6px">Alle Werkbank-Daten in diesem Browser löschen</button>
      </section>

      <!-- PROTOKOLL -->
      <section data-tab="log">
        <div class="row"><span class="hint" style="flex:1">Alle Änderungen, die über die Werkbank gemacht wurden (nur in diesem Browser gespeichert).</span>
          <button class="btn sm sec" data-ref="logClear">Protokoll leeren</button></div>
        <div data-ref="logView"></div>
      </section>
    </div>`;
  document.body.appendChild(panel);

  const $ = (ref, root = panel) => root.querySelector(`[data-ref="${ref}"]`);
  const el = (tag, attrs = {}, html = '') => { const e = document.createElement(tag); Object.assign(e, attrs); if (html) e.innerHTML = html; return e; };
  // Pop-up-Meldungen unten im Panel: bleiben sichtbar, egal wie weit man im Ticket gescrollt hat.
  // Klick schließt; gleiche Meldung erneut → Zeit wird verlängert statt doppelt angezeigt.
  const toastBox = el('div', { className: 'toasts' });
  panel.appendChild(toastBox);
  function toast(text, kind = 'err', ms = kind === 'err' ? 9000 : 6000) {
    const t = String(text || '').replace(/^\s*[✖⚠]\s*/, '').trim();
    if (!t) return;
    let n = [...toastBox.children].find(x => x.dataset.t === t);
    if (!n) {
      n = el('div', { className: 'toast ' + kind, role: 'alert', title: 'Klicken zum Schließen' });
      n.dataset.t = t;
      n.append(el('span', { textContent: (kind === 'err' ? '✖ ' : '⚠ ') + t }), el('b', { textContent: '×' }));
      n.addEventListener('click', () => n.remove());
      toastBox.appendChild(n);
      while (toastBox.children.length > 3) toastBox.firstChild.remove();
    }
    clearTimeout(n._timer); n._timer = setTimeout(() => n.remove(), ms);
  }
  const msg = (ref, text, kind = '') => {
    const m = $(ref); m.textContent = text; m.className = 'msg ' + kind;
    if (kind === 'err') toast(text, 'err');
  };
  const ui = { tab: 'search', ticket: null, comments: [], firm: null, editing: new Set(), sort: { key: 'SID', desc: true }, lastResult: null };

  // Zweistufige Bestätigung (statt Browser-Dialog, der MultiData blockieren würde)
  // label kann eine Funktion sein (z. B. Zusammenfassung „Mail an …“); wirft sie, wird der Fehler gemeldet und nicht scharf geschaltet.
  function confirmBtn(btn, label, fn, onError, ms = 4000) {
    let armed = false, timer, orig;
    btn.addEventListener('click', async ev => {
      ev.stopPropagation();
      if (!armed) {
        let l;
        try { l = typeof label === 'function' ? label() : label; } catch (e) { if (onError) onError(e); return; }
        orig = btn.textContent; armed = true; btn.textContent = l;
        timer = setTimeout(() => { armed = false; btn.textContent = orig; }, ms);
        return;
      }
      clearTimeout(timer); armed = false; btn.textContent = orig; btn.disabled = true;
      try { await fn(); } finally { btn.disabled = false; }
    });
  }
  // Auf-/Zugeklappt-Zustände dauerhaft speichern. Gespeichert wird nur, was der Nutzer selbst umschaltet
  // (programmatisches Öffnen, z. B. bei „Antworten“, ändert die gespeicherte Vorliebe nicht).
  const SECTIONS = CFG.store.sections;
  const secGet = (name, def) => { const v = store.get(SECTIONS, {})[name]; return v === undefined ? def : v; };
  const secSet = (name, v) => { const s = store.get(SECTIONS, {}); s[name] = v; store.set(SECTIONS, s); };
  function persistDetails(det, name, def) {
    det.open = secGet(name, def);
    det.querySelector('summary').addEventListener('click', () => setTimeout(() => secSet(name, det.open), 0));
    return det;
  }
  // Einmalige Anleitung je Dateityp: „Dateien dieses Typs immer öffnen“ im Browser einstellen
  function showOfficeHint(card, ext) {
    card.querySelector('.ohint')?.remove();
    const box = el('div', { className: 'warnbox ohint' });
    box.innerHTML = `<b>Einmalig einrichten:</b> Damit sich .${esc(ext)}-Dateien künftig direkt in Office öffnen, ohne Klick auf den Download: <span></span>
      <div class="row" style="margin-top:6px;gap:6px"><button class="btn sm" data-o="done">Erledigt</button><button class="btn sm sec" data-o="all">Nicht mehr anzeigen</button></div>`;
    box.querySelector('span').textContent = officeHowTo(ext);
    box.querySelector('[data-o=done]').addEventListener('click', () => { const a = officeAck(); if (!a.ext.includes(ext)) a.ext.push(ext); store.set(OFFICE_ACK, a); box.remove(); });
    box.querySelector('[data-o=all]').addEventListener('click', () => { store.set(OFFICE_ACK, { ...officeAck(), all: true }); box.remove(); });
    card.open = true; card.appendChild(box);
    box.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }
  // Kommentare: globaler Standard (alle auf/zu) + einzelne Abweichungen je Kommentar
  const COMMENTS = CFG.store.comments;
  const commentCollapsed = sid => { const c = store.get(COMMENTS, { all: false, ov: {} }); return c.ov[sid] ?? c.all; };
  function setCommentCollapsed(sid, v) {
    const c = store.get(COMMENTS, { all: false, ov: {} });
    if (v === c.all) delete c.ov[sid]; else c.ov[sid] = v;
    const keys = Object.keys(c.ov); if (keys.length > 3000) keys.slice(0, keys.length - 3000).forEach(k => delete c.ov[k]);
    store.set(COMMENTS, c);
  }
  const setAllComments = v => store.set(COMMENTS, { all: v, ov: {} });

  async function busy(btn, fn) { btn.disabled = true; try { return await fn(); } finally { btn.disabled = false; } }

  // Vorschlagsliste unter einem Eingabefeld. source(q) → [{ value, label }].
  // token=true: mehrere Werte mit „;“ – ergänzt wird immer der letzte Eintrag (für Mail-Empfänger).
  function attachSuggest(input, box, source, { onPick = null, token = false, max = 40 } = {}) {
    let items = [], idx = -1, timer;
    // Ohne eigene Eingabe (Feld zeigt nur den gewählten Wert) → alle Vorschläge zeigen
    // token-Modus: ist der letzte Eintrag schon eine vollständige Adresse, wird ein weiterer Empfänger vorgeschlagen (nicht ersetzt)
    const lastTok = () => input.value.split(';').pop().trim();
    const complete = () => token && validMail(lastTok());
    const current = () => token ? (complete() ? '' : lastTok()) : (input.dataset.typed ? input.value.trim() : '');
    const hide = () => { box.style.display = 'none'; idx = -1; };
    const mark = () => [...box.children].forEach((d, i) => d.style.background = i === idx ? '#e0f2f1' : '');
    async function show() {
      let list = [];
      try { list = (await source(current().toLowerCase())).slice(0, max); } catch { list = []; }
      items = list; idx = -1; box.innerHTML = '';
      if (!list.length || document.activeElement !== input) { hide(); return; }
      list.forEach((it, i) => { const d = el('div'); d.textContent = it.label; d.addEventListener('mousedown', e => { e.preventDefault(); pick(i); }); box.appendChild(d); });
      box.style.display = '';
    }
    function pick(i) {
      const it = items[i]; if (!it) return;
      if (token) {
        const parts = input.value.split(';');
        if (complete()) parts.push(it.value); else parts[parts.length - 1] = it.value;
        input.value = parts.map(s => s.trim()).filter(Boolean).join('; ') + '; ';
        input.focus(); show();
      } else { hide(); delete input.dataset.typed; if (onPick) onPick(it); }
    }
    input.addEventListener('input', () => { input.dataset.typed = '1'; clearTimeout(timer); timer = setTimeout(show, 150); });
    input.addEventListener('focus', show);
    input.addEventListener('blur', () => setTimeout(() => { hide(); delete input.dataset.typed; }, 150));
    input.addEventListener('keydown', e => {
      if (box.style.display === 'none') return;
      if (e.key === 'ArrowDown') { idx = Math.min(items.length - 1, idx + 1); mark(); e.preventDefault(); }
      else if (e.key === 'ArrowUp') { idx = Math.max(0, idx - 1); mark(); e.preventDefault(); }
      else if (e.key === 'Enter' && idx >= 0) { pick(idx); e.preventDefault(); e.stopPropagation(); }
      else if (e.key === 'Escape') hide();
    });
  }
  const userSource = (extra = []) => async q => {
    const users = await loadUsers();
    const hit = users.filter(u => !q || u.id.toLowerCase().includes(q) || u.name.toLowerCase().includes(q))
      .map(u => ({ value: u.id, label: `${u.name} (${u.id})${u.abt ? ' · ' + u.abt : ''}` }));
    return [...extra.filter(x => !q || x.label.toLowerCase().includes(q)), ...hit];
  };

  // --- Tabs / Öffnen ---
  const TAB_NAMES = { inbox: 'Inbox', search: 'Suche', ticket: 'Ticket', times: 'Zeiten', new: 'Neues Ticket', drafts: 'Entwürfe', log: 'Protokoll', settings: 'Einstellungen' };
  function showTab(tab) {
    ui.tab = tab;
    panel.querySelectorAll('.hdr .tab, section').forEach(x => x.classList.toggle('active', x.dataset.tab === tab));
    if (tab === 'log') renderLog();
    if (tab === 'new') { renderRecent(); restoreNewDraft(); }
    if (tab === 'inbox') renderInbox();
    if (tab === 'drafts') renderDrafts();
    if (tab === 'settings') renderSettings();
    if (tab === 'times') { renderTimes(); if (ui.timesStale && tm.data) { ui.timesStale = false; runTimes(); } }
    if (tab === 'search' && ui.searchStale) { ui.searchStale = false; runSearch(true); }
    setRoute(); updateTitle();
  }
  panel.querySelectorAll('.hdr .tab').forEach(b => b.addEventListener('click', () => showTab(b.dataset.tab)));
  function toggle() {
    if (panel.classList.contains('app')) return; // im App-Modus immer offen
    panel.classList.toggle('open');
    if (panel.classList.contains('open') && !ui.initDone) init();
  }
  fab.addEventListener('click', toggle);
  $('close').addEventListener('click', toggle);
  document.addEventListener('keydown', e => { if (e.altKey && (e.key === 'q' || e.key === 'Q')) { e.preventDefault(); toggle(); } });

  // --- Fenster verschieben (Kopfzeile ziehen) und Größe ändern (Ränder/Ecken) – Position & Größe werden gespeichert ---
  function applyGeom(g) {
    if (!g) return;
    const w = Math.max(480, Math.min(g.w, window.innerWidth)), h = Math.max(320, Math.min(g.h, window.innerHeight));
    const x = Math.max(100 - w, Math.min(g.x ?? (window.innerWidth - w - 12), window.innerWidth - 100));
    const y = Math.max(0, Math.min(g.y ?? 56, window.innerHeight - 40));
    Object.assign(panel.style, { width: w + 'px', height: h + 'px', left: x + 'px', top: y + 'px', right: 'auto', bottom: 'auto' });
  }
  const saveGeom = () => { const r = panel.getBoundingClientRect(); store.set(CFG.store.size, { x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) }); };
  applyGeom(store.get(CFG.store.size, null));
  window.addEventListener('resize', () => { if (store.get(CFG.store.size, null)) applyGeom(store.get(CFG.store.size, null)); });
  function dragWith(e, fn) {
    if (panel.classList.contains('full') || panel.classList.contains('app')) return;
    e.preventDefault();
    const r = panel.getBoundingClientRect(), x0 = e.clientX, y0 = e.clientY;
    const move = ev => applyGeom(fn(r, ev.clientX - x0, ev.clientY - y0));
    const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); document.body.style.userSelect = ''; saveGeom(); };
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
  }
  panel.querySelectorAll('[data-rz]').forEach(hd => hd.addEventListener('mousedown', e => {
    const m = hd.dataset.rz;
    dragWith(e, (r, dx, dy) => ({
      x: m.includes('l') ? r.left + dx : r.left, y: r.top,
      w: m.includes('l') ? r.width - dx : m.includes('r') ? r.width + dx : r.width,
      h: m.includes('b') ? r.height + dy : r.height,
    }));
  }));
  // Verschieben über die Kopfzeile (nicht über Tabs/Knöpfe); Doppelklick = Vollbild
  const hdr = panel.querySelector('.hdr');
  hdr.style.cursor = 'move';
  hdr.addEventListener('mousedown', e => {
    if (e.target.closest('button')) return;
    dragWith(e, (r, dx, dy) => ({ x: r.left + dx, y: r.top + dy, w: r.width, h: r.height }));
  });
  hdr.addEventListener('dblclick', e => { if (!e.target.closest('button')) panel.classList.toggle('full'); });
  $('full').addEventListener('click', () => panel.classList.toggle('full'));

  // --- App-Modus und mehrere Tabs ---
  // Adresse …/mdweb/rwi/HOME/0#werkbank[/ticket/123 | /search | /times …] öffnet die Werkbank als eigene Seite.
  // Modus und Ansicht werden pro Tab gemerkt (bleibt bei F5, auch wenn MultiData die Adresse umschreibt).
  const parseRoute = h => { const m = String(h || '').match(/^#werkbank(?:\/([a-z]+)(?:\/(\d+))?)?/); return m ? { tab: m[1] || null, sid: m[2] || null } : null; };
  const urlRoute = parseRoute(START_HASH);
  if (urlRoute) { tabStore.set('app', true); tabStore.set('route', urlRoute); }
  const appMode = tabStore.get('app', false) || new URLSearchParams(location.search).has('werkbank');
  const appUrl = r => location.origin + '/mdweb/rwi/HOME/0' + CFG.appHash + (r && r.tab ? '/' + r.tab + (r.sid ? '/' + r.sid : '') : '');
  // Links in der Werkbank sind echte Links auf die Werkbank-Adresse:
  //   Linksklick       → onLeft (meist: im aktuellen Tab öffnen)
  //   Mittelklick, Strg/⌘+Klick → Browser öffnet einen neuen Werkbank-Tab im Hintergrund (kein Tab-Wechsel)
  function wbLink(a, route, onLeft) {
    a.href = appUrl(route); a.dataset.wb = '1';
    a.addEventListener('click', e => {
      if (e.ctrlKey || e.metaKey || e.shiftKey) { e.stopPropagation(); return; }
      e.preventDefault();
      if (onLeft) { e.stopPropagation(); onLeft(e); }
    });
    return a;
  }
  // Ganze Tabellenzeile als Link: Zelleninhalte werden in Links verpackt, damit auch der Mittelklick überall greift
  function wbRow(tr, route, onLeft) {
    for (const td of tr.children) {
      const a = el('a', { className: 'rl' });
      while (td.firstChild) a.appendChild(td.firstChild);
      td.classList.add('l'); td.appendChild(a); wbLink(a, route, null);
    }
    tr.addEventListener('click', e => { if (!(e.ctrlKey || e.metaKey || e.shiftKey)) onLeft(e); });
  }
  const ticketRoute = sid => ({ tab: 'ticket', sid: String(sid) });
  // „In MultiData öffnen“: neuer Tab mit MultiData, dort wird das Ticket über die MultiData-Navigation geöffnet
  const mdTicketUrl = sid => location.origin + '/mdweb/rwi/HOME/0#mdopen/' + Number(sid);
  const mdOpen = (START_HASH.match(/^#mdopen\/(\d+)/) || [])[1];
  if (mdOpen) {
    try { history.replaceState(history.state, '', location.pathname + location.search); } catch { /* ignore */ }
    const t0 = Date.now();
    (function waitNav() {
      if (typeof window.navWI === 'function') return setTimeout(() => window.navWI('MDERRDA', JSON.stringify({ SID: Number(mdOpen) })), 300);
      if (Date.now() - t0 > 30000) { location.href = `/mdweb/rwi/MDERRDA/${encodeURIComponent(JSON.stringify({ SID: Number(mdOpen) }))}`; return; }
      setTimeout(waitNav, 250);
    })();
  }
  // Neuen Werkbank-Tab öffnen (Vordergrund). Filter werden über das sessionStorage mitgegeben, das der Browser beim Öffnen kopiert.
  function openInNewTab(route, { filters = null } = {}) {
    const keep = { app: tabStore.get('app', false), route: tabStore.get('route', null), filters: tabStore.get('filters', null) };
    tabStore.set('app', true); tabStore.set('route', route); if (filters) tabStore.set('filters', filters);
    window.open(appUrl(route), '_blank');
    tabStore.set('app', keep.app); tabStore.set('route', keep.route);
    if (filters) keep.filters ? tabStore.set('filters', keep.filters) : tabStore.del('filters');
  }
  const currentRoute = () => ({ tab: ui.tab, sid: ui.tab === 'ticket' && ui.ticket ? String(ui.ticket.SID) : null });
  function setRoute() {
    if (!appMode) return;
    const r = currentRoute(); tabStore.set('route', r);
    try { history.replaceState(history.state, '', appUrl(r)); } catch { /* ignore */ }
  }
  // Tab-Titel im App-Modus: „#123 Betreff – Werkbank“, „Suche – Werkbank“, mit Anzahl ungelesener Inbox-Einträge
  function updateTitle() {
    if (!appMode) return;
    const unread = Object.values(inboxState().items).filter(i => !i.read).length;
    const what = ui.tab === 'ticket' && ui.ticket ? `#${ui.ticket.SID} ${ui.ticket.TITEL || ''}` : TAB_NAMES[ui.tab] || 'Werkbank';
    ui.title = `${unread ? `(${unread}) ` : ''}${what} – Werkbank`;
    if (document.title !== ui.title) document.title = ui.title;
  }
  $('app').title = appMode ? 'Aktuelle Ansicht in neuem Tab öffnen' : 'Werkbank als eigene Seite in neuem Tab öffnen';
  wbLink($('app'), currentRoute(), () => openInNewTab(currentRoute(), { filters: getFilterState() }));
  $('app').addEventListener('mousedown', () => { $('app').href = appUrl(currentRoute()); }); // Mittelklick: aktuelle Ansicht
  if (appMode) {
    panel.classList.add('app', 'open'); fab.style.display = 'none';
    // MultiData setzt den Seitentitel selbst neu – wir halten ihn auf dem Werkbank-Titel
    new MutationObserver(() => { if (ui.title && document.title !== ui.title) document.title = ui.title; })
      .observe(document.querySelector('title') || document.head, { childList: true, subtree: true, characterData: true });
    setTimeout(() => init(), 0);
  }

  async function init() {
    ui.initDone = true;
    await loadUsers().catch(() => null);
    applyFilterState(tabStore.get('filters', null) || store.get(CFG.store.filters, null) || defaultFilterState());
    renderTemplates();
    if (appMode) {
      // gemerkte Ansicht dieses Tabs wiederherstellen
      const r = tabStore.get('route', null) || {};
      if (r.tab === 'ticket' && r.sid) openTicket(r.sid); else showTab(TAB_NAMES[r.tab] ? r.tab : 'search');
      return;
    }
    // Overlay über MultiData: mit dem dort geöffneten Ticket starten
    const t = await currentPageTicket().catch(() => null);
    if (t) openTicket(t.SID, false);
  }

  // ================= SUCHE =================
  function renderStatusChips() {
    const box = $('sStatus'); box.innerHTML = '';
    for (const [id, name, color] of CFG.statuses) {
      const c = el('button', { className: 'chip' + (ui.selStatus.has(id) ? ' sel' : '') }, `<span class="dot" style="background:${color}"></span>${esc(name)}`);
      c.addEventListener('click', () => { ui.selStatus.has(id) ? ui.selStatus.delete(id) : ui.selStatus.add(id); renderStatusChips(); scheduleFacets(); });
      box.appendChild(c);
    }
  }
  $('sOpen').addEventListener('click', () => { ui.selStatus = new Set(CFG.statuses.filter(s => s[0] !== 'Z').map(s => s[0])); renderStatusChips(); scheduleFacets(); });
  $('sAll').addEventListener('click', () => { ui.selStatus = new Set(); renderStatusChips(); scheduleFacets(); });

  // Schnell-Zeiträume; in Vorlagen wird nur der Schlüssel gespeichert (z. B. „Dieser Monat“ bleibt relativ)
  const RANGES = [['heute', 'Heute'], ['7tage', '7 Tage'], ['monat', 'Dieser Monat'], ['vormonat', 'Letzter Monat']];
  function rangeDates(key) {
    const t = todayStr(), d = new Date(), first = (y, m) => `${y}-${pad(m + 1)}-01`;
    const lastOf = (y, m) => addDays(first(m === 11 ? y + 1 : y, (m + 1) % 12), -1);
    const pm = d.getMonth() === 0 ? 11 : d.getMonth() - 1, py = d.getMonth() === 0 ? d.getFullYear() - 1 : d.getFullYear();
    return { heute: [t, t], '7tage': [addDays(t, -6), t], monat: [first(d.getFullYear(), d.getMonth()), lastOf(d.getFullYear(), d.getMonth())], vormonat: [first(py, pm), lastOf(py, pm)] }[key] || ['', ''];
  }
  function renderRanges() {
    const box = $('sRanges'); box.innerHTML = '';
    for (const [key, n] of [...RANGES, ['', 'Leeren']]) {
      const c = el('button', { className: 'chip' + (key && ui.rangeKey === key ? ' sel' : ''), textContent: n });
      c.addEventListener('click', () => {
        const [a, b] = rangeDates(key);
        $('sFrom').value = a; $('sTo').value = b; ui.rangeKey = key || null;
        if (a && $('sDateMode').value === 'none') $('sDateMode').value = 'booked';
        if (!a) $('sDateMode').value = 'none';
        renderRanges(); scheduleFacets();
      });
      box.appendChild(c);
    }
  }
  ['sFrom', 'sTo'].forEach(r => $(r).addEventListener('input', () => { ui.rangeKey = null; renderRanges(); }));

  // Mehrfachauswahl mit Suche: Auswahl als „Pillen“, Eingabefeld mit Vorschlagsliste
  function multiPick(container, { placeholder = '', source, onChange = null, pillLabel = null }) {
    let vals = []; // [{ value, label }]
    container.className = 'mp';
    container.innerHTML = `<input type="text" autocomplete="off" placeholder="${esc(placeholder)}"><div class="results" style="display:none"></div>`;
    const input = container.querySelector('input'), box = container.querySelector('.results');
    const render = () => {
      container.querySelectorAll('.pill').forEach(p => p.remove());
      for (const v of vals) {
        const p = el('span', { className: 'pill', title: v.label }); p.appendChild(el('span', { textContent: v.label }));
        const x = el('b', { textContent: '×', title: 'entfernen' }); x.addEventListener('click', e => { e.stopPropagation(); api_.remove(v.value); });
        p.appendChild(x); container.insertBefore(p, input);
      }
      input.placeholder = vals.length ? '' : placeholder;
    };
    const api_ = {
      get: () => vals.map(v => v.value), items: () => [...vals],
      has: v => vals.some(x => String(x.value) === String(v)),
      set(list) { vals = list.filter(Boolean); render(); },
      add(v) { if (!api_.has(v.value)) { vals.push(v); render(); onChange && onChange(); } },
      remove(value) { vals = vals.filter(x => String(x.value) !== String(value)); render(); onChange && onChange(); },
      toggle(v) { api_.has(v.value) ? api_.remove(v.value) : api_.add(v); },
    };
    container.addEventListener('click', () => input.focus());
    attachSuggest(input, box, async q => (await source(q)).filter(it => !api_.has(it.value)), { onPick: it => { api_.add(pillLabel ? { value: it.value, label: pillLabel(it) } : it); input.value = ''; setTimeout(() => input.focus(), 0); } });
    input.addEventListener('keydown', e => { if (e.key === 'Backspace' && !input.value && vals.length) api_.remove(vals[vals.length - 1].value); });
    return api_;
  }

  // Bearbeiter: „Ich“ = __ME__
  const userPick = multiPick($('sUsers'), {
    placeholder: 'alle Bearbeiter – Name tippen …',
    source: async q => [...(!q || 'ich'.includes(q) ? [{ value: '__ME__', label: '👤 Ich' }] : []),
      ...(!q || /^(o|oh|ohn|ohne|l|le|lee|leer|k|ke|kei|kein|n|ni|nic|nich|nicht)/.test(q) ? [{ value: NONE, label: '∅ ohne Bearbeiter' }] : []), ...(await userSource()(q))],
    onChange: () => scheduleFacets(),
  });
  const userItem = id => id === '__ME__' ? { value: '__ME__', label: '👤 Ich' } : id === NONE ? { value: NONE, label: '∅ ohne Bearbeiter' } : { value: id, label: userLabel(id) };

  // Kunden: Vorschläge aus der Übersicht (aktuelle Filter) + Suche über alle Tickets
  ui.custNames = {};
  const custLabel = (id, name) => `${(name || ui.custNames[id] || id).replace(/^\d+\s*-\s*/, '')}`;
  const custPick = multiPick($('sCusts'), {
    placeholder: 'alle Kunden – Name oder Nr. tippen …',
    source: async q => {
      const fac = (ui.facets || []).filter(c => !q || c.name.toLowerCase().includes(q) || String(c.id).includes(q));
      let more = [];
      if (q && q.length >= 2) more = (await searchCustomersByTickets(q).catch(() => [])).filter(c => !fac.some(f => f.id === c.id));
      for (const c of [...fac, ...more]) ui.custNames[c.id] = c.name;
      return [...fac.map(c => ({ value: c.id, label: `${custLabel(c.id, c.name)} (${c.n})` })), ...more.map(c => ({ value: c.id, label: `${custLabel(c.id, c.name)} · ${c.n} Tickets gesamt` }))];
    },
    onChange: () => { renderFacets(); },
    pillLabel: it => custLabel(it.value),
  });

  // Queues (Ticketeingang, Team 1, Intern …)
  const queuePick = multiPick($('sQueues'), {
    placeholder: 'alle Queues …',
    source: async q => [...(!q || /^(o|oh|ohn|ohne|l|le|lee|leer|k|ke|kei|kein)/.test(q) ? [{ value: NONE, label: '∅ ohne Queue' }] : []),
      ...(await loadQueues()).filter(x => !q || x.name.toLowerCase().includes(q)).map(x => ({ value: String(x.id), label: x.name }))],
    onChange: () => scheduleFacets(),
  });

  // Übersicht der Kunden, die in den aktuellen Filtern (Bearbeiter, Queues, Status, Zeitraum, Freitext) vorkommen
  let facetTimer, facetSeq = 0;
  function scheduleFacets() { clearTimeout(facetTimer); facetTimer = setTimeout(loadFacets, 600); }
  async function loadFacets() {
    const seq = ++facetSeq;
    $('sFacetHint').textContent = '(lädt …)';
    try {
      const { list, truncated } = await customerFacets(readFilters(false));
      if (seq !== facetSeq) return;
      ui.facets = list;
      for (const c of list) ui.custNames[c.id] = c.name;
      $('sFacetHint').textContent = `(${list.length} Kunden${truncated ? ', nur die neuesten 5000 Tickets ausgewertet' : ''} – Klick wählt aus/ab)`;
      renderFacets();
    } catch (e) { if (seq === facetSeq) $('sFacetHint').textContent = '(Fehler: ' + e.message + ')'; }
  }
  // Sortierung der Kundenübersicht (Anzahl/Name, auf-/absteigend) – gilt für alle Tabs
  const facetSort = () => store.get(CFG.store.facetSort, 'n-desc');
  $('sFacetSort').value = facetSort();
  $('sFacetSort').addEventListener('change', () => { store.set(CFG.store.facetSort, $('sFacetSort').value); renderFacets(); });
  function sortedFacets() {
    const [k, dir] = facetSort().split('-'), m = dir === 'asc' ? 1 : -1;
    const nm = c => custLabel(c.id, c.name).toLowerCase();
    return [...(ui.facets || [])].sort((a, b) => k === 'name' ? m * nm(a).localeCompare(nm(b), 'de') : (m * (a.n - b.n) || nm(a).localeCompare(nm(b), 'de')));
  }
  // Auf zwei Zeilen kürzen; „+N weitere“ klappt den Rest auf
  ui.facetsOpen = false;
  function clampFacets() {
    const box = $('sFacets'), chips = [...box.querySelectorAll('.chip:not(.more)')];
    box.querySelector('.chip.more')?.remove();
    chips.forEach(c => { c.style.display = ''; });
    if (!chips.length || !box.offsetWidth) return;
    const tops = [...new Set(chips.map(c => c.offsetTop))].sort((a, b) => a - b);
    if (tops.length <= 2) return;
    const more = el('button', { className: 'chip more' });
    if (ui.facetsOpen) {
      more.textContent = '▴ weniger'; box.appendChild(more);
    } else {
      const limit = tops[1];
      const hidden = chips.filter(c => c.offsetTop > limit);
      hidden.forEach(c => { c.style.display = 'none'; });
      box.appendChild(more);
      const visible = chips.filter(c => c.style.display !== 'none');
      more.textContent = `+${hidden.length} weitere ▾`;
      while (more.offsetTop > limit && visible.length) {
        const c = visible.pop(); c.style.display = 'none'; hidden.push(c); more.textContent = `+${hidden.length} weitere ▾`;
      }
    }
    more.addEventListener('click', () => { ui.facetsOpen = !ui.facetsOpen; clampFacets(); });
  }
  if ('ResizeObserver' in window) { let w = 0; new ResizeObserver(() => { const nw = $('sFacets').offsetWidth; if (nw !== w) { w = nw; clampFacets(); } }).observe($('sFacets')); }
  function renderFacets() {
    const box = $('sFacets'); box.innerHTML = '';
    for (const c of sortedFacets()) {
      const b = el('button', { className: 'chip' + (custPick.has(c.id) ? ' sel' : ''), title: c.name });
      b.textContent = custLabel(c.id, c.name); b.appendChild(el('span', { className: 'n', textContent: c.n }));
      b.addEventListener('click', () => custPick.toggle({ value: c.id, label: custLabel(c.id, c.name) }));
      box.appendChild(b);
    }
    // gewählte Kunden, die im aktuellen Filter nicht vorkommen, trotzdem anzeigen
    for (const v of custPick.items().filter(v => !(ui.facets || []).some(c => String(c.id) === String(v.value)))) {
      const b = el('button', { className: 'chip sel', title: 'gewählt, im aktuellen Filter ohne Treffer' }); b.textContent = v.label; b.appendChild(el('span', { className: 'n', textContent: 0 }));
      b.addEventListener('click', () => custPick.remove(v.value)); box.appendChild(b);
    }
    if (!box.children.length) box.innerHTML = '<span class="hint">Keine Kunden in den aktuellen Filtern.</span>';
    // gewählte Kunden immer sichtbar halten (vorne), Rest in gewählter Reihenfolge
    [...box.querySelectorAll('.chip.sel')].reverse().forEach(b => box.prepend(b));
    clampFacets();
  }
  ['sDateMode', 'sFrom', 'sTo'].forEach(r => $(r).addEventListener('change', scheduleFacets));
  $('sText').addEventListener('input', scheduleFacets);

  // Filterzustand der Suche. Gilt pro Browser-Tab (sessionStorage); der zuletzt benutzte ist Vorgabe für neue Tabs.
  const OPEN_STATUSES = CFG.statuses.filter(s => s[0] !== 'Z').map(s => s[0]);
  const defaultFilterState = () => ({ users: [userItem('__ME__')], custs: [], queues: [], statuses: OPEN_STATUSES, dateMode: 'none', from: '', to: '', range: null, text: '', group: true });
  function getFilterState() {
    return { users: userPick.items(), custs: custPick.items(), queues: queuePick.items(), statuses: [...ui.selStatus], dateMode: $('sDateMode').value,
      from: $('sFrom').value, to: $('sTo').value, range: ui.rangeKey || null, text: $('sText').value.trim(), group: $('sGroup').checked };
  }
  function applyFilterState(fs) {
    fs = { ...defaultFilterState(), ...fs };
    userPick.set((fs.users || []).map(u => typeof u === 'string' ? userItem(u) : u));       // ältere Version: nur IDs
    custPick.set((fs.custs || []).map(c => { ui.custNames[c.value] ||= c.label; return c; }));
    queuePick.set(fs.queues || []);
    ui.selStatus = new Set(fs.statuses || OPEN_STATUSES);
    $('sText').value = fs.text || ''; $('sDateMode').value = fs.dateMode || 'none';
    ui.rangeKey = fs.range || null;
    const [a, b] = fs.range ? rangeDates(fs.range) : [fs.from || '', fs.to || ''];
    $('sFrom').value = a; $('sTo').value = b;
    $('sGroup').checked = fs.group !== false;
    renderStatusChips(); renderRanges(); renderFacets(); scheduleFacets();
  }
  function readFilters(save = true) {
    const fs = getFilterState();
    const f = { users: userPick.get(), kndIds: custPick.get(), queues: queuePick.get(), statuses: fs.statuses, dateMode: fs.dateMode, from: fs.from, to: fs.to, text: fs.text };
    if (f.dateMode !== 'none' && !f.from && !f.to) f.dateMode = 'none';
    if (save) { tabStore.set('filters', fs); store.set(CFG.store.filters, fs); }
    return f;
  }

  // Suchvorlagen (für alle Tabs dieses Browsers; über ⚙ exportierbar)
  const templates = () => store.get(CFG.store.templates, null) ?? [
    { id: 'my-open', name: 'Meine offenen Tickets', fs: { users: [userItem('__ME__')], statuses: OPEN_STATUSES } },
    { id: 'all-open', name: 'Alle offenen Tickets', fs: { users: [], statuses: OPEN_STATUSES } },
  ];
  function renderTemplates() {
    const box = $('sTpl'); box.innerHTML = '';
    for (const tp of templates()) {
      const c = el('button', { className: 'chip', title: 'Vorlage anwenden und suchen' }); c.textContent = tp.name;
      const x = el('b', { textContent: ' ×', title: 'Vorlage löschen' }); x.style.marginLeft = '4px';
      c.appendChild(x);
      c.addEventListener('click', e => {
        if (e.target === x) return;
        applyFilterState({ users: [], custs: [], queues: [], ...tp.fs }); runSearch(); // Vorlage gilt vollständig (fehlende Felder = alle)
      });
      confirmBtn(x, ' Löschen?', async () => { store.set(CFG.store.templates, templates().filter(t => t.id !== tp.id)); renderTemplates(); });
      box.appendChild(c);
    }
  }
  $('sTplAdd').addEventListener('click', () => { $('sTplNew').style.display = ''; $('sTplName').focus(); });
  const saveTemplate = () => {
    const name = $('sTplName').value.trim(); if (!name) return;
    store.set(CFG.store.templates, [...templates().filter(t => t.name !== name), { id: 't' + Date.now(), name, fs: getFilterState() }]);
    $('sTplName').value = ''; $('sTplNew').style.display = 'none'; renderTemplates();
    msg('sMsg', `✔ Vorlage „${name}“ gespeichert.`, 'ok');
  };
  $('sTplOk').addEventListener('click', saveTemplate);
  $('sTplName').addEventListener('keydown', e => { if (e.key === 'Enter') saveTemplate(); if (e.key === 'Escape') $('sTplNew').style.display = 'none'; });
  // silent = automatische Aktualisierung (andere Tabs/MultiData haben etwas geändert): Scrollposition bleibt
  async function runSearch(silent = false) {
    if (!silent) msg('sMsg', 'Suche …');
    const body = panel.querySelector('.body'), top = body.scrollTop;
    try {
      const res = await searchTickets(readFilters(!silent));
      ui.lastResult = res; ui.lastSearchAt = Date.now();
      renderResults();
      if (silent) body.scrollTop = top;
      const shown = res.rows.length, total = res.total ?? shown;
      msg('sMsg', `${total} Ticket(s)${total > shown ? ` – angezeigt werden die neuesten ${shown}` : ''}.${silent ? ` · automatisch aktualisiert ${new Date().toLocaleTimeString('de-DE').slice(0, 5)}` : ''}`, 'ok');
    } catch (e) { msg('sMsg', '✖ ' + e.message, 'err'); }
  }
  $('sGo').addEventListener('click', () => busy($('sGo'), () => runSearch()));
  $('sText').addEventListener('keydown', e => { if (e.key === 'Enter') runSearch(); });

  function renderResults() {
    const res = ui.lastResult; if (!res) return;
    const bh = res.bookedHours;
    const cols = [['PRIORITY', 'Prio'], ['SID', 'ID'], ['TITEL', 'Betreff'], ['MATCHCODE', 'Kunde'], ['STATUS', 'Status'], ['ANLAGE_DATUM', 'Angelegt'],
      ['BEARBEITER_USER_ID', 'Bearbeiter'], ['SUM_STD', 'Std.'], ['MEMO_SUM_ABRSTD', 'abr.']];
    if (bh) cols.push(['_BSTD', 'Std. im Zeitraum'], ['_BABR', 'abr. im Zeitraum']);
    const rows = res.rows.map(r => ({ ...r, _BSTD: bh?.[r.SID]?.std ?? null, _BABR: bh?.[r.SID]?.abr ?? null }));
    const { key, desc } = ui.sort;
    const group = $('sGroup').checked;
    rows.sort((a, b) => {
      if (group) { const pa = a.PRIORITY ?? 99, pb = b.PRIORITY ?? 99; if (pa !== pb) return pa - pb; } // wichtigste Gruppe zuerst
      const x = a[key] ?? '', y = b[key] ?? ''; return (x > y ? 1 : x < y ? -1 : 0) * (desc ? -1 : 1);
    });

    const tbl = el('table');
    const thr = el('tr');
    for (const [k, n] of cols) {
      const th = el('th', { textContent: n + (key === k ? (desc ? ' ↓' : ' ↑') : '') });
      th.addEventListener('click', () => { ui.sort = { key: k, desc: key === k ? !desc : false }; renderResults(); });
      thr.appendChild(th);
    }
    tbl.appendChild(el('thead')).appendChild(thr);
    const tb = el('tbody'); tbl.appendChild(tb);
    const sum = { std: 0, abr: 0, b: 0, ba: 0 };
    const sumRow = (label, s_, n, cls = '') => {
      const tf = el('tr', { className: cls });
      tf.innerHTML = `<td></td><td></td><td><b>${esc(label)} (${n})</b></td><td></td><td></td><td></td><td></td><td class="n"><b>${fmtH(s_.std)}</b></td><td class="n"><b>${fmtH(s_.abr)}</b></td>` +
        (bh ? `<td class="n"><b>${fmtH(s_.b)}</b></td><td class="n"><b>${fmtH(s_.ba)}</b></td>` : '');
      tb.appendChild(tf);
    };
    let cur, gs = null, gn = 0;
    const flush = () => { if (gs && group) sumRow('Summe ' + prioInfo(cur)[1], gs, gn, 'gsum'); };
    for (const r of rows) {
      if (group && (gs === null || r.PRIORITY !== cur)) {
        flush(); cur = r.PRIORITY; gs = { std: 0, abr: 0, b: 0, ba: 0 }; gn = 0;
        const gh = el('tr', { className: 'ghead' }); gh.innerHTML = `<td colspan="${cols.length}">${prioHtml(cur)}</td>`; tb.appendChild(gh);
      }
      const [, sn, sc] = statusInfo(r.STATUS);
      const tr = el('tr', { className: 'r' });
      tr.innerHTML = `<td style="white-space:nowrap">${prioHtml(r.PRIORITY)}</td><td class="n">${r.SID}</td><td>${esc(r.TITEL)}</td><td>${esc((r.MATCHCODE || '').replace(/^\d+\s*-\s*/, ''))}</td>
        <td><span class="dot" style="background:${sc}"></span>${esc(sn)}</td><td>${fmtDate(r.ANLAGE_DATUM)}</td><td>${esc(r.BEARBEITER_USER_ID || '')}</td>
        <td class="n">${fmtH(r.SUM_STD)}</td><td class="n">${fmtH(r.MEMO_SUM_ABRSTD)}</td>` +
        (bh ? `<td class="n">${fmtH(r._BSTD)}</td><td class="n">${fmtH(r._BABR)}</td>` : '');
      tr.title = 'Klick: öffnen · Mittelklick oder Strg+Klick: in neuem Tab im Hintergrund';
      wbRow(tr, ticketRoute(r.SID), () => openTicket(r.SID));
      tb.appendChild(tr);
      for (const t of [sum, gs].filter(Boolean)) { t.std += r.SUM_STD || 0; t.abr += r.MEMO_SUM_ABRSTD || 0; t.b += r._BSTD || 0; t.ba += r._BABR || 0; }
      gn++;
    }
    flush();
    sumRow('Summe', sum, rows.length, 'tsum');
    const box = $('sResults'); box.innerHTML = ''; box.appendChild(tbl);
  }
  $('sGroup').addEventListener('change', () => { readFilters(); renderResults(); });
  $('sReset').addEventListener('click', () => { applyFilterState(defaultFilterState()); readFilters(); msg('sMsg', 'Filter zurückgesetzt.'); });

  // ================= TICKET =================
  async function openTicket(sid, switchTab = true) {
    if (switchTab) showTab('ticket');
    if (ui.editing.size && ui.ticket && String(ui.ticket.SID) !== String(sid)) msg('tMsg', 'Hinweis: ungespeicherte Kommentar-Änderungen wurden verworfen.');
    else msg('tMsg', 'Lade …');
    $('tBanner').style.display = 'none'; ui.ticketSig = null;
    try {
      const t = await ticketBySid(sid);
      if (!t) throw new Error(`Ticket ${sid} nicht gefunden.`);
      ui.ticket = t; ui.editing.clear();
      setRoute(); updateTitle();
      ticketSignature(t.SID).then(sig => { if (ui.ticket && ui.ticket.SID === t.SID) ui.ticketSig = sig; }).catch(() => null);
      [ui.comments, ui.docs, ui.children] = await Promise.all([loadComments(sid), loadDocs(sid).catch(() => []), loadChildren(sid).catch(() => [])]);
      $('tSid').value = t.SID;
      await renderTicket();
      msg('tMsg', '');
    } catch (e) { msg('tMsg', '✖ ' + e.message, 'err'); }
  }
  const reloadTicket = () => ui.ticket && openTicket(ui.ticket.SID, false);
  $('tLoad').addEventListener('click', () => { const v = $('tSid').value.trim(); if (v) openTicket(v); });
  $('tSid').addEventListener('keydown', e => { if (e.key === 'Enter') $('tLoad').click(); });
  $('tPage').addEventListener('click', async () => {
    const t = await currentPageTicket().catch(() => null);
    t ? openTicket(t.SID) : msg('tMsg', 'In MultiData ist gerade kein Ticket geöffnet.', 'err');
  });
  $('tReload').addEventListener('click', reloadTicket);

  async function renderTicket() {
    const t = ui.ticket, view = $('tView');
    const ka = await kontaktarten();
    view.innerHTML = '';

    // --- Kopf ---
    const head = el('div', { className: 'head' });
    head.innerHTML = `
      <h3>#${t.SID} <a href="#" data-ref="hOpen" target="_blank" style="font-size:12px;font-weight:400;margin-left:8px" title="Öffnet MultiData mit diesem Ticket in einem neuen Tab">in MultiData öffnen ↗</a>
        <a href="#" data-ref="hNewTab" style="font-size:12px;font-weight:400;margin-left:8px" title="Dieses Ticket in einem weiteren Werkbank-Tab öffnen">in neuem Tab ⧉</a></h3>
      <div class="row">
        <div class="c" style="flex:3"><label class="l">Betreff</label><input type="text" data-ref="hTitle" spellcheck="true" lang="de"></div>
        <div class="c"><label class="l">Status</label><select data-ref="hStatus"></select></div>
        <div class="c"><label class="l">Priorität</label><select data-ref="hPrio"></select></div>
        <div class="c"><label class="l">Queue</label><select data-ref="hQueue"></select></div>
        <div><button class="btn" data-ref="hSave">Speichern</button></div>
      </div>
      <div class="row" style="margin-top:6px;align-items:center">
        <div class="c" style="flex:3"><label class="l">Firma</label><div><b data-ref="hFirm"></b> <button class="btn sm sec" data-ref="hFirmBtn">ändern …</button></div></div>
        <div class="c sums" data-ref="hSums"></div>
        <div><button class="btn warn" data-ref="hQC">QuickClose</button></div>
      </div>
      <div class="row" style="margin-top:6px;align-items:flex-end">
        <div class="c" style="max-width:300px;position:relative"><label class="l">Bearbeiter <span class="hint">(suchen, mit „Speichern“ übernehmen)</span></label>
          <input type="text" data-ref="hUser" autocomplete="off"><div class="results" data-ref="hUserRes" style="display:none;position:absolute;left:0;right:0;z-index:5"></div></div>
        <label class="chk"><input type="checkbox" data-ref="hNoNotify"> Keine Benachrichtigungen <span class="hint">(gesetzt = keine Mails)</span></label>
        <div class="c sums" data-ref="hMail"></div>
      </div>
      <div class="row" style="margin-top:6px;align-items:flex-end">
        <div class="c" style="max-width:300px"><label class="l">Parent-Ticket <span class="hint">(Ticket-ID, leer = keins; mit „Speichern“)</span></label>
          <input type="text" data-ref="hParent" placeholder="Ticket-Nr."></div>
        <div class="c sums" data-ref="hParentInfo"></div>
        <span style="flex:1"></span>
        <button class="btn sm sec" data-ref="hKids" title="Child-Tickets anzeigen, zuordnen, erstellen">🌿 Child-Tickets</button>
        <button class="btn sm sec" data-ref="hMerge" title="Dieses Ticket in ein anderes zusammenführen">🔀 Zusammenführen</button>
      </div>
      <div data-ref="hExtra"></div>
      <div data-ref="hFirmBox" style="display:none;margin-top:6px">
        <input type="text" data-ref="hFirmSearch" placeholder="Neue Firma suchen (mind. 3 Zeichen) …">
        <div class="results" data-ref="hFirmRes" style="display:none"></div>
        <div class="msg" data-ref="hFirmMsg"></div>
      </div>`;
    view.appendChild(head);
    $('hTitle', head).value = t.TITEL || '';
    const sel = $('hStatus', head);
    const known = CFG.statuses.some(s => s[0] === t.STATUS);
    for (const [id, name] of (known ? CFG.statuses : [...CFG.statuses, [t.STATUS, t.STATUS]])) sel.appendChild(el('option', { value: id, textContent: name }));
    sel.value = t.STATUS;
    $('hFirm', head).textContent = t.MATCHCODE || t.KND_ID;
    // Priorität
    const pSel = $('hPrio', head);
    for (const [v, n] of (CFG.priorities.some(x => x[0] === Number(t.PRIORITY)) || t.PRIORITY == null ? CFG.priorities : [...CFG.priorities, [t.PRIORITY, String(t.PRIORITY)]]))
      pSel.appendChild(el('option', { value: v, textContent: n }));
    pSel.value = t.PRIORITY ?? '';
    const syncPrio = () => { pSel.style.borderLeft = `6px solid ${prioInfo(pSel.value)[2]}`; };
    pSel.addEventListener('change', syncPrio); syncPrio();
    // Queue
    const qSel = $('hQueue', head);
    loadQueues().then(qs => {
      if (!qs.some(q => String(q.id) === String(t.ERRQ_ID))) qs = [...qs, { id: t.ERRQ_ID, name: String(t.ERRQ_ID ?? '–') }];
      for (const q of qs) qSel.appendChild(el('option', { value: q.id, textContent: q.name }));
      qSel.value = t.ERRQ_ID;
    }).catch(() => qSel.appendChild(el('option', { value: t.ERRQ_ID, textContent: String(t.ERRQ_ID) })));
    // Parent
    $('hParent', head).value = t.MDERRDA_SID || '';
    if (t.MDERRDA_SID) ticketBySid(t.MDERRDA_SID).then(p => {
      const info = $('hParentInfo', head); info.innerHTML = '';
      if (!p) { info.textContent = 'Parent nicht gefunden'; return; }
      const a = wbLink(el('a', { textContent: `↑ #${p.SID} ${p.TITEL}` }), ticketRoute(p.SID), () => openTicket(p.SID));
      info.appendChild(a); info.appendChild(document.createTextNode(` · ${statusInfo(p.STATUS)[1]}`));
    }).catch(() => null);
    $('hNoNotify', head).checked = t.SUPPRESS_NOTIFY == 1;
    let newUser = t.BEARBEITER_USER_ID || null;
    wbLink($('hNewTab', head), ticketRoute(t.SID), () => openInNewTab(ticketRoute(t.SID)));
    const hUser = $('hUser', head);
    loadUsers().then(() => {
      hUser.value = t.BEARBEITER_USER_ID ? userLabel(t.BEARBEITER_USER_ID) : '';
      if (t.ANLAGE_USER_ID) $('hCreator', head).textContent = userLabel(t.ANLAGE_USER_ID); // Ersteller (MultiData-Feld „Ersteller“)
    }).catch(() => { hUser.value = t.BEARBEITER_USER_ID || ''; });
    attachSuggest(hUser, $('hUserRes', head), userSource([{ value: '__ME__', label: '👤 Mir zuweisen' }]), {
      onPick: async it => {
        newUser = it.value === '__ME__' ? await me() : it.value;
        hUser.value = userLabel(newUser);
        hUser.style.outline = newUser !== t.BEARBEITER_USER_ID ? '2px solid #ffb300' : '';
      },
    });
    hUser.addEventListener('focus', () => hUser.select());
    $('hMail', head).textContent = `An: ${t.EMAIL_TO || '–'}${t.EMAIL_CC ? ' · Cc: ' + t.EMAIL_CC : ''}`;
    $('hSums', head).innerHTML = `Angelegt ${fmtDate(t.ANLAGE_DATUM)} von <b data-ref="hCreator">${esc(t.ANLAGE_USER_ID || '–')}</b> · Bearbeiter ${esc(t.BEARBEITER_USER_ID || '–')}<br>
      Stunden <b>${fmtH(t.SUM_STD)}</b> · abrechenbar <b>${fmtH(t.MEMO_SUM_ABRSTD)}</b>`;
    $('hOpen', head).href = mdTicketUrl(t.SID); // immer neuer Tab (Linksklick: wechseln, Mittelklick: im Hintergrund)
    $('hSave', head).addEventListener('click', ev => busy(ev.target, async () => {
      try {
        const ch = {};
        const title = $('hTitle', head).value.trim();
        if (!title) throw new Error('Betreff darf nicht leer sein.');
        if (title !== t.TITEL) ch.TITEL = title;
        if (sel.value !== t.STATUS) ch.STATUS = sel.value;
        if (newUser && newUser !== t.BEARBEITER_USER_ID) ch.BEARBEITER_USER_ID = newUser;
        if (pSel.value && Number(pSel.value) !== Number(t.PRIORITY)) ch.PRIORITY = Number(pSel.value);
        if (qSel.value && String(qSel.value) !== String(t.ERRQ_ID)) {
          ch.ERRQ_ID = Number(qSel.value);
          const qn = (QUEUES || []).find(q => String(q.id) === String(qSel.value)); if (qn) ch.MDERRQ_FK__NAME_KB = qn.name;
        }
        const par = $('hParent', head).value.trim();
        if (par !== String(t.MDERRDA_SID || '')) {
          if (par && !/^\d+$/.test(par)) throw new Error('Parent: bitte eine Ticket-ID eingeben.');
          if (par && Number(par) === Number(t.SID)) throw new Error('Ein Ticket kann nicht sein eigenes Parent sein.');
          if (par && !(await ticketBySid(par))) throw new Error(`Parent-Ticket ${par} nicht gefunden.`);
          ch.MDERRDA_SID = par ? Number(par) : null;
        }
        const nn = $('hNoNotify', head).checked ? 1 : 0;
        if (nn !== (t.SUPPRESS_NOTIFY == 1 ? 1 : 0)) ch.SUPPRESS_NOTIFY = nn;
        if (!Object.keys(ch).length) { msg('tMsg', 'Keine Änderungen.'); return; }
        await updateTicket(t, ch);
        await reloadTicket();
        msg('tMsg', '✔ Ticket gespeichert.', 'ok');
      } catch (e) { msg('tMsg', '✖ ' + e.message, 'err'); }
    }));
    confirmBtn($('hQC', head), 'Wirklich schließen?', async () => {
      try { await quickClose(t); await reloadTicket(); msg('tMsg', '✔ Per QuickClose geschlossen.', 'ok'); }
      catch (e) { msg('tMsg', '✖ ' + e.message, 'err'); }
    });

    // Firma ändern
    $('hFirmBtn', head).addEventListener('click', () => { const b = $('hFirmBox', head); b.style.display = b.style.display ? '' : 'none'; $('hFirmSearch', head).focus(); });
    let ft;
    $('hFirmSearch', head).addEventListener('input', () => {
      clearTimeout(ft);
      const q = $('hFirmSearch', head).value.trim(), box = $('hFirmRes', head);
      if (q.length < 3) { box.style.display = 'none'; return; }
      ft = setTimeout(async () => {
        try {
          const res = await searchAddresses(q);
          box.innerHTML = ''; box.style.display = '';
          for (const a of res) {
            const d = el('div'); d.textContent = (a.IMAGE_NR === 36 ? '🏢 ' : '👤 ') + a.MATCH;
            d.addEventListener('click', () => {
              box.style.display = 'none';
              const m = $('hFirmMsg', head); m.className = 'msg'; m.innerHTML = '';
              m.appendChild(document.createTextNode(`Firma ändern auf „${a.MATCH}“? Abrechenbare Stunden können dabei zurückgesetzt werden. `));
              const b = el('button', { className: 'btn sm warn', textContent: 'Firma ändern' }); m.appendChild(b);
              confirmBtn(b, 'Wirklich ändern?', async () => {
                try { await changeFirm(t, a); await reloadTicket(); msg('tMsg', `✔ Firma geändert auf ${a.MATCH}. Bitte abrechenbare Stunden prüfen.`, 'ok'); }
                catch (e) { m.textContent = '✖ ' + e.message; m.className = 'msg err'; toast(e.message); }
              });
            });
            box.appendChild(d);
          }
          if (!res.length) box.innerHTML = '<div><small>Keine Treffer</small></div>';
        } catch (e) { $('hFirmMsg', head).textContent = '✖ ' + e.message; toast(e.message); }
      }, 300);
    });

    // --- Anlagetext ---
    const anl = el('details', { className: 'card' });
    anl.innerHTML = `<summary>📄 Anlagetext</summary><div class="txt"></div>`;
    persistDetails(anl, 'ticket.anlagetext', true);
    anl.querySelector('.txt').innerHTML = t.ANLAGE_NOTIZ ? sanitize(t.ANLAGE_NOTIZ) : '<span class="hint">– kein Anlagetext –</span>';
    if (t.ANLAGE_NOTIZ && t.EMAIL_TO) {
      const bar = el('div', { className: 'row' }); bar.style.marginTop = '6px';
      const srcA = { inbound: true, from: t.EMAIL_TO, cc: t.EMAIL_CC, date: t.ANLAGE_DATUM, subject: t.TITEL, html: t.ANLAGE_NOTIZ };
      for (const [m, l] of [['one', '↩ Antworten'], ['all', '↩↩ Allen antworten'], ['fwd', '↪ Weiterleiten']]) {
        const b = el('button', { className: 'btn sm sec', textContent: l }); b.addEventListener('click', () => ui.reply(srcA, m)); bar.appendChild(b);
      }
      anl.appendChild(bar);
    }
    view.appendChild(anl);

    // --- Anhänge ---
    const docs = ui.docs || [];
    const dd = el('details', { className: 'card' });
    dd.innerHTML = `<summary>📎 Anhänge (${docs.length})</summary><div data-ref="dList"></div><div class="msg" data-ref="dMsg"></div>`;
    persistDetails(dd, 'ticket.anhaenge', true);
    for (const d of docs) {
      const row = el('div', { className: 'doc' });
      const ext = fileExt(d.DOCUMENT_NAME), isOffice = OFFICE_EXT.has(ext), kind = previewKind(ext);
      const B = (a, label, title = '') => `<button class="btn sm sec" data-a="${a}"${title ? ` title="${title}"` : ''}>${label}</button>`;
      const btns = ext === 'eml' ? B('open', '✉ Anzeigen') + B('dl', '⬇ Herunterladen')
        : isOffice ? (kind ? B('preview', '👁 Vorschau', 'Im Browser ansehen, ohne Download') : '') + B('office', '🖥 In Office öffnen', 'Originaldatei herunterladen und im installierten Office öffnen')
        : kind ? B('preview', '👁 Anzeigen', 'Im Browser ansehen, ohne Download') + B('dl', '⬇ Herunterladen')
        : B('open', 'Öffnen') + B('dl', '⬇ Herunterladen');
      row.innerHTML = `<span class="nm"></span><span class="hint">${fmtDateTime(d.DATUM)} · ${esc(d.CREATED_BY || '')}${d.TITEL && d.TITEL !== t.TITEL ? ' · ' + esc(d.TITEL) : ''}</span>${btns}`;
      row.querySelector('.nm').textContent = '📄 ' + (d.DOCUMENT_NAME || '(ohne Namen)');
      for (const b of row.querySelectorAll('button')) b.addEventListener('click', async () => {
        const win = b.dataset.a === 'open' || b.dataset.a === 'preview' ? window.open('about:blank', '_blank') : null; // synchron öffnen → kein Popup-Blocker
        try {
          if (b.dataset.a === 'preview') {
            if (!win) throw new Error('Das Vorschaufenster wurde vom Browser blockiert – Pop-ups für diese Seite erlauben.');
            await openPreview(d, win); log('Anhang angezeigt', `#${t.SID}: ${d.DOCUMENT_NAME}`); return;
          }
          if (b.dataset.a === 'office') {
            await busy(b, () => openLocal(d)); log('Anhang in Office geöffnet', `#${t.SID}: ${d.DOCUMENT_NAME}`);
            if (officeHintNeeded(ext)) showOfficeHint(dd, ext);
          }
          else if (win && ext === 'eml') { await openEml(d, win); log('Anhang angezeigt', `#${t.SID}: ${d.DOCUMENT_NAME}`); }
          else { await openDocument(d, b.dataset.a === 'dl', win); log(b.dataset.a === 'dl' ? 'Anhang heruntergeladen' : 'Anhang geöffnet', `#${t.SID}: ${d.DOCUMENT_NAME}`); }
        } catch (e) {
          if (win && (ext === 'eml' || b.dataset.a === 'preview')) { if (!e.quiet && !win.closed) winMessage(win, '✖ ' + e.message); } else if (win) win.close();
          if (!e.quiet) msg('dMsg', '✖ ' + e.message, 'err');
        }
      });
      $('dList', dd).appendChild(row);
    }
    if (!docs.length) $('dList', dd).innerHTML = '<span class="hint">Keine Anhänge.</span>';
    view.appendChild(dd);

    // --- Child-Tickets ---
    const kids = ui.children || [];
    const kd = el('div', { className: 'card' });
    kd.style.display = 'none'; kd.style.background = 'var(--mq-bg)';
    kd.innerHTML = `<b>🌿 Child-Tickets (${kids.length})</b><div data-ref="kList"></div>
      <div class="row" style="margin-top:6px"><div class="c s" style="flex-basis:200px"><input type="text" data-ref="kAdd" placeholder="Ticket-ID als Child zuordnen"></div>
      <button class="btn sm sec" data-ref="kAddBtn">zuordnen</button><span class="hint" style="flex:1"></span>
      <div class="c s" style="flex-basis:260px"><input type="text" data-ref="kTitle" placeholder="Betreff für neues Child (optional)"></div><button class="btn sm" data-ref="kNew">＋ Child-Ticket erstellen</button></div>
      <div class="msg" data-ref="kMsg"></div>`;
    for (const k of kids) {
      const r = el('div', { className: 'doc' });
      r.innerHTML = `<span class="nm"></span><span class="hint">${esc(statusInfo(k.STATUS)[1])} · ${esc(k.BEARBEITER_USER_ID || '–')} · ${fmtH(k.SUM_STD)} h</span><a class="btn sm sec" data-a="open">Öffnen</a><button class="btn sm sec" data-a="rm">lösen</button>`;
      r.querySelector('.nm').textContent = `#${k.SID} ${k.TITEL}`;
      wbLink(r.querySelector('[data-a=open]'), ticketRoute(k.SID), () => openTicket(k.SID));
      confirmBtn(r.querySelector('[data-a=rm]'), 'Wirklich lösen?', async () => {
        try { await updateTicket(k, { MDERRDA_SID: null }); await reloadTicket(); msg('tMsg', `✔ #${k.SID} ist kein Child mehr.`, 'ok'); }
        catch (e) { msg('kMsg', '✖ ' + e.message, 'err'); }
      });
      $('kList', kd).appendChild(r);
    }
    if (!kids.length) $('kList', kd).innerHTML = '<span class="hint">Keine Child-Tickets.</span>';
    $('kAddBtn', kd).addEventListener('click', ev => busy(ev.target, async () => {
      try {
        const id = $('kAdd', kd).value.trim(); if (!/^\d+$/.test(id)) throw new Error('Bitte eine Ticket-ID eingeben.');
        if (Number(id) === Number(t.SID)) throw new Error('Ein Ticket kann nicht sein eigenes Child sein.');
        const k = await ticketBySid(id); if (!k) throw new Error(`Ticket ${id} nicht gefunden.`);
        await updateTicket(k, { MDERRDA_SID: Number(t.SID) }); await reloadTicket(); msg('tMsg', `✔ #${id} ist jetzt Child von #${t.SID}.`, 'ok');
      } catch (e) { msg('kMsg', '✖ ' + e.message, 'err'); }
    }));
    confirmBtn($('kNew', kd), 'Wirklich anlegen?', async () => {
      try {
        const sid = await createChildTicket(t);
        const title = $('kTitle', kd).value.trim();
        if (title) { const k = await ticketBySid(sid); if (k && k.TITEL !== title) await updateTicket(k, { TITEL: title }); }
        msg('tMsg', `✔ Child-Ticket #${sid} angelegt.`, 'ok');
        await openTicket(sid, false);
      } catch (e) { msg('kMsg', '✖ ' + e.message, 'err'); }
    });
    $('hExtra', head).appendChild(kd);

    // --- Zusammenführen ---
    const mg = el('div', { className: 'card' });
    mg.style.display = 'none'; mg.style.background = 'var(--mq-bg)';
    mg.innerHTML = `<b>🔀 In anderes Ticket zusammenführen</b>
      <div class="hint" style="margin-top:4px">Dieses Ticket (#${t.SID}) wird als Kommentar in das Zielticket übernommen.</div>
      <div class="row" style="margin-top:6px"><div class="c s" style="flex-basis:200px"><input type="text" data-ref="mTarget" placeholder="Ziel-Ticket-ID"></div>
      <span class="hint" data-ref="mInfo" style="flex:1"></span><button class="btn sm warn" data-ref="mGo">Zusammenführen</button></div>
      <div class="msg" data-ref="mMsg"></div>`;
    let mergeTarget = null;
    $('mTarget', mg).addEventListener('change', async () => {
      const id = $('mTarget', mg).value.trim(); mergeTarget = null; $('mInfo', mg).textContent = '';
      if (!/^\d+$/.test(id)) return;
      const z = await ticketBySid(id).catch(() => null);
      if (!z) { $('mInfo', mg).textContent = 'nicht gefunden'; return; }
      mergeTarget = z; $('mInfo', mg).textContent = `→ #${z.SID} ${z.TITEL} · ${(z.MATCHCODE || '').replace(/^\d+\s*-\s*/, '')} · ${statusInfo(z.STATUS)[1]}`;
    });
    confirmBtn($('mGo', mg), () => { if (!mergeTarget) throw new Error('Bitte zuerst ein gültiges Zielticket eingeben (Feld verlassen).'); return `Wirklich #${t.SID} → #${mergeTarget.SID}?`; }, async () => {
      try { await mergeTickets(t, mergeTarget); msg('tMsg', `✔ #${t.SID} wurde in #${mergeTarget.SID} zusammengeführt.`, 'ok'); await openTicket(mergeTarget.SID, false); }
      catch (e) { msg('mMsg', '✖ ' + e.message, 'err'); }
    }, e => msg('mMsg', '✖ ' + e.message, 'err'));
    $('hExtra', head).appendChild(mg);
    // Ein-/Ausblenden über die Knöpfe im Kopf (selten gebraucht → standardmäßig zu)
    const kb = $('hKids', head), mb = $('hMerge', head);
    if (kids.length) kb.textContent = `🌿 Child-Tickets (${kids.length})`;
    const panelState = () => { secSet('ticket.panel', kd.style.display !== 'none' ? 'kids' : mg.style.display !== 'none' ? 'merge' : ''); };
    const togglePanel = (show, hide) => { show.style.display = show.style.display === 'none' ? '' : 'none'; hide.style.display = 'none'; panelState(); };
    kb.addEventListener('click', () => togglePanel(kd, mg));
    mb.addEventListener('click', () => togglePanel(mg, kd));
    if (secGet('ticket.panel', '') === 'kids') kd.style.display = ''; else if (secGet('ticket.panel', '') === 'merge') mg.style.display = '';

    // --- Neuer Kommentar / Zeit / Mail ---
    const add = el('details', { className: 'card' });
    add.innerHTML = `<summary>＋ Kommentar / Zeit / Mail hinzufügen</summary>`;
    persistDetails(add, 'ticket.neuerKommentar', false);
    const form = timeForm({
      billableDefault: billableDefault({ kndId: t.KND_ID, name: t.MATCHCODE }), allowZero: true, kontaktarten: ka,
      allowMail: true, allowFiles: true, kndId: t.KND_ID,
      mailDefaults: { to: t.EMAIL_TO, cc: t.EMAIL_CC, more: ui.comments.flatMap(c => [...splitMails(c.EMAIL_TO), ...splitMails(c.EMAIL_CC)]) }, suppressNotify: t.SUPPRESS_NOTIFY == 1,
      onRemoveSuppress: async () => { await updateTicket(t, { SUPPRESS_NOTIFY: 0 }); t.SUPPRESS_NOTIFY = 0; $('hNoNotify', head).checked = false; },
      ticketDocs: ui.docs || [],
      onChange: st => st.empty && !st.mailOn ? drafts.del('ticket:' + t.SID)
        : drafts.set('ticket:' + t.SID, { kind: 'ticket', sid: t.SID, title: t.TITEL, cust: t.MATCHCODE, form: st }),
    });
    add.appendChild(form.el);
    // Entwurf wiederherstellen
    const dr = drafts.get('ticket:' + t.SID);
    if (dr && dr.form) {
      form.restore(dr.form); add.open = true;
      const note = el('div', { className: 'warnbox' }); note.textContent = `Entwurf vom ${new Date(dr.savedAt).toLocaleString('de-DE').slice(0, 16)} wiederhergestellt (Anhänge müssen neu hinzugefügt werden). `;
      const x = el('a', { href: '#', textContent: 'Entwurf verwerfen' }); x.addEventListener('click', e => { e.preventDefault(); drafts.del('ticket:' + t.SID); form.reset(); note.remove(); });
      note.appendChild(x); add.insertBefore(note, form.el);
    }
    // Antworten / Allen antworten / Weiterleiten (aus Kommentar oder Anlagetext)
    // Ungespeicherte Eingaben? Dann bei Änderungen von außen nicht automatisch neu laden, sondern Hinweis zeigen.
    ui.form = form; ui.addBox = add; ui.draftSeen = (drafts.get('ticket:' + t.SID) || {}).savedAt || null;
    ui.isDirty = () => ui.editing.size > 0 || !form.state().empty || $('hTitle', head).value.trim() !== (t.TITEL || '') || sel.value !== t.STATUS;
    ui.reply = (src, mode) => {
      const inbound = src.inbound;
      const from = inbound ? src.from : (src.fromLabel || userLabel(src.user));
      const quote = `<p><br></p>${form.signatureHtml()}<hr style="border:none;border-top:1px solid #ccc;margin:10px 0 4px">` +
        `<div style="font-size:9pt;color:#444"><b>Von:</b> ${esc(from || '–')}<br><b>Gesendet:</b> ${esc(fmtDateTime(src.date))}<br>` +
        `${src.to && !inbound ? `<b>An:</b> ${esc(src.to)}<br>` : ''}${src.cc ? `<b>Cc:</b> ${esc(src.cc)}<br>` : ''}<b>Betreff:</b> ${esc(src.subject || t.TITEL)}</div>` +
        `<blockquote style="margin:6px 0 0 0;padding-left:8px;border-left:2px solid #ccc">${sanitize(src.html)}</blockquote>`;
      const own = new Set(); // eigene Adressen nicht erneut anschreiben
      let to = [], cc = [];
      if (mode !== 'fwd') {
        to = inbound ? splitMails(src.from) : splitMails(src.to);
        if (mode === 'all') cc = [...(inbound ? splitMails(src.to2) : []), ...splitMails(src.cc)].filter(m => !to.includes(m) && !own.has(m));
      }
      add.open = true;
      form.setHtml(quote); form.setMail({ to, cc });
      add.scrollIntoView({ behavior: 'smooth', block: 'start' });
      setTimeout(() => form.focusText(), 300);
      msg('tMsg', mode === 'fwd' ? 'Weiterleiten: Empfänger eintragen, ggf. Ticket-Anhänge hinzufügen.' : 'Antwort vorbereitet – Text oben eingeben.');
    };
    const addBtn = el('button', { className: 'btn', textContent: 'Hinzufügen' }); addBtn.style.marginTop = '8px';
    form.onMailToggle = on => { addBtn.textContent = on ? '✉ Mail senden' : 'Hinzufügen'; addBtn.style.background = on ? '#1e88e5' : ''; };
    add.appendChild(addBtn);
    let pending = null;
    confirmBtn(addBtn, () => {
      pending = form.read();
      const v = pending;
      const anh = v.files.length ? `, ${v.files.length} ${v.files.length === 1 ? 'Anhang' : 'Anhänge'}` : '';
      if (!v.mail) return `Wirklich hinzufügen? (${v.minutes} min${anh})`;
      const n = v.mail.to.length + v.mail.cc.length + v.mail.bcc.length;
      return `✉ Wirklich an ${n} Empfänger senden? (${[...v.mail.to, ...v.mail.cc].slice(0, 2).join(', ')}${n > 2 ? ' …' : ''}${anh})`;
    }, async () => {
      try {
        const v = pending || form.read();
        msg('tMsg', v.mail ? 'Sende …' : 'Speichere …');
        const r = await addComment(t, v);
        drafts.del('ticket:' + t.SID);
        form.reset();
        await reloadTicket();
        msg('tMsg', v.mail
          ? `✔ Mail ${r.emailDate ? 'versendet (' + fmtDateTime(r.emailDate) + ')' : 'übergeben – Versandzeit noch nicht gesetzt, bitte in MultiData prüfen'} an ${v.mail.to.join(', ')}.`
          : `✔ Kommentar hinzugefügt (${v.minutes} min${v.billable && v.minutes ? ', abrechenbar' : ''}).`, 'ok');
      } catch (e) {
        msg('tMsg', '✖ ' + e.message, 'err');
        if (e.code === 'SUPPRESS') { form.setSuppress(true); $('hNoNotify', head).checked = true; } // Entwurf bleibt erhalten
      }
    }, e => msg('tMsg', '✖ ' + e.message, 'err'), 6000);
    view.appendChild(add);

    // --- Kommentare ---
    const h = el('div', { className: 'row' }); h.style.marginTop = '12px';
    h.innerHTML = `<b style="flex:1">Kommentare (${ui.comments.length})</b>
      <button class="btn sm sec" data-ref="cAllOpen">▾ alle aufklappen</button><button class="btn sm sec" data-ref="cAllClose">▸ alle einklappen</button>`;
    view.appendChild(h);
    const setAll = open => { setAllComments(!open); view.querySelectorAll('.card[data-csid]').forEach(cd => cd.classList.toggle('collapsed', !open)); };
    $('cAllOpen', h).addEventListener('click', () => setAll(true));
    $('cAllClose', h).addEventListener('click', () => setAll(false));
    for (const c of ui.comments) view.appendChild(commentCard(c, ka));
  }

  function commentCard(c, ka) {
    const card = el('div', { className: 'card' + (commentCollapsed(c.SID) ? ' collapsed' : '') });
    card.dataset.csid = c.SID;
    const kaName = (ka.find(k => k[0] === c.LST_ID) || [c.LST_ID, c.LST_ID])[1];
    const von = (c.LFZ_VON || '').slice(11, 16), bis = (c.LFZ_BIS || '').slice(11, 16);
    card.innerHTML = `
      <div class="meta" data-ref="cHead" title="Klicken zum Ein-/Ausklappen">
        <span class="tog"></span><b>${fmtDate(c.LST_DATUM)}</b>${von ? `<span>${von}–${bis}</span>` : ''}
        <span>${esc(c.BEARBEITER_USER_ID || '')}</span><span class="tag">${esc(kaName || '')}</span>
        <span>Std. <b>${fmtH(c.LFZ_STD)}</b></span>
        <span class="tag ${c.INABR_STD ? 'abr' : ''}">abr. ${fmtH(c.INABR_STD)}</span>
        ${c.LFZ_STD_MAN == 1 ? '<span class="tag">manuell</span>' : ''}
        <span style="flex:1"></span>
        ${(c.EMAIL_TO || c.MEMO_IA == '1') ? `<button class="btn sm sec" data-r="one" title="Antworten">↩</button><button class="btn sm sec" data-r="all" title="Allen antworten">↩↩</button><button class="btn sm sec" data-r="fwd" title="Weiterleiten">↪</button>` : ''}
        <button class="btn sm sec" data-ref="cEdit">✎ Bearbeiten</button>
        <button class="btn sm warn" data-ref="cDel">🗑 Löschen</button>
      </div>
      <div class="mailinfo" data-ref="cMail" style="display:none"></div>
      <div class="txt"></div>`;
    card.querySelector('.txt').innerHTML = sanitize(c.NOTIZ);
    if (c.EMAIL_TO || c.EMAIL_DATUM || c.MEMO_IA == '1') {
      const mi = $('cMail', card); mi.style.display = '';
      mi.textContent = c.MEMO_IA == '1' ? `📥 Eingehende Mail von ${c.EMAIL_TO || '–'}${c.EMAIL_CC ? ' · Cc: ' + c.EMAIL_CC : ''}` :
        `✉ ${c.EMAIL_DATUM ? 'Mail ' + fmtDateTime(c.EMAIL_DATUM) : 'Mail (nicht versendet)'} · An: ${c.EMAIL_TO || '–'}` +
        `${c.EMAIL_CC ? ' · Cc: ' + c.EMAIL_CC : ''}${c.EMAIL_BCC ? ' · Bcc: ' + c.EMAIL_BCC : ''}`;
    }
    // Vorschau der ersten Zeile im eingeklappten Zustand
    const prev = el('span', { className: 'prev' }); prev.textContent = htmlToText(c.NOTIZ).slice(0, 120);
    $('cHead', card).insertBefore(prev, $('cHead', card).querySelector('span[style*="flex:1"]'));
    $('cHead', card).addEventListener('click', e => {
      if (e.target.closest('button')) return;
      const col = card.classList.toggle('collapsed');
      setCommentCollapsed(c.SID, col);
    });
    confirmBtn($('cDel', card), 'Wirklich löschen?', async () => {
      try { await deleteComment(c); await reloadTicket(); msg('tMsg', '✔ Kommentar gelöscht (Inhalt steht im Protokoll).', 'ok'); }
      catch (e) { msg('tMsg', '✖ ' + e.message, 'err'); }
    });
    $('cEdit', card).addEventListener('click', () => { ui.editing.add(c.SID); card.replaceWith(commentEditor(c, ka)); });
    card.querySelectorAll('[data-r]').forEach(b => b.addEventListener('click', () => {
      const inbound = c.MEMO_IA == '1';
      ui.reply && ui.reply({ inbound, from: inbound ? c.EMAIL_TO : null, user: c.BEARBEITER_USER_ID, to: c.EMAIL_TO, cc: c.EMAIL_CC,
        date: c.EMAIL_DATUM || c.LST_DATUM, subject: c.LST_TITEL, html: c.NOTIZ }, b.dataset.r);
    }));
    return card;
  }

  function commentEditor(c, ka) {
    const card = el('div', { className: 'card edit' });
    const date = (c.LFZ_VON || c.LST_DATUM || '').slice(0, 10);
    card.innerHTML = `
      <div class="row">
        <div class="c s" style="flex-basis:150px"><label class="l">Datum</label><input type="date" data-ref="eDate"></div>
        <div class="c s" style="flex-basis:95px"><label class="l">von</label><input type="time" data-ref="eFrom" step="900"></div>
        <div class="c s" style="flex-basis:95px"><label class="l">bis</label><input type="time" data-ref="eTo" step="900"></div>
        <div class="c s" style="flex-basis:170px"><label class="l">Kontaktart</label><select data-ref="eKa"></select></div>
      </div>
      <div class="row">
        <label class="chk"><input type="checkbox" data-ref="eMan"> Manuell (Stunden unabhängig von von/bis)</label>
        <div class="c s"><label class="l">Stunden</label><input type="text" data-ref="eStd"></div>
        <div class="c s"><label class="l">abr. Stunden</label><input type="text" data-ref="eAbr"></div>
        <div class="c"><div class="hint" data-ref="eHint"></div></div>
      </div>
      <label class="l">Text</label>
      <div class="rte" contenteditable="true" spellcheck="true" lang="de" data-ref="eTxt"></div>
      <div class="row" style="margin-top:8px">
        <button class="btn" data-ref="eSave">Speichern</button>
        <button class="btn sec" data-ref="eCancel">Abbrechen</button>
      </div>`;
    $('eDate', card).value = date;
    $('eFrom', card).value = (c.LFZ_VON || '').slice(11, 16);
    $('eTo', card).value = (c.LFZ_BIS || '').slice(11, 16);
    for (const [id, n] of ka) $('eKa', card).appendChild(el('option', { value: id, textContent: n }));
    $('eKa', card).value = c.LST_ID;
    $('eMan', card).checked = c.LFZ_STD_MAN == 1;
    $('eStd', card).value = c.LFZ_STD ?? '';
    $('eAbr', card).value = c.INABR_STD ?? '';
    $('eTxt', card).innerHTML = sanitize(c.NOTIZ);
    const origHtml = $('eTxt', card).innerHTML;
    const sync = () => {
      const man = $('eMan', card).checked;
      $('eStd', card).disabled = !man; $('eAbr', card).disabled = !man;
      $('eHint', card).textContent = man ? 'Stunden und abr. Stunden werden manuell gesetzt.' : 'Stunden ergeben sich aus von/bis; abr. Stunden nur über „Manuell“ änderbar.';
    };
    $('eMan', card).addEventListener('change', sync); sync();

    $('eCancel', card).addEventListener('click', () => { ui.editing.delete(c.SID); card.replaceWith(commentCard(c, ka)); });
    $('eSave', card).addEventListener('click', ev => busy(ev.target, async () => {
      try {
        const ch = {};
        const html = $('eTxt', card).innerHTML;
        if (html !== origHtml) ch.html = sanitize(html);
        if ($('eKa', card).value !== c.LST_ID) ch.kontaktart = $('eKa', card).value;
        const d = $('eDate', card).value, f = $('eFrom', card).value, t = $('eTo', card).value;
        if (d !== date || f !== (c.LFZ_VON || '').slice(11, 16) || t !== (c.LFZ_BIS || '').slice(11, 16)) {
          if (!d || !f || !t) throw new Error('Datum, von und bis ausfüllen.');
          if (toMin(t) < toMin(f)) throw new Error('„bis“ liegt vor „von“.');
          Object.assign(ch, { date: d, from: toMin(f), to: toMin(t) });
        }
        const man = $('eMan', card).checked;
        if (man) {
          ch.manual = true;
          ch.std = parseHours($('eStd', card).value, false);
          ch.abr = parseHours($('eAbr', card).value, false);
        } else if (c.LFZ_STD_MAN == 1) ch.manual = false;
        if (!Object.keys(ch).length) { msg('tMsg', 'Keine Änderungen.'); return; }
        if (!await editComment(c, ch)) { msg('tMsg', 'Keine Änderungen.'); return; }
        ui.editing.delete(c.SID);
        await reloadTicket();
        msg('tMsg', '✔ Kommentar gespeichert.', 'ok');
      } catch (e) { msg('tMsg', '✖ ' + e.message, 'err'); }
    }));
    return card;
  }

  // Wiederverwendbares Formular: Datum/Dauer/Kontaktart, Text-Editor, Abrechenbar, optional Mail & Anhänge
  function timeForm({ billableDefault: bd = false, allowZero = false, kontaktarten: ka = [], allowMail = false, allowFiles = false,
    mailDefaults = {}, kndId = null, suppressNotify = false, onRemoveSuppress = null, ticketDocs = [], onChange = null }) {
    const w = el('div');
    w.innerHTML = `
      <div class="row">
        <div class="c s" style="flex-basis:150px"><label class="l">Datum</label><input type="date" data-ref="fDate"></div>
        <div class="c s" style="flex-basis:150px"><label class="l">Dauer in h (0,25 · 1,5 · 1:10 · 20m${allowZero ? ' · 0 = ohne Zeit' : ''})</label><input type="text" data-ref="fDur"></div>
        <div class="c s" style="flex-basis:170px"><label class="l">Kontaktart</label><select data-ref="fKa"></select></div>
        <div class="c"><div class="chips" data-ref="fChips"></div></div>
      </div>
      <div class="hint" data-ref="fCalc"></div>
      <label class="chk"><input type="checkbox" data-ref="fBill"> Abrechenbar</label>
      <span data-ref="fAbrBox" style="display:none"><label class="chk">abr. Stunden (leer = wie gebucht) <input type="text" data-ref="fAbr" style="width:80px"></label></span>
      ${allowMail ? `<label class="chk" style="margin-left:12px"><input type="checkbox" data-ref="fMailOn"> <b>✉ Als Mail senden</b> („Antwort sofort senden“)</label>
      <div class="mailbox" data-ref="fMailBox" style="display:none">
        <div class="warnbox" data-ref="fSupp" style="display:none">Am Ticket ist „Keine Benachrichtigungen“ gesetzt – so geht keine Mail raus.
          <button class="btn sm" data-ref="fSuppBtn">Haken am Ticket entfernen</button></div>
        <div class="row">
          <div class="c" style="position:relative"><label class="l">An * (mehrere mit ; trennen)</label><input type="text" data-ref="fTo" autocomplete="off">
            <div class="results" data-ref="fToRes" style="display:none;position:absolute;left:0;right:0;z-index:5"></div></div>
          <div class="c" style="position:relative"><label class="l">Cc</label><input type="text" data-ref="fCc" autocomplete="off">
            <div class="results" data-ref="fCcRes" style="display:none;position:absolute;left:0;right:0;z-index:5"></div></div>
          <div class="c" style="position:relative"><label class="l">Bcc</label><input type="text" data-ref="fBcc" autocomplete="off">
            <div class="results" data-ref="fBccRes" style="display:none;position:absolute;left:0;right:0;z-index:5"></div></div>
        </div>
        <div class="hint" data-ref="fSugg"></div>
      </div>` : ''}
      <label class="l">Text <span class="hint">(Formatierung und eingefügte Bilder werden übernommen)</span></label>
      <div class="tb">
        <select data-ref="fFont" title="Schriftart">${CFG.fonts.map(f => `<option value="${f}"${f === CFG.defaultFont ? ' selected' : ''}>${f}</option>`).join('')}</select>
        <select data-ref="fSize" title="Schriftgröße">${CFG.fontSizes.map(z => `<option value="${z}"${z === CFG.defaultFontSize ? ' selected' : ''}>${z} pt</option>`).join('')}</select>
        <button type="button" data-cmd="bold" title="Fett"><b>F</b></button><button type="button" data-cmd="italic" title="Kursiv"><i>K</i></button>
        <button type="button" data-cmd="underline" title="Unterstrichen"><u>U</u></button>
        <button type="button" data-cmd="insertUnorderedList">• Liste</button><button type="button" data-cmd="insertOrderedList">1. Liste</button>
        <button type="button" data-ref="tblBtn" title="Tabelle einfügen">⊞ Tabelle</button>
        <button type="button" data-tbl="row+" title="Zeile unter der aktuellen Zelle einfügen">+Zeile</button><button type="button" data-tbl="col+" title="Spalte rechts einfügen">+Spalte</button>
        <button type="button" data-tbl="row-" title="Aktuelle Zeile löschen">−Zeile</button><button type="button" data-tbl="col-" title="Aktuelle Spalte löschen">−Spalte</button>
        <button type="button" data-cmd="removeFormat" title="Formatierung entfernen">✕ Format</button>
        <button type="button" data-ref="sigBtn" title="Signatur einfügen (⚙ Einstellungen)">✍ Signatur</button>
        <select data-ref="tbSel" title="Textbaustein einfügen"><option value="">📋 Textbaustein …</option></select>
      </div>
      <div class="tb" data-ref="tblPop" style="display:none;align-items:center">Tabelle: <input type="number" data-ref="tblR" value="3" min="1" max="50" style="width:50px"> Zeilen ×
        <input type="number" data-ref="tblC" value="3" min="1" max="20" style="width:50px"> Spalten <label class="chk" style="margin:0 6px"><input type="checkbox" data-ref="tblH" checked> Kopfzeile</label>
        <button type="button" data-ref="tblIns">Einfügen</button></div>
      <div class="rte" contenteditable="true" spellcheck="true" lang="de" data-ref="fText" style="font-family:${CFG.defaultFont};font-size:${CFG.defaultFontSize}pt"></div>
      ${allowFiles ? `<div class="drop" data-ref="fDrop">📎 Dateien hierher ziehen oder <u>auswählen</u><input type="file" multiple data-ref="fFile" style="display:none"></div>
      ${ticketDocs.length ? `<details style="margin-top:4px"><summary style="font-weight:400">📎 Anhänge aus dem Ticket hinzufügen (${ticketDocs.length})</summary><div data-ref="fDocs"></div>
        <button type="button" class="btn sm sec" data-ref="fDocsAdd" style="margin-top:4px">Ausgewählte hinzufügen</button></details>` : ''}
      <div class="files" data-ref="fFiles"></div>` : ''}`;
    $('fDate', w).value = todayStr();
    for (const [id, n] of (ka.length ? ka : [[CFG.defaultKontaktart, 'Support']])) $('fKa', w).appendChild(el('option', { value: id, textContent: n }));
    $('fKa', w).value = CFG.defaultKontaktart;
    $('fBill', w).checked = bd;
    const sync = () => { $('fAbrBox', w).style.display = $('fBill', w).checked ? '' : 'none'; };
    $('fBill', w).addEventListener('change', sync); sync();
    const calc = () => {
      const raw = parseMinutes($('fDur', w).value);
      if (isNaN(raw)) { $('fCalc', w).textContent = ''; return null; }
      const m = roundUp(raw);
      $('fCalc', w).textContent = m === 0 ? '→ ohne Zeit' : `→ ${m} min = ${fmtH(m / 60)} h${m !== raw ? ` (aufgerundet von ${raw} min)` : ''}`;
      return m;
    };
    $('fDur', w).addEventListener('input', calc);
    for (const m of [15, 30, 45, 60, 90, 120, 240]) {
      const hv = String(m / 60).replace('.', ',');
      const c = el('button', { className: 'chip', textContent: `${hv} h` });
      c.addEventListener('click', () => { $('fDur', w).value = hv; calc(); });
      $('fChips', w).appendChild(c);
    }

    // Text-Editor: Toolbar + Bilder aus der Zwischenablage als eingebettete Bilder (wie MultiData)
    const rte = $('fText', w);
    // Auswahl im Editor merken (Toolbar-Felder nehmen den Fokus weg)
    let savedRange = null;
    const saveSel = () => { const sel = getSelection(); if (sel.rangeCount && rte.contains(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange(); };
    const restoreSel = () => { rte.focus(); if (savedRange) { const sel = getSelection(); sel.removeAllRanges(); sel.addRange(savedRange); } };
    ['keyup', 'mouseup', 'input'].forEach(t => rte.addEventListener(t, saveSel));
    document.addEventListener('selectionchange', () => { if (rte.contains(getSelection().anchorNode)) saveSel(); });
    const cmd = (c, v = null) => { restoreSel(); document.execCommand('styleWithCSS', false, true); document.execCommand(c, false, v); saveSel(); };
    w.querySelectorAll('.tb button[data-cmd]').forEach(b => b.addEventListener('mousedown', e => { e.preventDefault(); cmd(b.dataset.cmd); }));
    $('fFont', w).addEventListener('change', e => cmd('fontName', e.target.value));
    $('fSize', w).addEventListener('change', e => {
      // execCommand kennt nur Stufen 1–7 → Stufe 7 setzen und durch echte pt-Größe ersetzen
      cmd('fontSize', '7');
      rte.querySelectorAll('span[style*="xxx-large"],font[size="7"]').forEach(n => { n.removeAttribute('size'); n.style.fontSize = e.target.value + 'pt'; });
    });
    // Tabellen
    const cellStyle = 'border:1px solid #999;padding:4px 8px;min-width:40px;vertical-align:top';
    $('tblBtn', w).addEventListener('mousedown', e => { e.preventDefault(); const p = $('tblPop', w); p.style.display = p.style.display === 'none' ? 'flex' : 'none'; });
    $('tblIns', w).addEventListener('click', () => {
      const r = Math.max(1, Math.min(50, +$('tblR', w).value || 3)), c = Math.max(1, Math.min(20, +$('tblC', w).value || 3)), head = $('tblH', w).checked;
      let h = '<table style="border-collapse:collapse;margin:4px 0"><tbody>';
      for (let i = 0; i < r; i++) {
        h += '<tr>' + Array.from({ length: c }, () => head && i === 0
          ? `<th style="${cellStyle};background:#f0f0f0;text-align:left">&nbsp;</th>` : `<td style="${cellStyle}">&nbsp;</td>`).join('') + '</tr>';
      }
      h += '</tbody></table><p><br></p>';
      cmd('insertHTML', h); $('tblPop', w).style.display = 'none';
    });
    w.querySelectorAll('.tb button[data-tbl]').forEach(b => b.addEventListener('mousedown', e => {
      e.preventDefault();
      const node = savedRange && savedRange.startContainer;
      const cell = node && (node.nodeType === 1 ? node : node.parentElement).closest('td,th');
      if (!cell || !rte.contains(cell)) return;
      const tr = cell.parentElement, table = tr.closest('table'), ci = [...tr.children].indexOf(cell);
      const newCell = () => { const td = document.createElement('td'); td.setAttribute('style', cellStyle); td.innerHTML = '&nbsp;'; return td; };
      if (b.dataset.tbl === 'row+') { const nr = document.createElement('tr'); for (let i = 0; i < tr.children.length; i++) nr.appendChild(newCell()); tr.after(nr); }
      if (b.dataset.tbl === 'col+') table.querySelectorAll('tr').forEach(row => { const ref = row.children[ci]; const n = ref && ref.tagName === 'TH' ? Object.assign(document.createElement('th'), { innerHTML: '&nbsp;' }) : newCell(); if (n.tagName === 'TH') n.setAttribute('style', ref.getAttribute('style')); ref ? ref.after(n) : row.appendChild(n); });
      if (b.dataset.tbl === 'row-') { tr.remove(); if (!table.querySelector('tr')) table.remove(); }
      if (b.dataset.tbl === 'col-') { table.querySelectorAll('tr').forEach(row => row.children[ci] && row.children[ci].remove()); if (!table.querySelector('td,th')) table.remove(); }
    }));
    const insertImage = file => new Promise(res => {
      const fr = new FileReader();
      fr.onload = () => { rte.focus(); document.execCommand('insertHTML', false, `<img src="${fr.result}" alt="${esc(file.name || 'Bild')}">`); res(); };
      fr.readAsDataURL(file);
    });
    rte.addEventListener('paste', async e => {
      const items = [...(e.clipboardData?.items || [])];
      const imgs = items.filter(i => i.kind === 'file' && i.type.startsWith('image/')).map(i => i.getAsFile()).filter(Boolean);
      if (imgs.length) { e.preventDefault(); for (const f of imgs) await insertImage(f); return; }
      const html = e.clipboardData?.getData('text/html');
      if (html) { e.preventDefault(); document.execCommand('insertHTML', false, sanitize(html)); }
    });

    // Signatur & Textbausteine
    const sigHtml = () => { const sg = settings().signature; return sg ? `<div data-mdqt-sig="1">${sanitize(sg)}</div>` : ''; };
    const appendSignature = () => { if (!sigHtml() || rte.querySelector('[data-mdqt-sig]')) return; rte.insertAdjacentHTML('beforeend', '<p><br></p>' + sigHtml()); };
    $('sigBtn', w).addEventListener('mousedown', e => {
      e.preventDefault();
      if (!sigHtml()) { alertMsg('Noch keine Signatur hinterlegt – unter ⚙ Einstellungen anlegen.'); return; }
      cmd('insertHTML', sigHtml());
    });
    const tbSel = $('tbSel', w);
    (async () => {
      const own = settings().blocks || [];
      if (own.length) { const g = el('optgroup', { label: 'Eigene' }); own.forEach((b, i) => g.appendChild(el('option', { value: 'own:' + i, textContent: b.name }))); tbSel.appendChild(g); }
      if (settings().showMdBlocks) {
        const md = await mdTextBlocks();
        if (md.length) { const g = el('optgroup', { label: 'MultiData' }); md.forEach((b, i) => g.appendChild(el('option', { value: 'md:' + i, textContent: b.name + (b.group ? ` (${b.group})` : '') }))); tbSel.appendChild(g); }
      }
    })();
    tbSel.addEventListener('change', async () => {
      const v = tbSel.value; tbSel.value = ''; if (!v) return;
      const [src, i] = v.split(':');
      const b = src === 'own' ? (settings().blocks || [])[+i] : (await mdTextBlocks())[+i];
      if (b) cmd('insertHTML', sanitize(b.html));
    });
    const alertMsg = t => toast(t, /^\s*✖/.test(t) ? 'err' : 'warn');

    // Anhänge
    let files = [];
    const renderFiles = () => {
      const box = $('fFiles', w); if (!box) return; box.innerHTML = '';
      files.forEach((f, i) => {
        const s = el('span', { className: 'file' }); s.textContent = `${f.name} (${Math.max(1, Math.round(f.size / 1024))} KB)`;
        const x = el('b', { textContent: '×', title: 'entfernen' }); x.addEventListener('click', () => { files.splice(i, 1); renderFiles(); });
        s.appendChild(x); box.appendChild(s);
      });
    };
    if (allowFiles) {
      const drop = $('fDrop', w), inp = $('fFile', w);
      drop.addEventListener('click', () => inp.click());
      inp.addEventListener('change', () => { files.push(...inp.files); inp.value = ''; renderFiles(); });
      ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
      ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
      drop.addEventListener('drop', e => { files.push(...e.dataTransfer.files); renderFiles(); });
      if (ticketDocs.length) {
        const box = $('fDocs', w);
        ticketDocs.forEach((d, i) => { const l = el('label', { className: 'chk' }); l.innerHTML = `<input type="checkbox" value="${i}"> `; l.appendChild(document.createTextNode(d.DOCUMENT_NAME || '(ohne Namen)')); box.appendChild(l); });
        $('fDocsAdd', w).addEventListener('click', ev => busy(ev.target, async () => {
          const sel = [...box.querySelectorAll('input:checked')].map(c => ticketDocs[+c.value]);
          try { for (const d of sel) files.push(await docAsFile(d)); box.querySelectorAll('input').forEach(c => c.checked = false); renderFiles(); }
          catch (e) { alertMsg('✖ ' + e.message); }
        }));
      }
    }
    // Änderungen melden (für automatische Entwürfe)
    let chTimer;
    const changed = () => { if (!onChange) return; clearTimeout(chTimer); chTimer = setTimeout(() => onChange(api_.state()), 700); };
    rte.addEventListener('input', changed);
    w.addEventListener('change', changed);
    w.addEventListener('input', e => { if (e.target.tagName === 'INPUT') changed(); });

    // Mail
    const api_ = {
      el: w, onMailToggle: null,
      read() {
        const minutes = calc();
        if (minutes == null || (!allowZero && minutes === 0)) throw new Error('Bitte eine gültige Dauer eingeben.');
        if (minutes > CFG.maxHours * 60) throw new Error(`Dauer über ${CFG.maxHours} h – bitte prüfen.`);
        const date = $('fDate', w).value; if (!date) throw new Error('Bitte ein Datum wählen.');
        let html = sanitize(rte.innerHTML).trim();
        // nur reiner Text (keine Formatierung) → „- “-Zeilen werden zur Liste (wie bisher)
        if (!/<(?!\/?(div|p|br)\b)[a-z]/i.test(html) && !/style=/i.test(html)) html = textToHtml(rte.innerText);
        if (!rte.innerText.trim() && !/<(img|table)/i.test(html)) throw new Error('Bitte einen Text eingeben.');
        const billable = $('fBill', w).checked;
        let mail = null;
        if (allowMail && $('fMailOn', w).checked) {
          mail = { to: splitMails($('fTo', w).value), cc: splitMails($('fCc', w).value), bcc: splitMails($('fBcc', w).value) };
          if (!mail.to.length) throw new Error('Bitte mindestens einen Empfänger in „An“ eintragen.');
          const bad = [...mail.to, ...mail.cc, ...mail.bcc].filter(m => !validMail(m));
          if (bad.length) throw new Error('Ungültige Adresse(n): ' + bad.join(', '));
          if (suppressNotify) throw new Error('Am Ticket ist „Keine Benachrichtigungen“ gesetzt – erst den Haken entfernen.');
        }
        // Grundschrift (Standard Arial) für Mails mitgeben, damit der Empfänger sie so sieht
        if (mail) html = `<div style="font-family:${CFG.defaultFont}, sans-serif;font-size:${CFG.defaultFontSize}pt">${html}</div>`;
        return { date, minutes, html, kontaktart: $('fKa', w).value, billable, abrHours: billable ? parseHours($('fAbr', w).value) : null, mail, files: [...files] };
      },
      reset() {
        $('fDur', w).value = ''; rte.innerHTML = ''; $('fAbr', w).value = ''; files = []; renderFiles(); calc();
        if (allowMail) { $('fMailOn', w).checked = false; toggleMail(); }
      },
      setBillable(v) { $('fBill', w).checked = v; sync(); },
      // Zustand für Entwürfe (Dateien können nicht gespeichert werden)
      state() {
        return { date: $('fDate', w).value, dur: $('fDur', w).value, ka: $('fKa', w).value, bill: $('fBill', w).checked, abr: $('fAbr', w).value, html: rte.innerHTML,
          mailOn: allowMail ? $('fMailOn', w).checked : false, to: allowMail ? $('fTo', w).value : '', cc: allowMail ? $('fCc', w).value : '', bcc: allowMail ? $('fBcc', w).value : '',
          empty: !rte.innerText.trim() && !/<img|<table/i.test(rte.innerHTML) && !$('fDur', w).value };
      },
      restore(d) {
        if (!d) return;
        $('fDate', w).value = d.date || todayStr(); $('fDur', w).value = d.dur || ''; if (d.ka) $('fKa', w).value = d.ka;
        $('fBill', w).checked = !!d.bill; $('fAbr', w).value = d.abr || ''; sync(); calc();
        rte.innerHTML = sanitize(d.html || '');
        if (allowMail && d.mailOn) { $('fMailOn', w).checked = true; $('fTo', w).value = d.to || ''; $('fCc', w).value = d.cc || ''; $('fBcc', w).value = d.bcc || ''; toggleMail(true); }
      },
      setMail({ to = [], cc = [], bcc = [] }) {
        if (!allowMail) return;
        $('fMailOn', w).checked = true;
        $('fTo', w).value = to.join('; '); $('fCc', w).value = cc.join('; '); $('fBcc', w).value = bcc.join('; ');
        toggleMail(true);
      },
      setHtml(html) { rte.innerHTML = html; },
      setSuppress(v) { suppressNotify = !!v; if (allowMail) $('fSupp', w).style.display = v && $('fMailOn', w).checked ? '' : 'none'; },
      signatureHtml: () => sigHtml(),
      focusText() { rte.focus(); const r = document.createRange(); r.setStart(rte, 0); r.collapse(true); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); },
    };
    function toggleMail(noSignature = false) {
      const on = $('fMailOn', w).checked;
      if (on && noSignature !== true && settings().autoSignature) appendSignature();
      $('fMailBox', w).style.display = on ? '' : 'none';
      $('fKa', w).value = on ? ([...$('fKa', w).options].some(o => o.value === 'MAIL') ? 'MAIL' : $('fKa', w).value) : CFG.defaultKontaktart;
      $('fSupp', w).style.display = on && suppressNotify ? '' : 'none';
      if (on && !$('fTo', w).value && noSignature !== true) { $('fTo', w).value = splitMails(mailDefaults.to).join('; '); $('fCc', w).value = splitMails(mailDefaults.cc).join('; '); }
      if (on && kndId && !mailList) {
        mailList = emailSuggestions(kndId).then(list => {
          $('fSugg', w).textContent = list.length ? `${list.length} Kontakte der Firma als Vorschlag – ins Feld klicken oder tippen (auch nach „;“).` : '';
          return list;
        });
      }
      if (api_.onMailToggle) api_.onMailToggle(on);
    }
    let mailList = null;
    if (allowMail) {
      // Vorschläge: Kontakte der Firma + Adressen aus dem Ticket; bereits eingetragene werden ausgeblendet
      const extra = [...splitMails(mailDefaults.to), ...splitMails(mailDefaults.cc), ...(mailDefaults.more || [])];
      const source = input => async q => {
        const list = mailList ? await mailList : [];
        const have = new Set(['fTo', 'fCc', 'fBcc'].flatMap(r => splitMails($(r, w).value)).map(x => x.toLowerCase()));
        const all = [...extra.map(e => ({ email: e, label: e })), ...list];
        const seen = new Set();
        return all.filter(x => { const k = x.email.toLowerCase(); if (seen.has(k) || have.has(k)) return false; seen.add(k); return !q || k.includes(q) || x.label.toLowerCase().includes(q); })
          .map(x => ({ value: x.email, label: x.label === x.email ? x.email : `${x.label}` }));
      };
      for (const r of ['fTo', 'fCc', 'fBcc']) attachSuggest($(r, w), $(r + 'Res', w), source(r), { token: true });
      $('fMailOn', w).addEventListener('change', toggleMail);
      $('fSuppBtn', w).addEventListener('click', async ev => busy(ev.target, async () => {
        try { await onRemoveSuppress(); suppressNotify = false; $('fSupp', w).style.display = 'none'; }
        catch (e) { $('fSupp', w).textContent = '✖ ' + e.message; toast(e.message); }
      }));
    }
    return api_;
  }

  // ================= NEUES TICKET =================
  let newForm = null;
  async function ensureNewForm() {
    if (newForm) return;
    newForm = timeForm({ billableDefault: false, kontaktarten: await kontaktarten(), onChange: () => saveNewDraft() });
    $('nTimeBox').appendChild(newForm.el);
  }
  $('nWithTime').addEventListener('change', async () => { await ensureNewForm(); $('nTimeBox').style.display = $('nWithTime').checked ? '' : 'none'; saveNewDraft(); });
  function rememberCustomer(c) {
    const list = store.get(CFG.store.recent, []).filter(x => x.ROW_ID !== c.ROW_ID);
    list.unshift({ ROW_ID: c.ROW_ID, MATCH: c.MATCH });
    store.set(CFG.store.recent, list.slice(0, 12));
  }
  async function chooseFirm(c) {
    ui.firm = c; rememberCustomer(c);
    msg('nFirm', '🏢 ' + c.MATCH, 'ok');
    $('nResults').style.display = 'none'; $('nSearch').value = '';
    await ensureNewForm(); newForm.setBillable(billableDefault({ name: c.MATCH }));
    const knd = await kndForAddress(c);
    if (knd && ui.firm === c) newForm.setBillable(billableDefault({ kndId: knd, name: c.MATCH }));
    renderRecent(); saveNewDraft();
  }
  function renderRecent() {
    const box = $('nRecent'); box.innerHTML = '';
    for (const c of store.get(CFG.store.recent, [])) {
      const b = el('button', { className: 'chip' + (ui.firm && ui.firm.ROW_ID === c.ROW_ID ? ' sel' : ''), title: c.MATCH });
      b.textContent = c.MATCH.length > 34 ? c.MATCH.slice(0, 33) + '…' : c.MATCH;
      b.addEventListener('click', () => chooseFirm(c));
      box.appendChild(b);
    }
  }
  let nt;
  $('nSearch').addEventListener('input', () => {
    clearTimeout(nt);
    const q = $('nSearch').value.trim();
    if (q.length < 3) { $('nResults').style.display = 'none'; return; }
    nt = setTimeout(async () => {
      try {
        const res = await searchAddresses(q), box = $('nResults');
        box.innerHTML = ''; box.style.display = '';
        for (const c of res) { const d = el('div'); d.textContent = (c.IMAGE_NR === 36 ? '🏢 ' : '👤 ') + c.MATCH; d.addEventListener('click', () => chooseFirm(c)); box.appendChild(d); }
        if (!res.length) box.innerHTML = '<div><small>Keine Treffer</small></div>';
      } catch (e) { msg('nMsg', '✖ ' + e.message, 'err'); }
    }, 300);
  });
  confirmBtn($('nGo'), 'Wirklich anlegen?', async () => {
    try {
      if (!ui.firm) throw new Error('Bitte eine Firma wählen.');
      const title = $('nTitle').value.trim(); if (!title) throw new Error('Bitte einen Betreff eingeben.');
      const tv = $('nWithTime').checked ? newForm.read() : null;
      msg('nMsg', 'Lege Ticket an …');
      const t = await createTicket({ adrRowId: ui.firm.ROW_ID, title, html: textToHtml($('nDesc').value), suppressNotify: $('nNoNotify').checked });
      let out = `✔ Ticket #${t.SID} angelegt.`;
      drafts.del('new');
      if (tv) { await addComment(t, tv); out += `\n✔ ${tv.minutes} min gebucht${tv.billable ? ' (abrechenbar)' : ''}.`; newForm.reset(); }
      if ($('nClose').checked) { await quickClose(t); out += '\n✔ Per QuickClose geschlossen.'; }
      msg('nMsg', out, 'ok');
      const a = wbLink(el('a', { textContent: ' → in der Werkbank öffnen' }), ticketRoute(t.SID), () => openTicket(t.SID));
      $('nMsg').appendChild(a);
      $('nTitle').value = ''; $('nDesc').value = '';
    } catch (e) { msg('nMsg', '✖ ' + e.message, 'err'); }
  });

  // ================= INBOX =================
  ui.inboxTypes = new Set(Object.keys(INBOX_TYPES));
  function renderInboxTypes() {
    const box = $('iTypes'); box.innerHTML = '';
    for (const [k, n] of Object.entries(INBOX_TYPES)) {
      const c = el('button', { className: 'chip' + (ui.inboxTypes.has(k) ? ' sel' : ''), textContent: n });
      c.addEventListener('click', () => { ui.inboxTypes.has(k) ? ui.inboxTypes.delete(k) : ui.inboxTypes.add(k); renderInbox(); });
      box.appendChild(c);
    }
  }
  function inboxItems() {
    const st = inboxState(), days = +$('iRange').value, from = addDays(todayStr(), -(days - 1));
    return Object.values(st.items)
      .filter(i => ui.inboxTypes.has(i.type))
      .filter(i => !$('iUnread').checked || !i.read)
      .filter(i => !$('iCust').value || String(i.kndId) === $('iCust').value)
      .filter(i => (i.date || i.found || '').slice(0, 10) >= from)
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }
  function renderInbox() {
    renderInboxTypes();
    const st = inboxState();
    // Kundenauswahl aus vorhandenen Einträgen
    const cs = $('iCust'), cur = cs.value, custs = {};
    Object.values(st.items).forEach(i => { if (i.kndId) custs[i.kndId] = i.cust; });
    cs.innerHTML = '<option value="">Alle Kunden</option>';
    Object.entries(custs).sort((a, b) => String(a[1]).localeCompare(String(b[1]))).forEach(([id, n]) => cs.appendChild(el('option', { value: id, textContent: String(n || id).replace(/^\d+\s*-\s*/, '') })));
    cs.value = cur;
    $('iLast').textContent = st.last ? `zuletzt aktualisiert ${new Date(st.last).toLocaleTimeString('de-DE').slice(0, 5)}` : 'noch nicht geladen';
    const box = $('iList'); box.innerHTML = '';
    const items = inboxItems();
    for (const i of items) {
      const r = el('a', { className: 'ibx' + (i.read ? '' : ' unread') });
      r.innerHTML = `<div class="ty">${esc(INBOX_TYPES[i.type] || i.type)}</div><div class="bd"><div class="tt"></div><div class="sn"></div></div>
        <div class="dt">${esc(fmtDateTime(i.date))}</div><button class="btn sm sec" data-a="rd" title="${i.read ? 'als ungelesen markieren' : 'als gelesen markieren'}">${i.read ? '○' : '✓'}</button>`;
      r.querySelector('.tt').textContent = `${i.sid ? '#' + i.sid + ' ' : ''}${i.title || ''}${i.cust ? ' · ' + String(i.cust).replace(/^\d+\s*-\s*/, '') : ''}`;
      r.querySelector('.sn').textContent = `${i.from ? (i.type === 'comment' ? userLabel(i.from) : i.from) + ': ' : ''}${i.text || ''}`;
      r.querySelector('[data-a=rd]').addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); setInboxRead([i.key], !i.read); renderInbox(); updateBadges(); });
      const markRead = () => { setInboxRead([i.key], true); updateBadges(); };
      if (i.sid) {
        wbLink(r, ticketRoute(i.sid), () => { markRead(); openTicket(i.sid); });
        r.addEventListener('click', e => { if (e.ctrlKey || e.metaKey || e.shiftKey) { markRead(); setTimeout(renderInbox, 50); } });
        r.addEventListener('auxclick', e => { if (e.button === 1) { markRead(); setTimeout(renderInbox, 50); } });
      } else r.addEventListener('click', e => { e.preventDefault(); markRead(); renderInbox(); });
      box.appendChild(r);
    }
    if (!items.length) box.innerHTML = '<div class="hint" style="margin-top:10px">Keine Einträge für diese Filter.</div>';
  }
  ['iCust', 'iRange', 'iUnread'].forEach(r => $(r).addEventListener('change', renderInbox));
  $('iRefresh').addEventListener('click', ev => busy(ev.target, async () => {
    msg('iMsg', 'Aktualisiere …');
    try { await refreshInbox(); msg('iMsg', ''); renderInbox(); updateBadges(); } catch (e) { msg('iMsg', '✖ ' + e.message, 'err'); }
  }));
  $('iAllRead').addEventListener('click', () => { setInboxRead(inboxItems().map(i => i.key), true); renderInbox(); updateBadges(); });

  function updateBadges() {
    const st = inboxState();
    const unread = Object.values(st.items).filter(i => !i.read).length;
    const nd = Object.keys(drafts.all()).length;
    $('bInbox').textContent = unread || ''; $('bDrafts').textContent = nd || '';
    fab.querySelector('.badge').textContent = unread || '';
    updateTitle();
  }
  // Automatisch aktualisieren (auch bei geschlossener Werkbank, für den Zähler)
  async function inboxTick() {
    try { await refreshInbox(); updateBadges(); if (ui.tab === 'inbox' && panel.classList.contains('open')) renderInbox(); } catch { /* still */ }
  }
  updateBadges();
  // Nur EIN Tab ruft die Inbox ab (Web Locks); alle anderen bekommen die Ergebnisse über den gemeinsamen Speicher.
  // Wird dieser Tab geschlossen, übernimmt automatisch der nächste.
  const startInboxPolling = () => { setTimeout(inboxTick, 8000); setInterval(inboxTick, CFG.inboxIntervalMin * 60e3); };
  if (navigator.locks && navigator.locks.request) navigator.locks.request('mdqt-inbox-leader', () => { startInboxPolling(); return new Promise(() => {}); });
  else startInboxPolling();

  // ================= ENTWÜRFE =================
  function renderDrafts() {
    const box = $('dView'); box.innerHTML = '';
    const all = Object.values(drafts.all()).sort((a, b) => (b.savedAt || '').localeCompare(a.savedAt || ''));
    for (const d of all) {
      const c = el('div', { className: 'card' });
      const preview = htmlToText(d.form?.html || d.desc || '').slice(0, 200);
      c.innerHTML = `<div class="meta"><b></b><span>${esc(new Date(d.savedAt).toLocaleString('de-DE').slice(0, 16))}</span>
        ${d.form?.mailOn ? '<span class="tag abr">✉ Mail</span>' : ''}<span style="flex:1"></span>
        <a class="btn sm" data-a="open">Öffnen</a><button class="btn sm warn" data-a="del">Löschen</button></div><div class="txt hint"></div>`;
      c.querySelector('b').textContent = d.kind === 'new' ? `＋ Neues Ticket: ${d.title || '(ohne Betreff)'}${d.firm ? ' · ' + d.firm.MATCH : ''}` : `#${d.sid} ${d.title || ''}${d.cust ? ' · ' + String(d.cust).replace(/^\d+\s*-\s*/, '') : ''}`;
      c.querySelector('.txt').textContent = (d.form?.mailOn && d.form?.to ? `An: ${d.form.to} – ` : '') + preview;
      wbLink(c.querySelector('[data-a=open]'), d.kind === 'new' ? { tab: 'new' } : ticketRoute(d.sid), () => d.kind === 'new' ? showTab('new') : openTicket(d.sid));
      confirmBtn(c.querySelector('[data-a=del]'), 'Wirklich löschen?', async () => { drafts.del(d.key); renderDrafts(); });
      box.appendChild(c);
    }
    if (!all.length) box.innerHTML = '<div class="hint" style="margin-top:10px">Keine Entwürfe.</div>';
  }
  // Entwurf „Neues Ticket“
  let newDraftTimer;
  function saveNewDraft() {
    clearTimeout(newDraftTimer);
    newDraftTimer = setTimeout(() => {
      const title = $('nTitle').value, desc = $('nDesc').value, form = newForm && $('nWithTime').checked ? newForm.state() : null;
      if (!title.trim() && !desc.trim() && (!form || form.empty)) { drafts.del('new'); return; }
      drafts.set('new', { kind: 'new', title, desc, firm: ui.firm, withTime: $('nWithTime').checked, form });
    }, 700);
  }
  ['nTitle', 'nDesc'].forEach(r => $(r).addEventListener('input', saveNewDraft));
  let newDraftRestored = false;
  async function restoreNewDraft() {
    if (newDraftRestored) return; newDraftRestored = true;
    const d = drafts.get('new'); if (!d) return;
    if (d.firm) await chooseFirm(d.firm);
    $('nTitle').value = d.title || ''; $('nDesc').value = d.desc || '';
    if (d.withTime && d.form) { $('nWithTime').checked = true; await ensureNewForm(); $('nTimeBox').style.display = ''; newForm.restore(d.form); }
    msg('nMsg', `Entwurf vom ${new Date(d.savedAt).toLocaleString('de-DE').slice(0, 16)} wiederhergestellt.`);
  }

  // ================= EINSTELLUNGEN =================
  function renderSettings() {
    const st = settings();
    $('stTheme').value = st.theme || 'auto';
    $('stSig').innerHTML = sanitize(st.signature || '');
    $('stAuto').checked = st.autoSignature !== false;
    $('stMd').checked = st.showMdBlocks !== false;
    const box = $('stBlocks'); box.innerHTML = '';
    (st.blocks || []).forEach(b => addBlockEditor(b));
    if (!(st.blocks || []).length) box.innerHTML = '<div class="hint" data-empty="1">Noch keine eigenen Bausteine.</div>';
    renderTeam();
  }

  // --- Team-Konfiguration ---
  const custSource = async q => {
    if (!q || q.length < 2) return [];
    return (await searchCustomersByTickets(q).catch(() => [])).slice(0, 30)
      .map(c => ({ value: String(c.id), label: `${String(c.name).replace(/^\d+\s*-\s*/, '')} (${c.id})`, name: String(c.name).replace(/^\d+\s*-\s*/, '') }));
  };
  const toPick = list => list.map(c => ({ value: String(c.id), label: `${c.name} (${c.id})`, name: c.name }));
  const fromPick = items => items.map(it => ({ id: String(it.value), name: it.name || String(it.label).replace(/\s*\(\d+\)$/, '') }));
  let teamRendering = false;
  function saveTeam() {
    if (teamRendering) return;
    store.set(CFG.store.team, { ...store.get(CFG.store.team, {}), nonBillable: fromPick(nbPick.items()), intern: fromPick(inPick.items()), othersBillable: $('stOthers').checked });
    msg('stTeamMsg', '✔ Team-Konfiguration gespeichert.', 'ok');
    updateTeamHint();
  }
  const nbPick = multiPick($('stNb'), { placeholder: 'Kunde suchen (Name oder Nr.) …', source: custSource, onChange: saveTeam });
  const inPick = multiPick($('stIn'), { placeholder: 'Kunde suchen (Name oder Nr.) …', source: custSource, onChange: saveTeam });
  function renderOffice() {
    const a = officeAck();
    $('oStatus').textContent = a.all ? 'Einrichtungshinweise ausgeschaltet.'
      : a.ext.length ? `Eingerichtet für: ${a.ext.map(e => '.' + e).join(', ')}. Für weitere Typen erscheint beim ersten Öffnen ein Hinweis.`
      : 'Beim ersten Öffnen je Dateityp erscheint eine kurze Anleitung.';
    $('oReset').style.display = a.all || a.ext.length ? '' : 'none';
  }
  $('oReset').addEventListener('click', () => { store.set(OFFICE_ACK, { all: false, ext: [] }); renderOffice(); });
  function renderTeam() {
    renderOffice();
    const tc = team();
    teamRendering = true;
    nbPick.set(toPick(tc.nonBillable)); inPick.set(toPick(tc.intern));
    $('stOthers').checked = tc.othersBillable !== false;
    teamRendering = false;
  }
  $('stOthers').addEventListener('change', saveTeam);
  function updateTeamHint() {
    $('teamHint').style.display = !teamConfigured() && !store.get('mdqt.teamHintOff', false) ? '' : 'none';
  }
  $('teamHintGo').addEventListener('click', () => showTab('settings'));
  $('teamHintNo').addEventListener('click', () => { store.set('mdqt.teamHintOff', true); updateTeamHint(); });

  $('stExport').addEventListener('click', () => {
    const tc = team();
    const data = { type: 'mdqt-team', version: 1, exportedAt: new Date().toISOString(), nonBillable: tc.nonBillable, intern: tc.intern, othersBillable: tc.othersBillable !== false };
    if ($('stExpTpl').checked) data.templates = templates();
    download('werkbank-team-config.json', JSON.stringify(data, null, 2), 'application/json');
    msg('stTeamMsg', '✔ Exportiert.', 'ok');
  });
  $('stImport').addEventListener('click', () => $('stImportFile').click());
  $('stImportFile').addEventListener('change', async () => {
    const f = $('stImportFile').files[0]; $('stImportFile').value = '';
    if (!f) return;
    try {
      const d = JSON.parse(await f.text());
      if (!d || d.type !== 'mdqt-team') throw new Error('Keine Werkbank-Team-Konfiguration.');
      const clean = l => (Array.isArray(l) ? l : []).filter(c => c && c.id != null).map(c => ({ id: String(c.id), name: String(c.name || c.id) }));
      store.set(CFG.store.team, { nonBillable: clean(d.nonBillable), intern: clean(d.intern), othersBillable: d.othersBillable !== false });
      let extra = '';
      if (Array.isArray(d.templates) && d.templates.length) {
        const own = templates(), names = new Set(own.map(t => t.name));
        const add = d.templates.filter(t => t && t.name && t.fs && !names.has(t.name));
        store.set(CFG.store.templates, [...own, ...add]); renderTemplates();
        extra = `, ${add.length} Suchvorlage(n) ergänzt`;
      }
      renderTeam(); updateTeamHint();
      msg('stTeamMsg', `✔ Importiert: ${clean(d.nonBillable).length} nicht abrechenbar, ${clean(d.intern).length} intern${extra}.`, 'ok');
      log('Team-Konfiguration importiert', f.name);
    } catch (e) { msg('stTeamMsg', '✖ ' + e.message, 'err'); }
  });
  confirmBtn($('stWipe'), 'Wirklich alles löschen? Nochmal klicken', async () => {
    Object.keys(localStorage).filter(k => k.startsWith('mdqt.')).forEach(k => localStorage.removeItem(k));
    Object.keys(sessionStorage).filter(k => k.startsWith('mdqt.')).forEach(k => sessionStorage.removeItem(k));
    location.reload();
  });
  renderTeam(); updateTeamHint();
  function addBlockEditor(b = { name: '', html: '' }) {
    const box = $('stBlocks'); box.querySelector('[data-empty]')?.remove();
    const c = el('div', { className: 'card' });
    c.innerHTML = `<div class="row"><div class="c"><input type="text" data-a="name" placeholder="Name des Bausteins"></div><button class="btn sm warn" data-a="del">Entfernen</button></div>
      <div class="rte" contenteditable="true" spellcheck="true" lang="de" data-a="html" style="margin-top:4px;font-family:Arial;font-size:10pt"></div>`;
    c.querySelector('[data-a=name]').value = b.name || '';
    c.querySelector('[data-a=html]').innerHTML = sanitize(b.html || '');
    c.querySelector('[data-a=del]').addEventListener('click', () => c.remove());
    enablePaste(c.querySelector('[data-a=html]'));
    box.appendChild(c);
  }
  // Einfügen in einfache Editoren: Bilder als eingebettete Bilder, HTML entschärft
  function enablePaste(ed) {
    ed.addEventListener('paste', e => {
      const items = [...(e.clipboardData?.items || [])];
      const imgs = items.filter(i => i.kind === 'file' && i.type.startsWith('image/')).map(i => i.getAsFile()).filter(Boolean);
      if (imgs.length) {
        e.preventDefault();
        imgs.forEach(f => { const fr = new FileReader(); fr.onload = () => document.execCommand('insertHTML', false, `<img src="${fr.result}">`); fr.readAsDataURL(f); });
        return;
      }
      const html = e.clipboardData?.getData('text/html');
      if (html) { e.preventDefault(); document.execCommand('insertHTML', false, sanitize(html)); }
    });
  }
  enablePaste($('stSig'));
  $('stTheme').addEventListener('change', () => { saveSettings({ ...settings(), theme: $('stTheme').value }); applyTheme(); });

  // Dark Mode: automatisch nach System oder fest hell/dunkel
  const darkMq = window.matchMedia ? matchMedia('(prefers-color-scheme: dark)') : null;
  function applyTheme() {
    const mode = settings().theme || 'auto';
    const dark = mode === 'dark' || (mode === 'auto' && !!darkMq?.matches);
    panel.classList.toggle('dark', dark);
    $('theme').textContent = dark ? '☀' : '🌙';
  }
  darkMq?.addEventListener?.('change', applyTheme);
  $('theme').addEventListener('click', () => {
    saveSettings({ ...settings(), theme: panel.classList.contains('dark') ? 'light' : 'dark' }); applyTheme();
    if (ui.tab === 'settings') $('stTheme').value = settings().theme;
  });
  applyTheme();
  $('stAdd').addEventListener('click', () => addBlockEditor());
  $('stSave').addEventListener('click', () => {
    const blocks = [...$('stBlocks').querySelectorAll('.card')].map(c => ({ name: c.querySelector('[data-a=name]').value.trim(), html: sanitize(c.querySelector('[data-a=html]').innerHTML) }))
      .filter(b => b.name && (htmlToText(b.html) !== '' || /<img|<table/i.test(b.html)));
    saveSettings({ ...settings(), theme: $('stTheme').value, signature: sanitize($('stSig').innerHTML), autoSignature: $('stAuto').checked, showMdBlocks: $('stMd').checked, blocks });
    msg('stMsg', `✔ Gespeichert (${blocks.length} Baustein(e)). Gilt für neu geöffnete Formulare.`, 'ok');
  });

  // ================= ZEITEN =================
  // Nachbau und Erweiterung von „Berichte → Ticketstunden“: Tage → Intern/Kunde → Tickets, mit Farbcodierung,
  // Personio-Abgleich, Ausgleichsbuchung, Kundenübersicht und CSV-Export.
  const WD = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const weekday = d => new Date(d + 'T12:00:00').getDay();
  const isWorkday = d => { const w = weekday(d); return w >= 1 && w <= 5; };
  const personioRound = min => { const q = Math.floor(min / 15), r = min % 15; return (q + (r >= CFG.personioRoundMin ? 1 : 0)) * 0.25; };
  const personioStore = () => store.get(CFG.store.personio, {});

  async function loadTimes({ users, from, to }) {
    const meId = await me();
    const ids = [...new Set((users.length ? users : ['__ME__']).map(u => u === '__ME__' ? meId : u))];
    const cs = await load('MDERRDA_MDERRDAC', {
      take: 20000,
      filter: andAll([orFilter('BEARBEITER_USER_ID', ids), ['LST_DATUM', '>=', from + 'T00:00:00'], ['LST_DATUM', '<', addDays(to, 1) + 'T00:00:00']]),
      sort: [{ selector: 'LST_DATUM', desc: false }],
      sqlColumns: ['SID', 'MDERRDA_SID', 'BEARBEITER_USER_ID', 'LST_DATUM', 'LFZ_VON', 'LFZ_BIS', 'LFZ_STD', 'INABR_STD', 'LST_ID', 'NOTIZ'],
    });
    const book = cs.filter(c => Number(c.LFZ_STD) > 0);
    const sids = [...new Set(book.map(c => c.MDERRDA_SID))];
    const tickets = {};
    for (let i = 0; i < sids.length; i += 60) {
      for (const t of await load('MDERRDA', { take: 100, filter: orFilter('SID', sids.slice(i, i + 60)), sqlColumns: ['SID', 'TITEL', 'KND_ID', 'MATCHCODE', 'STATUS', 'PRIORITY'] })) tickets[t.SID] = t;
    }
    return { bookings: book, tickets, users: ids };
  }

  const tm = { data: null };
  const timesUserPick = multiPick($('zUsers'), {
    placeholder: 'Bearbeiter …',
    source: async q => [...(!q || 'ich'.includes(q) ? [{ value: '__ME__', label: '👤 Ich' }] : []), ...(await userSource()(q))],
  });
  timesUserPick.set([{ value: '__ME__', label: '👤 Ich' }]);
  persistDetails($('zPDet'), 'zeiten.personio', false);
  (function initTimes() {
    const z = store.get(CFG.store.timesView, {});
    const d = new Date(), first = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-01`;
    $('zFrom').value = z.from || first; $('zTo').value = z.to || todayStr();
    $('zColor').checked = z.color !== false; $('zWeekend').checked = z.weekend !== false; $('zDetails').checked = !!z.details;
    $('zLow').value = z.low ?? CFG.dayColors.low;
    $('zTicket').value = settings().timesTicket || '';
    const box = $('zRanges'), t = todayStr();
    const mon = s => addDays(s, -((weekday(s) + 6) % 7));
    const lastOfMonth = (y, m) => addDays(`${m === 11 ? y + 1 : y}-${pad((m + 1) % 12 + 1)}-01`, -1);
    const pm = d.getMonth() === 0 ? 11 : d.getMonth() - 1, py = d.getMonth() === 0 ? d.getFullYear() - 1 : d.getFullYear();
    for (const [n, a, b] of [['Diese Woche', mon(t), t], ['Letzte Woche', addDays(mon(t), -7), addDays(mon(t), -1)],
      ['Dieser Monat', first, t], ['Letzter Monat', `${py}-${pad(pm + 1)}-01`, lastOfMonth(py, pm)]]) {
      const c = el('button', { className: 'chip', textContent: n });
      c.addEventListener('click', () => { $('zFrom').value = a; $('zTo').value = b; runTimes(); });
      box.appendChild(c);
    }
  })();
  const saveTimesView = () => store.set(CFG.store.timesView, { from: $('zFrom').value, to: $('zTo').value, color: $('zColor').checked, weekend: $('zWeekend').checked, details: $('zDetails').checked, low: +$('zLow').value || CFG.dayColors.low });
  ['zColor', 'zWeekend', 'zDetails', 'zLow'].forEach(r => $(r).addEventListener('change', () => { saveTimesView(); renderTimes(); }));
  $('zTicket').addEventListener('change', () => saveSettings({ ...settings(), timesTicket: $('zTicket').value.trim() }));

  async function runTimes() {
    const from = $('zFrom').value, to = $('zTo').value;
    if (!from || !to || to < from) { msg('zMsg', 'Bitte einen gültigen Zeitraum wählen.', 'err'); return; }
    if ((new Date(to) - new Date(from)) / 864e5 > 186) { msg('zMsg', 'Zeitraum bitte auf max. 6 Monate begrenzen.', 'err'); return; }
    saveTimesView();
    msg('zMsg', 'Lade Zeiten …');
    try { tm.data = await loadTimes({ users: timesUserPick.get(), from, to }); tm.from = from; tm.to = to; renderTimes(); msg('zMsg', ''); }
    catch (e) { msg('zMsg', '✖ ' + e.message, 'err'); }
  }
  $('zGo').addEventListener('click', ev => busy(ev.target, runTimes));

  // Personio-CSV (Export „Arbeitszeit pro Tag“, Minuten) – per Drag & Drop oder Auswahl
  async function importPersonio(file) {
    const text = await file.text();
    const p = personioStore(); let n = 0;
    for (const line of text.split(/\r?\n/)) {
      const m = line.match(/(\d{2})\.(\d{2})\.(\d{4})\s*[;,\t]\s*(\d+)/);
      if (m) { p[`${m[3]}-${m[2]}-${m[1]}`] = +m[4]; n++; }
    }
    if (!n) throw new Error('Keine Zeilen im Format „TT.MM.JJJJ;Minuten“ gefunden.');
    store.set(CFG.store.personio, p);
    msg('zMsg', `✔ Personio: ${n} Tage übernommen.`, 'ok');
    renderTimes();
  }
  {
    const drop = $('zDrop'), inp = $('zFile');
    drop.addEventListener('click', () => inp.click());
    inp.addEventListener('change', () => { if (inp.files[0]) importPersonio(inp.files[0]).catch(e => msg('zMsg', '✖ ' + e.message, 'err')); inp.value = ''; });
    ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) importPersonio(f).catch(er => msg('zMsg', '✖ ' + er.message, 'err')); });
    $('zPClear').addEventListener('click', () => { store.set(CFG.store.personio, {}); renderTimes(); });
  }

  // Ausgleich: Differenz Personio − gebucht auf das Sammelticket buchen
  async function bookBalance(days) {
    const sid = $('zTicket').value.trim();
    if (!/^\d+$/.test(sid)) throw new Error('Bitte zuerst ein Ticket für den Ausgleich eintragen.');
    const t = await ticketBySid(sid); if (!t) throw new Error(`Ticket ${sid} nicht gefunden.`);
    for (const d of days) {
      await addComment(t, { date: d.date, minutes: Math.round(d.diff * 60), html: '<p>Interne Nacharbeiten und Abgleich der Arbeitszeit.</p>', kontaktart: CFG.defaultKontaktart, billable: false });
    }
    log('Zeitausgleich gebucht', `#${sid}: ` + days.map(d => `${fmtDate(d.date)} ${fmtH(d.diff)} h`).join(', '));
  }

  function renderTimes() {
    const box = $('zView'); box.innerHTML = '';
    const pStore = personioStore();
    const pDays = Object.keys(pStore).sort();
    $('zPInfo').textContent = pDays.length ? `Personio-Daten: ${fmtDate(pDays[0])} – ${fmtDate(pDays[pDays.length - 1])} (${pDays.length} Tage)` : 'keine Personio-Daten';
    $('zPClear').style.display = pDays.length ? '' : 'none';
    if (!tm.data) { box.innerHTML = '<div class="hint" style="margin-top:10px">Zeitraum wählen und „Laden“ klicken.</div>'; return; }
    const { bookings, tickets, users } = tm.data;
    const single = users.length === 1 && users[0] === ME; // Personio-Abgleich nur für die eigenen Zeiten
    const low = +$('zLow').value || CFG.dayColors.low;
    const showDetails = $('zDetails').checked;
    if (!team().intern.length) box.appendChild(el('div', { className: 'hint', textContent: 'Hinweis: Unter ⚙ sind keine internen Firmen hinterlegt – alle Zeiten erscheinen deshalb als „Kunde“.' }));

    // Aggregation: Tag → Gruppe → Ticket
    const days = {};
    for (const c of bookings) {
      const d = c.LST_DATUM.slice(0, 10), t = tickets[c.MDERRDA_SID] || { SID: c.MDERRDA_SID, TITEL: '(Ticket nicht geladen)', KND_ID: '' };
      const g = isIntern(t.KND_ID) ? 'Intern' : 'Kunde';
      const day = days[d] ||= { std: 0, abr: 0, groups: { Intern: {}, Kunde: {} } };
      const row = day.groups[g][t.SID] ||= { t, std: 0, abr: 0, items: [] };
      row.std += c.LFZ_STD || 0; row.abr += c.INABR_STD || 0; row.items.push(c);
      day.std += c.LFZ_STD || 0; day.abr += c.INABR_STD || 0;
    }
    // Kennzahlen
    const all = { std: 0, abr: 0, intern: 0, kunde: 0, wd: 0, wdBooked: 0, wdEmpty: [], soll: 0, diffPlus: [], diffMinus: 0 };
    const dayList = [];
    for (let d = tm.from; d <= tm.to; d = addDays(d, 1)) dayList.push(d);
    for (const d of dayList) {
      const day = days[d];
      if (day) { all.std += day.std; all.abr += day.abr; for (const r of Object.values(day.groups.Intern)) all.intern += r.std; for (const r of Object.values(day.groups.Kunde)) all.kunde += r.std; }
      const pMin = pStore[d];
      const soll = single && pMin != null ? personioRound(pMin) : null;
      if (isWorkday(d) && d <= todayStr() && (soll == null || soll > 0)) { all.wd++; if (day) all.wdBooked++; else all.wdEmpty.push(d); }
      if (soll != null) {
        all.soll += soll;
        const diff = Math.round((soll - (day?.std || 0)) * 100) / 100;
        if (diff > 0) all.diffPlus.push({ date: d, diff }); else if (diff < 0) all.diffMinus += -diff;
      }
    }
    const hasP = single && dayList.some(d => pStore[d] != null);
    const sumDiff = all.diffPlus.reduce((a, x) => a + x.diff, 0);
    const kpi = el('div', { className: 'kpis' });
    const k = (label, val, sub = '') => `<div class="kpi"><div class="kl">${label}</div><div class="kv">${val}</div>${sub ? `<div class="ks">${sub}</div>` : ''}</div>`;
    kpi.innerHTML = k('Gebucht', fmtH(all.std) + ' h') + k('Abrechenbar', fmtH(all.abr) + ' h', all.std ? Math.round(all.abr / all.std * 100) + ' %' : '') +
      k('Kunde / Intern', `${fmtH(all.kunde)} / ${fmtH(all.intern)}`) +
      k('Arbeitstage mit Buchung', `${all.wdBooked} / ${all.wd}`, all.wdEmpty.length ? `ohne: ${all.wdEmpty.map(d => d.slice(8, 10) + '.' + d.slice(5, 7) + '.').join(', ')}` : '') +
      k('Ø je Arbeitstag', all.wdBooked ? fmtH(all.std / Math.max(1, all.wdBooked)) + ' h' : '–') +
      (hasP ? k('Personio (gerundet)', fmtH(all.soll) + ' h', `fehlt: ${fmtH(sumDiff)} h · zu viel: ${fmtH(all.diffMinus)} h`) : '');
    box.appendChild(kpi);
    if (hasP && all.diffPlus.length) {
      const bar = el('div', { className: 'row' }); bar.style.margin = '6px 0';
      const b = el('button', { className: 'btn sm', textContent: `Alle Fehlzeiten ausgleichen (${all.diffPlus.length} Tage, ${fmtH(sumDiff)} h)` });
      confirmBtn(b, () => { if (!/^\d+$/.test($('zTicket').value.trim())) throw new Error('Bitte zuerst ein Ticket für den Ausgleich eintragen.'); return `Wirklich ${fmtH(sumDiff)} h auf #${$('zTicket').value.trim()} buchen?`; },
        async () => { try { msg('zMsg', 'Buche …'); await bookBalance(all.diffPlus); await runTimes(); msg('zMsg', '✔ Ausgleich gebucht.', 'ok'); } catch (e) { msg('zMsg', '✖ ' + e.message, 'err'); } },
        e => msg('zMsg', '✖ ' + e.message, 'err'), 6000);
      bar.appendChild(b); box.appendChild(bar);
    }

    // Tagesliste
    const tbl = el('table', { className: 'ztbl' });
    tbl.innerHTML = `<thead><tr><th>Tag / Ticket</th><th>Kunde</th><th class="n">Std.</th><th class="n">abr.</th>${hasP ? '<th class="n">Personio</th><th class="n">Diff.</th><th></th>' : ''}</tr></thead>`;
    const tb = el('tbody'); tbl.appendChild(tb);
    const colN = hasP ? 7 : 4;
    for (const d of dayList) {
      const day = days[d], wd = isWorkday(d);
      if (!day && !wd && $('zWeekend').checked) continue;
      const pMin = pStore[d], soll = hasP && pMin != null ? personioRound(pMin) : null;
      const std = day?.std || 0, diff = soll != null ? Math.round((soll - std) * 100) / 100 : null;
      let cls = 'zday';
      if ($('zColor').checked && wd && d <= todayStr()) cls += soll === 0 ? ' z-free' : std <= 0 ? ' z-red' : std < low ? ' z-yellow' : ' z-green';
      if (!wd) cls += ' z-we';
      const tr = el('tr', { className: cls });
      tr.innerHTML = `<td><b>${WD[weekday(d)]} ${fmtDate(d)}</b>${!wd && day ? ' <span class="tag">Wochenende</span>' : ''}${std > 10 ? ' <span class="tag abr">über 10 h</span>' : ''}${soll === 0 ? ' <span class="tag">frei laut Personio</span>' : ''}</td><td></td>
        <td class="n"><b>${fmtH(std)}</b></td><td class="n">${fmtH(day?.abr || 0)}</td>` +
        (hasP ? `<td class="n">${soll == null ? '–' : fmtH(soll)}</td><td class="n ${diff > 0 ? 'neg' : diff < 0 ? 'pos' : ''}">${diff == null || diff === 0 ? (diff === 0 ? '✓' : '') : (diff > 0 ? '−' : '+') + fmtH(Math.abs(diff))}</td><td></td>` : '');
      if (hasP && diff > 0) {
        const b = el('button', { className: 'btn sm sec', textContent: `+${fmtH(diff)} h buchen`, title: 'Fehlende Zeit auf das Ausgleichs-Ticket buchen' });
        confirmBtn(b, () => { if (!/^\d+$/.test($('zTicket').value.trim())) throw new Error('Bitte zuerst ein Ticket für den Ausgleich eintragen.'); return `Auf #${$('zTicket').value.trim()}?`; },
          async () => { try { await bookBalance([{ date: d, diff }]); await runTimes(); msg('zMsg', `✔ ${fmtH(diff)} h am ${fmtDate(d)} gebucht.`, 'ok'); } catch (e) { msg('zMsg', '✖ ' + e.message, 'err'); } },
          e => msg('zMsg', '✖ ' + e.message, 'err'));
        tr.lastElementChild.appendChild(b);
      }
      tb.appendChild(tr);
      if (!day) continue;
      for (const g of ['Kunde', 'Intern']) {
        const rows = Object.values(day.groups[g]); if (!rows.length) continue;
        const gs = rows.reduce((a, r) => ({ std: a.std + r.std, abr: a.abr + r.abr }), { std: 0, abr: 0 });
        const gh = el('tr', { className: 'zgrp' });
        gh.innerHTML = `<td style="padding-left:14px">${g === 'Intern' ? '🏠 Intern' : '🏢 Kunde'}</td><td></td><td class="n">${fmtH(gs.std)}</td><td class="n">${fmtH(gs.abr)}</td>${hasP ? '<td></td><td></td><td></td>' : ''}`;
        tb.appendChild(gh);
        for (const r of rows.sort((a, b) => b.std - a.std)) {
          const trt = el('tr', { className: 'r zt' });
          trt.innerHTML = `<td style="padding-left:28px"><b>#${r.t.SID}</b> <span></span></td><td class="hint"></td><td class="n">${fmtH(r.std)}</td><td class="n">${fmtH(r.abr)}</td>${hasP ? '<td></td><td></td><td></td>' : ''}`;
          trt.querySelector('span').textContent = r.t.TITEL || '';
          trt.children[1].textContent = (r.t.MATCHCODE || '').replace(/^\d+\s*-\s*/, '');
          wbRow(trt, ticketRoute(r.t.SID), () => openTicket(r.t.SID));
          tb.appendChild(trt);
          if (showDetails) for (const c of r.items) {
            const tri = el('tr', { className: 'zi' });
            const von = (c.LFZ_VON || '').slice(11, 16), bis = (c.LFZ_BIS || '').slice(11, 16);
            tri.innerHTML = `<td style="padding-left:44px" colspan="2" class="hint"></td><td class="n hint">${fmtH(c.LFZ_STD)}</td><td class="n hint">${fmtH(c.INABR_STD)}</td>${hasP ? '<td></td><td></td><td></td>' : ''}`;
            tri.firstElementChild.textContent = `${von ? von + '–' + bis + ' · ' : ''}${users.length > 1 ? c.BEARBEITER_USER_ID + ' · ' : ''}${htmlToText(c.NOTIZ).slice(0, 140)}`;
            tb.appendChild(tri);
          }
        }
      }
    }
    const tf = el('tr', { className: 'tsum' });
    tf.innerHTML = `<td><b>Summe</b></td><td></td><td class="n"><b>${fmtH(all.std)}</b></td><td class="n"><b>${fmtH(all.abr)}</b></td>` +
      (hasP ? `<td class="n"><b>${fmtH(all.soll)}</b></td><td class="n"><b>${sumDiff ? '−' + fmtH(sumDiff) : '✓'}</b></td><td></td>` : '');
    tb.appendChild(tf);
    box.appendChild(tbl);

    // Kundenübersicht
    const per = {};
    for (const day of Object.values(days)) for (const g of ['Kunde', 'Intern']) for (const r of Object.values(day.groups[g])) {
      const key = r.t.KND_ID || '?'; const x = per[key] ||= { name: (r.t.MATCHCODE || key).replace(/^\d+\s*-\s*/, ''), intern: g === 'Intern', std: 0, abr: 0, tickets: new Set() };
      x.std += r.std; x.abr += r.abr; x.tickets.add(r.t.SID);
    }
    const ct = el('details', { className: 'card' });
    ct.innerHTML = `<summary>📊 Kundenübersicht (${Object.keys(per).length})</summary>`;
    persistDetails(ct, 'zeiten.kunden', true);
    const t2 = el('table');
    t2.innerHTML = '<thead><tr><th>Kunde</th><th class="n">Tickets</th><th class="n">Std.</th><th class="n">abr.</th><th class="n">Anteil</th></tr></thead>';
    const tb2 = el('tbody'); t2.appendChild(tb2);
    for (const x of Object.values(per).sort((a, b) => b.std - a.std)) {
      const r = el('tr');
      r.innerHTML = `<td></td><td class="n">${x.tickets.size}</td><td class="n">${fmtH(x.std)}</td><td class="n">${fmtH(x.abr)}</td><td class="n">${all.std ? Math.round(x.std / all.std * 100) : 0} %</td>`;
      r.firstElementChild.textContent = (x.intern ? '🏠 ' : '🏢 ') + x.name;
      tb2.appendChild(r);
    }
    ct.appendChild(t2); box.appendChild(ct);
  }

  // CSV-Export (Excel-tauglich: Semikolon, deutsches Dezimalkomma)
  $('zCsv').addEventListener('click', () => {
    if (!tm.data) { msg('zMsg', 'Erst Zeiten laden.', 'err'); return; }
    const num = v => String(Math.round((v || 0) * 100) / 100).replace('.', ',');
    const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const lines = [['Datum', 'Wochentag', 'Bearbeiter', 'Ticket', 'Titel', 'Kunde', 'Gruppe', 'Von', 'Bis', 'Stunden', 'abr. Stunden', 'Text'].join(';')];
    for (const c of tm.data.bookings) {
      const t = tm.data.tickets[c.MDERRDA_SID] || {}, d = c.LST_DATUM.slice(0, 10);
      lines.push([fmtDate(d), WD[weekday(d)], c.BEARBEITER_USER_ID, c.MDERRDA_SID, q(t.TITEL), q((t.MATCHCODE || '').replace(/^\d+\s*-\s*/, '')), isIntern(t.KND_ID) ? 'Intern' : 'Kunde',
        (c.LFZ_VON || '').slice(11, 16), (c.LFZ_BIS || '').slice(11, 16), num(c.LFZ_STD), num(c.INABR_STD), q(htmlToText(c.NOTIZ).slice(0, 500))].join(';'));
    }
    download(`Zeiten_${tm.from}_${tm.to}.csv`, '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  });

  // ================= SYNCHRONISATION ZWISCHEN TABS & ÄNDERUNGEN VON AUSSEN =================
  // Kurzer „Fingerabdruck“ eines Tickets, um Änderungen in MultiData oder von Kolleg:innen zu erkennen
  async function ticketSignature(sid) {
    const [t] = await load('MDERRDA', { filter: [['SID', '=', Number(sid)]], sqlColumns: ['SID', 'DATE_MODIFIED', 'STATUS', 'SUM_STD', 'MEMO_SUM_ABRSTD', 'BEARBEITER_USER_ID', 'TITEL', 'PRIORITY', 'ERRQ_ID'] });
    const cs = await load('MDERRDA_MDERRDAC', { take: 3000, filter: [['MDERRDA_SID', '=', Number(sid)]], sqlColumns: ['SID'] });
    return JSON.stringify([t && [t.DATE_MODIFIED, t.STATUS, t.SUM_STD, t.MEMO_SUM_ABRSTD, t.BEARBEITER_USER_ID, t.TITEL, t.PRIORITY, t.ERRQ_ID], cs.length, Math.max(0, ...cs.map(c => c.SID))]);
  }
  function ticketChangedElsewhere(where) {
    if (ui.isDirty && ui.isDirty()) {
      const b = $('tBanner'); b.style.display = ''; b.firstChild.textContent = `Dieses Ticket wurde ${where} geändert.`;
      return;
    }
    reloadTicket().then(() => msg('tMsg', `↻ Automatisch aktualisiert – das Ticket wurde ${where} geändert.`, 'ok'));
  }
  $('tBannerReload').addEventListener('click', () => reloadTicket());

  let searchTimer2, timesTimer;
  function refreshSearchSoon() {
    if (!ui.lastResult) return;
    if (ui.tab !== 'search' || document.hidden || !panel.classList.contains('open')) { ui.searchStale = true; return; }
    clearTimeout(searchTimer2); searchTimer2 = setTimeout(() => runSearch(true), 800);
  }
  function refreshTimesSoon() {
    if (!tm.data) return;
    if (ui.tab !== 'times' || document.hidden || !panel.classList.contains('open')) { ui.timesStale = true; return; }
    clearTimeout(timesTimer); timesTimer = setTimeout(() => runTimes(), 1000);
  }
  // Tickets wurden über die Werkbank geändert (dieser oder ein anderer Tab)
  bus.on('tickets', (sids, remote) => {
    refreshSearchSoon(); refreshTimesSoon();
    if (remote && ui.ticket && sids.includes(Number(ui.ticket.SID))) ticketChangedElsewhere('in einem anderen Werkbank-Tab');
  });
  // Gemeinsam gespeicherte Daten wurden in einem anderen Tab geändert
  bus.on('store', key => {
    if (key === CFG.store.inbox) { updateBadges(); if (ui.tab === 'inbox') renderInbox(); }
    else if (key === CFG.store.drafts) {
      updateBadges(); if (ui.tab === 'drafts') renderDrafts();
      // derselbe Entwurf wurde in einem anderen Tab weitergeschrieben → anbieten zu übernehmen
      const d = ui.ticket && drafts.get('ticket:' + ui.ticket.SID);
      if (d && d.tab !== TAB_ID && d.savedAt !== ui.draftSeen && ui.form && ui.addBox) {
        ui.draftSeen = d.savedAt;
        let note = ui.addBox.querySelector('[data-ref=draftNote]');
        if (!note) { note = el('div', { className: 'warnbox' }); note.dataset.ref = 'draftNote'; ui.addBox.insertBefore(note, ui.form.el); }
        note.innerHTML = ''; note.appendChild(document.createTextNode('Dieser Entwurf wurde in einem anderen Tab geändert. '));
        const b = el('button', { className: 'btn sm', textContent: 'Übernehmen' });
        b.addEventListener('click', () => { ui.form.restore(drafts.get('ticket:' + ui.ticket.SID)?.form); ui.addBox.open = true; note.remove(); });
        note.appendChild(b);
      }
    }
    else if (key === CFG.store.settings) applyTheme();
    else if (key === CFG.store.templates) renderTemplates();
    else if (key === CFG.store.facetSort) { $('sFacetSort').value = facetSort(); renderFacets(); }
    else if (key === CFG.store.team) { renderTeam(); updateTeamHint(); }
    else if (key === CFG.store.log && ui.tab === 'log') renderLog();
  });
  // Regelmäßig prüfen: offenes Ticket (Änderungen in MultiData/von Kolleg:innen) und offene Suche – nur im sichtbaren Tab
  const lastPoll = { ticket: Date.now(), search: Date.now() };
  setInterval(async () => {
    if (document.hidden || !panel.classList.contains('open')) return;
    const now = Date.now();
    if (ui.tab === 'ticket' && ui.ticket && ui.ticketSig && now - lastPoll.ticket > CFG.pollTicketSec * 1000) {
      lastPoll.ticket = now;
      try {
        const sid = ui.ticket.SID, sig = await ticketSignature(sid);
        if (ui.ticket && ui.ticket.SID === sid && ui.ticketSig && sig !== ui.ticketSig) { ui.ticketSig = sig; ticketChangedElsewhere('in MultiData oder von Kolleg:innen'); }
      } catch { /* nächster Versuch */ }
    }
    const typing = document.activeElement && document.activeElement.closest && document.activeElement.closest('section[data-tab=search]');
    if (ui.tab === 'search' && ui.lastResult && !typing && now - Math.max(lastPoll.search, ui.lastSearchAt || 0) > CFG.pollSearchSec * 1000) {
      lastPoll.search = now; runSearch(true);
    }
  }, 15000);
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) return;
    if (ui.tab === 'search' && ui.searchStale) { ui.searchStale = false; runSearch(true); }
    if (ui.tab === 'times' && ui.timesStale && tm.data) { ui.timesStale = false; runTimes(); }
  });

  // ================= PROTOKOLL =================
  function renderLog() {
    const box = $('logView'); box.innerHTML = '';
    const l = store.get(CFG.store.log, []);
    if (!l.length) { box.innerHTML = '<div class="hint" style="margin-top:8px">Noch keine Einträge.</div>'; return; }
    const tbl = el('table');
    tbl.innerHTML = '<thead><tr><th>Zeit</th><th>Aktion</th><th>Details</th></tr></thead>';
    const tb = el('tbody'); tbl.appendChild(tb);
    for (const e of l) {
      const tr = el('tr'); const d = new Date(e.t);
      tr.innerHTML = `<td style="white-space:nowrap">${d.toLocaleDateString('de-DE')} ${d.toLocaleTimeString('de-DE').slice(0, 5)}</td><td>${esc(e.action)}</td><td>${esc(e.detail)}</td>`;
      tb.appendChild(tr);
    }
    box.appendChild(tbl);
  }
  confirmBtn($('logClear'), 'Wirklich leeren?', async () => { store.set(CFG.store.log, []); renderLog(); });
  } // main
})();
