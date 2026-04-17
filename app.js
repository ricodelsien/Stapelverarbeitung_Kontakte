/**
 * Adressliste – Excel-ähnliche Tabelle, Filter, Serienbrief-Export (Word)
 */
/* global XLSX, ExcelJS */

const STORAGE_KEY = "bwk-adressliste-json-v3";
const FAILSAFE_KEY = "bwk-adressliste-failsafe-json-v3";
const EVENT_KEY = "bwk-event-settings-v1";

const DEFAULT_APP_TITLE = "Adressliste Kooperationspartner";
/** @type {string} */
let appTitle = DEFAULT_APP_TITLE;
/** @type {boolean} */
let appTitleEditing = false;

function normalizeAppTitle(raw) {
  const s = typeof raw === "string" ? raw.trim() : "";
  if (!s) return DEFAULT_APP_TITLE;
  return s.slice(0, 120);
}

function applyAppTitleToDom() {
  const t = appTitle || DEFAULT_APP_TITLE;
  const h = document.getElementById("app-title-heading");
  const inp = document.getElementById("app-title-input");
  if (h && !appTitleEditing) h.textContent = t;
  if (inp && !appTitleEditing) inp.value = t;
  document.title = t;
}

function showAppTitleInput() {
  const h = document.getElementById("app-title-heading");
  const inp = document.getElementById("app-title-input");
  if (!h || !inp || appTitleEditing) return;
  appTitleEditing = true;
  inp.value = appTitle;
  h.classList.add("hidden");
  inp.classList.remove("hidden");
  inp.focus();
  inp.select();
}

/**
 * @param {boolean} commit Wenn false (Escape): Anzeige zurücksetzen, nicht speichern.
 */
function finishAppTitleEdit(commit) {
  const h = document.getElementById("app-title-heading");
  const inp = document.getElementById("app-title-input");
  if (!h || !inp || !appTitleEditing) return;
  appTitleEditing = false;
  if (commit) {
    const next = normalizeAppTitle(inp.value);
    if (next !== appTitle) pushUndoBeforeMutation();
    appTitle = next;
  }
  inp.classList.add("hidden");
  h.classList.remove("hidden");
  applyAppTitleToDom();
  if (commit) saveAll();
}

function filenameStemFromAppTitle() {
  const raw = (appTitle || DEFAULT_APP_TITLE)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "")
    .slice(0, 48);
  return raw || "Adressliste";
}

let rowsExtern = [];
let rowsTN = [];
/** @type {Record<string, string>} */
let plzMap = {};
let plzLoaded = false;

/** @type {Record<string, string>} */
const filters = {};
const FILTER_KEYS = [
  "status",
  "unternehmen",
  "anrede",
  "titel",
  "vorname",
  "nachname",
  "email",
  "abteilung",
  "strasse",
  "hausnr",
  "plz",
  "stadt",
  "anmerkungen",
];

let eventSettings = {
  datum: "",
  zeit: "",
  ort: "",
  terminZeile: "",
  /** Schreibdatum für „Berlin, …“ im Briefkopf (TT.MM.JJJJ); leer = heute */
  briefdatum: "",
  serienExportPrefix: "",
  serienMailSubject: "",
  /** Optional: Absender in der .eml (From:) — vermeidet „Kein Absender“ und hilft manchen Mail-Programmen */
  serienMailFrom: "",
  /** @type {"both" | "word" | "email"} */
  serienExportWhat: "both",
};

/** @type {Set<string>} */
const selectedIds = new Set();

let currentTab = "liste";
/** @type {string} "extern" | "tn" | "extra:<uuid>" */
let listSubTab = "extern";

function isListSubTabExternLike() {
  return listSubTab === "extern" || (typeof listSubTab === "string" && listSubTab.startsWith("extra:"));
}

function normalizeListSubTabFromStorage(raw) {
  if (raw === "tn") return "tn";
  if (typeof raw === "string" && raw.startsWith("extra:")) return raw;
  return "extern";
}

const STORAGE_VERSION = 5;

/** @typedef {{ id: string, name: string, useExtern?: boolean, useTn?: boolean, rowsExtern: typeof rowsExtern, rowsTN: typeof rowsTN, extraLists?: Array<{ id: string, name: string, rows: typeof rowsExtern }> }} ListSet */
/** @type {ListSet[]} */
let listSets = [];

/** Mindestens eine Unterliste (Extern oder TN) muss sichtbar sein. */
function normalizeSublistFlags(ls) {
  if (!ls || typeof ls !== "object") return { useExtern: true, useTn: true };
  let useExtern = ls.useExtern !== false;
  let useTn = ls.useTn !== false;
  if (!useExtern && !useTn) {
    useExtern = true;
    useTn = true;
  }
  return { useExtern, useTn };
}

function ensureValidListSubTab() {
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  const { useExtern, useTn } = normalizeSublistFlags(cur);
  cur.useExtern = useExtern;
  cur.useTn = useTn;
  if (!cur.extraLists) cur.extraLists = [];

  function firstExtraTab() {
    return cur.extraLists.length ? `extra:${cur.extraLists[0].id}` : "extern";
  }

  if (listSubTab === "extern" && !useExtern) {
    if (useTn) listSubTab = "tn";
    else listSubTab = firstExtraTab();
  }
  if (listSubTab === "tn" && !useTn) {
    if (useExtern) listSubTab = "extern";
    else listSubTab = firstExtraTab();
  }
  if (typeof listSubTab === "string" && listSubTab.startsWith("extra:")) {
    const eid = listSubTab.slice(6);
    if (!cur.extraLists.some((x) => x.id === eid)) {
      if (useExtern) listSubTab = "extern";
      else if (useTn) listSubTab = "tn";
      else listSubTab = firstExtraTab();
    }
  }
}

function listSetFromStorage(ls) {
  const { useExtern, useTn } = normalizeSublistFlags(ls);
  return {
    id: String(ls.id || "").trim() || crypto.randomUUID(),
    name: String(ls.name || "Liste").trim().slice(0, 96) || "Liste",
    useExtern,
    useTn,
    rowsExtern: (ls.rowsExtern || []).map(migrateRow),
    rowsTN: (ls.rowsTN && ls.rowsTN.length ? ls.rowsTN : [emptyTnRow()]).map(migrateTnRow),
    extraLists: migrateExtraLists(ls),
  };
}

function listSetFromImport(ls) {
  const { useExtern, useTn } = normalizeSublistFlags(ls);
  const rowsTNRaw = (ls.rowsTN || []).map(migrateTnRow);
  return {
    id: String(ls.id || "").trim() || crypto.randomUUID(),
    name: String(ls.name || "Liste").trim().slice(0, 96) || "Liste",
    useExtern,
    useTn,
    rowsExtern: (ls.rowsExtern || []).map(migrateRow),
    rowsTN: rowsTNRaw.length ? rowsTNRaw : [emptyTnRow()],
    extraLists: migrateExtraLists(ls),
  };
}
/** @type {string | null} */
let activeListId = null;

const emptyRow = () => ({
  id: crypto.randomUUID(),
  status: "",
  unternehmen: "",
  anrede: "",
  titel: "",
  vorname: "",
  nachname: "",
  email: "",
  abteilung: "",
  strasse: "",
  hausnr: "",
  plz: "",
  stadt: "",
  anmerkungen: "",
});

function migrateRow(r) {
  if (!r || typeof r !== "object") return emptyRow();
  if (r.adresse && typeof r.adresse === "string" && r.adresse.trim() && !r.strasse) {
    const lines = r.adresse.trim().split(/\r?\n/);
    r.strasse = lines[0] || "";
    if (lines[1]) {
      const m = lines[1].match(/^(\d{5})\s+(.+)$/);
      if (m) {
        r.plz = m[1];
        r.stadt = resolveStadtFromPlzAndOrt(m[1], m[2]);
      }
    }
    delete r.adresse;
  }
  if (r.strasse === undefined) r.strasse = "";
  if (r.hausnr === undefined) r.hausnr = "";
  if (r.plz === undefined) r.plz = "";
  if (r.stadt === undefined) r.stadt = "";
  const pz = String(r.plz || "").replace(/\s/g, "");
  if (pz.length === 5 && /^\d{5}$/.test(pz) && String(r.stadt || "").trim()) {
    r.stadt = resolveStadtFromPlzAndOrt(pz, r.stadt);
  }
  return r;
}

const emptyTnRow = () => ({
  id: crypto.randomUUID(),
  status: "",
  anrede: "",
  vorname: "",
  nachname: "",
});

const TN_STATUS_VALUES = ["", "einladen", "offen", "Zusage", "Absage"];

function migrateTnRow(r) {
  if (!r || typeof r !== "object") return emptyTnRow();
  const st = (r.status || "").trim();
  const lower = st.toLowerCase();
  if (TN_STATUS_VALUES.includes(lower)) r.status = lower;
  else r.status = "";
  if (r.anrede === undefined) r.anrede = "";
  if (r.vorname === undefined) r.vorname = "";
  if (r.nachname === undefined) r.nachname = "";
  return r;
}

function migrateExtraLists(ls) {
  return (ls.extraLists || []).map((e) => ({
    id: String(e.id || "").trim() || crypto.randomUUID(),
    name: String(e.name || "Liste").trim().slice(0, 96) || "Liste",
    rows: (e.rows && e.rows.length ? e.rows : [emptyRow()]).map(migrateRow),
  }));
}

/** data-status für Zeilenfarbe (TN-Tabelle) */
function tnRowDatasetStatus(st) {
  const s = (st || "").trim().toLowerCase();
  if (s === "einladen" || s === "offen" || s === "Absage" || s === "Zusage") return s;
  return "";
}

function rowStatusTn(r) {
  return (r.status || "").trim();
}

function tnRowHasText(r) {
  return (
    rowStatusTn(r) ||
    r.anrede.trim() ||
    r.vorname.trim() ||
    r.nachname.trim()
  );
}

function lfdTnForIndex(i) {
  let n = 0;
  for (let j = 0; j <= i; j++) {
    const row = rowsTN[j];
    if (rowStatusTn(row) || row.anrede.trim() || row.vorname.trim() || row.nachname.trim()) n += 1;
  }
  const r = rowsTN[i];
  if (!rowStatusTn(r) && !r.anrede.trim() && !r.vorname.trim() && !r.nachname.trim()) return "";
  return String(n);
}

function rowStatus(row) {
  return row.status || "";
}

/** Versand-Tabelle: einladen/offen klein; Zusage/Absage mit großem Anfangsbuchstaben. */
function formatExternStatusSerienDisplay(st) {
  const s = String(st || "").trim();
  if (!s) return "—";
  const labels = {
    einladen: "einladen",
    offen: "offen",
    zusage: "Zusage",
    absage: "Absage",
  };
  return labels[s] ?? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

/**
 * Zweiter Klick auf dasselbe gewählte Status-Segment hebt die Auswahl auf.
 * Nur per click auswerten (nicht mousedown + uncheck): sonst feuert der nachfolgende click
 * und wählt das Radio durch das Label wieder an.
 */
function bindStatusSegmentRepeatClickClears(label, inp, onClear) {
  label.addEventListener(
    "mousedown",
    (e) => {
      if (e.button !== 0) return;
      label.dataset.statusWasChecked = inp.checked ? "1" : "";
    },
    true
  );
  label.addEventListener(
    "click",
    (e) => {
      if (label.dataset.statusWasChecked !== "1") {
        delete label.dataset.statusWasChecked;
        return;
      }
      delete label.dataset.statusWasChecked;
      if (!inp.checked) return;
      e.preventDefault();
      e.stopPropagation();
      inp.checked = false;
      onClear();
    },
    true
  );
  inp.addEventListener("keydown", (e) => {
    if (e.key !== " " && e.key !== "Enter") return;
    if (!inp.checked) return;
    e.preventDefault();
    inp.checked = false;
    onClear();
  });
}

function getSerienStatusFilter() {
  const el = document.querySelector('input[name="serien-filter-status"]:checked');
  return el ? el.value : "";
}

function rowHasText(r) {
  return (
    r.unternehmen.trim() ||
    r.anrede.trim() ||
    r.titel.trim() ||
    r.vorname.trim() ||
    r.nachname.trim() ||
    r.email.trim() ||
    r.abteilung.trim() ||
    r.strasse.trim() ||
    r.hausnr.trim() ||
    r.plz.trim() ||
    r.stadt.trim() ||
    r.anmerkungen.trim()
  );
}

/**
 * Externe Liste: „unvollständig“ wenn eine der Kernkategorien fehlt (für Serienbrief/Kontakt).
 * — Status, Firma oder Name, E-Mail, PLZ+Ort
 */
function externRowIsIncomplete(r) {
  const st = rowStatus(r);
  const hasIdentity = r.unternehmen.trim() || r.vorname.trim() || r.nachname.trim();
  const hasEmail = r.email.trim();
  const hasOrt = r.plz.trim() && r.stadt.trim();
  return !st || !hasIdentity || !hasEmail || !hasOrt;
}

/** TN-Liste: Status, Anrede, Vor- und Nachname */
function tnRowIsIncomplete(r) {
  return (
    !rowStatusTn(r) || !r.anrede.trim() || !r.vorname.trim() || !r.nachname.trim()
  );
}

function applyRowIncompleteClass(tr, incomplete, withTitle = true) {
  const on = !!incomplete;
  tr.classList.toggle("row-incomplete", on);
  tr.dataset.rowIncomplete = on ? "1" : "0";
  if (on && withTitle) {
    tr.title =
      "Eintrag unvollständig: schraffierte Felder oder Status bitte ergänzen — mit Inhalt verschwindet die Schraffur sofort.";
  } else {
    tr.removeAttribute("title");
  }
}

/** Leere Felder: Schraffur nur auf der abgerundeten Eingabefläche (Input bzw. Status-Segmentbereich). */
function refreshEmptyCellSchraffurExtern(tr, r) {
  tr.querySelectorAll("input.cell-input").forEach((inp) => inp.classList.remove("cell-input--schraffiert"));
  tr.querySelectorAll("td.cell-status").forEach((td) => td.classList.remove("cell-status--schraffiert"));

  const tdStatus = tr.querySelector("td.cell-status");
  if (tdStatus && !rowStatus(r)) tdStatus.classList.add("cell-status--schraffiert");

  tr.querySelectorAll("input.cell-input[data-field]").forEach((inp) => {
    if (!String(inp.value || "").trim()) inp.classList.add("cell-input--schraffiert");
  });
}

function refreshEmptyCellSchraffurTn(tr, r) {
  tr.querySelectorAll("input.cell-input").forEach((inp) => inp.classList.remove("cell-input--schraffiert"));
  tr.querySelectorAll("td.cell-status").forEach((td) => td.classList.remove("cell-status--schraffiert"));

  const tdStatus = tr.querySelector("td.cell-status");
  if (tdStatus && !rowStatusTn(r)) tdStatus.classList.add("cell-status--schraffiert");

  tr.querySelectorAll("input.cell-input[data-field]").forEach((inp) => {
    if (!String(inp.value || "").trim()) inp.classList.add("cell-input--schraffiert");
  });
}

function applyExternRowIncompleteUi(tr, r, withTitle = true) {
  applyRowIncompleteClass(tr, externRowIsIncomplete(r), withTitle);
  refreshEmptyCellSchraffurExtern(tr, r);
}

function applyTnRowIncompleteUi(tr, r) {
  applyRowIncompleteClass(tr, tnRowIsIncomplete(r));
  refreshEmptyCellSchraffurTn(tr, r);
}

function hasMeaningfulData() {
  return (
    rowsExtern.some((r) => rowStatus(r) || rowHasText(r)) || rowsTN.some((r) => tnRowHasText(r))
  );
}

function normMatchPart(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

/** Nur Vor- und Nachname (beide nicht leer) — Duplikate in derselben Liste erkennen. */
function dupNamePairKey(r) {
  const vn = normMatchPart(r.vorname);
  const nn = normMatchPart(r.nachname);
  if (!vn || !nn) return null;
  return `${vn}\u0000${nn}`;
}

/**
 * Indizes aller Datenzeilen, deren Vor- und Nachname gemeinsam mindestens ein weiteres Mal vorkommen.
 * @param {typeof rowsExtern} rows
 * @param {(r: object) => boolean} isDataRow
 */
function indicesWithDuplicateNamePairs(rows, isDataRow) {
  const keyCounts = new Map();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!isDataRow(r)) continue;
    const k = dupNamePairKey(r);
    if (!k) continue;
    keyCounts.set(k, (keyCounts.get(k) || 0) + 1);
  }
  const dupKeys = new Set();
  for (const [k, c] of keyCounts) {
    if (c > 1) dupKeys.add(k);
  }
  const indices = new Set();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!isDataRow(r)) continue;
    const k = dupNamePairKey(r);
    if (k && dupKeys.has(k)) indices.add(i);
  }
  return indices;
}

function applyDuplicateNamePairRowUi(tr, nameDup) {
  tr.classList.toggle("row-duplicate-name", !!nameDup);
  tr.dataset.nameDuplicate = nameDup ? "1" : "0";
  if (!nameDup) return;
  const dupHint =
    "Gleicher Vor- und Nachname wie in einer anderen Zeile dieser Liste — prüfen, ob ein Duplikat (z. B. doppelte Zeile) entfernt werden soll.";
  const prev = tr.getAttribute("title");
  tr.title = prev ? dupHint + "\n\n" + prev : dupHint;
}

/** E-Mail bzw. Person+Firma+PLZ — für Abgleich bei erneutem Listen-Import */
function externContactMatchKey(r) {
  const em = normMatchPart(r.email).replace(/\s/g, "");
  if (em && em.includes("@")) return "mail:" + em;
  const u = normMatchPart(r.unternehmen);
  const vn = normMatchPart(r.vorname);
  const nn = normMatchPart(r.nachname);
  const plz = normMatchPart(r.plz).replace(/\s/g, "");
  return "addr:" + u + "|" + vn + "|" + nn + "|" + plz;
}

function tnContactMatchKey(r) {
  return (
    "tn:" + normMatchPart(r.vorname) + "|" + normMatchPart(r.nachname) + "|" + normMatchPart(r.anrede)
  );
}

function collectExternMatchKeys(rows) {
  const set = new Set();
  for (const r of rows) {
    if (!rowStatus(r) && !rowHasText(r)) continue;
    set.add(externContactMatchKey(r));
  }
  return set;
}

function collectTnMatchKeys(rows) {
  const set = new Set();
  for (const r of rows) {
    if (!tnRowHasText(r)) continue;
    set.add(tnContactMatchKey(r));
  }
  return set;
}

function dedupeRowsByContactKey(rows, keyFn) {
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    const k = keyFn(r);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(r);
  }
  return out;
}

/** Anzahl „überzähliger“ Datenzeilen mit bereits vorkommendem Kontakt-Schlüssel (ab 2. Vorkommen). */
function countDuplicateExternDataRows(rows) {
  const seen = new Set();
  let dups = 0;
  for (const r of rows) {
    if (!rowStatus(r) && !rowHasText(r)) continue;
    const k = externContactMatchKey(r);
    if (seen.has(k)) dups++;
    else seen.add(k);
  }
  return dups;
}

function countDuplicateTnDataRows(rows) {
  const seen = new Set();
  let dups = 0;
  for (const r of rows) {
    if (!tnRowHasText(r)) continue;
    const k = tnContactMatchKey(r);
    if (seen.has(k)) dups++;
    else seen.add(k);
  }
  return dups;
}

function trimEmptyExternRowsOnArray(rows) {
  const filtered = rows.filter((r) => rowStatus(r) || rowHasText(r));
  rows.length = 0;
  rows.push(...filtered);
  if (!rows.length) rows.push(emptyRow());
}

function trimEmptyTnRowsOnArray(rows) {
  const filtered = rows.filter((r) => tnRowHasText(r));
  rows.length = 0;
  rows.push(...filtered);
  if (!rows.length) rows.push(emptyTnRow());
}

/** Erste Zeile je Kontakt behalten (gleiche Logik wie beim Excel-Abgleich). */
function dedupeExternRowsArrayInPlace(rows) {
  const dataRows = rows.filter((r) => rowStatus(r) || rowHasText(r));
  const seen = new Set();
  const kept = [];
  for (const r of dataRows) {
    const k = externContactMatchKey(r);
    if (seen.has(k)) continue;
    seen.add(k);
    kept.push(r);
  }
  rows.length = 0;
  rows.push(...kept);
  trimEmptyExternRowsOnArray(rows);
}

function dedupeTnRowsArrayInPlace(rows) {
  const dataRows = rows.filter((r) => tnRowHasText(r));
  const seen = new Set();
  const kept = [];
  for (const r of dataRows) {
    const k = tnContactMatchKey(r);
    if (seen.has(k)) continue;
    seen.add(k);
    kept.push(r);
  }
  rows.length = 0;
  rows.push(...kept);
  trimEmptyTnRowsOnArray(rows);
}

function countDuplicatesInListSet(ls) {
  let n = 0;
  n += countDuplicateExternDataRows(ls.rowsExtern || []);
  n += countDuplicateTnDataRows(ls.rowsTN || []);
  for (const el of ls.extraLists || []) {
    n += countDuplicateExternDataRows(el.rows || []);
  }
  return n;
}

function countDuplicatesInAllListSets() {
  return listSets.reduce((sum, ls) => sum + countDuplicatesInListSet(ls), 0);
}

function dedupeAllListSetsInPlace() {
  for (const ls of listSets) {
    if (!ls.rowsExtern || !ls.rowsExtern.length) ls.rowsExtern = [emptyRow()];
    else dedupeExternRowsArrayInPlace(ls.rowsExtern);
    if (!ls.rowsTN || !ls.rowsTN.length) ls.rowsTN = [emptyTnRow()];
    else dedupeTnRowsArrayInPlace(ls.rowsTN);
    for (const el of ls.extraLists || []) {
      if (!el.rows || !el.rows.length) el.rows = [emptyRow()];
      else dedupeExternRowsArrayInPlace(el.rows);
    }
  }
  loadRowsForCurrentSubTab();
}

function pruneSerienSelectedIdsInvalid() {
  for (const id of [...selectedIds]) {
    if (!rowsExtern.some((r) => r.id === id)) selectedIds.delete(id);
  }
}

function maybePromptRemoveDuplicatesAfterSpreadsheetImport(targetTn, previousStatusLine) {
  const rows = targetTn ? rowsTN : rowsExtern;
  const n = targetTn ? countDuplicateTnDataRows(rows) : countDuplicateExternDataRows(rows);
  if (n <= 0) return;
  const tnHint = targetTn ? " TN: gleicher Vorname, Nachname und Anrede." : "";
  const ok = confirm(
    `In der Liste gibt es ${n} doppelte Kontakt-Zeile(n) (gleiche E-Mail oder gleiche Kombination Firma, Vorname, Nachname und PLZ).${tnHint}\n\n` +
      "Diese Duplikate entfernen und nur die erste Zeile je Kontakt behalten?",
  );
  if (!ok) return;
  pushUndoBeforeMutation();
  if (targetTn) dedupeTnRowsArrayInPlace(rowsTN);
  else dedupeExternRowsArrayInPlace(rowsExtern);
  syncActiveListIntoListSets();
  saveAll();
  pruneSerienSelectedIdsInvalid();
  render();
  const base = previousStatusLine ? String(previousStatusLine).replace(/\.$/, "") : "Liste";
  setStatus(`${base} · ${n} Duplikat(e) entfernt.`);
}

function maybePromptDuplicatesAfterJsonImport(importMsg) {
  const n = countDuplicatesInAllListSets();
  if (n <= 0) return;
  const ok = confirm(
    `In den importierten Listen gibt es ${n} doppelte Kontakt-Zeile(n) (gleicher Abgleich wie beim Tabellen-Import: E-Mail bzw. Firma/Vorname/Nachname/PLZ; TN: Vorname/Nachname/Anrede).\n\n` +
      "Diese Duplikate entfernen und nur die erste Zeile je Kontakt behalten?",
  );
  if (!ok) return;
  pushUndoBeforeMutation();
  dedupeAllListSetsInPlace();
  syncActiveListIntoListSets();
  saveAll();
  pruneSerienSelectedIdsInvalid();
  render();
  setStatus(`${importMsg || "Import abgeschlossen."} · ${n} Duplikat(e) entfernt.`);
}

function findFirstExternRowIndexByKey(key) {
  for (let i = 0; i < rowsExtern.length; i++) {
    if (externContactMatchKey(rowsExtern[i]) === key) return i;
  }
  return -1;
}

function findFirstTnRowIndexByKey(key) {
  for (let i = 0; i < rowsTN.length; i++) {
    if (tnContactMatchKey(rowsTN[i]) === key) return i;
  }
  return -1;
}

function normImportField(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

const EXTERN_IMPORT_DIFF_FIELDS = [
  ["status", "Status"],
  ["unternehmen", "Unternehmen"],
  ["anrede", "Anrede"],
  ["titel", "Titel"],
  ["vorname", "Vorname"],
  ["nachname", "Nachname"],
  ["email", "E-Mail"],
  ["abteilung", "Abteilung"],
  ["strasse", "Straße"],
  ["hausnr", "Nr."],
  ["plz", "PLZ"],
  ["stadt", "Stadt"],
  ["anmerkungen", "Anmerkungen"],
];

const TN_IMPORT_DIFF_FIELDS = [
  ["status", "Status"],
  ["anrede", "Anrede"],
  ["vorname", "Vorname"],
  ["nachname", "Nachname"],
];

function diffExternRowAgainstImport(existing, incoming) {
  const changes = [];
  for (const [key, label] of EXTERN_IMPORT_DIFF_FIELDS) {
    const a = normImportField(existing[key]);
    const b = normImportField(incoming[key]);
    if (a !== b) changes.push({ key, label, from: a, to: b });
  }
  return changes;
}

function diffTnRowAgainstImport(existing, incoming) {
  const changes = [];
  for (const [key, label] of TN_IMPORT_DIFF_FIELDS) {
    const a = normImportField(existing[key]);
    const b = normImportField(incoming[key]);
    if (a !== b) changes.push({ key, label, from: a, to: b });
  }
  return changes;
}

function applyExternRowFromImport(dest, src) {
  for (const [key] of EXTERN_IMPORT_DIFF_FIELDS) {
    dest[key] = src[key] ?? "";
  }
}

function applyTnRowFromImport(dest, src) {
  for (const [key] of TN_IMPORT_DIFF_FIELDS) {
    dest[key] = src[key] ?? "";
  }
}

/** Kurzbezeichnung für Meldungen (Name oder Firma) */
function importRowLabel(r, targetTn) {
  if (targetTn) {
    const n = [r.vorname, r.nachname].filter(Boolean).join(" ").trim();
    if (n) return n;
    return r.anrede || "TN";
  }
  const n = [r.vorname, r.nachname].filter(Boolean).join(" ").trim();
  if (n) return n;
  if (r.unternehmen && String(r.unternehmen).trim()) return String(r.unternehmen).trim();
  return r.email || "Kontakt";
}

function formatImportChangeLine(c) {
  const a = c.from ? `„${c.from}“` : "leer";
  const b = c.to ? `„${c.to}“` : "leer";
  return `  · ${c.label}: ${a} → ${b}`;
}

function buildImportConfirmText(sheetName, listLabel, newCount, updateItems, dupInFile, unchangedCount, targetTn) {
  const parts = [];
  if (newCount) parts.push(`${newCount} neue Zeile(n)`);
  if (updateItems.length) parts.push(`${updateItems.length} bestehende Zeile(n) mit geänderten Zellen (Excel-Stand übernehmen)`);
  let msg = `${parts.join(" · ")} aus „${sheetName}“ in die ${listLabel} übernehmen?`;
  const extras = [];
  if (dupInFile) extras.push(`${dupInFile} doppelt in der Datei (nur erste je Kontakt)`);
  if (unchangedCount) extras.push(`${unchangedCount} bereits identisch (keine Änderung)`);
  if (extras.length) msg += `\n\n${extras.join(" · ")}.`;
  if (updateItems.length) {
    const lines = [];
    for (const u of updateItems.slice(0, 8)) {
      const who = importRowLabel(u.incoming, targetTn);
      const flds = u.changes
        .slice(0, 4)
        .map((c) => `${c.label}`)
        .join(", ");
      lines.push(`• ${who}: ${flds}${u.changes.length > 4 ? ", …" : ""}`);
    }
    msg += `\n\nÄnderungen (Auszug):\n${lines.join("\n")}`;
    if (updateItems.length > 8) msg += `\n… und ${updateItems.length - 8} weitere`;
  }
  return msg;
}

/** Nach Import: leere Zeilen entfernen, mindestens eine Bearbeitungszeile behalten */
function trimEmptyExternRows() {
  rowsExtern = rowsExtern.filter((r) => rowStatus(r) || rowHasText(r));
  if (!rowsExtern.length) rowsExtern.push(emptyRow());
}

function trimEmptyTnRows() {
  rowsTN = rowsTN.filter((r) => tnRowHasText(r));
  if (!rowsTN.length) rowsTN.push(emptyTnRow());
}

const UNDO_STACK_MAX = 60;
let undoStack = [];
let redoStack = [];
let skipUndoRecording = false;
let cellUndoPrime = null;
let cellUndoCommitted = false;

function syncActiveListIntoListSets() {
  if (!activeListId || !listSets.length) return;
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  if (listSubTab === "extern") {
    cur.rowsExtern = rowsExtern;
  } else if (listSubTab === "tn") {
    cur.rowsTN = rowsTN;
  } else if (typeof listSubTab === "string" && listSubTab.startsWith("extra:")) {
    const eid = listSubTab.slice(6);
    const el = (cur.extraLists || []).find((x) => x.id === eid);
    if (el) el.rows = rowsExtern;
  }
}

function loadRowsForCurrentSubTab() {
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  if (!cur.extraLists) cur.extraLists = [];
  if (listSubTab === "extern") {
    rowsExtern = cur.rowsExtern;
  } else if (listSubTab === "tn") {
    rowsTN = cur.rowsTN;
  } else if (typeof listSubTab === "string" && listSubTab.startsWith("extra:")) {
    const eid = listSubTab.slice(6);
    const el = cur.extraLists.find((x) => x.id === eid);
    if (el) rowsExtern = el.rows;
    else if (cur.useExtern !== false) {
      listSubTab = "extern";
      rowsExtern = cur.rowsExtern;
    } else if (cur.useTn !== false) {
      listSubTab = "tn";
      rowsTN = cur.rowsTN;
    } else {
      listSubTab = "extern";
      rowsExtern = cur.rowsExtern;
    }
  }
}

function captureUndoState() {
  syncActiveListIntoListSets();
  const filt = {};
  FILTER_KEYS.forEach((k) => {
    filt[k] = filters[k] || "";
  });
  const cur = listSets.find((ls) => ls.id === activeListId);
  const sl = cur ? normalizeSublistFlags(cur) : { useExtern: true, useTn: true };
  const extraListsSnap = cur && cur.extraLists ? JSON.parse(JSON.stringify(cur.extraLists)) : [];
  return JSON.stringify({
    rowsExtern,
    rowsTN,
    eventSettings,
    filters: filt,
    selectedIds: Array.from(selectedIds),
    listSubTab,
    currentTab,
    useExtern: sl.useExtern,
    useTn: sl.useTn,
    extraLists: extraListsSnap,
    appTitle,
  });
}

function applyUndoState(jsonStr) {
  const o = JSON.parse(jsonStr);
  rowsExtern = (o.rowsExtern || []).map(migrateRow);
  rowsTN = (o.rowsTN || []).map(migrateTnRow);
  if (o.eventSettings && typeof o.eventSettings === "object") {
    eventSettings = { ...eventSettings, ...o.eventSettings };
  }
  if (o.filters && typeof o.filters === "object") {
    FILTER_KEYS.forEach((k) => {
      filters[k] = o.filters[k] ?? "";
    });
  }
  selectedIds.clear();
  (o.selectedIds || []).forEach((id) => {
    if (rowsExtern.some((r) => r.id === id)) selectedIds.add(id);
  });
  currentTab = o.currentTab === "serienbrief" ? "serienbrief" : "liste";
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (cur) {
    if (typeof o.useExtern === "boolean") cur.useExtern = o.useExtern;
    if (typeof o.useTn === "boolean") cur.useTn = o.useTn;
    const fixed = normalizeSublistFlags(cur);
    cur.useExtern = fixed.useExtern;
    cur.useTn = fixed.useTn;
    if (Array.isArray(o.extraLists)) {
      cur.extraLists = o.extraLists.map((e) => ({
        id: String(e.id || "").trim() || crypto.randomUUID(),
        name: String(e.name || "Liste").trim().slice(0, 96) || "Liste",
        rows: (e.rows || []).map(migrateRow),
      }));
    } else if (!cur.extraLists) cur.extraLists = [];
  }
  listSubTab = normalizeListSubTabFromStorage(o.listSubTab);
  if ("appTitle" in o && typeof o.appTitle === "string") {
    appTitle = normalizeAppTitle(o.appTitle);
  }
  ensureValidListSubTab();
  if (typeof listSubTab === "string" && listSubTab.startsWith("extra:") && cur) {
    const eid = listSubTab.slice(6);
    const el = (cur.extraLists || []).find((x) => x.id === eid);
    if (el) rowsExtern = el.rows;
  }
  syncActiveListIntoListSets();
}

function updateUndoRedoUi() {
  const ub = document.getElementById("btn-undo");
  const rb = document.getElementById("btn-redo");
  if (ub) ub.disabled = undoStack.length === 0;
  if (rb) rb.disabled = redoStack.length === 0;
}

function pushSnapshotToUndo(snapStr) {
  if (skipUndoRecording) return;
  if (undoStack.length && undoStack[undoStack.length - 1] === snapStr) return;
  undoStack.push(snapStr);
  if (undoStack.length > UNDO_STACK_MAX) undoStack.shift();
  redoStack.length = 0;
  updateUndoRedoUi();
}

function pushUndoBeforeMutation() {
  pushSnapshotToUndo(captureUndoState());
}

function onCellUndoInputCommit() {
  if (!cellUndoCommitted && cellUndoPrime) {
    pushSnapshotToUndo(cellUndoPrime);
    cellUndoCommitted = true;
  }
}

function undoLastAction() {
  if (!undoStack.length) return;
  const cur = captureUndoState();
  const prev = undoStack.pop();
  redoStack.push(cur);
  skipUndoRecording = true;
  applyUndoState(prev);
  applyAppTitleToDom();
  saveAll();
  loadEventFieldsToDom();
  switchTab(currentTab);
  switchListSubTab(listSubTab);
  skipUndoRecording = false;
  setStatus("Rückgängig.");
  updateUndoRedoUi();
}

function redoLastAction() {
  if (!redoStack.length) return;
  const cur = captureUndoState();
  const next = redoStack.pop();
  undoStack.push(cur);
  skipUndoRecording = true;
  applyUndoState(next);
  applyAppTitleToDom();
  saveAll();
  loadEventFieldsToDom();
  switchTab(currentTab);
  switchListSubTab(listSubTab);
  skipUndoRecording = false;
  setStatus("Wiederholt.");
  updateUndoRedoUi();
}

function undoRedoTargetIsTextField(el) {
  if (!el) return false;
  if (el instanceof HTMLTextAreaElement) return true;
  if (el instanceof HTMLInputElement) {
    const t = el.type;
    if (["checkbox", "radio", "button", "submit", "reset", "file", "hidden"].includes(t)) return false;
    return true;
  }
  return false;
}

function onGlobalUndoRedoKeydown(e) {
  const mod = e.ctrlKey || e.metaKey;
  if (!mod) return;
  if (undoRedoTargetIsTextField(e.target)) return;
  const k = e.key;
  if ((k === "z" || k === "Z") && !e.shiftKey) {
    if (undoStack.length === 0) return;
    e.preventDefault();
    undoLastAction();
    return;
  }
  if (k === "y" || k === "Y") {
    if (redoStack.length === 0) return;
    e.preventDefault();
    redoLastAction();
    return;
  }
  if ((k === "z" || k === "Z") && e.shiftKey) {
    if (redoStack.length === 0) return;
    e.preventDefault();
    redoLastAction();
  }
}

function lfdForIndex(i) {
  let n = 0;
  for (let j = 0; j <= i; j++) {
    const r = rowsExtern[j];
    if (rowStatus(r) || rowHasText(r)) n += 1;
  }
  const r = rowsExtern[i];
  if (!rowStatus(r) && !rowHasText(r)) return "";
  return String(n);
}

/** Index der Zeile in der externen Liste (für Lfd / Serienbrief-Platzhalter). */
function externIndexOfRow(r) {
  return rowsExtern.findIndex((row) => row.id === r.id);
}

const DEFAULT_STATUS_TOAST_MS = 4000;

let _toastTtlTimer = 0;
let _toastFallbackClearTimer = 0;
let _toastGen = 0;

function _toastEls() {
  const root = document.getElementById("app-toast");
  if (!root) return null;
  const text = root.querySelector(".app-toast__text");
  if (!text) return null;
  return { root, text };
}

function _hideToastAfterTransition(hideGen) {
  const els = _toastEls();
  if (!els) return;
  const { root, text } = els;
  const done = (e) => {
    if (e && e.propertyName && e.propertyName !== "opacity") return;
    root.removeEventListener("transitionend", done);
    if (hideGen !== _toastGen) return;
    text.textContent = "";
  };
  root.addEventListener("transitionend", done);
  if (_toastFallbackClearTimer) clearTimeout(_toastFallbackClearTimer);
  _toastFallbackClearTimer = setTimeout(() => {
    root.removeEventListener("transitionend", done);
    if (hideGen === _toastGen) text.textContent = "";
    _toastFallbackClearTimer = 0;
  }, 500);
}

function showStatusToast(msg, opts) {
  const els = _toastEls();
  if (!els) return;
  const { root, text } = els;
  _toastGen += 1;
  const showGen = _toastGen;

  const variant = opts && opts.variant === "serien" ? "serien" : "liste";
  root.classList.toggle("app-toast--serien", variant === "serien");

  const trimmed = (msg || "").trim();
  if (_toastTtlTimer) {
    clearTimeout(_toastTtlTimer);
    _toastTtlTimer = 0;
  }
  if (_toastFallbackClearTimer) {
    clearTimeout(_toastFallbackClearTimer);
    _toastFallbackClearTimer = 0;
  }

  if (!trimmed) {
    root.classList.remove("app-toast--visible");
    text.textContent = "";
    return;
  }

  let ttlMs = DEFAULT_STATUS_TOAST_MS;
  if (opts && typeof opts.ttlMs === "number" && opts.ttlMs > 0) ttlMs = opts.ttlMs;

  text.textContent = trimmed;
  root.classList.remove("app-toast--visible");
  void root.offsetWidth;
  requestAnimationFrame(() => {
    if (showGen !== _toastGen) return;
    root.classList.add("app-toast--visible");
  });

  _toastTtlTimer = setTimeout(() => {
    _toastTtlTimer = 0;
    if (showGen !== _toastGen) return;
    root.classList.remove("app-toast--visible");
    _hideToastAfterTransition(showGen);
  }, ttlMs);
}

function setStatus(msg, opts) {
  showStatusToast(msg, { ...(opts || {}), variant: "liste" });
}

function setStatusSerien(msg, opts) {
  showStatusToast(msg, { ...(opts || {}), variant: "serien" });
}

function yieldToUi() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => setTimeout(resolve, 0));
  });
}

/**
 * @param {number} total
 * @param {{ unit?: "empfänger" | "pdf"; title?: string }} [opts]
 */
function openExportProgressDialog(total, opts) {
  const dlg = document.getElementById("dlg-export-progress");
  const bar = document.getElementById("export-progress-bar");
  const lbl = document.getElementById("export-progress-text");
  const titleEl = document.getElementById("export-progress-title");
  if (!dlg || !bar || !lbl) return;
  const o = opts || {};
  const isPdf = o.unit === "pdf";
  const unit = isPdf ? "PDF" : "Empfänger";
  const prep = isPdf ? "PDF wird vorbereitet…" : "Export wird vorbereitet…";
  if (titleEl) titleEl.textContent = o.title || "Export läuft";
  const n = Math.max(1, total);
  bar.max = n;
  bar.value = 0;
  lbl.textContent = total <= 1 ? prep : `0 / ${total} ${unit}`;
  dlg.showModal();
}

/**
 * @param {number} done
 * @param {number} total
 * @param {string} [detail] — z. B. ZIP-Fortschritt; sonst Zähler nach unit
 * @param {{ unit?: "empfänger" | "pdf" }} [opts]
 */
function updateExportProgressDialog(done, total, detail, opts) {
  const bar = document.getElementById("export-progress-bar");
  const lbl = document.getElementById("export-progress-text");
  const o = opts || {};
  const unit = o.unit === "pdf" ? "PDF" : "Empfänger";
  const n = Math.max(1, total);
  if (bar) {
    bar.max = n;
    bar.value = Math.min(done, n);
  }
  if (lbl) {
    lbl.textContent = detail || (total <= 1 ? "Bitte warten…" : `${done} / ${total} ${unit}`);
  }
}

function closeExportProgressDialog() {
  const titleEl = document.getElementById("export-progress-title");
  if (titleEl) titleEl.textContent = "Export läuft";
  document.getElementById("dlg-export-progress")?.close();
}

/**
 * Nach Word-Export in einen Ordner: Hinweis, dass ein druckidentisches PDF nur über Microsoft Word
 * möglich ist (Browser kann Word-Layout nicht 1:1 nachbilden).
 */
function showPdfWordHinweisDialog() {
  return new Promise((resolve) => {
    const dlg = document.getElementById("dlg-pdf-word-hinweis");
    const btn = document.getElementById("dlg-pdf-word-hinweis-ok");
    if (!dlg || !btn) {
      resolve();
      return;
    }
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      btn.removeEventListener("click", finish);
      dlg.removeEventListener("cancel", onCancel);
      dlg.close();
      resolve();
    };
    const onCancel = () => finish();
    btn.addEventListener("click", finish);
    dlg.addEventListener("cancel", onCancel);
    dlg.showModal();
  });
}

async function maybeShowPdfWordHinweisNachExport(nDoc, wantWord) {
  if (!wantWord || nDoc <= 0) return;
  await showPdfWordHinweisDialog();
}

function buildStoragePayload() {
  syncActiveListIntoListSets();
  return JSON.stringify({
    version: STORAGE_VERSION,
    listSets,
    activeListId,
    listSubTab,
    eventSettings,
    appTitle,
  });
}

function writeStorageNow() {
  try {
    const payload = buildStoragePayload();
    localStorage.setItem(STORAGE_KEY, payload);
    localStorage.setItem(FAILSAFE_KEY, payload);
  } catch (e) {
    console.warn(e);
  }
}

let _saveTimer = 0;
let _savePending = false;
/** Debounced save used for high-frequency edits (Zellen, Status, Ereignisfelder). */
function scheduleSave(delayMs) {
  _savePending = true;
  if (_saveTimer) return;
  _saveTimer = setTimeout(() => {
    _saveTimer = 0;
    if (_savePending) {
      _savePending = false;
      writeStorageNow();
    }
  }, typeof delayMs === "number" ? delayMs : 350);
}

function flushScheduledSave() {
  if (_saveTimer) {
    clearTimeout(_saveTimer);
    _saveTimer = 0;
  }
  if (_savePending) {
    _savePending = false;
    writeStorageNow();
  }
}

/** @param {string} [statusMsg] Meldung im Toast; Standard „Gespeichert.“ */
function saveAll(statusMsg) {
  flushScheduledSave();
  writeStorageNow();
  setStatus(statusMsg ?? "Gespeichert.", { ttlMs: 1200 });
}

function saveEventOnly() {
  try {
    localStorage.setItem(EVENT_KEY, JSON.stringify(eventSettings));
  } catch (_) {}
  scheduleSave();
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.listSets) && parsed.listSets.length) {
        listSets = parsed.listSets.map((ls) => listSetFromStorage(ls));
        activeListId =
          parsed.activeListId && listSets.some((l) => l.id === parsed.activeListId)
            ? parsed.activeListId
            : listSets[0].id;
        if (typeof parsed.appTitle === "string") appTitle = normalizeAppTitle(parsed.appTitle);
        if (parsed.eventSettings && typeof parsed.eventSettings === "object") {
          eventSettings = { ...eventSettings, ...parsed.eventSettings };
        }
        listSubTab = normalizeListSubTabFromStorage(parsed.listSubTab);
        ensureValidListSubTab();
        loadRowsForCurrentSubTab();
        return;
      }
      if (parsed && Array.isArray(parsed.rowsExtern)) {
        rowsExtern = parsed.rowsExtern.map(migrateRow);
        rowsTN = (parsed.rowsTN && parsed.rowsTN.length ? parsed.rowsTN : [emptyTnRow()]).map(migrateTnRow);
        if (parsed.eventSettings && typeof parsed.eventSettings === "object") {
          eventSettings = { ...eventSettings, ...parsed.eventSettings };
        }
        if (parsed.listSubTab === "tn") listSubTab = "tn";
        else listSubTab = "extern";
        const id = crypto.randomUUID();
        listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
        activeListId = id;
        saveAll();
        return;
      }
      if (parsed && Array.isArray(parsed.rows) && parsed.rows.length) {
        rowsExtern = parsed.rows.map(migrateRow);
        rowsTN = [emptyTnRow()];
        if (parsed.eventSettings && typeof parsed.eventSettings === "object") {
          eventSettings = { ...eventSettings, ...parsed.eventSettings };
        }
        const id = crypto.randomUUID();
        listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
        activeListId = id;
        saveAll();
        return;
      }
      if (Array.isArray(parsed) && parsed.length) {
        rowsExtern = parsed.map(migrateRow);
        rowsTN = [emptyTnRow()];
        const id = crypto.randomUUID();
        listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
        activeListId = id;
        saveAll();
        return;
      }
    }
  } catch (_) {}
  try {
    const legacy = localStorage.getItem("bwk-adressliste-json-v2");
    if (legacy) {
      const parsed = JSON.parse(legacy);
      if (Array.isArray(parsed) && parsed.length) {
        rowsExtern = parsed.map(migrateRow);
        rowsTN = [emptyTnRow()];
        const id = crypto.randomUUID();
        listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
        activeListId = id;
        saveAll();
        return;
      }
      if (parsed && Array.isArray(parsed.rows)) {
        rowsExtern = parsed.rows.map(migrateRow);
        rowsTN = [emptyTnRow()];
        const id = crypto.randomUUID();
        listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
        activeListId = id;
        saveAll();
        return;
      }
    }
  } catch (_) {}
  try {
    const ev = localStorage.getItem(EVENT_KEY);
    if (ev) eventSettings = { ...eventSettings, ...JSON.parse(ev) };
  } catch (_) {}
  rowsExtern = [emptyRow()];
  rowsTN = [emptyTnRow()];
  const id = crypto.randomUUID();
  listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
  activeListId = id;
}

function loadEventFieldsToDom() {
  document.getElementById("ev-datum").value = eventSettings.datum || "";
  document.getElementById("ev-zeit").value = eventSettings.zeit || "";
  document.getElementById("ev-ort").value = eventSettings.ort || "";
  document.getElementById("ev-termin-zeile").value = eventSettings.terminZeile || "";
  const bd = document.getElementById("ev-briefdatum");
  if (bd) bd.value = eventSettings.briefdatum || "";
  const sep = document.getElementById("serien-export-prefix");
  if (sep) sep.value = eventSettings.serienExportPrefix || "";
  const subj = document.getElementById("serien-mail-subject");
  if (subj) subj.value = eventSettings.serienMailSubject || "";
  const sfrom = document.getElementById("serien-mail-from");
  if (sfrom) sfrom.value = eventSettings.serienMailFrom || "";
  const mode = eventSettings.serienExportWhat || "both";
  document.querySelectorAll('input[name="serien-export-what"]').forEach((el) => {
    el.checked = el.value === mode;
  });
}

function readEventFieldsFromDom() {
  eventSettings = {
    ...eventSettings,
    datum: document.getElementById("ev-datum").value.trim(),
    zeit: document.getElementById("ev-zeit").value.trim(),
    ort: document.getElementById("ev-ort").value.trim(),
    terminZeile: document.getElementById("ev-termin-zeile").value.trim(),
    briefdatum: document.getElementById("ev-briefdatum")?.value.trim() ?? "",
    serienExportPrefix: document.getElementById("serien-export-prefix")?.value.trim() ?? "",
    serienMailSubject: document.getElementById("serien-mail-subject")?.value ?? "",
    serienMailFrom: document.getElementById("serien-mail-from")?.value.trim() ?? "",
    serienExportWhat: document.querySelector('input[name="serien-export-what"]:checked')?.value || "both",
  };
}

function applyImportedPayload(data, msg) {
  if (data && Array.isArray(data.listSets) && data.listSets.length) {
    listSets = data.listSets.map((ls) => listSetFromImport(ls));
    if (!listSets.length) {
      listSets = [
        {
          id: crypto.randomUUID(),
          name: "Hauptliste",
          useExtern: true,
          useTn: true,
          extraLists: [],
          rowsExtern: [emptyRow()],
          rowsTN: [emptyTnRow()],
        },
      ];
    }
    activeListId =
      data.activeListId && listSets.some((l) => l.id === data.activeListId)
        ? data.activeListId
        : listSets[0].id;
    if (typeof data.appTitle === "string") appTitle = normalizeAppTitle(data.appTitle);
    if (data.eventSettings) eventSettings = { ...eventSettings, ...data.eventSettings };
    listSubTab = normalizeListSubTabFromStorage(data.listSubTab);
    ensureValidListSubTab();
    loadRowsForCurrentSubTab();
    trimEmptyExternRows();
    trimEmptyTnRows();
    selectedIds.clear();
    undoStack.length = 0;
    redoStack.length = 0;
    updateUndoRedoUi();
    applyAppTitleToDom();
    saveAll();
    loadEventFieldsToDom();
    switchListSubTab(listSubTab);
    setStatus(msg || "Import abgeschlossen.");
    maybePromptDuplicatesAfterJsonImport(msg || "Import abgeschlossen.");
    return;
  }
  if (data && Array.isArray(data.rowsExtern)) {
    rowsExtern = data.rowsExtern.map(migrateRow);
    rowsTN = (data.rowsTN || []).map(migrateTnRow);
    if (!rowsTN.length) rowsTN.push(emptyTnRow());
    trimEmptyExternRows();
    trimEmptyTnRows();
    if (data.eventSettings) eventSettings = { ...eventSettings, ...data.eventSettings };
    const id = crypto.randomUUID();
    listSets = [{ id, name: "Import", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
    activeListId = id;
    selectedIds.clear();
    undoStack.length = 0;
    redoStack.length = 0;
    updateUndoRedoUi();
    saveAll();
    loadEventFieldsToDom();
    listSubTab = "extern";
    switchListSubTab(listSubTab);
    setStatus(msg || "Import abgeschlossen.");
    maybePromptDuplicatesAfterJsonImport(msg || "Import abgeschlossen.");
    return;
  }
  if (data && Array.isArray(data.rows)) {
    rowsExtern = data.rows.map(migrateRow);
    rowsTN = (data.rowsTN || [emptyTnRow()]).map(migrateTnRow);
    trimEmptyExternRows();
    trimEmptyTnRows();
    if (data.eventSettings) eventSettings = { ...eventSettings, ...data.eventSettings };
    const id = crypto.randomUUID();
    listSets = [{ id, name: "Import", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
    activeListId = id;
    selectedIds.clear();
    undoStack.length = 0;
    redoStack.length = 0;
    updateUndoRedoUi();
    saveAll();
    loadEventFieldsToDom();
    listSubTab = "extern";
    switchListSubTab(listSubTab);
    setStatus(msg || "Import abgeschlossen.");
    maybePromptDuplicatesAfterJsonImport(msg || "Import abgeschlossen.");
    return;
  }
  if (Array.isArray(data)) {
    rowsExtern = data.map(migrateRow);
    rowsTN = [emptyTnRow()];
    trimEmptyExternRows();
    trimEmptyTnRows();
    const id = crypto.randomUUID();
    listSets = [{ id, name: "Import", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
    activeListId = id;
    selectedIds.clear();
    undoStack.length = 0;
    redoStack.length = 0;
    updateUndoRedoUi();
    saveAll();
    listSubTab = "extern";
    switchListSubTab(listSubTab);
    setStatus(msg || "Import abgeschlossen.");
    maybePromptDuplicatesAfterJsonImport(msg || "Import abgeschlossen.");
  }
}

function parseJsonFile(text) {
  const data = JSON.parse(text);
  if (data && Array.isArray(data.listSets) && data.listSets.length) return data;
  if (data && Array.isArray(data.rowsExtern)) return data;
  if (data && Array.isArray(data.rows)) return data;
  if (Array.isArray(data)) return { version: 2, rows: data };
  throw new Error("Ungültiges Format (erwarte { rows: [...] } oder { rowsExtern, rowsTN } oder mehrere Listen).");
}

async function loadPlzMap() {
  try {
    const res = await fetch("plz-berlin-brandenburg.json", { cache: "force-cache" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    plzMap = await res.json();
    plzLoaded = true;
  } catch (e) {
    console.warn("PLZ-Daten nicht geladen:", e);
    plzMap = {};
    plzLoaded = false;
  }
}

function lookupStadt(plz) {
  const p = String(plz || "").replace(/\s/g, "");
  if (p.length !== 5 || !/^\d{5}$/.test(p)) return "";
  return plzMap[p] || "";
}

/** Kleinschreibung + Umlaute vereinheitlichen (nur für Vergleich) */
function normStadtCompareKey(s) {
  let t = String(s ?? "")
    .normalize("NFC")
    .replace(/\s+/g, " ")
    .trim();
  t = t.split(",")[0].trim();
  const paren = t.indexOf("(");
  if (paren > 0) t = t.slice(0, paren).trim();
  t = t.toLowerCase();
  return t.replace(/ä/g, "a").replace(/ö/g, "o").replace(/ü/g, "u").replace(/ß/g, "ss");
}

function levenshteinDistance(a, b) {
  if (a === b) return 0;
  const al = a.length;
  const bl = b.length;
  if (!al) return bl;
  if (!bl) return al;
  const m = new Array(al + 1);
  for (let i = 0; i <= al; i++) {
    m[i] = new Uint16Array(bl + 1);
    m[i][0] = i;
  }
  for (let j = 0; j <= bl; j++) m[0][j] = j;
  for (let i = 1; i <= al; i++) {
    for (let j = 1; j <= bl; j++) {
      const cost = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1;
      const x = m[i - 1][j] + 1;
      const y = m[i][j - 1] + 1;
      const z = m[i - 1][j - 1] + cost;
      m[i][j] = x < y ? (x < z ? x : z) : y < z ? y : z;
    }
  }
  return m[al][bl];
}

/**
 * Importierte Ortsbezeichnung an PLZ-Katalog angleichen (Tippfehler, „Berli“ statt „Berlin“, etc.).
 * Nur wenn PLZ in plz-berlin-brandenburg.json bekannt ist.
 */
function resolveStadtFromPlzAndOrt(plz, rawStadt) {
  const raw = String(rawStadt ?? "").replace(/\s+/g, " ").trim();
  const p = String(plz ?? "").replace(/\s/g, "");
  const canonical = lookupStadt(p);
  if (!canonical) return raw;
  if (!raw) return canonical;

  const a = normStadtCompareKey(raw);
  const b = normStadtCompareKey(canonical);
  if (a === b) return canonical;

  if (a.length >= 4 && b.startsWith(a)) return canonical;
  if (b.length >= 4 && a.startsWith(b)) return canonical;

  const dist = levenshteinDistance(a, b);
  const maxLen = Math.max(a.length, b.length);
  let maxDist = 2;
  if (maxLen <= 6) maxDist = 1;
  else if (maxLen <= 14) maxDist = 2;
  else maxDist = 3;

  if (dist <= maxDist) return canonical;
  return raw;
}

const FILTER_STATUS_OPTIONS = [
  { value: "", label: "∗" },
  { value: "__empty__", label: "— (leer)" },
  { value: "einladen", label: "Einladen" },
  { value: "offen", label: "Offen" },
  { value: "absage", label: "Absage" },
  { value: "zusage", label: "Zusage" },
];

function uniqueValuesForColumn(key) {
  const set = new Set();
  let hasEmpty = false;
  for (const row of rowsExtern) {
    const v = key === "status" ? rowStatus(row) : String(row[key] ?? "");
    const t = v.trim();
    if (!t) hasEmpty = true;
    else set.add(t);
  }
  return {
    values: Array.from(set).sort((a, b) => a.localeCompare(b, "de")),
    hasEmpty,
  };
}

function matchesFilters(r) {
  for (const key of FILTER_KEYS) {
    const q = filters[key];
    if (q === undefined || q === null || q === "") continue;
    if (key === "status") {
      if (q === "__empty__") {
        if (rowStatus(r) !== "") return false;
      } else if (rowStatus(r) !== q) return false;
      continue;
    }
    const val = String(r[key] ?? "").trim();
    if (q === "__empty__") {
      if (val !== "") return false;
    } else if (val !== q) return false;
  }
  return true;
}

function filterOptionLabel(key, raw) {
  const v = String(raw ?? "");
  if (key === "abteilung") return abteilungTextFuerAnzeige(v, 58);
  return v.length > 48 ? v.slice(0, 45) + "…" : v;
}

function appendFilterSelect(th, key) {
  const sel = document.createElement("select");
  sel.className = "filter-select";
  sel.dataset.filterKey = key;
  sel.title = "Filter wählen";
  sel.setAttribute("aria-label", "Filter " + key);

  if (key === "status") {
    FILTER_STATUS_OPTIONS.forEach(({ value, label: lbl }) => {
      const o = document.createElement("option");
      o.value = value;
      o.textContent = lbl;
      sel.appendChild(o);
    });
  } else {
    const o0 = document.createElement("option");
    o0.value = "";
    o0.textContent = "∗";
    sel.appendChild(o0);
    const { values, hasEmpty } = uniqueValuesForColumn(key);
    if (hasEmpty) {
      const oe = document.createElement("option");
      oe.value = "__empty__";
      oe.textContent = "(leer)";
      sel.appendChild(oe);
    }
    values.forEach((v) => {
      const o = document.createElement("option");
      o.value = v;
      const lbl = filterOptionLabel(key, v);
      o.textContent = lbl;
      if (lbl !== v) o.title = v;
      sel.appendChild(o);
    });
  }

  const fv = filters[key] || "";
  if (fv && !Array.from(sel.options).some((opt) => opt.value === fv)) {
    const o = document.createElement("option");
    o.value = fv;
    const lbl = filterOptionLabel(key, fv);
    o.textContent = lbl;
    if (lbl !== fv) o.title = fv;
    sel.appendChild(o);
  }
  sel.value = fv;

  const syncFilterSelectTitle = () => {
    if (key === "status") {
      sel.title = "Filter wählen";
      return;
    }
    const opt = sel.options[sel.selectedIndex];
    if (!opt || opt.value === "" || opt.value === "__empty__") {
      sel.title = "Filter wählen";
      return;
    }
    sel.title = String(opt.value);
  };
  syncFilterSelectTitle();

  sel.addEventListener("change", () => {
    pushUndoBeforeMutation();
    filters[key] = sel.value;
    syncFilterSelectTitle();
    renderListeTable();
    renderSerienTable();
  });
  th.appendChild(sel);
}

function getFilteredIndices() {
  const out = [];
  for (let i = 0; i < rowsExtern.length; i++) {
    if (matchesFilters(rowsExtern[i])) out.push(i);
  }
  return out;
}

function briefanrede(r) {
  const an = (r.anrede || "").trim();
  const ti = (r.titel || "").trim();
  const vn = (r.vorname || "").trim();
  const nn = (r.nachname || "").trim();
  const titelPart = ti ? `${ti} ` : "";
  if (/^frau/i.test(an)) return `Sehr geehrte Frau ${titelPart}${nn}`.trim();
  if (/^herr/i.test(an)) return `Sehr geehrter Herr ${titelPart}${nn}`.trim();
  if (an) return `Sehr geehrte ${an} ${titelPart}${nn}`.trim();
  if (nn) return `Sehr geehrte Damen und Herren`;
  return "";
}

/** Wie Briefanrede, aber mit Vor- und Nachnamen (falls vorhanden) — für vollständige persönliche Anrede */
function briefanredeMitVorname(r) {
  const an = (r.anrede || "").trim();
  const ti = (r.titel || "").trim();
  const vn = (r.vorname || "").trim();
  const nn = (r.nachname || "").trim();
  const titelPart = ti ? `${ti} ` : "";
  const namePart = [vn, nn].filter(Boolean).join(" ").trim() || nn;
  if (/^frau/i.test(an)) return `Sehr geehrte Frau ${titelPart}${namePart}`.trim();
  if (/^herr/i.test(an)) return `Sehr geehrter Herr ${titelPart}${namePart}`.trim();
  if (an) return `Sehr geehrte ${an} ${titelPart}${namePart}`.trim();
  if (nn) return `Sehr geehrte Damen und Herren`;
  return "";
}

/** Aus Anrede-Feld: „Frau“ / „Herr“ oder leer */
function anredeGeschlechtKurz(r) {
  const an = (r.anrede || "").trim();
  if (/^frau/i.test(an)) return "Frau";
  if (/^herr/i.test(an)) return "Herr";
  return "";
}

/** Ein Feldzeile fürs Fenster: Zeilenumbrüche aus Import → einzeilig, Mehrfach-Leerzeichen weg */
function normalizeBriefkopfFeld(s) {
  return String(s ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\s*\n\s*/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

/**
 * Abteilung aus Behörden-Listen ist oft ein Absatz — für den Fensterbrief auf eine knappe Zeile begrenzen.
 * (Volltext bleibt in der Tabelle / CSV.)
 */
function kuerzeBriefkopfZeile(s, maxLen) {
  const t = normalizeBriefkopfFeld(s);
  if (!t) return "";
  if (t.length <= maxLen) return t;
  return `${t.slice(0, Math.max(0, maxLen - 1)).trimEnd()}…`;
}

/**
 * Abteilung für Anzeige (Filter-Dropdown, Briefkopf): Aussage am Anfang behalten.
 * Reihenfolge: Mittelpunkt · Semikolon; Gedankenstrich –/—; ggf. „ - “; Komma; dann Wortkürzung.
 */
function abteilungTextFuerAnzeige(s, maxLen) {
  const max = typeof maxLen === "number" && maxLen > 12 ? maxLen : 48;
  let t = normalizeBriefkopfFeld(s);
  if (!t) return "";
  if (t.length <= max) return t;

  const shortenIfStillLong = () => {
    if (t.length <= max) return;
    const byDot = t.split(/\s*·\s*/);
    if (byDot.length > 1) {
      const first = byDot[0].trim();
      const restLong = byDot.slice(1).some((x) => String(x).trim().length > 18);
      if (first.length >= 10 || restLong) t = first;
    }
    if (t.length <= max) return;

    const bySemi = t.split(/\s*;\s*/);
    if (bySemi.length > 1) t = bySemi[0].trim();

    if (t.length <= max) return;

    const byDash = t.split(/\s+[–—]\s+/);
    if (byDash.length > 1) {
      const first = byDash[0].trim();
      if (first.length >= 8) t = first;
    }

    if (t.length <= max) return;

    const hy = t.match(/^(.{10,90}?)\s+-\s+(.{20,})$/);
    if (hy) t = hy[1].trim();

    if (t.length <= max) return;

    const byComma = t.split(/\s*,\s*/);
    if (byComma.length > 1) {
      const first = byComma[0].trim();
      if (first.length >= Math.min(28, max - 6)) t = first;
    }

    if (t.length <= max) return;

    let cut = t.slice(0, max - 1);
    const sp = cut.lastIndexOf(" ");
    if (sp > max * 0.45) cut = cut.slice(0, sp);
    t = cut.trimEnd() + "…";
  };

  shortenIfStillLong();
  return t;
}

/** Briefkopf: eine knappe Zeile (gleiche Logik wie Anzeige, etwas kürzer). */
function kompakteAbteilungFuerBriefkopf(s) {
  return abteilungTextFuerAnzeige(s, 48);
}

/** Zeilen für Briefkopf/Fenster wie in der Liste: Firma, ggf. Abteilung, Person, Straße, Ort */
function empfaengerBriefkopfZeilen(r) {
  const lines = [];
  const u = normalizeBriefkopfFeld(r.unternehmen);
  if (u) lines.push(u);
  const abt = kompakteAbteilungFuerBriefkopf(r.abteilung);
  if (abt) lines.push(abt);
  const person = [r.anrede, r.titel, r.vorname, r.nachname]
    .map((x) => normalizeBriefkopfFeld(x))
    .filter(Boolean)
    .join(" ");
  if (person) lines.push(person);
  const s1 = normalizeBriefkopfFeld([r.strasse, r.hausnr].filter(Boolean).join(" "));
  if (s1) lines.push(s1);
  const s2 = normalizeBriefkopfFeld([r.plz, r.stadt].filter(Boolean).join(" "));
  if (s2) lines.push(s2);
  return lines;
}

function padEmpfaengerZeilen(lines, max) {
  const o = {};
  for (let i = 0; i < max; i++) {
    o[`Empfaenger_Zeile${i + 1}`] = lines[i] || "";
  }
  return o;
}

function buildTerminZeileExport() {
  if (eventSettings.terminZeile) return eventSettings.terminZeile;
  const parts = [eventSettings.datum, eventSettings.zeit, eventSettings.ort].filter(Boolean);
  return parts.join(", ");
}

/** Datum aus Anzeigestring → YYYYMMDD für Dateinamen; fallback: heute */
function parseDatumToYyyymmdd(datumStr) {
  const s = String(datumStr || "");
  let m = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return `${m[3]}${m[2].padStart(2, "0")}${m[1].padStart(2, "0")}`;
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}${m[2]}${m[3]}`;
  return "";
}

/** Kürzel für Dateinamen: Fr. / Hr. / leer */
function anredeKurzDatei(r) {
  const an = (r.anrede || "").trim();
  if (/^frau/i.test(an)) return "Fr.";
  if (/^herr/i.test(an)) return "Hr.";
  return "";
}

/** Mindestens eine verwendbare Briefadresse (PLZ + Ort + Straße oder Hausnummer). */
function hasPostalAddressForBrief(r) {
  const plz = String(r.plz || "").replace(/\s/g, "");
  if (plz.length !== 5 || !/^\d{5}$/.test(plz)) return false;
  if (!(r.stadt || "").trim()) return false;
  if (!(r.strasse || "").trim() && !(r.hausnr || "").trim()) return false;
  return true;
}

function hasEmailForBrief(r) {
  return String(r.email || "").trim().includes("@");
}

/**
 * Status „Einladung“: ohne Briefadresse und ohne E-Mail → komplett weglassen.
 * Ist nur eines von beidem da: Word nur bei Adresse, E-Mail nur bei Mail — siehe rowQualifiesForWordBrief / Export-Schleife.
 */
function shouldOmitEinladungOhneAdresseUndMail(r) {
  if (rowStatus(r) !== "einladen") return false;
  return !hasPostalAddressForBrief(r) && !hasEmailForBrief(r);
}

/** Einladung: Word nur mit Briefadresse; andere Stati: Word wie bisher. */
function rowQualifiesForWordBrief(r) {
  if (rowStatus(r) !== "einladen") return true;
  return hasPostalAddressForBrief(r);
}

function mergeRowObject(r) {
  const extIdx = externIndexOfRow(r);
  const lfdListe = extIdx >= 0 ? lfdForIndex(extIdx) : "";
  const terminZeile = buildTerminZeileExport();
  const nameKomplett = [r.vorname, r.nachname].filter(Boolean).join(" ").trim();
  const adr1 = [r.strasse, r.hausnr].filter(Boolean).join(" ").trim();
  const adr2 = [r.plz, r.stadt].filter(Boolean).join(" ").trim();
  const ymd =
    parseDatumToYyyymmdd(eventSettings.datum) || todayStamp().replace(/-/g, "");
  const titel = (r.titel || "").trim();
  const an = (r.anrede || "").trim();
  const nameMitAnredeTitel = [an, titel, r.vorname, r.nachname]
    .map((x) => String(x ?? "").trim())
    .filter(Boolean)
    .join(" ");
  const briefkopfLines = empfaengerBriefkopfZeilen(r);
  const briefkopfBlock = briefkopfLines.join("\n");
  return {
    /** Laufende Nr. wie Spalte „Nr.“ / Lfd in der externen Liste */
    Lfd: lfdListe,
    Lfd_Nr: lfdListe,
    Vorname: r.vorname || "",
    Nachname: r.nachname || "",
    Name_komplett: nameKomplett,
    Anrede: r.anrede || "",
    Titel: r.titel || "",
    Briefanrede: briefanrede(r),
    Briefanrede_mit_Vorname: briefanredeMitVorname(r),
    Name_mit_Anrede_Titel: nameMitAnredeTitel,
    Name_Titel_Nachname: [titel, r.nachname].map((x) => String(x ?? "").trim()).filter(Boolean).join(" "),
    Anrede_Geschlecht: anredeGeschlechtKurz(r),
    Geschlecht: anredeGeschlechtKurz(r),
    Datum_YYYYMMDD: ymd,
    Anrede_Kurz: anredeKurzDatei(r),
    Straße: r.strasse || "",
    Hausnummer: r.hausnr || "",
    PLZ: r.plz || "",
    Ort: r.stadt || "",
    Adresse_Zeile1: adr1,
    Adresse_Zeile2: adr2,
    Briefkopf_Block: briefkopfBlock,
    /** Für Vorlage: {#Briefkopf_zeilen}{.}{/Briefkopf_zeilen} — eine Zeile pro Absatz, Format bleibt stabil */
    Briefkopf_zeilen: briefkopfLines.length ? briefkopfLines.slice() : [""],
    Empfaenger_Adresse_komplett: briefkopfBlock,
    ...padEmpfaengerZeilen(briefkopfLines, 6),
    E_Mail: r.email || "",
    Unternehmen: r.unternehmen || "",
    Abteilung: r.abteilung || "",
    /** Schreibdatum Briefkopf rechts „Berlin, …“ — nicht mit Veranstaltungs-Anzeige verwechseln */
    Briefdatum: (eventSettings.briefdatum || "").trim() || briefdatumGermanToday(),
    /** Alias für ältere/neu benannte Vorlagen (gleicher Wert wie Briefdatum) */
    Ausstellungsdatum: (eventSettings.briefdatum || "").trim() || briefdatumGermanToday(),
    Schreibdatum: (eventSettings.briefdatum || "").trim() || briefdatumGermanToday(),
    Veranstaltung_Datum: eventSettings.datum || "",
    Veranstaltung_Zeit: eventSettings.zeit || "",
    Veranstaltung_Ort: eventSettings.ort || "",
    Veranstaltung_Terminzeile: terminZeile,
    /** Gleicher Stamm wie .docx / .pdf / .eml (ohne Endung); siehe buildSerienExportBasename */
    Export_Dateiname: buildSerienExportBasename(r),
    /** Nur Anrede(kurz)_Titel_ggf_Nachname — gleiche Formatierung wie im Dateinamen */
    Export_Namenszeile_Datei: buildSerienExportNameSuffix(r),
  };
}

const LISTE_HEADERS = [
  "",
  "Nr.",
  "Status",
  "Unternehmen",
  "Anrede",
  "Titel",
  "Vorname",
  "Nachname",
  "E-Mail",
  "Abteilung",
  "Straße",
  "Nr.",
  "PLZ",
  "Stadt",
  "Anmerkungen",
  " ",
];

/** Spalten mit Vorschlagsliste (Werte, die mindestens 2× vorkommen) */
const EXTERN_AUTOCOMPLETE_FIELDS = [
  "unternehmen",
  "anrede",
  "titel",
  "vorname",
  "nachname",
  "email",
  "abteilung",
  "strasse",
  "hausnr",
  "plz",
  "stadt",
  "anmerkungen",
];

const TN_AUTOCOMPLETE_FIELDS = ["anrede", "vorname", "nachname"];

const AUTOCOMPLETE_MIN_COUNT = 2;
const AUTOCOMPLETE_MAX_OPTIONS = 60;

const CELL_ENTER_COMPLETE_TITLE =
  "Eingabetaste: Vervollständigen aus Werten, die in dieser Spalte mindestens zweimal vorkommen. Erneut Enter = nächster Treffer. Tab = zur nächsten Zelle.";

/** @type {WeakMap<HTMLInputElement, { prefix: string; idx: number }>} */
const cellEnterCompleteState = new WeakMap();

/**
 * @param {unknown[]} rows
 * @param {string[]} fields
 * @param {(row: unknown, field: string) => string} getValue
 * @returns {Record<string, string[]>}
 */
function computeAutocompleteSuggestions(rows, fields, getValue) {
  /** @type {Record<string, string[]>} */
  const out = {};
  const counts = new Map();
  fields.forEach((f) => {
    out[f] = [];
    counts.set(f, new Map());
  });
  for (const r of rows) {
    for (const f of fields) {
      const v = String(getValue(r, f) ?? "").trim();
      if (!v) continue;
      const m = counts.get(f);
      m.set(v, (m.get(v) || 0) + 1);
    }
  }
  for (const f of fields) {
    const m = counts.get(f);
    const arr = [];
    for (const [val, c] of m) {
      if (c >= AUTOCOMPLETE_MIN_COUNT) arr.push({ val, c });
    }
    arr.sort((a, b) => b.c - a.c || a.val.localeCompare(b.val, "de"));
    out[f] = arr.slice(0, AUTOCOMPLETE_MAX_OPTIONS).map((x) => x.val);
  }
  return out;
}

/**
 * @param {string} prefix
 * @param {string} field
 * @param {boolean} isTn
 * @returns {string[]}
 */
function getCellCompletionMatches(prefix, field, isTn) {
  const rows = isTn ? rowsTN : rowsExtern;
  const fields = isTn ? TN_AUTOCOMPLETE_FIELDS : EXTERN_AUTOCOMPLETE_FIELDS;
  if (!fields.includes(field)) return [];
  const pre = prefix.trim().toLowerCase();
  if (!pre) return [];
  const sug = computeAutocompleteSuggestions(rows, fields, (r, f) => r[f]);
  const candidates = sug[field] || [];
  return candidates.filter((c) => c.toLowerCase().startsWith(pre));
}

/** Vervollständigung mit Enter (Tab bleibt für Navigation); Umschalt+Enter = neue Zeile (onListeShiftEnter). */
function handleCellEnterComplete(e) {
  if (e.key !== "Enter" || e.shiftKey) return;
  if (currentTab !== "liste") return;
  const t = e.target;
  if (!(t instanceof HTMLInputElement)) return;
  if (!t.classList.contains("cell-input") || !t.dataset.field) return;
  const isTn = t.dataset.tn === "1";
  const field = t.dataset.field;
  const fields = isTn ? TN_AUTOCOMPLETE_FIELDS : EXTERN_AUTOCOMPLETE_FIELDS;
  if (!fields.includes(field)) return;

  const prefix = t.value;
  const matches = getCellCompletionMatches(prefix, field, isTn);
  if (!matches.length) return;

  e.preventDefault();
  e.stopPropagation();

  let st = cellEnterCompleteState.get(t);
  if (!st || st.prefix !== prefix) {
    st = { prefix, idx: 0 };
  } else {
    st = { prefix, idx: (st.idx + 1) % matches.length };
  }
  cellEnterCompleteState.set(t, st);

  pushUndoBeforeMutation();

  const chosen = matches[st.idx];
  t.value = chosen;
  const row = Number(t.dataset.row);
  if (isTn) {
    rowsTN[row][field] = chosen;
    saveAll();
    updateTnLfdCells();
    const rowTr = t.closest("tr");
    if (rowTr) applyTnRowIncompleteUi(rowTr, rowsTN[row]);
  } else {
    rowsExtern[row][field] = chosen;
    saveAll();
    updateLfdCells();
    const rowTr = t.closest("tr");
    if (rowTr) applyExternRowIncompleteUi(rowTr, rowsExtern[row]);
    renderSerienTable();
  }
  cellUndoPrime = captureUndoState();
  cellUndoCommitted = false;
  requestAnimationFrame(() => {
    t.setSelectionRange(chosen.length, chosen.length);
  });
}

function renderListeTable() {
  const thead = document.getElementById("liste-thead");
  const tbody = document.getElementById("liste-tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  LISTE_HEADERS.forEach((h, hi) => {
    const th = document.createElement("th");
    th.textContent = h === " " ? "" : h;
    if (h === " ") th.className = "col-action";
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const trf = document.createElement("tr");
  trf.className = "filter-row";
  LISTE_HEADERS.forEach((h, hi) => {
    const th = document.createElement("th");
    if (h === " ") {
      th.innerHTML = "";
    } else if (hi === 0) {
      th.innerHTML = "";
    } else if (hi === 1) {
      th.className = "filter-corner";
      th.innerHTML =
        '<span class="filter-row-label filter-row-label--sign" title="Dropdown-Filter pro Spalte" aria-label="Filterzeile"></span>';
    } else {
      const key = FILTER_KEYS[hi - 2];
      appendFilterSelect(th, key);
    }
    trf.appendChild(th);
  });
  thead.appendChild(trf);

  const nameDupIndices = indicesWithDuplicateNamePairs(rowsExtern, (r) => rowStatus(r) || rowHasText(r));
  const indices = getFilteredIndices();
  indices.forEach((i) => {
    const r = rowsExtern[i];
    const tr = document.createElement("tr");
    tr.dataset.status = rowStatus(r);
    tr.dataset.rowIndex = String(i);

    const tdIcon = document.createElement("td");
    tdIcon.className = "cell-status-icon";
    tr.appendChild(tdIcon);

    const tdLfd = document.createElement("td");
    tdLfd.className = "cell-lfd";
    tdLfd.textContent = lfdForIndex(i) || "";
    tr.appendChild(tdLfd);

    const tdSt = document.createElement("td");
    tdSt.className = "cell-status";
    const grp = document.createElement("div");
    grp.className = "cell-status-radios status-segmented";
    grp.setAttribute("role", "radiogroup");
    grp.setAttribute("aria-label", "Status");
    const radioName = "liste-status-" + r.id;
    const statusOpts = [
      { value: "einladen", label: "Einladen", title: "Einladen" },
      { value: "offen", label: "Offen", title: "Offen" },
      { value: "absage", label: "Absage", title: "Absage" },
      { value: "zusage", label: "Zusage", title: "Zusage" },
    ];
    statusOpts.forEach(({ value: opt, label: lbl, title: ttl }) => {
      const label = document.createElement("label");
      label.className = "status-segment";
      if (ttl) label.title = ttl;
      const inp = document.createElement("input");
      inp.type = "radio";
      inp.name = radioName;
      inp.value = opt;
      inp.className = "status-segment-input";
      if (rowStatus(r) === opt) inp.checked = true;
      inp.addEventListener("change", () => {
        if (inp.checked) {
          pushUndoBeforeMutation();
          rowsExtern[i].status = opt;
          saveAll();
          const rowTr = inp.closest("tr");
          if (rowTr) {
            rowTr.dataset.status = rowStatus(rowsExtern[i]);
            applyExternRowIncompleteUi(rowTr, rowsExtern[i]);
            const dups = indicesWithDuplicateNamePairs(rowsExtern, (r) => rowStatus(r) || rowHasText(r));
            applyDuplicateNamePairRowUi(rowTr, dups.has(i));
          }
          renderSerienTable();
        }
      });
      bindStatusSegmentRepeatClickClears(label, inp, () => {
        pushUndoBeforeMutation();
        rowsExtern[i].status = "";
        saveAll();
        const rowTr = inp.closest("tr");
        if (rowTr) {
          rowTr.dataset.status = rowStatus(rowsExtern[i]);
          applyExternRowIncompleteUi(rowTr, rowsExtern[i]);
          const dups = indicesWithDuplicateNamePairs(rowsExtern, (r) => rowStatus(r) || rowHasText(r));
          applyDuplicateNamePairRowUi(rowTr, dups.has(i));
        }
        renderSerienTable();
      });
      const span = document.createElement("span");
      span.className = "status-segment-text";
      span.textContent = lbl;
      label.appendChild(inp);
      label.appendChild(span);
      grp.appendChild(label);
    });
    tdSt.appendChild(grp);
    tr.appendChild(tdSt);

    EXTERN_AUTOCOMPLETE_FIELDS.forEach((field) => {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.type = field === "email" ? "email" : "text";
      inp.className = "cell-input";
      inp.value = r[field] || "";
      inp.dataset.row = String(i);
      inp.dataset.field = field;
      inp.setAttribute("autocomplete", "off");
      inp.title = CELL_ENTER_COMPLETE_TITLE;
      if (field === "plz") {
        inp.maxLength = 5;
        inp.inputMode = "numeric";
      }
      inp.addEventListener("focusin", () => {
        cellUndoPrime = captureUndoState();
        cellUndoCommitted = false;
      });
      inp.addEventListener("blur", () => {
        if (field === "plz") onPlzBlurForRow(i);
        else {
          cellUndoPrime = null;
          cellUndoCommitted = false;
        }
        if (field === "vorname" || field === "nachname") render();
      });
      inp.addEventListener("input", () => {
        cellEnterCompleteState.delete(inp);
        onCellUndoInputCommit();
        rowsExtern[i][field] = inp.value;
        scheduleSave();
        updateLfdCells();
        const rowTr = inp.closest("tr");
        if (rowTr) applyExternRowIncompleteUi(rowTr, rowsExtern[i]);
      });
      td.appendChild(inp);
      tr.appendChild(td);
    });

    const tdDel = document.createElement("td");
    tdDel.className = "col-action";
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn-tiny danger";
    btn.textContent = "×";
    btn.title = "Zeile löschen";
    btn.addEventListener("click", () => {
      pushUndoBeforeMutation();
      rowsExtern.splice(i, 1);
      if (!rowsExtern.length) rowsExtern.push(emptyRow());
      selectedIds.clear();
      saveAll();
      render();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    applyExternRowIncompleteUi(tr, r);
    applyDuplicateNamePairRowUi(tr, nameDupIndices.has(i));
    tbody.appendChild(tr);
  });

  if (!indices.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = LISTE_HEADERS.length;
    td.className = "empty-hint";
    td.textContent = "Keine Treffer";
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
}

function renderTnTable() {
  const thead = document.getElementById("tn-thead");
  const tbody = document.getElementById("tn-tbody");
  if (!thead || !tbody) return;
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  ["", "Lfd", "Status", "Anrede", "Vorname", "Nachname", " "].forEach((h, hi) => {
    const th = document.createElement("th");
    th.textContent = h === " " ? "" : h;
    if (h === " ") th.className = "col-action";
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tnNameDupIndices = indicesWithDuplicateNamePairs(rowsTN, (r) => tnRowHasText(r));
  rowsTN.forEach((r, i) => {
    const tr = document.createElement("tr");
    const st = rowStatusTn(r);
    tr.dataset.status = tnRowDatasetStatus(st);
    tr.dataset.rowIndex = String(i);

    const tdIcon = document.createElement("td");
    tdIcon.className = "cell-status-icon";
    tr.appendChild(tdIcon);

    const tdLfd = document.createElement("td");
    tdLfd.className = "cell-lfd";
    tdLfd.textContent = lfdTnForIndex(i) || "";
    tr.appendChild(tdLfd);

    const tdSt = document.createElement("td");
    tdSt.className = "cell-status";
    const grp = document.createElement("div");
    grp.className = "cell-status-radios status-segmented tn-status-radios";
    grp.setAttribute("role", "radiogroup");
    grp.setAttribute("aria-label", "Status");
    const radioName = "tn-status-" + r.id;
    [
      { value: "einladen", label: "Einladung", title: "Einladen" },
      { value: "offen", label: "Offen", title: "Offen" },
      { value: "zusage", label: "Zusage", title: "Zusage" },
      { value: "absage", label: "Absage", title: "Absage" },
    ].forEach(({ value: opt, label: lbl, title: ttl }) => {
      const label = document.createElement("label");
      label.className = "status-segment";
      if (ttl) label.title = ttl;
      const inp = document.createElement("input");
      inp.type = "radio";
      inp.name = radioName;
      inp.value = opt;
      inp.className = "status-segment-input";
      if (st === opt) inp.checked = true;
      inp.addEventListener("change", () => {
        if (inp.checked) {
          pushUndoBeforeMutation();
          rowsTN[i].status = opt;
          saveAll();
          const rowTr = inp.closest("tr");
          if (rowTr) {
            rowTr.dataset.status = tnRowDatasetStatus(rowStatusTn(rowsTN[i]));
            applyTnRowIncompleteUi(rowTr, rowsTN[i]);
            const dups = indicesWithDuplicateNamePairs(rowsTN, (r) => tnRowHasText(r));
            applyDuplicateNamePairRowUi(rowTr, dups.has(i));
          }
        }
      });
      bindStatusSegmentRepeatClickClears(label, inp, () => {
        pushUndoBeforeMutation();
        rowsTN[i].status = "";
        saveAll();
        const rowTr = inp.closest("tr");
        if (rowTr) {
          rowTr.dataset.status = tnRowDatasetStatus(rowStatusTn(rowsTN[i]));
          applyTnRowIncompleteUi(rowTr, rowsTN[i]);
          const dups = indicesWithDuplicateNamePairs(rowsTN, (r) => tnRowHasText(r));
          applyDuplicateNamePairRowUi(rowTr, dups.has(i));
        }
      });
      const span = document.createElement("span");
      span.className = "status-segment-text";
      span.textContent = lbl;
      label.appendChild(inp);
      label.appendChild(span);
      grp.appendChild(label);
    });
    tdSt.appendChild(grp);
    tr.appendChild(tdSt);

    TN_AUTOCOMPLETE_FIELDS.forEach((field) => {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.type = "text";
      inp.className = "cell-input";
      inp.value = r[field] || "";
      inp.dataset.row = String(i);
      inp.dataset.field = field;
      inp.dataset.tn = "1";
      inp.setAttribute("autocomplete", "off");
      inp.title = CELL_ENTER_COMPLETE_TITLE;
      inp.addEventListener("focusin", () => {
        cellUndoPrime = captureUndoState();
        cellUndoCommitted = false;
      });
      inp.addEventListener("blur", () => {
        cellUndoPrime = null;
        cellUndoCommitted = false;
        if (field === "vorname" || field === "nachname") render();
      });
      inp.addEventListener("input", () => {
        cellEnterCompleteState.delete(inp);
        onCellUndoInputCommit();
        rowsTN[i][field] = inp.value;
        scheduleSave();
        updateTnLfdCells();
        const rowTr = inp.closest("tr");
        if (rowTr) applyTnRowIncompleteUi(rowTr, rowsTN[i]);
      });
      td.appendChild(inp);
      tr.appendChild(td);
    });

    const tdDel = document.createElement("td");
    tdDel.className = "col-action";
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn-tiny danger";
    btn.textContent = "×";
    btn.title = "Zeile löschen";
    btn.addEventListener("click", () => {
      pushUndoBeforeMutation();
      rowsTN.splice(i, 1);
      if (!rowsTN.length) rowsTN.push(emptyTnRow());
      saveAll();
      render();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    applyTnRowIncompleteUi(tr, r);
    applyDuplicateNamePairRowUi(tr, tnNameDupIndices.has(i));
    tbody.appendChild(tr);
  });
}

function updateTnLfdCells() {
  const nameDupIndices = indicesWithDuplicateNamePairs(rowsTN, (r) => tnRowHasText(r));
  document.querySelectorAll("#tn-tbody tr[data-row-index]").forEach((tr) => {
    const i = Number(tr.dataset.rowIndex);
    const td = tr.querySelector(".cell-lfd");
    if (td) td.textContent = lfdTnForIndex(i) || "";
    const st = rowStatusTn(rowsTN[i]);
    tr.dataset.status = tnRowDatasetStatus(st);
    applyTnRowIncompleteUi(tr, rowsTN[i]);
    applyDuplicateNamePairRowUi(tr, nameDupIndices.has(i));
  });
}

function updateLfdCells() {
  const nameDupIndices = indicesWithDuplicateNamePairs(rowsExtern, (r) => rowStatus(r) || rowHasText(r));
  document.querySelectorAll("#liste-tbody tr[data-row-index]").forEach((tr) => {
    const i = Number(tr.dataset.rowIndex);
    const td = tr.querySelector(".cell-lfd");
    if (td) td.textContent = lfdForIndex(i) || "";
    tr.dataset.status = rowStatus(rowsExtern[i]);
    applyExternRowIncompleteUi(tr, rowsExtern[i]);
    applyDuplicateNamePairRowUi(tr, nameDupIndices.has(i));
  });
}

function onPlzBlurForRow(i) {
  const r = rowsExtern[i];
  const plz = String(r.plz || "").replace(/\s/g, "");
  cellUndoPrime = null;
  cellUndoCommitted = false;
  if (plz.length !== 5 || !/^\d{5}$/.test(plz)) return;
  const stadt = lookupStadt(plz);
  if (!stadt) return;
  const cur = (r.stadt || "").trim();
  if (!cur || cur === stadt) {
    if ((r.stadt || "").trim() !== stadt) pushUndoBeforeMutation();
    r.stadt = stadt;
    saveAll();
    renderListeTable();
    renderSerienTable();
  }
}

/** Straße + Nr., PLZ + Ort in einer Zeile — Serienbrief-Liste zur Ansicht/Validierung. */
function anschriftSerienExportAnsicht(r) {
  const adr1 = [r.strasse, r.hausnr].filter(Boolean).join(" ").trim();
  const adr2 = [r.plz, r.stadt].filter(Boolean).join(" ").trim();
  const parts = [adr1, adr2].filter(Boolean);
  return parts.length ? parts.join(", ") : "—";
}

function renderSerienTable() {
  readEventFieldsFromDom();
  const tbody = document.getElementById("serien-tbody");
  tbody.innerHTML = "";

  const statusFilter = getSerienStatusFilter();

  const indices = getFilteredIndices().filter((i) => {
    if (!statusFilter) return true;
    return rowStatus(rowsExtern[i]) === statusFilter;
  });

  indices.forEach((i) => {
    const r = rowsExtern[i];
    const tr = document.createElement("tr");
    tr.dataset.rowId = r.id;
    tr.dataset.status = rowStatus(r) || "";
    applyRowIncompleteClass(tr, externRowIsIncomplete(r), false);

    const td0 = document.createElement("td");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedIds.has(r.id);
    cb.addEventListener("change", () => {
      pushUndoBeforeMutation();
      if (cb.checked) selectedIds.add(r.id);
      else selectedIds.delete(r.id);
      setStatusSerien(selectedIds.size + " ausgewählt.");
    });
    td0.appendChild(cb);
    tr.appendChild(td0);

    const name =
      [r.titel, r.vorname, r.nachname].map((x) => String(x || "").trim()).filter(Boolean).join(" ") || "—";

    [lfdForIndex(i) || "", formatExternStatusSerienDisplay(rowStatus(r)), name, r.email || "", anschriftSerienExportAnsicht(r)].forEach((text) => {
      const td = document.createElement("td");
      td.textContent = text;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  if (!indices.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6" class="empty-hint">Keine Treffer</td>`;
    tbody.appendChild(tr);
  }
}

function render() {
  if (listSubTab === "tn") {
    renderTnTable();
  } else {
    renderListeTable();
  }
  renderSerienTable();
  updateUndoRedoUi();
  refreshListSetToolbar();
  refreshSublistTabs();
}

function applyListSubTabPanelUi() {
  const subNow = listSubTab;
  const sheetExtern = document.getElementById("sheet-extern");
  const sheetTn = document.getElementById("sheet-tn");
  const btnClearFilters = document.getElementById("btn-clear-filters");
  const kbdHint = document.getElementById("kbd-hint");

  if (sheetExtern && sheetTn) {
    if (subNow === "tn") {
      sheetTn.classList.remove("hidden");
      sheetTn.hidden = false;
      sheetExtern.classList.add("hidden");
      sheetExtern.hidden = true;
    } else {
      sheetExtern.classList.remove("hidden");
      sheetExtern.hidden = false;
      sheetTn.classList.add("hidden");
      sheetTn.hidden = true;
    }
  }
  if (btnClearFilters) btnClearFilters.hidden = subNow === "tn";
  if (kbdHint) {
    const tip =
      subNow === "tn"
        ? "In einer Zelle der TN-Liste: Umschalt+Eingabe = neue Zeile"
        : "In einer Tabellenzelle oder im Filter: Umschalt+Eingabe = neue Zeile";
    kbdHint.title = tip;
    kbdHint.setAttribute("aria-label", "Tastenkürzel: Umschalttaste und Eingabe für neue Zeile. " + tip);
  }
  render();
}

function switchListSubTab(sub) {
  syncActiveListIntoListSets();
  listSubTab = sub;
  ensureValidListSubTab();
  loadRowsForCurrentSubTab();
  applyListSubTabPanelUi();
}

function switchActiveList(newId) {
  if (newId === activeListId) return;
  if (!listSets.some((ls) => ls.id === newId)) return;
  syncActiveListIntoListSets();
  activeListId = newId;
  ensureValidListSubTab();
  loadRowsForCurrentSubTab();
  selectedIds.clear();
  undoStack.length = 0;
  redoStack.length = 0;
  updateUndoRedoUi();
  saveAll();
  applyListSubTabPanelUi();
  setStatus("Hauptliste gewechselt.");
}

function refreshSublistTabs() {
  const inner = document.getElementById("list-subtabs-inner");
  if (!inner) return;
  syncActiveListIntoListSets();
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  const { useExtern, useTn } = normalizeSublistFlags(cur);
  cur.useExtern = useExtern;
  cur.useTn = useTn;
  if (!cur.extraLists) cur.extraLists = [];
  const canRemoveStd = useExtern && useTn;

  inner.innerHTML = "";

  const restoreWrap = document.createElement("div");
  restoreWrap.className = "list-sublist-restore";
  if (!useExtern) {
    const rb = document.createElement("button");
    rb.type = "button";
    rb.className = "list-sublist-restore-btn";
    rb.textContent = "Externe Liste einblenden";
    rb.addEventListener("click", () => showSublist("extern"));
    restoreWrap.appendChild(rb);
  }
  if (!useTn) {
    const rb = document.createElement("button");
    rb.type = "button";
    rb.className = "list-sublist-restore-btn";
    rb.textContent = "TN-Liste einblenden";
    rb.addEventListener("click", () => showSublist("tn"));
    restoreWrap.appendChild(rb);
  }
  if (restoreWrap.childElementCount) inner.appendChild(restoreWrap);

  function appendSubTab(which) {
    const isExtern = which === "extern";
    const active = listSubTab === which;
    const row = document.createElement("div");
    row.className =
      "list-subtab-wrap" +
      (canRemoveStd ? " has-remove" : "") +
      (active ? " active" : "");
    const b = document.createElement("button");
    b.type = "button";
    b.id = isExtern ? "list-sub-extern" : "list-sub-tn";
    b.className =
      "list-subtab " + (isExtern ? "list-subtab--extern" : "list-subtab--tn") + (active ? " active" : "");
    b.dataset.listSub = which;
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", active ? "true" : "false");
    const exCount = (cur.rowsExtern || []).filter((r) => rowStatus(r) || rowHasText(r)).length;
    const tnCount = (cur.rowsTN || []).filter((r) => tnRowHasText(r)).length;
    const subLine = isExtern
      ? `Kooperationspartner · ${formatEntryCountLabel(exCount)}`
      : `Zusage / Absage · ${formatEntryCountLabel(tnCount)}`;
    b.title = (isExtern ? "Externe Liste" : "TN-Liste") + " — " + subLine;
    const t1 = document.createElement("span");
    t1.className = "list-subtab__title";
    t1.textContent = isExtern ? "Externe Liste" : "TN-Liste";
    const t2 = document.createElement("span");
    t2.className = "list-subtab__sub";
    t2.textContent = subLine;
    b.appendChild(t1);
    b.appendChild(t2);
    row.appendChild(b);
    if (canRemoveStd) {
      const rm = document.createElement("button");
      rm.type = "button";
      rm.className = "list-subtab-remove";
      rm.dataset.sub = which;
      const lab = isExtern ? "Externe Liste" : "TN-Liste";
      rm.title = `${lab} ausblenden (Daten bleiben erhalten)`;
      rm.setAttribute("aria-label", `${lab} ausblenden`);
      rm.textContent = "×";
      row.appendChild(rm);
    }
    inner.appendChild(row);
  }

  function appendExtraTab(entry) {
    const subKey = `extra:${entry.id}`;
    const active = listSubTab === subKey;
    const row = document.createElement("div");
    row.className = "list-subtab-wrap has-remove" + (active ? " active" : "");
    const b = document.createElement("button");
    b.type = "button";
    b.className = "list-subtab list-subtab--extra" + (active ? " active" : "");
    b.dataset.listSub = subKey;
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", active ? "true" : "false");
    const n = (entry.rows || []).filter((r) => rowStatus(r) || rowHasText(r)).length;
    const subLine = `${formatEntryCountLabel(n)}`;
    b.title = `${entry.name} — ${subLine}`;
    const t1 = document.createElement("span");
    t1.className = "list-subtab__title";
    t1.textContent = entry.name;
    const t2 = document.createElement("span");
    t2.className = "list-subtab__sub";
    t2.textContent = subLine;
    b.appendChild(t1);
    b.appendChild(t2);
    row.appendChild(b);
    const rm = document.createElement("button");
    rm.type = "button";
    rm.className = "list-subtab-remove";
    rm.dataset.extraRemove = entry.id;
    rm.title = "Unterliste löschen";
    rm.setAttribute("aria-label", `Unterliste „${entry.name}“ löschen`);
    rm.textContent = "×";
    row.appendChild(rm);
    inner.appendChild(row);
  }

  if (useExtern) appendSubTab("extern");
  if (useTn) appendSubTab("tn");
  cur.extraLists.forEach((e) => appendExtraTab(e));
}

function hideSublist(which) {
  if (which !== "extern" && which !== "tn") return;
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  const { useExtern, useTn } = normalizeSublistFlags(cur);
  if (!useExtern || !useTn) return;
  const label = which === "extern" ? "Externe Liste" : "TN-Liste";
  if (
    !confirm(
      `${label} für diese Hauptliste ausblenden? Die Daten bleiben gespeichert — zum Wiederanzeigen den Button „${label} einblenden“ unter den Unterlisten verwenden.`,
    )
  )
    return;
  pushUndoBeforeMutation();
  if (which === "extern") cur.useExtern = false;
  else cur.useTn = false;
  ensureValidListSubTab();
  saveAll();
  switchListSubTab(listSubTab);
}

function showSublist(which) {
  if (which !== "extern" && which !== "tn") return;
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  pushUndoBeforeMutation();
  if (which === "extern") cur.useExtern = true;
  else cur.useTn = true;
  saveAll();
  switchListSubTab(which);
}

function createNewExtraSublist() {
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur) return;
  if (!cur.extraLists) cur.extraLists = [];
  const n = cur.extraLists.length + 1;
  const defaultName = `Zusatzliste ${n}`;
  const name = prompt("Neue Unterliste anlegen:", defaultName);
  if (name === null) return;
  const trimmed = String(name).trim();
  if (!trimmed) {
    setStatus("Keine Unterliste angelegt (leerer Name).");
    return;
  }
  pushUndoBeforeMutation();
  const id = crypto.randomUUID();
  cur.extraLists.push({ id, name: trimmed.slice(0, 96), rows: [emptyRow()] });
  saveAll();
  switchListSubTab(`extra:${id}`);
  setStatus(`Unterliste „${trimmed}“ angelegt.`);
}

function removeExtraSublist(extraId) {
  const cur = listSets.find((ls) => ls.id === activeListId);
  if (!cur || !cur.extraLists) return;
  const el = cur.extraLists.find((e) => e.id === extraId);
  if (!el) return;
  if (!confirm(`Unterliste „${el.name}“ inkl. aller Zeilen löschen?`)) return;
  pushUndoBeforeMutation();
  cur.extraLists = cur.extraLists.filter((e) => e.id !== extraId);
  if (listSubTab === `extra:${extraId}`) {
    listSubTab = cur.useExtern !== false ? "extern" : cur.useTn !== false ? "tn" : "extern";
    ensureValidListSubTab();
    loadRowsForCurrentSubTab();
  }
  saveAll();
  applyListSubTabPanelUi();
  setStatus("Unterliste gelöscht.");
}

/** Zählt „volle“ Zeilen pro Hauptliste (Extern: Status oder Zelltext; TN: Zelltext). */
function countListSetEntries(ls) {
  const exMain = (ls.rowsExtern || []).filter((r) => rowStatus(r) || rowHasText(r)).length;
  const exExtra = (ls.extraLists || []).reduce((sum, e) => {
    return sum + (e.rows || []).filter((r) => rowStatus(r) || rowHasText(r)).length;
  }, 0);
  const ex = exMain + exExtra;
  const tn = (ls.rowsTN || []).filter((r) => tnRowHasText(r)).length;
  return { ex, exMain, exExtra, tn, total: ex + tn };
}

function formatEntryCountLabel(n) {
  if (n === 1) return "1 Datensatz";
  return `${n} Datensätze`;
}

function refreshListSetToolbar() {
  const wrap = document.getElementById("list-set-buttons");
  if (!wrap) return;
  wrap.innerHTML = "";
  const canDelete = listSets.length > 1;
  listSets.forEach((ls) => {
    const { ex, exMain, exExtra, tn, total } = countListSetEntries(ls);
    const row = document.createElement("div");
    row.className =
      "list-set-tab-wrap" +
      (ls.id === activeListId ? " active" : "") +
      (canDelete ? " has-remove" : "");

    const b = document.createElement("button");
    b.type = "button";
    b.className = "list-subtab list-set-tab";
    b.dataset.listId = ls.id;
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", ls.id === activeListId ? "true" : "false");
    b.title = `${ls.name} — ${total} Datensätze (${exMain} Extern${exExtra ? ` + ${exExtra} Zusatz` : ""}, ${tn} TN)`;
    const body = document.createElement("div");
    body.className = "list-set-tab__body";
    const title = document.createElement("span");
    title.className = "list-subtab__title list-set-tab__title";
    title.textContent = ls.name;
    body.appendChild(title);
    const sub = document.createElement("span");
    sub.className = "list-subtab__sub";
    sub.textContent = formatEntryCountLabel(total);
    body.appendChild(sub);
    b.appendChild(body);
    row.appendChild(b);

    if (canDelete) {
      const rm = document.createElement("button");
      rm.type = "button";
      rm.className = "list-set-tab-remove";
      rm.dataset.listId = ls.id;
      rm.title = "Hauptliste löschen";
      rm.setAttribute("aria-label", `Hauptliste „${ls.name}“ löschen`);
      rm.textContent = "×";
      row.appendChild(rm);
    }

    wrap.appendChild(row);
  });
}

function createNewList() {
  const defaultName = `Hauptliste ${listSets.length + 1}`;
  const name = prompt("Neue Hauptliste anlegen:", defaultName);
  if (name === null) return;
  const trimmed = String(name).trim();
  if (!trimmed) {
    setStatus("Keine neue Hauptliste angelegt (leerer Name).");
    return;
  }
  const ls = {
    id: crypto.randomUUID(),
    name: trimmed.slice(0, 96),
    useExtern: true,
    useTn: true,
    extraLists: [],
    rowsExtern: [emptyRow()],
    rowsTN: [emptyTnRow()],
  };
  listSets.push(ls);
  activeListId = ls.id;
  rowsExtern = ls.rowsExtern;
  rowsTN = ls.rowsTN;
  selectedIds.clear();
  undoStack.length = 0;
  redoStack.length = 0;
  updateUndoRedoUi();
  saveAll();
  switchListSubTab(listSubTab);
  setStatus(`Hauptliste „${trimmed}“ angelegt.`);
}

function deleteListById(id) {
  if (listSets.length <= 1) {
    setStatus("Die letzte Hauptliste kann nicht gelöscht werden.");
    return;
  }
  const cur = listSets.find((ls) => ls.id === id);
  if (!cur) return;
  const name = cur.name || "Hauptliste";
  if (!confirm(`Hauptliste „${name}“ wirklich löschen? Alle Unterlisten (Extern, TN und Zusatzlisten) gehen verloren.`)) return;
  const wasActive = activeListId === id;
  syncActiveListIntoListSets();
  listSets = listSets.filter((ls) => ls.id !== id);
  if (wasActive) {
    activeListId = listSets[0].id;
    selectedIds.clear();
    undoStack.length = 0;
    redoStack.length = 0;
    ensureValidListSubTab();
    loadRowsForCurrentSubTab();
  }
  updateUndoRedoUi();
  saveAll();
  applyListSubTabPanelUi();
  setStatus("Hauptliste gelöscht.");
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function csvCell(v) {
  if (Array.isArray(v)) return csvCell(v.join(" · "));
  const s = String(v ?? "");
  if (/[;\r\n"]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function focusListeRowInput(rowIndex, fieldName) {
  const field = fieldName || "unternehmen";
  requestAnimationFrame(() => {
    const tr = document.querySelector(`#liste-tbody tr[data-row-index="${rowIndex}"]`);
    if (!tr) return;
    const inp = tr.querySelector(`input.cell-input[data-field="${field}"]`);
    if (inp) {
      inp.focus();
      inp.select();
    }
  });
}

/** @param {{ focusField?: string }} [opts] */
function addExternRow(opts) {
  pushUndoBeforeMutation();
  const newIndex = rowsExtern.length;
  rowsExtern.push(emptyRow());
  saveAll("Neue Zeile.");
  render();
  const focusField = (opts && opts.focusField) || "unternehmen";
  focusListeRowInput(newIndex, focusField);
}

function focusTnRowInput(rowIndex, fieldName) {
  const field = fieldName || "vorname";
  requestAnimationFrame(() => {
    const tr = document.querySelector(`#tn-tbody tr[data-row-index="${rowIndex}"]`);
    if (!tr) return;
    const inp = tr.querySelector(`input.cell-input[data-field="${field}"]`);
    if (inp) {
      inp.focus();
      inp.select();
    }
  });
}

function addTnRow(opts) {
  pushUndoBeforeMutation();
  const newIndex = rowsTN.length;
  rowsTN.push(emptyTnRow());
  saveAll("Neue Zeile.");
  render();
  const focusField = (opts && opts.focusField) || "vorname";
  focusTnRowInput(newIndex, focusField);
}

function addRowByContext(opts) {
  if (listSubTab === "tn") addTnRow(opts);
  else addExternRow(opts);
}

function onListeShiftEnter(e) {
  if (!e.shiftKey || e.key !== "Enter") return;
  if (currentTab !== "liste") return;
  const el = e.target;
  if (!(el instanceof HTMLInputElement)) return;

  if (listSubTab === "tn") {
    const tnTable = document.getElementById("tn-table");
    if (!tnTable || !tnTable.contains(el)) return;
    e.preventDefault();
    addTnRow({ focusField: el.type === "radio" ? "vorname" : el.dataset.field || "vorname" });
    return;
  }

  const listeTable = document.getElementById("liste-table");
  if (!listeTable || !listeTable.contains(el)) return;
  if (el.classList.contains("cell-input")) {
    e.preventDefault();
    addExternRow({ focusField: el.dataset.field || "unternehmen" });
    return;
  }
  if (el.classList.contains("filter-select")) {
    e.preventDefault();
    addExternRow({ focusField: "unternehmen" });
    return;
  }
  if (el.type === "radio" && el.closest(".status-segmented")) {
    e.preventDefault();
    addExternRow({ focusField: "unternehmen" });
  }
}

function exportJsonDownload() {
  readEventFieldsFromDom();
  syncActiveListIntoListSets();
  const blob = new Blob(
    [
      JSON.stringify(
        {
          version: STORAGE_VERSION,
          exportedAt: new Date().toISOString(),
          listSets,
          activeListId,
          listSubTab,
          eventSettings,
          appTitle,
        },
        null,
        2
      ),
    ],
    { type: "application/json;charset=utf-8" }
  );
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${filenameStemFromAppTitle()}_${todayStamp()}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  setStatus("JSON heruntergeladen.");
}

/** Rohdaten für CSV/XLSX-Export (Extern / Zusatzlisten — aktueller Filter). */
function getExternExportTable() {
  const headers = [
    "Lfd",
    "Status",
    "Unternehmen",
    "Anrede",
    "Titel",
    "Vorname",
    "Nachname",
    "E-Mail",
    "Abteilung",
    "Straße",
    "Hausnr",
    "PLZ",
    "Stadt",
    "Anmerkungen",
  ];
  const rows = [];
  getFilteredIndices().forEach((i) => {
    const r = rowsExtern[i];
    rows.push([
      lfdForIndex(i),
      rowStatus(r),
      r.unternehmen,
      r.anrede,
      r.titel,
      r.vorname,
      r.nachname,
      r.email,
      r.abteilung,
      r.strasse,
      r.hausnr,
      r.plz,
      r.stadt,
      r.anmerkungen,
    ]);
  });
  return { headers, rows };
}

function getTnExportTable() {
  const headers = ["Lfd", "Zusage_Absage", "Anrede", "Vorname", "Nachname"];
  const rows = [];
  rowsTN.forEach((r, i) => {
    if (!tnRowHasText(r)) return;
    rows.push([lfdTnForIndex(i), rowStatusTn(r) || "", r.anrede, r.vorname, r.nachname]);
  });
  return { headers, rows };
}

/** Dateiname und Kurzbezeichnung für Exporte der aktuellen Unterliste (nicht TN). */
function getExternExportFileInfo(ext) {
  const stamp = todayStamp();
  let fname = `Adressliste_Extern_${stamp}.${ext}`;
  let label = "Extern";
  if (typeof listSubTab === "string" && listSubTab.startsWith("extra:")) {
    const cur = listSets.find((ls) => ls.id === activeListId);
    const eid = listSubTab.slice(6);
    const el = cur?.extraLists?.find((e) => e.id === eid);
    const safe = (el?.name || "Zusatz").replace(/[^\w\-äöüÄÖÜß]+/g, "_").slice(0, 48);
    fname = `Adressliste_${safe}_${stamp}.${ext}`;
    label = el?.name || "Zusatz";
  }
  return { fname, label };
}

function excelSheetName(s) {
  const cleaned = String(s ?? "")
    .replace(/[[\]:*?/\\]/g, "")
    .trim();
  return cleaned.slice(0, 31) || "Liste";
}

const EXTERN_XLSX_COLW = [
  { wch: 5 },
  { wch: 14 },
  { wch: 28 },
  { wch: 8 },
  { wch: 10 },
  { wch: 12 },
  { wch: 14 },
  { wch: 28 },
  { wch: 14 },
  { wch: 22 },
  { wch: 6 },
  { wch: 8 },
  { wch: 14 },
  { wch: 28 },
];
const TN_XLSX_COLW = [{ wch: 5 }, { wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 18 }];

const TN_EXCEL_FILTER_KEYS = ["status", "anrede", "vorname", "nachname"];

function getExcelJsNamespace() {
  const w = typeof globalThis !== "undefined" ? globalThis : window;
  if (!w.ExcelJS) return null;
  if (typeof w.ExcelJS.Workbook === "function") return w.ExcelJS;
  if (w.ExcelJS.default && typeof w.ExcelJS.default.Workbook === "function") return w.ExcelJS.default;
  return null;
}

/** A=1, B=2, … Z, AA, AB … */
function excelColumnLetter(n) {
  let s = "";
  let c = n;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function uniqueValuesForTnColumn(key) {
  const set = new Set();
  let hasEmpty = false;
  for (const row of rowsTN) {
    const v = key === "status" ? rowStatusTn(row) : String(row[key] ?? "");
    const t = v.trim();
    if (!t) hasEmpty = true;
    else set.add(t);
  }
  return { values: Array.from(set).sort((a, b) => a.localeCompare(b, "de")), hasEmpty };
}

/** Hintergrundfarbe Status-Spalte (an UI angelehnt). */
function excelStatusFillArgb(status) {
  const s = String(status || "")
    .trim()
    .toLowerCase();
  if (s === "einladen") return "FFFEF9C3";
  if (s === "offen") return "FFE0F2FE";
  if (s === "Zusage") return "FFD1FAE5";
  if (s === "Absage") return "FFFFE4E6";
  return null;
}

/**
 * Verstecktes Blatt „Listen“: Werte pro Spalte für Dropdown-Validierung.
 * @returns {Record<string, { listenCol: number, fromR: number, toR: number } | null>}
 */
function fillExternListenSheet(listenWs) {
  /** @type {Record<string, { listenCol: number, fromR: number, toR: number } | null>} */
  const ranges = {};
  FILTER_KEYS.forEach((key, idx) => {
    const listenCol = idx + 1;
    let { values, hasEmpty } = uniqueValuesForColumn(key);
    if (key === "status") {
      values = ["einladen", "offen", "Zusage", "Absage"];
      hasEmpty = uniqueValuesForColumn("status").hasEmpty;
    }
    if (!values.length && !hasEmpty) {
      ranges[key] = null;
      return;
    }
    let startRow = 1;
    if (hasEmpty) {
      listenWs.getCell(1, listenCol).value = "";
      startRow = 2;
    }
    values.forEach((v, i) => {
      listenWs.getCell(startRow + i, listenCol).value = v;
    });
    const endRow = startRow + values.length - 1;
    const fromR = hasEmpty ? 1 : startRow;
    ranges[key] = { listenCol, fromR, toR: endRow };
  });
  return ranges;
}

function fillTnListenSheet(listenWs) {
  /** @type {Record<string, { listenCol: number, fromR: number, toR: number } | null>} */
  const ranges = {};
  TN_EXCEL_FILTER_KEYS.forEach((key, idx) => {
    const listenCol = idx + 1;
    let { values, hasEmpty } = uniqueValuesForTnColumn(key);
    if (key === "status") {
      values = ["einladen", "offen", "zusage", "absage"];
      hasEmpty = uniqueValuesForTnColumn("status").hasEmpty;
    }
    if (!values.length && !hasEmpty) {
      ranges[key] = null;
      return;
    }
    let startRow = 1;
    if (hasEmpty) {
      listenWs.getCell(1, listenCol).value = "";
      startRow = 2;
    }
    values.forEach((v, i) => {
      listenWs.getCell(startRow + i, listenCol).value = v;
    });
    const endRow = startRow + values.length - 1;
    const fromR = hasEmpty ? 1 : startRow;
    ranges[key] = { listenCol, fromR, toR: endRow };
  });
  return ranges;
}

function applyExternExcelValidations(dataWs, listenRanges, lastDataRow) {
  if (lastDataRow < 2) return;
  const dv = dataWs.dataValidations;
  if (!dv || typeof dv.add !== "function") return;
  FILTER_KEYS.forEach((key, idx) => {
    const meta = listenRanges[key];
    if (!meta) return;
    const dataCol = idx + 2;
    const L = excelColumnLetter(dataCol);
    const ref = `${L}2:${L}${lastDataRow}`;
    const lL = excelColumnLetter(meta.listenCol);
    const f = `Listen!$${lL}$${meta.fromR}:$${lL}$${meta.toR}`;
    try {
      dv.add(ref, {
        type: "list",
        allowBlank: true,
        formulae: [f],
        showErrorMessage: true,
        errorTitle: "Ungültiger Wert",
        error: "Bitte einen Wert aus der Liste wählen (siehe Blatt „Listen“).",
      });
    } catch (e) {
      console.warn("Excel DV", key, e);
    }
  });
}

function applyTnExcelValidations(dataWs, listenRanges, lastDataRow) {
  if (lastDataRow < 2) return;
  const dv = dataWs.dataValidations;
  if (!dv || typeof dv.add !== "function") return;
  TN_EXCEL_FILTER_KEYS.forEach((key, idx) => {
    const meta = listenRanges[key];
    if (!meta) return;
    const dataCol = idx + 2;
    const L = excelColumnLetter(dataCol);
    const ref = `${L}2:${L}${lastDataRow}`;
    const lL = excelColumnLetter(meta.listenCol);
    const f = `Listen!$${lL}$${meta.fromR}:$${lL}$${meta.toR}`;
    try {
      dv.add(ref, {
        type: "list",
        allowBlank: true,
        formulae: [f],
        showErrorMessage: true,
        errorTitle: "Ungültiger Wert",
        error: "Bitte einen Wert aus der Liste wählen (siehe Blatt „Listen“).",
      });
    } catch (e) {
      console.warn("Excel DV TN", key, e);
    }
  });
}

async function exportListeXlsxExcelJS() {
  const Excel = getExcelJsNamespace();
  if (!Excel) throw new Error("ExcelJS nicht geladen");

  if (listSubTab === "tn") {
    const { headers, rows } = getTnExportTable();
    const wb = new Excel.Workbook();
    const dataWs = wb.addWorksheet(excelSheetName("TN_Zusage_Absage"), {
      views: [{ state: "frozen", ySplit: 1 }],
    });
    const listenWs = wb.addWorksheet("Listen", { state: "hidden" });
    const listenRanges = fillTnListenSheet(listenWs);

    dataWs.addRow(headers);
    rows.forEach((line) => dataWs.addRow(line.map((c) => (c == null ? "" : c))));
    const lastRow = rows.length + 1;
    const lastColL = excelColumnLetter(headers.length);
    dataWs.autoFilter = `A1:${lastColL}1`;
    dataWs.getRow(1).font = { bold: true };
    dataWs.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF1F5F9" },
    };
    TN_XLSX_COLW.forEach((w, i) => {
      dataWs.getColumn(i + 1).width = w.wch;
    });
    for (let r = 2; r <= lastRow; r++) {
      const st = dataWs.getCell(r, 2).value;
      const argb = excelStatusFillArgb(st);
      if (argb) {
        dataWs.getCell(r, 2).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb },
        };
      }
    }
    applyTnExcelValidations(dataWs, listenRanges, lastRow);

    const buf = await wb.xlsx.writeBuffer();
    downloadBlobBinary(
      new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `TN_Zusage_Absage_${todayStamp()}.xlsx`,
    );
    setStatus("Excel-Liste heruntergeladen.");
    return;
  }

  const { headers, rows } = getExternExportTable();
  const { fname, label } = getExternExportFileInfo("xlsx");
  const sheetLabel =
    typeof listSubTab === "string" && listSubTab.startsWith("extra:") ? label : "Adressliste";

  const wb = new Excel.Workbook();
  const dataWs = wb.addWorksheet(excelSheetName(sheetLabel), {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  const listenWs = wb.addWorksheet("Listen", { state: "hidden" });
  const listenRanges = fillExternListenSheet(listenWs);

  dataWs.addRow(headers);
  rows.forEach((line) => dataWs.addRow(line.map((c) => (c == null ? "" : c))));
  const lastRow = rows.length + 1;
  dataWs.autoFilter = `A1:${excelColumnLetter(headers.length)}1`;
  dataWs.getRow(1).font = { bold: true };
  dataWs.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF1F5F9" },
  };
  EXTERN_XLSX_COLW.forEach((w, i) => {
    dataWs.getColumn(i + 1).width = w.wch;
  });
  for (let r = 2; r <= lastRow; r++) {
    const st = dataWs.getCell(r, 2).value;
    const argb = excelStatusFillArgb(st);
    if (argb) {
      dataWs.getCell(r, 2).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb },
      };
    }
  }
  applyExternExcelValidations(dataWs, listenRanges, lastRow);

  const buf = await wb.xlsx.writeBuffer();
  downloadBlobBinary(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    fname,
  );
  setStatus(`Excel (${label}) heruntergeladen.`);
}

function exportListeCsv() {
  const sep = ";";
  if (listSubTab === "tn") {
    const { headers, rows } = getTnExportTable();
    const lines = [headers.join(sep)];
    rows.forEach((cells) => {
      lines.push(cells.map(csvCell).join(sep));
    });
    downloadBlob("\ufeff" + lines.join("\r\n"), `TN_Zusage_Absage_${todayStamp()}.csv`, "text/csv;charset=utf-8");
    setStatus("CSV (TN) heruntergeladen.");
    return;
  }
  const { headers, rows } = getExternExportTable();
  const lines = [headers.join(sep)];
  rows.forEach((cells) => {
    lines.push(cells.map(csvCell).join(sep));
  });
  const { fname, label } = getExternExportFileInfo("csv");
  downloadBlob("\ufeff" + lines.join("\r\n"), fname, "text/csv;charset=utf-8");
  setStatus(`CSV (${label}) heruntergeladen.`);
}

async function exportListeXlsx() {
  if (typeof XLSX === "undefined") {
    setStatus("Excel-Bibliothek nicht geladen — bitte Seite neu laden.");
    return;
  }
  let excelJsFallback = false;
  const ExcelNs = getExcelJsNamespace();
  if (ExcelNs) {
    try {
      await exportListeXlsxExcelJS();
      return;
    } catch (e) {
      console.warn("ExcelJS-Export fehlgeschlagen, Fallback auf einfaches XLSX:", e);
      excelJsFallback = true;
    }
  }
  if (listSubTab === "tn") {
    const { headers, rows } = getTnExportTable();
    const aoa = [headers, ...rows.map((line) => line.map((c) => (c == null ? "" : c)))];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = TN_XLSX_COLW;
    if (ws["!ref"]) ws["!autofilter"] = { ref: ws["!ref"] };
    XLSX.utils.book_append_sheet(wb, ws, excelSheetName("TN_Zusage_Absage"));
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlobBinary(
      new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `TN_Zusage_Absage_${todayStamp()}.xlsx`,
    );
    setStatus(
      excelJsFallback
        ? "Excel (TN) heruntergeladen (ohne Dropdowns/Farben)."
        : "Excel (TN) heruntergeladen.",
    );
    return;
  }
  const { headers, rows } = getExternExportTable();
  const aoa = [headers, ...rows.map((line) => line.map((c) => (c == null ? "" : c)))];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = EXTERN_XLSX_COLW;
  if (ws["!ref"]) ws["!autofilter"] = { ref: ws["!ref"] };
  const { fname, label } = getExternExportFileInfo("xlsx");
  const sheetLabel =
    typeof listSubTab === "string" && listSubTab.startsWith("extra:") ? label : "Adressliste";
  XLSX.utils.book_append_sheet(wb, ws, excelSheetName(sheetLabel));
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  downloadBlobBinary(
    new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    fname,
  );
  setStatus(
    excelJsFallback
      ? `Excel (${label}) heruntergeladen (ohne Dropdowns/Farben).`
      : `Excel (${label}) heruntergeladen.`,
  );
}

function exportMergeCsv() {
  readEventFieldsFromDom();
  saveEventOnly();

  const mergeHeaders = [
    "Lfd",
    "Vorname",
    "Nachname",
    "Name_komplett",
    "Anrede",
    "Titel",
    "Briefanrede",
    "Briefanrede_mit_Vorname",
    "Name_mit_Anrede_Titel",
    "Name_Titel_Nachname",
    "Anrede_Geschlecht",
    "Geschlecht",
    "Datum_YYYYMMDD",
    "Anrede_Kurz",
    "Straße",
    "Hausnummer",
    "PLZ",
    "Ort",
    "Adresse_Zeile1",
    "Adresse_Zeile2",
    "Briefkopf_Block",
    "Briefkopf_zeilen",
    "Empfaenger_Adresse_komplett",
    "Empfaenger_Zeile1",
    "Empfaenger_Zeile2",
    "Empfaenger_Zeile3",
    "Empfaenger_Zeile4",
    "Empfaenger_Zeile5",
    "Empfaenger_Zeile6",
    "E_Mail",
    "Unternehmen",
    "Abteilung",
    "Briefdatum",
    "Ausstellungsdatum",
    "Schreibdatum",
    "Veranstaltung_Datum",
    "Veranstaltung_Zeit",
    "Veranstaltung_Ort",
    "Veranstaltung_Terminzeile",
    "Export_Dateiname",
    "Export_Namenszeile_Datei",
  ];

  const pickedRaw = rowsExtern.filter((r) => selectedIds.has(r.id));
  if (!pickedRaw.length) {
    alert("Bitte mindestens einen Empfänger mit der Checkbox auswählen.");
    return;
  }

  const picked = pickedRaw.filter((r) => !shouldOmitEinladungOhneAdresseUndMail(r));
  const nOmitMerge = pickedRaw.length - picked.length;
  if (!picked.length) {
    alert(
      "Kein exportierbarer Eintrag: Ausgewählte Einladungen ohne Adresse und ohne E-Mail werden nicht in die Merge-CSV übernommen. Bitte PLZ, Ort und Straße bzw. eine E-Mail ergänzen.",
    );
    return;
  }

  const sep = ";";
  const lines = [mergeHeaders.join(sep)];
  picked.forEach((r) => {
    const m = mergeRowObject(r);
    const row = mergeHeaders.map((h) => csvCell(m[h]));
    lines.push(row.join(sep));
  });

  downloadBlob("\ufeff" + lines.join("\r\n"), `Seriendruck_Auswahl_${todayStamp()}.csv`, "text/csv;charset=utf-8");
  let mergeMsg = "Merge-CSV heruntergeladen (" + picked.length + " Empfänger).";
  if (nOmitMerge) mergeMsg += " " + nOmitMerge + " Einladung(en) ohne Adresse/E-Mail ausgelassen.";
  setStatusSerien(mergeMsg);
}

function downloadBlob(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function downloadBlobBinary(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

const SERIEN_IDB = "bwk-adressliste-serien-v1";
const SERIEN_STORE = "kv";

function openSerienIdb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(SERIEN_IDB, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(SERIEN_STORE)) req.result.createObjectStore(SERIEN_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function serienIdbGet(key) {
  const db = await openSerienIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SERIEN_STORE, "readonly");
    const q = tx.objectStore(SERIEN_STORE).get(key);
    q.onsuccess = () => resolve(q.result);
    q.onerror = () => reject(q.error);
  });
}

async function serienIdbSet(key, val) {
  const db = await openSerienIdb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SERIEN_STORE, "readwrite");
    tx.objectStore(SERIEN_STORE).put(val, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function sanitizeFilenameSegment(s) {
  return String(s || "")
    .replace(/[\\/:*?"<>|]+/g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

/**
 * Standard-Stamm ohne Endung (DOCX / PDF / E-Mail-Dateiname gleich):
 * Präfix_JJJJMMTT_Anr._Titel_ggf_Nachname
 * Präfix leer → Einladung_AVöD_Auftaktveranstaltung; Datum aus Veranstaltungsfeld (ev. Datum), sonst heute.
 */
function buildSerienExportBasename(r) {
  const raw = (eventSettings.serienExportPrefix || "").trim();
  const prefix =
    sanitizeFilenameSegment(raw) ||
    sanitizeFilenameSegment("Einladung_AVöD_Auftaktveranstaltung") ||
    "Einladung";
  const ymd =
    parseDatumToYyyymmdd(eventSettings.datum) || todayStamp().replace(/-/g, "");
  const parts = [prefix, ymd, sanitizeFilenameSegment(anredeKurzDatei(r))];
  const ti = (r.titel || "").trim();
  if (ti) parts.push(sanitizeFilenameSegment(ti));
  parts.push(sanitizeFilenameSegment(r.nachname || "") || "Kontakt");
  return parts.filter(Boolean).join("_");
}

/** Nur der Namens-/Anrede-Teil wie im Dateinamen (Fr._Dr._Mustermann …). */
function buildSerienExportNameSuffix(r) {
  const parts = [sanitizeFilenameSegment(anredeKurzDatei(r))];
  const ti = (r.titel || "").trim();
  if (ti) parts.push(sanitizeFilenameSegment(ti));
  parts.push(sanitizeFilenameSegment(r.nachname || "") || "Kontakt");
  return parts.filter(Boolean).join("_");
}

function getDocxtemplaterConstructor() {
  const w = typeof globalThis !== "undefined" ? globalThis : window;
  if (typeof w.Docxtemplater === "function") return w.Docxtemplater;
  if (w.docxtemplater && typeof w.docxtemplater.Docxtemplater === "function")
    return w.docxtemplater.Docxtemplater;
  if (w.docxtemplater && typeof w.docxtemplater.default === "function") return w.docxtemplater.default;
  return null;
}

/** Word OOXML: .dotx (Vorlage) und .docx sind ZIP-Pakete („PK“). Altes .doc/.dot (OLE): D0 CF 11 E0 … */
function isWordOoxmlZipBuffer(buffer) {
  if (!buffer || buffer.byteLength < 4) return false;
  const u = new Uint8Array(buffer.slice(0, 4));
  return u[0] === 0x50 && u[1] === 0x4b;
}

function isOleWordDocBuffer(buffer) {
  if (!buffer || buffer.byteLength < 4) return false;
  const u = new Uint8Array(buffer.slice(0, 4));
  return u[0] === 0xd0 && u[1] === 0xcf && u[2] === 0x11 && u[3] === 0xe0;
}

/** Entfernt Steuerzeichen, die in WordprocessingML zu ungültigem XML / nicht lesbaren .docx führen können */
function sanitizeDocxMergeValue(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  return s;
}

function sanitizeMergeDataForDocx(data) {
  const out = {};
  for (const k of Object.keys(data)) {
    const v = data[k];
    if (v === null || v === undefined) {
      out[k] = "";
    } else if (Array.isArray(v)) {
      out[k] = v.map((item) => sanitizeDocxMergeValue(String(item)));
    } else if (typeof v === "number" && Number.isFinite(v)) {
      out[k] = sanitizeDocxMergeValue(String(v));
    } else if (typeof v === "number") {
      out[k] = "";
    } else {
      out[k] = sanitizeDocxMergeValue(v);
    }
  }
  return out;
}

/**
 * .dotx-Vorlagen deklarieren document.xml als „template.main“; exportierte .docx müssen „document.main“ sein,
 * sonst meldet Word oft „beschädigt“, während Pages/Preview toleranter sind.
 */
function normalizeDocxContentTypesForExportedDocument(ctXml) {
  return String(ctXml || "").replace(
    /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.template\.main\+xml/g,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  );
}

/** Verweis auf die .dotx-Quelle in app.xml kann bei .docx-Export irritieren — entfernen. */
function stripDocPropsAppTemplateXml(appXml) {
  return String(appXml || "").replace(/<Template>[\s\S]*?<\/Template>/g, "");
}

/** Nach docxtemplater: gleiches PizZip-Paket, aber OOXML-Typen wie bei echtem .docx (nicht .dotx). */
function applyWordExportFixesToPizZip(zip) {
  if (!zip || typeof zip.file !== "function") return;
  try {
    const ct = zip.file("[Content_Types].xml");
    if (ct && typeof ct.asText === "function") {
      zip.file("[Content_Types].xml", normalizeDocxContentTypesForExportedDocument(ct.asText()));
    }
    const app = zip.file("docProps/app.xml");
    if (app && typeof app.asText === "function") {
      zip.file("docProps/app.xml", stripDocPropsAppTemplateXml(app.asText()));
    }
  } catch (e) {
    console.warn("applyWordExportFixesToPizZip:", e);
  }
}

/**
 * Paket mit JSZip neu schreiben (DEFLATE, Zentraldirectory) — manche Word-Versionen sind beim Roh-PizZip-Output pingelig.
 */
async function repackDocxUint8ArrayWithJsZip(u8) {
  const JSZipCtor = typeof window !== "undefined" ? window.JSZip : null;
  if (!JSZipCtor || !(u8 instanceof Uint8Array)) return u8;
  try {
    const zip = await JSZipCtor.loadAsync(u8);
    return await zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  } catch (e) {
    console.warn("DOCX-Repack (ZIP):", e);
    return u8;
  }
}

async function renderDocxFromTemplate(arrayBuffer, data) {
  const PizZip = typeof window !== "undefined" ? window.PizZip : null;
  const DocT = getDocxtemplaterConstructor();
  if (!PizZip || !DocT) {
    throw new Error("Word-Bibliothek (PizZip/Docxtemplater) nicht geladen. Bitte Seite neu laden.");
  }
  const zip = new PizZip(arrayBuffer);
  const doc = new DocT(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter() {
      return "";
    },
  });
  doc.setData(sanitizeMergeDataForDocx(data));
  doc.render();
  const zipOut = doc.getZip();
  applyWordExportFixesToPizZip(zipOut);
  const raw = zipOut.generate({ type: "uint8array" });
  const out = await repackDocxUint8ArrayWithJsZip(raw);
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

/** Einfache Platzhalter wie {Briefanrede} — Keys wie in mergeRowObject */
function applySerienTextTemplate(templateStr, data) {
  return String(templateStr || "").replace(/\{([A-Za-z_0-9]+)\}/g, (_, key) => {
    const v = data[key];
    if (v === undefined || v === null) return "";
    if (Array.isArray(v)) return sanitizeDocxMergeValue(v.join(" · "));
    return sanitizeDocxMergeValue(String(v));
  });
}

function getMammothApi() {
  const w = typeof window !== "undefined" ? window : globalThis;
  if (w.mammoth && typeof w.mammoth.extractRawText === "function") return w.mammoth;
  return null;
}

/** Gespeicherte E-Mail-Vorlage: Word-OOXML (.docx/.dotx) oder reiner Text (ohne kind: aus Inhalt erraten) */
function getMailTemplateKind(tplMail) {
  if (!tplMail || !tplMail.buffer) return "none";
  if (tplMail.kind === "text" || tplMail.kind === "docx") return tplMail.kind;
  const buf = tplMail.buffer;
  if (isOleWordDocBuffer(buf)) return "ole";
  if (isWordOoxmlZipBuffer(buf)) return "docx";
  return "text";
}

/** E-Mail-Body: bei Word-Vorlage docxtemplater + Mammoth (Rohtext); sonst Platzhalter im Text ersetzen */
async function buildMailBodyFromTemplate(tplMail, r) {
  readEventFieldsFromDom();
  saveEventOnly();
  const data = mergeRowObject(r);
  const kind = getMailTemplateKind(tplMail);
  if (kind === "ole") {
    throw new Error("E-Mail-Vorlage ist altes Word (.doc/.dot) — bitte als .docx speichern.");
  }
  if (kind === "docx") {
    const mammoth = getMammothApi();
    if (!mammoth) {
      throw new Error("Mammoth-Bibliothek nicht geladen (Seite neu laden).");
    }
    const filled = await renderDocxFromTemplate(tplMail.buffer.slice(0), data);
    const ab = await filled.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: ab });
    return String(result.value || "").replace(/\r\n/g, "\n").trim();
  }
  const mailTplText = new TextDecoder("utf-8").decode(tplMail.buffer);
  return applySerienTextTemplate(mailTplText, data);
}

/** Datum-Zeile für EML-Köpfe (RFC 2822-ähnlich, lokale Zeitzone). */
function buildEmlDateHeader(d = new Date()) {
  const w = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const m = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const tz = -d.getTimezoneOffset();
  const sign = tz >= 0 ? "+" : "-";
  const off = Math.abs(tz);
  const oh = Math.floor(off / 60);
  const om = off % 60;
  const tzs = `${sign}${String(oh).padStart(2, "0")}${String(om).padStart(2, "0")}`;
  return `${w[d.getDay()]}, ${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()} ${String(d.getHours()).padStart(
    2,
    "0",
  )}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")} ${tzs}`;
}

/**
 * From-Zeile für .eml: reine E-Mail oder „Name <mail@…>“.
 */
function pickEmlFromHeader(raw) {
  const s = String(raw || "")
    .replace(/\r?\n/g, " ")
    .trim();
  if (!s) return "export-placeholder@invalid";
  const br = s.match(/<([^>]+)>/);
  const inner = (br ? br[1] : s).trim();
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/i.test(inner)) return s;
  return "export-placeholder@invalid";
}

/**
 * .eml für Serienexport: u. a. From/Date/Message-ID und X-Unsent (Entwurf-Kennzeichnung).
 * Hinweis: Apple Mail zeigt .eml oft trotzdem wie Posteingang — Outlook wertet X-Unsent zuverlässiger.
 */
function buildEmlBlob(toAddr, subject, bodyText, fromAddrOpt) {
  const nl = "\r\n";
  const escHeader = (h) =>
    String(h || "")
      .replace(/\r?\n/g, " ")
      .trim();
  const from = pickEmlFromHeader(fromAddrOpt);
  const msgId =
    "<" + Date.now() + "." + Math.random().toString(36).slice(2, 12) + "@adressliste-export>";
  const head =
    `From: ${from}${nl}` +
    `To: ${escHeader(toAddr)}${nl}` +
    `Subject: ${escHeader(subject)}${nl}` +
    `Date: ${escHeader(buildEmlDateHeader())}${nl}` +
    `Message-ID: ${msgId}${nl}` +
    `X-Unsent: 1${nl}` +
    `MIME-Version: 1.0${nl}` +
    `Content-Type: text/plain; charset=UTF-8${nl}` +
    `Content-Transfer-Encoding: 8bit${nl}${nl}`;
  return new Blob([head + bodyText.replace(/\r?\n/g, nl)], { type: "message/rfc822" });
}

async function refreshSerienTemplateLabels() {
  const dx = await serienIdbGet("tpl_docx");
  const ml = await serienIdbGet("tpl_mail");
  const elDx = document.getElementById("label-template-docx");
  const elMl = document.getElementById("label-template-mail");
  if (elDx) elDx.textContent = dx && dx.name ? dx.name : "Keine Vorlage gespeichert";
  if (elMl) elMl.textContent = ml && ml.name ? ml.name : "Keine Vorlage gespeichert";
}

async function refreshExportDirLabel() {
  const h = await serienIdbGet("export_dir_handle");
  const el = document.getElementById("label-export-dir");
  if (!el) return;
  if (!h || !h.name) {
    el.textContent = "— (beim Export wählen oder nicht gespeichert)";
    return;
  }
  el.textContent = `Ordner: ${h.name}`;
}

async function pickAndStoreExportDirectory() {
  if (typeof window.showDirectoryPicker !== "function") {
    alert(
      "Ordnerwahl wird von diesem Browser nicht unterstützt. Beim Export wird eine ZIP-Datei angeboten.",
    );
    return;
  }
  try {
    const dir = await window.showDirectoryPicker({ mode: "readwrite" });
    try {
      await serienIdbSet("export_dir_handle", dir);
      refreshExportDirLabel();
      setStatusSerien("Exportordner gespeichert.");
    } catch (e) {
      refreshExportDirLabel();
      setStatusSerien("Ordner aktiv, konnte aber nicht dauerhaft gemerkt werden — beim nächsten Export erneut wählen.");
      console.warn(e);
    }
  } catch (e) {
    if (e && e.name === "AbortError") return;
    throw e;
  }
}

async function getWritableExportDirectory() {
  if (typeof window.showDirectoryPicker !== "function") return null;
  let handle = await serienIdbGet("export_dir_handle");
  if (handle) {
    try {
      const st = await handle.requestPermission({ mode: "readwrite" });
      if (st === "granted") return handle;
    } catch (_) {
      handle = null;
    }
  }
  try {
    handle = await window.showDirectoryPicker({ mode: "readwrite" });
    try {
      await serienIdbSet("export_dir_handle", handle);
    } catch (e) {
      console.warn(e);
    }
    refreshExportDirLabel();
    return handle;
  } catch (e) {
    if (e && e.name === "AbortError") return null;
    throw e;
  }
}

async function writeBlobToDirectory(dirHandle, filename, blob) {
  const fh = await dirHandle.getFileHandle(filename, { create: true });
  const writable = await fh.createWritable();
  await writable.write(blob);
  await writable.close();
}

async function exportSerienBriefFiles() {
  readEventFieldsFromDom();
  saveEventOnly();

  const mode = eventSettings.serienExportWhat || "both";
  const wantWord = mode === "both" || mode === "word";
  const wantEmail = mode === "both" || mode === "email";

  const tplDocx = await serienIdbGet("tpl_docx");
  const tplMail = await serienIdbGet("tpl_mail");

  if (wantWord) {
    if (!tplDocx || !tplDocx.buffer) {
      alert(
        "Für den Word-Export: Bitte zuerst eine Word-Vorlage wählen und speichern (empfohlen: .dotx; alternativ .docx — aus klassischem .doc/.dot in Word entsprechend speichern).",
      );
      return;
    }
    if (isOleWordDocBuffer(tplDocx.buffer) || !isWordOoxmlZipBuffer(tplDocx.buffer)) {
      alert(
        "Die gespeicherte Word-Vorlage ist kein verwendbares OOXML (.dotx/.docx). Bitte eine Vorlage im Format „Word-Vorlage (*.dotx)“ oder „Word-Dokument (*.docx)“ wählen bzw. neu laden.",
      );
      return;
    }
  }
  if (wantEmail) {
    if (!tplMail || !tplMail.buffer) {
      alert("Für den E-Mail-Export: Bitte zuerst eine E-Mail-Vorlage wählen (Word .docx/.dotx oder Text/HTML mit Platzhaltern).");
      return;
    }
    const mailKindPre = getMailTemplateKind(tplMail);
    if (mailKindPre === "ole") {
      alert(
        "Die gespeicherte E-Mail-Vorlage ist ein altes Word-Format — bitte als .docx/.dotx speichern und neu laden.",
      );
      return;
    }
    if (mailKindPre === "docx" && !getMammothApi()) {
      alert("Bibliothek für E-Mail-Text aus Word (Mammoth) nicht geladen. Bitte Seite neu laden.");
      return;
    }
  }

  const pickedRaw = rowsExtern.filter((r) => selectedIds.has(r.id));
  if (!pickedRaw.length) {
    alert("Bitte mindestens einen Empfänger mit der Checkbox auswählen.");
    return;
  }

  const picked = pickedRaw.filter((r) => !shouldOmitEinladungOhneAdresseUndMail(r));
  const nOmitVorlage = pickedRaw.length - picked.length;
  if (!picked.length) {
    alert(
      "Kein exportierbarer Eintrag: Ausgewählte Einladungen ohne Adresse und ohne E-Mail werden nicht in die Vorlagen übernommen. Bitte PLZ, Ort und Straße bzw. eine E-Mail ergänzen.",
    );
    return;
  }

  if (wantEmail && !wantWord && !picked.some((r) => hasEmailForBrief(r))) {
    alert("Nur E-Mail-Export: Keiner der verbleibenden Kontakte hat eine E-Mail-Adresse.");
    return;
  }

  if (wantWord && !wantEmail && !picked.some((r) => rowQualifiesForWordBrief(r))) {
    alert(
      "Nur Word-Export: Bei Status „Einladung“ ist eine Briefadresse nötig — keiner der ausgewählten Einträge hat PLZ, Ort und Straße/Hausnr.",
    );
    return;
  }

  const subjectTpl = eventSettings.serienMailSubject || "Einladung — {Veranstaltung_Ort}";

  let dirHandle = null;
  if (typeof window.showDirectoryPicker === "function") {
    dirHandle = await getWritableExportDirectory();
  }
  const zip =
    !dirHandle && typeof window.JSZip === "function" ? new window.JSZip() : null;
  if (!dirHandle && !zip) {
    alert(
      "Kein Speicherort: Ordnerwahl abgebrochen/nicht unterstützt, und JSZip nicht geladen. Bitte Seite neu laden oder einen unterstützten Browser nutzen.",
    );
    return;
  }

  const usedNames = new Map();
  let nDoc = 0;
  let nMail = 0;
  let skippedMail = 0;
  let skippedWordEinladungNoAddr = 0;
  const warnings = [];
  const total = picked.length;

  openExportProgressDialog(total);
  try {
    let done = 0;
    for (const r of picked) {
      const data = mergeRowObject(r);
      const base = buildSerienExportBasename(r);
      const num = (usedNames.get(base) || 0) + 1;
      usedNames.set(base, num);
      const fnameBase = num > 1 ? `${base}_${num}` : base;

      const email = String(r.email || "").trim();
      const hasMail = hasEmailForBrief(r);

      let docxBlob = null;
      if (wantWord && rowQualifiesForWordBrief(r)) {
        const docxBuf = tplDocx.buffer.slice(0);
        try {
          docxBlob = await renderDocxFromTemplate(docxBuf, data);
        } catch (err) {
          throw new Error(
            `Word für „${fnameBase}“: ${err && err.message ? err.message : String(err)} — Vorlage (.dotx/.docx) und Platzhalter prüfen.`,
          );
        }
      } else if (wantWord && rowStatus(r) === "einladen" && !hasPostalAddressForBrief(r)) {
        skippedWordEinladungNoAddr += 1;
      }

      let bodyText = "";
      let subject = "";
      if (wantEmail && hasMail) {
        try {
          bodyText = await buildMailBodyFromTemplate(tplMail, r);
        } catch (err) {
          throw new Error(
            `E-Mail-Text für „${fnameBase}“: ${err && err.message ? err.message : String(err)}`,
          );
        }
        subject = applySerienTextTemplate(subjectTpl, data);
      }

      if (dirHandle) {
        if (wantWord && docxBlob) {
          await writeBlobToDirectory(dirHandle, fnameBase + ".docx", docxBlob);
          nDoc += 1;
        }
        if (wantEmail) {
          if (hasMail) {
            await writeBlobToDirectory(
              dirHandle,
              fnameBase + ".eml",
              buildEmlBlob(email, subject, bodyText, eventSettings.serienMailFrom),
            );
            nMail += 1;
          } else {
            skippedMail += 1;
            warnings.push(fnameBase);
          }
        }
      } else if (zip) {
        if (wantWord && docxBlob) {
          zip.file(fnameBase + ".docx", docxBlob);
          nDoc += 1;
        }
        if (wantEmail) {
          if (hasMail) {
            zip.file(fnameBase + ".eml", buildEmlBlob(email, subject, bodyText, eventSettings.serienMailFrom));
            nMail += 1;
          } else {
            skippedMail += 1;
            warnings.push(fnameBase);
          }
        }
      }

      done += 1;
      updateExportProgressDialog(done, total);
      await yieldToUi();
    }

    if (dirHandle) {
      const bits = [];
      if (wantWord) bits.push(`${nDoc} Word-Datei(en)`);
      if (wantEmail) bits.push(`${nMail} E-Mail(s)`);
      let msg = `${bits.join(" · ")} im Ordner gespeichert.`;
      if (nOmitVorlage) msg += ` ${nOmitVorlage} Einladung(en) ohne Adresse/E-Mail von vornherein ausgelassen.`;
      if (wantWord && skippedWordEinladungNoAddr) {
        msg += ` ${skippedWordEinladungNoAddr} Einladung(en) nur per E-Mail (kein Brief ohne vollständige Adresse).`;
      }
      if (wantEmail && skippedMail) msg += ` ${skippedMail} ohne E-Mail-Adresse (.eml ausgelassen).`;
      setStatusSerien(msg);
    } else {
      updateExportProgressDialog(total, total, "ZIP-Archiv wird erstellt…");
      await yieldToUi();
      const blob = await zip.generateAsync({ type: "blob" }, (meta) => {
        if (meta && typeof meta.percent === "number") {
          updateExportProgressDialog(total, total, `ZIP: ${Math.round(meta.percent)} %`);
        }
      });
      downloadBlobBinary(blob, `Serienbriefe_${todayStamp()}.zip`);
      const bits = [];
      if (wantWord) bits.push(`${nDoc} Word-Datei(en)`);
      if (wantEmail) bits.push(`${nMail} E-Mail(s)`);
      let msg = `ZIP mit ${bits.join(" · ")} heruntergeladen.`;
      if (nOmitVorlage) msg += ` ${nOmitVorlage} Einladung(en) ohne Adresse/E-Mail von vornherein ausgelassen.`;
      if (wantWord && skippedWordEinladungNoAddr) {
        msg += ` ${skippedWordEinladungNoAddr} Einladung(en) nur per E-Mail (kein Brief ohne vollständige Adresse).`;
      }
      if (wantEmail && skippedMail) {
        msg += ` ${skippedMail} ohne E-Mail (.eml fehlen): ${warnings.slice(0, 5).join(", ")}${warnings.length > 5 ? "…" : ""}.`;
      }
      setStatusSerien(msg);
    }
  } finally {
    closeExportProgressDialog();
  }
  await maybeShowPdfWordHinweisNachExport(nDoc, wantWord);
}

function todayStamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/** TT.MM.JJJJ für Briefkopf „Berlin, …“ (wenn kein eigenes Schreibdatum gesetzt) */
function briefdatumGermanToday() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
}

function clearAll() {
  let label = "die Liste „TN Zusage / Absage“";
  if (listSubTab === "extern") label = "die Liste „Extern“";
  else if (typeof listSubTab === "string" && listSubTab.startsWith("extra:")) {
    const cur = listSets.find((ls) => ls.id === activeListId);
    const eid = listSubTab.slice(6);
    const el = cur?.extraLists?.find((e) => e.id === eid);
    label = el ? `die Liste „${el.name}“` : "die aktuelle Liste";
  }
  if (!confirm(`Alle Einträge in ${label} löschen?`)) return;
  pushUndoBeforeMutation();
  if (listSubTab === "tn") {
    rowsTN = [emptyTnRow()];
  } else {
    rowsExtern = [emptyRow()];
    selectedIds.clear();
  }
  saveAll();
  render();
  setStatus("Liste geleert.");
}

function backupLooksFilled(raw) {
  try {
    const parsed = JSON.parse(raw);
    if (parsed && parsed.version >= 5 && Array.isArray(parsed.listSets)) {
      return parsed.listSets.some((ls) => {
        const ex = (ls.rowsExtern || []).map(migrateRow);
        const tn = (ls.rowsTN || []).map(migrateTnRow);
        const extraHas =
          Array.isArray(ls.extraLists) &&
          ls.extraLists.some((e) => (e.rows || []).some((r) => rowStatus(r) || rowHasText(r)));
        return (
          ex.some((r) => rowStatus(r) || rowHasText(r)) ||
          tn.some((r) => tnRowHasText(r)) ||
          extraHas
        );
      });
    }
    if (parsed && Array.isArray(parsed.rowsExtern)) {
      const ex = parsed.rowsExtern.map(migrateRow);
      const tn = (parsed.rowsTN || []).map(migrateTnRow);
      return (
        ex.some((r) => rowStatus(r) || rowHasText(r)) || tn.some((r) => tnRowHasText(r))
      );
    }
    const arr = parsed.rows || parsed;
    if (!Array.isArray(arr)) return false;
    return arr.map(migrateRow).some((r) => rowStatus(r) || rowHasText(r));
  } catch (_) {
    return false;
  }
}

function tryRestoreFailsafe() {
  try {
    const fb = localStorage.getItem(FAILSAFE_KEY);
    if (!fb) return false;
    const parsed = JSON.parse(fb);
    const canRestore =
      (parsed && parsed.version >= 5 && Array.isArray(parsed.listSets)) ||
      (parsed && Array.isArray(parsed.rowsExtern)) ||
      (parsed && Array.isArray(parsed.rows)) ||
      Array.isArray(parsed);
    if (!canRestore) return false;
    pushUndoBeforeMutation();
    if (parsed && parsed.version >= 5 && Array.isArray(parsed.listSets) && parsed.listSets.length) {
      listSets = parsed.listSets.map((ls) => listSetFromStorage(ls));
      activeListId =
        parsed.activeListId && listSets.some((l) => l.id === parsed.activeListId)
          ? parsed.activeListId
          : listSets[0].id;
      if (typeof parsed.appTitle === "string") appTitle = normalizeAppTitle(parsed.appTitle);
      if (parsed.eventSettings) eventSettings = { ...eventSettings, ...parsed.eventSettings };
      listSubTab = normalizeListSubTabFromStorage(parsed.listSubTab);
      ensureValidListSubTab();
      loadRowsForCurrentSubTab();
    } else if (parsed && Array.isArray(parsed.rowsExtern)) {
      rowsExtern = parsed.rowsExtern.map(migrateRow);
      rowsTN = (parsed.rowsTN && parsed.rowsTN.length ? parsed.rowsTN : [emptyTnRow()]).map(migrateTnRow);
      if (parsed.eventSettings) eventSettings = { ...eventSettings, ...parsed.eventSettings };
      const id = crypto.randomUUID();
      listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
      activeListId = id;
    } else if (parsed.rows) {
      rowsExtern = parsed.rows.map(migrateRow);
      rowsTN = [emptyTnRow()];
      if (parsed.eventSettings) eventSettings = { ...eventSettings, ...parsed.eventSettings };
      const id = crypto.randomUUID();
      listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
      activeListId = id;
    } else {
      rowsExtern = parsed.map(migrateRow);
      rowsTN = [emptyTnRow()];
      const id = crypto.randomUUID();
      listSets = [{ id, name: "Hauptliste", useExtern: true, useTn: true, extraLists: [], rowsExtern, rowsTN }];
      activeListId = id;
    }
    trimEmptyExternRows();
    trimEmptyTnRows();
    applyAppTitleToDom();
    saveAll();
    loadEventFieldsToDom();
    switchListSubTab(listSubTab);
    setStatus("Backup wiederhergestellt.");
    return true;
  } catch (_) {
    return false;
  }
}

/** Max. Zeilen pro Excel/CSV-Import (Schutz vor Hängern) */
const SHEET_IMPORT_MAX_ROWS = 8000;

function normalizeHeaderCell(s) {
  return String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/** Einzelnes Wort als akademischen Titel werten (Herr/Frau sind keine Titel hier) */
function isTitleToken(tok) {
  const t = tok.trim();
  if (!t) return false;
  return (
    /^dr\.?$/i.test(t) ||
    /^prof\.?$/i.test(t) ||
    /^dipl\.-?ing\.?$/i.test(t) ||
    /^univ\.-?prof\.?$/i.test(t) ||
    /^ing\.?$/i.test(t) ||
    /^(med|rer|nat|jur|phil)\.?$/i.test(t) ||
    /^habil\.?$/i.test(t) ||
    /^mba$/i.test(t) ||
    /^m(sc|b)\.?c?\.?$/i.test(t) ||
    /^pd\.?$/i.test(t) ||
    /^dipl\.?$/i.test(t)
  );
}

/**
 * Eine Zelle wie "Herr Dr. Max Mustermann" oder "Mustermann, Dr. Max" in Anrede/Titel/Vor-/Nachname zerlegen.
 */
function parseNameCell(raw) {
  let s = String(raw ?? "").replace(/\s+/g, " ").trim();
  if (!s) return { anrede: "", titel: "", vorname: "", nachname: "" };

  if (/^[^,]+,\s*.+$/.test(s)) {
    const idx = s.indexOf(",");
    const left = s.slice(0, idx).trim();
    const right = s.slice(idx + 1).trim();
    if (left && right) {
      const p = parseNameCell(right);
      return {
        anrede: p.anrede,
        titel: p.titel,
        vorname: p.vorname,
        nachname: left,
      };
    }
  }

  const tokens = s.split(" ").filter(Boolean);
  if (!tokens.length) return { anrede: "", titel: "", vorname: "", nachname: "" };

  let i = 0;
  let anrede = "";
  const t0 = tokens[0];
  if (/^herrn?$/i.test(t0) || /^hr\.?$/i.test(t0)) {
    anrede = "Herr";
    i = 1;
  } else if (/^frau$/i.test(t0) || /^fr\.?$/i.test(t0)) {
    anrede = "Frau";
    i = 1;
  }

  const titleParts = [];
  while (i < tokens.length && isTitleToken(tokens[i])) {
    titleParts.push(tokens[i]);
    i++;
  }
  const titel = titleParts.join(" ");

  const rest = tokens.slice(i);
  if (rest.length === 0) return { anrede, titel, vorname: "", nachname: "" };
  if (rest.length === 1) {
    if (!titel) return { anrede, titel, vorname: "", nachname: rest[0] };
    if (titleParts.length >= 2) return { anrede, titel, vorname: "", nachname: rest[0] };
    return { anrede, titel, vorname: rest[0], nachname: "" };
  }

  const particles = new Set(["von", "van", "de", "zu", "vom", "der", "den", "la", "le", "das"]);
  let last = rest.length - 1;
  let nachname = rest[last];
  if (last >= 1 && particles.has(rest[last - 1].toLowerCase())) {
    nachname = rest[last - 1] + " " + rest[last];
    last -= 2;
  } else {
    last -= 1;
  }
  const vorname = rest.slice(0, last + 1).join(" ");
  return { anrede, titel, vorname, nachname };
}

/** Kein Personennamen-Parsing für reine Abteilungs-/Organisationsbezeichnungen */
function shouldTreatAsPersonNameForParsing(s) {
  const t = String(s ?? "").trim();
  if (!t) return false;
  if (/^abteilung\b/i.test(t)) return false;
  if (/^(team|referat|bereich|zentral|verwaltung|vertrieb|einkauf|personal|produktion|lager|buchhaltung)\b/i.test(t))
    return false;
  if (/^(marketing|it|service|logistik|beschaffung|entwicklung|forschung|design|qualität|controlling|revision)\b/i.test(t))
    return false;
  if (/^(abteilungs|bereichs|fach|referats)/i.test(t)) return false;
  if (/\b(gmbh|ag|kg|ug|e\.?\s*v\.?|ohg|gbr|ltd|inc)\b/i.test(t)) return false;
  return true;
}

/**
 * Wenn Vor- oder Nachname-Spalte den vollen Namen enthält (z. B. „Herr Sigurd Romrod“ nur in Vorname),
 * in Anrede/Titel/Vor-/Nachname zerlegen.
 */
function mergeSplitNameFields(anredeIn, titelIn, vnIn, nnIn) {
  let anrede = String(anredeIn ?? "").trim();
  let titel = String(titelIn ?? "").trim();
  let vorname = String(vnIn ?? "").trim();
  let nachname = String(nnIn ?? "").trim();

  const tryParse = (s) => {
    if (!shouldTreatAsPersonNameForParsing(s)) return null;
    const p = parseNameCell(s);
    const hasStructure = !!(p.nachname || p.anrede || p.titel);
    if (hasStructure) return p;
    const parts = s.split(/\s+/).filter(Boolean);
    if (parts.length >= 2 && (p.vorname || p.nachname)) return p;
    return null;
  };

  if (vorname && !nachname) {
    const p = tryParse(vorname);
    if (p) {
      if (!anrede && p.anrede) anrede = p.anrede;
      if (!titel && p.titel) titel = p.titel;
      vorname = p.vorname;
      nachname = p.nachname;
    }
  } else if (!vorname && nachname) {
    const p = tryParse(nachname);
    if (p) {
      if (!anrede && p.anrede) anrede = p.anrede;
      if (!titel && p.titel) titel = p.titel;
      vorname = p.vorname;
      nachname = p.nachname;
    }
  } else if (vorname && nachname && /^(herr|frau|herrn|hr\.?|fr\.?)\s+/i.test(vorname)) {
    const p = tryParse(vorname);
    if (p && (p.vorname || p.nachname)) {
      if (!anrede && p.anrede) anrede = p.anrede;
      if (!titel && p.titel) titel = p.titel;
      vorname = p.vorname;
      nachname = p.nachname || nachname;
    }
  }

  return { anrede, titel, vorname, nachname };
}

/** Häufige deutsche Vornamen → Anrede, nur wenn Excel keine Anrede liefert */
const INFER_FEMALE_FIRST = new Set(
  "anna,maria,petra,sabine,ursula,beate,birgit,christine,eva,gabi,gabriele,heike,ingrid,karin,katja,kerstin,silke,stefanie,susanne,ute,andrea,angelika,anja,barbara,britta,claudia,dagmar,doris,edith,elisabeth,franziska,friederike,gerda,gisela,hannelore,helga,ilse,inge,johanna,judith,julia,kirsten,lisa,magdalena,margarete,marlene,martina,monika,nadine,nicole,nina,renate,rita,sandra,sigrid,sonja,tanja,vera,waltraud,yvonne,lotte,nora,emma,hannah,sophie,laura,mia,mila,lea,lina,finja,melanie,stephanie,katharina,elisabeth,miriam,simone,anja"
    .split(",")
    .map((x) => x.trim().toLowerCase())
);

const INFER_MALE_FIRST = new Set(
  "sigurd,klaus,hans,peter,wolfgang,jürgen,dieter,manfred,uwe,horst,gerhard,norbert,ralf,bernd,günter,heinz,helmut,herbert,johannes,karl,lothar,reinhard,rudolf,werner,achim,alexander,andreas,anton,armin,axel,bodo,christian,clemens,dietmar,erich,felix,florian,frank,fritz,georg,gert,günther,harald,heiko,henning,holger,jan,jens,joachim,jochen,jörg,kai,kurt,lars,ludwig,markus,martin,matthias,michael,olaf,otto,paul,rainer,richard,rolf,sebastian,stefan,sven,thomas,thorsten,torsten,ulrich,volker,wilhelm,wolfram,ingo,dirk,oliver,tim,tobias,benjamin,daniel,jonathan,maximilian,philipp,simon,edgar,erwin,guido,mehmet,mustafa"
    .split(",")
    .map((x) => x.trim().toLowerCase())
);

function inferAnredeFromVorname(vorname) {
  const raw = String(vorname ?? "").trim();
  if (!raw) return "";
  const first = raw.split(/\s+/)[0].replace(/\.$/, "");
  const k = first.toLowerCase();
  if (INFER_FEMALE_FIRST.has(k)) return "Frau";
  if (INFER_MALE_FIRST.has(k)) return "Herr";
  return "";
}

/** „10963 Berlin“ → PLZ + Ort */
function splitPlzOrt(s) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  if (!t) return { plz: "", stadt: "" };
  const m = t.match(/^(\d{5})\s+(.+)$/);
  if (m) return { plz: m[1], stadt: m[2].trim() };
  return { plz: "", stadt: t };
}

/** „Bernburger Straße 27“ / „Platz der Luftbrücke 6“ */
function splitStrasseHausnr(s) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  if (!t) return { strasse: "", hausnr: "" };
  const m = t.match(/^(.+?)\s+(\d+\s*[a-zA-ZüäöÜÄÖß]?)$/u);
  if (m) return { strasse: m[1].trim(), hausnr: m[2].trim() };
  return { strasse: t, hausnr: "" };
}

/**
 * Arbeitsblatt wählen: „Adressliste gesamt“ bevorzugen (wie BWK-Vorlage).
 */
function pickExternImportSheetName(wb) {
  const names = wb.SheetNames || [];
  const norm = (n) => String(n).replace(/\s+/g, " ").trim();
  let hit = names.find((n) => {
    const j = norm(n);
    return /adressliste.*gesamt/i.test(j) || /gesamt.*adressliste/i.test(j);
  });
  if (hit) return hit;
  hit = names.find((n) => /kooperationspartner/i.test(norm(n)));
  if (hit) return hit;
  hit = names.find((n) => /adressliste/i.test(norm(n)) && !/gästeliste\s*tn/i.test(norm(n)));
  if (hit) return hit;
  return names[0] || "";
}

/** TN-Blatt, z. B. „Gästeliste TN ZUSAGE ABSAGE“. */
function pickTnImportSheetName(wb) {
  const names = wb.SheetNames || [];
  const norm = (n) => String(n).replace(/\s+/g, " ").trim();
  let hit = names.find((n) => /gästeliste\s*tn|tn\s*zusage|tn\s*absage/i.test(norm(n)));
  if (hit) return hit;
  hit = names.find((n) => {
    const j = norm(n);
    return /tn/i.test(j) && /(zusage|absage|gästeliste)/i.test(j);
  });
  if (hit) return hit;
  hit = names.find((n) => /(^|\s)tn(\s|$)/i.test(norm(n)));
  if (hit) return hit;
  return names[0] || "";
}

/**
 * Zeile 2 in der Vorlage: nur Zähler (4, 1, 10, 33) in den Status-Spalten, kein Name.
 */
/** Datenzeile ohne jeden Inhalt (leere Strings / null; Zahlen/Booleans zählen als Inhalt). */
function importSheetRowIsCompletelyEmpty(cells) {
  for (const c of cells) {
    if (c === null || c === undefined) continue;
    if (typeof c === "boolean") return false;
    if (typeof c === "number") {
      if (!Number.isNaN(c)) return false;
      continue;
    }
    if (String(c).trim() !== "") return false;
  }
  return true;
}

/** Wenn die Datei eine Spalte „Lfd Nr“ hat: Zeile ohne Lfd oder Excel-Fehler (#…) überspringen. */
function importSheetHasLfdNr(cells, map) {
  if (map.lfdCol < 0) return true;
  const v = cells[map.lfdCol];
  if (v === null || v === undefined) return false;
  if (typeof v === "number" && Number.isFinite(v)) return true;
  const s = String(v).trim();
  if (!s) return false;
  if (/^#/i.test(s)) return false;
  return true;
}

function rowLooksLikeImportSummaryRow(cells, map) {
  const g = (i) => (i >= 0 && i < cells.length ? String(cells[i] ?? "").trim() : "");
  if (map.statusSingle >= 0) {
    const st = g(map.statusSingle);
    if (st && st.includes("|") && /\d/.test(st) && /zusagen|absagen|offen|einladen/i.test(st)) return true;
  }
  const boolIdx = [map.statusZusage, map.statusAbsage, map.statusOffen, map.statusEinladen].filter((i) => i >= 0);
  if (typeof map.statusNichtGeladen === "number" && map.statusNichtGeladen >= 0) boolIdx.push(map.statusNichtGeladen);
  if (boolIdx.length < 2) return false;
  const allNumericCounts = boolIdx.every((i) => {
    const v = cells[i];
    const n = typeof v === "number" && !Number.isNaN(v) ? v : Number(String(v).replace(",", ".").trim());
    return Number.isFinite(n) && n > 1;
  });
  if (!allNumericCounts) return false;
  const hasNameOrCompany =
    (map.full >= 0 && g(map.full)) ||
    (map.unternehmen >= 0 && g(map.unternehmen)) ||
    (map.vorname >= 0 && g(map.vorname)) ||
    (map.nachname >= 0 && g(map.nachname));
  return !hasNameOrCompany;
}

/**
 * Zelle als WAHR/FALSCH (Excel, deutsch/englisch, 1/0) lesen.
 * @returns {boolean|null} null = leer / nicht erkennbar
 */
function parseSpreadsheetBool(raw) {
  if (raw === true) return true;
  if (raw === false) return false;
  if (typeof raw === "number" && !Number.isNaN(raw)) {
    if (raw === 1) return true;
    if (raw === 0) return false;
  }
  const s = String(raw ?? "").trim().toLowerCase();
  if (!s) return null;
  if (/^(wahr|true|ja|j|yes|y|x)$/.test(s)) return true;
  if (/^(falsch|false|nein|n|no)$/.test(s)) return false;
  if (s === "1") return true;
  if (s === "0") return false;
  return null;
}

/**
 * Status aus Spalten Zusage / Einladen / Offen / Absage (je WAHR/FALSCH).
 * Bei mehreren WAHR: Zusage > Absage > Einladen > Offen.
 */
function statusFromBoolColumns(cells, map) {
  const candidates = [
    { idx: map.statusZusage, value: "zusage" },
    { idx: map.statusAbsage, value: "absage" },
    { idx: map.statusEinladen, value: "einladen" },
    { idx: map.statusOffen, value: "offen" },
  ];
  const priority = { zusage: 4, absage: 3, einladen: 2, offen: 1 };
  let best = "";
  let bestP = -1;
  for (const { idx, value } of candidates) {
    if (idx < 0) continue;
    if (parseSpreadsheetBool(cells[idx]) === true) {
      const p = priority[value] ?? 0;
      if (p > bestP) {
        bestP = p;
        best = value;
      }
    }
  }
  return best;
}

/** Excel-Dropdown „Zusage / Absage / offen / einladen“ (neue Vorlage) → interne Status-Werte */
function normalizeStatusDropdownCell(raw) {
  const s = String(raw ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
  if (!s) return "";
  if (s === "zusage") return "zusage";
  if (s === "absage") return "absage";
  if (s === "offen") return "offen";
  if (s === "einladen" || s === "einladung") return "einladen";
  return "";
}

/**
 * TN-Blatt: optional „Nicht geladen“ = WAHR → kein Status (leer).
 */
function statusFromBoolColumnsTn(cells, map) {
  const st = statusFromBoolColumns(cells, map);
  if (st) return st;
  if (map.statusNichtGeladen >= 0 && parseSpreadsheetBool(cells[map.statusNichtGeladen]) === true) return "";
  return "";
}

/**
 * Erste Tabellenzeile auswerten (u. a. BWK „Adressliste gesamt“: Unternehmen, Name, Straße Hausnr, PLZ Ort, …).
 * Optional: eine Spalte „Status“ mit Werten Zusage/Absage/offen/einladen (Excel-Dropdown) oder
 * getrennte Spalten Zusage / Absage / offen / einladen (WAHR/FALSCH) → Status-Segment.
 * TN-Blatt: Spalten „Name“ (= Nachname) + „Vorname“.
 */
function inferColumnMap(headers) {
  const map = {
    anrede: -1,
    titel: -1,
    vorname: -1,
    nachname: -1,
    full: -1,
    lfdCol: -1,
    unternehmen: -1,
    abteilung: -1,
    abteilungZusatz: -1,
    strasseHaus: -1,
    plzOrt: -1,
    plz: -1,
    stadt: -1,
    email: -1,
    anmerkungen: -1,
    statusSingle: -1,
    statusZusage: -1,
    statusEinladen: -1,
    statusOffen: -1,
    statusAbsage: -1,
  };
  const n = headers.length;
  for (let i = 0; i < n; i++) {
    const h = normalizeHeaderCell(headers[i]);
    if (!h) continue;
    if ((h.includes("lfd") && (h.includes("nr") || h.includes("nummer"))) || h === "lfd nr") map.lfdCol = i;
    else if (
      h === "unternehmen" ||
      h.includes("firmenname") ||
      h.includes("unternehmensname") ||
      (h.includes("unternehmen") && !h.includes("name"))
    )
      map.unternehmen = i;
    else if (h.includes("zusatz") && h.includes("abteilung")) map.abteilungZusatz = i;
    else if (h.includes("abteilung") && !h.includes("zusatz")) map.abteilung = i;
    else if ((h.includes("straße") || h.includes("strasse")) && (h.includes("haus") || h.includes("hausnr"))) map.strasseHaus = i;
    else if (h.includes("plz") && h.includes("ort")) map.plzOrt = i;
    else if (h === "plz" || h.includes("postleitzahl")) map.plz = i;
    else if (h === "stadt" || h === "ort" || h.includes("wohnort")) map.stadt = i;
    else if (h.includes("e-mail") || h === "email" || h.includes("email adresse") || h.includes("e-mail adresse"))
      map.email = i;
    else if (h.includes("anmerkung") || h.includes("bemerkung")) map.anmerkungen = i;
    else if (h === "status") map.statusSingle = i;
    else if (h === "anrede" || h.startsWith("anrede ")) map.anrede = i;
    else if (h === "titel" || /^titel(\s|$)/.test(h)) map.titel = i;
    else if (h.includes("vorname") && !h.includes("abteilung")) map.vorname = i;
    else if ((h.includes("nachname") || h.includes("familienname") || h === "zuname") && !h.includes("abteilung"))
      map.nachname = i;
  }

  let nameCol = -1;
  for (let i = 0; i < n; i++) {
    const h = normalizeHeaderCell(headers[i]);
    if (h === "name") {
      nameCol = i;
      break;
    }
  }
  if (map.vorname >= 0 && map.nachname < 0 && nameCol >= 0) {
    map.nachname = nameCol;
  } else if (map.vorname < 0 && map.nachname < 0 && nameCol >= 0) {
    map.full = nameCol;
  }

  if (map.full < 0) {
    for (let i = 0; i < n; i++) {
      const h = normalizeHeaderCell(headers[i]);
      if (!h) continue;
      if (h.includes("vorname") || h.includes("nachname") || h.includes("familienname")) continue;
      if ((h.includes("lfd") && (h.includes("nr") || h.includes("nummer"))) || h === "lfd nr") continue;
      if (h.includes("unternehmen") || h.includes("firma") || h.includes("firmenname")) continue;
      if (h.includes("email") || h.includes("e-mail") || h.includes("mail")) continue;
      if (h.includes("plz") && h.includes("ort")) continue;
      if (h === "plz" || h.includes("postleitzahl")) continue;
      if (h.includes("straße") || h === "strasse" || h.includes(" str.")) continue;
      if (h === "ort" || h === "stadt" || h.includes("wohnort")) continue;
      if (h === "name") continue;
      if (h.includes("abteilung") || h.includes("bereich") || h.includes("referat")) continue;
      if (h.includes("benutzer") || h.includes("login")) continue;
      if (h.includes("vollständig") && h.includes("name")) {
        map.full = i;
        break;
      }
      if (h.includes("teilnehmer") || h.includes("person") || (h.includes("kontakt") && !h.includes("daten"))) {
        map.full = i;
        break;
      }
      if (
        h.includes("name") &&
        !h.includes("vorname") &&
        !h.includes("nachname") &&
        !h.includes("firm") &&
        !h.includes("stadt")
      ) {
        map.full = i;
        break;
      }
    }
  }

  for (let i = 0; i < n; i++) {
    const h = normalizeHeaderCell(headers[i]);
    if (!h) continue;
    if (h === "zusage" && map.statusZusage < 0) map.statusZusage = i;
    else if ((h === "einladen" || h === "einladung") && map.statusEinladen < 0) map.statusEinladen = i;
    else if ((h === "offen" || h === "ausstehend") && map.statusOffen < 0) map.statusOffen = i;
    else if (h === "absage" && map.statusAbsage < 0) map.statusAbsage = i;
  }

  if (map.full < 0 && map.vorname < 0 && map.nachname < 0) {
    const reserved = new Set(
      [
        map.lfdCol,
        map.unternehmen,
        map.abteilung,
        map.abteilungZusatz,
        map.strasseHaus,
        map.plzOrt,
        map.plz,
        map.stadt,
        map.email,
        map.anmerkungen,
        map.statusSingle,
        map.anrede,
        map.titel,
        map.vorname,
        map.nachname,
        map.statusZusage,
        map.statusEinladen,
        map.statusOffen,
        map.statusAbsage,
      ].filter((idx) => idx >= 0)
    );
    for (let i = 0; i < n; i++) {
      if (!reserved.has(i)) {
        map.full = i;
        break;
      }
    }
  }
  return map;
}

/**
 * TN-Liste (z. B. Blatt „Gästeliste TN …“): Name = Nachname, Vorname, Status per WAHR/FALSCH, optional „Nicht geladen“.
 */
function inferTnImportColumnMap(headers) {
  const map = {
    anrede: -1,
    vorname: -1,
    nachname: -1,
    full: -1,
    lfdCol: -1,
    statusSingle: -1,
    statusZusage: -1,
    statusEinladen: -1,
    statusOffen: -1,
    statusAbsage: -1,
    statusNichtGeladen: -1,
  };
  const n = headers.length;
  for (let i = 0; i < n; i++) {
    const h = normalizeHeaderCell(headers[i]);
    if (!h) continue;
    if ((h.includes("lfd") && (h.includes("nr") || h.includes("nummer"))) || h === "lfd nr") map.lfdCol = i;
    else if (h === "status") map.statusSingle = i;
    else if (h.includes("nicht") && h.includes("geladen")) map.statusNichtGeladen = i;
    else if (h === "anrede" || h.startsWith("anrede ")) map.anrede = i;
    else if (h.includes("vorname") && !h.includes("abteilung")) map.vorname = i;
    else if ((h.includes("nachname") || h.includes("familienname") || h === "zuname") && !h.includes("abteilung"))
      map.nachname = i;
    else if (h === "zusage" && map.statusZusage < 0) map.statusZusage = i;
    else if ((h === "einladen" || h === "einladung") && map.statusEinladen < 0) map.statusEinladen = i;
    else if ((h === "offen" || h === "ausstehend") && map.statusOffen < 0) map.statusOffen = i;
    else if (h === "absage" && map.statusAbsage < 0) map.statusAbsage = i;
  }
  let nameCol = -1;
  for (let i = 0; i < n; i++) {
    if (normalizeHeaderCell(headers[i]) === "name") {
      nameCol = i;
      break;
    }
  }
  if (map.vorname >= 0 && map.nachname < 0 && nameCol >= 0) {
    map.nachname = nameCol;
  } else if (map.vorname < 0 && map.nachname < 0 && nameCol >= 0) {
    map.full = nameCol;
  }
  if (map.full < 0 && map.vorname < 0 && map.nachname < 0) {
    const reserved = new Set(
      [
        map.lfdCol,
        map.anrede,
        map.vorname,
        map.nachname,
        map.statusSingle,
        map.statusZusage,
        map.statusEinladen,
        map.statusOffen,
        map.statusAbsage,
        map.statusNichtGeladen,
      ].filter((idx) => idx >= 0)
    );
    for (let i = 0; i < n; i++) {
      if (!reserved.has(i)) {
        map.full = i;
        break;
      }
    }
  }
  return map;
}

function buildExternRowFromCells(cells, map) {
  const r = emptyRow();
  const g = (i) => (i >= 0 && i < cells.length ? String(cells[i] ?? "").replace(/\s+/g, " ").trim() : "");

  let an = map.anrede >= 0 ? g(map.anrede) : "";
  let ti = map.titel >= 0 ? g(map.titel) : "";
  let vn = map.vorname >= 0 ? g(map.vorname) : "";
  let nn = map.nachname >= 0 ? g(map.nachname) : "";
  const full = map.full >= 0 ? g(map.full) : "";

  const merged = mergeSplitNameFields(an, ti, vn, nn);
  an = merged.anrede;
  ti = merged.titel;
  vn = merged.vorname;
  nn = merged.nachname;

  const hasSplit = vn || nn || an || ti;
  if (hasSplit) {
    r.anrede = an;
    r.titel = ti;
    r.vorname = vn;
    r.nachname = nn;
    if (full && !r.vorname && !r.nachname) {
      const p = parseNameCell(full);
      if (!r.anrede) r.anrede = p.anrede;
      if (!r.titel) r.titel = p.titel;
      r.vorname = p.vorname;
      r.nachname = p.nachname;
    }
  } else if (full) {
    const p = parseNameCell(full);
    r.anrede = p.anrede;
    r.titel = p.titel;
    r.vorname = p.vorname;
    r.nachname = p.nachname;
  } else if (map.unternehmen >= 0 && g(map.unternehmen)) {
    r.unternehmen = g(map.unternehmen);
  } else {
    return null;
  }

  if (!r.anrede.trim()) {
    const guess = inferAnredeFromVorname(r.vorname);
    if (guess) r.anrede = guess;
  }

  const hasPerson = r.vorname.trim() || r.nachname.trim() || r.anrede.trim() || r.titel.trim();
  if (!hasPerson && !r.unternehmen.trim()) return null;

  if (map.unternehmen >= 0) r.unternehmen = g(map.unternehmen);
  if (map.abteilung >= 0 || map.abteilungZusatz >= 0) {
    let ab = map.abteilung >= 0 ? g(map.abteilung) : "";
    const zu = map.abteilungZusatz >= 0 ? g(map.abteilungZusatz) : "";
    if (zu) ab = ab ? `${ab} · ${zu}` : zu;
    r.abteilung = ab;
  }
  if (map.strasseHaus >= 0) {
    const sp = splitStrasseHausnr(g(map.strasseHaus));
    r.strasse = sp.strasse;
    r.hausnr = sp.hausnr;
  }
  if (map.plzOrt >= 0) {
    const po = splitPlzOrt(g(map.plzOrt));
    r.plz = po.plz;
    r.stadt = po.stadt;
  } else {
    if (map.plz >= 0) {
      const rawCell = map.plz < cells.length ? cells[map.plz] : "";
      let pv = g(map.plz);
      if (typeof rawCell === "number" && Number.isFinite(rawCell)) {
        pv = String(Math.round(rawCell));
      }
      r.plz = String(pv ?? "")
        .replace(/\s/g, "")
        .replace(/\..*$/, "")
        .slice(0, 5);
    }
    if (map.stadt >= 0) r.stadt = g(map.stadt);
  }
  const pz = String(r.plz || "").replace(/\s/g, "");
  if (pz.length === 5 && /^\d{5}$/.test(pz)) {
    r.stadt = resolveStadtFromPlzAndOrt(pz, r.stadt || "");
  }
  if (map.email >= 0) {
    const em = g(map.email);
    r.email = em.split(/\r?\n/)[0].trim();
  }
  if (map.anmerkungen >= 0) r.anmerkungen = g(map.anmerkungen);

  const stBool = statusFromBoolColumns(cells, map);
  if (stBool) r.status = stBool;
  else if (map.statusSingle >= 0) {
    const stDrop = normalizeStatusDropdownCell(g(map.statusSingle));
    if (stDrop) r.status = stDrop;
  }

  return migrateRow(r);
}

function buildTnRowFromCells(cells, map) {
  const r = emptyTnRow();
  const g = (i) => (i >= 0 && i < cells.length ? String(cells[i] ?? "").replace(/\s+/g, " ").trim() : "");

  let an = map.anrede >= 0 ? g(map.anrede) : "";
  let vn = map.vorname >= 0 ? g(map.vorname) : "";
  let nn = map.nachname >= 0 ? g(map.nachname) : "";
  const full = map.full >= 0 ? g(map.full) : "";

  const merged = mergeSplitNameFields(an, "", vn, nn);
  an = merged.anrede;
  vn = merged.vorname;
  nn = merged.nachname;

  if (vn || nn || an) {
    r.anrede = an;
    r.vorname = vn;
    r.nachname = nn;
    if (full && !r.vorname && !r.nachname) {
      const p = parseNameCell(full);
      r.anrede = r.anrede || p.anrede;
      r.vorname = p.vorname;
      r.nachname = p.nachname;
    }
  } else if (full) {
    const p = parseNameCell(full);
    r.anrede = p.anrede;
    r.vorname = p.vorname;
    r.nachname = p.nachname;
  } else {
    return null;
  }

  if (!r.anrede.trim()) {
    const guess = inferAnredeFromVorname(r.vorname);
    if (guess) r.anrede = guess;
  }

  if (!r.vorname.trim() && !r.nachname.trim() && !r.anrede.trim()) return null;

  let st = statusFromBoolColumnsTn(cells, map);
  if (!st && map.statusSingle >= 0) st = normalizeStatusDropdownCell(g(map.statusSingle));
  r.status = st || "";
  return migrateTnRow(r);
}

async function loadSpreadsheetMatrix(file, kind = "extern") {
  if (typeof XLSX === "undefined") throw new Error("Tabellen-Bibliothek nicht geladen (XLSX). Bitte Seite neu laden.");
  const name = (file.name || "").toLowerCase();
  let wb;
  if (name.endsWith(".csv")) {
    const text = await file.text();
    wb = XLSX.read(text, { type: "string", raw: true });
  } else {
    const buf = await file.arrayBuffer();
    wb = XLSX.read(buf, { type: "array" });
  }
  if (!wb.SheetNames || !wb.SheetNames.length) throw new Error("Leere Arbeitsmappe.");
  const sheetName = kind === "tn" ? pickTnImportSheetName(wb) : pickExternImportSheetName(wb);
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
  return { matrix, sheetName };
}

async function importNamesFromSpreadsheet(file) {
  const targetTn = listSubTab === "tn";
  const { matrix, sheetName } = await loadSpreadsheetMatrix(file, targetTn ? "tn" : "extern");
  if (!matrix.length) throw new Error("Keine Zeilen in der Datei.");

  if (currentTab !== "liste") switchTab("liste");

  const headers = matrix[0].map((c) => String(c ?? "").trim());
  const map = targetTn ? inferTnImportColumnMap(headers) : inferColumnMap(headers);

  const newRows = [];
  for (let r = 1; r < matrix.length; r++) {
    if (newRows.length >= SHEET_IMPORT_MAX_ROWS) break;
    const row = matrix[r];
    if (!row || !row.length) continue;
    const padded = row.map((c) => (c == null ? "" : c));
    if (importSheetRowIsCompletelyEmpty(padded)) continue;
    if (!importSheetHasLfdNr(padded, map)) continue;
    if (rowLooksLikeImportSummaryRow(padded, map)) continue;
    const built = targetTn ? buildTnRowFromCells(padded, map) : buildExternRowFromCells(padded, map);
    if (built) newRows.push(built);
  }

  if (!newRows.length) {
    setStatus("Keine importierbaren Zeilen gefunden (erste Zeile = Überschriften).");
    return;
  }

  const keyFn = targetTn ? tnContactMatchKey : externContactMatchKey;
  const toProcess = dedupeRowsByContactKey(newRows, keyFn);
  const dupInFile = newRows.length - toProcess.length;

  const newOnes = [];
  const updates = [];
  let unchangedExisting = 0;

  if (hasMeaningfulData()) {
    for (const r of toProcess) {
      const k = keyFn(r);
      const idx = targetTn ? findFirstTnRowIndexByKey(k) : findFirstExternRowIndexByKey(k);
      if (idx < 0) {
        newOnes.push(r);
        continue;
      }
      const existing = targetTn ? rowsTN[idx] : rowsExtern[idx];
      const changes = targetTn ? diffTnRowAgainstImport(existing, r) : diffExternRowAgainstImport(existing, r);
      if (!changes.length) {
        unchangedExisting += 1;
        continue;
      }
      updates.push({ index: idx, incoming: r, changes });
    }
  } else {
    for (const r of toProcess) newOnes.push(r);
  }

  if (!newOnes.length && !updates.length) {
    const parts = [`Keine neuen oder geänderten Daten · Blatt „${sheetName}“`];
    if (unchangedExisting) parts.push(`${unchangedExisting} Zeilen unverändert zur Liste`);
    if (dupInFile) parts.push(`${dupInFile} doppelt in der Datei`);
    setStatus(parts.join(" · ") + ".");
    return;
  }

  const listLabel = targetTn ? "TN-Liste" : "externe Liste";
  if (hasMeaningfulData()) {
    const ok = confirm(
      buildImportConfirmText(
        sheetName,
        listLabel,
        newOnes.length,
        updates,
        dupInFile,
        unchangedExisting,
        targetTn,
      ),
    );
    if (!ok) {
      setStatus("Import abgebrochen.");
      return;
    }
  }

  pushUndoBeforeMutation();
  for (const u of updates) {
    if (targetTn) applyTnRowFromImport(rowsTN[u.index], u.incoming);
    else applyExternRowFromImport(rowsExtern[u.index], u.incoming);
  }
  if (targetTn) {
    rowsTN.push(...newOnes);
    trimEmptyTnRows();
  } else {
    rowsExtern.push(...newOnes);
    trimEmptyExternRows();
  }
  saveAll();
  render();
  const bits = [];
  if (newOnes.length) bits.push(`${newOnes.length} neu`);
  if (updates.length) bits.push(`${updates.length} aktualisiert`);
  if (unchangedExisting) bits.push(`${unchangedExisting} unverändert`);
  let statusMsg = `Import · ${bits.join(" · ")} · Blatt „${sheetName}“ · ${listLabel}`;
  if (dupInFile) statusMsg += ` · ${dupInFile} doppelt in der Datei`;
  const rowsForNameDup = targetTn ? rowsTN : rowsExtern;
  const nNameDupPair = indicesWithDuplicateNamePairs(
    rowsForNameDup,
    targetTn ? (r) => tnRowHasText(r) : (r) => rowStatus(r) || rowHasText(r),
  ).size;
  if (nNameDupPair) {
    statusMsg += ` · ${nNameDupPair} Zeile(n) mit gleichem Vor- und Nachnamen wie eine andere (orange markiert)`;
  }
  const statusLine = statusMsg + ".";
  setStatus(statusLine);
  maybePromptRemoveDuplicatesAfterSpreadsheetImport(targetTn, statusLine);
}

async function handleImportFile(file) {
  const text = await file.text();
  const data = parseJsonFile(text);
  if (hasMeaningfulData()) {
    const ok = confirm("Aktuelle Daten durch die Import-Datei ersetzen?");
    if (!ok) {
      setStatus("Import abgebrochen.");
      return;
    }
  }
  pushUndoBeforeMutation();
  applyImportedPayload(data, "Import abgeschlossen.");
}

function switchTab(name) {
  currentTab = name;
  const pListe = document.getElementById("panel-liste");
  const pSer = document.getElementById("panel-serienbrief");
  const bListe = document.getElementById("tab-btn-liste");
  const bSer = document.getElementById("tab-btn-serienbrief");

  if (name === "liste") {
    pListe.classList.remove("hidden");
    pListe.hidden = false;
    pSer.classList.add("hidden");
    pSer.hidden = true;
    bListe.classList.add("active");
    bListe.setAttribute("aria-selected", "true");
    bSer.classList.remove("active");
    bSer.setAttribute("aria-selected", "false");
    render();
  } else {
    pSer.classList.remove("hidden");
    pSer.hidden = false;
    pListe.classList.add("hidden");
    pListe.hidden = true;
    bSer.classList.add("active");
    bSer.setAttribute("aria-selected", "true");
    bListe.classList.remove("active");
    bListe.setAttribute("aria-selected", "false");
    renderSerienTable();
  }
}

function onPageHide() {
  readEventFieldsFromDom();
  flushScheduledSave();
  writeStorageNow();
}

/** Spalten-Reihenfolge für Pfeiltasten (wie Excel) */
const EXTERN_COL_KEYS = [
  "status",
  "unternehmen",
  "anrede",
  "titel",
  "vorname",
  "nachname",
  "email",
  "abteilung",
  "strasse",
  "hausnr",
  "plz",
  "stadt",
  "anmerkungen",
];

const TN_COL_KEYS = ["status", "anrede", "vorname", "nachname"];

function getVisibleTnIndices() {
  return Array.from({ length: rowsTN.length }, (_, i) => i);
}

function focusExternCell(rowIndexStr, colKey) {
  requestAnimationFrame(() => {
    const tr = document.querySelector(`#liste-tbody tr[data-row-index="${rowIndexStr}"]`);
    if (!tr) return;
    if (colKey === "status") {
      const inp =
        tr.querySelector("input.status-segment-input:checked") ||
        tr.querySelector("input.status-segment-input");
      inp?.focus();
      return;
    }
    const inp = tr.querySelector(`input.cell-input[data-field="${colKey}"]`);
    if (inp) {
      inp.focus();
      inp.select();
    }
  });
}

function focusTnCell(rowIndexStr, colKey) {
  requestAnimationFrame(() => {
    const tr = document.querySelector(`#tn-tbody tr[data-row-index="${rowIndexStr}"]`);
    if (!tr) return;
    if (colKey === "status") {
      const inp =
        tr.querySelector("input.status-segment-input:checked") ||
        tr.querySelector("input.status-segment-input");
      inp?.focus();
      return;
    }
    const inp = tr.querySelector(`input.cell-input[data-field="${colKey}"]`);
    if (inp) {
      inp.focus();
      inp.select();
    }
  });
}

function focusAdjacentStatusSegment(radio, dir) {
  const name = radio.name;
  const radios = Array.from(document.querySelectorAll(`input.status-segment-input[name="${CSS.escape(name)}"]`));
  if (!radios.length) return false;
  const idx = radios.indexOf(radio);
  if (idx < 0) return false;
  if (dir === "left" && idx > 0) {
    radios[idx - 1].focus();
    return true;
  }
  if (dir === "right" && idx < radios.length - 1) {
    radios[idx + 1].focus();
    return true;
  }
  return false;
}

function textAtLeftBoundary(el) {
  return el.selectionStart === 0 && el.selectionEnd === 0;
}

function textAtRightBoundary(el) {
  const len = el.value.length;
  return el.selectionStart === len && el.selectionEnd === len;
}

function handleSpreadsheetArrowNav(e) {
  if (!["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)) return;
  if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;
  if (currentTab !== "liste") return;

  const t = e.target;
  if (!(t instanceof HTMLInputElement)) return;

  if (isListSubTabExternLike()) {
    if (!t.closest("#liste-tbody")) return;
    const tr = t.closest("tr[data-row-index]");
    if (!tr) return;
    const rowIndexStr = tr.dataset.rowIndex;
    const i = Number(rowIndexStr);
    const vis = getFilteredIndices();
    const p = vis.indexOf(i);
    if (p < 0) return;

    if (t.classList.contains("status-segment-input")) {
      if (e.key === "ArrowLeft") {
        if (focusAdjacentStatusSegment(t, "left")) {
          e.preventDefault();
          return;
        }
        if (p > 0) {
          e.preventDefault();
          focusExternCell(String(vis[p - 1]), "anmerkungen");
        }
        return;
      }
      if (e.key === "ArrowRight") {
        if (focusAdjacentStatusSegment(t, "right")) {
          e.preventDefault();
          return;
        }
        e.preventDefault();
        focusExternCell(rowIndexStr, "unternehmen");
        return;
      }
      if (e.key === "ArrowUp" && p > 0) {
        e.preventDefault();
        focusExternCell(String(vis[p - 1]), "status");
        return;
      }
      if (e.key === "ArrowDown" && p < vis.length - 1) {
        e.preventDefault();
        focusExternCell(String(vis[p + 1]), "status");
        return;
      }
      return;
    }

    if (!t.classList.contains("cell-input") || !t.dataset.field) return;
    const field = t.dataset.field;
    const fi = EXTERN_COL_KEYS.indexOf(field);
    if (fi < 1) return;

    if (e.key === "ArrowUp") {
      if (p > 0) {
        e.preventDefault();
        focusExternCell(String(vis[p - 1]), field);
      }
      return;
    }
    if (e.key === "ArrowDown") {
      if (p < vis.length - 1) {
        e.preventDefault();
        focusExternCell(String(vis[p + 1]), field);
      }
      return;
    }
    if (e.key === "ArrowLeft") {
      if (!textAtLeftBoundary(t)) return;
      e.preventDefault();
      if (fi > 1) {
        focusExternCell(rowIndexStr, EXTERN_COL_KEYS[fi - 1]);
      } else {
        focusExternCell(rowIndexStr, "status");
      }
      return;
    }
    if (e.key === "ArrowRight") {
      if (!textAtRightBoundary(t)) return;
      e.preventDefault();
      if (fi < EXTERN_COL_KEYS.length - 1) {
        focusExternCell(rowIndexStr, EXTERN_COL_KEYS[fi + 1]);
      } else if (p < vis.length - 1) {
        focusExternCell(String(vis[p + 1]), "status");
      }
      return;
    }
    return;
  }

  if (listSubTab === "tn") {
    if (!t.closest("#tn-tbody")) return;
    const tr = t.closest("tr[data-row-index]");
    if (!tr) return;
    const rowIndexStr = tr.dataset.rowIndex;
    const i = Number(rowIndexStr);
    const vis = getVisibleTnIndices();
    const p = vis.indexOf(i);
    if (p < 0) return;

    if (t.classList.contains("status-segment-input")) {
      if (e.key === "ArrowLeft") {
        if (focusAdjacentStatusSegment(t, "left")) {
          e.preventDefault();
          return;
        }
        if (p > 0) {
          e.preventDefault();
          focusTnCell(String(vis[p - 1]), "nachname");
        }
        return;
      }
      if (e.key === "ArrowRight") {
        if (focusAdjacentStatusSegment(t, "right")) {
          e.preventDefault();
          return;
        }
        e.preventDefault();
        focusTnCell(rowIndexStr, "anrede");
        return;
      }
      if (e.key === "ArrowUp" && p > 0) {
        e.preventDefault();
        focusTnCell(String(vis[p - 1]), "status");
        return;
      }
      if (e.key === "ArrowDown" && p < vis.length - 1) {
        e.preventDefault();
        focusTnCell(String(vis[p + 1]), "status");
        return;
      }
      return;
    }

    if (!t.classList.contains("cell-input") || !t.dataset.field) return;
    const field = t.dataset.field;
    const fi = TN_COL_KEYS.indexOf(field);
    if (fi < 1) return;

    if (e.key === "ArrowUp") {
      if (p > 0) {
        e.preventDefault();
        focusTnCell(String(vis[p - 1]), field);
      }
      return;
    }
    if (e.key === "ArrowDown") {
      if (p < vis.length - 1) {
        e.preventDefault();
        focusTnCell(String(vis[p + 1]), field);
      }
      return;
    }
    if (e.key === "ArrowLeft") {
      if (!textAtLeftBoundary(t)) return;
      e.preventDefault();
      if (fi > 1) {
        focusTnCell(rowIndexStr, TN_COL_KEYS[fi - 1]);
      } else {
        focusTnCell(rowIndexStr, "status");
      }
      return;
    }
    if (e.key === "ArrowRight") {
      if (!textAtRightBoundary(t)) return;
      e.preventDefault();
      if (fi < TN_COL_KEYS.length - 1) {
        focusTnCell(rowIndexStr, TN_COL_KEYS[fi + 1]);
      } else if (p < vis.length - 1) {
        focusTnCell(String(vis[p + 1]), "status");
      }
      return;
    }
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  FILTER_KEYS.forEach((k) => {
    filters[k] = "";
  });

  await loadPlzMap();
  loadFromStorage();
  applyAppTitleToDom();
  loadEventFieldsToDom();
  switchListSubTab(listSubTab);
  refreshSerienTemplateLabels();
  refreshExportDirLabel();

  const dlgImport = document.getElementById("dlg-import");
  const fileInput = document.getElementById("file-import-json");

  if (!hasMeaningfulData()) {
    const fb = localStorage.getItem(FAILSAFE_KEY);
    if (fb && backupLooksFilled(fb)) {
      if (confirm("Automatisches Browser-Backup wiederherstellen?")) tryRestoreFailsafe();
    }
    if (!hasMeaningfulData() && dlgImport) dlgImport.showModal();
  }

  document.getElementById("dlg-import-json")?.addEventListener("click", () => {
    dlgImport?.close();
    fileInput?.click();
  });
  document.getElementById("dlg-import-extern")?.addEventListener("click", () => {
    dlgImport?.close();
    switchTab("liste");
    switchListSubTab("extern");
    document.getElementById("file-import-sheet")?.click();
  });
  document.getElementById("dlg-import-skip")?.addEventListener("click", () => dlgImport?.close());

  fileInput?.addEventListener("change", async () => {
    const f = fileInput.files?.[0];
    fileInput.value = "";
    if (!f) return;
    try {
      await handleImportFile(f);
    } catch (err) {
      setStatus("Fehler: " + (err && err.message ? err.message : String(err)));
    }
  });

  const titleH = document.getElementById("app-title-heading");
  const titleInp = document.getElementById("app-title-input");
  titleH?.addEventListener("click", (e) => {
    e.preventDefault();
    showAppTitleInput();
  });
  titleH?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      showAppTitleInput();
    }
  });
  titleInp?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      titleInp.blur();
    } else if (e.key === "Escape") {
      e.preventDefault();
      finishAppTitleEdit(false);
    }
  });
  titleInp?.addEventListener("blur", () => finishAppTitleEdit(true));

  document.getElementById("btn-clear-filters")?.addEventListener("click", () => {
    pushUndoBeforeMutation();
    FILTER_KEYS.forEach((k) => {
      filters[k] = "";
    });
    render();
  });

  document.getElementById("list-subtabs-sub-wrap")?.addEventListener("click", (e) => {
    const el = e.target instanceof Element ? e.target : null;
    if (!el) return;
    const rm = el.closest("button.list-subtab-remove");
    if (rm && rm instanceof HTMLButtonElement) {
      e.preventDefault();
      e.stopPropagation();
      const extraRm = rm.dataset.extraRemove;
      if (extraRm) {
        removeExtraSublist(extraRm);
        return;
      }
      const sub = rm.dataset.sub;
      if (sub === "extern" || sub === "tn") hideSublist(sub);
      return;
    }
    const tab = el.closest("button[data-list-sub]");
    if (tab && tab instanceof HTMLButtonElement) {
      const sub = tab.dataset.listSub;
      if (sub === "extern" || sub === "tn" || (typeof sub === "string" && sub.startsWith("extra:"))) {
        switchListSubTab(sub);
      }
    }
  });

  document.getElementById("btn-sublist-add")?.addEventListener("click", () => createNewExtraSublist());

  document.getElementById("list-set-buttons")?.addEventListener("click", (e) => {
    const t = e.target;
    const el = t instanceof Element ? t : t && t.parentNode instanceof Element ? t.parentNode : null;
    if (!el) return;
    const rm = el.closest("button.list-set-tab-remove");
    if (rm && rm instanceof HTMLButtonElement) {
      e.preventDefault();
      e.stopPropagation();
      const rid = rm.dataset.listId;
      if (rid) deleteListById(rid);
      return;
    }
    const tabBtn = el.closest("button.list-set-tab[data-list-id]");
    if (!tabBtn || !(tabBtn instanceof HTMLButtonElement)) return;
    const id = tabBtn.dataset.listId;
    if (id) switchActiveList(id);
  });
  document.getElementById("btn-list-add")?.addEventListener("click", () => createNewList());

  document.getElementById("btn-add")?.addEventListener("click", () => addRowByContext());

  document.getElementById("btn-undo")?.addEventListener("click", () => undoLastAction());
  document.getElementById("btn-redo")?.addEventListener("click", () => redoLastAction());
  document.addEventListener("keydown", onGlobalUndoRedoKeydown);

  document.addEventListener("keydown", onListeShiftEnter);
  document.addEventListener("keydown", handleCellEnterComplete, true);
  document.addEventListener("keydown", handleSpreadsheetArrowNav, true);
  document.getElementById("btn-export-liste")?.addEventListener("click", () => {
    document.getElementById("dlg-export-liste")?.showModal();
  });
  document.getElementById("dlg-export-liste-csv")?.addEventListener("click", () => {
    document.getElementById("dlg-export-liste")?.close();
    exportListeCsv();
  });
  document.getElementById("dlg-export-liste-xlsx")?.addEventListener("click", async () => {
    document.getElementById("dlg-export-liste")?.close();
    try {
      await exportListeXlsx();
    } catch (e) {
      setStatus("Excel-Export: " + (e && e.message ? e.message : String(e)));
    }
  });
  document.getElementById("dlg-export-liste-json")?.addEventListener("click", () => {
    document.getElementById("dlg-export-liste")?.close();
    exportJsonDownload();
  });
  document.getElementById("dlg-export-liste-close")?.addEventListener("click", () => {
    document.getElementById("dlg-export-liste")?.close();
  });
  document.getElementById("btn-import-daten")?.addEventListener("click", () => {
    document.getElementById("dlg-import")?.showModal();
  });
  const fileImportSheet = document.getElementById("file-import-sheet");
  fileImportSheet?.addEventListener("change", async () => {
    const f = fileImportSheet.files?.[0];
    fileImportSheet.value = "";
    if (!f) return;
    try {
      await importNamesFromSpreadsheet(f);
    } catch (err) {
      setStatus("Import: " + (err && err.message ? err.message : String(err)));
    }
  });
  document.getElementById("btn-clear")?.addEventListener("click", clearAll);

  document.getElementById("panel-serienbrief")?.addEventListener("change", (e) => {
    if (e.target && e.target.name === "serien-export-what") {
      readEventFieldsFromDom();
      saveEventOnly();
    }
  });

  [
    "ev-datum",
    "ev-zeit",
    "ev-ort",
    "ev-termin-zeile",
    "ev-briefdatum",
    "serien-export-prefix",
    "serien-mail-subject",
    "serien-mail-from",
  ].forEach((id) => {
    document.getElementById(id)?.addEventListener("change", () => {
      pushUndoBeforeMutation();
      readEventFieldsFromDom();
      saveEventOnly();
    });
    document.getElementById(id)?.addEventListener("blur", () => {
      readEventFieldsFromDom();
      saveEventOnly();
    });
  });

  document.getElementById("btn-template-docx")?.addEventListener("click", () => {
    document.getElementById("file-template-docx")?.click();
  });
  document.getElementById("btn-template-mail")?.addEventListener("click", () => {
    document.getElementById("file-template-mail")?.click();
  });
  document.getElementById("file-template-docx")?.addEventListener("change", async () => {
    const inp = document.getElementById("file-template-docx");
    const f = inp?.files?.[0];
    if (inp) inp.value = "";
    if (!f) return;
    try {
      const buffer = await f.arrayBuffer();
      if (isOleWordDocBuffer(buffer)) {
        alert(
          "Alte Word-Dateien (.doc / .dot, OLE) können hier nicht als Vorlage genutzt werden. " +
            "Bitte in Word „Speichern unter“: „Word-Vorlage (*.dotx)“ (empfohlen) oder „Word-Dokument (*.docx)“, dann erneut laden.",
        );
        setStatusSerien("Vorlage: bitte als .dotx oder .docx speichern (kein klassisches .doc/.dot).");
        return;
      }
      if (!isWordOoxmlZipBuffer(buffer)) {
        alert(
          "Die Datei ist keine gültige Word-OOXML-Datei (.dotx oder .docx). " +
            "Vorlage idealerweise als „Word-Vorlage (*.dotx)“ anlegen oder ein Dokument als .docx speichern.",
        );
        setStatusSerien("Vorlage: ungültig — erwartet wird OOXML (.dotx/.docx, ZIP).");
        return;
      }
      await serienIdbSet("tpl_docx", { buffer, name: f.name });
      refreshSerienTemplateLabels();
      setStatusSerien("Word-Vorlage gespeichert (" + f.name + ").");
    } catch (err) {
      setStatusSerien("Vorlage: " + (err && err.message ? err.message : String(err)));
    }
  });
  document.getElementById("file-template-mail")?.addEventListener("change", async () => {
    const inp = document.getElementById("file-template-mail");
    const f = inp?.files?.[0];
    if (inp) inp.value = "";
    if (!f) return;
    try {
      const buffer = await f.arrayBuffer();
      if (isOleWordDocBuffer(buffer)) {
        alert(
          "Alte Word-Dateien (.doc / .dot, OLE) können nicht als E-Mail-Vorlage genutzt werden. " +
            "Bitte „Speichern unter“: „Word-Dokument (*.docx)“ oder „Word-Vorlage (*.dotx)“.",
        );
        setStatusSerien("E-Mail-Vorlage: bitte .docx/.dotx (kein klassisches .doc/.dot).");
        return;
      }
      let kind = "text";
      if (isWordOoxmlZipBuffer(buffer)) {
        kind = "docx";
      } else {
        const lower = (f.name || "").toLowerCase();
        if (lower.endsWith(".docx") || lower.endsWith(".dotx")) {
          alert("Die Datei wirkt beschädigt oder ist kein Word-OOXML (ZIP). Bitte gültige .docx/.dotx wählen.");
          setStatusSerien("E-Mail-Vorlage: Datei nicht lesbar (kein OOXML).");
          return;
        }
      }
      await serienIdbSet("tpl_mail", { buffer, name: f.name, kind });
      refreshSerienTemplateLabels();
      setStatusSerien(
        "E-Mail-Vorlage gespeichert (" +
          f.name +
          (kind === "docx" ? ", Word → Text für .eml" : ", Text") +
          ").",
      );
    } catch (err) {
      setStatusSerien("Vorlage: " + (err && err.message ? err.message : String(err)));
    }
  });
  document.getElementById("btn-export-dir")?.addEventListener("click", () => {
    pickAndStoreExportDirectory().catch((err) => {
      setStatusSerien("Ordner: " + (err && err.message ? err.message : String(err)));
    });
  });
  document.getElementById("btn-export-serien-files")?.addEventListener("click", () => {
    exportSerienBriefFiles().catch((err) => {
      console.error(err);
      setStatusSerien("Export: " + (err && err.message ? err.message : String(err)));
      alert(err && err.message ? err.message : String(err));
    });
  });

  document.getElementById("tab-btn-liste")?.addEventListener("click", () => switchTab("liste"));
  document.getElementById("tab-btn-serienbrief")?.addEventListener("click", () => switchTab("serienbrief"));

  document.getElementById("serien-filter-wrap")?.addEventListener("change", (e) => {
    if (e.target && e.target.name === "serien-filter-status") renderSerienTable();
  });

  document.getElementById("btn-select-all-serien")?.addEventListener("click", () => {
    pushUndoBeforeMutation();
    const statusFilter = getSerienStatusFilter();
    getFilteredIndices()
      .filter((i) => !statusFilter || rowStatus(rowsExtern[i]) === statusFilter)
      .forEach((i) => selectedIds.add(rowsExtern[i].id));
    renderSerienTable();
    setStatusSerien(selectedIds.size + " ausgewählt.");
  });

  document.getElementById("btn-select-none-serien")?.addEventListener("click", () => {
    pushUndoBeforeMutation();
    selectedIds.clear();
    renderSerienTable();
    setStatusSerien("Auswahl leer.");
  });

  document.getElementById("btn-export-merge")?.addEventListener("click", exportMergeCsv);

  window.addEventListener("pagehide", onPageHide);
  window.addEventListener("beforeunload", onPageHide);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) onPageHide();
  });

  render();
});
