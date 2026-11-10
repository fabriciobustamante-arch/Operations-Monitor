"use strict";
/* ============ Estado ============ */
const BOOT = JSON.parse(document.getElementById("boot").textContent);
const MAESTRO = new Map(BOOT.maestro.map(r => [r[0], {empresa: r[1], tipo: r[2], source: r[3], central: r[4]}]));
const DOW = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"];
const GROUPS = [
  {k: "Hidro", src: ["4. HYDRO"], v: "--s-hidro"},
  {k: "Gas natural", src: ["5. GAS", "6. GAS/DIESEL"], v: "--s-gas"},
  {k: "Eólica", src: ["2. WIND"], v: "--s-eol"},
  {k: "Solar", src: ["3. SOLAR"], v: "--s-sol"},
  {k: "Biomasa", src: ["1. BIOMASS"], v: "--s-bio"},
  {k: "Diésel, residual y carbón", src: ["7. COAL", "8. DIESEL", "9. DIESEL/R500", "91. R500"], v: "--s-die"},
  {k: "Sin clasificar", src: [], v: "--s-sin"}
];
const TIPOS = ["HIDROELÉCTRICA", "TERMOELÉCTRICA", "EÓLICA", "SOLAR"];
const BARRAS = [{k: "Santa Rosa", z: "Centro", v: "--b-centro"}, {k: "Trujillo", z: "Norte", v: "--b-norte"}, {k: "Socabaya", z: "Sur", v: "--b-sur"}];
const $ = id => document.getElementById(id);
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const sum = a => a.reduce((s, x) => s + (x || 0), 0);
const fmt = (x, d = 0) => x == null || isNaN(x) ? "—" : Number(x).toLocaleString("es-PE", {minimumFractionDigits: d, maximumFractionDigits: d});
const sgn = (x, d = 0) => (x > 0 ? "+" : "") + fmt(x, d);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
function groupOf(central) {
  const m = MAESTRO.get(central);
  if (!m) return "Sin clasificar";
  const g = GROUPS.find(g => g.src.includes(m.source));
  return g ? g.k : "Sin clasificar";
}
const tipoOf = c => MAESTRO.get(c)?.tipo || null;

/* ============ Semana embebida ============ */
function fromEmbedded(e) {
  return {semana: e.semana, anio: e.anio, inicio: e.inicio, fin: e.fin, fuente: e.fuente, times: e.times,
    centrales: e.centrales, cmg: e.cmg, re: e.re,
    tot: {"Demanda YUPANA": e.dr["Demanda YUPANA"], "Total CoGeneracion": e.dr["Total CoGeneracion"], "Total Renovables": e.dr["Total Renovables"]}};
}

/* ============ Lectura de Excel del COES (misma lógica que el Power Query) ============ */
const EXCL = ["Cmg", "YUPANA", "FLUJO", "Flujo", "Demanda", "Oferta", "Total", "Reserva", "NCP"];
function timeMinutes(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") { if (v < 0 || v > 1.01) return null; return Math.round(v * 1440); }
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  const iso = String(v).match(/T(\d{2}):(\d{2})/); if (iso) return (+iso[1]) * 60 + (+iso[2]);
  const m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : null;
}
function slotOf(min) { // 0..47 ; 23:59, 24:00 y 00:00 = último intervalo del día
  if (min == null) return null;
  if (min === 0 || min === 1439 || min === 1440) return 47;
  if (min % 30 !== 0) return null;
  const s = min / 30 - 1; return s >= 0 && s < 48 ? s : null;
}
function dateFromText(t) {
  const iso = String(t).match(/(\d{4})-(\d{2})-(\d{2})/); if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const m = String(t).match(/(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
  if (!m) return null;
  return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
}
function addDays(iso, n) { const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); }
function labelsFor(start) {
  const out = [];
  for (let d = 0; d < 7; d++) { const day = addDays(start, d); for (let s = 0; s < 48; s++) { const mins = (s + 1) * 30; out.push(day + "T" + (mins === 1440 ? "23:59" : String(Math.floor(mins / 60)).padStart(2, "0") + ":" + String(mins % 60).padStart(2, "0"))); } }
  return out;
}
/* Lee una hoja del COES: busca la fila "Día Hora" (encabezados); 2 filas arriba = empresa, 1 arriba = código.
   Devuelve {slots: {0..47: {columna: valor}}, meta: {columna: {empresa, codigo}}, titulo, fecha} */
function readSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null});
  const hi = rows.findIndex(r => (r || []).some(v => v != null && String(v).trim() === "Día Hora"));
  if (hi < 0) return null;
  const hdr = rows[hi].map(h => h == null ? null : String(h).trim()), tc = hdr.indexOf("Día Hora");
  const emp = rows[hi - 2] || [], cod = rows[hi - 1] || [], meta = {}; let lastEmp = "";
  for (let c = tc + 1; c < hdr.length; c++) { if (emp[c] != null && String(emp[c]).trim()) lastEmp = String(emp[c]).trim(); if (hdr[c]) meta[hdr[c]] = {empresa: lastEmp, codigo: cod[c] ?? null}; }
  const top = rows.slice(0, hi).flat().filter(v => v != null).map(String);
  const titulo = top.find(v => /PROGRAMACION|REPROGRAMA/i.test(v)) || "", fecha = top.map(dateFromText).find(Boolean) || null;
  const slots = {};
  for (const r of rows.slice(hi + 1)) {
    const s = slotOf(timeMinutes(r[tc])); if (s == null) continue;
    const rec = slots[s] || (slots[s] = {});
    for (let c = tc + 1; c < hdr.length; c++) {
      const h = hdr[c]; if (!h) continue;
      const x = r[c], v = typeof x === "number" ? x : (x != null && x !== "" && !isNaN(+String(x).replace(",", ".")) ? +String(x).replace(",", ".") : null);
      if (v != null) rec[h] = v;
    }
  }
  return Object.keys(slots).length ? {slots, meta, titulo, fecha} : null;
}
const ZONAS = ["Demanda Centro", "Demanda Norte", "Demanda Sur"];
function classify(dayMaps, start) { // dayMaps: [{date, slots}] -> semana
  const times = labelsFor(start), n = 336, A = () => Array(n).fill(null);
  const centrales = {}, cmg = {"Santa Rosa": A(), "Trujillo": A(), "Socabaya": A()};
  const tot = {"Demanda YUPANA": A(), "Total CoGeneracion": A(), "Total Renovables": A()};
  const re = A(), rf = A(), flujo = {}, zona = {}; ZONAS.forEach(z => zona[z] = A());
  for (const {date, slots} of dayMaps) {
    const d = Math.round((new Date(date + "T12:00:00") - new Date(start + "T12:00:00")) / 864e5);
    if (d < 0 || d > 6) continue;
    for (const [s, rec] of Object.entries(slots)) {
      const i = d * 48 + (+s);
      for (const [h, v] of Object.entries(rec)) {
        if (/cmg/i.test(h)) {
          const b = /santa/i.test(h) ? "Santa Rosa" : /truji/i.test(h) ? "Trujillo" : /socab/i.test(h) ? "Socabaya" : null;
          if (b && (cmg[b][i] == null || /yupana/i.test(h))) cmg[b][i] = v;
        } else if (h === "Reserva Eficiente") re[i] = v;
        else if (h === "Reserva Fría Térmica") rf[i] = v;
        else if (/^FLUJO/i.test(h)) (flujo[h] || (flujo[h] = A()))[i] = v;
        else if (zona[h]) zona[h][i] = v;
        else if (tot[h]) tot[h][i] = v;
        else if (!EXCL.some(x => h.includes(x))) { (centrales[h] || (centrales[h] = Array(n).fill(0)))[i] = v; }
      }
    }
  }
  return {times, centrales, cmg, tot, re, rf, flujo, zona};
}
async function parseAnexo(file) {
  const wb = XLSX.read(await file.arrayBuffer(), {type: "array"});
  const days = []; let meta = {};
  for (const name of wb.SheetNames) {
    const r = readSheet(wb.Sheets[name]); if (!r) continue;
    const date = dateFromText(name) || r.fecha; if (!date) continue;
    days.push({date, slots: r.slots}); meta = {...meta, ...r.meta};
  }
  if (!days.length) throw new Error(`${file.name}: no encontré hojas con la tabla "Día Hora".`);
  days.sort((a, b) => a.date.localeCompare(b.date));
  const start = days[0].date, w = classify(days, start);
  const sem = (file.name.match(/Sem\s*(\d+)/i) || [])[1];
  return {semana: sem ? +sem : null, anio: +start.slice(0, 4), inicio: start, fin: addDays(start, 6), fuente: file.name, meta, ...w};
}
async function parseDaily(file) { // Reprog_DDMMX.xlsx o Anexo1_Despacho_AAAAMMDD.xlsx (PDO)
  const wb = XLSX.read(await file.arrayBuffer(), {type: "array"});
  let best = null;
  for (const n of wb.SheetNames) { const r = readSheet(wb.Sheets[n]); if (r && (!best || Object.keys(r.slots).length > Object.keys(best.slots).length)) best = {...r, sheet: n}; }
  if (!best) throw new Error(file.name + ': no encontré la tabla "Día Hora"');
  const date = best.fecha || dateFromText(best.sheet);
  const lt = (best.titulo.match(/REPROGRAMA\s+([A-Z]+)/i) || [])[1] || (file.name.match(/Reprog_?\s*\d{4}([A-Za-z]+)/i) || [])[1];
  const kind = /REPROGRAMA/i.test(best.titulo) || /reprog/i.test(file.name) ? "reprog" : "pdo";
  if (!date) throw new Error(file.name + ": no encontré la fecha");
  return {kind, date, letter: lt ? lt.toUpperCase() : null, slots: best.slots, meta: best.meta};
}
async function parseIndices(file) { // indicesAAAAMMDD_X.xlsx: HORA, alfa, beta, gamma, RDO, CHM
  const m = file.name.match(/indices\s*(\d{4})(\d{2})(\d{2})_?([A-Za-z]+)/i); if (!m) throw new Error(file.name + ": nombre esperado indicesAAAAMMDD_X.xlsx");
  const wb = XLSX.read(await file.arrayBuffer(), {type: "array"}), rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1, raw: true, defval: null});
  const hi = rows.findIndex(r => (r || []).some(v => String(v).trim().toLowerCase() === "alfa")); if (hi < 0) throw new Error(file.name + ": no encontré la columna alfa");
  const h = rows[hi].map(x => String(x ?? "").trim().toLowerCase()), col = k => h.indexOf(k), out = [];
  for (const r of rows.slice(hi + 1)) {
    const t = r[col("hora")]; const mins = typeof t === "number" ? Math.round((t % 1) * 1440) : timeMinutes(t); const s = slotOf(mins); if (s == null) continue;
    const chm = String(r[col("chm")] ?? ""); 
    out.push({s, alfa: +r[col("alfa")] || 0, beta: +r[col("beta")] || 0, gamma: +r[col("gamma")] || 0, rdo: +r[col("rdo")] || 0, noMon: (chm.match(/([A-ZÁÉÍÓÚÑ ]+?) - no monitoreada/g) || []).map(x => x.replace(/ - no monitoreada/, "").trim())});
  }
  return {date: `${m[1]}-${m[2]}-${m[3]}`, letter: m[4].toUpperCase(), rows: out};
}
async function parseListado(file) {
  const wb = XLSX.read(await file.arrayBuffer(), {type: "array"});
  const out = [];
  for (const n of wb.SheetNames) {
    for (const r of XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, raw: false, defval: null})) {
      const f = r[0] && dateFromText(r[0]); if (!f || !r[5]) continue;
      out.push({fecha: f, operador: r[2] || "", causa: r[3] || "", hora: r[4] || "", reprograma: String(r[5]).trim().toUpperCase(), publicacion: r[6] || ""});
    }
  }
  return out;
}
/* ============ Texto de informes ============ */
let pdfReady = null;
function loadPdfJs() {
  if (pdfReady) return pdfReady;
  pdfReady = (async () => {
    const lib = window.pdfjsLib; if (!lib) throw new Error("No cargó el lector de PDF");
    try {
      const src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      const blob = new Blob([`importScripts("${src}");`], {type: "application/javascript"});
      lib.GlobalWorkerOptions.workerPort = new Worker(URL.createObjectURL(blob));
    } catch (e) { lib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"; }
    return lib;
  })();
  return pdfReady;
}
async function docText(file) {
  const n = file.name.toLowerCase();
  if (n.endsWith(".txt")) return await file.text();
  if (n.endsWith(".docx")) { const r = await mammoth.extractRawText({arrayBuffer: await file.arrayBuffer()}); return r.value; }
  if (n.endsWith(".pdf")) {
    const lib = await loadPdfJs();
    const pdf = await lib.getDocument({data: new Uint8Array(await file.arrayBuffer())}).promise;
    let out = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const tc = await (await pdf.getPage(p)).getTextContent();
      let line = "", lastY = null, txt = [];
      for (const it of tc.items) { const y = Math.round(it.transform[5]); if (lastY != null && Math.abs(y - lastY) > 2) { txt.push(line); line = ""; } line += it.str; lastY = y; }
      txt.push(line);
      out.push(`[Página ${p}]\n` + txt.join("\n"));
    }
    return out.join("\n\n");
  }
  throw new Error("Formato no soportado: usa PDF, DOCX o TXT");
}
function excerpt(text, keywords, budget) {
  if (!text) return "";
  if (text.length <= budget) return text;
  const paras = text.split(/\n\s*\n|(?=\[Página \d+\])/).map(p => p.trim()).filter(Boolean);
  const kw = keywords.filter(Boolean).map(k => k.toLowerCase());
  const scored = paras.map((p, i) => { const l = p.toLowerCase(); let s = 0; for (const k of kw) if (l.includes(k)) s += k.length > 6 ? 3 : 1; return {i, p, s: s + (i < 4 ? 2 : 0)}; });
  const pick = new Set(); let used = 0;
  for (const x of [...scored].sort((a, b) => b.s - a.s)) { if (used + x.p.length > budget) continue; pick.add(x.i); used += x.p.length; }
  return paras.filter((_, i) => pick.has(i)).join("\n\n");
}
const GENERIC_KW = ["mantenimiento", "indisponib", "hidrolog", "caudal", "embalse", "gas", "camisea", "demanda", "costo marginal", "congesti", "restricci", "línea", "transmisión", "reserva", "despacho", "térmic", "renovable", "eólic", "solar", "precipitaci", "avenida", "estiaje", "falla", "intervenci", "programa", "limitaci"];

/* ============ Cálculos ============ */
function derive(w) {
  const n = w.times.length, bySrc = {}, byTipo = {};
  GROUPS.forEach(g => bySrc[g.k] = Array(n).fill(0));
  TIPOS.forEach(t => byTipo[t] = Array(n).fill(0));
  for (const [c, a] of Object.entries(w.centrales)) {
    const g = groupOf(c), t = tipoOf(c);
    for (let i = 0; i < n; i++) { bySrc[g][i] += a[i] || 0; if (t && byTipo[t]) byTipo[t][i] += a[i] || 0; }
  }
  const D = w.tot["Demanda YUPANA"], C = w.tot["Total CoGeneracion"], R = w.tot["Total Renovables"];
  const resid = w.times.map((_, i) => D[i] == null ? null : Math.max(0, D[i] - (C[i] || 0) - (R[i] || 0) - byTipo["HIDROELÉCTRICA"][i]));
  const gen = w.times.map((_, i) => GROUPS.reduce((s, g) => s + bySrc[g.k][i], 0));
  return {bySrc, byTipo, resid, gen};
}
const gwh = a => 0.5 * sum(a) / 1000;
const days = w => [...new Set(w.times.map(t => t.slice(0, 10)))];
function dayLabel(iso) { const d = new Date(iso + "T12:00:00"); return DOW[d.getDay()] + " " + d.getDate(); }
function tLabel(t) { return dayLabel(t.slice(0, 10)) + " " + t.slice(11); }
function stats(a, w) {
  let mx = -Infinity, mn = Infinity, imx = -1, imn = -1, s = 0, c = 0;
  a.forEach((v, i) => { if (v == null) return; s += v; c++; if (v > mx) { mx = v; imx = i; } if (v < mn) { mn = v; imn = i; } });
  return {max: mx, min: mn, avg: c ? s / c : null, tmax: imx >= 0 ? w.times[imx] : null, tmin: imn >= 0 ? w.times[imn] : null};
}
function dailyAvg(a) { const o = []; for (let d = 0; d < 7; d++) { const s = a.slice(d * 48, d * 48 + 48).filter(v => v != null); o.push(s.length ? sum(s) / s.length : null); } return o; }
function dailyGwh(a) { const o = []; for (let d = 0; d < 7; d++) o.push(gwh(a.slice(d * 48, d * 48 + 48))); return o; }

