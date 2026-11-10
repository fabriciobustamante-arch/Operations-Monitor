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

/* Caídas de despacho por central (candidatas a intervención) */
function detectDrops(w) {
  const out = [], dd = days(w);
  for (const [c, a] of Object.entries(w.centrales)) {
    const wmax = Math.max(...a); if (wmax < 15) continue;
    const dmax = dd.map((_, d) => Math.max(...a.slice(d * 48, d * 48 + 48)));
    const t = tipoOf(c);
    if (t === "SOLAR" || t === "EÓLICA") continue; // variables por recurso
    dd.forEach((day, d) => {
      const r = dmax[d] / wmax;
      if (r < 0.35) {
        const lost = sum(a.slice(d * 48, d * 48 + 48).map(v => wmax - v)) * 0.5;
        out.push({central: c, tipo: t || "Sin clasificar", empresa: MAESTRO.get(c)?.empresa || "", dia: day, wmax, dmax: dmax[d], tipoCaida: dmax[d] < 1 ? "Fuera de servicio" : "Reducción", mwh: lost});
      }
    });
  }
  const merged = {}; // agrupa días consecutivos por central
  for (const x of out) { (merged[x.central] || (merged[x.central] = [])).push(x); }
  const res = Object.values(merged).map(list => ({...list[0], dias: list.map(x => x.dia), dmaxs: list.map(x => x.dmax), mwh: sum(list.map(x => x.mwh)), tipoCaida: list.every(x => x.dmax < 1) ? "Fuera de servicio" : "Reducción"}));
  return res.sort((a, b) => b.mwh - a.mwh).slice(0, 20);
}
/* Reprogramas: programa vigente por día (última letra con dato en cada media hora) */
function vigente(date) {
  const R = ST.reprog[date]; if (!R) return null;
  const letters = Object.keys(R).sort((a, b) => a.length - b.length || a.localeCompare(b));
  const slots = {}, from = {};
  for (const L of letters) for (const [s, rec] of Object.entries(R[L])) { slots[s] = rec; from[s] = L; }
  return {letters, slots, from};
}
function rdoDay(date) {
  const w = ST.cur, V = vigente(date); if (!w || !V) return null;
  const d = days(w).indexOf(date); if (d < 0) return null;
  const o = d * 48, pso = {}, rdo = {}, cen = {}, diff = {};
  // Base: programa diario (PDO) si está cargado; si no, el PSO de ese día
  const P = ST.baseMode === "PSO" ? null : ST.pdo?.[date], B = P ? classify([{date, slots: P}], date) : null, bo = P ? 0 : o, src = B || w;
  const baseLabel = P ? "PDO" : "PSO";
  const sl = a => (a || Array(336).fill(null)).slice(bo, bo + 48);
  TIPOS.forEach(t => { pso[t] = Array(48).fill(0); rdo[t] = Array(48).fill(null); });
  const cm = {pso: {}, rdo: {}}; BARRAS.forEach(b => { cm.pso[b.k] = sl(src.cmg[b.k]); cm.rdo[b.k] = Array(48).fill(null); });
  const re = {pso: sl(src.re), rdo: Array(48).fill(null)};
  for (const [c, a] of Object.entries(src.centrales)) { const t = tipoOf(c); const seg = sl(a); cen[c] = {pso: seg, rdo: Array(48).fill(null)}; if (t && pso[t]) seg.forEach((v, i) => pso[t][i] += v || 0); }
  const T = k => sl(src.tot[k]);
  const resid = {pso: T("Demanda YUPANA").map((v, i) => v == null ? null : Math.max(0, v - (T("Total CoGeneracion")[i] || 0) - (T("Total Renovables")[i] || 0) - pso["HIDROELÉCTRICA"][i])), rdo: Array(48).fill(null)};
  for (let s = 0; s < 48; s++) {
    const rec = V.slots[s]; if (!rec) continue;
    TIPOS.forEach(t => rdo[t][s] = 0);
    const tot = {};
    for (const [h, v] of Object.entries(rec)) {
      if (/cmg/i.test(h)) { const b = /santa/i.test(h) ? "Santa Rosa" : /truji/i.test(h) ? "Trujillo" : /socab/i.test(h) ? "Socabaya" : null; if (b && (cm.rdo[b][s] == null || /yupana/i.test(h))) cm.rdo[b][s] = v; continue; }
      if (h === "Reserva Eficiente") { re.rdo[s] = v; continue; }
      if (["Demanda YUPANA", "Total CoGeneracion", "Total Renovables"].includes(h)) { tot[h] = v; continue; }
      if (EXCL.some(x => h.includes(x))) continue;
      const t = tipoOf(h); if (t && rdo[t]) rdo[t][s] += v;
      (cen[h] || (cen[h] = {pso: Array(48).fill(0), rdo: Array(48).fill(null)})).rdo[s] = v;
    }
    if (tot["Demanda YUPANA"] != null) resid.rdo[s] = Math.max(0, tot["Demanda YUPANA"] - (tot["Total CoGeneracion"] || 0) - (tot["Total Renovables"] || 0) - rdo["HIDROELÉCTRICA"][s]);
  }
  // celda vacía = 0 MW en las horas que cubre el reprograma
  for (const [c, x] of Object.entries(cen)) { let m = 0; for (let s = 0; s < 48; s++) if (V.slots[s]) { if (x.rdo[s] == null) x.rdo[s] = 0; m += (x.rdo[s] - (x.pso[s] || 0)) * 0.5; } diff[c] = {mwh: m}; }
  const firstSlot = Math.min(...Object.keys(V.slots).map(Number));
  const diffs = Object.entries(diff).map(([c, x]) => ({c, mwh: x.mwh})).filter(x => Math.abs(x.mwh) > 1).sort((a, b) => Math.abs(b.mwh) - Math.abs(a.mwh));
  return {date, d, baseLabel, pso, rdo, cm, re, resid, cen, diffs, letters: V.letters, from: V.from, firstSlot, causas: ST.listado.filter(x => x.fecha === date).sort((a, b) => a.reprograma.localeCompare(b.reprograma))};
}

/* ============ Gráficos ============ */

/* ============ Estado ============ */
const WEEKS = new Map(); // clave "2026-42" -> {w, ipso, ipsoName, ipsi, ipsiName, ai:{}}
const ST = {key: null, reprog: {}, pdo: {}, ipdo: {}, indices: {}, listado: [], tab: "pso", charts: {}, dropSel: null, rdoCur: null,
  get cur() { return WEEKS.get(this.key)?.w; }, get wk() { return WEEKS.get(this.key); }};
const keyOf = w => `${w.anio}-${String(w.semana).padStart(2, "0")}`;
function coesWeek(start) {
  const y = +start.slice(0, 4), j1 = new Date(y + "-01-01T12:00:00");
  const firstSat = new Date(j1); firstSat.setDate(1 + (6 - j1.getDay() + 7) % 7);
  return Math.floor((new Date(start + "T12:00:00") - firstSat) / 6048e5) + 1;
}
function addWeek(w, doc = {}) {
  if (!w.semana) w.semana = coesWeek(w.inicio);
  const k = keyOf(w), old = WEEKS.get(k);
  WEEKS.set(k, {w, ipso: old?.ipso || doc.ipso || "", ipsoName: old?.ipsoName || doc.ipsoName || "", ipsi: old?.ipsi || doc.ipsi || "", ipsiName: old?.ipsiName || doc.ipsiName || "", ai: {}});
  return k;
}
const sortedKeys = () => [...WEEKS.keys()].sort();
const prevWeek = () => { const c = ST.cur; if (!c) return null; const s = addDays(c.inicio, -7); for (const v of WEEKS.values()) if (v.w.inicio === s) return v; return null; };

const dayList = ds => { const r = []; for (const d of [...ds].sort()) { const l = r[r.length - 1], nx = l && new Date(new Date(l[1] + "T12:00").getTime() + 864e5).toISOString().slice(0, 10); if (l && nx === d) l[1] = d; else r.push([d, d]); }
  return r.map(([a, b]) => a === b ? dayLabel(a) : `${dayLabel(a)}–${dayLabel(b).split(" ")[1]}`).join(", "); };
const gfmt = x => fmt(x, x < 10 ? 1 : 0);
/* ============ Hallazgos (dato = corto para pantalla, det = detalle para la IA) ============ */
function psoFindings(w, dv) {
  const f = [], dd = days(w), tot = gwh(dv.gen);
  const sh = GROUPS.map(g => ({k: g.k, e: gwh(dv.bySrc[g.k])})).filter(x => x.e > 0.05).sort((a, b) => b.e - a.e);
  f.push({id: "P1", chart: "src", titulo: "Matriz", dato: sh.slice(0, 4).map(s => `${s.k} ${fmt(100 * s.e / tot)} %`).join(" · "),
    det: `Total ${fmt(tot, 1)} GWh. ` + sh.map(s => `${s.k} ${fmt(s.e, 1)} GWh (${fmt(100 * s.e / tot, 1)} %)`).join("; ")});
  const hd = dailyGwh(dv.bySrc["Hidro"]), gd = dailyGwh(dv.bySrc["Gas natural"]);
  const hmin = hd.indexOf(Math.min(...hd)), hmax = hd.indexOf(Math.max(...hd));
  f.push({id: "P2", chart: "src", titulo: "Hidro y gas por día", dato: `Hidro de ${fmt(hd[hmin], 1)} (${dayLabel(dd[hmin])}) a ${fmt(hd[hmax], 1)} GWh (${dayLabel(dd[hmax])})`,
    det: dd.map((d, i) => `${dayLabel(d)}: hidro ${fmt(hd[i], 1)}, gas ${fmt(gd[i], 1)} GWh`).join("; ")});
  const th = Object.entries(w.centrales).filter(([c]) => tipoOf(c) === "TERMOELÉCTRICA").map(([c, a]) => ({c, e: gwh(a), mx: Math.max(...a)})).sort((a, b) => b.e - a.e);
  f.push({id: "P3", chart: "src", titulo: "Mayores térmicas", dato: th.slice(0, 3).map(x => x.c).join(", "), det: th.slice(0, 8).map(x => `${x.c} ${fmt(x.e, 1)} GWh (máx ${fmt(x.mx)} MW)`).join("; ")});
  const um = Object.keys(w.centrales).filter(c => !MAESTRO.has(c));
  if (um.length) f.push({id: "P4", chart: "src", local: true, titulo: "Sin clasificar", dato: `${um.length} centrales, ${fmt(gwh(dv.bySrc["Sin clasificar"]), 1)} GWh`, det: um.join(", ")});
  BARRAS.forEach((b, j) => { const s = stats(w.cmg[b.k], w);
    f.push({id: "C" + (j + 1), chart: "cmg", titulo: `${b.z} (${b.k})`, dato: `media ${fmt(s.avg, 1)} · máx ${fmt(s.max, 1)} ${tLabel(s.tmax)}`,
      det: `Promedio ${fmt(s.avg, 2)} US$/MWh; máximo ${fmt(s.max, 2)} el ${tLabel(s.tmax)}; mínimo ${fmt(s.min, 2)} el ${tLabel(s.tmin)}. Diario: ` + dailyAvg(w.cmg[b.k]).map((v, i) => `${dayLabel(dd[i])} ${fmt(v, 1)}`).join(", ")}); });
  const sp = stats(w.times.map((_, i) => (w.cmg["Trujillo"][i] ?? 0) - (w.cmg["Santa Rosa"][i] ?? 0)), w);
  const sp2 = stats(w.times.map((_, i) => (w.cmg["Socabaya"][i] ?? 0) - (w.cmg["Santa Rosa"][i] ?? 0)), w);
  f.push({id: "C4", chart: "cmg", titulo: "Desacople entre zonas", dato: `Norte +${fmt(sp.max, 0)} · Sur +${fmt(sp2.max, 0)} US$/MWh sobre Centro`,
    det: `Norte − Centro: máximo ${fmt(sp.max, 2)} US$/MWh (${tLabel(sp.tmax)}), promedio ${fmt(sp.avg, 2)}. Sur − Centro: máximo ${fmt(sp2.max, 2)} (${tLabel(sp2.tmax)}), promedio ${fmt(sp2.avg, 2)}.`});
  const fCN = w.flujo?.["FLUJO ENLACE CENTRO - NORTE"], fCS = w.flujo?.["FLUJO ENLACE CENTRO - SUR"];
  if (fCN || fCS) { const tn = stats(w.cmg["Trujillo"], w), it = w.times.indexOf(tn.tmax), a = fCN ? stats(fCN, w) : null, b = fCS ? stats(fCS, w) : null;
    const extra = Object.entries(w.flujo).filter(([h]) => /500 KV/.test(h)).map(([h, arr]) => `${h.replace("FLUJO ", "")} ${fmt(arr[it])} MW`).join("; ");
    f.push({id: "C5", chart: "cmg", titulo: "Flujos por los enlaces", dato: `C–N máx ${a ? fmt(a.max) : "—"} MW · C–S máx ${b ? fmt(b.max) : "—"} MW`,
      det: `Enlace Centro–Norte: máx ${a ? fmt(a.max) : "—"} MW (${a ? tLabel(a.tmax) : ""}), en el pico de CMg Norte (${tLabel(tn.tmax)}) ${fCN ? fmt(fCN[it]) : "—"} MW. Enlace Centro–Sur: máx ${b ? fmt(b.max) : "—"} MW (${b ? tLabel(b.tmax) : ""}). Líneas de 500 kV en el pico Norte: ${extra}.`}); }
  const dem = stats(w.tot["Demanda YUPANA"], w), demd = dailyGwh(w.tot["Demanda YUPANA"]);
  f.push({id: "D1", chart: "dr", titulo: "Demanda", dato: `máx ${fmt(dem.max)} MW ${tLabel(dem.tmax)}`,
    det: `Máxima ${fmt(dem.max)} MW el ${tLabel(dem.tmax)}; mínima ${fmt(dem.min)} MW el ${tLabel(dem.tmin)}. Diaria: ` + dd.map((d, i) => `${dayLabel(d)} ${fmt(demd[i], 1)} GWh`).join(", ")});
  const rs = stats(dv.resid, w);
  f.push({id: "D2", chart: "dr", titulo: "Residual (cubre la térmica)", dato: `${fmt(rs.min)} a ${fmt(rs.max)} MW`, det: `Máxima ${fmt(rs.max)} MW el ${tLabel(rs.tmax)}; mínima ${fmt(rs.min)} MW el ${tLabel(rs.tmin)}; media ${fmt(rs.avg)} MW.`});
  if (w.re.some(v => v != null)) { const r = stats(w.re, w); f.push({id: "D3", chart: "dr", titulo: "Reserva eficiente", dato: `hasta ${fmt(r.max)} MW`, det: `En ${w.re.filter(v => v != null).length} medias horas; máxima ${fmt(r.max, 1)} MW el ${tLabel(r.tmax)}.`}); }
  return f;
}
function findingsFor(tab) {
  const w = ST.cur; if (!w) return [];
  if (tab === "pso") return [...cmgEvents(w), ...psoFindings(w, derive(w))];
  if (tab === "ipsi") return detectDrops(w).map((x, i) => ({id: "I" + (i + 1), chart: "drop", central: x.central, titulo: x.central, dato: `${x.tipoCaida} · ${dayList(x.dias)}`, auto: typeof dropJust === "function" ? dropJust(x) : null,
    det: `${x.empresa} · ${x.tipo}. ${x.tipoCaida} los días ${x.dias.join(", ")}. Máximo semanal ${fmt(x.wmax)} MW; máximo esos días ${x.dmaxs.map(v => fmt(v)).join("/")} MW; ${fmt(x.mwh)} MWh no despachados.`}));
  if (tab === "rdo") { const R = ST.rdoCur; if (!R) return []; const f = [];
    const de = TIPOS.map(t => ({t, d: sum(R.rdo[t].map((v, s) => v == null ? 0 : v - R.pso[t][s])) * 0.5})).filter(x => Math.abs(x.d) > 1);
    f.push({id: "R1", chart: "rdo", titulo: "Cambio por tipo", dato: de.map(x => `${x.t[0] + x.t.slice(1, 5).toLowerCase()}. ${sgn(x.d)} MWh`).join(" · ") || "sin cambios", det: de.map(x => `${x.t} ${sgn(x.d)} MWh`).join("; ") || "Sin cambios"});
    f.push({id: "R2", chart: "rdo", titulo: "Centrales que más cambian", dato: R.diffs.slice(0, 3).map(x => `${x.c} ${sgn(x.mwh)}`).join(" · ") || "—", det: R.diffs.slice(0, 10).map(x => `${x.c} ${sgn(x.mwh)} MWh`).join("; ")});
    const lab = s => { const m = (s + 1) * 30; return m === 1440 ? "23:59" : String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0"); };
    R.letters.forEach(L => { const sl = Object.keys(ST.reprog[R.date][L]).map(Number), ini = Math.min(...sl), ix = ST.indices[R.date]?.[L];
      let dat = `desde ${lab(ini)}`, det = `Reprograma ${L} rige desde ${lab(ini)}.`;
      if (ix?.length) { const act = ix.filter(x => x.rdo === 1), mx = k => Math.max(...ix.map(x => x[k])), nm = [...new Set(ix.flatMap(x => x.noMon))];
        dat += ` · RDO=1 en ${act.length} medias horas · alfa máx ${fmt(mx("alfa"), 2)}`;
        det += ` Índices previos: RDO=1 en ${act.length} medias horas (${act.length ? lab(act[0].s) + "–" + lab(act[act.length - 1].s) : "—"}); alfa máx ${fmt(mx("alfa"), 2)}, beta máx ${fmt(mx("beta"), 2)}, gamma máx ${fmt(mx("gamma"), 0)}. Centrales hidro no monitoreadas: ${nm.join(", ") || "ninguna"}.`; }
      f.push({id: "L" + L, chart: "rdo", titulo: `Reprograma ${L}`, dato: dat, det}); });
    R.causas.forEach((c, i) => f.push({id: "K" + (i + 1), chart: "rdo", titulo: `Reprog. ${c.reprograma} · ${c.hora}`, dato: c.causa, det: `Causa publicada: ${c.causa} (${c.operador})`, causa: true}));
    if (typeof rdoCmgEvents === "function") f.push(...rdoCmgEvents().filter(e => e.t.slice(0, 10) === R.date));
    BARRAS.forEach((b, j) => { const cp = stats(R.cm.pso[b.k].map((v, i) => R.cm.rdo[b.k][i] == null ? null : v), {times: []}), cr = stats(R.cm.rdo[b.k], {times: []}); if (cr.avg == null) return;
      f.push({id: "M" + (j + 1), chart: "rdocmg", titulo: `CMg ${b.z}`, dato: `${R.baseLabel} ${fmt(cp.avg, 1)} → ${fmt(cr.avg, 1)} (máx ${fmt(cr.max, 0)})`, det: `${b.k}: ${R.baseLabel} medio ${fmt(cp.avg, 2)}, vigente medio ${fmt(cr.avg, 2)} US$/MWh, máx vigente ${fmt(cr.max, 2)}`}); });
    const rp = stats(R.resid.pso.map((v, i) => R.resid.rdo[i] == null ? null : v), {times: []}), rr = stats(R.resid.rdo, {times: []});
    if (rr.avg != null) f.push({id: "R4", chart: "rdores", titulo: "Demanda residual", dato: `media ${fmt(rp.avg)} → ${fmt(rr.avg)} MW`, det: `${R.baseLabel} medio ${fmt(rp.avg)} MW, vigente medio ${fmt(rr.avg)} MW, máx vigente ${fmt(rr.max)} MW`});
    const ep = sum(R.re.pso.map((v, i) => R.from[i] ? v : 0)) * 0.5, er = sum(R.re.rdo) * 0.5;
    if (ep || er) f.push({id: "R5", chart: "rdores", titulo: "Reserva eficiente", dato: `${fmt(ep)} → ${fmt(er)} MWh`, det: `Reserva eficiente ${R.baseLabel} ${fmt(ep)} MWh, vigente ${fmt(er)} MWh`});
    return f; }
  if (tab === "cmp") { const C = cmpData(); if (!C) return []; const f = [];
    const big = [...C.src].sort((a, b) => Math.abs(b.b - b.a) - Math.abs(a.b - a.a));
    f.push({id: "S1", chart: "cmpsrc", titulo: "Matriz", dato: big.slice(0, 2).map(x => `${x.k} ${sgn(x.b - x.a, 1)} GWh`).join(" · "), det: C.src.map(x => `${x.k} ${fmt(x.a, 1)} → ${fmt(x.b, 1)} GWh`).join("; ")});
    f.push({id: "S2", chart: "cmpsrc", titulo: "Demanda", dato: `${sgn(100 * (C.demB / C.demA - 1), 1)} %`, det: `${fmt(C.demA, 1)} → ${fmt(C.demB, 1)} GWh`});
    f.push({id: "S3", chart: "cmpcen", titulo: "CMg medio", dato: C.cm.map(x => `${x.z} ${sgn(x.b - x.a, 1)}`).join(" · "), det: C.cm.map(x => `${x.z} ${fmt(x.a, 2)} → ${fmt(x.b, 2)} US$/MWh`).join("; ")});
    f.push({id: "S4", chart: "cmpcen", titulo: "Centrales que más cambian", dato: C.cen.slice(0, 3).map(x => `${x.c} ${sgn(x.d, 1)}`).join(" · "), det: C.cen.slice(0, 10).map(x => `${x.c} ${sgn(x.d, 2)} GWh${x.nueva ? " (nueva)" : x.sale ? " (sale)" : ""}`).join("; ")});
    return f; }
  return [];
}

function renderIPSI() {
  const w = ST.cur, drops = detectDrops(w), ai = aiMap("ipsi");
  if (!drops.length) { $("tblDrops").innerHTML = `<p class="muted">Sin caídas relevantes.</p>`; $("why-drop").innerHTML = ""; return; }
  if (!drops.some(x => x.central === ST.dropSel)) ST.dropSel = drops[0].central;
  const F = findingsFor("ipsi");
  $("tblDrops").innerHTML = `<table><thead><tr><th>Central</th><th>Días</th><th>Estado</th><th class="n">MWh</th><th>IPSI</th></tr></thead><tbody>${drops.map((x, i) => { const h = ai.get(F[i].id), c = (h?.confianza || "").toLowerCase();
    return `<tr data-central="${esc(x.central)}" class="${x.central === ST.dropSel ? "sel" : ""}"><td><b>${esc(x.central)}</b></td><td style="white-space:nowrap">${dayList(x.dias)}</td><td>${x.tipoCaida}</td><td class="n">${fmt(x.mwh)}</td><td>${h ? `<span class="pill ${c}">${esc(c === "ninguna" ? "sin sustento" : c)}</span>` : "—"}</td></tr>`; }).join("")}</tbody></table>`;
  const sel = drops.find(x => x.central === ST.dropSel);
  $("dropTitle").textContent = sel.central;
  mkChart("cDrop", {type: "line", data: {labels: w.times, datasets: [{label: sel.central, data: w.centrales[sel.central], borderColor: css("--accent"), backgroundColor: css("--accent-soft"), fill: true}]}, options: baseOpts(w, "")});
  renderWhy("why-drop", "ipsi", F.filter(f => f.central === sel.central), `${esc(sel.empresa)} · ${esc(sel.tipo)}`);
}
function cmpData() {
  const P = prevWeek(); if (!P) return null;
  const a = P.w, b = ST.cur, da = derive(a), db = derive(b);
  const src = GROUPS.map(g => ({k: g.k, a: gwh(da.bySrc[g.k]), b: gwh(db.bySrc[g.k])})).filter(x => x.a + x.b > 0.05);
  const names = new Set([...Object.keys(a.centrales), ...Object.keys(b.centrales)]);
  const cen = [...names].map(c => ({c, a: a.centrales[c] ? gwh(a.centrales[c]) : 0, b: b.centrales[c] ? gwh(b.centrales[c]) : 0, nueva: !a.centrales[c], sale: !b.centrales[c]})).map(x => ({...x, d: x.b - x.a})).sort((x, y) => Math.abs(y.d) - Math.abs(x.d));
  const cm = BARRAS.map(br => ({k: br.k, z: br.z, a: stats(a.cmg[br.k], a).avg, b: stats(b.cmg[br.k], b).avg}));
  return {a, b, da, db, src, cen, cm, demA: gwh(a.tot["Demanda YUPANA"]), demB: gwh(b.tot["Demanda YUPANA"])};
}
function renderCMP() {
  const C = cmpData(); $("cmpEmpty").hidden = !!C; $("cmpBody").hidden = !C;
  if (!C) { $("cmpEmptyTxt").textContent = `Agrega el Anexo1 de la semana ${ST.cur.semana - 1}.`; return; }
  const ta = gwh(C.da.gen), tb = gwh(C.db.gen);
  $("cmpKpis").innerHTML = kpiHtml([["Energía", sgn(tb - ta, 1), "GWh"], ["Demanda", sgn(100 * (C.demB / C.demA - 1), 1), "%"], ...C.cm.map(x => ["CMg " + x.z, sgn(x.b - x.a, 1), "US$/MWh"])]);
  const fg2 = css("--fg-2"), line = css("--line");
  mkChart("cCmp", {type: "bar", data: {labels: C.src.map(x => x.k), datasets: [
    {label: "Sem " + C.a.semana, data: C.src.map(x => x.a), backgroundColor: css("--fg-3"), borderRadius: 4},
    {label: "Sem " + C.b.semana, data: C.src.map(x => x.b), backgroundColor: css("--accent"), borderRadius: 4}]},
    options: {responsive: true, maintainAspectRatio: false, animation: false, indexAxis: "y", plugins: {legend: {display: false}, tooltip: {callbacks: {label: c => ` ${c.dataset.label}: ${fmt(c.parsed.x, 1)} GWh`}}},
      scales: {x: {grid: {color: line}, ticks: {color: fg2}}, y: {grid: {display: false}, ticks: {color: fg2}}}}});
  legend("lgCmp", [{label: "Semana " + C.a.semana, color: css("--fg-3")}, {label: "Semana " + C.b.semana, color: css("--accent")}]);
  mkChart("cCmpCmg", {type: "line", data: {labels: days(C.b).map(dayLabel), datasets: [
    {label: "Sem " + C.a.semana, data: dailyAvg(C.a.cmg["Santa Rosa"]), borderColor: css("--fg-3"), borderDash: [5, 4], pointRadius: 4},
    {label: "Sem " + C.b.semana, data: dailyAvg(C.b.cmg["Santa Rosa"]), borderColor: css("--b-centro"), pointRadius: 4}]},
    options: {responsive: true, maintainAspectRatio: false, animation: false, interaction: {mode: "index", intersect: false}, plugins: {legend: {display: false}}, scales: {x: {grid: {display: false}, ticks: {color: fg2}}, y: {grid: {color: line}, ticks: {color: fg2}}}}});
  legend("lgCmpCmg", [{label: "Semana " + C.a.semana, color: css("--fg-3"), line: true, dash: true}, {label: "Semana " + C.b.semana, color: css("--b-centro"), line: true}]);
  $("tblCmp").innerHTML = `<table><thead><tr><th>Central</th><th class="n">GWh</th><th></th></tr></thead><tbody>${C.cen.slice(0, 6).map(x => `<tr><td>${esc(x.c)}</td><td class="n ${x.d > 0 ? "pos" : "neg"}">${sgn(x.d, 1)}</td><td class="faint">${x.nueva ? "nueva" : x.sale ? "sale" : ""}</td></tr>`).join("")}</tbody></table>`;
  const F = findingsFor("cmp");
  renderWhy("why-cmpsrc", "cmp", F.filter(f => f.chart === "cmpsrc")); renderWhy("why-cmpcen", "cmp", F.filter(f => f.chart === "cmpcen"));
}

/* ============ Panel "Por qué" + IA ============ */
let SAMPLE = null;
const NEEDS = {pso: ["ipso"], ipsi: ["ipsi"], rdo: [], cmp: ["ipso"]};
const aiKey = tab => tab === "rdo" ? "rdo:" + (ST.rdoCur?.date || "") : tab;
const aiState = tab => (ST.wk?.ai || {})[aiKey(tab)] || {};
const aiMap = tab => new Map((aiState(tab).result?.hallazgos || []).map(h => [h.id, h]));
function renderWhy(boxId, tab, list, sub = "") {
  const ai = aiMap(tab), st = aiState(tab), el = $(boxId); if (!el) return;
  el.innerHTML = `<div class="lbl">Justificación</div>${sub ? `<div class="unit">${sub}</div>` : ""}` +
    (st.loading ? `<div class="thinking">Leyendo informes…</div>` : "") +
    list.map(f => { const h = ai.get(f.id) || (f.auto ? {explicacion: f.auto.texto, cita: f.auto.cita || "", fuente: f.auto.fuente, confianza: f.auto.conf, regla: true} : null), c = (h?.confianza || "").toLowerCase();
      return `<div class="item ${c}"><div class="s"></div><div><div class="t">${esc(f.titulo)}</div><div class="k">${esc(f.dato)}</div>${h ? `<p class="x">${esc(h.explicacion)}</p>${h.cita ? `<q>“${esc(h.cita)}”</q>` : ""}<div class="src2"><span class="pill ${c}">${esc(c === "ninguna" ? "sin sustento" : c)}</span>${esc(h.fuente || "")}${h.regla ? ' · <span class="faint">cruce automático</span>' : ai.get(f.id) ? ' · <span class="faint">IA</span>' : ""}</div>` : ""}</div></div>`; }).join("");
}
function renderAIBar(tab) {
  const el = $("aiBar-" + tab); if (!el) return;
  const st = aiState(tab), missing = NEEDS[tab].filter(n => !ST.wk?.[n]);
  const ready = SAMPLE && findingsFor(tab).length;
  let msg = !SAMPLE ? "La IA funciona al abrir la página dentro de Claude." : missing.length ? `Agrega el ${missing.join(" y ").toUpperCase()} de la semana ${ST.cur.semana} para justificar.` : st.result ? "" : "";
  if (st.error) msg = st.error;
  el.innerHTML = `<button class="primary" type="button" data-run="${tab}" ${ready && !st.loading ? "" : "disabled"}>${st.loading ? "Justificando…" : st.result ? "Volver a justificar" : "Justificar con IA"}</button>${st.loading ? `<button type="button" data-stop="${tab}">Detener</button>` : ""}<span class="unit ${st.error ? "neg" : ""}">${esc(msg)}</span>`;
}
function contextDocs(tab, findings, budget = 70000) {
  const kw = [...GENERIC_KW, ...findings.flatMap(f => (f.titulo + " " + f.det).split(/[\s,;:()]+/).filter(x => x.length > 4 && /[A-ZÁÉÍÓÚÑ]{4,}/.test(x)))];
  const W = ST.wk, parts = [];
  if (W?.ipso) parts.push(`=== IPSO (${W.ipsoName}) ===\n` + excerpt(W.ipso, kw, tab === "ipsi" ? budget * 0.3 : budget * 0.7));
  if (W?.ipsi) parts.push(`=== IPSI (${W.ipsiName}) ===\n` + excerpt(W.ipsi, kw, tab === "ipsi" ? budget * 0.7 : budget * 0.3));
  if (tab === "rdo" && ST.ipdo[ST.rdoCur?.date]) parts.unshift(`=== IPDO del ${ST.rdoCur.date} (${ST.ipdo[ST.rdoCur.date].name}) ===\n` + excerpt(ST.ipdo[ST.rdoCur.date].text, kw, budget * 0.5));
  if (tab === "rdo" && ST.rdoCur?.causas.length) parts.push("=== Listado de reprogramas ===\n" + ST.rdoCur.causas.map(c => `${c.reprograma} | ${c.hora} | ${c.operador} | ${c.causa}`).join("\n"));
  return parts.join("\n\n");
}
function prompt(tab, findings) {
  const w = ST.cur;
  return `Eres analista de operación del SEIN (Perú). Explicas el Programa Semanal de Operación del COES a un ingeniero de una generadora.
Semana ${w.semana} (${w.inicio} a ${w.fin}). Los HALLAZGOS vienen del Anexo1 de despacho: sus cifras son correctas, no las cambies.
Para cada hallazgo explica en 1-2 frases cortas por qué ocurre, usando SOLO los documentos. Si no lo explican, dilo y usa confianza "ninguna". Una razón técnica general que no esté en el documento va con "baja". Incluye una cita literal breve (máx. 20 palabras) cuando exista, y la página si aparece.

HALLAZGOS:
${findings.map(f => `[${f.id}] ${f.titulo}: ${f.det}`).join("\n")}

DOCUMENTOS:
${contextDocs(tab, findings) || "(ninguno)"}

Responde SOLO JSON: {"hallazgos":[{"id":"P1","explicacion":"...","cita":"","fuente":"IPSO p. 4","confianza":"alta|media|baja|ninguna"}]}`;
}
async function runAI(tab) {
  const findings = findingsFor(tab).filter(f => !f.local && !f.causa);
  const W = ST.wk, k = aiKey(tab), st = W.ai[k] = {loading: true, ctl: new AbortController()};
  renderTab();
  try { const res = await SAMPLE.json(prompt(tab, findings), {signal: st.ctl.signal, cache: false}); st.result = res?.hallazgos ? res : null; if (!st.result) st.error = "Respuesta sin formato. Intenta de nuevo."; }
  catch (e) { st.error = {cancelled: "", not_granted: "Sin permiso para usar Claude.", rate_limited: "Demasiadas consultas. Espera un momento.", prompt_too_large: "Informes muy largos. Pega solo las secciones útiles.", invalid_json: "Respuesta sin formato. Intenta de nuevo."}[e.code] ?? `No se pudo completar (${e.code || "error"}).`;
    if (e.code === "not_granted" || e.code === "sampling_disabled") SAMPLE = null; }
  st.loading = false; if (ST.wk === W) renderTab();
}
/* Chat */
const CHAT = [];
function renderChat() {
  $("chatLog").innerHTML = CHAT.map(m => `<div class="msg ${m.role === "user" ? "u" : "a"}">${esc(m.content)}</div>`).join("");
  $("chatSend").disabled = !SAMPLE;
  $("chatNote").textContent = !SAMPLE ? "El chat funciona dentro de Claude." : (ST.wk?.ipso || ST.wk?.ipsi) ? "" : "Sin IPSO ni IPSI: solo responde con las cifras.";
}
async function sendChat(q) {
  const w = ST.cur, f = psoFindings(w, derive(w));
  const ctx = `Eres analista de operación del SEIN (COES). Responde en español, breve, citando IPSO, IPSI o listado cuando los uses. Si los documentos no lo dicen, dilo.
Semana ${w.semana} (${w.inicio} a ${w.fin}). Cifras:\n${f.map(x => `- ${x.titulo}: ${x.det}`).join("\n")}
Caídas: ${detectDrops(w).slice(0, 10).map(x => `${x.central} ${x.dias.join("/")} (${x.tipoCaida})`).join("; ")}

DOCUMENTOS:\n${contextDocs("chat", f, 60000) || "(ninguno)"}`;
  CHAT.push({role: "user", content: q}, {role: "assistant", content: "…"}); renderChat();
  const bubble = $("chatLog").lastElementChild;
  try { const {text} = await SAMPLE([{role: "user", content: ctx}, {role: "assistant", content: "Listo, ¿cuál es tu pregunta?"}, ...CHAT.slice(0, -1).slice(-8)], {cache: false, onText: ({text}) => { bubble.textContent = text; }}); CHAT[CHAT.length - 1].content = text; }
  catch (e) { CHAT[CHAT.length - 1].content = e.text || `No se pudo responder (${e.code || "error"}).`; }
  renderChat();
}

function persist() {
  try {
    const weeks = [...WEEKS.entries()].map(([k, v]) => ({k, w: v.w.embedded ? null : v.w, ipso: v.ipso, ipsoName: v.ipsoName, ipsi: v.ipsi, ipsiName: v.ipsiName}));
    localStorage.setItem("pso-state", JSON.stringify({weeks, reprog: ST.reprog, pdo: ST.pdo, ipdo: ST.ipdo, indices: ST.indices, listado: ST.listado}));
  } catch (e) { /* sin espacio: se pierde al recargar */ }
}
const MESES = {enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12};
function dateFromSpanish(t) { const m = String(t).toLowerCase().match(/(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de(?:l)?\s+(\d{4})/); return m && MESES[m[2]] ? `${m[3]}-${String(MESES[m[2]]).padStart(2, "0")}-${m[1].padStart(2, "0")}` : null; }
async function handleFiles(files, forceWeek) {
  const msgs = [], errs = [];
  for (const f of files) {
    const n = f.name.toLowerCase();
    try {
      const pm = typeof handlePostFile === "function" ? await handlePostFile(f) : null; if (pm) { msgs.push(pm); continue; }
      if (/\.(pdf|docx|txt)$/.test(n)) {
        const t = await docText(f); if (t.replace(/\s/g, "").length < 200) throw new Error(`${f.name}: casi sin texto (¿escaneado?). Usa "Pegar texto".`);
        if (/ipdo|programa diario/.test(n) || /PROGRAMA DIARIO DE OPERACI/i.test(t.slice(0, 600))) {
          const d = dateFromSpanish(t.slice(0, 1500)); if (!d) throw new Error(`${f.name}: no encontré la fecha del IPDO`);
          ST.ipdo[d] = {text: t, name: f.name}; msgs.push(`IPDO del ${dayLabel(d)}`);
        } else {
          const kind = /ipsi|interv/.test(n) || /INTERVENCIONES/i.test(t.slice(0, 600)) ? "ipsi" : "ipso";
          ST.wk[kind] = t; ST.wk[kind + "Name"] = f.name; ST.wk.ai = {}; msgs.push(`${kind.toUpperCase()} → semana ${ST.cur.semana}`);
        }
      } else if (/^indices/.test(n)) {
        const r = await parseIndices(f); (ST.indices[r.date] || (ST.indices[r.date] = {}))[r.letter] = r.rows; msgs.push(`Índices ${r.letter} del ${dayLabel(r.date)}`);
      } else if (/anexo[2-8]|rs_reprog|cvar|costo|hidrolog|combust|restric|cogenera|vincula|delega/.test(n)) {
        msgs.push(`${f.name}: aún no se usa`);
      } else if (/reprog|despacho_\d{8}/.test(n)) {
        const r = await parseDaily(f);
        if (r.kind === "reprog") { (ST.reprog[r.date] || (ST.reprog[r.date] = {}))[r.letter || "A"] = r.slots; msgs.push(`Reprog ${r.letter || "A"} del ${dayLabel(r.date)}`); }
        else { ST.pdo[r.date] = r.slots; msgs.push(`PDO del ${dayLabel(r.date)}`); }
      } else if (forceWeek || /anexo|despacho|sem/.test(n)) {
        const w = await parseAnexo(f), k = addWeek(w); msgs.push(`Semana ${w.semana} cargada`); if (forceWeek === "week" || (forceWeek !== "prev" && k > ST.key)) ST.key = k;
      } else {
        const rows = await parseListado(f);
        if (rows.length) { const ks = new Set(ST.listado.map(x => x.fecha + x.reprograma)); ST.listado.push(...rows.filter(x => !ks.has(x.fecha + x.reprograma))); msgs.push(`Listado: ${rows.length} reprogramas`); }
        else { const w = await parseAnexo(f), k = addWeek(w); if (k > ST.key) ST.key = k; msgs.push(`Semana ${w.semana} cargada`); }
      }
    } catch (e) { errs.push(e.message || String(e)); }
  }
  persist(); renderAll();
  toast([...msgs, ...errs].join(" · "), errs.length && !msgs.length);
}

/* ============ Arranque ============ */
// Datos precargados (semana 42 real) y luego lo que este navegador haya guardado
for (const x of BOOT.weeks) { x.w.embedded = true; addWeek(x.w, x); }
ST.key = sortedKeys()[sortedKeys().length - 1];
Object.assign(ST.reprog, BOOT.reprog || {}); Object.assign(ST.pdo, BOOT.pdo || {}); Object.assign(ST.ipdo, BOOT.ipdo || {}); Object.assign(ST.indices, BOOT.indices || {}); ST.listado.push(...(BOOT.listado || [])); ST.limites = BOOT.limites || {};
try {
  const s = JSON.parse(localStorage.getItem("pso-state") || "null");
  if (s) { for (const x of s.weeks || []) { if (x.w) addWeek(x.w, x); else if (WEEKS.has(x.k)) { const W = WEEKS.get(x.k); for (const f of ["ipso", "ipsoName", "ipsi", "ipsiName"]) if (x[f]) W[f] = x[f]; } }
    for (const [d, L] of Object.entries(s.reprog || {})) ST.reprog[d] = {...(ST.reprog[d] || {}), ...L};
    Object.assign(ST.pdo, s.pdo || {}); Object.assign(ST.ipdo, s.ipdo || {}); for (const [d, L] of Object.entries(s.indices || {})) ST.indices[d] = {...(ST.indices[d] || {}), ...L};
    const ks = new Set(ST.listado.map(x => x.fecha + x.reprograma)); ST.listado.push(...(s.listado || []).filter(x => !ks.has(x.fecha + x.reprograma))); }
  const k = localStorage.getItem("pso-week"); if (k && WEEKS.has(k)) ST.key = k;
} catch (e) {}

/* ============ Filtros cruzados ============ */
ST.F = {days: new Set(), src: new Set(), central: ""}; ST.srcClosed = new Set();
ST.baseMode = "PDO"; ST.moment = null;
const TITLES = {pso: "Weekly Dispatch Plan", rdo: "Daily Redispatch Plan", th: "Thermal Redispatch", ren: "Renewable Redispatch", ipsi: "Intervenciones", cmp: "Vs semana anterior", chat: "Preguntar"};
const THERM_SRC = ["5. GAS", "6. GAS/DIESEL", "7. COAL", "8. DIESEL", "9. DIESEL/R500", "91. R500", "1. BIOMASS"];
const isStk = c => (MAESTRO.get(c)?.empresa || "").toUpperCase().includes("STATKRAFT");
const empOf = c => MAESTRO.get(c)?.empresa || ST.cur?.meta?.[c]?.empresa || "";
function pass(c) {
  const F = ST.F, m = MAESTRO.get(c);
  if (F.src.size && !F.src.has(m?.source || "")) return false;
  if (F.central) { if (F.central.startsWith("Empresa: ")) { if (empOf(c) !== F.central.slice(9)) return false; } else if (c !== F.central) return false; }
  return true;
}
const isFiltered = () => !!(ST.F.src.size || ST.F.central);
const idxs = w => w.times.map((_, i) => i).filter(i => !ST.F.days.size || ST.F.days.has(w.times[i].slice(0, 10)));
const pick = (a, I) => I.map(i => a ? a[i] : null);
function deriveF(w) { // como derive(), con centrales filtradas; la residual es del sistema (sin filtro)
  const n = w.times.length, bySrc = {}, byTipo = {}, all = derive(w);
  GROUPS.forEach(g => bySrc[g.k] = Array(n).fill(0)); TIPOS.forEach(t => byTipo[t] = Array(n).fill(0));
  for (const [c, a] of Object.entries(w.centrales)) { if (!pass(c)) continue; const g = groupOf(c), t = tipoOf(c); for (let i = 0; i < n; i++) { bySrc[g][i] += a[i] || 0; if (t && byTipo[t]) byTipo[t][i] += a[i] || 0; } }
  return {bySrc, byTipo, resid: all.resid, termAll: all.byTipo["TERMOELÉCTRICA"]};
}

/* ============ Reprogramas de toda la semana (vigente vs base) ============ */
function vigWeek(w) {
  const n = w.times.length, A = () => Array(n).fill(null), o = {cmg: {}, cmgB: {}, cen: {}, cenB: {}, re: A(), reB: [...w.re], resid: A(), residB: [...derive(w).resid], letter: Array(n).fill(null), days: [], marks: [], baseOf: Array(n).fill("PSO")};
  BARRAS.forEach(b => { o.cmg[b.k] = A(); o.cmgB[b.k] = [...w.cmg[b.k]]; });
  for (const [c, a] of Object.entries(w.centrales)) { o.cen[c] = A(); o.cenB[c] = [...a]; }
  days(w).forEach((date, d) => {
    if (!ST.reprog[date]) return; const R = rdoDay(date); if (!R) return; o.days.push({date, R});
    R.letters.forEach(L => { const sl = Math.min(...Object.keys(ST.reprog[date][L]).map(Number)); o.marks.push({t: w.times[d * 48 + sl], text: L}); });
    for (let s = 0; s < 48; s++) {
      const i = d * 48 + s, cov = R.from?.[s] != null; o.baseOf[i] = R.baseLabel;
      BARRAS.forEach(b => { o.cmgB[b.k][i] = R.cm.pso[b.k][s]; if (cov) o.cmg[b.k][i] = R.cm.rdo[b.k][s]; });
      o.reB[i] = R.re.pso[s]; o.residB[i] = R.resid.pso[s];
      if (cov) { o.letter[i] = R.from[s]; o.re[i] = R.re.rdo[s]; o.resid[i] = R.resid.rdo[s]; }
      for (const [c, x] of Object.entries(R.cen)) { (o.cenB[c] || (o.cenB[c] = Array(n).fill(0)))[i] = x.pso[s] || 0; if (!o.cen[c]) o.cen[c] = A(); if (cov) o.cen[c][i] = x.rdo[s] ?? 0; }
      for (const c of Object.keys(o.cenB)) if (!R.cen[c]) { if (R.baseLabel === "PDO") o.cenB[c][i] = 0; if (cov) { if (!o.cen[c]) o.cen[c] = A(); o.cen[c][i] = 0; } }
    }
  });
  return o;
}
let VW = null, VWsig = "";
function getVW() { const sig = ST.key + ST.baseMode + JSON.stringify(Object.keys(ST.reprog).map(d => d + Object.keys(ST.reprog[d]).join(""))) + Object.keys(ST.pdo).join(); if (sig !== VWsig) { VW = vigWeek(ST.cur); VWsig = sig; } return VW; }
const sumCen = (map, I, f) => I.map(i => { let s = 0, any = false; for (const [c, a] of Object.entries(map)) { if (!f(c)) continue; const v = a?.[i]; if (v != null) { s += v; any = true; } } return any ? s : null; });
function rdoFocusDay() { const V = getVW(); const sel = [...ST.F.days].filter(d => V.days.some(x => x.date === d)); return sel.length ? sel.sort().pop() : V.days.length ? V.days[V.days.length - 1].date : null; }

/* ============ Gráficos: opciones, marcas e inspector ============ */
const plugMoment = {id: "moment", afterDatasetsDraw(ch) { try { const t = ST.moment, L = ch.data.labels; if (!t || !L) return; const i = L.indexOf(t); if (i < 0) return;
  const x = ch.scales.x.getPixelForValue(i), {top, bottom} = ch.chartArea, c = ch.ctx; c.save(); c.strokeStyle = css("--fg"); c.lineWidth = 1.5; c.beginPath(); c.moveTo(x, top); c.lineTo(x, bottom); c.stroke(); c.restore(); } catch (e) {} }};
const plugMarks = {id: "marks", afterDatasetsDraw(ch, a, opts) { try { const M = opts?.items || [], L = ch.data.labels; if (!M.length || !L) return; const {top, bottom} = ch.chartArea, c = ch.ctx;
  c.save(); c.setLineDash([3, 3]); c.strokeStyle = css("--warn"); c.fillStyle = css("--warn"); c.font = "600 11px " + (css("--f-mono").split(",")[0] || "monospace"); c.textAlign = "center";
  for (const m of M) { const i = L.indexOf(m.t); if (i < 0) continue; const x = ch.scales.x.getPixelForValue(i); c.beginPath(); c.moveTo(x, top + 12); c.lineTo(x, bottom); c.stroke(); c.fillText(m.text, x, top + 9); }
  c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugMoment, plugMarks); } catch (e) {} }
function tsOpts(labels, {stacked = false, marks = [], y = ""} = {}) {
  const fg2 = css("--fg-2"), line = css("--line"), nd = new Set(labels.map(t => t.slice(0, 10))).size;
  return {responsive: true, maintainAspectRatio: false, animation: false, interaction: {mode: "index", intersect: false},
    onClick: (e, els, ch) => { try { const p = ch.getElementsAtEventForMode(e, "index", {intersect: false}, false); if (p.length) setMoment(ch.data.labels[p[0].index]); } catch (x) {} },
    plugins: {legend: {display: false}, marks: {items: marks}, tooltip: {filter: it => it.parsed.y != null, backgroundColor: css("--panel"), titleColor: css("--fg"), bodyColor: css("--fg-2"), borderColor: line, borderWidth: 1, padding: 10, boxPadding: 4,
      callbacks: {title: it => it.length ? tLabel(labels[it[0].dataIndex]) + (getVW().letter[ST.cur.times.indexOf(labels[it[0].dataIndex])] ? ` · Reprog ${getVW().letter[ST.cur.times.indexOf(labels[it[0].dataIndex])]}` : "") : "", label: c => c.parsed.y == null ? null : ` ${c.dataset.label}: ${fmt(c.parsed.y, 1)}`}}},
    scales: {x: {grid: {display: false}, ticks: {color: fg2, autoSkip: false, maxRotation: 0, callback: (v, i) => nd <= 1 ? (i % 6 === 5 ? labels[i].slice(11) : "") : (labels[i].slice(11) === "00:30" ? dayLabel(labels[i].slice(0, 10)) : "")}},
      y: {stacked, grid: {color: line}, border: {display: false}, ticks: {color: fg2, callback: v => fmt(v)}, title: {display: !!y, text: y, color: fg2, font: {size: 11}}}},
    elements: {point: {radius: 0, hoverRadius: 4}, line: {borderWidth: 2, tension: 0}}};
}
const DASH = {borderDash: [5, 4], borderWidth: 1.5};
function mkChart(id, cfg) { if (ST.charts[id]) ST.charts[id].destroy(); ST.charts[id] = new Chart($(id), cfg); }
function legend(id, items) { $(id).innerHTML = items.map(x => `<span><i class="${x.line ? "ln" : ""}" style="background:${x.dash ? "repeating-linear-gradient(90deg," + x.color + " 0 4px,transparent 4px 7px)" : x.color}"></i>${esc(x.label)}</span>`).join(""); }
const kpiHtml = list => list.map(k => `<div class="kpi"><span class="lbl">${k[0]}</span><span class="v">${k[1]}<small>${k[2]}</small></span>${k[3] ? `<span class="d">${esc(k[3])}</span>` : ""}</div>`).join("");
const SERIES = ["--s-hidro", "--s-gas", "--s-eol", "--s-sol", "--c-mag", "--s-bio", "--s-die", "--c-red"];
const TV = {"HIDROELÉCTRICA": "--s-hidro", "TERMOELÉCTRICA": "--s-gas", "EÓLICA": "--s-eol", "SOLAR": "--s-sol"};


/* ============ Justificación automática del costo marginal ============ */
const tShort = t => `${dayLabel(t.slice(0, 10))} ${t.slice(11)}`;
const toMin = t => new Date(t + ":00").getTime();
function groupSlots(times, flags, val) { // agrupa medias horas marcadas (tolera 1 hueco)
  const ev = []; let cur = null;
  flags.forEach((f, i) => { if (f) { if (cur && i - cur.i1 <= 2) cur.i1 = i; else { cur = {i0: i, i1: i}; ev.push(cur); } } });
  return ev.map(e => { let mx = -Infinity, im = e.i0; for (let i = e.i0; i <= e.i1; i++) if (val[i] != null && val[i] > mx) { mx = val[i]; im = i; } return {...e, t0: times[e.i0], t1: times[e.i1], tmax: times[im], imax: im, max: mx}; });
}
function limitesFor(zona, t0, t1) {
  const L = ST.limites?.[ST.key] || [], a = toMin(t0) - 30 * 6e4, b = toMin(t1), out = [];
  for (const x of L) { if (zona && x.zona !== zona) continue;
    for (const [w0, w1] of x.ventanas) { const c0 = toMin(w0), c1 = toMin(w1); if (c0 <= b && c1 >= a) { const ov = Math.min(b, c1) - Math.max(a, c0), dur = c1 - c0; out.push({...x, w0, w1, cover: ov / Math.max(1, b - a), dur}); } } }
  return out.sort((p, q) => p.dur - q.dur);
}
const lc1 = t => t ? t[0].toLowerCase() + t.slice(1) : t;
function justifyCmg(ev, zona) {
  const lim = limitesFor(zona, ev.t0, ev.t1), w = ST.cur;
  const spec = lim.find(x => x.dur < 3 * 864e5), gen = lim.find(x => x.dur >= 3 * 864e5);
  const fl = zona === "Norte" ? w.flujo?.["FLUJO ENLACE CENTRO - NORTE"] : zona === "Sur" ? w.flujo?.["FLUJO ENLACE CENTRO - SUR"] : null;
  const flMax = fl ? Math.max(...fl.filter(v => v != null)) : null, flV = fl ? fl[ev.imax] : null, ratio = fl ? flV / flMax : null;
  const flRow = fl ? [`Flujo Centro–${zona}`, `${fmt(flV)} MW en el pico · ${fmt(100 * ratio)} % del máximo de la semana (${fmt(flMax)} MW)`] : null, F = (...r) => r.filter(Boolean);
  const flTxt = fl ? ` Flujo Centro–${zona} en el pico: ${fmt(flV)} MW (${fmt(100 * ratio)} % del máximo programado de la semana, ${fmt(flMax)} MW).` : "";
  if (spec) { const ch = typeof ipsiChain === "function" ? ipsiChain(spec.motivo) : {txt: "", src: ""}; return {texto: `Coincide con ${spec.equipo} (${spec.tramo}) operando a su límite del ${tShort(spec.w0)} al ${tShort(spec.w1)}: ${lc1(spec.motivo)}.${ch.txt}${flTxt}`, fuente: spec.fuente + ch.src, conf: spec.cover >= 0.5 ? "alta" : "media", cita: spec.motivo, filas: F([spec.fuente, `${spec.equipo} (${spec.tramo}) al límite, ${tShort(spec.w0)}–${spec.w1.slice(11)}`], ["Motivo", spec.motivo], ch.it && ["IPSI 3.1", `${ch.it.equipo} fuera de servicio ${winText(ch.it.ventanas)}: ${lc1(ch.it.desc)}`], flRow), corto: `${spec.equipo} al límite${ch.code ? ` · ${ch.code} fuera de servicio` : (spec.motivo.match(/L-\d{4}/) ? ` · mant. ${spec.motivo.match(/L-\d{4}/)[0]}` : "")}`}; }
  if (gen && ratio != null && ratio >= 0.85) return {texto: `El enlace Centro–${zona} va muy cargado (${fmt(100 * ratio)} % de su máximo de la semana) y el IPSO indica que opera con límites: ${lc1(gen.motivo)}. Eso separa el precio de la zona ${zona}.${flTxt}`, fuente: gen.fuente + " + flujos del Anexo1", conf: "media", cita: gen.motivo, filas: F([gen.fuente, `${gen.tramo} con límites toda la semana`], ["Motivo", gen.motivo], flRow), corto: `Enlace Centro–${zona} al ${fmt(100 * ratio)} % · con límites`};
  if (gen) return {texto: `El ${gen.tramo} opera con límites toda la semana (${lc1(gen.motivo)}), pero en este episodio el flujo no está cerca de su máximo.${flTxt}`, fuente: gen.fuente, conf: "baja", filas: F([gen.fuente, `${gen.tramo} con límites toda la semana`], flRow, ["Lectura", "El flujo no está cerca de su máximo: el límite no explica solo el precio"]), corto: `Enlace con límites, flujo ${fmt(100 * ratio)} %`};
  const dem = w.tot["Demanda YUPANA"], res = derive(w).resid, pr = (a, v) => a.filter(x => x != null && x > v).length / a.filter(x => x != null).length;
  const rD = pr(dem, dem[ev.imax]), rR = pr(res, res[ev.imax]), h = +ev.tmax.slice(11, 13);
  if (rR <= 0.15) return {texto: `Coincide con demanda residual alta: ${fmt(res[ev.imax])} MW (15 % más alto de la semana). Las térmicas más caras quedan marginando${h >= 17 && h <= 23 ? " en horas punta" : ""}.`, fuente: "Datos del Anexo1", conf: "media", filas: F(["Demanda residual", `${fmt(res[ev.imax])} MW · entre el 15 % más alto de la semana`], h >= 17 && h <= 23 && ["Hora", "Punta: entran las térmicas más caras"]), corto: `Demanda residual alta (${fmt(res[ev.imax])} MW)`};
  if (rD <= 0.1) return {texto: `Coincide con demanda alta: ${fmt(dem[ev.imax])} MW (10 % más alto de la semana).`, fuente: "Datos del Anexo1", conf: "media", filas: [["Demanda", `${fmt(dem[ev.imax])} MW · entre el 10 % más alto de la semana`]], corto: `Demanda alta (${fmt(dem[ev.imax])} MW)`};
  return {texto: `No encontré una restricción ni un cambio de demanda que lo explique en los documentos cargados.${flTxt}`, fuente: "Ninguna", conf: "ninguna", filas: F(["Revisado", "IPSO, IPSI, flujos y demanda: sin una causa clara"], flRow), corto: "Sin causa en los documentos"};
}
function cmgEvents(w) {
  const out = [], SR = w.cmg["Santa Rosa"], T = w.times;
  [{k: "Trujillo", z: "Norte"}, {k: "Socabaya", z: "Sur"}].forEach(b => {
    const sp = T.map((_, i) => w.cmg[b.k][i] == null || SR[i] == null ? null : w.cmg[b.k][i] - SR[i]);
    groupSlots(T, sp.map((v, i) => v != null && v > Math.max(10, 0.25 * (SR[i] || 0))), w.cmg[b.k]).sort((p, q) => q.max - p.max).slice(0, 3)
      .forEach(e => out.push({zona: b.z, barra: b.k, ...e, spread: sp[e.imax]}));
  });
  const vals = SR.filter(v => v != null).sort((a, b) => a - b), p95 = vals[Math.floor(vals.length * 0.95)];
  groupSlots(T, SR.map(v => v != null && v >= p95), SR).sort((p, q) => q.max - p.max).slice(0, 2).forEach(e => out.push({zona: "Centro", barra: "Santa Rosa", ...e, spread: null}));
  return out.map((e, j) => ({id: "E" + (j + 1), chart: "cmg", titulo: `${e.zona} · ${tShort(e.t0)}–${e.t1.slice(11)}${e.t0.slice(0, 10) !== e.t1.slice(0, 10) ? " (" + dayLabel(e.t1.slice(0, 10)) + ")" : ""}`,
    dato: `${e.barra} hasta ${fmt(e.max)} US$/MWh${e.spread != null ? ` · ${sgn(e.spread)} sobre Centro` : " · máximo de la semana"}`, auto: justifyCmg(e, e.zona), t: e.tmax, ep: {zona: e.zona, barra: e.barra, t0: e.t0, t1: e.t1, tmax: e.tmax, max: e.max, spread: e.spread},
    det: `${e.barra} (${e.zona}) entre ${e.t0} y ${e.t1}: máximo ${fmt(e.max, 2)} US$/MWh el ${tLabel(e.tmax)}${e.spread != null ? `, ${fmt(e.spread, 2)} sobre Santa Rosa` : ""}.`}));
}
function rdoCmgEvents() {
  const w = ST.cur, V = getVW(), T = w.times, out = [];
  const d = T.map((_, i) => V.cmg["Santa Rosa"][i] == null || V.cmgB["Santa Rosa"][i] == null ? null : V.cmg["Santa Rosa"][i] - V.cmgB["Santa Rosa"][i]);
  const I = new Set(idxs(w));
  const flags = d.map((v, i) => I.has(i) && v != null && Math.abs(v) > Math.max(20, 0.5 * (V.cmgB["Santa Rosa"][i] || 0)));
  const evs = []; ["A","B","C","D","E","F","G","H"].forEach(Lt => evs.push(...groupSlots(T, flags.map((f, i) => f && V.letter[i] === Lt), V.cmg["Santa Rosa"])));
  evs.sort((p, q) => q.max - p.max).slice(0, 5).sort((p, q) => p.i0 - q.i0).forEach((e, j) => {
    const date = e.t0.slice(0, 10), L = V.letter[e.i0], cz = ST.listado.find(x => x.fecha === date && x.reprograma === L);
    const lim = limitesFor("Centro", e.t0, e.t1).filter(x => x.dur < 3 * 864e5)[0];
    let auto; const later = ST.listado.find(x => x.fecha === date && x.reprograma > L && x.hora <= e.t1.slice(11) && !V.days.find(y => y.date === date)?.R.letters.includes(x.reprograma));
    if (cz) auto = {texto: `Lo introduce el reprograma ${L} (desde ${cz.hora}): ${cz.causa}.${lim ? ` Además, ${lim.equipo} (${lim.tramo}) opera a su límite en ese periodo.` : ""}`, fuente: "Listado de reprogramas" + (cz.fuente ? " (captura)" : ""), conf: later ? "media" : "alta", cita: cz.causa, filas: [["Reprograma " + L, `desde ${cz.hora}`], ["Causa", cz.causa], ...(lim ? [[lim.fuente, `${lim.equipo} (${lim.tramo}) al límite`]] : []), ...(later ? [["Ojo", `Desde las ${later.hora} rige el Reprog ${later.reprograma} (${lc1(later.causa)}), que no está cargado: esta parte puede haber cambiado`]] : [])], corto: cz.causa.length > 48 ? cz.causa.slice(0, 46) + "…" : cz.causa};
    else if (L) auto = {texto: `Lo introduce el reprograma ${L}; falta su causa en el listado.${lim ? ` En ese periodo ${lim.equipo} (${lim.tramo}) opera a su límite.` : ""}`, fuente: "Reprograma " + L, conf: lim ? "media" : "baja", filas: [["Reprograma " + L, "causa no cargada en el listado"], ...(lim ? [[lim.fuente, `${lim.equipo} (${lim.tramo}) al límite`]] : [])], corto: lim ? `${lim.equipo} al límite` : "Falta la causa en el listado"};
    else auto = {texto: "Sin reprograma asociado.", fuente: "Ninguna", conf: "ninguna"};
    if (cz && auto && typeof postFacts === "function") { const fx = postFacts(date, cz.causa, cz.hora); if (fx.length) { auto = {...auto, filas: [...(auto.filas || []), ...fx.map(x => ["IEOD (hecho)", factTxt(x)])], fuente: auto.fuente + " + IEOD", texto: auto.texto + " IEOD: " + fx.map(factTxt).join("; ") + "."}; } }
    out.push({id: "G" + (j + 1), chart: "rdocmg", titulo: `Reprog ${V.letter[e.i0]} · ${tShort(e.t0)}–${e.t1.slice(11)}`, dato: `CMg Centro hasta ${fmt(e.max)} US$/MWh · ${sgn(d[e.imax])} vs ${V.baseOf[e.imax]}`, auto, t: e.tmax, ep: {zona: "Reprog " + L, barra: "Santa Rosa", t0: e.t0, t1: e.t1, tmax: e.tmax, max: e.max, spread: d[e.imax], letter: L},
      det: `CMg Santa Rosa reprogramado entre ${e.t0} y ${e.t1}: máximo ${fmt(e.max, 2)} US$/MWh (${fmt(d[e.imax], 2)} sobre la base) con el reprograma ${L || "—"}.`});
  });
  return out;
}

/* ============ Páginas ============ */
function renderTipos() {
  const w = ST.cur, I = idxs(w), dv = deriveF(w);
  $("tipoList").innerHTML = TIPOS.map(t => { const a = pick(dv.byTipo[t], I); return `<div class="tipo"><span class="ti" style="color:${css(TV[t])};background:color-mix(in srgb,${css(TV[t])} 14%,transparent)">${typeof tipoIcon === "function" ? tipoIcon(t) : ""}</span><div><b>${t[0] + t.slice(1).toLowerCase()}</b><div class="v">${fmt(gwh(a), 1)} GWh · máx ${fmt(Math.max(0, ...a.map(v => v || 0)))} MW</div></div></div>`; }).join("");
}
function renderPSO() {
  const w = ST.cur, I = idxs(w), L = pick(w.times, I), dv = deriveF(w);
  const cm = {}; BARRAS.forEach(b => cm[b.k] = pick(w.cmg[b.k], I));
  const allc = Object.values(cm).flat().filter(v => v != null).sort((a, b) => a - b), p95 = allc[Math.floor(allc.length * 0.95)] || 0, cmax = allc[allc.length - 1] || 0;
  const peaks = BARRAS.map(b => { const s = stats(cm[b.k], {times: L}); return s.tmax ? {t: s.tmax, text: `${b.z} ${fmt(s.max)}`} : null; }).filter(Boolean);
  const sc = ST.cmgScale === "det" ? "det" : "todo", out = cmax > 3 * p95 && p95 > 0;
  const co = tsOpts(L, {y: "US$/MWh", marks: sc === "det" && out ? peaks.filter(p => +p.text.split(" ").pop().replace(/,/g, "") > 3 * p95) : []});
  let note = "";
  if (sc === "det" && out) { co.scales.y.max = Math.ceil(p95 * 1.6 / 10) * 10; co.scales.y.min = 0; note = "Detalle: picos fuera de escala marcados arriba. Usa «Todo» para verlos completos."; }
  else { co.scales.y.min = 0; co.scales.y.max = Math.ceil(cmax * 1.08 / 100) * 100; if (out) note = "Escala completa. Usa «Detalle» para ver mejor los valores habituales."; }
  document.querySelectorAll("#cmgScale button").forEach(b => b.setAttribute("aria-pressed", b.dataset.sc === sc));
  $("cmgNote").textContent = note;
  mkChart("cCmg", {type: "line", data: {labels: L, datasets: BARRAS.map(b => ({label: b.z, data: cm[b.k], borderColor: css(b.v), backgroundColor: css(b.v)}))}, options: co});
  legend("lgCmg", BARRAS.map(b => ({label: `${b.z} (barra ${b.k})`, color: css(b.v), line: true})).concat(note ? [{label: "Etiqueta arriba = pico fuera de escala", color: css("--warn"), line: true, dash: true}] : []));
  const order = GROUPS.filter(g => gwh(pick(dv.bySrc[g.k], I)) >= 0.05), surf = css("--panel");
  mkChart("cSrc", {type: "line", data: {labels: L, datasets: order.map(g => ({label: g.k, data: pick(dv.bySrc[g.k], I), fill: true, backgroundColor: css(g.v), borderColor: surf, borderWidth: 1, stack: "s"}))}, options: tsOpts(L, {stacked: true, y: "MW"})});
  legend("lgSrc", order.map(g => ({label: `${g.k} ${gfmt(gwh(pick(dv.bySrc[g.k], I)))} GWh`, color: css(g.v)})));
  mkChart("cDr", {type: "line", data: {labels: L, datasets: [{label: "Residual", data: pick(dv.resid, I), borderColor: css("--accent"), backgroundColor: css("--accent-soft"), fill: true},
    {label: isFiltered() ? "Térmica (filtro)" : "Térmica", data: pick(dv.byTipo["TERMOELÉCTRICA"], I), borderColor: css("--fg-3"), ...DASH, fill: false}]}, options: tsOpts(L, {y: "MW"})});
  legend("lgDr", [{label: "Demanda residual (lo que deben cubrir las térmicas)", color: css("--accent")}, {label: "Generación térmica programada", color: css("--fg-3"), line: true, dash: true}]);
  const F = findingsFor("pso"); renderEpis("why-cmg", "pso", "cCmg", F.filter(f => f.chart === "cmg" && (!f.t || !ST.F.days.size || ST.F.days.has(f.t.slice(0, 10))))); [["why-src", "src"], ["why-dr", "dr"]].forEach(([b, c]) => renderWhy(b, "pso", F.filter(f => f.chart === c)));
}
function rdoBaseName(V, I) { const s = new Set(I.map(i => V.baseOf[i])); return s.size === 1 ? [...s][0] : "PDO/PSO"; }
function renderRDO() {
  const w = ST.cur, V = getVW(), I = idxs(w), L = pick(w.times, I), has = V.days.length > 0;
  $("rdoEmpty").hidden = has; document.querySelector("#tab-rdo .rdoBody").hidden = !has; if (!has) { ST.rdoCur = null; renderWhyEmpty(); return; }
  const B = rdoBaseName(V, I.filter(i => V.letter[i])), marks = V.marks.filter(m => L.includes(m.t)), cov = I.filter(i => V.letter[i]);
  const focus = rdoFocusDay(); ST.rdoCur = focus ? rdoDay(focus) : null;
  // KPIs
  const gV = sumCen(V.cen, cov, pass), gB = sumCen(V.cenB, cov, pass), dE = (sum(gV) - sum(gB)) * 0.5;
  const crV = stats(pick(V.cmg["Santa Rosa"], cov), {times: pick(w.times, cov)}), crB = stats(pick(V.cmgB["Santa Rosa"], cov), {times: []});
  const nRep = V.days.filter(x => !ST.F.days.size || ST.F.days.has(x.date)).reduce((s, x) => s + x.R.letters.length, 0);
  $("rdoKpis").innerHTML = kpiHtml([["Reprogramas", nRep, "", `${V.days.length} día(s) con reprograma`], ["Energía vs " + B, sgn(dE / 1000, 2), "GWh", isFiltered() ? "con filtro" : "sistema"],
    ["CMg Centro vigente", fmt(crV.avg, 1), "US$/MWh", crV.avg != null ? sgn(crV.avg - crB.avg, 1) + " vs " + B : ""], ["CMg máx. vigente", fmt(crV.max, 0), "US$/MWh", crV.tmax ? tLabel(crV.tmax) : ""]].concat((() => { if (typeof ejecWeek !== "function") return []; const ej = ejecWeek(w), J = cov.filter(i => ej[i] != null); if (!J.length) return [];
      const dE2 = (sum(J.map(i => ej[i])) - sum(sumCen(V.cen, J, pass))) * 0.5; return [["Ejecutado vs vigente", sgn(dE2 / 1000, 2), "GWh", "IEOD · " + [...new Set(J.map(i => dayLabel(w.times[i].slice(0, 10))))].join(", ")]]; })()));
  renderRdoCharts(w, V, I, L);
  // Tabla RDO por central
  const rows = Object.keys(V.cen).filter(pass).map(c => { const v = pick(V.cen[c], cov), b = pick(V.cenB[c], cov); return {c, mx: Math.max(0, ...v.map(x => x || 0)), d: (sum(v) - sum(b)) * 0.5}; }).filter(x => x.mx > 0 || Math.abs(x.d) > 1).sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
  $("tblRdo").innerHTML = `<table><thead><tr><th>RDO</th><th>Tipo</th><th class="n">Máx. MW</th><th class="n">Δ MWh vs ${B}</th></tr></thead><tbody>${rows.slice(0, 60).map(x => `<tr data-filter="${esc(x.c)}" class="${ST.F.central === x.c ? "sel" : ""}"><td>${esc(x.c)}${isStk(x.c) ? ' <span class="pill media">STK</span>' : ""}</td><td class="faint">${esc((tipoOf(x.c) || "—").toLowerCase())}</td><td class="n">${fmt(x.mx, 1)}</td><td class="n ${x.d > 0 ? "pos" : x.d < 0 ? "neg" : ""}">${sgn(x.d)}</td></tr>`).join("")}</tbody></table>`;
  // Causas
  const dates = new Set(days(w)), cz = ST.listado.filter(x => dates.has(x.fecha) && (!ST.F.days.size || ST.F.days.has(x.fecha))).sort((a, b) => (b.fecha + b.hora).localeCompare(a.fecha + a.hora));
  $("causasSrc").textContent = cz.some(x => x.fuente) ? "incluye causas copiadas de la captura del Power BI" : "";
  $("tblCausas").innerHTML = cz.length ? `<table><thead><tr><th>Fecha</th><th>Hora</th><th>RDO</th><th>Causas</th></tr></thead><tbody>${cz.map(x => `<tr data-day="${x.fecha}"><td class="mono">${x.fecha.slice(8)}/${x.fecha.slice(5, 7)}</td><td class="mono">${esc(x.hora)}</td><td class="mono"><b>${esc(x.reprograma)}</b></td><td>${esc(x.causa)}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Agrega el listado de reprogramas para ver las causas.</p>`;
  const Fd = findingsFor("rdo"), sub = ST.rdoCur ? `Hallazgos del ${dayLabel(ST.rdoCur.date)} (elige un día arriba)` : "";
  renderEpis("why-rcmg", "rdo", "cRCmg", rdoCmgEvents(), sub, Fd.filter(f => f.chart === "rdocmg" && !/^G/.test(f.id)));
  renderWhy("why-rgen", "rdo", Fd.filter(f => f.chart === "rdo"), sub);
  renderWhy("why-rres", "rdo", Fd.filter(f => f.chart === "rdores"), sub);
}
function renderWhyEmpty() { ["why-rcmg", "why-rgen", "why-rres", "why-th", "why-ren"].forEach(id => { const el = $(id); if (el) el.innerHTML = ""; }); }
function renderTH() {
  const w = ST.cur, V = getVW(), I = idxs(w), L = pick(w.times, I), cov = I.filter(i => V.letter[i]), B = rdoBaseName(V, cov);
  const isTh = c => pass(c) && THERM_SRC.includes(MAESTRO.get(c)?.source), onlyCov = a => I.map(i => V.letter[i] ? (a[i] ?? 0) : null);
  const th = Object.keys(V.cen).filter(isTh).map(c => ({c, a: onlyCov(V.cen[c])})).map(x => ({...x, e: sum(x.a)})).filter(x => x.e > 0).sort((p, q) => q.e - p.e);
  const top = th.slice(0, 7), rest = th.slice(7), surf = css("--panel");
  const ds = top.map((x, k) => ({label: x.c, data: x.a, fill: true, backgroundColor: css(SERIES[k]), borderColor: surf, borderWidth: 1, stack: "v", order: 2}));
  if (rest.length) ds.push({label: `Otras (${rest.length})`, data: L.map((_, j) => rest.reduce((s, x) => s + (x.a[j] || 0), 0)).map((v, j) => V.letter[I[j]] ? v : null), fill: true, backgroundColor: css("--s-sin"), borderColor: surf, borderWidth: 1, stack: "v", order: 2});
  ds.push(...baseLines(w, V, I, a => sumCen(a, w.times.map((_, i) => i), isTh), "Total térmico", css("--fg")));
  mkChart("cTh", {type: "line", data: {labels: L, datasets: ds}, options: (o => { o.scales.y.stacked = true; return o; })(biOpts(L, {y: "Potencia (MW)", min: 0}))});
  legend("lgTh", ds.map(x => ({label: x.label, color: x.fill ? x.backgroundColor : x.borderColor, line: !x.fill})));
  const rows = th.map(x => ({c: x.c, mxV: Math.max(0, ...pick(V.cen[x.c], cov).map(v => v || 0)), mxB: Math.max(0, ...pick(V.cenB[x.c], cov).map(v => v || 0)), d: (sum(pick(V.cen[x.c], cov)) - sum(pick(V.cenB[x.c], cov))) * 0.5})).sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
  $("tblTh").innerHTML = `<table><thead><tr><th>Central</th><th>Fuente</th><th class="n">Máx. ${B}</th><th class="n">Máx. reprog</th><th class="n">Δ MWh</th></tr></thead><tbody>${rows.slice(0, 30).map(x => `<tr data-filter="${esc(x.c)}"><td>${esc(x.c)}</td><td class="faint">${esc(MAESTRO.get(x.c)?.source || "")}</td><td class="n">${fmt(x.mxB)}</td><td class="n">${fmt(x.mxV)}</td><td class="n ${x.d > 0 ? "pos" : x.d < 0 ? "neg" : ""}">${sgn(x.d)}</td></tr>`).join("")}</tbody></table>`;
  const focus = rdoFocusDay(); ST.rdoCur = focus ? rdoDay(focus) : null;
  renderWhy("why-th", "rdo", findingsFor("rdo").filter(f => f.chart === "rdo"), ST.rdoCur ? `Hallazgos del ${dayLabel(ST.rdoCur.date)}` : "Sin reprogramas");
}
function renderREN() {
  const w = ST.cur, V = getVW(), I = idxs(w), L = pick(w.times, I), cov = I.filter(i => V.letter[i]), B = rdoBaseName(V, cov);
  const ds = [], all = w.times.map((_, i) => i);
  ["EÓLICA", "SOLAR"].forEach(t => { const f = c => pass(c) && tipoOf(c) === t, nm = t === "EÓLICA" ? "Eólica" : "Solar";
    ds.push(...baseLines(w, V, I, a => sumCen(a, all, f), nm, css(TV[t])));
    ds.push({label: `${nm} RDO`, data: sumCen(V.cen, I, f), borderColor: css(TV[t]), backgroundColor: css(TV[t]), borderDash: [2, 3], borderWidth: 3, borderCapStyle: "round"}); });
  mkChart("cRen", {type: "line", data: {labels: L, datasets: ds}, options: biOpts(L, {y: "Potencia (MW)", min: 0})});
  legend("lgRen", ds.map(x => ({label: x.label, color: x.borderColor, line: true, dash: !!x.borderDash})));
  const rows = Object.keys(V.cen).filter(c => pass(c) && ["EÓLICA", "SOLAR"].includes(tipoOf(c))).map(c => ({c, d: (sum(pick(V.cen[c], cov)) - sum(pick(V.cenB[c], cov))) * 0.5, mx: Math.max(0, ...pick(V.cen[c], cov).map(v => v || 0))})).sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
  $("tblRen").innerHTML = rows.length ? `<table><thead><tr><th>Central</th><th>Tipo</th><th class="n">Máx. reprog</th><th class="n">Δ MWh vs ${B}</th></tr></thead><tbody>${rows.slice(0, 30).map(x => `<tr data-filter="${esc(x.c)}"><td>${esc(x.c)}</td><td class="faint">${esc(tipoOf(x.c).toLowerCase())}</td><td class="n">${fmt(x.mx)}</td><td class="n ${x.d > 0 ? "pos" : x.d < 0 ? "neg" : ""}">${sgn(x.d)}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Sin datos.</p>`;
  const focus = rdoFocusDay(); ST.rdoCur = focus ? rdoDay(focus) : null;
  const ren = findingsFor("rdo").filter(f => f.chart === "rdo" && (f.causa || /Cambio/.test(f.titulo)));
  renderWhy("why-ren", "rdo", ren, ST.rdoCur ? `Hallazgos del ${dayLabel(ST.rdoCur.date)}` : "Sin reprogramas");
}

/* ============ Inspector de momento ============ */
let playH = null;
function setMoment(t) { ST.moment = t; ST.momentAI = ""; $("moment").hidden = !t; document.body.classList.toggle("has-moment", !!t); renderMoment(); Object.values(ST.charts).forEach(c => { try { c.draw(); } catch (e) {} }); }
function stepMoment(k) { const w = ST.cur, I = idxs(w); if (!I.length) return; let j = I.findIndex(i => w.times[i] === ST.moment); j = j < 0 ? 0 : Math.max(0, Math.min(I.length - 1, j + k)); setMoment(w.times[I[j]]); }
function momentData(t) {
  const w = ST.cur, i = w.times.indexOf(t); if (i < 0) return null; const V = getVW(), L = V.letter[i], B = V.baseOf[i];
  const cm = BARRAS.map(b => ({z: b.z, pso: w.cmg[b.k][i], base: V.cmgB[b.k][i], rdo: V.cmg[b.k][i]}));
  const fl = Object.entries(w.flujo || {}).map(([h, a]) => ({h: h.replace("FLUJO ", ""), v: a[i]})).filter(x => x.v != null);
  const ch = Object.keys(V.cen).filter(pass).map(c => ({c, b: V.cenB[c]?.[i] || 0, v: V.cen[c]?.[i]})).filter(x => x.v != null).map(x => ({...x, d: x.v - x.b})).filter(x => Math.abs(x.d) > 0.5).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 6);
  const big = Object.keys(w.centrales).filter(pass).map(c => ({c, v: (L ? V.cen[c]?.[i] : null) ?? w.centrales[c][i] ?? 0})).sort((a, b) => b.v - a.v).slice(0, 5);
  const causas = ST.listado.filter(x => x.fecha === t.slice(0, 10) && (!L || x.reprograma <= L));
  return {t, i, L, B, cm, fl, ch, big, causas, dem: w.tot["Demanda YUPANA"][i], res: L ? V.resid[i] : V.residB[i], re: L ? V.re[i] : V.reB[i]};
}
function renderMoment() {
  const t = ST.moment; if (!t) return; const M = momentData(t); if (!M) { setMoment(null); return; }
  $("mTime").textContent = tLabel(t); $("mLetter").textContent = M.L ? `Reprograma ${M.L} vigente · base ${M.B}` : "Programa (sin reprograma)";
  const c = (k, v) => `<div class="c"><span>${k}</span><b>${v}</b></div>`;
  $("mGrid").innerHTML = M.cm.map(x => c(`CMg ${x.z}`, `${fmt(x.rdo ?? x.base, 1)}${x.rdo != null ? ` <span class="faint">(${M.B} ${fmt(x.base, 1)})</span>` : ""}`)).join("")
    + c("Demanda YUPANA", fmt(M.dem) + " MW") + c("Residual", fmt(M.res) + " MW") + c("Reserva eficiente", M.re != null ? fmt(M.re) + " MW" : "—")
    + M.fl.filter(x => /ENLACE/.test(x.h)).map(x => c(x.h.replace("ENLACE ", "Enlace "), fmt(x.v) + " MW")).join("")
    + (M.ch.length ? `<div class="c wide"><span>Mayores cambios vs ${M.B}</span><b>${M.ch.map(x => `${esc(x.c)} ${sgn(x.d)}`).join(" · ")}</b></div>` : "")
    + `<div class="c wide"><span>Mayores despachos${isFiltered() ? " (filtro)" : ""}</span><b>${M.big.map(x => `${esc(x.c)} ${fmt(x.v)}`).join(" · ")}</b></div>`
    + (M.causas.length ? `<div class="c wide"><span>Causas del día</span><b style="font-family:var(--f-body)">${M.causas.map(x => `${esc(x.reprograma)} ${esc(x.hora)}: ${esc(x.causa)}`).join("<br>")}</b></div>` : "");
  $("mAi").textContent = ST.momentAI || (SAMPLE ? "Pulsa «Explicar este momento» para que la IA lo cruce con el IPSO, el IPDO y las causas." : "La IA funciona al abrir la página dentro de Claude.");
  $("mAsk").disabled = !SAMPLE || !!ST.momentBusy; $("mPlay").textContent = playH ? "■ Detener" : "▶ Reproducir";
}
async function askMoment() {
  const M = momentData(ST.moment); if (!M || !SAMPLE) return;
  const det = `Momento: ${tLabel(M.t)} (${M.t}). ${M.L ? `Reprograma ${M.L} vigente; base de comparación ${M.B}.` : "Sin reprograma: rige el programa."}
CMg (US$/MWh): ${M.cm.map(x => `${x.z} ${fmt(x.rdo ?? x.base, 2)}${x.rdo != null ? ` (base ${fmt(x.base, 2)})` : ""}`).join("; ")}.
Demanda ${fmt(M.dem)} MW; residual ${fmt(M.res)} MW; reserva eficiente ${M.re != null ? fmt(M.re) : "—"} MW.
Flujos: ${M.fl.map(x => `${x.h} ${fmt(x.v)} MW`).join("; ")}.
Cambios vs base: ${M.ch.map(x => `${x.c} ${sgn(x.d, 1)} MW`).join("; ") || "—"}.
Causas publicadas del día: ${M.causas.map(x => `${x.reprograma} ${x.hora}: ${x.causa}`).join(" | ") || "—"}.`;
  const kw = [...GENERIC_KW, ...M.ch.map(x => x.c), ...M.fl.map(x => x.h), "límite", "capacidad", tLabel(M.t).split(" ")[1]];
  const W = ST.wk, ip = ST.ipdo[M.t.slice(0, 10)];
  const docs = [W?.ipso ? `=== IPSO ===\n${excerpt(W.ipso, kw, 30000)}` : "", ip ? `=== IPDO ${M.t.slice(0, 10)} ===\n${excerpt(ip.text, kw, 25000)}` : "", W?.ipsi ? `=== IPSI ===\n${excerpt(W.ipsi, kw, 20000)}` : ""].filter(Boolean).join("\n\n");
  const q = `Eres analista de operación del SEIN (COES, Perú). Explica en 3 a 5 frases cortas qué está pasando en este momento de la operación programada y por qué, usando SOLO los documentos y las causas publicadas. Si algo no está explicado en ellos, dilo. Cita entre comillas el fragmento del documento que lo justifica (máx. 20 palabras) e indica de qué documento sale.\n\n${det}\n\nDOCUMENTOS:\n${docs || "(ninguno)"}`;
  ST.momentBusy = true; ST.momentAI = "Leyendo informes…"; renderMoment();
  try { const {text} = await SAMPLE(q, {cache: false, onText: ({text}) => { ST.momentAI = text; $("mAi").textContent = text; }}); ST.momentAI = text; }
  catch (e) { ST.momentAI = e.text || ({not_granted: "Sin permiso para usar Claude.", rate_limited: "Demasiadas consultas. Espera un momento."}[e.code] ?? `No se pudo completar (${e.code || "error"}).`); }
  ST.momentBusy = false; renderMoment();
}
function togglePlay() {
  if (playH) { clearInterval(playH); playH = null; renderMoment(); return; }
  if (!ST.moment) { const I = idxs(ST.cur); setMoment(ST.cur.times[I[0]]); }
  playH = setInterval(() => { const w = ST.cur, I = idxs(w), j = I.findIndex(i => w.times[i] === ST.moment); if (j < 0 || j >= I.length - 1) { clearInterval(playH); playH = null; renderMoment(); return; } setMoment(w.times[I[j + 1]]); }, 450);
  renderMoment();
}

/* ============ Barra, filtros y navegación ============ */
function toast(msg, err) { const t = $("toast"); t.textContent = msg || ""; t.classList.toggle("err", !!err); clearTimeout(toast.h); if (msg) toast.h = setTimeout(() => { t.textContent = ""; }, 7000); }
function renderBar() {
  const w = ST.cur, keys = sortedKeys(), i = keys.indexOf(ST.key);
  $("wkNum").textContent = w.semana; $("wkFrom").textContent = w.inicio.slice(8) + "/" + w.inicio.slice(5, 7); $("wkTo").textContent = w.fin.slice(8) + "/" + w.fin.slice(5, 7);
  $("wkSel").innerHTML = keys.map(k => { const x = WEEKS.get(k).w; return `<option value="${k}">Sem ${x.semana} · ${x.inicio.slice(8)}/${x.inicio.slice(5, 7)}</option>`; }).join(""); $("wkSel").value = ST.key;
  $("wkPrev").disabled = i <= 0; $("wkNext").disabled = i >= keys.length - 1;
  const dates = days(w);
  for (const d of [...ST.F.days]) if (!dates.includes(d)) ST.F.days.delete(d);
  $("dayChips").innerHTML = dates.map(d => `<button type="button" data-d="${d}" aria-pressed="${ST.F.days.has(d)}" class="${ST.reprog[d] ? "has" : ""}" title="${ST.reprog[d] ? "Con reprogramas: " + Object.keys(ST.reprog[d]).join(", ") : ""}"><small>${DOW[new Date(d + "T12:00").getDay()]}</small>${d.slice(8)}/${d.slice(5, 7)}</button>`).join("");
  renderSrcTree();
  const names = Object.keys(w.centrales).sort(), emps = [...new Set(names.map(empOf).filter(Boolean))].sort();
  $("dlCentral").innerHTML = emps.map(e => `<option value="Empresa: ${esc(e)}">`).join("") + names.map(c => `<option value="${esc(c)}">${esc(empOf(c))}</option>`).join("");
  $("fCentral").value = ST.F.central;
  $("baseBox").hidden = !["rdo", "th", "ren"].includes(ST.tab); if (!$("baseBox").hidden) renderBase();
  const nRep = dates.reduce((s, d) => s + Object.keys(ST.reprog[d] || {}).length, 0), nLis = ST.listado.filter(x => dates.includes(x.fecha)).length, nPdo = dates.filter(d => ST.pdo[d]).length, nIpdo = dates.filter(d => ST.ipdo[d]).length;
  $("srcs").innerHTML = [["IPSO", !!ST.wk.ipso], ["IPSI", !!ST.wk.ipsi], [`PDO ${nPdo || ""}`.trim(), nPdo > 0], [`IPDO ${nIpdo || ""}`.trim(), nIpdo > 0], [`Reprog ${nRep || ""}`.trim(), nRep > 0], [`Causas ${nLis || ""}`.trim(), nLis > 0], [`Sem ${w.semana - 1}`, !!prevWeek()]]
    .map(([k, ok]) => `<span class="src ${ok ? "ok" : ""}" data-pick role="button" tabindex="0"><i></i>${k}</span>`).join("") + `<span class="src" id="pasteChip" role="button" tabindex="0">Pegar texto</span>`;
  $("pageTitle").textContent = TITLES[ST.tab];
  $("filters").hidden = ["chat"].includes(ST.tab);
  renderTipos(); if (typeof renderNow === "function") renderNow();
  $("foot").textContent = `Despacho: ${w.fuente}. Clasificación: CENTRALES.xlsx. Residual = demanda YUPANA − cogeneración − RER − hidro. Clic en un gráfico para inspeccionar la media hora.`;
}
const EMBsources = () => [...MAESTRO.values()].map(m => m.source).filter(Boolean);
function renderTab() {
  if (!ST.cur) return; const t = ST.tab;
  if (t === "pso") renderPSO(); if (t === "rdo") renderRDO(); if (t === "th") renderTH(); if (t === "ren") renderREN();
  if (t === "ipsi") renderIPSI(); if (t === "post" && typeof renderPOST === "function") renderPOST(); if (t === "cmp") renderCMP(); if (t === "chat") renderChat();
  if (["pso", "ipsi", "cmp"].includes(t)) renderAIBar(t); else if (["rdo", "th", "ren"].includes(t)) renderAIBarRdo(t);
  if (ST.moment) renderMoment();
}
function renderAIBarRdo(t) { const el = $("aiBar-" + t); if (!el) return; const st = aiState("rdo"), ready = SAMPLE && ST.rdoCur;
  el.innerHTML = `<button class="primary" type="button" data-run="rdo" ${ready && !st.loading ? "" : "disabled"}>${st.loading ? "Justificando…" : st.result ? "Volver a justificar" : "Justificar con IA"}</button>${st.loading ? `<button type="button" data-stop="rdo">Detener</button>` : ""}<span class="unit ${st.error ? "neg" : ""}">${esc(st.error || (!SAMPLE ? "La IA funciona al abrir la página dentro de Claude." : ST.rdoCur ? `Reprogramas del ${dayLabel(ST.rdoCur.date)}` : ""))}</span>`; }
function renderAll() { renderBar(); renderTab(); }
function setWeek(k) { if (!WEEKS.has(k)) return; ST.key = k; ST.rdoCur = null; ST.F.days.clear(); setMoment(null); try { localStorage.setItem("pso-week", k); } catch (e) {} renderAll(); }
function setTab(t) {
  ST.tab = t;
  document.querySelectorAll("#navTabs button").forEach(b => b.setAttribute("aria-selected", b.dataset.tab === t));
  document.querySelectorAll("main > section").forEach(s => s.hidden = s.id !== "tab-" + t);
  try { localStorage.setItem("pso-tab", t); } catch (e) {}
  renderAll();
}

/* ============ Intervenciones (IPSI) y restricciones de suministro ============ */
ST.ipsiItems = BOOT.ipsiItems || {}; ST.restric = BOOT.restric || {};
try { const s = JSON.parse(localStorage.getItem("pso-extra") || "null"); if (s) { Object.assign(ST.ipsiItems, s.ipsiItems || {}); Object.assign(ST.limites, s.limites || {}); } } catch (e) {}
function persistExtra() { try { const keep = o => Object.fromEntries(Object.entries(o).filter(([k]) => !(BOOT.ipsiItems || {})[k] && !(BOOT.limites || {})[k])); localStorage.setItem("pso-extra", JSON.stringify({ipsiItems: keep(ST.ipsiItems), limites: keep(ST.limites)})); } catch (e) {} }
const items = () => ST.ipsiItems[ST.key] || [];
function winText(V) {
  const w = ST.cur; if (V.length === 1 && V[0][0] <= w.inicio + "T00:30" && V[0][1] >= w.fin + "T23:30") return "toda la semana";
  return V.map(([a, b]) => a.slice(0, 10) === b.slice(0, 10) ? `${dayLabel(a.slice(0, 10))} ${a.slice(11)}–${b.slice(11)}` : `${tShort(a)} → ${tShort(b)}`).join(" · ");
}
const overlapDay = (V, day) => V.some(([a, b]) => a.slice(0, 10) <= day && b.slice(0, 10) >= day);
function ipsiChain(text) { // "mantenimiento de la línea L-5008" -> intervención del IPSI
  const codes = String(text).match(/L-\d{4}/g) || []; const it = items().find(x => codes.some(c => x.equipo.includes(c)) && x.tipo === "Transmisión");
  return it ? {txt: ` El IPSI lo confirma: ${it.equipo} (${it.ubic}) fuera de servicio ${winText(it.ventanas)} por ${lc1(it.desc)}.`, src: " + IPSI 3.1", code: it.equipo, it} : {txt: "", src: ""};
}
function dropJust(x) {
  const pj = typeof dropPostJust === "function" ? dropPostJust(x) : null; if (pj) return pj;
  const hit = items().filter(it => it.centrales.includes(x.central) && x.dias.some(d => overlapDay(it.ventanas, d)));
  if (hit.length) return {texto: hit.map(it => `IPSI: ${it.ubic} ${it.equipo}, ${lc1(it.desc)} (${winText(it.ventanas)}).`).join(" "), fuente: "IPSI 3.1", conf: "alta", cita: hit[0].desc};
  const nz = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  const base = nz(MAESTRO.get(x.central)?.central || x.central).replace(/[-_]/g, " ").replace(/\s+(TG|TV|G|U|CC)\s*\d.*$/, "").trim();
  const lzW = base.length >= 4 ? ST.listado.filter(r => nz(r.causa).replace(/[-_]/g, " ").includes(base) && days(ST.cur).includes(r.fecha)) : [], lz = lzW.filter(r => x.dias.includes(r.fecha));
  if (!lz.length && lzW.length) return {texto: `En el PSO sale de servicio el ${dayList(x.dias)}. Aparte, el ${lzW.map(r => `${dayLabel(r.fecha)} el Reprog ${r.reprograma} la retira por ${lc1(r.causa)}`).join("; ")}: confirma que la central tiene un problema esta semana, aunque el IPSI no lo detalla.`, fuente: "Listado de reprogramas + PSO", conf: "media", cita: lzW[0].causa};
  if (lz.length) return {texto: lz.map(r => `Reprograma ${r.reprograma} del ${dayLabel(r.fecha)} (desde ${r.hora}): ${lc1(r.causa)}.`).join(" ") + (x.dias.length > lz.length ? " El IPSI no la incluye; los demás días de la caída siguen sin causa publicada." : ""), fuente: "Listado de reprogramas", conf: x.dias.length > 1 ? "media" : "alta", cita: lz[0].causa};
  const near = items().filter(it => it.centrales.includes(x.central));
  if (near.length) return {texto: `El IPSI tiene una intervención en ${x.central} (${near.map(it => it.equipo + ": " + winText(it.ventanas)).join("; ")}), pero no en los días de la caída.`, fuente: "IPSI 3.1", conf: "baja"};
  if (x.tipo === "TERMOELÉCTRICA") return {texto: "No figura entre las intervenciones relevantes del IPSI. Al ser térmica, lo más probable es que quede fuera por costo (orden de mérito frente al CMg) o por una intervención menor que solo aparece en el Anexo 1 de intervenciones.", fuente: "IPSI (no figura)", conf: "baja"};
  return {texto: "No figura entre las intervenciones relevantes del IPSI. Revisar los caudales del Anexo 2 de hidrología del PDO o el Anexo 1 de intervenciones.", fuente: "IPSI (no figura)", conf: "ninguna"};
}
const plugBands = {id: "bands", beforeDatasetsDraw(ch, a, opts) { try { const B = opts?.items || [], L = ch.data.labels; if (!B.length || !L) return; const {top, bottom} = ch.chartArea, c = ch.ctx, x = ch.scales.x;
  c.save(); c.fillStyle = css("--warn"); c.globalAlpha = 0.13;
  for (const [t0, t1] of B) { let i0 = L.findIndex(t => t >= t0), i1 = L.length - 1 - [...L].reverse().findIndex(t => t <= t1); if (i0 < 0 || i1 < i0) continue; const p0 = x.getPixelForValue(i0), p1 = x.getPixelForValue(i1); c.fillRect(p0, top, Math.max(2, p1 - p0), bottom - top); }
  c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugBands); } catch (e) {} }
function effectOf(it) {
  if (!it.centrales.length) return "";
  const w = ST.cur, P = prevWeek(); return it.centrales.map(c => { const a = w.centrales[c]; if (!a) return ""; const idx = w.times.map((t, i) => i).filter(i => it.ventanas.some(([s, e]) => w.times[i] >= s && w.times[i] <= e));
    const mx = Math.max(0, ...idx.map(i => a[i] || 0)), pm = P?.w.centrales[c] ? Math.max(...P.w.centrales[c]) : null;
    return `${c}: máx ${fmt(mx)} MW${pm != null ? ` (sem ${P.w.semana}: ${fmt(pm)})` : ""}`; }).filter(Boolean).join(" · ");
}
function renderIPSI() {
  const w = ST.cur, drops = detectDrops(w), F = findingsFor("ipsi"), ai = aiMap("ipsi");
  if (typeof renderImpact === "function") renderImpact();
  const its = items(), sel0 = ST.dropSel && (w.centrales[ST.dropSel]) ? ST.dropSel : drops[0]?.central; ST.dropSel = sel0;
  $("tblDrops").innerHTML = drops.length ? `<table><thead><tr><th>Central</th><th>Días</th><th>Estado</th><th class="n">MWh no desp.</th><th>Justificación</th></tr></thead><tbody>${drops.map((x, i) => { const h = ai.get(F[i].id) || F[i].auto, c = (h?.confianza || h?.conf || "").toLowerCase();
    return `<tr data-central="${esc(x.central)}" class="${x.central === ST.dropSel ? "sel" : ""}"><td><b>${esc(x.central)}</b>${isStk(x.central) ? ' <span class="pill media">STK</span>' : ""}<div class="faint">${esc(x.empresa)}</div></td><td style="white-space:nowrap">${dayList(x.dias)}</td><td>${x.tipoCaida}</td><td class="n">${fmt(x.mwh)}</td><td>${/^(alta|media)$/.test(c) ? `<span class="faint">${esc(h?.fuente || "")}</span>` : "—"}</td></tr>`; }).join("")}</tbody></table>` : `<p class="muted">Sin caídas relevantes.</p>`;
  if (sel0) {
    const fi = F.filter(f => f.central === sel0), extra = its.filter(it => it.centrales.includes(sel0) && !drops.some(d => d.central === sel0));
    renderWhy("why-drop", "ipsi", fi.length ? fi : extra.map((it, j) => ({id: "X" + j, titulo: `${it.ubic} ${it.equipo}`, dato: winText(it.ventanas), auto: {texto: `${it.desc}. ${effectOf(it)}`, fuente: "IPSI 3.1", conf: "alta"}})), `${esc(empOf(sel0))} · ${esc((tipoOf(sel0) || "").toLowerCase())}`);
  }
  const days0 = [...ST.F.days], vis = its.filter(it => (!days0.length || days0.some(d => overlapDay(it.ventanas, d))) && (!isFiltered() || it.centrales.some(pass)));
  $("ipsiSrc").textContent = its.length ? `${its.length} intervenciones · ${ST.wk.ipsiName || "IPSI"}` : (ST.wk.ipsi ? "IPSI cargado: pulsa «Extraer tablas con IA»" : "Falta el IPSI");
  $("tblIpsi").innerHTML = vis.length ? `<table><thead><tr><th>Tipo</th><th>Equipo</th><th>Ubicación</th><th>Cuándo</th><th>Descripción</th><th>Efecto en el programa</th></tr></thead><tbody>${vis.map(it => `<tr ${it.centrales[0] ? `data-central="${esc(it.centrales[0])}"` : ""}><td class="faint">${esc(it.tipo)}</td><td class="mono"><b>${esc(it.equipo)}</b></td><td>${esc(it.ubic)}${it.centrales.some(isStk) ? ' <span class="pill media">STK</span>' : ""}</td><td style="min-width:150px">${esc(winText(it.ventanas))}</td><td>${esc(it.desc)}</td><td class="faint">${esc(effectOf(it) || linkCmg(it))}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">${its.length ? "Ninguna intervención con estos filtros." : "Sin tabla de intervenciones."}</p>`;
  const R = [], el0 = $("tblRestric"); if (!el0) return;
  el0.innerHTML = R.length ? R.map(r => { const stk = r.filas.filter(f => /STATKRAFT/i.test(f[0])), tot = r.filas.reduce((s, f) => s + (f[5] || 0), 0);
    return `<div style="margin-bottom:12px"><div class="t"><b>${esc(r.motivo)}</b></div><div class="unit">${esc(r.horario)} · rechazo total ${fmt(tot, 2)} MW${stk.length ? ` · Statkraft: ${fmt(stk.reduce((s, f) => s + f[5], 0), 2)} MW en ${stk.length} cliente(s)` : ""} · ${esc(r.fuente)}</div>
      <div class="tblwrap"><table><thead><tr><th>Suministrador</th><th>Cliente</th><th>Subestación</th><th class="n">kV</th><th class="n">Demanda MW</th><th class="n">Rechazo MW</th></tr></thead><tbody>${r.filas.map(f => `<tr class="${/STATKRAFT/i.test(f[0]) ? "sel" : ""}"><td>${esc(f[0])}${/STATKRAFT/i.test(f[0]) ? ' <span class="pill media">STK</span>' : ""}</td><td>${esc(f[1])}</td><td>${esc(f[2])}</td><td class="n">${f[3] ?? "—"}</td><td class="n">${fmt(f[4], 2)}</td><td class="n">${fmt(f[5], 2)}</td></tr>`).join("")}</tbody></table></div></div>`; }).join("") : `<p class="muted">Sin datos de restricción de suministros.</p>`;
}
function linkCmg(it) { // relaciona una línea con los episodios de CMg
  const ev = cmgEvents(ST.cur).filter(e => e.auto?.texto?.includes(it.equipo)); return ev.length ? `Explica: ${ev.map(e => e.titulo).join(" · ")}` : "";
}
async function extractTables() {
  if (!SAMPLE) { toast("La IA funciona al abrir la página dentro de Claude.", true); return; }
  const W = ST.wk; if (!W.ipsi && !W.ipso) { toast("Carga primero el IPSI o el IPSO.", true); return; }
  const btn = $("extractBtn"); btn.disabled = true; btn.textContent = "Extrayendo…";
  const q = `Del IPSI y del IPSO del COES (semana ${ST.cur.semana}, ${ST.cur.inicio} a ${ST.cur.fin}) extrae dos tablas. Las tablas del PDF vienen con columnas desalineadas: reconstruye cada fila con cuidado y no inventes datos.
1) "intervenciones": cada intervención de la sección 3.1 del IPSI y las conexiones provisionales. Campos: tipo ("Generación", "Transmisión" o "Conexión provisional"), empresa, equipo, ubic, desc, ventanas (lista de [inicio, fin] en formato AAAA-MM-DDTHH:MM; 24:00 = 23:59 del mismo día; 00:00 del día siguiente a la semana = ${ST.cur.fin}T23:59), centrales (para generación, el nombre de la central tal como aparece en esta lista: ${Object.keys(ST.cur.centrales).slice(0, 400).join(", ")}; si no está, lista vacía).
2) "limites": cada fila de la sección "Líneas que operan a su capacidad nominal y/o límite de transmisión" del IPSO. Campos: equipo, tramo, zona ("Norte", "Centro" o "Sur" según la ubicación en el SEIN), ventanas (igual formato), motivo, fuente ("IPSO 4.2").
Responde SOLO JSON: {"intervenciones":[...],"limites":[...]}

IPSI:
${(W.ipsi || "").slice(0, 60000)}

IPSO:
${(W.ipso || "").slice(0, 40000)}`;
  try { const r = await SAMPLE.json(q, {cache: false, modelTier: "complex"});
    if (Array.isArray(r?.intervenciones) && r.intervenciones.length) ST.ipsiItems[ST.key] = r.intervenciones.map(x => ({...x, centrales: x.centrales || [], ventanas: x.ventanas || []}));
    if (Array.isArray(r?.limites) && r.limites.length) ST.limites[ST.key] = r.limites;
    persistExtra(); toast(`Extraídas ${r?.intervenciones?.length || 0} intervenciones y ${r?.limites?.length || 0} líneas al límite.`);
  } catch (e) { toast(e.code === "not_granted" ? "Sin permiso para usar Claude." : `No se pudo extraer (${e.code || "error"}).`, true); }
  btn.disabled = false; btn.textContent = "Extraer tablas con IA"; renderAll();
}

document.addEventListener("click", e => { const r = e.target.closest(".irow[data-central]"); if (r) { ST.dropSel = r.dataset.central; renderIPSI(); } });

/* ============ Filtro Source en árbol (como el BI) ============ */
const TIPO_LBL = {"EÓLICA": "Eólica", "HIDROELÉCTRICA": "Hidroeléctrica", "SOLAR": "Solar", "TERMOELÉCTRICA": "Termoeléctrica"};
let _tree = null;
function SRC_TREE() {
  if (_tree) return _tree; const g = {};
  for (const m of MAESTRO.values()) { if (!m.tipo || !m.source) continue; const t = g[m.tipo] || (g[m.tipo] = {}); t[m.source] = (t[m.source] || 0) + 1; }
  const num = s => parseFloat(s) || 0;
  return _tree = Object.keys(g).sort().map(tipo => ({tipo, srcs: Object.entries(g[tipo]).sort((a, b) => num(a[0]) - num(b[0])).map(([s, n]) => ({s, n}))}));
}
function renderSrcTree() {
  const S = ST.F.src, T = SRC_TREE();
  const box = (attrs, on, mixed, label, n) => `<label class="ck"><input type="checkbox" ${attrs} ${on ? "checked" : ""} ${mixed ? 'data-mixed="1"' : ""}><span>${label}</span>${n != null ? `<em>${n}</em>` : ""}</label>`;
  $("fSourcePop").innerHTML = box("data-all", !S.size, false, "Todas") + T.map(g => {
    const k = g.srcs.filter(x => S.has(x.s)).length, open = !ST.srcClosed.has(g.tipo);
    return `<div class="grp"><button type="button" class="tw" data-tw="${esc(g.tipo)}" aria-expanded="${open}" aria-label="${open ? "Contraer" : "Expandir"} ${esc(g.tipo)}"><svg viewBox="0 0 12 12" width="11" height="11"><path d="${open ? "M2 8l4-4 4 4" : "M2 4l4 4 4-4"}" fill="none" stroke="currentColor" stroke-width="1.6"/></svg></button>${box(`data-tipo="${esc(g.tipo)}"`, k === g.srcs.length, k > 0 && k < g.srcs.length, esc(g.tipo), g.srcs.reduce((a, x) => a + x.n, 0))}</div>`
      + (open ? `<div class="subs">${g.srcs.map(x => box(`data-src="${esc(x.s)}"`, S.has(x.s), false, esc(x.s), x.n)).join("")}</div>` : "");
  }).join("");
  $("fSourcePop").querySelectorAll("input[data-mixed]").forEach(i => { i.indeterminate = true; });
  // texto del botón
  let txt = "Todas";
  if (S.size) { const parts = T.map(g => { const k = g.srcs.filter(x => S.has(x.s)); const n = TIPO_LBL[g.tipo] || g.tipo;
      return !k.length ? null : k.length === g.srcs.length ? n : k.length === 1 ? k[0].s : `${n} (${k.length}/${g.srcs.length})`; }).filter(Boolean);
    txt = parts.length <= 2 ? parts.join(" + ") : `${S.size} sources`; }
  $("fSourceTxt").textContent = txt; $("fSource").classList.toggle("on", S.size > 0);
}

/* ============ Comparar contra (PSO / PDO) ============ */
let _baseCache = {sig: "", v: null};
function baseSummary(mode) { // resumen del reprograma vigente frente a una base
  const keep = ST.baseMode; ST.baseMode = mode;
  try { const w = ST.cur, V = vigWeek(w), cov = idxs(w).filter(i => V.letter[i]);
    const dE = (sum(sumCen(V.cen, cov, pass)) - sum(sumCen(V.cenB, cov, pass))) * 0.5;
    const cm = a => { const v = pick(a, cov).filter(x => x != null); return v.length ? sum(v) / v.length : null; };
    const dC = cm(V.cmg["Santa Rosa"]) != null && cm(V.cmgB["Santa Rosa"]) != null ? cm(V.cmg["Santa Rosa"]) - cm(V.cmgB["Santa Rosa"]) : null;
    let nC = 0, top = null; for (const c of Object.keys({...V.cen, ...V.cenB})) { if (!pass(c)) continue; let m = 0; for (const i of cov) m += ((V.cen[c]?.[i] ?? V.cenB[c]?.[i] ?? 0) - (V.cenB[c]?.[i] || 0)) * 0.5; if (Math.abs(m) >= 50) nC++; if (!top || Math.abs(m) > Math.abs(top.m)) top = {c, m}; }
    const used = new Set(cov.map(i => V.baseOf[i]));
    return {dE, dC, nC, top, used, n: cov.length};
  } finally { ST.baseMode = keep; }
}
function renderBase() {
  const w = ST.cur, V = getVW(), dates = V.days.map(x => x.date).filter(d => !ST.F.days.size || ST.F.days.has(d));
  const sig = ST.key + JSON.stringify([...ST.F.days, ...ST.F.src, ST.F.central]) + VWsig;
  if (_baseCache.sig !== sig) _baseCache = {sig, v: {PSO: baseSummary("PSO"), PDO: baseSummary("PDO")}};
  const R = _baseCache.v, pdoDays = dates.filter(d => ST.pdo[d]), noPdo = dates.filter(d => !ST.pdo[d]);
  if (!dates.length) { $("baseFlow").innerHTML = `<span class="muted">No hay reprogramas en los días elegidos.</span>`; $("baseSum").textContent = ""; return; }
  const opt = (v, t, sub) => { const r = R[v], dis = v === "PDO" && !pdoDays.length, on = ST.baseMode === v && !dis || (v === "PSO" && ST.baseMode === "PDO" && !pdoDays.length);
    return `<button type="button" class="bopt ${on ? "on" : ""}" data-v="${v}" ${dis ? "disabled" : ""} aria-pressed="${on}"><b>${t}</b><span>${sub}</span>${dis ? `<i>no cargado</i>` : `<i class="${r.dE >= 0 ? "up" : "dn"}">${sgn(r.dE / 1000, 2)} GWh${r.dC != null ? ` · CMg ${sgn(r.dC, 1)}` : ""}</i>`}</button>`; };
  const day = rdoFocusDay(), lets = (V.days.find(x => x.date === day)?.R.letters) || [], Rd = day ? rdoDay(day) : null;
  const chain = lets.map((L, k) => { const s = Object.entries(Rd?.from || {}).filter(([, x]) => x === L).map(([s]) => +s); const t0 = s.length ? Math.min(...s) : null;
    return `<span class="lt ${k === lets.length - 1 ? "last" : ""}" title="Reprograma ${L}${t0 != null ? " rige desde " + w.times[t0].slice(11) : ""}">${L}${t0 != null ? `<small>${w.times[t0].slice(11)}</small>` : ""}</span>`; }).join("");
  const miss = day ? ST.listado.filter(r => r.fecha === day && !lets.includes(r.reprograma)).sort((a, b) => a.reprograma.localeCompare(b.reprograma)) : [];
  const chain2 = chain + miss.map(r => `<span class="lt miss" title="Publicado (${r.causa}) pero no cargado">${r.reprograma}<small>${r.hora}</small></span>`).join("");
  $("baseFlow").innerHTML = `<div class="bopts">${opt("PSO", "PSO", "Programa semanal (sáb–vie)")}${opt("PDO", "PDO", "Programa diario del día")}</div>
    <svg class="arr" viewBox="0 0 40 12" width="40" height="12" aria-hidden="true"><path d="M0 6h34M30 1l6 5-6 5" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
    <div class="chain"><span class="lbl">Reprogramas ${day ? dayLabel(day) : ""}</span><div>${chain2 || "—"}</div></div>`;
  const cur = pdoDays.length && ST.baseMode === "PDO" ? "PDO" : "PSO", r = R[cur], other = R[cur === "PDO" ? "PSO" : "PDO"];
  const fl = isFiltered() ? " (con el filtro de centrales)" : "";
  const chip = (k, v, cls = "") => `<span class="bchip ${cls}"><em>${k}</em>${v}</span>`;
  $("baseSum").innerHTML = chip(`Vigente vs ${cur}${isFiltered() ? " (filtro)" : ""}`, `${sgn(r.dE / 1000, 2)} GWh`, r.dE >= 0 ? "up" : "dn")
    + (r.dC != null ? chip("CMg Santa Rosa", `${sgn(r.dC, 1)} US$/MWh`) : "")
    + chip("Centrales > 50 MWh", r.nC) + (r.top && Math.abs(r.top.m) >= 50 ? chip("Mayor cambio", `${esc(r.top.c)} ${sgn(r.top.m, 0)} MWh`) : "")
    + (miss.length ? chip("Falta cargar", `Reprog ${miss.map(r => r.reprograma).join(", ")} · ${miss[0].hora}`, "warn") : "")
    + (pdoDays.length && noPdo.length && cur === "PDO" ? chip("Sin PDO", `${noPdo.map(dayLabel).join(", ")} → PSO`, "faint") : "");
}

/* ============ Justificación compacta de episodios de CMg ============ */
ST.ep = {}; ST.epArgs = {};
const ZCOL = z => css(z === "Norte" ? "--b-norte" : z === "Sur" ? "--b-sur" : "--b-centro");
const plugEpis = {id: "epis", afterDatasetsDraw(ch, a, opts) { try { const E = opts?.items || [], L = ch.data.labels; if (!E.length || !L) return;
  const c = ch.ctx, x = ch.scales.x, y = ch.scales.y, {top, bottom} = ch.chartArea; c.save();
  for (const e of E) { const i = L.indexOf(e.t); if (i < 0) continue; const px = x.getPixelForValue(i); let py = y.getPixelForValue(e.v); if (!(py >= top + 10)) py = top + 10; if (py > bottom - 10) py = bottom - 10;
    const r = e.sel ? 10 : 8; c.beginPath(); c.arc(px, py, r, 0, 2 * Math.PI); c.fillStyle = e.color; c.globalAlpha = e.dim ? 0.35 : 1; c.fill();
    c.lineWidth = 2; c.strokeStyle = css("--panel"); c.stroke(); c.fillStyle = "#fff"; c.font = `600 ${e.sel ? 11 : 10}px IBM Plex Sans, sans-serif`; c.textAlign = "center"; c.textBaseline = "middle"; c.fillText(String(e.n), px, py + 0.5); }
  c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugEpis); } catch (e) {} }
function epHelp(f, ai) { const r = epHelp0(f, ai); return /^(alta|media)$/.test(r.conf) && r.txt ? r : {txt: "", corto: "", conf: "", filas: null}; }
function epHelp0(f, ai) { const h = ai.get(f.id), a = f.auto;
  if (h) return {txt: h.explicacion, corto: (h.explicacion || "").split(/(?<=\.)\s/)[0].slice(0, 70), cita: h.cita, fuente: h.fuente, conf: (h.confianza || "").toLowerCase(), by: "IA"};
  return a ? {filas: a.filas, txt: a.texto, corto: a.corto || a.texto.slice(0, 60), cita: a.cita, fuente: a.fuente, conf: (a.conf || "").toLowerCase(), by: "cruce automático"} : {txt: "", corto: "Sin justificar", conf: "ninguna"}; }
function itemHtml(f, ai) { const h = epHelp(f, ai), c = h.conf;
  return `<div class="item ${c}"><div class="s"></div><div><div class="t">${esc(f.titulo)}</div><div class="k">${esc(f.dato)}</div>${h.txt ? `<p class="x">${esc(h.txt)}</p><div class="src2"><span class="pill ${c}">${esc(c === "ninguna" ? "sin sustento" : c)}</span>${esc(h.fuente || "")}</div>` : ""}</div></div>`; }
function renderEpis(boxId, tab, chartId, list, sub = "", extra = []) {
  ST.epArgs[boxId] = [tab, chartId, list, sub, extra];
  const el = $(boxId); if (!el) return; const S = ST.ep[boxId] || (ST.ep[boxId] = {sel: null, zona: ""}), ai = aiMap(tab), st = aiState(tab);
  const eps = list.filter(f => f.ep).sort((p, q) => p.ep.t0.localeCompare(q.ep.t0)).map((f, k) => ({f, n: k + 1, h: epHelp(f, ai)}));
  if (S.sel && !eps.some(e => e.f.id === S.sel)) S.sel = null;
  const zonas = [...new Set(eps.map(e => e.f.ep.zona))], vis = eps.filter(e => !S.zona || e.f.ep.zona === S.zona);
  const chips = zonas.length > 1 && eps.length > zonas.length ? `<div class="epf"><button type="button" data-epz="" data-box="${boxId}" aria-pressed="${!S.zona}">Todas ${eps.length}</button>${zonas.map(z => `<button type="button" data-epz="${esc(z)}" data-box="${boxId}" aria-pressed="${S.zona === z}"><i style="background:${ZCOL(z)}"></i>${esc(z)} ${eps.filter(e => e.f.ep.zona === z).length}</button>`).join("")}</div>` : "";
  const row = e => { const p = e.f.ep, open = S.sel === e.f.id, c = e.h.conf, col = ZCOL(p.zona), rng = `${tShort(p.t0)}–${p.t1.slice(11)}${p.t0.slice(0, 10) !== p.t1.slice(0, 10) ? " (" + dayLabel(p.t1.slice(0, 10)) + ")" : ""}`;
    return `<div class="ep ${open ? "open" : ""}"><button type="button" class="eph" data-ep="${e.f.id}" data-box="${boxId}" aria-expanded="${open}">
      <span class="n" style="background:${col}">${e.n}</span>
      <span class="w"><b>${esc(p.letter ? "Reprog " + p.letter : p.zona)}</b> <span class="k">${esc(rng)}</span></span>
      <span class="v" title="${p.spread != null ? (p.letter ? "Diferencia contra la base" : "Diferencia contra el Centro") : "Máximo de la semana"}">${fmt(p.max)}<small>${p.spread != null ? sgn(p.spread) : "máx"}</small></span>
      ${e.h.corto ? `<span class="c">${esc(e.h.corto)}</span>` : ""}</button>
      ${open ? `<div class="epx">${e.h.filas?.length ? `<dl>${e.h.filas.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join("")}</dl>` : `<p>${esc(e.h.txt)}</p>`}<div class="src2"><span class="faint">Fuente: ${esc(e.h.fuente || "")}</span><button type="button" class="lnk" data-mom="${p.tmax}">Ver el pico ${p.tmax.slice(11)} →</button></div></div>` : ""}</div>`; };
  el.innerHTML = `<div class="ephd"><span class="lbl">Justificación</span>${st.loading ? `<span class="thinking">Leyendo informes…</span>` : ""}</div>${sub ? `<div class="unit">${sub}</div>` : ""}${chips}`
    + (vis.length ? `<div class="eplist">${vis.map(row).join("")}</div>` : `<p class="muted">Sin episodios relevantes${S.zona ? " en " + esc(S.zona) : ""}.</p>`)

    + (() => { const ok = extra.filter(f => epHelp(f, ai).txt); return ok.length ? `<div class="eplist">${ok.map(f => whyRow(boxId, f, ai, S.xopen === f.id)).join("")}</div>` : ""; })();
  // marcadores numerados y ventana del episodio en el gráfico
  const ch = ST.charts[chartId]; if (!ch?.options?.plugins) return;
  const sel = eps.find(e => e.f.id === S.sel);
  ch.options.plugins.epis = {items: eps.map(e => ({t: e.f.ep.tmax, v: e.f.ep.max, n: e.n, color: ZCOL(e.f.ep.zona), sel: e === sel, dim: (S.zona && e.f.ep.zona !== S.zona) || (sel && e !== sel)}))};
  ch.options.plugins.bands = {items: sel ? [[sel.f.ep.t0, sel.f.ep.t1]] : []};
  try { ch.update("none"); } catch (e) {}
}
const reEpis = id => { const a = ST.epArgs[id]; if (a) renderEpis(id, ...a); };
document.addEventListener("click", e => {
  const T = e.target, h = T.closest(".eph[data-ep]"); if (h) { const S = ST.ep[h.dataset.box]; S.sel = S.sel === h.dataset.ep ? null : h.dataset.ep; reEpis(h.dataset.box); return; }
  const z = T.closest("[data-epz]"); if (z) { const S = ST.ep[z.dataset.box]; S.zona = z.dataset.epz; S.sel = null; reEpis(z.dataset.box); return; }
  const m = T.closest("[data-mom]"); if (m) setMoment(m.dataset.mom);
});
document.addEventListener("mouseover", e => { const h = e.target.closest?.(".eph[data-ep]"); if (!h || h === ST.epHover) return; ST.epHover = h;
  const [, chartId, list] = ST.epArgs[h.dataset.box] || [], f = list?.find(x => x.id === h.dataset.ep), ch = ST.charts[chartId]; if (!f || !ch?.options?.plugins) return;
  ch.options.plugins.bands = {items: [[f.ep.t0, f.ep.t1]]}; try { ch.update("none"); } catch (x) {} });
document.addEventListener("mouseout", e => { const h = e.target.closest?.(".eph[data-ep]"); if (!h || h.contains(e.relatedTarget)) return; ST.epHover = null;
  const S = ST.ep[h.dataset.box], [, chartId, list] = ST.epArgs[h.dataset.box] || [], ch = ST.charts[chartId]; if (!ch?.options?.plugins) return;
  const f = list?.find(x => x.id === S?.sel); ch.options.plugins.bands = {items: f ? [[f.ep.t0, f.ep.t1]] : []}; try { ch.update("none"); } catch (x) {} });

/* ============ Daily Redispatch Plan con el formato del Power BI ============ */
const MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const biTitle = t => { const [d, h] = t.split("T"), [Y, M, D] = d.split("-"); let [hh, mm] = h.split(":").map(Number); const pm = hh >= 12, h12 = hh % 12 || 12;
  return `${D}/${M}/${Y} ${String(h12).padStart(2, "0")}:${String(mm).padStart(2, "0")}:00 ${pm ? "p.m." : "a.m."}`; };
const plugLabels = {id: "dlabels", afterDatasetsDraw(ch, a, opts) { try { if (!opts?.on) return; const L = ch.data.labels, c = ch.ctx, used = [];
  c.save(); c.font = "600 11px Segoe UI, IBM Plex Sans, sans-serif"; c.textAlign = "center"; c.fillStyle = css("--fg-2");
  ch.data.datasets.forEach((ds, k) => { if (ds.dl === false || !ch.isDatasetVisible(k)) return; const meta = ch.getDatasetMeta(k), by = {};
    ds.data.forEach((v, i) => { if (v == null) return; const d = L[i].slice(0, 10); const b = by[d] || (by[d] = {mx: i, mn: i}); if (v > ds.data[b.mx]) b.mx = i; if (v < ds.data[b.mn]) b.mn = i; });
    for (const b of Object.values(by)) for (const [i, up] of [[b.mx, true], [b.mn, false]]) { if (!up && ds.data[b.mx] - ds.data[b.mn] < 8) continue;
      const p = meta.data[i]; if (!p) continue; const x = p.x, y = p.y + (up ? -7 : 14);
      if (y < ch.chartArea.top + 8 || used.some(u => Math.abs(u[0] - x) < 22 && Math.abs(u[1] - y) < 12)) continue; used.push([x, y]); c.fillText(fmt(ds.data[i]), x, y); } });
  c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugLabels); } catch (e) {} }
function biOpts(L, {y = "", min, max, step, y2 = false, labels = false, area = false} = {}) {
  const o = tsOpts(L, {y}), fg2 = css("--fg-2"), grid = css("--fg-3");
  o.plugins.marks = {items: []}; o.plugins.dlabels = {on: labels};
  o.elements = {point: {radius: 0, hoverRadius: 4}, line: {borderWidth: 2, tension: 0}};
  o.scales.x = {grid: {display: true, drawTicks: false, color: c => (L[c.index] || "").endsWith("T00:30") ? grid : "transparent", lineWidth: 1}, border: {display: false, dash: [2, 3]},
    ticks: {color: fg2, autoSkip: false, maxRotation: 0, padding: 6, font: {size: 12}, callback: (v, i) => (L[i] || "").endsWith("T12:00") ? `${+L[i].slice(8, 10)} ${MES[+L[i].slice(5, 7) - 1]}` : ""}};
  o.scales.y = {min, max, grid: {color: grid, lineWidth: 1}, border: {display: false, dash: [2, 3]}, ticks: {color: fg2, stepSize: step, callback: v => step && v % step ? "" : fmt(v), font: {size: 12}}, title: {display: !!y, text: y, color: css("--fg"), font: {size: 13}}};
  if (y2) o.scales.y2 = {position: "right", min, max, grid: {display: false}, border: {display: false}, ticks: {color: fg2, stepSize: step, callback: v => step && v % step ? "" : `${fmt(v / 1000)} mil`, font: {size: 11}}};
  o.plugins.tooltip.callbacks.title = it => it.length ? biTitle(L[it[0].dataIndex]) : "";
  o.plugins.tooltip.callbacks.label = c => c.parsed.y == null ? null : ` ${c.dataset.label}   ${fmt(c.parsed.y, 2)}`;
  o.plugins.tooltip.callbacks.afterBody = it => { if (!it.length) return []; const i = ST.cur.times.indexOf(L[it[0].dataIndex]), Lt = getVW().letter[i]; return Lt ? [`Reprograma   ${Lt}`, "Software   YUPANA"] : []; };
  return o;
}
const niceUp = (v, s) => Math.ceil(v / s) * s, niceDn = (v, s) => Math.floor(v / s) * s;
const BI = () => ({pdoL: css("--bi-pdo"), week: css("--bi-week"), cen: css("--bi-center"), sur: css("--bi-south"), nor: css("--bi-north"), rdo: css("--bi-rdo"), pso: css("--bi-pso"), rep: css("--bi-reps"), rer: css("--bi-rerdo"), drp: css("--bi-drpso"), drr: css("--bi-drrdo"), pdo: css("--fg-3")});
function renderRdoCharts(w, V, I, L) {
  const C = BI(), pdoOn = ST.baseMode === "PDO" && I.some(i => V.baseOf[i] === "PDO" && V.letter[i]);
  const pdoOnly = a => I.map(i => V.baseOf[i] === "PDO" && V.letter[i] ? a[i] : null);
  const DOT = {borderDash: [2, 3], borderWidth: 3, pointRadius: 0, borderCapStyle: "round"};
  // 1) CMg: igual que el BI (PSO semanal en Santa Rosa + RDO por barra)
  const faint = c => c + "55";
  const ds = pdoOn ? [{label: "PDO Spot Price Center", data: pdoOnly(V.cmgB["Santa Rosa"]), borderColor: C.pdoL, backgroundColor: C.pdoL, borderWidth: 2.5},
      {label: "Weekly Avg Spot Price", data: pick(w.cmg["Santa Rosa"], I), borderColor: faint(C.week), backgroundColor: faint(C.week), borderWidth: 1.5, dl: false}]
    : [{label: "Weekly Avg Spot Price", data: pick(w.cmg["Santa Rosa"], I), borderColor: C.week, backgroundColor: C.week, borderWidth: 2}];
  ds.push(
{label: "RDO Spot Price Center", data: pick(V.cmg["Santa Rosa"], I), borderColor: C.cen, backgroundColor: C.cen, ...DOT},
    {label: "RDO Spot Price South", data: pick(V.cmg["Socabaya"], I), borderColor: C.sur, backgroundColor: C.sur, ...DOT},
    {label: "RDO Spot Price North", data: pick(V.cmg["Trujillo"], I), borderColor: C.nor, backgroundColor: C.nor, ...DOT});
  if (typeof cmgEjWeek === "function" && ST.tc > 0) { const ce = pick(cmgEjWeek(w, "sr"), I); if (ce.some(v => v != null)) ds.push({label: "Ejecutado Center (IEOD)", data: ce, borderColor: css("--ejec"), backgroundColor: css("--ejec"), borderWidth: 2.5, dl: false}); }
  const vals = ds.flatMap(d => d.data).filter(v => v != null), cmax = Math.max(...vals), st = cmax > 600 ? 200 : 100;
  mkChart("cRCmg", {type: "line", data: {labels: L, datasets: ds}, options: biOpts(L, {y: "CMg (USD/MWh)", min: 0, max: niceUp(cmax * 1.08, st), step: st, labels: true})});
  legend("lgRCmg", ds.map(d => ({label: d.label, color: d.borderColor})));
  renderByBarra(w, V, I, C);
  // 2) Potencia: RDO/PSO MW (eje izq.) y reserva eficiente (eje der., misma escala)
  const gP = sumCen(w.centrales, I, pass), gR = sumCen(V.cen, I, pass), reP = pick(w.re, I), reR = pick(V.re, I);
  const gmax = Math.round(Math.max(...[...gP, ...gR].filter(v => v != null)) * 1.04);
  const gd = [{label: "RDO MW", data: gR, borderColor: C.rdo, backgroundColor: C.rdo, yAxisID: "y"}, {label: "PSO MW", data: gP, borderColor: C.pso, backgroundColor: C.pso, yAxisID: "y", borderWidth: 1.5},
    {label: "RES EF PSO MW", data: reP, borderColor: C.rep, backgroundColor: C.rep, yAxisID: "y2", ...DOT}, {label: "RES EF RDO MW", data: reR, borderColor: C.rer, backgroundColor: C.rer, yAxisID: "y2", borderWidth: 3}];
  if (pdoOn) gd.push({label: "PDO MW", data: pdoOnly(sumCen(V.cenB, w.times.map((_, i) => i), pass)), borderColor: C.pdo, backgroundColor: C.pdo, borderDash: [6, 4], borderWidth: 1.5, yAxisID: "y"});
  if (typeof ejecWeek === "function") { const ej = pick(ejecWeek(w), I); if (ej.some(v => v != null)) gd.splice(2, 0, {label: "Ejecutado MW (IEOD)", data: ej, borderColor: css("--ejec"), backgroundColor: css("--ejec"), borderWidth: 2.5, yAxisID: "y"}); }
  mkChart("cRGen", {type: "line", data: {labels: L, datasets: gd}, options: biOpts(L, {y: "Potencia (MW)", min: 0, max: gmax, step: 2000, y2: true})});
  legend("lgRGen", gd.map(d => ({label: d.label, color: d.borderColor})));
  // 3) Residual Demand: áreas, eje desde 2 500
  const rP = pick(derive(w).resid, I), rR = pick(V.resid, I), rv = [...rP, ...rR].filter(v => v != null);
  const rd = [{label: "RDO DR", data: rR, borderColor: C.drr, backgroundColor: C.drr, fill: "origin", borderWidth: 1.5, order: 1}, {label: "PSO DR", data: rP, borderColor: C.drp, backgroundColor: C.drp + "99", fill: "origin", borderWidth: 1.5, order: 2}];
  if (pdoOn) rd.push({label: "PDO DR", data: pdoOnly(V.residB), borderColor: C.pdo, borderDash: [6, 4], borderWidth: 1.5, fill: false, order: 0});
  mkChart("cRRes", {type: "line", data: {labels: L, datasets: rd}, options: biOpts(L, {y: "Potencia (MW)", min: niceDn(Math.min(...rv) * 0.97, 500), max: niceUp(Math.max(...rv) * 1.02, 500), step: 500})});
  legend("lgRRes", [rd[1], rd[0], ...rd.slice(2)].map(d => ({label: d.label, color: d.borderColor})));
}

/* Efecto del reprograma por barra: cada barra contra su propia base */
function renderByBarra(w, V, I, C) {
  const cov = I.filter(i => V.letter[i]), L = pick(w.times, cov), pdo = ST.baseMode === "PDO";
  const bases = new Set(cov.map(i => V.baseOf[i])), bl = bases.size === 1 ? [...bases][0] : "PDO/PSO";
  $("bybSub").textContent = cov.length ? `${dayList([...new Set(L.map(t => t.slice(0, 10)))])} · base: ${bl}` : "";
  const B = [{k: "Santa Rosa", z: "Centro", c: C.cen}, {k: "Trujillo", z: "Norte", c: C.nor}, {k: "Socabaya", z: "Sur", c: C.sur}];
  const st = B.map(b => { const d = cov.map(i => V.cmg[b.k][i] != null && V.cmgB[b.k][i] != null ? V.cmg[b.k][i] - V.cmgB[b.k][i] : null), v = d.filter(x => x != null);
    let mx = 0, mn = 0; d.forEach((x, j) => { if (x != null && x > d[mx]) mx = j; if (x != null && x < d[mn]) mn = j; });
    return {d, avg: v.length ? sum(v) / v.length : null, mx, mn}; });
  const ord = st.map((x, k) => [Math.abs(x.avg || 0), k]).sort((a, b) => b[0] - a[0]), [a1, k1] = ord[0], a2 = ord[1][0];
  const top = cov.length && a1 >= 1 && a1 - a2 >= 10 && a1 >= 1.2 * a2 ? k1 : -1, avgs = st.map(x => x.avg).filter(v => v != null);
  const dc = k => st[k].avg != null && st[0].avg != null ? st[k].avg - st[0].avg : null;
  $("bybNote").innerHTML = !cov.length ? "" : top >= 0 ? `<b>Mayor efecto en el ${B[top].z}</b>: ${sgn(st[top].avg, 1)} US$/MWh en promedio, frente a ${sgn(st[ord[1][1]].avg, 1)} en la siguiente barra.`
    : `<b>Efecto similar en las tres barras</b> (${sgn(Math.min(...avgs), 0)} a ${sgn(Math.max(...avgs), 0)} US$/MWh): cambio general del sistema, no de una zona.`;
  $("bybNote").innerHTML += cov.length ? ` <span class="faint">Desacople respecto al Centro: Norte ${sgn(dc(1), 1)} · Sur ${sgn(dc(2), 1)} US$/MWh${Math.max(Math.abs(dc(1) || 0), Math.abs(dc(2) || 0)) >= 20 ? " — el reprograma cambió la transferencia entre zonas" : ""}.</span>` : "";
  B.forEach((b, k) => { const s = st[k], cell = $("bcell" + k); cell.classList.toggle("top", k === top);
    if (!cov.length) { $("bst" + k).innerHTML = `<span class="muted">Sin reprogramas en los días elegidos.</span>`; return; }
    const wk = {label: `Weekly Spot Price ${b.z}`, data: pick(w.cmg[b.k], cov), borderColor: C.week, backgroundColor: C.week, borderWidth: 2, dl: false};
    const ds = pdo ? [{label: `PDO ${b.z}`, data: cov.map(i => V.baseOf[i] === "PDO" ? V.cmgB[b.k][i] : null), borderColor: C.pdoL, backgroundColor: C.pdoL, borderWidth: 2.2, dl: false}, {...wk, borderColor: C.week + "66", backgroundColor: C.week + "66", borderWidth: 1.5}] : [wk];
    ds.push({label: `RDO ${b.z}`, data: pick(V.cmg[b.k], cov), borderColor: b.c, backgroundColor: b.c, borderDash: [2, 3], borderWidth: 3, borderCapStyle: "round", dl: false});
    const o = biOpts(L, {y: "US$/MWh", min: 0}); o.plugins.letters = {items: cov.map(i => V.letter[i]), labels: L}; o.scales.y.title.display = false; o.scales.x.ticks.callback = (v, i) => (L[i] || "").slice(11, 16).match(/^(00:30|06:00|12:00|18:00)$/) ? (L[i].slice(11) === "00:30" ? `${+L[i].slice(8, 10)} ${MES[+L[i].slice(5, 7) - 1]}` : L[i].slice(11, 13) + "h") : "";
    mkChart("cRB" + k, {type: "line", data: {labels: L, datasets: ds}, options: o});
    legend("lgRB" + k, ds.map(d => ({label: d.label, color: d.borderColor})));
    const t = j => L[j] ? `${dayLabel(L[j].slice(0, 10))} ${L[j].slice(11)}` : "";
    $("bst" + k).innerHTML = `<div class="bv ${s.avg > 0.5 ? "up" : s.avg < -0.5 ? "dn" : ""}"><span class="lbl">Δ medio</span><b>${sgn(s.avg, 1)}</b><small>${k && dc(k) != null ? `vs Centro ${sgn(dc(k), 1)}` : "US$/MWh"}</small></div>
      <div class="bv"><span class="lbl">Mayor alza</span>${(s.d[s.mx] || 0) >= 0.5 ? `<b class="up">${sgn(s.d[s.mx], 0)}</b><small>${t(s.mx)}</small>` : `<b>—</b><small>sin alzas</small>`}</div>
      <div class="bv"><span class="lbl">Mayor baja</span>${(s.d[s.mn] || 0) <= -0.5 ? `<b class="dn">${sgn(s.d[s.mn], 0)}</b><small>${t(s.mn)}</small>` : `<b>—</b><small>sin bajas</small>`}</div>`; });
}

/* Líneas de base coherentes con «¿Contra qué programa se mide?»: PSO continuo; con PDO, PDO continuo en sus días y PSO tenue */
function baseLines(w, V, I, sumOf, nm, col) {
  const pso = sumOf(w.centrales), pdoOn = ST.baseMode === "PDO" && I.some(i => V.baseOf[i] === "PDO" && V.letter[i]);
  if (!pdoOn) return [{label: `${nm} PSO`, data: pick(pso, I), borderColor: col, backgroundColor: col, borderWidth: 2, fill: false, stack: "b1", order: 1, dl: false}];
  const pdo = sumOf(V.cenB), fa = /^#[0-9a-f]{6}$/i.test(col) ? col + "55" : col;
  return [{label: `${nm} PDO`, data: I.map(i => V.baseOf[i] === "PDO" && V.letter[i] ? pdo[i] : null), borderColor: col, backgroundColor: col, borderWidth: 2.5, fill: false, stack: "b1", order: 1, dl: false},
    {label: `${nm} PSO`, data: pick(pso, I), borderColor: fa, backgroundColor: fa, borderWidth: 1.5, fill: false, stack: "b2", order: 1, dl: false}];
}

/* Franjas por reprograma: separa dónde empieza cada letra */
const plugLetters = {id: "letters", beforeDatasetsDraw(ch, a, opts) { try { const T = opts?.items; if (!T?.length) return; const x = ch.scales.x, {top, bottom} = ch.chartArea, c = ch.ctx, L = opts.labels;
  const segs = []; T.forEach((t, i) => { const prevDay = i && L[i].slice(0, 10) !== L[i - 1].slice(0, 10); if (!segs.length || t !== segs[segs.length - 1].t || prevDay) segs.push({t, i0: i, i1: i}); else segs[segs.length - 1].i1 = i; });
  c.save(); segs.forEach((g, k) => { const p0 = x.getPixelForValue(g.i0) - (g.i0 ? (x.getPixelForValue(g.i0) - x.getPixelForValue(g.i0 - 1)) / 2 : 0), p1 = x.getPixelForValue(g.i1) + (g.i1 < T.length - 1 ? (x.getPixelForValue(g.i1 + 1) - x.getPixelForValue(g.i1)) / 2 : 0);
    if (k % 2) { c.fillStyle = css("--panel-2"); c.fillRect(p0, top, p1 - p0, bottom - top); }
    if (k) { c.strokeStyle = css("--warn"); c.lineWidth = 1.2; c.setLineDash([3, 3]); c.beginPath(); c.moveTo(p0, top); c.lineTo(p0, bottom); c.stroke(); c.setLineDash([]); }
    if (p1 - p0 > 18 && g.t) { c.fillStyle = css("--warn"); c.font = "700 11px IBM Plex Sans, sans-serif"; c.textAlign = "left"; c.textBaseline = "top"; c.fillText(g.t + " " + L[g.i0].slice(11), p0 + 4, top + 3); } });
  c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugLetters); } catch (e) {} }

/* ============ Justificación compacta para todos los paneles ============ */
ST.wopen = {}; ST.wmore = {}; ST.whyArgs = {}; ST.lgOff = {};
function whyRow(boxId, f, ai, open) {
  const h = epHelp(f, ai), c = h.conf || "ninguna";
  return `<div class="ep ${open ? "open" : ""}"><button type="button" class="eph gen" data-wf="${esc(f.id)}" data-box="${boxId}" aria-expanded="${open}">
    <span class="w"><b>${esc(f.titulo)}</b></span><svg class="chev" viewBox="0 0 12 12" width="11" height="11" aria-hidden="true"><path d="M2 4l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>
    <span class="c">${esc(h.txt ? h.corto : f.dato)}</span></button>
    ${open ? `<div class="epx">${h.txt ? `<div class="k">${esc(f.dato)}</div>` : ""}${h.filas?.length ? `<dl>${h.filas.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${esc(v)}</dd>`).join("")}</dl>` : h.txt ? `<p>${esc(h.txt)}</p>` : `<p class="faint">Sin justificación todavía: usa «Justificar con IA».</p>`}
      ${h.txt ? `<div class="src2"><span class="faint">Fuente: ${esc(h.fuente || "")}</span>${f.t ? `<button type="button" class="lnk" data-mom="${f.t}">Ver ${f.t.slice(11)} →</button>` : ""}</div>` : ""}</div>` : ""}</div>`;
}
function whyHead(list, ai, extra = "") {
  return `<div class="ephd"><span class="lbl">Justificación</span>${extra}</div>`;
}
function renderWhy(boxId, tab, list, sub = "") {
  ST.whyArgs[boxId] = [tab, list, sub];
  const el = $(boxId); if (!el) return; const ai = aiMap(tab), st = aiState(tab), op = ST.wopen[boxId];
  el.innerHTML = whyHead(list, ai, st.loading ? `<span class="thinking">Leyendo…</span>` : "") + (sub ? `<div class="unit">${sub}</div>` : "")
    + (() => { const ok = list.filter(f => epHelp(f, ai).txt); return ok.length ? `<div class="eplist">${ok.map(f => whyRow(boxId, f, ai, op === f.id)).join("")}</div>` : ""; })();
  const any = list.some(f => epHelp(f, ai).txt) || st.loading; el.hidden = !any; el.closest(".viz")?.classList.toggle("nowhy", !any);
}
function itemHtml(f, ai) { return whyRow("_x", f, ai, false); }
const reWhy = id => { const a = ST.whyArgs[id]; if (a) renderWhy(id, ...a); };
document.addEventListener("click", e => { const mo = e.target.closest("[data-wmore]"); if (mo) { const bx = mo.dataset.wmore; ST.wmore[bx] = !ST.wmore[bx]; ST.whyArgs[bx] ? reWhy(bx) : reEpis(bx); return; } const h = e.target.closest(".eph.gen"); if (!h) return; const b = h.dataset.box, id = h.dataset.wf;
  ST.wopen[b] = ST.wopen[b] === id ? null : id; if (ST.whyArgs[b]) reWhy(b); else if (ST.epArgs[b]) { ST.ep[b] = ST.ep[b] || {}; ST.ep[b].xopen = ST.wopen[b]; reEpis(b); } });

/* ============ Leyendas que ocultan/muestran series y gráficos animados ============ */
function legend(id, items) {
  const chId = id.replace(/^lg/, "c"), ch = ST.charts[chId], ds = ch?.data?.datasets || [], off = ST.lgOff[chId] || (ST.lgOff[chId] = new Set());
  const idx = x => { const k = ds.findIndex(d => d.label === x.label); return k >= 0 ? k : ds.findIndex(d => x.label.startsWith(d.label + " ") || x.label.startsWith(d.label + ":") || x.label.startsWith(d.label + " (")); };
  $(id).innerHTML = items.map(x => { const k = idx(x), hid = k >= 0 && off.has(ds[k].label);
    if (k >= 0 && hid) { try { ch.setDatasetVisibility(k, false); } catch (e) {} }
    return `<span ${k >= 0 ? `role="button" tabindex="0" data-lgc="${chId}" data-lgk="${k}" class="lgi ${hid ? "off" : ""}" title="Clic para ${hid ? "mostrar" : "ocultar"}"` : ""}><i class="${x.line ? "ln" : ""}" style="background:${x.dash ? "repeating-linear-gradient(90deg," + x.color + " 0 4px,transparent 4px 7px)" : x.color}"></i>${esc(x.label)}</span>`; }).join("");
  if (off.size) { try { ch.update("none"); } catch (e) {} }
}
document.addEventListener("click", e => { const s = e.target.closest("[data-lgc]"); if (!s) return; const ch = ST.charts[s.dataset.lgc], k = +s.dataset.lgk; if (!ch) return;
  const off = ST.lgOff[s.dataset.lgc], lab = ch.data.datasets[k].label, vis = !off.has(lab);
  vis ? off.add(lab) : off.delete(lab); try { ch.setDatasetVisibility(k, !vis); ch.update(); } catch (x) {} s.classList.toggle("off", vis); s.title = `Clic para ${vis ? "mostrar" : "ocultar"}`; });
document.addEventListener("keydown", e => { if ((e.key === "Enter" || e.key === " ") && e.target.matches?.("[data-lgc]")) { e.preventDefault(); e.target.click(); } });
function mkChart(id, cfg) {
  if (ST.charts[id]) ST.charts[id].destroy();
  if (cfg.options && !matchMedia("(prefers-reduced-motion: reduce)").matches) cfg.options.animation = {duration: 380, easing: "easeOutCubic"};
  ST.charts[id] = new Chart($(id), cfg);
}

/* ============ Centro de control: reloj, "ahora", modo sala ============ */
const pad2 = n => String(n).padStart(2, "0");
function nowLabel(d = new Date()) { // media hora vigente en el formato del PSO (00:30 … 23:59)
  let m = d.getHours() * 60 + d.getMinutes(), s = Math.ceil((m + 0.001) / 30) * 30; const iso = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  return s >= 1440 ? `${iso}T23:59` : `${iso}T${pad2(Math.floor(s / 60))}:${pad2(s % 60)}`;
}
const nowIdx = w => w ? w.times.indexOf(nowLabel()) : -1;
const plugNow = {id: "nowline", afterDatasetsDraw(ch) { try { const L = ch.data.labels, i = L ? L.indexOf(nowLabel()) : -1; if (i < 0) return;
  const x = ch.scales.x.getPixelForValue(i), {top, bottom} = ch.chartArea, c = ch.ctx, col = css("--live");
  c.save(); c.strokeStyle = col; c.lineWidth = 1.5; c.setLineDash([4, 3]); c.beginPath(); c.moveTo(x, top + 14); c.lineTo(x, bottom); c.stroke(); c.setLineDash([]);
  c.fillStyle = col; c.font = "600 10px IBM Plex Sans, sans-serif"; const t = "AHORA", tw = c.measureText(t).width + 10;
  c.beginPath(); c.roundRect ? c.roundRect(x - tw / 2, top, tw, 14, 7) : c.rect(x - tw / 2, top, tw, 14); c.fill(); c.fillStyle = "#fff"; c.textAlign = "center"; c.textBaseline = "middle"; c.fillText(t, x, top + 7.5); c.restore(); } catch (e) {} }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugNow); } catch (e) {} }

function tickClock() { const d = new Date(); $("clkT").textContent = `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; $("clkD").textContent = `${DOW[d.getDay()]} ${d.getDate()} ${MES[d.getMonth()]}`; }
let nowSlot = "", tickI = 0;
function renderNow() {
  const el = $("nowBar"), w = ST.cur; if (!el || !w) return;
  if (ST.tab === "chat") { el.hidden = true; return; }
  const i = nowIdx(w);
  if (i < 0) { const cur = sortedKeys().find(k => WEEKS.get(k).w.times.includes(nowLabel()));
    el.hidden = false; el.className = "nowbar off"; el.innerHTML = `<span class="live off"><i></i>Semana ${w.semana}</span><span class="muted">Estás viendo una semana que no incluye la hora actual.</span>${cur ? `<button type="button" class="lnk" data-wk="${cur}">Ir a la semana actual →</button>` : ""}`; return; }
  const V = getVW(), Lt = V.letter[i], day = w.times[i].slice(0, 10), lis = ST.listado.filter(x => x.fecha === day).sort((a, b) => a.reprograma.localeCompare(b.reprograma));
  const pub = lis[lis.length - 1], miss = pub && pub.reprograma !== Lt && pub.hora <= w.times[i].slice(11);
  const cm = b => Lt && V.cmg[b][i] != null ? V.cmg[b][i] : w.cmg[b][i], cmN = b => { const j = Math.min(i + 2, w.times.length - 1); return Lt && V.cmg[b][j] != null ? V.cmg[b][j] : w.cmg[b][j]; };
  const cm2 = j => Lt && V.letter[j] && V.cmg["Santa Rosa"][j] != null ? V.cmg["Santa Rosa"][j] : w.cmg["Santa Rosa"][j];
  const arrow = (a, b) => b == null || a == null || Math.abs(b - a) < 0.5 ? "" : b > a ? `<em class="up" title="Sube en la próxima hora">▲</em>` : `<em class="dn" title="Baja en la próxima hora">▼</em>`;
  const tipo = {}; if (Lt) for (const [c, a] of Object.entries(V.cen)) { const t = tipoOf(c); if (t && a[i] != null) tipo[t] = (tipo[t] || 0) + a[i]; }
  else { const dv = derive(w); TIPOS.forEach(t => tipo[t] = dv.byTipo[t]?.[i] || 0); }
  const tot = Object.values(tipo).reduce((s, v) => s + v, 0) || 1, dem = w.tot["Demanda YUPANA"]?.[i];
  const bar = TIPOS.filter(t => tipo[t] > 0).map(t => `<i style="width:${(tipo[t] / tot * 100).toFixed(1)}%;background:${css(TV[t])}" title="${t[0] + t.slice(1).toLowerCase()} ${fmt(tipo[t])} MW"></i>`).join("");
  el.hidden = false; el.className = "nowbar";
  el.innerHTML = `<span class="live"><i></i>Ahora <b class="mono">${w.times[i].slice(11)}</b></span>
    <span class="nb map">${typeof seinMap === "function" ? seinMap(w, V, i, Lt) : ""}</span>
    <span class="nb"><span class="lbl">Programa vigente</span><b>${Lt ? `Reprog ${Lt}` : "PSO"}</b>${miss ? `<small class="neg" title="${esc(pub.causa)}">${pub.reprograma} publicado ${pub.hora}, sin cargar</small>` : `<small>${Lt ? "desde " + (V.marks.filter(m => m.t <= w.times[i]).pop()?.t.slice(11) || "") : "sin reprograma"}</small>`}</span>
    ${BARRAS.map(b => `<span class="nb cm"><span class="lbl"><i style="background:${css(b.v)}"></i>CMg ${b.z}</span><b class="mono">${fmt(cm(b.k), 1)}${arrow(cm(b.k), cmN(b.k))}</b><small>US$/MWh</small></span>`).join("")}
    <span class="nb"><span class="lbl">Demanda prog.</span><b class="mono">${fmt(dem)}</b><small>MW</small></span>
    <span class="nb mix"><span class="lbl">Despacho por tipo</span><span class="stack">${bar}</span><span class="mixl">${TIPOS.filter(t => tipo[t] > 0).map(t => `<span><i style="background:${css(TV[t])}"></i>${t[0] + t.slice(1, 4).toLowerCase()}. <b>${fmt(100 * tipo[t] / tot)}%</b></span>`).join("")}</span></span>
    <span class="nb"><span class="lbl">Reserva eficiente</span><b class="mono">${fmt(Lt && V.re[i] != null ? V.re[i] : w.re?.[i])}</b><small>MW</small></span>
    ${(() => { const d0 = w.times.findIndex(t => t.startsWith(day)), J = [...Array(48).keys()].map(k => d0 + k), v = J.map(j => cm2(j)), k = v.indexOf(Math.max(...v.filter(x => x != null))); return k < 0 ? "" : `<span class="nb"><span class="lbl">Pico CMg Centro hoy</span><b class="mono">${fmt(v[k], 1)}</b><small>${w.times[J[k]].slice(11)} · ${J[k] < i ? "ya pasó" : J[k] === i ? "ahora" : "por venir"}</small></span>`; })()}
    ${lis.length ? `<span class="nb tick"><span class="lbl">Último reprograma</span><span class="tk" id="tk"></span></span>` : ""}`;
  if (lis.length) rotTick(lis);
}
let tkH = null;
function rotTick(lis) { const L = [...lis].reverse(); let k = 0; const show = () => { const el = $("tk"); if (!el) return; const r = L[k % L.length]; el.innerHTML = `<b>${r.reprograma} · ${r.hora}</b> ${esc(r.causa)}`; el.classList.remove("in"); void el.offsetWidth; el.classList.add("in"); k++; };
  clearInterval(tkH); show(); if (L.length > 1) tkH = setInterval(show, 7000); }
setInterval(() => { tickClock(); const s = nowLabel(); if (s !== nowSlot) { nowSlot = s; renderNow(); Object.values(ST.charts).forEach(c => { try { c.draw(); } catch (e) {} }); } }, 15000);

/* KPIs que cuentan hasta su valor */
function countUp(root = document) {
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  root.querySelectorAll(".kpi .v").forEach(v => { const n = v.firstChild; if (!n || n.nodeType !== 3) return; const txt = n.textContent, m = txt.match(/^([+−-]?)([\d,]*\.?\d+)$/); if (!m || v.dataset.done === txt) return;
    v.dataset.done = txt; const dec = (m[2].split(".")[1] || "").length, target = parseFloat(m[2].replace(/,/g, "")), t0 = performance.now();
    const step = t => { const p = Math.min(1, (t - t0) / 600), e = 1 - Math.pow(1 - p, 3); n.textContent = m[1] + fmt(target * e, dec); if (p < 1) requestAnimationFrame(step); else n.textContent = txt; }; requestAnimationFrame(step); });
}
new MutationObserver(() => countUp()).observe(document.querySelector("main"), {childList: true, subtree: true});

/* Modo sala: tema oscuro, pantalla completa, sin menú y rotación de páginas */
const SALA_TABS = ["pso", "rdo", "th", "ren", "ipsi"], SALA_MS = 25000;
let salaH = null, salaT0 = 0, salaPaused = false, salaPrevTheme = null, wakeL = null;
function salaTick() { if (salaPaused) return; const p = (performance.now() - salaT0) / SALA_MS; $("salaProg").style.width = Math.min(100, p * 100) + "%";
  if (p >= 1) { const i = SALA_TABS.indexOf(ST.tab); setTab(SALA_TABS[(i + 1) % SALA_TABS.length]); salaT0 = performance.now(); } salaLbl(); }
function salaLbl() { const i = SALA_TABS.indexOf(ST.tab), nx = TITLES[SALA_TABS[(i + 1) % SALA_TABS.length]], s = Math.max(0, Math.ceil((SALA_MS - (performance.now() - salaT0)) / 1000)); $("salaLbl").textContent = salaPaused ? "Rotación en pausa" : `Siguiente: ${nx} en ${s} s`; }
async function salaOn() {
  const R = document.documentElement; salaPrevTheme = R.getAttribute("data-theme"); R.setAttribute("data-theme", "dark"); R.classList.add("sala");
  $("salaBar").hidden = false; salaPaused = false; $("salaPause").textContent = "Pausar"; if (!SALA_TABS.includes(ST.tab)) setTab("pso");
  salaT0 = performance.now(); clearInterval(salaH); salaH = setInterval(salaTick, 250);
  try { await R.requestFullscreen?.(); } catch (e) {} try { wakeL = await navigator.wakeLock?.request("screen"); } catch (e) {}
  setTimeout(renderAll, 60);
}
function salaOff() {
  const R = document.documentElement; R.classList.remove("sala"); salaPrevTheme ? R.setAttribute("data-theme", salaPrevTheme) : R.removeAttribute("data-theme");
  $("salaBar").hidden = true; clearInterval(salaH); salaH = null; try { if (document.fullscreenElement) document.exitFullscreen(); } catch (e) {} try { wakeL?.release(); } catch (e) {} wakeL = null; setTimeout(renderAll, 60);
}
$("salaBtn").onclick = () => document.documentElement.classList.contains("sala") ? salaOff() : salaOn();
$("salaExit").onclick = salaOff;
$("salaPause").onclick = () => { salaPaused = !salaPaused; if (!salaPaused) salaT0 = performance.now(); $("salaPause").textContent = salaPaused ? "Reanudar" : "Pausar"; salaLbl(); };
document.addEventListener("keydown", e => { if (!document.documentElement.classList.contains("sala") || e.target.matches?.("input,textarea,select")) return;
  if (e.key === "Escape" && !ST.moment) salaOff(); if (e.key === " ") { e.preventDefault(); $("salaPause").click(); }
  if (e.key === "PageDown" || e.key === "PageUp") { const i = SALA_TABS.indexOf(ST.tab), n = SALA_TABS.length; setTab(SALA_TABS[(i + (e.key === "PageDown" ? 1 : n - 1)) % n]); salaT0 = performance.now(); } });
document.addEventListener("click", e => { const wk = e.target.closest("#nowBar [data-wk]"); if (wk) setWeek(wk.dataset.wk); });
tickClock();

/* ============ Iconos e ilustraciones (SVG en línea) ============ */
const IC = {
  pso: '<path d="M3 5h14v12H3zM3 9h14M7 3v4M13 3v4" />',
  rdo: '<path d="M16 7a6.5 6.5 0 1 0 .8 5M16 3v4h-4" /><path d="M10 7v4l2.5 1.5" />',
  th: '<path d="M10 2c2.5 3 5 5.2 5 9a5 5 0 0 1-10 0c0-2 1-3.4 2.2-4.5.3 1.6 1.1 2.5 2.2 2.8C9 6.8 9.2 4.4 10 2z" />',
  ren: '<path d="M10 9.5V18M7 18h6" /><circle cx="10" cy="8" r="1.3" /><path d="M10 6.7 9.3 1.8c.9-.3 1.8 0 2 .7zM11.2 8.6l4.6 1.8c-.2.9-1 1.4-1.7 1.2zM8.8 8.6 5 11.7c-.6-.7-.6-1.6 0-2.1z" />',
  ipsi: '<path d="M12.5 3.5a4 4 0 0 0-5 5L3 13l2.5 2.5 4.5-4.5a4 4 0 0 0 5-5L12.6 8.4 10.9 8l-.4-1.7z" />',
  cmp: '<path d="M4 7h11l-3-3M16 13H5l3 3" />',
  chat: '<path d="M3 4h14v9H8l-4 3v-3H3z" />'
};
const ICT = {
  "HIDROELÉCTRICA": '<path d="M10 2.5C7 6.5 5 9.2 5 12a5 5 0 0 0 10 0c0-2.8-2-5.5-5-9.5z" />',
  "TERMOELÉCTRICA": IC.th,
  "EÓLICA": IC.ren,
  "SOLAR": '<circle cx="10" cy="10" r="3.5" /><path d="M10 2v2.5M10 15.5V18M2 10h2.5M15.5 10H18M4.3 4.3l1.8 1.8M13.9 13.9l1.8 1.8M4.3 15.7l1.8-1.8M13.9 6.1l1.8-1.8" />'
};
const svgI = (p, s = 18) => `<svg class="ico" viewBox="0 0 20 20" width="${s}" height="${s}" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
document.querySelectorAll("#navTabs button").forEach(b => { if (IC[b.dataset.tab] && !b.querySelector(".ico")) b.insertAdjacentHTML("afterbegin", svgI(IC[b.dataset.tab])); });
document.querySelector(".brand .mark")?.insertAdjacentHTML("afterbegin", '<svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true" fill="none" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"><path d="M12 2 7 22M12 2l5 20M8.6 15h6.8M9.8 9.5h4.4M5 7h14M8 7l4 4 4-4"/></svg>');
// silueta de torres de alta tensión para la franja del encabezado
(() => { const W = 1600, H = 120, tw = (x, h) => { const t = H - h; return `<path d="M${x - 14} ${H}L${x} ${t}L${x + 14} ${H}M${x - 9} ${t + h * .55}h18M${x - 6} ${t + h * .3}h12M${x - 22} ${t + 10}h44M${x - 16} ${t + 24}h32M${x - 22} ${t + 10}l22 14 22-14"/>`; };
  let p = "", wire = ""; const xs = [80, 330, 580, 830, 1080, 1330, 1580];
  xs.forEach((x, k) => { const h = 70 + (k % 2) * 14; p += tw(x, h); if (k) { const x0 = xs[k - 1], h0 = 70 + ((k - 1) % 2) * 14; [10, 24].forEach(o => { const y0 = H - h0 + o, y1 = H - h + o; wire += `<path d="M${x0 + 22} ${y0}Q${(x0 + x) / 2} ${Math.max(y0, y1) + 22} ${x - 22} ${y1}"/>`; }); } });
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMaxYMax slice"><g fill="none" stroke="white" stroke-opacity=".16" stroke-width="1.6">${p}</g><g fill="none" stroke="white" stroke-opacity=".12" stroke-width="1">${wire}</g></svg>`;
  document.documentElement.style.setProperty("--towers", `url("data:image/svg+xml,${encodeURIComponent(svg)}")`); })();

/* Esquema del SEIN para el panel «Ahora»: Norte – Centro – Sur con sus enlaces */
function seinMap(w, V, i, Lt) {
  const cm = b => Lt && V.cmg[b][i] != null ? V.cmg[b][i] : w.cmg[b][i];
  const fl = k => { const a = w.flujo?.[k]; if (!a) return null; const mx = Math.max(...a.filter(v => v != null).map(Math.abs)); return {v: a[i], p: mx ? Math.abs(a[i]) / mx : 0}; };
  const N = cm("Trujillo"), C = cm("Santa Rosa"), S = cm("Socabaya"), fN = fl("FLUJO ENLACE CENTRO - NORTE"), fS = fl("FLUJO ENLACE CENTRO - SUR");
  const lo = Math.min(N, C, S), hi = Math.max(N, C, S), heat = v => hi - lo < 5 ? 0.25 : 0.15 + 0.85 * (v - lo) / (hi - lo);
  const node = (x, y, z, v, col) => `<g><circle cx="${x}" cy="${y}" r="15" fill="${col}" fill-opacity="${heat(v).toFixed(2)}" stroke="${col}" stroke-width="2"/><text x="${x + 22}" y="${y - 3}" class="sz">${z}</text><text x="${x + 22}" y="${y + 11}" class="sv">${fmt(v, 0)}</text></g>`;
  const link = (y0, y1, f, lbl) => { if (!f) return ""; const hot = f.p >= .85, col = hot ? css("--bad") : css("--fg-3");
    return `<line x1="30" y1="${y0}" x2="30" y2="${y1}" stroke="${col}" stroke-width="${2 + 3 * f.p}" stroke-linecap="round" class="${hot ? "flow hot" : "flow"}"/><text x="40" y="${(y0 + y1) / 2 + 4}" class="sl" fill="${col}">${fmt(f.v)} MW · ${fmt(100 * f.p)}%</text>`; };
  return `<svg class="sein" viewBox="0 0 150 168" width="150" height="168" role="img" aria-label="Esquema del SEIN: CMg por zona y flujo por enlace">
    ${link(30, 84, fN, "C–N")}${link(84, 138, fS, "C–S")}
    ${node(30, 26, "Norte", N, css("--b-norte"))}${node(30, 84, "Centro", C, css("--b-centro"))}${node(30, 142, "Sur", S, css("--b-sur"))}</svg>`;
}
function tipoIcon(t) { return ICT[t] ? svgI(ICT[t], 16) : ""; }

document.addEventListener("click", e => { const b = e.target.closest("#cmgScale [data-sc]"); if (!b) return; ST.cmgScale = b.dataset.sc; try { localStorage.setItem("pso-cmgscale", ST.cmgScale); } catch (x) {} renderTab(); });
try { ST.cmgScale = localStorage.getItem("pso-cmgscale") || "todo"; } catch (x) { ST.cmgScale = "todo"; }

/* ============ Post operación (IEOD / IDCOS, D+1) ============ */
ST.post = Object.assign({}, BOOT.post || {});
try { const s = JSON.parse(localStorage.getItem("pso-post") || "null"); if (s) Object.assign(ST.post, s); } catch (e) {}
try { ST.tc = +localStorage.getItem("pso-tc") || 0; } catch (e) { ST.tc = 0; }
function persistPost() { try { const own = Object.fromEntries(Object.entries(ST.post).filter(([d]) => !(BOOT.post || {})[d])); localStorage.setItem("pso-post", JSON.stringify(own)); } catch (e) {} }
const dmy2iso = s => { const m = String(s || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/); return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` + (m[4] ? `T${m[4].padStart(2, "0")}:${m[5]}` : "") : (/^\d{4}-\d\d-\d\dT/.test(String(s)) ? String(s).slice(0, 16) : ""); };
const slotIx = h => { const m = String(h || "").match(/(\d{1,2}):(\d{2})/); if (!m) return -1; const t = +m[1] * 60 + +m[2]; return t === 0 ? 47 : Math.round(t / 30) - 1; };
const tbl = (rows, key) => { const i = rows.findIndex(r => r.some(v => String(v || "").trim().toUpperCase() === key)); return i < 0 ? null : {h: rows[i].map(v => String(v || "").replace(/\s+/g, " ").trim()), rows: rows.slice(i + 1).filter(r => r.some(v => v != null && v !== ""))}; };
const col = (t, re) => t.h.findIndex(h => re.test(h));
// AnexoA del IEOD (ya convertido en filas)
function parseAnexoARows(S, fname) {
  const out = {fuente: fname};
  const DE = S["DESPACHO_EJECUTADO"]; let date = null;
  if (DE) { const hi = DE.findIndex(r => String(r[0] || "").toUpperCase() === "FECHA"), zona = DE[hi + 1], tipo = DE[hi + 2], emp = DE[hi + 3], cen = DE[hi + 4], ej = {};
    for (const r of DE.slice(hi + 5)) { const s = slotIx(r[1]); if (s < 0) continue; date = date || dmy2iso(r[0]).slice(0, 10);
      for (let j = 2; j < r.length; j++) { const c = cen[j]; if (!c || /^MW$/i.test(String(c).trim())) continue; const k = String(c).trim(); (ej[k] || (ej[k] = Array(48).fill(null)))[s] = typeof r[j] === "number" ? r[j] : (r[j] == null ? null : +r[j] || 0); } }
    out.ejec = ej; out.ejMeta = Object.fromEntries(cen.map((c, j) => c && j > 1 ? [String(c).trim(), {zona: zona[j], tipo: tipo[j], empresa: emp[j]}] : null).filter(Boolean)); }
  out.date = date;
  const DD = S["DESV_DEMANDA_PRONOS"]; if (DD) { const t = tbl(DD, "EJECUTADA") || tbl(DD, "FECHA"), ie = col(t, /EJECUTADA/i), ip = col(t, /PROG/i), ih = col(t, /HORA/i); out.dem = {ej: Array(48).fill(null), pd: Array(48).fill(null)};
    for (const r of t.rows) { const s = slotIx(r[ih]); if (s >= 0) { out.dem.ej[s] = r[ie]; out.dem.pd[s] = r[ip]; } } }
  const IC = S["INTERCONEXIONES"]; if (IC) { const hi = IC.findIndex(r => String(r[0] || "").toUpperCase() === "FECHA"), H = IC[hi].map(v => String(v || "")), grp = IC[hi - 1].map(v => String(v || ""));
    let g = ""; const which = H.map((h, j) => { if (/CENTRO\s*-\s*NORTE/i.test(grp[j])) g = "cn"; else if (/CENTRO\s*-\s*SUR/i.test(grp[j])) g = "cs"; return /TOTAL/i.test(h) ? g : null; });
    out.inter = {cn: Array(48).fill(null), cs: Array(48).fill(null)};
    for (const r of IC.slice(hi + 1)) { const s = slotIx(r[1]); if (s < 0) continue; which.forEach((k, j) => { if (k && typeof r[j] === "number") out.inter[k][s] = Math.abs(r[j]); }); }
    if (!which.includes("cs")) { // sin columna TOTAL para el Sur: suma de enlaces del grupo Centro–Sur
      const js = H.map((h, j) => j).filter(j => j > 1 && grp.slice(0, j + 1).reverse().find(x => x) && /CENTRO\s*-\s*SUR/i.test(grp.slice(0, j + 1).reverse().find(x => x)));
      for (const r of IC.slice(hi + 1)) { const s = slotIx(r[1]); if (s >= 0) out.inter.cs[s] = js.reduce((a, j) => a + (typeof r[j] === "number" ? r[j] : 0), 0); } } }
  const CO = S["COSTO_OPE_EJEC"]; if (CO) { const t = tbl(CO, "FECHA"); out.costo = t ? t.rows.map(r => ({d: String(r[t.h.indexOf("Fecha")] || "").trim(), ej: r[col(t, /Ejecutado/i)], pr: r[col(t, /Programado/i)]})).filter(x => typeof x.ej === "number") : [];
    out.costoTxt = (CO.find(r => r.some(v => /costo total de la operaci/i.test(String(v || "")))) || []).find(v => /costo total/i.test(String(v || ""))) || ""; }
  const lst = (sh, key, map) => { const R = S[sh]; if (!R) return []; const t = tbl(R, key); return t ? t.rows.map(r => map(Object.fromEntries(t.h.map((h, j) => [h, r[j]])))).filter(Boolean) : []; };
  const prog = p => ({P: "Programado", R: "Reprogramado", F: "Forzado", E: "Emergencia"}[String(p || "").trim()] || String(p || "").trim());
  out.indisp = lst("INDISP", "EMPRESA", x => x.Equipo ? {empresa: x.Empresa, ubic: x["Ubicación"], equipo: String(x.Equipo).trim(), ini: dmy2iso(x.Inicio), fin: dmy2iso(x.Final), desc: String(x["Descripción"] || "").trim(), mw: +x["MW Indisp."] || 0, prog: prog(x["Prog."]), tipo: String(x.Tipo || "").trim()} : null);
  out.mant = lst("MANTENIMIENTO EJECUTADOS", "EMPRESA", x => x.Equipo ? {empresa: x.Empresa, ubic: x["Ubicación"], equipo: String(x.Equipo).trim(), ini: dmy2iso(x.Inicio), fin: dmy2iso(x.Final), desc: String(x["Descripción"] || "").trim(), mw: +x["MW Indisp."] || 0, prog: prog(x["Prog."]), tipo: String(x.Tipo || "").trim()} : null);
  out.eventos = lst("EVENTOS", "TIPO DE EVENTO", x => x.EQUIPO ? {tipo: x["TIPO DE EVENTO"], empresa: x.EMPRESA, ubic: x["UBICACIÓN"], equipo: String(x.EQUIPO).trim(), ini: dmy2iso(x.INICIO), fin: dmy2iso(x.FINAL), desc: String(x["DESCRIPCIÓN"] || "").trim(), mw: +x["MW INDISP."] || 0, interr: x["INTERRUPCIÓN (SI/NO)"]} : null);
  out.restric = lst("RESTRIC_OPE", "FECHA", x => x["DESCRIPCIÓN"] ? {ini: x["HORA INICIO"], fin: x["HORA FINAL"], empresa: x.EMPRESA, ubic: x["UBICACIÓN"], equipo: x.EQUIPO, desc: String(x["DESCRIPCIÓN"]).trim()} : null);
  const cal = S["HOP"] || S["CALIFICA_OPE_UG"];
  out.calif = cal ? lst(S["HOP"] ? "HOP" : "CALIFICA_OPE_UG", "EMPRESA", x => x["Tipo Operación"] ? {empresa: x.Empresa, central: String(x.Central || "").trim(), grupo: String(x.Grupo || "").trim(), modo: String(x["Modo de Operación"] || "").trim(), ini: dmy2iso(x.Inicio), fin: dmy2iso(x.Final), arr: dmy2iso(x["O. Arranque"]), par: dmy2iso(x["O. Parada"]), tipo: String(x["Tipo Operación"]).trim(), lim: String(x["Lim. Transm."] || "").trim(), causa: String(x.Causa || "").replace(/^\.+$/, "").trim(), sis: String(x.Sistema || "").trim()} : null) : [];
  return out;
}
function parseCmgRows(rows) {
  const k = rows.findIndex(r => r.some(v => /S\/\.?\/MWh/i.test(String(v || "")))); if (k < 0) return null; const H = rows[k].map(v => String(v || "").trim());
  const pick = n => H.indexOf(n), ix = {sr: pick("SANTA ROSA 220"), tr: pick("TRUJILLO 220"), so: pick("SOCABAYA 220")}, out = {sr: Array(48).fill(null), tr: Array(48).fill(null), so: Array(48).fill(null)}; let date = null;
  for (const r of rows.slice(k + 1)) { const t = String(r[1] || ""), m = t.match(/(\d{4}-\d\d-\d\d)[T ](\d\d:\d\d)/); if (!m) continue; date = date || m[1]; const s = slotIx(m[2]); if (s < 0) continue; for (const [kk, j] of Object.entries(ix)) if (j >= 0) out[kk][s] = r[j]; }
  return {date, cmg: out, unidad: "S/./MWh"};
}
function parseIdcosText(t) {
  const d = dateFromSpanish(t.slice(0, 3000)) || null, reps = [];
  const sec = (t.match(/Reprogramaciones del d[ií]a([\s\S]*?)(?:\n\s*\d+\.\s+Observaciones|$)/i) || [])[1] || "";
  const re = /(\d{1,2}:\d{2})\s+SCO\s*-\s*(\d+)\s*-\s*([A-Z])\b/g; let m; while ((m = re.exec(sec))) reps.push({hora: m[1].padStart(5, "0"), reprograma: m[3], codigo: `SCO-${m[2]}-${m[3]}`});
  const maxcap = ((t.match(/operaron al m[aá]ximo de su capacidad de transmisi[oó]n\s*([\s\S]*?)\n\s*\d+\.\s/i) || [])[1] || "").replace(/\s+/g, " ").trim();
  const s6 = (t.match(/Operaci[oó]n por regulaci[oó]n de tensi[oó]n([\s\S]*?)\n\s*7\.\s/i) || [])[1] || "";
  const tCent = ((s6.match(/6\.1\.[^\n]*\n([\s\S]*?)6\.2\./) || [])[1] || "").replace(/\s+/g, " ").trim();
  const tLin = [...s6.matchAll(/(L-\d{3,5})\s+(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})/g)].map(m => ({linea: m[1], ini: m[2].padStart(5, "0"), fin: m[3].padStart(5, "0"), motivo: ((s6.slice(s6.indexOf("Motivo")).replace(/\s+/g, " ").match(/Por regulaci[oó]n de tensi[oó]n[^.]*\./i) || [])[0] || "Por regulación de tensión")}));
  const s8 = (t.match(/\n\s*8\.\s+Pruebas([\s\S]*?)\n\s*9\.\s/i) || [])[1] || "", s81 = (s8.match(/8\.1\.([\s\S]*?)8\.2\./) || [])[1] || "";
  const pru = [...s81.matchAll(/([A-ZÁÉÍÓÚÑ.]{3,})[ \t]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ ]{2,}?)[ \t]+((?:G|TG|U|TV)[ \t]?\d+)[ \t]+(\d{1,2}:\d{2})[ \t]+(\d{1,2}:\d{2})/g)].map(m => ({empresa: m[1], central: m[2].trim(), unidad: m[3], ini: m[4].padStart(5, "0"), fin: m[5].padStart(5, "0"), desc: /OPERATIVAS/i.test(s81) ? "Pruebas operativas" : "Pruebas"}));
  const sorteo = ((s8.match(/Resultado del sorteo:[ \t]*([^\n]+)/i) || [])[1] || "").trim();
  return {date: d, reps, maxcap, tension: {centrales: tCent, lineas: tLin}, pruebas: pru, sorteo};
}
function mergePost(date, part) { const P = ST.post[date] || (ST.post[date] = {}); Object.assign(P, part); persistPost(); }
async function sheetsOf(file) { const wb = XLSX.read(await file.arrayBuffer(), {type: "array", cellDates: true}); const S = {};
  for (const n of wb.SheetNames) S[n.trim()] = XLSX.utils.sheet_to_json(wb.Sheets[n], {header: 1, raw: true, defval: null}).map(r => r.map(v => v instanceof Date ? `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}T${pad2(v.getHours())}:${pad2(v.getMinutes())}` : v));
  return S; }
async function handlePostFile(f) { // devuelve un mensaje o null si no es de post operación
  const n = f.name.toLowerCase();
  if (/^anexoa_/.test(n)) { const r = parseAnexoARows(await sheetsOf(f), f.name); if (!r.date) throw new Error(`${f.name}: no encontré la fecha`); mergePost(r.date, r); return `IEOD del ${dayLabel(r.date)}`; }
  if (/^cmgcp/.test(n)) { const S = await sheetsOf(f), r = parseCmgRows(S["Cmg_Barra"] || Object.values(S)[0]); if (!r?.date) throw new Error(`${f.name}: no encontré las barras`); mergePost(r.date, {cmg: r.cmg, cmgUnit: r.unidad}); return `CMg ejecutado del ${dayLabel(r.date)}`; }
  if (/\.pdf$/.test(n) && /idcos|^lun_|^mar_|^mi[eé]_|^jue_|^vie_|^s[aá]b_|^dom_/.test(n)) { const t = await docText(f);
    if (/IDCOS/i.test(t.slice(0, 4000))) { const r = parseIdcosText(t); if (!r.date) throw new Error(`${f.name}: no encontré la fecha del IDCOS`); mergePost(r.date, {idcos: {reps: r.reps, maxcap: r.maxcap, tension: r.tension, pruebas: r.pruebas, sorteo: r.sorteo, name: f.name}});
      const have = new Set(ST.listado.filter(x => x.fecha === r.date).map(x => x.reprograma)); r.reps.filter(x => !have.has(x.reprograma)).forEach(x => ST.listado.push({fecha: r.date, hora: x.hora, reprograma: x.reprograma, causa: "(ver IDCOS)", operador: "", fuente: "IDCOS"})); return `IDCOS del ${dayLabel(r.date)}`; }
    if (/INFORME DE EVALUACI[OÓ]N DE LA OPERACI[OÓ]N DIARIA|IEOD/i.test(t.slice(0, 4000))) { const d = dateFromSpanish(t.slice(0, 3000)); if (d) { mergePost(d, {ieodTxt: t, ieodName: f.name}); return `Informe IEOD del ${dayLabel(d)}`; } } }
  return null;
}

/* --- cruces con hechos --- */
const nzS = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
function postFacts(date, words, before) { // indisponibilidades/eventos del día cuya ubicación aparece en el texto
  const P = ST.post[date]; if (!P) return []; const W = nzS(words);
  return [...(P.indisp || []).map(x => ({...x, k: "indisp"})), ...(P.eventos || []).map(x => ({...x, k: "evento"}))]
    .filter(x => { const u = nzS(x.ubic).replace(/^(C\.?\s?[HTSE]\.?|S\.?E\.?)\s+/, "").split(/\s+/)[0]; return u.length >= 4 && W.includes(u) && (!before || !x.ini || x.ini.slice(11) <= before); })
    .filter(x => x.k === "evento" || /Forzado|Reprogramado|Emergencia/.test(x.prog) || /trip|falla|indisponible/i.test(x.desc))
    .slice(0, 3);
}
const sCase = t => { t = String(t || "").replace(/\s+/g, " ").trim(); const up = (t.match(/[A-ZÁÉÍÓÚÑ]/g) || []).length, lo = (t.match(/[a-záéíóúñ]/g) || []).length; if (up <= lo) return t;
  return t.toLowerCase().replace(/\b([a-z]{0,4}-?\d+[a-z0-9]*)\b/g, m => m.toUpperCase()).replace(/\bkv\b/g, "kV").replace(/\bmw\b/g, "MW").replace(/\b(c\.[htse]\.|s\.e\.|cc\.hh\.)/g, m => m.toUpperCase()); };
const factTxt = x => `${x.ubic} ${x.equipo}: ${sCase(x.desc).slice(0, 110)}${x.ini ? ` (${x.ini.slice(11)}${x.prog ? ", " + x.prog.toLowerCase() : ""})` : ""}`;
function dropPostJust(x) { // caídas de centrales vs indisponibilidades ejecutadas
  const base = nzS(MAESTRO.get(x.central)?.central || x.central).replace(/[-_]/g, " ").replace(/\s+(TG|TV|G|U|CC)\s*\d.*$/, "").trim();
  for (const d of x.dias) { const P = ST.post[d]; if (!P) continue; const hit = [...(P.indisp || []), ...(P.mant || [])].find(m => base.length >= 4 && nzS(m.ubic).replace(/[-_]/g, " ").includes(base) && (m.mw > 0 || /indispon|f\/s|trip|falla/i.test(m.desc)));
    if (hit) return {texto: `IEOD ${dayLabel(d)}: ${factTxt(hit)}.`, fuente: "IEOD (indisponibilidades)", conf: "alta", cita: hit.desc, filas: [["IEOD " + dayLabel(d), factTxt(hit)], ["MW indisponibles", fmt(hit.mw, 1)]]}; }
  return null;
}

/* ============ Ejecutado en Daily + página Post operación ============ */
const postDays = () => Object.keys(ST.post).filter(d => days(ST.cur).includes(d) && ST.post[d].ejec).sort();
function ejecWeek(w, f = pass) { // MW ejecutados por media hora en la semana (null sin dato)
  const out = Array(w.times.length).fill(null);
  days(w).forEach((d, k) => { const P = ST.post[d]; if (!P?.ejec) return; for (let s = 0; s < 48; s++) { let v = 0, any = false; for (const [c, a] of Object.entries(P.ejec)) { if (!f(c) || a[s] == null) continue; v += a[s]; any = true; } out[k * 48 + s] = any ? v : null; } });
  return out;
}
function cmgEjWeek(w, k) { const out = Array(w.times.length).fill(null); if (!(ST.tc > 0)) return out;
  days(w).forEach((d, j) => { const a = ST.post[d]?.cmg?.[k]; if (a) for (let s = 0; s < 48; s++) out[j * 48 + s] = a[s] != null ? a[s] / ST.tc : null; }); return out; }

const POSTTAB = {id: "post", title: "Post operación"};
TITLES.post = "Post operación (D+1)";
(() => { const nav = document.querySelector('#navTabs button[data-tab="cmp"]'); if (nav && !document.querySelector('#navTabs [data-tab="post"]')) { nav.insertAdjacentHTML("beforebegin", `<button role="tab" data-tab="post" aria-selected="false">${typeof svgI === "function" ? svgI('<path d="M4 16V9M8 16V5M12 16v-6M16 16V3M3 17h14"/>') : ""}Post operación</button>`); document.querySelector('#navTabs [data-tab="post"]').addEventListener("click", () => setTab("post")); } })();

function renderPOST() {
  const w = ST.cur, D = postDays(), el = $("tab-post"); if (!el) return;
  if (!D.length) { $("postBody").hidden = true; $("postEmpty").hidden = false; return; } $("postBody").hidden = false; $("postEmpty").hidden = true;
  if (!D.includes(ST.postDay)) ST.postDay = D[D.length - 1]; const d = ST.postDay, P = ST.post[d], k0 = days(w).indexOf(d) * 48, I = [...Array(48).keys()].map(s => k0 + s), L = pick(w.times, I);
  $("postDays").innerHTML = D.map(x => `<button type="button" data-pday="${x}" aria-pressed="${x === d}">${dayLabel(x)}</button>`).join("");
  $("postTc").value = ST.tc || "";
  // KPIs
  const V = getVW(), ej = pick(ejecWeek(w), I), vig = I.map(i => V.letter[i] ? sumCen(V.cen, [i], pass)[0] : null), pso = sumCen(w.centrales, I, pass);
  const E = a => a.some(v => v != null) ? sum(a) * 0.5 / 1000 : null, eE = E(ej), eV = E(vig.map((v, j) => v ?? pso[j])), eP = E(pso);
  const cRow = (P.costo || []).find(x => new RegExp(`\\b${+d.slice(8)}\\b`).test(x.d)), cPct = cRow ? 100 * (cRow.ej - cRow.pr) / cRow.pr : null;
  const dmx = P.dem ? P.dem.ej.reduce((a, v, s) => v != null && Math.abs(v - P.dem.pd[s]) > Math.abs(a.v) ? {v: v - P.dem.pd[s], s} : a, {v: 0, s: -1}) : null;
  const reps = (P.idcos?.reps?.length ? P.idcos.reps : ST.listado.filter(x => x.fecha === d)).map(x => x.reprograma).sort();
  $("postKpis").innerHTML = kpiHtml([
    ["Energía ejecutada", eE != null ? fmt(eE, 1) : "—", "GWh", eE != null ? `${sgn(eE - eV, 2)} vs programa vigente · ${sgn(eE - eP, 2)} vs PSO` : ""],
    ["Mayor desvío de demanda", dmx && dmx.s >= 0 ? sgn(dmx.v, 0) : "—", "MW", dmx && dmx.s >= 0 ? `${L[dmx.s].slice(11)} · ejecutada ${fmt(P.dem.ej[dmx.s])} MW` : ""],
    ["Reprogramas del día", reps.length || "—", "", reps.join(" · ") + (P.idcos ? " (IDCOS)" : "")]]);
  // Generación: PSO, vigente, ejecutado
  const C = BI();
  const gd = [{label: "PSO MW", data: pso, borderColor: C.pso, backgroundColor: C.pso, borderWidth: 1.5}, {label: "RDO vigente MW", data: vig, borderColor: C.rdo, backgroundColor: C.rdo, borderDash: [2, 3], borderWidth: 3},
    {label: "Ejecutado MW", data: ej, borderColor: css("--ejec"), backgroundColor: css("--ejec"), borderWidth: 2.5}];
  mkChart("cPostGen", {type: "line", data: {labels: L, datasets: gd}, options: postOpts(L, "Potencia (MW)")}); legend("lgPostGen", gd.map(x => ({label: x.label, color: x.borderColor, line: true, dash: !!x.borderDash})));
  // Demanda
  if (P.dem) { const dd = [{label: "Demanda programada (PDO)", data: P.dem.pd, borderColor: C.pso, backgroundColor: C.pso, borderDash: [5, 4], borderWidth: 1.8}, {label: "Demanda ejecutada", data: P.dem.ej, borderColor: css("--ejec"), backgroundColor: css("--ejec"), borderWidth: 2.5}];
    mkChart("cPostDem", {type: "line", data: {labels: L, datasets: dd}, options: postOpts(L, "MW")}); legend("lgPostDem", dd.map(x => ({label: x.label, color: x.borderColor, line: true, dash: !!x.borderDash}))); }
  // CMg ejecutado por barra
  if (P.cmg) { const tc = ST.tc > 0, cv = a => a.map(v => v == null ? null : tc ? v / ST.tc : v), u = tc ? "US$/MWh" : "S/./MWh";
    const cd = [["sr", "Santa Rosa (Centro)", "--b-centro", "Santa Rosa"], ["tr", "Trujillo (Norte)", "--b-norte", "Trujillo"], ["so", "Socabaya (Sur)", "--b-sur", "Socabaya"]].flatMap(([k, n, v, b]) => {
      const out = [{label: `Ejecutado ${n}`, data: cv(P.cmg[k]), borderColor: css(v), backgroundColor: css(v), borderWidth: 2.2}];
      if (tc) out.push({label: `RDO ${n}`, data: I.map(i => V.letter[i] ? V.cmg[b][i] : w.cmg[b][i]), borderColor: css(v) + "77", backgroundColor: css(v) + "77", borderDash: [2, 3], borderWidth: 2}); return out; });
    mkChart("cPostCmg", {type: "line", data: {labels: L, datasets: cd}, options: postOpts(L, u)}); legend("lgPostCmg", cd.map(x => ({label: x.label, color: x.borderColor, line: true, dash: !!x.borderDash})));
    $("postCmgNote").textContent = tc ? `Convertido con ${fmt(ST.tc, 3)} S/ por US$. Punteado: CMg del programa vigente.` : "En soles. Ingresa el tipo de cambio para compararlo con el CMg programado (US$/MWh)."; }
  // Desvíos por central (ejecutado − vigente)
  const cov = I.filter(i => V.letter[i]), dv = Object.keys(P.ejec || {}).filter(pass).map(c => { const e = P.ejec[c], b = I.map((i, s) => V.letter[i] ? (V.cen[c]?.[i] ?? 0) : (w.centrales[c]?.[i] ?? 0));
      const m = e.reduce((a, v, s) => a + ((v ?? 0) - b[s]) * 0.5, 0); return {c, m, ej: e.reduce((a, v) => a + (v ?? 0) * 0.5, 0)}; }).filter(x => Math.abs(x.m) >= 20).sort((a, b) => Math.abs(b.m) - Math.abs(a.m)).slice(0, 25);
  $("tblPostDev").innerHTML = dv.length ? `<table><thead><tr><th>Central</th><th>Tipo</th><th class="n">Ejecutado MWh</th><th class="n">Desvío vs programa vigente</th><th>Hecho registrado</th></tr></thead><tbody>${dv.map(x => { const f = postFacts(d, MAESTRO.get(x.c)?.central || x.c)[0];
      return `<tr data-filter="${esc(x.c)}"><td>${esc(x.c)}${isStk(x.c) ? ' <span class="pill media">STK</span>' : ""}</td><td class="faint">${esc((tipoOf(x.c) || P.ejMeta?.[x.c]?.tipo || "").toLowerCase())}</td><td class="n">${fmt(x.ej)}</td><td class="n ${x.m > 0 ? "pos" : "neg"}">${sgn(x.m)}</td><td class="faint">${f ? esc(factTxt(f)) : typeof calFactTxt === "function" ? esc(calFactTxt(d, x.c)) : ""}</td></tr>`; }).join("")}</tbody></table>` : `<p class="muted">Sin desvíos mayores a 20 MWh.</p>`;
  // Indisponibilidades y eventos
  const chip = p => `<span class="pchip ${/Forzado/.test(p) ? "f" : /Reprog/.test(p) ? "r" : /Emerg/.test(p) ? "f" : "p"}">${esc(p || "—")}</span>`;
  $("tblPostInd").innerHTML = (P.indisp || []).length ? `<table><thead><tr><th>Equipo</th><th>Inicio – fin</th><th>Descripción</th><th class="n">MW</th><th>Tipo</th></tr></thead><tbody>${[...P.indisp].sort((a, b) => (/Forzado/.test(b.prog) - /Forzado/.test(a.prog)) || b.mw - a.mw).map(x => `<tr class="${/Forzado/.test(x.prog) ? "hot" : ""}"><td><b>${esc(x.ubic)}</b> ${esc(x.equipo)}<div class="faint">${esc(x.empresa)}</div></td><td class="mono" style="white-space:nowrap">${(x.ini || "").slice(11)} – ${(x.fin || "").slice(0, 10) !== d ? "24:00" : (x.fin || "").slice(11)}</td><td>${esc(sCase(x.desc))}</td><td class="n">${x.mw ? fmt(x.mw, 1) : "—"}</td><td>${chip(x.prog)}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Sin indisponibilidades.</p>`;
  $("tblPostEv").innerHTML = (P.eventos || []).length ? `<table><thead><tr><th>Hora</th><th>Equipo</th><th>Descripción</th><th>Interrupción</th></tr></thead><tbody>${P.eventos.map(x => `<tr><td class="mono">${(x.ini || "").slice(11)}</td><td><b>${esc(x.equipo)}</b><div class="faint">${esc(x.ubic)}</div></td><td>${esc(sCase(x.desc))}</td><td>${/SI/i.test(x.interr) ? '<span class="pchip f">Sí</span>' : "No"}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Sin eventos.</p>`;
  // Mantenimientos del IPSI vs ejecutados
  const ips = (ST.ipsiItems[ST.key] || []).filter(it => it.ventanas.some(([a, b]) => a.slice(0, 10) <= d && b.slice(0, 10) >= d));
  const codes = s => (String(s).match(/[A-Z]{1,3}-?\d{3,4}|G\d+|TG\d+|U\d+/gi) || []).map(nzS);
  $("tblPostMant").innerHTML = ips.length ? `<table><thead><tr><th>Intervención del IPSI</th><th>Programada</th><th>Ejecutado (IEOD)</th></tr></thead><tbody>${ips.map(it => { const cs = codes(it.equipo), U = nzS(it.ubic).split(/\s+/)[0];
      const EQ = nzS(it.equipo).replace(/[^A-Z0-9]/g, ""), ex = [...(P.mant || []), ...(P.indisp || [])].find(m => (EQ.length >= 4 && nzS(m.equipo).replace(/[^A-Z0-9]/g, "") === EQ) || cs.some(c => nzS(m.equipo).replace(/\s/g, "").includes(c.replace(/\s/g, ""))) && (U.length < 4 || nzS(m.ubic).includes(U) || /^L-?\d/.test(cs[0] || "")));
      return `<tr><td><b>${esc(it.ubic)}</b> ${esc(it.equipo)}<div class="faint">${esc(it.desc)}</div></td><td class="mono">${esc(winText(it.ventanas))}</td><td>${ex ? `<span class="pchip ok">Ejecutado</span> <span class="mono">${(ex.ini || "").slice(11)}–${(ex.fin || "").slice(0, 10) !== d ? "24:00" : (ex.fin || "").slice(11)}</span>` : `<span class="pchip">No figura</span>`}</td></tr>`; }).join("")}</tbody></table>` : `<p class="muted">Sin intervenciones del IPSI para este día.</p>`;
  $("postMaxcap").innerHTML = P.idcos ? `<b>Equipos al máximo de su capacidad de transmisión (IDCOS 7):</b> ${esc(P.idcos.maxcap || "—")}` : "";
  if (typeof renderCalif === "function") renderCalif(d);
  $("postSrc").textContent = [P.fuente, P.cmg ? "CMgCP (IEOD)" : "", P.idcos?.name ? "IDCOS " + P.idcos.name : ""].filter(Boolean).join(" · ");
}
function postOpts(L, y) { const o = biOpts(L, {y}); o.scales.x.ticks.callback = (v, i) => /T(0[0-9]|1[0-9]|2[0-3]):00$/.test(L[i] || "") && +L[i].slice(11, 13) % 3 === 0 ? L[i].slice(11, 13) + "h" : ""; return o; }
document.addEventListener("click", e => { const b = e.target.closest("[data-pday]"); if (b) { ST.postDay = b.dataset.pday; renderPOST(); } });
document.addEventListener("change", e => { if (e.target.id === "postTc") { const v = parseFloat(String(e.target.value).replace(",", ".")); ST.tc = v > 0 && v < 20 ? v : 0; try { localStorage.setItem("pso-tc", ST.tc || ""); } catch (x) {} renderTab(); } });

/* ============ Calificación de la operación, pruebas y regulación de tensión ============ */
const MERIT = /POTENCIA O ENERG/i, COGEN = /COGENERACI/i;
const calColor = t => css(/PRUEBA/i.test(t) ? "--warn" : /SEGURIDAD/i.test(t) ? "--bad" : /TENSI/i.test(t) ? "--s-die" : COGEN.test(t) ? "--s-bio" : MERIT.test(t) ? "--b-centro" : "--c-mag");
const calName = t => sCase(String(t || "").replace(/^POR\s+/i, "por ")).replace(/^por /, "Por ").replace(/\bpr-(\d+)/i, "PR-$1").replace(/energia\b/i, "energía").replace(/cogeneracion\b/i, "cogeneración").replace(/tension\b/i, "tensión").replace(/regulacion\b/i, "regulación");
const cmp = s => nzS(s).replace(/[^A-Z0-9]/g, "");
const dice = (a, b) => { if (!a || !b) return 0; const g = s => { const o = new Map(); for (let i = 0; i < s.length - 1; i++) { const k = s.slice(i, i + 2); o.set(k, (o.get(k) || 0) + 1); } return o; }, A = g(a), B = g(b); let n = 0; for (const [k, v] of A) n += Math.min(v, B.get(k) || 0); return 2 * n / Math.max(1, a.length + b.length - 2); };
const CALMAP = new Map();
function calUnits(d, r) { // columnas del despacho ejecutado que corresponden a una fila de calificación
  const P = ST.post[d]; if (!P?.ejec) return []; const key = d + "|" + r.empresa + "|" + r.modo; if (CALMAP.has(key)) return CALMAP.get(key);
  const e1 = nzS(r.empresa).split(/\s+/)[0], cand = Object.keys(P.ejec).filter(k => !e1 || nzS(P.ejMeta?.[k]?.empresa || "").split(/\s+/)[0] === e1);
  const core = cmp(r.central.replace(/^(C\.?\s?T\.?\s?B?\.?|CT|C\.H\.)\s*/i, "")), mk = cmp(r.modo.split(/\s+-\s+/)[0]);
  let res = [];
  if (/^CENTRAL$/i.test(r.grupo)) { res = cand.filter(k => core.length >= 4 && cmp(k).startsWith(core)); if (!res.length) res = cand.filter(k => cmp(k).length >= 4 && core.startsWith(cmp(k)));
    if (!res.length) { const sc = cand.map(k => [k, Math.max(dice(cmp(k), core), dice(cmp(k), mk))]).sort((a, b) => b[1] - a[1]); if (sc[0] && sc[0][1] >= 0.6) res = [sc[0][0]]; } }
  else { const g = cmp(r.grupo), sc = cand.map(k => [k, dice(cmp(k), mk) + (cmp(k).endsWith(g) ? 0.3 : 0)]).sort((a, b) => b[1] - a[1]); if (sc[0] && sc[0][1] >= 0.75) res = [sc[0][0]]; }
  CALMAP.set(key, res); return res;
}
function calMWh(d, r) { const P = ST.post[d], U = calUnits(d, r); if (!U.length) return null; const a = toMin(r.ini || d + "T00:00"), b = toMin(r.fin || d + "T23:59"), d0 = toMin(d + "T00:00"); let e = 0;
  for (let s = 0; s < 48; s++) { const s0 = d0 + s * 30 * 6e4, s1 = s0 + 30 * 6e4, ov = Math.max(0, Math.min(b, s1) - Math.max(a, s0)) / (30 * 6e4); if (ov <= 0) continue; for (const u of U) e += (P.ejec[u][s] || 0) * 0.5 * ov; }
  return e; }
function calFactTxt(d, c) { // hecho de calificación para una central del PSO
  const P = ST.post[d]; if (!P?.calif) return ""; const rs = P.calif.filter(r => calUnits(d, r).includes(c)); if (!rs.length) return "";
  const nm = rs.find(r => !MERIT.test(r.tipo) && !COGEN.test(r.tipo)); if (nm) return `${calName(nm.tipo)} (${(nm.ini || "").slice(11)}–${(nm.fin || "").slice(0, 10) !== d ? "24:00" : (nm.fin || "").slice(11)})`;
  const ar = rs.find(r => r.arr); return ar ? `Arranque ordenado ${ar.arr.slice(11)}, en servicio ${ar.ini.slice(11)} · ${calName(ar.tipo).toLowerCase()}` : "";
}
ST.calF = "";
function renderCalif(d) {
  const P = ST.post[d], el = $("calCard"); if (!el) return; const C = P.calif || []; el.hidden = !C.length; if (!C.length) return;
  const rows = C.map(r => ({...r, mwh: calMWh(d, r), h: Math.max(0, (toMin(r.fin || d + "T23:59") - toMin(r.ini || d + "T00:00")) / 36e5)}));
  const by = {}; rows.forEach(r => { const g = by[r.tipo] || (by[r.tipo] = {n: new Set(), h: 0, e: 0, ok: true}); g.n.add(r.central + r.grupo); g.h += r.h; if (r.mwh == null) g.ok = false; else g.e += r.mwh; });
  const order = Object.keys(by).sort((a, b) => (MERIT.test(a) || COGEN.test(a)) - (MERIT.test(b) || COGEN.test(b)) || by[b].e - by[a].e);
  $("calSum").innerHTML = `<button type="button" class="calchip ${!ST.calF ? "on" : ""}" data-calf="">Todas <b>${rows.length}</b></button>` + order.map(t => { const g = by[t], nm = !MERIT.test(t) && !COGEN.test(t);
    return `<button type="button" class="calchip ${ST.calF === t ? "on" : ""} ${nm ? "nm" : ""}" data-calf="${esc(t)}" style="--c:${calColor(t)}"><i></i><span>${esc(calName(t))}</span><small>${g.n.size} unid. · ${fmt(g.h, 0)} h · ${fmt(g.e / 1000, 2)} GWh${g.ok ? "" : "*"}</small></button>`; }).join("");
  const vis = rows.filter(r => !ST.calF || r.tipo === ST.calF).sort((a, b) => (MERIT.test(a.tipo) || COGEN.test(a.tipo)) - (MERIT.test(b.tipo) || COGEN.test(b.tipo)) || (a.ini || "").localeCompare(b.ini || ""));
  const d0 = toMin(d + "T00:00"), pct = t => Math.max(0, Math.min(100, (toMin(t) - d0) / 864e3));
  const units = [...new Map(vis.map(r => [r.central + "|" + r.grupo, r])).values()];
  $("calGantt").innerHTML = `<div class="gax">${[0, 3, 6, 9, 12, 15, 18, 21, 24].map(h => `<span style="left:${h / 24 * 100}%">${pad2(h)}h</span>`).join("")}</div>`
    + units.map(u => `<div class="grow2"><span class="gl" title="${esc(u.empresa)}">${esc(u.central.replace(/^C\.T\.B?\.?\s*/i, ""))} <em>${esc(/^CENTRAL$/i.test(u.grupo) ? "" : u.grupo)}</em></span><span class="gt">${vis.filter(r => r.central + "|" + r.grupo === u.central + "|" + u.grupo).map(r => { const a = pct(r.ini || d + "T00:00"), b = pct(r.fin || d + "T23:59");
        return `<i style="left:${a}%;width:${Math.max(0.6, b - a)}%;background:${calColor(r.tipo)}" title="${esc(calName(r.tipo))} · ${(r.ini || "").slice(11)}–${(r.fin || "").slice(11)}${r.mwh != null ? " · " + fmt(r.mwh) + " MWh" : ""}"></i>${r.arr ? `<b class="arr" style="left:${pct(r.arr)}%" title="Orden de arranque ${r.arr.slice(11)}"></b>` : ""}`; }).join("")}</span></div>`).join("");
  $("tblCal").innerHTML = `<table><thead><tr><th>Unidad</th><th>Modo de operación</th><th>Calificación</th><th>Inicio – fin</th><th>Arranque / parada</th><th class="n">MWh</th></tr></thead><tbody>${vis.map(r => `<tr class="${!MERIT.test(r.tipo) && !COGEN.test(r.tipo) ? "nmrow" : ""}"><td><b>${esc(r.central)}</b> ${esc(/^CENTRAL$/i.test(r.grupo) ? "" : r.grupo)}<div class="faint">${esc(r.empresa)}</div></td><td class="faint">${esc(r.modo)}</td><td><span class="calt" style="--c:${calColor(r.tipo)}">${esc(calName(r.tipo))}</span>${r.causa ? `<div class="faint">${esc(r.causa)}</div>` : ""}</td><td class="mono" style="white-space:nowrap">${(r.ini || "").slice(11)} – ${(r.fin || "").slice(0, 10) !== d ? "24:00" : (r.fin || "").slice(11)}</td><td class="mono" style="white-space:nowrap">${r.arr ? r.arr.slice(11) : "—"} / ${r.par ? r.par.slice(11) : "—"}</td><td class="n">${r.mwh != null ? fmt(r.mwh) : "—"}</td></tr>`).join("")}</tbody></table>`;
  $("calNote").textContent = Object.values(by).some(g => !g.ok) ? "* Algunas unidades no se pudieron cruzar con el despacho ejecutado; su energía no está incluida." : "";
  // Pruebas y regulación de tensión (IDCOS)
  const I = P.idcos || {}, pr = I.pruebas || [], tl = I.tension?.lineas || [], cal25 = rows.filter(r => /PRUEBA/i.test(r.tipo));
  $("tblPru").innerHTML = (pr.length || cal25.length || I.sorteo) ? `<table><thead><tr><th>Tipo</th><th>Unidad</th><th>Horario</th><th>Detalle</th></tr></thead><tbody>${pr.map(x => `<tr><td><span class="calt" style="--c:${css("--warn")}">Requerimiento propio</span></td><td><b>${esc(x.central)}</b> ${esc(x.unidad)}<div class="faint">${esc(x.empresa)}</div></td><td class="mono">${x.ini}–${x.fin}</td><td>${esc(x.desc)}</td></tr>`).join("")}${cal25.map(r => `<tr><td><span class="calt" style="--c:${css("--warn")}">Aleatoria PR-25</span></td><td><b>${esc(r.central)}</b> ${esc(r.grupo)}<div class="faint">${esc(r.empresa)}</div></td><td class="mono">${(r.ini || "").slice(11)}–${(r.fin || "").slice(11)}</td><td>${r.mwh != null ? fmt(r.mwh) + " MWh" : ""}${I.sorteo ? ` · Sorteo: ${esc(sCase(I.sorteo))}` : ""}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Sin pruebas registradas.</p>`;
  $("tblTen").innerHTML = `<p class="tenc"><b>Centrales que operaron por regulación de tensión:</b> ${esc(I.tension?.centrales || "—")}</p>` + (tl.length ? `<table><thead><tr><th>Línea desconectada</th><th>Horario</th><th>Motivo</th></tr></thead><tbody>${tl.map(x => `<tr><td class="mono"><b>${esc(x.linea)}</b></td><td class="mono">${x.ini}–${x.fin}</td><td>${esc(x.motivo)}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">Sin líneas desconectadas por tensión.</p>`);
}
document.addEventListener("click", e => { const b = e.target.closest("[data-calf]"); if (b) { ST.calF = b.dataset.calf; renderCalif(ST.postDay); } });

/* ============ Intervenciones: mapa de impacto y comparación con la semana anterior ============ */
function impactRows(w) {
  const P = prevWeek()?.w, dd = days(w), drops = detectDrops(w), its = (ST.ipsiItems[ST.key] || []).filter(it => it.tipo === "Generación");
  const names = [...new Set([...drops.map(x => x.central), ...its.flatMap(it => it.centrales)])].filter(c => w.centrales[c] && pass(c));
  return names.map(c => { const a = w.centrales[c], p = P?.centrales?.[c], ref = Math.max(1, ...(p || a));
    const cells = dd.map((d, k) => { const sl = a.slice(k * 48, k * 48 + 48), ps = p ? p.slice(k * 48, k * 48 + 48) : null, mx = Math.max(...sl);
      const lost = ps ? Math.max(0, sum(ps.map((v, s) => (v || 0) - (sl[s] || 0))) * 0.5) : sum(sl.map(v => ref - (v || 0))) * 0.5;
      const ip = its.filter(it => it.centrales.includes(c) && overlapDay(it.ventanas, d)), iv = ST.post[d] && typeof dropPostJust === "function" ? (ST.post[d].indisp || []).find(m => nzS(m.ubic).includes(nzS(MAESTRO.get(c)?.central || c).split(/[\s_-]/)[0])) : null;
      return {d, frac: Math.max(0, Math.min(1, lost / (ref * 24))), lost, ip, iv, off: mx < 1 && ref > 1}; });
    return {c, ref, cells, tot: sum(cells.map(x => x.lost)), ipsi: its.filter(it => it.centrales.includes(c))}; })
    .filter(r => r.tot > 50 || r.ipsi.length).sort((a, b) => b.tot - a.tot).slice(0, 14);
}
function renderImpact() {
  const w = ST.cur, R = impactRows(w), P = prevWeek()?.w, dd = days(w), el = $("impGrid"); if (!el) return;
  if (!R.length) { el.innerHTML = `<p class="muted">Sin caídas relevantes ni intervenciones de generación esta semana.</p>`; return; }
  if (!ST.dropSel || !R.some(r => r.c === ST.dropSel)) ST.dropSel = R[0].c;
  const tot = sum(R.map(r => r.tot)), sus = R.filter(r => r.cells.some(x => x.ip.length || x.iv)).length, capOut = Math.max(...dd.map((_, k) => sum(R.map(r => r.cells[k].lost / 24))));
  $("impKpis").innerHTML = kpiHtml([["Energía no despachada", fmt(tot / 1000, 2), "GWh", P ? `frente a la semana ${P.semana}` : "frente a su máximo"], ["Capacidad fuera (día pico)", fmt(capOut), "MW", "MW medios no despachados"],
    ["Centrales afectadas", R.length, "", `${sus} con sustento (IPSI o IEOD)`]]);
  el.innerHTML = `<div class="ihd"><span></span>${dd.map(d => `<span class="${ST.F.days.has(d) ? "on" : ""}">${dayLabel(d)}</span>`).join("")}<span class="n">Total</span></div>`
    + R.map(r => `<button type="button" class="irow ${r.c === ST.dropSel ? "sel" : ""}" data-central="${esc(r.c)}"><span class="inm"><b>${esc(r.c)}</b>${isStk(r.c) ? ' <span class="pill media">STK</span>' : ""}<small>${esc((tipoOf(r.c) || "").toLowerCase())} · ref. ${fmt(r.ref)} MW</small></span>
      ${r.cells.map(x => `<span class="icell ${x.off ? "off" : ""}" style="--f:${(x.frac * 100).toFixed(0)}%" title="${dayLabel(x.d)}: ${x.off ? "fuera de servicio" : fmt(100 * x.frac) + " % de su energía de referencia"} · ${fmt(x.lost)} MWh menos${x.ip.length ? " · IPSI: " + x.ip.map(i => i.equipo).join(", ") : ""}${x.iv ? " · IEOD: " + x.iv.equipo : ""}">${x.lost >= 1 ? `<em>${x.lost >= 1000 ? fmt(x.lost / 1000, 1) + "k" : fmt(x.lost)}</em>` : ""}${x.ip.length ? '<i class="tg ipsi">IPSI</i>' : ""}${x.iv ? '<i class="tg ieod">IEOD</i>' : ""}</span>`).join("")}
      <span class="itot n">${fmt(r.tot)}</span></button>`).join("")
    + `<div class="ilg"><span><i class="sw" style="--f:15%"></i>leve</span><span><i class="sw" style="--f:55%"></i>media</span><span><i class="sw" style="--f:100%"></i>fuera de servicio</span><span><i class="tg ipsi">IPSI</i>intervención programada</span><span><i class="tg ieod">IEOD</i>indisponibilidad ejecutada</span><span class="faint">MWh menos ${P ? "que el mismo día de la semana " + P.semana : "que su máximo"} · clic = ver curva</span></div>`;
  // curva de la central elegida: esta semana vs anterior
  const c = ST.dropSel, I = idxs(w), L = pick(w.times, I), cur = pick(w.centrales[c], I), prv = P?.centrales?.[c] ? pick(P.centrales[c], I) : null;
  const bands = (ST.ipsiItems[ST.key] || []).filter(it => it.centrales.includes(c)).flatMap(it => it.ventanas);
  const ds = [{label: `Semana ${w.semana}`, data: cur, borderColor: css("--accent"), backgroundColor: css("--accent"), borderWidth: 2.2, fill: false, order: 1}];
  if (prv) ds.push({label: `Semana ${P.semana} (mismo día)`, data: prv, borderColor: css("--fg-3"), backgroundColor: color(css("--bad"), 0.22), borderDash: [5, 4], borderWidth: 1.5, fill: {target: 0, above: color(css("--bad"), 0.22), below: color(css("--good"), 0.18)}, order: 2});
  const o = biOpts(L, {y: "MW", min: 0}); o.plugins.bands = {items: bands};
  mkChart("cDrop", {type: "line", data: {labels: L, datasets: ds}, options: o});
  const r = R.find(x => x.c === c);
  $("dropTitle").textContent = `${c}: ${prv ? `${fmt(r.tot)} MWh menos que la semana ${P.semana}` : `${fmt(r.tot)} MWh bajo su máximo`}`;
  legend("lgDrop", [{label: `Semana ${w.semana}`, color: css("--accent"), line: true}].concat(prv ? [{label: `Semana ${P.semana}`, color: css("--fg-3"), line: true, dash: true}, {label: "Energía perdida", color: color(css("--bad"), 0.5)}] : []).concat(bands.length ? [{label: "Intervención IPSI", color: color(css("--warn"), 0.4)}] : []));
}
function color(c, a) { const m = String(c).trim().match(/^#([0-9a-f]{6})$/i); if (!m) return c; const n = parseInt(m[1], 16); return `rgba(${n >> 16},${(n >> 8) & 255},${n & 255},${a})`; }

/* ============ Intervenciones: ranking de energía perdida + cronograma del IPSI ============ */
const SUS = {IPSI: "--accent", IEOD: "--ejec", "Sin sustento": "--fg-3"};
const plugBarVal = {id: "barval", afterDatasetsDraw(ch, a, o) { if (!o?.on) return; const c = ch.ctx; c.save(); c.font = "600 11px IBM Plex Mono, monospace"; c.fillStyle = css("--fg-2"); c.textBaseline = "middle";
  ch.getDatasetMeta(0).data.forEach((b, i) => { const v = ch.data.datasets[0].data[i]; c.fillText(v >= 1000 ? fmt(v / 1000, 1) + " GWh" : fmt(v) + " MWh", b.x + 6, b.y); }); c.restore(); }};
if (typeof Chart !== "undefined" && Chart.register) { try { Chart.register(plugBarVal); } catch (e) {} }
function renderImpact() {
  const w = ST.cur, R = impactRows(w), P = prevWeek()?.w, dd = days(w);
  if (!R.length) { $("impKpis").innerHTML = ""; $("impBarsBox").innerHTML = `<p class="muted">Sin caídas relevantes esta semana.</p>`; return; }
  if (!ST.dropSel || !R.some(r => r.c === ST.dropSel)) ST.dropSel = R[0].c;
  R.forEach(r => r.sus = r.cells.some(x => x.ip.length) || r.ipsi.length ? "IPSI" : r.cells.some(x => x.iv) ? "IEOD" : "Sin sustento");
  const tot = sum(R.map(r => r.tot)), nS = R.filter(r => r.sus !== "Sin sustento").length;
  $("impKpis").innerHTML = kpiHtml([["Energía no despachada", fmt(tot / 1000, 1), "GWh", P ? `frente a la semana ${P.semana}` : "frente a su máximo"],
    ["Centrales afectadas", R.length, "", `${nS} con sustento en IPSI o IEOD`], ["Intervenciones IPSI", (ST.ipsiItems[ST.key] || []).length, "", "generación y transmisión"]]);
  // 1) ranking horizontal
  const B = R.filter(r => r.tot >= 1);
  $("impBarsBox").style.height = (B.length * 28 + 30) + "px";
  mkChart("cImpBars", {type: "bar", data: {labels: B.map(r => r.c), datasets: [{data: B.map(r => r.tot), backgroundColor: B.map(r => r.c === ST.dropSel ? css("--fg") : css(SUS[r.sus])), borderRadius: 4, borderSkipped: "start", barThickness: 16}]},
    options: {indexAxis: "y", responsive: true, maintainAspectRatio: false, layout: {padding: {right: 78}}, animation: {duration: 380},
      plugins: {legend: {display: false}, barval: {on: true}, tooltip: {callbacks: {title: it => it[0].label, label: c => { const r = B[c.dataIndex]; return [` ${fmt(r.tot)} MWh menos ${P ? "que la semana " + P.semana : "que su máximo"}`, ` Sustento: ${r.sus}${r.ipsi.length ? " · " + r.ipsi.map(i => i.equipo).join(", ") : ""}`]; }}}},
      scales: {x: {beginAtZero: true, grid: {color: css("--line")}, border: {display: false}, ticks: {color: css("--fg-3"), callback: v => v >= 1000 ? fmt(v / 1000, 0) + "k" : fmt(v)}},
        y: {grid: {display: false}, border: {display: false}, ticks: {color: css("--fg"), font: {size: 12, weight: "600"}}}},
      onClick: (e, els) => { if (els[0]) { ST.dropSel = B[els[0].index].c; renderIPSI(); } }, onHover: (e, els) => { e.native.target.style.cursor = els.length ? "pointer" : "default"; }}});
  $("lgImpBars").innerHTML = Object.entries(SUS).map(([k, v]) => `<span><i style="background:${css(v)}"></i>${k === "IPSI" ? "Intervención programada (IPSI)" : k === "IEOD" ? "Indisponibilidad ejecutada (IEOD)" : "Sin sustento en los informes"}</span>`).join("");
  // 2) cronograma de intervenciones del IPSI
  const its = (ST.ipsiItems[ST.key] || []).filter(it => !ST.F.central || it.centrales.some(pass)), t0 = toMin(w.inicio + "T00:00"), span = 7 * 864e5, pc = t => Math.max(0, Math.min(100, (toMin(t) - t0) / span * 100));
  const col = it => css(it.tipo === "Generación" ? "--accent" : it.tipo === "Transmisión" ? "--warn" : "--fg-3");
  const now = nowLabel(), nowP = w.times.includes(now) ? pc(now) : null;
  const rows = [...its].sort((a, b) => (a.ventanas[0]?.[0] || "").localeCompare(b.ventanas[0]?.[0] || ""));
  $("ipsiGantt").innerHTML = `<div class="gh"><span></span><div class="gd">${dd.map(d => `<span>${dayLabel(d)}</span>`).join("")}</div></div>`
    + rows.map(it => { const imp = R.find(r => it.centrales.includes(r.c)), sel = it.centrales.includes(ST.dropSel);
      return `<div class="gr ${sel ? "sel" : ""} ${it.centrales[0] ? "click" : ""}" ${it.centrales[0] ? `data-gc="${esc(it.centrales[0])}"` : ""} title="${esc(it.desc)}">
        <span class="gn"><b>${esc(/^CENTRAL$/i.test(it.equipo) ? it.ubic : /^(G|TG|TV|U|GP)\d/.test(it.equipo) ? it.ubic + " " + it.equipo : it.equipo)}</b><small>${esc(/^(G|TG|TV|U|GP)\d|^CENTRAL$/i.test(it.equipo) ? it.tipo + (it.mw > 0 ? ` · ${fmt(it.mw)} MW` : "") : it.ubic)}${imp && imp.tot >= 1 ? ` · <em>−${fmt(imp.tot)} MWh</em>` : ""}</small></span>
        <span class="gtk">${it.ventanas.map(([a, b]) => `<i style="left:${pc(a)}%;width:${Math.max(0.8, pc(b) - pc(a))}%;background:${col(it)}"></i>`).join("")}${nowP != null ? `<b class="gnow" style="left:${nowP}%"></b>` : ""}</span></div>`; }).join("")
    || `<p class="muted">Sin intervenciones del IPSI cargadas.</p>`;
  $("lgGantt").innerHTML = [["Generación", "--accent"], ["Transmisión", "--warn"], ["Conexión provisional", "--fg-3"]].map(([k, v]) => `<span><i style="background:${css(v)}"></i>${k}</span>`).join("") + (nowP != null ? `<span><i style="background:${css("--live")};width:2px"></i>Ahora</span>` : "");
  // 3) detalle de la central elegida
  const c = ST.dropSel, I = idxs(w), L = pick(w.times, I), cur = pick(w.centrales[c], I), prv = P?.centrales?.[c] ? pick(P.centrales[c], I) : null, r = R.find(x => x.c === c);
  const bands = (ST.ipsiItems[ST.key] || []).filter(it => it.centrales.includes(c)).flatMap(it => it.ventanas);
  const ds = [{label: `Semana ${w.semana}`, data: cur, borderColor: css("--accent"), backgroundColor: css("--accent"), borderWidth: 2.2, fill: false}];
  if (prv) ds.push({label: `Semana ${P.semana}`, data: prv, borderColor: css("--fg-3"), backgroundColor: css("--fg-3"), borderDash: [4, 4], borderWidth: 1.5, fill: {target: 0, above: color(css("--bad"), 0.16), below: "transparent"}});
  const o = biOpts(L, {y: "MW", min: 0}); o.plugins.bands = {items: bands};
  mkChart("cDrop", {type: "line", data: {labels: L, datasets: ds}, options: o});
  $("dropTitle").innerHTML = `${esc(c)} <span class="faint">· ${prv ? `${fmt(r?.tot || 0)} MWh menos que la semana ${P.semana}` : "sin dato de la semana anterior"}</span>`;
  legend("lgDrop", [{label: `Semana ${w.semana}`, color: css("--accent"), line: true}].concat(prv ? [{label: `Semana ${P.semana}`, color: css("--fg-3"), line: true, dash: true}, {label: "Energía no despachada", color: color(css("--bad"), 0.35)}] : []).concat(bands.length ? [{label: "Ventana del IPSI", color: color(css("--warn"), 0.35)}] : []));
}
document.addEventListener("click", e => { const g = e.target.closest(".gr[data-gc]"); if (g) { ST.dropSel = g.dataset.gc; renderIPSI(); } });

/* ============ Energía por tipo: cambia según la pestaña ============ */
function tiposCtx() {
  const w = ST.cur, t = ST.tab, I0 = idxs(w), dd = days(w);
  const byT = (I, val) => { const o = {}; TIPOS.forEach(x => o[x] = I.map(() => 0)); return {o, add(c, tp, j, v) { if (tp && o[tp] && v != null) o[tp][j] += v; }}; };
  const fromMap = (map, I) => { const b = byT(I); for (const [c, a] of Object.entries(map)) { if (!pass(c)) continue; const tp = tipoOf(c); I.forEach((i, j) => b.add(c, tp, j, a?.[i])); } return b.o; };
  const week = () => ({head: "Programado PSO", sub: `Semana ${w.semana}${ST.F.days.size ? " · días elegidos" : ""}`, cur: fromMap(w.centrales, I0)});
  if (["rdo", "th", "ren"].includes(t)) {
    const V = getVW(), {day, note} = dailyDay(), k = dd.indexOf(day), I = [...Array(48).keys()].map(s => k * 48 + s), R = V.days.find(x => x.date === day)?.R, lt = R?.letters || [];
    const tag = `${dayLabel(day)}/${day.slice(5, 7)}${day === todayISO() ? " · hoy" : ""}`;
    // programa del día: PDO si está cargado, si no el PSO de ese día
    const PD = ST.pdo?.[day] ? classify([{date: day, slots: ST.pdo[day]}], day) : null, prog = byT(I);
    if (PD) for (const [c, a] of Object.entries(PD.centrales)) { if (!pass(c)) continue; const tp = tipoOf(c); for (let j = 0; j < 48; j++) prog.add(c, tp, j, a?.[j]); }
    else for (const [c, a] of Object.entries(w.centrales)) { if (!pass(c)) continue; const tp = tipoOf(c); I.forEach((i, j) => prog.add(c, tp, j, a?.[i])); }
    const vig = byT(I); TIPOS.forEach(x => vig.o[x] = [...prog.o[x]]);
    if (lt.length) { TIPOS.forEach(x => I.forEach((i, j) => { if (V.letter[i]) vig.o[x][j] = 0; }));
      for (const [c, a] of Object.entries(V.cen)) { if (!pass(c)) continue; const tp = tipoOf(c); I.forEach((i, j) => { if (V.letter[i]) vig.add(c, tp, j, a?.[i]); }); } }
    const base = PD ? "PDO" : "PSO";
    if (t === "rdo") return {head: `Programado ${base}${PD ? "" : " (falta el PDO)"}`, sub: tag, note, cur: prog.o, ...(lt.length ? {ref: vig.o, refLbl: `con Reprog ${lt[lt.length - 1]}`, rev: true} : {})};
    return {head: lt.length ? `Vigente · Reprog ${lt[lt.length - 1]}` : `Programado ${base}`, sub: `${tag}${lt.length ? " · cambio vs " + base : ""}`, note, cur: vig.o, ...(lt.length ? {ref: prog.o, refLbl: base} : {}),
      hl: t === "th" ? ["TERMOELÉCTRICA"] : ["EÓLICA", "SOLAR"]};
  }
  if (t === "post") {
    const D = postDays(); if (!D.length) return {...week(), note: "Sin IEOD cargado: se muestra la semana."};
    const day = D.includes(ST.postDay) ? ST.postDay : D[D.length - 1], P = ST.post[day], cur = {};
    TIPOS.forEach(x => cur[x] = Array(48).fill(0));
    for (const [c, a] of Object.entries(P.ejec)) { if (!pass(c)) continue; const tp = P.ejMeta?.[c]?.tipo || tipoOf(c); if (cur[tp]) a.forEach((v, s) => { if (v != null) cur[tp][s] += v; }); }
    const V = getVW(), k = dd.indexOf(day), I = [...Array(48).keys()].map(s => k * 48 + s), ref = {};
    TIPOS.forEach(x => ref[x] = Array(48).fill(0));
    for (const c of new Set([...Object.keys(V.cen), ...Object.keys(V.cenB)])) { if (!pass(c)) continue; const tp = tipoOf(c); if (!ref[tp]) continue; I.forEach((i, j) => { const v = V.letter[i] ? V.cen[c]?.[i] : V.cenB[c]?.[i]; if (v != null) ref[tp][j] += v; }); }
    return {head: "Ejecutado (IEOD)", sub: `${dayLabel(day)}/${day.slice(5, 7)} · cambio vs programa vigente`, cur, ref, refLbl: "vigente"};
  }
  if (t === "cmp") { const P = prevWeek()?.w, base = week(); if (!P) return {...base, note: "Falta la semana anterior."};
    return {...base, sub: `Semana ${w.semana} · cambio vs semana ${P.semana}`, ref: fromMap(P.centrales, I0.filter(i => i < P.times.length)), refLbl: "sem " + P.semana}; }
  return week();
}
function renderTipos() {
  const x = tiposCtx(), hl = x.hl || [];
  const rows = TIPOS.map(t => { const a = x.cur[t], e = gwh(a), r = x.ref ? gwh(x.ref[t]) : null, dl = r != null ? (x.rev ? r - e : e - r) : null, col = css(TV[t]);
    return `<div class="tipo ${hl.includes(t) ? "hl" : ""} ${hl.length && !hl.includes(t) ? "dim" : ""}"><span class="ti" style="color:${col};background:color-mix(in srgb,${col} 14%,transparent)">${typeof tipoIcon === "function" ? tipoIcon(t) : ""}</span><div><b>${t[0] + t.slice(1).toLowerCase()}</b><div class="v">${fmt(e, 1)} GWh · máx ${fmt(Math.max(0, ...a.map(v => v || 0)))} MW</div>${dl != null && Math.abs(dl) >= 0.05 ? `<div class="d">${dl > 0 ? "▲ +" : "▼ −"}${fmt(Math.abs(dl), 1)} GWh ${x.rev ? "" : "vs "}${esc(x.refLbl)}</div>` : dl != null ? `<div class="d faint">sin cambio ${x.rev ? "" : "vs "}${esc(x.refLbl)}</div>` : ""}</div></div>`; }).join("");
  $("tipoList").innerHTML = `<div class="tphd"><b>${esc(x.head)}</b><small>${esc(x.sub)}</small>${x.note ? `<small class="faint">${esc(x.note)}</small>` : ""}</div>` + rows;
}
function renderTab() {
  if (!ST.cur) return; const t = ST.tab;
  if (t === "pso") renderPSO(); if (t === "rdo") renderRDO(); if (t === "th") renderTH(); if (t === "ren") renderREN();
  if (t === "ipsi") renderIPSI(); if (t === "post" && typeof renderPOST === "function") renderPOST(); if (t === "cmp") renderCMP(); if (t === "chat") renderChat();
  if (["pso", "ipsi", "cmp"].includes(t)) renderAIBar(t); else if (["rdo", "th", "ren"].includes(t)) renderAIBarRdo(t);
  if (ST.moment) renderMoment();
  try { renderTipos(); } catch (e) {}
}
document.addEventListener("click", e => { if (e.target.closest("[data-pday]")) setTimeout(() => { try { renderTipos(); } catch (x) {} }, 0); });

const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
function dailyDay() { // día elegido (último clic); si no hay, hoy; si hoy no está en la semana, el último día con PDO o reprogramas
  const dd = days(ST.cur), sel = [...ST.F.days].filter(d => dd.includes(d));
  if (sel.length) return {day: sel[sel.length - 1]};
  const td = todayISO(); if (dd.includes(td)) return {day: td};
  const alt = [...dd].reverse().find(d => ST.pdo?.[d] || ST.reprog[d]) || dd[0];
  return {day: alt, note: `Hoy no está en la semana ${ST.cur.semana}: se muestra el ${dayLabel(alt)}/${alt.slice(5, 7)}.`};
}

/* ============ Versión web (GitHub/Cloudflare): datos desde data/boot.json, sin IA por ahora ============ */
(() => {
  const st = document.createElement("style");
  st.textContent = `.aibar{display:none!important}#navTabs [data-tab="chat"]{display:none}.brand{flex-wrap:wrap}.brand .upd{flex-basis:100%;font-family:var(--f-mono);font-size:11px;color:var(--fg-3);font-weight:400;margin-top:2px}`;
  document.head.appendChild(st);
  if (BOOT.actualizado) { const [d, h] = BOOT.actualizado.split("T"); document.querySelector(".brand")?.insertAdjacentHTML("beforeend", `<small class="upd" title="Última revisión del portal del COES">Actualizado ${d.slice(8)}/${d.slice(5, 7)} ${h}</small>`); }
  // abrir en la semana que contiene hoy
  const hoy = typeof todayISO === "function" ? todayISO() : new Date().toISOString().slice(0, 10);
  for (const [k, v] of WEEKS) if (v.w.inicio <= hoy && hoy <= v.w.fin) { ST.key = k; break; }
})();

function baseOpts(w) { return tsOpts(w.times); } // compatibilidad (Intervenciones)
const refilter = () => { VWsig = ""; renderAll(); };
document.querySelectorAll("#navTabs button").forEach(b => b.addEventListener("click", () => setTab(b.dataset.tab)));
$("wkPrev").onclick = () => { const ks = sortedKeys(), i = ks.indexOf(ST.key); if (i > 0) setWeek(ks[i - 1]); };
$("wkNext").onclick = () => { const ks = sortedKeys(), i = ks.indexOf(ST.key); if (i < ks.length - 1) setWeek(ks[i + 1]); };
$("wkSel").onchange = e => setWeek(e.target.value);
$("addFiles").onclick = () => $("fileAny").click();
$("addPrev").onclick = () => { $("fileWeek").dataset.mode = "prev"; $("fileWeek").click(); };
$("fileAny").onchange = e => { const f = [...e.target.files]; e.target.value = ""; if (f.length) handleFiles(f).then(() => { VWsig = ""; renderAll(); }); };
$("fileWeek").onchange = e => { const f = [...e.target.files]; e.target.value = ""; if (f.length) handleFiles(f, e.target.dataset.mode).then(() => { VWsig = ""; renderAll(); }); };
$("fSource").onclick = () => { const p = $("fSourcePop"); p.hidden = !p.hidden; $("fSource").setAttribute("aria-expanded", !p.hidden); };
$("fSourcePop").addEventListener("change", e => { const T = e.target, S = ST.F.src;
  if (T.dataset.all != null) S.clear();
  else if (T.dataset.tipo) { const ss = SRC_TREE().find(g => g.tipo === T.dataset.tipo).srcs.map(x => x.s); ss.forEach(s => T.checked ? S.add(s) : S.delete(s)); }
  else if (T.dataset.src) T.checked ? S.add(T.dataset.src) : S.delete(T.dataset.src);
  if (S.size === SRC_TREE().reduce((n, g) => n + g.srcs.length, 0)) S.clear(); refilter(); });
$("fCentral").addEventListener("change", e => { const v = e.target.value.trim(), ok = !v || v.startsWith("Empresa: ") || ST.cur.centrales[v]; if (ok) { ST.F.central = v; refilter(); } else toast("No encontré esa central o empresa.", true); });
$("fClear").onclick = () => { ST.F = {days: new Set(), src: new Set(), central: ""}; refilter(); };
document.addEventListener("click", e => {
  const T = e.target;
  if (T.closest("[data-pick]")) $("fileAny").click();
  if (T.closest("#pasteChip")) { $("pasteTxt").value = ""; $("pasteDlg").showModal(); }
  const r = T.closest("[data-run]"); if (r && !r.disabled) runAI(r.dataset.run);
  const s = T.closest("[data-stop]"); if (s) aiState(s.dataset.stop).ctl?.abort();
  const tr = T.closest("tr[data-central]"); if (tr) { ST.dropSel = tr.dataset.central; renderIPSI(); }
  const tf = T.closest("tr[data-filter]"); if (tf) { ST.F.central = ST.F.central === tf.dataset.filter ? "" : tf.dataset.filter; refilter(); }
  const td = T.closest("tr[data-day]"); if (td) { ST.F.days = new Set([td.dataset.day]); refilter(); }
  const dc = T.closest("#dayChips button"); if (dc) { const d = dc.dataset.d; ST.F.days.has(d) ? ST.F.days.delete(d) : ST.F.days.add(d); refilter(); }
  const bs = T.closest("#baseFlow button[data-v]"); if (bs && !bs.disabled) { ST.baseMode = bs.dataset.v; refilter(); }
  const tw = T.closest("#fSourcePop [data-tw]"); if (tw) { const k = tw.dataset.tw; ST.srcClosed.has(k) ? ST.srcClosed.delete(k) : ST.srcClosed.add(k); renderSrcTree(); }
  if (!T.closest(".dd") && !$("fSourcePop").hidden) { $("fSourcePop").hidden = true; $("fSource").setAttribute("aria-expanded", "false"); }
});
document.addEventListener("keydown", e => {
  if ((e.key === "Enter" || e.key === " ") && e.target.matches?.(".src[role=button]")) { e.preventDefault(); e.target.click(); }
  if (e.key === "Escape" && !$("fSourcePop").hidden) { $("fSourcePop").hidden = true; $("fSource").focus(); return; }
  if (ST.moment && !e.target.matches?.("input,textarea,select")) { if (e.key === "ArrowRight") { stepMoment(1); e.preventDefault(); } if (e.key === "ArrowLeft") { stepMoment(-1); e.preventDefault(); } if (e.key === "Escape") setMoment(null); }
});
$("mPrev").onclick = () => stepMoment(-1); $("mNext").onclick = () => stepMoment(1); $("mClose").onclick = () => { if (playH) togglePlay(); setMoment(null); };
$("mPlay").onclick = togglePlay; $("mAsk").onclick = askMoment; $("extractBtn").onclick = extractTables;
$("pasteDlg").addEventListener("close", () => { const v = $("pasteDlg").returnValue, t = $("pasteTxt").value.trim();
  if ((v === "ipso" || v === "ipsi") && t) { ST.wk[v] = t; ST.wk[v + "Name"] = "texto pegado"; ST.wk.ai = {}; persist(); renderAll(); toast(`${v.toUpperCase()} → semana ${ST.cur.semana}`); } });
let dragN = 0;
document.addEventListener("dragenter", e => { e.preventDefault(); dragN++; document.body.classList.add("dropping"); });
document.addEventListener("dragleave", () => { if (--dragN <= 0) { dragN = 0; document.body.classList.remove("dropping"); } });
document.addEventListener("dragover", e => e.preventDefault());
document.addEventListener("drop", e => { e.preventDefault(); dragN = 0; document.body.classList.remove("dropping"); const f = [...(e.dataTransfer?.files || [])]; if (f.length) handleFiles(f).then(() => { VWsig = ""; renderAll(); }); });
$("chatForm").addEventListener("submit", e => { e.preventDefault(); const q = $("chatIn").value.trim(); if (!q || !SAMPLE) return; $("chatIn").value = ""; sendChat(q); });
$("chatIn").addEventListener("keydown", e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); $("chatForm").requestSubmit(); } });
window.matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change", renderTab);
new MutationObserver(renderTab).observe(document.documentElement, {attributes: true, attributeFilter: ["data-theme"]});
if (typeof Chart !== "undefined" && Chart.defaults) Chart.defaults.font.family = "IBM Plex Sans";
let startTab = "pso"; try { startTab = localStorage.getItem("pso-tab") || "pso"; } catch (e) {}
setTab(TITLES[startTab] ? startTab : "pso");
(async () => { try { SAMPLE = window.claude ? await window.claude.use("sample") : null; } catch (e) { SAMPLE = null; } renderTab(); })();
