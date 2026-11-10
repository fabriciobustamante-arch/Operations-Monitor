// Ejecuta los MISMOS lectores del monitor (scripts/js/p1.js y p15.js) sobre los Excel ya volcados a JSON.
// Uso: node scripts/lectores.mjs entrada.json salida.json
//   entrada: {maestro: [...], tareas: [{tipo, nombre, wb?, texto?}]}
//   tipo: semanal | diario (PDO o reprograma) | indices | anexoa | cmgcp | idcos
import fs from "node:fs";
import vm from "node:vm";
import path from "node:path";

const [, , entrada, salida] = process.argv;
const IN = JSON.parse(fs.readFileSync(entrada, "utf8"));
const dir = path.join(path.dirname(new URL(import.meta.url).pathname), "js");

// ---- Celdas: el volcado guarda fechas como {d: "AAAA-MM-DDTHH:MM:SS"} y horas como {t: fracción}
const serial = iso => { const [a, b = "00:00:00"] = iso.split("T"); const [y, m, d] = a.split("-").map(Number); const [h, mi, s] = b.split(":").map(Number);
  return (Date.UTC(y, m - 1, d, h, mi, s || 0) - Date.UTC(1899, 11, 30)) / 864e5; };
const asDate = iso => { const [a, b = "00:00:00"] = iso.split("T"); const [y, m, d] = a.split("-").map(Number); const [h, mi, s] = b.split(":").map(Number); return new Date(y, m - 1, d, h, mi, s || 0); };
const conv = (v, dates) => {
  if (v && typeof v === "object") {
    if ("d" in v) return dates ? asDate(v.d) : serial(v.d);
    if ("t" in v) { if (!dates) return v.t; const mins = Math.round(v.t * 1440); return new Date(1899, 11, 30, Math.floor(mins / 60), mins % 60); }
  }
  return v;
};
const XLSX = {
  read: (wb, opts = {}) => ({SheetNames: wb.SheetNames, Sheets: Object.fromEntries(wb.SheetNames.map(n => [n, {rows: wb.Sheets[n], dates: !!opts.cellDates}]))}),
  utils: {sheet_to_json: (ws) => ws.rows.map(r => r.map(v => conv(v, ws.dates)))},
};

// ---- Contexto mínimo para que los archivos del monitor carguen fuera del navegador
const MESES = {enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12};
const ctx = vm.createContext({
  console, XLSX, Map, Set, Date, Math, JSON, Object, Array, String, Number, RegExp, Error, Promise, isNaN, parseFloat, parseInt,
  document: {getElementById: id => id === "boot" ? {textContent: JSON.stringify({maestro: IN.maestro, weeks: [], post: {}})} : null, documentElement: {}},
  getComputedStyle: () => ({getPropertyValue: () => ""}),
  localStorage: {getItem: () => null, setItem: () => {}},
  window: {},
  ST: {listado: [], post: {}},
  MESES,
  pad2: n => String(n).padStart(2, "0"),
});
vm.runInContext(`function dateFromSpanish(t) { const m = String(t).toLowerCase().match(/(\\d{1,2})\\s+de\\s+([a-záéíóú]+)\\s+de(?:l)?\\s+(\\d{4})/); return m && MESES[m[2]] ? \`\${m[3]}-\${String(MESES[m[2]]).padStart(2, "0")}-\${m[1].padStart(2, "0")}\` : null; }`, ctx);
for (const f of ["p1.js", "p15.js"]) vm.runInContext(fs.readFileSync(path.join(dir, f), "utf8"), ctx, {filename: f});

const file = (nombre, wb) => ({name: nombre, arrayBuffer: async () => wb});
const out = [];
for (const t of IN.tareas) {
  try {
    let r;
    if (t.tipo === "semanal") r = await ctx.parseAnexo(file(t.nombre, t.wb));
    else if (t.tipo === "diario") r = await ctx.parseDaily(file(t.nombre, t.wb));
    else if (t.tipo === "indices") r = await ctx.parseIndices(file(t.nombre, t.wb));
    else if (t.tipo === "anexoa") r = ctx.parseAnexoARows(await ctx.sheetsOf(file(t.nombre, t.wb)), t.nombre);
    else if (t.tipo === "cmgcp") { const S = await ctx.sheetsOf(file(t.nombre, t.wb)); r = ctx.parseCmgRows(S["Cmg_Barra"] || Object.values(S)[0]); }
    else if (t.tipo === "idcos") r = ctx.parseIdcosText(t.texto);
    out.push({nombre: t.nombre, tipo: t.tipo, ok: true, r: JSON.parse(JSON.stringify(r))});
  } catch (e) {
    out.push({nombre: t.nombre, tipo: t.tipo, ok: false, error: String(e.message || e)});
  }
}
fs.writeFileSync(salida, JSON.stringify(out));
console.log(`lectores: ${out.filter(x => x.ok).length} archivos leídos, ${out.filter(x => !x.ok).length} con error`);
