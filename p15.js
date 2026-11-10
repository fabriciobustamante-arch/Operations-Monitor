
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
