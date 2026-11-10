"""Convierte los archivos del COES descargados en raw/ en data/boot.json, que es lo que lee el monitor.

Estructura esperada (la crea coes.py):
  raw/semanal/SemNN/PSO/    Anexo1_Despacho_SemNN.xlsx, SPR-IPSO-*.pdf
  raw/semanal/SemNN/IPSI/   SPR-IPSI-*.pdf, Anexo1_Intervenciones_SemNN.zip
  raw/diario/AAAAMMDD/PDO/    Anexo1_Despacho_AAAAMMDD.xlsx, SPR-IPDO-*.pdf
  raw/diario/AAAAMMDD/Reprog/ Reprog_DDMMX.xlsx, indicesAAAAMMDD_X.xlsx
  raw/diario/AAAAMMDD/IEOD/   AnexoA_DDMM.xlsx, CMgCP*.xlsx
  raw/diario/AAAAMMDD/IDCOS/  <día>_DDMM.pdf

Los Excel se vuelcan a JSON con Python y se leen con los mismos lectores del monitor (scripts/lectores.mjs).
"""
import datetime as dt, hashlib, io, json, os, re, subprocess, sys, tempfile, unicodedata, zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
RAW, CACHE, DATA = RAIZ / "raw", RAIZ / "cache", RAIZ / "data"
CACHE.mkdir(exist_ok=True); DATA.mkdir(exist_ok=True)


# ---------------------------------------------------------------- volcado de archivos
def _celda(v):
    if isinstance(v, dt.datetime):
        return {"d": v.strftime("%Y-%m-%dT%H:%M:%S")}
    if isinstance(v, dt.date):
        return {"d": v.strftime("%Y-%m-%dT00:00:00")}
    if isinstance(v, dt.time):
        return {"t": (v.hour * 3600 + v.minute * 60 + v.second) / 86400}
    if isinstance(v, float):
        return round(v, 4)
    return v


def _recortar(fila):
    fila = list(fila)
    while fila and fila[-1] is None:
        fila.pop()
    return fila


def volcar_xlsx(datos: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
    hojas = {}
    for ws in wb.worksheets:
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        hojas[ws.title] = [_recortar(_celda(v) for v in fila) for fila in ws.iter_rows(values_only=True)]
    return {"SheetNames": list(hojas), "Sheets": hojas}


def volcar_xls(datos: bytes):
    try:
        import xlrd
    except ImportError:  # sin xlrd: se convierte con LibreOffice
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "x.xls"; src.write_bytes(datos)
            subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, str(src)], capture_output=True, timeout=300)
            return volcar_xlsx((Path(tmp) / "x.xlsx").read_bytes())
    wb = xlrd.open_workbook(file_contents=datos)
    hojas = {}
    for sh in wb.sheets():
        filas = []
        for i in range(sh.nrows):
            fila = []
            for c in sh.row(i):
                if c.ctype == xlrd.XL_CELL_DATE:
                    fila.append(_celda(xlrd.xldate.xldate_as_datetime(c.value, wb.datemode)))
                elif c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    fila.append(None)
                else:
                    fila.append(_celda(c.value))
            filas.append(_recortar(fila))
        hojas[sh.name] = filas
    return {"SheetNames": list(hojas), "Sheets": hojas}


def texto_pdf(datos: bytes):
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(datos))
    return "\n\n".join(f"[Página {i + 1}]\n" + (p.extract_text() or "") for i, p in enumerate(r.pages))


def volcado(ruta: Path):
    """Volcado con memoria: solo se procesa de nuevo si el archivo cambió."""
    datos = ruta.read_bytes()
    clave = hashlib.sha1(datos).hexdigest()[:16]
    f = CACHE / f"{clave}.json"
    if f.exists():
        return json.loads(f.read_text("utf-8"))
    n = ruta.name.lower()
    if n.endswith(".xlsx"):
        v = volcar_xlsx(datos)
    elif n.endswith(".xls"):
        v = volcar_xls(datos)
    elif n.endswith(".pdf"):
        v = texto_pdf(datos)
    else:
        raise ValueError(ruta.name)
    f.write_text(json.dumps(v, ensure_ascii=False, separators=(",", ":")), "utf-8")
    return v


# ---------------------------------------------------------------- utilidades
def nz(s):
    return unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper().strip()


def redondear(o, d=3):
    if isinstance(o, float):
        return round(o, d)
    if isinstance(o, list):
        return [redondear(x, d) for x in o]
    if isinstance(o, dict):
        return {k: redondear(v, d) for k, v in o.items()}
    return o


def archivos(carpeta: Path, patron):
    return sorted(p for p in carpeta.glob("*") if p.is_file() and re.search(patron, p.name, re.I)) if carpeta.exists() else []


# ---------------------------------------------------------------- IPSI: Anexo1_Intervenciones (hoja MANTTOS)
def intervenciones(zip_path: Path, centrales_semana, maestro):
    """Devuelve la lista de intervenciones con el formato que usa el monitor (ipsiItems)."""
    with zipfile.ZipFile(zip_path) as z:
        nombre = next((n for n in z.namelist() if re.search(r"Agentes\)?\.xlsx?$", n, re.I)), None) \
            or next((n for n in z.namelist() if re.search(r"\.xlsx?$", n, re.I)), None)
        if not nombre:
            return []
        tmp = CACHE / ("interv_" + zip_path.stem + Path(nombre).suffix)
        tmp.write_bytes(z.read(nombre))
    wb = volcado(tmp)
    filas = wb["Sheets"].get("MANTTOS") or next(iter(wb["Sheets"].values()))
    hi = next((i for i, r in enumerate(filas) if r and "ITEM" in [str(x).strip().upper() for x in r if x is not None] and any(str(x).strip().upper() == "EMPRESA" for x in r if x is not None)), None)
    if hi is None:
        return []
    H = [nz(x) for x in filas[hi]]
    c = lambda *k: next((i for i, h in enumerate(H) if any(h.startswith(x) for x in k)), None)
    iE, iU, iQ, iI, iF, iD, iMW, iDisp, iProv = c("EMPRESA"), c("UBICACION"), c("EQUIPO"), c("INICIO"), c("FINAL"), c("DESCRIPCION"), c("MW INDISP"), c("DISPON"), c("INST. PROV", "INST PROV")
    # columnas del despacho por central (para cruzar con la generación)
    cen_maestro = {k: nz(v.get("central") or k) for k, v in maestro.items()}
    PREF = r"^(CENTRAL (EOLICA|SOLAR|HIDROELECTRICA|TERMICA|TERMOELECTRICA) |C\.?H\.? |C\.?T\.? |CTB |PCH |CS |CE |RESERVA FRIA( DE GENERACION)? |EL |LA )+"
    sin_esp = lambda x: re.sub(r"[^A-Z0-9]", "", nz(x))

    def centrales_de(ubic, equipo):
        u = re.sub(PREF, "", nz(ubic)).strip()
        if len(u) < 3:
            return []
        cand = [k for k in centrales_semana if nz(k).startswith(u) or cen_maestro.get(k, "").startswith(u) or nz(k) == u]
        eq = sin_esp(re.sub(r"\b" + re.escape(u) + r"\b", "", nz(equipo)))
        con_eq = [k for k in cand if eq and eq in sin_esp(k)]
        return con_eq or cand

    def iso(v):
        if isinstance(v, dict) and "d" in v:
            return v["d"][:16]
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?", str(v or "").strip())
        return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}T{int(m[4] or 0):02d}:{m[5] or '00'}" if m else None

    grupos = {}
    for r in filas[hi + 1:]:
        if not r or len(r) <= max(x for x in (iE, iU, iQ, iI, iF) if x is not None):
            continue
        ini, fin = iso(r[iI]), iso(r[iF])
        if not ini or not fin or not r[iQ]:
            continue
        disp = nz(r[iDisp]) if iDisp is not None and iDisp < len(r) else ""
        mw = r[iMW] if iMW is not None and iMW < len(r) and isinstance(r[iMW], (int, float)) else 0
        prov = iProv is not None and iProv < len(r) and nz(r[iProv]) == "SI"
        equipo, ubic = str(r[iQ]).strip(), str(r[iU] or "").strip()
        eqn = nz(equipo)
        es_unidad = bool(re.match(r"^(G|TG|TV|U|GP|CC|GE|RF-)\S*\d", eqn)) or eqn == "CENTRAL"
        cen = centrales_de(ubic, equipo) if es_unidad or mw > 0 else []
        if prov:
            tipo = "Conexión provisional"
        elif es_unidad and (disp == "F/S" or mw > 0):
            tipo = "Generación"
        elif cen and mw > 0 and not re.match(r"^(L|CL|T|SSEE|SS\.EE|BARRA)", eqn):
            tipo = "Generación"
        elif disp == "F/S" and re.match(r"^L-?[125]\d{3}", eqn):
            tipo = "Transmisión"
        else:
            continue
        desc = re.sub(r"\s+", " ", str(r[iD] or "")).strip()
        k = (str(r[iE] or "").strip(), ubic, equipo, desc)
        g = grupos.setdefault(k, {"tipo": tipo, "empresa": k[0], "equipo": equipo, "ubic": ubic, "centrales": cen if tipo == "Generación" else [], "ventanas": [], "desc": desc, "mw": mw})
        g["ventanas"].append([ini, fin])
        g["mw"] = max(g["mw"], mw)
    out = []
    for g in grupos.values():  # une ventanas contiguas (el anexo trae una fila por día)
        v = sorted(g["ventanas"]); u = [v[0]]
        for a, b in v[1:]:
            if a <= u[-1][1]:
                u[-1][1] = max(u[-1][1], b)
            else:
                u.append([a, b])
        g["ventanas"] = [[a, b.replace("T00:00", "T00:00") if b[11:] != "00:00" else (dt.datetime.fromisoformat(b) - dt.timedelta(minutes=1)).strftime("%Y-%m-%dT%H:%M")] for a, b in u]
        out.append(g)
    orden = {"Generación": 0, "Transmisión": 1, "Conexión provisional": 2}
    out.sort(key=lambda g: (orden[g["tipo"]], g["ventanas"][0][0]))
    return out


# ---------------------------------------------------------------- IDCOS: causas de los reprogramas (sección 13)
def causas_idcos(texto, fecha):
    sec = re.search(r"Reprogramaciones del d[ií]a(.*?)(?:\n\s*\d+\.\s+Observaciones|$)", texto, re.S | re.I)
    if not sec:
        return []
    num = (re.search(r"IDCOS\s+N[°º]\s*(\d+/\d{4})", texto) or [None, ""])[1]
    s = sec.group(1)
    marcas = list(re.finditer(r"(\d{1,2}:\d{2})\s+SCO\s*-\s*\d+\s*-\s*([A-Z])\b(?:\s*/\s*Sem\s*\d+)?", s))
    out = []
    for i, m in enumerate(marcas):
        fin = marcas[i + 1].start() if i + 1 < len(marcas) else len(s)
        causa = re.sub(r"\[Página \d+\]|IDCOS N° \d+/\d{4}|Sub-Dirección de Coordinación SCO/COES-SINAC|Página \d+ de \d+", " ", s[m.end():fin])
        causa = re.sub(r"\s+", " ", causa).strip().rstrip(".")
        out.append({"fecha": fecha, "operador": "", "causa": causa or "(sin motivo en el IDCOS)", "hora": m.group(1).zfill(5), "reprograma": m.group(2),
                    "publicacion": "", "fuente": f"IDCOS {num}".strip()})
    return out


# ---------------------------------------------------------------- principal
def main():
    maestro_rows = json.loads((DATA / "maestro.json").read_text("utf-8"))["maestro"]
    maestro = {r[0]: {"empresa": r[1], "tipo": r[2], "source": r[3], "central": r[4]} for r in maestro_rows}
    tareas, meta = [], []

    def tarea(tipo, ruta, **extra):
        t = {"tipo": tipo, "nombre": ruta.name}
        if tipo == "idcos":
            t["texto"] = volcado(ruta)
        else:
            t["wb"] = volcado(ruta)
        tareas.append(t); meta.append({"ruta": ruta, **extra})

    semanas = sorted((RAW / "semanal").glob("Sem*")) if (RAW / "semanal").exists() else []
    for s in semanas:
        for f in archivos(s / "PSO", r"^Anexo1_Despacho_Sem\d+\.xlsx$"):
            tarea("semanal", f, sem=s.name)
    dias = sorted((RAW / "diario").glob("2*")) if (RAW / "diario").exists() else []
    for d in dias:
        for f in archivos(d / "PDO", r"^Anexo1_Despacho_\d{8}\.xlsx$"):
            tarea("diario", f)
        for f in archivos(d / "Reprog", r"^Reprog_\d{4}[A-Z]+\.xlsx$"):
            tarea("diario", f)
        for f in archivos(d / "Reprog", r"^indices\d{8}_?[A-Z]+\.xlsx$"):
            tarea("indices", f)
        for f in archivos(d / "IEOD", r"^AnexoA_\d+\.xlsx$"):
            tarea("anexoa", f)
        for f in archivos(d / "IEOD", r"^CMgCP\d+\.xlsx$"):
            tarea("cmgcp", f)
        for f in archivos(d / "IDCOS", r"\.pdf$"):
            tarea("idcos", f)

    with tempfile.TemporaryDirectory() as tmp:
        ent, sal = Path(tmp) / "in.json", Path(tmp) / "out.json"
        ent.write_text(json.dumps({"maestro": maestro_rows, "tareas": tareas}, ensure_ascii=False), "utf-8")
        subprocess.run(["node", str(RAIZ / "scripts" / "lectores.mjs"), str(ent), str(sal)], check=True)
        res = json.loads(sal.read_text("utf-8"))

    boot = {"maestro": maestro_rows, "weeks": [], "pdo": {}, "reprog": {}, "indices": {}, "ipdo": {}, "listado": [], "limites": {}, "ipsiItems": {}, "restric": {}, "post": {}}
    errores = []
    for i, (r, m) in enumerate(zip(res, meta)):
        if not r["ok"]:
            errores.append(f'{r["nombre"]}: {r["error"]}'); continue
        x = r["r"]
        if r["tipo"] == "semanal":
            w = x; carpeta = m["ruta"].parent.parent
            w["fuente"] = f'{r["nombre"]} (COES, versión final)'
            doc = {"w": w}
            for sub, k in (("PSO", "ipso"), ("IPSI", "ipsi")):
                pdf = archivos(carpeta / sub, rf"SPR-{k.upper()}.*\.pdf$")
                if pdf:
                    doc[k] = volcado(pdf[0]); doc[k + "Name"] = pdf[0].name
            boot["weeks"].append(doc)
            zipi = archivos(carpeta / "IPSI", r"^Anexo1_Intervenciones.*\.zip$")
            if zipi:
                try:
                    boot["ipsiItems"][f'{w["anio"]}-{int(w["semana"]):02d}'] = intervenciones(zipi[0], list(w["centrales"]), maestro)
                except Exception as e:
                    errores.append(f"{zipi[0].name}: {e}")
        elif r["tipo"] == "diario":
            if x["kind"] == "pdo":
                boot["pdo"][x["date"]] = x["slots"]
            else:
                boot["reprog"].setdefault(x["date"], {})[x["letter"]] = x["slots"]
        elif r["tipo"] == "indices":
            boot["indices"].setdefault(x["date"], {})[x["letter"]] = x["rows"]
        elif r["tipo"] == "anexoa" and x.get("date"):
            boot["post"].setdefault(x["date"], {}).update(x)
        elif r["tipo"] == "cmgcp" and x and x.get("date"):
            boot["post"].setdefault(x["date"], {}).update({"cmg": x["cmg"], "cmgUnit": x["unidad"]})
        elif r["tipo"] == "idcos" and x.get("date"):
            texto = tareas[i]["texto"]
            boot["post"].setdefault(x["date"], {})["idcos"] = {k: x[k] for k in ("reps", "maxcap", "tension", "pruebas", "sorteo")} | {"name": r["nombre"]}
            boot["listado"] += causas_idcos(texto, x["date"])

    # IPDO (texto del informe del programa diario)
    for d in dias:
        for f in archivos(d / "PDO", r"IPDO.*\.pdf$"):
            fecha = f"{d.name[:4]}-{d.name[4:6]}-{d.name[6:]}"
            boot["ipdo"][fecha] = {"text": volcado(f), "name": f.name}

    boot["weeks"].sort(key=lambda x: x["w"]["inicio"])
    boot["listado"].sort(key=lambda x: (x["fecha"], x["reprograma"]))
    for k in ("pdo", "reprog", "post"):
        boot[k] = redondear(boot[k])
    for x in boot["weeks"]:
        x["w"] = redondear(x["w"])
    ahora = dt.datetime.now(dt.timezone(dt.timedelta(hours=-5)))
    boot["actualizado"] = ahora.strftime("%Y-%m-%dT%H:%M")
    (DATA / "boot.json").write_text(json.dumps(boot, ensure_ascii=False, separators=(",", ":")), "utf-8")
    resumen = {"actualizado": boot["actualizado"], "semanas": [x["w"]["semana"] for x in boot["weeks"]], "pdo": sorted(boot["pdo"]),
               "reprogramas": {d: "".join(sorted(v)) for d, v in sorted(boot["reprog"].items())}, "post": sorted(boot["post"]),
               "intervenciones": {k: len(v) for k, v in boot["ipsiItems"].items()}, "causas": len(boot["listado"]), "errores": errores}
    (DATA / "estado.json").write_text(json.dumps(resumen, ensure_ascii=False, indent=1), "utf-8")
    print(json.dumps(resumen, ensure_ascii=False, indent=1))
    print("boot.json:", round((DATA / "boot.json").stat().st_size / 1e6, 2), "MB")


if __name__ == "__main__":
    main()
