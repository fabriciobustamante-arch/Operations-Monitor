"""Descarga del portal del COES (www.coes.org.pe) los archivos que usa el monitor.

Ventana: la semana operativa actual y la anterior (sábado a viernes), y sus días hasta mañana.
Solo descarga lo que falta; un día ya cerrado (con PDO, IDCOS e IEOD) no se vuelve a consultar.
Deja todo en raw/ con la estructura que espera procesar.py y escribe raw/hay_cambios si bajó algo nuevo.
"""
import datetime as dt, html, io, json, re, shutil, sys, time, urllib.parse, urllib.request, zipfile
from pathlib import Path

PORTAL = "https://www.coes.org.pe/Portal/browser/"
CABECERAS = {"User-Agent": "Mozilla/5.0 (monitor-pso; descarga automatica)", "X-Requested-With": "XMLHttpRequest"}
RAIZ = Path(__file__).resolve().parent.parent
RAW = RAIZ / "raw"
ESTADO = RAW / "estado.json"

PS = "Operación/Programa de Operación/Programa Semanal Operación/"
MS = "Operación/Programa de Mantenimiento/Programa Semanal/"
PD = "Operación/Programa de Operación/Programa Diario/"
RP = "Operación/Programa de Operación/Reprograma Diario Operación/"
ID = "Post Operación/Reportes/IDCOS/"
IE = "Post Operación/Reportes/IEOD/"

consultas = 0
nuevos = []


# ---------------------------------------------------------------- portal
def listar(base, rel):
    """Nombres (carpetas y archivos) de una carpeta del portal."""
    global consultas
    consultas += 1
    datos = urllib.parse.urlencode({"baseDirectory": base, "url": rel, "indicadorHeader": "S", "tipoVista": "B", "orderFolder": "D", "origen": "Local"}).encode()
    req = urllib.request.Request(PORTAL + "vistadatos", data=datos, headers={**CABECERAS, "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    for intento in range(3):
        try:
            h = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
            break
        except Exception:
            if intento == 2:
                raise
            time.sleep(5)
    nombres = []
    for fila in re.findall(r"<tr[^>]*>(.*?)</tr>", h, re.S):
        celdas = [html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", fila, re.S)]
        celdas = [c for c in celdas if c]
        if celdas:
            nombres.append(celdas[0])
    time.sleep(0.3)  # sin apuro: no cargar el servidor del COES
    return nombres


def bajar(ruta_portal, destino: Path):
    global consultas
    consultas += 1
    req = urllib.request.Request(PORTAL + "download?url=" + urllib.parse.quote(ruta_portal, safe=""), headers=CABECERAS)
    datos = urllib.request.urlopen(req, timeout=180).read()
    if destino.suffix.lower() in (".xlsx", ".zip") and datos[:2] != b"PK":
        raise RuntimeError(f"{destino.name}: el COES no devolvió un archivo válido")
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(datos)
    nuevos.append(str(destino.relative_to(RAW)))
    print("  +", destino.relative_to(RAW), f"({len(datos) / 1e3:,.0f} kB)")
    return datos


def traer(base, carpeta, destino: Path, patron, ya=None):
    """Descarga de una carpeta del portal los archivos que calzan con el patrón y aún no están."""
    try:
        nombres = listar(base, carpeta)
    except Exception as e:
        print("  ! no se pudo listar", carpeta, e); return []
    out = []
    for n in nombres:
        if not re.search(patron, n, re.I):
            continue
        f = destino / n
        if not f.exists():
            try:
                bajar(carpeta + n, f)
            except Exception as e:
                print("  !", n, e); continue
        out.append(f)
    return out


def sub(base, carpeta, regex):
    """Primera subcarpeta cuyo nombre calza con la expresión."""
    try:
        return next((n for n in listar(base, carpeta) if re.search(regex, n, re.I)), None)
    except Exception as e:
        print("  ! no se pudo listar", carpeta, e); return None


# ---------------------------------------------------------------- ventana de fechas
def semanas_objetivo(hoy):
    out = []
    for anio in sorted({hoy.year, (hoy - dt.timedelta(days=14)).year}, reverse=True):
        for n in listar(PS, f"{PS}{anio}/"):
            m = re.match(r"(\d+)_SEMANAL N°\s*(\d+)\s*\((\d{2})\.(\d{2})\.(\d{4})\s*-\s*(\d{2})\.(\d{2})\.(\d{4})\)", n)
            if m:
                ini = dt.date(int(m[5]), int(m[4]), int(m[3]))
                if ini <= hoy:
                    out.append({"num": int(m[2]), "anio": anio, "carpeta": n, "ini": ini, "fin": ini + dt.timedelta(days=6)})
    out.sort(key=lambda s: s["ini"], reverse=True)
    return out[:2]


MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def carpeta_dia(base, d, sin_dia=False):
    """Ruta de la carpeta del día en el portal (los nombres de mes y día varían entre secciones)."""
    mes = sub(base, f"{base}{d.year}/", rf"^{d.month:02d}_")
    if not mes:
        return None
    dia = sub(base, f"{base}{d.year}/{mes}/", rf"^(Día\s*)?{d.day:02d}$")
    return f"{base}{d.year}/{mes}/{dia}/" if dia else None


# ---------------------------------------------------------------- principal
def main():
    estado = json.loads(ESTADO.read_text("utf-8")) if ESTADO.exists() else {"dias_cerrados": [], "reprogramas": {}}
    ahora = dt.datetime.now(dt.timezone(dt.timedelta(hours=-5)))  # hora de Lima
    hoy = ahora.date()
    semanas = semanas_objetivo(hoy)
    if not semanas:
        print("No encontré semanas en el portal"); sys.exit(1)
    print("Semanas:", ", ".join(f"Sem{s['num']}" for s in semanas))

    # 1) Programa semanal (Final) e intervenciones (Final)
    for s in semanas:
        dst = RAW / "semanal" / f"Sem{s['num']:02d}"
        if not list((dst / "PSO").glob("Anexo1_Despacho_*.xlsx")) or not list((dst / "PSO").glob("SPR-IPSO*.pdf")):
            if sub(PS, f"{PS}{s['anio']}/{s['carpeta']}/", r"^Final$"):
                traer(PS, f"{PS}{s['anio']}/{s['carpeta']}/Final/", dst / "PSO", r"^(Anexo1_Despacho_Sem\d+\.xlsx|SPR-IPSO.*\.pdf)$")
        if not list((dst / "IPSI").glob("Anexo1_Intervenciones*.zip")) or not list((dst / "IPSI").glob("SPR-IPSI*.pdf")):
            ms = sub(MS, f"{MS}{s['anio']}/", rf"^{s['num']:02d}_")
            if ms and sub(MS, f"{MS}{s['anio']}/{ms}/", r"^Final$"):
                traer(MS, f"{MS}{s['anio']}/{ms}/Final/", dst / "IPSI", r"^(Anexo1_Intervenciones.*\.zip|SPR-IPSI.*\.pdf)$")

    # 2) Días: programa diario, reprogramas, IDCOS e IEOD
    ini = min(s["ini"] for s in semanas)
    fin = min(max(s["fin"] for s in semanas), hoy + dt.timedelta(days=1))
    d = ini
    while d <= fin:
        clave = d.isoformat()
        dst = RAW / "diario" / d.strftime("%Y%m%d")
        if clave in estado["dias_cerrados"]:
            d += dt.timedelta(days=1); continue
        print(clave)
        if not list((dst / "PDO").glob("Anexo1_Despacho_*.xlsx")):
            c = carpeta_dia(PD, d)
            if c:
                traer(PD, c, dst / "PDO", r"^(Anexo1_Despacho_\d{8}\.xlsx|SPR-IPDO.*\.pdf)$")
        if d <= hoy:
            c = carpeta_dia(RP, d)
            if c:
                hechos = set(estado["reprogramas"].get(clave, []))
                for rep in listar(RP, c):
                    if rep in hechos or "." in rep:
                        continue
                    ok = traer(RP, f"{c}{rep}/", dst / "Reprog", r"^(Reprog_\d{4}[A-Z]+\.xlsx|indices\d{8}_?[A-Z]+\.xlsx)$")
                    if any(f.name.startswith("Reprog_") for f in ok):
                        hechos.add(rep)
                estado["reprogramas"][clave] = sorted(hechos)
        if d < hoy:
            if not list((dst / "IDCOS").glob("*.pdf")):
                c = carpeta_dia(ID, d)
                if c:
                    traer(ID, c, dst / "IDCOS", r"\.pdf$")
            if not list((dst / "IEOD").glob("AnexoA_*.xlsx")) or not list((dst / "IEOD").glob("CMgCP*.xlsx")):
                c = carpeta_dia(IE, d)
                if c:
                    traer(IE, c, dst / "IEOD", r"^AnexoA_\d+\.xlsx$")
                    for z in traer(IE, c, dst / "IEOD", r"^CMg\d{8}\.zip$"):
                        with zipfile.ZipFile(z) as zf:  # el zip trae PSS/E y GAMS: solo se guarda el CMgCP
                            for n in zf.namelist():
                                if re.search(r"CMgCP.*\.xlsx$", n, re.I):
                                    (dst / "IEOD" / Path(n).name).write_bytes(zf.read(n))
                        z.unlink()
            completo = all(list((dst / sub_).glob(p)) for sub_, p in (("PDO", "Anexo1_Despacho_*.xlsx"), ("IDCOS", "*.pdf"), ("IEOD", "AnexoA_*.xlsx"), ("IEOD", "CMgCP*.xlsx")))
            if completo:
                estado["dias_cerrados"].append(clave)
        d += dt.timedelta(days=1)

    # 3) Limpieza: fuera de la ventana no se guarda nada
    nums = {f"Sem{s['num']:02d}" for s in semanas}
    for p in (RAW / "semanal").glob("Sem*") if (RAW / "semanal").exists() else []:
        if p.name not in nums:
            shutil.rmtree(p); nuevos.append("borrado " + p.name)
    for p in (RAW / "diario").glob("2*") if (RAW / "diario").exists() else []:
        f = dt.date(int(p.name[:4]), int(p.name[4:6]), int(p.name[6:]))
        if f < ini or f > fin:
            shutil.rmtree(p); nuevos.append("borrado " + p.name)
    estado["dias_cerrados"] = sorted(x for x in set(estado["dias_cerrados"]) if x >= ini.isoformat())
    estado["reprogramas"] = {k: v for k, v in estado["reprogramas"].items() if k >= ini.isoformat()}
    RAW.mkdir(exist_ok=True)
    ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=1), "utf-8")
    (RAW / "hay_cambios").write_text("1" if nuevos else "0")
    print(f"\n{len(nuevos)} cambios · {consultas} consultas al portal")


if __name__ == "__main__":
    main()
