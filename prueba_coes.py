"""Prueba de acceso al portal del COES desde GitHub Actions (solo librería estándar)."""
import datetime, html, re, sys, urllib.parse, urllib.request

PORTAL = "https://www.coes.org.pe/Portal/browser/"
UA = {"User-Agent": "Mozilla/5.0 (monitor-pso prueba)", "X-Requested-With": "XMLHttpRequest"}
MESES = ["01_ENERO", "02_FEBRERO", "03_MARZO", "04_ABRIL", "05_MAYO", "06_JUNIO", "07_JULIO",
         "08_AGOSTO", "09_SETIEMBRE", "10_OCTUBRE", "11_NOVIEMBRE", "12_DICIEMBRE"]

def listar(base, rel):
    datos = urllib.parse.urlencode({"baseDirectory": base, "url": rel, "indicadorHeader": "S",
                                    "tipoVista": "B", "orderFolder": "D", "origen": "Local"}).encode()
    req = urllib.request.Request(PORTAL + "vistadatos", data=datos, headers={**UA,
          "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
    h = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    nombres = []
    for fila in re.findall(r"<tr[^>]*>(.*?)</tr>", h, re.S):
        celdas = [html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", fila, re.S)]
        celdas = [c for c in celdas if c]
        if celdas:
            nombres.append(celdas[0])
    return nombres

def descargar(ruta):
    req = urllib.request.Request(PORTAL + "download?url=" + urllib.parse.quote(ruta, safe=""), headers=UA)
    return urllib.request.urlopen(req, timeout=120).read()

def main():
    hoy = datetime.datetime.utcnow() - datetime.timedelta(hours=5)  # hora de Lima
    ok = True
    # 1) Programa semanal
    ps = "Operación/Programa de Operación/Programa Semanal Operación/"
    semanas = listar(ps, f"{ps}{hoy.year}/")
    print("Semanas publicadas:", semanas[:3])
    if not semanas:
        ok = False
    # 2) Reprogramas de hoy (o de ayer si aún no hay)
    rp = "Operación/Programa de Operación/Reprograma Diario Operación/"
    for d in (hoy, hoy - datetime.timedelta(days=1)):
        carpeta = f"{rp}{d.year}/{MESES[d.month - 1]}/Día {d.day:02d}/"
        reps = listar(rp, carpeta)
        print(f"Reprogramas del {d:%d/%m}:", reps)
        if reps:
            archivos = listar(rp, carpeta + reps[-1] + "/")
            print(f"  Archivos de {reps[-1]}:", archivos)
            xlsx = next((a for a in archivos if a.lower().startswith("reprog_")), None)
            if xlsx:
                b = descargar(carpeta + reps[-1] + "/" + xlsx)
                es_excel = b[:2] == b"PK"
                print(f"  Descarga de {xlsx}: {len(b):,} bytes · {'Excel válido' if es_excel else 'NO es un Excel'}")
                ok = ok and es_excel
            break
    else:
        ok = False
    print("\nRESULTADO:", "GitHub SÍ puede descargar del COES" if ok else "Algo falló: revisa los mensajes de arriba")
    sys.exit(0 if ok else 1)

if __name__ == "__main__":
    main()
