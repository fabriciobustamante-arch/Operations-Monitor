# Operations-Monitor · Monitor PSO

Monitor del Programa Semanal de Operación (PSO), el programa diario (PDO), los reprogramas y la post operación del SEIN, con datos públicos del portal del COES.

## Cómo funciona
1. **GitHub Actions** (`.github/workflows/actualizar.yml`) corre cada 30 minutos.
2. `scripts/coes.py` revisa el portal del COES y descarga solo lo nuevo de la semana actual y la anterior:
   PSO e IPSO (versión Final), IPSI y Anexo 1 de intervenciones (Final), PDO e IPDO, reprogramas e índices, IDCOS e IEOD (AnexoA y CMgCP).
3. `scripts/procesar.py` convierte esos archivos en `data/boot.json` usando los mismos lectores del monitor (`scripts/js/`, ejecutados con `scripts/lectores.mjs`).
4. La página (`index.html` + `app.js`) se publica en GitHub Pages y lee `data/boot.json`.

Si el COES no publicó nada nuevo, no se procesa ni se publica nada.

## Archivos
| Archivo | Qué es |
|---|---|
| `index.html`, `app.js` | El monitor |
| `data/maestro.json` | Clasificación de centrales (empresa, tipo, fuente). Editable |
| `scripts/coes.py` | Descarga del portal del COES |
| `scripts/procesar.py` | Arma `data/boot.json` |
| `scripts/lectores.mjs`, `scripts/js/` | Lectores de Excel del monitor, reutilizados |

## Correr en local (opcional)
```
pip install -r requirements.txt
python scripts/coes.py
python scripts/procesar.py
python -m http.server 8000      # y abrir http://localhost:8000
```
