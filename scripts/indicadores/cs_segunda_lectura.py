#!/usr/bin/env python3
"""Soporte común de CS.VRC y CS.RPO para la segunda lectura (no es un indicador).

- muestra_20(): toma el 20 % (redondeado hacia arriba) de las filas de una tabla de clasificación con semilla fija
  random.Random(20261109) y escribe outputs/indicadores/segunda_lectura/muestra_<tema>.csv (CSV ";" UTF-8 con BOM) con las
  mismas columnas más clasificacion_lectora_2, acuerdo_si_no y comentario (vacías, para la usuaria).
  Si el archivo ya existe y la usuaria llenó clasificacion_lectora_2, NO se reescribe (se avisa).
- escribir_leeme(): LEEME.md con las instrucciones para llenar las muestras (sin datos personales).
Uso: lo importan cs_vrc.py y cs_rpo.py.
"""
import math
import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C

SEMILLA = 20261109
FRACCION = 0.20
DIR_LECTURA = C.DIR_SALIDA / "segunda_lectura"
COLS_LECTORA = ["clasificacion_lectora_2", "acuerdo_si_no", "comentario"]


def indices_muestra(n: int) -> list[int]:
    k = math.ceil(FRACCION * n)
    return sorted(random.Random(SEMILLA).sample(range(n), k))


def muestra_20(filas: list[dict], columnas: list[str], tema: str) -> tuple[Path, int, bool]:
    """Escribe muestra_<tema>.csv. Devuelve (ruta, n_filas_muestra, escrito)."""
    ruta = DIR_LECTURA / f"muestra_{tema}.csv"
    if ruta.exists():
        previas = C.leer_csv(ruta)
        if any((r.get("clasificacion_lectora_2") or "").strip() for r in previas):
            print(f"AVISO: {ruta} ya tiene respuestas en clasificacion_lectora_2; NO se reescribe.")
            return ruta, len(previas), False
    ordenadas = sorted(filas, key=lambda r: str(r["id"]))
    sel = [ordenadas[i] for i in indices_muestra(len(ordenadas))]
    salida = [{**{c: r.get(c) for c in columnas}, **{c: "" for c in COLS_LECTORA}} for r in sel]
    C.escribir_csv(ruta, salida, columnas + COLS_LECTORA)
    return ruta, len(salida), True


LEEME = """# Segunda lectura de las clasificaciones de CS.VRC y CS.RPO

Estos archivos permiten que una segunda persona revise, de forma independiente, una muestra de las clasificaciones hechas por el
Ejecutor (un solo lector). Sirven para estimar el acuerdo entre lectores antes de usar los resultados.

## Archivos
- `muestra_vrc.csv`: 20 % de los eventos de rendición de cuentas de `outputs/indicadores/parciales/CS/clasificacion_vrc.csv`.
- `muestra_rpo.csv`: 20 % de las observaciones ciudadanas de `outputs/indicadores/parciales/CS/clasificacion_rpo.csv`.
- Formato: CSV con separador `;` y codificación UTF-8 con BOM (ábralos en Excel con "Datos > Desde texto/CSV").
- Selección: `random.Random(20261109)`, 20 % de cada tabla redondeado hacia arriba. Es reproducible: con la misma tabla se obtiene la misma muestra.

## Cómo llenarlas
1. Para cada fila, abra el documento indicado (columnas `documento`, `pagina` y `enlace`; la ruta local está en `data_raw_IIC/`) y lea la cita y el contexto.
2. **Sin mirar** las columnas de banderas ni la clasificación del primer lector, escriba en `clasificacion_lectora_2` su propia clasificación:
   - **VRC** (evento): `mencion_veedurias` (si / no / desconocida) y `documento_integral` (1 o 0), por ejemplo `si; integral=1`.
     Regla: `si` = el documento del evento menciona veedurías presentes; `no` = documento integral (acta o memoria con asistentes) sin mención;
     `desconocida` = documento no integral (cuadro de preguntas, presentación, justificación) sin mención.
   - **RPO** (observación): tipo (`pregunta`, `sugerencia`, `felicitacion`, `tramite_individual`), `tema_pdm` (1 o 0) y respuesta de fondo
     (`estricto=1/0` y `laxo=1/0`), por ejemplo `pregunta; tema_pdm=1; estricto=1; laxo=1`.
     Regla estricta: la respuesta atiende el asunto con un dato, una decisión, un compromiso con fecha o una negativa motivada; no cuentan el
     acuse, el agradecimiento, "se tendrá en cuenta" ni la remisión sin contenido. Regla laxa: cualquier texto de respuesta salvo acuse explícito.
3. Después de clasificar, compare con las columnas del primer lector y escriba en `acuerdo_si_no` `si` o `no`.
4. Si hay desacuerdo, explique en `comentario` qué regla aplicó y por qué (sin nombres de personas, cédulas, teléfonos ni correos).
5. Guarde el archivo con el mismo nombre y formato (`;` y UTF-8 con BOM).

## Qué pasa después
- Los scripts `cs_vrc.py` y `cs_rpo.py` no sobrescriben `muestra_*.csv` si ya tienen respuestas en `clasificacion_lectora_2` (avisan por pantalla).
- El acuerdo = filas con `acuerdo_si_no = si` / filas revisadas. Los desacuerdos se llevan a los profesores como decisión metodológica.
- Las clasificaciones dudosas del primer lector están marcadas con `dudoso = 1` y `nota_dudoso` en las tablas de clasificación completas.

## Datos personales
Los documentos originales (cuadros de preguntas y cartas) contienen nombres, teléfonos y correos de ciudadanos. Las tablas de este
repositorio no los reproducen: las citas reemplazan los nombres por `[ciudadano/a]` o `[funcionario/a]`. Haga lo mismo en `comentario`.
"""


def escribir_leeme() -> Path:
    DIR_LECTURA.mkdir(parents=True, exist_ok=True)
    ruta = DIR_LECTURA / "LEEME.md"
    ruta.write_text(LEEME, encoding="utf-8")
    return ruta
