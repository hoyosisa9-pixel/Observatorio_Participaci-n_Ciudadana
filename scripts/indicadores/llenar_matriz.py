#!/usr/bin/env python3
"""Diligencia una COPIA de la matriz de los profesores con los valores del escenario por defecto.

Original (solo lectura): Documentos guias/Matriz de informacion de indicadores del observatorio ciudadano.xlsx
Salida: outputs/indicadores/matriz_diligenciada.xlsx (hojas Registro A:Q y PP_Proyectos A:M; nada más se toca).
Estado siempre "Pendiente" (o "No disponible" si no hay valor). Nunca "Validado".

Uso (desde la raíz, sin -I):
    python3 scripts/indicadores/llenar_matriz.py [--base DIR] [--salida RUTA]
    python3 scripts/indicadores/llenar_matriz.py --verificar [RUTA_SALIDA] [--zonas]

--verificar produce además una copia TEMPORAL con estado "Validado" (solo en el scratchpad de la sesión),
la recalcula con LibreOffice headless y compara la hoja Resultados con resultados_indicadores (escenario por
defecto); imprime las diferencias y borra la copia temporal.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C  # noqa: E402
import openpyxl  # noqa: E402

SCRATCH = Path("/tmp/claude-1000/-workspaces-Observatorio-Participaci-n-Ciudadana/"
               "8695b82b-4d33-43a8-8a3c-c302ba25f047/scratchpad")
SALIDA_DEFECTO = C.DIR_SALIDA / "matriz_diligenciada.xlsx"
MAX_Q = 600
FILA_INI_REG, FILA_FIN_REG = 6, 1005
FILA_INI_PP, FILA_FIN_PP = 6, 305


def sha(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _fecha(s):
    if isinstance(s, (dt.date, dt.datetime)) or s in (None, ""):
        return s or None
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def _num(s):
    if s in (None, ""):
        return None
    try:
        v = float(s)
        return int(v) if v == int(v) and abs(v) < 1e15 and "." not in str(s) else v
    except ValueError:
        return None


def cargar_entradas(base: Path):
    """Lee valores_variables y pp_proyectos ensamblados; si no existen, de los parciales."""
    def leer(nombre):
        p = base / f"{nombre}.csv"
        if p.exists():
            return C.leer_csv(p)
        filas = []
        pdir = base / "parciales"
        for s in C.SISTEMAS:
            q = pdir / s / f"{nombre}.csv"
            if q.exists():
                filas += C.leer_csv(q)
        return filas
    return leer("valores_variables"), leer("pp_proyectos")


def _obs_q(v: dict, alternativas: list[str]) -> str:
    partes = [f"ESCENARIO POR DEFECTO PROVISIONAL: {v.get('parametros') or 'defecto'}."]
    real = v.get("fecha_corte_real")
    if real and real != v.get("fecha_corte"):
        partes.append(f"Corte real: {real}.")
    if v.get("observaciones"):
        partes.append(v["observaciones"].strip())
    txt = " ".join(partes)
    if alternativas:
        alt = " Alternativas: " + "; ".join(alternativas)
        txt = txt + alt if len(txt) + len(alt) <= MAX_Q else txt
    return txt if len(txt) <= MAX_Q else txt[:MAX_Q - 1] + "…"


def llenar(wb, valores: list[dict], proyectos: list[dict], validado: bool = False) -> dict:
    """Escribe Registro y PP_Proyectos. Devuelve estadísticas."""
    reg = wb["Registro"]
    indice = {}                 # (id, anio, periodo, territorio) -> fila
    base_fila = {}              # id_variable -> fila base (A:E)
    libre = None
    for r in range(FILA_INI_REG, FILA_FIN_REG + 1):
        idv = reg.cell(r, 3).value
        if idv:
            indice[(idv, reg.cell(r, 6).value, reg.cell(r, 7).value, reg.cell(r, 8).value)] = r
            base_fila.setdefault(idv, r)
        elif libre is None:
            libre = r
    est = {"actualizadas": 0, "agregadas": 0, "aux_omitidas": 0, "alternativas": 0}
    defecto = [v for v in valores if v.get("es_defecto") == "True"]
    alt_por_clave = {}
    for v in valores:
        if v.get("es_defecto") != "True" and v.get("valor"):
            k = (v["id_variable"], int(v["anio"]), v["periodo"], v["territorio"])
            alt_por_clave.setdefault(k, []).append(f"{v['escenario_id']}={v['valor']}")
    for v in sorted(defecto, key=lambda x: (x["id_variable"], int(x["anio"]), x["periodo"], x["territorio"])):
        idv = v["id_variable"]
        if idv not in base_fila:           # variables auxiliares: no tienen fila en la Matriz
            est["aux_omitidas"] += 1
            continue
        clave = (idv, int(v["anio"]), v["periodo"], v["territorio"])
        r = indice.get(clave)
        if r is None:
            r = libre
            if r is None or r > FILA_FIN_REG:
                raise SystemExit("Registro lleno: sin filas libres hasta la 1005")
            b = base_fila[idv]
            for col in range(1, 6):
                reg.cell(r, col).value = reg.cell(b, col).value
            reg.cell(r, 6).value, reg.cell(r, 7).value, reg.cell(r, 8).value = clave[1:]
            indice[clave] = r
            libre = next((x for x in range(r + 1, FILA_FIN_REG + 1) if not reg.cell(x, 3).value), None)
            est["agregadas"] += 1
        else:
            est["actualizadas"] += 1
        valor = _num(v.get("valor"))
        estado = v.get("estado_informacion") or "Pendiente"
        if estado == "Validado":
            raise SystemExit("estado Validado en la entrada")
        if valor is None and estado == "Pendiente":
            estado = "No disponible"
        if valor is not None and estado == "No disponible":
            estado = "Pendiente"
        reg.cell(r, 9).value = valor
        reg.cell(r, 10).value = "Validado" if (validado and valor is not None) else estado
        reg.cell(r, 11).value = v.get("fuente") or None
        reg.cell(r, 12).value = C.normalizar_url(v.get("enlace", "")) or None
        reg.cell(r, 13).value = v.get("pagina_tabla") or None
        reg.cell(r, 14).value = _fecha(v.get("fecha_corte"))
        reg.cell(r, 15).value = _fecha(v.get("fecha_consulta"))
        hz = _num(v.get("horizonte_meta"))
        reg.cell(r, 16).value = hz
        alt = alt_por_clave.get(clave, [])
        est["alternativas"] += len(alt)
        reg.cell(r, 17).value = _obs_q(v, alt[:6])
        for col in (14, 15):
            reg.cell(r, col).number_format = "dd/mm/yyyy"

    pp = wb["PP_Proyectos"]
    filas = [p for p in proyectos if p.get("territorio", "Municipio") == "Municipio"]
    filas.sort(key=lambda p: (int(p["anio"]), p.get("periodo", ""), int(p["zona"] or 0), p["codigo_proyecto"]))
    vistos, utiles = set(), []
    for p in filas:
        k = (p["anio"], p["periodo"], p["codigo_proyecto"])
        if k not in vistos:
            vistos.add(k)
            utiles.append(p)
    if len(utiles) > FILA_FIN_PP - FILA_INI_PP + 1:
        raise SystemExit(f"PP_Proyectos admite 300 filas; hay {len(utiles)}")
    for i, p in enumerate(utiles):
        r = FILA_INI_PP + i
        m = (f"Zona {p['zona']}; APP {p.get('apropiacion_definitiva') or 'N/D'}; OPP {p.get('obligacion') or 'N/D'}; "
             f"tipología sugerida: {p.get('tipologia_sugerida') or 'N/D'}; correspondencia EFPC: "
             f"{p.get('tipo_correspondencia') or 'N/D'}. PROVISIONAL.")
        vals = [int(p["anio"]), p["periodo"] or "Anual", "Municipio", str(p["codigo_proyecto"]), p["nombre_proyecto"],
                p.get("aplica_meta_atencion") or "Pendiente", None, None, p.get("fuente") or None,
                C.normalizar_url(p.get("enlace", "")) or None, _fecha(p.get("fecha_corte")), "Pendiente", m]
        for col, val in enumerate(vals, 1):
            pp.cell(r, col).value = val
        pp.cell(r, 11).number_format = "dd/mm/yyyy"
    est["proyectos"] = len(utiles)
    return est


def firma_integridad(wb) -> dict:
    """Texto de fórmulas y estructura que no deben cambiar."""
    fir = {}
    for ws in wb:
        formulas = {}
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    formulas[c.coordinate] = v
        resto = {}
        if ws.title not in ("Registro", "PP_Proyectos"):
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None and not (isinstance(c.value, str) and c.value.startswith("=")):
                        resto[c.coordinate] = c.value
        fir[ws.title] = {"formulas": formulas, "otros": resto,
                         "validaciones": sorted((str(dv.sqref), dv.formula1, dv.type) for dv in ws.data_validations.dataValidation),
                         "cf": sorted(str(cf.sqref) for cf in ws.conditional_formatting),
                         "dim": ws.dimensions, "combinadas": sorted(str(m) for m in ws.merged_cells.ranges)}
    return fir


def comprobar_integridad(orig_path, nueva_path) -> list[str]:
    a = firma_integridad(openpyxl.load_workbook(orig_path))
    b = firma_integridad(openpyxl.load_workbook(nueva_path))
    probs = []
    if list(a) != list(b):
        probs.append(f"hojas distintas: {list(a)} vs {list(b)}")
    for h in a:
        for k in ("formulas", "otros", "validaciones", "cf", "combinadas"):
            if a[h][k] != b[h][k]:
                probs.append(f"{h}: difiere {k}")
        if h not in ("Registro", "PP_Proyectos") and a[h]["dim"] != b[h]["dim"]:
            probs.append(f"{h}: dimensiones")
    return probs


def construir(base: Path, salida: Path, validado=False, anio_periodo_terr=None):
    original = C.RUTA_MATRIZ
    h0 = sha(original)
    valores, proyectos = cargar_entradas(base)
    wb = openpyxl.load_workbook(original)
    est = llenar(wb, valores, proyectos, validado)
    if anio_periodo_terr:
        res = wb["Resultados"]
        res["B4"], res["E4"], res["H4"] = anio_periodo_terr
    wb.calculation.fullCalcOnLoad = True
    salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(salida)
    assert sha(original) == h0, "el original cambió"
    return est


def recalcular_lo(xlsx: Path, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = ["soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", str(outdir), str(xlsx)]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=240,
                       env={**__import__("os").environ, "HOME": str(SCRATCH / "lo_home")})
    out = outdir / xlsx.name
    if r.returncode != 0 or not out.exists():
        raise RuntimeError(f"LibreOffice falló: {r.stdout} {r.stderr}")
    return out


def verificar(base: Path, zonas: bool, tmp: Path) -> int:
    resultados = [r for r in C.leer_csv(base / "resultados_indicadores.csv")] if (base / "resultados_indicadores.csv").exists() \
        else [r for s in C.SISTEMAS if (base / "parciales" / s / "resultados_indicadores.csv").exists()
              for r in C.leer_csv(base / "parciales" / s / "resultados_indicadores.csv")]
    defecto = [r for r in resultados if r["es_defecto"] == "True" and not r.get("no_calculado_por_etapa") == "True"]
    combos = sorted({(int(r["anio"]), r["periodo"], r["territorio"]) for r in defecto
                     if r["territorio"] == "Municipio" or zonas})
    if not combos:
        combos = [(2024, "Anual", "Municipio")]
        print("Sin resultados de entrada: se verifica solo que LibreOffice recalcula la hoja Resultados (datos vacíos).")
    shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True)
    difs, comparados, nd_matriz = [], 0, 0
    try:
        for (anio, per, terr) in combos:
            f = tmp / f"validado_{anio}_{per.replace(' ', '')}_{terr.replace(' ', '')}.xlsx"
            construir(base, f, validado=True, anio_periodo_terr=(anio, per, terr))
            calc = recalcular_lo(f, tmp / "recalculado")
            ws = openpyxl.load_workbook(calc, data_only=True)["Resultados"]
            esperado = {r["codigo"]: r for r in defecto if (int(r["anio"]), r["periodo"], r["territorio"]) == (anio, per, terr)}
            for row in range(8, 31):
                cod = ws.cell(row, 2).value
                if cod not in esperado:
                    continue
                e = esperado[cod]
                m = ws.cell(row, 12).value
                comparados += 1
                py = float(e["puntaje"]) if e["puntaje"] else e["disponibilidad"]
                if isinstance(m, (int, float)) and isinstance(py, float):
                    if abs(m - py) > 1e-6:
                        difs.append((cod, anio, per, terr, m, py, "valor distinto"))
                elif str(m) != str(py):
                    nd_matriz += 1
                    difs.append((cod, anio, per, terr, m, py, "matriz vs python"))
            print(f"  recalculado {anio}/{per}/{terr}: LibreOffice OK; L8 (DSI)={ws['L8'].value!r}; N4={ws['N4'].value!r}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        print(f"copia temporal borrada: {not tmp.exists()}")
    print(f"comparados {comparados} resultados; diferencias {len(difs)}")
    for d in difs:
        print("  DIF", d)
    return 0 if not difs else 2


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--base", default=str(C.DIR_SALIDA), help="directorio con valores_variables.csv y pp_proyectos.csv (o parciales/)")
    ap.add_argument("--salida", default=None, help="ruta del xlsx diligenciado")
    ap.add_argument("--verificar", nargs="?", const="", default=None, metavar="RUTA_SALIDA",
                    help="además de diligenciar, valida con una copia temporal Validado recalculada en LibreOffice")
    ap.add_argument("--zonas", action="store_true", help="en --verificar, incluye también los resultados por zona")
    a = ap.parse_args()
    base = Path(a.base)
    salida = Path(a.salida or a.verificar or SALIDA_DEFECTO)
    est = construir(base, salida)
    probs = comprobar_integridad(C.RUTA_MATRIZ, salida)
    print(f"matriz diligenciada: {salida}  {est}")
    if probs:
        salida.unlink(missing_ok=True)
        print("INTEGRIDAD: diferencias respecto al original, la copia se descartó:\n  " + "\n  ".join(probs))
        return 1
    print("integridad OK: fórmulas, validaciones, formato condicional y demás hojas idénticos al original")
    if a.verificar is not None:
        return verificar(base, a.zonas, SCRATCH / "verif_matriz")
    return 0


if __name__ == "__main__":
    sys.exit(main())
