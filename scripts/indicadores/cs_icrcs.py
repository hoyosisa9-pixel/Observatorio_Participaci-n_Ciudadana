#!/usr/bin/env python3
"""CS.ICRCS = (DSI + PCS + VRC + RPO) / 4, Municipio, anual (2024 y 2025).

Lee los puntajes de los cuatro componentes en outputs/indicadores/parciales/CS/resultados_indicadores.csv
(por eso va DESPUES de cs_pcs, cs_dsi, cs_vrc y cs_rpo). Reglas de la Guia 15: solo se calcula con las 4 entradas
numericas y con el mismo corte; no se redistribuyen pesos; si falta una, N/D con el componente que falta.
Escenarios (plan 3.3): defecto + un factor a la vez (C.escenarios) + el escenario nombrado
control_institucional_545_546 (el de cs_vrc, aplicado a los cuatro componentes) + extremos_evaluados_min / _max
(extremos de los escenarios calculables de este script; no son cota de todas las combinaciones) +
envolvente_componentes_min / _max (promedio de los minimos / maximos de cada componente entre sus escenarios con numero;
combina opciones que pueden ser incompatibles y exige apartarse de la regla de VRC).
Resultados PROVISIONALES. Ejecutar sin -I desde la raiz: python3 scripts/indicadores/cs_icrcs.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C

COD = "CS.ICRCS"
SCRIPT = "scripts/indicadores/cs_icrcs.py"
ANIOS = (2024, 2025)
COMP = C.COMPONENTES[COD]
NOMBRE_CONTROL = "control_institucional_545_546"
PARAMS_CONTROL = {"vrc_universo": "institucional_pi_545_546", "vrc_evento_sin_mencion": "cero_siempre",
                  "ventana_documental_cs": "hasta_fecha_consulta"}   # igual que cs_vrc.py


def num(x):
    try:
        return float(x) if x not in (None, "") else None
    except ValueError:
        return None


def indexar():
    """(codigo, anio, parametros_str) -> fila del componente."""
    idx = {}
    for r in C.leer_csv(C.ruta_parcial("CS", "resultados_indicadores")):
        if r["codigo"] in COMP and r["territorio"] == "Municipio" and r["periodo"] == "Anual":
            idx[(r["codigo"], int(r["anio"]), r["parametros"])] = r
    return idx


def componente(idx, comp, anio, params):
    """Fila del componente para los parametros del escenario del indice (restringidos a los del componente)."""
    sub = {k: params[k] for k in C.parametros_aplicables(comp)}
    return idx.get((comp, anio, C.parametros_str(sub)))


def evaluar(idx, esc, anio):
    filas = [componente(idx, c, anio, esc["params"]) for c in COMP]
    punt = [num(f["puntaje"]) if f else None for f in filas]
    nd = [c for c, f, p in zip(COMP, filas, punt) if f is None or p is None]
    cortes = {f["fecha_corte_comun"] for f, p in zip(filas, punt) if f and p is not None}
    if not nd and len(cortes) > 1:
        res = C._resultado(None, "N/D", "Cortes distintos entre componentes (" + ", ".join(sorted(cortes))
                           + "); la Guia exige el mismo corte", COD)
        return res, punt, None, filas
    res = C.indice(COD, punt, {"periodo": "Anual", "territorio": "Municipio"})
    if nd:
        motivo = {c: (f["motivo_nd"] if f else "sin fila") for c, f in zip(COMP, filas) if c in nd}
        res["motivo"] = ("Falta puntaje de " + ", ".join(nd) + " (Guia 15; no se redistribuyen pesos). "
                         + " | ".join(f"{c}: {m}" for c, m in motivo.items()))
    corte = next(iter(cortes)) if len(cortes) == 1 else None
    return res, punt, corte, filas


def fila(esc, anio, res, punt, corte, extra_obs=""):
    ids = ", ".join(f"{c.split('.')[1]}={'N/D' if p is None else round(p, 2)}" for c, p in zip(COMP, punt))
    return C.fila_resultado(COD, anio, esc, res, punt, nombre="Índice de control y rendición social (ICRCS)",
        tipo="Índice", fecha_corte_comun=corte, script=SCRIPT,
        observaciones=f"ICRCS = promedio simple de DSI, PCS, VRC y RPO ({ids}). Provisional. {extra_obs}".strip())


def main():
    idx = indexar()
    escs = C.escenarios(COD) + [C.escenario_nombrado(NOMBRE_CONTROL, PARAMS_CONTROL, COD)]
    resultados, calculables = [], {a: [] for a in ANIOS}
    for esc in escs:
        for a in ANIOS:
            res, punt, corte, _ = evaluar(idx, esc, a)
            extra = ("Escenario de control: universo institucional (consecutivos 545 y 546), evento sin mención = cero, "
                     "ventana hasta la fecha de consulta, aplicado a los cuatro componentes." if esc["escenario_id"] == NOMBRE_CONTROL else "")
            resultados.append(fila(esc, a, res, punt, corte, extra))
            if res["disponibilidad"] == "Calculado":
                calculables[a].append((res["puntaje"], esc, punt, corte))
    ESC_DEFECTO = escs[0]
    # (a) extremos de los escenarios evaluados del propio índice (no es una cota de todas las combinaciones)
    for a in ANIOS:
        cal = calculables[a]
        _, punt_def, _, _ = evaluar(idx, ESC_DEFECTO, a)
        for nombre, pick in (("extremos_evaluados_min", min), ("extremos_evaluados_max", max)):
            etiqueta = "mínimo" if pick is min else "máximo"
            if cal:
                p, esc0, punt, corte = pick(cal, key=lambda t: t[0])
                esc = C.escenario_nombrado(nombre, esc0["desviaciones"], COD)
                res = C.indice(COD, punt, {"periodo": "Anual", "territorio": "Municipio"})
                obs = (f"Extremo {etiqueta} de los {len(cal)} escenarios calculables de {a} (escenario de origen: "
                       f"'{esc0['escenario_id']}'). Extremos de los escenarios evaluados, no cota de todas las combinaciones. "
                       "No es un valor oficial.")
                resultados.append(fila(esc, a, res, punt, corte, obs))
            else:
                esc = C.escenario_nombrado(nombre, {}, COD)
                res = C._resultado(None, "N/D", f"Ningún escenario del índice es calculable en {a} (un factor a la vez)", COD)
                resultados.append(fila(esc, a, res, punt_def, None,
                                       f"Extremo {etiqueta}: sin escenarios calculables de un factor en {a}. Extremos de los "
                                       "escenarios evaluados, no cota de todas las combinaciones. Ver 'envolvente_componentes'."))
        # (b) envolvente por componente: mínimo y máximo de cada componente entre sus escenarios con número
        mins, maxs, detalle = [], [], []
        for c in COMP:
            vals = [num(r["puntaje"]) for (cc, aa, _), r in idx.items()
                    if cc == c and aa == a and num(r["puntaje"]) is not None
                    and "productos_plan_indicativo" not in r["escenario_id"]]
            mins.append(min(vals) if vals else None)
            maxs.append(max(vals) if vals else None)
            detalle.append(f"{c.split('.')[1]} {('N/D' if not vals else str(round(min(vals), 2)) + ' a ' + str(round(max(vals), 2)))}")
        for nombre, punt, etiqueta in (("envolvente_componentes_min", mins, "mínima"), ("envolvente_componentes_max", maxs, "máxima")):
            esc = C.escenario_nombrado(nombre, {}, COD)
            res = C.indice(COD, punt, {"periodo": "Anual", "territorio": "Municipio"})
            obs = (f"ENVOLVENTE_COMPONENTES {etiqueta} {a}: promedio de los {'mínimos' if punt is mins else 'máximos'} de cada "
                   f"componente entre sus escenarios con número ({'; '.join(detalle)}; se excluye dsi_lista_metas="
                   "productos_plan_indicativo, solo de control). ADVERTENCIA: combina opciones que pueden ser incompatibles entre "
                   "componentes (ventanas, universos) y exige apartarse de la regla de VRC (presencia desconocida = N/D); es un rango "
                   "de sensibilidad, no un valor oficial.")
            if any(x is None for x in punt):
                res["motivo"] = "Algún componente no tiene ningún escenario con número"
            resultados.append(fila(esc, a, res, punt, None, obs))
    C.escribir_parcial("CS", "resultados_indicadores", resultados, codigos=[COD])
    print(f"{len(resultados)} filas CS.ICRCS")
    print("escenario | anio | DSI | PCS | VRC | RPO | ICRCS | semaforo")
    for r in sorted(resultados, key=lambda r: (r["anio"], not r["es_defecto"])):
        if r["es_defecto"] or r["disponibilidad"] == "Calculado" or r["escenario_id"].startswith(("extremos_evaluados", "envolvente_componentes")):
            print(r["escenario_id"][:60], r["anio"], *(r[f"entrada_{i}"] for i in range(1, 5)),
                  r["puntaje"], r["semaforo"], r["motivo_nd"][:70])


if __name__ == "__main__":
    main()
