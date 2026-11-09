#!/usr/bin/env python3
"""PP.EPP 2024 (aproximado): correspondencia listado EFPC 2024 <-> rubros/proyectos de Hacienda al 31-dic-2024.

EPP = 100 x OPP / APP. Municipal = Σ OPP / Σ APP (nunca el promedio de zonas, Guía 13). Todo PROVISIONAL:
variables "Pendiente", resultados "Provisional", rotulado "aproximado".

Método (no hay plan de acción 2024 en el repositorio; `pp_fuente_calculo` solo admite Hacienda en 2024):
  1. EFPC24: 51 registros (zona, nombre, valor); el de la zona 8 viene sin "$" en el PDF (registro dudoso).
  2. HAC24: PDF de ejecución presupuestal de egresos al 31-dic-2024; el sufijo de 4 dígitos del rubro es el código
     de proyecto. APP = Apropiación Definitiva y OPP = Ejecución (obligaciones), sumadas por proyecto.
  3. PI24X (proyectosdeinversion2024.xlsx): código, nombre y dependencia de cada proyecto; da la zona del proyecto
     de Hacienda (el nombre dice "ZONA n").
  4. Tipo de coincidencia, un proyecto EFPC con un único código de Hacienda (1 a 1):
       exacta     : APP de Hacienda = valor EFPC (tolerancia de $1) y la zona del nombre del proyecto en PI24X
                    coincide con la zona EFPC.
       aproximada : sin igualdad de valor; mismo número de zona en el nombre del proyecto de PI24X y coincidencia
                    de palabras del nombre (>= 0,50 sobre el nombre más corto), asignación 1 a 1 por mayor
                    coincidencia. El valor de Hacienda difiere del EFPC (adiciones/ajustes) y se informa la diferencia.
       sin_correspondencia : ningún código libre de la misma zona alcanza 0,50.
  5. `pp_2024_metodo`: completa (exactas + aproximadas, defecto) | solo_exactos | nd (EPP 2024 = N/D).
     Una zona sin ningún proyecto con correspondencia en el escenario queda N/D con motivo; si tiene cobertura
     parcial se calcula y se anota la cobertura (valor EFPC con correspondencia / valor EFPC de la zona).

No se invoca solo: `pp_epp.py` lo importa y escribe todo junto (valores, resultados, pp_proyectos y
clasificacion_pp_2024) para no pisar las filas de 2025-2026. Ejecutado directamente, delega en `pp_epp.main()`.
"""
from __future__ import annotations

import datetime as dt
import difflib
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
sys.path.insert(0, str(RAIZ / "scripts"))
import comun as C            # noqa: E402
import pp_epp as P           # noqa: E402

SCRIPT = "scripts/indicadores/pp_epp_2024.py"
ER = RAIZ / "data_raw_IIC" / "02_ejecucion_y_presupuesto" / "envigado_gov_co"
F_HAC24 = ER / "ano-2024-325557" / "ejecucion-presupuestal-egresos-al-31122024.pdf"
F_PI24 = ER / "informe-de-gestion-alcaldia-de-envigado" / "proyectosdeinversion2024.xlsx"
COSTO_PROCESO = "0553"
CORTE_2024 = dt.date(2024, 12, 31)
UMBRAL_NOMBRE = 0.50
TOL_PESO = 1
ESPERADO = {"n_efpc": 51, "valor_efpc_total": 18464271721}
PUBLICADO_PREVIO = {"app": 17332821550, "opp": 17199119162}      # estimación de la nota de disponibilidad (47/50)


def rel(p: Path) -> str:
    return str(p.relative_to(RAIZ))


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9 ]", " ", s)


def _tokens(s: str) -> set:
    return {w for w in _norm(s).split() if len(w) > 3}


def _solape(a: str, b: str) -> float:
    A, B = _tokens(a), _tokens(b)
    return len(A & B) / max(1, min(len(A), len(B)))


# ----------------------------------------------------------------------------------------------
def leer_pi24() -> dict:
    """{codigo '0xxx': {nombre, dependencia, zona (del nombre o None), distribucion}} de proyectosdeinversion2024."""
    wb = openpyxl.load_workbook(F_PI24, read_only=True, data_only=True)
    pi = {}
    for ws in wb:
        for r in ws.iter_rows(min_row=2, values_only=True):
            if len(r) > 6 and r[0] is not None and re.fullmatch(r"\d+(\.0)?", str(r[0]).strip()) and r[1]:
                c = f"{int(float(r[0])):04d}"
                m = re.search(r"zona\s*(\d+)", str(r[1]), re.I)
                pi[c] = {"nombre": str(r[1]).strip(), "dependencia": ws.title, "zona": int(m.group(1)) if m else None,
                         "distribucion": r[6]}
    return pi


def clasificar() -> dict:
    """Tabla de correspondencia EFPC24 <-> Hacienda. Devuelve {'filas': [...], 'hac': {...}, 'pi': {...}}."""
    import pp_no_viables as N      # lectura del listado EFPC 2024 (sin duplicar el parser)
    efpc = N.leer_efpc24()
    assert len(efpc) == ESPERADO["n_efpc"] and sum(r["valor"] for r in efpc) == ESPERADO["valor_efpc_total"], \
        (len(efpc), sum(r["valor"] for r in efpc))
    pi = leer_pi24()
    hac = P.leer_hacienda(F_HAC24, set(pi))
    cands = {c: h for c, h in hac.items() if h["ap"] > 0}
    asignado = {}                                        # índice EFPC -> (codigo, tipo, solape)
    usados = set()
    # (a) exactas: APP de Hacienda = valor EFPC (±$1) y misma zona en el nombre de PI24X
    for i, r in enumerate(efpc):
        c = [c for c, h in cands.items() if abs(h["ap"] - r["valor"]) <= TOL_PESO
             and pi[c]["zona"] == r["zona"] and c not in usados]
        assert len(c) <= 1, ("coincidencia exacta ambigua", r["zona"], r["n"], c)
        if c:
            asignado[i] = (c[0], "exacta", _solape(r["nombre"], pi[c[0]]["nombre"]))
            usados.add(c[0])
    # (b) aproximadas: misma zona, mayor solape de nombre >= umbral, asignación 1 a 1 por solape descendente
    pares = []
    for i, r in enumerate(efpc):
        if i in asignado:
            continue
        for c in cands:
            if c in usados or pi[c]["zona"] != r["zona"]:
                continue
            s = _solape(r["nombre"], pi[c]["nombre"])
            if s >= UMBRAL_NOMBRE:
                pares.append((s, i, c))
    for s, i, c in sorted(pares, key=lambda x: (-x[0], x[1], x[2])):
        if i in asignado or c in usados:
            continue
        asignado[i] = (c, "aproximada", s)
        usados.add(c)
    filas = []
    for i, r in enumerate(efpc):
        base = {"zona": r["zona"], "n_en_zona": r["n"], "nombre_efpc": r["nombre"],
                "valor_efpc": r["valor"], "pagina_efpc": r["pag"],
                "registro_dudoso_sin_signo": bool(r["valor_sin_signo"])}
        if i in asignado:
            c, tipo, s = asignado[i]
            h, p = hac[c], pi[c]
            dif = h["ap"] - r["valor"]
            if tipo == "exacta":
                crit = (f"APP de Hacienda del proyecto {c} = valor EFPC (diferencia ${dif:,.0f}".replace(",", ".")
                        + ") y el nombre del proyecto en PI24X trae la misma zona")
            else:
                crit = (f"Sin igualdad de valor; misma zona ({r['zona']}) en el nombre PI24X y coincidencia de palabras "
                        f"{s:.2f} >= {UMBRAL_NOMBRE:.2f}; el valor de Hacienda difiere del EFPC")
            obs = []
            if r["valor_sin_signo"]:
                obs.append("Registro dudoso: el PDF EFPC trae el valor sin '$'. El valor coincide al peso con el APP de "
                           f"Hacienda del proyecto {c} y con PI24X (corroborado por dos fuentes); se trata como exacta "
                           "y se incluye (sensibilidad sin este registro en el resumen del script)")
            if tipo == "aproximada" and abs(dif) / r["valor"] > 0.20:
                obs.append(f"Diferencia relativa > 20 % ({100 * dif / r['valor']:+.1f} %): correspondencia débil en valor")
            if tipo == "aproximada" and s < 0.60:
                obs.append("Coincidencia de nombre baja; el objetivo del listado EFPC menciona el mismo componente "
                           "(revisar con Planeación)")
            base.update({"tipo_coincidencia": tipo, "codigo_hacienda": c, "nombre_pi24x": p["nombre"],
                         "dependencia_pi24x": p["dependencia"], "app_hacienda": round(h["ap"]),
                         "opp_hacienda": round(h["ob"]), "compromiso_hacienda": round(h["comp"]),
                         "diferencia_app_vs_efpc": round(dif),
                         "diferencia_relativa_pct": round(100 * dif / r["valor"], 2),
                         "solape_nombre": round(s, 2), "criterio": crit, "observaciones": ". ".join(obs)})
        else:
            mejor = max((_solape(r["nombre"], pi[c]["nombre"]) for c in cands
                         if c not in usados and pi[c]["zona"] == r["zona"]), default=None)
            base.update({"tipo_coincidencia": "sin_correspondencia", "codigo_hacienda": "", "nombre_pi24x": "",
                         "dependencia_pi24x": "", "app_hacienda": None, "opp_hacienda": None,
                         "compromiso_hacienda": None, "diferencia_app_vs_efpc": None,
                         "diferencia_relativa_pct": None, "solape_nombre": round(mejor, 2) if mejor is not None else None,
                         "criterio": ("Ningún código de Hacienda libre de la misma zona (nombre PI24X) alcanza "
                                      f"coincidencia de nombre >= {UMBRAL_NOMBRE:.2f} ni iguala el valor EFPC"),
                         "observaciones": "Sin APP/OPP reconstruible: queda fuera del EPP 2024 en todos los escenarios"})
        filas.append(base)
    codigos = [f["codigo_hacienda"] for f in filas if f["codigo_hacienda"]]
    assert len(codigos) == len(set(codigos)), "un código de Hacienda asignado a dos proyectos EFPC"
    return {"filas": filas, "hac": hac, "pi": pi, "efpc": efpc}


# ----------------------------------------------------------------------------------------------
def universo_escenario(cl: dict, params: dict) -> dict:
    """Códigos usados por zona, cobertura y 0553 según `pp_2024_metodo`."""
    metodo = params["pp_2024_metodo"]
    tipos = {"hacienda_correspondencia_completa": ("exacta", "aproximada"), "solo_exactos": ("exacta",), "nd": ()}[metodo]
    por_zona = {z: [] for z in range(1, 14)}
    cob = {z: [0, 0] for z in range(1, 14)}        # [valor EFPC con correspondencia, valor EFPC total]
    for f in cl["filas"]:
        cob[f["zona"]][1] += f["valor_efpc"]
        if f["tipo_coincidencia"] in tipos:
            por_zona[f["zona"]].append(f["codigo_hacienda"])
            cob[f["zona"]][0] += f["valor_efpc"]
    return {"por_zona": por_zona, "cob": cob, "tipos": tipos}


def _suma(hac: dict, codigos: list):
    return sum(hac[c]["ob"] for c in codigos), sum(hac[c]["ap"] for c in codigos)


def construir(e: dict, cl: dict, filas_v: list, filas_r: list, log: dict) -> None:
    params = e["params"]
    hac = cl["hac"]
    n_tot = len(cl["filas"])
    fdoc = F_HAC24
    doc = rel(fdoc)
    enl = P.enlace(fdoc)
    pagina = "Rubros con sufijo = código de proyecto: Apropiación Definitiva y Ejecución (obligaciones), al 31-dic-2024"
    fcorte = C.fecha_corte_registro(2024, "Anual", CORTE_2024, params)
    esc_txt = "; ".join(f"{k}={v}" for k, v in e["desviaciones"].items()) or "defecto"
    metodo = params["pp_2024_metodo"]
    inc = params["pp_incluir_0553"] == "true"
    zonas_univ = list(range(1, 14)) if params["zonas_oficiales"] == "13" else list(range(2, 14))
    u = universo_escenario(cl, params)
    sin_plan = params["pp_identificacion"] == "solo_plan_accion_zona"
    n_con = sum(len(v) for v in u["por_zona"].values())
    ex = sum(1 for f in cl["filas"] if f["tipo_coincidencia"] == "exacta" and "exacta" in u["tipos"])
    ap_ = sum(1 for f in cl["filas"] if f["tipo_coincidencia"] == "aproximada" and "aproximada" in u["tipos"])
    base_obs = (f"APROXIMADO (2024 sin plan de acción en el repositorio): APP y OPP de Hacienda al 31-dic-2024 ({doc}; "
                f"sufijo del rubro = código de proyecto) cruzados con el listado EFPC 2024 y PI24X; "
                f"{n_con} de {n_tot} proyectos EFPC con correspondencia en este escenario ({ex} exactas, {ap_} aproximadas)")
    terr_vals = {}
    for z in zonas_univ:
        cods = u["por_zona"][z]
        if z == 1:
            terr_vals["Zona 1"] = (None, None) if params["pp_zona1_tratamiento"] == "N/D" else (0, 0)
        elif sin_plan or metodo == "nd" or not cods:
            terr_vals[f"Zona {z}"] = ("N/D", "N/D")
        else:
            terr_vals[f"Zona {z}"] = _suma(hac, cods)
    todos = [c for z in range(2, 14) for c in u["por_zona"][z]]
    if inc:
        todos = todos + [COSTO_PROCESO]
    terr_vals["Municipio"] = ("N/D", "N/D") if (sin_plan or metodo == "nd" or not todos) else _suma(hac, todos)
    if terr_vals["Municipio"][0] != "N/D":      # coherencia: Σ zonas (+0553) = municipal
        sz = [terr_vals[f"Zona {z}"] for z in range(2, 14) if terr_vals[f"Zona {z}"][0] != "N/D"]
        extra = hac[COSTO_PROCESO] if inc else {"ob": 0, "ap": 0}
        assert round(sum(o for o, _ in sz) + extra["ob"]) == round(terr_vals["Municipio"][0])
        assert round(sum(a for _, a in sz) + extra["ap"]) == round(terr_vals["Municipio"][1])

    for terr in ["Municipio"] + [f"Zona {z}" for z in zonas_univ]:
        opp, app = terr_vals[terr]
        obs = [base_obs]
        alerta = "Aproximado: correspondencia EFPC x Hacienda por valor y nombre (2024 sin plan de acción)"
        motivo_extra = ""
        if terr == "Municipio":
            obs.append("Municipal = Σ OPP / Σ APP (Guía 13), no promedio de zonas; "
                       + ("incluye" if inc else "excluye") + " el proyecto 0553 (pp_incluir_0553=" + params["pp_incluir_0553"] + ")")
            if opp == "N/D":
                res = C.calcular("PP.EPP", ["N/D", "N/D"], {"periodo": "Anual", "territorio": "Municipio"})
            else:
                res = C.agregar_municipal_epp([(opp, app)])
            cob_m = sum(v[0] for v in u["cob"].values()) / sum(v[1] for v in u["cob"].values())
            obs.append(f"Cobertura en valor EFPC: {100 * cob_m:.1f} %")
        else:
            z = int(terr.split()[1])
            cz = u["cob"][z]
            if z == 1:
                res = C.epp_zona(terr, opp, app, params, None)
                obs.append("Zona 1 sin proyectos PP en el listado EFPC 2024 (PDM p. 558: 13 zonas, 12 comités); "
                           "la Guía exige confirmarlo con Planeación")
            else:
                res = C.calcular("PP.EPP", [opp, app], {"periodo": "Anual", "territorio": terr})
                n_z = sum(1 for f in cl["filas"] if f["zona"] == z)
                obs.append(f"{len(u['por_zona'][z])} de {n_z} proyectos EFPC de la zona con correspondencia; "
                           f"cobertura en valor EFPC {100 * cz[0] / cz[1]:.1f} %")
        if res["disponibilidad"] == "N/D" and (opp == "N/D" or app == "N/D"):
            if sin_plan:
                motivo_extra = ("N/D: pp_identificacion=solo_plan_accion_zona exige plan de acción 2024 por proyecto, "
                                "que no está en el repositorio")
            elif metodo == "nd":
                motivo_extra = "N/D por escenario pp_2024_metodo=nd (EPP 2024 no se calcula)"
            else:
                motivo_extra = (f"N/D: ningún proyecto EFPC de la zona tiene correspondencia suficiente con Hacienda "
                                f"en el escenario pp_2024_metodo={metodo}")
            res["motivo"] = motivo_extra
        obs_txt = ". ".join(obs) + f". Escenario: {esc_txt}"
        filas_r.append(C.fila_resultado(
            "PP.EPP", 2024, e, res, [opp if opp is not None else "N/D", app if app is not None else "N/D"],
            territorio=terr, periodo="Anual", fecha_corte_comun=fcorte, alerta=alerta, script=SCRIPT,
            observaciones=obs_txt))
        for idv, nombre, v in (("PP.EPP.OPP", "Obligaciones del presupuesto participativo (OPP)", opp),
                               ("PP.EPP.APP", "Apropiación definitiva del presupuesto participativo (APP)", app)):
            num = v if C.es_numero(v) else None
            filas_v.append(C.fila_valor(
                "PP", idv, 2024, num, "COP", e, nombre_variable=nombre, tipo_valor="construido", periodo="Anual",
                territorio=terr, fuente=doc, enlace=enl, pagina_tabla=pagina, fecha_corte=fcorte,
                fecha_corte_real=CORTE_2024, categoria_disponibilidad="Construible" if num is not None else "N/D",
                provisional=True, script=SCRIPT, observaciones=obs_txt))
    log[e["escenario_id"]] = terr_vals


def filas_proyectos(cl: dict, params: dict) -> list:
    hac, pi = cl["hac"], cl["pi"]
    fcorte = C.fecha_corte_registro(2024, "Anual", CORTE_2024, params)
    doc, enl = rel(F_HAC24), P.enlace(F_HAC24)
    rows = []
    for f in cl["filas"]:
        c = f["codigo_hacienda"]
        con = f["tipo_coincidencia"] != "sin_correspondencia"
        obs = [f"Correspondencia {f['tipo_coincidencia']}: {f['criterio']}"]
        if f["observaciones"]:
            obs.append(f["observaciones"])
        rows.append({
            "anio": 2024, "periodo": "Anual", "territorio": "Municipio", "zona": f["zona"],
            "codigo_proyecto": c if con else f"EFPC2024-Z{f['zona']}-{f['n_en_zona']}",
            "bpin": "", "nombre_proyecto": pi[c]["nombre"] if con else f["nombre_efpc"],
            "unidad_ejecutora": pi[c]["dependencia"] if con else "",
            "en_listado_efpc": True, "valor_listado_efpc": f["valor_efpc"], "tipo_correspondencia": f["tipo_coincidencia"],
            "apropiacion_definitiva": f["app_hacienda"], "compromiso": f["compromiso_hacienda"],
            "obligacion": f["opp_hacienda"], "fuente_valor": "hacienda" if con else "",
            "fecha_corte": fcorte, "fecha_corte_real": CORTE_2024, "es_0553": False,
            "incluido_epp_defecto": con, "estado_informacion": "Pendiente", "fuente": doc, "enlace": enl,
            "pagina_hoja": f"EFPC 2024 p. {f['pagina_efpc']}; Hacienda rubros con sufijo {c}" if con
            else f"EFPC 2024 p. {f['pagina_efpc']}",
            "observaciones": ". ".join(obs)})
    h0 = hac[COSTO_PROCESO]
    rows.append({
        "anio": 2024, "periodo": "Anual", "territorio": "Municipio", "zona": "N/A", "codigo_proyecto": COSTO_PROCESO,
        "bpin": "", "nombre_proyecto": pi[COSTO_PROCESO]["nombre"], "unidad_ejecutora": pi[COSTO_PROCESO]["dependencia"],
        "en_listado_efpc": False, "valor_listado_efpc": None, "tipo_correspondencia": "sin_correspondencia",
        "apropiacion_definitiva": round(h0["ap"]), "compromiso": round(h0["comp"]), "obligacion": round(h0["ob"]),
        "fuente_valor": "hacienda", "fecha_corte": fcorte, "fecha_corte_real": CORTE_2024, "es_0553": True,
        "incluido_epp_defecto": False, "estado_informacion": "Pendiente", "fuente": doc, "enlace": enl,
        "pagina_hoja": "Hacienda rubros con sufijo 0553",
        "observaciones": "Costo del proceso (priorización participativa), no un proyecto priorizado; excluido del EPP por "
                         "defecto (pp_incluir_0553=false); zona no aplica"})
    return rows


def filas_clasificacion(cl: dict) -> list:
    out = []
    for f in cl["filas"]:
        g = dict(f)
        g.update({"anio": 2024, "estado_informacion": "Pendiente",
                  "incluido_completa": f["tipo_coincidencia"] in ("exacta", "aproximada"),
                  "incluido_solo_exactos": f["tipo_coincidencia"] == "exacta",
                  "fuente_hacienda": rel(F_HAC24), "fuente_efpc": "data_raw_IIC/01_propuestas_y_acuerdos/"
                  "efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf", "fuente_pi24x": rel(F_PI24)})
        out.append(g)
    return out


def esc_aplicables(esc_list: list) -> list:
    """Escenarios de PP.EPP con filas 2024: todos salvo `pp_fuente_calculo=cuipo` (CUIPO exige BPIN, que 2024 no tiene;
    solo contraste)."""
    return [e for e in esc_list if e["params"]["pp_fuente_calculo"] != "cuipo"]


def construir_todo(esc_list: list) -> dict:
    cl = clasificar()
    filas_v, filas_r, log = [], [], {}
    for e in esc_aplicables(esc_list):
        construir(e, cl, filas_v, filas_r, log)
    defecto = esc_list[0]
    return {"cl": cl, "filas_v": filas_v, "filas_r": filas_r, "log": log,
            "pj": filas_proyectos(cl, defecto["params"]), "clasif": filas_clasificacion(cl)}


def pct(x):
    return f"{x:.2f}".replace(".", ",")


def imprimir(res: dict, esc_list: list) -> None:
    cl, log = res["cl"], res["log"]
    from collections import Counter
    cnt = Counter(f["tipo_coincidencia"] for f in cl["filas"])
    print(f"\nPP.EPP 2024: {len(cl['filas'])} proyectos EFPC; correspondencia: {dict(cnt)}")
    for e in esc_aplicables(esc_list):
        rs = {r["territorio"]: r for r in res["filas_r"] if r["escenario_id"] == e["escenario_id"]}
        m = rs["Municipio"]
        zs = [r["puntaje"] for t, r in rs.items() if t != "Municipio" and r["puntaje"] not in (None, "")]
        txt = pct(m["puntaje"]) if m["puntaje"] not in (None, "") else m["disponibilidad"]
        print(f"  {e['escenario_id']:<48} Mun={txt:<8} {m['semaforo']:<8} zonas calculadas={len(zs)}")
    d = log["defecto"]
    print(f"  Defecto: APP {d['Municipio'][1]:,.0f}  OPP {d['Municipio'][0]:,.0f} (estimación previa "
          f"{PUBLICADO_PREVIO['app']:,} / {PUBLICADO_PREVIO['opp']:,})")


def main() -> None:
    P.main()


if __name__ == "__main__":
    main()
