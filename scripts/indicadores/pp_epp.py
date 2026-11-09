#!/usr/bin/env python3
"""PP.EPP (ejecución presupuestal del presupuesto participativo): 2025 y 2026 (Semestre 1, provisional).

EPP = 100 x OPP / APP. Municipal = Σ OPP / Σ APP (nunca el promedio de zonas, Guía 13). Zona 1 según
`pp_zona1_tratamiento` (NE / N/D). Todo es PROVISIONAL: variables "Pendiente", resultados "Provisional".

Uso (desde la raíz, SIN -I):  python3 scripts/indicadores/pp_epp.py

Reutiliza la lógica validada de scripts/epp_pp_2025.py (se importa; no se modifica):
  plan de acción IV-trim-2025 (61 proyectos "ZONA n", sin el 0553) y contraste con el listado EFPC.
2026: plan de acción II-trim-2026 (66 proyectos "ZONA n"); OPP según `pp_2026_opp_fuente`.
Fuentes alternas por `pp_fuente_calculo`: Hacienda (sufijo del rubro = código de proyecto) y CUIPO (por BPIN).

Escribe SOLO en outputs/indicadores/parciales/PP/ (valores_variables, resultados_indicadores con codigos=["PP.EPP"],
y pp_proyectos).
"""
from __future__ import annotations

import csv
import datetime as dt
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import pymupdf

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
sys.path.insert(0, str(RAIZ / "scripts"))
import comun as C            # noqa: E402
import epp_pp_2025 as E25    # noqa: E402  (lógica validada de 2025; se importa sin ejecutar main)

SCRIPT = "scripts/indicadores/pp_epp.py"
ER = RAIZ / "data_raw_IIC" / "02_ejecucion_y_presupuesto"
F_PLAN25 = E25.F_PLAN
HOJA_PLAN25 = E25.HOJA_PLAN
HOJA_ROCA25 = "ROCA07.01.26"
F_PLAN26 = ER / "envigado_gov_co" / "seguimiento-a-proyectos-2026" / "evaluaciondefinitivaplanaccion2026segundotrimestre20260730.xlsx"
HOJA_PLAN26 = "Plan-Acción"
HOJA_ROCA26 = "ROCA-2026-07-01"
F_HAC25 = ER / "envigado_gov_co" / "ano-2025-576917" / "ejecucion-presupuestal-egresos-4-al-31122025.pdf"
F_HAC26 = ER / "envigado_gov_co" / "ano-2026-358126" / "ejecucion-presupuestal-egresos-2-al-30062026.pdf"
F_CUIPO_EJ = ER / "datos_gov_co" / "datosgov_4f7r-epif_CUIPO_Ejecucion_Gastos_Envigado.csv"
F_CUIPO_PR = ER / "datos_gov_co" / "datosgov_d9mu-h6ar_CUIPO_Programacion_Gastos_Envigado.csv"
PERIODO_CUIPO = {2025: "20251201", 2026: "20260601"}
COSTO_PROCESO = E25.COSTO_PROCESO
RE_ZONA = E25.RE_ZONA
ZONAS = list(range(2, 14))

# APP y OPP esperados (regresión): al peso, escenario por defecto
ESPERADO_2025 = {"app": 20928690218, "opp": 20684427720}
ESPERADO_APP_2026 = 24519738737
OPP_2026_REPO = 5702528148          # Hacienda 30-jun-2026 (PDF) = ROCA-2026-07-01 del xlsx


def rel(p: Path) -> str:
    return str(p.relative_to(RAIZ))


def enlace(p: Path, alterno: str = "") -> str:
    u = C.url_de_ruta(rel(p))
    return u if u.startswith("http") else alterno


# ----------------------------------------------------------------------------------------------
# Lectura de fuentes
# ----------------------------------------------------------------------------------------------
def leer_plan(f: Path, hoja: str, c_cod: int, c_nom: int, c_ue: int, c_ap: int, c_ob: int, fila0: int = 4) -> dict:
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)[hoja]
    proy = {}
    for r in ws.iter_rows(min_row=fila0, values_only=True):
        if not r[c_cod] or not r[c_nom]:
            continue
        p = proy.setdefault(str(r[c_cod]), {"nombre": str(r[c_nom]), "ue": str(r[c_ue] or ""), "ap": [], "ob": []})
        if r[c_ap] not in (None, ""):
            p["ap"].append(float(r[c_ap]))
        if r[c_ob] not in (None, ""):
            p["ob"].append(float(r[c_ob]))
    for cod, p in proy.items():
        assert len(p["ap"]) <= 1 and len(p["ob"]) <= 1, f"proyecto {cod} con varias filas financieras"
        p["ap"] = p["ap"][0] if p["ap"] else 0.0
        p["ob"] = p["ob"][0] if p["ob"] else 0.0
    return proy


def leer_roca(f: Path, hoja: str, codigos: set) -> dict:
    """Por proyecto: BPIN, compromiso acumulado, fuentes de financiación (descripción) y filas, desde la hoja ROCA."""
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)[hoja]
    hdr = None
    ix = {}
    out = defaultdict(lambda: {"bpin": set(), "comp": 0.0, "ap": 0.0, "ob": 0.0, "fuentes": set(), "filas": 0})
    for r in ws.iter_rows(values_only=True):
        if r[0] == "SECCIÓN":
            hdr = [str(h or "").replace("\n", " ").strip() for h in r]
            ix = {
                "proy": hdr.index("PROYECTO"),
                "bpin": next(i for i, h in enumerate(hdr) if h.startswith("BANCO PROYECTO DE INVERSI")),
                "fuente": hdr.index("DESCRIPCIÓN FUENTES DE FINANCIACIÓN"),
                "ap": hdr.index("APROPIACIÓN DEFINITIVA"),
                "comp": hdr.index("REGISTRO COMPROMISO ACUMULADO"),
                "ob": hdr.index("EJECUCIÓN OBLIGACIÓN ACUMULADO"),
            }
            ix["rubro"] = hdr.index("RUBRO PRESUPUESTAL")
            continue
        if not hdr or r[ix["proy"]] is None or r[ix["rubro"]] is None:
            continue
        c = str(r[ix["proy"]])
        if c not in codigos:
            continue
        o = out[c]
        if r[ix["bpin"]] not in (None, ""):
            o["bpin"].add(str(r[ix["bpin"]]).strip())
        o["fuentes"].add(str(r[ix["fuente"]] or "").strip())
        o["ap"] += float(r[ix["ap"]] or 0)
        o["comp"] += float(r[ix["comp"]] or 0)
        o["ob"] += float(r[ix["ob"]] or 0)
        o["filas"] += 1
    return dict(out)


_MONEY = re.compile(r"^\$ ?-?[\d.]+,\d\d$")


def _m(x: str) -> float:
    return float(x.replace("$", "").replace(".", "").replace(",", ".").strip())


def leer_hacienda(f: Path, codigos: set) -> dict:
    """PDF de ejecución presupuestal de egresos: por rubro (47 dígitos), 11 importes. El sufijo de 4 dígitos del
    rubro es el código de proyecto. Devuelve {codigo: {ap, comp, ob, filas}}."""
    lineas = [l.strip() for pg in pymupdf.open(f) for l in pg.get_text().splitlines() if l.strip()]
    out = defaultdict(lambda: {"ap": 0.0, "comp": 0.0, "ob": 0.0, "filas": 0})
    i = 0
    while i < len(lineas):
        if re.fullmatch(r"\d{40,}", lineas[i]) and i + 1 < len(lineas) and re.fullmatch(r"\d", lineas[i + 1]):
            j = i + 2
            if not _MONEY.match(lineas[j]):
                j += 1
            vals = []
            while j < len(lineas) and _MONEY.match(lineas[j]) and len(vals) < 11:
                vals.append(_m(lineas[j]))
                j += 1
            cod = lineas[i][-4:]
            if cod in codigos and len(vals) == 11:
                o = out[cod]
                o["ap"] += vals[5]
                o["comp"] += vals[7]
                o["ob"] += vals[8]
                o["filas"] += 1
            i = j
        else:
            i += 1
    return dict(out)


def leer_cuipo(bpins: set, periodo: str) -> tuple[dict, dict]:
    """CUIPO por BPIN en el periodo de corte, solo vigencia actual (cod 1). Devuelve (apropiacion, obligacion)."""
    ap, ob = defaultdict(float), defaultdict(float)
    with open(F_CUIPO_PR, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if r["periodo"] == periodo and r["bpin"] in bpins and r["cod_vigencia_del_gasto"] == "1":
                ap[r["bpin"]] += float(r["apropiacion_definitiva"] or 0)
    with open(F_CUIPO_EJ, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if r["periodo"] == periodo and r["bpin"] in bpins and r["cod_vigencia_del_gasto"] == "1":
                ob[r["bpin"]] += float(r["obligaciones"] or 0)
    return dict(ap), dict(ob)


def emparejar_efpc(pp: dict) -> int:
    """Misma lógica que scripts/epp_pp_2025.py: por zona, primero valores idénticos y luego el resto en orden."""
    efpc = E25.leer_efpc()
    efpc_z = defaultdict(list)
    for z, v in efpc:
        efpc_z[z].append(v)
    for z in sorted({p["zona"] for p in pp.values()}):
        libres = list(efpc_z[z])
        prz = sorted(c for c in pp if pp[c]["zona"] == z)
        for c in prz:
            v = round(pp[c]["plan_ap"])
            if v in libres:
                libres.remove(v)
                pp[c]["efpc"], pp[c]["tipo_corr"] = v, "exacta"
        for c in prz:
            if "efpc" not in pp[c]:
                pp[c]["efpc"] = libres.pop(0) if libres else None
                pp[c]["tipo_corr"] = "aproximada" if pp[c]["efpc"] is not None else "sin_correspondencia"
    return len(efpc)


# ----------------------------------------------------------------------------------------------
# Construcción de los universos 2025 y 2026
# ----------------------------------------------------------------------------------------------
def universo(anio: int) -> dict:
    if anio == 2025:
        plan = leer_plan(F_PLAN25, HOJA_PLAN25, 0, 1, 3, 13, 14)
        f_roca, h_roca, f_plan, hoja, f_hac = F_PLAN25, HOJA_ROCA25, F_PLAN25, HOJA_PLAN25, F_HAC25
    else:
        plan = leer_plan(F_PLAN26, HOJA_PLAN26, 1, 2, 4, 14, 15)
        f_roca, h_roca, f_plan, hoja, f_hac = F_PLAN26, HOJA_ROCA26, F_PLAN26, HOJA_PLAN26, F_HAC26
    pp = {}
    for c, p in plan.items():
        m = RE_ZONA.search(p["nombre"].upper())
        if m and c != COSTO_PROCESO:
            pp[c] = {"nombre": p["nombre"], "ue": p["ue"], "zona": int(m.group(1)),
                     "plan_ap": p["ap"], "plan_ob": p["ob"], "es_0553": False}
    c0 = plan[COSTO_PROCESO]
    pp[COSTO_PROCESO] = {"nombre": c0["nombre"], "ue": c0["ue"], "zona": None, "plan_ap": c0["ap"],
                         "plan_ob": c0["ob"], "es_0553": True}
    codigos = set(pp)
    roca = leer_roca(f_roca, h_roca, codigos)
    hac = leer_hacienda(f_hac, codigos)
    for c, p in pp.items():
        r = roca.get(c)
        p["bpin"] = sorted(r["bpin"])[0] if r and len(r["bpin"]) == 1 else ""
        p["bpin_varios"] = bool(r and len(r["bpin"]) > 1)
        p["comp"] = r["comp"] if r else None
        p["fuentes"] = r["fuentes"] if r else set()
        p["hac_ap"] = hac[c]["ap"] if c in hac else None
        p["hac_ob"] = hac[c]["ob"] if c in hac else None
    bpins = {p["bpin"] for p in pp.values() if p["bpin"]}
    cap, cob = leer_cuipo(bpins, PERIODO_CUIPO[anio])
    for p in pp.values():
        p["cuipo_ap"] = cap.get(p["bpin"]) if p["bpin"] else None
        p["cuipo_ob"] = cob.get(p["bpin"]) if p["bpin"] else None
    n_efpc = 0
    if anio == 2025:
        n_efpc = emparejar_efpc({c: p for c, p in pp.items() if not p["es_0553"]})
        for p in pp.values():
            p.setdefault("efpc", None)
            p.setdefault("tipo_corr", "sin_correspondencia")
    return {"anio": anio, "pp": pp, "n_efpc": n_efpc, "f_plan": f_plan, "hoja": hoja, "f_hac": f_hac}


def valores(u: dict, params: dict) -> dict:
    """APP y OPP por proyecto según los parámetros del escenario. None = N/D (no reconstruible)."""
    anio = u["anio"]
    fuente = params["pp_fuente_calculo"]
    out = {}
    for c, p in u["pp"].items():
        if fuente == "hacienda":
            app, opp = p["hac_ap"], p["hac_ob"]
            src = "hacienda"
        elif fuente == "cuipo":
            app, opp = p["cuipo_ap"], p["cuipo_ob"]
            src = "cuipo"
        else:  # plan_accion_roca
            app = p["plan_ap"]
            if anio == 2026 and params["pp_2026_opp_fuente"] == "hacienda_30jun":
                opp, src = p["hac_ob"], "plan_accion_roca (APP); hacienda 30-jun-2026 (OPP)"
            else:
                opp, src = p["plan_ob"], "plan_accion_roca"
        out[c] = {"app": app, "opp": opp, "fuente": src}
    return out


def suma(vals: dict, codigos) -> tuple:
    """(OPP, APP) de los proyectos; 'N/D' si algún proyecto no tiene valor."""
    o = [vals[c]["opp"] for c in codigos]
    a = [vals[c]["app"] for c in codigos]
    if any(x is None for x in o + a):
        return "N/D", "N/D"
    return sum(o), sum(a)


# ----------------------------------------------------------------------------------------------
# Escenarios -> filas
# ----------------------------------------------------------------------------------------------
def periodo_y_corte(anio: int, params: dict):
    """(periodo, fecha_corte, fecha_corte_real, contexto de cálculo, alerta)."""
    if anio == 2025:
        real = dt.date(2026, 1, 7) if params["pp_fuente_calculo"] == "plan_accion_roca" else dt.date(2025, 12, 31)
        fc = C.fecha_corte_registro(2025, "Anual", real, params)
        return "Anual", fc, real, None, ""
    real = dt.date(2026, 6, 30)
    if params["tratamiento_2026"] == "anual_provisional":
        return ("Anual", dt.date(2026, 6, 30), real, None,
                "2026 sin cierre rotulado Anual con corte 30-jun (contradice la Guía 4; solo sensibilidad)")
    return ("Semestre 1", dt.date(2026, 6, 30), real, {"ignorar_contexto": True},
            "Provisional, fuera del contexto de la Matriz (la fórmula exige Anual); corte 30-jun-2026")


def nota_fuente(anio: int, params: dict, u: dict) -> str:
    f = params["pp_fuente_calculo"]
    if f == "hacienda":
        return (f"APP y OPP de Hacienda ({rel(u['f_hac'])}; sufijo del rubro = código de proyecto)"
                + ("; en 2026 pp_2026_opp_fuente no aplica" if anio == 2026 else ""))
    if f == "cuipo":
        return (f"CUIPO por BPIN, periodo {PERIODO_CUIPO[anio]}, vigencia actual (solo contraste; "
                "APP de Programación de gastos, OPP de Ejecución de gastos)")
    if anio == 2026:
        o = ("OPP de Hacienda al 30-jun-2026 (PDF; coincide al peso con la hoja ROCA-2026-07-01)"
             if params["pp_2026_opp_fuente"] == "hacienda_30jun"
             else "OPP del plan de acción xlsx (hoja ROCA-2026-07-01; coincide con Hacienda al 30-jun-2026; "
                  "el encabezado de la columna dice 01-04-2026)")
        return f"APP del plan de acción II-trim-2026; {o}"
    return "Plan de acción IV-trim-2025 (hoja PlanAcción27.01.2026, corte real 07-01-2026)"


def construir(e: dict, anio: int, u: dict, filas_v: list, filas_r: list, ctx_log: dict) -> None:
    params = e["params"]
    periodo, fcorte, freal, ctx_extra, alerta_base = periodo_y_corte(anio, params)
    inc = params["pp_incluir_0553"] == "true"
    vals = valores(u, params)
    cods_pp = sorted(c for c, p in u["pp"].items() if not p["es_0553"])
    por_zona = {z: [c for c in cods_pp if u["pp"][c]["zona"] == z] for z in range(1, 14)}
    zonas_univ = list(range(1, 14)) if params["zonas_oficiales"] == "13" else ZONAS
    fuente_txt = nota_fuente(anio, params, u)
    doc = rel(u["f_plan"]) if params["pp_fuente_calculo"] == "plan_accion_roca" else (
        rel(u["f_hac"]) if params["pp_fuente_calculo"] == "hacienda" else rel(F_CUIPO_EJ))
    fdoc = Path(RAIZ / doc)
    enl = enlace(fdoc, "https://www.datos.gov.co/d/4f7r-epif")
    pagina = {"plan_accion_roca": f"Hoja {u['hoja']}, col. APROPIACIÓN DEFINITIVA y EJECUCIÓN OBLIGACIÓN"
                                  + (" (OPP 2026: ver observaciones)" if anio == 2026 else ""),
              "hacienda": "Rubros con sufijo = código de proyecto: Apropiación Definitiva y Ejecución/Obligación",
              "cuipo": f"Periodo {PERIODO_CUIPO[anio]}: apropiacion_definitiva (Programación) y obligaciones (Ejecución)"
              }[params["pp_fuente_calculo"]]
    esc_txt = "; ".join(f"{k}={v}" for k, v in e["desviaciones"].items()) or "defecto"

    terr_vals = {}   # territorio -> (opp, app)
    for z in zonas_univ:
        if z == 1 and not por_zona[1]:
            terr_vals["Zona 1"] = (None, None) if params["pp_zona1_tratamiento"] == "N/D" else (0, 0)
        else:
            terr_vals[f"Zona {z}"] = suma(vals, por_zona[z])
    # Municipal: Σ OPP / Σ APP de los proyectos PP (+ 0553 si se incluye), nunca promedio de zonas
    mun_cods = cods_pp + ([COSTO_PROCESO] if inc else [])
    terr_vals["Municipio"] = suma(vals, mun_cods)
    if inc and terr_vals["Municipio"][0] != "N/D":
        terr_vals["Municipio"] = (terr_vals["Municipio"][0], terr_vals["Municipio"][1])

    for terr in ["Municipio"] + [f"Zona {z}" for z in zonas_univ]:
        opp, app = terr_vals[terr]
        alerta = alerta_base
        obs = [fuente_txt]
        if terr == "Municipio":
            # se calcula con las sumas por proyecto (equivale a Σ zonas + 0553); ver comprobación en main()
            if ctx_extra:
                res = C.calcular("PP.EPP", [opp, app], ctx_extra)
            else:
                res = C.agregar_municipal_epp([(opp, app)])
            if inc:
                obs.append("Incluye el proyecto 0553 (costo del proceso) en OPP y APP (pp_incluir_0553=true)")
            else:
                obs.append("Excluye el proyecto 0553 (costo del proceso, pp_incluir_0553=false)")
            obs.append("Municipal = Σ OPP / Σ APP (Guía 13), no promedio de zonas")
        else:
            z = int(terr.split()[1])
            if z == 1:
                res = C.epp_zona(terr, opp, app, params, ctx_extra)
                obs.append("Zona 1 sin proyectos PP en el plan de acción (PDM p. 558: 13 zonas, 12 comités); "
                           "la Guía exige confirmarlo con Planeación")
            else:
                res = C.epp_zona(terr, opp, app, params, ctx_extra or {"periodo": periodo, "territorio": terr})
        if res["disponibilidad"] == "N/D" and (opp == "N/D" or app == "N/D"):
            res["motivo"] = "Algún proyecto sin valor reconstruible en la fuente elegida (N/D)"
            nd = [c for c in (mun_cods if terr == "Municipio" else por_zona[int(terr.split()[1])])
                  if vals[c]["opp"] is None or vals[c]["app"] is None]
            obs.append("Proyectos sin valor en la fuente: " + ", ".join(nd))
        if anio == 2025 and params["pp_fuente_calculo"] == "hacienda":
            obs.append("Hacienda: 58 de 61 proyectos coinciden al peso con el plan de acción (notas_PP)")
        obs_txt = ". ".join(obs) + f". Escenario: {esc_txt}"
        entradas = [opp if opp is not None else "N/D", app if app is not None else "N/D"]
        filas_r.append(C.fila_resultado(
            "PP.EPP", anio, e, res, entradas, territorio=terr, periodo=periodo, fecha_corte_comun=fcorte,
            alerta=alerta, script=SCRIPT, observaciones=obs_txt))
        # variables
        for idv, nombre, v in (("PP.EPP.OPP", "Obligaciones del presupuesto participativo (OPP)", opp),
                               ("PP.EPP.APP", "Apropiación definitiva del presupuesto participativo (APP)", app)):
            num = v if C.es_numero(v) else None
            filas_v.append(C.fila_valor(
                "PP", idv, anio, num, "COP", e, nombre_variable=nombre, tipo_valor="construido",
                periodo=periodo, territorio=terr, fuente=doc, enlace=enl, pagina_tabla=pagina,
                fecha_corte=fcorte, fecha_corte_real=freal,
                categoria_disponibilidad="Construible" if num is not None else "N/D",
                provisional=True, script=SCRIPT, observaciones=obs_txt))
    ctx_log[(anio, e["escenario_id"])] = terr_vals


# ----------------------------------------------------------------------------------------------
# pp_proyectos
# ----------------------------------------------------------------------------------------------
def filas_proyectos(u: dict, params: dict) -> list:
    anio = u["anio"]
    periodo, fcorte, freal, _, _ = periodo_y_corte(anio, params)
    vals = valores(u, params)
    rows = []
    doc = rel(u["f_plan"])
    enl_plan = enlace(u["f_plan"])
    for c, p in sorted(u["pp"].items(), key=lambda kv: (kv[1]["zona"] or 99, kv[0])):
        obs = []
        if p["bpin_varios"]:
            obs.append("Varios BPIN en la hoja ROCA; BPIN no asignado")
        nofte = [f for f in p["fuentes"] if f and "LIBRE DESTINACI" not in f.upper()]
        obs.append("Fuente de financiación: " + ("ICLD" if not nofte else "revisar (" + "; ".join(sorted(nofte)) + ")"))
        if p["es_0553"]:
            obs.append("Costo del proceso (asambleas), no un proyecto priorizado; excluido del EPP por defecto "
                       "(pp_incluir_0553=false); zona no aplica")
        if anio == 2026 and params["pp_2026_opp_fuente"] == "hacienda_30jun":
            obs.append(f"OPP al 30-jun-2026 del PDF de Hacienda ({rel(u['f_hac'])}); coincide con la hoja "
                       f"ROCA-2026-07-01 (plan de acción: {round(p['plan_ob'])})")
        if anio == 2026:
            obs.append("Sin listado EFPC 2026 publicado: identificación solo por 'ZONA n' en el plan de acción")
        if anio == 2025 and p["tipo_corr"] != "exacta" and not p["es_0553"]:
            obs.append("Apropiación distinta del valor EFPC (adiciones) o sin correspondencia")
        if p["es_0553"] or True:
            app, opp = vals[c]["app"], vals[c]["opp"]
        rows.append({
            "anio": anio, "periodo": periodo, "territorio": "Municipio",
            "zona": p["zona"] if p["zona"] is not None else "N/A",
            "codigo_proyecto": c, "bpin": p["bpin"], "nombre_proyecto": p["nombre"], "unidad_ejecutora": p["ue"],
            "en_listado_efpc": (p["efpc"] is not None) if anio == 2025 and not p["es_0553"] else False,
            "valor_listado_efpc": p["efpc"] if anio == 2025 else None,
            "tipo_correspondencia": p["tipo_corr"] if anio == 2025 and not p["es_0553"] else "sin_correspondencia",
            "apropiacion_definitiva": round(app) if app is not None else None,
            "compromiso": round(p["comp"]) if p["comp"] is not None and anio == 2025 else None,
            "obligacion": round(opp) if opp is not None else None,
            "fuente_valor": vals[c]["fuente"], "fecha_corte": fcorte, "fecha_corte_real": freal,
            "es_0553": p["es_0553"], "incluido_epp_defecto": not p["es_0553"],
            "estado_informacion": "Pendiente",
            "fuente": doc, "enlace": enl_plan,
            "pagina_hoja": f"Hoja {u['hoja']}; BPIN de la hoja ROCA",
            "observaciones": ". ".join(obs),
        })
    return rows


# ----------------------------------------------------------------------------------------------
def pct(x):
    return f"{x:.2f}".replace(".", ",") + " %"


def main():
    sup = C.cargar_supuestos()
    esc_list = C.escenarios("PP.EPP", sup)
    d = esc_list[0]
    assert d["es_defecto"]

    u25, u26 = universo(2025), universo(2026)
    n25 = sum(not p["es_0553"] for p in u25["pp"].values())
    n26 = sum(not p["es_0553"] for p in u26["pp"].values())
    assert (n25, n26) == (61, 66), (n25, n26)
    # fuente de financiación: todo ICLD en los proyectos PP (Guía 13), verificado en el ROCA
    for u in (u25, u26):
        no_icld = {c: p["fuentes"] for c, p in u["pp"].items() if any(f and "LIBRE DESTINACI" not in f.upper() for f in p["fuentes"])}
        print(f"{u['anio']}: proyectos con fuente distinta de ICLD en el ROCA: {len(no_icld)} {list(no_icld.items())[:3]}")

    filas_v, filas_r, log = [], [], {}
    for e in esc_list:
        construir(e, 2025, u25, filas_v, filas_r, log)
        if e["params"]["tratamiento_2026"] != "excluir":
            construir(e, 2026, u26, filas_v, filas_r, log)

    # --- regresión 2025 (defecto) --------------------------------------------------------------
    opp, app = log[(2025, "defecto")]["Municipio"]
    assert round(app) == ESPERADO_2025["app"] and round(opp) == ESPERADO_2025["opp"], (opp, app)
    r_def = next(r for r in filas_r if r["anio"] == 2025 and r["territorio"] == "Municipio" and r["escenario_id"] == "defecto")
    assert f"{r_def['puntaje']:.2f}" == "98.83", r_def["puntaje"]
    # con 0553 = 98,86; municipal Σ/Σ distinto del promedio de zonas
    e0553 = next(e for e in esc_list if e["escenario_id"] == "pp_incluir_0553=true")
    o2, a2 = log[(2025, e0553["escenario_id"])]["Municipio"]
    assert f"{100 * o2 / a2:.2f}" == "98.86", 100 * o2 / a2
    zs = [log[(2025, "defecto")][f"Zona {z}"] for z in ZONAS]
    assert round(sum(o for o, _ in zs)) == round(opp) and round(sum(a for _, a in zs)) == round(app)
    assert round(u26["pp"][COSTO_PROCESO]["plan_ap"]) > 0
    a26 = sum(p["plan_ap"] for p in u26["pp"].values() if not p["es_0553"])
    assert round(a26) == ESPERADO_APP_2026, a26
    o26h = sum(p["hac_ob"] for p in u26["pp"].values() if not p["es_0553"] and p["hac_ob"] is not None)
    o26x = sum(p["plan_ob"] for p in u26["pp"].values() if not p["es_0553"])
    print(f"OPP 2026: Hacienda PDF 30-jun = {o26h:,.0f}; plan de acción/ROCA = {o26x:,.0f}; "
          f"(la cifra 11.555.041.247 de octubre fue un error de transcripción)")
    assert abs(o26h - o26x) <= 1 and round(o26h) == OPP_2026_REPO, (o26h, o26x)
    for c, p in u26["pp"].items():     # por proyecto, Hacienda = ROCA dentro de $1
        assert p["hac_ob"] is not None and abs(p["hac_ob"] - p["plan_ob"]) <= 1, (c, p["hac_ob"], p["plan_ob"])

    ruta_v = C.escribir_parcial("PP", "valores_variables", filas_v, codigos=["PP.EPP"])
    ruta_r = C.escribir_parcial("PP", "resultados_indicadores", filas_r, codigos=["PP.EPP"])
    pj = filas_proyectos(u25, d["params"]) + filas_proyectos(u26, d["params"])
    ruta_p = C.escribir_parcial("PP", "pp_proyectos", pj)

    # --- resumen --------------------------------------------------------------------------------
    print(f"\nProyectos PP: 2025 = {n25} (+0553), 2026 = {n26} (+0553); filas pp_proyectos = {len(pj)}")
    print(f"Filas: valores={len(filas_v)}, resultados={len(filas_r)}")
    for e in esc_list:
        for anio in (2025, 2026):
            if (anio, e["escenario_id"]) not in log:
                print(f"{e['escenario_id']:<48} {anio}: sin filas (excluir)")
                continue
            res = {r["territorio"]: r for r in filas_r if r["anio"] == anio and r["escenario_id"] == e["escenario_id"]}
            mun = res["Municipio"]
            z = [r["puntaje"] for t, r in res.items() if t != "Municipio" and r["puntaje"] not in (None, "")]
            z1 = res.get("Zona 1")
            print(f"{e['escenario_id']:<48} {anio} {mun['periodo']:<10} Mun={pct(mun['puntaje']) if mun['puntaje'] not in (None, '') else mun['disponibilidad']:<9}"
                  f" zonas={pct(min(z)) + '-' + pct(max(z)) if z else 'N/D'} Z1={z1['disponibilidad'] if z1 else '-'}")
    print("\nEscrito:", rel(ruta_v), rel(ruta_r), rel(ruta_p), sep="\n  ")


if __name__ == "__main__":
    main()
