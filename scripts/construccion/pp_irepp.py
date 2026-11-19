#!/usr/bin/env python3
"""Dimension Presupuesto participativo (PP): EPP, EJ90, FID, IREPP y brecha de informacion, 2020-2025.

IREPP = 0,20*EPP + 0,40*EJ90 + 0,40*FID (pesos fijos; si falta un componente, N/D; sin redistribuir).
Insumo  : outputs/construccion/PP/pp_cruce_proyectos.csv (lo genera scripts/construccion/pp_cruce.py: cruce por proyecto
          valor votado EFPC = apropiacion inicial de Hacienda, sin el rubro 2.3.7.06) y, solo para el respaldo de
          "terminados" 2025, outputs/indicadores/clasificacion_pp_entregas.csv
Salida  : outputs/construccion/PP/pp_indicadores.csv  (; y UTF-8 con BOM)
Uso     : python3 scripts/construccion/pp_cruce.py && python3 scripts/construccion/pp_irepp.py   (desde la raiz)
Reglas  : universo = proyectos del listado EFPC (valor votado). Un proyecto no verificado cuenta en el denominador de
          EJ90 y FID, no suma al numerador y queda fuera de EPP (D52). Un anio con menos de COBERTURA_MIN de proyectos
          verificados queda N/D (no se fuerzan cruces; D53). Todo resultado es PROVISIONAL.
"""
import csv
from collections import defaultdict
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
CRUCE = RAIZ / "outputs/construccion/PP/pp_cruce_proyectos.csv"
ENTREGAS = RAIZ / "outputs/indicadores/clasificacion_pp_entregas.csv"
SALIDA = RAIZ / "outputs/construccion/PP/pp_indicadores.csv"

PERIODO = list(range(2020, 2026))          # periodo de referencia de la brecha
COBERTURA_MIN = 0.50                       # D56 (decision de la usuaria): minimo de proyectos verificados para calcular el anio
TOL_FID = 0.01                             # +-1 % entre apropiacion definitiva y valor votado
UMBRAL_EJ = 0.90
PESOS = {"EPP": 0.20, "EJ90": 0.40, "FID": 0.40}
UMBRALES = [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")]   # Rojo por debajo (propuesta)
PROV = "Provisional"
FECHA = "2026-11-18"
FECHA_ACT = {}   # fecha_consulta por anio (del manifiesto de cada fuente; se llena desde el cruce)
MOT_2026 = "vigencia en curso: 2026 no tiene cierre; fuera del periodo y del indice"
NOTA_2020 = ("2020: la vigencia 2020 se priorizo en 2019 (convencion: el anio de la votacion es t-1; Acuerdo 038/2016, art. 15A) y no hay cifra de votos publicada; "
             "el 'valor votado' del listado EFPC de 2020 corresponde a esa priorizacion de 2019 y es en realidad el valor PRIORIZADO (la que no se voto fue la vigencia 2019, en la priorizacion de 2018: informe de auditoria de la priorizacion 2019, p. 3); "
             "ademas, varios proyectos verificados quedaron con apropiacion definitiva en cero (el PDF muestra el traslado contracredito de todo su valor; "
             "la causa no consta en el documento).")

COLS = ["indicador", "codigo", "tipo", "anio", "numerador", "denominador", "valor", "puntaje_0_100",
        "peso_en_indice", "semaforo", "referente", "fuente", "enlace", "fecha_consulta", "estado", "brecha", "nota"]


def num(x):
    return float(x) if x not in ("", None) else None


def semaforo(v):
    if v is None:
        return ""
    for m, c in UMBRALES:
        if v >= m:
            return c
    return "Rojo"


def f2(v):
    return "" if v is None else f"{round(v, 2):.2f}"


def leer(p):
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


CR = leer(CRUCE)
ENT = leer(ENTREGAS)
FILAS = []


def fila(ind, cod, tipo, anio, n=None, d=None, v=None, peso="", ref="100 %", fuente="", enlace="", brecha=None, nota="",
         sem=True, puntaje=True, num_vacio=False):
    if brecha is None:
        brecha = "con dato" if v is not None else "sin dato"
    nd = v is None
    FILAS.append({
        "indicador": ind, "codigo": cod, "tipo": tipo, "anio": anio,
        "numerador": "" if num_vacio else ("N/D" if nd else ("" if n is None else n)),
        "denominador": "" if num_vacio else ("N/D" if nd else ("" if d is None else d)),
        "valor": "N/D" if nd else f2(v),
        "puntaje_0_100": f2(v) if (puntaje and not nd) else ("N/D" if (puntaje and nd) else ""),
        "peso_en_indice": peso, "semaforo": semaforo(v) if (sem and not nd) else "",
        "referente": ref, "fuente": fuente, "enlace": enlace, "fecha_consulta": FECHA_ACT.get(anio, FECHA), "estado": PROV, "brecha": brecha, "nota": nota})


# ------------------------------------------------------------------ calculos por anio
def calcular(anio):
    P = [r for r in CR if r["anio"] == str(anio) and r["tipo_cruce"] != "duplicado del listado"]   # D55: duplicados fuera del universo
    ver = [r for r in P if r["tipo_cruce"] != "no verificado"]
    n_vot, n_ver = len(P), len(ver)
    res = {"n_vot": n_vot, "n_ver": n_ver, "cobertura": n_ver / n_vot,
           "exactos": sum(r["tipo_cruce"] == "exacto" for r in P), "minimos": sum(r["tipo_cruce"] == "diferencia minima" for r in P),
           "noverif": [r["id_efpc"] for r in P if r["tipo_cruce"] == "no verificado"],
           "fecha": P[0]["fecha_consulta"], "fecha_l": P[0]["fecha_consulta_listado"], "fuente": P[0]["fuente"], "enlace": P[0]["enlace_ejecucion"], "efpc": P[0]["enlace_listado_efpc"], "P": P, "ver": ver}
    if not ver:
        return res

    def ap(r, incl=False):
        return num(r["apropiacion_definitiva_sin_2_3_7_06"]) + (num(r["excluido_2_3_7_06_apropiacion"] or "0") if incl else 0)

    def ob(r, incl=False):
        return num(r["obligacion_sin_2_3_7_06"]) + (num(r["excluido_2_3_7_06_obligacion"] or "0") if incl else 0)

    def calc(incl):
        o, a = sum(ob(r, incl) for r in ver), sum(ap(r, incl) for r in ver)
        n90 = sum(1 for r in ver if ap(r, incl) > 0 and ob(r, incl) / ap(r, incl) >= UMBRAL_EJ)
        fid = sum(1 for r in ver if abs(ap(r, incl) - num(r["valor_votado"])) <= TOL_FID * num(r["valor_votado"]))
        return {"epp": (round(o), round(a), 100 * o / a), "ej90": (n90, n_vot, 100 * n90 / n_vot), "fid": (fid, n_vot, 100 * fid / n_vot),
                "ej90_ver": (n90, n_ver, 100 * n90 / n_ver), "fid_ver": (fid, n_ver, 100 * fid / n_ver)}
    res.update(calc(False))
    res["con_rubro"] = calc(True)
    res["excluido"] = (round(sum(num(r["excluido_2_3_7_06_apropiacion"] or "0") for r in ver)),
                       round(sum(num(r["excluido_2_3_7_06_obligacion"] or "0") for r in ver)),
                       sum(1 for r in ver if num(r["excluido_2_3_7_06_apropiacion"] or "0") > 0))
    res["sin_ap"] = sum(1 for r in ver if ap(r) <= 0)
    # respaldos
    at, vt = sum(ap(r) for r in ver), sum(num(r["valor_votado"]) for r in ver)
    res["fid_mon"] = (round(at), round(vt), 100 * at / vt)
    c = sum(num(r["compromiso_sin_2_3_7_06"]) for r in ver)
    res["comp"] = (round(c), round(at), 100 * c / at) if at else None
    res["comp_igual_obl"] = all(r["compromiso_sin_2_3_7_06"] == r["obligacion_sin_2_3_7_06"] for r in ver)
    zon = defaultdict(lambda: [0.0, 0.0])
    for r in ver:
        zon[r["zona"]][0] += ob(r)
        zon[r["zona"]][1] += ap(r)
    z = {k: 100 * v[0] / v[1] for k, v in zon.items() if v[1]}
    mx, mn = max(z.items(), key=lambda t: t[1]), min(z.items(), key=lambda t: t[1])
    res["equidad"] = (mx, mn, mx[1] - mn[1], len(z))
    return res


R_ALL = {a: calcular(a) for a in PERIODO}
for _a, _r in R_ALL.items():
    FECHA_ACT[_a] = f"{_r['fecha']} (ejecucion); {_r['fecha_l']} (listado EFPC)" if _r["fecha"] != _r["fecha_l"] else _r["fecha"]
# anios con indicador: hay cruce y la cobertura alcanza el minimo
R = {a: r for a, r in R_ALL.items() if r["n_ver"] and r["cobertura"] >= COBERTURA_MIN}


def irepp(a):
    r = R.get(a)
    if r is None:
        return None
    return sum(PESOS[k] * r[k.lower()][2] for k in PESOS)


def motivo_nd(a):
    r = R_ALL[a]
    return (f"solo {r['n_ver']} de {r['n_vot']} proyectos votados se pudieron verificar con Hacienda ({100 * r['cobertura']:.0f} %; minimo "
            f"{COBERTURA_MIN * 100:.0f} %): en los otros {r['n_vot'] - r['n_ver']} el valor votado no coincide con la apropiacion inicial de "
            "ningun sufijo de proyecto (ni, salvo un caso, con la definitiva) o coincide con varios; un indice con ese denominador mediria "
            "la falta de cruce, no la fidelidad. D56 (decision de la usuaria): 2021 = N/D con brecha; hay que pedir a Planeacion el plan de accion "
            "2021 con codigos de proyecto, el decreto de liquidacion o las fichas BPIN")


# ------------------------------------------------------------------ filas
for a in PERIODO + [2026]:
    r = R.get(a)
    if a == 2026:
        for ind, cod, pes in (("EPP, ejecucion presupuestal del PP", "PP.EPP", PESOS["EPP"]),
                              ("EJ90, proyectos con ejecucion >= 90 %", "PP.EJ90", PESOS["EJ90"]),
                              ("FID, fidelidad a lo votado", "PP.FID", PESOS["FID"])):
            fila(ind, cod, "principal", a, None, None, None, pes, brecha="sin dato: " + MOT_2026, nota="Fuera del periodo.")
        fila("IREPP, indice de ejecucion y fidelidad del PP", "PP.IREPP", "indice", a, None, None, None, "",
             brecha="sin dato: " + MOT_2026, nota="No se redistribuyen pesos.")
        continue
    ra = R_ALL[a]
    efpc = ra["efpc"]
    if r is None:
        mot = motivo_nd(a)
        for ind, cod, pes in (("EPP, ejecucion presupuestal del PP", "PP.EPP", PESOS["EPP"]),
                              ("EJ90, proyectos con ejecucion >= 90 %", "PP.EJ90", PESOS["EJ90"]),
                              ("FID, fidelidad a lo votado", "PP.FID", PESOS["FID"])):
            fila(ind, cod, "principal", a, None, None, None, pes, fuente=ra["fuente"], enlace=ra["enlace"] + " ; " + efpc,
                 brecha="sin dato: " + mot, nota=f"Listado EFPC {a}: {ra['n_vot']} proyectos votados; verificados {ra['n_ver']}.")
        fila("IREPP, indice de ejecucion y fidelidad del PP", "PP.IREPP", "indice", a, None, None, None, "",
             fuente=ra["fuente"], enlace=ra["enlace"] + " ; " + efpc, brecha="sin dato: faltan los tres componentes (" + mot + ")",
             nota="No se redistribuyen pesos; el indice solo se calcula con los tres componentes.")
        # respaldo (solo contexto): lo que si se pudo cruzar
        if ra["n_ver"]:
            fila("Cobertura del cruce (proyectos verificados / votados)", "PP.COBERTURA", "respaldo", a, ra["n_ver"], ra["n_vot"],
                 100 * ra["cobertura"], "", "", ra["fuente"], ra["enlace"] + " ; " + efpc, sem=False, puntaje=False,
                 nota=f"Exactos {ra['exactos']}, diferencia minima {ra['minimos']}, no verificados {len(ra['noverif'])}. "
                      "Detalle en pp_cruce_proyectos.csv. No entra al indice.")
        continue
    enl = ra["enlace"] + " ; " + efpc
    fu = ra["fuente"] + " + listado de proyectos votados (EFPC)"
    cob = (f"Cruce: {ra['exactos']} exactos, {ra['minimos']} con diferencia minima y {len(ra['noverif'])} no verificados de {ra['n_vot']} "
           "votados; los no verificados cuentan en el denominador de EJ90 y FID, no suman al numerador y quedan fuera de EPP (D52).")
    cr = ra["con_rubro"]
    ex_ap, ex_ob, ex_n = ra["excluido"]
    nota_ex = ""
    if ex_n:
        pe = lambda v: f"{v:,}".replace(",", ".")
        nota_ex = (f" Se excluyo el rubro 2.3.7.06 (deficit fiscal, pagos de vigencias anteriores; D51): ${pe(ex_ap)} de apropiacion y "
                   f"${pe(ex_ob)} de obligacion en {ex_n} proyectos. Con el rubro incluido: EPP {cr['epp'][2]:.2f}, "
                   f"EJ90 {cr['ej90'][0]}/{cr['ej90'][1]} = {cr['ej90'][2]:.2f}, FID {cr['fid'][0]}/{cr['fid'][1]} = {cr['fid'][2]:.2f}.")
    extra = (" " + NOTA_2020) if a == 2020 else ""
    if a == 2024:
        extra += (" Los nombres de proyecto usados para confirmar y desempatar cruces de 2024 vienen de un archivo derivado "
                  "(outputs/indicadores/pp_proyectos.csv), sin manifiesto propio. D54: EFPC2024-Z9-4 cruza con el sufijo 0755.")
    if a == 2020:
        ob_v = sum(num(r["obligacion_sin_2_3_7_06"]) for r in ra["ver"])
        vo_v = sum(num(r["valor_votado"]) for r in ra["ver"])
        extra += (f" El EPP (obligacion / apropiacion definitiva) NO refleja los proyectos con contracredito total; "
                  f"obligacion / valor votado de los verificados = {100 * ob_v / vo_v:.2f} %.")
    if a == 2025:
        extra += (" 2025 sale del Plan de accion IV trimestre (hoja ROCA, misma estructura de rubros), no de un PDF de Hacienda; "
                  "valores votados de 0673 y 0786 corregidos (estaban intercambiados en pp_proyectos.csv).")
    o_, a_, v = ra["epp"]
    fila("EPP, ejecucion presupuestal del PP", "PP.EPP", "principal", a, o_, a_, v, PESOS["EPP"], fuente=fu, enlace=enl,
         nota=("Obligaciones / apropiacion definitiva de los " + f"{ra['n_ver']} proyectos verificados" + (" (proyecto 0553 'costo del proceso' fuera del universo)" if a in (2024, 2025) else "") + ". "
               + (f"El universo son los {ra['n_vot']} proyectos del listado (1 duplicado del listado excluido, D55). " if a == 2022 else "")
               + cob + nota_ex + extra))
    n90, nv, v = ra["ej90"]
    nn, nc, vc = ra["ej90_ver"]
    fila("EJ90, proyectos con ejecucion >= 90 %", "PP.EJ90", "principal", a, n90, nv, v, PESOS["EJ90"], fuente=fu, enlace=enl,
         nota=(f"Universo: proyectos votados ({nv}). Ejecucion = obligacion / apropiacion definitiva >= 90 %"
               + (f"; {ra['sin_ap']} proyectos verificados quedaron con apropiacion definitiva en cero y cuentan como no ejecutados" if ra["sin_ap"] else "")
               + f". Variante solo proyectos verificados: {nn}/{nc} = {vc:.2f}. " + cob + extra))
    ok, nv, v = ra["fid"]
    nok, nc, vc = ra["fid_ver"]
    fila("FID, fidelidad a lo votado", "PP.FID", "principal", a, ok, nv, v, PESOS["FID"], fuente=fu, enlace=enl,
         nota=(f"Proyectos cuya apropiacion definitiva (sin 2.3.7.06) esta a +-{TOL_FID * 100:.0f} % del valor votado / proyectos votados ({nv}). "
               f"Variante solo proyectos verificados: {nok}/{nc} = {vc:.2f}. " + cob + extra))
    iv = irepp(a)
    alt = sum(PESOS[k] * ra[k.lower() + ("_ver" if k != "EPP" else "")][2] for k in PESOS)
    fila("IREPP, indice de ejecucion y fidelidad del PP", "PP.IREPP", "indice", a, None, None, iv, "", fuente=fu, enlace=enl, num_vacio=True,
         nota=("0,20*EPP + 0,40*EJ90 + 0,40*FID con valores sin redondear; se redondea solo el resultado. "
               f"Con EJ90 y FID de solo proyectos verificados (mismo universo) daria {alt:.2f}." + extra))

    # --- respaldos
    fila("Cobertura del cruce (proyectos verificados / votados)", "PP.COBERTURA", "respaldo", a, ra["n_ver"], ra["n_vot"],
         100 * ra["cobertura"], "", "", fu, enl, sem=False, puntaje=False,
         nota="Exactos: valor votado = apropiacion inicial al peso; diferencia minima: hasta max($1.000; 0,02 %) con nombre y zona. No entra al indice.")
    if a == 2020:
        cc = [r for r in ra["ver"] if num(r["apropiacion_definitiva_sin_2_3_7_06"]) <= 0]
        tot = sum(num(r["apropiacion_inicial"]) for r in cc)
        fila("Proyectos con contracredito total (apropiacion definitiva en cero)", "PP.CONTRACREDITO", "respaldo", a, len(cc), ra["n_vot"],
             100 * len(cc) / ra["n_vot"], "", "", fu, enl, sem=False, puntaje=False,
             nota=f"{len(cc)} de {ra['n_vot']} proyectos con contracredito total, ${round(tot):,} (causa no consta en el PDF); cuentan como no ejecutados y no fieles. No entra al indice.".replace(",", "."))
    at, vt, v = ra["fid_mon"]
    fila("Fidelidad monetaria (apropiacion total / valor votado total)", "PP.FID_MON", "respaldo", a, at, vt, v, "", "", fu, enl,
         nota=f"Solo contexto; no entra al indice. {ra['n_ver']} proyectos verificados. Puede superar 100.", sem=False, puntaje=False)
    if ra["comp"] and not (a == 2024 and ra["comp_igual_obl"]):
        c, ap_, v = ra["comp"]
        fila("Tasa de compromiso", "PP.COMP", "respaldo", a, c, ap_, v, "", "", fu, enl,
             nota="Solo contexto; compromisos / apropiacion definitiva de los proyectos verificados.", sem=False, puntaje=False)
    else:
        fila("Tasa de compromiso", "PP.COMP", "respaldo", a, None, None, None, "", "", fu, enl,
             "sin dato: en el PDF de Hacienda 2024 el compromiso es igual a la obligacion en todos los proyectos (no hay compromisos aparte)",
             nota="Compromiso = obligacion en todos los proyectos verificados de 2024.", sem=False, puntaje=False)
    mx, mn, dif, nz = ra["equidad"]
    fila("Equidad entre zonas (brecha de ejecucion)", "PP.EQUIDAD", "respaldo", a, None, None, dif, "", "0 puntos (sin brecha)", fu, enl,
         num_vacio=True, sem=False, puntaje=False,
         nota=(f"Definicion de scripts/sistema_simplificado/calcular.py: EPP zonal mas alto (zona {mx[0]}: {mx[1]:.2f}) menos EPP zonal "
               f"mas bajo (zona {mn[0]}: {mn[1]:.2f}), en puntos; {nz} zonas con proyectos verificados. Solo contexto."))
    if a == 2025:
        f = [e for e in ENT if e["anio"] == "2025" and e["cuenta_en_defecto_pe"] == "1" and e["pec_todas_filas_dic_100"] != ""]
        pec = sum(1 for e in f if e["pec_todas_filas_dic_100"] == "1")
        fila("Proyectos terminados segun avance fisico (AUTORREPORTE)", "PP.TERMINADOS", "respaldo", a, pec, len(f),
             100 * pec / len(f), "", "", "Plan de accion IV trimestre 2025 (autorreporte de las dependencias)",
             "https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003057/152845_evaluaciondefinitivaplanaccionivtri.xlsx",
             nota="AUTORREPORTE de las dependencias, sin acta de recibo: solo contexto, no entra al indice.", sem=False, puntaje=False)
    else:
        fila("Proyectos terminados segun avance fisico (AUTORREPORTE)", "PP.TERMINADOS", "respaldo", a, None, None, None, "", "", "", enl,
             f"sin dato: no hay avance fisico por proyecto de {a} en los documentos reunidos", nota="Autorreporte; solo contexto.",
             sem=False, puntaje=False)

# ------------------------------------------------------------------ brecha (indicadores e indice)
for cod, nombre in (("PP.EPP", "EPP"), ("PP.EJ90", "EJ90"), ("PP.FID", "FID"), ("PP.IREPP", "IREPP")):
    con = [a for a in PERIODO if a in R]
    sin = [a for a in PERIODO if a not in R]
    mot = "; ".join(f"{a}: {R_ALL[a]['n_ver']} de {R_ALL[a]['n_vot']} proyectos verificados" for a in sin)
    fila(f"Brecha de informacion de {nombre} (2020-2025)", f"{cod}.BRECHA", "brecha", "2020-2025", len(sin), len(PERIODO),
         100 * len(sin) / len(PERIODO), "", "", "", "", sem=False, puntaje=False,
         nota="Anios con dato: " + ", ".join(map(str, con)) + "; sin dato: " + (", ".join(map(str, sin)) or "ninguno")
              + (f" ({mot})" if sin else "") + ". Brecha = anios sin dato / anios del periodo. El 2026 queda fuera (vigencia en curso). "
              "2020 se calcula con el valor priorizado en la priorizacion de 2019 (la vigencia 2019, priorizada en 2018, fue la que no se voto). 2021 N/D por D56 (decision de la usuaria): solo se calcula un anio si se verifica al menos el 50 % de sus proyectos.")
    FILAS[-1]["brecha"] = f"{len(sin)} de {len(PERIODO)} anios sin dato"

SALIDA.parent.mkdir(parents=True, exist_ok=True)
with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
    w.writeheader()
    w.writerows(FILAS)

# ------------------------------------------------------------------ resumen en pantalla
for a in PERIODO:
    r = R_ALL[a]
    if a in R:
        print(a, "EPP %.2f (%s/%s)" % (r["epp"][2], r["epp"][0], r["epp"][1]), "| EJ90", r["ej90"][:2], "%.2f" % r["ej90"][2],
              "| FID", r["fid"][:2], "%.2f" % r["fid"][2], "| IREPP %.2f %s" % (irepp(a), semaforo(irepp(a))),
              "| exactos", r["exactos"], "min", r["minimos"], "noverif", len(r["noverif"]))
        print("    con rubro:", {k: round(v[2], 2) for k, v in r["con_rubro"].items()}, "| equidad", r["equidad"][2:], "| sin_ap", r["sin_ap"])
    else:
        print(a, "N/D:", r["n_ver"], "de", r["n_vot"])
print("filas:", len(FILAS), "->", SALIDA)
