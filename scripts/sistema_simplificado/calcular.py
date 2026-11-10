#!/usr/bin/env python3
"""Sistema de indicadores simplificado (PROVISIONAL): escribe outputs/sistema_simplificado/indicadores_simplificados.csv.

Lee los CSV ya calculados de outputs/indicadores/ (resultados_indicadores, valores_variables, pp_proyectos,
clasificacion_pp_entregas) SIN recalcularlos y calcula lo nuevo (fidelidad a lo votado y complementos del presupuesto
participativo, respaldos de participación ciudadana y control social, series de primera infancia, Saber 11 y cobertura
de salud, referencial del ICPD-6).

Uso (desde la raíz, sin -I):  python3 scripts/sistema_simplificado/calcular.py
Antes (una vez): scripts/sistema_simplificado/descargas_nuevas.py y scripts/sistema_simplificado/poblacion_edad.py
(este script los invoca solo si faltan sus archivos).
Los archivos descargados se leen como datos (CSV, ZIP de texto, XLSX): nunca se ejecutan.
"""
from __future__ import annotations

import csv
import re
import subprocess
import sys
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
import comun as C  # noqa: E402  (solo UMBRALES y semaforo; no se recalcula nada)

OUT = RAIZ / "outputs" / "sistema_simplificado"
IND = RAIZ / "outputs" / "indicadores"
DESC = RAIZ / "data_raw_IIC" / "06_descargas_indicadores"
F_SALIDA = OUT / "indicadores_simplificados.csv"
F_POB_EDAD = OUT / "trabajo" / "poblacion_envigado_edad_pped.csv"

# ---------------------------------------------------------------- parámetros (provisionales)
TOL_FIDELIDAD = 0.01          # ±1 % entre apropiación y valor votado
UMBRALES_PP = [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")]   # Rojo por debajo (umbrales de PP, propuesta)
LB_ILE, META_ILE = 16.3, 19.0  # PDM 2024-2027 p. 650 (línea base 2023) y meta del plan
TOL_LB = 0.1
ANIOS_SERIE = list(range(2019, 2026))
CAMPOS = ["componente", "indicador", "tipo", "anio", "numerador", "denominador", "valor", "semaforo", "estado",
          "fuente", "enlace", "nota"]
PROV = "Provisional"

URL_EFPC24 = "https://efpc.envigado.gov.co/wp-content/uploads/2025/07/Proyectos_PP_2024.pdf"
URL_TERRIDATA = "https://terridata.dnp.gov.co/"
URL_DANE = "https://www.dane.gov.co/index.php/estadisticas-por-tema/demografia-y-poblacion/proyecciones-de-poblacion"
URL_SABER = "https://www.datos.gov.co/d/kgxf-xxbe"
URL_SABER_2020B = "https://www.datos.gov.co/d/rnvb-vnyh"
URL_RLCPD = "https://www.datos.gov.co/d/kjje-h58h"
URL_MINSALUD_DIC25 = "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip"
URL_MINSALUD_AGO26 = "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip"

FILAS: list[dict] = []
INFO_PP: dict = {}


# ---------------------------------------------------------------- utilidades
def num(x):
    try:
        return float(str(x).replace(",", "."))
    except (TypeError, ValueError):
        return None


def r2(x):
    return None if x is None else round(float(x), 2)


def leer(path):
    return C.leer_csv(path)


def sem_pp(v):
    if v is None:
        return "Gris"
    for m, c in UMBRALES_PP:
        if v >= m:
            return c
    return "Rojo"


def add(comp, ind, tipo, anio, num_=None, den=None, valor=None, sem=None, fuente="", enlace="", nota=""):
    FILAS.append({"componente": comp, "indicador": ind, "tipo": tipo, "anio": anio,
                  "numerador": "" if num_ is None else num_, "denominador": "" if den is None else den,
                  "valor": "" if valor is None else r2(valor),
                  "semaforo": sem if sem else ("Gris" if valor is None else ""), "estado": PROV,
                  "fuente": fuente, "enlace": enlace, "nota": nota})


def nd(comp, ind, tipo, anio, motivo, fuente="", enlace="", construible=False):
    add(comp, ind, tipo, anio, None, None, None, "Gris", fuente, enlace,
        ("Construible: " if construible else "No disponible: ") + motivo)


# ---------------------------------------------------------------- lectura de lo ya calculado
RES = leer(IND / "resultados_indicadores.csv")
VAL = leer(IND / "valores_variables.csv")
PPP = leer(IND / "pp_proyectos.csv")
ENT = leer(IND / "clasificacion_pp_entregas.csv")


def fila_def(codigo, anio, periodo="Anual", terr="Municipio"):
    for r in RES:
        if (r["codigo"] == codigo and r["anio"] == str(anio) and r["periodo"] == periodo and r["territorio"] == terr
                and r["es_defecto"] == "True"):
            return r
    raise KeyError((codigo, anio))


def fila_esc(codigo, anio, **pares):
    """Fila de resultados cuyo escenario tiene exactamente esos parámetros (se ignora corte_fecha_registro=real)."""
    for r in RES:
        if r["codigo"] != codigo or r["anio"] != str(anio) or r["periodo"] != "Anual" or r["territorio"] != "Municipio":
            continue
        tok = dict(t.split("=", 1) for t in r["escenario_id"].split(";") if "=" in t and not t.startswith("corte_fecha_registro"))
        if set(tok) == set(pares) and all(tok[k].startswith(v) for k, v in pares.items()):
            return r
    raise KeyError((codigo, anio, pares))


def val_def(id_var, anio, periodo="Anual"):
    for r in VAL:
        if (r["id_variable"] == id_var and r["anio"] == str(anio) and r["periodo"] == periodo
                and r["escenario_id"] == "defecto" and r["territorio"] == "Municipio"):
            return r
    return None


def enlace_de(codigo, anio):
    for r in VAL:
        if r["codigo_indicador"] == codigo and r["anio"] == str(anio) and r["escenario_id"] == "defecto" and r["enlace"]:
            return r["enlace"]
    for r in VAL:
        if r["codigo_indicador"] == codigo and r["escenario_id"] == "defecto" and r["enlace"]:
            return r["enlace"]
    return ""


FUENTES = {
    "CS.PCS": "Plan Indicativo de Productos (DAP)", "CS.DSI": "Plan Indicativo, informes de gestión y fuentes oficiales",
    "CS.VRC": "Actas y cuadros de rendición de cuentas", "CS.RPO": "Cuadros de preguntas y respuestas de la rendición",
    "CS.ICRCS": "Los cuatro indicadores de control social",
    "PC.I1": "Listado EFPC, Hacienda y plan de acción", "PC.I2": "Informe de gestión 2025, PDM, ANTE",
    "PC.I3": "Informe de gestión 2025 y plan indicativo", "PC.I4": "Portal de la Alcaldía; Subdirección de Participación",
    "PDM.CM": "Forensis (INMLCF); Policía Nacional (datos.gov.co)", "PDM.CA": "POT; PDM; informe de gestión",
    "PDM.CP": "PDM 2024-2027", "PDM.CD": "Informe de gestión (Índer)",
    "PP.EPP": "Hacienda (ejecución presupuestal) y plan de acción",
}


def ya_calculado(comp, codigo, nombre, tipo, anios, periodo="Anual", extra_nota="", num_dec=None):
    """Copia (sin recalcular) la fila por defecto de resultados_indicadores.csv."""
    for a in anios:
        r = fila_def(codigo, a, periodo)
        v = num(r["puntaje"])
        n1, n2 = num(r["entrada_1"]), num(r["entrada_2"])
        enl = enlace_de(codigo, a)
        if v is None:
            nd(comp, nombre, tipo, a, (r["motivo_nd"] or "sin dato").rstrip(".") + ".", FUENTES.get(codigo, ""), enl)
        else:
            add(comp, nombre, tipo, a, n1, n2, v, r["semaforo"], FUENTES.get(codigo, ""), enl,
                ("Tomado de los resultados ya calculados. " + extra_nota).strip())


# =============================================================================== PRESUPUESTO PARTICIPATIVO
def presupuesto_participativo():
    comp = "Presupuesto participativo"
    for anio in (2024, 2025):
        fil = [r for r in PPP if r["anio"] == str(anio) and r["en_listado_efpc"] == "True"]
        cruzados = [r for r in fil if num(r["apropiacion_definitiva"]) is not None]
        n_votados = len(fil)
        ok = sum(1 for r in cruzados
                 if abs(num(r["apropiacion_definitiva"]) - num(r["valor_listado_efpc"])) <= TOL_FIDELIDAD * num(r["valor_listado_efpc"]))
        fid = 100 * ok / n_votados
        enl = URL_EFPC24 if anio == 2024 else enlace_de("PP.CEP", 2025)
        nota = (f"Provisional. Tolerancia ±{TOL_FIDELIDAD*100:.0f} % entre apropiación definitiva y valor votado (parámetro). "
                f"Denominador: {n_votados} proyectos votados del listado de la EFPC.")
        if anio == 2024:
            nota += (f" Si solo se cuentan los {len(cruzados)} proyectos que se pudieron cruzar con Hacienda, daría "
                     f"{ok}/{len(cruzados)} = " + f"{100*ok/len(cruzados):.2f}".replace(".", ",") + ".")
        nota += " Semáforo con los umbrales de PP: propuesta."
        add(comp, "Fidelidad a lo votado", "principal", anio, ok, n_votados, fid, sem_pp(fid),
            "Listado de proyectos votados (EFPC) y ejecución presupuestal (Hacienda)", enl, nota)
        # complemento monetario: apropiación total / valor votado total (de los proyectos cruzados)
        at = sum(num(r["apropiacion_definitiva"]) for r in cruzados)
        vt = sum(num(r["valor_listado_efpc"]) for r in cruzados)
        nm = f"Proyectos considerados: {len(cruzados)} de {n_votados}"
        add(comp, "Fidelidad a lo votado, versión monetaria", "complemento", anio, round(at), round(vt), 100 * at / vt, "",
            "Listado de proyectos votados (EFPC) y ejecución presupuestal (Hacienda)", enl,
            f"Complemento de la fidelidad, no un respaldo aparte. {nm}. Sin semáforo (puede superar 100).")
    # 2026: el listado no trae valor votado
    vot26 = [r for r in PPP if r["anio"] == "2026" and r["valor_listado_efpc"]]
    assert not vot26, "el listado 2026 ahora trae valor votado: revisar"
    nd(comp, "Fidelidad a lo votado", "principal", 2026,
       "el listado de 2026 no trae el valor votado de cada proyecto; se puede calcular al obtenerlo.",
       "Listado de proyectos votados (EFPC)", "", construible=True)

    # EPP: ya calculado
    for a, per in ((2024, "Anual"), (2025, "Anual"), (2026, "Semestre 1")):
        r = fila_def("PP.EPP", a, per)
        v = num(r["puntaje"])
        nota = "Tomado de los resultados ya calculados."
        if per != "Anual":
            nota += " Corte al primer semestre de 2026; sin semáforo (el año no ha terminado)."
        if a == 2024:
            nota += " Aproximado: 50 de los 51 proyectos votados se pudieron cruzar con Hacienda."
        add(comp, "EPP, ejecución presupuestal del presupuesto participativo" + (" (1.er semestre)" if per != "Anual" else ""),
            "principal", a, num(r["entrada_1"]), num(r["entrada_2"]), v, r["semaforo"] if per == "Anual" else "Gris",
            FUENTES["PP.EPP"], enlace_de("PP.EPP", a), nota)

    # tasa de compromiso 2025 (2024: Hacienda no separa compromisos)
    f25 = [r for r in PPP if r["anio"] == "2025" and r["en_listado_efpc"] == "True"]
    comp25 = sum(num(r["compromiso"]) for r in f25)
    app25 = sum(num(r["apropiacion_definitiva"]) for r in f25)
    tc = 100 * comp25 / app25
    add(comp, "Tasa de compromiso", "respaldo", 2025, round(comp25), round(app25), tc, sem_pp(tc), "Hacienda (ejecución presupuestal)",
        enlace_de("PP.EPP", 2025), "Compromisos / apropiación definitiva de los 61 proyectos. Semáforo con los umbrales de PP: propuesta.")
    f24 = [r for r in PPP if r["anio"] == "2024" and r["en_listado_efpc"] == "True" and r["apropiacion_definitiva"]]
    assert all(r["compromiso"] == r["obligacion"] for r in f24)
    nd(comp, "Tasa de compromiso", "respaldo", 2024, "Hacienda no separa los compromisos de las obligaciones en 2024.",
       "Hacienda (ejecución presupuestal)", enlace_de("PP.EPP", 2024))

    # proyectos con ejecución >= 90 %
    for anio in (2024, 2025):
        fil = [r for r in PPP if r["anio"] == str(anio) and r["en_listado_efpc"] == "True" and num(r["apropiacion_definitiva"])]
        n90 = sum(1 for r in fil if num(r["obligacion"]) / num(r["apropiacion_definitiva"]) >= 0.9)
        v = 100 * n90 / len(fil)
        add(comp, "Proyectos con ejecución de al menos 90 %", "respaldo", anio, n90, len(fil), v, sem_pp(v), "Hacienda (ejecución presupuestal)",
            enlace_de("PP.EPP", anio),
            "Proyectos con obligaciones iguales o mayores al 90 % de su apropiación. Semáforo con los umbrales de PP: propuesta.")

    # equidad entre zonas: máx − mín del EPP zonal
    for anio in (2024, 2025):
        zon = [(r["territorio"], num(r["puntaje"])) for r in RES if r["codigo"] == "PP.EPP" and r["anio"] == str(anio)
               and r["periodo"] == "Anual" and r["es_defecto"] == "True" and r["territorio"].startswith("Zona") and num(r["puntaje"]) is not None]
        mx, mn = max(zon, key=lambda z: z[1]), min(zon, key=lambda z: z[1])
        add(comp, "Equidad entre zonas (brecha de ejecución)", "respaldo", anio, round(mx[1], 2), round(mn[1], 2), mx[1] - mn[1], "",
            "Hacienda (ejecución presupuestal)", enlace_de("PP.EPP", anio),
            f"Puntos de diferencia entre la zona más alta ({mx[0]}) y la más baja ({mn[0]}) en EPP; {len(zon)} zonas con proyectos. "
            "Numerador = EPP más alto; denominador = EPP más bajo. Sin semáforo.")

    # proyectos votados que se terminaron (principal; reemplaza al antiguo "avance físico") y su respaldo monetario
    f = [r for r in ENT if r["anio"] == "2025" and r["cuenta_en_defecto_pe"] == "1"]
    pec = sum(1 for r in f if r["pec_todas_filas_dic_100"] == "1")
    pe = len(f)
    assert (pec, pe) == (int(val_def("PP.CEP.PEC", 2025)["valor"]), int(val_def("PP.CEP.PE", 2025)["valor"]))
    v = 100 * pec / pe
    ROT = "Autorreporte de las dependencias, sin acta de recibo."
    nom_a, nom_b = "Proyectos votados que se terminaron", "Ejecución presupuestal de los proyectos terminados"
    add(comp, nom_a, "principal", 2025, pec, pe, v, sem_pp(v), "Plan de acción 2025 (autorreporte de las dependencias)",
        enlace_de("PP.CEP", 2025),
        ROT + " Proyectos votados con todas las filas de diciembre en 100 % de avance físico (misma regla por defecto del avance físico) "
        "sobre proyectos votados. Semáforo con los umbrales de PP: propuesta.")
    for a_, mot in ((2024, "sin avance físico por proyecto de 2024 en los documentos reunidos."),
                    (2026, "el plan de acción de 2026 solo reporta avance al 30 de junio; no hay cierre de la vigencia.")):
        nd(comp, nom_a, "principal", a_, mot, "Plan de acción (autorreporte de las dependencias)", "")
    # (b) obligaciones / apropiación SOLO de los proyectos terminados (cruce pp_proyectos x clasificacion_pp_entregas por código de proyecto)
    terminados = {r["codigo_proyecto"] for r in f if r["pec_todas_filas_dic_100"] == "1"}
    pp25 = [r for r in PPP if r["anio"] == "2025" and r["en_listado_efpc"] == "True"]
    assert terminados <= {r["codigo_proyecto"] for r in pp25}
    sel = [r for r in pp25 if r["codigo_proyecto"] in terminados]
    assert len(sel) == pec
    ob_t = sum(num(r["obligacion"]) for r in sel)
    ap_t = sum(num(r["apropiacion_definitiva"]) for r in sel)
    vb = 100 * ob_t / ap_t
    add(comp, nom_b, "respaldo", 2025, round(ob_t), round(ap_t), vb, sem_pp(vb), "Hacienda (ejecución) y plan de acción 2025 (autorreporte)",
        enlace_de("PP.CEP", 2025),
        ROT + f" Obligaciones / apropiación definitiva solo de los {pec} proyectos con avance físico de 100 % "
        "(cruce por código de proyecto entre la ejecución de Hacienda y el plan de acción). Semáforo con los umbrales de PP: propuesta.")
    for a_, mot in ((2024, "sin avance físico por proyecto de 2024 para identificar los proyectos terminados."),
                    (2026, "sin cierre de la vigencia 2026 ni avance físico al cierre.")):
        nd(comp, nom_b, "respaldo", a_, mot, "Hacienda y plan de acción", "")
    INFO_PP.update({"terminados": pec, "ob_t": ob_t, "ap_t": ap_t})


# =============================================================================== PARTICIPACIÓN CIUDADANA
def participacion():
    comp = "Participación ciudadana"
    # I1: principal N/D y respaldo con los datos de la fidelidad
    for a in (2024, 2025):
        ya_calculado(comp, "PC.I1", "I1 - Incorporación de aportes", "principal", [a])
        fil = [r for r in PPP if r["anio"] == str(a) and r["en_listado_efpc"] == "True"]
        con = sum(1 for r in fil if num(r["apropiacion_definitiva"]))
        v = 100 * con / len(fil)
        add(comp, "I1 - Respaldo: proyectos votados con presupuesto asignado", "respaldo", a, con, len(fil), v, C.semaforo("PC.I1", v),
            "Listado de proyectos votados (EFPC), Hacienda y plan de acción", URL_EFPC24 if a == 2024 else enlace_de("PP.CEP", 2025),
            "Usa los mismos datos de la fidelidad de presupuesto participativo; no sumar en dos componentes (evita el doble conteo).")
    # I2
    ya_calculado(comp, "PC.I2", "I2 - Cobertura territorial", "principal", [2024, 2025])
    for a in (2024, 2025):
        r = fila_esc("PC.I2", a, pc_i2_actividades="asambleas_pp")
        add(comp, "I2 - Respaldo: zonas con asamblea de presupuesto participativo", "respaldo", a, num(r["entrada_1"]), num(r["entrada_2"]),
            num(r["puntaje"]), C.semaforo("PC.I2", num(r["puntaje"])), FUENTES["PC.I2"], enlace_de("PC.I2", a),
            "Solo asambleas de presupuesto participativo; la zona 1 no tiene comité zonal.")
    # I3
    ya_calculado(comp, "PC.I3", "I3 - Oferta formativa", "principal", [2024, 2025])
    for a in (2024, 2025):
        r = fila_esc("PC.I3", a, pc_i3_excluir_servidores="false")
        add(comp, "I3 - Respaldo: acciones formativas culminadas (número)", "respaldo", a, num(r["entrada_1"]), None, num(r["entrada_1"]), "",
            FUENTES["PC.I3"], enlace_de("PC.I3", a), "Conteo de acciones culminadas, incluidas las dirigidas a servidores. Sin semáforo.")
    # I4
    for a in (2024, 2025):
        ya_calculado(comp, "PC.I4", "I4 - Publicación de resultados", "principal", [a])
        v = val_def("PC.I4.V2", a)
        add(comp, "I4 - Respaldo: procesos con documento de resultados (universo exploratorio)", "respaldo", a, num(v["valor"]), None, num(v["valor"]), "",
            FUENTES["PC.I4"], v["enlace"] or enlace_de("PC.I4", a),
            "Universo exploratorio no oficial: construible al contar con el inventario de procesos participativos. Sin semáforo.")
    # ISPC
    i2 = {a: num(fila_def("PC.I2", a)["puntaje"]) for a in (2024, 2025)}
    i3 = {a: num(fila_def("PC.I3", a)["puntaje"]) for a in (2024, 2025)}
    for a in (2024, 2025):
        nd(comp, "ISPC - Índice (fórmula original: 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4)", "principal", a,
           "faltan I1 e I4.", "I1 a I4", "")
        if i2[a] is None or i3[a] is None:
            nd(comp, "ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)", "respaldo", a, "falta I3 (sin programación de acciones formativas en 2024).",
               "I2 e I3", "")
        else:
            v = (0.25 * i2[a] + 0.20 * i3[a]) / 0.45
            add(comp, "ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)", "respaldo", a, None, None, v, C.semaforo("PC.ISPC", v),
                "I2 e I3", "", "Opción para los profesores mientras I1 e I4 no sean viables: pesos 0,25/0,45 y 0,20/0,45. No es el índice original.")


# =============================================================================== CONTROL SOCIAL
def control_social():
    comp = "Control social"
    ya_calculado(comp, "CS.PCS", "PCS - Promoción del control social", "principal", [2024, 2025])
    for a in (2024, 2025):
        r = fila_def("CS.PCS", a)
        add(comp, "PCS - Respaldo: estrategias ejecutadas (número)", "respaldo", a, num(r["entrada_1"]), None, num(r["entrada_1"]), "",
            FUENTES["CS.PCS"], enlace_de("CS.PCS", a), "Conteo de estrategias de promoción ejecutadas. Sin semáforo.")
    ya_calculado(comp, "CS.DSI", "DSI - Disponibilidad de información de seguimiento", "principal", [2024, 2025])
    r = fila_esc("CS.DSI", 2024, ventana_documental_cs="hasta_fecha_consulta")
    add(comp, "DSI - Respaldo: con todo lo publicado hasta la fecha de consulta", "respaldo", 2024, num(r["entrada_1"]), num(r["entrada_2"]),
        num(r["puntaje"]), r["semaforo"], FUENTES["CS.DSI"], enlace_de("CS.DSI", 2024),
        "Cuenta también lo publicado después del cierre de la vigencia, hasta la consulta (10 de noviembre de 2026).")
    ya_calculado(comp, "CS.VRC", "VRC - Veedurías en la rendición de cuentas", "principal", [2024, 2025])
    vr = {}
    for a in (2024, 2025):
        r = fila_esc("CS.VRC", a, vrc_evento_sin_mencion="cero_siempre")
        vr[a] = r
        add(comp, "VRC - Respaldo: evento sin mención de veeduría = ausencia", "respaldo", a, num(r["entrada_1"]), num(r["entrada_2"]),
            num(r["puntaje"]), r["semaforo"], FUENTES["CS.VRC"], enlace_de("CS.VRC", a),
            "Supuesto: si el acta no menciona veedurías se cuenta como ausencia. Provisional.")
    ya_calculado(comp, "CS.RPO", "RPO - Respuesta a observaciones ciudadanas", "principal", [2024, 2025])
    rp = {2024: fila_def("CS.RPO", 2024), 2025: fila_esc("CS.RPO", 2025, rpo_registro_incompleto="calcular_con_publicado")}
    for a in (2024, 2025):
        r = rp[a]
        if a == 2024:
            continue
        add(comp, "RPO - Respaldo: con el cuadro de preguntas y respuestas publicado", "respaldo", a, num(r["entrada_1"]), num(r["entrada_2"]),
            num(r["puntaje"]), r["semaforo"], FUENTES["CS.RPO"], enlace_de("CS.RPO", a),
            "2025: solo hay un cuadro publicado con 2 observaciones; el resultado es frágil.")
    ya_calculado(comp, "CS.ICRCS", "ICRCS - Índice de Condiciones y Respuesta para el Control Social", "principal", [2024, 2025])
    for a in (2024, 2025):
        dsi, pcs = num(fila_def("CS.DSI", a)["puntaje"]), num(fila_def("CS.PCS", a)["puntaje"])
        v = (dsi + pcs + num(vr[a]["puntaje"]) + num(rp[a]["puntaje"])) / 4
        esperado = {2024: 65.196, 2025: 42.5}[a]
        assert abs(v - esperado) < 0.01, (a, v)
        add(comp, "ICRCS - Respaldo: con los respaldos de VRC y RPO", "respaldo", a, None, None, v, C.semaforo("CS.ICRCS", v),
            "Los cuatro indicadores de control social", "", "Promedio de DSI, PCS, VRC y RPO usando los respaldos de VRC y RPO. Provisional.")


# =============================================================================== PDM
def leer_terridata():
    t = zipfile.ZipFile(DESC.parent / "05_fuentes_externas_y_pot" / "dnp_terridata" / "TerriData05266f.txt.zip").read("TerriData05266f.txt").decode("utf-8")
    d = {}
    for l in t.splitlines()[1:]:
        p = l.split("|")
        if p[2] == "05266" and p[7]:
            d.setdefault(p[6], {})[int(p[9])] = float(p[7].replace(".", "").replace(",", "."))
    return d


def poblacion_pped_total():
    import openpyxl
    wb = openpyxl.load_workbook(DESC / "dane" / "PPED-AreaMun-2018-2042_VP.xlsx", read_only=True)
    out = {}
    for row in wb["PobMunicipalxÁrea"].iter_rows(values_only=True):
        if row and str(row[2]) == "05266" and row[5] == "Total":
            out[int(row[4])] = int(row[6])
    return out


def poblacion_postcovid_total():
    import openpyxl
    wb = openpyxl.load_workbook(DESC / "dane" / "DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx", read_only=True)
    out = {}
    for row in wb.active.iter_rows(values_only=True):
        if row and str(row[2]) == "05266" and row[5] == "Total":
            out[int(row[4])] = int(row[6])
    return out


def poblacion_edad():
    if not F_POB_EDAD.exists():
        subprocess.run([sys.executable, str(RAIZ / "scripts" / "sistema_simplificado" / "poblacion_edad.py")], check=True)
    d = {}
    with open(F_POB_EDAD, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r["edad"] != "total":
                d.setdefault(int(r["anio"]), {})[int(r["edad"])] = int(r["poblacion"])
    return d


INFO_ILE = {}


def ile_primera_infancia(td, pob_edad):
    comp = "Cumplimiento del PDM"
    atend = {a: td["Número de niños, niñas y mujeres gestantes con educación inicial en el marco de la atención integral a la primera infancia"][a]
             for a in range(2019, 2024)}
    for a in (2024, 2025):
        atend[a] = num(val_def("AUX.PDM.ILE.A.ATENDIDOS", a)["valor"])
    fuente_atend = {a: ("TerriData (MinEducación)" if a <= 2023 else "Informe de gestión de la Alcaldía") for a in atend}
    enl_atend = {a: (URL_TERRIDATA if a <= 2023 else val_def("AUX.PDM.ILE.A.ATENDIDOS", a)["enlace"]) for a in atend}
    # calibración en el año base del PDM (2023)
    cal = {}
    for nombre, hasta in (("0 a 5 años", 5), ("0 a 6 años", 6)):
        pob = sum(pob_edad[2023][e] for e in range(0, hasta + 1))
        cal[nombre] = (atend[2023], pob, 100 * atend[2023] / pob)
    ok = [n for n, (_, _, v) in cal.items() if abs(v - LB_ILE) <= TOL_LB]
    if len(ok) == 1:
        elegido, motivo_edad = ok[0], "reproduce la línea base"
    else:
        elegido = "0 a 5 años"
        motivo_edad = ("ninguno reproduce la línea base de 16,3 con ±0,1 (se usa 0 a 5 años por regla)" if not ok
                       else "ambos reproducen la línea base (se usa 0 a 5 años por regla)")
    hasta = 5 if elegido == "0 a 5 años" else 6
    INFO_ILE.update({"cal": cal, "elegido": elegido, "motivo": motivo_edad})
    nota_edad = ("Año base del PDM: 2023 (línea base 16,3). " +
                 "Con la población DANE por edad simple (proyección 2018-2042), 2023 da " +
                 " y ".join(f"{v:.2f} con {n}".replace(".", ",") for n, (_, _, v) in cal.items()) + f"; se usa {elegido}: {motivo_edad}.")
    serie, prev = {}, None
    for a in ANIOS_SERIE:
        pob = sum(pob_edad[a][e] for e in range(0, hasta + 1))
        v = 100 * atend[a] / pob
        serie[a] = v
        sem = ""
        var = None if prev is None else v - serie[a - 1]
        n_ = nota_edad if a == 2023 else ""
        add(comp, "ILE primera infancia - cobertura de atención integral", "principal", a, atend[a], pob, v, "", fuente_atend[a] + "; DANE (proyección de población)",
            enl_atend[a], (f"Niños y niñas de {elegido} con atención integral / población DANE de la misma edad. " + n_).strip())
        if var is not None:
            add(comp, "ILE primera infancia - variación anual (puntos porcentuales)", "principal", a, None, None, var, "", "Cálculo propio sobre la serie", "",
                f"Valor de {a} menos el de {a-1}.")
        av = max(0.0, min(100.0, 100 * (v - LB_ILE) / (META_ILE - LB_ILE)))
        add(comp, "ILE primera infancia - avance hacia la meta 2027", "principal", a, None, None, av, "", "PDM 2024-2027 (línea base y meta)", "",
            "100 × (valor − 16,3) / (19 − 16,3), con tope entre 0 y 100. Línea base 16,3 (2023) y meta 19.")
        prev = v
    INFO_ILE["serie"] = serie
    INFO_ILE["atend"] = (atend, fuente_atend, enl_atend)
    INFO_ILE["pob_edad"] = {a: sum(pob_edad[a][e] for e in range(0, hasta + 1)) for a in ANIOS_SERIE}
    INFO_ILE["rango"] = elegido


def ile_saber():
    comp = "Cumplimiento del PDM"
    filas = []
    for nombre in ("saber11_envigado_agregado_kgxf-xxbe.csv", "saber11_envigado_agregado_rnvb-vnyh_2020-2.csv"):
        with open(DESC / "icfes" / nombre, encoding="utf-8-sig", newline="") as f:
            filas += list(csv.DictReader(f))
    def serie(filtro):
        acc = {}
        for r in filas:
            if not filtro(r):
                continue
            a = int(r["periodo"][:4])
            s, n = acc.get(a, (0.0, 0))
            acc[a] = (s + float(r["puntaje_promedio"]) * int(r["n_estudiantes"]), n + int(r["n_estudiantes"]))
        return acc
    per = {}
    for r in filas:
        per.setdefault(int(r["periodo"][:4]), set()).add(r["periodo"])
    ofi = serie(lambda r: r["cole_naturaleza"] == "OFICIAL" and r["cole_area_ubicacion"] == "URBANO")
    tot = serie(lambda r: True)
    completo = {a for a, p in per.items() if any(x.endswith("4") for x in p) and any(x.endswith("1") for x in p)}
    ref = "Saber 11 (ICFES), Resultados únicos, datos.gov.co"
    res = {"oficial_urbano": {}, "municipio": {}}
    for a in ANIOS_SERIE:
        # instituciones oficiales urbanas (comparable con el indicador del PDM: 276 en 2023)
        if a in ofi:
            v = ofi[a][0] / ofi[a][1]
            res["oficial_urbano"][a] = v
            add(comp, "ILE Saber 11 - puntaje global promedio, instituciones oficiales urbanas", "respaldo", a, ofi[a][1], None, v, "", ref, URL_SABER,
                "Puntaje promedio ponderado por número de estudiantes (el indicador del PDM tiene línea base 276 en 2023). Respaldo mientras no se tenga la clasificación de planteles. Sin semáforo.")
        else:
            nd(comp, "ILE Saber 11 - puntaje global promedio, instituciones oficiales urbanas", "respaldo", a,
               "el conjunto público de ICFES en datos.gov.co no trae resultados de ese año." if a > 2022 else
               "falta el calendario A de ese año en el conjunto público.", ref, URL_SABER)
        # municipio, todos los planteles (solo si están ambos calendarios)
        if a in tot and a in completo:
            v = tot[a][0] / tot[a][1]
            res["municipio"][a] = v
            add(comp, "ILE Saber 11 - puntaje global promedio, todos los planteles del municipio", "respaldo", a, tot[a][1], None, v, "", ref, URL_SABER,
                "Promedio ponderado de los dos calendarios (A y B). Sin semáforo.")
        else:
            nd(comp, "ILE Saber 11 - puntaje global promedio, todos los planteles del municipio", "respaldo", a,
               "el conjunto público de ICFES no trae resultados de ese año." if a > 2022 else
               "solo está el calendario B (menos de 300 estudiantes), no es representativo.", ref, URL_SABER)
    for k, nombre in (("oficial_urbano", "oficiales urbanas"), ("municipio", "todos los planteles")):
        for a in ANIOS_SERIE:
            if a in res[k] and a - 1 in res[k]:
                add(comp, f"ILE Saber 11 - variación anual del puntaje ({nombre})", "respaldo", a, None, None, res[k][a] - res[k][a - 1], "", ref, URL_SABER,
                    f"Puntos de diferencia entre {a} y {a-1}.")
    add(comp, "ILE Saber 11 - clasificación de planteles (meta 14 %)", "principal", 2025, None, None, None, "Gris", "ICFES", "",
        "No disponible: no se encontró en datos abiertos la clasificación de planteles de Envigado (solo hay la de otros municipios); "
        "el avance hacia la meta de 14 % solo aplica a esa clasificación. Se solicita a la Secretaría de Educación.")
    INFO_ILE["saber"] = res


def cobertura_salud(td, pop):
    comp = "Cumplimiento del PDM"
    tot = td["Afiliados al SGSSS"]
    contrib, sub, esp = td["Afiliados al régimen contributivo"], td["Afiliados al régimen subsidado"], td["Afiliados a regímenes especiales"]
    for a in (2021, 2022, 2023, 2024, 2025):   # verificación: total = contributivo + subsidiado + especiales
        assert abs(contrib[a] + sub[a] + esp[a] - tot[a]) < 1, a
    nota0 = ("Afiliados (contributivo + subsidiado + especiales, corte diciembre; TerriData/MinSalud) / población DANE de la proyección 2018-2042 "
             "(la misma para todos los años, comparable). ")
    prev = None
    ser = {}
    for a in ANIOS_SERIE:
        if a not in tot:
            nd(comp, "Cobertura de salud", "principal", a,
               "TerriData no publica el total de afiliados ni los regímenes especiales de ese año (la serie completa empieza en 2020).",
               "TerriData (MinSalud)", URL_TERRIDATA)
            continue
        v = 100 * tot[a] / pop[a]
        ser[a] = v
        add(comp, "Cobertura de salud", "principal", a, int(tot[a]), pop[a], v, "", "TerriData (MinSalud), diciembre; DANE (proyección de población 2018-2042)",
            URL_TERRIDATA, (nota0 if a == 2020 else "") + ("Sin meta oficial: no tiene semáforo." if a == 2020 else ""))
    d = val_def("AUX.PDM.CS.COB_TOTAL_SGSSS", 2025)
    add(comp, "Cobertura de salud, cifra publicada por MinSalud (diciembre)", "referencial", 2025, 219795, 250767, num(d["valor"]), "", "MinSalud, cifras de afiliación a diciembre de 2025",
        URL_MINSALUD_DIC25, "Usa la población de referencia que MinSalud tenía en diciembre de 2025.")
    d2 = [r for r in VAL if r["id_variable"] == "AUX.PDM.CS.COB_TOTAL_SGSSS" and r["anio"] == "2026" and r["escenario_id"] == "defecto"][0]
    add(comp, "Cobertura de salud, cifra publicada por MinSalud (agosto)", "referencial", 2026, 222151, 268990, num(d2["valor"]), "", "MinSalud, cifras de afiliación a agosto de 2026",
        URL_MINSALUD_AGO26, "La diferencia con diciembre de 2025 se debe a que MinSalud actualizó la población con otra proyección, no a una pérdida de afiliados.")
    INFO_ILE["salud"] = ser
    INFO_ILE["afiliados"] = {"Afiliados al régimen contributivo": contrib, "Afiliados al régimen subsidiado": sub,
                             "Afiliados a regímenes especiales": esp, "Afiliados al SGSSS": tot}
    INFO_ILE["pop_pped"] = pop


def discapacidad():
    comp = "Cumplimiento del PDM"
    with open(DESC / "minsalud_datosgov" / "rlcpd_envigado_por_anio_registro_kjje-h58h.csv", encoding="utf-8-sig", newline="") as f:
        filas = list(csv.DictReader(f))
    total = sum(float(r["personas_registradas"]) for r in filas)
    ult = max(int(r["anio_registro"]) for r in filas)
    INFO_ILE["discap"] = (ult, total)
    add(comp, "Inclusión: personas con discapacidad en el registro nacional (acumulado)", "referencial", ult, int(total), None, total, "",
        "MinSalud, Registro de Localización y Caracterización de Personas con Discapacidad (datos.gov.co)", URL_RLCPD,
        f"Personas registradas desde 2003 hasta {ult}; el conjunto público no se actualiza desde 2022 y es un registro acumulado, no una tasa. "
        "Sigue por verificar con la Secretaría de Salud o de Inclusión Social. Sin semáforo.")


def pdm_otros():
    comp = "Cumplimiento del PDM"
    ya_calculado(comp, "PDM.CM", "CM - Mortalidad vial", "principal", [2024], extra_nota="Valor = logro frente a la meta (7,25 muertes por 100.000 habitantes frente a 8).")
    r = [x for x in RES if x["codigo"] == "PDM.CM" and x["anio"] == "2024" and x["periodo"] == "Anual" and x["territorio"] == "Municipio"
         and "pdm_cm_fuente=policia_datosgov" in x["escenario_id"] and "poblacion_dane=dane_postcovid" in x["escenario_id"] and "horizonte" not in x["escenario_id"]][0]
    add(comp, "CM - Respaldo: cifra de la Policía", "respaldo", 2024, num(r["entrada_1"]), num(r["entrada_2"]), num(r["puntaje"]), r["semaforo"],
        FUENTES["PDM.CM"], enlace_de("PDM.CM", 2024), "Homicidios en accidente de tránsito (Policía) por 100.000 habitantes con la población DANE post-COVID: 8,05 frente a la meta 8.")
    ya_calculado(comp, "PDM.CA", "CA - Espacio público verde", "principal", [2024], extra_nota="Valor = logro frente a la meta (5,75 m² por habitante frente a 3,88).")
    r = fila_esc("PDM.CA", 2024, pdm_ca_fuente="pot_epv_total")
    add(comp, "CA - Respaldo: espacio público verde total del POT", "respaldo", 2024, num(r["entrada_1"]), num(r["entrada_2"]), num(r["puntaje"]), r["semaforo"],
        FUENTES["PDM.CA"], enlace_de("PDM.CA", 2024), "Usa otra serie (EPV total del POT, 69,41 m² por habitante); no es comparable con la serie de la línea base.")
    ya_calculado(comp, "PDM.CP", "CP - Procesos participativos", "principal", [2024, 2025])
    d = val_def("AUX.PDM.CP.INDICE_SITUACIONAL", 2022)
    add(comp, "CP - Respaldo: índice situacional de implementación", "respaldo", 2022, num(d["valor"]), None, num(d["valor"]), "", "PDM 2024-2027 p. 561",
        d["enlace"], "Índice del diagnóstico del PDM (3,23 en 2022); sin meta comparable. Sin semáforo.")
    ya_calculado(comp, "PDM.CD", "CD - Servicios deportivos", "principal", [2025])
    dd, dm = val_def("AUX.PDM.CD.DISCIPLINAS", 2025), val_def("AUX.PDM.CD.DISCIPLINAS_META", 2025)
    v = 100 * num(dd["valor"]) / num(dm["valor"])
    add(comp, "CD - Respaldo: disciplinas deportivas atendidas", "respaldo", 2025, num(dd["valor"]), num(dm["valor"]), v, C.semaforo("PDM.CD", v),
        FUENTES["PDM.CD"], dd["enlace"], "Disciplinas implementadas por el Índer frente a las 40 previstas; no mide cobertura de población.")


def icpd6():
    comp = "Cumplimiento del PDM"
    nota_ref = "referencial: no cumple la regla de la Guía de no calcular índices con entradas faltantes"
    for a in (2024, 2025):
        ent = []   # (nombre, valor)
        for cod, nom in (("PDM.CM", "CM"), ("PDM.CA", "CA"), ("PDM.CP", "CP")):
            try:
                v = num(fila_def(cod, a)["puntaje"])
            except KeyError:
                v = None
            if v is not None:
                ent.append((nom, v))
        if a == 2025:
            d = [f for f in FILAS if f["indicador"] == "CD - Respaldo: disciplinas deportivas atendidas" and f["anio"] == 2025][0]
            ent.append(("CD (respaldo)", float(d["valor"])))
        av = INFO_ILE["serie"][a]
        ent.append(("ILE primera infancia (avance)", max(0.0, min(100.0, 100 * (av - LB_ILE) / (META_ILE - LB_ILE)))))
        v = sum(x for _, x in ent) / len(ent)
        add(comp, "ICPD-6 - Promedio parcial con los indicadores disponibles", "referencial", a, round(sum(x for _, x in ent), 2), len(ent), v, "Gris",
            "Los indicadores del PDM disponibles", "",
            f"Promedio simple de {len(ent)} de los 6 componentes ({', '.join(n for n, _ in ent)}); numerador = suma de puntajes, denominador = número de entradas. "
            f"Sin semáforo. {nota_ref[0].upper() + nota_ref[1:]}.")


def politicas():
    add("Cumplimiento del PDM", "Logros de políticas públicas (nueva dimensión)", "principal", 2025, None, None, None, "Gris",
        "Concejo Municipal (acuerdos), Planeación", "",
        "No disponible: no hay inventario oficial de políticas públicas vigentes en los documentos reunidos (solo constan, por separado, "
        "las de participación ciudadana, víctimas, salud mental y adicciones y biodiversidad). Construible armando el inventario con los acuerdos del Concejo; "
        "a) políticas con informe de seguimiento publicado / políticas vigentes sería calculable de inmediato; b) y c) quedan para después.")


# =============================================================================== INVENTARIO DE VARIABLES POR DIMENSIÓN
F_VARS = OUT / "variables_por_dimension.csv"
CAMPOS_VARS = ["dimension", "id_variable", "variable", "origen", "para_indicador", "estado", "anio", "valor", "unidad",
               "detalle_valor", "fuente", "enlace", "que_falta", "limitacion"]
VARS: list[dict] = []
DIM_CS, DIM_PC, DIM_PDM, DIM_PP = "Control social", "Participación ciudadana", "Cumplimiento del PDM", "Presupuesto participativo"


def fil_ind(ind, anio, tipo=None):
    c = [r for r in FILAS if r["indicador"] == ind and r["anio"] == anio and (tipo is None or r["tipo"] == tipo)]
    assert len(c) == 1, (ind, anio)
    return c[0]


def vf(ind, anio, campo="valor"):
    return num(fil_ind(ind, anio)[campo])


def vv(id_var, anio, periodo=None):
    for r in VAL:
        if (r["id_variable"] == id_var and r["anio"] == str(anio) and r["escenario_id"] == "defecto" and r["territorio"] == "Municipio"
                and (periodo is None or r["periodo"] == periodo)):
            return num(r["valor"])
    return None


def enl_var(id_var):
    for r in VAL:
        if r["id_variable"] == id_var and r["escenario_id"] == "defecto" and r["enlace"]:
            return r["enlace"].split("\n")[0]
    return ""


def var(dim, id_, nombre, origen, para, estado, fuente, enlace, falta, limit, unidad, valores):
    """valores: lista de (anio, valor, detalle). Una fila por año con valor; una fila sin año si no hay ninguno."""
    assert estado in ("Disponible", "Construible", "No disponible"), estado
    base = {"dimension": dim, "id_variable": id_, "variable": nombre, "origen": origen, "para_indicador": para, "estado": estado,
            "unidad": unidad, "fuente": fuente, "enlace": enlace, "que_falta": falta, "limitacion": limit}
    vals = [x for x in valores if x[1] is not None]
    if not vals:
        VARS.append({**base, "anio": "", "valor": "", "detalle_valor": ""})
    for a, v, d in vals:
        VARS.append({**base, "anio": a, "valor": round(float(v), 4) if isinstance(v, float) else v, "detalle_valor": d})


def variables_por_dimension():
    M, N = "Matriz", "Nueva"
    # ------------------------------------------------------------------ CONTROL SOCIAL
    ur_dsi = enl_var("CS.PCS.EE")
    var(DIM_CS, "CS.DSI.MS", "Metas del plan seleccionadas para el seguimiento (MS)", M, "DSI", "Construible",
        "Decisión del equipo sobre el Apéndice 9.1 del PDM", enl_var("CS.DSI.MS"),
        "Lista de metas aprobada por los profesores (la Matriz sugiere unas 10)", "No es un dato de la Alcaldía: es una decisión de diseño; con otra lista cambia el resultado",
        "metas", [(a, vv("CS.DSI.MS", a), "lista provisional del índice del PDM") for a in (2024, 2025)])
    dsi_r = "DSI - Respaldo: con todo lo publicado hasta la fecha de consulta"
    var(DIM_CS, "CS.DSI.MD", "Metas seleccionadas con valor observado publicado a tiempo (MD)", M, "DSI", "Construible",
        "Plan Indicativo de Productos y evaluaciones del PDM (DAP)", enl_var("CS.DSI.MD"),
        "Valores observados 2024 y 2025 de las metas seleccionadas, publicados oficialmente",
        f"Las evaluaciones publicadas solo cubren metas de producto; con lo publicado hasta hoy serían {fm_n(vf(dsi_r, 2024, 'numerador'))} de {fm_n(vf(dsi_r, 2024, 'denominador'))} en 2024",
        "metas", [(a, vv("CS.DSI.MD", a), "al cierre de la vigencia") for a in (2024, 2025)])
    var(DIM_CS, "CS.PCS.EE", "Estrategias de promoción del control social ejecutadas (EE)", M, "PCS", "Disponible",
        "Plan Indicativo de Productos (DAP)", enl_var("CS.PCS.EE"), "Soportes de ejecución (documento de estrategia 2024, informe de ciclos 2025); cierre de 2026",
        "Dato autodeclarado por la dependencia; el DAP anotó reparos a las evidencias",
        "estrategias", [(2024, vv("CS.PCS.EE", 2024), ""), (2025, vv("CS.PCS.EE", 2025), ""), (2026, vv("CS.PCS.EE", 2026), "corte 30 de junio")])
    var(DIM_CS, "CS.PCS.EP", "Estrategias de promoción del control social programadas (EP)", M, "PCS", "Disponible",
        "Plan Indicativo de Productos (DAP)", enl_var("CS.PCS.EP"), "Confirmar con el DAP que no hubo modificaciones de la programación",
        "Programación oficial de 1 por año (meta total 4)", "estrategias",
        [(2024, vv("CS.PCS.EP", 2024), ""), (2025, vv("CS.PCS.EP", 2025), ""), (2026, vv("CS.PCS.EP", 2026), "corte 30 de junio")])
    vr = "VRC - Respaldo: evento sin mención de veeduría = ausencia"
    var(DIM_CS, "CS.VRC.EV", "Eventos de rendición de cuentas con veedurías presentes (EV)", M, "VRC", "Construible",
        "Actas de rendición de cuentas (Secretaría de Salud y otras)", enl_var("CS.VRC.EV"),
        "Listas de asistencia anexas a las actas y acta del evento institucional",
        "Solo hay menciones textuales de veedurías; sin listas, la presencia no está comprobada. Valor con el supuesto del respaldo",
        "eventos", [(a, vf(vr, a, "numerador"), "supuesto: sin mención = ausencia") for a in (2024, 2025)])
    var(DIM_CS, "CS.VRC.ET", "Eventos municipales de rendición de cuentas realizados (ET)", M, "VRC", "Construible",
        "Actas y cuadros de rendición de cuentas", enl_var("CS.VRC.ET"), "Definición oficial de qué cuenta como evento municipal",
        "Según el universo elegido (solo la audiencia del Alcalde o todos los eventos con documento) la cifra va de 1 a 8 por año",
        "eventos", [(a, vv("CS.VRC.ET", a), "universo exploratorio") for a in (2024, 2025)])
    rp = "RPO - Respaldo: con el cuadro de preguntas y respuestas publicado"
    var(DIM_CS, "CS.RPO.OR", "Observaciones ciudadanas con respuesta de fondo publicada (OR)", M, "RPO", "Construible",
        "Cuadros de preguntas y respuestas de la rendición", enl_var("CS.RPO.OR"), "Clasificar cada respuesta como \"de fondo\" y fijar el universo evaluado",
        "2025: solo un cuadro publicado con 2 observaciones; el resultado es frágil", "observaciones",
        [(2024, vv("CS.RPO.OR", 2024), ""), (2025, vf(rp, 2025, "numerador"), "con el cuadro publicado")])
    var(DIM_CS, "CS.RPO.OT", "Observaciones ciudadanas registradas en rendiciones cerradas (OT)", M, "RPO", "Construible",
        "Cuadros de preguntas y respuestas de la rendición", enl_var("CS.RPO.OT"), "Separar felicitaciones y temas ajenos al PDM",
        "Los cuadros no distinguen las observaciones sobre el PDM", "observaciones",
        [(2024, vv("CS.RPO.OT", 2024), ""), (2025, vf(rp, 2025, "denominador"), "con el cuadro publicado")])

    # ------------------------------------------------------------------ PARTICIPACIÓN CIUDADANA
    var(DIM_PC, "PC.I1.V1", "Procesos cerrados que incorporaron un aporte ciudadano en una decisión formal", M, "I1", "No disponible",
        "Sin fuente: no hay registro", "", "Inventario de procesos cerrados que vincule cada aporte con la decisión",
        "Solo hay un vínculo narrativo en presupuesto participativo y en el PDM; la Matriz exige verlo proceso por proceso", "procesos", [])
    var(DIM_PC, "PC.I1.V2", "Procesos cerrados en que se recibieron aportes ciudadanos", M, "I1 e I4 (respaldo de I1)", "Construible",
        "Universo propuesto por el equipo", enl_var("PC.I1.V2"), "Definición oficial de proceso participativo cerrado",
        "Universo no oficial; riesgo de subconteo (consultas sectoriales e instancias no inventariadas)", "procesos",
        [(a, vv("PC.I1.V2", a), "exploratorio") for a in (2024, 2025)])
    var(DIM_PC, "PC.I2.V1", "Zonas con al menos una actividad participativa", M, "I2", "Construible",
        "Informe de gestión 2025, PDM y cronogramas", enl_var("PC.I2.V1"), "Actas o listas de asistencia por zona (fecha, zona y acta)",
        "Hay cronogramas y reportes agregados, no actas por zona", "zonas", [(a, vv("PC.I2.V1", a), "") for a in (2024, 2025)])
    var(DIM_PC, "PC.I2.V2", "Zonas oficiales del municipio", M, "I2", "Disponible",
        "Acuerdos municipales sobre zonificación", enl_var("PC.I2.V2"), "Verificar el texto del Acuerdo 007 de 2025",
        "Cifra sin verificar contra el acto que crea las zonas", "zonas", [(a, vv("PC.I2.V2", a), "") for a in (2024, 2025)])
    var(DIM_PC, "PC.I3.V1", "Acciones formativas para la ciudadanía realizadas", M, "I3", "Construible",
        "Informe de gestión 2025 (Escuela de Formación)", enl_var("PC.I3.V1"), "Relación oficial de acciones realizadas, sin formación a servidores",
        "El reporte oficial cuenta personas formadas, no acciones; 2024 se reconstruye de una lista", "acciones",
        [(a, vv("PC.I3.V1", a), "") for a in (2024, 2025)])
    var(DIM_PC, "PC.I3.V2", "Acciones formativas programadas con finalización en el periodo", M, "I3", "Construible",
        "Informe de gestión 2025 (Escuela de Formación)", enl_var("PC.I3.V2"), "Plan de formación aprobado con las acciones programadas",
        "No hay programación oficial de acciones (la meta está en personas); 2025 es una reconstrucción; 2024 sin dato", "acciones",
        [(2025, vv("PC.I3.V2", 2025), "inferido")])
    var(DIM_PC, "PC.I4.V1", "Procesos cerrados con resultados publicados y consultables", M, "I4", "No disponible",
        "Portal de la Alcaldía", enl_var("PC.I4.V1"), "Inventario de procesos y revisión de cada publicación de resultados",
        "No se ha verificado que cada publicación comunique aportes y decisiones", "procesos", [])
    var(DIM_PC, "PC.I4.V2", "Procesos participativos cerrados (universo)", M, "I4", "Construible",
        "Universo propuesto por el equipo", enl_var("PC.I4.V2"), "Definición oficial de proceso participativo cerrado",
        "Mismo universo no oficial que I1: una sola lista debe servir a ambos indicadores", "procesos",
        [(a, vv("PC.I4.V2", a), "exploratorio") for a in (2024, 2025)])

    # ------------------------------------------------------------------ CUMPLIMIENTO DEL PDM
    pn = lambda *x: x
    var(DIM_PDM, "PDM.CM.O", "Tasa de mortalidad vial por 100.000 habitantes", M, "CM", "Construible",
        "Forensis (INMLCF); Policía Nacional", "https://terridata.dnp.gov.co/", "Cifra de 2025 (Forensis aún no la publica) y de 2026",
        "Solo 2024; la Policía y el INMLCF miden distinto (la Policía da más en 2024)", "por 100.000 hab.", [(2024, vv("PDM.CM.O", 2024), "INMLCF")])
    var(DIM_PDM, "PDM.CM.M", "Meta de mortalidad vial", M, "CM", "Disponible", "PDM 2024-2027", enl_var("PDM.CM.M"), "Nada",
        "Solo meta final 2027 (sin metas anuales); línea base 10,2", "por 100.000 hab.", [(2027, vv("PDM.CM.M", 2024), "meta final del plan")])
    var(DIM_PDM, "PDM.CA.O", "Espacio público verde observado", M, "CA", "Construible",
        "Diagnóstico del POT 2024", enl_var("PDM.CA.O"), "Cifras de 2025 y 2026; conciliar con la línea base del plan",
        "No concilia con la línea base del plan (3,88) ni con la del AMVA (4,66): otra definición de área", "m² por habitante",
        [(2024, vv("PDM.CA.O", 2024), "EPV urbano")])
    var(DIM_PDM, "PDM.CA.M", "Meta de espacio público verde", M, "CA", "Disponible", "PDM 2024-2027", enl_var("PDM.CA.M"), "Nada",
        "Solo meta final 2027", "m² por habitante", [(2027, vv("PDM.CA.M", 2024), "meta final del plan")])
    var(DIM_PDM, "PDM.CP.O", "Implementación de los procesos de participación (método oficial)", M, "CP", "No disponible",
        "PDM 2024-2027 (solo diagnóstico)", enl_var("AUX.PDM.CP.INDICE_SITUACIONAL"), "Ficha oficial del indicador (fórmula, valor y fuente) pedida al DAP",
        f"El PDM usa dos magnitudes distintas (un índice de {fm_n(vv('AUX.PDM.CP.INDICE_SITUACIONAL', 2022))} en 2022 y un porcentaje de 75 en 2023)", "%", [])
    var(DIM_PDM, "PDM.CP.M", "Meta de implementación de los procesos de participación", M, "CP", "Disponible", "PDM 2024-2027", enl_var("PDM.CP.M"), "Nada",
        "Solo meta final 2027; línea base 75", "%", [(2027, vv("PDM.CP.M", 2024), "meta final del plan")])
    dd, dm = vv("AUX.PDM.CD.DISCIPLINAS", 2025), vv("AUX.PDM.CD.DISCIPLINAS_META", 2025)
    var(DIM_PDM, "PDM.CD.O", "Cobertura de los servicios deportivos", M, "CD", "No disponible",
        "Informe de gestión (Índer)", enl_var("AUX.PDM.CD.DISCIPLINAS"), "Numerador, denominador y valor de la cobertura (pedido al Índer)",
        f"Solo hay disciplinas implementadas ({fm_n(dd)} de {fm_n(dm)} en 2025): mide oferta, no población atendida", "%", [])
    var(DIM_PDM, "PDM.CD.M", "Meta de cobertura de servicios deportivos", M, "CD", "Disponible", "PDM 2024-2027", enl_var("PDM.CD.M"), "Nada",
        "Solo meta final 2027; línea base 5,3", "%", [(2027, vv("PDM.CD.M", 2024), "meta final del plan")])
    ile = "ILE primera infancia - cobertura de atención integral"
    rango = INFO_ILE["rango"]
    var(DIM_PDM, "PDM.ILE.A", f"Cobertura de atención integral de primera infancia ({rango})", M, "ILE (primera infancia)", "Construible",
        "Cálculo propio: atendidos / población DANE", URL_TERRIDATA, "Confirmar con la Secretaría de Educación el rango de edad de la línea base",
        f"Con {rango}, 2023 da {fm_n(vf(ile, 2023))}, no los 16,3 de la línea base: el rango exacto del plan no está documentado", "%",
        [(a, vf(ile, a), "") for a in ANIOS_SERIE])
    var(DIM_PDM, "PDM.ILE.MA", "Meta de cobertura de atención integral", M, "ILE (primera infancia)", "Disponible", "PDM 2024-2027", enl_var("PDM.ILE.MA"), "Nada",
        "Solo meta final 2027; línea base 16,3", "%", [(2027, vv("PDM.ILE.MA", 2024), "meta final del plan")])
    var(DIM_PDM, "PDM.ILE.S", "Instituciones oficiales que mejoran su categoría en Saber 11", M, "ILE (Saber 11)", "No disponible",
        "ICFES (solo menciones de casos)", URL_SABER, "Clasificación de planteles 2023 a 2025 y total de instituciones oficiales",
        "Hay casos aislados (una institución en 2024 y una en 2025), no el porcentaje", "%", [])
    var(DIM_PDM, "PDM.ILE.MS", "Meta de instituciones que mejoran su categoría", M, "ILE (Saber 11)", "Disponible", "PDM 2024-2027", enl_var("PDM.ILE.MS"), "Nada",
        "Solo meta final 2027; línea base 0", "%", [(2027, vv("PDM.ILE.MS", 2024), "meta final del plan")])
    # nuevas: salud
    af = INFO_ILE["afiliados"]
    for k, (nom, idv) in enumerate((("Afiliados al régimen contributivo", "PDM.SAL.CONTRIB"), ("Afiliados al régimen subsidiado", "PDM.SAL.SUBS"),
                                    ("Afiliados a regímenes especiales", "PDM.SAL.ESP"), ("Afiliados al SGSSS", "PDM.SAL.TOTAL"))):
        serie = af[nom]
        est = "Disponible" if all(a in serie for a in (2024, 2025)) else "Construible"
        var(DIM_PDM, idv, nom.replace("SGSSS", "sistema de salud (total)"), N, "Cobertura de salud", est,
            "TerriData (MinSalud), diciembre", URL_TERRIDATA,
            {"Afiliados a regímenes especiales": "Cifra de 2026 al cierre; 2019 y 2020 sin dato publicado",
             "Afiliados al SGSSS": "Cifra de 2026 al cierre; 2019 sin total publicado"}.get(nom, "Cifra de 2026 al cierre"),
            "Cortes de diciembre",
            "personas", [(a, serie[a], "") for a in ANIOS_SERIE if a in serie])
    pp_ = INFO_ILE["pop_pped"]
    var(DIM_PDM, "PDM.SAL.POB", "Población total del municipio (proyección DANE 2018-2042)", N, "Cobertura de salud", "Disponible",
        "DANE, proyecciones de población", URL_DANE, "Nada",
        "Una sola proyección para toda la serie (comparable); MinSalud usa en cada corte otra población, por eso difiere de sus cifras", "personas",
        [(a, pp_[a], "") for a in ANIOS_SERIE])
    atend, fat, eat = INFO_ILE["atend"]
    var(DIM_PDM, "PDM.ILE.ATEND", "Niños, niñas y gestantes atendidos de forma integral (educación inicial)", N, "ILE (primera infancia)", "Disponible",
        "TerriData (MinEducación) hasta 2023; informe de gestión de la Alcaldía 2024 y 2025", URL_TERRIDATA,
        "Confirmar con la Secretaría de Educación las cifras de 2024 y 2025", "Fuente distinta antes y después de 2023; incluye gestantes", "personas",
        [(a, atend[a], fat[a]) for a in ANIOS_SERIE])
    var(DIM_PDM, "PDM.ILE.POB05", f"Población de {rango} (DANE, edad simple)", N, "ILE (primera infancia)", "Disponible",
        "DANE, proyección de población 2018-2042 por edad simple", URL_DANE, "Nada", "Proyección, no censo; se revisa cada vez que el DANE actualiza", "personas",
        [(a, INFO_ILE["pob_edad"][a], "") for a in ANIOS_SERIE])
    sab = INFO_ILE["saber"]
    var(DIM_PDM, "PDM.ILE.SABER_OU", "Puntaje global promedio de Saber 11, instituciones oficiales urbanas", N, "ILE (Saber 11, respaldo)", "Construible",
        "ICFES, resultados Saber 11 (datos.gov.co); promedio propio ponderado", URL_SABER, "Resultados de 2021 y de 2023 a 2025",
        "El conjunto público termina en 2022; es el respaldo mientras no haya clasificación de planteles", "puntos",
        [(a, sab["oficial_urbano"][a], "") for a in ANIOS_SERIE if a in sab["oficial_urbano"]])
    var(DIM_PDM, "PDM.ILE.SABER_TD", "Puntaje global promedio de Saber 11, todos los planteles", N, "ILE (Saber 11, respaldo)", "Construible",
        "ICFES, resultados Saber 11 (datos.gov.co); promedio propio ponderado", URL_SABER, "Resultados de 2021 y de 2023 a 2025",
        "Promedio de los calendarios A y B; el conjunto público termina en 2022", "puntos",
        [(a, sab["municipio"][a], "") for a in ANIOS_SERIE if a in sab["municipio"]])
    var(DIM_PDM, "PDM.ILE.CLASIF", "Clasificación de planteles de Saber 11 (A+, A, B, C, D) por institución oficial", N, "ILE (Saber 11)", "No disponible",
        "ICFES (no publica la de Envigado en datos abiertos)", "", "Solicitarla a la Secretaría de Educación, 2023 a 2025",
        "Sin ella no se puede calcular el avance hacia la meta de 14 %", "instituciones", [])
    ult, tot_d = INFO_ILE["discap"]
    var(DIM_PDM, "PDM.INC.DISCAP", "Personas con discapacidad en el registro nacional (acumulado)", N, "Inclusión (respaldo de salud)", "Construible",
        "MinSalud, Registro de Localización y Caracterización (datos.gov.co)", URL_RLCPD, "Registro actualizado del municipio (Secretaría de Salud o de Inclusión Social)",
        f"Acumulado desde 2003; el conjunto público no se actualiza desde {ult}; no es una tasa", "personas", [(ult, int(tot_d), "acumulado")])
    pol = "Concejo Municipal (acuerdos) y Planeación"
    var(DIM_PDM, "PDM.POL.VIG", "Políticas públicas municipales vigentes (inventario)", N, "Logros de políticas públicas", "Construible", pol, "",
        "Armar la lista con los acuerdos del Concejo y confirmarla con Planeación", "Hoy solo constan, por separado, 4 políticas (participación, víctimas, salud mental y biodiversidad)", "políticas", [])
    var(DIM_PDM, "PDM.POL.INF", "Políticas vigentes con informe de seguimiento publicado", N, "Logros de políticas públicas", "Construible", pol, "",
        "Revisar el portal por política una vez exista el inventario", "Depende del inventario de vigentes", "políticas", [])
    var(DIM_PDM, "PDM.POL.PLAN", "Políticas vigentes con plan de acción vigente", N, "Logros de políticas públicas (futuro)", "No disponible", pol, "",
        "Planes de acción por política", "Sin fuente identificada", "políticas", [])
    var(DIM_PDM, "PDM.POL.METAS", "Metas de las políticas cumplidas y metas totales", N, "Logros de políticas públicas (futuro)", "No disponible", pol, "",
        "Metas por política con su seguimiento", "Sin fuente identificada", "metas", [])

    # ------------------------------------------------------------------ PRESUPUESTO PARTICIPATIVO
    est24 = {a: [r for r in PPP if r["anio"] == str(a) and r["en_listado_efpc"] == "True"] for a in (2024, 2025)}
    url25 = enlace_de("PP.CEP", 2025)
    var(DIM_PP, "PP.VOT", "Valor votado por proyecto (listado de proyectos votados)", N, "Fidelidad a lo votado", "Disponible",
        "Listado de proyectos votados (EFPC)", URL_EFPC24, "Valor votado de cada proyecto de 2026",
        "Es el valor del proyecto en el listado; 2026 no lo trae", "pesos",
        [(a, round(sum(num(r["valor_listado_efpc"]) for r in est24[a])), f"suma de {len(est24[a])} proyectos") for a in (2024, 2025)])
    var(DIM_PP, "PP.CEP.PE", "Proyectos votados (número)", M, "Fidelidad, terminados y respaldos", "Disponible",
        "Listado de proyectos votados (EFPC)", URL_EFPC24, "Lista oficial de códigos de proyecto del presupuesto participativo",
        "No hay marcador de presupuesto participativo en Hacienda: el universo se reconstruye cruzando el listado con los códigos del plan de acción", "proyectos",
        [(a, vv("PP.CEP.PE", a), "") for a in (2024, 2025, 2026)])
    var(DIM_PP, "PP.EPP.APP", "Apropiación definitiva por proyecto (APP)", M, "EPP, fidelidad y respaldos", "Construible",
        "Hacienda y plan de acción", enl_var("PP.EPP.APP"), "Cierre de 2026; en 2024 un proyecto de 51 no se pudo cruzar con Hacienda",
        "Universo reconstruido y por validar con Planeación; 2024 es aproximado (cruce por nombre y valor)", "pesos",
        [(2024, vv("PP.EPP.APP", 2024), "50 de 51 proyectos"), (2025, vv("PP.EPP.APP", 2025), "61 proyectos"), (2026, vv("PP.EPP.APP", 2026), "corte 30 de junio")])
    comp25 = round(sum(num(r["compromiso"]) for r in est24[2025]))
    var(DIM_PP, "PP.CMP", "Compromisos por proyecto", N, "Tasa de compromiso", "Construible",
        "Hacienda (plan de acción 2025)", url25, "Compromisos separados de obligaciones en 2024 y 2026",
        "En 2024 Hacienda no separa compromisos de obligaciones; 2026 sin dato", "pesos", [(2025, comp25, "61 proyectos")])
    var(DIM_PP, "PP.EPP.OPP", "Obligaciones por proyecto (OPP)", M, "EPP y respaldos", "Construible",
        "Hacienda y plan de acción", enl_var("PP.EPP.OPP"), "Cierre de 2026",
        "Mismas limitaciones que la apropiación; 2026 es provisional (primer semestre)", "pesos",
        [(2024, vv("PP.EPP.OPP", 2024), "50 de 51 proyectos"), (2025, vv("PP.EPP.OPP", 2025), "61 proyectos"), (2026, vv("PP.EPP.OPP", 2026), "corte 30 de junio")])
    e25 = [r for r in ENT if r["anio"] == "2025" and r["cuenta_en_defecto_pe"] == "1"]
    con_dato = sum(1 for r in e25 if r["n_filas_dic"])
    var(DIM_PP, "PP.AVF", "Avance físico por proyecto (plan de acción)", N, "Proyectos terminados", "Construible",
        "Plan de acción 2025 (autorreporte)", url25, "Avance físico por proyecto de 2024 y al cierre de 2026",
        "Autorreporte, sin acta de recibo; 3 proyectos traen valores no numéricos", "proyectos con avance reportado", [(2025, con_dato, "de 61 proyectos")])
    var(DIM_PP, "PP.CEP.PEC", "Proyectos votados con avance físico de 100 % (autorreporte)", M, "Proyectos terminados", "Construible",
        "Plan de acción 2025 (autorreporte)", enl_var("PP.CEP.PEC"), "Actas de recibo, informes finales o de supervisión por proyecto",
        "Autorreporte: con otra regla de agregación el número va de 45 a 50 (la cifra de 55 citada antes no se reproduce)", "proyectos", [(2025, vv("PP.CEP.PEC", 2025), "todas las filas de diciembre al 100 %")])
    var(DIM_PP, "PP.TER.APP", "Apropiación de los proyectos terminados", N, "Proyectos terminados (respaldo)", "Construible",
        "Hacienda y plan de acción 2025", url25, "Avance físico de 2024 y 2026", "Hereda las limitaciones del autorreporte", "pesos",
        [(2025, round(INFO_PP["ap_t"]), f"{INFO_PP['terminados']} proyectos")])
    var(DIM_PP, "PP.TER.OPP", "Obligaciones de los proyectos terminados", N, "Proyectos terminados (respaldo)", "Construible",
        "Hacienda y plan de acción 2025", url25, "Avance físico de 2024 y 2026", "Hereda las limitaciones del autorreporte", "pesos",
        [(2025, round(INFO_PP["ob_t"]), f"{INFO_PP['terminados']} proyectos")])
    var(DIM_PP, "PP.ACTA", "Actas de recibo o informes finales por proyecto", N, "Proyectos terminados (verificación)", "No disponible",
        "Sin fuente: no se encontraron publicadas", "", "Pedirlas a la Subdirección de Participación y a las unidades ejecutoras",
        "Sin ellas el indicador sigue siendo autorreporte", "documentos", [])
    for idv, nom in (("PP.CMA.NA", "Proyectos con meta de atención aplicable (NA)"), ("PP.CMA.MP", "Meta anual de personas por proyecto (MP)"),
                     ("PP.CMA.PA", "Personas atendidas por proyecto (PA)")):
        var(DIM_PP, idv, nom, M, "Meta de atención (fuera por ahora)", "No disponible", "Plan de acción (mide actividades, no personas)", enl_var(idv),
            "Ficha del proyecto con meta anual de personas", "El indicador de meta de atención sale por ahora del sistema", "proyectos" if "NA" in idv else "personas", [])
    return VARS


def fm_n(x):
    if x is None:
        return "sin dato"
    return (f"{int(round(x)):,}".replace(",", ".") if abs(x - round(x)) < 1e-9 else f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


def escribir_variables():
    variables_por_dimension()
    C.escribir_csv(F_VARS, VARS, CAMPOS_VARS)
    ids = {v["id_variable"] for v in VARS}
    print(f"{len(VARS)} filas ({len(ids)} variables) -> {F_VARS.relative_to(RAIZ)}")


def main():
    for f, s in ((DESC / "icfes" / "saber11_envigado_agregado_kgxf-xxbe.csv", "descargas_nuevas.py"),):
        if not f.exists():
            subprocess.run([sys.executable, str(RAIZ / "scripts" / "sistema_simplificado" / s)], check=True)
    td = leer_terridata()
    pop_pped = poblacion_pped_total()
    pob_edad = poblacion_edad()
    presupuesto_participativo()
    participacion()
    control_social()
    pdm_otros()
    ile_primera_infancia(td, pob_edad)
    ile_saber()
    cobertura_salud(td, pop_pped)
    discapacidad()
    icpd6()
    politicas()
    C.escribir_csv(F_SALIDA, FILAS, CAMPOS)
    escribir_variables()
    print(f"{len(FILAS)} filas -> {F_SALIDA.relative_to(RAIZ)}")
    print("ILE calibración:", INFO_ILE["cal"], "->", INFO_ILE["elegido"], "|", INFO_ILE["motivo"])


if __name__ == "__main__":
    main()
