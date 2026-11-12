"""Excel de presentación con fórmulas vivas del sistema de indicadores simplificado.

Lee (sin recalcular resultados):
  - outputs/sistema_simplificado/variables_por_dimension.csv (variables por año, fuente, enlace, estado);
  - outputs/sistema_simplificado/indicadores_simplificados.csv (conteos de apoyo, enlaces);
  - outputs/indicadores/resultados_indicadores.csv (ejecución por zona del presupuesto participativo);
  - outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx, hoja "Indicadores" (veredicto de la Matriz original).
Escribe outputs/sistema_simplificado/indicadores_OCE_presentacion.xlsx. Todo es PROVISIONAL.
Cada indicador es una fórmula de Excel que referencia las celdas de variables de su misma hoja.
Ejecutar desde la raíz: python3 scripts/sistema_simplificado/generar_excel_formulas.py [ruta_mapa_json]
"""
import csv
import json
import math
import re
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl
import openpyxl.worksheet.properties
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

RAIZ = Path(__file__).resolve().parents[2]
D = RAIZ / "outputs/sistema_simplificado"
F_VAR = D / "variables_por_dimension.csv"
F_IND = D / "indicadores_simplificados.csv"
F_RES = RAIZ / "outputs/indicadores/resultados_indicadores.csv"
F_DISP = RAIZ / "outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx"
SAL = D / "indicadores_OCE_presentacion.xlsx"
MAPA = Path(sys.argv[1]) if len(sys.argv) > 1 else None

AZUL, AZUL_CLARO, ENC, GRIS_TXT = "1F4E78", "DDEBF7", "BDD7EE", "7F7F7F"
FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
F_SEM = {"V": "C6EFCE", "A": "FFEB9C", "N": "F8CBAD", "R": "FFC7CE"}
FN = "Calibri"
POR_QUE_FILL = {"Sí": "C6EFCE", "Con ajuste": "FFEB9C", "No": "FFC7CE", "Eliminado": "D9D9D9", "Nuevo": "BDD7EE"}
EST_VAR_FILL = {"Disponible": "C6EFCE", "Construible": "FFEB9C", "No disponible": "FFC7CE"}

# Umbrales por indicador (Matriz / propuesta): lista descendente (umbral, color); último = resto.
U_9070 = [(90, "V"), (70, "A"), (None, "R")]
U_7550 = [(75, "V"), (50, "A"), (None, "R")]
U_8050 = [(80, "V"), (50, "A"), (None, "R")]
U_PDM = [(100, "V"), (75, "A"), (None, "R")]
U_PP = [(90, "V"), (75, "A"), (50, "N"), (None, "R")]
NOMBRE_SEM = {"V": "Verde", "A": "Amarillo", "N": "Naranja", "R": "Rojo"}
TXT_UMBRAL = {
    "U_9070": "Verde ≥ 90 · Amarillo 70 a < 90 · Rojo < 70", "U_7550": "Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50",
    "U_8050": "Verde ≥ 80 · Amarillo 50 a < 80 · Rojo < 50", "U_PDM": "Verde solo 100 · Amarillo 75 a < 100 · Rojo < 75",
    "U_PP": "Verde ≥ 90 · Amarillo 75 a < 90 · Naranja 50 a < 75 · Rojo < 50",
}
UMB = {"U_9070": U_9070, "U_7550": U_7550, "U_8050": U_8050, "U_PDM": U_PDM, "U_PP": U_PP}


def num(s):
    s = (s or "").strip()
    if s == "":
        return None
    v = float(s)
    return int(v) if v == int(v) and abs(v) < 1e15 else v


def leer(f, **kw):
    with open(f, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh, **kw))


VARS_CSV = leer(F_VAR, delimiter=";")
IND_CSV = leer(F_IND, delimiter=";")
LINK_IND = {}
for r in IND_CSV:
    LINK_IND.setdefault(r["indicador"], (r["fuente"], r["enlace"]))

# ---------------------------------------------------------------- variables
DIM_HOJA = {"Control social": "Control social", "Participación ciudadana": "Participación ciudadana",
            "Cumplimiento del PDM": "Plan de Desarrollo", "Presupuesto participativo": "Presupuesto participativo"}
VARS = OrderedDict()  # id -> dict
for r in VARS_CSV:
    v = VARS.setdefault(r["id_variable"], dict(id=r["id_variable"], hoja=DIM_HOJA[r["dimension"]], nombre=r["variable"],
                                              unidad=r["unidad"], estado=r["estado"], vals={}, fuente=r["fuente"],
                                              enlace=r["enlace"], falta=r["limitacion"] or r["que_falta"]))
    if r["anio"] and r["valor"] != "":
        v["vals"][int(r["anio"])] = num(r["valor"])
VARS["PP.VOT"]["nombre"] = "Valor votado total del listado de proyectos votados"
VARS.pop("PDM.ILE.A", None)  # se recalcula a partir de atendidos y población


def extra(id_, hoja, nombre, unidad, estado, vals, fuente=None, enlace=None, falta="", ref_ind=None, antes=None):
    if ref_ind:
        fuente, enlace = LINK_IND[ref_ind]
    VARS[id_] = dict(id=id_, hoja=hoja, nombre=nombre, unidad=unidad, estado=estado, vals=vals, fuente=fuente,
                     enlace=enlace, falta=falta)


PDM, CS, PC, PP = "Plan de Desarrollo", "Control social", "Participación ciudadana", "Presupuesto participativo"
extra("CS.DSI.MDT", CS, "Metas seleccionadas con valor publicado hasta la fecha de consulta (sin exigir oportunidad)", "metas",
      "Construible", {2024: 2}, ref_ind="DSI - Respaldo: con todo lo publicado hasta la fecha de consulta",
      falta="Para 2025 aún no se ha contado")
extra("PC.I2.ASAM", PC, "Zonas con asamblea de presupuesto participativo", "zonas", "Construible", {2024: 12, 2025: 12},
      ref_ind="I2 - Respaldo: zonas con asamblea de presupuesto participativo", falta="La zona 1 no tiene comité zonal")
extra("PC.I3.CULM", PC, "Acciones formativas culminadas (incluye las dirigidas a servidores)", "acciones", "Construible",
      {2024: 7, 2025: 21}, ref_ind="I3 - Respaldo: acciones formativas culminadas (número)",
      falta="Incluye acciones para servidores; sin meta")
extra("PDM.CM.POL", PDM, "Tasa de mortalidad vial según la Policía Nacional (homicidios en accidente de tránsito)",
      "por 100.000 hab.", "Construible", {2024: 8.054642696050003}, ref_ind="CM - Respaldo: cifra de la Policía",
      falta="Otra definición que la del Instituto de Medicina Legal")
extra("PDM.CA.POT", PDM, "Espacio público verde total del POT", "m² por habitante", "Construible", {2024: 69.41},
      ref_ind="CA - Respaldo: espacio público verde total del POT", falta="Otra serie: no comparable con la línea base del plan")
extra("PDM.CD.IMP", PDM, "Disciplinas deportivas atendidas", "disciplinas", "Construible", {2025: 38},
      ref_ind="CD - Respaldo: disciplinas deportivas atendidas", falta="Mide oferta, no población atendida")
extra("PDM.CD.PREV", PDM, "Disciplinas deportivas previstas", "disciplinas", "Construible", {2025: 40},
      ref_ind="CD - Respaldo: disciplinas deportivas atendidas", falta="Mide oferta, no población atendida")
extra("PDM.CP.IDX", PDM, "Índice situacional de implementación de los procesos de participación (diagnóstico del plan)",
      "índice", "Disponible", {2022: 3.23}, ref_ind="CP - Respaldo: índice situacional de implementación",
      falta="Sin meta comparable")
extra("PDM.ILE.LB", PDM, "Línea base de cobertura de atención integral de primera infancia (PDM)", "%", "Disponible",
      {"M": 16.3}, ref_ind="ILE primera infancia - cobertura de atención integral",
      falta="Con la población de 0 a 5 años, 2023 da 16,56 y no 16,3")
extra("PDM.SAL.POBMS", PDM, "Población usada por MinSalud en su cifra publicada", "personas", "Disponible",
      {2025: 250767, 2026: 268990}, ref_ind="Cobertura de salud, cifra publicada por MinSalud (diciembre)",
      falta="Población distinta de la proyección DANE usada en la serie")
extra("PDM.SAL.TOT26", PDM, "Afiliados al sistema de salud, corte de agosto de 2026", "personas", "Disponible",
      {2026: 222151}, ref_ind="Cobertura de salud, cifra publicada por MinSalud (agosto)", falta="Corte parcial del año")
# Presupuesto participativo
ind_pp = {r["indicador"]: r for r in IND_CSV if r["componente"] == "Presupuesto participativo"}


def csvnum(ind, anio, campo):
    for r in IND_CSV:
        if r["indicador"] == ind and r["anio"] == str(anio):
            return num(r[campo])


extra("PP.VOT.C", PP, "Valor votado de los proyectos considerados en la versión monetaria", "pesos", "Construible",
      {2024: csvnum("Fidelidad a lo votado, versión monetaria", 2024, "denominador"),
       2025: csvnum("Fidelidad a lo votado, versión monetaria", 2025, "denominador")},
      ref_ind="Fidelidad a lo votado, versión monetaria", falta="2024: 50 de los 51 proyectos se pudieron cruzar con Hacienda")
extra("PP.FID.N", PP, "Proyectos con apropiación definitiva igual al valor votado (tolerancia de ± 1 %)", "proyectos", "Construible",
      {2024: 39, 2025: 58}, ref_ind="Fidelidad a lo votado",
      falta="La tolerancia de 1 % es un parámetro de la propuesta; 2026 sin el valor votado de cada proyecto")
extra("PP.PRES", PP, "Proyectos votados con apropiación asignada", "proyectos", "Construible", {2024: 50, 2025: 61},
      ref_ind="I1 - Respaldo: proyectos votados con presupuesto asignado", falta="Usa los mismos datos de la fidelidad")
extra("PP.PROY.EV", PP, "Proyectos con ejecución presupuestal identificada (universo del indicador de 90 %)", "proyectos",
      "Construible", {2024: 50, 2025: 61}, ref_ind="Proyectos con ejecución de al menos 90 %",
      falta="2024: 50 proyectos cruzados con Hacienda")
extra("PP.PROY90", PP, "Proyectos con obligaciones de al menos 90 % de su apropiación", "proyectos", "Construible",
      {2024: 48, 2025: 57}, ref_ind="Proyectos con ejecución de al menos 90 %", falta="")
# ejecución por zona
ZONA = {}
for r in leer(F_RES, delimiter=";"):
    if r["codigo"] == "PP.EPP" and r["es_defecto"] == "True" and r["territorio"].startswith("Zona"):
        a = int(r["anio"])
        z = int(r["territorio"].split()[1])
        ZONA[(z, a)] = (num(r["entrada_1"]), num(r["entrada_2"]))
fz, lz = LINK_IND["Equidad entre zonas (brecha de ejecución)"]
for z in range(1, 14):
    for k, (nombre, i) in enumerate([("Obligaciones de la zona %d" % z, 0), ("Apropiación definitiva de la zona %d" % z, 1)]):
        extra("PP.Z%d.%s" % (z, "O" if i == 0 else "A"), PP, nombre + " (proyectos de presupuesto participativo)", "pesos",
              "Construible", {a: ZONA[(z, a)][i] for a in (2024, 2025, 2026) if (z, a) in ZONA},
              fuente=fz, enlace=lz, falta="Zona 1 sin proyectos" if z == 1 else "")
VARS["PP.EPP.OPP"]["falta"] = VARS["PP.EPP.OPP"]["falta"] or ""

# ---------------------------------------------------------------- matriz original
DISP = {}
wb0 = openpyxl.load_workbook(F_DISP, read_only=True)
for row in list(wb0["Indicadores"].iter_rows(values_only=True))[1:]:
    if row[1]:
        DISP[row[1]] = row


def bloque_a():
    """Por dimensión: (código y nombre original, ¿se puede?, por qué, qué lo reemplaza)."""
    return {
        CS: [
            ("CS.PCS Cumplimiento de la promoción del control social", "Sí",
             "Las estrategias programadas y ejecutadas están publicadas en el plan indicativo.",
             "Se mantiene como PCS; respaldo: número de estrategias ejecutadas."),
            ("CS.DSI Disponibilidad de información de seguimiento", "Con ajuste",
             "Hay que fijar la lista de metas a seguir; con la lista provisional de 7 metas el resultado es 0.",
             "Se mantiene como DSI; respaldo: con todo lo publicado hasta la fecha de consulta."),
            ("CS.VRC Presencia de veedurías en la rendición de cuentas", "Con ajuste",
             "La presencia de veedurías solo consta como mención en las actas; sin listas de asistencia queda sin dato.",
             "Se mantiene como VRC; respaldo: evento sin mención de veeduría cuenta como ausencia."),
            ("CS.RPO Respuesta pública a observaciones ciudadanas", "Con ajuste",
             "El cuadro de preguntas y respuestas de 2024 sí permite contar; el de 2025 está incompleto.",
             "Se mantiene como RPO; respaldo 2025: con el único cuadro publicado."),
            ("CS.ICRCS Índice de Condiciones y Respuesta para el Control Social", "Con ajuste",
             "Necesita las cuatro entradas y VRC y RPO 2025 no tienen dato completo.",
             "Se mantiene el ICRCS; respaldo: promedio con los respaldos de VRC y RPO."),
        ],
        PC: [
            ("PC.I1 Incorporación de aportes ciudadanos", "No",
             "No existe un inventario oficial de procesos cerrados ni un registro de aporte y decisión.",
             "Respaldo: proyectos votados con presupuesto asignado (mismos datos de la fidelidad de PP; no sumar dos veces)."),
            ("PC.I2 Cobertura territorial de la participación", "Con ajuste",
             "Hay 13 zonas oficiales; las actividades por zona salen de reportes agregados, sin actas por zona.",
             "Se mantiene como I2; respaldo: zonas con asamblea de presupuesto participativo."),
            ("PC.I3 Cumplimiento de la oferta formativa", "Con ajuste",
             "La programación oficial no está en acciones; la de 2025 se reconstruyó y 2024 no tiene programación.",
             "Se mantiene como I3; respaldo: número de acciones culminadas."),
            ("PC.I4 Publicación de resultados participativos", "No",
             "Comparte el universo de procesos cerrados que no existe y no se auditó documento por documento.",
             "Respaldo: procesos del universo exploratorio (conteo, sin semáforo)."),
            ("PC.I5 Cumplimiento de sesiones de participación", "Eliminado",
             "Los profesores lo eliminaron de la Matriz.", "Se elimina; los pesos del ISPC no cambian."),
            ("PC.I6 Disponibilidad de participación virtual", "Eliminado",
             "Los profesores lo eliminaron de la Matriz.", "Se elimina; los pesos del ISPC no cambian."),
            ("PC.ISPC Índice Simplificado de Participación Ciudadana", "No",
             "Faltan I1 e I4 y la Guía impide redistribuir pesos o calcular con entradas faltantes.",
             "Respaldo: ISPC con pesos redistribuidos entre I2 e I3 (0,556 y 0,444), rotulado como opción."),
        ],
        PDM: [
            ("PDM.CS Logro de la cobertura de aseguramiento subsidiado", "No",
             "La cobertura con el método del plan no se publica; con los datos públicos no se puede reproducir.",
             "Reemplazado por Cobertura de salud (afiliados / población), que ya está construida."),
            ("PDM.CM Logro de la reducción de mortalidad vial", "Con ajuste",
             "Solo 2024 con Medicina Legal; la línea base del plan no tiene fuente declarada.",
             "Se mantiene como CM; respaldo: cifra de la Policía Nacional."),
            ("PDM.CA Logro del mantenimiento de espacio público verde", "Con ajuste",
             "Solo hay la cifra 2024 del POT, que no concilia con la línea base del plan.",
             "Se mantiene como CA; respaldo: espacio público verde total del POT."),
            ("PDM.CP Logro de implementación de procesos participativos", "No",
             "No hay valor ni método oficial del porcentaje de implementación.",
             "Respaldo: índice situacional del diagnóstico del plan (3,23, año 2022); falta pedir la ficha técnica."),
            ("PDM.CD Logro de cobertura de servicios deportivos", "No",
             "Solo hay oferta de disciplinas, que la Matriz no permite convertir en cobertura.",
             "Respaldo: disciplinas atendidas de las previstas (mide oferta); falta la cobertura del Índer."),
            ("PDM.ILE Índice de Logro Educativo", "No",
             "Falta la clasificación de planteles de Saber 11 por institución, necesaria para uno de sus dos componentes.",
             "Reemplazado por el ILE como evolución: serie de primera infancia, variación anual y avance hacia la meta 2027."),
            ("PDM.ICPD-6 Índice de cumplimiento de resultados seleccionados del PDM", "No",
             "Cuatro de sus seis entradas no tienen dato y la Guía impide calcularlo incompleto.",
             "Respaldo referencial: promedio de las entradas disponibles (no cumple la regla de la Guía)."),
            ("Cobertura de salud", "Nuevo", "No estaba en la Matriz; se construye con afiliados (TerriData) y población DANE.",
             "Reemplaza al indicador de aseguramiento subsidiado; sin meta oficial."),
            ("ILE como evolución", "Nuevo", "No estaba en la Matriz; mide la serie y el avance de primera infancia 2019-2025.",
             "Reemplaza al ILE original en el sistema simplificado."),
            ("Logros de políticas públicas", "Nuevo",
             "No estaba en la Matriz; depende de un inventario de políticas vigentes que no existe todavía.",
             "Indicador nuevo, construible cuando se tenga el inventario."),
        ],
        PP: [
            ("PP.EPP Ejecución presupuestal del presupuesto participativo", "Con ajuste",
             "La lista de proyectos se reconstruyó cruzando el listado de proyectos votados con el plan de acción; falta confirmarla.",
             "Se mantiene como EPP; respaldos: tasa de compromiso, proyectos con ejecución ≥ 90 % y equidad entre zonas."),
            ("PP.CEP Cumplimiento de entrega de proyectos", "No",
             "No hay actas de recibo ni informes finales; solo el autorreporte de las dependencias.",
             "Reemplazado por Proyectos votados que se terminaron (autorreporte) y su ejecución presupuestal."),
            ("PP.CMA Cumplimiento de metas de atención", "No",
             "No hay meta anual de personas ni personas atendidas por proyecto.",
             "Reemplazado por Fidelidad a lo votado (el proyecto recibió lo que la comunidad votó)."),
            ("PP.IREPP Índice de Resultados de Ejecución del Presupuesto Participativo", "No",
             "Depende de CEP y CMA, que no se pueden calcular, y la Guía impide compensarlos con EPP.",
             "Reemplazado por tres indicadores sueltos: Fidelidad a lo votado, EPP y Proyectos votados que se terminaron."),
            ("Fidelidad a lo votado", "Nuevo", "No estaba en la Matriz; se calcula con el listado de proyectos votados y Hacienda.",
             "Indicador principal nuevo (semáforo de PP)."),
            ("Proyectos votados que se terminaron", "Nuevo", "No estaba en la Matriz; se calcula con el avance físico del plan de acción.",
             "Indicador principal nuevo (autorreporte, provisional)."),
        ],
    }


# ---------------------------------------------------------------- indicadores
def pct(n, d):
    return ('=IF(AND(ISNUMBER({%s}),ISNUMBER({%s})),IF({%s}>0,100*{%s}/{%s},"N/D"),"N/D")' % (n, d, d, n, d))


def ref(x):
    return '=IF(ISNUMBER({%s}),{%s},"N/D")' % (x, x)


def ind(key, nombre, tipo, fb, palabras, lectura, f, years=None, fmt="0.00", sem=None, semy=None, csv=None):
    return dict(key=key, nombre=nombre, tipo=tipo, fb=fb, palabras=palabras, lectura=lectura, f=f, years=years,
                fmt=fmt, sem=sem, semy=semy, csv=csv)


def logro_reduccion(o):
    return ('=IF(ISNUMBER({%s}),IF({%s}=0,100,MIN(100,100*{PDM.CM.M@M}/{%s})),"N/D")' % (o, o, o))


def logro_aumento(o, m):
    return '=IF(ISNUMBER({%s}),MIN(100,100*{%s}/{%s@M}),"N/D")' % (o, o, m)


I_CS = [
    ind("pcs", "PCS Promoción del control social", "Principal", "No se puede", "100 × estrategias ejecutadas / programadas",
        "Dato autodeclarado por la dependencia (con observaciones del DAP). 2026 sin cierre.", pct("CS.PCS.EE", "CS.PCS.EP"),
        [2024, 2025], sem="U_7550", csv="PCS - Promoción del control social"),
    ind("pcs_r", "PCS: estrategias ejecutadas", "Respaldo", "No se puede", "Conteo de estrategias ejecutadas",
        "Conteo sin meta: no tiene semáforo.", ref("CS.PCS.EE"), [2024, 2025], fmt="0",
        csv="PCS - Respaldo: estrategias ejecutadas (número)"),
    ind("dsi", "DSI Disponibilidad de información de seguimiento", "Principal", "No se puede",
        "100 × metas con valor publicado a tiempo / metas seleccionadas",
        "Depende de la lista provisional de 7 metas; con otra lista cambia el resultado.", pct("CS.DSI.MD", "CS.DSI.MS"),
        [2024, 2025], sem="U_9070", csv="DSI - Disponibilidad de información de seguimiento"),
    ind("dsi_r", "DSI: con todo lo publicado hasta la consulta", "Respaldo", "No se puede",
        "Igual, sin exigir que el valor se haya publicado a tiempo",
        "Cuenta lo publicado después del cierre de la vigencia.", pct("CS.DSI.MDT", "CS.DSI.MS"), [2024, 2025],
        sem="U_9070", csv="DSI - Respaldo: con todo lo publicado hasta la fecha de consulta"),
    ind("vrc", "VRC Veedurías en la rendición de cuentas", "Principal", "Construible",
        "100 × eventos con veeduría / eventos de rendición de cuentas",
        "Sin listas de asistencia no consta la presencia de veedurías: queda sin dato (N/D).", "N/D", [2024, 2025],
        sem="U_7550", csv="VRC - Veedurías en la rendición de cuentas"),
    ind("vrc_r", "VRC: evento sin mención de veeduría = ausencia", "Respaldo", "No se puede",
        "Igual, contando como ausencia el evento que no menciona veedurías",
        "El supuesto puede subestimar la presencia real de veedurías.", pct("CS.VRC.EV", "CS.VRC.ET"), [2024, 2025],
        sem="U_7550", csv="VRC - Respaldo: evento sin mención de veeduría = ausencia"),
    ind("rpo", "RPO Respuesta a observaciones ciudadanas", "Principal", "Construible",
        "100 × observaciones con respuesta de fondo publicada / observaciones registradas",
        "2025 sin cuadro completo de preguntas y respuestas: queda sin dato.",
        {2024: pct("CS.RPO.OR", "CS.RPO.OT"), 2025: "N/D"}, [2024, 2025], sem="U_9070",
        csv="RPO - Respuesta a observaciones ciudadanas"),
    ind("rpo_r", "RPO: con el cuadro publicado", "Respaldo", "No se puede", "Igual, con el único cuadro publicado de 2025",
        "Solo 2 observaciones: resultado frágil.", pct("CS.RPO.OR", "CS.RPO.OT"), [2025], sem="U_9070",
        csv="RPO - Respaldo: con el cuadro de preguntas y respuestas publicado"),
    ind("icrcs", "ICRCS Índice de Condiciones y Respuesta para el Control Social", "Principal", "Construible",
        "(DSI + PCS + VRC + RPO) / 4, solo si las cuatro entradas tienen valor",
        "No se calcula si falta algún componente (regla de la Guía).",
        '=IF(COUNT({i:dsi},{i:pcs},{i:vrc},{i:rpo})=4,({i:dsi}+{i:pcs}+{i:vrc}+{i:rpo})/4,"N/D")', [2024, 2025], sem="U_7550",
        csv="ICRCS - Índice de Condiciones y Respuesta para el Control Social"),
    ind("icrcs_r", "ICRCS: con los respaldos de VRC y RPO", "Respaldo", "No se puede",
        "Mismo promedio, con los respaldos de VRC y RPO", "Mezcla principales y respaldos; provisional.",
        '=IF(AND(ISNUMBER({i:dsi}),ISNUMBER({i:pcs}),ISNUMBER({i:vrc_r}),OR(ISNUMBER({i:rpo}),ISNUMBER({i:rpo_r}))),'
        '({i:dsi}+{i:pcs}+{i:vrc_r}+IF(ISNUMBER({i:rpo}),{i:rpo},{i:rpo_r}))/4,"N/D")', [2024, 2025], sem="U_7550",
        csv="ICRCS - Respaldo: con los respaldos de VRC y RPO"),
]
I_PC = [
    ind("i1", "I1 Incorporación de aportes", "Principal", "No se puede", "100 × procesos con aporte incorporado / procesos cerrados",
        "Sin inventario oficial de procesos cerrados: queda sin dato.", pct("PC.I1.V1", "PC.I1.V2"), sem="U_7550",
        csv="I1 - Incorporación de aportes"),
    ind("i1_r", "I1: proyectos votados con presupuesto asignado", "Respaldo", "No se puede",
        "100 × proyectos votados con presupuesto / proyectos votados",
        "Usa los datos de la fidelidad de PP: no sumar en dos dimensiones.", pct("PP.PRES", "PP.CEP.PE"), sem="U_7550",
        csv="I1 - Respaldo: proyectos votados con presupuesto asignado"),
    ind("i2", "I2 Cobertura territorial", "Principal", "No se puede", "100 × zonas con actividad participativa / 13 zonas oficiales",
        "Las actividades por zona salen de reportes agregados, sin actas por zona.", pct("PC.I2.V1", "PC.I2.V2"), sem="U_8050",
        csv="I2 - Cobertura territorial"),
    ind("i2_r", "I2: zonas con asamblea de presupuesto participativo", "Respaldo", "No se puede",
        "100 × zonas con asamblea de presupuesto participativo / 13", "La zona 1 no tiene comité zonal.",
        pct("PC.I2.ASAM", "PC.I2.V2"), sem="U_8050", csv="I2 - Respaldo: zonas con asamblea de presupuesto participativo"),
    ind("i3", "I3 Oferta formativa", "Principal", "No se puede", "100 × acciones formativas realizadas / programadas",
        "2024 sin programación publicada; la de 2025 es una reconstrucción.", pct("PC.I3.V1", "PC.I3.V2"), sem="U_7550",
        csv="I3 - Oferta formativa"),
    ind("i3_r", "I3: acciones formativas culminadas", "Respaldo", "No se puede", "Conteo de acciones culminadas",
        "Incluye acciones para servidores; sin meta, sin semáforo.", ref("PC.I3.CULM"), fmt="0",
        csv="I3 - Respaldo: acciones formativas culminadas (número)"),
    ind("i4", "I4 Publicación de resultados", "Principal", "No se puede",
        "100 × procesos con resultados publicados / procesos cerrados",
        "Sin inventario oficial de procesos cerrados: queda sin dato.", pct("PC.I4.V1", "PC.I4.V2"), sem="U_7550",
        csv="I4 - Publicación de resultados"),
    ind("i4_r", "I4: procesos del universo exploratorio", "Respaldo", "No se puede", "Conteo de procesos del universo exploratorio",
        "Universo no oficial; sin semáforo.", ref("PC.I4.V2"), fmt="0",
        csv="I4 - Respaldo: procesos con documento de resultados (universo exploratorio)"),
    ind("ispc", "ISPC Índice Simplificado de Participación Ciudadana", "Principal", "No se puede",
        "0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4, solo con las cuatro entradas",
        "Faltan I1 e I4: el índice original no se calcula.",
        '=IF(COUNT({i:i1},{i:i2},{i:i3},{i:i4})=4,0.3*{i:i1}+0.25*{i:i2}+0.2*{i:i3}+0.25*{i:i4},"N/D")', sem="U_7550",
        csv="ISPC - Índice (fórmula original: 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4)"),
    ind("ispc_r", "ISPC: opción redistribuida (0,556·I2 + 0,444·I3)", "Respaldo", "No se puede",
        "(0,25·I2 + 0,20·I3) / 0,45", "No es el índice original; 2024 sin I3. Umbral 75/50 como propuesta.",
        '=IF(AND(ISNUMBER({i:i2}),ISNUMBER({i:i3})),(0.25*{i:i2}+0.2*{i:i3})/0.45,"N/D")', sem="U_7550",
        csv="ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)"),
]
Y_ALL = list(range(2019, 2027))
Y19_25 = list(range(2019, 2026))
I_PDM = [
    ind("sal", "Cobertura de salud (reemplaza el aseguramiento subsidiado)", "Principal", "No se puede",
        "100 × afiliados (contributivo + subsidiado + especiales) / población DANE",
        "Sin meta oficial: no tiene semáforo. Una sola proyección DANE para toda la serie.",
        pct("PDM.SAL.TOTAL", "PDM.SAL.POB"), Y19_25, csv="Cobertura de salud"),
    ind("sal_d", "Cobertura de salud, cifra publicada por MinSalud (diciembre)", "Referencial", "No se puede",
        "100 × afiliados de diciembre / población usada por MinSalud", "Otra población: solo para comparar.",
        pct("PDM.SAL.TOTAL", "PDM.SAL.POBMS"), [2025], csv="Cobertura de salud, cifra publicada por MinSalud (diciembre)"),
    ind("sal_a", "Cobertura de salud, cifra publicada por MinSalud (agosto)", "Referencial", "No se puede",
        "100 × afiliados de agosto de 2026 / población usada por MinSalud", "Corte parcial del año: solo para comparar.",
        pct("PDM.SAL.TOT26", "PDM.SAL.POBMS"), [2026], csv="Cobertura de salud, cifra publicada por MinSalud (agosto)"),
    ind("disc", "Inclusión: personas con discapacidad en el registro (acumulado)", "Referencial", "No se puede",
        "Número de personas del registro nacional", "Registro sin actualizar desde 2020; no es una tasa.",
        ref("PDM.INC.DISCAP"), [2020], fmt="#,##0",
        csv="Inclusión: personas con discapacidad en el registro nacional (acumulado)"),
    ind("cm", "CM Mortalidad vial", "Principal", "No se puede", "Logro de reducción: 100 si el observado es 0; si no, MIN(100; 100 × meta / observado)",
        "Solo 2024; meta final 2027, sin metas anuales.", logro_reduccion("PDM.CM.O"), [2024, 2025], sem="U_PDM",
        csv="CM - Mortalidad vial"),
    ind("cm_r", "CM: cifra de la Policía", "Respaldo", "No se puede", "Igual, con la cifra de homicidios en accidente de tránsito de la Policía",
        "Otra definición que la de Medicina Legal.", logro_reduccion("PDM.CM.POL"), [2024, 2025], sem="U_PDM",
        csv="CM - Respaldo: cifra de la Policía"),
    ind("ca", "CA Espacio público verde", "Principal", "No se puede", "Logro de aumento: MIN(100; 100 × observado / meta)",
        "No concilia con la línea base del plan; solo 2024.", logro_aumento("PDM.CA.O", "PDM.CA.M"), [2024, 2025], sem="U_PDM",
        csv="CA - Espacio público verde"),
    ind("ca_r", "CA: espacio público verde total del POT", "Respaldo", "No se puede", "Igual, con el total del POT",
        "Otra serie: no comparable con la línea base.", logro_aumento("PDM.CA.POT", "PDM.CA.M"), [2024, 2025], sem="U_PDM",
        csv="CA - Respaldo: espacio público verde total del POT"),
    ind("cp", "CP Procesos participativos", "Principal", "No se puede", "Logro de aumento: MIN(100; 100 × observado / meta 100 %)",
        "No hay método ni cifra oficial del porcentaje de implementación.", logro_aumento("PDM.CP.O", "PDM.CP.M"),
        [2024, 2025], sem="U_PDM", csv="CP - Procesos participativos"),
    ind("cp_r", "CP: índice situacional", "Respaldo", "No se puede", "Índice del diagnóstico del plan (año 2022)",
        "Sin meta comparable; sin semáforo.", ref("PDM.CP.IDX"), [2022], csv="CP - Respaldo: índice situacional de implementación"),
    ind("cd", "CD Servicios deportivos", "Principal", "No se puede", "Logro de aumento: MIN(100; 100 × cobertura observada / meta 6 %)",
        "Falta la cifra de cobertura del Índer.", logro_aumento("PDM.CD.O", "PDM.CD.M"), [2025], sem="U_PDM",
        csv="CD - Servicios deportivos"),
    ind("cd_r", "CD: disciplinas atendidas", "Respaldo", "No se puede", "100 × disciplinas implementadas / previstas",
        "Mide oferta, no población atendida.", pct("PDM.CD.IMP", "PDM.CD.PREV"), [2025], sem="U_PDM",
        csv="CD - Respaldo: disciplinas deportivas atendidas"),
    ind("ile_c", "ILE primera infancia: cobertura de atención integral (0 a 5 años)", "Principal", "No se puede",
        "100 × atendidos de forma integral / población de 0 a 5 años", "Con 0 a 5 años, 2023 da 16,56 y no 16,3 de la línea base.",
        pct("PDM.ILE.ATEND", "PDM.ILE.POB05"), Y19_25, csv="ILE primera infancia - cobertura de atención integral"),
    ind("ile_v", "ILE primera infancia: variación anual (puntos porcentuales)", "Complemento", "No se puede",
        "Cobertura del año − cobertura del año anterior", "Puntos porcentuales, no porcentaje.",
        '=IF(AND(ISNUMBER({i:ile_c}),ISNUMBER({i:ile_c@prev})),{i:ile_c}-{i:ile_c@prev},"N/D")', list(range(2020, 2026)),
        csv="ILE primera infancia - variación anual (puntos porcentuales)"),
    ind("ile_a", "ILE primera infancia: avance hacia la meta 2027", "Complemento", "No se puede",
        "MAX(0; MIN(100; 100 × (valor − línea base 16,3) / (meta 19 − 16,3)))",
        "Sin semáforo. Por debajo de la línea base el avance es 0.",
        '=IF(ISNUMBER({i:ile_c}),MAX(0,MIN(100,100*({i:ile_c}-{PDM.ILE.LB@M})/({PDM.ILE.MA@M}-{PDM.ILE.LB@M}))),"N/D")', Y19_25,
        csv="ILE primera infancia - avance hacia la meta 2027"),
    ind("ile_s", "ILE: planteles que mejoran su categoría en Saber 11", "Principal", "No se puede",
        "Logro de aumento: MIN(100; 100 × observado / meta 14 %)", "Falta la clasificación de planteles por institución.",
        logro_aumento("PDM.ILE.S", "PDM.ILE.MS"), [2025], csv="ILE Saber 11 - clasificación de planteles (meta 14 %)"),
    ind("sab_o", "ILE: puntaje Saber 11, oficiales urbanas", "Respaldo", "No se puede", "Puntaje global promedio",
        "Datos públicos solo hasta 2022; sin semáforo.", ref("PDM.ILE.SABER_OU"), Y19_25,
        csv="ILE Saber 11 - puntaje global promedio, instituciones oficiales urbanas"),
    ind("sab_t", "ILE: puntaje Saber 11, todos los planteles", "Respaldo", "No se puede", "Puntaje global promedio",
        "Datos públicos solo hasta 2022; sin semáforo.", ref("PDM.ILE.SABER_TD"), Y19_25,
        csv="ILE Saber 11 - puntaje global promedio, todos los planteles del municipio"),
    ind("sab_ov", "ILE: variación anual del puntaje (oficiales urbanas)", "Respaldo", "No se puede", "Puntaje del año − puntaje del año anterior",
        "Solo donde hay dos años consecutivos.",
        '=IF(AND(ISNUMBER({i:sab_o}),ISNUMBER({i:sab_o@prev})),{i:sab_o}-{i:sab_o@prev},"N/D")', list(range(2020, 2026)),
        csv="ILE Saber 11 - variación anual del puntaje (oficiales urbanas)"),
    ind("sab_tv", "ILE: variación anual del puntaje (todos los planteles)", "Respaldo", "No se puede", "Puntaje del año − puntaje del año anterior",
        "Solo donde hay dos años consecutivos.",
        '=IF(AND(ISNUMBER({i:sab_t}),ISNUMBER({i:sab_t@prev})),{i:sab_t}-{i:sab_t@prev},"N/D")', list(range(2020, 2026)),
        csv="ILE Saber 11 - variación anual del puntaje (todos los planteles)"),
    ind("icpd", "ICPD-6 Índice de cumplimiento de resultados seleccionados del PDM", "Principal", "No se puede",
        "(IBS + ITG) / 2 con IBS = (salud + ILE + CD) / 3 e ITG = (CM + CA + CP) / 3",
        "Faltan salud (sin meta), CP, CD y planteles: el índice completo no se calcula.", "N/D", [2024, 2025]),
    ind("icpd_n", "ICPD-6 referencial: entradas disponibles (de 6)", "Referencial", "No se puede", "Cuántas de las 6 entradas tienen valor",
        "Conteo de apoyo.", '=COUNT({i:cm},{i:ca},{i:cd_r},{i:ile_a})', [2024, 2025], fmt="0"),
    ind("icpd_p", "ICPD-6 referencial: promedio con las entradas disponibles", "Referencial", "No se puede",
        "Promedio simple de las entradas con valor (CM, CA, CD de respaldo y avance del ILE)",
        "Referencial: no cumple la regla de la Guía de no calcular índices con entradas faltantes.",
        '=IF(COUNT({i:cm},{i:ca},{i:cd_r},{i:ile_a})>0,AVERAGE({i:cm},{i:ca},{i:cd_r},{i:ile_a}),"N/D")', [2024, 2025],
        csv="ICPD-6 - Promedio parcial con los indicadores disponibles"),
    ind("pol", "Logros de políticas públicas (nueva dimensión)", "Principal", "Construible",
        "100 × políticas vigentes con informe de seguimiento publicado / políticas vigentes",
        "Sin inventario oficial de políticas vigentes: aún sin valor.", pct("PDM.POL.INF", "PDM.POL.VIG"), [2025],
        csv="Logros de políticas públicas (nueva dimensión)"),
    ind("pol_r", "Políticas: con plan de acción vigente", "Respaldo", "No se puede", "100 × políticas con plan de acción / políticas vigentes",
        "Para más adelante; sin fuente identificada.", pct("PDM.POL.PLAN", "PDM.POL.VIG"), [2025]),
]
I_PP = [
    ind("fid", "Fidelidad a lo votado", "Principal", "No se puede",
        "100 × proyectos con apropiación igual al valor votado (± 1 %) / proyectos votados",
        "2024 aproximada (denominador 51; con los 50 cruzados daría 78,00). 2026 sin el valor votado por proyecto.",
        pct("PP.FID.N", "PP.CEP.PE"), sem="U_PP", semy=[2024, 2025], csv="Fidelidad a lo votado"),
    ind("fid_m", "Fidelidad a lo votado, versión monetaria", "Complemento", "No se puede",
        "100 × apropiación total / valor votado total (proyectos considerados)", "Complemento de la fidelidad; puede superar 100.",
        pct("PP.EPP.APP", "PP.VOT.C"), [2024, 2025], csv="Fidelidad a lo votado, versión monetaria"),
    ind("ter", "Proyectos votados que se terminaron", "Principal", "Construible",
        "100 × proyectos votados con avance físico de 100 % / proyectos votados",
        "Autorreporte de las dependencias, sin acta de recibo. 2024 y 2026 sin dato.", pct("PP.CEP.PEC", "PP.CEP.PE"),
        sem="U_PP", csv="Proyectos votados que se terminaron"),
    ind("epp", "EPP Ejecución presupuestal del presupuesto participativo", "Principal", "No se puede",
        "100 × obligaciones (OPP) / apropiación definitiva (APP)",
        "2024 aproximada (50 de 51 proyectos cruzados); 2026 es solo el primer semestre y no lleva semáforo.",
        pct("PP.EPP.OPP", "PP.EPP.APP"), sem="U_PP", semy=[2024, 2025],
        csv={2024: "EPP, ejecución presupuestal del presupuesto participativo",
             2025: "EPP, ejecución presupuestal del presupuesto participativo",
             2026: "EPP, ejecución presupuestal del presupuesto participativo (1.er semestre)"}),
    ind("ter_e", "Ejecución presupuestal de los proyectos terminados", "Respaldo", "Construible",
        "100 × obligaciones / apropiación, solo de los proyectos con avance físico de 100 %",
        "Autorreporte de las dependencias, sin acta de recibo.", pct("PP.TER.OPP", "PP.TER.APP"), sem="U_PP",
        csv="Ejecución presupuestal de los proyectos terminados"),
    ind("comp", "Tasa de compromiso", "Respaldo", "Construible", "100 × compromisos / apropiación definitiva",
        "2024: Hacienda no separa compromisos.", pct("PP.CMP", "PP.EPP.APP"), [2024, 2025], sem="U_PP", csv="Tasa de compromiso"),
    ind("p90", "Proyectos con ejecución de al menos 90 %", "Respaldo", "Construible",
        "100 × proyectos con obligaciones ≥ 90 % de su apropiación / proyectos", "2024 sobre 50 proyectos cruzados.",
        pct("PP.PROY90", "PP.PROY.EV"), [2024, 2025], sem="U_PP", csv="Proyectos con ejecución de al menos 90 %"),
]
for z in range(1, 14):
    I_PP.append(ind("z%d" % z, "EPP de la zona %d" % z, "Complemento", "Construible",
                    "100 × obligaciones de la zona / apropiación de la zona",
                    "Zona 1: sin proyectos de presupuesto participativo." if z == 1 else "Insumo de la equidad entre zonas.",
                    pct("PP.Z%d.O" % z, "PP.Z%d.A" % z)))
I_PP.append(ind("eq", "Equidad entre zonas (brecha de ejecución)", "Respaldo", "Construible",
                "EPP de la zona más alta − EPP de la zona más baja (puntos)", "Sin umbral: no tiene semáforo. 12 zonas con proyectos.",
                "EQ", [2024, 2025], csv="Equidad entre zonas (brecha de ejecución)"))

HOJAS = OrderedDict([
    (CS, dict(ys=[2024, 2025, 2026], meta=False, wy=13, ind=I_CS, titulo="Control social",
              desc="Mide si la ciudadanía tiene información y respuestas para vigilar la gestión pública.",
              siglas="Siglas: PCS = Promoción del control social; DSI = Disponibilidad de información de seguimiento; "
                     "VRC = Veedurías en la rendición de cuentas; RPO = Respuesta a observaciones ciudadanas; "
                     "ICRCS = Índice de Condiciones y Respuesta para el Control Social; PDM = Plan de Desarrollo Municipal; "
                     "DAP = Departamento Administrativo de Planeación.")),
    (PC, dict(ys=[2024, 2025], meta=False, wy=17, ind=I_PC, titulo="Participación ciudadana",
              desc="Mide si la participación llega a todo el territorio, si la oferta formativa se cumple y si los aportes "
                   "ciudadanos se incorporan y se publican. Los profesores eliminaron I5 e I6.",
              siglas="Siglas: I1 = Incorporación de aportes; I2 = Cobertura territorial; I3 = Oferta formativa; "
                     "I4 = Publicación de resultados; ISPC = Índice Simplificado de Participación Ciudadana.")),
    (PDM, dict(ys=Y_ALL, meta=True, wy=9.5, ind=I_PDM, titulo="Plan de Desarrollo Municipal (PDM)",
               desc="Mide el avance de resultados seleccionados del Plan de Desarrollo 2024-2027.",
               siglas="Siglas: PDM = Plan de Desarrollo Municipal; CM = Mortalidad vial; CA = Espacio público verde; "
                      "CP = Procesos participativos; CD = Servicios deportivos; ILE = Índice de Logro Educativo; "
                      "ICPD-6 = Índice de cumplimiento de resultados seleccionados del PDM; POT = Plan de Ordenamiento Territorial; "
                      "DANE = Departamento Administrativo Nacional de Estadística; MinSalud = Ministerio de Salud; LB = línea base.")),
    (PP, dict(ys=[2024, 2025, 2026], meta=False, wy=14, ind=I_PP, titulo="Presupuesto participativo (PP)",
              desc="Mide si los proyectos votados por la comunidad recibieron lo votado, se ejecutaron y se terminaron.",
              siglas="Siglas: PP = presupuesto participativo; EPP = ejecución presupuestal del PP; CEP = Cumplimiento de entrega de "
                     "proyectos; CMA = Cumplimiento de metas de atención; IREPP = Índice de Resultados de Ejecución del PP; "
                     "APP = apropiación definitiva; OPP = obligaciones.")),
])

# ---------------------------------------------------------------- utilidades de estilo
def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def fnt(**k):
    k.setdefault("name", FN)
    k.setdefault("size", 10)
    return Font(**k)


def lineas(txt, ancho):
    if not txt:
        return 1
    return sum(max(1, math.ceil(len(p) / max(ancho * 1.05, 1))) for p in str(txt).split("\n"))


def alto(pares):
    """pares = [(texto, ancho_total)] -> altura de fila."""
    return max(16, max(lineas(t, w) for t, w in pares) * 13.2 + 3)


REG_V, REG_I, YC, METAC = {}, {}, {}, {}  # registros: id->(hoja,fila), (hoja,key)->fila, hoja->{año:col}
WB = openpyxl.Workbook()
WB.remove(WB.active)
CELDAS = []  # para el mapa de verificación
NFORM = 0


def q(h):
    return "'%s'!" % h


def resolver(tpl, hoja, anio):
    def sub(m):
        esi, name, mod = m.group(1), m.group(2), m.group(3)
        if esi:
            h, fila = hoja, REG_I[(hoja, name)]
        else:
            h, fila = REG_V[name]
        y = anio
        if mod == "prev":
            y = anio - 1
        elif mod and mod.isdigit():
            y = int(mod)
        col = METAC[h] if mod == "M" else YC[h][y]
        return (q(h) if h != hoja else "") + "%s%d" % (L(col), fila)
    return re.sub(r"\{(i:)?([A-Za-z0-9_.]+)(?:@(\w+))?\}", sub, tpl)


def banda(ws, fila, texto, ncols, color=AZUL, tam=14, fcolor="FFFFFF", negrita=True, alto_f=None):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, texto)
    c.font = fnt(size=tam, bold=negrita, color=fcolor)
    c.fill = fill(color)
    c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    for k in range(1, ncols + 1):
        ws.cell(fila, k).fill = fill(color)
    if alto_f:
        ws.row_dimensions[fila].height = alto_f


def nota(ws, fila, texto, ncols, ancho_total, negrita=False, color="404040", italic=False):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, texto)
    c.font = fnt(size=9.5, bold=negrita, color=color, italic=italic)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    ws.row_dimensions[fila].height = alto([(texto, ancho_total)])


def enc(ws, fila, textos, desde=1):
    for i, t in enumerate(textos):
        c = ws.cell(fila, desde + i, t)
        c.font = fnt(bold=True, color="1F3864")
        c.fill = fill(ENC)
        c.border = BORDE
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center" if i else "left")
    ws.row_dimensions[fila].height = 30


def aplicar_cf(ws, fila, cols, umb):
    """Formato condicional por umbrales sobre las celdas de resultado de una fila."""
    if not umb:
        return
    a, b = L(cols[0]), L(cols[-1])
    rng = "%s%d:%s%d" % (a, fila, b, fila)
    x = "%s%d" % (a, fila)
    prev = None
    for t, col in umb:
        cond = ["ISNUMBER(%s)" % x]
        if t is not None:
            cond.append("%s>=%s" % (x, t))
        if prev is not None:
            cond.append("%s<%s" % (x, prev))
        ws.conditional_formatting.add(rng, FormulaRule(formula=["AND(%s)" % ",".join(cond)], fill=fill(F_SEM[col]),
                                                       font=Font(color={"V": "006100", "A": "7F6000", "N": "833C0B", "R": "9C0006"}[col])))
        prev = t


def formato_num(v):
    return v


def construir_hoja(nombre, cfg, a_rows):
    global NFORM
    ws = WB.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    ys, meta = cfg["ys"], cfg["meta"]
    ny = len(ys)
    YC[nombre] = {y: 4 + i for i, y in enumerate(ys)}
    metac = 4 + ny if meta else None
    METAC[nombre] = metac
    f1 = 4 + ny + (1 if meta else 0)
    f2 = f1 + 1
    ncols = f2
    wy = cfg["wy"]
    w = {1: 44, 2: 14, 3: 15}
    for y in ys:
        w[YC[nombre][y]] = wy
    if meta:
        w[metac] = 11
    w[f1] = 60
    w[f2] = 42
    for c, v in w.items():
        ws.column_dimensions[L(c)].width = v
    total_w = sum(w.values())
    banda(ws, 1, "Observatorio Ciudadano de Envigado · " + cfg["titulo"], ncols, tam=15, alto_f=30)
    nota(ws, 2, "PROVISIONAL: ningún resultado está validado. Las celdas de la sección B son los datos de entrada; las fórmulas "
                "de la sección C se recalculan solas si se cambia un dato. " + cfg["desc"], ncols, total_w, negrita=True, color="C00000")
    nota(ws, 3, cfg["siglas"], ncols, total_w, italic=True)
    fila = 5
    # ---------- A
    banda(ws, fila, "A. Matriz original: qué pasó con cada indicador", ncols, color=ENC, tam=11, fcolor="1F3864", alto_f=22)
    fila += 1
    enc(ws, fila, ["Indicador original (código y nombre de la Matriz)", "¿Se puede calcular?"])
    ultimo_y = metac or (3 + ny)
    ws.merge_cells(start_row=fila, start_column=3, end_row=fila, end_column=ultimo_y)
    enc(ws, fila, ["Por qué"], desde=3)
    ws.merge_cells(start_row=fila, start_column=f1, end_row=fila, end_column=f2)
    enc(ws, fila, ["Qué lo reemplaza o respalda en el sistema simplificado"], desde=f1)
    for k in range(3, ncols + 1):
        ws.cell(fila, k).border = BORDE
        ws.cell(fila, k).fill = fill(ENC)
    fila += 1
    a0 = fila
    wporque = sum(w[k] for k in range(3, ultimo_y + 1))
    wrepl = w[f1] + w[f2]
    for orig, se, porque, repl in a_rows:
        ws.cell(fila, 1, orig)
        ws.cell(fila, 2, se)
        ws.merge_cells(start_row=fila, start_column=3, end_row=fila, end_column=ultimo_y)
        ws.cell(fila, 3, porque)
        ws.merge_cells(start_row=fila, start_column=f1, end_row=fila, end_column=f2)
        ws.cell(fila, f1, repl)
        for k in range(1, ncols + 1):
            c = ws.cell(fila, k)
            c.font = fnt(bold=(k <= 2))
            c.border = BORDE
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center" if k == 2 else "left")
        ws.cell(fila, 2).fill = fill(POR_QUE_FILL[se])
        ws.row_dimensions[fila].height = alto([(orig, w[1]), (porque, wporque), (repl, wrepl)])
        fila += 1
    a1 = fila - 1
    fila += 1
    # ---------- B
    banda(ws, fila, "B. Variables (datos de entrada)", ncols, color=ENC, tam=11, fcolor="1F3864", alto_f=22)
    fila += 1
    nota(ws, fila, "Cada fila es un dato de entrada. Estado: Disponible = publicado y verificable; Construible = se arma con datos "
                   "públicos pero requiere una decisión o una lectura propia; No disponible = depende de la Alcaldía. "
                   "N/D = sin dato para ese año. Los indicadores de la sección C leen estas celdas.", ncols, total_w, italic=True)
    fila += 1
    enc(ws, fila, ["Variable", "Unidad", "Estado"] + [str(y) for y in ys] + (["Meta 2027 / línea base"] if meta else [])
        + ["Fuente (con enlace)", "Qué falta o limitación"])
    fila += 1
    vars_h = [v for v in VARS.values() if v["hoja"] == nombre]
    b0 = fila
    for v in vars_h:
        REG_V[v["id"]] = (nombre, fila)
        c = ws.cell(fila, 1, v["nombre"])
        ws.cell(fila, 2, v["unidad"])
        ws.cell(fila, 3, v["estado"]).fill = fill(EST_VAR_FILL[v["estado"]])
        entero = v["unidad"] in ("pesos", "personas", "proyectos", "zonas", "metas", "eventos", "observaciones", "estrategias",
                                 "acciones", "procesos", "políticas", "disciplinas", "documentos", "instituciones")
        nf = "#,##0" if entero else "#,##0.00"
        for y in ys:
            val = v["vals"].get(y)
            cc = ws.cell(fila, YC[nombre][y])
            if val is None:
                cc.value, cc.font, cc.alignment = "N/D", fnt(color=GRIS_TXT), Alignment(horizontal="center", vertical="center")
            else:
                cc.value, cc.number_format = val, nf
                cc.font, cc.alignment = fnt(color="0000FF"), Alignment(horizontal="center", vertical="center")
        if meta:
            cc = ws.cell(fila, metac)
            mv = v["vals"].get(2027, v["vals"].get("M"))
            if mv is None:
                cc.value, cc.font, cc.alignment = "—", fnt(color=GRIS_TXT), Alignment(horizontal="center", vertical="center")
            else:
                cc.value, cc.number_format = mv, nf
                cc.font, cc.alignment = fnt(color="0000FF", bold=True), Alignment(horizontal="center", vertical="center")
        # fuente con hipervínculo
        cf = ws.cell(fila, f1, v["fuente"] or "—")
        if v["enlace"]:
            cf.hyperlink = v["enlace"]
            cf.font = fnt(color="0563C1", underline="single")
        ws.cell(fila, f2, v["falta"] or "")
        for k in range(1, ncols + 1):
            c = ws.cell(fila, k)
            c.border = BORDE
            if k in (1, 2, 3, f1, f2):
                c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center" if k in (2, 3) else "left")
                if k in (1, 2, 3, f2):
                    c.font = fnt()
        ws.row_dimensions[fila].height = alto([(v["nombre"], w[1]), (v["fuente"], w[f1]), (v["falta"], w[f2])])
        fila += 1
    b1 = fila - 1
    fila += 1
    # ---------- C
    banda(ws, fila, "C. Indicadores calculados (fórmulas vivas sobre las variables de la sección B)", ncols, color=ENC, tam=11,
          fcolor="1F3864", alto_f=22)
    fila += 1
    nota(ws, fila, "Tipo: Principal = indicador del sistema; Respaldo = alternativa o apoyo del principal; Complemento = lectura "
                   "adicional del mismo indicador; Referencial = solo para comparar. Estado: Construido = ya tiene valor; "
                   "Construible = se puede armar con datos públicos y aún no tiene valor; No se puede = depende de datos de la Alcaldía. "
                   "Colores de semáforo según los umbrales de la Matriz de cada indicador (sin color = sin semáforo). Los umbrales "
                   "se indican en la columna de lectura.", ncols, total_w, italic=True)
    fila += 1
    enc(ws, fila, ["Indicador", "Tipo", "Estado"] + [str(y) for y in ys] + (["Meta 2027"] if meta else [])
        + ["Fórmula en palabras", "Lectura y limitación"])
    fila += 1
    c0 = fila
    zona_filas = []
    for s in cfg["ind"]:
        REG_I[(nombre, s["key"])] = fila
        if s["key"].startswith("z") and s["key"][1:].isdigit():
            zona_filas.append(fila)
        ws.cell(fila, 1, s["nombre"]).font = fnt(bold=(s["tipo"] == "Principal"))
        ws.cell(fila, 2, s["tipo"])
        yrs = s["years"] or ys
        ycol_first, ycol_last = L(YC[nombre][ys[0]]), L(YC[nombre][ys[-1]])
        ce = ws.cell(fila, 3, '=IF(COUNT(%s%d:%s%d)>0,"Construido","%s")' % (ycol_first, fila, ycol_last, fila, s["fb"]))
        NFORM += 1
        for y in ys:
            cc = ws.cell(fila, YC[nombre][y])
            cc.alignment = Alignment(horizontal="center", vertical="center")
            if y not in yrs:
                cc.value, cc.font, cc.alignment = "—", fnt(color="BFBFBF"), Alignment(horizontal="center", vertical="center")
                continue
            f = s["f"]
            if isinstance(f, dict):
                f = f[y]
            if f == "EQ":
                rng = "%s%d:%s%d" % (L(YC[nombre][y]), zona_filas[0], L(YC[nombre][y]), zona_filas[-1])
                f = '=IF(COUNT(%s)>=2,MAX(%s)-MIN(%s),"N/D")' % (rng, rng, rng)
            if f.startswith("="):
                cc.value = resolver(f, nombre, y)
                NFORM += 1
                cc.number_format = s["fmt"]
                cc.font = fnt(bold=(s["tipo"] == "Principal"))
            else:
                cc.value, cc.font, cc.alignment = f, fnt(color=GRIS_TXT), Alignment(horizontal="center", vertical="center")
            if s["csv"]:
                nm = s["csv"][y] if isinstance(s["csv"], dict) else s["csv"]
                CELDAS.append(dict(hoja=nombre, celda="%s%d" % (L(YC[nombre][y]), fila), indicador=nm, anio=y))
        if meta:
            cm = ws.cell(fila, metac)
            mkey = {"cm": "PDM.CM.M", "cm_r": "PDM.CM.M", "ca": "PDM.CA.M", "ca_r": "PDM.CA.M", "cp": "PDM.CP.M",
                    "cd": "PDM.CD.M", "ile_a": "PDM.ILE.MA", "ile_s": "PDM.ILE.MS"}.get(s["key"])
            if mkey:
                cm.value = "=%s%d" % (L(metac), REG_V[mkey][1])
                NFORM += 1
                cm.number_format = "0.00"
                cm.font = fnt(color="404040")
            cm.alignment = Alignment(horizontal="center", vertical="center")
        palabras = s["palabras"]
        if s["sem"]:
            lect = s["lectura"] + " Umbrales: " + TXT_UMBRAL[s["sem"]] + "."
        else:
            lect = s["lectura"]
        ws.cell(fila, f1, palabras)
        ws.cell(fila, f2, lect)
        for k in range(1, ncols + 1):
            c = ws.cell(fila, k)
            c.border = BORDE
            if k in (1, 2, 3, f1, f2):
                c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center" if k in (2, 3) else "left")
                if k != 1:
                    c.font = fnt()
        if s["sem"]:
            sy = s["semy"] or ys
            cols = [YC[nombre][y] for y in sy if y in yrs]
            if cols:
                aplicar_cf(ws, fila, [min(cols), max(cols)], UMB[s["sem"]])
        ws.row_dimensions[fila].height = alto([(s["nombre"], w[1]), (palabras, w[f1]), (lect, w[f2])])
        fila += 1
    c1 = fila - 1
    # formato condicional de tipo y estado
    ws.conditional_formatting.add("C%d:C%d" % (c0, c1), CellIsRule(operator="equal", formula=['"Construido"'], fill=fill("C6EFCE"), font=Font(color="006100")))
    ws.conditional_formatting.add("C%d:C%d" % (c0, c1), CellIsRule(operator="equal", formula=['"Construible"'], fill=fill("FFEB9C"), font=Font(color="7F6000")))
    ws.conditional_formatting.add("C%d:C%d" % (c0, c1), CellIsRule(operator="equal", formula=['"No se puede"'], fill=fill("FFC7CE"), font=Font(color="9C0006")))
    ws.freeze_panes = "B4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "Provisional · Observatorio Ciudadano de Envigado · &P de &N"
    return dict(a=(a0, a1), b=(b0, b1), c=(c0, c1), ncols=ncols, total_w=total_w)


UBICA = {}
A_ROWS = bloque_a()
for nombre in [PP, CS, PC, PDM]:
    UBICA[nombre] = construir_hoja(nombre, HOJAS[nombre], A_ROWS[nombre])

# ---------------------------------------------------------------- Series en el tiempo
SE = "Series en el tiempo"
ws = WB.create_sheet(SE)
ws.sheet_view.showGridLines = False
ys = Y19_25
ws.column_dimensions["A"].width = 58
for i in range(len(ys)):
    ws.column_dimensions[L(2 + i)].width = 11
ws.column_dimensions[L(2 + len(ys))].width = 60
nc = 2 + len(ys)
banda(ws, 1, "Observatorio Ciudadano de Envigado · Series en el tiempo", nc, tam=15, alto_f=30)
nota(ws, 2, "PROVISIONAL: ningún resultado está validado. Las series se calculan con fórmulas sobre las variables de la hoja "
            "'Plan de Desarrollo'. N/D = sin dato para ese año. Salud: una sola proyección de población DANE para todos los años.",
     nc, 58 + 11 * len(ys) + 60, negrita=True, color="C00000")
nota(ws, 3, "Siglas: ILE = Índice de Logro Educativo; DANE = Departamento Administrativo Nacional de Estadística; Saber 11 = prueba "
            "de Estado del ICFES para estudiantes de grado 11; pp = puntos porcentuales.", nc, 130, italic=True)
SER_COL = {y: 2 + i for i, y in enumerate(ys)}


def vref(id_, y):
    h, r = REG_V[id_]
    return q(h) + "%s%d" % (L(YC[h][y]), r)


def bloque_serie(fila, titulo, filas):
    banda(ws, fila, titulo, nc, color=ENC, tam=11, fcolor="1F3864", alto_f=22)
    fila += 1
    enc(ws, fila, ["Serie"] + [str(y) for y in ys] + ["Cómo se calcula"])
    fila += 1
    pos = {}
    for key, nombre, tpl, nf, txt in filas:
        pos[key] = fila
        ws.cell(fila, 1, nombre).font = fnt(bold=key in ("cov", "cob", "avn"))
        for y in ys:
            cc = ws.cell(fila, SER_COL[y])
            if tpl.startswith("ref:"):
                x = vref(tpl[4:], y)
                f = '=IF(ISNUMBER(%s),%s,"N/D")' % (x, x)
            elif tpl.startswith("pct:"):
                n, d = tpl[4:].split(",")
                f = '=IF(AND(ISNUMBER(%s),ISNUMBER(%s)),IF(%s>0,100*%s/%s,"N/D"),"N/D")' % (vref(n, y), vref(d, y), vref(d, y), vref(n, y), vref(d, y))
            elif tpl.startswith("var:"):
                k = pos[tpl[4:]]
                if y == ys[0]:
                    cc.value, cc.font, cc.alignment = "—", fnt(color="BFBFBF"), Alignment(horizontal="center")
                    cc.border = BORDE
                    continue
                a, b = L(SER_COL[y]), L(SER_COL[y] - 1)
                f = '=IF(AND(ISNUMBER(%s%d),ISNUMBER(%s%d)),%s%d-%s%d,"N/D")' % (a, k, b, k, a, k, b, k)
            elif tpl.startswith("avn:"):
                k = pos[tpl[4:]]
                a = "%s%d" % (L(SER_COL[y]), k)
                lb, ma = q(PDM) + "%s%d" % (L(METAC[PDM]), REG_V["PDM.ILE.LB"][1]), q(PDM) + "%s%d" % (L(METAC[PDM]), REG_V["PDM.ILE.MA"][1])
                f = '=IF(ISNUMBER(%s),MAX(0,MIN(100,100*(%s-%s)/(%s-%s))),"N/D")' % (a, a, lb, ma, lb)
            cc.value = f
            cc.number_format = nf
            cc.alignment = Alignment(horizontal="center", vertical="center")
            cc.font = fnt(bold=key in ("cov", "cob", "avn"))
            cc.border = BORDE
            CELDAS_SERIE.append((key, y, "%s%d" % (L(SER_COL[y]), fila)))
        ws.cell(fila, 1).border = BORDE
        c = ws.cell(fila, nc, txt)
        c.font, c.border = fnt(), BORDE
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(fila, 1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[fila].height = alto([(nombre, 58), (txt, 60)])
        fila += 1
    return fila + 1, pos


CELDAS_SERIE = []
r = 5
r, p1 = bloque_serie(r, "Cobertura de salud (reemplaza el aseguramiento subsidiado)", [
    ("afi", "Afiliados al sistema de salud (personas)", "ref:PDM.SAL.TOTAL", "#,##0", "Dato de entrada (hoja Plan de Desarrollo)."),
    ("pob", "Población total del municipio, proyección DANE (personas)", "ref:PDM.SAL.POB", "#,##0", "Dato de entrada (hoja Plan de Desarrollo)."),
    ("cov", "Cobertura de salud (%)", "pct:PDM.SAL.TOTAL,PDM.SAL.POB", "0.00", "100 × afiliados / población."),
    ("covv", "Variación anual (pp)", "var:cov", "0.00", "Cobertura del año − cobertura del año anterior."),
])
r, p2 = bloque_serie(r, "Primera infancia: cobertura de atención integral (ILE como evolución)", [
    ("ate", "Niños, niñas y gestantes atendidos de forma integral (personas)", "ref:PDM.ILE.ATEND", "#,##0", "Dato de entrada."),
    ("p05", "Población de 0 a 5 años, DANE (personas)", "ref:PDM.ILE.POB05", "#,##0", "Dato de entrada."),
    ("cob", "Cobertura de atención integral (%)", "pct:PDM.ILE.ATEND,PDM.ILE.POB05", "0.00", "100 × atendidos / población de 0 a 5 años."),
    ("cobv", "Variación anual (pp)", "var:cob", "0.00", "Cobertura del año − cobertura del año anterior."),
    ("avn", "Avance hacia la meta 2027 (%)", "avn:cob", "0.00", "MAX(0; MIN(100; 100 × (cobertura − 16,3) / (19 − 16,3))); por debajo de la línea base el avance es 0."),
])
r, p3 = bloque_serie(r, "Saber 11: puntaje global promedio", [
    ("so", "Instituciones oficiales urbanas (puntos)", "ref:PDM.ILE.SABER_OU", "0.00", "Dato de entrada; datos públicos solo hasta 2022."),
    ("sov", "Variación anual, oficiales urbanas (puntos)", "var:so", "0.00", "Solo donde hay dos años consecutivos."),
    ("st", "Todos los planteles del municipio (puntos)", "ref:PDM.ILE.SABER_TD", "0.00", "Dato de entrada; datos públicos solo hasta 2022."),
    ("stv", "Variación anual, todos los planteles (puntos)", "var:st", "0.00", "Solo donde hay dos años consecutivos."),
])
NFORM += len(CELDAS_SERIE)
ws.freeze_panes = "B4"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws.oddFooter.center.text = "Provisional · Observatorio Ciudadano de Envigado · &P de &N"

# ---------------------------------------------------------------- Resumen
RS = WB.create_sheet("Resumen", 0)
RS.sheet_view.showGridLines = False
anchos = [34, 44, 11, 14, 14, 12, 14]
for i, a in enumerate(anchos):
    RS.column_dimensions[L(i + 1)].width = a
tw = sum(anchos)
banda(RS, 1, "Observatorio Ciudadano de Envigado · Sistema de indicadores simplificado", 7, tam=16, alto_f=34)
nota(RS, 2, "PROVISIONAL: ningún resultado está validado. Es una propuesta de trabajo para discutir con los profesores.", 7, tw,
     negrita=True, color="C00000")
nota(RS, 3, "Cómo leer este libro. 1) Cada hoja de dimensión tiene tres bloques: A (qué pasó con cada indicador de la Matriz "
            "original), B (datos de entrada con su fuente) y C (indicadores calculados con fórmulas de Excel que leen los datos de B). "
            "2) Si se cambia un dato en B, el indicador y este resumen se recalculan. "
            "3) Construido = ya tiene valor; Construible = se puede armar con datos públicos y aún no tiene valor; "
            "No se puede por ahora = depende de datos que debe entregar la Alcaldía.", 7, tw)
nota(RS, 4, "Siglas: ISPC = Índice Simplificado de Participación Ciudadana; ICRCS = Índice de Condiciones y Respuesta para el "
            "Control Social; DSI = Disponibilidad de información de seguimiento; ILE = Índice de Logro Educativo; "
            "PDM = Plan de Desarrollo Municipal; EPP = ejecución presupuestal del presupuesto participativo; "
            "Matriz = matriz de indicadores de los profesores.", 7, tw, italic=True)
nota(RS, 5, "Cambio importante: la Cobertura de salud (afiliados / población) REEMPLAZA al indicador original de aseguramiento "
            "subsidiado (PDM.CS), cuya cobertura con el método del plan no se puede reproducir con datos públicos. Del mismo modo, "
            "Fidelidad a lo votado y Proyectos votados que se terminaron reemplazan a CEP, CMA e IREPP del presupuesto participativo.",
     7, tw, negrita=True, color=AZUL)
RS.row_dimensions[5].height = 44
RS.row_dimensions[3].height = 70
fila = 7
banda(RS, fila, "1. Conteo de indicadores principales por dimensión (se calcula con fórmulas sobre las hojas)", 7, color=ENC, tam=11,
      fcolor="1F3864", alto_f=22)
fila += 1
enc(RS, fila, ["Dimensión", "Hoja", "Construidos ya", "Construibles", "No se pueden por ahora", "Total principales",
               "Cuadra"])
RS.row_dimensions[fila].height = 32
fila += 1
cuentas = []
r_ini = fila
for nombre, etq in [(CS, "Control social"), (PC, "Participación ciudadana"), (PDM, "Plan de Desarrollo Municipal (PDM)"),
                    (PP, "Presupuesto participativo")]:
    u = UBICA[nombre]
    c0, c1 = u["c"]
    rb = "%s$B$%d:$B$%d" % (q(nombre), c0, c1)
    rc = "%s$C$%d:$C$%d" % (q(nombre), c0, c1)
    RS.cell(fila, 1, etq)
    RS.cell(fila, 2, nombre)
    RS.cell(fila, 3, '=COUNTIFS(%s,"Principal",%s,"Construido")' % (rb, rc))
    RS.cell(fila, 4, '=COUNTIFS(%s,"Principal",%s,"Construible")' % (rb, rc))
    RS.cell(fila, 5, '=COUNTIFS(%s,"Principal",%s,"No se puede")' % (rb, rc))
    RS.cell(fila, 6, '=COUNTIF(%s,"Principal")' % rb)
    RS.cell(fila, 7, '=IF(C%d+D%d+E%d=F%d,"Sí","Revisar")' % (fila, fila, fila, fila))
    NFORM += 5
    for k in range(1, 8):
        c = RS.cell(fila, k)
        c.border = BORDE
        c.font = fnt(bold=(k == 1))
        c.alignment = Alignment(horizontal="left" if k <= 2 else "center", vertical="center", wrap_text=True)
    RS.row_dimensions[fila].height = 18
    fila += 1
r_fin = fila - 1
RS.cell(fila, 1, "Total")
for k in (3, 4, 5, 6):
    RS.cell(fila, k, "=SUM(%s%d:%s%d)" % (L(k), r_ini, L(k), r_fin))
    NFORM += 1
RS.cell(fila, 7, '=IF(C%d+D%d+E%d=F%d,"Sí","Revisar")' % (fila, fila, fila, fila))
NFORM += 1
for k in range(1, 8):
    c = RS.cell(fila, k)
    c.border, c.fill = BORDE, fill(AZUL_CLARO)
    c.font = fnt(bold=True)
    c.alignment = Alignment(horizontal="left" if k <= 2 else "center")
TOTAL_ROW = fila
fila += 2
banda(RS, fila, "2. Qué pasó con los indicadores de la Matriz original (se calcula con fórmulas sobre el bloque A de cada hoja)", 7,
      color=ENC, tam=11, fcolor="1F3864", alto_f=22)
fila += 1
enc(RS, fila, ["Dimensión", "Hoja", "Originales de la Matriz", "Eliminados o reemplazados", "Con ajuste o sin dato (se mantienen)",
               "Nuevos (no estaban en la Matriz)", ""])
RS.row_dimensions[fila].height = 44
fila += 1
r2_ini = fila
for nombre, etq in [(CS, "Control social"), (PC, "Participación ciudadana"), (PDM, "Plan de Desarrollo Municipal (PDM)"),
                    (PP, "Presupuesto participativo")]:
    a0, a1 = UBICA[nombre]["a"]
    rb = "%s$B$%d:$B$%d" % (q(nombre), a0, a1)
    rf = "%s$F$%d:$F$%d" % (q(nombre), a0, a1)
    rf = "%s$%s$%d:$%s$%d" % (q(nombre), L(HOJAS[nombre]["ys"] and (4 + len(HOJAS[nombre]["ys"]) + (1 if HOJAS[nombre]["meta"] else 0))), a0,
                              L(4 + len(HOJAS[nombre]["ys"]) + (1 if HOJAS[nombre]["meta"] else 0)), a1)
    RS.cell(fila, 1, etq)
    RS.cell(fila, 2, nombre)
    RS.cell(fila, 6, '=COUNTIF(%s,"Nuevo")' % rb)
    RS.cell(fila, 3, '=ROWS(%s)-F%d' % (rb, fila))
    RS.cell(fila, 4, '=COUNTIF(%s,"Eliminado")+COUNTIFS(%s,"No",%s,"Reemplazado*")' % (rb, rb, rf))
    RS.cell(fila, 5, '=C%d-D%d' % (fila, fila))
    NFORM += 4
    for k in range(1, 8):
        c = RS.cell(fila, k)
        c.border = BORDE
        c.font = fnt(bold=(k == 1))
        c.alignment = Alignment(horizontal="left" if k <= 2 else "center", vertical="center", wrap_text=True)
    fila += 1
r2_fin = fila - 1
RS.cell(fila, 1, "Total")
for k in (3, 4, 5, 6):
    RS.cell(fila, k, "=SUM(%s%d:%s%d)" % (L(k), r2_ini, L(k), r2_fin))
    NFORM += 1
for k in range(1, 8):
    c = RS.cell(fila, k)
    c.border, c.fill = BORDE, fill(AZUL_CLARO)
    c.font = fnt(bold=True)
    c.alignment = Alignment(horizontal="left" if k <= 2 else "center")
TOTAL2 = fila
fila += 2
banda(RS, fila, "3. Indicadores principales: valor más reciente y semáforo (leídos de las hojas de dimensión)", 7, color=ENC, tam=11,
      fcolor="1F3864", alto_f=22)
fila += 1
enc(RS, fila, ["Dimensión", "Indicador", "Año", "Valor", "Semáforo", "Umbrales", ""])
RS.merge_cells(start_row=fila, start_column=6, end_row=fila, end_column=7)
fila += 1
PRINC = [
    (CS, "pcs", 2025, "PCS Promoción del control social"), (CS, "dsi", 2025, "DSI Disponibilidad de información de seguimiento"),
    (CS, "rpo", 2024, "RPO Respuesta a observaciones ciudadanas"),
    (CS, "icrcs_r", 2025, "ICRCS con los respaldos de VRC y RPO (el principal no tiene valor)"),
    (PC, "i2", 2025, "I2 Cobertura territorial"), (PC, "i3", 2025, "I3 Oferta formativa"),
    (PC, "ispc_r", 2025, "ISPC con pesos redistribuidos (el original no tiene valor)"),
    (PDM, "sal", 2025, "Cobertura de salud (reemplaza el aseguramiento subsidiado)"), (PDM, "cm", 2024, "CM Mortalidad vial"),
    (PDM, "ca", 2024, "CA Espacio público verde"), (PDM, "ile_a", 2025, "ILE primera infancia: avance hacia la meta 2027"),
    (PP, "fid", 2025, "Fidelidad a lo votado"), (PP, "epp", 2025, "EPP Ejecución presupuestal"),
    (PP, "ter", 2025, "Proyectos votados que se terminaron"),
]
SEMKEY = {(h, s["key"]): s["sem"] for h in HOJAS for s in HOJAS[h]["ind"]}
p0 = fila
for h, key, y, etq in PRINC:
    filai = REG_I[(h, key)]
    RS.cell(fila, 1, h)
    RS.cell(fila, 2, etq)
    RS.cell(fila, 3, y)
    RS.cell(fila, 4, "=%s%s%d" % (q(h), L(YC[h][y]), filai))
    NFORM += 1
    RS.cell(fila, 4).number_format = "0.00"
    sk = SEMKEY[(h, key)]
    if sk:
        v = "D%d" % fila
        expr = '"Sin dato"'
        for t, col in reversed(UMB[sk]):
            if t is None:
                expr = '"%s"' % NOMBRE_SEM[col]
            else:
                expr = 'IF(%s>=%s,"%s",%s)' % (v, t, NOMBRE_SEM[col], expr)
        RS.cell(fila, 5, '=IF(ISNUMBER(%s),%s,"Sin dato")' % (v, expr))
        NFORM += 1
        RS.cell(fila, 6, TXT_UMBRAL[sk])
    else:
        RS.cell(fila, 5, "Sin semáforo")
        RS.cell(fila, 6, "La Matriz no define umbral para este indicador")
    RS.merge_cells(start_row=fila, start_column=6, end_row=fila, end_column=7)
    for k in range(1, 8):
        c = RS.cell(fila, k)
        c.border = BORDE
        c.font = fnt(bold=(k == 4))
        c.alignment = Alignment(horizontal="left" if k in (1, 2, 6) else "center", vertical="center", wrap_text=True)
    RS.cell(fila, 6).font = fnt(size=9)
    RS.row_dimensions[fila].height = alto([(etq, 44), (TXT_UMBRAL.get(sk, ""), 26)])
    CELDAS.append(dict(hoja="Resumen", celda="E%d" % fila, semaforo_de=(h, key, y)))
    fila += 1
p1_ = fila - 1
for nombre_s, col in [("Verde", "C6EFCE"), ("Amarillo", "FFEB9C"), ("Naranja", "F8CBAD"), ("Rojo", "FFC7CE")]:
    RS.conditional_formatting.add("E%d:E%d" % (p0, p1_), CellIsRule(operator="equal", formula=['"%s"' % nombre_s], fill=fill(col)))
    RS.conditional_formatting.add("D%d:D%d" % (p0, p1_), FormulaRule(formula=['E%d="%s"' % (p0, nombre_s)], fill=fill(col)))
RS.conditional_formatting.add("G9:G%d" % TOTAL_ROW, CellIsRule(operator="equal", formula=['"Revisar"'], fill=fill("FFC7CE")))
fila += 1
nota(RS, fila, "Los valores son provisionales. Detalle del cálculo de cada indicador, sus datos de entrada y sus fuentes: "
               "ver las hojas de cada dimensión y 'Series en el tiempo'.", 7, tw, italic=True)
RS.freeze_panes = "A2"
RS.page_setup.orientation = "landscape"
RS.page_setup.fitToWidth = 1
RS.page_setup.fitToHeight = 1
RS.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
RS.oddFooter.center.text = "Provisional · Observatorio Ciudadano de Envigado · &P de &N"

WB._sheets = [WB["Resumen"], WB[CS], WB[PC], WB[PDM], WB[PP], WB[SE]]
WB.properties.title = "Indicadores del Observatorio Ciudadano de Envigado (provisional)"
WB.properties.creator = "Observatorio Ciudadano de Envigado"
WB.calculation.fullCalcOnLoad = True
WB.save(SAL)
print("Guardado", SAL, "· fórmulas:", NFORM)
if MAPA:
    json.dump(dict(celdas=CELDAS, serie=[(k, y, c) for k, y, c in CELDAS_SERIE], nform=NFORM,
                   total_row=TOTAL_ROW, total2=TOTAL2, princ_ini=p0, princ_fin=p1_), open(MAPA, "w"), ensure_ascii=False)
