#!/usr/bin/env python3
"""CS.DSI: disponibilidad de información de seguimiento = 100 * MD / MS (sin tope; MD > MS -> Revisar).

MS = metas seleccionadas (parámetro dsi_lista_metas). MD = metas con valor observado publicado para la vigencia,
construido meta por meta con tabla de soporte auditable (soporte_conteos: una fila por meta x año x escenario).
Años 2024 y 2025 (2026 sin cierre: no se calcula). Todo PROVISIONAL.

Regla de decisión por (meta, año, escenario):
  1. Candidatos = evidencias documentadas cuya condición de parámetros pdm_cm_fuente / pdm_ca_fuente / pdm_cs_fuente se cumple.
  2. dsi_fuente_valor: solo_alcaldia = publica la Alcaldía y la definición coincide; oficial_misma_definicion = la definición
     y la unidad coinciden (Alcaldía o entidad oficial); oficial_incluye_proxy = también cifras oficiales de otra definición.
  3. ventana_documental_cs: guia_31mar_siguiente = solo cuenta si la fecha de publicación documentada es <= 31-mar del año
     siguiente (si la fecha no se pudo verificar, NO cuenta: criterio conservador, rotulado); hasta_fecha_consulta = todo lo publicado.
  4. Una meta cuenta en MD si algún candidato cumple 2 y 3.
Defecto = 7 metas del ICPD-6; la matriz sugería unas 10; la lista la definen los profesores.
Ejecutar sin -I desde la raíz: python3 scripts/indicadores/cs_dsi.py
"""
import datetime as dt
import sys
import warnings
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C

warnings.filterwarnings("ignore")
COD = "CS.DSI"
SCRIPT = "scripts/indicadores/cs_dsi.py"
NOTA_LISTA = "Defecto = 7 metas del ICPD-6; la matriz sugería unas 10; la lista la definen los profesores."
AN = (2024, 2025)
D = "data_raw_IIC/"
R_AP91 = D + "05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx"
R_PI26 = (D + "02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/"
          "evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx")
R_G24 = D + "02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf"
R_G25 = D + "02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf"
R_POT = D + "05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/8-dimension-funcional-_-espacio-publico.pdf"
R_BAL = D + "05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf"
R_TD = D + "05_fuentes_externas_y_pot/dnp_terridata/TerriData05266f.txt.zip"
R_MS = D + "06_descargas_indicadores/minsalud/cifras-afiliacion-salud-dic-2025.zip"
R_POL = D + "06_descargas_indicadores/policia_datosgov/uav5-b85g_homicidio_accidente_transito_envigado.csv"
U_FOR = "https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf"
U_MS = "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip"
U_POL = "https://www.datos.gov.co/resource/uav5-b85g.csv?cod_muni=05266"
U_TD = "https://terridata.dnp.gov.co/assets/docs/txt/entidades/TerriData05266f.txt.zip"

# Lista de metas (Apéndice 9.1.2, hoja 'Apéndice9.1.2.'): fila -> (clave, nombre del indicador de resultado, unidad)
METAS = {
    16: ("CS", "Cobertura del aseguramiento en salud del Régimen Subsidiado", "Porcentaje"),
    24: ("CD", "Cobertura de servicios deportivos de disciplinas tradicionales y nuevas tendencias", "Porcentaje"),
    71: ("ILE.MA", "Tasa de cobertura en atención integral de niños y niñas (primera infancia)", "Porcentaje"),
    74: ("ILE.MS", "Instituciones educativas oficiales que mejoran la clasificación de planteles", "Porcentaje"),
    119: ("CM", "Tasa de mortalidad por siniestros viales", "Tasa por cada 100.000 habitantes"),
    129: ("CA", "Espacio público verde", "Metros cuadrados por habitante"),
    146: ("CP", "Implementación de los procesos de participación ciudadana", "Porcentaje"),
    147: ("G147", "Percepción ciudadana sobre la claridad y oportunidad de la información (Encuesta de Opinión, DAP)", "Porcentaje"),
    148: ("G148", "Índice de Desempeño Institucional (DAFP)", "Número"),
    151: ("G151", "Índice de Gobierno Digital (MinTIC)", "Número"),
}
LISTAS = {"icpd7": [16, 24, 71, 74, 119, 129, 146],
          "icpd7_mas3_gobernanza": [16, 24, 71, 74, 119, 129, 146, 147, 148, 151]}


def ev(anio, cond, valor, pub, defn, doc, ruta, enlace, pagina, fecha, base_fecha, nota=""):
    return dict(anio=anio, cond=cond, valor=valor, pub=pub, defn=defn, doc=doc, ruta=ruta, enlace=enlace, pagina=pagina,
                fecha=fecha, base_fecha=base_fecha, nota=nota)


F_G24, F_G25 = dt.date(2025, 1, 30), dt.date(2026, 1, 30)
B_INV = "fecha del inventario + orden de carpeta content/files (la fecha del inventario no es confiable por sí sola)"
B_NV = "fecha de publicación no verificada"
# Evidencias documentadas (notas_PDM.md, variables_PDM.csv, notas_CS.md). cond: (parámetro, opción) o None.
EVID = {
    16: [
        ev(2024, ("pdm_cs_fuente", "minsalud_rs_sobre_pob_total"), "12,20 % (afiliados RS / población total)", "externo", "otra",
           "TerriData (DNP), Cobertura del régimen subsidiado", R_TD, U_TD, "Archivo TerriData05266f, serie 2024", None, B_NV,
           "Otra definición: RS sobre población total, no el 99,5 del plan (BDUA/SisMaster)."),
        ev(2025, ("pdm_cs_fuente", "minsalud_rs_sobre_pob_total"), "29.794 afiliados RS / 250.767 hab. = 11,88 %", "externo", "otra",
           "MinSalud, Cifras de afiliación en salud, dic-2025 (hoja CoberturaMunicipio)", R_MS, U_MS, "Hoja CoberturaMunicipio, fila ENVIGADO",
           None, B_NV, "Otra definición: RS sobre población total, no el 99,5 del plan."),
        ev(2025, ("pdm_cs_fuente", "minsalud_cobertura_total"), "87,65 % (cobertura total SGSSS)", "externo", "otra",
           "MinSalud, Cifras de afiliación en salud, dic-2025 (hoja CoberturaMunicipio)", R_MS, U_MS, "Hoja CoberturaMunicipio, fila ENVIGADO",
           None, B_NV, "Otra definición; la Matriz prohíbe la cobertura total para PDM.CS. Dic-2024 solo en Tableau (no consultado)."),
    ],
    24: [
        ev(2025, None, "38 de 40 disciplinas implementadas (Índer)", "alcaldia", "otra",
           "Informe de gestión por dependencias 2025 (Índer)", R_G25, None, "PDF p. 5", F_G25, B_INV,
           "Conteo de disciplinas (producto), no cobertura porcentual."),
    ],
    71: [
        ev(2024, None, "2.161 niños con educación inicial con atención integral (meta del producto 1.950)", "alcaldia", "otra",
           "Balance PDM 2024, Educación oficial (ppt)", R_BAL, None, "PDF p. 24 (producto 2.1.1.1.1)", None, B_NV + "; inventario 2021-09-14 inválido; carpeta 002717 posterior a la 002703 (doc. del 31-mar-2025)",
           "Numerador del producto; el porcentaje de cobertura no está publicado."),
        ev(2025, None, "2.092 niños ≤5 años con atención integral", "alcaldia", "otra",
           "Informe de gestión por dependencias 2025", R_G25, None, "PDF p. 150", F_G25, B_INV,
           "Numerador del producto; el porcentaje de cobertura no está publicado."),
    ],
    74: [],
    119: [
        ev(2024, ("pdm_cm_fuente", "inmlcf_forensis"), "7,25 muertes por transporte por 100.000 hab. (18 muertes)", "externo", "misma",
           "INMLCF, Forensis 2024 (cifra citada; PDF local pendiente de descarga manual)", "", U_FOR, "PDF p. 440 (tabla 14) y p. 37 (tabla 5)",
           None, B_NV + " (Forensis 2024 se publica habitualmente después de marzo; verificar)", "Cifra citada en estado Pendiente; verificar contra el PDF local."),
        ev(2024, ("pdm_cm_fuente", "policia_datosgov"), "7,7 por 100.000 (homicidios en accidente de tránsito; 20 casos)", "externo", "otra",
           "TerriData (DNP) / Policía Nacional vía datos.gov.co (uav5-b85g)", R_TD, U_TD, "TerriData, Convivencia y seguridad, 2024", None, B_NV,
           "Otra definición: homicidio en accidente de tránsito, no muertes por transporte."),
        ev(2025, ("pdm_cm_fuente", "policia_datosgov"), "11,3 por 100.000 (30 casos; 30/264.637)", "externo", "otra",
           "Policía Nacional vía datos.gov.co (uav5-b85g) y TerriData", R_POL, U_POL, "Conjunto uav5-b85g, cod_muni 05266, 2025", None, B_NV,
           "Otra definición: homicidio en accidente de tránsito. INMLCF 2025 no publicado."),
    ],
    129: [
        ev(2024, ("pdm_ca_fuente", "pot_epv_urbano"), "5,75 m2/hab (EPV urbano 2024)", "alcaldia", "misma",
           "Documento diagnóstico POT, 8. Dimensión funcional - espacio público", R_POT, None, "PDF p. 87, tabla 39",
           dt.date(2025, 10, 24), B_INV, "Serie distinta de la línea base 3,88 (AMVA 2022: 4,66); se acepta por el parámetro pdm_ca_fuente."),
        ev(2024, ("pdm_ca_fuente", "pot_epv_total"), "69,41 m2/hab (EPV total municipal)", "alcaldia", "otra",
           "Documento diagnóstico POT, 8. Dimensión funcional - espacio público", R_POT, None, "PDF p. 86, tabla 38",
           dt.date(2025, 10, 24), B_INV, "EPV total: no comparable con la meta 3,88."),
    ],
    146: [],
    147: [], 148: [], 151: [],
}
SIN_EVID = {
    16: "Sin cifra con el método del plan (99,5 por BDUA/SisMaster, población objetivo) publicada para 2024-2025.",
    24: "Sin cobertura porcentual de servicios deportivos publicada para esta vigencia.",
    71: "Sin porcentaje de cobertura ni numerador publicado para esta vigencia.",
    74: "Sin valor del indicador: solo menciones de casos aislados (balance 2024: p. 101 narra una mejora, p. 26 reporta logro 0 en incentivos); clasificación ICFES por plantel no descargada.",
    119: "Sin cifra para esta vigencia con la fuente del escenario (INMLCF 2025 no publicado).",
    129: "Sin valor para esta vigencia con la fuente del escenario.",
    146: "Sin valor observado: evaluaciones del PDM (IV trim. 2024 y 2025) solo presentan productos; el PDM p. 561 trae 3,23 de 2022 (otra magnitud).",
    147: "Sin evidencia en el repositorio (Encuesta de Opinión del DAP no localizada); búsqueda externa no realizada en esta tanda.",
    148: "Sin evidencia en el repositorio (resultado FURAG/IDI del DAFP no localizado ni en TerriData); búsqueda externa no realizada en esta tanda.",
    151: "Sin evidencia en el repositorio (índice de Gobierno Digital de MinTIC no localizado); búsqueda externa no realizada en esta tanda.",
}
CRITERIO = {
    "solo_alcaldia": "Cuenta solo si la Alcaldía publica el valor con definición y unidad de la meta",
    "oficial_misma_definicion": "Cuenta si la Alcaldía o una entidad oficial publica el valor con igual definición y unidad que la meta",
    "oficial_incluye_proxy": "Cuenta también una cifra oficial de otra definición (proxy)",
}
VENT = {"guia_31mar_siguiente": "publicación <= 31-mar del año siguiente (fecha no verificada = no cuenta)",
        "hasta_fecha_consulta": "todo lo publicado hasta la fecha de consulta"}


def productos_pi26():
    """Control (dsi_lista_metas = productos_plan_indicativo): productos con meta > 0 y avance numérico publicado."""
    ws = openpyxl.load_workbook(C.RAIZ / R_PI26, data_only=True)["Plan-Indicativo-Productos"]
    cols = {2024: ("AS", "BM"), 2025: ("AT", "BN")}
    out = {}
    for a, (cm, ce) in cols.items():
        assert str(a) in str(ws[f"{cm}1"].value) and str(a) in str(ws[f"{ce}1"].value)
        ms = md = 0
        for r in range(2, ws.max_row + 1):
            if not isinstance(ws[f"A{r}"].value, (int, float)):
                continue
            m = ws[f"{cm}{r}"].value
            if isinstance(m, (int, float)) and m > 0:
                ms += 1
                if isinstance(ws[f"{ce}{r}"].value, (int, float)):
                    md += 1
        out[a] = (md, ms)
    return out


def evaluar(fila, anio, p):
    """Devuelve (incluida, evidencia o None, motivo)."""
    cands = [e for e in EVID[fila] if e["anio"] == anio and (e["cond"] is None or p[e["cond"][0]] == e["cond"][1])]
    if not cands:
        return False, None, SIN_EVID[fila]
    fuente, vent = p["dsi_fuente_valor"], p["ventana_documental_cs"]
    motivos = []
    for e in cands:
        if fuente == "solo_alcaldia" and not (e["pub"] == "alcaldia" and e["defn"] == "misma"):
            motivos.append("no cumple solo_alcaldia (publicador externo o definición distinta)")
            continue
        if fuente == "oficial_misma_definicion" and e["defn"] != "misma":
            motivos.append("definición distinta a la de la meta (proxy): no cuenta con oficial_misma_definicion")
            continue
        if vent == "guia_31mar_siguiente":
            lim = dt.date(anio + 1, 3, 31)
            if e["fecha"] is None:
                motivos.append("fecha de publicación no verificada: no cuenta con guia_31mar_siguiente")
                continue
            if e["fecha"] > lim:
                motivos.append(f"publicado {e['fecha']:%d/%m/%Y}, después del límite {lim:%d/%m/%Y}")
                continue
        return True, e, "Cumple los criterios del escenario."
    return False, cands[0], "; ".join(dict.fromkeys(motivos))


def main():
    prod = productos_pi26()
    url_ap91 = C.url_de_ruta(R_AP91) or C.normalizar_url(
        "https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002340/116989_apendice91planindicativo27052024.xlsx")
    soporte, valores, resultados = [], [], []
    cont = {}

    def nuevo_id(anio):
        cont[anio] = cont.get(anio, 0) + 1
        return f"CS-DSI-{anio}-{cont[anio]:03d}"

    for esc in C.escenarios(COD):
        p = esc["params"]
        lista = p["dsi_lista_metas"]
        for anio in AN:
            corte = C.fecha_corte_registro(anio, "Anual", dt.date(anio, 12, 31), p)
            ids_inc = []
            if lista == "productos_plan_indicativo":
                md, ms = prod[anio]
                sid = nuevo_id(anio)
                ids_inc.append(sid)
                soporte.append(dict(
                    id_soporte=sid, codigo_indicador=COD, id_variable="CS.DSI.MD;CS.DSI.MS", anio=anio, escenario_id=esc["escenario_id"],
                    elemento_tipo="producto_agregado", elemento_id=f"PI26-productos-meta-mayor-0-{anio}",
                    descripcion=f"Productos del Plan Indicativo con meta programada > 0 en {anio}: {ms}; con avance numérico publicado: {md}",
                    documento="Plan Indicativo de Productos, evaluación definitiva 2T-2026",
                    ruta_local=R_PI26, enlace=C.url_de_ruta(R_PI26), pagina=f"Hoja Plan-Indicativo-Productos, columnas {'AS/BM' if anio == 2024 else 'AT/BN'}",
                    fecha_publicacion_doc=dt.date(2026, 7, 30), banderas="solo_control;DSI_100_por_construccion;publicado_despues_de_ventana",
                    clasificacion="control", lector="ejecutor_CS", cuenta_en_defecto=1, incluido="sí",
                    justificacion="Opción de control: MD = MS por construcción. Las evaluaciones EV24/EV25 (dentro de ventana) presentan estos mismos productos pero no se contaron uno a uno.",
                    criterio="Todos los productos con meta > 0; " + VENT[p["ventana_documental_cs"]]))
                fuente_md, pag_md = "Plan Indicativo de Productos, evaluación definitiva 2T-2026", "Hoja Plan-Indicativo-Productos (control)"
            else:
                filas = LISTAS[lista]
                md = ms = 0
                for fila in filas:
                    ms += 1
                    inc, e, motivo = evaluar(fila, anio, p)
                    md += inc
                    sid = nuevo_id(anio)
                    clave, nombre, unidad = METAS[fila]
                    if inc:
                        ids_inc.append(sid)
                    if e:
                        clasif = "valor_misma_definicion" if e["defn"] == "misma" else "valor_otra_definicion_proxy"
                        desc = f"{nombre} ({unidad}). Valor publicado {anio}: {e['valor']}"
                        doc, ruta, enlace, pag = e["doc"], e["ruta"], C.normalizar_url(e["enlace"] or C.url_de_ruta(e["ruta"])), e["pagina"]
                        fpub, bandera = e["fecha"], f"publicador={e['pub']};definicion={e['defn']};{e['base_fecha'] if e['fecha'] is None else 'fecha_por_inventario_y_carpeta'}"
                        just = (motivo + " " + e["nota"]).strip()
                    else:
                        clasif, desc = "sin_valor_publicado", f"{nombre} ({unidad}). Sin valor observado publicado para {anio}"
                        doc = "Sin documento con valor: revisados evaluaciones PDM IV trim. 2024 y 2025, informes de gestión 2024 y 2025, TerriData y fuentes de las notas de disponibilidad"
                        ruta, enlace, pag, fpub, bandera, just = "", "", "No aplica (sin valor publicado)", None, "sin_evidencia", motivo
                    soporte.append(dict(
                        id_soporte=sid, codigo_indicador=COD, id_variable="CS.DSI.MD;CS.DSI.MS", anio=anio, escenario_id=esc["escenario_id"],
                        elemento_tipo="meta", elemento_id=f"AP91-fila{fila}-{clave}", descripcion=desc, fecha_elemento=fpub,
                        documento=doc, ruta_local=ruta, enlace=enlace, pagina=pag, fecha_publicacion_doc=fpub, banderas=bandera,
                        clasificacion=clasif, justificacion=just,
                        criterio=f"{CRITERIO[p['dsi_fuente_valor']]}; ventana: {VENT[p['ventana_documental_cs']]}. Meta de la lista del parámetro dsi_lista_metas={lista} (Apéndice 9.1.2 fila {fila}).",
                        lector="ejecutor_CS", cuenta_en_defecto=int(inc), incluido="sí" if inc else "no"))
                fuente_md = "Soporte meta por meta (soporte_conteos): documentos de la Alcaldía y fuentes oficiales externas"
                pag_md = "Ver ids_soporte en soporte_conteos.csv"
            res = C.calcular(COD, [md, ms], {"periodo": "Anual", "territorio": "Municipio"})
            obs = (f"{NOTA_LISTA} Lista del escenario: {lista} (MS = {ms}). Resultado PROVISIONAL; los valores observados son "
                   f"proxies o cifras no verificadas en varios casos (ver soporte_conteos).")
            com = dict(unidad="Número de metas", tipo_valor="Construido", fecha_corte=corte, fecha_corte_real=dt.date(anio, 12, 31),
                       fecha_publicacion=None, script=SCRIPT, categoria_disponibilidad="Construible", horizonte_meta=str(anio),
                       n_elementos_soporte=ms)
            valores.append(C.fila_valor("CS", "CS.DSI.MD", anio, md, "Número de metas", esc=esc,
                nombre_variable="Metas seleccionadas con valor observado publicado para la vigencia",
                fuente=fuente_md, documento_ruta_local="outputs/indicadores/parciales/CS/soporte_conteos.csv", enlace=url_ap91,
                pagina_tabla=pag_md, ids_soporte=",".join(ids_inc), observaciones=obs,
                **{k: v for k, v in com.items() if k != "unidad"}))
            valores.append(C.fila_valor("CS", "CS.DSI.MS", anio, ms, "Número de metas", esc=esc,
                nombre_variable="Metas seleccionadas (lista del parámetro dsi_lista_metas)",
                fuente="Apéndice 9.1.2 del Plan Indicativo del PDM 2024-2027 (lista de metas definida por parámetro)" if lista != "productos_plan_indicativo"
                       else "Plan Indicativo de Productos, evaluación definitiva 2T-2026",
                documento_ruta_local=R_AP91 if lista != "productos_plan_indicativo" else R_PI26,
                enlace=url_ap91 if lista != "productos_plan_indicativo" else C.url_de_ruta(R_PI26),
                pagina_tabla=("Hoja Apéndice9.1.2., filas " + ", ".join(map(str, LISTAS[lista]))) if lista != "productos_plan_indicativo"
                             else "Hoja Plan-Indicativo-Productos, productos con meta > 0",
                observaciones=obs, **{k: v for k, v in com.items() if k != "unidad"}))
            resultados.append(C.fila_resultado(COD, anio, esc, res, [md, ms], nombre="Disponibilidad de información de seguimiento (DSI)",
                tipo="Indicador", fecha_corte_comun=corte, script=SCRIPT,
                observaciones=f"MD={md}, MS={ms}; puntaje = 100*MD/MS. {NOTA_LISTA} Provisional."))
    C.escribir_parcial("CS", "soporte_conteos", soporte, codigos=[COD])
    C.escribir_parcial("CS", "valores_variables", valores, codigos=["CS.DSI"])
    C.escribir_parcial("CS", "resultados_indicadores", resultados, codigos=[COD])
    print(prod)
    for r in resultados:
        print(r["escenario_id"][:90].ljust(92), r["anio"], r["entrada_1"], r["entrada_2"], r["puntaje"], r["semaforo"])
    print(len(soporte), "soporte;", len(valores), "variables;", len(resultados), "resultados")


if __name__ == "__main__":
    main()
