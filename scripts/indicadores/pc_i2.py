"""PC.I2 - Cobertura territorial de la participacion (principal). PROVISIONAL.

I2 = 100 * V1 / V2; V1 = zonas con al menos una actividad participativa realizada, V2 = zonas oficiales.
Anios 2024 y 2025 (Anual, Municipio). Una fila de soporte por zona-anio-actividad (soporte_conteos, codigos=["PC.I2"]);
V1 de cada escenario se DERIVA filtrando esa tabla segun los parametros pc_i2_actividades, pc_i2_evidencia y zonas_oficiales.
Fuentes: ANTE p. 53 (asambleas zonales PDM 2024), PI26 consecutivo 531 (asambleas PP 2024), G25 p. 92 (asambleas PP 2025),
POTC (cronograma de asambleas POT 2025, solo programacion), PDM p. 558 (13 zonas; zona 1 sin comite zonal).
Ejecutar desde la raiz: python3 scripts/indicadores/pc_i2.py
"""
import sys
import datetime as dt
from pathlib import Path
RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
import comun as C

SIS, IND, SCRIPT = "PC", "PC.I2", "scripts/indicadores/pc_i2.py"
NOMBRES = {"PC.I2.V1": "Número de zonas con al menos una actividad participativa realizada para su comunidad",
           "PC.I2.V2": "Número total de zonas oficiales del municipio"}
_B = "data_raw_IIC/"
DOCS = {
    "ANTE": ("Anteproyecto del Plan de Desarrollo Municipal 2024-2027 (29-feb-2024)",
             _B + "05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/anteproyecto29022024.pdf"),
    "PI26": ("Plan Indicativo de Productos, evaluación definitiva 2T-2026",
             _B + "02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/"
                  "evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx"),
    "G25": ("Informe de gestión por dependencias 2025",
            _B + "02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf"),
    "POTC": ("Programación de las asambleas zonales del POT (cronograma, publicado 14-feb-2025)",
             _B + "03_actas_y_comites_tecnicos/envigado_gov_co/programacion-de-las-asambleas-zonales-del-plan-de-ordenamiento/"
                  "cronograma-asambleas-pot.pdf"),
    "PDM": ("Plan de Desarrollo 2024-2027 (Acuerdo 017 de 2024)",
            _B + "05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf"),
}
# Fechas leidas de la imagen del cronograma (ANTE p. 53, Figura 3) y del cronograma POT (imagen/texto, 1 pagina)
F_PDM = {9: (2024, 1, 29), 11: (2024, 1, 30), 7: (2024, 1, 31), 12: (2024, 2, 3), 13: (2024, 2, 3), 5: (2024, 2, 5),
         1: (2024, 2, 6), 8: (2024, 2, 6), 2: (2024, 2, 7), 6: (2024, 2, 8), 3: (2024, 2, 9), 4: (2024, 2, 10),
         10: (2024, 2, 10)}
F_POT = {13: (2025, 2, 22), 3: (2025, 2, 24), 11: (2025, 3, 1), 5: (2025, 3, 5), 1: (2025, 3, 8), 2: (2025, 3, 11),
         6: (2025, 3, 12), 12: (2025, 3, 15), 7: (2025, 3, 18), 8: (2025, 3, 20), 9: (2025, 3, 25), 10: (2025, 4, 1),
         4: (2025, 4, 2)}
CORTE_REAL = {2024: None, 2025: None}


def enl(clave):
    return C.url_de_ruta(DOCS[clave][1])


def construir_soporte():
    filas = []
    n = {2024: 0, 2025: 0}

    def add(anio, zona, tipo, evid, clave, pagina, desc, fecha, incluida, criterio, just, extra=""):
        n[anio] += 1
        doc, ruta = DOCS[clave]
        filas.append(dict(
            id_soporte=f"PC-I2-{anio}-{n[anio]:03d}", codigo_indicador=IND, id_variable="PC.I2.V1", anio=anio,
            territorio=f"Zona {zona}", escenario_id="defecto", elemento_tipo="zona_actividad",
            elemento_id=f"Z{zona:02d}-{anio}-{tipo.upper()}", descripcion=desc,
            fecha_elemento=dt.date(*fecha) if fecha else None, documento=doc, ruta_local=ruta, enlace=enl(clave),
            pagina=pagina, banderas=f"zona={zona};tipo={tipo};evidencia={evid}{extra}",
            clasificacion={"realizacion": "realizada (reportada)", "inferida": "realizada (inferida del agregado)",
                           "programacion": "solo programada", "ninguna": "sin actividad"}[evid],
            justificacion=just, criterio=criterio, lector="ejecutor_PC", incluido="sí" if incluida else "no",
            cuenta_en_defecto=1 if incluida else 0))

    crit = ("Cuenta si es una asamblea zonal que permite intervenir, proponer o deliberar (Matriz) y hay reporte de su "
            "realización en el documento citado; zona = 1 sola vez; actas por zona no publicadas (verificación pendiente)")
    for z in range(1, 14):   # PDM 2024
        add(2024, z, "pdm", "realizacion", "ANTE", "p. 53 (texto y Figura 3, cronograma en imagen)",
            "Asamblea zonal de formulación participativa del PDM 2024-2027", F_PDM[z], True, crit,
            "La p. 53 dice que las asambleas tuvieron lugar entre el 29-ene y el 17-feb-2024 en cada una de las 13 zonas; "
            "la fecha por zona sale del cronograma (imagen), que llega hasta el 10-feb",
            ";fecha_fuente=cronograma_imagen")
    for z in range(2, 14):   # PP 2024
        add(2024, z, "pp", "inferida", "PI26", "Hoja Plan-Indicativo-Productos, consecutivo 531, fila 534 (col. BM = 12; EG y EH)",
            "Asamblea zonal de presupuesto participativo (vigencia 2024)", None, True, crit,
            "El producto 531 reporta 12 asambleas zonales ejecutadas en 2024 'en las 12 zonas'; el listado por zona no se "
            "publica: se asignan a las zonas 2 a 13 porque la zona 1 no tiene comité zonal (PDM p. 558). Inferencia del equipo",
            ";nivel_evidencia=agregado_anual")
    add(2024, 1, "pp", "ninguna", "PI26", "Consecutivo 531, fila 534", "Sin asamblea de presupuesto participativo en la zona 1",
        None, False, "No cuenta: el reporte del PP no incluye la zona 1", "Zona 1 sin comité zonal (PDM p. 558, <1.000 habitantes)")
    for z in range(2, 14):   # PP 2025
        add(2025, z, "pp", "inferida", "G25", "p. 92 (Asambleas Zonales) y PI26 consecutivo 531, col. BN = 12",
            "Asamblea zonal de presupuesto participativo (vigencia 2026), ago-2025", None, True, crit,
            "El informe dice que se realizaron 12 asambleas zonales en agosto de 2025 'cubriendo la totalidad de las zonas'; "
            "no lista las zonas: se asignan a las zonas 2 a 13 (zona 1 sin comité, PDM p. 558). Inferencia del equipo",
            ";nivel_evidencia=agregado_anual")
    add(2025, 1, "pp", "ninguna", "G25", "p. 92", "Sin asamblea de presupuesto participativo en la zona 1", None, False,
        "No cuenta: el reporte del PP no incluye la zona 1", "Zona 1 sin comité zonal (PDM p. 558)")
    for z in range(1, 14):   # POT 2025
        add(2025, z, "pot", "programacion", "POTC", "p. 1 (cronograma, imagen)",
            "Asamblea zonal de revisión del POT (programada)", F_POT[z], False,
            "Solo cuenta si pc_i2_evidencia = incluye_programacion; en el defecto no entra porque no hay constancia de realización",
            "El archivo es un cronograma (publicado 14-feb-2025); no hay acta, informe ni asistencia de cada asamblea en el repositorio")
    return filas


def v1_escenario(soporte, params, anio):
    """Zonas cubiertas en el escenario, derivadas del soporte. Devuelve (conjunto_zonas, ids)."""
    if params["pc_i2_evidencia"] == "acta_por_zona":
        return None, []
    universo = set(range(1, 14)) if params["zonas_oficiales"] == "13" else set(range(2, 14))
    zonas, ids = set(), []
    for f in soporte:
        if f["anio"] != anio:
            continue
        b = dict(kv.split("=", 1) for kv in f["banderas"].split(";"))
        if b["evidencia"] == "ninguna":
            continue
        if params["pc_i2_actividades"] == "asambleas_pp" and b["tipo"] != "pp":
            continue
        if b["evidencia"] == "programacion" and params["pc_i2_evidencia"] != "incluye_programacion":
            continue
        z = int(b["zona"])
        if z in universo:
            zonas.add(z)
            ids.append(f["id_soporte"])
    return zonas, ids


def main():
    soporte = construir_soporte()
    C.escribir_parcial(SIS, "soporte_conteos", soporte, codigos=[IND])
    valores, resultados = [], []
    comunes = dict(codigo_indicador=IND, categoria_disponibilidad="Construible", script=SCRIPT, tipo_valor="observado")
    for esc in C.escenarios(IND):
        p = esc["params"]
        v2 = 13 if p["zonas_oficiales"] == "13" else 12
        for anio in (2024, 2025):
            fc = C.fecha_corte_registro(anio, "Anual", None, p)
            zonas, ids = v1_escenario(soporte, p, anio)
            v1 = None if zonas is None else len(zonas)
            if v1 is None:
                valores.append(C.fila_valor(SIS, "PC.I2.V1", anio, None, "Conteo", esc, nombre_variable=NOMBRES["PC.I2.V1"],
                    estado_informacion="No disponible", **comunes, fecha_corte=fc,
                    observaciones="pc_i2_evidencia=acta_por_zona: no hay actas ni asistencia por zona publicadas; N/D"))
            else:
                clave = "G25" if anio == 2025 else ("ANTE" if p["pc_i2_actividades"] == "asambleas_pp_pdm_pot" else "PI26")
                pag = {"G25": "p. 92 (asambleas zonales) y cronograma POT p. 1 si aplica", "ANTE": "p. 53 (asambleas PDM) y PI26 consecutivo 531",
                       "PI26": "Hoja Plan-Indicativo-Productos, consecutivo 531, fila 534"}[clave]
                valores.append(C.fila_valor(SIS, "PC.I2.V1", anio, v1, "Conteo", esc, nombre_variable=NOMBRES["PC.I2.V1"],
                    estado_informacion="Pendiente", **comunes, fuente=DOCS[clave][0], documento_ruta_local=DOCS[clave][1],
                    enlace=enl(clave), pagina_tabla=pag, fecha_corte=fc, n_elementos_soporte=len(ids),
                    ids_soporte=",".join(ids), observaciones=(
                        "Zonas distintas con actividad según soporte_conteos (una fila por zona-año-actividad). Asambleas PP "
                        "por zona inferidas del agregado (zonas 2-13); actas por zona no publicadas. Universo exploratorio, "
                        "no validado por Planeación")))
            valores.append(C.fila_valor(SIS, "PC.I2.V2", anio, v2, "Conteo", esc, nombre_variable=NOMBRES["PC.I2.V2"],
                estado_informacion="Pendiente", **{**comunes, "categoria_disponibilidad": "Disponible"},
                fuente=DOCS["PDM"][0], documento_ruta_local=DOCS["PDM"][1], enlace=enl("PDM"),
                pagina_tabla="p. 558 (13 zonas; 12 comités; zona 1 pendiente) y Plan Indicativo columnas C:O (ZONA 1 a ZONA 13)",
                fecha_corte=fc, observaciones="Falta el acto oficial que fije las zonas en el repositorio; validar con Planeación"
                + ("" if v2 == 13 else ". Escenario 12_con_comite: solo zonas 2-13")))
            res = C.calcular(IND, [v1, v2])
            if res["disponibilidad"] != "Calculado" and v1 is None:
                res["motivo"] = "Sin actas por zona publicadas (pc_i2_evidencia=acta_por_zona)"
            resultados.append(C.fila_resultado(IND, anio, esc, res, [v1, v2], sistema=C._sis_nombre(SIS),
                nombre="Cobertura territorial de la participación", tipo="Principal", fecha_corte_comun=fc, script=SCRIPT,
                observaciones="Solo Municipio; cobertura por zona (sí/no) en soporte_conteos"))
    C.escribir_parcial(SIS, "valores_variables", valores, codigos=[IND])
    C.escribir_parcial(SIS, "resultados_indicadores", resultados, codigos=[IND])
    print(f"soporte={len(soporte)} valores={len(valores)} resultados={len(resultados)}")
    for r in resultados:
        if r["escenario_id"] == "defecto" or r["entrada_1"] != 12 and r["entrada_1"] != 13:
            pass
    defs = [r for r in resultados if r["es_defecto"]]
    for r in defs:
        print("DEFECTO", r["anio"], r["entrada_1"], r["entrada_2"], r["puntaje"], r["semaforo"], r["disponibilidad"])
    from collections import Counter
    print(Counter((r["anio"], r["disponibilidad"], r["semaforo"], r["puntaje"]) for r in resultados))


if __name__ == "__main__":
    main()
