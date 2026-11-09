#!/usr/bin/env python3
"""CS.PCS: estrategias de promoción del control social (EE ejecutadas / EP programadas, tope 100).

Fuente: Plan Indicativo de Productos, evaluación definitiva 2T-2026, hoja Plan-Indicativo-Productos,
CONSECUTIVO 544 (producto 5.1.2.1.2). Años 2024 y 2025 con puntaje; 2026 (sin cierre) solo como variable.
Resultados PROVISIONALES. Ejecutar sin -I desde la raíz: python3 scripts/indicadores/cs_pcs.py
"""
import datetime as dt
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C

COD = "CS.PCS"
SCRIPT = "scripts/indicadores/cs_pcs.py"
RUTA = ("data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/"
        "evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx")
HOJA, CONSEC = "Plan-Indicativo-Productos", 544
COL_EP = {2024: "AS", 2025: "AT", 2026: "AU"}     # META PROGRAMADA
COL_EE = {2024: "BM", 2025: "BN", 2026: "BO"}     # AVANCE EJECUTADO
CORTE_ARCHIVO = dt.date(2026, 6, 30)              # evaluación definitiva del 2T-2026
PUBLICACION = dt.date(2026, 7, 30)                # fecha en el nombre del archivo


def leer_fuente():
    ws = openpyxl.load_workbook(C.RAIZ / RUTA, data_only=True)[HOJA]
    fila = next(c.row for c in ws["A"] if c.value == CONSEC)
    assert "Estrategias de promoción del control social" in str(ws[f"AM{fila}"].value)
    datos = {}
    for a in (2024, 2025, 2026):
        assert str(a) in str(ws[f"{COL_EP[a]}1"].value) and str(a) in str(ws[f"{COL_EE[a]}1"].value)
        datos[a] = (ws[f"{COL_EE[a]}{fila}"].value, ws[f"{COL_EP[a]}{fila}"].value)
    return fila, datos


def main():
    fila_x, datos = leer_fuente()
    url = C.normalizar_url(C.url_de_ruta(RUTA))
    assert url.startswith(C.ESPEJO), f"URL no normalizada: {url!r}"
    valores, resultados = [], []
    # soporte: la fuente solo trae el total anual por año; no hay listado de estrategias (se deja constancia)
    soporte = []
    for anio in (2024, 2025, 2026):
        ee, ep = datos[anio]
        soporte.append(dict(
            id_soporte=f"CS-PCS-{anio}-001", codigo_indicador=COD, id_variable="CS.PCS.EE;CS.PCS.EP", anio=anio,
            escenario_id="defecto", elemento_tipo="agregado_anual", elemento_id=f"PCS-{anio}-AGREGADO",
            descripcion=f"Total anual de estrategias de promoción del control social: ejecutadas={ee}, programadas={ep}",
            documento="Plan Indicativo de Productos, evaluación definitiva 2T-2026 (Alcaldía de Envigado, DAP)",
            ruta_local=RUTA, enlace=C.normalizar_url(C.url_de_ruta(RUTA)),
            pagina=f"Hoja {HOJA}, fila {fila_x}, cols. {COL_EP[anio]} (programada) y {COL_EE[anio]} (ejecutada)",
            clasificacion="total reportado", justificacion="La fuente reporta solo la cifra agregada del año; no lista las "
            "estrategias una a una, así que no hay filas por estrategia", criterio="Cifra tal como la reporta el Plan Indicativo",
            lector="ejecutor_CS", incluido="sí", cuenta_en_defecto=1 if anio < 2026 else 0))
    C.escribir_parcial("CS", "soporte_conteos", soporte, codigos=[COD])
    for esc in C.escenarios(COD):
        p = esc["params"]
        trat = p["tratamiento_2026"]
        for anio in (2024, 2025, 2026):
            if anio == 2026 and trat == "excluir":
                continue
            periodo = "Semestre 1" if (anio == 2026 and trat == "semestre1_provisional") else "Anual"
            ee, ep = datos[anio]
            if anio == 2026:
                corte = CORTE_ARCHIVO
            else:
                corte = C.fecha_corte_registro(anio, "Anual", CORTE_ARCHIVO, p)
            obs_base = (f"Cifra del Plan Indicativo de Productos, evaluación definitiva 2T-2026 (corte {CORTE_ARCHIVO:%d/%m/%Y}, "
                        f"publicado {PUBLICACION:%d/%m/%Y}). Producto 5.1.2.1.2 'Estrategias de promoción del control social', "
                        f"unidad Número, indicador de producto 3.4.3F.")
            if anio == 2026:
                obs_base += (" 2026 sin cierre de vigencia: el avance es al 2T y NO se puntúa (solo variable, provisional).")
            elif p["corte_fecha_registro"] != "real":
                obs_base += f" Corte real del archivo {CORTE_ARCHIVO:%d/%m/%Y}; fecha de registro normalizada al cierre de vigencia."
            comunes = dict(esc=esc, periodo=periodo, unidad="Número", tipo_valor="Observado",
                           fuente="Plan Indicativo de Productos, evaluación definitiva 2T-2026 (Alcaldía de Envigado, DAP)",
                           documento_ruta_local=RUTA, enlace=url, fecha_corte=corte, fecha_corte_real=CORTE_ARCHIVO,
                           fecha_publicacion=PUBLICACION, script=SCRIPT, categoria_disponibilidad="Disponible",
                           horizonte_meta=str(anio), observaciones=obs_base)
            comunes["ids_soporte"] = f"CS-PCS-{anio}-001"
            valores.append(C.fila_valor("CS", "CS.PCS.EE", anio, ee, "Número",
                nombre_variable="Estrategias de promoción del control social ejecutadas",
                pagina_tabla=f"Hoja {HOJA}, fila {fila_x}, col. {COL_EE[anio]} (AVANCE EJECUTADO {anio}), CONSECUTIVO {CONSEC}",
                **{k: v for k, v in comunes.items() if k not in ("esc", "unidad")}, esc=esc))
            valores.append(C.fila_valor("CS", "CS.PCS.EP", anio, ep, "Número",
                nombre_variable="Estrategias de promoción del control social programadas",
                pagina_tabla=f"Hoja {HOJA}, fila {fila_x}, col. {COL_EP[anio]} (META PROGRAMADA {anio}), CONSECUTIVO {CONSEC}",
                **{k: v for k, v in comunes.items() if k not in ("esc", "unidad")}, esc=esc))
            if anio == 2026:
                continue
            res = C.calcular(COD, [ee, ep], {"periodo": "Anual", "territorio": "Municipio"})
            resultados.append(C.fila_resultado(COD, anio, esc, res, [ee, ep], nombre="Promoción del control social (PCS)",
                tipo="Indicador", fecha_corte_comun=corte, script=SCRIPT,
                observaciones=f"EE={ee}, EP={ep}; puntaje = MIN(100, 100*EE/EP). Provisional."))
    C.escribir_parcial("CS", "valores_variables", valores, codigos=[COD])
    C.escribir_parcial("CS", "resultados_indicadores", resultados, codigos=[COD])
    for esc in C.escenarios(COD):
        print(esc["escenario_id"])
    print({a: datos[a] for a in datos})
    for r in resultados:
        if r["escenario_id"] == "defecto" or r["anio"] == 2024:
            print(r["escenario_id"], r["anio"], r["puntaje"], r["semaforo"])
    print(len(valores), "variables;", len(resultados), "resultados")


if __name__ == "__main__":
    main()
