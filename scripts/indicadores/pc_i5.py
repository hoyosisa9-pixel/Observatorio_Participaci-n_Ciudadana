"""PC.I5 - Cumplimiento de sesiones de las instancias seleccionadas (complementario). PROVISIONAL.

I5 = 100 * V1 / V2; V1 = sesiones realizadas, V2 = sesiones previstas (V1 > V2 -> Revisar).
Fuente: PPSS 2025 (Seguimiento al Plan de Accion de la Politica de Participacion Social en Salud),
hojas PROGRAMACION (col. I "EXPRESION NUMERICA DE LA ACTIVIDAD") y SEGUIMIENTO (col. G "RESULTADO DE LA ACTIVIDAD"),
lineas E5Ld-M01-A01 (CMSSS, trimestral) y E5Ld-M02-A01 (COPACO, trimestral). E5Ld-M02-A02 (reuniones de trabajo
rotuladas "VIGENCIA 2024") se registra aparte como adicional y no entra en V1 ni V2.
Ejecutar desde la raiz: python3 scripts/indicadores/pc_i5.py
"""
import sys, datetime as dt, warnings
from pathlib import Path
RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
import comun as C
import openpyxl
warnings.filterwarnings("ignore")

SIS, IND, SCRIPT = "PC", "PC.I5", "scripts/indicadores/pc_i5.py"
R25 = ("data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/"
       "pps190ppps20251231mu000000005266.xlsx")
R24 = ("data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion/"
       "pps190spps20241231mu000000005266.xlsx")
D25 = "Seguimiento al Plan de Acción de la Política de Participación Social en Salud PPSS 2025"
D24 = "Seguimiento al Plan de Acción de la Política de Participación Social en Salud PPSS 2024"
CORTE_REAL = dt.date(2025, 12, 31)
NOMBRES = {"PC.I5.V1": "Número de sesiones realizadas de las instancias seleccionadas",
           "PC.I5.V2": "Número de sesiones previstas de las instancias seleccionadas"}
INSTANCIAS = [("CMSSS", "E5Ld", "M01", "A01", "Comité Municipal de Participación Social en Salud (CMSSS)"),
              ("COPACO", "E5Ld", "M02", "A01", "Comité de Participación Comunitaria en Salud (COPACO)")]
ADICIONAL = ("COPACO", "E5Ld", "M02", "A02")


def enlace(ruta, fallback=""):
    return C.url_de_ruta(ruta) or C.normalizar_url(fallback)


def leer_ppss25():
    """Devuelve {(linea, meta, act): dict(prog, res, fila_prog, fila_seg, desc)} desde el xlsx local."""
    wb = openpyxl.load_workbook(RAIZ / R25, data_only=True)
    out = {}
    for i, r in enumerate(wb["PROGRAMACIÓN"].iter_rows(values_only=True), 1):
        if r[3] == "E5Ld":
            out[(r[3], r[4], r[6])] = {"prog": int(r[8]), "fila_prog": i, "desc": str(r[7]).strip(),
                                       "ini": r[10], "fin": r[11]}
    for i, r in enumerate(wb["SEGUIMIENTO"].iter_rows(values_only=True), 1):
        if r[3] == "E5Ld":
            d = out[(r[3], r[4], r[5])]
            d["res"] = int(r[6]); d["fila_seg"] = i
    return out


def main():
    hoy = dt.date.today()
    datos = leer_ppss25()
    enl25 = enlace(R25)
    enl24 = enlace(R24)
    escs = C.escenarios(IND)
    # ---- soporte (una fila por sesión prevista; un bloque aparte para adicionales). Escenario de referencia: defecto
    soporte, ids_inc = [], []
    n = 0
    for sigla, lin, meta, act, nombre in INSTANCIAS:
        d = datos[(lin, meta, act)]
        assert d["prog"] == 4 and d["res"] == 4, (sigla, d)   # control: plan §control PC.I5 = 8/8
        for t in range(1, d["prog"] + 1):
            n += 1
            idr = f"PC-I5-2025-{n:03d}"
            realizada = 1 if t <= d["res"] else 0
            soporte.append(dict(
                id_soporte=idr, codigo_indicador=IND, id_variable="PC.I5.V1;PC.I5.V2", anio=2025, escenario_id="defecto",
                elemento_tipo="sesion", elemento_id=f"{sigla}-2025-T{t}",
                descripcion=f"Sesión ordinaria trimestral T{t} de {nombre}; la fecha exacta no figura en el seguimiento",
                fecha_elemento=None, documento=D25, ruta_local=R25, enlace=enl25,
                pagina=f"Hojas PROGRAMACIÓN fila {d['fila_prog']} (col. I = {d['prog']}) y SEGUIMIENTO fila {d['fila_seg']} "
                       f"(col. G = {d['res']}); línea {lin}-{meta}-{act}",
                banderas=f"instancia={sigla};prevista=1;realizada={realizada};nivel_evidencia=agregado_anual;adicional=0",
                clasificacion="prevista y realizada" if realizada else "prevista no realizada",
                justificacion=f"La programación fija {d['prog']} reuniones ordinarias al año (trimestrales) y el seguimiento "
                              f"reporta {d['res']} ejecutadas; se asigna una sesión por trimestre sin fecha individual",
                criterio="Cuenta si hace parte de la programación de la línea y el seguimiento informa su ejecución; "
                         "las actas están en SharePoint institucional sin acceso local (verificación pendiente)",
                lector="ejecutor_PC", cuenta_en_defecto=1))
            ids_inc.append(idr)
    for t in range(1, datos[("E5Ld", "M02", "A02")]["prog"] + 1):
        d = datos[("E5Ld", "M02", "A02")]
        n += 1
        soporte.append(dict(
            id_soporte=f"PC-I5-2025-{n:03d}", codigo_indicador=IND, id_variable="PC.I5.V1;PC.I5.V2", anio=2025,
            escenario_id="defecto", elemento_tipo="sesion", elemento_id=f"COPACO-TRABAJO-2025-{t}",
            descripcion="Reunión de trabajo para fortalecimiento del COPACO (adicional; línea rotulada vigencia 2024)",
            documento=D25, ruta_local=R25, enlace=enl25,
            pagina=f"Hojas PROGRAMACIÓN fila {d['fila_prog']} y SEGUIMIENTO fila {d['fila_seg']}; línea E5Ld-M02-A02",
            banderas="instancia=COPACO;prevista=1;realizada=1;nivel_evidencia=agregado_anual;adicional=1",
            clasificacion="adicional", justificacion="No son reuniones ordinarias de la instancia; la fila de la línea "
            "está rotulada 'VIGENCIA 2024' dentro del archivo 2025", criterio="Adicionales se reportan aparte y no entran en V1 ni V2",
            lector="ejecutor_PC", cuenta_en_defecto=0, incluido="no"))
    C.escribir_parcial(SIS, "soporte_conteos", soporte, codigos=[IND])
    v1 = sum(d["res"] for (l, m, a), d in datos.items() if (m, a) in (("M01", "A01"), ("M02", "A01")))
    v2 = sum(d["prog"] for (l, m, a), d in datos.items() if (m, a) in (("M01", "A01"), ("M02", "A01")))
    assert (v1, v2) == (8, 8) and len(ids_inc) == 8

    valores, resultados = [], []
    comunes = dict(codigo_indicador=IND, categoria_disponibilidad="Parcial", script=SCRIPT, tipo_valor="observado")

    def fv(idv, anio, periodo, valor, esc, **kw):
        return C.fila_valor(SIS, idv, anio, valor, "Conteo", esc, nombre_variable=NOMBRES[idv], periodo=periodo,
                            **comunes, **kw)

    for esc in escs:
        amp = esc["params"]["pc_i5_instancias"] != "salud_copaco_cmsss"
        fcorte = lambda a, p, real=None: C.fecha_corte_registro(a, p, real, esc["params"])
        # --- 2025 Anual
        for idv, val in (("PC.I5.V1", v1), ("PC.I5.V2", v2)):
            if amp:
                valores.append(fv(idv, 2025, "Anual", None, esc, estado_informacion="No disponible",
                    observaciones="Consejo Municipal de Participación, CTP y comités zonales: solo asistencia técnica por demanda, "
                                  "sin calendario de sesiones previstas; N/D por diseño del escenario"))
            else:
                valores.append(fv(idv, 2025, "Anual", val, esc, estado_informacion="Pendiente", fuente=D25,
                    documento_ruta_local=R25, enlace=enl25,
                    pagina_tabla="Hoja " + ("SEGUIMIENTO col. G" if idv.endswith("V1") else "PROGRAMACIÓN col. I") +
                                 ", líneas E5Ld-M01-A01 (CMSSS) y E5Ld-M02-A01 (COPACO)",
                    fecha_corte=fcorte(2025, "Anual", CORTE_REAL), fecha_corte_real=CORTE_REAL,
                    n_elementos_soporte=8, ids_soporte=",".join(ids_inc), observaciones=(
                        "4 CMSSS + 4 COPACO; V1 es resultado agregado reportado por la dependencia, sin actas locales (evidencias en "
                        "SharePoint). Adicionales (E5Ld-M02-A02, 4 reuniones de trabajo, rotulada VIGENCIA 2024) aparte. "
                        "El PDF 20251125_instancia-copaco-1 es ficha descriptiva, no acta.")))
        # --- 2025 semestres: sin fecha por sesión
        for per in ("Semestre 1", "Semestre 2"):
            for idv in ("PC.I5.V1", "PC.I5.V2"):
                valores.append(fv(idv, 2025, per, None, esc, estado_informacion="No disponible",
                    observaciones="El seguimiento PPSS no fecha cada sesión (solo total anual): no se puede asignar a semestres"))
        # --- 2024: PPSS 2024 sin programación
        for idv in ("PC.I5.V1", "PC.I5.V2"):
            valores.append(fv(idv, 2024, "Anual", None, esc, estado_informacion="No disponible",
                observaciones="El archivo PPSS 2024 local (Hoja1) solo contiene enlaces de evidencias por línea; no trae "
                              "programación ni resultado de sesiones. N/D"))
        # --- resultados
        casos = [(2025, "Anual", [None if amp else v1, None if amp else v2],
                  "Instancias del escenario sin calendario de sesiones previstas" if amp else None),
                 (2025, "Semestre 1", [None, None], "Seguimiento sin fecha por sesión: no hay corte semestral"),
                 (2025, "Semestre 2", [None, None], "Seguimiento sin fecha por sesión: no hay corte semestral"),
                 (2024, "Anual", [None, None], "PPSS 2024 sin programación ni resultado de sesiones")]
        for anio, per, ent, motivo in casos:
            res = C.calcular(IND, ent)
            if motivo and res["disponibilidad"] != "Calculado":
                res["motivo"] = motivo
            real = CORTE_REAL if (anio == 2025 and per == "Anual" and not amp) else None
            resultados.append(C.fila_resultado(IND, anio, esc, res, ent, periodo=per, sistema=C._sis_nombre(SIS),
                nombre="Cumplimiento de sesiones de instancias de participación", tipo="Complementario",
                fecha_corte_comun=fcorte(anio, per, real), script=SCRIPT,
                observaciones=("Complementario; solo instancias de salud del escenario por defecto (COPACO y CMSSS)"
                               if not amp else "Escenario ampliado: sin calendario")))
    C.escribir_parcial(SIS, "valores_variables", valores, codigos=[IND])
    C.escribir_parcial(SIS, "resultados_indicadores", resultados, codigos=[IND])
    print(f"V1={v1} V2={v2} soporte={len(soporte)} valores={len(valores)} resultados={len(resultados)}")
    for r in resultados:
        if r["periodo"] == "Anual" and r["anio"] == 2025:
            print(r["escenario_id"], r["puntaje"], r["semaforo"], r["disponibilidad"])


if __name__ == "__main__":
    main()
