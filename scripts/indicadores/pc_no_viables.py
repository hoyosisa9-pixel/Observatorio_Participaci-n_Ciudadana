"""Variables de los indicadores PC NO viables por ahora (PC.I1, PC.I4, PC.I6). PROVISIONAL; NO se calcula ningun indicador.

- PC.I1.V2 y PC.I4.V2: universo EXPLORATORIO de procesos participativos cerrados (2024: 3; 2025: 5), con una fila de soporte por
  proceso. El universo NO es oficial: lo definen los profesores (parametro pc_universo_procesos). V1 de I1 e I4 queda "No
  disponible": no se verifico proceso por proceso la incorporacion del aporte en una decision formal (I1) ni que el resultado
  publicado identifique el proceso, sintetice aportes y comunique decisiones (I4).
- PC.I6.V1: lista parcial (cota inferior) de actividades con canal virtual para intervenir: 1 en 2024 y 1 en 2025; PC.I6.V2 N/D.
Resultados: se escriben filas con puntaje vacio, disponibilidad N/D y no_calculado_por_etapa=True.
Ejecutar desde la raiz: python3 scripts/indicadores/pc_no_viables.py
"""
import sys
import datetime as dt
from pathlib import Path
RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
import comun as C

SIS, SCRIPT = "PC", "scripts/indicadores/pc_no_viables.py"
_B = "data_raw_IIC/"
D = {
    "ANTE": ("Anteproyecto del Plan de Desarrollo Municipal 2024-2027 (29-feb-2024)",
             _B + "05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/anteproyecto29022024.pdf"),
    "I88": ("Comunicado: 88 iniciativas ciudadanas postuladas al presupuesto participativo (resumen)",
            _B + "01_propuestas_y_acuerdos/envigado_gov_co/88-iniciativas-ciudadanas-fueron-postuladas-para-el-515924/"
                 "resumeniniciativascomunitariaspp2024.pdf"),
    "I90": ("Comunicado: 90 iniciativas ciudadanas postuladas al presupuesto participativo 2025-2026",
            _B + "01_propuestas_y_acuerdos/envigado_gov_co/90-iniciativas-ciudadanas-fueron-postuladas-para-el-694496/"
                 "90-iniciativas-ciudadanas-fueron-postuladas-para-el-proceso-de-presupuesto-participativo-2025--2026.pdf"),
    "QA24": ("Preguntas y recomendaciones, audiencia pública de rendición de cuentas del Alcalde 2024",
             _B + "03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/"
                  "129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf"),
    "QA25": ("Preguntas y respuestas, rendición de cuentas 2025 (Alcalde)",
             _B + "02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/"
                  "preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf"),
    "G25": ("Informe de gestión por dependencias 2025",
            _B + "02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf"),
    "POTC": ("Programación de las asambleas zonales del POT (cronograma, publicado 14-feb-2025)",
             _B + "03_actas_y_comites_tecnicos/envigado_gov_co/programacion-de-las-asambleas-zonales-del-plan-de-ordenamiento/"
                  "cronograma-asambleas-pot.pdf"),
    "PLZ": ("Planes zonales de desarrollo PLAN-ZONA-2 a PLAN-ZONA-13 (EFPC)",
            _B + "01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-06/PLAN-ZONA-2.pdf"),
    "PI26": ("Plan Indicativo de Productos, evaluación definitiva 2T-2026",
             _B + "02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/"
                  "evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx"),
}
for _k, (_n, _r) in D.items():
    assert (RAIZ / _r).exists(), _r

# (anio, clave_doc, pagina, nombre, fecha de cierre o None, descripcion)
PROCESOS = [
    (2024, "ANTE", "p. 53", "Formulación participativa del PDM 2024-2027 (asambleas zonales)", None,
     "Asambleas en las 13 zonas entre el 29-ene y el 17-feb-2024; la fecha de cierre (aprobación del PDM, Acuerdo 017 de 2024) no se verificó"),
    (2024, "I88", "p. 1", "Postulación de iniciativas de presupuesto participativo vigencia 2025", dt.date(2024, 4, 30),
     "Postulación del 17-mar al 30-abr-2024; 88 iniciativas"),
    (2024, "QA24", "pp. 1-24", "Audiencia pública de rendición de cuentas del Alcalde 2024", dt.date(2024, 11, 26),
     "Audiencia del 26-nov-2024 con preguntas y recomendaciones ciudadanas"),
    (2025, "G25", "p. 92 (Ejercicios de Presupuesto Participativo)", "Ciclo de presupuesto participativo vigencia 2026", dt.date(2025, 8, 24),
     "Convocatoria de iniciativas 28-feb a 31-mar-2025; votación 11 a 24-ago-2025 (comunicado de 90 iniciativas, p. 1)"),
    (2025, "POTC", "p. 1", "Asambleas zonales de revisión del POT", None,
     "Cronograma de asambleas del 22-feb al 2-abr-2025; la realización y la síntesis no se verificaron localmente"),
    (2025, "PLZ", "documentos PLAN-ZONA-2 a PLAN-ZONA-13", "Actualización de los planes zonales de desarrollo", None,
     "12 documentos publicados en la EFPC (junio de 2025); fecha de cierre no verificada"),
    (2025, "G25", "p. 92 (Rendición de Cuentas)", "Rendición de cuentas del presupuesto participativo 2024", dt.date(2025, 2, 12),
     "Realizada el 12-feb-2025 en el auditorio de la Personería, con asistencia presencial y transmisión"),
    (2025, "QA25", "pp. 1-5", "Rendición de cuentas del Alcalde 2025", None,
     "Audiencia de fines de 2025 (nov-2025 según las notas de disponibilidad; fecha exacta no verificada)"),
]
ACT_VIRTUALES = [
    (2024, "QA24", "pp. 6-8", "Audiencia de rendición de cuentas del Alcalde, 26-nov-2024", dt.date(2024, 11, 26),
     "Preguntas de participantes virtuales (filas rotuladas 'Virtual' o 'Participante virtual' en el cuadro de preguntas)"),
    (2025, "G25", "p. 92 (Plataforma Web)", "Votación y priorización del presupuesto participativo en la plataforma web", dt.date(2025, 8, 24),
     "Votación del 11 al 24-ago-2025 en la plataforma web; más de 12.000 ciudadanos registrados (PI26 consecutivo 529: plataforma en operación)"),
]
NOMBRES = {
    "PC.I1.V1": "Número de procesos participativos cerrados que incorporaron al menos un aporte ciudadano en una decisión formal",
    "PC.I1.V2": "Número total de procesos participativos cerrados en los que se recibieron aportes ciudadanos",
    "PC.I4.V1": "Número de procesos participativos cerrados con resultados publicados y consultables",
    "PC.I4.V2": "Número total de procesos participativos cerrados",
    "PC.I6.V1": "Número de actividades participativas realizadas que habilitaron un canal virtual para intervenir",
    "PC.I6.V2": "Número total de actividades participativas realizadas"}
NOM_IND = {"PC.I1": "Incorporación de aportes ciudadanos", "PC.I4": "Publicación de resultados participativos",
           "PC.I6": "Disponibilidad de participación virtual"}
OBS_UNIV = ("Universo exploratorio NO oficial: lo definen los profesores (no existe inventario oficial de procesos participativos "
            "cerrados); indicador no calculado en esta etapa")


def soporte():
    filas, n = [], {}

    def add(ind, idv, anio, clave, pag, tipo, nombre, fecha, desc, extra, clasif):
        n[(ind, anio)] = n.get((ind, anio), 0) + 1
        doc, ruta = D[clave]
        filas.append(dict(
            id_soporte=f"PC-{ind.split('.')[1]}-{anio}-{n[(ind, anio)]:03d}", codigo_indicador=ind, id_variable=idv,
            anio=anio, escenario_id="defecto", elemento_tipo=tipo, elemento_id=f"{ind.split('.')[1]}-{anio}-{n[(ind, anio)]:02d}",
            descripcion=f"{nombre}. {desc}", fecha_elemento=fecha, documento=doc, ruta_local=ruta, enlace=C.url_de_ruta(ruta),
            pagina=pag, banderas=extra, clasificacion=clasif,
            justificacion=OBS_UNIV if ind != "PC.I6" else "Lista parcial (cota inferior); el universo de actividades no está totalizado",
            criterio=("Cuenta como proceso del universo exploratorio si hubo participación ciudadana con aportes y cierre en el año"
                      if ind != "PC.I6" else "Cuenta si la actividad habilitó un canal virtual para preguntar, proponer o deliberar; "
                      "una transmisión para observar no basta"),
            lector="ejecutor_PC", incluido="sí", cuenta_en_defecto=1))

    for anio, clave, pag, nombre, fecha, desc in PROCESOS:
        add("PC.I1", "PC.I1.V2;PC.I4.V2", anio, clave, pag, "proceso", nombre, fecha, desc,
            "universo=exploratorio;oficial=0;incorporacion_verificada=0;resultados_verificados=0;tambien_universo_I4=1",
            "proceso del universo exploratorio")
    for anio, clave, pag, nombre, fecha, desc in ACT_VIRTUALES:
        add("PC.I6", "PC.I6.V1", anio, clave, pag, "actividad_virtual", nombre, fecha, desc,
            "canal_virtual=1;tipo_canal=pregunta_o_votacion;lista=parcial;oficial=0", "actividad con canal virtual")
    return filas


def main():
    sp = soporte()
    C.escribir_parcial(SIS, "soporte_conteos", sp, codigos=["PC.I1", "PC.I4", "PC.I6"])
    ids_proc = {a: [f["id_soporte"] for f in sp if f["codigo_indicador"] == "PC.I1" and f["anio"] == a] for a in (2024, 2025)}
    ids_virt = {a: [f["id_soporte"] for f in sp if f["codigo_indicador"] == "PC.I6" and f["anio"] == a] for a in (2024, 2025)}
    valores, resultados = [], []
    ref = {"PC.I1": ("G24/G25, ANTE, POTC, PI26 consecutivos 530 y 531, comunicados 88/90 iniciativas",),
           }
    for ind in ("PC.I1", "PC.I4", "PC.I6"):
        for esc in C.escenarios(ind):
            exploratorio = esc["params"]["pc_universo_procesos"] == "exploratorio_pp_pdm_pot_rc"
            for anio in (2024, 2025):
                fc = C.fecha_corte_registro(anio, "Anual", None, esc["params"])
                base = dict(codigo_indicador=ind, categoria_disponibilidad="Parcial", script=SCRIPT, tipo_valor="observado",
                            fecha_corte=fc)

                def val(idv, valor, ids, clave, pag, obs, estado=None):
                    kw = dict(estado_informacion=estado or ("Pendiente" if valor is not None else "No disponible"))
                    if valor is not None:
                        kw.update(fuente=D[clave][0], documento_ruta_local=D[clave][1], enlace=C.url_de_ruta(D[clave][1]),
                                  pagina_tabla=pag, n_elementos_soporte=len(ids), ids_soporte=",".join(ids))
                    valores.append(C.fila_valor(SIS, idv, anio, valor, "Conteo", esc, nombre_variable=NOMBRES[idv],
                                                **base, **kw, observaciones=obs))
                sin_inv = "Sin inventario oficial de procesos participativos (pc_universo_procesos=inventario_oficial): N/D"
                if ind in ("PC.I1", "PC.I4"):
                    if exploratorio:
                        k = {2024: 3, 2025: 5}[anio]
                        assert len(ids_proc[anio]) == k
                        val(f"{ind}.V2", k, ids_proc[anio], "G25" if anio == 2025 else "ANTE",
                            "Ver soporte_conteos (una fila por proceso, con documento y página)", OBS_UNIV +
                            ("; PC.I4.V2 = universo de PC.I1" if ind == "PC.I4" else ""))
                    else:
                        val(f"{ind}.V2", None, [], "G25", "", sin_inv)
                    why = ("No se verificó proceso por proceso la incorporación del aporte en una decisión formal; el vínculo "
                           "aporte-decisión solo está documentado de forma narrativa (PP y PDM)" if ind == "PC.I1" else
                           "No se verificó documento por documento que el resultado publicado identifique el proceso, sintetice "
                           "aportes y comunique decisiones (lista de chequeo sin diligenciar)")
                    val(f"{ind}.V1", None, [], "G25", "", why + ". " + OBS_UNIV)
                else:
                    if exploratorio:
                        val("PC.I6.V1", 1, ids_virt[anio], "G25" if anio == 2025 else "QA24",
                            "p. 92" if anio == 2025 else "pp. 6-8",
                            "Cota inferior: lista parcial con 1 actividad con canal virtual para intervenir; transmisiones "
                            "(YouTube, Facebook) solo para observar no se cuentan. No es un total oficial")
                    else:
                        val("PC.I6.V1", None, [], "G25", "", sin_inv)
                    val("PC.I6.V2", None, [], "G25", "", "No existe registro único de actividades participativas realizadas; "
                        "el universo no se totaliza. N/D")
                res = {"puntaje": None, "semaforo": "Gris", "disponibilidad": "N/D",
                       "motivo": "Indicador no viable por ahora: sin universo oficial de " + (
                           "procesos participativos cerrados" if ind != "PC.I6" else "actividades participativas") +
                       "; solo se construyen variables"}
                resultados.append(C.fila_resultado(ind, anio, esc, res, None, sistema=C._sis_nombre(SIS), nombre=NOM_IND[ind],
                    tipo="Principal" if ind != "PC.I6" else "Complementario", fecha_corte_comun=fc, script=SCRIPT,
                    no_calculado_por_etapa=True, observaciones="No calculado en esta etapa (No viable por ahora)"))
    cod = ["PC.I1", "PC.I4", "PC.I6"]
    C.escribir_parcial(SIS, "valores_variables", valores, codigos=cod)
    C.escribir_parcial(SIS, "resultados_indicadores", resultados, codigos=cod)
    print(f"soporte={len(sp)} valores={len(valores)} resultados={len(resultados)}")
    for v in valores:
        if v["es_defecto"]:
            print(v["id_variable"], v["anio"], v["valor"], v["estado_informacion"])


if __name__ == "__main__":
    main()
