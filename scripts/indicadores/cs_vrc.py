#!/usr/bin/env python3
"""CS.VRC: veedurías en rendición de cuentas = 100 * EV / ET (EV ⊆ ET; EV > ET -> Revisar; denominador 0 -> NE).

ET = eventos municipales de rendición de cuentas del universo (parámetro vrc_universo) asignados al año según vrc_asignacion_anio
     y cuya evidencia documental esté publicada dentro de la ventana (ventana_documental_cs).
EV = eventos de ET con presencia de veedurías acreditada (vrc_acreditacion_presencia, vrc_evento_sin_mencion).
Regla de la Matriz: "si hay casos con presencia desconocida, VRC = N/D" (EV = N/D).

Insumo: tabla de clasificación de eventos (una fila por evento) escrita en parciales/CS/clasificacion_vrc.csv. Cada escenario se
recalcula SOLO desde las banderas de esa tabla (se verifica leyendo el CSV de nuevo). Todo PROVISIONAL.
Ejecutar sin -I desde la raíz: python3 scripts/indicadores/cs_vrc.py
"""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import cs_segunda_lectura as SL

COD = "CS.VRC"
SCRIPT = "scripts/indicadores/cs_vrc.py"
AN = (2024, 2025)
D = dt.date
B2 = "data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/"
B3 = "data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/"
B1 = "data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/"
R_PI26 = (B2 + "informe-de-gestion-alcaldia-de-envigado/"
          "evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx")
R_QA24 = B3 + "audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf"
R_QA100 = B2 + "preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf"
R_ACS24A = B2 + "primera-rendicion-de-cuentas-de-la-secretaria-de-salud/20240321primerarendicincuentassalud-1.pdf"
R_ACS24B = B3 + "acta-audiencia-publica-de-rendicion-de-cuentas-2024/2024actarendicioncuentassecretariasalud02docx.pdf"
R_ACS25 = B3 + "audiencia-publica-de-rendicion-de-cuentas-202501-secretaria/2025actarendicioncuentassecretariasalud01final.pdf"
R_QA25 = B2 + "rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf"
R_CMD24A = B3 + "rendicion-de-cuentas-del-comite-municipal-de-discapacidad/rendicion-de-cuentas-semestre-1.pdf"
R_CMD24B_CARTA = B3 + "rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/respuesta_rendicion_cuentas_ricardo_gaviria.pdf"
R_CMD25A = B3 + "rendicion-de-cuentas-del-comite-municipal-de-discapacidad-961114/memoriadeleventorddiscapacidad20251.pdf"
R_CMD25B = B3 + "rendicion-de-cuentas-comite-municipal-de-discapacidad-571162/memoriadeleventordcdiscapacidadii.pdf"
R_VCA24_JUS = B2 + "justificacion-de-no-elaboracion-de-plan-de-mejoramiento/justificacionnoplandemejoramientovca2024.pdf"
R_VCA25 = B2 + "rendicion-de-cuentas-de-la-politica-publica-de-victimas/infrendcuentasvca2025.pdf"
R_ETN_JUS = B2 + "rendicion-de-cuentas-de-la-poblacion-afro-indigena-y/justificacionrendiciondecuentasetnicos.pdf"
R_PP_JUS = B1 + "rendicion-de-cuentas-presupuesto-participativo-2024/justificacionnoplandemejoramiento.pdf"
R_JUV = B2 + "secretaria-de-educacion-426550/rendicion-cuentas-ppjuventud-sec-educacion_2025.pdf"

# Marcas para quitar nombres de personas de rutas y enlaces (cartas de respuesta del CMD)
OMITIR_NOMBRE = {R_CMD24B_CARTA}

OPT_UNIVERSO = ["alcalde", "institucional_pi_545_546", "alcalde_mas_sectoriales_con_acta", "amplio_todo_documento"]
COL_UNIVERSO = {"alcalde": "u_alcalde", "institucional_pi_545_546": "u_institucional_545_546",
                "alcalde_mas_sectoriales_con_acta": "u_alcalde_mas_sectoriales_con_acta",
                "amplio_todo_documento": "u_amplio_todo_documento"}
COL_PRESENCIA = {"nd_siempre": "presencia_nd_siempre", "cero_si_documento_integral": "presencia_cero_si_integral",
                 "cero_siempre": "presencia_cero_siempre"}

B_INV = "fecha del inventario (no confiable por sí sola) + orden de carpeta content/files"
EV = []  # eventos clasificados (una fila por evento)


def ev(id_, anio_ev, anio_vig, fecha, fecha_nota, evento, es546, es545, sect, confirmado, tipo_doc, integral, mencion,
       doc, ruta, pagina, cita, fpub, base_fpub, criterio, dudoso=0, nota="", compl="", enlace=None):
    assert mencion in ("sí", "no", "desconocida")
    assert not (mencion == "no" and not integral), id_       # 'no' exige documento integral
    assert not (mencion == "desconocida" and integral), id_  # integral sin mención = 'no'
    EV.append(dict(id=id_, anio_evento=anio_ev, anio_vigencia=anio_vig, fecha_evento=fecha, fecha_evento_nota=fecha_nota,
                   evento=evento, es_546_alcalde=int(es546), es_545_veedurias=int(es545),
                   sectorial_con_acta_o_memoria=int(sect), evento_confirmado=int(confirmado), tipo_documento=tipo_doc,
                   documento_integral=int(integral), mencion_veedurias=mencion, documento=doc, documentos_complementarios=compl,
                   ruta_local=ruta, pagina=pagina, enlace=enlace, cita=cita, fecha_pub_evidencia=fpub,
                   base_fecha_pub=base_fpub, criterio=criterio, dudoso=int(dudoso), nota_dudoso=nota))


# ------------------------------------------------------------------ 2024
ev("CS-VRC-2024-001", 2024, 2024, D(2024, 3, 21), "", "Rendición de cuentas Secretaría de Salud 2024-01",
   0, 0, 1, 1, "acta_informe_de_rendicion", 1, "sí", "ACS24a: informe-acta de la primera rendición 2024 de la Secretaría de Salud",
   R_ACS24A, "pp. 1 y 17",
   "Asistentes: Líderes sociales y comunitarios [...] integrantes de las veedurías en salud, integrantes del COPACO, Juntas de Acción Comunal [...] y comunidad en general.",
   D(2024, 5, 8), B_INV + " (carpeta 002317)",
   "Sectorial con acta/informe publicado; el documento integral describe a los asistentes por categoría y nombra las veedurías en salud (pp. 1 y 17). Mención textual = presencia (mencion_en_documento). Listas de asistencia anexas no publicadas.")
ev("CS-VRC-2024-002", 2024, 2024, D(2024, 4, 10), "", "Rendición de cuentas del Alcalde: primeros 100 días",
   0, 0, 0, 1, "cuadro_preguntas_respuestas", 0, "desconocida", "QA100: preguntas y respuestas de la rendición de los 100 primeros días",
   R_QA100, "p. 1",
   "Preguntas y Respuestas de la Rendición de Cuentas 100 primeros días [...] rendición de cuentas que se llevó a cabo el pasado 10 de abril [...] y que fueron respondidas durante la transmisión.",
   D(2024, 4, 30), B_INV + " (carpeta 002312)",
   "Solo universo amplio: no es el consecutivo 546 (la observación de la UE del 546 solo cita la audiencia del 26-nov-2024) ni tiene acta o memoria. Documento = cuadro de preguntas, sin asistentes: sin mención, presencia desconocida.",
   1, "Rendición del Alcalde fuera del consecutivo 546; podría argumentarse que es evento del Alcalde (universo 'alcalde'). No se incluyó.",
   compl="JUS: justificación de no plan de mejoramiento (10-abr-2024)")
ev("CS-VRC-2024-003", 2024, 2024, D(2024, 7, 30), "", "Rendición de cuentas Comité Municipal de Discapacidad, semestre 1 de 2024",
   0, 0, 0, 1, "presentacion_y_cartas", 0, "desconocida",
   "CMD S1-2024: presentación del evento (carpeta con dos cartas de respuesta y justificación)",
   R_CMD24A, "p. 2",
   "Rendición de Cuentas Comité Municipal de Discapacidad Semestre 1 2024. Espacio de socialización de las gestiones realizadas por esta instancia y los logros más significativos en este primer semestre.",
   D(2024, 7, 22), B_INV + " (carpetas 002459-002462; las cartas llevan fecha 2-ago-2024, posterior a la del inventario)",
   "Sin acta ni memoria del evento: solo presentación y cartas. Ningún documento menciona veedurías: presencia desconocida. Fuera del universo alcalde+sectoriales por no tener acta/memoria publicada.",
   1, "Si se aceptara la presentación como 'memoria publicada' (como se hizo con el CMD de 2025), entraría al universo alcalde+sectoriales.",
   compl="Dos cartas de respuesta (nombre de persona omitido) y justificación de no plan de mejoramiento")
ev("CS-VRC-2024-004", 2024, 2024, D(2024, 11, 26), "PI26 dice 26-nov-2024; la encuesta de la jornada está fechada 20-nov-2024",
   "Audiencia pública de rendición de cuentas del Alcalde 2024 (consecutivo 546)",
   1, 0, 0, 1, "cuadro_preguntas_respuestas", 0, "desconocida",
   "QA24: preguntas y recomendaciones de la audiencia del Alcalde 2024",
   R_QA24, "p. 1",
   "Ejercicio de Rendición de Cuentas 2024 [...] se consolidan las preguntas, comentarios, sugerencias y/ felicitaciones en el siguiente cuadro (sin lista de asistentes ni mención de veedurías).",
   D(2024, 12, 12), B_INV + " (carpeta 002595; copia duplicada del 12-dic-2024)",
   "Evento del consecutivo 546. Documentos publicados: cuadro de preguntas, encuesta de jornada y plan de mejoramiento; ninguno es acta o memoria con asistentes ni menciona veedurías: presencia desconocida.",
   1, "El informe digital de rendición del Alcalde 2024 (54 pp.) no menciona veedurías y no lista asistentes; se trató como documento no integral.",
   compl="Encuesta de la jornada (xlsx), plan de mejoramiento, informe digital de rendición 2024")
ev("CS-VRC-2024-005", 2024, 2024, D(2024, 11, 26), "El acta dice 26-nov-2024; la invitación y la justificación dicen 27-nov-2024",
   "Rendición de cuentas Secretaría de Salud 2024-02", 0, 0, 1, 1, "acta", 1, "sí",
   "ACS24b: acta de la audiencia pública de rendición de cuentas 2024-02 de la Secretaría de Salud",
   R_ACS24B, "pp. 1, 8-9 y 12",
   "Continúan sus presentaciones los integrantes de las veedurías [ciudadano/a], [ciudadano/a] y [ciudadano/a] [...] Quienes presentaron el objetivo de la veeduría, los hallazgos y las recomendaciones.",
   D(2024, 12, 27), B_INV + " (carpeta 002602)",
   "Sectorial con acta publicada. El acta documenta la presencia de las veedurías (asistentes, saludo especial y presentación de sus informes). Listas de asistencia anexas no publicadas.",
   0, "", compl="Justificación de no plan de mejoramiento; informe y presentación de la rendición")
ev("CS-VRC-2024-006", 2024, 2024, D(2024, 12, 3), "La justificación dice 3-dic-2024; la solicitud a la mesa de víctimas anunciaba el 5-dic-2024",
   "Rendición de cuentas de la política pública de víctimas 2024", 0, 0, 0, 1, "presentacion_y_justificacion", 0, "desconocida",
   "VCA-2024: justificación de no plan de mejoramiento (la presentación del evento no menciona veedurías)",
   R_VCA24_JUS, "p. 1",
   "Rendición de Cuentas de la política pública de víctimas realizada el 3 de diciembre de 2024, y [las preguntas] fueron respondidas durante el encuentro.",
   D(2024, 12, 10), B_INV + " (carpeta 002594)",
   "Sin acta ni memoria: presentación, evidencia de la consulta a la mesa de víctimas y justificación. No menciona veedurías: presencia desconocida.",
   1, "Fecha del evento contradictoria (3 o 5 de diciembre); ambas caen en 2024.",
   compl="Presentación del programa de víctimas y evidencia de la consulta a la mesa de víctimas")
ev("CS-VRC-2024-007", 2024, 2024, D(2024, 12, 12), "",
   "Evento de rendición de cuentas por parte de las veedurías ciudadanas (consecutivo 545)", 0, 1, 0, 1,
   "observacion_plan_indicativo", 0, "sí", "PI26: Plan Indicativo de Productos, evaluación definitiva 2T-2026, consecutivo 545",
   R_PI26, "Hoja Plan-Indicativo-Productos, fila 548, col. EG (observaciones UE 2024)",
   "Se realiza el evento de rendición de cuentas por parte de las veedurías ciudadanas el 12 de dciembre [sic] de 2024, en las instalaciones de la Casa de la Cultura de Envigado.",
   D(2026, 7, 30), "fecha en el nombre del archivo (30-jul-2026); no se localizó una versión anterior del Plan Indicativo en el repositorio",
   "Producto 5.1.2.1.3 'Eventos de rendición pública de cuentas con veedurías ciudadanas': por definición del producto las veedurías participan (mención textual en la observación de la UE). Única evidencia: PI26, publicado después de la ventana del 31-mar-2025.",
   1, "Evento autodeclarado sin acta, memoria ni lista de asistencia; la evidencia solo está en un documento publicado fuera de la ventana. No hay documento propio del evento.")
ev("CS-VRC-2024-008", 2024, 2024, D(2024, 12, 17), "", "Rendición de cuentas Comité Municipal de Discapacidad, semestre 2 de 2024",
   0, 0, 0, 1, "presentacion_y_cartas", 0, "sí",
   "CMD S2-2024: carta de respuesta a una veeduría (nombre de persona omitido); la mención está en la carta, no en un acta",
   R_CMD24B_CARTA, "p. 1",
   "Respuesta a preguntas realizadas por [ciudadano/a] el día 17 de diciembre de 2024 en el marco de la Rendición de Cuentas del Comité Municipal de Discapacidad. Veeduría 284 en Discapacidad.",
   D(2025, 1, 9), B_INV + " (carpeta 002627, posterior a la 002602 del 27-dic-2024)",
   "Sin acta ni memoria: presentación, cartas y justificación. Una carta de respuesta identifica a un integrante de la Veeduría 284 que formuló preguntas en el evento del 17-dic-2024: mención textual de presencia.",
   1, "La mención proviene de una carta de respuesta, no de un acta; con acreditación por lista de asistencia sería desconocida.",
   compl="Presentación iirendiciondecuentasdiscapacidad.pdf, otra carta de respuesta y justificación (nombres omitidos)")
# ------------------------------------------------------------------ 2025
ev("CS-VRC-2025-001", 2025, 2024, D(2025, 1, 12), "La justificación dice 12-ene-2025; una noticia sugiere 12-feb-2025 (no verificada)",
   "Rendición de cuentas de presupuesto participativo (proyectos ejecutados en 2024)", 0, 0, 0, 1,
   "presentaciones_y_justificacion", 0, "desconocida",
   "PP-2024: justificación de no plan de mejoramiento de la rendición de presupuesto participativo (paquete de presentaciones por secretaría)",
   R_PP_JUS, "p. 1",
   "Rendición de Cuentas de Presupuesto Participativo de los proyectos ejecutados en el año 2024, realizada el día 12 de enero de 2025, y [las preguntas] fueron respondidas durante la transmisión.",
   D(2025, 2, 28), B_INV + " (carpeta 002666)",
   "Evento que rinde la vigencia 2024 realizado en 2025 (parámetro vrc_asignacion_anio). Sin acta ni memoria. El paquete incluye el plan de acción de una veeduría de PP, que no acredita su asistencia al evento: presencia desconocida.",
   1, "El paquete publica el plan de acción de la Veeduría 118 (p. 2); no es mención de asistencia. Fecha del evento dudosa (enero o febrero de 2025).",
   compl="Presentaciones por secretaría y plan de acción de la Veeduría 118 (carpeta 003232)")
ev("CS-VRC-2025-002", 2025, 2025, D(2025, 3, 19), "", "Rendición de cuentas Secretaría de Salud 2025-01",
   0, 0, 1, 1, "acta", 1, "sí", "ACS25: acta de la audiencia pública de rendición de cuentas 2025-01 de la Secretaría de Salud",
   R_ACS25, "pp. 1, 2 y 8",
   "Asistentes: Líderes sociales y comunitarios [...] integrantes de las veedurías en salud, integrantes del COPACO, y comunidad en general. Se anexan listados de asistencia.",
   D(2025, 4, 3), B_INV + " (carpeta 002690; copia duplicada del 3-mar-2025)",
   "Sectorial con acta publicada. Mención textual de veedurías entre los asistentes y en el saludo (pp. 1, 2 y 8). El archivo duplicado en otra carpeta se cuenta una vez.")
ev("CS-VRC-2025-003", 2025, 2025, D(2025, 7, 18), "", "Rendición de cuentas Comité Municipal de Discapacidad, semestre 1 de 2025",
   0, 0, 1, 1, "memoria_presentacion", 0, "desconocida",
   "CMD S1-2025: memoria del evento (presentación) con informe de gestión, preguntas/respuestas y justificación",
   R_CMD25A, "p. 2",
   "Rendición de cuentas Comité Municipal de Discapacidad 2025-1.",
   D(2025, 7, 14), B_INV + " (carpetas 002864-002865; la carta de respuesta lleva fecha 30-jul-2025)",
   "El archivo se rotula 'memoria del evento' (cuenta como memoria publicada), pero es una presentación sin asistentes: no es documento integral. El informe de gestión menciona veedurías en otra reunión (COPACO), no en el evento: presencia desconocida.",
   1, "Que una presentación rotulada 'memoria' califique como memoria publicada es un juicio; el informe de gestión (p. 20) nombra veedurías en una sesión del COPACO.")
ev("CS-VRC-2025-004", 2025, 2025, D(2025, 10, 30), "", "Rendición de cuentas de la política pública de víctimas 2025",
   0, 0, 1, 1, "informe_del_evento", 1, "no", "VCA-2025: informe de rendición de cuentas inversión en población víctima 2025",
   R_VCA25, "p. 3",
   "Este espacio contó con la presencia de la Mesa de Participación Efectiva de Víctimas-MPEV- e integrantes de la misma.",
   D(2025, 10, 15), B_INV + " (carpeta 002934; la fecha del inventario es anterior al evento)",
   "Informe del evento (convocatoria, desarrollo, asistentes, evaluación) tratado como memoria integral. Describe la presencia de la mesa de víctimas y no menciona veedurías: ausencia documentada.",
   1, "Se trató el informe como documento integral aunque no trae lista ni conteo de asistentes; si se considerara no integral, la presencia sería desconocida.",
   compl="Justificación de no plan de mejoramiento (30-oct-2025)")
ev("CS-VRC-2025-005", 2025, 2025, D(2025, 12, 10), "", "Rendición de cuentas Comité Municipal de Discapacidad, semestre 2 de 2025",
   0, 0, 1, 1, "memoria_presentacion", 0, "desconocida",
   "CMD S2-2025: memoria del evento (presentación) y justificación (no se recibieron preguntas)",
   R_CMD25B, "p. 20",
   "Veeduría PAE Veeduría PIC REPRESENTANTE PERSONAS CON DISCAPACIDAD FÍSICA (diapositiva de composición del Comité).",
   D(2025, 12, 9), B_INV + " (carpeta 002994, anterior a la 003187 de feb-2026)",
   "Memoria del evento = presentación sin asistentes (no integral). Las veedurías PAE y PIC aparecen como integrantes del Comité, no como asistentes: no acredita presencia, presencia desconocida.",
   1, "Las veedurías figuran en la composición del Comité; podría leerse como presencia (si es así, sería 'sí').")
ev("CS-VRC-2025-006", 2025, 2025, None, "Sin fecha en el repositorio; el cuadro de preguntas cita el 10-dic-2025 como 'ayer' (hacia el 11-dic-2025, inferido)",
   "Audiencia pública de rendición de cuentas del Alcalde 2025 (consecutivo 546)", 1, 0, 0, 1,
   "cuadro_preguntas_respuestas", 0, "desconocida", "QA25: preguntas y respuestas de la rendición de cuentas 2025 del Alcalde",
   R_QA25, "p. 1",
   "Ejercicio de Rendición de Cuentas 2025 [...] se consolidan las preguntas, sugerencias y felicitaciones en el siguiente cuadro (sin lista de asistentes ni mención de veedurías).",
   D(2025, 12, 16), "Last-Modified HTTP 16-dic-2025 (notas_CS.md); el inventario dice 2025-09-16, inconsistente con el contenido",
   "Evento del consecutivo 546. Solo se publicó el cuadro de preguntas (3 intervenciones); no hay acta, memoria ni informe con asistentes: presencia desconocida.",
   1, "Fecha del evento no verificada; no hay informe 2025 del Alcalde en el repositorio.")
ev("CS-VRC-2025-007", 2025, 2025, None, "Sin fecha ni descripción del evento", "Evento de rendición de cuentas con veedurías ciudadanas 2025 (consecutivo 545)",
   0, 1, 0, 1, "observacion_plan_indicativo", 0, "sí", "PI26: Plan Indicativo de Productos, evaluación definitiva 2T-2026, consecutivo 545",
   R_PI26, "Hoja Plan-Indicativo-Productos, fila 548, cols. AM, BN (avance 2025 = 1) y EH",
   "Indicador: Eventos de rendición pública de cuentas con veedurías ciudadanas realizados. Avance 2025 = 1. Observación UE 2025: Esta actividad se realiza con la capacidad instalada.",
   D(2026, 7, 30), "fecha en el nombre del archivo (30-jul-2026); sin versión anterior en el repositorio",
   "Solo consta la cifra autodeclarada (1) del producto 'con veedurías ciudadanas'; presencia = por definición del producto. Sin descripción, fecha ni documento del evento; evidencia publicada después de la ventana del 31-mar-2026.",
   1, "Presencia inferida de la definición del producto; no hay documento del evento. Si se exigiera mención en un documento del evento, sería desconocida.")
ev("CS-VRC-2025-008", 2025, 2025, None, "Sin fecha ni evento verificado", "Posible rendición de la política pública de juventud (Secretaría de Educación)",
   0, 0, 0, 0, "presentacion_sin_evento_confirmado", 0, "desconocida",
   "Presentación 'rendición de cuentas PP juventud' de la Secretaría de Educación; no se confirmó que sea un evento municipal",
   R_JUV, "p. 2",
   "Línea 5 Promoción de educación con calidad y pertinencia para los y las jóvenes (diapositiva de la presentación).",
   D(2025, 3, 1), "solo orden de carpeta 002773 (entre 002755 de ene-2025 y 002830 de abr-2025); el inventario dice 2021-09-14, inválido",
   "No se halló fecha, convocatoria, acta ni noticia del evento; las fuentes previas de CS no lo listaron. Se deja en la tabla fuera de todos los universos (evento_confirmado = 0).",
   1, "Posible evento adicional de 2025 sin verificar; si se confirmara, entraría al universo amplio (y ET 2025 amplio pasaría de 7 a 8).")
# ------------------------------------------------------------------ 2026 (vigencia 2025)
ev("CS-VRC-2026-001", 2026, 2025, D(2026, 9, 3), "", "Rendición de cuentas de la política pública de grupos étnicos (vigencia 2025)",
   0, 0, 0, 1, "presentacion_y_justificacion", 0, "desconocida",
   "ETN: justificación de no plan de mejoramiento de la rendición de grupos étnicos",
   R_ETN_JUS, "p. 1",
   "Rendición de Cuentas de la Política Pública de Grupos Étnicos correspondiente a la vigencia 2025, realizada el día 03 de septiembre de 2026.",
   D(2026, 9, 2), "inventario 2-sep-2026 (anterior al evento) + carpeta 003671",
   "Evento realizado en 2026 que rinde la vigencia 2025: cuenta en 2025 solo con vrc_asignacion_anio=vigencia_rendida y su evidencia queda fuera de la ventana del 31-mar-2026. Presentación y justificación; sin mención de veedurías: presencia desconocida.",
   1, "Fecha del documento anterior a la del evento; asignación de año depende del parámetro.")


# ------------------------------------------------------------------ derivación de banderas
def enlace_de(ruta):
    u = C.url_de_ruta(ruta)
    if ruta in OMITIR_NOMBRE:
        u = u.rsplit("/", 1)[0] + "/"
    assert u, f"sin URL para {ruta}"
    return u


def derivar(e):
    f = dict(e)
    ruta = f["ruta_local"]
    f["enlace"] = enlace_de(ruta)
    if ruta in OMITIR_NOMBRE:
        f["ruta_local"] = ruta.rsplit("/", 1)[0] + "/[archivo con nombre de persona omitido]"
    c = f["evento_confirmado"]
    f["u_alcalde"] = int(f["es_546_alcalde"] and c)
    f["u_institucional_545_546"] = int((f["es_545_veedurias"] or f["es_546_alcalde"]) and c)
    f["u_alcalde_mas_sectoriales_con_acta"] = int((f["es_546_alcalde"] or f["sectorial_con_acta_o_memoria"]) and c)
    f["u_amplio_todo_documento"] = int(c)
    m = f["mencion_veedurias"]
    f["mencion_si"], f["mencion_no"], f["mencion_desconocida"] = int(m == "sí"), int(m == "no"), int(m == "desconocida")
    f["presencia_nd_siempre"] = 1 if m == "sí" else "ND"
    f["presencia_cero_si_integral"] = 1 if m == "sí" else (0 if m == "no" else "ND")
    f["presencia_cero_siempre"] = 1 if m == "sí" else 0
    f["presencia_lista_asistencia"] = "ND"       # listas de asistencia anexas no publicadas
    f["evidencia_en_ventana_fecha_evento"] = int(f["fecha_pub_evidencia"] <= D(f["anio_evento"] + 1, 3, 31))
    f["evidencia_en_ventana_vigencia"] = int(f["fecha_pub_evidencia"] <= D(f["anio_vigencia"] + 1, 3, 31))
    f["clasificacion"] = (f"evento_del_universo: alcalde={f['u_alcalde']}, institucional_545_546={f['u_institucional_545_546']}, "
                          f"alcalde_mas_sectoriales_con_acta={f['u_alcalde_mas_sectoriales_con_acta']}, amplio={f['u_amplio_todo_documento']} | "
                          f"veedurías: {m} | documento {'integral' if f['documento_integral'] else 'no integral'} ({f['tipo_documento']})")
    return f


COLS = ["id", "anio_evento", "anio_vigencia", "fecha_evento", "fecha_evento_nota", "evento", "documento", "pagina", "enlace", "cita",
        "clasificacion", "criterio", "tipo_documento", "documento_integral", "mencion_veedurias",
        "es_546_alcalde", "es_545_veedurias", "sectorial_con_acta_o_memoria", "evento_confirmado",
        "u_alcalde", "u_institucional_545_546", "u_alcalde_mas_sectoriales_con_acta", "u_amplio_todo_documento",
        "mencion_si", "mencion_no", "mencion_desconocida",
        "presencia_nd_siempre", "presencia_cero_si_integral", "presencia_cero_siempre", "presencia_lista_asistencia",
        "fecha_pub_evidencia", "base_fecha_pub", "evidencia_en_ventana_fecha_evento", "evidencia_en_ventana_vigencia",
        "documentos_complementarios", "ruta_local", "dudoso", "nota_dudoso"]


# ------------------------------------------------------------------ cálculo desde la tabla (filas como texto)
def a_texto(filas):
    return [{c: C._fmt(f.get(c)) for c in COLS} for f in filas]


def evaluar(filas, p, anio):
    """Devuelve dict(et_ids, ev_ids, desconocidos, ET, EV) para el escenario p y el año, SOLO con banderas de la tabla."""
    col_anio = "anio_evento" if p["vrc_asignacion_anio"] == "fecha_evento" else "anio_vigencia"
    col_vent = "evidencia_en_ventana_fecha_evento" if p["vrc_asignacion_anio"] == "fecha_evento" else "evidencia_en_ventana_vigencia"
    et = []
    for f in filas:
        if int(f[col_anio]) != anio or f[COL_UNIVERSO[p["vrc_universo"]]] != "1":
            continue
        if p["ventana_documental_cs"] == "guia_31mar_siguiente" and f[col_vent] != "1":
            continue
        et.append(f)
    col_pres = "presencia_lista_asistencia" if p["vrc_acreditacion_presencia"] == "lista_asistencia" \
        else COL_PRESENCIA[p["vrc_evento_sin_mencion"]]
    ev_ids, desc = [], []
    for f in et:
        v = f[col_pres]
        if v == "1":
            ev_ids.append(f["id"])
        elif v == "ND":
            desc.append(f["id"])
    return dict(et_ids=[f["id"] for f in et], ev_ids=ev_ids, desconocidos=desc, ET=len(et), EV=len(ev_ids))


def entradas(r):
    if r["ET"] == 0:
        return [0, 0]                      # sin universo: NE
    if r["desconocidos"]:
        return [None, r["ET"]]             # presencia desconocida: EV = N/D -> VRC = N/D
    return [r["EV"], r["ET"]]


def main():
    filas = [derivar(e) for e in EV]
    assert len({f["id"] for f in filas}) == len(filas)
    # --- validación de citas / textos libres
    for f in filas:
        for c in ("cita", "criterio", "nota_dudoso", "fecha_evento_nota", "clasificacion", "documento"):
            hall = C.validar_sin_datos_personales(f.get(c, ""))
            assert not hall, (f["id"], c, hall)
        assert len(f["cita"]) <= 200, (f["id"], len(f["cita"]))
    C.escribir_parcial("CS", "clasificacion_vrc", [{c: f.get(c) for c in COLS} for f in filas])
    ruta_cl = C.ruta_parcial("CS", "clasificacion_vrc")
    tabla = C.leer_csv(ruta_cl)                       # el cálculo parte de lo escrito
    assert len(tabla) == len(filas)

    escs = C.escenarios(COD) + [C.escenario_nombrado(
        "control_institucional_545_546",
        {"vrc_universo": "institucional_pi_545_546", "vrc_evento_sin_mencion": "cero_siempre",
         "ventana_documental_cs": "hasta_fecha_consulta"}, COD)]
    defecto = C.escenario_defecto(COD)["params"]
    url_et = C.normalizar_url(C.url_de_ruta(R_PI26))
    soporte, valores, resultados = [], [], []
    cuenta_def = {}  # el soporte describe cada elemento con el conteo del escenario por defecto
    for a in AN:
        r0 = evaluar(tabla, defecto, a)
        for i in r0["et_ids"]:
            cuenta_def[i] = (a, 1 if i in r0["ev_ids"] else 0)
    for f in filas:
        i = f["id"]
        dentro = i in cuenta_def
        soporte.append(dict(
            id_soporte=i, codigo_indicador=COD, id_variable="CS.VRC.ET;CS.VRC.EV", anio=f["anio_evento"], escenario_id="defecto",
            elemento_tipo="evento", elemento_id=i, descripcion=f["evento"], fecha_elemento=f["fecha_evento"],
            documento=f["documento"], ruta_local=f["ruta_local"], enlace=f["enlace"], pagina=f["pagina"],
            fecha_publicacion_doc=f["fecha_pub_evidencia"],
            banderas=(f"mencion={f['mencion_veedurias']};integral={f['documento_integral']};u_alcalde={f['u_alcalde']};"
                      f"u_inst={f['u_institucional_545_546']};u_sect={f['u_alcalde_mas_sectoriales_con_acta']};u_amplio={f['u_amplio_todo_documento']};"
                      f"en_ventana={f['evidencia_en_ventana_fecha_evento']};dudoso={f['dudoso']}"),
            clasificacion=f["mencion_veedurias"], justificacion=f["criterio"],
            criterio="Evento contado en ET si está en el universo del parámetro vrc_universo, en el año según vrc_asignacion_anio y con evidencia dentro de la ventana; "
                     "en EV si la presencia de veedurías se acredita según vrc_acreditacion_presencia y vrc_evento_sin_mencion (ver clasificacion_vrc.csv)",
            lector="ejecutor_CS", cuenta_en_defecto=int(dentro), incluido="sí" if dentro else "no"))

    resumen = []
    for esc in escs:
        p = esc["params"]
        for a in AN:
            r = evaluar(tabla, p, a)
            ent = entradas(r)
            res = C.calcular(COD, ent, {"periodo": "Anual", "territorio": "Municipio"})
            if r["ET"] > 0 and r["desconocidos"]:
                res = dict(res, motivo="Presencia de veedurías desconocida en " + ", ".join(r["desconocidos"])
                           + " (regla de la Matriz: con presencia desconocida, VRC = N/D)")
            corte = C.fecha_corte_registro(a, "Anual", D(a, 12, 31), p)
            obs_p = (f"Escenario {esc['escenario_id']}. ET = {r['ET']} ({', '.join(r['et_ids']) or 'ninguno'}); "
                     f"EV = {'N/D' if r['desconocidos'] else r['EV']}"
                     + (f"; presencia desconocida: {', '.join(r['desconocidos'])}" if r["desconocidos"] else "")
                     + ". Clasificación y banderas en parciales/CS/clasificacion_vrc.csv. PROVISIONAL.")
            com = dict(tipo_valor="Construido", fecha_corte=corte, fecha_corte_real=D(a, 12, 31),
                       fecha_publicacion=None, script=SCRIPT, categoria_disponibilidad="Parcial", horizonte_meta=str(a),
                       documento_ruta_local="outputs/indicadores/parciales/CS/clasificacion_vrc.csv", enlace=url_et,
                       pagina_tabla="Ver ids_soporte (clasificacion_vrc.csv y soporte_conteos.csv); Plan Indicativo consecutivos 545 y 546 en Hoja Plan-Indicativo-Productos, filas 548 y 549",
                       fuente="Documentos de rendición de cuentas de la Alcaldía de Envigado (actas, cuadros de preguntas, memorias, cartas, Plan Indicativo 2T-2026) clasificados evento por evento",
                       observaciones=obs_p)
            valores.append(C.fila_valor("CS", "CS.VRC.EV", a, ent[0] if not (r["ET"] > 0 and r["desconocidos"]) else None, "Conteo", esc=esc,
                nombre_variable="Número de eventos municipales de rendición de cuentas realizados con presencia documentada de veedurías",
                n_elementos_soporte=r["ET"], ids_soporte=",".join(r["ev_ids"]), **com))
            valores.append(C.fila_valor("CS", "CS.VRC.ET", a, r["ET"], "Conteo", esc=esc,
                nombre_variable="Número total de eventos municipales de rendición de cuentas realizados en el periodo",
                n_elementos_soporte=r["ET"], ids_soporte=",".join(r["et_ids"]), **com))
            resultados.append(C.fila_resultado(COD, a, esc, res, ent, nombre="Veedurías en rendición de cuentas (VRC)", tipo="Indicador",
                fecha_corte_comun=corte, script=SCRIPT, observaciones=obs_p))
            resumen.append((esc["escenario_id"], a, r, res))

    C.escribir_parcial("CS", "soporte_conteos", soporte, codigos=[COD])
    C.escribir_parcial("CS", "valores_variables", valores, codigos=[COD])
    C.escribir_parcial("CS", "resultados_indicadores", resultados, codigos=[COD])

    # --- verificación: los resultados reproducidos desde el CSV coinciden con lo escrito
    escritos = {(r["escenario_id"], str(r["anio"])): r for r in C.leer_csv(C.ruta_parcial("CS", "resultados_indicadores")) if r["codigo"] == COD}
    for eid, a, r, res in resumen:
        w = escritos[(eid, str(a))]
        assert w["semaforo"] == res["semaforo"] and w["disponibilidad"] == res["disponibilidad"], (eid, a)

    # --- muestra del 20 % para segunda lectura
    cols_m = COLS
    ruta_m, n_m, escrito = SL.muestra_20(tabla, cols_m, "vrc")
    SL.escribir_leeme()

    print(f"clasificacion_vrc: {len(filas)} filas; muestra_vrc: {n_m} filas ({'escrita' if escrito else 'NO reescrita'})")
    for eid, a, r, res in resumen:
        pj = res["puntaje"]
        ev_txt = "N/D" if (r["ET"] > 0 and r["desconocidos"]) else r["EV"]
        print(f"{eid[:100]:<100} {a} EV={ev_txt} ET={r['ET']} -> {('%.1f' % pj) if pj is not None else res['disponibilidad']} {res['semaforo']}")
    print("dudosos:", [f["id"] for f in filas if f["dudoso"]])
    print(len(soporte), "soporte;", len(valores), "variables;", len(resultados), "resultados")


if __name__ == "__main__":
    main()
