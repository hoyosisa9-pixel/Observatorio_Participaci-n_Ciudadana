#!/usr/bin/env python3
"""Genera outputs/indicadores/entregables/base_indicadores_OCE.xlsx para los profesores del OCE.

Solo presenta lo ya calculado: lee los CSV de outputs/indicadores/, config/supuestos.yaml y
outputs/indicadores/resumen_calculos.md (sección 5). No recalcula ningún valor ni supuesto.
Reutiliza por import las funciones de presentación de scripts/generar_disponibilidad.py (sin modificarlo).

Ejecutar desde la raíz:  python3 scripts/entregables/generar_base_oce.py
"""
import csv
import datetime as dt
import re
import sys
from functools import lru_cache
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts'))
sys.path.insert(0, str(RAIZ / 'scripts' / 'indicadores'))
import generar_disponibilidad as gd  # noqa: E402
from comun import validar_sin_datos_personales  # noqa: E402

DIR = RAIZ / 'outputs' / 'indicadores'
SALIDA = DIR / 'entregables' / 'base_indicadores_OCE.xlsx'
SUPUESTOS = RAIZ / 'config' / 'supuestos.yaml'
RESUMEN = DIR / 'resumen_calculos.md'

# colores del semáforo (se agregan al diccionario de rellenos del módulo importado, solo en memoria)
for k, v in {'Verde': 'D6F0D6', 'Amarillo': 'FFF3BF', 'Rojo': 'F8D4D4', 'Gris': 'D9D9D9'}.items():
    gd.RELLENO[k] = PatternFill('solid', fgColor=v)

ORDEN_SIS = ['Control social', 'Participación ciudadana', 'Cumplimiento del PDM', 'Presupuesto participativo']
ORDEN_COD = ['CS.PCS', 'CS.DSI', 'CS.VRC', 'CS.RPO', 'CS.ICRCS', 'PC.I1', 'PC.I2', 'PC.I3', 'PC.I4', 'PC.I5', 'PC.I6',
             'PDM.CA', 'PDM.CM', 'PDM.CS', 'PDM.CP', 'PDM.CD', 'PDM.ILE', 'PP.EPP']
ORDEN_PER = {'Anual': 0, 'Semestre 1': 1, 'Semestre 2': 2}


def leer(nombre):
    with open(DIR / nombre, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f, delimiter=';'))


# --------------------------------------------------------------------------- parámetros en lenguaje claro
# id: (etiqueta corta, pregunta, {opción: texto corto})
P = {
 'corte_fecha_registro': ('Fecha de corte registrada',
   'Para los datos de cierre de vigencia cuyo corte real es posterior al 31 de diciembre (por ejemplo, 7 de enero de 2026), ¿se registra la fecha de corte real o el 31 de diciembre?',
   {'real': 'Fecha de corte real', 'normalizada_cierre_vigencia': '31 de diciembre del año (el corte real queda en las observaciones)'}),
 'tratamiento_2026': ('Tratamiento de 2026',
   '¿Cómo se trata 2026, que no tiene cierre de vigencia (corte 30 de junio de 2026)?',
   {'excluir': 'No calcular 2026', 'semestre1_provisional': 'Primer semestre de 2026, con corte 30 de junio y rotulado provisional',
    'anual_provisional': 'Año 2026 con corte 30 de junio rotulado anual (solo para sensibilidad)'}),
 'regla_estado_calculo': ('Regla de estado para calcular',
   '¿Se calcula con datos en estado Pendiente (aún sin validar) o solo con datos Validados?',
   {'estricta_matriz': 'Solo con datos validados (todo resulta N/D por ahora)', 'preliminar_pendiente': 'Con datos pendientes, rotulado preliminar'}),
 'zonas_oficiales': ('Zonas oficiales', '¿Cuántas zonas oficiales se consideran: 13 o 12 con comité zonal?',
   {'13': '13 zonas (PDM p. 558; la zona 1 no tiene comité zonal)', '12_con_comite': '12 zonas con comité zonal'}),
 'poblacion_dane': ('Población para las tasas', '¿Qué población se usa para calcular las tasas?',
   {'serie_publicada_por_fuente': 'La tasa tal como la publica la fuente (Medicina Legal, población 248.304)',
    'dane_postcovid_2020_2035': 'Proyección DANE 2020-2035 actualizada tras la pandemia',
    'dane_vp_2018_2042': 'Proyección DANE 2018-2042'}),
 'horizonte_meta_pdm': ('Horizonte de la meta del PDM', '¿Se evalúa 2024-2026 contra la meta final de 2027 o contra metas anuales?',
   {'meta_2027': 'Logro al corte frente a la meta final (2027)', 'metas_anuales': 'Metas anuales oficiales (no existen: resultado N/D)'}),
 'ventana_documental_cs': ('Ventana de publicación (control social)', '¿Hasta cuándo se acepta una publicación como correspondiente a la vigencia?',
   {'guia_31mar_siguiente': 'Hasta el 31 de marzo del año siguiente (RPO: respuestas publicadas al 31 de diciembre)',
    'hasta_fecha_consulta': 'Todo lo publicado hasta la fecha de consulta'}),
 'dsi_lista_metas': ('Lista de metas del DSI', '¿Qué metas forman el denominador del DSI? La Matriz sugiere iniciar con unas 10 metas.',
   {'icpd7': '7 resultados del ICPD-6 (Apéndice 9.1.2)',
    'icpd7_mas3_gobernanza': '7 del ICPD-6 más 3 de gobernanza (10 metas; propuesta del equipo técnico, no de los profesores)',
    'productos_plan_indicativo': 'Todos los productos del plan indicativo con meta (da 100 por construcción; solo control)'}),
 'dsi_fuente_valor': ('Fuente del valor publicado (DSI)', '¿Qué cuenta como valor observado publicado?',
   {'solo_alcaldia': 'Solo publicaciones de la Alcaldía',
    'oficial_misma_definicion': 'Alcaldía o entidad oficial (Medicina Legal, MinSalud, DANE, ICFES) con igual definición y unidad que la meta',
    'oficial_incluye_proxy': 'También cifras oficiales con otra definición (aproximadas)'}),
 'vrc_universo': ('Universo de eventos de rendición (VRC)', '¿Qué es un evento municipal de rendición de cuentas?',
   {'alcalde': 'Solo la audiencia pública del Alcalde (consecutivo 546)',
    'institucional_pi_545_546': 'Audiencia del Alcalde y evento de veedurías (consecutivos 545 y 546)',
    'alcalde_mas_sectoriales_con_acta': 'Audiencia del Alcalde y rendiciones sectoriales con acta o memoria publicada',
    'amplio_todo_documento': 'Todo evento con documento propio (de 1 a 8 por año)'}),
 'vrc_acreditacion_presencia': ('Acreditación de presencia de veedurías (VRC)', '¿La mención de las veedurías en el acta acredita su presencia?',
   {'mencion_en_documento': 'Sí: basta la mención textual en acta, memoria o carta', 'lista_asistencia': 'Solo la lista de asistencia (no publicada: N/D)'}),
 'vrc_evento_sin_mencion': ('Evento sin mención de veedurías (VRC)', 'Si el documento de un evento no menciona veedurías, ¿se cuenta como ausencia o la presencia queda desconocida?',
   {'nd_siempre': 'Presencia desconocida: VRC = N/D (lectura literal de la Matriz)',
    'cero_si_documento_integral': 'Ausencia si el documento es un acta o memoria integral con asistentes; si es un cuadro de preguntas, desconocida',
    'cero_siempre': 'Ausencia siempre'}),
 'vrc_asignacion_anio': ('Año de asignación del evento (VRC)', 'Un evento realizado en el año t+1 que rinde cuentas del año t, ¿a qué año se asigna?',
   {'fecha_evento': 'Al año de realización', 'vigencia_rendida': 'Al año de la vigencia rendida'}),
 'rpo_universo': ('Ejercicios que integran el RPO', '¿Qué cuadros de preguntas integran el indicador?',
   {'alcalde': 'Audiencia pública del Alcalde', 'alcalde_y_100dias': 'Alcalde y rendición de los 100 días (abril de 2024)',
    'todos_los_cuadros': 'Alcalde, 100 días, Salud y Comité de Discapacidad'}),
 'rpo_filtro_pdm': ('Observaciones sobre el PDM (RPO)', '¿Qué es una observación "sobre el PDM"?',
   {'excluir_felicitaciones': 'Todas las preguntas, sugerencias u observaciones; se excluyen felicitaciones y agradecimientos',
    'tematico_pdm': 'Solo las que remiten a un programa, meta o sector del PDM; se excluyen trámites individuales',
    'todas': 'Todas las intervenciones'}),
 'rpo_criterio_fondo': ('Respuesta de fondo (RPO)', '¿Qué es una respuesta de fondo?',
   {'estricto_rubrica': 'Criterio estricto: atiende el asunto con dato, decisión, compromiso o negativa motivada',
    'laxo_texto_respuesta': 'Criterio laxo: cualquier texto de respuesta salvo acuse explícito'}),
 'rpo_registro_incompleto': ('Cuadro publicado incompleto (RPO)', 'Cuando el cuadro publicado es claramente incompleto (Alcalde 2025: 3 intervenciones), ¿se calcula con lo publicado?',
   {'nd': 'No: N/D (un listado incompleto no permite conocer el total)', 'calcular_con_publicado': 'Sí: calcular con lo publicado, rotulado'}),
 'pc_i2_actividades': ('Actividades que cubren una zona (I2)', '¿Qué actividades cuentan como cobertura de una zona?',
   {'asambleas_pp': 'Solo asambleas zonales de presupuesto participativo (máximo 12: zonas 2 a 13)',
    'asambleas_pp_pdm_pot': 'Asambleas de presupuesto participativo, de formulación del PDM (2024) y del POT (2025)'}),
 'pc_i2_evidencia': ('Evidencia de actividad en la zona (I2)', '¿Qué acredita que la actividad se hizo en la zona?',
   {'realizacion_reportada': 'Informe de gestión, anteproyecto o ejecución del plan indicativo',
    'incluye_programacion': 'También cronogramas o programación', 'acta_por_zona': 'Solo acta o asistencia por zona (no publicadas: N/D)'}),
 'pc_i3_unidad': ('Unidad de medida de I3', '¿I3 se mide en acciones formativas o en personas?',
   {'acciones': 'Acciones formativas (cursos, diplomados, seminarios)', 'personas_producto_521': 'Personas programadas frente a ejecutadas (producto 521 del plan indicativo)'}),
 'pc_i3_excluir_servidores': ('Formación exclusiva para servidores (I3)', '¿Se excluye la formación dirigida solo a servidores públicos (la Matriz lo exige)?',
   {'true': 'Excluir', 'false': 'Incluir'}),
 'pc_i3_v1_mayor_v2': ('Ejecución mayor que programación (I3)', 'Si lo ejecutado supera lo programado, ¿se marca "Revisar" o se limita a 100?',
   {'revisar': 'Revisar (fórmula de la Matriz)', 'tope_100': 'Limitar a 100'}),
 'pc_i5_instancias': ('Instancias seleccionadas (I5)', '¿Qué instancias integran el indicador?',
   {'salud_copaco_cmsss': 'COPACO y Consejo Municipal de Seguridad Social en Salud (únicas con programación publicada)',
    'ampliado_cmpc_ctp_comites': 'Ampliado: Consejo Municipal de Participación, CTP y comités zonales (sin calendario: N/D)'}),
 'pc_universo_procesos': ('Universo de procesos participativos cerrados', '¿Qué procesos participativos cerrados forman el universo (solo variables; I1, I4 e I6 no se calculan)?',
   {'exploratorio_pp_pdm_pot_rc': 'Universo exploratorio: presupuesto participativo, PDM, POT y rendición (3 en 2024; 5 en 2025)',
    'inventario_oficial': 'Inventario oficial de Participación (no existe: N/D)'}),
 'pdm_cm_fuente': ('Fuente de mortalidad vial (CM)', '¿Qué fuente de muertes por siniestros viales se usa?',
   {'inmlcf_forensis': 'Medicina Legal, Forensis (muertes por transporte; tasa publicada)',
    'policia_datosgov': 'Policía: homicidio en accidente de tránsito (otra definición)'}),
 'pdm_ca_fuente': ('Espacio público verde observado (CA)', '¿Qué cifra de espacio público verde es comparable con la meta?',
   {'pot_epv_urbano': 'POT 2024, espacio público verde urbano: 5,75 m²/hab. (serie distinta de la línea base 3,88)',
    'pot_epv_total': 'POT 2024, espacio público verde total: 69,41 m²/hab.',
    'estricto_serie_linea_base': 'Solo serie comparable con la línea base (no existe: N/D)'}),
 'pdm_cs_fuente': ('Cobertura del régimen subsidiado (PDM.CS)', '¿Qué fuente se usa para la cobertura del régimen subsidiado (solo variables; PDM.CS no se calcula)?',
   {'metodo_plan': 'Método del plan (BDUA/SisMaster, población objetivo): N/D',
    'minsalud_rs_sobre_pob_total': 'Afiliados del régimen subsidiado sobre población total (aprox. 12 %; otra definición)',
    'minsalud_cobertura_total': 'Cobertura total del sistema de salud (la Matriz lo prohíbe)'}),
 'pdm_ile_a_rango_edad': ('Población de la atención integral inicial (ILE.A)', '¿Qué rango de edad forma el denominador de ILE.A?',
   {'0_6': 'Población DANE de 0 a 6 años', '0_5': 'Población DANE de 0 a 5 años'}),
 'pp_identificacion': ('Identificación de proyectos de PP', '¿Cómo se identifican los proyectos de presupuesto participativo?',
   {'efpc_x_plan_accion': 'Listado de la EFPC cruzado con los proyectos "ZONA n" del plan de acción',
    'solo_plan_accion_zona': 'Solo proyectos con "ZONA n" en el plan de acción'}),
 'pp_incluir_0553': ('Proyecto 0553 (costo del proceso)', '¿Se incluye el proyecto 0553 (costo del proceso de presupuesto participativo)?',
   {'false': 'Excluir', 'true': 'Incluir'}),
 'pp_fuente_calculo': ('Fuente de APP y OPP', '¿De dónde se toman la apropiación y las obligaciones?',
   {'plan_accion_roca': 'Plan de acción del DAP (2025 y 2026)', 'hacienda': 'Ejecución presupuestal de egresos de Hacienda',
    'cuipo': 'CUIPO por BPIN (solo contraste)'}),
 'pp_2024_metodo': ('Método de PP 2024', '¿Se aceptan los cruces aproximados entre el listado de la EFPC y Hacienda en 2024?',
   {'hacienda_correspondencia_completa': '50 de 51 proyectos (35 exactos y 15 aproximados), rotulado aproximado',
    'solo_exactos': 'Solo los 35 proyectos con correspondencia exacta', 'nd': 'EPP 2024 = N/D'}),
 'pp_2026_opp_fuente': ('Fuente de las obligaciones 2026', '¿De dónde se toman las obligaciones (OPP) de 2026 al corte?',
   {'hacienda_30jun': 'Hacienda al 30 de junio de 2026 ($5.702.528.148)', 'plan_accion_xlsx': 'Plan de acción 2T-2026, hoja ROCA ($5.702.528.148; coincide con Hacienda)'}),
 'pp_cep_regla_agregacion': ('Regla de entrega cumplida (CEP)', '¿Qué regla se usa para el proxy de entrega cumplida (solo variable; CEP no se calcula)?',
   {'todas_filas_dic_100': 'Todas las filas de diciembre con ejecución física neta de 100 % (2025: 49)', 'solo_31dic': 'Solo filas del 31 de diciembre (2025: 45)',
    'ultima_fila_por_indicador': 'Última fila por indicador (2025: 50)'}),
 'pp_cma_aplicabilidad': ('Aplicabilidad de la meta de atención (CMA)', '¿A qué proyectos aplica la meta de atención?',
   {'pendiente': 'Pendiente para todos (tipología sugerida solo en observaciones)', 'tipologia_sugerida': 'Infraestructura y medios = No (con motivo); servicios a personas = Sí'}),
 'pp_zona1_tratamiento': ('Tratamiento de la Zona 1', '¿La Zona 1 es "No evaluable" (NE) o "No disponible" (N/D) en presupuesto participativo?',
   {'NE': 'No evaluable: APP = 0 y ningún proyecto de PP se asigna a la Zona 1', 'N/D': 'Dato no disponible: no se puede afirmar que APP sea cero hasta confirmarlo'}),
}
ESPECIALES = {
 'control_institucional_545_546': 'Control: universo institucional (consecutivos 545 y 546) en todos los componentes, evento sin mención = ausencia, publicaciones hasta la fecha de consulta',
 'envolvente_componentes_min': 'Rango de sensibilidad (mínimo): promedio de los puntajes mínimos de cada componente entre sus escenarios; no es un valor oficial',
 'envolvente_componentes_max': 'Rango de sensibilidad (máximo): promedio de los puntajes máximos de cada componente entre sus escenarios; no es un valor oficial',
 'extremos_evaluados_min': 'Extremo mínimo entre los escenarios de un solo cambio evaluados; no es cota de todas las combinaciones',
 'extremos_evaluados_max': 'Extremo máximo entre los escenarios de un solo cambio evaluados; no es cota de todas las combinaciones',
}
TEMA = lambda p: ('General' if p in ('corte_fecha_registro', 'tratamiento_2026', 'regla_estado_calculo', 'zonas_oficiales', 'poblacion_dane', 'horizonte_meta_pdm')
                  else 'Control social' if p.startswith(('ventana', 'dsi', 'vrc', 'rpo')) else 'Participación ciudadana' if p.startswith('pc_')
                  else 'Cumplimiento del PDM' if p.startswith('pdm_') else 'Presupuesto participativo')

# --------------------------------------------------------------------------- documentos y rutas
DOCS = {
 'PLAN-ZONA-2.pdf': 'Planes zonales de desarrollo PLAN-ZONA-2 a PLAN-ZONA-13 (Escuela de Formación para la Participación Ciudadana, EFPC)',
 'Proyectos_PP_2024.pdf': 'Listado de proyectos de presupuesto participativo 2024 (Escuela de Formación para la Participación Ciudadana, EFPC)',
 'resumeniniciativascomunitariaspp2024.pdf': 'Comunicado: 88 iniciativas ciudadanas postuladas al presupuesto participativo 2024 (resumen)',
 'justificacionnoplandemejoramiento.pdf': 'Justificación de no plan de mejoramiento de la rendición de presupuesto participativo 2024',
 'pps190ppps20251231mu000000005266.xlsx': 'Seguimiento al Plan de Acción de la Política de Participación Social en Salud (PPSS) 2025',
 'datosgov_4f7r-epif_CUIPO_Ejecucion_Gastos_Envigado.csv': 'CUIPO, ejecución de gastos de Envigado (datos.gov.co)',
 'ejecucion-presupuestal-egresos-al-31122024.pdf': 'Ejecución presupuestal de egresos al 31 de diciembre de 2024 (Secretaría de Hacienda)',
 '24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf': 'Informe de gestión por dependencias 2024 (Alcaldía de Envigado)',
 'ejecucion-presupuestal-egresos-4-al-31122025.pdf': 'Ejecución presupuestal de egresos, 4.º trimestre, al 31 de diciembre de 2025 (Secretaría de Hacienda)',
 'ejecucion-presupuestal-egresos-2-al-30062026.pdf': 'Ejecución presupuestal de egresos, 2.º trimestre, al 30 de junio de 2026 (Secretaría de Hacienda)',
 '2025informegestiondependencias.pdf': 'Informe de gestión por dependencias 2025 (Alcaldía de Envigado)',
 'evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx': 'Plan Indicativo de Productos, evaluación definitiva 2T-2026 (Alcaldía de Envigado, DAP)',
 'proyectosdeinversion2024.xlsx': 'Proyectos de inversión 2024 (informe de gestión de la Alcaldía de Envigado)',
 'justificacionnoplandemejoramientovca2024.pdf': 'Justificación de no plan de mejoramiento, rendición de la política de víctimas 2024',
 'preguntasyrespuestasrendicioncuentasabril2024.pdf': 'Preguntas y respuestas de la rendición de cuentas de los 100 primeros días (abril de 2024)',
 'preguntasyrespuestasrendicioncuentassaludi2025.pdf': 'Preguntas y respuestas, rendición de cuentas de la Secretaría de Salud 2025-01',
 '20240321primerarendicincuentassalud-1.pdf': 'Informe-acta de la primera rendición de cuentas 2024 de la Secretaría de Salud',
 'preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf': 'Preguntas y respuestas de la rendición de cuentas 2025 del Alcalde',
 'justificacionrendiciondecuentasetnicos.pdf': 'Justificación de no plan de mejoramiento, rendición de cuentas de grupos étnicos',
 'infrendcuentasvca2025.pdf': 'Informe de rendición de cuentas de la inversión en población víctima 2025',
 'rendicion-cuentas-ppjuventud-sec-educacion_2025.pdf': 'Presentación "rendición de cuentas PP juventud" de la Secretaría de Educación 2025',
 'evaluaciondefinitivaplanaccionivtri.xlsx': 'Plan de acción, evaluación definitiva IV trimestre de 2025 (Alcaldía de Envigado, DAP)',
 'evaluaciondefinitivaplanaccion2026segundotrimestre20260730.xlsx': 'Plan de acción, evaluación definitiva 2T-2026 (Alcaldía de Envigado, DAP)',
 '2024actarendicioncuentassecretariasalud02docx.pdf': 'Acta de la audiencia pública de rendición de cuentas 2024-02 de la Secretaría de Salud',
 '129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf': 'Preguntas y recomendaciones, audiencia pública de rendición de cuentas del Alcalde 2024',
 '2025actarendicioncuentassecretariasalud01final.pdf': 'Acta de la audiencia pública de rendición de cuentas 2025-01 de la Secretaría de Salud',
 'cronograma-asambleas-pot.pdf': 'Programación de las asambleas zonales del POT (cronograma, publicado el 14 de febrero de 2025)',
 'memoriadeleventordcdiscapacidadii.pdf': 'Memoria del evento de rendición de cuentas del Comité Municipal de Discapacidad, semestre 2 de 2025',
 'memoriadeleventorddiscapacidad20251.pdf': 'Memoria del evento de rendición de cuentas del Comité Municipal de Discapacidad, semestre 1 de 2025',
 'preguntasyrespuestas1rddiscapacidad.pdf': 'Carta de respuesta del Comité Municipal de Discapacidad, semestre 1 de 2025 (nombre de persona omitido)',
 'rendicion-de-cuentas-semestre-1.pdf': 'Presentación de la rendición de cuentas del Comité Municipal de Discapacidad, semestre 1 de 2024',
 'TerriData05266f.txt.zip': 'DNP, TerriData Envigado',
 '8-dimension-funcional-_-espacio-publico.pdf': 'Documento diagnóstico de la revisión del POT, 8. Dimensión funcional - Espacio público (2025)',
 'anteproyecto29022024.pdf': 'Anteproyecto del Plan de Desarrollo Municipal 2024-2027 (29 de febrero de 2024)',
 'apendice91planindicativo27052024.xlsx': 'Apéndice 9.1.2 del Plan Indicativo del PDM 2024-2027',
 'plandesarrollo20242027.pdf': 'Plan de Desarrollo 2024-2027 (Acuerdo 017 de 2024)',
 'ppt-balance-pdm-2024_educacion_oficial.pdf': 'Balance del PDM 2024, Secretaría de Educación',
 'DCD-area-sexo-edad-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx': 'DANE, proyecciones de población municipal por área, sexo y edad simple 2020-2035 (post-COVID)',
 'Forensis_2024.pdf': 'INMLCF, Forensis 2024',
 'cifras-afiliacion-salud-agosto2026.zip': 'MinSalud, Cifras de aseguramiento en salud, agosto de 2026',
 'cifras-afiliacion-salud-dic-2025.zip': 'MinSalud, Cifras de aseguramiento en salud, diciembre de 2025',
 'uav5-b85g_homicidio_accidente_transito_envigado.csv': 'Policía Nacional vía datos.gov.co: homicidios en accidente de tránsito, Envigado',
}
RE_RUTA = re.compile(r'data_raw_IIC/(?:\[archivo con nombre de persona omitido\]|[^\s;|,()"\]\[])+')


def nombre_doc(ruta):
    ruta = ruta.rstrip('.')
    if '[archivo con nombre de persona omitido]' in ruta:
        if '905995' in ruta:
            return 'Carta de respuesta del Comité Municipal de Discapacidad, semestre 2 de 2024 (nombre de persona omitido)'
        return 'Carta de respuesta del Comité Municipal de Discapacidad, semestre 1 de 2024 (nombre de persona omitido)'
    base = ruta.rsplit('/', 1)[-1]
    if base in DOCS:
        return DOCS[base]
    raise KeyError(ruta)


# --------------------------------------------------------------------------- depuración de texto
OPC_BARE = {  # identificadores de opciones que aparecen sueltos en textos
 'ultima_fila_por_indicador': 'última fila por indicador', 'solo_31dic': 'solo filas del 31 de diciembre',
 'todas_filas_dic_100': 'todas las filas de diciembre al 100 %', 'solo_exactos': 'solo cruces exactos',
 'plan_accion_xlsx': 'plan de acción', 'anual_provisional': 'anual provisional', '12_con_comite': '12 zonas con comité',
 'metas_anuales': 'metas anuales', 'acta_por_zona': 'acta por zona', 'serie_publicada_por_fuente': 'serie publicada por la fuente',
 'dane_vp_2018_2042': 'proyección DANE 2018-2042', 'dane_postcovid_2020_2035': 'proyección DANE 2020-2035',
 'solo_plan_accion_zona': 'solo plan de acción por zona', 'hasta_fecha_consulta': 'hasta la fecha de consulta',
 'guia_31mar_siguiente': 'hasta el 31 de marzo del año siguiente', 'tope_100': 'tope de 100', 'cero_siempre': 'ausencia siempre',
 'nd_siempre': 'presencia desconocida', 'lista_asistencia': 'lista de asistencia', 'vigencia_rendida': 'vigencia rendida',
 'institucional_pi_545_546': 'consecutivos 545 y 546', 'amplio_todo_documento': 'todo documento', 'productos_plan_indicativo': 'productos del plan indicativo',
 'icpd7_mas3_gobernanza': '10 metas (con gobernanza)', 'icpd7': '7 metas del ICPD-6', 'inmlcf_forensis': 'Medicina Legal (Forensis)',
 'policia_datosgov': 'Policía (datos.gov.co)', 'excluir_felicitaciones': 'se excluyen felicitaciones', 'tipologia_sugerida': 'tipología sugerida',
 'posible_no_pp': 'posible no PP', 'control_institucional_545_546': 'control con consecutivos 545 y 546', 'envolvente_componentes': 'rango de sensibilidad por componente',
 'plan_accion_roca': 'plan de acción', 'sin_correspondencia': 'sin correspondencia', 'hacienda_30jun': 'Hacienda al 30 de junio',
 'ie_mejoran': 'instituciones que mejoran', 'fecha_hecho': 'fecha del hecho', 'total_e': 'total',
}
REGLAS = [  # (patrón, reemplazo) en orden
 (r'\s*\(?(?:ver|según|en|de)\s+(?:los\s+)?(?:notas?_\w+(?:\.md)?|clasificacion_\w+(?:\.csv)?|soporte_conteos(?:\.csv)?)\)?', ''),
 (r'\bnotas_(?:PDM|PP|CS|PC)(?:\.md)?\b', 'notas de trabajo'),
 (r'\bdsi\.md\b', 'notas de trabajo'),
 (r'\bclasificacion_vrc(?:\.csv)?\b', 'la hoja Clasificación VRC'),
 (r'\bclasificacion_rpo(?:\.csv)?\b', 'la hoja Clasificación RPO'),
 (r'\bclasificacion_pp_entregas(?:\.csv)?\b', 'la hoja Entregas PP'),
 (r'\bclasificacion_pp_2024(?:\.csv)?\b', 'la hoja PP por proyecto'),
 (r'\bsoporte_conteos(?:\.csv)?\b', 'la hoja Soporte de conteos'),
 (r'\bids_soporte\b', 'identificadores de soporte'),
 (r'\bAUX\.PDM\.ILE\.S\.IE_MEJORAN\b', 'la variable auxiliar de instituciones que mejoran'),
 (r'\bplan_accion_roca\b', 'plan de acción'),
 (r'\bapropiacion_definitiva\b', 'apropiación definitiva'),
 (r'\bcuadro_registro_incompleto\b', 'cuadro incompleto'),
 (r'\bpresencia_cero_si_integral\b', 'ausencia si el documento es integral'),
 (r'\bpresencia_nd_siempre\b', 'presencia desconocida'),
 (r'\bpresencia_cero_siempre\b', 'ausencia siempre'),
 (r'\bpresencia_lista_asistencia\b', 'lista de asistencia'),
 (r'\bno [Vv]alidad([oa])s?\b', r'sin validar'),
 (r'\brepositorio\b', 'archivo del observatorio'),
 (r'\bparciales\b', 'archivos de trabajo'),
 (r'\bparcial(?:es)?\b', 'incompleto'),
 (r'\bPARCIAL\b', 'INCOMPLETO'),
 (r'\bscripts?\b', 'cálculo del observatorio'),
 (r'\bRegistro!N\b', 'la fecha de registro de la Matriz'),
 (r'\bRegistro col\. ([A-Z])\b', r'columna \1 de la hoja Registro de la Matriz'),
 (r'`', ''),
]
REGLAS = [(re.compile(a), b) for a, b in REGLAS]
RE_ESC_FINAL = re.compile(r'\s*\.?\s*Escenario:\s*[^|]*$')
RE_PARAM_OPC = re.compile(r'\b(' + '|'.join(sorted(P, key=len, reverse=True)) + r')=([A-Za-z0-9_/.\-]+)')
RE_PARAM = re.compile(r'\b(' + '|'.join(sorted(P, key=len, reverse=True)) + r')\b')
RE_BARE = re.compile(r'\b(' + '|'.join(sorted(OPC_BARE, key=len, reverse=True)) + r')\b')
RE_PREFIJO_DOC = re.compile(r'^(?:[A-Z]{2,}[A-Z0-9\-]*(?: S\d-\d{4})?):\s+')
RE_SNAKE = re.compile(r'(?<![\w.\-/])[a-z0-9áéíóúñ]+(?:_[a-z0-9áéíóúñ]+)+(?![\w\-/]|\.\w)', re.I)
RE_FLAG = re.compile(r'\b([a-z0-9_áéíóúñ]+)=([01])\b', re.I)
ENLACE_RUTA = {}  # ruta -> URL, se llena al leer los datos


def _opcion(m):
    p, o = m.group(1), m.group(2).rstrip('.')
    txt = P[p][2].get(o, o)
    return f'{P[p][0]}: {txt[0].lower() + txt[1:] if txt[:2] not in ("N/", "PD", "PO", "CU") else txt}'


@lru_cache(maxsize=None)
def depurar(s):
    s = str(s or '')
    s = RE_ESC_FINAL.sub('', s)
    s = RE_RUTA.sub(lambda m: nombre_doc(m.group(0)), s)
    s = RE_PARAM_OPC.sub(_opcion, s)
    s = RE_PARAM.sub(lambda m: P[m.group(1)][0].lower(), s)
    for rx, rep in REGLAS:
        s = rx.sub(rep, s)
    s = RE_BARE.sub(lambda m: OPC_BARE[m.group(1)], s)
    s = gd.limpiar(s)
    s = re.sub(r'\bparcial', 'incompleto', s)
    s = RE_SNAKE.sub(lambda m: m.group(0).replace('_', ' '), s)
    for a, b in (('mas sectoriales', 'más sectoriales'), ('rendicion', 'rendición'), ('mencion en', 'mención en'), ('institucional 545 546', 'institucional 545 y 546'), ('tramite', 'trámite'), ('programacion', 'programación'), ('tambien', 'también'), ('despues', 'después'), ('construccion', 'construcción')):
        s = s.replace(a, b)
    return re.sub(r'\s{2,}', ' ', s).strip()


def doc_legible(fuente, ruta=''):
    """Nombre de documento legible (sin código interno ni ruta)."""
    t = RE_RUTA.sub(lambda m: nombre_doc(m.group(0)), fuente or '')
    if not t and ruta.startswith('data_raw'):
        t = nombre_doc(ruta)
    t = RE_PREFIJO_DOC.sub('', t)
    t = depurar(t)
    return t[:1].upper() + t[1:] if t else t


@lru_cache(maxsize=None)
def desc_escenario(esc):
    if esc == 'defecto':
        return 'Por defecto'
    if esc in ESPECIALES:
        return ESPECIALES[esc]
    partes = []
    for tok in esc.split(';'):
        p, o = tok.split('=', 1)
        partes.append(f'{P[p][0]}: {P[p][2][o]}')
    return '; '.join(partes)


def fecha(s):
    return dt.date.fromisoformat(s) if s else None


def num(s):
    if s in ('', None):
        return None
    try:
        f = float(s)
    except ValueError:
        return depurar(s)
    return int(f) if f.is_integer() and abs(f) < 1e15 else f


def sino(s):
    return {'1': 'Sí', 'True': 'Sí', 'true': 'Sí', '0': 'No', 'False': 'No', 'false': 'No', 'ND': 'N/D', 'sí': 'Sí', 'no': 'No'}.get(s, depurar(s))


def clave_terr(t):
    return 0 if t == 'Municipio' else int(t.split()[1])


def clave_cod(c):
    return ORDEN_COD.index(c) if c in ORDEN_COD else 99


def es(x, dec=2):
    """Número con formato español (coma decimal, punto de miles), sin ceros finales."""
    if x is None:
        return ''
    if abs(x - round(x)) < 1e-9:
        t = f'{int(round(x)):,}'
    else:
        t = f'{x:,.{dec}f}'
    return t.replace(',', '#').replace('.', ',').replace('#', '.')


REGLAS.insert(0, (re.compile('ENVOLVENTE_COMPONENTES'), 'rango de sensibilidad por componente'))


def legible(s):
    """Como depurar, y además convierte marcas clave=1 / clave=0 en 'clave: sí' / 'clave: no'."""
    s = RE_FLAG.sub(lambda m: f'{m.group(1).replace("_", " ")}: {"sí" if m.group(2) == "1" else "no"}', str(s or ''))
    return depurar(s)


# --------------------------------------------------------------------------- constructor de hojas
class Col:
    def __init__(self, header, ancho, tipo, campo=None, desc='', llave='', get=None, tabla=None):
        self.header, self.ancho, self.tipo, self.campo = header, ancho, tipo, campo
        self.desc, self.llave, self.tabla = desc, llave, tabla
        self.get = get or self._por_campo

    def _por_campo(self, r):
        v = r[self.campo]
        return {'txt': depurar, 'leg': legible, 'int': num, 'num': num, 'money': num, 'date': fecha, 'sino': sino, 'raw': str, 'pts': num}[self.tipo](v)


DICC_CSV = {(r['tabla'], r['columna']): r['descripcion'] for r in leer('diccionario_datos.csv')}
DICCIONARIO = []  # filas del diccionario del libro
FORMATOS = {'pts': '#,##0.00', 'int': '0', 'num': '#,##0.00', 'money': '#,##0', 'date': 'dd/mm/yyyy'}


def url_de(enlace, ruta=''):
    urls = [gd.normalizar_url(u.rstrip('.')) for u in gd.RE_URL.findall(enlace or '')]
    if not urls and ruta in ENLACE_RUTA:
        urls = [ENLACE_RUTA[ruta]]
    return urls


def hoja(wb, titulo, cols, filas, tabla_csv=None, notas_de=()):
    """Escribe una hoja a partir de especificaciones de columnas; devuelve la hoja."""
    datos, enlaces = [], {i: [] for i, c in enumerate(cols) if c.tipo == 'link'}
    for r in filas:
        fila = []
        for i, c in enumerate(cols):
            v = c.get(r)
            if c.tipo == 'link':
                urls = v if isinstance(v, list) else []
                enlaces[i].append(urls[0] if urls else None)
                v = '\n'.join(urls) if urls else gd.SIN_URL
            fila.append(v)
        datos.append(fila)
    # texto largo: el excedente pasa a Notas (misma convención del libro de disponibilidad)
    idx_largo = [i for i, c in enumerate(cols) if c.tipo in ('txt', 'leg')]
    notas = []
    for fila in datos:
        extra = []
        for i in idx_largo:
            if isinstance(fila[i], str) and len(fila[i]) > gd.MAX_CELDA:
                cab, resto = gd.corte_frase(fila[i], gd.MAX_CELDA)
                fila[i] = cab + ' (continúa en Notas)'
                extra.append(f'{cols[i].header} (continuación): {resto}')
        partes, t = [], ' | '.join(extra)
        while t:
            cab, t = gd.corte_frase(t, gd.MAX_CELDA)
            partes.append(cab)
        notas.append(partes)
    n_notas = max((len(p) for p in notas), default=0)
    cab = [c.header for c in cols]
    anch = [c.ancho for c in cols]
    if n_notas:
        cab += ['Notas'] + [f'Notas (continuación {i})' for i in range(1, n_notas)]
        anch += [55] * n_notas
        for fila, p in zip(datos, notas):
            fila += p + [''] * (n_notas - len(p))
    colorear = {i for i, c in enumerate(cols) if c.tipo == 'sem'}
    ws = gd.escribir_hoja(wb, titulo, cab, datos, anch, enlaces=enlaces or None, colorear=colorear)
    for i, c in enumerate(cols):
        fmt = FORMATOS.get(c.tipo)
        for cell in ws.iter_cols(min_col=i + 1, max_col=i + 1, min_row=2):
            for x in cell:
                if c.tipo in ('val', 'mix'):
                    if isinstance(x.value, float):
                        x.number_format = '#,##0.00'
                    elif isinstance(x.value, int):
                        x.number_format = '#,##0'
                elif fmt and x.value is not None:
                    if c.tipo == 'num' and isinstance(x.value, int):
                        x.number_format = '#,##0'
                    else:
                        x.number_format = fmt
    # diccionario
    for i, c in enumerate(cols):
        vals = [fila[i] for fila in datos if fila[i] not in (None, '')]
        distintos = sorted({str(v) for v in vals})
        if c.tipo in ('txt', 'leg', 'sino', 'sem') and 0 < len(distintos) <= 12:
            perm = ' | '.join(distintos)
        else:
            perm = {'int': 'Número entero', 'pts': 'Número entre 0 y 100 (dos decimales visibles)', 'num': 'Número decimal (dos decimales visibles)', 'money': 'Pesos colombianos (COP), número entero',
                    'date': 'Fecha dd/mm/aaaa', 'link': 'Enlace (URL) o "Solo en archivo local del observatorio"', 'txt': 'Texto libre',
                    'leg': 'Texto libre', 'val': 'Número, o N/D, NE o Revisar', 'mix': 'Número o texto'}.get(c.tipo, 'Texto')
        tipo = {'int': 'Entero', 'num': 'Decimal', 'pts': 'Decimal', 'money': 'Entero (COP)', 'date': 'Fecha', 'sino': 'Sí/No', 'sem': 'Texto (con color)',
                'link': 'Enlace', 'val': 'Número o texto', 'mix': 'Número o texto'}.get(c.tipo, 'Texto')
        desc = c.desc or (depurar(DICC_CSV.get((tabla_csv, c.campo), '')) if tabla_csv and c.campo else '') or c.header + '.'
        DICCIONARIO.append([titulo, c.header, tipo, desc, perm, c.llave])
    for j in range(n_notas):
        DICCIONARIO.append([titulo, 'Notas' if j == 0 else f'Notas (continuación {j})', 'Texto',
                            'Excedente de las celdas de texto que superan 600 caracteres.', 'Texto libre', ''])
    return ws


# --------------------------------------------------------------------------- lectura de datos
RES = leer('resultados_indicadores.csv')
VAR = leer('valores_variables.csv')
SOP = leer('soporte_conteos.csv')
PPY = leer('pp_proyectos.csv')
C_VRC, C_RPO = leer('clasificacion_vrc.csv'), leer('clasificacion_rpo.csv')
C_PP24, C_ENT = leer('clasificacion_pp_2024.csv'), leer('clasificacion_pp_entregas.csv')
YAML = yaml.safe_load(open(SUPUESTOS, encoding='utf-8'))
NOM_VAR = {v['id_variable']: v['nombre_variable'] for v in VAR}
NOM_IND = {r['codigo']: r['nombre'] for r in RES}
NOM_IND_CORTO = NOM_IND

for fuente, camp_r, camp_e in ((VAR, 'documento_ruta_local', 'enlace'), (SOP, 'ruta_local', 'enlace'), (C_VRC, 'ruta_local', 'enlace'),
                               (C_RPO, 'ruta_local', 'enlace'), (PPY, 'fuente', 'enlace')):
    for r in fuente:
        if r[camp_r].startswith('data_raw') and r[camp_e].startswith('http'):
            ENLACE_RUTA.setdefault(r[camp_r], gd.normalizar_url(r[camp_e].split(';')[0].strip()))
IDX = {(r['codigo'], r['anio'], r['periodo'], r['territorio'], r['escenario_id']): r for r in RES}


def sem_vis(r):
    d, s = r['disponibilidad'], r['semaforo']
    if d == 'Revisar':
        return 'Revisar'
    if s == 'Gris' and d == 'Calculado':
        return 'Sin semáforo'
    return s


def valor_res(r):
    if r['puntaje'] != '':
        return float(r['puntaje'])
    return r['disponibilidad']


def iguales(a, b):
    return a['disponibilidad'] == b['disponibilidad'] and (
        (a['puntaje'] == '' and b['puntaje'] == '') or (a['puntaje'] != '' and b['puntaje'] != '' and abs(float(a['puntaje']) - float(b['puntaje'])) < 1e-6)
    ) and sem_vis(a) == sem_vis(b)


def txt_res(r):
    if r is None:
        return 'Sin cálculo'
    v = valor_res(r)
    return es(v) if isinstance(v, float) else v


# --------------------------------------------------------------------------- efecto de cada alternativa
def calcular_efectos():
    out = []
    for p, (lab, preg, ops) in P.items():
        dflt = str(YAML['parametros'][p]['defecto'])
        for o, otxt in ops.items():
            if o == dflt:
                continue
            esc = f'{p}={o}'
            alts = [r for r in RES if r['escenario_id'] == esc and r['no_calculado_por_etapa'] == 'False']
            if p == 'regla_estado_calculo':
                out.append(dict(p=p, o=otxt, codigo='', anio='', periodo='', vd=None, sd='', va=None, sa='', cambia='Sí', zonas=0,
                                nota='Con la regla estricta de la Matriz todos los resultados serían N/D mientras no existan datos validados '
                                     '(columna "Puntaje con la regla estricta" de la hoja Resultados).'))
                continue
            if not alts:
                out.append(dict(p=p, o=otxt, codigo='', anio='', periodo='', vd=None, sd='', va=None, sa='', cambia='No', zonas=0,
                                nota='Ningún resultado calculado cambia: la opción solo afecta variables o indicadores que no se calculan en esta etapa.'))
                continue
            codigos = sorted({r['codigo'] for r in alts}, key=clave_cod)
            claves = sorted({(r['codigo'], r['anio'], r['periodo']) for r in alts} |
                            {(r['codigo'], r['anio'], r['periodo']) for r in RES if r['escenario_id'] == 'defecto' and r['territorio'] == 'Municipio'
                             and r['codigo'] in codigos and r['no_calculado_por_etapa'] == 'False'},
                            key=lambda k: (clave_cod(k[0]), k[1], ORDEN_PER[k[2]]))
            for cod, anio, per in claves:
                d = IDX.get((cod, anio, per, 'Municipio', 'defecto'))
                a = IDX.get((cod, anio, per, 'Municipio', esc))
                if d is None and a is None:
                    continue
                zc = sum(1 for z in range(1, 14) if (IDX.get((cod, anio, per, f'Zona {z}', esc)) is not None and IDX.get((cod, anio, per, f'Zona {z}', 'defecto')) is not None
                         and not iguales(IDX[(cod, anio, per, f'Zona {z}', esc)], IDX[(cod, anio, per, f'Zona {z}', 'defecto')])))
                cambia = 'Sí' if (d is None or a is None or not iguales(d, a)) else 'No'
                nota = ''
                if cambia == 'No' and zc:
                    nota = 'El resultado municipal no cambia; cambia en zonas.'
                elif cambia == 'No':
                    nota = 'No cambia el resultado.'
                out.append(dict(p=p, o=otxt, codigo=cod, anio=anio, periodo=per,
                                vd=None if d is None else valor_res(d), sd='' if d is None else sem_vis(d),
                                va=None if a is None else valor_res(a), sa='' if a is None else sem_vis(a),
                                cambia=cambia, zonas=zc, nota=nota, _d=d, _a=a))
    return out


EFECTOS = calcular_efectos()


def resumen_alternativa(p, otxt):
    fs = [e for e in EFECTOS if e['p'] == p and e['o'] == otxt]
    et = otxt if len(otxt) > 28 else f'{P[p][0]}, {otxt[0].lower() + otxt[1:]}'
    if p == 'regla_estado_calculo':
        return f'{et}: todos los resultados serían N/D mientras no existan datos validados.'
    if not fs[0]['codigo']:
        return f'{et}: no cambia ningún resultado calculado (solo afecta variables).'
    cam = [e for e in fs if e['cambia'] == 'Sí']
    if not cam:
        zon = sum(e['zonas'] for e in fs)
        return f'{et}: no cambia el resultado municipal' + (f' (cambia en {zon} resultados de zonas).' if zon else '.')
    partes = [f"{e['codigo']} {e['anio']}{'' if e['periodo'] == 'Anual' else ' S1'}: {txt_res(e['_d'])} → {txt_res(e['_a'])}" for e in cam]
    return f'{et}: ' + '; '.join(partes) + '.'


def tabla_seccion5():
    t = RESUMEN.read_text(encoding='utf-8')
    sec = t.split('## 5.')[1].split('## 6.')[0]
    filas = []
    for ln in sec.splitlines():
        if ln.startswith('|') and not ln.startswith('|---') and not ln.startswith('| Decisión'):
            c = [x.strip() for x in ln.strip().strip('|').split(' | ')]
            if len(c) == 5:
                filas.append(c)
    return filas


S5_MAP = {
 'Evento sin mención de veedurías': ['vrc_evento_sin_mencion'], 'Universo de eventos de rendición': ['vrc_universo'],
 'Acreditación de presencia y año de asignación': ['vrc_acreditacion_presencia', 'vrc_asignacion_anio'],
 'Ventana de publicación': ['ventana_documental_cs'], 'Lista de metas del DSI': ['dsi_lista_metas'],
 'Fuente del valor publicado (DSI)': ['dsi_fuente_valor'], 'Observaciones sobre el plan de desarrollo (RPO)': ['rpo_filtro_pdm'],
 'Universo de ejercicios (RPO)': ['rpo_universo'], 'Respuesta de fondo (RPO)': ['rpo_criterio_fondo'],
 'Cuadro publicado incompleto (RPO 2025)': ['rpo_registro_incompleto'], 'Actividades que cubren una zona (I2)': ['pc_i2_actividades'],
 'Evidencia de actividad en la zona (I2)': ['pc_i2_evidencia'], 'Zonas oficiales': ['zonas_oficiales'], 'Unidad de I3': ['pc_i3_unidad'],
 'Formación para servidores (I3)': ['pc_i3_excluir_servidores'], 'Instancias seleccionadas (I5)': ['pc_i5_instancias'],
 'Horizonte de la meta del plan de desarrollo': ['horizonte_meta_pdm'], 'Espacio público verde (CA)': ['pdm_ca_fuente'],
 'Mortalidad vial (CM) y población': ['pdm_cm_fuente', 'poblacion_dane'], 'Método de PP 2024': ['pp_2024_metodo'],
 'Proyecto 0553 (costo del proceso)': ['pp_incluir_0553'], 'Fuente de cálculo del PP': ['pp_fuente_calculo'], 'Zona 1': ['pp_zona1_tratamiento'],
 'Tratamiento de 2026': ['tratamiento_2026'], 'Fecha de registro': ['corte_fecha_registro'], 'Regla de estado para calcular': ['regla_estado_calculo'],
}


# --------------------------------------------------------------------------- hojas de datos
def ent_txt(r):
    ids = [r[f'entrada_{i}_id'] for i in range(1, 5) if r[f'entrada_{i}_id']]
    if r['codigo'] == 'CS.ICRCS':
        return 'Puntajes de los cuatro componentes del índice: 1) DSI; 2) PCS; 3) VRC; 4) RPO (promedio simple)'
    return depurar('; '.join(f'{k}) {NOM_VAR.get(x, x)}' for k, x in enumerate(ids, 1)))


def obs_res(r):
    a, o = depurar(r['alerta']), depurar(r['observaciones'])
    return ' | '.join(dict.fromkeys(x for x in (a, o) if x))


def orden_res(r):
    return (0 if r['es_defecto'] == 'True' else 1, ORDEN_SIS.index(r['sistema']), clave_cod(r['codigo']), r['anio'],
            ORDEN_PER[r['periodo']], r['escenario_id'] if r['es_defecto'] != 'True' else '', clave_terr(r['territorio']), r['escenario_id'])


def cols_resultados():
    return [
     Col('Sistema', 20, 'txt', 'sistema', 'Sistema del Observatorio al que pertenece el indicador.', tabla='resultados_indicadores'),
     Col('Código del indicador', 11, 'txt', 'codigo', 'Código único del indicador o índice.', 'Llave: código del indicador'),
     Col('Indicador', 38, 'txt', 'nombre', 'Nombre del indicador o índice.'),
     Col('Año', 7, 'int', 'anio', 'Año observado.', 'Llave: año'),
     Col('Periodo', 11, 'txt', 'periodo', 'Anual o Semestre 1 (2026, sin cierre de vigencia).', 'Llave: periodo'),
     Col('Territorio', 11, 'txt', 'territorio', 'Municipio o Zona 1 a Zona 13.', 'Llave: territorio'),
     Col('Escenario (qué cambia respecto al valor por defecto)', 52, 'txt', get=lambda r: desc_escenario(r['escenario_id']),
         desc='Descripción legible de los supuestos que cambian respecto al escenario por defecto; "Por defecto" para el escenario por defecto.'),
     Col('¿Es el escenario por defecto?', 11, 'sino', 'es_defecto', 'Sí si el resultado corresponde al escenario por defecto (provisional).'),
     Col('Entrada 1 (numerador)', 16, 'num', 'entrada_1', 'Primera entrada de la fórmula: numerador (observado en índices y metas).'),
     Col('Entrada 2 (denominador)', 16, 'num', 'entrada_2', 'Segunda entrada de la fórmula: denominador (meta en indicadores de cumplimiento del PDM).'),
     Col('Entrada 3', 11, 'num', 'entrada_3', 'Tercera entrada (solo índices de cuatro componentes).'),
     Col('Entrada 4', 11, 'num', 'entrada_4', 'Cuarta entrada (solo índices de cuatro componentes).'),
     Col('Qué son las entradas', 46, 'txt', get=ent_txt, desc='Nombre de las variables o componentes que corresponden a cada entrada.'),
     Col('Puntaje', 10, 'pts', 'puntaje', 'Resultado del indicador en escala 0 a 100 (sin redondear; se muestra con dos decimales). Vacío si es N/D, NE o Revisar.'),
     Col('Semáforo', 13, 'sem', get=sem_vis, desc='Verde, Amarillo, Rojo según los umbrales de la Matriz; Gris si es N/D o NE; "Revisar" y "Sin semáforo" (primer semestre de 2026) no tienen color.'),
     Col('Estado del cálculo', 12, 'txt', 'disponibilidad', 'Calculado, N/D (no se puede calcular), NE (no evaluable) o Revisar (numerador mayor que denominador).'),
     Col('Variables usadas', 28, 'txt', get=lambda r: r['variables_usadas'].replace(';', '; '), desc='Códigos de las variables (hoja Variables) que alimentan el resultado.', llave='Llave: código de variable (hoja Variables)'),
     Col('Motivo si N/D', 44, 'txt', 'motivo_nd', 'Razón por la que el resultado es N/D, NE o Revisar.'),
     Col('Alertas y observaciones', 60, 'txt', get=obs_res, desc='Alertas sobre el dato y observaciones del cálculo.'),
     Col('Puntaje con la regla estricta de la Matriz', 14, 'txt', 'puntaje_regla_matriz', 'Resultado si solo se aceptaran datos validados: N/D en todos los casos mientras no haya validación.'),
     Col('Estado', 12, 'txt', 'estado', 'Siempre Provisional: ningún resultado ha pasado por validación.'),
     Col('Fecha de corte', 12, 'date', 'fecha_corte_comun', 'Fecha de corte de la información (31/12 del año; 30/06/2026 para el primer semestre de 2026).'),
     Col('Clave de escenario', 30, 'raw', 'escenario_id', 'Identificador técnico del escenario ("defecto" o lista de cambios parámetro=opción); sirve para unir hojas.', 'Llave: escenario'),
    ]


def cols_variables():
    def doc(r):
        return doc_legible(r['fuente'], r['documento_ruta_local'])
    return [
     Col('Sistema', 20, 'txt', 'sistema', tabla='valores_variables'),
     Col('Código del indicador', 11, 'txt', 'codigo_indicador', 'Indicador al que pertenece la variable.', 'Llave: código del indicador'),
     Col('Código de la variable', 18, 'txt', 'id_variable', 'Código de la variable en la Matriz (o variable auxiliar AUX...).', 'Llave: código de variable'),
     Col('Variable', 50, 'txt', 'nombre_variable', 'Nombre completo de la variable.'),
     Col('Unidad', 14, 'txt', 'unidad', 'Unidad de medida.'),
     Col('Tipo de valor', 12, 'txt', get=lambda r: r['tipo_valor'].capitalize(), desc='Observado, Meta, Construido, Proxy o Auxiliar.'),
     Col('Año', 7, 'int', 'anio', 'Año del dato.', 'Llave: año'),
     Col('Periodo', 11, 'txt', 'periodo', 'Anual o Semestre 1.', 'Llave: periodo'),
     Col('Territorio', 11, 'txt', 'territorio', 'Municipio o Zona 1 a Zona 13.', 'Llave: territorio'),
     Col('Escenario (qué cambia respecto al valor por defecto)', 48, 'txt', get=lambda r: desc_escenario(r['escenario_id']), desc='Descripción legible del escenario; "Por defecto" para el escenario por defecto.'),
     Col('¿Es el escenario por defecto?', 11, 'sino', 'es_defecto', 'Sí si el valor corresponde al escenario por defecto (provisional).'),
     Col('Valor', 16, 'num', 'valor', 'Valor de la variable sin redondear; vacío si no hay dato.'),
     Col('Estado de la información', 14, 'txt', 'estado_informacion', 'Pendiente (sin validar), No disponible o No evaluable.'),
     Col('Disponibilidad del dato', 14, 'txt', get=lambda r: {'Parcial': 'Incompleta'}.get(r['categoria_disponibilidad'], r['categoria_disponibilidad']),
         desc='Categoría de disponibilidad (Disponible, Construible, Incompleta, Proxy, No disponible). "Incompleta" equivale a la categoría intermedia de la hoja de disponibilidad.'),
     Col('Fuente (documento)', 48, 'txt', get=doc, desc='Nombre del documento o publicación de origen.'),
     Col('Enlace al documento', 40, 'link', get=lambda r: url_de(r['enlace'], r['documento_ruta_local']), desc='Enlace clicable al documento de origen (dominio espejo para la Alcaldía).'),
     Col('Página, hoja o tabla', 40, 'txt', 'pagina_tabla', 'Ubicación del dato en el documento.'),
     Col('Fecha de corte', 12, 'date', 'fecha_corte', 'Fecha de corte registrada (según el supuesto de fecha de corte).'),
     Col('Fecha de corte real del dato', 12, 'date', 'fecha_corte_real', 'Fecha de corte real del dato en la fuente.'),
     Col('Fecha de publicación', 12, 'date', 'fecha_publicacion', 'Fecha interna de publicación del documento, si se conoce.'),
     Col('Fecha de consulta', 12, 'date', 'fecha_consulta', 'Fecha en que se consultó la fuente.'),
     Col('Horizonte de la meta', 10, 'int', 'horizonte_meta', 'Año de la meta (p. ej. 2027).'),
     Col('N.º de elementos de soporte', 11, 'int', 'n_elementos_soporte', 'Cantidad de filas de la hoja Soporte de conteos que suman el valor.'),
     Col('Identificadores de soporte', 30, 'txt', get=lambda r: r['ids_soporte'].replace(',', '; '), desc='Identificadores de las filas de la hoja Soporte de conteos.', llave='Llave: ID de soporte'),
     Col('Observaciones', 70, 'txt', 'observaciones', 'Brechas, discrepancias y alternativas del dato.'),
     Col('Clave de escenario', 30, 'raw', 'escenario_id', 'Identificador técnico del escenario; sirve para unir hojas.', 'Llave: escenario'),
    ]


def build_pp_zona():
    base = [r for r in RES if r['codigo'] == 'PP.EPP' and r['escenario_id'] == 'defecto']
    base.sort(key=lambda r: (r['anio'], ORDEN_PER[r['periodo']], clave_terr(r['territorio'])))
    etiq = {'2024': '2024 (aproximado)', '2025': '2025', '2026': '2026, primer semestre (corte 30 de junio; sin semáforo)'}

    def alt(esc):
        return lambda r: (lambda x: None if x is None or x['puntaje'] == '' else float(x['puntaje']))(
            IDX.get((r['codigo'], r['anio'], r['periodo'], r['territorio'], esc)))
    cols = [
     Col('Año', 7, 'int', 'anio', 'Año.', 'Llave: año'), Col('Periodo', 11, 'txt', 'periodo', 'Periodo.', 'Llave: periodo'),
     Col('Año (rótulo)', 26, 'txt', get=lambda r: etiq[r['anio']], desc='Año con su rótulo (2024 aproximado; 2026 primer semestre sin semáforo).'),
     Col('Territorio', 11, 'txt', 'territorio', 'Municipio o Zona. El municipal es la suma de obligaciones sobre la suma de apropiaciones, no un promedio de zonas.', 'Llave: territorio'),
     Col('Apropiación definitiva (APP, COP)', 18, 'money', 'entrada_2', 'Apropiación definitiva del presupuesto participativo, en pesos.'),
     Col('Obligaciones (OPP, COP)', 18, 'money', 'entrada_1', 'Obligaciones del presupuesto participativo, en pesos.'),
     Col('% de ejecución (OPP / APP)', 12, 'pts', 'puntaje', 'Puntaje del indicador de ejecución presupuestal (EPP), escala 0 a 100. Vacío en la Zona 1 (NE).'),
     Col('Semáforo', 13, 'sem', get=sem_vis, desc='Semáforo del escenario por defecto; el primer semestre de 2026 no tiene semáforo.'),
     Col('Estado del cálculo', 12, 'txt', 'disponibilidad', 'Calculado, o NE (no evaluable) para la Zona 1.'),
     Col('% si se incluye el proyecto 0553', 14, 'pts', get=alt('pp_incluir_0553=true'), desc='Alternativa: incluir el costo del proceso (proyecto 0553).'),
     Col('% solo cruces exactos (2024)', 14, 'pts', get=alt('pp_2024_metodo=solo_exactos'), desc='Alternativa 2024: solo los 35 proyectos con correspondencia exacta.'),
     Col('% con Hacienda como fuente', 14, 'pts', get=alt('pp_fuente_calculo=hacienda'), desc='Alternativa: apropiación y obligaciones tomadas de Hacienda (2025 y 2026).'),
     Col('Alertas y observaciones', 60, 'txt', get=obs_res, desc='Alertas y observaciones del cálculo.'),
     Col('Estado', 12, 'txt', 'estado', 'Siempre Provisional.'),
     Col('Clave de escenario', 20, 'raw', 'escenario_id', 'Escenario por defecto.', 'Llave: escenario'),
    ]
    return cols, base


def build_pp_proyecto():
    c24 = {(c['zona'], c['codigo_hacienda']): c for c in C_PP24}

    def c(r):
        return c24.get((r['zona'], r['codigo_proyecto'])) if r['anio'] == '2024' else None

    def ejec(r):
        a, o = num(r['apropiacion_definitiva']), num(r['obligacion'])
        return None if not a or o is None else round(o / a * 100, 6)
    cols = [
     Col('Año', 7, 'int', 'anio', 'Año.', 'Llave: año'), Col('Periodo', 11, 'txt', 'periodo', 'Anual o Semestre 1 (2026).', 'Llave: periodo'),
     Col('Zona', 8, 'txt', get=lambda r: 'Sin zona' if r['zona'] == 'N/A' else int(r['zona']), desc='Zona del presupuesto participativo (2 a 13); "Sin zona" para el costo del proceso.', llave='Llave: territorio (Zona n)'),
     Col('Código del proyecto', 16, 'txt', 'codigo_proyecto', 'Código del proyecto de inversión (o código provisional EFPC cuando no hay cruce).'),
     Col('BPIN', 16, 'txt', 'bpin', 'Código BPIN del proyecto, si se conoce.'),
     Col('Nombre del proyecto', 60, 'txt', 'nombre_proyecto', 'Nombre del proyecto.'),
     Col('Unidad ejecutora', 30, 'txt', 'unidad_ejecutora', 'Dependencia que ejecuta el proyecto.'),
     Col('¿En el listado de la EFPC?', 11, 'sino', 'en_listado_efpc', 'Sí si el proyecto figura en el listado de la Escuela de Formación para la Participación Ciudadana.'),
     Col('Valor en el listado de la EFPC (COP)', 16, 'money', 'valor_listado_efpc', 'Valor del proyecto en el listado de la EFPC.'),
     Col('Tipo de correspondencia', 16, 'txt', get=lambda r: legible(r['tipo_correspondencia']).capitalize(), desc='Exacta, aproximada o sin correspondencia entre el listado y la fuente presupuestal.'),
     Col('Apropiación definitiva (APP, COP)', 17, 'money', 'apropiacion_definitiva', 'Apropiación definitiva, en pesos.'),
     Col('Compromisos (COP)', 17, 'money', 'compromiso', 'Compromisos, en pesos.'),
     Col('Obligaciones (OPP, COP)', 17, 'money', 'obligacion', 'Obligaciones, en pesos.'),
     Col('% de ejecución (OPP / APP)', 12, 'pts', get=ejec, desc='Obligaciones sobre apropiación por cien (calculado a partir de las dos columnas anteriores).'),
     Col('¿Incluido en el cálculo por defecto?', 12, 'sino', 'incluido_epp_defecto', 'Sí si el proyecto entra en el escenario por defecto (el 0553 se excluye).'),
     Col('Fuente del valor', 26, 'txt', 'fuente_valor', 'Fuente de los valores de APP y OPP.'),
     Col('Documento fuente', 48, 'txt', get=lambda r: doc_legible(r['fuente']), desc='Documento de origen.'),
     Col('Enlace al documento', 40, 'link', get=lambda r: url_de(r['enlace'], r['fuente']), desc='Enlace clicable al documento.'),
     Col('Página u hoja', 36, 'txt', 'pagina_hoja', 'Ubicación del dato en el documento.'),
     Col('Fecha de corte', 12, 'date', 'fecha_corte', 'Fecha de corte registrada.'),
     Col('Fecha de corte real', 12, 'date', 'fecha_corte_real', 'Fecha de corte real del dato.'),
     Col('Estado de la información', 13, 'txt', 'estado_informacion', 'Pendiente (sin validar).'),
     Col('Observaciones', 60, 'txt', 'observaciones', 'Observaciones del cruce y del cálculo.'),
     Col('Tipo de cruce 2024 (listado × Hacienda)', 16, 'txt', get=lambda r: legible(c(r)['tipo_coincidencia']).capitalize() if c(r) else '', desc='Solo 2024: exacta, aproximada o sin correspondencia.'),
     Col('Posible no PP (2024)', 11, 'sino', get=lambda r: sino(c(r)['posible_no_pp']) if c(r) else '', desc='Sí si el proyecto de 2024 incluye rubros de déficit fiscal que posiblemente no son presupuesto participativo.'),
     Col('APP del rubro de déficit fiscal 2.3.7.06 (2024, COP)', 17, 'money', get=lambda r: num(c(r)['app_rubro_2_3_7_06']) if c(r) else None, desc='Solo 2024: apropiación del rubro de disminución de pasivos incluida en el proyecto.'),
     Col('Criterio del cruce 2024', 60, 'txt', get=lambda r: depurar(c(r)['criterio']) if c(r) else '', desc='Solo 2024: criterio con el que se cruzó el proyecto del listado con Hacienda.'),
    ]
    filas = sorted(PPY, key=lambda r: (r['anio'], ORDEN_PER[r['periodo']], 99 if r['zona'] == 'N/A' else int(r['zona']), r['codigo_proyecto']))
    return cols, filas


def doc_sop(r):
    d = r['documento']
    if re.search(r'\.\w{3,4}$', d) and r['ruta_local']:
        d = ''
    return doc_legible(d, r['ruta_local']) if d or r['ruta_local'] else ''


def cols_soporte():
    return [
     Col('ID de soporte', 17, 'txt', 'id_soporte', 'Identificador de la fila de soporte.', 'Llave: ID de soporte'),
     Col('Sistema', 20, 'txt', 'sistema'), Col('Código del indicador', 11, 'txt', 'codigo_indicador', 'Indicador al que apoya el conteo.', 'Llave: código del indicador'),
     Col('Código de la variable', 20, 'txt', get=lambda r: r['id_variable'].replace(';', '; '), desc='Variable(s) que el elemento alimenta.', llave='Llave: código de variable'),
     Col('Año', 7, 'int', 'anio', 'Año.', 'Llave: año'), Col('Periodo', 10, 'txt', 'periodo', 'Periodo.', 'Llave: periodo'),
     Col('Territorio', 11, 'txt', 'territorio', 'Territorio.', 'Llave: territorio'),
     Col('Escenario (qué cambia respecto al valor por defecto)', 44, 'txt', get=lambda r: desc_escenario(r['escenario_id']), desc='Descripción legible del escenario.'),
     Col('¿Cuenta en el escenario por defecto?', 11, 'sino', 'cuenta_en_defecto', 'Sí si el elemento se cuenta en el escenario por defecto.'),
     Col('¿Incluido en este escenario?', 11, 'sino', 'incluido', 'Sí o No: si el elemento se incluye en el conteo del escenario de la fila.'),
     Col('Tipo de elemento', 18, 'leg', 'elemento_tipo', 'Tipo de elemento contado (zona, evento, observación, proyecto, meta...).'),
     Col('Identificador del elemento', 22, 'txt', 'elemento_id', 'Identificador del elemento contado.'),
     Col('Descripción', 60, 'txt', 'descripcion', 'Descripción del elemento contado.'),
     Col('Fecha del elemento', 12, 'date', 'fecha_elemento', 'Fecha del elemento (evento, actividad), si aplica.'),
     Col('Documento', 50, 'txt', get=doc_sop, desc='Nombre del documento que respalda el elemento.'),
     Col('Enlace al documento', 38, 'link', get=lambda r: url_de(r['enlace'], r['ruta_local']), desc='Enlace clicable al documento.'),
     Col('Página u hoja', 30, 'txt', 'pagina', 'Ubicación en el documento.'),
     Col('Fecha de publicación del documento', 12, 'date', 'fecha_publicacion_doc', 'Fecha interna de publicación del documento.'),
     Col('Banderas', 40, 'leg', 'banderas', 'Marcas de clasificación del elemento (sí o no).'),
     Col('Clasificación', 32, 'leg', 'clasificacion', 'Clasificación asignada al elemento.'),
     Col('Justificación', 60, 'leg', 'justificacion', 'Justificación de la clasificación.'),
     Col('Criterio aplicado', 60, 'leg', 'criterio', 'Criterio con el que se clasificó el elemento.'),
     Col('Clave de escenario', 30, 'raw', 'escenario_id', 'Identificador técnico del escenario.', 'Llave: escenario'),
    ]


def f(h, w, t, k, d='', l='', **kw):
    return Col(h, w, t, k, desc=d, llave=l, **kw)


def cols_vrc():
    doc = lambda r: doc_legible(r['documento'], r['ruta_local'])
    return [f('ID del evento', 16, 'txt', 'id', 'Identificador del evento clasificado.', 'Llave: ID de soporte'), f('Año del evento', 9, 'int', 'anio_evento'), f('Año de la vigencia', 9, 'int', 'anio_vigencia', 'Año de la vigencia rendida.'),
     f('Fecha del evento', 12, 'date', 'fecha_evento'), f('Nota sobre la fecha', 40, 'txt', 'fecha_evento_nota'), f('Evento', 44, 'txt', 'evento'),
     Col('Documento', 50, 'txt', get=doc, desc='Documento que respalda la clasificación.'), f('Página', 10, 'txt', 'pagina'),
     Col('Enlace al documento', 38, 'link', get=lambda r: url_de(r['enlace'], r['ruta_local']), desc='Enlace clicable al documento.'),
     f('Cita del documento', 60, 'txt', 'cita', 'Cita textual (ya anonimizada).'), f('Clasificación', 50, 'leg', 'clasificacion'), f('Criterio', 60, 'leg', 'criterio'),
     f('Tipo de documento', 22, 'leg', 'tipo_documento'), f('¿Documento integral?', 10, 'sino', 'documento_integral'),
     Col('¿Menciona veedurías?', 12, 'txt', get=lambda r: r['mencion_veedurias'].capitalize(), desc='Sí, No o Desconocida.'),
     f('¿Es la audiencia del Alcalde (consecutivo 546)?', 12, 'sino', 'es_546_alcalde'), f('¿Es el evento de veedurías (consecutivo 545)?', 12, 'sino', 'es_545_veedurias'),
     f('¿Sectorial con acta o memoria?', 12, 'sino', 'sectorial_con_acta_o_memoria'), f('¿Evento confirmado?', 10, 'sino', 'evento_confirmado'),
     f('¿En el universo "solo Alcalde"?', 11, 'sino', 'u_alcalde'), f('¿En el universo 545 y 546?', 11, 'sino', 'u_institucional_545_546'),
     f('¿En el universo Alcalde y sectoriales con acta?', 12, 'sino', 'u_alcalde_mas_sectoriales_con_acta'), f('¿En el universo amplio?', 11, 'sino', 'u_amplio_todo_documento'),
     f('Presencia si "desconocida = N/D"', 12, 'sino', 'presencia_nd_siempre'), f('Presencia si "ausencia solo en documento integral"', 14, 'sino', 'presencia_cero_si_integral'),
     f('Presencia si "ausencia siempre"', 12, 'sino', 'presencia_cero_siempre'), f('Presencia según lista de asistencia', 12, 'sino', 'presencia_lista_asistencia'),
     f('Fecha de publicación de la evidencia', 12, 'date', 'fecha_pub_evidencia'), f('Base de la fecha de publicación', 44, 'txt', 'base_fecha_pub'),
     f('¿Evidencia dentro de la ventana (fecha del evento)?', 12, 'sino', 'evidencia_en_ventana_fecha_evento'), f('¿Evidencia dentro de la ventana (vigencia)?', 12, 'sino', 'evidencia_en_ventana_vigencia'),
     f('Documentos complementarios', 44, 'txt', 'documentos_complementarios'), f('¿Clasificación dudosa?', 10, 'sino', 'dudoso'), f('Nota sobre la duda', 60, 'txt', 'nota_dudoso')]


def cols_rpo():
    doc = lambda r: doc_legible(r['documento'], r['ruta_local'])
    return [f('ID de la observación', 16, 'txt', 'id', 'Identificador de la observación.', 'Llave: ID de soporte'), f('Año', 7, 'int', 'anio', '', 'Llave: año'), f('Ejercicio', 38, 'txt', 'ejercicio'),
     f('N.º en el cuadro', 9, 'int', 'num_en_cuadro'), f('Fecha del evento', 12, 'date', 'fecha_evento'),
     Col('Documento', 50, 'txt', get=doc, desc='Documento que respalda la clasificación.'), f('Página', 8, 'txt', 'pagina'),
     Col('Enlace al documento', 38, 'link', get=lambda r: url_de(r['enlace'], r['ruta_local']), desc='Enlace clicable al documento.'),
     f('Cita de la observación', 60, 'txt', 'cita', 'Cita textual (ya anonimizada).'), f('Cita de la respuesta', 60, 'txt', 'cita_respuesta', 'Cita textual de la respuesta (ya anonimizada).'),
     f('Clasificación', 44, 'leg', 'clasificacion'), f('Criterio', 60, 'leg', 'criterio'), f('Tipo de intervención', 18, 'leg', 'tipo_intervencion'),
     f('¿Es felicitación?', 10, 'sino', 'es_felicitacion'), f('¿Es trámite individual?', 10, 'sino', 'es_tramite_individual'), f('¿Remite al PDM?', 10, 'sino', 'tema_pdm'),
     f('¿Tiene texto de respuesta?', 10, 'sino', 'tiene_texto_respuesta'), f('¿Respuesta solo de acuse o remisión?', 11, 'sino', 'respuesta_solo_acuse_o_remision'),
     f('¿Respuesta de fondo (criterio estricto)?', 12, 'sino', 'fondo_estricto'), f('¿Respuesta de fondo (criterio laxo)?', 12, 'sino', 'fondo_laxo'),
     f('Fecha de publicación de la respuesta', 12, 'date', 'fecha_publicacion_respuesta'), f('Base de la fecha de publicación', 44, 'txt', 'base_fecha_pub'),
     f('¿Publicada al 31 de diciembre?', 11, 'sino', 'publicada_al_31dic'), f('¿En el universo "solo Alcalde"?', 11, 'sino', 'en_universo_alcalde'),
     f('¿En el universo Alcalde y 100 días?', 11, 'sino', 'en_universo_alcalde_y_100dias'), f('¿En el universo de todos los cuadros?', 11, 'sino', 'en_universo_todos_los_cuadros'),
     f('¿Cuenta (se excluyen felicitaciones)?', 12, 'sino', 'en_ot_excluir_felicitaciones'), f('¿Cuenta (solo temático del PDM)?', 12, 'sino', 'en_ot_tematico_pdm'),
     f('¿Cuenta (todas)?', 10, 'sino', 'en_ot_todas'), f('¿Cuadro publicado incompleto?', 11, 'sino', 'cuadro_registro_incompleto'),
     f('¿Clasificación dudosa?', 10, 'sino', 'dudoso'), f('Nota sobre la duda', 60, 'txt', 'nota_dudoso')]


def cols_entregas():
    return [f('Año', 7, 'int', 'anio', '', 'Llave: año'), f('Zona', 8, 'txt', 'zona', 'Zona (2 a 13) o N/A.'), f('Código del proyecto', 16, 'txt', 'codigo_proyecto'),
     f('Nombre del proyecto', 60, 'txt', 'nombre_proyecto'), f('¿Es el proyecto 0553?', 10, 'sino', 'es_0553'), f('¿Tiene entrega exigible (PE)?', 11, 'sino', 'pe_cuenta'),
     f('¿Cuenta en el escenario por defecto (PE)?', 11, 'sino', 'cuenta_en_defecto_pe'),
     f('Entrega cumplida: todas las filas de diciembre al 100 %', 14, 'sino', 'pec_todas_filas_dic_100'), f('Entrega cumplida: solo filas del 31 de diciembre', 14, 'sino', 'pec_solo_31dic'),
     f('Entrega cumplida: última fila por indicador', 14, 'sino', 'pec_ultima_fila'), f('Regla por defecto', 24, 'leg', 'pec_regla_defecto'),
     f('Entrega cumplida según la regla por defecto', 14, 'sino', 'pec_defecto'), f('N.º de productos o actividades', 11, 'int', 'n_productos_o_actividades'),
     f('N.º de filas de diciembre', 11, 'int', 'n_filas_dic'), f('N.º de filas no numéricas', 11, 'int', 'n_filas_dic_no_numericas'),
     f('Tipología sugerida', 20, 'txt', 'tipologia_sugerida'), f('¿Aplica la meta de atención? (por defecto)', 12, 'txt', 'aplica_meta_atencion_defecto'),
     f('¿Aplica la meta de atención? (según tipología)', 12, 'txt', 'aplica_meta_atencion_tipologia'), f('Motivo de exclusión', 40, 'txt', 'motivo_exclusion'),
     f('Personas atendidas', 11, 'int', 'personas_atendidas_PA'), f('Meta de atención', 11, 'int', 'meta_atencion_MP'), f('Estado de la información', 13, 'txt', 'estado_informacion'),
     Col('Documento fuente', 48, 'txt', get=lambda r: doc_legible(r['fuente']), desc='Documento de origen.'),
     Col('Enlace al documento', 38, 'link', get=lambda r: url_de('', r['fuente']), desc='Enlace clicable al documento.'),
     f('Página u hoja', 30, 'txt', 'pagina_hoja'), f('Observaciones', 60, 'txt', 'observaciones')]


# --------------------------------------------------------------------------- supuestos
def build_supuestos():
    s5 = tabla_seccion5()
    por_param = {}
    extras = []
    for dec, preg, ops, ind, rango in s5:
        if dec in S5_MAP:
            for p in S5_MAP[dec]:
                por_param[p] = (rango, ind)
        else:
            extras.append((dec, preg, ops, ind, rango))
    assert len(extras) == 3 and len(P) == 36 and set(P) == set(YAML['parametros'])
    filas = []
    for n, (p, (lab, preg, ops)) in enumerate(P.items(), 1):
        y = YAML['parametros'][p]
        dflt = str(y['defecto'])
        afecta = ', '.join(y['afecta']).replace('todos', 'Todos los indicadores')
        alts = [t for o, t in ops.items() if o != dflt]
        deriv = [resumen_alternativa(p, t) for t in alts]
        sint = depurar(por_param[p][0]) if p in por_param else 'Detalle por alternativa en la columna siguiente.'
        filas.append(dict(n=n, tipo='Parámetro', tema=TEMA(p), clave=p, nombre=lab, preg=preg,
                          defecto=ops[dflt] + ' (PROVISIONAL)', alts='\n'.join(f'• {t}' for t in alts), afecta=afecta,
                          sint=sint, detalle=gd.lista_acotada(deriv, 'alternativas más en la hoja Efecto de alternativas'),
                          estado='PROVISIONAL: decisión pendiente de los profesores'))
    for k, (dec, preg, ops, ind, rango) in enumerate(extras, 37):
        nombre = re.sub(r'\s*\(.*fuera de las opciones configuradas\)', '', dec) if 'fuera de' in dec else dec
        nombre = re.sub(r'\s*\((I3; )?fuera de las opciones configuradas\)', '', dec).strip()
        nombre = re.sub(r'\s*\([^)]*fuera de las opciones configuradas\)', '', dec).strip()
        filas.append(dict(n=k, tipo='Sensibilidad adicional', tema='', clave='', nombre=nombre, preg=preg,
                          defecto=ops.split(' / ')[0] + ' (PROVISIONAL)', alts='\n'.join(f'• {o}' for o in ops.split(' / ')[1:]), afecta=ind,
                          sint=depurar(rango), detalle='Se evaluó fuera de los parámetros configurables del cálculo; el resultado está en la columna anterior.',
                          estado='PROVISIONAL: decisión pendiente de los profesores'))
    return filas


def cols_supuestos():
    return [
     Col('N.º', 5, 'int', get=lambda r: r['n'], desc='Número de orden.'),
     Col('Tipo', 16, 'txt', get=lambda r: r['tipo'], desc='Parámetro (configurable) o Sensibilidad adicional evaluada fuera de los parámetros.'),
     Col('Tema', 20, 'txt', get=lambda r: r['tema'], desc='Sistema o tema al que pertenece el supuesto.'),
     Col('Supuesto', 36, 'txt', get=lambda r: r['nombre'], desc='Nombre corto del supuesto.'),
     Col('Pregunta para los profesores', 62, 'txt', get=lambda r: r['preg'], desc='Decisión que corresponde a los profesores, en lenguaje claro.'),
     Col('Valor por defecto (PROVISIONAL)', 46, 'txt', get=lambda r: r['defecto'], desc='Opción usada en el escenario por defecto; es un supuesto de trabajo, no una decisión de los profesores.'),
     Col('Alternativas', 56, 'txt', get=lambda r: r['alts'], desc='Otras opciones razonables que se calcularon como escenarios.'),
     Col('Indicadores afectados', 30, 'txt', get=lambda r: r['afecta'], desc='Códigos de los indicadores (o variables) afectados por el supuesto.'),
     Col('Cuánto cambia el resultado (síntesis)', 60, 'txt', get=lambda r: r['sint'], desc='Rango de resultados municipales según el resumen de cálculos (o resumen derivado de los escenarios calculados).'),
     Col('Cuánto cambia con cada alternativa (nivel municipal)', 70, 'txt', get=lambda r: r['detalle'], desc='Una línea por alternativa: resultado municipal por defecto → con la alternativa. Detalle completo en la hoja Efecto de alternativas.'),
     Col('Estado', 26, 'txt', get=lambda r: r['estado'], desc='Siempre PROVISIONAL.'),
     Col('Clave del parámetro', 26, 'txt', get=lambda r: r['clave'], desc='Identificador técnico del parámetro (aparece en la clave de escenario).', llave='Llave: parámetro (dentro de la clave de escenario)'),
    ]


def cols_efectos():
    return [
     Col('Tema', 20, 'txt', get=lambda e: TEMA(e['p']), desc='Sistema o tema del supuesto.'),
     Col('Supuesto', 36, 'txt', get=lambda e: P[e['p']][0], desc='Supuesto que cambia.'),
     Col('Alternativa', 56, 'txt', get=lambda e: e['o'], desc='Opción alternativa frente al valor por defecto.'),
     Col('Código del indicador', 11, 'txt', get=lambda e: e['codigo'], desc='Indicador afectado.', llave='Llave: código del indicador'),
     Col('Indicador', 38, 'txt', get=lambda e: NOM_IND.get(e['codigo'], ''), desc='Nombre del indicador.'),
     Col('Año', 7, 'int', get=lambda e: int(e['anio']) if e['anio'] else None, desc='Año.', llave='Llave: año'),
     Col('Periodo', 11, 'txt', get=lambda e: e['periodo'], desc='Periodo.', llave='Llave: periodo'),
     Col('Resultado municipal por defecto', 14, 'val', get=lambda e: e['vd'], desc='Puntaje municipal en el escenario por defecto (o N/D, NE, Revisar).'),
     Col('Semáforo por defecto', 13, 'sem', get=lambda e: e['sd'], desc='Semáforo del escenario por defecto.'),
     Col('Resultado municipal con la alternativa', 14, 'val', get=lambda e: e['va'] if e['va'] is not None or not e['codigo'] else 'Sin cálculo', desc='Puntaje municipal con la alternativa (o N/D, NE, Revisar, Sin cálculo).'),
     Col('Semáforo con la alternativa', 13, 'sem', get=lambda e: e['sa'], desc='Semáforo con la alternativa.'),
     Col('¿Cambia el resultado municipal?', 11, 'txt', get=lambda e: e['cambia'], desc='Sí si cambia el puntaje, el estado o el color.'),
     Col('N.º de zonas cuyo resultado cambia', 11, 'int', get=lambda e: e['zonas'], desc='Cantidad de zonas (PP por zona, I2) cuyo resultado cambia con la alternativa.'),
     Col('Nota', 60, 'txt', get=lambda e: e['nota'], desc='Aclaración.'),
     Col('Clave de escenario', 36, 'txt', get=lambda e: f"{e['p']}={[k for k, v in P[e['p']][2].items() if v == e['o']][0]}", desc='Clave técnica del escenario evaluado.', llave='Llave: escenario'),
    ]


def cols_nocalc():
    return [
     Col('Sistema', 20, 'txt', 'sistema'), Col('Código del indicador', 11, 'txt', 'codigo', '', 'Llave: código del indicador'), Col('Indicador', 44, 'txt', 'nombre'),
     Col('Año', 7, 'int', 'anio', '', 'Llave: año'), Col('Periodo', 10, 'txt', 'periodo', '', 'Llave: periodo'),
     Col('Escenario', 40, 'txt', get=lambda r: desc_escenario(r['escenario_id'])), Col('Motivo por el que no se calcula', 70, 'txt', 'motivo_nd', 'Razón por la que el indicador no se calcula en esta etapa.'),
     Col('Variables construidas', 28, 'txt', get=lambda r: r['variables_usadas'].replace(';', '; '), desc='Variables que sí se construyeron (hoja Variables).'),
     Col('Observaciones', 50, 'txt', get=obs_res), Col('Estado', 12, 'txt', 'estado', 'Siempre Provisional.'),
     Col('Clave de escenario', 30, 'raw', 'escenario_id', '', 'Llave: escenario')]


# --------------------------------------------------------------------------- Léeme y diccionario
def leeme(n):
    return [
     ['AVISO IMPORTANTE: TODOS LOS RESULTADOS SON PROVISIONALES',
      'Ningún dato de este libro ha pasado por validación: las variables están en estado "Pendiente" (o "No disponible") y los resultados en estado "Provisional". '
      'Todos los resultados dependen de decisiones metodológicas que corresponden a los profesores Francisco Correa y Eber Gutiérrez (ver la hoja Supuestos): '
      'los valores "por defecto" son supuestos de trabajo propuestos por el equipo y adoptados de forma provisional, no decisiones de los profesores. '
      'Léanse como una primera aproximación, no como una evaluación de desempeño.'],
     ['Qué es este libro',
      'Base de resultados del cálculo provisional de los indicadores viables del Sistema de Indicadores del Observatorio Ciudadano de Envigado (OCE), con fecha de corte del 9 de noviembre de 2026. '
      'Contiene los resultados, las variables que los alimentan, el soporte de cada conteo, los supuestos que los condicionan y cuánto cambia cada resultado si se elige otra alternativa. '
      'Se calcularon los indicadores clasificados como "Viable" y "Viable con ajuste" para 2024 y 2025, y para el primer semestre de 2026 en presupuesto participativo.'],
     ['Fecha de corte',
      'Cálculo y consulta de fuentes: 9 de noviembre de 2026. Cortes de los datos: 31 de diciembre de 2024; 31 de diciembre de 2025 (algunas fuentes tienen corte real del 7 de enero de 2026, que por defecto se registra como 31 de diciembre) y 30 de junio de 2026 (primer semestre de 2026, sin cierre de vigencia).'],
     ['Contenido (número de filas)', '\n'.join(f'{k}: {v:,}'.replace(',', '.') for k, v in n.items())],
     ['Hoja Resultados',
      'Una fila por indicador × año × periodo × territorio × escenario. Primero están todas las filas del escenario por defecto (por sistema, indicador, año y territorio, con el Municipio antes de las zonas) y luego las alternativas. '
      'La columna "Escenario" describe en palabras qué supuestos cambian respecto al valor por defecto. Los indicadores que no se calculan en esta etapa están en la hoja No calculados.'],
     ['Hoja Variables', 'Valor de cada variable de entrada con su fuente (nombre del documento y enlace), página o tabla, fechas de corte, estado y escenario. Primero el escenario por defecto.'],
     ['Hoja PP por zona', 'Ejecución presupuestal del presupuesto participativo (EPP) por zona y año en el escenario por defecto, con apropiación (APP), obligaciones (OPP), porcentaje y semáforo, y tres alternativas relevantes en columnas. 2024 es aproximado; el primer semestre de 2026 no tiene semáforo.'],
     ['Hoja PP por proyecto', 'Proyectos de presupuesto participativo con apropiación, compromisos, obligaciones y porcentaje de ejecución, más la marca de "posible no PP" y el tipo de cruce de 2024.'],
     ['Hojas Soporte de conteos, Clasificación VRC, Clasificación RPO y Entregas PP',
      'Detalle auditable de cada elemento contado (zona, evento, observación, proyecto, meta) y de su clasificación. Las citas de ciudadanos y funcionarios están anonimizadas. La segunda lectura de una muestra del 20 % de VRC y RPO está pendiente; hasta entonces la clasificación es de un solo lector.'],
     ['Hojas Supuestos y Efecto de alternativas', 'Las 36 decisiones metodológicas (parámetros) que condicionan los resultados, con la pregunta para los profesores, el valor por defecto, las alternativas y cuánto cambia el resultado con cada una. La hoja Efecto de alternativas lo detalla por indicador y año a nivel municipal.'],
     ['Hoja Diccionario de datos', 'Para quien construirá el tablero: cada columna de cada hoja, con su tipo, descripción, valores permitidos y las llaves para unir hojas.'],
     ['Llaves para unir hojas', 'Código del indicador (y código de la variable), año, periodo, territorio y clave de escenario. La clave de escenario es "defecto" o la lista de cambios parámetro=opción, y aparece al final de las hojas que la usan. El ID de soporte une Variables con Soporte de conteos y las hojas de clasificación.'],
     ['Convenciones',
      'N/D: no se puede calcular con la información disponible o con la regla de la Matriz.\nNE: no evaluable (la Zona 1 no tiene proyectos de presupuesto participativo).\n'
      'Revisar: el numerador supera al denominador y la Matriz pide revisión; no tiene color.\nSin semáforo: resultado informativo del primer semestre de 2026, periodo que la Matriz no prevé; no tiene color.\n'
      'Provisional: estado de todos los resultados. Pendiente: estado de todas las variables con dato.\nEscenario por defecto: conjunto de supuestos de trabajo adoptados provisionalmente. '
      'Los demás escenarios cambian uno o varios supuestos.\nVRC: veedurías en rendición de cuentas. RPO: respuesta a observaciones ciudadanas. DSI: divulgación de la situación. PCS: promoción del control social. ICRCS: índice de control y rendición social. '
      'I2, I3, I5: indicadores de participación ciudadana. CA, CM: indicadores del PDM. EPP: ejecución presupuestal del presupuesto participativo. APP: apropiación definitiva. OPP: obligaciones. EFPC: Escuela de Formación para la Participación Ciudadana.'],
     ['Colores', 'Semáforo: verde, amarillo y rojo según los umbrales de la Matriz; gris para N/D y NE. Los semáforos se asignan sin redondear antes de clasificar.'],
     ['Alertas principales',
      '1. El EPP 2024 es aproximado: no hay plan de acción de 2024 y se cruzan 50 de 51 proyectos del listado con los rubros de Hacienda (35 exactos y 15 aproximados).\n'
      '2. La mortalidad vial usa una cifra citada de Medicina Legal (Forensis 2024) pendiente de verificar contra el documento.\n'
      '3. El DSI queda en Rojo por la ventana de publicación (hasta el 31 de marzo del año siguiente); con todo lo publicado hasta hoy, 2024 da 28,57.\n'
      '4. VRC e ICRCS son N/D por defecto porque en varios eventos la presencia de las veedurías es desconocida; cualquier valor numérico exige apartarse de la regla de la Matriz.\n'
      '5. Los rubros de déficit fiscal de 2024 (7 proyectos, $627.204.058) no se excluyeron por defecto: la decisión es de los profesores.\n'
      '6. En el primer semestre de 2026, el 23,26 compara medio año de ejecución con una apropiación anual.'],
     ['Versión', 'Versión 1.0 (provisional), 9 de noviembre de 2026. Preparado para los profesores Francisco Correa y Eber Gutiérrez.'],
    ]


# --------------------------------------------------------------------------- ensamblaje
def main():
    calc = sorted((r for r in RES if r['no_calculado_por_etapa'] == 'False'), key=orden_res)
    nocalc = sorted((r for r in RES if r['no_calculado_por_etapa'] == 'True'), key=orden_res)
    var = sorted(VAR, key=lambda r: (0 if r['es_defecto'] == 'True' else 1, ORDEN_SIS.index(r['sistema']), clave_cod(r['codigo_indicador']),
                                     r['codigo_indicador'], r['id_variable'], r['anio'], ORDEN_PER[r['periodo']],
                                     r['escenario_id'] if r['es_defecto'] != 'True' else '', clave_terr(r['territorio']), r['escenario_id']))
    cz, fz = build_pp_zona()
    cp, fp = build_pp_proyecto()
    sup = build_supuestos()
    efe = EFECTOS
    cuentas = {'Resultados': len(calc), 'No calculados': len(nocalc), 'Variables': len(var), 'PP por zona': len(fz), 'PP por proyecto': len(fp),
               'Soporte de conteos': len(SOP), 'Clasificación VRC': len(C_VRC), 'Clasificación RPO': len(C_RPO), 'Entregas PP': len(C_ENT),
               'Supuestos': len(sup), 'Efecto de alternativas': len(efe)}
    wb = Workbook()
    ws0 = gd.escribir_hoja(wb, 'Léeme', ['Tema', 'Contenido'], leeme(cuentas), [34, 120], primera=True)
    ws0.auto_filter.ref = None
    for c in ws0[2]:
        c.fill = PatternFill('solid', fgColor='F8D4D4')
        c.font = Font(bold=True, color='9C0006', size=12)
    hoja(wb, 'Resultados', cols_resultados(), calc, 'resultados_indicadores')
    hoja(wb, 'No calculados', cols_nocalc(), nocalc)
    hoja(wb, 'Variables', cols_variables(), var, 'valores_variables')
    hoja(wb, 'PP por zona', cz, fz)
    hoja(wb, 'PP por proyecto', cp, fp)
    hoja(wb, 'Soporte de conteos', cols_soporte(), SOP, 'soporte_conteos')
    hoja(wb, 'Clasificación VRC', cols_vrc(), C_VRC)
    hoja(wb, 'Clasificación RPO', cols_rpo(), C_RPO)
    hoja(wb, 'Entregas PP', cols_entregas(), C_ENT)
    hoja(wb, 'Supuestos', cols_supuestos(), sup)
    hoja(wb, 'Efecto de alternativas', cols_efectos(), efe)
    gd.escribir_hoja(wb, 'Diccionario de datos', ['Hoja', 'Columna', 'Tipo', 'Descripción', 'Valores permitidos', 'Llave para unir hojas'],
                     DICCIONARIO, [24, 40, 16, 70, 50, 34])
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    wb.save(SALIDA)
    print('Guardado', SALIDA)
    for k, v in cuentas.items():
        print(f'  {k}: {v}')
    print('  Diccionario:', len(DICCIONARIO))
    return cuentas


def verificar(cuentas):
    """Comprobaciones de integridad sobre el libro ya guardado."""
    wb = load_workbook(SALIDA)
    fallos = []

    def hoja_dict(nombre):
        ws = wb[nombre]
        h = [c.value for c in ws[1]]
        return h, [dict(zip(h, [c.value for c in fila])) for fila in ws.iter_rows(min_row=2)]
    # 1. filas por hoja
    for nombre, n in cuentas.items():
        _, f = hoja_dict(nombre)
        if len(f) != n:
            fallos.append(f'filas {nombre}: {len(f)} != {n}')
    esperado = {'Resultados': len(RES) - sum(r['no_calculado_por_etapa'] == 'True' for r in RES), 'No calculados': sum(r['no_calculado_por_etapa'] == 'True' for r in RES),
                'Variables': len(VAR), 'Soporte de conteos': len(SOP), 'PP por proyecto': len(PPY), 'Clasificación VRC': len(C_VRC),
                'Clasificación RPO': len(C_RPO), 'Entregas PP': len(C_ENT), 'Supuestos': 39}
    for k, v in esperado.items():
        if cuentas[k] != v:
            fallos.append(f'origen {k}: {cuentas[k]} != {v}')
    # 2. puntajes por defecto
    h, res = hoja_dict('Resultados')
    esp = {('CS.PCS', 2024): 100, ('CS.PCS', 2025): 100, ('CS.DSI', 2024): 0, ('CS.DSI', 2025): 0, ('CS.RPO', 2024): 94.12, ('CS.VRC', 2024): 'N/D',
           ('CS.ICRCS', 2024): 'N/D', ('PC.I2', 2024): 100, ('PC.I2', 2025): 92.31, ('PC.I3', 2025): 83.33, ('PC.I5', 2025): 100, ('PDM.CA', 2024): 100,
           ('PDM.CM', 2024): 100, ('PP.EPP', 2024): 98.69, ('PP.EPP', 2025): 98.83}
    for (cod, anio), v in esp.items():
        fila = [r for r in res if r['Código del indicador'] == cod and r['Año'] == anio and r['Territorio'] == 'Municipio' and r['Clave de escenario'] == 'defecto'][0]
        got = fila['Puntaje'] if fila['Puntaje'] is not None else fila['Estado del cálculo']
        if (isinstance(v, str) and got != v) or (not isinstance(v, str) and round(got, 2) != v):
            fallos.append(f'puntaje defecto {cod} {anio}: {got} != {v}')
    s1 = [r for r in res if r['Código del indicador'] == 'PP.EPP' and r['Año'] == 2026 and r['Territorio'] == 'Municipio' and r['Clave de escenario'] == 'defecto'][0]
    if round(s1['Puntaje'], 2) != 23.26 or s1['Semáforo'] != 'Sin semáforo':
        fallos.append(f'EPP 2026 S1: {s1["Puntaje"]} {s1["Semáforo"]}')
    # todos los puntajes contra el CSV
    csvs = {(r['codigo'], r['anio'], r['periodo'], r['territorio'], r['escenario_id']): r for r in RES}
    for r in res:
        o = csvs[(r['Código del indicador'], str(r['Año']), r['Periodo'], r['Territorio'], r['Clave de escenario'])]
        if (o['puntaje'] == '') != (r['Puntaje'] is None) or (o['puntaje'] != '' and abs(float(o['puntaje']) - r['Puntaje']) > 1e-9):
            fallos.append(f'puntaje difiere {r["Código del indicador"]} {r["Año"]} {r["Clave de escenario"]}')
    # el orden: primero todos los por defecto
    marcas = [r['¿Es el escenario por defecto?'] for r in res]
    if marcas != sorted(marcas, key=lambda x: x != 'Sí'):
        fallos.append('orden: los escenarios por defecto no van primero')
    # 3. textos: Validado, rutas locales, datos personales, hipervínculos
    n_enlaces = 0
    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for c in fila:
                if isinstance(c.value, str):
                    v = c.value
                    if re.search(r'\bValidado\b', v):
                        fallos.append(f'"Validado" en {ws.title}!{c.coordinate}')
                    if 'data_raw' in v or 'ruta_local' in v or re.search(r'(?<!\w)(scripts?|outputs)/', v):
                        fallos.append(f'ruta local en {ws.title}!{c.coordinate}')
                    sinurl = re.sub(r'https?://\S+', '', v)
                    ph = validar_sin_datos_personales(sinurl)
                    ph = [x for x in ph if not (x == 'celular' and re.fullmatch(r'[\d.\s$,-]+', sinurl.strip()))]
                    if ph and not re.fullmatch(r'[\d\s.,$-]+', sinurl.strip()):
                        # los montos de 10 dígitos son falsos positivos del detector de celulares
                        if not (ph == ['celular'] and re.search(r'\$\s?[\d.]{10,}', sinurl)):
                            fallos.append(f'posible dato personal ({ph}) en {ws.title}!{c.coordinate}: {sinurl[:60]}')
                if c.hyperlink:
                    n_enlaces += 1
                    u = c.hyperlink.target or ''
                    if not u.startswith('http') or ' ' in u or 'www.envigado.gov.co/sites' in u:
                        fallos.append(f'hipervínculo no válido en {ws.title}!{c.coordinate}: {u}')
    print(f'Verificación: {n_enlaces} hipervínculos; fallos: {len(fallos)}')
    for x in fallos[:40]:
        print('  FALLO', x)
    return fallos


if __name__ == '__main__':
    verificar(main())
