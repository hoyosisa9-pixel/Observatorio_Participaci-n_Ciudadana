#!/usr/bin/env python3
"""CS.RPO: respuesta a observaciones ciudadanas = 100 * OR / OT (OR ⊆ OT; OR > OT -> Revisar; denominador 0 -> NE; sin tope).

OT = observaciones ciudadanas registradas en los cuadros de preguntas y respuestas del universo (rpo_universo) y año, que pasan el
     filtro rpo_filtro_pdm (excluir_felicitaciones | tematico_pdm | todas).
OR = observaciones de OT con respuesta de fondo publicada al corte: rpo_criterio_fondo (estricto_rubrica | laxo_texto_respuesta) y
     ventana_documental_cs (guia_31mar_siguiente: respuesta publicada al 31-dic del año; hasta_fecha_consulta: cualquier fecha).
2025 solo se calcula con rpo_registro_incompleto=calcular_con_publicado; con 'nd' (defecto) el cuadro del Alcalde 2025 (3 intervenciones)
se considera incompleto y RPO = N/D.

Insumo: tabla de clasificación (una fila por observación) en parciales/CS/clasificacion_rpo.csv. Cada escenario se recalcula SOLO desde
las banderas de esa tabla (se lee de nuevo el CSV). Sin datos personales: las citas reemplazan nombres por [ciudadano/a] o [funcionario/a].
Todo PROVISIONAL. Ejecutar sin -I desde la raíz: python3 scripts/indicadores/cs_rpo.py
"""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import cs_segunda_lectura as SL

COD = "CS.RPO"
SCRIPT = "scripts/indicadores/cs_rpo.py"
AN = (2024, 2025)
D = dt.date
B2 = "data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/"
B3 = "data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/"
R = dict(
    QA24=B3 + "audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf",
    QA100=B2 + "preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf",
    ACS24B=B3 + "acta-audiencia-publica-de-rendicion-de-cuentas-2024/2024actarendicioncuentassecretariasalud02docx.pdf",
    C1=B3 + "rendicion-de-cuentas-del-comite-municipal-de-discapacidad/respuesta_rendicion_cuentas_alejandro_hernandez-sin-firma.pdf",
    C2=B3 + "rendicion-de-cuentas-del-comite-municipal-de-discapacidad/respuesta_rendicion_cuentas_carlos_velez-sin-firma.pdf",
    C3=B3 + "rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/respuesta_rendicion_cuentas_patricia_mejia.pdf",
    C4=B3 + "rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/respuesta_rendicion_cuentas_ricardo_gaviria.pdf",
    QA25=B2 + "rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf",
    QAS25=B2 + "preguntas-y-respuestas-rendicion-de-cuentas-secretaria/preguntasyrespuestasrendicioncuentassaludi2025.pdf",
    C5=B3 + "rendicion-de-cuentas-del-comite-municipal-de-discapacidad-961114/preguntasyrespuestas1rddiscapacidad.pdf",
)
CARTAS = {R["C1"], R["C2"], R["C3"], R["C4"]}      # nombres de personas en el nombre del archivo: se omiten

# grupo -> (ejercicio, fecha del evento, documento, universos a los que pertenece)
GRUPOS = {
    "QA24": ("Audiencia pública del Alcalde 2024", D(2024, 11, 26), "QA24: preguntas y recomendaciones, audiencia del Alcalde 2024", R["QA24"],
             ("alcalde", "alcalde_y_100dias", "todos_los_cuadros")),
    "QA100": ("Rendición de cuentas del Alcalde: primeros 100 días", D(2024, 4, 10), "QA100: preguntas y respuestas, 100 primeros días", R["QA100"],
              ("alcalde_y_100dias", "todos_los_cuadros")),
    "S2402": ("Rendición de cuentas Secretaría de Salud 2024-02", D(2024, 11, 26), "ACS24b: acta de la rendición de Salud 2024-02 (preguntas y respuestas en pp. 12-15)",
              R["ACS24B"], ("todos_los_cuadros",)),
    "CMD24": ("Rendición de cuentas Comité Municipal de Discapacidad 2024", None, "Cartas de respuesta del Comité Municipal de Discapacidad (nombres de personas omitidos)",
              None, ("todos_los_cuadros",)),
    "QA25": ("Rendición de cuentas del Alcalde 2025", None, "QA25: preguntas y respuestas de la rendición de cuentas 2025 del Alcalde", R["QA25"],
             ("alcalde", "alcalde_y_100dias", "todos_los_cuadros")),
    "QAS25": ("Rendición de cuentas Secretaría de Salud 2025-01", D(2025, 3, 19), "QAS25: preguntas y respuestas, Salud 2025-01", R["QAS25"],
              ("todos_los_cuadros",)),
    "CMD25": ("Rendición de cuentas Comité Municipal de Discapacidad, semestre 1 de 2025", D(2025, 7, 18),
              "Carta de respuesta del Comité Municipal de Discapacidad, semestre 1 de 2025 (nombre de persona omitido)", R["C5"], ("todos_los_cuadros",)),
}
B_INV = "fecha del inventario + orden de carpeta content/files (la fecha del inventario no es confiable por sí sola)"
OB = []


def ob(grupo, num, anio, pagina, cita, resp, tipo, fel, tram, tema, texto, acuse, est, laxo, fpub, base, pub31, dudoso=0, nota="",
       ruta=None, fecha_evento=None, sub=None):
    assert est <= laxo <= texto or (est == laxo == 0), (grupo, num)
    assert not (fel and texto and False)
    ej, fev, doc, ruta_g, _ = GRUPOS[grupo]
    OB.append(dict(grupo=grupo, num=num, anio=anio, ejercicio=ej, fecha_evento=fecha_evento or fev, documento=sub or doc,
                   ruta_local=ruta or ruta_g, pagina=pagina, cita=cita, cita_respuesta=resp, tipo_intervencion=tipo,
                   es_felicitacion=fel, es_tramite_individual=tram, tema_pdm=tema, tiene_texto_respuesta=texto,
                   respuesta_solo_acuse_o_remision=acuse, fondo_estricto=est, fondo_laxo=laxo, fecha_publicacion_respuesta=fpub,
                   base_fecha_pub=base, publicada_al_31dic=pub31, dudoso=dudoso, nota_dudoso=nota))


# --------------------------------------------------------------------------- QA24 (22 intervenciones, cuadro numerado 1-22)
P24, F24, BQ24 = 2024, D(2024, 12, 12), B_INV + " (carpeta 002595; copia duplicada del 12-dic-2024)"
def q24(n, pag, cita, resp, tipo, fel=0, tram=0, tema=1, texto=1, acuse=0, est=1, laxo=1, dudoso=0, nota=""):
    ob("QA24", n, P24, pag, cita, resp, tipo, fel, tram, tema, texto, acuse, est, laxo, F24, BQ24, 1, dudoso, nota)

q24(1, "1", "[...] ¿tendrá Envigado una casa del cuidado?", "No vamos a contar con una casa de cuidado en nuestro municipio, pero sí se cuenta con una estrategia [...] Hacer un análisis y un estudio.", "pregunta")
q24(2, "1-2", "[...] queremos saber cuándo se dotarán nuevamente a los niños del uniforme [...]", "El próximo año vamos a estar entregando uniformes a los 1.380 niños que tenemos en EFISAE.", "pregunta")
q24(3, "2", "Alcalde Dios lo bendiga mil y mil gracias por tanto apoyo con la comunidad, usted ha cumplido con todo lo que prometió.", "N/A", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)
q24(4, "2-3", "Señor Alcalde y equipo de trabajo quiero o queremos felicitarlo en todo lo que hacen, gracias a Dios por permitirle triunfar de nuevo.", "N/A", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)
q24(5, "3", "Es un honor estar en este lugar y ver como de tu mano Envigado va para adelante [...]", "N/A", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)
q24(6, "3", "Actualmente los andenes están en mal estado, ¿qué han pensado o qué acciones piensan tomar al respecto?", "El recurso [...] para el cuatrienio es de $4.950 millones para andenes [...] la meta [...] 20 mil metros cuadrados y esperamos cumplirlos.", "pregunta")
q24(7, "3-4", "Más que una sugerencia quiero hacer un agradecimiento al señor Alcalde por confiar en nosotros, los niños [...]", "N/A", "agradecimiento", fel=1, tema=0, texto=0, est=0, laxo=0)
q24(8, "4", "Para mí todo lo que el Dr. [funcionario/a] ha hecho en estos 10 meses ha sido muy maravilloso.", "N/A", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)
q24(9, "4-5", "Espero que para el próximo año [...] Sugiero a la medida de lo posible ayudar con mejoras de hogar para nuestras familias.", "Se da respuesta en el evento en la intervención de la Secretaria de Bienestar Social.", "sugerencia",
    acuse=1, est=0, laxo=1, dudoso=1, nota="La respuesta remite a lo dicho en el evento sin transcribirlo: remisión sin contenido (estricto 0); con la regla laxa cuenta como texto de respuesta.")
q24(10, "5-6", "Quisiera saber el estado actual del subsidio para cuidadores primarios.", "El subsidio para cuidadores primarios se piensa implementar para el próximo año a través de la Secretaria de Bienestar Social.", "pregunta")
q24(11, "6", "¿Qué acciones se están tomando para mejorar la movilidad, principalmente en el centro de la ciudad donde varias zonas se están convirtiendo en áreas de parqueo [...]?", "[...] el próximo año una actualización a las zonas de estacionamiento regulado ZER.", "pregunta")
q24(12, "7", "¿Cuál es el estado actual de la Política de Libertad Religiosa?", "La política se encuentra en una etapa de diagnóstico para un estudio de factibilidad, esperamos terminar esta etapa en febrero [...]", "pregunta")
q24(13, "7-8", "¿Cuál es la estrategia de recolección de basuras, reciclaje y manejo de los residuos sólidos?", "El municipio solo cuenta con una ruta selectiva [...] sancionatorio con 20 operativos en conjunto con la Policía con 20 comparendos [...]", "pregunta")
q24(14, "8", "¿Cómo quedó asignado el presupuesto para Juventud 2025?", "$1.955 millones, el presupuesto se reduce era de $2.300 millones [...]", "pregunta")
q24(15, "8-9", "¿Por qué en la rendición de cuentas no se habló nada de las acciones que ha realizado esta administración durante este año por la paz? [...]", "[...] acordamos convocar a la próxima sesión el día 27 de diciembre. El tema a trabajar serán los indicadores [...]", "pregunta")
q24(16, "9-13", "¿Qué acciones en términos de movilidad tomará el Municipio Envigado con el incremento de los parqueos en las calles [...]?", "[...] documento de formulación del Plan de Movilidad Sostenible y Segura PMSYS [...] el programa de ESTACIONAMIENTOS plantea reorganizar el espacio público de la zona céntrica [...]", "pregunta")
q24(17, "13-19", "¿Qué plan de trabajo o estrategia tiene la oficina de comunicaciones [...] ¿qué presupuesto manejó para ese plan de comunicaciones?", "Se transmitieron 11 programas con todas las zonas del municipio (menos la zona 4) [...] Llegando a una audiencia de más de 9.000 personas.", "pregunta",
    dudoso=1, nota="La respuesta da datos del plan de comunicación pero no responde la parte sobre el presupuesto; el tema (comunicación con los comités zonales) es discutible como 'sobre el PDM'.")
q24(18, "19-21", "¿Cuál fue la deuda real que recibió el actual gobierno y qué pasó con las obras de infraestructura [...]?", "El saldo de la deuda a 31 de diciembre de 2023 fue de $194.859.709.750,34. [...] no se han efectuado reparaciones en obras de infraestructura recientemente entregadas.", "pregunta",
    dudoso=1, nota="Tema de finanzas públicas e infraestructura; puede discutirse si es 'sobre el PDM'.")
q24(19, "21-22", "¿Qué acciones tomará el municipio de Envigado siendo socio de la empresa Metroplus en virtud de la irregular tardanza [...]?", "[...] ha solicitado tanto a la empresa encargada como al Ministerio de Transporte acelerar la finalización del tramo 12 en su territorio [...]", "pregunta")
q24(20, "22", "[...] qué tiene proyectado el municipio de Envigado [...] lo de la terminal del Sur de Metroplus, ¿qué pasará con ese lote?", "[...] la estación cabecera ya no se construirá en el lote denominado \"Casa de la Abuela\".", "pregunta")
q24(21, "23", "[...] ¿por qué la insistencia del departamento administrativo de planeación en no cobrar la plusvalía cuando es instrumento que está en la normatividad [...]?", "[...] durante este gobierno se llevarán a cabo los procesos necesarios para implementar este tributo. Además, se incluirá la plusvalía en el cambio del POT [...]", "pregunta",
    dudoso=1, nota="Compromiso sin fecha precisa ('durante este gobierno'); se leyó como decisión (estricto 1). Tema: instrumento de financiación del POT.")
q24(22, "23-24", "[...] ¿dónde se puede encontrar la calificación de desempeño de los funcionarios públicos del Departamento Administrativo de Planeación y la Subdirección de Participación Ciudadana?", "Las evaluaciones de desempeño [...] reposan en las respectivas historias laborales en la secretaria de talento humano.", "pregunta",
    tema=0, dudoso=1, nota="Gestión interna de talento humano: no remite a un programa o meta del PDM (tema_pdm 0). La respuesta indica dónde reposa la información (estricto 1).")

# --------------------------------------------------------------------------- QA100 (12 intervenciones por ciudadano/a)
F100, B100 = D(2024, 4, 30), B_INV + " (carpeta 002312)"
def q100(n, pag, cita, resp, tipo, tram=0, tema=1, est=1, laxo=1, acuse=0, dudoso=0, nota=""):
    ob("QA100", n, 2024, pag, cita, resp, tipo, 0, tram, tema, 1, acuse, est, laxo, F100, B100, 1, dudoso, nota)

q100(1, "1", "¿Cómo está la deuda en el municipio y como afecta el Plan de Desarrollo?", "[...] recibe una deuda de $194.000.000.000 este endeudamiento no afecta los planes y proyectos de este nuevo gobierno ya que contamos con una capacidad de endeudamiento adicional [...]", "pregunta")
q100(2, "1", "¿Por qué no se incentiva a pagar el impuesto predial con descuento si se paga la anualidad?", "[...] no recibir ese 5% [...] equivaldría a más de $4.500.000 que [...] el municipio necesita para sus finanzas [...] no vamos a dar el descuento del 5%.", "pregunta")
q100(3, "2", "¿Por qué no ha sido citado el Concejo Municipal de Paz y a su Comité Técnico para cumplir con sus funciones frente al plan de desarrollo municipal?", "El Comité Municipal de Paz ya está agendado con el señor alcalde [...] ese Concejo Municipal ya esta citado para los próximos días.", "pregunta",
     dudoso=1, nota="Compromiso 'para los próximos días' sin fecha exacta; se leyó como decisión (estricto 1).")
q100(4, "2-3", "¿Para Cuando el Parque del Inder, para cuando los CDI, para cuando los Andenes, para cuando las Basuras?", "[...] ya pudimos visitar [...] estos temas de seguridad ya están resueltos [...] la próxima semana también iniciaremos ya con unas clases [...]", "pregunta",
     dudoso=1, nota="Pregunta de cuatro componentes: responden parque, CDI y basuras (esta última en la respuesta siguiente); los andenes no tienen respuesta explícita.")
q100(5, "3", "¿Es viable que trabajen desde Enviaseo y Seguridad en la labor para darle control a las personas indisciplinadas en el manejo de las basuras, excretas y escombros?", "[...] a finales de este mes se realizarán operativos sancionatorios con la Secretaría de Seguridad y la Inspección Ambiental [...]", "pregunta")
q100(6, "3-4", "[...] el compromiso de la administración fue adecuar con la calidad requerida el vivero de Medio Ambiente como vivero Pedagógico-Comunitario ¿Qué paso entonces?", "En este momento no lo tenemos en funcionamiento [...] tenemos que esperar a que la Secretaría de Obras Públicas [...] desarrolle el acceso independiente [...]", "pregunta")
q100(7, "4", "¿Cuánto vale la transmisión de estos programas en vivo?", "Estas transmisiones por streaming tiene un valor de $1.300.000, implica el uso de tres cámaras [...]", "pregunta", tema=0,
     dudoso=1, nota="Costo del evento de rendición: pregunta institucional que no remite a un programa o meta del PDM (tema_pdm 0).")
q100(8, "4", "[...] cuanto estoy extrañando los paseos que nos hacen a diferentes sitios.", "Las salidas de día de sol, arrancarán en el segundo semestre [...] Vamos a hacer dos paseos al año.", "sugerencia")
q100(9, "5", "¿Nos pueden indicar a nivel de cultura cuando arrancan todos los programas que funcionaban por presupuesto participativo?", "[...] Red de Escuelas de Música: que da inicio desde marzo [...] Iniciativas Artísticas [...] estamos finalizando la fase precontractual [...]", "pregunta")
q100(10, "5", "[...] tengo mucha necesidad de hablar con el alcalde y soy emprendedora.", "Nos puede dejar sus datos y se le suministra contacto para que agende la cita, el alcalde es de puertas abiertas.", "tramite_individual",
     tram=1, tema=0, est=0, laxo=1, acuse=1, dudoso=1, nota="Solicitud de cita personal: la respuesta es una remisión sin contenido (estricto 0).")
q100(11, "5", "Podrá implementar donde nos capaciten en tema de cuidadores a personas que como yo, cuidamos a nuestros padres [...]", "[...] tenemos la técnica de Cuidadores, empezamos nuevamente para mayo las inscripciones del segundo semestre para iniciar en agosto [...]", "pregunta")
q100(12, "5", "Nuestra Parroquia [...] es patrimonio cultural, nuestra parroquia requiere mejoramiento, es muy visitada.", "[...] hay una negociación [...] vamos a hacer unas permutaciones [...] se haría una restauración de la iglesia [...]", "sugerencia")

# --------------------------------------------------------------------------- Salud 2024-02 (acta ACS24b): 5 preguntas + 3 comentarios
FS, BS = D(2024, 12, 27), B_INV + " (carpeta 002602)"
def s24(n, pag, cita, resp, tipo, fel=0, tram=0, tema=1, texto=1, acuse=0, est=1, laxo=1, dudoso=0, nota=""):
    ob("S2402", n, 2024, pag, cita, resp, tipo, fel, tram, tema, texto, acuse, est, laxo, FS, BS, 1, dudoso, nota)

s24(1, "12-14", "¿Continuidad del Programa de salud Mental?", "El programa de Salud Mental, continua con todas sus estrategias [...] continuara más fortalecido [...]", "pregunta")
s24(2, "14", "El programa para acompañamiento a los adultos mayores, inicie más rápido [...] los adultos mayores requieren un acompañamiento más continuo [...]", "[...] el proceso de contratación se iniciara en el mes de diciembre, con el fin de que pueda iniciar antes [...]", "pregunta")
s24(3, "14-15", "[...] implementar puntos de escuchaderos en los barrios dado que muchas personas no acceden a dicho servicio.", "[...] durante el año 2024, se realizó una prueba piloto [...] Se evaluara para el próximo año cual será la estrategia [...]", "sugerencia")
s24(4, "15", "Sobre la casa para personas adictas a sustancias psicoactivas de envigado como y cuando inicia.", "[...] el Centro de Atención de las Adicciones inicio desde el 18 de noviembre de 2024 [...]", "pregunta")
s24(5, "15", "[...] Que debo hacer y donde me debo dirigir.", "[...] a la fecha ya no se tienen cupos para el programa de Visión [...] será contactada para evaluar más a fondo su problemática [...]", "tramite_individual",
    tram=1, tema=0, dudoso=1, nota="Caso individual de acceso a servicios; la respuesta da orientación concreta (cupos, rango de Sisbén, contacto), por eso estricto 1.")
s24(6, "12", "La verdad me pareció un encuentro muy agradable, ya que toda la explicación de los distintos temas fue clara y concisa. [...]", "Sin respuesta (comentario).", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)
s24(7, "12", "Gracias por estos espacios es importante conocer cómo trabaja la secretaría y cómo invierte el presupuesto, muchas gracias.", "Sin respuesta (comentario).", "agradecimiento", fel=1, tema=0, texto=0, est=0, laxo=0)
s24(8, "12", "Excelente.", "Sin respuesta (comentario).", "felicitacion", fel=1, tema=0, texto=0, est=0, laxo=0)

# --------------------------------------------------------------------------- Cartas del Comité Municipal de Discapacidad 2024 (una carta = una intervención)
def c24(n, grupo_ruta, fev, pag, cita, resp, tipo, fpub, base, pub31, tram=0, tema=1, est=1, laxo=1, dudoso=0, nota="", sub=""):
    ob("CMD24", n, 2024, pag, cita, resp, tipo, 0, tram, tema, 1, 0, est, laxo, fpub, base, pub31, dudoso, nota, ruta=grupo_ruta,
       fecha_evento=fev, sub=sub)

c24(1, R["C1"], D(2024, 7, 30), "1-2", "¿Dónde se encuentra información respecto al censo actualizado de número de personas con discapacidad [...] mapa georreferenciado [...] inclusión laboral [...]?",
    "[...] mensualmente se extrae base de datos la cual contiene más de 51 variables [...] a la fecha no se tiene un mapa georreferenciado.", "pregunta", D(2024, 8, 2),
    "fecha de la carta (2-ago-2024); el inventario dice 22-jul-2024 (no confiable); carpeta 002460 anterior a la 002590 (3-dic-2024)", 1,
    sub="Carta de respuesta del CMD (semestre 1 de 2024), nombre de persona omitido")
c24(2, R["C2"], D(2024, 7, 30), "1-3", "Información acerca de la ejecución del presupuesto para población con discapacidad del año 2024 [...] Información sobre el proceso de Valoración de apoyos.",
    "[...] se le brindará la información presupuestal del primer trimestre [...] Total de la porción del valor ejecutado aplicado efectivamente a la PcD: $ 1.808.603.153 [...]", "pregunta", D(2024, 8, 2),
    "fecha de la carta (2-ago-2024); el inventario dice 22-jul-2024 (no confiable); carpeta 002460 anterior a la 002590 (3-dic-2024)", 1,
    dudoso=1, nota="Dos preguntas: la de presupuesto recibe datos del primer trimestre; la de valoración de apoyos recibe una definición legal y una remisión a la Personería.",
    sub="Carta de respuesta del CMD (semestre 1 de 2024), nombre de persona omitido")
c24(3, R["C3"], D(2024, 12, 17), "1-3", "¿Cómo acceder al subsidio de discapacidad? [...] ¿Cómo encontrar un empleo con mi técnica, tomando en cuenta mi discapacidad visual?",
    "El subsidio [...] se otorga de acuerdo al trámite establecido [...] no es posible asumir un compromiso que garantice la vinculación laboral [...]", "tramite_individual", D(2024, 12, 26),
    "fecha de la carta (26-dic-2024); el inventario dice 9-ene-2025 y la carpeta 002627 es posterior a la 002602 (27-dic-2024): publicación probablemente posterior al 31-dic-2024, no verificada", 0,
    tram=1, tema=0, dudoso=1, nota="Pregunta mixta (trámite del subsidio y caso laboral personal): se clasificó como trámite individual. Publicación posterior al 31-dic-2024 no verificada.",
    sub="Carta de respuesta del CMD (semestre 2 de 2024), nombre de persona omitido")
c24(4, R["C4"], D(2024, 12, 17), "1-3", "¿Cuál es la población con discapacidad a la fecha en Envigado [...] ¿En qué va el diseño de la Política Pública? [...] ¿Por qué no se ha vinculado a la veeduría 284 [...]?",
    "[...] es de 1.838 con corte al 30 de noviembre de 2024 [...] 0.74%. [...] La actualización de la Política Pública [...] iniciará en el año 2025 [...]", "pregunta", D(2024, 12, 26),
    "fecha de la carta (26-dic-2024); el inventario dice 9-ene-2025 y la carpeta 002627 es posterior a la 002602 (27-dic-2024): publicación probablemente posterior al 31-dic-2024, no verificada", 0,
    dudoso=1, nota="Publicación posterior al 31-dic-2024 no verificada (fuera de la ventana del 31-dic con la regla guia_31mar_siguiente).",
    sub="Carta de respuesta del CMD (semestre 2 de 2024), nombre de persona omitido")

# --------------------------------------------------------------------------- 2025
F25, B25 = D(2025, 12, 16), "Last-Modified HTTP 16-dic-2025 (notas_CS.md); el inventario dice 2025-09-16, inconsistente con el contenido (la respuesta cita el 10-dic-2025)"
def q25(n, pag, cita, resp, tipo, fel=0, tram=0, tema=1, texto=1, acuse=0, est=1, laxo=1, dudoso=0, nota=""):
    ob("QA25", n, 2025, pag, cita, resp, tipo, fel, tram, tema, texto, acuse, est, laxo, F25, B25, 1, dudoso, nota)

q25(1, "1", "Seguir con el cuidado, la protección, o cuidado ambiental y protección ciudadana [...] felicitaciones por sus gestiones y administración en salud aseo.", "Felicitaciones, por ende no se dio respuesta.", "felicitacion",
    fel=1, tema=0, texto=0, est=0, laxo=0, dudoso=1, nota="Felicitación con una sugerencia genérica ('seguir con el cuidado ambiental'); el propio cuadro la rotula felicitación.")
q25(2, "1-3", "Cuál es la razón de haber terminado con el festival de música religiosa un gran referente de envigado a nivel nacional y el exterior. [...]", "[...] la priorización de nuestra programación cultural responde directamente a los productos y metas establecidos en el Plan de Desarrollo Municipal [...] disponibilidad presupuestal anual [...]", "pregunta",
    dudoso=1, nota="La respuesta da una razón general (priorización por el PDM y presupuesto) sin decir expresamente que se canceló; se leyó como negativa motivada (estricto 1).")
q25(3, "4-5", "¿Qué acciones ha desarrollado el municipio para garantizar la inclusión educativa de personas con discapacidad especialmente estudiantes con autismo? [...]", "Se recomienda que la presente solicitud sea remitida a la oficina de bienestar institucional de la Institución Universitaria de Envigado [...]", "pregunta",
    acuse=1, est=0, laxo=1, dudoso=1, nota="La respuesta trata el caso particular y remite a la universidad; no responde la pregunta general sobre acciones del municipio. Se clasificó como pregunta temática (tema_pdm 1, trámite 0).")

FQ, BQ = D(2025, 4, 1), B_INV + " (carpeta 002690)"
def qs(n, pag, cita, resp, tipo, fel=0, tram=0, tema=1, texto=1, acuse=0, est=1, laxo=1, dudoso=0, nota=""):
    ob("QAS25", n, 2025, pag, cita, resp, tipo, fel, tram, tema, texto, acuse, est, laxo, FQ, BQ, 1, dudoso, nota)

qs(1, "1", "¿Qué acciones de vigilancia y control tiene la Secretaria de Salud con los hogares Geriátricos que operan en el municipio? [...]", "Desde la Dirección de Control de Factores de Riesgos [...] somos los responsables de las visitas de inspección sanitaria [...] se toman medidas sanitarias de seguridad [...]", "pregunta")
qs(2, "1", "Me parece una gestión muy interesante, y a su vez beneficia a toda la comunidad Envigado, y persona de escasos recursos.", "Agradecemos su participación en este espacio que contribuye al conocimiento de la gestión realizada por la Secretaría de Salud.", "felicitacion",
   fel=1, tema=0, texto=1, acuse=1, est=0, laxo=0)
qs(3, "1-3", "¿si los recursos de salud deben destinarse en un 60% a acciones de prevención y en un 40% a acciones de atención, cuantos negocios se han visitado en Envigado [...]? Deseo conocer registros", "[...] se deben aumentar las visitas [...] la tarea es de corresponsabilidad [...] desde el año 2024 se vienen realizando visitas a las farmacias [...]", "pregunta",
   est=0, laxo=1, dudoso=1, nota="La respuesta no entrega las cifras de visitas solicitadas ni un compromiso fechado (estricto 0); da explicación general (laxo 1).")
qs(4, "3", "Observación: Revisar los negocios callejeros", "Sin respuesta registrada en el acta.", "sugerencia", texto=0, est=0, laxo=0)
qs(5, "3", "Observación: Trabajar en equipo para fortalecer los procesos de pérdidas y desperdicio de alimentos. Y apoyo a tenderos [...]", "Sin respuesta registrada en el acta.", "sugerencia", texto=0, est=0, laxo=0)
qs(6, "3", "Observación: Control de basuras en ciertos sectores", "[...] reconoce la importancia de compartir las experiencias de los asistentes, en relación al manejo de las basuras con Enviaseo [...] se ofrece a propiciar esta reunión.", "sugerencia",
   est=0, laxo=1, dudoso=1, nota="Ofrecimiento de una reunión sin fecha: no es compromiso fechado (estricto 0).")

ob("CMD25", 1, 2025, "1-2", "Aún se encuentran muchos andenes en mucho deterioro [...] ¿Para cuándo? [...] Con las señalizaciones para los de movilidad reducida [...] ¿Qué hacen al respecto?",
   "[...] el municipio cuenta con un contrato destinado exclusivamente al mejoramiento y mantenimiento de andenes [...] no podrá cubrir las necesidades del municipio en un 100%.", "pregunta",
   0, 0, 1, 1, 0, 1, 1, D(2025, 7, 30), "fecha de la carta (30-jul-2025); el inventario dice 14-jul-2025 (no confiable); carpeta 002865 anterior a la 002890 (5-sep-2025)", 1,
   1, "Respuesta con dato sobre el contrato de andenes pero sin fecha para la pregunta '¿para cuándo?'; se leyó como respuesta de fondo (estricto 1).")


# --------------------------------------------------------------------------- derivación
def enlace_de(ruta):
    u = C.url_de_ruta(ruta)
    assert u, f"sin URL para {ruta}"
    return u.rsplit("/", 1)[0] + "/" if ruta in CARTAS else u


def derivar(o, k):
    f = dict(o)
    f["id"] = f"CS-RPO-{f['anio']}-{k:03d}"
    ruta = f["ruta_local"]
    f["enlace"] = enlace_de(ruta)
    if ruta in CARTAS:
        f["ruta_local"] = ruta.rsplit("/", 1)[0] + "/[archivo con nombre de persona omitido]"
    univ = GRUPOS[f["grupo"]][4]
    f["en_universo_alcalde"] = int("alcalde" in univ)
    f["en_universo_alcalde_y_100dias"] = int("alcalde_y_100dias" in univ)
    f["en_universo_todos_los_cuadros"] = 1
    f["en_ot_excluir_felicitaciones"] = int(not f["es_felicitacion"])
    f["en_ot_tematico_pdm"] = int(bool(f["tema_pdm"]) and not f["es_tramite_individual"])
    f["en_ot_todas"] = 1
    f["cuadro_registro_incompleto"] = int(f["grupo"] == "QA25")
    f["num_en_cuadro"] = f["num"]
    f["clasificacion"] = (f"{f['tipo_intervencion']} | tema_pdm={f['tema_pdm']} | trámite_individual={f['es_tramite_individual']} | "
                          f"respuesta: {'texto' if f['tiene_texto_respuesta'] else 'sin texto'}"
                          f"{' (acuse o remisión)' if f['respuesta_solo_acuse_o_remision'] else ''} | fondo estricto={f['fondo_estricto']}, laxo={f['fondo_laxo']}")
    reglas = []
    if f["es_felicitacion"]:
        reglas.append("felicitación o agradecimiento: sale de OT con excluir_felicitaciones y del filtro temático")
    if f["es_tramite_individual"]:
        reglas.append("trámite individual: sale del filtro temático")
    if not f["tema_pdm"] and not f["es_felicitacion"] and not f["es_tramite_individual"]:
        reglas.append("no remite a un programa, meta o sector del PDM: sale del filtro temático")
    if f["fondo_estricto"]:
        reglas.append("rúbrica estricta: la respuesta atiende el asunto con dato, decisión, compromiso o negativa motivada")
    elif f["fondo_laxo"]:
        reglas.append("rúbrica estricta no cumplida (acuse, remisión o respuesta sin dato ni compromiso fechado); la regla laxa sí cuenta el texto")
    else:
        reglas.append("sin texto de respuesta de fondo ni en la regla laxa")
    reglas.append("ventana: respuesta publicada al 31-dic" + (" (sí)" if f["publicada_al_31dic"] else " (no o no verificado)"))
    f["criterio"] = "; ".join(reglas)
    return f


COLS = ["id", "anio", "ejercicio", "num_en_cuadro", "fecha_evento", "documento", "pagina", "enlace", "cita", "cita_respuesta", "clasificacion",
        "criterio", "tipo_intervencion", "es_felicitacion", "es_tramite_individual", "tema_pdm", "tiene_texto_respuesta",
        "respuesta_solo_acuse_o_remision", "fondo_estricto", "fondo_laxo", "fecha_publicacion_respuesta", "base_fecha_pub",
        "publicada_al_31dic", "en_universo_alcalde", "en_universo_alcalde_y_100dias", "en_universo_todos_los_cuadros",
        "en_ot_excluir_felicitaciones", "en_ot_tematico_pdm", "en_ot_todas", "cuadro_registro_incompleto", "ruta_local", "dudoso", "nota_dudoso"]
COL_UNIV = {"alcalde": "en_universo_alcalde", "alcalde_y_100dias": "en_universo_alcalde_y_100dias", "todos_los_cuadros": "en_universo_todos_los_cuadros"}
COL_FILTRO = {"excluir_felicitaciones": "en_ot_excluir_felicitaciones", "tematico_pdm": "en_ot_tematico_pdm", "todas": "en_ot_todas"}
COL_FONDO = {"estricto_rubrica": "fondo_estricto", "laxo_texto_respuesta": "fondo_laxo"}


def evaluar(filas, p, anio):
    """OT/OR del escenario p y año, SOLO con las banderas de la tabla (texto)."""
    ot = [f for f in filas if int(f["anio"]) == anio and f[COL_UNIV[p["rpo_universo"]]] == "1" and f[COL_FILTRO[p["rpo_filtro_pdm"]]] == "1"]
    orr = [f for f in ot if f[COL_FONDO[p["rpo_criterio_fondo"]]] == "1"
           and (p["ventana_documental_cs"] == "hasta_fecha_consulta" or f["publicada_al_31dic"] == "1")]
    incompleto = anio == 2025 and p["rpo_registro_incompleto"] == "nd" and any(f["cuadro_registro_incompleto"] == "1" for f in ot)
    return dict(ot_ids=[f["id"] for f in ot], or_ids=[f["id"] for f in orr], OT=len(ot), OR=len(orr), incompleto=incompleto)


def main():
    orden = {2024: 0, 2025: 0}
    filas = []
    for o in OB:
        orden[o["anio"]] += 1
        filas.append(derivar(o, orden[o["anio"]]))
    assert len({f["id"] for f in filas}) == len(filas)
    for f in filas:
        for c in ("cita", "cita_respuesta", "criterio", "nota_dudoso", "clasificacion", "documento", "base_fecha_pub"):
            hall = C.validar_sin_datos_personales(f.get(c, ""))
            assert not hall, (f["id"], c, hall)
        assert len(f["cita"]) <= 200 and len(f["cita_respuesta"]) <= 200, (f["id"], len(f["cita"]), len(f["cita_respuesta"]))
    C.escribir_parcial("CS", "clasificacion_rpo", [{c: f.get(c) for c in COLS} for f in filas])
    tabla = C.leer_csv(C.ruta_parcial("CS", "clasificacion_rpo"))
    assert len(tabla) == len(filas)

    escs = C.escenarios(COD)
    defecto = C.escenario_defecto(COD)["params"]
    url_ev = C.normalizar_url(C.url_de_ruta(R["QA24"]))
    cuenta_def, or_def = {}, {}
    for a in AN:
        p0 = dict(defecto, rpo_registro_incompleto="calcular_con_publicado")     # solo para describir el elemento en el soporte
        r0 = evaluar(tabla, p0, a)
        for i in r0["ot_ids"]:
            cuenta_def[i] = a
        for i in r0["or_ids"]:
            or_def[i] = a
    soporte = []
    for f in filas:
        i = f["id"]
        soporte.append(dict(
            id_soporte=i, codigo_indicador=COD, id_variable="CS.RPO.OT;CS.RPO.OR", anio=f["anio"], escenario_id="defecto",
            elemento_tipo="observacion", elemento_id=i, descripcion=f"{f['ejercicio']}, intervención {f['num_en_cuadro']}: {f['tipo_intervencion']}",
            fecha_elemento=f["fecha_evento"], documento=f["documento"], ruta_local=f["ruta_local"], enlace=f["enlace"], pagina=f["pagina"],
            fecha_publicacion_doc=f["fecha_publicacion_respuesta"],
            banderas=(f"felicitacion={f['es_felicitacion']};tramite={f['es_tramite_individual']};tema_pdm={f['tema_pdm']};"
                      f"estricto={f['fondo_estricto']};laxo={f['fondo_laxo']};pub31dic={f['publicada_al_31dic']};dudoso={f['dudoso']}"),
            clasificacion=f["tipo_intervencion"],
            justificacion=f"{f['criterio']}. Cuenta en OR (defecto): {'sí' if i in or_def else 'no'}.",
            criterio="OT: observación del universo y año que pasa el filtro rpo_filtro_pdm; OR: con respuesta de fondo según rpo_criterio_fondo publicada dentro de la ventana (ver clasificacion_rpo.csv)",
            lector="ejecutor_CS", cuenta_en_defecto=int(i in cuenta_def), incluido="sí" if i in cuenta_def else "no"))

    valores, resultados, resumen = [], [], []
    for esc in escs:
        p = esc["params"]
        for a in AN:
            r = evaluar(tabla, p, a)
            if r["incompleto"]:
                ent = [None, None]
            else:
                ent = [r["OR"], r["OT"]] if r["OT"] > 0 else [0, 0]
            res = C.calcular(COD, ent, {"periodo": "Anual", "territorio": "Municipio"})
            if r["incompleto"]:
                res = dict(res, motivo="N/D: cuadro de preguntas del Alcalde 2025 publicado claramente incompleto (3 intervenciones); "
                                       "rpo_registro_incompleto=nd (un listado parcial no permite conocer OT)")
            corte = C.fecha_corte_registro(a, "Anual", D(a, 12, 31), p)
            obs_p = (f"Escenario {esc['escenario_id']}. "
                     + (f"OT/OR no se calculan: cuadro 2025 incompleto (publicado: OT = {r['OT']}, OR = {r['OR']}, solo informativo). "
                        if r["incompleto"] else f"OT = {r['OT']}; OR = {r['OR']}. ")
                     + "Clasificación y banderas en parciales/CS/clasificacion_rpo.csv. PROVISIONAL.")
            com = dict(tipo_valor="Construido", fecha_corte=corte, fecha_corte_real=D(a, 12, 31), fecha_publicacion=None, script=SCRIPT,
                       categoria_disponibilidad="Construible" if not r["incompleto"] else "Parcial", horizonte_meta=str(a),
                       documento_ruta_local="outputs/indicadores/parciales/CS/clasificacion_rpo.csv", enlace=url_ev,
                       pagina_tabla="Ver ids_soporte (clasificacion_rpo.csv y soporte_conteos.csv); cuadros de preguntas y respuestas (QA24, QA100, QA25, QAS25), acta de Salud 2024-02 y cartas del CMD",
                       fuente="Cuadros de preguntas y respuestas de las rendiciones de cuentas de la Alcaldía de Envigado, clasificados observación por observación",
                       observaciones=obs_p)
            vOR = None if r["incompleto"] else r["OR"]
            vOT = None if r["incompleto"] else r["OT"]
            valores.append(C.fila_valor("CS", "CS.RPO.OR", a, vOR, "Conteo", esc=esc,
                nombre_variable="Número de observaciones ciudadanas del universo evaluado con respuesta de fondo publicada al corte",
                n_elementos_soporte=len(r["or_ids"]), ids_soporte=",".join(r["or_ids"]) if not r["incompleto"] else "", **com))
            valores.append(C.fila_valor("CS", "CS.RPO.OT", a, vOT, "Conteo", esc=esc,
                nombre_variable="Número total de observaciones ciudadanas sobre el PDM registradas públicamente en ejercicios cerrados de rendición de cuentas del periodo",
                n_elementos_soporte=len(r["ot_ids"]), ids_soporte=",".join(r["ot_ids"]) if not r["incompleto"] else "", **com))
            resultados.append(C.fila_resultado(COD, a, esc, res, ent, nombre="Respuesta a observaciones ciudadanas (RPO)", tipo="Indicador",
                fecha_corte_comun=corte, script=SCRIPT, observaciones=obs_p))
            resumen.append((esc["escenario_id"], a, r, res))

    C.escribir_parcial("CS", "soporte_conteos", soporte, codigos=[COD])
    C.escribir_parcial("CS", "valores_variables", valores, codigos=[COD])
    C.escribir_parcial("CS", "resultados_indicadores", resultados, codigos=[COD])
    escritos = {(r["escenario_id"], str(r["anio"])): r for r in C.leer_csv(C.ruta_parcial("CS", "resultados_indicadores")) if r["codigo"] == COD}
    for eid, a, r, res in resumen:
        w = escritos[(eid, str(a))]
        assert w["semaforo"] == res["semaforo"] and w["disponibilidad"] == res["disponibilidad"], (eid, a)

    ruta_m, n_m, escrito = SL.muestra_20(tabla, COLS, "rpo")
    SL.escribir_leeme()
    print(f"clasificacion_rpo: {len(filas)} filas; muestra_rpo: {n_m} filas ({'escrita' if escrito else 'NO reescrita'})")
    for eid, a, r, res in resumen:
        pj = res["puntaje"]
        txt = "N/D (cuadro incompleto)" if r["incompleto"] else f"OR={r['OR']} OT={r['OT']}"
        print(f"{eid[:100]:<100} {a} {txt} -> {('%.1f' % pj) if pj is not None else res['disponibilidad']} {res['semaforo']}")
    print("dudosos:", [f["id"] for f in filas if f["dudoso"]])
    print(len(soporte), "soporte;", len(valores), "variables;", len(resultados), "resultados")


if __name__ == "__main__":
    main()
