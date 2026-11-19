"""Contenido de las fichas metodológicas y de los bloques «indicador de la matriz original» (entrega al profe).

Los textos de método están escritos aquí; todas las cifras (valores, líneas base, años, brechas, pesos) se leen de los CSV de
construcción y, donde el texto cita una cifra de método (referente, meta), `req()` comprueba que exista en el CSV del indicador.
Si los CSV se actualizan, al volver a correr los generadores todo se actualiza solo.
Lo usan generar_detalle_profe.py, generar_ficha_metodologica.py y generar_base_powerbi.py.
Sin números de decisión y sin atribuir decisiones a personas.
"""
import re

import comun_entregas as ce
from comun_entregas import (pp_anios_sin_dato, texto_csv, ESTADO, FECHA, INDICE_DIM, NOMBRE_DIM, NOMBRE_INDICE, NOMBRE_TABLERO, PRIMER_ANIO, SEMAFORO_TXT,
                            anios_con_indice, brecha_texto, dim_de, fmt, fuentes_lista, matriz, meses_fecha, num, para_profe,
                            peso_txt, principales, req, serie_indice_txt, serie_texto, unidad_sentido, una_linea, val, vf)

TERRI = 'https://terridata.dnp.gov.co/'
MASIVO = 'https://terridata.dnp.gov.co/assets/docs/txt/entidades/TerriData05266.txt.zip'
POB = ('Población total del municipio [personas]. DANE, proyecciones de población VP (8-ago-2025), publicadas en TerriData '
       f'(DNP), dimensión «Descripción general», indicador «Población total»: {TERRI}')
CAMPOS = [  # clave, rótulo (en el mismo orden en el Word y en el Excel)
    ('codigo', 'Código'), ('nombre', 'Nombre'), ('tipo', 'Tipo'), ('dimension', 'Dimensión'), ('objetivo', 'Objetivo'),
    ('definicion', 'Definición'), ('formula', 'Fórmula'), ('puntaje', 'Cálculo del puntaje de 0 a 100'),
    ('pesos', 'Pesos de los componentes (solo índices)'), ('regla', 'Regla de agregación (solo índices)'),
    ('anios', 'Años calculados (solo índices)'),
    ('variables', 'Variables (nombre completo, unidad y fuente)'), ('unidad', 'Unidad de medida'), ('sentido', 'Sentido'),
    ('periodicidad', 'Periodicidad'), ('cobertura', 'Cobertura geográfica'),
    ('referente', 'Línea base y referente o meta, con su justificación'), ('peso', 'Peso en el índice'),
    ('semaforo', 'Semáforo (propuesta)'), ('serie', 'Serie disponible (años)'), ('brecha', 'Brecha de información'),
    ('faltantes', 'Datos faltantes, rangos (mínimo y máximo) y casos especiales'), ('limitaciones', 'Limitaciones y alertas'),
    ('fuentes', 'Fuentes'), ('actualizacion', 'Fecha de actualización'), ('estado', 'Estado'),
]


def lista(*items):
    return '\n'.join('• ' + i for i in items if i)


# =====================================================================================================  INDICADORES
def _salud():
    c = 'PDM.SALUD'
    req(c, '100 % (cobertura universal')
    return dict(
        nombre='Cobertura en salud (afiliación al sistema de seguridad social en salud)',
        objetivo='Medir qué parte de la población de Envigado está afiliada al sistema general de seguridad social en salud.',
        definicion='Afiliados al sistema (régimen contributivo, régimen subsidiado y regímenes especiales y de excepción) como '
                   'porcentaje de la población total del municipio. Reemplaza al «aseguramiento al régimen subsidiado» de la matriz original.',
        formula='COB = 100 × (AC + AS + AE) / P',
        puntaje='Puntaje = mín(100; 100 × COB / 100) = mín(100; COB). El referente es la cobertura universal (100 %).',
        variables=lista('AC: afiliados al régimen contributivo [personas]. Ministerio de Salud, publicado en TerriData (DNP): ' + TERRI,
                        'AS: afiliados al régimen subsidiado [personas]. Misma fuente.',
                        'AE: afiliados a regímenes especiales y de excepción [personas]. Misma fuente.',
                        'P: ' + POB),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista(f'Línea base: no se fija una; la serie arranca en 2012 con {vf(c, 2012)} %.',
                        'Referente: 100 % (cobertura universal).',
                        'Justificación: la meta del PDM (99,5 %) es solo del régimen subsidiado, es decir, de otro universo. '
                        'El dato usa la población total; el referente debe ser del mismo universo.'),
        faltantes=lista('Sin años faltantes en 2012–2025.', 'El puntaje tiene tope de 100.',
                        '2017 es un año atípico: los afiliados del régimen contributivo caen frente a 2016 y se recuperan en 2018.'),
        limitaciones=lista('La afiliación mide aseguramiento, no acceso efectivo ni calidad del servicio.',
                           'El indicador «Afiliados al SGSSS» que publica TerriData excluye los regímenes especiales entre 2012 y 2017; '
                           'por eso el numerador suma los tres regímenes.',
                           'Las cifras de aseguramiento las reporta el Ministerio de Salud; no hay verificación local.',
                           'La meta del PDM no se usa porque es de otro universo.'))


def _ile():
    c = 'PDM.ILE'
    req(c, '19 %', '100,44 %')
    return dict(
        nombre='Logro educativo, ILE (atención a la primera infancia y cobertura neta)',
        objetivo='Combinar el logro en atención a la primera infancia y la cobertura neta en educación del municipio.',
        definicion='Promedio simple de dos puntajes: la cobertura de niñas y niños de 0 a 5 años atendidos en educación inicial y '
                   'la cobertura neta total en educación, cada una frente a su meta del PDM. Reemplaza el componente de Saber 11 de la matriz original.',
        formula='A = 100 × AT / P05\nLA = mín(100; 100 × A / 19)\nLN = mín(100; 100 × N / 100,44)\nILE = (LA + LN) / 2',
        puntaje='El ILE ya es un puntaje de 0 a 100: cada componente se limita a 100 antes de promediar.',
        variables=lista('AT: niñas y niños de 0 a 5 años atendidos en educación inicial [personas]. Hasta 2023, Ministerio de Educación Nacional '
                        '(datos.gov.co, conjunto 3y4s-dmxy, «Total educación inicial ICBF + MEN»; enlace: https://datos.gov.co/resource/3y4s-dmxy.csv). '
                        '2024 y 2025: Alcaldía de Envigado (balance del PDM 2024 de Educación y informe de gestión 2025).',
                        'P05: población de 0 a 5 años [personas]. DANE, proyecciones VP (PPED 2018–2042; retroproyección 2005–2017 para 2015–2017).',
                        f'N: cobertura neta total en educación [%]. Ministerio de Educación, publicada en TerriData: {TERRI}',
                        '19 % y 100,44 %: metas del PDM de cobertura inicial y de cobertura neta [%].'),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista('Referentes: 19 % de atención inicial y 100,44 % de cobertura neta, las metas del PDM (la matriz original trae MA = 19 %).',
                        'Edad: el PDM habla de 0 a 6 años; se usan 0 a 5 porque reproduce mejor la línea base del PDM.'),
        faltantes=lista('2025 sin dato: TerriData no publica la cobertura neta de 2025; no se sustituye la componente faltante (ILE queda sin dato).',
                        '2015–2018 se muestran con alerta de comparabilidad; no afectan el índice del PDM, que arranca en 2019.'),
        limitaciones=lista('Hasta 2018 el numerador es solo ICBF; desde 2019 es ICBF + preescolar del Ministerio de Educación: la serie cambia de composición.',
                           'La cobertura neta de TerriData salta de 70,34 (2014) a 95,88 (2015); la causa no está documentada.',
                           'En 2024 y 2025 el numerador sale de la Alcaldía (dato de gestión, no verificado con la fuente nacional) y cuenta solo niñas y niños, sin gestantes; '
                           'hasta 2023 incluía gestantes.',
                           'La matriz original proponía Saber 11, pero la clasificación de planteles no es pública.'))


def _sp():
    c = 'PDM.SP'
    req(c, '100 % (cobertura universal')
    return dict(
        nombre='Servicios públicos domiciliarios (acueducto, alcantarillado, aseo y energía)',
        objetivo='Medir la cobertura de los cuatro servicios públicos domiciliarios básicos en todo el municipio.',
        definicion='Promedio simple de las coberturas totales (urbana y rural) de acueducto, alcantarillado, aseo y energía eléctrica. '
                   'Reemplaza a la cobertura deportiva de la matriz original.',
        formula='SP = (CAC + CAL + CAS + CEN) / 4',
        puntaje='Puntaje = mín(100; 100 × SP / 100) = mín(100; SP).',
        variables=lista('CAC, CAL y CAS: cobertura de acueducto, alcantarillado y aseo [%]. Reportadas por la Alcaldía al Sistema Único de Información '
                        f'(Superservicios) y republicadas por TerriData: {TERRI}',
                        f'CEN: cobertura de energía eléctrica [%]. UPME, publicada en TerriData: {TERRI}'),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia), zonas urbana y rural.',
        referente=lista('Referente: 100 % (cobertura universal).',
                        'Justificación: no se usa la meta de la encuesta de calidad de vida porque es de otro universo; se compara cada cobertura con la cobertura universal.'),
        faltantes=lista('El indicador completo empieza en 2019 porque la energía total solo existe desde ese año.',
                        '2025 sin dato: TerriData no publica las coberturas de servicios públicos de 2025 en su archivo masivo.',
                        'Acueducto, alcantarillado y aseo tienen serie desde 2008; entre 2009 y 2018 se muestran solo como contexto, sin puntaje.'),
        limitaciones=lista('Entre 2016 y 2020 el reporte de cobertura rural es 0 y el urbano 100 en acueducto y alcantarillado; desde 2021 el rural va de 8 a 61 % '
                           'y el urbano de 87 a 92 %. La serie 2019–2020 no es comparable con la de 2021 en adelante (el 2019 se mantiene, con alerta).',
                           'La energía tiene un tope de la UPME: 100 % entre 2019 y 2022 y luego 96,27 y 95,63.',
                           'Las coberturas las reporta la propia Alcaldía al SUI; se solicita a Planeación la cobertura rural de 2016 a 2020.'))


def _cm():
    c = 'PDM.CM'
    req(c, '7.3301', '9,3459')
    return dict(
        nombre='Mortalidad vial (muertes por siniestros viales por 100.000 habitantes)',
        objetivo='Evaluar si la tasa anual de mortalidad por siniestros viales alcanza el máximo fijado por el plan.',
        definicion='Tasa anual de muertes por siniestros de tránsito por 100.000 habitantes, comparada con un referente del mismo universo.',
        formula='O = 100.000 × M / P',
        puntaje='Puntaje = mín(100; 100 × R / O), con R = 7,33. Si O = 0, el puntaje es 100.',
        variables=lista('M: muertes por homicidio culposo en accidente de tránsito [personas]. Policía Nacional, datos.gov.co, conjunto ha6j-pa2r '
                        '(2010–2025): https://www.datos.gov.co/resource/ha6j-pa2r.csv',
                        'P: ' + POB,
                        f'Para 2003–2009 el conjunto de la Policía no existe: se usa la tasa publicada por TerriData ({TERRI}), de la misma fuente original.',
                        'R: referente de 7,33 muertes por 100.000 habitantes (ver línea base y referente).'),
        periodicidad='Anual (año completo).', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista(f'Línea base: tasa 2019 de la Policía Nacional = {vf(c, 2019)} por 100.000.',
                        'Referente: 7,33 = 9,35 × (8 / 10,2), es decir, la reducción que fija el PDM (de 10,2 a 8) aplicada a la línea base 2019 de la Policía.',
                        'Justificación: la meta de 8 del PDM es de Movilidad, otro universo (otra fuente y definición); se aplica solo su reducción relativa '
                        'a la línea base del mismo universo del dato.'),
        faltantes=lista('Sin años faltantes en 2003–2025.', 'Si la tasa es 0, el puntaje es 100 sin dividir.',
                        'Control: la tasa propia redondeada a un decimal coincide con la de TerriData.'),
        limitaciones=lista('Mezclar fuentes (por ejemplo, Movilidad y la Policía) cambia la medida; por eso se usa una sola, la Policía Nacional.',
                           'Con pocas muertes por año (entre 11 y 43 desde 2010), la tasa varía mucho entre años.',
                           'El referente 7,33 es una construcción del observatorio, no una meta oficial publicada.'))


def _ca():
    c = 'PDM.CA'
    req(c, '4,54', '9 m2/hab')
    return dict(
        nombre='Espacio público verde urbano (metros cuadrados por habitante)',
        objetivo='Medir el avance de la disponibilidad de espacio público verde urbano por habitante hacia el referente del POT.',
        definicion='Metros cuadrados de espacio público verde urbano por habitante, y su avance desde la línea base de 2019 hacia 9 m² por habitante.',
        formula='CA = metros cuadrados de espacio público verde urbano / habitantes urbanos',
        puntaje='Avance = 100 × (CA − 4,54) / (9 − 4,54), acotado entre 0 y 100.',
        variables=lista('Espacio público verde urbano [m²]. 2019: AMVA con la Universidad Nacional (reporte ODS 11.7); 2022: AMVA con el Instituto Humboldt; '
                        '2024: diagnóstico de la revisión del POT de Envigado, tabla 39.',
                        'Habitantes urbanos [personas]. Según el estudio de cada corte (DANE).',
                        'Referente: 9 m² por habitante (diagnóstico del POT, tabla 39).'),
        periodicidad='No tiene periodicidad fija: son cortes de estudios (2019, 2022 y 2024). Se solicita a Planeación la medición anual.',
        cobertura='Zona urbana del municipio de Envigado.',
        referente=lista('Línea base: 2019 = 4,54 m² por habitante (AMVA con la Universidad Nacional).',
                        'Referente: 9 m² por habitante (diagnóstico del POT, tabla 39).',
                        'Justificación: la meta del PDM (3,88) no tiene fuente ni método declarados y daba 100 sin exigencia; se descartan los 5,00 de 2020 y 2023 por ser autorreporte.'),
        faltantes=lista('Solo hay tres cortes: 2019, 2022 y 2024. Los demás años quedan sin dato y se reportan en la brecha.',
                        'El verde total de 2024 (69,41 m² por habitante, con reservas rurales) solo se muestra como contexto.'),
        limitaciones=lista('Los cortes son de estudios distintos (productores, poblaciones y delimitación distintos): no deben leerse como una tendencia.',
                           'El 4,54 de 2019 no se reproduce con los datos del propio reporte (92,06 ha / 232.903 habitantes = 3,95).',
                           'Sensibilidad: con la población VP, el valor de 2024 sería 5,54 en lugar del 5,75 publicado.'))


def _veed():
    c = 'PDM.VEED'
    media = sum(num(val(c, a)) for a in range(2019, 2024)) / 5
    assert abs(media - 5.4) < 1e-9, media
    req(c, 'LB = 5,4', '291 y 292')
    return dict(
        nombre='Veedurías ciudadanas inscritas en la Personería (por año)',
        objetivo='Observar el número de veedurías ciudadanas que se inscriben cada año como expresión de la participación y el control ciudadano.',
        definicion='Número de veedurías ciudadanas inscritas en la Personería de Envigado en el año, contadas por la fecha de inscripción. '
                   'Reemplaza a «procesos participativos implementados» de la matriz original.',
        formula='V = número de veedurías inscritas en el año (por fecha de inscripción)',
        puntaje='Puntaje = mín(100; 100 × V / 5,4).',
        variables=lista('V: veedurías ciudadanas inscritas por año [conteo]. Personería de Envigado, datos.gov.co, conjunto jju5-br3w '
                        '(campos no personales): https://www.datos.gov.co/resource/jju5-br3w.csv',
                        'Consecutivos de inscripción [número]. Misma fuente, para calcular la brecha.'),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista('Línea base: promedio de 2019 a 2023 = 5,4 veedurías por año (2, 4, 11, 4 y 6).',
                        'Referente: la línea base (5,4).',
                        'Justificación: el PDM no fija meta; la línea base se toma de los años recientes con registro. '
                        'Si se contaran los faltantes seguros de 2019–2021, la línea base sería 7,8.'),
        faltantes=lista('El listado público es incompleto: los consecutivos van de 118 a 293 y el listado trae solo 45. Los años 2012–2018 también llevan puntaje, con la nota de listado incompleto.',
                        'Brecha de cada año: consecutivos que faltan dentro del año / (listadas + faltantes). Los consecutivos sin año atribuible se anotan aparte.',
                        '2026 está en curso y no se califica.'),
        limitaciones=lista('El puntaje refleja el listado publicado, no el total real de veedurías.',
                           'Los consecutivos 291 y 292 tienen año de registro 2026 y fecha de inscripción 2025; se cuentan por fecha de inscripción.',
                           'Una veeduría inscrita no implica control ciudadano efectivo.',
                           'La línea base depende del listado incompleto; se solicita el listado completo a la Personería.'))


def _hom():
    c = 'PDM.HOM'
    req(c, '2,7 por 100.000', '3,52')
    return dict(
        nombre='Homicidio (tasa por 100.000 habitantes)',
        objetivo='Medir la seguridad del municipio mediante la tasa de homicidios, como séptimo componente del índice del PDM.',
        definicion='Tasa anual de homicidios intencionales (con dolo), incluido el feminicidio, por 100.000 habitantes. '
                   'Es un componente nuevo respecto de la matriz original.',
        formula='O = 100.000 × H / P',
        puntaje='Puntaje = mín(100; 100 × 2,7 / O). Si O = 0, el puntaje es 100.',
        variables=lista('H: homicidios, incluido el feminicidio [personas]. Policía Nacional, datos.gov.co, conjunto m8fd-ahd9: https://www.datos.gov.co/resource/m8fd-ahd9.csv',
                        'P: ' + POB),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista(f'Línea base del PDM: 3,7 (2023); con la población VP la tasa de 2023 es {vf(c, 2023)}.',
                        'Meta: 2,7 por 100.000, la del PDM.',
                        'Justificación: la meta del PDM viene de la Policía Nacional, el mismo universo del dato.'),
        faltantes=lista('Sin años faltantes en 2003–2025.', 'Control: la tasa propia redondeada a un decimal coincide con la de TerriData en los 23 años.'),
        limitaciones=lista('La línea base 3,7 del PDM se calculó con la población posterior al COVID; con la población VP, la de 2023 es 3,52.',
                           'Con pocos casos por año (3 a 61), la tasa es sensible a variaciones pequeñas.',
                           'Mide homicidios registrados por la Policía, no percepción de seguridad ni otros delitos.'))


# ---- Participación ciudadana
def _i1():
    c = 'PC.I1'
    lb = num(val(c, 2017, 'valor'))
    assert abs(lb - 3.679) < 0.001
    req(c, 'LB 3,68', 'referente 7 %', '8.221/223.456', 'Decreto 242 de 2017, p. 31', '9.956')
    return dict(
        nombre='Votación del presupuesto participativo (IPCM)',
        objetivo='Medir qué parte de la población vota en la priorización participativa del presupuesto.',
        definicion='Índice de Participación Ciudadana Municipal (IPCM): votos de la priorización del presupuesto participativo como porcentaje de la población. '
                   'Reemplaza a «incorporación de aportes ciudadanos» de la matriz original.',
        formula='IPCM = 100 × VOT / P',
        puntaje='Puntaje = mín(100; máx(0; 100 × (IPCM − 3,68) / (7 − 3,68))).',
        variables=lista('VOT: votos (o votantes) de la votación del presupuesto participativo del año [personas]. Informe de priorización participativa de Planeación (2017), '
                        'documento «Proyectos priorizados 2021» (votación de 2020), noticias del portal de la Alcaldía (2021–2023) e informe de gestión 2025.',
                        'P: ' + POB),
        periodicidad='Anual, en cada votación. El año es el de la votación; los proyectos se ejecutan en el año siguiente (Acuerdo 038 de 2016, art. 15A).',
        cobertura='Municipio de Envigado (Antioquia).',
        referente=lista(f'Línea base: 2017 = {fmt(lb)} % (8.221 votos / 223.456 habitantes, población DANE VP). La cifra oficial de la Alcaldía es 3,62 %, con una población de referencia que no publica.',
                        'Referente: 7 %, el tope de la escala oficial de tramos del IPCM (informe de 2017, p. 6).',
                        'Justificación: el PDM 2024–2027 no fija meta para este indicador.'),
        faltantes=lista('2018: no hubo votación (vigencia 2019), el dato no aplica.',
                        f'{ce.gep.pc_anios_nd("PC.I1")}: la Alcaldía no publica la cifra de votos; quedan sin dato.',
                        f'La votación de 2020 (para la vigencia 2021) suma {ce.fmt(val("PC.I1", 2020, "numerador"))} votos, según «Proyectos priorizados 2021», p. 1; reemplaza los 5.100 votos parciales de una nota de prensa, que no se usan.',
                        'El puntaje tiene tope de 100 y piso de 0.'),
        limitaciones=lista(f'{ce.comprimir([a for a in range(2017, 2025) if a in ce.vals_i1()])} cuentan «votos» (según los documentos) y 2025 cuenta «votantes directos» (informe de gestión); la unidad no es idéntica.',
                           'Decreto 242 de 2017 (p. 31): el índice de participación por zona se define como población que participa / población total de la zona, lo que respalda usar la población total y no la de 14 años y más; '
                           '(p. 13) el voto exige acreditar residencia o domicilio en el municipio.',
                           'Pueden votar los mayores de 14 años, pero el IPCM oficial y la línea base usan la población total.',
                           'La votación se mide en el año t y los proyectos se ejecutan en t+1.',
                           'Se solicita a Planeación la serie 2018–2024 por zona, con la población de referencia y la unidad.'))


def _i2():
    c = 'PC.I2'
    req(c, '12 zonas', 'Acuerdo 007 de 2025', '93 fichas', 'sin actas de asistencia')
    return dict(
        nombre='Cobertura territorial de la participación',
        objetivo='Identificar el alcance territorial de las actividades participativas realizadas.',
        definicion='Porcentaje de zonas del municipio con al menos una actividad participativa realizada para su comunidad.',
        formula='I2 = 100 × Z / 12',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = I2.',
        variables=lista('Z: número de zonas con al menos una actividad participativa [conteo]. Informe de gestión 2025 de la Alcaldía; plan indicativo; listados del presupuesto participativo.',
                        '12: número de zonas del universo (zonas 2 a 13) [conteo].'),
        periodicidad='Anual.', cobertura='12 zonas del municipio de Envigado (zonas 2 a 13); la zona 1, industrial, queda fuera.',
        referente=lista('Referente: 12 de 12 zonas (100 %).',
                        'Justificación: la zona 1 es industrial y no tiene proyectos del presupuesto participativo ni comité zonal; contarla en el denominador sesgaría el indicador.'),
        faltantes=lista(f'{ce.gep.pc_anios_nd("PC.I2")} sin dato: no hay publicación de actividades participativas por zona.',
                        'La actividad debe permitir intervenir, proponer o deliberar (criterio de la matriz).'),
        limitaciones=lista('2020–2022: la cobertura es documental (12 de 12 zonas con fichas, programación, agenda de socialización y votación), no se mide la actividad efectiva; no hay actas de asistencia. '
                           'En 2020 hay 93 fichas de los espacios zonales de deliberación; en 2022, la agenda de socialización y votación por zonas (25 de agosto a 6 de septiembre) y los votos por zona de los planes zonales (las 12 zonas con votos; mínimo 146).',
                           'Marco normativo: Decreto 242 de 2017 (pp. 13 y 15: espacios zonales de divulgación y priorización) y Acuerdo 007 de 2025 (comités zonales de planeación participativa en las 13 zonas del municipio).',
                           '2025 = 12/12 es un agregado reportado por la Alcaldía (12 asambleas zonales «cubriendo la totalidad de las zonas»), sin lista por zona ni actas publicadas.',
                           'El respaldo de los proyectos de presupuesto participativo también da 12/12.',
                           'Se solicita la lista de asambleas por zona a la Subdirección de Participación Ciudadana.'))


def _i3():
    c = 'PC.I3'
    req(c, '21', '25')
    return dict(
        nombre='Cumplimiento de la oferta formativa',
        objetivo='Verificar la realización de la formación programada para la ciudadanía.',
        definicion='Porcentaje de acciones formativas (cursos, diplomados y seminarios completos) realizadas frente a las programadas para la ciudadanía.',
        formula='I3 = 100 × AR / AP',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = I3.',
        variables=lista('AR: acciones formativas realizadas en el año [conteo]. Informe de gestión 2025 de la Alcaldía, p. 88.',
                        'AP: acciones formativas programadas [conteo]. No hay programación publicada: se reconstruye sumando las realizadas y las no ejecutadas.'),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista('Referente: todas las acciones programadas (100 %).',
                        'Justificación: así lo define la matriz (acciones realizadas / programadas).'),
        faltantes=lista('Solo 2025 tiene tasa; 2020–2024 quedan sin dato porque no hay programación publicada.',
                        'Como respaldo se publica el conteo de acciones para la ciudadanía en 2023, 2024 y 2025.'),
        limitaciones=lista('El denominador de 2025 (25) es reconstruido: 21 realizadas (3 diplomados, 9 seminarios y 9 cursos) más 4 no ejecutadas.',
                           'Se cuenta el diplomado en Gestión Pública porque la fuente no dice que fuera solo para servidores.',
                           'Los semilleros no se cuentan.',
                           'Se solicita la programación anual a la Subdirección de Participación Ciudadana.'))


def _i4():
    c = 'PC.I4'
    req(c, '3 procesos fijos')
    return dict(
        nombre='Publicación de resultados participativos',
        objetivo='Verificar la devolución pública de los resultados de los procesos participativos.',
        definicion='Porcentaje de los tres procesos participativos de una lista fija (votación del presupuesto participativo, rendición de cuentas del Alcalde '
                   'y rendición de cuentas del presupuesto participativo) cuyos resultados están publicados y son consultables.',
        formula='I4 = 100 × PP / 3',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = I4. Al índice entra el mínimo del rango.',
        variables=lista('PP: procesos con resultados publicados y consultables en el portal [conteo]. Portal de la Alcaldía (www.envigado.gov.co): '
                        'https://www.envigado.gov.co/',
                        '3: procesos de la lista fija [conteo].'),
        periodicidad='Anual (año de la vigencia).', cobertura='Municipio de Envigado (Antioquia).',
        referente=lista('Referente: los 3 procesos con resultados publicados (100 %).',
                        'Justificación: no existe un inventario oficial de procesos participativos cerrados; el observatorio fija una lista de procesos recurrentes y '
                        'verifica cada año en el portal si se publicaron sus resultados (no es autorreporte).'),
        faltantes=lista('2025 se reporta como rango: el mínimo (2/3) entra al índice; el máximo (3/3) se alcanzaría si se publica la rendición del presupuesto participativo; brecha 1/3.',
                        '2020: no se encontró la rendición de cuentas del Alcalde.',
                        '2 de 18 proceso-año sin resultado publicado (ver la brecha).'),
        limitaciones=lista('La lista de tres procesos no cubre todos los procesos participativos del municipio.',
                           'La rendición del presupuesto participativo de la vigencia 2025 no está publicada; se hará en el último trimestre de 2026. '
                           'El plan indicativo marca avance 2026 = 1, aunque su observación dice que se realizará en el último trimestre (contradicción de la fuente).',
                           'Publicar no implica que se hayan incorporado los aportes.'))


# ---- Presupuesto participativo
_PP_LIM = ('Se excluye el rubro 2.3.7.06 (déficit fiscal, pagos de vigencias anteriores) de los proyectos de presupuesto participativo.')


def _sin_dato_pp():
    """Texto de los años de PP sin índice, tomado del CSV."""
    sd = pp_anios_sin_dato()
    if not sd:
        return 'Todos los años de 2020 a 2025 tienen índice.'
    return '; '.join(f'{a} queda sin dato ({m})' for a, m in sd) + '.'


def _contracredito():
    """Frase sobre los proyectos con contracrédito total de 2020, solo si el CSV la trae."""
    try:
        n, d = val('PP.CONTRACREDITO', 2020, 'numerador'), val('PP.CONTRACREDITO', 2020, 'denominador')
    except KeyError:
        return ''
    monto = re.search(r'\$([\d\.]+)', texto_csv('PP.CONTRACREDITO'))
    m = re.search(r'valor votado de los verificados = ([\d\.]+) %', texto_csv('PP.EPP'))
    return (f'En 2020, {n} de {d} proyectos tuvieron un contracrédito total' + (f' (${monto.group(1)})' if monto else '') +
            '; el EPP de 2020 no los refleja' + (f' (obligación / valor votado = {m.group(1).replace(".", ",")} %)' if m else '') + '.')


def _universo_2021():
    return ce.gep.pp_universo_2021().replace('(publicada', '(publicada').replace('2021: el universo', 'En 2021 el universo') + (
        ' Hay un cálculo de sensibilidad en las filas PP.SENS2021 de la hoja «Datos».' if ce.gep.pp_universo_2021() else '')


def _pp_comun():
    return ['Los cruces entre el valor votado y la apropiación de Hacienda dudosos cuentan como «no verificados»: están en el denominador de EJ90 y FID, '
            'no suman al numerador y quedan fuera de EPP.',
            'Un año se calcula solo si se verifica al menos el 50 % de sus proyectos. ' + _sin_dato_pp()]


def _epp():
    c = 'PP.EPP'
    return dict(
        nombre='Ejecución presupuestal del presupuesto participativo (EPP)',
        objetivo='Medir qué proporción de los recursos disponibles del presupuesto participativo se convirtió en obligaciones presupuestales reconocidas.',
        definicion='Obligaciones de los proyectos de presupuesto participativo como porcentaje de su apropiación definitiva.',
        formula='EPP = 100 × OPP / APP',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = EPP.',
        variables=lista('OPP: valor de las obligaciones presupuestales de los proyectos verificados [pesos corrientes (COP)]. Ejecución de egresos de Hacienda al cierre de cada vigencia y, para 2025, Plan de acción IV trimestre.',
                        'APP: valor de la apropiación definitiva de los mismos proyectos [COP]. Misma fuente.',
                        'Universo: proyectos votados, según el listado del presupuesto participativo (EFPC): https://efpc.envigado.gov.co/'),
        periodicidad='Anual (cierre de la vigencia; la matriz admite corte trimestral si hay reporte).',
        cobertura='Municipio de Envigado: proyectos votados en las zonas del presupuesto participativo.',
        referente=lista('Referente: 100 % de los recursos convertidos en obligaciones.', 'Justificación: es la definición de la matriz; no hay otra meta publicada.'),
        faltantes=lista(_PP_LIM, *_pp_comun(), 'Se excluye el proyecto 0553 («costo del proceso»).'),
        limitaciones=lista(_contracredito(), ce.gep.pp_rubros24(),
                           'La vigencia 2020 se priorizó en 2019; la que no se votó fue la vigencia 2019. El «valor votado» de 2020 es en realidad un valor priorizado.',
                           _universo_2021(),
                           'El 2025 sale del Plan de acción del IV trimestre y no de un PDF de Hacienda.',
                           'La apropiación definitiva puede haberse ajustado durante el año; el indicador no mide la calidad del gasto.'))


def _ej90():
    return dict(
        nombre='Proyectos con ejecución mayor o igual a 90 % (EJ90)',
        objetivo='Medir la proporción de proyectos votados que alcanzan una ejecución presupuestal de al menos 90 %.',
        definicion='Proyectos de presupuesto participativo cuya obligación es al menos 90 % de su apropiación definitiva, como porcentaje de los proyectos votados. '
                   'Reemplaza a «cumplimiento de entrega de proyectos» de la matriz original.',
        formula='EJ90 = 100 × PE90 / PV',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = EJ90.',
        variables=lista('PE90: proyectos con obligación / apropiación definitiva ≥ 90 % [conteo]. Hacienda (ejecución de egresos) y Plan de acción IV trimestre 2025.',
                        'PV: proyectos votados [conteo]. Listado del presupuesto participativo (EFPC): https://efpc.envigado.gov.co/'),
        periodicidad='Anual.', cobertura='Municipio de Envigado: proyectos votados en las zonas del presupuesto participativo.',
        referente=lista('Referente: todos los proyectos votados con ejecución de al menos 90 % (100 %).',
                        'Justificación: no hay actas de recibo ni «terminado» verificable (el avance físico es autorreporte); la ejecución presupuestal sí es verificable en Hacienda. '
                        '«Terminados» queda solo como respaldo, con advertencia.'),
        faltantes=lista(_PP_LIM, *_pp_comun()),
        limitaciones=lista('La ejecución presupuestal no equivale a entrega ni a satisfacción con el proyecto.',
                           'El umbral de 90 % es una convención del observatorio.',
                           'En 2020, los proyectos con contracrédito total cuentan como no ejecutados.'))


def _fid():
    return dict(
        nombre='Fidelidad a lo votado (FID)',
        objetivo='Medir si los proyectos que se ejecutan corresponden a lo que la ciudadanía votó.',
        definicion='Proyectos cuya apropiación definitiva (sin el rubro de déficit fiscal) está a ±1 % del valor votado, como porcentaje de los proyectos votados. '
                   'Reemplaza a «cumplimiento de metas de atención» de la matriz original.',
        formula='FID = 100 × PF / PV',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = FID.',
        variables=lista('PF: proyectos con apropiación definitiva a ±1 % del valor votado [conteo]. Hacienda y Plan de acción IV trimestre 2025.',
                        'PV: proyectos votados [conteo]. Listado del presupuesto participativo (EFPC): https://efpc.envigado.gov.co/'),
        periodicidad='Anual.', cobertura='Municipio de Envigado: proyectos votados en las zonas del presupuesto participativo.',
        referente=lista('Referente: todos los proyectos con apropiación igual a lo votado (100 %).',
                        'Justificación: el plan de acción mide actividades, no personas atendidas, así que no permite calcular metas de atención; la fidelidad compara lo votado con lo apropiado.'),
        faltantes=lista(_PP_LIM, *_pp_comun(),
                        'La llave de cruce es: valor votado = apropiación inicial de Hacienda; el nombre solo confirma.'),
        limitaciones=lista('Se mide la fidelidad monetaria, no el contenido del proyecto.',
                           'Algunos cruces dependen de nombres y sufijos de proyecto; los dudosos cuentan como no fieles.',
                           'En 2020 el valor votado es en realidad el valor priorizado.'))


# ---- Control social
def _pqrsdf():
    c = 'CS.PQRSDF'
    req(c, '17-may-2022', '+64 %', '90.37')
    return dict(
        nombre='Oportunidad de respuesta a las PQRSDF',
        objetivo='Medir qué proporción de las peticiones, quejas, reclamos, sugerencias, denuncias y felicitaciones se atiende dentro de los términos de ley.',
        definicion='PQRSDF atendidas en términos como porcentaje de las recibidas, según los informes trimestrales de la Alcaldía. '
                   'Reemplaza a «cumplimiento de la promoción del control social» de la matriz original, que era un autorreporte del 100 %.',
        formula='PQRSDF = 100 × ET / R',
        puntaje='El indicador ya es un puntaje de 0 a 100 (tope 100): puntaje = PQRSDF. Rango: máximo = 100 × (ET + PT + SR) / R.',
        variables=lista('ET: PQRSDF atendidas en términos [conteo]. Informes trimestrales de PQRSDF de la Secretaría General (sistema BPMS hasta el primer trimestre de 2025 y Débora desde mayo de 2025).',
                        'R: PQRSDF recibidas [conteo]. Misma fuente.',
                        'PT: pendientes dentro del término [conteo]; SR: «sin respuesta» (solo 2020–2021) [conteo]. Misma fuente; entran al máximo del rango.'),
        periodicidad='Trimestral (informes); se agrega por año con informes que cubren los 12 meses sin repetirse.',
        cobertura='Municipio de Envigado (Alcaldía).',
        referente=lista('Referente: todas las PQRSDF atendidas en términos (100 %).', 'Justificación: es la definición de la matriz; el término lo fija la ley (Ley 1755 de 2015).'),
        faltantes=lista('Serie completa: 24 de 24 trimestres leídos (2023 T1 se lee a mano porque el alcance del informe trae un error de copia).',
                        f'Rango: la columna «sin respuesta» de 2020–2021 y las pendientes en término van al máximo; el máximo de 2020 es {vf("CS.PQRSDF.MAX", 2020)} y el de 2021 {vf("CS.PQRSDF.MAX", 2021)}.',
                        'Si hay informes duplicados del mismo periodo con cifras distintas, se usa el más reciente.'),
        limitaciones=lista('Decreto 491 de 2020: amplió los plazos hasta el 17-may-2022; 2020, 2021 y 2022 T1 se miden con plazos ampliados, 2022 T2 es mixto y desde 2022 T3 con plazos ordinarios. No comparar sin esta salvedad.',
                           'En 2024 las recibidas suben 64 % (de 19.502 a 31.970) por reclasificación de lo que ingresa; no debe leerse como tendencia.',
                           'El sistema Débora rige desde mayo de 2025.',
                           'Es un registro administrativo de la Alcaldía; subestima la oportunidad real porque hay PQRSDF respondidas a tiempo pero cerradas tarde en el sistema (lo confirma una auditoría de Control Interno).'))


def _dsi():
    c = 'CS.DSI'
    req(c, '10 metas')
    return dict(
        nombre='Disponibilidad de información de seguimiento (DSI)',
        objetivo='Identificar metas del plan con cifras públicas utilizables para vigilarlo.',
        definicion='Metas de una lista fija de 10 metas de resultado del PDM 2024–2027 cuyo valor observado se publicó, con fecha verificable, a más tardar el 31 de marzo del año siguiente, '
                   'como porcentaje de las 10 metas.',
        formula='Mínimo: DSI = 100 × MV / 10\nMáximo: DSI = 100 × (MV + MN) / 10',
        puntaje='El indicador ya es un puntaje de 0 a 100: al índice entra el mínimo; el máximo suma las metas no verificables.',
        variables=lista('MV: metas con valor observado publicado y fecha de publicación verificada en la ventana (≤ 31 de marzo del año siguiente) [conteo]. Fuentes que nombra el PDM '
                        '(TerriData, Ministerio de Educación en datos.gov.co, Policía Nacional, informes de gestión).',
                        'MN: metas cuyo valor no se puede verificar (fuera de ventana o sin fecha) [conteo]. Misma fuente.',
                        'Lista de 10 metas: Plan de Desarrollo 2024–2027, Apéndice 9.1.2 (plan indicativo de resultados).'),
        periodicidad='Anual (corte con ventana hasta el 31 de marzo del año siguiente).', cobertura='Municipio de Envigado (PDM 2024–2027).',
        referente=lista('Referente: las 10 metas con valor publicado dentro de la ventana (100 %).',
                        'Justificación: la matriz propone empezar con 10 metas; se actualiza la lista a las metas del PDM 2024–2027 ligadas a los siete componentes del ICPD-7.'),
        faltantes=lista('Solo aplica en 2024 y 2025: la lista de metas es del PDM 2024–2027; 2020–2023 quedan como «no aplica».',
                        f'Rango: 2024 = {vf(c, 2024, "valor")}–{vf("CS.DSI.MAX", 2024)} y 2025 = {vf(c, 2025, "valor")}–{vf("CS.DSI.MAX", 2025)}; la diferencia es la brecha.'),
        limitaciones=lista('El DSI mide la oportunidad de la información: las fuentes nacionales con rezago (encuesta de calidad de vida, Ministerio de Educación, Policía) fallan por diseño aunque el dato exista.',
                           'Hoy se publica poco con fecha verificable.',
                           'La selección de 10 metas no representa estadísticamente todo el plan.'))


def _vrc():
    c = 'CS.VRC'
    req(c, 'desconocida')
    return dict(
        nombre='Veedurías en la rendición de cuentas (VRC)',
        objetivo='Medir la presencia documentada de veedurías en los eventos de rendición de cuentas.',
        definicion='Eventos de rendición de cuentas con presencia documentada de veedurías como porcentaje de los eventos del universo (audiencia del Alcalde y rendiciones sectoriales con acta o memoria).',
        formula='Mínimo: VRC = 100 × ED / ET\nMáximo: VRC = 100 × (ED + EN) / ET\nBrecha = 100 × EN / ET',
        puntaje='El indicador ya es un puntaje de 0 a 100: al índice entra el mínimo; EN son los eventos sin información sobre veedurías.',
        variables=lista('ED: eventos con veeduría documentada [conteo]. Actas, memorias y cuadros de preguntas publicados por la Alcaldía.',
                        'EN: eventos sin información sobre veedurías [conteo]. Misma fuente.',
                        'ET: eventos del universo [conteo]. Misma fuente.'),
        periodicidad='Anual (eventos realizados en la vigencia).', cobertura='Municipio de Envigado (Alcaldía).',
        referente=lista('Referente: todos los eventos con veeduría documentada (100 %).', 'Justificación: es la definición de la matriz.'),
        faltantes=lista('Rango: el mínimo cuenta como no documentado lo que no se sabe; el máximo lo cuenta como documentado; la diferencia es la brecha de información. Cambia la regla de la matriz, que lo trataba como N/D.',
                        '2020 sin dato (sin eventos clasificados).',
                        'La brecha de cada año se muestra en la serie del indicador.'),
        limitaciones=lista('«Con veeduría» solo si el documento lo dice expresamente; una invitación no acredita asistencia.',
                           'La audiencia del Alcalde no publica lista de asistentes.',
                           '2021–2023 solo incluyen la audiencia del Alcalde y las actas de Salud; el universo no es el mismo de 2024–2025 (Discapacidad y Víctimas).',
                           'Para seguir midiéndolo se necesitan listas de asistencia con una casilla de veeduría.'))


def _rpo():
    c = 'CS.RPO'
    req(c, '2 observaciones')
    return dict(
        nombre='Respuesta pública a observaciones ciudadanas (RPO)',
        objetivo='Observar la devolución institucional publicada a las observaciones ciudadanas hechas en la rendición de cuentas.',
        definicion='Observaciones de la audiencia pública del Alcalde (sin felicitaciones) con respuesta de fondo publicada dentro de la ventana (hasta el 31 de diciembre, con búsqueda hasta el 31 de marzo del año siguiente), '
                   'como porcentaje de las observaciones.',
        formula='RPO = 100 × OR / OT',
        puntaje='El indicador ya es un puntaje de 0 a 100: puntaje = RPO. Cuando la fecha de publicación no se puede probar, se reporta un rango.',
        variables=lista('OR: observaciones con respuesta de fondo publicada en la ventana [conteo]. Cuadro de preguntas y respuestas de la audiencia pública de rendición de cuentas del Alcalde (portal de la Alcaldía).',
                        'OT: observaciones registradas, sin felicitaciones [conteo]. Misma fuente.'),
        periodicidad='Anual.', cobertura='Municipio de Envigado (Alcaldía).',
        referente=lista('Referente: todas las observaciones con respuesta de fondo (100 %).',
                        'Justificación: es la definición de la matriz; un acuse de recibo o una remisión no cuenta como respuesta de fondo.'),
        faltantes=lista('2020 y 2022 sin dato: 2020, no se encontró la rendición del Alcalde; 2022, el documento trae solo respuestas, sin la lista de preguntas (no hay denominador).',
                        '2021 se reporta como rango de 0 a 100: la respuesta se publicó, pero no está probado que fuera antes del 31 de diciembre de 2021; daría 11/13 = 84,62 si se confirma la fecha.',
                        '2023 = 0 de 6 dentro de la ventana.'),
        limitaciones=lista('2025 = 1 de 2 con alerta: el cuadro publicado es corto (3 intervenciones frente a 537 asistentes) y puede no registrar todas.',
                           'Los cuadros son los que publica la Alcaldía; no se sabe si registran todas las intervenciones.',
                           'La serie depende de la fecha de publicación del documento, que no siempre se puede verificar.'))


# =====================================================================================================  ÍNDICES
def _indice(cod):
    dim = dim_de(cod)
    a_idx = anios_con_indice(cod)
    comunes = ['No se redistribuyen pesos: si falta un componente, el índice de ese año queda sin dato y se anota en la brecha.',
               'Los componentes usan fuentes y poblaciones distintas; ver la ficha de cada uno.']
    base = dict(unidad='puntaje 0–100', sentido='mayor es mejor', periodicidad='Anual.', cobertura='Municipio de Envigado (Antioquia).',
                peso='No aplica (es el índice de la dimensión).',
                fuentes='Cálculo del observatorio a partir de los componentes; las fuentes están en la ficha de cada componente.')
    if cod == 'PDM.ICPD7':
        a = dict(
            nombre='Índice de cumplimiento del PDM (ICPD-7)',
            objetivo='Sintetizar en un solo puntaje el avance de siete resultados seleccionados del PDM.',
            definicion='Promedio simple de los puntajes de siete componentes: cobertura en salud, logro educativo, servicios públicos domiciliarios, mortalidad vial, '
                       'espacio público verde, veedurías inscritas y homicidio. Reemplaza al ICPD-6 de la matriz original.',
            formula='ICPD-7 = (SALUD + ILE + SP + CM + CA + VEED + HOM) / 7',
            puntaje='Cada componente es un puntaje de 0 a 100; el índice es su promedio simple y se redondea solo el resultado.',
            variables=lista('Siete puntajes de 0 a 100 [puntaje]: SALUD, ILE, SP, CM, CA, VEED y HOM; cada uno con su ficha.'),
            referente=lista('Peso: 1/7 por componente. La matriz original tenía dos subíndices de tres componentes (ICPD-6); se pasa a un promedio simple de siete para conservar la regla de igual peso por componente.',
                            'Cada componente se compara con su propio referente (ver su ficha); no hay un referente del índice.'),
            semaforo=SEMAFORO_TXT,
            faltantes=lista('El índice se calcula solo en los años en que están los siete componentes: hoy 2019, 2022 y 2024.',
                            'El factor limitante es el espacio verde (solo tres cortes); en 2025 faltan además ILE y servicios públicos.',
                            'Sensibilidad: con la población VP, el espacio verde de 2024 sería 5,54 y el ICPD-7 de 2024 sería 66,17.'),
            limitaciones=lista(*comunes,
                               'El espacio verde son cortes de estudios distintos, no una tendencia.',
                               'El servicio público de 2019 se mantiene con alerta de comparabilidad.',
                               'Los años del índice no forman una serie continua.'))
        req('PDM.ICPD7', 'Promedio simple de 7 componentes')
    elif cod == 'PC.ISPC':
        a = dict(
            nombre='Índice Simplificado de Participación Ciudadana (ISPC)',
            objetivo='Sintetizar votación del presupuesto participativo, cobertura territorial, oferta formativa y publicación de resultados.',
            definicion='Promedio ponderado de los puntajes de los cuatro indicadores de participación. Mantiene los pesos de la matriz original; el primer indicador pasa a medir la votación del presupuesto participativo.',
            formula='ISPC = 0,30 × I1 + 0,25 × I2 + 0,20 × I3 + 0,25 × I4',
            puntaje='Cada indicador es un puntaje de 0 a 100; el índice es el promedio ponderado, sin redondear los componentes.',
            variables=lista('I1 (30 %), I2 (25 %), I3 (20 %) e I4 (25 %), cada uno con su ficha.'),
            referente=lista('Pesos: 30 %, 25 %, 20 % y 25 %, como en la matriz.',
                            'I5 (sesiones) e I6 (participación virtual) se eliminaron y no entran al cálculo.'),
            semaforo=SEMAFORO_TXT,
            faltantes=lista('Solo 2025 tiene los cuatro componentes.',
                            f'I4 de 2025 entra con su mínimo (2/3); con la rendición del presupuesto participativo publicada, el ISPC de 2025 sería {vf("PC.ISPC.MAX", 2025)}.'),
            limitaciones=lista(*comunes, 'I2 2025 es un agregado reportado por la Alcaldía; I3 usa un denominador reconstruido; I1 mezcla votos y votantes.',
                               'La convención del año es el año calendario de la actividad participativa: en I1, la votación para la vigencia siguiente; en I4, el año de la vigencia.'))
    elif cod == 'PP.IREPP':
        a = dict(
            nombre='Índice de Resultados de Ejecución del Presupuesto Participativo (IREPP)',
            objetivo='Resumir ejecución presupuestal, ejecución mayor o igual a 90 % y fidelidad a lo votado.',
            definicion='Promedio ponderado de EPP, EJ90 y FID, con el 80 % del peso en los dos últimos. Mantiene los pesos de la matriz original; EJ90 y FID reemplazan a CEP y CMA.',
            formula='IREPP = 0,20 × EPP + 0,40 × EJ90 + 0,40 × FID',
            puntaje='Cada componente es un puntaje de 0 a 100; el índice usa valores sin redondear y se redondea solo el resultado.',
            variables=lista('EPP (20 %), EJ90 (40 %) y FID (40 %), cada uno con su ficha.'),
            referente=lista('Pesos: 20 %, 40 % y 40 %, como en la matriz.'),
            semaforo=SEMAFORO_TXT,
            faltantes=lista('El mínimo para calcular un año es verificar el 50 % de sus proyectos. ' + _sin_dato_pp(),
                            '2026 está en curso y no se calcula.'),
            limitaciones=lista(*comunes, 'En 2020 el valor votado es un valor priorizado.', _contracredito(),
                               'El cruce entre proyectos votados y ejecución se hace por monto, con el nombre como confirmación.'))
    else:
        a = dict(
            nombre='Índice de Condiciones y Respuesta para el Control Social (ICRCS)',
            objetivo='Resumir con igual peso las cuatro dimensiones documentales del control social.',
            definicion='Promedio simple de los puntajes de DSI, PQRSDF, VRC y RPO. Mantiene la fórmula de la matriz original; PQRSDF reemplaza a PCS.',
            formula='ICRCS = (DSI + PQRSDF + VRC + RPO) / 4',
            puntaje='Cada indicador es un puntaje de 0 a 100; DSI, VRC y RPO entran con su mínimo. El máximo del índice usa los máximos de los rangos.',
            variables=lista('DSI, PQRSDF, VRC y RPO (25 % cada uno), cada uno con su ficha.'),
            referente=lista('Peso: 25 % por indicador, como en la matriz.'),
            semaforo=SEMAFORO_TXT,
            faltantes=lista('Solo 2024 y 2025 tienen índice: el DSI solo aplica desde 2024.',
                            f'Con los máximos de los componentes, el ICRCS sería {vf("CS.ICRCS.MAX", 2024)} en 2024 y {vf("CS.ICRCS.MAX", 2025)} en 2025.'),
            limitaciones=lista(*comunes, 'PQRSDF no es homogénea entre años (plazos ampliados, reclasificación en 2024, cambio de sistema en 2025) y subestima la oportunidad real.',
                               'DSI mide oportunidad y entra con su mínimo (0 en 2024 y 2025); VRC y RPO 2025 llevan alerta.'))
    a.update(base)
    comp = [c for c in principales(dim)]
    if cod == 'PDM.ICPD7':
        a['pesos'] = 'Siete componentes con el mismo peso: 1/7 (0,1429) cada uno.'
    else:
        a['pesos'] = '; '.join(f"{c.split('.')[1]}: {peso_txt(c)}" for c in comp) + '.'
    a['regla'] = ('El índice se calcula solo en los años en que están todos sus componentes; no se redistribuyen pesos entre los demás. '
                  'Los componentes se usan sin redondear y se redondea solo el resultado.')
    if cod == 'CS.ICRCS':
        a['regla'] += ' DSI, VRC y RPO entran con su valor mínimo.'
    if cod == 'PC.ISPC':
        a['regla'] += ' I4 entra con su valor mínimo.'
    todos = list(range(PRIMER_ANIO[cod], 2026))
    con = anios_con_indice(cod)
    a['anios'] = (f"Calculados: {', '.join(str(x) for x in con) or 'ninguno'}. Sin índice: "
                  f"{', '.join(str(x) for x in todos if x not in con) or 'ninguno'} (periodo {todos[0]}–{todos[-1]}).")
    a['serie'] = f"Años con índice: {', '.join(str(x) for x in a_idx)}. Valores: {serie_indice_txt(cod)}."
    a['brecha'] = brecha_texto(dim, cod)
    return a


# =====================================================================================================  REFERENTES CORTOS (hoja Indicadores)
def referentes_cortos():
    veed = sum(num(val('PDM.VEED', a)) for a in range(2019, 2024)) / 5
    req('PDM.VEED', 'LB = 5,4')
    assert abs(veed - 5.4) < 1e-9
    req('PDM.CM', '7.3301'); req('PDM.HOM', '2,7 por 100.000'); req('PDM.CA', '4,54', '9 m2/hab'); req('PDM.ILE', '19 %', '100,44 %')
    req('PC.I1', 'LB 3,68', 'referente 7 %'); req('PC.I2', '12 zonas')
    return {
        'PDM.SALUD': '100 % (cobertura universal)', 'PDM.ILE': 'Metas del PDM: 19 % (primera infancia) y 100,44 % (cobertura neta)',
        'PDM.SP': '100 % (cobertura universal)', 'PDM.CM': f'7,33 por 100.000 (línea base 2019 de la Policía: {vf("PDM.CM", 2019)})',
        'PDM.CA': f'Avance desde 4,54 (2019) hacia 9 m²/hab', 'PDM.VEED': 'Línea base: 5,4 veedurías al año (promedio 2019–2023)',
        'PDM.HOM': 'Meta del PDM: 2,7 por 100.000', 'PDM.ICPD7': 'Promedio simple de 7 componentes',
        'PC.I1': 'Avance desde 3,68 % (2017) hacia 7 %', 'PC.I2': '12 de 12 zonas', 'PC.I3': 'Todas las acciones programadas',
        'PC.I4': '3 de 3 procesos con resultados publicados', 'PC.ISPC': '0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4',
        'PP.EPP': '100 % de la apropiación', 'PP.EJ90': '100 % de los proyectos con ejecución ≥ 90 %', 'PP.FID': '100 % de proyectos fieles a lo votado',
        'PP.IREPP': '0,20·EPP + 0,40·EJ90 + 0,40·FID', 'CS.PQRSDF': '100 % atendidas en términos', 'CS.DSI': '10 de 10 metas con valor publicado a tiempo',
        'CS.VRC': 'Todos los eventos con veeduría documentada', 'CS.RPO': 'Todas las observaciones con respuesta de fondo',
        'CS.ICRCS': '(DSI + PQRSDF + VRC + RPO) / 4',
    }


def cambio_por_codigo():
    """{código nuevo: texto de qué cambió frente a la matriz} y lista de indicadores eliminados."""
    mz = matriz()
    out, eliminados = {}, []
    for dim, orig, etq, nuevos, porque in ORIGINALES:
        if orig:
            ref = f"Matriz original: {orig}, {datos_matriz(orig)['nombre']}."
        else:
            ref = 'No estaba en la matriz original.'
        txt = para_profe(f'{etq.capitalize()}. {ref} {porque}')
        if not nuevos:
            eliminados.append((dim, orig, datos_matriz(orig)['nombre'], txt))
        for n in nuevos:
            out[n] = txt
    return out, eliminados


# =====================================================================================================  ENSAMBLE
_INDICADORES = {
    'PDM.SALUD': _salud, 'PDM.ILE': _ile, 'PDM.SP': _sp, 'PDM.CM': _cm, 'PDM.CA': _ca, 'PDM.VEED': _veed, 'PDM.HOM': _hom,
    'PC.I1': _i1, 'PC.I2': _i2, 'PC.I3': _i3, 'PC.I4': _i4,
    'PP.EPP': _epp, 'PP.EJ90': _ej90, 'PP.FID': _fid,
    'CS.PQRSDF': _pqrsdf, 'CS.DSI': _dsi, 'CS.VRC': _vrc, 'CS.RPO': _rpo,
}


def _fuentes_txt(cod):
    out = []
    for f, enl, mas in fuentes_lista(cod):
        out.append('• ' + f + (('\n   Enlace: ' + '\n   '.join(enl)) if enl else '') + (f'\n   (y {mas} enlaces más; ver la hoja «Datos» del Excel)' if mas else ''))
    return '\n'.join(out) or 'Ver la hoja «Datos» del Excel.'


def construir_fichas():
    """Lista de fichas (una por indicador principal y una por índice), en el orden de las dimensiones."""
    fichas = []
    for dim in ('PDM', 'PC', 'PP', 'CS'):
        for cod in principales(dim):
            d = _INDICADORES[cod]()
            unidad, sentido = unidad_sentido(cod)
            d.update(pesos='—', regla='—', anios='—')
            d.update(codigo=cod, tipo='Indicador principal', dimension=NOMBRE_DIM[dim], unidad=unidad, sentido=sentido, peso=peso_txt(cod),
                      semaforo=SEMAFORO_TXT, serie=f'{serie_texto(dim, cod)} (años con dato; el año sigue la convención de la dimensión)',
                      brecha=brecha_texto(dim, cod), fuentes=_fuentes_txt(cod))
            fichas.append(d)
        cod = INDICE_DIM[dim]
        d = _indice(cod)
        d.update(codigo=cod, tipo='Índice', dimension=NOMBRE_DIM[dim])
        fichas.append(d)
    for f in fichas:
        f['actualizacion'] = f'{FECHA} (datos consultados hasta {meses_fecha(f["codigo"]) or "el 2026-11-19"})'
        f['estado'] = ESTADO
        f['ref_corto'] = referentes_cortos()[f['codigo']]
        for k, _ in CAMPOS:
            assert f.get(k), (f['codigo'], k)
            f[k] = para_profe(f[k])
    return fichas


# =====================================================================================================  INDICADORES DE LA MATRIZ ORIGINAL
# etiqueta: se mantiene / se ajusta / se reemplaza / se elimina / nuevo
ORIGINALES = [
    # dimensión, código original, etiqueta, códigos nuevos, por qué, qué cambió dentro del indicador
    ('PDM', 'PDM.CS', 'se reemplaza', ['PDM.SALUD'],
     'La meta del PDM (99,5 %) es solo del régimen subsidiado y ese dato no era fácil de conseguir con una serie pública. En la reunión de revisión se acordó pasar a la cobertura de salud del municipio '
     'total, con una meta del mismo universo: la cobertura universal.'),
    ('PDM', 'PDM.CM', 'se mantiene', ['PDM.CM'],
     'Mezclar fuentes cambia la medida, y la meta de 8 por 100.000 es de Movilidad, otro universo. Se mantiene con una sola fuente (Policía Nacional, que publica datos abiertos) '
     'y un referente del mismo universo del dato, construido con la reducción que fija el PDM.'),
    ('PDM', 'PDM.CA', 'se ajusta', ['PDM.CA'],
     'Con la meta del PDM (3,88 m² por habitante), el indicador daba 100 sin exigencia, y esa meta no tiene fuente ni método declarados. Se ajusta a un avance desde la línea base 2019 (AMVA) hacia 9 m² por habitante, '
     'solo para espacio verde urbano, con los cortes que existen.'),
    ('PDM', 'PDM.CP', 'se reemplaza', ['PDM.VEED'],
     'No hay cifra ni método oficial publicado del porcentaje de implementación de procesos participativos. Se reemplaza por el número anual de veedurías ciudadanas inscritas en la Personería '
     '(datos abiertos), con línea base.'),
    ('PDM', 'PDM.CD', 'se reemplaza', ['PDM.SP'],
     'No hay una tasa oficial de cobertura deportiva con serie pública. Se reemplaza por servicios públicos domiciliarios (acueducto, alcantarillado, aseo y energía), '
     'con cobertura total y serie anual en TerriData.'),
    ('PDM', 'PDM.ILE', 'se ajusta', ['PDM.ILE'],
     'La clasificación de planteles en Saber 11 no es pública. Se reemplaza ese componente por la cobertura neta total en educación y se conserva la atención a la primera infancia (0 a 5 años). Se conserva el nombre.'),
    ('PDM', 'PDM.ICPD-6', 'se ajusta', ['PDM.ICPD7'],
     'Al reemplazar y agregar componentes quedan siete (se agrega el homicidio como componente de seguridad), así que los dos subíndices de tres componentes no se sostienen. '
     'Se pasa a un promedio simple de siete, para conservar la regla de igual peso por componente.'),
    ('PDM', None, 'nuevo', ['PDM.HOM'],
     'El sexto componente del PDM quedaba abierto en la reunión del 12 de noviembre. Se agrega el homicidio (seguridad), con la meta del PDM y la misma fuente y universo '
     '(Policía Nacional). Este indicador no existía en la matriz original.'),
    ('PC', 'PC.I1', 'se reemplaza', ['PC.I1'],
     'No hay inventario de procesos participativos, y el respaldo repetía los datos de la fidelidad del presupuesto participativo. Para diferenciarlos, se reemplaza por la votación del presupuesto participativo (IPCM), '
     'un indicador oficial de la Alcaldía con línea base.'),
    ('PC', 'PC.I2', 'se mantiene', ['PC.I2'],
     f'Se conserva la cobertura territorial. La matriz pedía validar el universo de zonas: se usan 12 zonas (la zona 1 es industrial). El dato principal existe para {serie_texto("PC", "PC.I2")}, con alerta: es cobertura documental por zona.'),
    ('PC', 'PC.I3', 'se mantiene', ['PC.I3'],
     'Se conserva la fórmula. Solo 2025 tiene tasa, con un denominador reconstruido porque no hay programación publicada; se agrega como respaldo el conteo de acciones formativas para la ciudadanía.'),
    ('PC', 'PC.I4', 'se mantiene', ['PC.I4'],
     'No hay inventario oficial de procesos participativos cerrados. Se conserva el indicador con una lista fija de tres procesos que el observatorio verifica cada año en el portal (no es autorreporte).'),
    ('PC', 'PC.I5', 'se elimina', [],
     'Se acordó en la primera reunión eliminar el cumplimiento de sesiones de las instancias de participación; no entra al índice ni se calcula.'),
    ('PC', 'PC.I6', 'se elimina', [],
     'Se acordó en la primera reunión eliminar la disponibilidad de participación virtual; no entra al índice ni se calcula.'),
    ('PC', 'PC.ISPC', 'se mantiene', ['PC.ISPC'],
     'Se mantienen la fórmula y los pesos (30, 25, 20 y 25 %). Cambia lo que mide I1 y se eliminan I5 e I6, que ya eran complementarios y no entraban al cálculo.'),
    ('PP', 'PP.EPP', 'se mantiene', ['PP.EPP'],
     'Hay dato de Hacienda por proyecto. Se mantiene la definición; se excluye el rubro de déficit fiscal (pagos de vigencias anteriores) de los proyectos de presupuesto participativo.'),
    ('PP', 'PP.CEP', 'se reemplaza', ['PP.EJ90'],
     'No hay actas de recibo y «terminado» solo existe como autorreporte. Se reemplaza por los proyectos con ejecución de al menos 90 % (verificable en Hacienda); «terminados» queda solo como respaldo con advertencia.'),
    ('PP', 'PP.CMA', 'se reemplaza', ['PP.FID'],
     'El plan de acción mide actividades, no personas atendidas, así que las metas de atención no se pueden calcular. Se reemplaza por la fidelidad a lo votado.'),
    ('PP', 'PP.IREPP', 'se mantiene', ['PP.IREPP'],
     'Se mantienen la fórmula y los pesos (20, 40 y 40 %). EJ90 y FID reemplazan a CEP y CMA. La serie se amplía a 2020–2025 y se reporta la brecha.'),
    ('CS', 'CS.PCS', 'se reemplaza', ['CS.PQRSDF'],
     'Era un autorreporte: 1 de 1 estrategia ejecutada daba 100 %. Se reemplaza por la oportunidad de respuesta a las PQRSDF, que se calcula con informes trimestrales públicos.'),
    ('CS', 'CS.DSI', 'se mantiene', ['CS.DSI'],
     'Se mantiene el indicador, con la lista de metas actualizada (10 metas de resultado del PDM 2024–2027) y como rango con brecha de información, porque la oportunidad de publicación no siempre se puede verificar.'),
    ('CS', 'CS.VRC', 'se mantiene', ['CS.VRC'],
     'Se mantiene el indicador; como en varios eventos no se sabe si hubo veeduría, se reporta como rango (mínimo y máximo) con la brecha de información, y al índice entra el mínimo.'),
    ('CS', 'CS.RPO', 'se mantiene', ['CS.RPO'],
     'Se mantiene el indicador con el universo de la audiencia del Alcalde, sin felicitaciones. Donde no se puede probar la fecha de publicación, se reporta un rango; 2025 lleva alerta.'),
    ('CS', 'CS.ICRCS', 'se mantiene', ['CS.ICRCS'],
     'Se mantienen la fórmula y el peso de 25 % por indicador; PQRSDF ocupa el lugar de PCS. VRC, RPO y DSI entran con su mínimo.'),
]
ETIQUETAS = ('se mantiene', 'se ajusta', 'se reemplaza', 'se elimina', 'nuevo')
TITULO_ETQ_LISTA = ('[SE MANTIENE]', '[SE AJUSTA]', '[SE REEMPLAZA]', '[SE ELIMINA]', '[NUEVO]')


def datos_matriz(cod):
    m = matriz()[cod]
    return dict(nombre=m['Indicador / índice'].strip().split('\n')[0].strip(),
                objetivo=una_linea(m['Objetivo']), formula=m['Fórmula y metodología'].strip(), semaforo=una_linea(m['Semáforo del anexo']),
                periodicidad=una_linea(m['Periodicidad']))


def tabla_cambios():
    """Una fila por indicador de la matriz original (y los nuevos), para el Excel y el Word."""
    fichas = {f['codigo']: f for f in construir_fichas()}
    mz = matriz()
    out = []
    for dim, orig, etq, nuevos, porque in ORIGINALES:
        assert etq in ETIQUETAS
        if orig:
            assert orig in mz, orig
            nom = datos_matriz(orig)['nombre']
        else:
            nom = '(no existía en la matriz original)'
        if nuevos:
            quedo = '; '.join(f"{n}: {fichas[n]['nombre']}. Fórmula: {una_linea(fichas[n]['formula'])}" for n in nuevos)
        else:
            quedo = 'No se calcula ni entra al índice.'
        out.append(dict(dimension=NOMBRE_DIM[dim], codigo_original=orig or '—', indicador_original=nom, que_paso=etq,
                        por_que=para_profe(porque), como_quedo=para_profe(quedo), codigo_nuevo=', '.join(nuevos) or '—'))
    # todos los indicadores de la matriz aparecen
    assert {o[1] for o in ORIGINALES if o[1]} == set(mz), set(mz) ^ {o[1] for o in ORIGINALES if o[1]}
    return out


# =====================================================================================================  ANEXO: DEFINICIONES COMUNES
def definiciones_comunes():
    pob18, pob24 = val('PDM.POB', 2018, 'valor'), val('PDM.POB', 2024, 'valor')
    assert pob18 == '228836' and pob24 == '260430', (pob18, pob24)
    return [
        ('Brecha de información',
         'Parte del periodo de un indicador (años, trimestres, eventos o metas) para la que no hay dato publicado o verificable. '
         'Se reporta siempre: brecha = unidades sin dato / unidades del periodo. En un índice, es la proporción de años sin índice por falta de algún componente. '
         f'Ejemplo (IREPP): {brecha_texto("PP", "PP.IREPP")}. Una brecha alta no significa un mal desempeño: significa que no se puede medir.'),
        ('Rango mínimo–máximo',
         'Cuando un dato no se puede verificar del todo (por ejemplo, no se sabe si hubo una veeduría en un evento, o si una respuesta se publicó a tiempo), '
         'el indicador se reporta con un mínimo y un máximo. El mínimo cuenta como cumplido solo lo que está documentado; el máximo suma lo no verificable. '
         'Al índice entra el mínimo, y la diferencia entre ambos es la brecha de información. Se usa en PQRSDF, DSI, VRC, RPO e I4 (2025).'),
        ('Mismo universo',
         'Un dato solo se compara con una meta o línea base del mismo universo: la misma población, el mismo concepto y el mismo periodo. '
         'Por eso no se usan la meta de 99,5 % del régimen subsidiado para la cobertura total en salud, ni la meta de 8 de Movilidad para la tasa de la Policía; '
         'se construye un referente del mismo universo del dato.'),
        ('Población DANE VP',
         'Todas las tasas y coberturas por habitante usan una sola serie de población: las proyecciones del DANE vigentes al 8 de agosto de 2025 (VP), '
         'tal como las publica TerriData (DNP) en la dimensión «Descripción general», indicador «Población total» '
         f'(2018 = {fmt(pob18)}; 2024 = {fmt(pob24)}). No se usa la proyección anterior de la dimensión «Censo 2005 y proyecciones DANE» (2018 = 238.221).'),
        ('Convención de años',
         'Presupuesto participativo: el año es el de la vigencia (el año en que se ejecutan los proyectos). '
         'Votación del presupuesto participativo (I1): el año es el de la votación; los proyectos se ejecutan en el año siguiente. '
         'Publicación de resultados (I4): año de la vigencia. PQRSDF: año de los trimestres. Los demás indicadores: el año del dato. '
         'Cada indicador trae todos los años que tiene; el índice usa solo los años en que coinciden todos sus componentes. El año 2026 está en curso y no se califica.'),
        ('Puntaje de 0 a 100',
         'Cada indicador se convierte a una escala de 0 a 100 comparándolo con su referente (línea base y meta). '
         'Los indicadores que ya son porcentajes de cumplimiento usan el valor como puntaje.'),
        ('Semáforo (propuesta)', 'Verde: 90 a 100. Amarillo: 75 a menos de 90. Naranja: 50 a menos de 75. Rojo: menos de 50. '
                                 'Se aplica al puntaje y es una propuesta, igual en las cuatro dimensiones, por validar con el profesor. '
                                 'En los indicadores con rango se asigna con el mínimo.'),
        ('Sin redistribuir pesos', 'Un índice se calcula solo si están todos sus componentes. Si falta uno, el índice de ese año queda sin dato; '
                                   'no se reparte su peso entre los demás.'),
        ('Sin dato y no aplica', '«Sin dato» (N/D): el dato no existe o no se pudo verificar. «No aplica»: el concepto no corresponde a ese año '
                                 '(por ejemplo, no hubo votación del presupuesto participativo en 2018, o la lista de metas del DSI solo existe desde 2024).'),
        ('Estado', 'Todos los resultados son provisionales: pueden cambiar cuando se cierren brechas de información o se actualicen las fuentes.'),
    ]


# =====================================================================================================  FUENTE DE LA SERIE COMPLETA (hoja Indicadores)
_FUENTE_SERIE = {
    'PDM.SALUD': 'TerriData (DNP): afiliados al sistema de salud (Ministerio de Salud) y población DANE VP',
    'PDM.ILE': 'Ministerio de Educación (datos.gov.co) y Alcaldía de Envigado (balance del PDM e informe de gestión) para la atención inicial; cobertura neta de TerriData; población DANE VP',
    'PDM.SP': 'TerriData (DNP): coberturas de acueducto, alcantarillado y aseo (SUI) y energía (UPME)',
    'PDM.CM': 'Policía Nacional (datos.gov.co) desde 2010 y TerriData antes; población DANE VP',
    'PDM.CA': 'Tres cortes de estudios: AMVA con la Universidad Nacional, AMVA con el Instituto Humboldt y diagnóstico del POT de Envigado',
    'PDM.VEED': 'Personería de Envigado (datos.gov.co): veedurías inscritas',
    'PDM.HOM': 'Policía Nacional (datos.gov.co); población DANE VP',
    'PC.I1': 'Informe de priorización de Planeación (2017), documentos y noticias de la Alcaldía (2020–2023) e informe de gestión 2025; población DANE VP',
    'PC.I2': 'Fichas y agendas zonales de la Alcaldía (2020–2022), plan indicativo e informe de gestión (2024–2025)',
    'PC.I3': 'Informe de gestión 2025 de la Alcaldía de Envigado',
    'PC.I4': 'Portal de la Alcaldía: documentos de rendición de cuentas y de resultados de la votación de PP',
    'PP.EPP': 'Ejecuciones presupuestales de la Alcaldía (egresos de Hacienda y plan de acción de 2025) + listas de proyectos de PP (lista oficial en 2021; listado EFPC en los demás años)',
    'PP.EJ90': 'Ejecuciones presupuestales de la Alcaldía (egresos de Hacienda y plan de acción de 2025) + listas de proyectos de PP (lista oficial en 2021; listado EFPC en los demás años)',
    'PP.FID': 'Ejecuciones presupuestales de la Alcaldía (egresos de Hacienda y plan de acción de 2025) + listas de proyectos de PP (lista oficial en 2021; listado EFPC en los demás años)',
    'CS.PQRSDF': 'Informes trimestrales de PQRSDF de la Secretaría General de la Alcaldía (24 informes)',
    'CS.DSI': 'Plan de Desarrollo 2024–2027 (Apéndice 9.1.2) y las fuentes que nombra para cada meta',
    'CS.VRC': 'Actas, memorias y cuadros de preguntas de las rendiciones de cuentas de la Alcaldía (detalle en «Datos», filas CS.VRC.EVENTO)',
    'CS.RPO': 'Cuadros de preguntas y respuestas de la audiencia pública de rendición de cuentas del Alcalde',
}


def fuente_serie(cod):
    """Descripción de la fuente de la serie completa del indicador (no solo del primer año)."""
    if cod in INDICE_DIM.values():
        return 'Cálculo del observatorio con los componentes'
    return _FUENTE_SERIE[cod] + f'. Años con dato: {serie_texto(dim_de(cod), cod)}'
