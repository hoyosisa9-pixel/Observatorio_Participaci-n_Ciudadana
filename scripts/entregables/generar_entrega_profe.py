"""Genera los dos Word de la revisión de indicadores (nov-2026).

1. outputs/entrega_profe/indicadores_cambios_y_datos_OCE.docx: para el profesor. Corto, con cuadros sencillos.
2. outputs/entrega_profe/detalle_revision_indicadores_OCE.docx: para Isabella. Qué pasó con cada indicador, en detalle.

Uso (desde la raíz, sin -I):  python3 scripts/entregables/generar_entrega_profe.py
Contenido: outputs/entrega_profe/*.md y outputs/revision_indicadores/qa_construibilidad.md (15-nov-2026).
Las dudas se leen de outputs/entrega_profe/dudas_para_el_profe.md. Todo es PROVISIONAL.
"""
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / 'outputs' / 'entrega_profe'
FUENTE = 'Cambria'
GRIS = 'E7E6E6'
AZUL = RGBColor(0x1F, 0x3A, 0x5F)
FECHA = '15 de noviembre de 2026'


# ------------------------------------------------------------------ utilidades de Word
def fijar_fuente(estilo, nombre, tam=None, negrita=None, color=None):
    rpr = estilo.element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.insert(0, rf)
    for a in list(rf.attrib):
        if a.endswith('Theme'):
            del rf.attrib[a]
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), nombre)
    if tam:
        estilo.font.size = Pt(tam)
    if negrita is not None:
        estilo.font.bold = negrita
    if color is not None:
        estilo.font.color.rgb = color


def nuevo_documento(horizontal):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
    if horizontal:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ('left_margin', 'right_margin'):
        setattr(sec, m, Cm(2.0))
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = doc.styles
    fijar_fuente(st['Normal'], FUENTE, 10.5)
    st['Normal'].paragraph_format.space_after = Pt(4)
    st['Normal'].paragraph_format.line_spacing = 1.1
    for nombre, tam in (('Title', 18), ('Heading 1', 14), ('Heading 2', 12), ('Heading 3', 11)):
        fijar_fuente(st[nombre], FUENTE, tam, True, AZUL)
        pf = st[nombre].paragraph_format
        pf.space_before, pf.space_after, pf.keep_with_next = Pt(10), Pt(4), True
    for nombre in ('List Bullet', 'List Number'):
        fijar_fuente(st[nombre], FUENTE, 10.5)
    pie(sec)
    return doc


def pie(sec):
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Observatorio Ciudadano de Envigado · Provisional · página ')
    r.font.size = Pt(8)
    for tipo, texto in (('begin', None), (None, 'PAGE'), ('end', None)):
        run = p.add_run()
        run.font.size = Pt(8)
        if tipo:
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), tipo)
            run._r.append(fc)
        else:
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = texto
            run._r.append(it)


def escribir(p, texto, tam=None):
    """Escribe texto con **negritas** al estilo Markdown."""
    for i, trozo in enumerate(re.split(r'\*\*', texto)):
        if trozo:
            r = p.add_run(trozo)
            r.bold = i % 2 == 1
            if tam:
                r.font.size = Pt(tam)
    return p


def parrafo(doc, texto, tam=None):
    return escribir(doc.add_paragraph(), texto, tam)


def vinetas(doc, items, numerada=False):
    for it in items:
        escribir(doc.add_paragraph(style='List Number' if numerada else 'List Bullet'), it)


def sombrear(celda, color):
    tcpr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcpr.append(shd)


def tabla(doc, anchos_cm, encabezado, filas, tam=8.5):
    t = doc.add_table(rows=1, cols=len(encabezado))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for j, txt in enumerate(encabezado):
        c = t.rows[0].cells[j]
        c.text = ''
        escribir(c.paragraphs[0], f'**{txt}**', tam)
        sombrear(c, GRIS)
    # repetir encabezado en cada página
    trpr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement('w:tblHeader')
    th.set(qn('w:val'), 'true')
    trpr.append(th)
    for fila in filas:
        celdas = t.add_row().cells
        for j, txt in enumerate(fila):
            celdas[j].text = ''
            escribir(celdas[j].paragraphs[0], txt, tam)
    for fila in t.rows:
        for j, a in enumerate(anchos_cm):
            fila.cells[j].width = Cm(a)
            for p in fila.cells[j].paragraphs:
                p.paragraph_format.space_after = Pt(1)
    doc.add_paragraph()
    return t


# ------------------------------------------------------------------ contenido común
ENC = ['Indicador original', 'Situación', 'Propuesta', 'Origen', 'Datos']

PDM = [
    ['**Aseguramiento al régimen subsidiado**', 'El dato del subsidiado no era fácil de conseguir; la meta del PDM (99,5 %) es solo de ese régimen',
     '**Se reemplaza** por la cobertura de salud total (afiliados / población DANE)', 'Profe (reunión 1) e Isabella; confirmado el 12-nov',
     '2018–2025 (TerriData). Sin meta del mismo universo'],
    ['**Mortalidad vial**', 'Forensis 2025 no se puede consultar y mezclar fuentes cambia la medida',
     '**Se mantiene**, con una sola fuente: Policía Nacional', 'Profe; fuente única: Claude', '2019–2025. Meta 8 por 100.000, de Movilidad (otro universo)'],
    ['**Espacio público verde**', 'Con la meta del PDM (3,88 m²) daba 100 sin exigencia',
     '**Se ajusta**: avance desde la **línea base 2019 = 4,54 m²/hab** (AMVA) hacia el referente de 9 m²/hab', 'Profe (12-nov: avance hacia 9 m²); línea base 2019: Isabella',
     'Cortes de espacio verde urbano: 2019: 4,54 (AMVA) · 2022: 4,66 (AMVA) · 2024: 5,75 (POT). Avance: 2,7 % en 2022 y 27,1 % en 2024. No hay dato anual (ver cuadro de fuentes)'],
    ['**Implementación de procesos participativos**', 'No hay cifra ni método oficial publicado',
     '**Se reemplaza** por el número anual de veedurías inscritas en la Personería', 'Profe (notas de reunión) e Isabella',
     '2012–2026 (datos.gov.co). El listado parece incompleto'],
    ['**Cobertura deportiva**', 'No hay tasa oficial del Índer',
     '**Se reemplaza** por servicios públicos: acueducto, alcantarillado, aseo y energía (promedio simple)', 'Profe (12-nov) e Isabella',
     '2019–2024 (TerriData). Meta 99,9 %, de otro universo'],
    ['**Índice de logro educativo**', 'La clasificación de planteles en Saber 11 no es pública',
     '**Se ajusta**: Saber 11 → cobertura neta; se conserva primera infancia (0–5 años)', 'Profe (reunión 1 y 12-nov) e Isabella',
     '2018–2024 (TerriData, MinEducación)'],
    ['— Sexta medida', 'Quedó abierta el 12-nov', 'Ver duda 1', 'Profe', 'Homicidio, hurto y lesiones personales 2019–2025 (Policía)'],
    ['— Logros de políticas públicas', 'Dimensión para el observatorio futuro', 'Solo variables por solicitar', 'Profe (reunión 1)', 'Sin datos'],
]

PDM_COMP = [
    ['Cobertura de salud', 'Afiliados y población', '2018–2025', 'No: la del PDM es solo del subsidiado'],
    ['Logro educativo', 'Cobertura neta, niños atendidos y población de 0–5 años', '2018–2024', 'Sí (19 % y 100,44 %)'],
    ['Servicios públicos', 'Coberturas de acueducto, alcantarillado, aseo y energía', '2019–2024', 'Sí, 99,9 %, pero de otro universo (el aseo no tiene)'],
    ['Mortalidad vial', 'Muertes en tránsito y población', '2019–2025', 'Sí, 8 por 100.000, pero de otro universo'],
    ['Espacio verde urbano', 'm² verdes urbanos por habitante', '**2019 (línea base), 2022 y 2024** (cortes de distintos estudios)', 'Referente 9 m² (no se usa la meta del PDM de 3,88)'],
    ['Veedurías (en vez de procesos participativos)', 'Veedurías inscritas por año', '2012–2026', 'No hay meta en el PDM'],
]

EV_ENC = ['Año', 'm² por habitante', 'Avance hacia 9 m²', 'Quién lo midió', 'Fuente (documento, tabla y página)']
EV_CORTES = [
    ['**2019**', '**4,54**', '0 % (línea base)', 'Área Metropolitana (AMVA) con la Universidad Nacional, estudio de conectividad ecológica (2020)',
     'Reporte Metropolitano Voluntario ODS 11.7, AMVA y ONU-Hábitat (2022), tabla 4, p. 24'],
    ['2022', '4,66', '2,7 %', 'AMVA con el Instituto Humboldt, actualización del Plan Maestro de Espacios Públicos Verdes Urbanos',
     'Nota del Instituto Humboldt (2023), gráfico por municipio; repetido en el diagnóstico del POT, tabla 39, pp. 86-87'],
    ['2024', '5,75', '27,1 %', 'Alcaldía de Envigado, diagnóstico de la revisión del POT (criterio del AMVA, inventario propio)',
     'Diagnóstico del POT, dimensión funcional (espacio público), tabla 39, pp. 86-87: 1.382.762 m² / 240.655 habitantes urbanos (DANE 2024)'],
]
EV_NOTA = ('Avance = (valor del año − 4,54) / (9 − 4,54). Son **cortes de estudios distintos, no una serie anual**: el dato no se publica cada año '
           'y el de 2024 usa otro inventario, así que parte del avance puede deberse al cambio de método. '
           '**Recomendación para el informe:** solicitar a **Planeación** la serie completa desde 2019 con la metodología del AMVA, y proponer que '
           '**Planeación la mida cada año** en el seguimiento del POT (expediente municipal), con el mismo método, y que la **Secretaría de Medio Ambiente** '
           'la reporte en el seguimiento del PDM. Se descartan el 3,88 del PDM (sin fuente ni método), los 5,00 de 2020 y 2023 (autorreporte) y el '
           'verde total 2024 (69,41, incluye las reservas rurales).')
EV_URLS = [
    '2019 (4,54): Reporte ODS 11.7, AMVA y ONU-Hábitat (2022), tabla 4, p. 24, y tabla 5, p. 25: https://urbanpolicyplatform.org/wp-content/uploads/2023/06/VMR-EPV-VALLE-DE-ABURRA.pdf',
    '2022 (4,66): Instituto Humboldt (2023), https://www.humboldt.org.co/noticias/espacios-publicos-verdes-urbanos-en-el-valle-de-aburra-un-reto-de-conservacion '
    '(gráfico: https://cdn.prod.website-files.com/64df6dd37ac6a0dbb9d03cb3/6591f459139e1daa37e628c9_Instituto-Humboldt-plan.jpg, columna 2022, fila Envigado).',
    '2024 (5,75): diagnóstico del POT, dimensión funcional (espacio público), tabla 39, pp. 86-87: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002925/146211_8-dimension-funcional-_-espacio-publico.pdf',
    'Contexto 2013 (4,16): Minambiente, Informe Nacional de Calidad Ambiental Urbana 2013, tabla 16, p. 41: https://www.minambiente.gov.co/wp-content/uploads/2021/06/Informe_Nacional_de_Calidad_Ambiental_Urbana_poblacion_100.000_y_500.000_habitantes._Ano_2013.pdf',
    'Descartado 3,88 (LB 2023 del PDM 2024-2027): apéndice 9.1.2 del PDM, fila 129: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002340/116989_apendice91planindicativo27052024.xlsx',
    'Búsqueda completa (todos los años, valores descartados y motivos): outputs/revision_indicadores/PDM/espacio_verde_serie.md; verificación de QA: outputs/revision_indicadores/PDM/qa_espacio_verde.md.',
]

I1_ENC = ['Año de la votación', 'Votos', 'IPCM', 'Avance hacia 7 %', 'Fuente (documento y página)']
I1_DATOS = [
    ['**2017**', '8.221 votos', '**3,62 %** (oficial)', '0 % (línea base)', 'Planeación, Informe de Priorización Participativa del Presupuesto (11-oct-2017), p. 7 y 12'],
    ['2018', '—', '—', '—', 'No publicado: por solicitar'],
    ['2019', '0', '—', '—', 'No hubo votación: los recursos alcanzaron para todos los proyectos (informe de auditoría de la priorización 2019)'],
    ['2020–2024', '—', '—', '—', 'No publicado: por solicitar (el documento de 2023 es una imagen escaneada)'],
    ['2025', '10.662 votantes', '4,03 % (calculado con la población DANE: 264.637)', '12,1 %', 'Informe de gestión por dependencias 2025, p. 92 del PDF (votación del 11 al 24 de agosto)'],
]
I1_NOTA = ('IPCM = votos de la priorización del presupuesto participativo / población × 100. Avance = (IPCM − 3,62) / (7 − 3,62). '
           '**El PDM 2024-2027 no tiene meta** de participación en el presupuesto participativo (solo «ejercicios realizados: 1 al año»), así que se usa como referente '
           'el **7 %, tope de la escala oficial** con la que Planeación reparte los recursos (tramos de 1–2,99 %, 3–4,99 % y 5–6,99 %; informe de 2017, p. 6). '
           '**Cuidado:** 2017 cuenta votos y 2025 votantes; la Alcaldía usa su propia población (el 3,62 % no sale exacto con el DANE). '
           '**Recomendación para el informe:** solicitar a **Planeación** los votos y el IPCM de 2018 a 2024, total y por zona, la población de referencia y si cuentan votos o votantes.')
I1_URLS = [
    '2017 (8.221 votos; IPCM 3,62 %; escala de tramos en la p. 6): https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/000288/14376_informepriorizacionparticipativadelpresupuesto11deoctubrede2017.pdf',
    '2019 (sin votación): https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/000106/5254_informe-ejecutivo-auditoria-priorizacion-participativa-del-presupuesto-2019.pdf',
    '2025 (10.662 votantes, p. 92 del PDF): https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003058/152890_2025informegestiondependencias.pdf',
    'Sin meta en el PDM: apéndice 9.1.1 del plan indicativo, fila 539 («Ejercicios de presupuesto participativo realizados», 1 por año): https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002340/116989_apendice91planindicativo27052024.xlsx',
    'Población 2025: DANE, proyección PPED 2018-2042 (264.637), la misma de las demás tasas (outputs/revision_indicadores/PDM/series_pdm.csv).',
]

VRC_ENC = ['Año', 'Eventos', 'Con veeduría documentada', 'Sin veeduría', 'No se sabe', 'VRC mínimo', 'VRC máximo', 'Brecha de información']
VRC_DATOS = [
    ['2024', '3', '2', '0', '1 (audiencia del Alcalde)', '**66,7 %**', '100 %', '**33,3 %**'],
    ['2025', '5', '1', '1', '3 (audiencia del Alcalde y Discapacidad S1 y S2)', '**20 %**', '80 %', '**60 %**'],
]
VRC_NOTA = ('**Cómo se construye con la brecha:** VRC mínimo = eventos con veeduría documentada / eventos (lo que no se sabe cuenta como no documentado; '
            'coincide con el nombre del indicador, «presencia documentada»). VRC máximo = (documentada + no se sabe) / eventos. '
            '**Brecha de información** = eventos sin soporte / eventos: dice cuánto del resultado depende de documentos que la Alcaldía no publica. '
            'Para el índice se usaría el **mínimo**, siempre acompañado de la brecha. Eventos: audiencia del Alcalde y rendiciones sectoriales con acta o memoria '
            '(Salud, Discapacidad, Víctimas). Con todos los eventos publicados: 2024: 50 %–100 % (brecha 50 %); 2025: 28,6 %–85,7 % (brecha 57,1 %). '
            '**Para cerrar la brecha:** solicitar a la Alcaldía las listas de asistencia de cada evento con la casilla «¿pertenece a una veeduría? ¿cuál?»; '
            'confirmar con la Personería que la veeduría esté inscrita, y que el observatorio asista a la audiencia del Alcalde y registre las veedurías presentes. '
            '**Se deja así:** VRC se reporta cada año como rango, con la brecha al lado. La brecha también sirve para seguir la calidad de la información: se espera que **baje año a año** a medida que la Alcaldía publique las listas de asistencia y el observatorio lleve su propio registro, hasta que el mínimo y el máximo coincidan.')
VRC_NECESITA = [
    'Lista fija de eventos del año (denominador): audiencia del Alcalde y rendiciones sectoriales (Salud, Discapacidad, Víctimas, presupuesto participativo, grupos étnicos), '
    'confirmada en enero con el cronograma de rendición de cuentas que publica la Alcaldía. La recoge el observatorio.',
    'Lista de asistencia de cada evento con la casilla «¿pertenece a una veeduría? ¿cuál?». La produce la dependencia que organiza el evento (por confirmar quién coordina la rendición de cuentas en Envigado); la recoge el observatorio después de cada evento.',
    'Listado de veedurías inscritas, para verificar que quien firma como veedor pertenece a una veeduría inscrita. Lo produce la Personería; el observatorio lo pide una vez al año.',
    'Registro propio del observatorio: una fila por evento (fecha, dependencia, documento, enlace y veeduría: sí, no o sin soporte). Cierre al 31 de diciembre; búsqueda de documentos hasta el 31 de marzo (regla 17 de la guía).',
    'Observación directa: que el observatorio asista a la audiencia del Alcalde y registre las veedurías presentes, porque es el evento que hoy queda sin dato.',
]

PC = [
    ['**I1** Incorporación de aportes', 'No hay inventario de procesos; el respaldo usaba los mismos datos de la fidelidad de PP',
     '**Se reemplaza** por la votación del presupuesto participativo (IPCM = votos / población), para diferenciarlo de PP. '
     'Avance desde la **línea base 2017 = 3,62 %** hacia el **referente 7 %** (tope de la escala oficial del IPCM)',
     'Profe (diferenciar); Isabella y Claude; LB 2017 y referente 7 %: Isabella',
     '2017: 3,62 % (oficial) · 2025: 4,03 % (10.662 votantes / población DANE): avance 12,1 %. 2019 sin votación; 2018–2024 por solicitar (ver cuadro de fuentes)'],
    ['**I2** Cobertura territorial', 'El dato principal solo existe para 2024–2025',
     '**Se mantiene**; respaldo: zonas con proyectos de PP', 'Profe; respaldo: Claude',
     '2024: 13/13; 2025: 12/13. Respaldo 2020–2025: 12/13 cada año (la zona 1 es industrial)'],
    ['**I3** Oferta formativa', 'Solo 2025 tiene denominador, y es reconstruido',
     '**Se mantiene**; respaldo: número de acciones realizadas', '—', 'Tasa 2025: 83,3 %. Conteo 2023–2025: 17, 7 y 20 (criterios por unificar)'],
    ['**I4** Publicación de resultados', 'No hay inventario oficial de procesos cerrados',
     '**Se mantiene** con una lista fija de 3 procesos anuales que el observatorio verifica en el portal', 'Isabella y Claude', '2021–2025'],
    ['**I5** e **I6**', '—', '**Eliminados**', 'Profe (reunión 1)', '—'],
]

PP = [
    ['**EPP** Ejecución presupuestal', 'Hay dato de Hacienda por proyecto', '**Se mantiene**', 'Profe', '2024: 98,7 %; 2025: 98,8 %'],
    ['**CEP** Entrega de proyectos', 'No hay actas de recibo; «terminado» solo existe como autorreporte',
     '**Se reemplaza** por la ejecución según los proyectos que sí se hicieron. Propuesta: proyectos con ejecución ≥ 90 % (Hacienda); terminados como respaldo',
     'Profe; variante verificable: Claude', '2024: 96 %; 2025: 93,4 %'],
    ['**CMA** Metas de atención', 'El plan de acción mide actividades, no personas', '**Se reemplaza** por la fidelidad a lo votado', 'Profe',
     '2024: 76,5 %; 2025: 95,1 %'],
    ['Respaldos', '—', 'Fidelidad en dinero, tasa de compromiso y equidad entre zonas', 'Claude (aceptado)', '2024–2025'],
]

CS = [
    ['**PCS** Promoción del control social', 'Autorreporte: 1 de 1 estrategia = 100 % en 2024 y 2025',
     '**Se reemplaza** por la oportunidad de respuesta a las PQRSDF (atendidas a tiempo / recibidas)',
     'Isabella (no al autorreporte); reemplazo: Claude', '2020–2025 (informes trimestrales)'],
    ['**DSI** Información de seguimiento', 'Verificable; 0 % en 2024 y 2025', '**Se mantiene**; hay que actualizar la lista de metas del PDM', 'Profe', '2024–2025'],
    ['**VRC** Veedurías en rendición de cuentas', 'En varios eventos no se sabe si hubo veeduría (incluida la audiencia del Alcalde): con la regla de la matriz queda N/D',
     '**Se mantiene**, reportado como **rango** con la **brecha de información** al lado; la brecha debe bajar año a año (ver cuadro)', 'Profe; rango y brecha: Isabella, con la propuesta de Claude',
     '2024: 66,7 %–100 % (brecha 33,3 %) · 2025: 20 %–80 % (brecha 60 %)'],
    ['**RPO** Respuesta a observaciones', 'El cuadro de 2025 está incompleto', '**Se mantiene**', 'Profe', '2024: 94,1 %; 2025 frágil; ampliable a 2021 y 2023'],
]

INDICES = [
    ['**ICPD-6** (PDM)', 'Bienestar = salud + logro educativo + servicios públicos; territorio = mortalidad vial + espacio verde + veedurías',
     '2019, 2022 y 2024', 'Dudas 1 a 5: referentes de salud y veedurías, metas de otro universo. Espacio verde con un mismo método cada año'],
    ['**ISPC** (PC)', '0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4', 'Solo 2025', 'Dudas 6 a 8: I1, denominador de I2 y pesos'],
    ['**IREPP** (PP)', '0,20·EPP + 0,40·ejecución ≥ 90 % + 0,40·fidelidad', '2024 y 2025', 'Dudas 9 y 10. Resueltas, se puede construir ya'],
    ['**ICRCS** (CS)', '(DSI + PQRSDF + VRC + RPO) / 4', '2024 (y 2025 si se completa RPO), con VRC mínimo y sumando los trimestres de PQRSDF', 'VRC como rango con brecha (duda 14) y RPO 2025 incompleto. Dudas 12 a 14'],
]

SOLICITAR = [
    'Espacio público verde urbano por habitante, año por año desde 2019 y con la metodología del AMVA (Planeación). Que Planeación lo mida cada año en el seguimiento del POT y la Secretaría de Medio Ambiente lo reporte en el PDM.',
    'Votos de presupuesto participativo e IPCM por año y por zona, 2018–2024, con la población de referencia y si se cuentan votos o votantes (Planeación).',
    'Apropiación, compromisos y obligaciones por proyecto de presupuesto participativo, 2020–2023 (Hacienda o Planeación).',
    'Listado histórico completo de veedurías inscritas (Personería).',
    'Programación anual de la oferta formativa y actividades participativas por zona (Escuela de Formación, Participación).',
    'Inventario de procesos participativos con los aportes recibidos y la decisión tomada.',
    'Actas de recibo y personas atendidas por proyecto de presupuesto participativo.',
    'Listas de asistencia de las rendiciones de cuentas 2024–2025, con la casilla de veeduría (Alcaldía), para cerrar la brecha de VRC; y cuadro completo de preguntas y respuestas de 2025.',
    'Clasificación de planteles en Saber 11 (Secretaría de Educación).',
    'Logros de políticas públicas: políticas vigentes, informes de seguimiento y metas cumplidas.',
]


def leer_dudas():
    """Devuelve [(sección, [texto])] desde dudas_para_el_profe.md."""
    secciones, actual = [], None
    for linea in (OUT / 'dudas_para_el_profe.md').read_text(encoding='utf-8').splitlines():
        if linea.startswith('## '):
            actual = (linea[3:].strip(), [])
            secciones.append(actual)
        elif actual and re.match(r'^\d+b?\. ', linea):
            actual[1].append(linea.strip())
        elif actual and linea.strip().startswith('- ') and actual[1]:
            actual[1][-1] += ' ' + linea.strip()[2:]
    return [s for s in secciones if s[1]]


# ------------------------------------------------------------------ documento 1: para el profe
def doc_profe():
    doc = nuevo_documento(horizontal=True)
    doc.add_paragraph(style='Title').add_run('Indicadores del Observatorio: cambios y datos disponibles')
    parrafo(doc, f'Isabella Hoyos Álvarez · {FECHA} · Resultados **provisionales**')
    parrafo(doc, 'Revisión de los indicadores de la matriz, dimensión por dimensión, según lo acordado en las reuniones. '
                 'Criterios: que haya **serie en el tiempo**, que los datos sean **públicos y encontrables** sin solicitarlos a Planeación, '
                 'que **no se use autorreporte** y que la meta sea del **mismo universo** que el dato. '
                 'Donde no había datos, se buscó una variable parecida que sí los tuviera.')
    # sin la columna «Origen»: es solo para la trazabilidad de Isabella (va en el documento detallado)
    sin_origen = [c for c in range(len(ENC)) if ENC[c] != 'Origen']
    w = [4.8, 6.2, 7.6, 5.8]
    for titulo, filas in (('1. Cumplimiento del Plan de Desarrollo (PDM)', PDM), ('2. Participación ciudadana', PC),
                          ('3. Presupuesto participativo', PP), ('4. Control social', CS)):
        doc.add_heading(titulo, level=1)
        tabla(doc, w, [ENC[c] for c in sin_origen], [[f[c] for c in sin_origen] for f in filas])
        if filas is PDM:
            doc.add_heading('Variables para construir el índice del PDM', level=2)
            tabla(doc, [6.0, 8.4, 3.8, 6.2], ['Componente', 'Variables', 'Años con datos', 'Meta para pasarlo a puntaje 0–100'], PDM_COMP)
            doc.add_heading('Espacio público verde: datos y fuentes', level=2)
            tabla(doc, [1.6, 2.6, 2.8, 7.6, 9.8], EV_ENC, EV_CORTES)
            parrafo(doc, EV_NOTA)
        if filas is PC:
            doc.add_heading('I1, votación del presupuesto participativo: datos y fuentes', level=2)
            tabla(doc, [2.6, 2.8, 4.4, 2.8, 11.8], I1_ENC, I1_DATOS)
            parrafo(doc, I1_NOTA)
        if filas is CS:
            doc.add_heading('VRC, veedurías en la rendición de cuentas: rango y brecha de información', level=2)
            tabla(doc, [1.6, 2.0, 3.4, 2.4, 6.0, 2.8, 2.8, 3.4], VRC_ENC, VRC_DATOS)
            parrafo(doc, VRC_NOTA)
    doc.add_heading('5. ¿Se pueden construir los índices?', level=1)
    parrafo(doc, 'Todos los componentes tienen variables con datos públicos, pero ningún índice se puede calcular todavía sin resolver las dudas de abajo. '
                 'La matriz no permite calcular con componentes faltantes ni redistribuir pesos.')
    tabla(doc, [3.4, 8.8, 3.6, 8.6], ['Índice', 'Componentes con los cambios', 'Años con todos los componentes', 'Qué falta'], INDICES)
    doc.add_heading('6. Dudas para decidir', level=1)
    for seccion, items in leer_dudas():
        doc.add_heading(seccion, level=2)
        for it in items:
            escribir(doc.add_paragraph(), it)
    doc.add_heading('7. Variables por solicitar (informe futuro)', level=1)
    vinetas(doc, SOLICITAR)
    ruta = OUT / 'indicadores_cambios_y_datos_OCE.docx'
    doc.save(ruta)
    return ruta


# ------------------------------------------------------------------ documento 2: detalle para Isabella
DETALLE = {
    'Cumplimiento del Plan de Desarrollo (PDM)': [
        ('Cobertura de salud (antes: aseguramiento al régimen subsidiado)', {
            'Qué medía la matriz': 'Logro = 100 × cobertura observada del régimen subsidiado / meta (99,5 %). Línea base 99,5 % en 2022.',
            'Qué pasó y por qué': 'En la reunión 1 el profe dijo que, si el dato no estaba fácil, se reemplazara por discapacidad, inclusión o cobertura de salud. '
                                  'Elegiste la cobertura total, que cubre a toda la población. El 12-nov se confirmó, con la advertencia de que la meta debe ser del mismo universo.',
            'Datos': 'Afiliados al SGSSS (contributivo + subsidiado + especiales) / población DANE (una sola proyección). TerriData, archivo completo de Envigado. '
                     '2018: 82,71 % · 2019: 82,37 · 2020: 82,29 · 2021: 85,21 · 2022: 85,03 · 2023: 84,23 · 2024: 83,76 · 2025: 83,06.',
            'Limitaciones': 'La única meta del PDM es del subsidiado (99,5 %), así que no sirve para la cobertura total. Falta decidir el referente (duda 2).',
            'Estado': 'Construible en cuanto se decida el referente.'}),
        ('Mortalidad vial', {
            'Qué medía la matriz': 'Logro de reducción frente a la meta de 8 muertes por 100.000 habitantes (línea base 10,2 en 2023, Secretaría de Movilidad).',
            'Qué pasó y por qué': 'Se mantiene (12-nov). Se usa solo la Policía Nacional en toda la serie: Forensis 2025 no se puede consultar y mezclar fuentes cambia la medida (decisión de control de calidad).',
            'Datos': 'Muertes en accidentes de tránsito en Envigado (datos.gov.co, ha6j-pa2r) / población DANE. Tasas: 2019: 9,35 · 2020: 4,95 · 2021: 8,48 · '
                     '2022: 11,51 · 2023: 16,79 · 2024: 7,68 · 2025: 11,34.',
            'Limitaciones': 'La meta 8 es de Movilidad, otro universo. El salto de 2023 (43 muertes) no tiene explicación en los datos.',
            'Estado': 'Construible; falta decidir cómo usar una meta de otro universo (duda 4b).'}),
        ('Espacio público verde', {
            'Qué medía la matriz': 'Logro = espacio verde observado / meta del PDM (3,88 m² por habitante, estable).',
            'Qué pasó y por qué': 'El 12-nov el profe cuestionó que diera 100 sin exigencia y propuso medir el avance desde la línea base del AMVA (4,66 m², 2022) hacia el referente de 9 m². '
                                  'El 15-nov decidiste usar el espacio verde URBANO y tomar como línea base el corte de 2019 (4,54, AMVA), el primero con la metodología del AMVA, para ver la evolución en 2022 y 2024. '
                                  'No se espera una serie continua: el dato no se publica cada año.',
            'Datos': 'Espacio verde urbano en tres cortes: 2019: 4,54 (AMVA con la Universidad Nacional, Reporte ODS 11.7, tabla 4); 2022: 4,66 (AMVA con el Instituto Humboldt; tabla 39 del POT); 2024: 5,75 = 1.382.762 m² / 240.655 habitantes urbanos (diagnóstico del POT, tabla 39). Primer dato propio de Envigado: 2013, 4,16 (Minambiente, ICAU). Contexto: verde total 2024 = 69,41 (incluye las reservas rurales; el POT advierte que está distorsionado). No hay variable de espacio verde en TerriData.',
            'Limitaciones': 'Los tres cortes no son comparables: distintos productores, poblaciones y delimitación (2019 incluye suelo de expansión; 2024 solo suelo urbano). El 4,54 no se puede reproducir con los datos del reporte. Se descartan el 3,88 del PDM (sin método declarado) y los 5,00 de 2020 y 2023 (autorreporte). El diagnóstico del POT da 4,22 en una página y 4,66 en su tabla 39.',
            'Avance hacia 9 m² (línea base 2019 = 4,54)': '2019: 0 % · 2022: (4,66 − 4,54) / (9 − 4,54) = 2,7 % · 2024: (5,75 − 4,54) / (9 − 4,54) = 27,1 %. En valores: +0,12 m² (2019–2022) y +1,09 m² (2022–2024).',
            'Fuentes (trazabilidad)': ' · '.join(EV_URLS),
            'Recomendación para el informe': 'Solicitar a Planeación la serie completa desde 2019 con la metodología del AMVA. Proponer que Planeación la mida cada año en el seguimiento del POT (expediente municipal) y que la Secretaría de Medio Ambiente, responsable del indicador en el PDM, la reporte con ese mismo método.',
            'Estado': 'Construible para 2019 (línea base), 2022 y 2024 como cortes, no como tendencia.'}),
        ('Veedurías ciudadanas inscritas (antes: implementación de procesos participativos)', {
            'Qué medía la matriz': 'Logro de implementación de los procesos de participación ciudadana (meta 100 %, Planeación). Era un autorreporte y no había cifra oficial.',
            'Qué pasó y por qué': 'Tus notas dicen «cambiar por número de veedurías ciudadanas»; confirmaste que era para este indicador, cuya categoría en la matriz es «Participación ciudadana».',
            'Datos': 'Conjunto «Veedurías Ciudadanas Personería Envigado» (datos.gov.co, jju5-br3w). Inscritas por año: 2019: 2 · 2020: 4 · 2021: 11 · 2022: 4 · 2023: 6 · 2024: 1 · 2025: 5.',
            'Limitaciones': 'Los consecutivos van del 118 al 293, pero solo hay 45 registros: el listado parece incompleto. No hay meta en el PDM (duda 3). '
                            'El conjunto trae nombres y correos: solo se usan conteos.',
            'Estado': 'Construible con un referente; conviene pedir a la Personería el listado completo.'}),
        ('Servicios públicos domiciliarios (antes: cobertura deportiva)', {
            'Qué medía la matriz': 'Logro de cobertura de servicios deportivos (meta 6 %, Índer). No hay tasa oficial publicada.',
            'Qué pasó y por qué': 'El 12-nov se acordó trabajar con acueducto, alcantarillado, energía y aseo, con cobertura total (no solo urbana) y sin internet. '
                                  'Recordaste que reemplaza a deporte. Se propone promedio simple, como el resto de la matriz.',
            'Datos': 'TerriData (REC de Superservicios; UPME/IPSE para energía). Promedio de las 4 coberturas: 2019: 91,97 · 2020: 77,22 · 2021: 89,74 · 2022: 89,56 · 2023: 89,76 · 2024: 89,31.',
            'Limitaciones': 'Caída a ~68 % en 2020 y zona rural en 0 de 2016 a 2020 (problemas de registro). Cobertura rural baja (acueducto 17,9 % en 2024), '
                            'probablemente porque los acueductos veredales no están registrados. La meta (99,9 %) viene de la Encuesta de Calidad de Vida: otro universo. El aseo no tiene meta.',
            'Estado': 'Construible; falta decidir el uso de la meta (dudas 4b y 5).'}),
        ('Índice de logro educativo', {
            'Qué medía la matriz': 'ILE = (logro de atención integral a primera infancia + logro de colegios oficiales que mejoran su categoría en Saber 11) / 2, con tope de 100 en cada parte.',
            'Qué pasó y por qué': 'En la reunión 1 el profe pidió transformarlo para ver la evolución en años. El 12-nov encontraste en TerriData la cobertura neta y se acordó combinarla con primera infancia. '
                                  'Decidiste usar 0–5 años porque reproduce mejor la línea base del PDM (16,56 frente a 16,3; con 0–6 da 14,02).',
            'Datos': 'Cobertura neta (TerriData/MinEducación) y niños atendidos / población 0–5 (TerriData, DANE). ILE: 2018: 83,39 · 2019: 83,48 · 2020: 86,24 · 2021: 89,18 · 2022: 90,91 · 2023: 93,21 · 2024: 95,15.',
            'Limitaciones': 'El numerador de primera infancia incluye mujeres gestantes y la fuente cambia en 2024 (informe de gestión). La cobertura neta de 2025 no está publicada.',
            'Estado': 'Construible 2018–2024.'}),
        ('Sexta medida', {
            'Qué pasó y por qué': 'El 12-nov se acordaron cinco líneas y quedó abierta una sexta. Con las veedurías en el lugar de procesos participativos, el índice ya tiene seis componentes; '
                                  'falta saber si el profe quiere además una medida de seguridad (duda 1).',
            'Datos': 'Policía/MinDefensa, 2019–2025 (tasas por 100.000). Homicidio: 4,25 · 2,89 · 1,21 · 3,57 · 3,52 · 4,61 · 3,02 (meta 2,7). '
                     'Hurto a personas: 680,1 … 501,4 (meta 631,5). Lesiones personales: 159,7 … 100,1 (meta 112,7).',
            'Limitaciones': 'Hurto y lesiones ya están por debajo de su meta desde 2024: darían 100 fijo. El homicidio tiene pocos casos al año y la tasa salta mucho.',
            'Estado': 'Construible si el profe lo pide; se recomienda homicidio.'}),
    ],
    'Participación ciudadana': [
        ('I1: votación del presupuesto participativo (antes: incorporación de aportes)', {
            'Qué medía la matriz': 'Procesos cerrados que incorporaron un aporte ciudadano / procesos cerrados con aportes.',
            'Qué pasó y por qué': 'No hay inventario de procesos. El respaldo que se tenía usaba los mismos datos de la fidelidad de PP, y el profe pidió diferenciarlos. '
                                  'Elegiste la opción A: la fidelidad se queda solo en PP y I1 pasa a medir la participación en la votación (IPCM, el índice oficial que usa la Alcaldía). '
                                  'El 15-nov elegiste medirlo como avance desde la línea base 2017 (3,62 %) hacia el referente de 7 %, el tope de la escala oficial del IPCM, porque el PDM no tiene meta.',
            'Datos': '2017: 8.221 votos, IPCM 3,62 % (informe de Planeación, 11-oct-2017, p. 7 y 12). 2019: no hubo votación. 2025: 10.662 votantes (informe de gestión 2025, p. 92 del PDF; verificado el 15-nov) / 264.637 habitantes (DANE) = 4,03 %.',
            'Avance hacia 7 % (línea base 2017 = 3,62 %)': '2017: 0 % · 2025: (4,03 − 3,62) / (7 − 3,62) = 12,1 %. Puntaje simple (IPCM / 7): 51,7 en 2017 y 57,6 en 2025.',
            'Limitaciones': 'Solo dos años confiables. 2017 cuenta votos y 2025 votantes. El 3,62 % no se reproduce con la población del DANE (da 3,53 o 3,68): la Alcaldía usa su propia población. El 7 % es el tope de la tabla de reparto, no una meta adoptada.',
            'Fuentes (trazabilidad)': ' · '.join(I1_URLS),
            'Recomendación para el informe': 'Solicitar a Planeación los votos y el IPCM de 2018 a 2024, total y por zona, la población de referencia y si cuentan votos o votantes. Con los datos por zona, I1 se puede cruzar con I2.',
            'Estado': 'Construible para 2017 (línea base) y 2025; hay que solicitar 2018–2024.'}),
        ('I2: cobertura territorial de la participación', {
            'Qué medía la matriz': 'Zonas con al menos una actividad participativa / 13 zonas oficiales.',
            'Qué pasó y por qué': 'Se mantiene (tu nota). Se agregó un respaldo con serie larga: zonas con proyectos de PP en los listados de la EFPC.',
            'Datos': 'Principal: 2024: 13/13; 2025: 12/13. Respaldo 2020–2025: 12/13 todos los años (falta siempre la zona 1, que es industrial según el profe).',
            'Limitaciones': 'Falta decidir si el denominador es 13 o 12 zonas (duda 7b). Los listados de la EFPC se titulan «proyectos ejecutados».',
            'Estado': 'Construible 2024–2025; respaldo 2020–2025.'}),
        ('I3: oferta formativa', {
            'Qué medía la matriz': 'Acciones formativas para la ciudadanía realizadas / programadas.',
            'Qué pasó y por qué': 'Se mantiene, con el número de acciones como respaldo.',
            'Datos': 'Tasa 2025: 20/24 (83,3 %), con denominador reconstruido. Conteo: 2023: 17 (informe de Control Interno, proyecto 0541); 2024: 7; 2025: 20.',
            'Limitaciones': 'Los criterios de conteo no son iguales (2024 incluye acciones para servidores). El dato de 2023 lo reportó Planeación sin evidencias cargadas.',
            'Estado': 'Construible solo para 2025; el conteo hay que unificarlo.'}),
        ('I4: publicación de resultados participativos', {
            'Qué medía la matriz': 'Procesos cerrados con resultados publicados / procesos cerrados.',
            'Qué pasó y por qué': 'Se descartó el Índice de Participación del DNP porque mide participación electoral. Decidiste conservar I4 con una lista fija de procesos anuales que verifica el observatorio.',
            'Datos': 'Lista: votación de PP, rendición de cuentas del Alcalde, rendición de cuentas de PP. 2021: 3/3 · 2022: 3/3 (por confirmar) · 2023: 3/3 · 2024: 3/3 · 2025: 2/2.',
            'Limitaciones': 'Con tres procesos cada uno pesa 33 puntos. Falta confirmar en el portal las respuestas de la rendición del Alcalde 2022.',
            'Estado': 'Construible 2021–2025.'}),
        ('I5 e I6', {'Qué pasó y por qué': 'Eliminados por decisión del profe en la reunión 1.', 'Estado': 'Fuera.'}),
    ],
    'Presupuesto participativo': [
        ('EPP: ejecución presupuestal', {
            'Qué medía la matriz': 'Obligaciones / apropiación definitiva de los recursos de PP.',
            'Qué pasó y por qué': 'Se mantiene; el 12-nov el indicador de PP se dio por definido.',
            'Datos': '2024: 98,69 % (50 de 51 proyectos cruzados con Hacienda); 2025: 98,83 %; 2026, primer semestre: 23,26 %.',
            'Limitaciones': 'Solo 2024–2025: la ejecución de Hacienda de 2020–2023 viene por rubro presupuestal, no por proyecto.',
            'Estado': 'Construible 2024–2025.'}),
        ('CEP → ejecución según los proyectos que sí se hicieron', {
            'Qué medía la matriz': 'Proyectos con todas las entregas completas y verificadas / proyectos con entrega exigible.',
            'Qué pasó y por qué': 'No hay actas de recibo. El profe propuso medir la ejecución según los proyectos que sí se hicieron (tu nota «algo sobre ejecución de presupuesto»). '
                                  'Como el único dato de «terminado» es autorreporte, se propone como principal la proporción de proyectos con ejecución ≥ 90 % (Hacienda).',
            'Datos': 'Ejecución ≥ 90 %: 2024: 48/50 (96 %); 2025: 57/61 (93,4 %). Terminados (autorreporte): 2025: 49/61 (80,3 %).',
            'Limitaciones': 'Ejecución presupuestal no es lo mismo que entrega física. Falta la decisión del profe (duda 9).',
            'Estado': 'Construible 2024–2025.'}),
        ('CMA → fidelidad a lo votado', {
            'Qué medía la matriz': 'Personas atendidas frente a la meta de atención de cada proyecto.',
            'Qué pasó y por qué': 'El plan de acción mide actividades, no personas. Se reemplaza por la fidelidad a lo votado, la propuesta del profe.',
            'Datos': 'Proyectos cuya apropiación coincide con el valor votado (±1 %) / proyectos votados. 2024: 39/51 (76,47 %); 2025: 58/61 (95,08 %).',
            'Limitaciones': '2026 necesita el valor votado por proyecto.',
            'Estado': 'Construible 2024–2025.'}),
    ],
    'Control social': [
        ('PCS → oportunidad de respuesta a las PQRSDF', {
            'Qué medía la matriz': 'Estrategias de promoción del control social ejecutadas / programadas.',
            'Qué pasó y por qué': 'Es un autorreporte que da 100 % (1 de 1), y dijiste que eso no sirve. Se propone la oportunidad de respuesta a las PQRSDF, que sale del sistema de registro.',
            'Datos': 'Informes trimestrales de PQRSDF (Secretaría General), años completos 2020–2025. Verificado: 4.º trimestre de 2023: 3.544 / 4.149 = 85 %.',
            'Limitaciones': 'El total incluye solicitudes pendientes dentro del plazo; el informe usa dos sistemas (BPMS y SAC); algunas se cierran tarde en el sistema.',
            'Estado': 'Construible con trabajo: sumar los trimestres con una regla fija.'}),
        ('DSI: información de seguimiento', {
            'Qué medía la matriz': 'Metas del PDM con valor observado publicado / metas seleccionadas.',
            'Qué pasó y por qué': 'Se mantiene. Hay que actualizar la lista de metas, usando indicadores que el PDM sí tiene con meta.',
            'Datos': '2024: 0/7; 2025: 0/7. Con todo lo publicado hasta hoy: 2/7 en 2024.', 'Estado': 'Construible 2024–2025.'}),
        ('VRC: veedurías en la rendición de cuentas', {
            'Qué medía la matriz': 'Eventos de rendición de cuentas con veeduría documentada / eventos.',
            'Qué pasó y por qué': 'Se mantiene, pero en varios eventos no se sabe si hubo veeduría: la Alcaldía publica cuadros de preguntas o presentaciones, casi nunca el acta con asistentes. '
                                  'La audiencia del Alcalde de 2024 y la de 2025 quedan sin dato. Con la regla de la matriz (presencia desconocida = N/D), VRC no tiene ningún año. '
                                  'El 15-nov preguntaste cómo construirlo con la brecha y decidiste dejarlo así: se reporta como rango (mínimo y máximo) con la brecha de información al lado, y se espera que la información mejore año a año. Se le informa al profe, porque cambia la regla N/D de la matriz.',
            'Datos': 'Eventos: audiencia del Alcalde + rendiciones sectoriales con acta o memoria. 2024: 3 eventos (2 con veeduría, 0 sin, 1 sin soporte) → 66,7 %–100 %, brecha 33,3 %. '
                     '2025: 5 eventos (1 con, 1 sin, 3 sin soporte) → 20 %–80 %, brecha 60 %. Con todos los eventos publicados (sensibilidad): 2024: 8 eventos → 50 %–100 % (brecha 50 %); 2025: 7 → 28,6 %–85,7 % (brecha 57,1 %).',
            'Cómo se construye': 'VRC mínimo = documentada / eventos (es el que entraría al índice; equivale al respaldo anterior «sin mención = ausencia»). VRC máximo = (documentada + sin soporte) / eventos. Brecha = sin soporte / eventos. '
                                 'Clasificación por evento: outputs/indicadores/clasificacion_vrc.csv (con documento, página, enlace y cita). Cálculo: scripts/revision_indicadores/vrc_brecha.py → outputs/revision_indicadores/CS/vrc_brecha.csv.',
            'Qué se necesita para seguir midiéndolo': ' · '.join(VRC_NECESITA),
            'Limitaciones': 'El mínimo subestima si hubo veedores que no quedaron registrados. El universo cambia según qué eventos se cuenten (3 o 8 en 2024). Una invitación a veedores no prueba asistencia.',
            'Cómo se hizo (nota de método)': 'Se revisaron los documentos publicados de cada evento de rendición de cuentas de 2024 y 2025 (actas, memorias, presentaciones, cuadros de preguntas y justificaciones). Cada evento se clasificó como «con veeduría» si el documento lo dice expresamente, «sin veeduría» si el documento integral del evento la describe sin veedurías, y «no se sabe» si el documento no permite saberlo. Una invitación no cuenta como asistencia. Con esa clasificación se calculan el mínimo, el máximo y la brecha.',
            'Seguimiento': 'La brecha se reporta cada año y debe bajar: es la medida de cuánto mejora la información. Cuando la brecha llegue a 0, el mínimo y el máximo coinciden y VRC vuelve a ser un solo número, como pide la matriz.',
            'Estado': 'Construible como rango para 2024 y 2025. Decisión de Isabella (15-nov); se le informa al profe en la duda 14.'}),
        ('RPO: respuesta a observaciones ciudadanas', {
            'Qué medía la matriz': 'Observaciones ciudadanas con respuesta de fondo publicada / observaciones.',
            'Qué pasó y por qué': 'Se mantiene.', 'Datos': '2024: 16/17 (94,12 %). 2025: cuadro incompleto (respaldo 1/2). Hay cuadros publicados de 2021 y 2023.',
            'Estado': 'Construible 2024; 2021 y 2023 con trabajo.'}),
    ],
}

QA_RESUMEN = [
    '**Veredicto global:** la propuesta sirve como inventario de datos disponibles. Ningún índice se puede calcular todavía sin resolver las dudas del profe.',
    '**IREPP (PP)** es el más cercano: resueltas las dudas 9 y 10, se puede construir para 2024 y 2025.',
    '**ICPD-6 (PDM)** saldría solo para 2019, 2022 y 2024 (los cortes de espacio verde); **ISPC (PC)** solo para 2025; **ICRCS (CS)** no tiene ningún año sin usar respaldos.',
    '**No construibles con los datos públicos:** I1 original, CEP y CMA originales y logros de políticas públicas. VRC solo como rango con la brecha de información (propuesta).',
    'QA corrigió 13 errores en los archivos de entrega (años, metas de otro universo, criterios de conteo, títulos de los listados de la EFPC). Ya están corregidos.',
    'Los datos crudos de 2024–2025 que usaron los cálculos anteriores no están en este Codespace (están en Google Drive); las descargas de hoy quedaron en data_raw_IIC/07_revision_nov2026/.',
]


def doc_detalle():
    doc = nuevo_documento(horizontal=False)
    doc.add_paragraph(style='Title').add_run('Revisión de indicadores: qué pasó con cada uno')
    parrafo(doc, f'Documento de trabajo para Isabella · {FECHA} · Todo es **provisional**')
    doc.add_heading('Cómo se hizo la revisión', level=1)
    vinetas(doc, [
        'Se partió de la matriz de los profesores, de tus notas de las reuniones 1 y 2 y del informe de la reunión del 12-nov.',
        'Se revisó dimensión por dimensión: PDM, Participación ciudadana, Presupuesto participativo y Control social.',
        'Criterios: serie en el tiempo, datos públicos y encontrables, sin autorreporte, meta del mismo universo y las mismas dimensiones.',
        'Los datos se buscaron en TerriData, datos.gov.co, MinEducación, la Policía, la Personería, la EFPC y los informes de la Alcaldía.',
        'Cada paso lo revisó un agente de control de calidad (QA), que verificó cifras contra la fuente y encontró errores que ya se corrigieron.',
        'Los archivos de apoyo están en outputs/entrega_profe/ (cuadros, dudas y registro de cambios) y outputs/revision_indicadores/ (series, scripts y reportes de QA).',
    ])
    doc.add_heading('Resultado del control de calidad: ¿se pueden construir?', level=1)
    vinetas(doc, QA_RESUMEN)
    tabla(doc, [2.6, 6.2, 3.0, 5.6], ['Índice', 'Componentes', 'Años con todos', 'Qué falta'], INDICES, tam=8)
    for dim, indicadores in DETALLE.items():
        doc.add_heading(dim, level=1)
        for nombre, campos in indicadores:
            doc.add_heading(nombre, level=2)
            for campo, texto in campos.items():
                escribir(doc.add_paragraph(), f'**{campo}:** {texto}')
    doc.add_heading('Variables por solicitar (informe futuro)', level=1)
    vinetas(doc, SOLICITAR)
    ruta = OUT / 'detalle_revision_indicadores_OCE.docx'
    doc.save(ruta)
    return ruta


def main():
    for ruta in (doc_profe(), doc_detalle()):
        d = Document(ruta)
        print(ruta.relative_to(RAIZ), '| párrafos:', len(d.paragraphs), '| tablas:', len(d.tables))


if __name__ == '__main__':
    main()
