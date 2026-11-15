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
     '**Se ajusta**: avance desde 4,66 m²/hab (AMVA, 2022) hacia 9 m²/hab', 'Profe (12-nov)', 'Solo 2022 (4,66) y 2024 (5,75, POT)'],
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
    ['Espacio verde', 'm² verdes por habitante', '**Solo 2022 y 2024**', 'Sí (avance hacia 9 m²)'],
    ['Veedurías (en vez de procesos participativos)', 'Veedurías inscritas por año', '2012–2026', 'No hay meta en el PDM'],
]

PC = [
    ['**I1** Incorporación de aportes', 'No hay inventario de procesos; el respaldo usaba los mismos datos de la fidelidad de PP',
     '**Se reemplaza** por la votación del presupuesto participativo (IPCM = votos / población), para diferenciarlo de PP',
     'Profe (diferenciar); Isabella y Claude', '2017 (3,62 %) y 2025 (10.662 votantes); 2019 sin votación; 2018–2024 por solicitar'],
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
    ['**VRC** Veedurías en rendición de cuentas', 'En varios eventos no se sabe si hubo veeduría', '**Se mantiene**', 'Profe',
     'Sin dato principal; respaldo 2024 (66,7 %) y 2025 (20 %)'],
    ['**RPO** Respuesta a observaciones', 'El cuadro de 2025 está incompleto', '**Se mantiene**', 'Profe', '2024: 94,1 %; 2025 frágil; ampliable a 2021 y 2023'],
]

INDICES = [
    ['**ICPD-6** (PDM)', 'Bienestar = salud + logro educativo + servicios públicos; territorio = mortalidad vial + espacio verde + veedurías',
     '2022 y 2024', 'Dudas 1 a 5: referentes de salud y veedurías, metas de otro universo. Espacio verde año por año'],
    ['**ISPC** (PC)', '0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4', 'Solo 2025', 'Dudas 6 a 8: I1, denominador de I2 y pesos'],
    ['**IREPP** (PP)', '0,20·EPP + 0,40·ejecución ≥ 90 % + 0,40·fidelidad', '2024 y 2025', 'Dudas 9 y 10. Resueltas, se puede construir ya'],
    ['**ICRCS** (CS)', '(DSI + PQRSDF + VRC + RPO) / 4', 'Ninguno sin respaldos', 'VRC sin dato principal y RPO 2025 incompleto. Dudas 12 a 14'],
]

SOLICITAR = [
    'Espacio público verde por habitante, año por año (AMVA o Planeación).',
    'Votos de presupuesto participativo e IPCM por año y por zona, 2018–2024 (Planeación).',
    'Apropiación, compromisos y obligaciones por proyecto de presupuesto participativo, 2020–2023 (Hacienda o Planeación).',
    'Listado histórico completo de veedurías inscritas (Personería).',
    'Programación anual de la oferta formativa y actividades participativas por zona (Escuela de Formación, Participación).',
    'Inventario de procesos participativos con los aportes recibidos y la decisión tomada.',
    'Actas de recibo y personas atendidas por proyecto de presupuesto participativo.',
    'Presencia de veedurías en las rendiciones de cuentas y cuadro completo de preguntas y respuestas de 2025.',
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
            'Qué pasó y por qué': 'El 12-nov el profe cuestionó que diera 100 sin exigencia y propuso medir el avance desde la línea base del AMVA (4,66 m², 2022) hacia el referente de 9 m².',
            'Datos': '2022: 4,66 (AMVA, tabla 39 del diagnóstico del POT); 2024: 5,75 (diagnóstico del POT). Avance 2024 ≈ 25 %. No hay variable de espacio verde en TerriData ni en los informes de gestión 2024–2025.',
            'Limitaciones': 'Hay tres cifras que no coinciden (3,88 del PDM, 4,66 del AMVA y 5,75 del POT). Solo dos años.',
            'Estado': 'Construible para 2022 y 2024. Para tener serie hay que solicitar el dato año por año.'}),
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
                                  'Elegiste la opción A: la fidelidad se queda solo en PP y I1 pasa a medir la participación en la votación (IPCM, el índice oficial que usa la Alcaldía).',
            'Datos': '2017: 8.221 votos, IPCM 3,62 % (informe de Planeación, 11-oct-2017). 2019: no hubo votación. 2025: 10.662 votantes (informe de gestión 2025).',
            'Limitaciones': 'Solo dos años confiables. 2025 fue el primer año con la nueva plataforma web. El 3,62 % no se reproduce con la población del DANE.',
            'Estado': 'Construible solo para algunos años; hay que solicitar los votos 2018–2024.'}),
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
            'Qué pasó y por qué': 'Se mantiene, pero en varios eventos no se sabe si hubo veeduría.',
            'Datos': 'Respaldo (sin mención = ausencia): 2024: 2/3 (66,7 %); 2025: 1/5 (20 %).',
            'Limitaciones': 'Con la regla de la matriz no hay dato principal en ningún año.', 'Estado': 'No construible hoy sin respaldos.'}),
        ('RPO: respuesta a observaciones ciudadanas', {
            'Qué medía la matriz': 'Observaciones ciudadanas con respuesta de fondo publicada / observaciones.',
            'Qué pasó y por qué': 'Se mantiene.', 'Datos': '2024: 16/17 (94,12 %). 2025: cuadro incompleto (respaldo 1/2). Hay cuadros publicados de 2021 y 2023.',
            'Estado': 'Construible 2024; 2021 y 2023 con trabajo.'}),
    ],
}

QA_RESUMEN = [
    '**Veredicto global:** la propuesta sirve como inventario de datos disponibles. Ningún índice se puede calcular todavía sin resolver las dudas del profe.',
    '**IREPP (PP)** es el más cercano: resueltas las dudas 9 y 10, se puede construir para 2024 y 2025.',
    '**ICPD-6 (PDM)** saldría solo para 2022 y 2024 (por espacio verde); **ISPC (PC)** solo para 2025; **ICRCS (CS)** no tiene ningún año sin usar respaldos.',
    '**No construibles con los datos públicos:** I1 original, CEP y CMA originales, VRC principal y logros de políticas públicas.',
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
