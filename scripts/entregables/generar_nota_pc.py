"""Genera la nota técnica de participación ciudadana (Word) para el profesor Eber Gutiérrez.

Uso (desde la raíz del repo, sin -I):  python3 scripts/entregables/generar_nota_pc.py
Estilo visual tomado del informe de disponibilidad (Cambria 11, interlineado 1,15, tablas tipo booktabs,
encabezado con número de página, colores de semáforo). Las cifras de resultados se leen de
outputs/indicadores/resultados_indicadores.csv y valores_variables.csv; no se recalcula nada.
"""
import csv, re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parents[2]
IND = RAIZ / 'outputs' / 'indicadores'
SALIDA = IND / 'entregables' / 'nota_tecnica_participacion.docx'
FUENTE = 'Cambria'
MARGEN = 2.3
ANCHO = 21.59 - 2 * MARGEN
COL = {'Verde': 'D6F0D6', 'Amarillo': 'FFF3BF', 'Naranja': 'FFDEB8', 'Rojo': 'F8D4D4', 'Gris': 'E7E6E6'}
MIR = ('https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/'
       'alcaldia-municipal-de-envigado-en-antioquia/content/files/')
URL = {
    'PDM': MIR + '002562/128075_plandesarrollo20242027.pdf',
    'G25': MIR + '003058/152890_2025informegestiondependencias.pdf',
    'G24': MIR + '002640/131997_24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf',
    'PI': MIR + '003611/180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx',
    'PPSS': MIR + '003452/172562_pps190ppps20251231mu000000005266.xlsx',
    'QA24': MIR + '002595/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf',
    'POT': MIR + '002670/133472_cronograma-asambleas-pot.pdf',
    'CAL': MIR + '002443/122120_calendarioparticipacionciudadana2024.pdf',
    'ANTE': MIR + '002315/115749_anteproyecto29022024.pdf',
    'PP26': MIR + '003451/172548_asambleas-zonales-y-puntos-descentralizados-pp2027.pdf',
}


# ------------------------------------------------------------------ lectura de datos
def leer(nombre):
    with open(IND / nombre, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f, delimiter=';'))


RES = leer('resultados_indicadores.csv')
VAR = leer('valores_variables.csv')


def fila(codigo, anio, esc='defecto', terr='Municipio', periodo='Anual'):
    c = [r for r in RES if r['codigo'] == codigo and r['anio'] == str(anio) and r['territorio'] == terr
         and r['periodo'] == periodo and (r['es_defecto'] == 'True' if esc == 'defecto' else r['escenario_id'] == esc)]
    assert len(c) == 1, (codigo, anio, esc, terr, len(c))
    return c[0]


def num(x):
    x = float(x)
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f'{x:.2f}'.replace('.', ',')


def punt(r):
    """Texto del puntaje: número, 'Revisar' o N/D."""
    if r['puntaje'] != '':
        return num(r['puntaje'])
    return 'Revisar' if r['disponibilidad'] == 'Revisar' else 'N/D'


def entradas(r):
    return r['entrada_1'], r['entrada_2']


def variable(idv, anio, terr='Municipio'):
    c = [r for r in VAR if r['id_variable'] == idv and r['anio'] == str(anio) and r['territorio'] == terr
         and r['es_defecto'] == 'True' and r['periodo'] in ('Anual', '')]
    assert len(c) == 1, (idv, anio, len(c))
    return c[0]['valor']


# cifras de resultados (todas desde los CSV)
I2 = {a: fila('PC.I2', a) for a in (2024, 2025)}
I3 = {a: fila('PC.I3', a) for a in (2024, 2025)}
I5 = {a: fila('PC.I5', a) for a in (2024, 2025)}
assert (I2[2024]['semaforo'], I2[2025]['semaforo'], I3[2025]['semaforo'], I5[2025]['semaforo']) == ('Verde',) * 4
for cod in ('PC.I1', 'PC.I4', 'PC.I6'):
    for a in (2024, 2025):
        assert fila(cod, a)['puntaje'] == ''
I2_PP = fila('PC.I2', 2024, 'pc_i2_actividades=asambleas_pp')
I2_12 = fila('PC.I2', 2025, 'zonas_oficiales=12_con_comite')
I2_PROG = fila('PC.I2', 2025, 'pc_i2_evidencia=incluye_programacion')
I2_ACTA = fila('PC.I2', 2025, 'pc_i2_evidencia=acta_por_zona')
I3_PERS = {a: fila('PC.I3', a, 'pc_i3_unidad=personas_producto_521') for a in (2024, 2025)}
I3_SERV = fila('PC.I3', 2025, 'pc_i3_excluir_servidores=false')
I5_AMP = fila('PC.I5', 2025, 'pc_i5_instancias=ampliado_cmpc_ctp_comites')
EPP = {a: fila('PP.EPP', a, periodo='Anual') for a in (2024, 2025)}
EPP26 = fila('PP.EPP', 2026, periodo='Semestre 1')
EPP_Z10 = fila('PP.EPP', 2024, terr='Zona 10')
EPP_EX = fila('PP.EPP', 2024, 'pp_2024_metodo=solo_exactos')
EPP_0553 = fila('PP.EPP', 2025, 'pp_incluir_0553=true')
CP = {a: fila('PDM.CP', a) for a in (2024, 2025)}
assert all(r['puntaje'] == '' for r in CP.values())
sem_libres = I3[2025]['entrada_1'], I3[2025]['entrada_2']
V_I6 = (variable('PC.I6.V1', 2024), variable('PC.I6.V1', 2025))
V_PROC = (variable('PC.I1.V2', 2024), variable('PC.I1.V2', 2025))
V_PE = tuple(variable('PP.CEP.PE', a) for a in (2024, 2025, 2026))
V_PEC = variable('PP.CEP.PEC', 2025)
V_INDSIT = [r['valor'] for r in VAR if r['id_variable'] == 'AUX.PDM.CP.INDICE_SITUACIONAL'][0]
assert V_PROC == ('3', '5') and V_I6 == ('1', '1')
# sensibilidad de semilleros: 5 ejecutados de 7 planeados que se suman a 20/24
SEMILLEROS = 100 * (int(sem_libres[0]) + 5) / (int(sem_libres[1]) + 7)


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


def estilos(doc):
    st = doc.styles
    fijar_fuente(st['Normal'], FUENTE, 11)
    st['Normal'].paragraph_format.line_spacing = 1.15
    st['Normal'].paragraph_format.space_after = Pt(5)
    for n, t in (('Title', 17), ('Heading 1', 14), ('Heading 2', 12)):
        fijar_fuente(st[n], FUENTE, t, True, RGBColor(0, 0, 0))
        st[n].paragraph_format.space_before = Pt(11 if n == 'Heading 1' else 8)
        st[n].paragraph_format.space_after = Pt(4)
        st[n].paragraph_format.keep_with_next = True
    ppr = st['Title'].element.get_or_add_pPr()
    for b in ppr.findall(qn('w:pBdr')):
        ppr.remove(b)
    for n in ('List Bullet', 'List Number'):
        fijar_fuente(st[n], FUENTE)
        st[n].paragraph_format.space_after = Pt(3)


def sombrear(pr, color):
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), color)
    pr.append(s)


def borde(pr, etiqueta, lados, color='auto'):
    b = pr.find(qn(etiqueta))
    if b is None:
        b = OxmlElement(etiqueta)
        pr.append(b)
    for lado, (val, sz) in lados.items():
        e = OxmlElement('w:' + lado)
        e.set(qn('w:val'), val)
        if val != 'nil':
            e.set(qn('w:sz'), str(sz))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), color)
        b.append(e)


TOK = re.compile(r'(\*\*.+?\*\*|\[[^\]]+\]\([^)]+\))')


def escribir(p, texto, tam=None, base_b=False):
    """Mini marcado: **negrita** y [texto](clave_url|url)."""
    for parte in TOK.split(texto):
        if not parte:
            continue
        if parte.startswith('**'):
            r = p.add_run(parte[2:-2])
            r.bold = True
            if tam:
                r.font.size = Pt(tam)
        elif parte.startswith('['):
            m = re.match(r'\[([^\]]+)\]\(([^)]+)\)', parte)
            url = URL.get(m.group(2), m.group(2))
            rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
            hl = OxmlElement('w:hyperlink')
            hl.set(qn('r:id'), rid)
            r = p.add_run(m.group(1))
            r.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            r.font.underline = True
            if tam:
                r.font.size = Pt(tam)
            hl.append(r._r)
            p._p.append(hl)
        else:
            r = p.add_run(parte)
            if base_b:
                r.bold = True
            if tam:
                r.font.size = Pt(tam)


def parrafo(doc, texto, **kw):
    p = doc.add_paragraph()
    escribir(p, texto)
    for k, v in kw.items():
        setattr(p.paragraph_format, k, v)
    return p


def lista(doc, items, numerada=False):
    for t in items:
        p = doc.add_paragraph(style='List Number' if numerada else 'List Bullet')
        escribir(p, t)


def tabla(doc, anchos, enc, filas, tam=8.5):
    """filas: lista de listas; cada celda es texto o (texto, color_relleno)."""
    t = doc.add_table(rows=1 + len(filas), cols=len(anchos))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tp = t._tbl.tblPr
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tp.append(lay)
    borde(tp, 'w:tblBorders', {'top': ('single', 12), 'left': ('nil', 0), 'bottom': ('single', 12),
                               'right': ('nil', 0), 'insideH': ('nil', 0), 'insideV': ('nil', 0)})
    tot = sum(anchos)
    anchos = [a * ANCHO / tot for a in anchos]
    for k, gc in enumerate(t._tbl.tblGrid.findall(qn('w:gridCol'))):
        gc.set(qn('w:w'), str(int(anchos[k] / 2.54 * 1440)))
    todas = [enc] + filas
    for i, celdas in enumerate(todas):
        tr = t.rows[i]._tr
        trpr = tr.get_or_add_trPr()
        trpr.append(OxmlElement('w:cantSplit'))
        if i == 0:
            trpr.append(OxmlElement('w:tblHeader'))
        for k, c in enumerate(celdas):
            relleno = None
            if isinstance(c, tuple):
                c, relleno = c
            cell = t.rows[i].cells[k]
            cell.width = Cm(anchos[k])
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_after = pf.space_before = Pt(1)
            pf.line_spacing = 1.0
            for j, linea in enumerate(c.split('\n')):
                if j:
                    p = cell.add_paragraph()
                    p.paragraph_format.space_after = p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.line_spacing = 1.0
                escribir(p, linea, tam, base_b=(i == 0))
            tcpr = cell._tc.get_or_add_tcPr()
            if relleno:
                sombrear(tcpr, COL.get(relleno, relleno))
            if i == 0:
                borde(tcpr, 'w:tcBorders', {'bottom': ('single', 4)})
                sombrear(tcpr, 'DCE6F1')
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 0.6
    return t


def encabezado(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
    for l in ('left_margin', 'right_margin'):
        setattr(sec, l, Cm(MARGEN))
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
    p = sec.header.paragraphs[0]
    for tab in doc.styles['Header'].paragraph_format.tab_stops:
        p.paragraph_format.tab_stops.add_tab_stop(tab.position, WD_TAB_ALIGNMENT.CLEAR)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(ANCHO), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run('OCE – Nota técnica: participación ciudadana\t').font.size = Pt(9.5)
    for tipo, txt in (('begin', None), (None, ' PAGE '), ('separate', None), ('t', '1'), ('end', None)):
        r = OxmlElement('w:r')
        if tipo in ('begin', 'separate', 'end'):
            f = OxmlElement('w:fldChar')
            f.set(qn('w:fldCharType'), tipo)
            r.append(f)
        elif tipo is None:
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = txt
            r.append(it)
        else:
            tt = OxmlElement('w:t')
            tt.text = txt
            r.append(tt)
        rpr = OxmlElement('w:rPr')
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '19')
        rpr.append(sz)
        r.insert(0, rpr)
        p._p.append(r)
    borde(p._p.get_or_add_pPr(), 'w:pBdr', {'bottom': ('single', 4)})


# ------------------------------------------------------------------ contenido
def construir():
    doc = Document()
    estilos(doc)
    encabezado(doc)
    body = doc.element.body
    for p in list(body.iterchildren(qn('w:p'))):
        body.remove(p)

    doc.add_paragraph('Indicadores de participación ciudadana: avance, hallazgos y decisiones', style='Title')
    for et, tx in (('Nota técnica', 'Observatorio Ciudadano de Envigado (OCE)'),
                   ('Fecha', '9 de noviembre de 2026'),
                   ('Dirigido a', 'profesor Eber Gutiérrez')):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(2.8))
        escribir(p, f'**{et}:**\t{tx}')
    borde(p._p.get_or_add_pPr(), 'w:pBdr', {'bottom': ('single', 4)})
    p.paragraph_format.space_after = Pt(6)

    aviso = doc.add_table(rows=1, cols=1)
    aviso.alignment = WD_TABLE_ALIGNMENT.CENTER
    tp = aviso._tbl.tblPr
    borde(tp, 'w:tblBorders', {x: ('single', 8) for x in ('top', 'left', 'bottom', 'right')}, '990000')
    c = aviso.cell(0, 0)
    sombrear(c._tc.get_or_add_tcPr(), 'F8D4D4')
    pa = c.paragraphs[0]
    pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pa.paragraph_format.space_before = pa.paragraph_format.space_after = Pt(2)
    escribir(pa, '**PROVISIONAL, POR VALIDAR.** Todas las variables están pendientes de validación y '
                 'ningún resultado de esta nota está validado.', 10)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

    # ---- resumen
    parrafo(doc, f'De los siete elementos del sistema de participación ciudadana, tres tienen un resultado '
                 f'provisional (I2, I3 e I5). I1, I4 e I6 no son viables por ahora porque no existe un inventario '
                 f'oficial de procesos participativos, y por eso el índice PC.ISPC no puede calcularse. Esta nota '
                 f'resume el estado de cada indicador, los hallazgos que condicionan su lectura, las decisiones que '
                 f'necesitamos de usted y los datos que conviene pedir a la Alcaldía.')

    # ---- 1
    doc.add_paragraph('1. Estado de los indicadores del sistema', style='Heading 1')
    parrafo(doc, 'La viabilidad sigue la clasificación del informe de disponibilidad y brecha. Los resultados aplican '
                 'la fórmula y los umbrales de la Matriz de información de indicadores con los supuestos por defecto '
                 'de la sección 4. N/D significa no disponible. El índice se define como '
                 'PC.ISPC = 0,30 × I1 + 0,25 × I2 + 0,20 × I3 + 0,25 × I4.')
    def res(r):
        return (f"{punt(r)} ({r['semaforo']})" if r['puntaje'] != '' else punt(r),
                r['semaforo'] if r['puntaje'] != '' else 'Gris')
    def dos(r24, r25, ind=None):
        a, b = res(r24), res(r25)
        return a, b
    def celda_res(r24, r25):
        a, b = res(r24), res(r25)
        # el color corresponde al resultado de 2025 (el más reciente con cálculo)
        return (f'2024: {a[0]}\n2025: {b[0]}', b[1] if b[1] != 'Gris' else a[1])
    nv, vc = ('No viable por ahora', 'Naranja'), ('Viable con ajuste', 'Amarillo')
    f1 = [
        ['**I1** Incorporación de aportes ciudadanos', nv, ('2024: N/D\n2025: N/D', 'Gris'),
         f'Universo exploratorio de procesos participativos: {V_PROC[0]} en 2024 y {V_PROC[1]} en 2025 (propuesta '
         'del equipo, no oficial). No se contó cuántos incorporaron un aporte en una decisión formal.'],
        ['**I2** Cobertura territorial', vc, celda_res(I2[2024], I2[2025]),
         f"Zonas con actividad sobre zonas oficiales: {'/'.join(entradas(I2[2024]))} en 2024 y "
         f"{'/'.join(entradas(I2[2025]))} en 2025. El denominador (13 zonas) se toma del [PDM, p. 558](PDM)."],
        ['**I3** Oferta formativa', vc, celda_res(I3[2024], I3[2025]),
         f"2025: {I3[2025]['entrada_1']} acciones formativas ejecutadas de {I3[2025]['entrada_2']} programadas. "
         f"2024: {I3[2024]['entrada_1']} ejecutadas, sin programación publicada. En personas: "
         f"{I3_PERS[2024]['entrada_1']} frente a {I3_PERS[2024]['entrada_2']} (2024) y "
         f"{I3_PERS[2025]['entrada_1']} frente a {I3_PERS[2025]['entrada_2']} (2025)."],
        ['**I4** Publicación de resultados', nv, ('2024: N/D\n2025: N/D', 'Gris'),
         'Mismo universo de I1. Se identificó un documento de resultados para 3 de 3 procesos (2024) y 4 de 5 '
         '(2025), sin verificar su contenido.'],
        ['**I5** Sesiones de instancias (complementario)', vc, celda_res(I5[2024], I5[2025]),
         f"2025: {I5[2025]['entrada_1']} sesiones realizadas de {I5[2025]['entrada_2']} previstas, solo en COPACO y en "
         "el Consejo Municipal de Seguridad Social en Salud (CMSSS), 4 y 4. 2024: sin programación ni resultado."],
        ['**I6** Participación virtual (complementario)', nv, ('2024: N/D\n2025: N/D', 'Gris'),
         f'Cota inferior de {V_I6[0]} actividad con canal virtual para intervenir por año: rendición del alcalde '
         '(2024) y votación web del presupuesto participativo (2025). No hay registro único de actividades.'],
        ['**PC.ISPC** Índice', nv, ('No calculable', 'Gris'),
         'Faltan I1 e I4 y la Guía no permite redistribuir pesos. I2 e I3 pueden presentarse como indicadores sueltos.'],
    ]
    tabla(doc, [3.0, 2.0, 2.4, 7.8], ['Indicador', 'Viabilidad', 'Resultado provisional', 'Variables obtenidas'], f1)

    # ---- 2
    doc.add_paragraph('2. Indicadores relacionados de otros sistemas', style='Heading 1')
    parrafo(doc, f"**PDM.CP, implementación de procesos participativos.** El Plan de Desarrollo fija una meta de "
                 f"100 con línea base 75 (PDF, p. 742). No hay valor observado ni método publicado: solo un índice "
                 f"situacional distinto, {V_INDSIT.replace('.', ',')} en 2022 ([PDM](PDM), p. 561), sin ficha que lo "
                 f"relacione con el indicador. Por eso queda N/D en 2024 y 2025. Debe solicitarse la ficha al Departamento "
                 f"Administrativo de Planeación (DAP); de no existir, podría redefinirse con I2 o I3 como insumo.")
    parrafo(doc, f"**Presupuesto participativo (PP).** Es en sí un proceso participativo y forma parte del universo "
                 f"exploratorio de I1 e I4: los comunicados reportan 88 y 90 iniciativas postuladas en los ciclos de "
                 f"2024 y 2025, con votación web del 11 al 24 de agosto de 2025. Su ejecución presupuestal (EPP) fue "
                 f"**{num(EPP[2024]['puntaje'])}** en 2024 (aproximado, {EPP[2024]['semaforo']}) y "
                 f"**{num(EPP[2025]['puntaje'])}** en 2025 ({EPP[2025]['semaforo']}); el primer semestre de 2026 da "
                 f"{num(EPP26['puntaje'])}, sin semáforo porque compara medio año con una apropiación anual. Por zonas, "
                 f"solo la zona 10 en 2024 baja a {num(EPP_Z10['puntaje'])} ({EPP_Z10['semaforo']}), por un cruce "
                 f"aproximado de un proyecto; la zona 1 es no evaluable porque no tiene proyectos de PP. Los proyectos "
                 f"con entrega exigible son {V_PE[0]}, {V_PE[1]} y {V_PE[2]} (2024 a 2026); la entrega cumplida de 2025 "
                 f"(49 por defecto) es un autorreporte sin acta de recibo, y la cifra de 55 citada en un documento "
                 f"previo no se reproduce con ninguna regla de conteo.")

    # ---- 3
    doc.add_paragraph('3. Hallazgos específicos', style='Heading 1')
    H = [
        '**No existe un inventario oficial de procesos participativos.** Tampoco hay un registro que vincule cada '
        'aporte con una decisión formal, que es el denominador y el numerador de I1 e I4. El plan indicativo reporta '
        'un ejercicio de PP por año (producto 530), cifra que no equivale al universo. El universo exploratorio '
        '(2024: formulación participativa del PDM, postulación de iniciativas de PP y rendición del alcalde; 2025: ciclo '
        'de PP, asambleas de revisión del POT, planes zonales, rendición de PP y rendición del alcalde) corre riesgo de '
        'subconteo. Los informes del Observatorio de Participación (producto 536) se reportan, pero la sección '
        'correspondiente del sitio de la EFPC está vacía.',
        '**Cifras que no cuadran: 65 frente a 170 personas capacitadas.** Ambas corresponden al producto 520, personas '
        'que participan en actividades de capacitación para la participación. En el [plan indicativo](PI) (segundo '
        'trimestre de 2026) el avance ejecutado es 105 en 2024, 65 en 2025 y 22 en 2026, con acumulados de 105, 170 y '
        '192; la población beneficiada de 2025 figura como 65. El [informe de gestión 2025](G25) (PDF, p. 87) dice '
        '«capacitación por demanda: 170 personas capacitadas frente a una meta de 30», y 30 es la meta anual de 2025 '
        'del mismo producto. Por tanto, 65 es el valor del año y 170 coincide con el acumulado 2024-2025 (105 + 65) '
        'comparado con una meta anual. Es una lectura aritmética: el informe no lo declara y debe confirmarla la '
        'Subdirección. El DAP además comentó el reporte del producto (ejecución anual que no concuerda con su sistema '
        'de seguimiento y evidencias «Sí, pero con observaciones»). La cifra de 186 personas formadas (producto 521) '
        'sí coincide en ambas fuentes.',
        '**Trece zonas, una sin comité y una asignación inferida.** El [PDM](PDM) (p. 558) establece 13 zonas con 12 '
        'comités zonales; la zona 1 queda pendiente hasta superar 1.000 habitantes y no tiene asamblea ni proyectos de '
        'PP. El informe de gestión 2025 (p. 92) dice que se hicieron 12 asambleas en agosto «cubriendo la totalidad '
        'de las zonas», sin listarlas; asignamos esas asambleas a las zonas 2 a 13. Si «totalidad» significa 13 zonas, '
        'I2 2025 sería 13/13. Las asambleas de revisión del [POT](POT) (2025) sí incluyeron la zona 1, y el producto '
        '531 del plan indicativo habla de «12 zonas». No hay actas por zona ni se obtuvo el texto del Acuerdo 007 de '
        '2025 que oficializaría la zonificación.',
        '**Participación virtual: solo dos actividades documentadas.** La rendición de cuentas del alcalde de 2024 '
        'incluyó preguntas de participantes virtuales ([cuadro de preguntas](QA24), pp. 6-8) y el PP de 2025 usó una '
        'plataforma web para votar ([informe de gestión](G25), p. 92; más de 12.000 ciudadanos registrados). Las '
        'transmisiones por YouTube y Facebook (776 vistas en la rendición del 12 de febrero de 2025; 11 programas '
        '«Conversando con los Comités Zonales» en 2024) solo permiten observar y no se cuentan. Los cursos virtuales '
        'de la EFPC son formación, no actividad participativa, salvo decisión contraria.',
        '**Instancias con calendario solo en salud.** COPACO y el Consejo Municipal de Seguridad Social en Salud (CMSSS) '
        'tienen programación y seguimiento en el [plan de acción en salud 2025](PPSS): 4 de 4 sesiones cada una, 8 de 8 '
        'en total. Es un reporte agregado; las actas y evidencias están en carpetas de SharePoint sin acceso, por lo que '
        'no se pudo verificar 2024. El PDF publicado sobre COPACO es una ficha descriptiva, no un acta. El Consejo '
        'Municipal de Participación, el CTP y los comités zonales solo reportan asistencia técnica «por demanda» '
        '(100 % cada año), que no equivale a sesiones previstas y realizadas.',
        f"**I3: denominador reconstruido y diplomado excluido.** No hay programación fechada de acciones; la meta oficial "
        f"está en personas. Con el [informe de gestión 2025](G25) (pp. 87-88) se reconstruyó el denominador: 21 acciones "
        f"culminadas (3 diplomados, 9 seminarios y 9 cursos) más 4 no ejecutadas (1 diplomado y 3 seminarios), es decir "
        f"25. Se excluyó el diplomado en Gestión Pública por presumirlo dirigido a servidores, algo que la fuente no "
        f"acredita; así se obtiene {I3[2025]['entrada_1']}/{I3[2025]['entrada_2']}. Los semilleros (5 de 7) se excluyen. "
        f"Para 2026 la Escuela figura «en planeación» y su oferta web lista cuatro ofertas sin fechas.",
        f"**Medida en personas, I3 da «Revisar».** La meta programada del producto 521 es {I3_PERS[2025]['entrada_2']} "
        f"personas en 2025 ({I3_PERS[2024]['entrada_2']} en 2024) frente a {I3_PERS[2025]['entrada_1']} "
        f"({I3_PERS[2024]['entrada_1']}) formadas. La Matriz exige «Revisar» cuando el numerador supera al denominador, "
        f"y el informe califica el resultado de sobrecumplimiento significativo, lo que sugiere una meta subestimada.",
        '**Otros hallazgos.** (a) Las fechas de dos eventos difieren entre documentos: la audiencia del alcalde de 2024 '
        '(20 o 26 de noviembre) y la rendición de PP de 2024 (12 de enero o 12 de febrero de 2025). (b) Los cronogramas '
        '([POT](POT), [calendario de participación 2024](CAL)) son programación, no ejecución, y los planes zonales no '
        'prueban actividad. (c) El portal institucional presentó un error de certificado, por lo que los documentos se '
        'consultaron en el espejo de la Alcaldía. (d) La programación de asambleas de PP de 2026 ([documento](PP26)) '
        'lista explícitamente las zonas 2 a 12; la zona 13 no aparece y no se pudo confirmar si es omisión.',
    ]
    lista(doc, H, numerada=True)

    # ---- 4
    doc.add_paragraph('4. Decisiones que necesitamos de usted', style='Heading 1')
    parrafo(doc, 'Los rangos son del nivel municipal; el valor por defecto está en negrita y se adoptó solo de forma '
                 'provisional. «Revisar» y N/D no tienen color de semáforo.')
    D = [
        ['Universo de procesos participativos (I1, I4, I6, ISPC)',
         '**Exploratorio** (3 en 2024; 5 en 2025) / inventario oficial',
         'En ninguna opción hay cálculo hoy: el universo exploratorio solo produce variables y el inventario oficial '
         'deja N/D hasta recibirlo.'],
        ['Actividades que cubren una zona (I2)', '**Asambleas de PP, PDM y POT** / solo PP',
         f"2024: {punt(I2[2024])} a {punt(I2_PP)}; 2025 no cambia ({punt(I2[2025])})."],
        ['Evidencia de la actividad (I2)', '**Realización reportada** / incluye programación / acta por zona',
         f"2025: {punt(I2[2025])} a {punt(I2_PROG)} con programación; con acta por zona, {punt(I2_ACTA)} en ambos años."],
        ['Zonas oficiales (I2)', '**13** / 12 con comité zonal',
         f"2025: {punt(I2[2025])} a {punt(I2_12)}. No cambia el PP."],
        ['Unidad de I3', '**Acciones formativas** / personas',
         f"2025: {punt(I3[2025])} a {punt(I3_PERS[2025])} ({'/'.join(entradas(I3_PERS[2025]))}); 2024: N/D a "
         f"{punt(I3_PERS[2024])} ({'/'.join(entradas(I3_PERS[2024]))})."],
        ['Formación para servidores (I3)', '**Excluir** / incluir',
         f"2025: {punt(I3[2025])} a {punt(I3_SERV)}."],
        ['Semilleros como acción formativa (I3)', '**No** / sí',
         f"2025: {punt(I3[2025])} a {f'{SEMILLEROS:.1f}'.replace('.', ',')} (sensibilidad fuera de las opciones configuradas)."],
        ['Instancias de I5', '**COPACO y CMSSS** / ampliado (Consejo Municipal de Participación, CTP, comités zonales)',
         f"2025: {punt(I5[2025])} a {punt(I5_AMP)}."],
        ['Tratamiento de 2026', '**Primer semestre provisional** / excluir / anual provisional',
         f"PP 2026: {num(EPP26['puntaje'])} sin semáforo / sin cálculo / rotulado anual (contradice la Guía; solo "
         "sensibilidad). Afecta también a I2 e I3."],
        ['Regla de estado para calcular', '**Preliminar** (acepta datos pendientes) / estricta (solo validados)',
         'Con la regla estricta de la Matriz todo es N/D mientras no haya datos validados.'],
        ['Presupuesto participativo (relacionado)',
         '**Cruces aproximados en 2024** / solo exactos; proyecto 0553 (costo del proceso) fuera / dentro',
         f"EPP 2024: {num(EPP[2024]['puntaje'])} a {num(EPP_EX['puntaje'])} o N/D; EPP 2025: "
         f"{num(EPP[2025]['puntaje'])} a {num(EPP_0553['puntaje'])} con el proyecto 0553."],
    ]
    tabla(doc, [3.6, 4.9, 6.7], ['Decisión', 'Opciones', 'Rango de resultados'], D)

    # ---- 5
    doc.add_paragraph('5. Qué pedir a la Alcaldía', style='Heading 1')
    doc.add_paragraph('Subdirección de Participación Ciudadana (DAP)', style='Heading 2')
    lista(doc, [
        'Inventario de procesos participativos cerrados de 2024, 2025 y 2026, con fecha de cierre, producto de cierre, '
        'enlace público al resultado publicado, acto (acta, decreto o acuerdo) donde se incorporó el aporte y el '
        'criterio de universo usado.',
        'Actas o listados de asistencia, solo con conteos, de las asambleas zonales del PDM, el POT y el PP, con fecha y '
        'zona, y el listado de las 12 asambleas de PP de 2025 por zona.',
        'Registro de actividades participativas con fecha, zona, modalidad (presencial, virtual o híbrida), canal '
        'habilitado para intervenir y evidencia; datos de participación virtual de las rendiciones de 2024 y 2025.',
        'Calendario aprobado de sesiones y actas, o certificación de sesiones con quórum, de las instancias que usted '
        'seleccione (Consejo Municipal de Participación, CTP, comités zonales) y soportes de COPACO y del CMSSS de 2024.',
        'Aclaración del producto 520 (65 frente a 170), ficha técnica del indicador de implementación de procesos '
        'participativos del PDM, informes del Observatorio de Participación (producto 536) y acto vigente de '
        'zonificación (Acuerdo 007 de 2025), con la situación de la zona 1.',
    ])
    doc.add_paragraph('Escuela de Formación Política y Ciudadana (EFPC)', style='Heading 2')
    lista(doc, [
        'Plan de formación aprobado y relación de acciones formativas (curso, diplomado, seminario, cátedra) con fechas '
        'de inicio y fin, público (ciudadanía o servidores), cupos y certificados, para 2024, 2025 y 2026.',
        'Asistentes certificados por acción, sin datos personales, para conciliar la medición en acciones con la '
        'medición en personas, y la meta anual de los productos 520 y 521.',
        'Relación anual de actividades con modalidad y evidencia del canal virtual, empezando por 2026.',
    ])
    parrafo(doc, 'Con la Oficina de Comunicaciones convendría pedir los informes públicos de rendición de cuentas de '
                 '2024 y 2025 y la evidencia de los canales virtuales de preguntas.', space_before=Pt(4))
    return doc


def main():
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    construir().save(SALIDA)
    print('Generado', SALIDA.relative_to(RAIZ))


if __name__ == '__main__':
    main()
