"""Genera los dos Word de la revisión de indicadores (nov-2026).

1. outputs/entrega_profe/indicadores_cambios_y_datos_OCE.docx: para el profesor. Corto, con cuadros sencillos.
2. outputs/entrega_profe/detalle_revision_indicadores_OCE.docx: para Isabella. Qué pasó con cada indicador, en detalle.

Uso (desde la raíz, sin -I):  python3 scripts/entregables/generar_entrega_profe.py
Contenido: los resultados salen de outputs/construccion/{PDM,PC,PP,CS}/*_indicadores.csv (no se copian cifras a mano);
los textos, de las decisiones (.claude/memoria/decisiones.md). Al final se verifica el contenido de los Word.
El archivo outputs/entrega_profe/dudas_para_el_profe.md se verifica (todas respondidas), pero ya no se vuelca al Word. Todo es PROVISIONAL.
"""
import csv
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
FECHA = '19 de noviembre de 2026'


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




# ------------------------------------------------------------------ datos de la construcción (CSV)
CONS = RAIZ / 'outputs' / 'construccion'
DIMS = {'PDM': 'pdm', 'PC': 'pc', 'PP': 'pp', 'CS': 'cs'}
ANIO_MAX = 2025


def cargar(dim):
    ruta = CONS / dim / f'{DIMS[dim]}_indicadores.csv'
    with open(ruta, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f, delimiter=';'))


DATA = {d: cargar(d) for d in DIMS}


def num(s):
    try:
        return float(str(s).strip().replace(',', '.'))
    except ValueError:
        return None


def fmt(s, dec=2):
    """Formatea un valor del CSV con coma decimal; los enteros llevan punto de miles."""
    s = str(s).strip()
    if re.fullmatch(r'-?\d+', s):
        return f'{int(s):,}'.replace(',', '.')
    x = num(s)
    if x is None:
        return s
    return f'{x:.{dec}f}'.replace('.', ',')


def filas(dim, codigo, tipo=None):
    return [r for r in DATA[dim] if r['codigo'] == codigo and (tipo is None or r['tipo'] == tipo)]


def anual(dim, codigo, tipo):
    return [r for r in filas(dim, codigo, tipo) if r['anio'].isdigit() and int(r['anio']) <= ANIO_MAX]


def con_dato(rs):
    return [r for r in rs if num(r['valor']) is not None]


def comprimir(anios):
    anios = sorted(set(anios))
    trozos, i = [], 0
    while i < len(anios):
        j = i
        while j + 1 < len(anios) and anios[j + 1] == anios[j] + 1:
            j += 1
        trozos.append(str(anios[i]) if i == j else f'{anios[i]}–{anios[j]}')
        i = j + 1
    return ', '.join(trozos)


def anios_dato(dim, codigo):
    return comprimir([int(r['anio']) for r in con_dato(anual(dim, codigo, 'principal'))]) or 'ninguno'


def maximo(dim, codigo, anio):
    """Valor máximo del rango (fila CODIGO.MAX) si existe; solo PC y CS reportan rangos."""
    if dim not in ('PC', 'CS'):
        return None
    for r in filas(dim, codigo + '.MAX'):
        if r['anio'] == anio and num(r['valor']) is not None:
            return r['valor']
    return None


def quitar_parentesis(s):
    while True:
        n = re.sub(r'\([^()]*\)', '', s)
        if n == s:
            return s
        s = n


COMP_PDM = {'Salud': 'Salud', 'ILE': 'Logro educativo', 'Servicios publicos': 'Servicios públicos',
            'Mortalidad vial': 'Mortalidad vial', 'Espacio verde': 'Espacio verde',
            'Veedurias': 'Veedurías', 'Homicidio': 'Homicidio'}


def componentes_faltantes(dim, r):
    """Lista legible de los componentes que faltan en un año del índice."""
    b = quitar_parentesis(r['brecha'])
    if dim == 'PDM':
        resto = b.split('faltan:', 1)[1] if 'faltan:' in b else ''
        return [COMP_PDM.get(x.strip(), x.strip()) for x in resto.split(';') if x.strip()]
    if dim == 'PC':
        resto = b.split('faltan', 1)[1] if 'faltan' in b else ''
        return [x.strip() for x in resto.split(',') if x.strip()]
    if dim == 'PP':
        return ['EPP', 'EJ90', 'FID'] if 'tres componentes' in b else []
    if dim == 'CS':
        out = []
        for x in re.findall(r'falta (\w+)', b):
            if x == 'DSI':
                out.append('DSI (no aplica antes de 2024)')
            elif x == 'PQRSDF':
                out.append('PQRSDF')
            else:
                out.append(x)
        return out
    return []


INDICES_DEF = [
    # nombre, dimensión, código, año inicial, año final, composición
    ('ICPD-7 (Cumplimiento del PDM)', 'PDM', 'PDM.ICPD7', 2019, 2025,
     'Promedio simple de 7 componentes: salud, logro educativo, servicios públicos, mortalidad vial, espacio verde, veedurías y homicidio.'),
    ('ISPC (Participación ciudadana)', 'PC', 'PC.ISPC', 2020, 2025,
     '0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4.'),
    ('IREPP (Presupuesto participativo)', 'PP', 'PP.IREPP', 2020, 2025,
     '0,20·EPP + 0,40·EJ90 + 0,40·FID.'),
    ('ICRCS (Control social)', 'CS', 'CS.ICRCS', 2020, 2025,
     '(DSI + PQRSDF + VRC + RPO) / 4.'),
]
SEMAFORO = 'Semáforo (propuesta): Verde ≥ 90 · Amarillo 75 a < 90 · Naranja 50 a < 75 · Rojo < 50.'


def datos_indice(dim, codigo, a0, a1):
    """Filas [Año, Valor, Semáforo, Componentes que faltan] y (años con índice, años del periodo)."""
    out, k = [], 0
    rs = [r for r in anual(dim, codigo, 'indice') if a0 <= int(r['anio']) <= a1]
    assert len(rs) == a1 - a0 + 1, (codigo, len(rs))
    for r in rs:
        v = num(r['valor'])
        if v is not None:
            k += 1
            out.append([r['anio'], f'**{fmt(r["valor"])}**', r['semaforo'] or '—', 'Ninguno'])
            continue
        valor = 'N/D'
        m = re.search(r'Rango posible segun PQRSDF.*?: (\d+\.\d+) \(P = 0\) a (\d+\.\d+)', r['nota'])
        if m:
            valor = f'N/D (entre {fmt(m.group(1))} y {fmt(m.group(2))} según PQRSDF)'
        faltan = componentes_faltantes(dim, r)
        assert faltan, (codigo, r['anio'], r['brecha'])
        out.append([r['anio'], valor, '—', ', '.join(faltan)])
    return out, k, len(rs)


def linea_brecha(k, n):
    pct = str(round((n - k) / n * 100, 1)).replace('.', ',')
    return f'**Brecha:** {n - k} de {n} años sin índice ({pct} %).'


# ------------------------------------------------------------------ cuadros de indicadores por dimensión
NOMBRES = {
    'PDM': [('PDM.SALUD', 'Cobertura de salud (%)'), ('PDM.ILE', 'Logro educativo, ILE (0–100)'),
            ('PDM.SP', 'Servicios públicos (%)'), ('PDM.CM', 'Mortalidad vial (por 100.000)'),
            ('PDM.CA', 'Espacio verde urbano (m² por habitante)'), ('PDM.VEED', 'Veedurías inscritas (por año)'),
            ('PDM.HOM', 'Homicidio (por 100.000)')],
    'PC': [('PC.I1', 'I1 · Votación del presupuesto participativo (IPCM, %)'), ('PC.I2', 'I2 · Cobertura territorial (%)'),
           ('PC.I3', 'I3 · Oferta formativa (%)'), ('PC.I4', 'I4 · Publicación de resultados (%)')],
    'PP': [('PP.EPP', 'EPP · Ejecución presupuestal (%)'), ('PP.EJ90', 'EJ90 · Proyectos con ejecución ≥ 90 % (%)'),
           ('PP.FID', 'FID · Fidelidad a lo votado (%)')],
    'CS': [('CS.PQRSDF', 'PQRSDF · Oportunidad de respuesta (%)'), ('CS.DSI', 'DSI · Información de seguimiento (%)'),
           ('CS.VRC', 'VRC · Veedurías en rendición de cuentas (%)'), ('CS.RPO', 'RPO · Respuesta a observaciones (%)')],
}


def valor_fila(r):
    """Valor a mostrar. I1 (IPCM) se recalcula con numerador / denominador para no redondear dos veces."""
    if r['codigo'] == 'PC.I1' and num(r['numerador']) is not None and num(r['denominador']):
        exacto = 100 * num(r['numerador']) / num(r['denominador'])
        assert abs(exacto - num(r['valor'])) < 0.001, (r['anio'], exacto, r['valor'])
        return repr(exacto)
    return r['valor']


def ultimo_valor(dim, cod):
    con = con_dato(anual(dim, cod, 'principal'))
    if not con:
        return 'Sin dato'
    r = con[-1]
    v, p, a = valor_fila(r), r['puntaje_0_100'], r['anio']
    mx = maximo(dim, cod, a)
    if mx is not None and num(mx) != num(v):
        return f'{fmt(v)}–{fmt(mx)} (mín.–máx.; {a})'
    if num(p) is not None and abs(num(p) - num(v)) > 0.005:
        return f'{fmt(v)} (puntaje {fmt(p)}; {a})'
    return f'{fmt(v)} ({a})'


def texto_brecha(dim, cod, corto=True):
    """Formato homogéneo: «N de M <unidad> sin dato (X %)»; el detalle va después solo si corto=False."""
    rs = filas(dim, cod + '.BRECHA', 'brecha')
    if not rs:
        return '—'
    r = rs[0]
    t = r['brecha'].replace('anios', 'años')
    pct = f'{fmt(r["valor"], 1)} %'
    pares = re.findall(r'(\d+)/(\d+)', t)
    if cod == 'CS.DSI' and pares:
        n, m = sum(int(a) for a, _ in pares), sum(int(b) for _, b in pares)
        assert abs(100 * n / m - num(r['valor'])) < 0.01, (n, m, r['valor'])
        base, resto = f'{n} de {m} metas (año × meta) no verificables ({pct})', t
    elif re.match(r'\d+ de \d+ ', t):
        cabeza, _, resto = t.partition(';')
        if '%' in cabeza:  # ya trae el porcentaje (veedurías)
            base = cabeza.strip()
        else:
            extra = re.search(r'\s*\(([^()]*)\)$', cabeza)
            base = f'{re.sub(r"\s*\([^()]*\)$", "", cabeza).strip()} ({pct})'
            resto = ((extra.group(1) + '; ') if extra else '') + resto
    else:
        base, resto = t, ''
    return base if corto or not resto.strip() else f'{base}; {resto.strip()}'



# ------------------------------------------------------------------ textos que dependen de los CSV (PP y PC)
def txt_csv(dim, cod):
    return ' '.join(r['nota'] + ' ' + r['brecha'] + ' ' + r['referente'] for r in DATA[dim] if r['codigo'] == cod)


def pp_anios_sin_dato():
    """[(año, motivo corto)] de los años 2020–2025 sin IREPP."""
    out = []
    for r in anual('PP', 'PP.IREPP', 'indice'):
        if num(r['valor']) is None and 2020 <= int(r['anio']) <= 2025:
            m = quitar_parentesis(re.sub(r'^sin dato:\s*', '', r['brecha'])).split(';')[0].strip().replace('anios', 'años')
            out.append((int(r['anio']), m[:200]))
    return out


def pp_sin_dato_txt():
    sd = pp_anios_sin_dato()
    return ('; '.join(f'{a} queda sin dato' for a, _m in sd) + '.') if sd else 'Todos los años de 2020 a 2025 tienen índice.'


def pp_universo_2021():
    """Universo de PP 2021 (lista oficial sin los proyectos del INDER) y sensibilidad con todos, según el CSV."""
    t = txt_csv('PP', 'PP.EPP')
    r21 = [r for r in DATA['PP'] if r['codigo'] == 'PP.EPP' and r['anio'] == '2021']
    if 'lista oficial' not in ' '.join(r['nota'] for r in r21):
        return ''
    n21 = r21[0]['nota']
    tot = re.search(r'(\d+) proyectos, \$', n21)
    inder = re.search(r'sin los (\d+) proyectos del INDER', n21)
    sens = [r for r in DATA['PP'] if r['codigo'] == 'PP.SENS2021.IREPP' and r['anio'] == '2021']
    ej = [r for r in DATA['PP'] if r['codigo'] == 'PP.EJ90' and r['anio'] == '2021'][0]
    txt = f'2021: el universo es la lista oficial de proyectos de la vigencia (publicada el 23-feb-2022), sin los {inder.group(1) if inder else "…"} proyectos del INDER: {ej["denominador"]} proyectos'
    if sens and tot:
        txt += f'; con los {tot.group(1)} proyectos el IREPP sería {fmt(sens[0]["valor"])}'
    return txt + '.'


def pp_rubros24():
    m = re.search(r'incluye los rubros 2\.4 \(Fondo Local de Salud\): (\d+) proyectos', txt_csv('PP', 'PP.EPP'))
    return f'El cruce de 2020 incluye los rubros 2.4 (Fondo Local de Salud): {m.group(1)} proyectos de Salud cruzan al peso.' if m else ''


def aviso_pp2020():
    c = [r for r in DATA['PP'] if r['codigo'] == 'PP.CONTRACREDITO' and r['anio'] == '2020']
    if not c:
        return ''
    monto = re.search(r'\$([\d\.]+)', c[0]['nota'])
    e = [r for r in DATA['PP'] if r['codigo'] == 'PP.EPP' and r['anio'] == '2020'][0]
    m = re.search(r'valor votado de los verificados = ([\d\.]+) %', e['nota'])
    t = (f'**PP 2020:** {c[0]["numerador"]} de {c[0]["denominador"]} proyectos con contracrédito total' + (f' (${monto.group(1)})' if monto else '') +
         f'; la ejecución presupuestal ({fmt(e["valor"])}) no lo refleja' + (f'; obligación / valor votado = {fmt(m.group(1))} %' if m else '') +
         '; la vigencia 2020 se priorizó en 2019.')
    return ' '.join(x for x in (t, pp_rubros24(), pp_universo_2021()) if x)


def pc_votos_txt():
    rs = [r for r in anual('PC', 'PC.I1', 'principal') if num(r['valor']) is not None]
    return '; '.join(f'{r["anio"]}: {fmt(r["numerador"])} votos' for r in rs)


def pc_anios_nd(cod):
    """Años 2017–2025 con «sin dato» (no «no aplica») del indicador de PC."""
    return comprimir([int(r['anio']) for r in anual('PC', cod, 'principal') if r['valor'] == 'N/D' and r['brecha'].startswith('sin dato')])


def tabla_indicadores_dim(dim):
    return [[nombre, anios_dato(dim, cod), ultimo_valor(dim, cod), texto_brecha(dim, cod)] for cod, nombre in NOMBRES[dim]]


AVISO_PP2020 = aviso_pp2020()
assert 'PP 2020' in AVISO_PP2020 and '$' in AVISO_PP2020, 'las cifras de PP 2020 cambiaron en el CSV: revisar aviso_pp2020()'


ENC_IND = ['Indicador', 'Años con dato', 'Último valor (puntaje)', 'Brecha de información']
W_IND = [7.0, 4.4, 5.6, 6.8]


# ------------------------------------------------------------------ cuadros de cambios (los mismos para los dos Word)
ENC = ['Indicador original', 'Situación', 'Decisión', 'Origen', 'Datos']


def serie(dim, cod, extra=''):
    return f'{anios_dato(dim, cod)}{extra}'


PDM = [
    ['**Aseguramiento al régimen subsidiado**', 'La meta del PDM (99,5 %) es solo del régimen subsidiado y ese dato no era fácil de conseguir',
     '**Se reemplaza** por la cobertura de salud total (afiliados / población DANE). Referente: **100 %** (cobertura universal)',
     'Profe (reuniones 1 y 2; referente 100 %)', serie('PDM', 'PDM.SALUD')],
    ['**Mortalidad vial**', 'Mezclar fuentes cambia la medida; la meta de 8 es de Movilidad, otro universo',
     '**Se mantiene** con una fuente (Policía Nacional). Referente del mismo universo: **7,33** por 100.000 (la reducción del PDM aplicada a la línea base 2019 de la Policía)',
     'Profe (no mezclar universos); referente: Isabella', serie('PDM', 'PDM.CM')],
    ['**Espacio público verde**', 'Con la meta del PDM (3,88 m²) daba 100 sin exigencia',
     '**Se ajusta**: avance desde la **línea base 2019 = 4,54 m²/hab** hacia **9 m²/hab**. Cortes de 2019, 2022 y 2024; se pide a Planeación la serie anual',
     'Profe (hacia 9 m²); línea base 2019: Isabella', serie('PDM', 'PDM.CA', ' (cortes de estudios distintos)')],
    ['**Implementación de procesos participativos**', 'No hay cifra ni método oficial publicado',
     '**Se reemplaza** por las veedurías inscritas en la Personería, con línea base **5,4** al año (promedio 2019–2023)',
     'Profe e Isabella', serie('PDM', 'PDM.VEED', '; el listado parece incompleto')],
    ['**Cobertura deportiva**', 'No hay tasa oficial del Índer',
     '**Se reemplaza** por servicios públicos: promedio simple de acueducto, alcantarillado, aseo y energía. Referente: **100 %**',
     'Profe e Isabella', serie('PDM', 'PDM.SP', '; 2019 con alerta de comparabilidad')],
    ['**Índice de logro educativo**', 'La clasificación de planteles en Saber 11 no es pública',
     '**Se ajusta**: Saber 11 → cobertura neta; se conserva primera infancia (0–5 años). Promedio simple', 'Profe e Isabella',
     serie('PDM', 'PDM.ILE', '; 2015–2018 con alerta')],
    ['**Sexta medida**', 'Quedó abierta el 12-nov',
     '**Se agrega homicidio como 7.º componente** (seguridad). Meta del PDM 2,7 por 100.000 (Policía, mismo universo). El índice pasa a ICPD-7: promedio simple, 1/7 cada uno',
     'Profe', serie('PDM', 'PDM.HOM')],
    ['**Logros de políticas públicas**', 'Dimensión para el observatorio futuro', 'Solo variables por solicitar', 'Profe', 'Sin datos'],
]

PC = [
    ['**I1** Incorporación de aportes', 'No hay inventario de procesos; el respaldo repetía los datos de la fidelidad de PP',
     '**Se reemplaza** por la votación del presupuesto participativo (IPCM = votos / población). Avance desde la **línea base 2017 = 3,68 %** hacia **7 %** (tope de la escala oficial)',
     'Profe (diferenciar de PP); Isabella (línea base y referente)', serie('PC', 'PC.I1', '; 2018 sin votación; sin cifra: ' + pc_anios_nd('PC.I1'))],
    ['**I2** Cobertura territorial', 'El dato principal existe para ' + anios_dato('PC', 'PC.I2'),
     '**Se mantiene**, con **12 zonas** (la zona 1 es industrial). Con alerta: es cobertura documental (fichas, programación y votación por zona), sin actas de asistencia; se solicita la lista de asambleas por zona', 'Profe', serie('PC', 'PC.I2')],
    ['**I3** Oferta formativa', 'Solo 2025 tiene denominador, y es reconstruido',
     '**Se mantiene**: 2025 = **21 de 25** acciones. Respaldo: número de acciones para la ciudadanía (2023–2025)', 'Isabella (D49, D72)', serie('PC', 'PC.I3')],
    ['**I4** Publicación de resultados', 'No hay inventario oficial de procesos cerrados',
     '**Se mantiene** con una **lista fija de 3 procesos** que el observatorio verifica en el portal. 2025 como rango (falta la rendición de PP)',
     'Profe e Isabella', serie('PC', 'PC.I4')],
    ['**I5** e **I6**', '—', '**Eliminados**', 'Profe (reunión 1)', '—'],
]

PP = [
    ['**EPP** Ejecución presupuestal', 'Hay dato de Hacienda por proyecto', '**Se mantiene**. Sin el rubro de déficit fiscal', 'Profe (mantener); Isabella (excluir déficit, D51)', serie('PP', 'PP.EPP', ' (' + pp_sin_dato_txt() + ')' if pp_anios_sin_dato() else '')],
    ['**CEP** Entrega de proyectos', 'No hay actas de recibo; «terminado» solo existe como autorreporte',
     '**Se reemplaza** por **EJ90**: proyectos con ejecución ≥ 90 % (Hacienda). «Terminados» queda solo como respaldo', 'Profe', serie('PP', 'PP.EJ90')],
    ['**CMA** Metas de atención', 'El plan de acción mide actividades, no personas', '**Se reemplaza** por **FID**: fidelidad a lo votado', 'Profe', serie('PP', 'PP.FID')],
    ['**IREPP**', 'Serie corta (solo 2024–2025)', 'Pesos 20/40/40. La serie se amplía a **2020–2025**. ' + pp_sin_dato_txt() + ' ' + pp_universo_2021(),
     'Profe (pesos); Isabella (serie)', '2020–2025'],
]

CS = [
    ['**PCS** Promoción del control social', 'Autorreporte: 1 de 1 estrategia = 100 %',
     '**Se reemplaza** por **PQRSDF**: oportunidad de respuesta (atendidas en términos / recibidas)', 'Profe; Isabella (no al autorreporte)',
     serie('CS', 'CS.PQRSDF', '. Alertas: plazos ampliados por el Decreto 491 solo hasta el 17-may-2022; en 2024 las recibidas suben 64 % por reclasificación (no es tendencia); sistema Débora desde mayo de 2025')],
    ['**DSI** Información de seguimiento', 'La lista de metas estaba desactualizada',
     '**Se mantiene** con **10 metas** de resultado del PDM; rango (mín.–máx.) con brecha', 'Profe (mantener); Isabella (lista de 10, D75/D79)', serie('CS', 'CS.DSI')],
    ['**VRC** Veedurías en rendición de cuentas', 'En varios eventos no se sabe si hubo veeduría',
     '**Se mantiene**, como **rango** (mín.–máx.) con la **brecha de información**; al índice entra el mínimo', 'Profe (mantener); Isabella (rango y brecha, D33; extensión, D77)', serie('CS', 'CS.VRC')],
    ['**RPO** Respuesta a observaciones', 'El cuadro de 2025 es corto',
     '**Se mantiene**; 2021 como rango 0–100 (mínimo 0 hasta confirmar la fecha de publicación); 2025 con alerta', 'Profe (mantener); Isabella (2021 y 2025, D76/D80)', serie('CS', 'CS.RPO')],
]

# Espacio verde: cortes y fuentes (los avances salen del CSV)
EV_ENC = ['Año', 'm² por habitante', 'Avance hacia 9 m²', 'Quién lo midió', 'Fuente (documento, tabla y página)']


def ev_cortes():
    av = {r['anio']: r['puntaje_0_100'] for r in con_dato(anual('PDM', 'PDM.CA', 'principal'))}
    val = {r['anio']: r['valor'] for r in con_dato(anual('PDM', 'PDM.CA', 'principal'))}
    assert (val['2019'], val['2022'], val['2024']) == ('4.54', '4.66', '5.75'), val
    return [
        ['**2019**', f'**{fmt(val["2019"])}**', '0 % (línea base)', 'Área Metropolitana (AMVA) con la Universidad Nacional (2020)',
         'Reporte Metropolitano Voluntario ODS 11.7, AMVA y ONU-Hábitat (2022), tabla 4, p. 24'],
        ['2022', fmt(val['2022']), f'{fmt(av["2022"], 1)} %', 'AMVA con el Instituto Humboldt',
         'Instituto Humboldt (2023); diagnóstico del POT, tabla 39, pp. 86-87'],
        ['2024', fmt(val['2024']), f'{fmt(av["2024"], 1)} %', 'Alcaldía de Envigado (diagnóstico del POT, criterio del AMVA)',
         'Diagnóstico del POT, tabla 39, pp. 86-87: 1.382.762 m² / 240.655 habitantes urbanos'],
    ]


EV_NOTA_CORTA = ('Son cortes de estudios distintos, no una serie anual. Se solicita a Planeación la serie completa desde 2019 con el método del AMVA, '
                 'para que la mida cada año en el seguimiento del POT.')
EV_NOTA = ('Avance = (valor del año − 4,54) / (9 − 4,54). Son **cortes de estudios distintos, no una serie anual**: el dato no se publica cada año '
           'y el de 2024 usa otro inventario, así que parte del avance puede deberse al cambio de método. '
           '**Recomendación para el informe:** solicitar a **Planeación** la serie completa desde 2019 con la metodología del AMVA, y proponer que '
           '**Planeación la mida cada año** en el seguimiento del POT (expediente municipal), con el mismo método, y que la **Secretaría de Medio Ambiente** '
           'la reporte en el seguimiento del PDM. Se descartan el 3,88 del PDM (sin fuente ni método), los 5,00 de 2020 y 2023 (autorreporte) y el '
           'verde total 2024 (69,41, incluye las reservas rurales). Con la población de la VP de 2024 el valor sería 5,54 (nota de sensibilidad).')
EV_URLS = [
    '2019 (4,54): Reporte ODS 11.7, AMVA y ONU-Hábitat (2022), tabla 4, p. 24, y tabla 5, p. 25: https://urbanpolicyplatform.org/wp-content/uploads/2023/06/VMR-EPV-VALLE-DE-ABURRA.pdf',
    '2022 (4,66): Instituto Humboldt (2023), https://www.humboldt.org.co/noticias/espacios-publicos-verdes-urbanos-en-el-valle-de-aburra-un-reto-de-conservacion',
    '2024 (5,75): diagnóstico del POT, dimensión funcional (espacio público), tabla 39, pp. 86-87: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002925/146211_8-dimension-funcional-_-espacio-publico.pdf',
    'Descartado 3,88 (LB 2023 del PDM 2024-2027): apéndice 9.1.2 del PDM, fila 129: https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002340/116989_apendice91planindicativo27052024.xlsx',
    'Búsqueda completa: outputs/revision_indicadores/PDM/espacio_verde_serie.md; verificación de QA: outputs/revision_indicadores/PDM/qa_espacio_verde.md.',
]

VRC_NECESITA = [
    'Lista fija de eventos del año (denominador): audiencia del Alcalde y rendiciones sectoriales, confirmada en enero con el cronograma de rendición de cuentas. La recoge el observatorio.',
    'Lista de asistencia de cada evento con la casilla «¿pertenece a una veeduría? ¿cuál?». La produce la dependencia que organiza el evento; la recoge el observatorio.',
    'Listado de veedurías inscritas, para verificar que quien firma como veedor pertenece a una veeduría inscrita (Personería, una vez al año).',
    'Registro propio del observatorio: una fila por evento (fecha, dependencia, documento, enlace y veeduría: sí, no o sin soporte).',
    'Observación directa: que el observatorio asista a la audiencia del Alcalde y registre las veedurías presentes.',
]

SOLICITAR = [
    f'Votos de presupuesto participativo de {pc_anios_nd("PC.I1").replace(", ", " y ")}, y desde 2018 por zona (Planeación).',
    'Población de referencia del IPCM y si se cuentan votos o votantes (Planeación).',
    'Lista de asambleas y actas de asistencia por zona, para respaldar el indicador de cobertura territorial (Subdirección de Participación Ciudadana).',
    'Programación anual de la Escuela de Formación (oferta formativa programada, 2020–2024).',
    'REC rural de 2016 a 2020 (cobertura de servicios públicos): hoy esos años tienen rural = 0 (Planeación).',
    'Espacio verde urbano año por año desde 2019 con el método del AMVA (Planeación, en el seguimiento del POT).',
    'Listado completo de veedurías inscritas (Personería): el listado público parece incompleto.',
    'Listas de asistencia de las rendiciones de cuentas, con la casilla de veeduría (Alcaldía), para cerrar la brecha de VRC.',
    'Rendición de cuentas del presupuesto participativo 2025 (se hará en el último trimestre de 2026).',
    'Logros de políticas públicas: políticas vigentes, informes de seguimiento y metas cumplidas.',
]

# Decisiones: (D, tema, decisión, origen, va en el Word del profe)
DECISIONES = [
    ('D34, D48', 'Índice del PDM', 'Siete componentes: salud, logro educativo, servicios públicos, mortalidad vial, espacio verde, veedurías y homicidio como séptimo (seguridad). Promedio simple: cada uno pesa 1/7.', 'Profe (7 componentes); Isabella (promedio simple)', True),
    ('D35', 'Referente de salud', '100 % (cobertura universal), porque la meta del PDM (99,5 %) es solo del régimen subsidiado.', 'Profe', True),
    ('D36, D46, D68', 'Referente de veedurías', 'Línea base = promedio de 2019 a 2023 = 5,4 veedurías al año. Puntaje = 100 × valor / 5,4, con tope 100.', 'Profe (con línea base, reunión 2); Isabella (promedio 2019–2023 = 5,4)', True),
    ('D31, D37, D59', 'Espacio verde', 'Línea base 2019 = 4,54 (AMVA) hacia 9 m²/hab; cortes 2019, 2022 y 2024 (5,75, valor publicado). Se pide a Planeación la serie anual con el método del AMVA.', 'Profe (hacia 9); Isabella (línea base 2019, urbano)', True),
    ('D38, D47, D83', 'Metas de otro universo', 'No se usan. Servicios públicos: referente 100 % (decisión). Mortalidad vial: 7,33 por 100.000 = la reducción del PDM (8 / 10,2) aplicada a la línea base 2019 de la Policía (9,35).', 'Profe (no mezclar universos); Isabella (7,33 y 100 % de servicios públicos)', True),
    ('D39, D50, D65', 'Servicios públicos', 'Promedio simple de las 4 coberturas desde 2019; referente 100 % (decisión). Acueducto, alcantarillado y aseo: 2009–2018 solo como contexto (2008 y 2010 vienen vacíos). 2019 se mantiene, con alerta (el REC rural de 2016–2020 no es comparable).', 'Profe; Isabella', True),
    ('D15, D16, D27, D57', 'Logro educativo', 'Promedio simple de cobertura neta y primera infancia (0–5 años). 2015–2018 con alerta de comparabilidad.', 'Isabella', True),
    ('D58, D63', 'Homicidio', 'Se mantiene la meta 2,7 del PDM (misma fuente y universo: Policía). Incluye el feminicidio. Nota: la línea base 3,7 del PDM usó otra población.', 'Profe (meta 2,7); Isabella (mantenerla con nota)', True),
    ('D60, D84', 'Semáforo (propuesta)', 'Propuesta por validar, igual en todas las dimensiones: Verde ≥ 90, Amarillo 75 a < 90, Naranja 50 a < 75, Rojo < 50.', 'Propuesta aceptada; rótulo «propuesta»: Isabella', True),
    ('D61, D62, D64, D66', 'Población y veedurías (método)', 'Población única: DANE vigente (8-ago-2025). Brecha de veedurías = consecutivos que faltan en el año / (listados + faltantes).', 'Propuesta aceptada; Isabella', False),
    ('D32, D40, D69', 'I1: votación de PP', 'Se reemplaza por el IPCM (votos / población). Línea base 2017 = 3,68 % con población DANE (la cifra oficial, 3,62 %, queda en nota); referente 7 %. El año es el de la votación. Votos usados: ' + pc_votos_txt() + '.', 'Profe; Isabella', True),
    ('D41, D73', 'I2: zonas', '12 zonas (sin la zona 1, industrial). Los años con dato (' + anios_dato('PC', 'PC.I2') + ') dan 12/12 con alerta: es cobertura documental por zona, sin actas de asistencia; se solicita la lista de asambleas por zona.', 'Profe; Isabella', True),
    ('D29, D40, D70, D74', 'I4: lista fija', 'Tres procesos: votación de PP, rendición de cuentas del Alcalde y rendición de PP. 2025 como rango (mín. 2/3, máx. 3/3); se actualiza cuando se publique la rendición de PP.', 'Profe; Isabella', True),
    ('D42', 'Pesos del ISPC', 'Se mantienen: 30/25/20/25.', 'Profe', True),
    ('D49, D72', 'I3: oferta formativa', 'Tasa solo en 2025: 21/25 = 84,00 (denominador reconstruido; se cuenta el diplomado en Gestión Pública). Respaldo: conteo 17, 7 y 21 acciones para la ciudadanía.', 'Isabella', True),
    ('D71', 'Consulta del portal', 'Se autoriza consultar el portal de la Alcaldía (incluida su API) para los votos de PP y archivar las noticias como fuente.', 'Isabella', False),
    ('D43', 'EJ90 en lugar de CEP', 'Principal: proyectos con ejecución ≥ 90 % (Hacienda, verificable). «Terminados» (autorreporte) solo como respaldo.', 'Profe', True),
    ('D24', 'FID en lugar de CMA', 'Fidelidad a lo votado: proyectos cuya apropiación coincide con el valor votado (±1 %) / proyectos votados.', 'Profe', True),
    ('D44', 'Pesos del IREPP', 'Se mantienen: 20/40/40.', 'Profe', True),
    ('D44, D53, D56', 'Serie de PP', 'Se calcula 2020–2025 y se reporta la brecha. Un año se calcula solo si se verifica al menos el 50 % de sus proyectos. ' + pp_sin_dato_txt() + ' ' + pp_universo_2021(), 'Profe; Isabella', True),
    ('D51, D52, D54, D55', 'Reglas de cruce de PP', 'Se excluye el rubro de déficit fiscal (2.3.7.06). Llave de cruce: valor votado = apropiación inicial; los cruces dudosos cuentan como no verificados.', 'Isabella', False),
    ('D45, D78, D82', 'PQRSDF en lugar de PCS', 'Oportunidad de respuesta (atendidas en términos / total). El año se arma con informes que cubran 12 meses sin repetirse; principal: sistema BPMS.', 'Profe; Isabella', True),
    ('D45, D75, D79', 'DSI', 'Lista de 10 metas de resultado del PDM, como rango (mín. = valores con fecha de publicación verificada) con brecha. Años 2024 y 2025.', 'Profe (mantener); Isabella (lista de 10, D75/D79)', True),
    ('D33, D45, D77, D81', 'VRC', 'Rango: mínimo = veeduría documentada / eventos (entra al índice); máximo suma los eventos sin soporte; la brecha debe bajar año a año. Se extiende a 2021–2023 con las actas de Salud, clasificadas a mano; faltan las rendiciones de Discapacidad y Víctimas.', 'Profe (mantener); Isabella (rango y brecha, D33; extensión, D77)', True),
    ('D76, D80', 'RPO', '2025 = 1/2 con alerta (cuadro corto). 2021 como rango (84,62–100). Universo: audiencia del Alcalde, sin felicitaciones.', 'Profe (mantener); Isabella (2021 y 2025, D76/D80)', True),
    ('D45', 'Brecha de información', 'Siempre se reporta: años (o eventos) sin dato / años del periodo.', 'Isabella (regla de la usuaria)', True),
    ('—', 'Sin redistribuir pesos', 'Un índice se calcula solo si están todos sus componentes; no se redistribuyen pesos.', 'Matriz / guía metodológica (regla 15)', True),
]


# ------------------------------------------------------------------ bloques reutilizables
def bloque_indices(doc, anchos=(2.2, 6.0, 3.4, 7.2)):
    """Un cuadro por índice: Año | Valor | Semáforo | Componentes que faltan."""
    parrafo(doc, 'Un índice se calcula solo en los años en que están todos sus componentes; no se redistribuyen pesos. ' + SEMAFORO)
    for nombre, dim, cod, a0, a1, formula in INDICES_DEF:
        doc.add_heading(nombre, level=2)
        parrafo(doc, f'{formula} Periodo: {a0}–{a1}.')
        filas_, k, n = datos_indice(dim, cod, a0, a1)
        brecha_csv = filas(dim, cod + '.BRECHA', 'brecha')[0]['brecha']
        assert f'{n - k} de {n}' in brecha_csv, (cod, k, n, brecha_csv)
        tabla(doc, list(anchos), ['Año', 'Valor', 'Semáforo (propuesta)', 'Componentes que faltan'], filas_, tam=9)
        parrafo(doc, linea_brecha(k, n))
        if cod == 'CS.ICRCS':
            mx = {r['anio']: r for r in filas('CS', 'CS.ICRCS.MAX')}
            parrafo(doc, 'El ICRCS se calcula en **2024 y 2025**, con PQRSDF leída de los 24 informes trimestrales de 2020–2025. 2020–2023 no tienen índice porque DSI solo aplica desde 2024. '
                         f'Con los máximos de los componentes (DSI, VRC y PQRSDF en su rango alto) el ICRCS sería {fmt(mx["2024"]["valor"])} en 2024 y {fmt(mx["2025"]["valor"])} en 2025. '
                         'Alertas: PQRSDF no es homogénea entre años (plazos ampliados por el Decreto 491 solo hasta el 17-may-2022; salto del volumen en 2024; sistema Débora desde mayo de 2025) y subestima la oportunidad real.')
        if cod == 'PC.ISPC':
            mx = filas('PC', 'PC.ISPC.MAX')[0]
            parrafo(doc, f'En 2025, I4 entra con su mínimo (2/3). Cuando se publique la rendición de PP 2025, el ISPC sería {fmt(mx["valor"])}.')


def cuadro_indicadores(doc, dim, titulo='Indicadores: datos disponibles', anchos=None):
    doc.add_heading(titulo, level=2)
    tabla(doc, anchos or W_IND, ENC_IND, tabla_indicadores_dim(dim), tam=9)


# ------------------------------------------------------------------ documento 1: para el profe
TITULOS = (('PDM', '1. Cumplimiento del Plan de Desarrollo (PDM)', PDM), ('PC', '2. Participación ciudadana', PC),
           ('PP', '3. Presupuesto participativo', PP), ('CS', '4. Control social', CS))


def doc_profe():
    doc = nuevo_documento(horizontal=True)
    doc.add_paragraph(style='Title').add_run('Indicadores del Observatorio: cambios y resultados')
    parrafo(doc, f'Isabella Hoyos Álvarez · {FECHA} · Resultados **provisionales**')
    parrafo(doc, 'Cambios acordados en la matriz y resultados de la construcción de los cuatro índices. '
                 'Criterios: serie en el tiempo, datos públicos, sin autorreporte y meta del mismo universo que el dato.')
    # el Word del profe no lleva la columna «Origen» (va solo en el documento detallado)
    sin_origen = [c for c in range(len(ENC)) if ENC[c] != 'Origen']
    w = [4.6, 5.6, 9.2, 4.4]
    for dim, titulo, filas_ in TITULOS:
        doc.add_heading(titulo, level=1)
        tabla(doc, w, [ENC[c] for c in sin_origen], [[f[c] for c in sin_origen] for f in filas_])
        cuadro_indicadores(doc, dim)
        if dim == 'PP':
            parrafo(doc, AVISO_PP2020)
        if dim == 'PDM':
            doc.add_heading('Espacio verde: cortes y fuentes', level=2)
            tabla(doc, [1.4, 2.6, 2.8, 7.0, 10.0], EV_ENC, ev_cortes())
            parrafo(doc, EV_NOTA_CORTA)
    doc.add_heading('5. Resultados de los índices', level=1)
    bloque_indices(doc)
    doc.add_heading('6. Decisiones tomadas', level=1)
    tabla(doc, [5.0, 18.8], ['Tema', 'Decisión'], [[f'**{t}**', d] for _, t, d, _, v in DECISIONES if v], tam=9)
    doc.add_heading('7. Variables por solicitar (informe futuro)', level=1)
    vinetas(doc, SOLICITAR)
    ruta = OUT / 'indicadores_cambios_y_datos_OCE.docx'
    doc.save(ruta)
    return ruta


# ------------------------------------------------------------------ documento 2: detalle para Isabella
SCRIPTS = [
    ('PP', 'outputs/construccion/PP/pp_indicadores.csv',
     'scripts/construccion/pp_lectores.py → pp_cruce.py → pp_muestra_manual.py → pp_irepp.py',
     'También: pp_cruce_proyectos.csv, pp_ajustes.csv y pp_muestra_manual.csv'),
    ('PDM', 'outputs/construccion/PDM/pdm_indicadores.csv', 'scripts/construccion/pdm_icpd7.py',
     '44 asserts; tarda unos 3 minutos (lee la PPED cruda del DANE)'),
    ('PC', 'outputs/construccion/PC/pc_indicadores.csv', 'scripts/construccion/pc_ispc.py', '56 asserts; tarda unos 3 minutos'),
    ('CS', 'outputs/construccion/CS/cs_indicadores.csv', 'scripts/construccion/cs_descargas.py → cs_icrcs.py',
     '64 asserts; el CSV solo se escribe si pasan todos'),
]
NOMBRE_DIM = {'PDM': 'Cumplimiento del Plan de Desarrollo (PDM)', 'PC': 'Participación ciudadana',
              'PP': 'Presupuesto participativo', 'CS': 'Control social'}
CMD_PQRSDF = ('python3 scripts/construccion/cs_descargas.py && python3 scripts/construccion/cs_icrcs.py '
              '&& python3 scripts/entregables/generar_entrega_profe.py && python3 scripts/entregables/generar_base_powerbi.py')


def primer_enlace(s):
    m = re.search(r'https?://[^\s;]+', s)
    return m.group(0) if m else s


def resumen(texto, n=230):
    texto = re.sub(r'\s+', ' ', texto).strip().replace('anios', 'años').replace('anio', 'año')
    return texto if len(texto) <= n else texto[:n].rsplit(' ', 1)[0] + '…'


def alertas(dim, cod, maximo_n=3):
    """Alertas distintas de la columna nota, resumidas."""
    vistas, out = set(), []
    for r in anual(dim, cod, 'principal'):
        for trozo in re.split(r'\bALERTAS?(?: DE COMPARABILIDAD)?\b', r['nota'])[1:]:
            t = resumen(trozo.strip(' :'))
            clave = t[:50]
            if clave not in vistas:
                vistas.add(clave)
                out.append(t)
    return out[:maximo_n]


def pegar_serie(doc, dim, cod, nombre):
    doc.add_heading(nombre, level=3)
    rs = anual(dim, cod, 'principal')
    con = con_dato(rs)
    ref = con[-1] if con else rs[-1]
    parrafo(doc, f'**Referente:** {resumen(ref["referente"], 200)}. **Fuente:** {resumen(ref["fuente"], 190)}. '
                 f'**Enlace:** {primer_enlace(ref["enlace"])}')
    cuerpo = []
    for r in rs:
        nd = ''
        if num(r['valor']) is None:
            nd = resumen(r['brecha'].replace('sin dato: ', '').replace('anios', 'años'), 95)
        elif r['brecha'] and r['brecha'] != 'con dato' and dim == 'PDM' and cod == 'PDM.VEED':
            nd = resumen(r['brecha'].replace('con dato; ', ''), 95)
        mx = maximo(dim, cod, r['anio'])
        vv = valor_fila(r)
        val = fmt(vv) if mx is None or num(mx) == num(r['valor']) else f'{fmt(vv)}–{fmt(mx)}'
        nd_den = '' if r['numerador'] in ('', 'N/D') else f'{fmt(r["numerador"])} / {fmt(r["denominador"])}' if r['denominador'] else fmt(r['numerador'])
        cuerpo.append([r['anio'], nd_den, val, fmt(r['puntaje_0_100']) if num(r['puntaje_0_100']) is not None else '—',
                       r['semaforo'] or '—', nd])
    tabla(doc, [1.4, 3.4, 2.2, 1.8, 2.4, 6.1], ['Año', 'Numerador / denominador', 'Valor', 'Puntaje', 'Semáforo (propuesta)', 'Dato faltante / nota'],
          cuerpo, tam=8)
    parrafo(doc, f'**Brecha:** {texto_brecha(dim, cod, corto=False)}.')
    al = alertas(dim, cod)
    if not al:  # sin alertas explícitas: se resume el método de la última fila
        al = [resumen(ref['nota'], 230)]
    al = [('Alerta ' + t if t.startswith(('de ', '(')) else t) for t in al]
    for t in al:
        escribir(doc.add_paragraph(style='List Bullet'), f'**Alerta o método:** {t[0].upper() + t[1:]}', 9)


def doc_detalle():
    doc = nuevo_documento(horizontal=False)
    doc.add_paragraph(style='Title').add_run('Revisión de indicadores: cambios y resultados de la construcción')
    parrafo(doc, f'Documento de trabajo para Isabella · {FECHA} · Todo es **provisional**')
    doc.add_heading('Resumen', level=1)
    vinetas(doc, [
        'Se construyeron los cuatro índices con todos los años que tiene cada indicador. Los CSV están en outputs/construccion/<dimensión>/ y los genera un script por dimensión.',
        '**IREPP (PP):** ' + comprimir([a for a, (v, _s) in ((int(r['anio']), (num(r['valor']), '')) for r in anual('PP', 'PP.IREPP', 'indice')) if v is not None]) + (' (' + pp_sin_dato_txt() + ')' if pp_anios_sin_dato() else '') + '. **ICPD-7 (PDM):** 2019, 2022 y 2024 (espacio verde solo tiene tres cortes). '
        '**ISPC (PC):** solo 2025. **ICRCS (CS):** 2024 y 2025 (PQRSDF ya está completa 2020–2025; 2020–2023 no tienen índice porque DSI solo aplica desde 2024).',
        'Todo es Provisional. QA aprobó con observaciones las cuatro dimensiones (18 y 19-nov).',
        'Los datos crudos están en Google Drive y las descargas de esta revisión en data_raw_IIC/07_revision_nov2026/ (con manifiesto).',
    ])
    doc.add_heading('Resultados de los índices', level=1)
    bloque_indices(doc, (2.0, 5.0, 3.0, 7.4))

    for dim, _, filas_ in TITULOS:
        doc.add_heading(NOMBRE_DIM[dim], level=1)
        doc.add_heading('Cambios frente a la matriz', level=2)
        tabla(doc, [3.0, 3.2, 5.2, 2.7, 3.4], ENC, filas_, tam=8)
        cuadro_indicadores(doc, dim, 'Resumen de los indicadores', [4.6, 3.4, 4.8, 4.6])
        doc.add_heading('Serie completa de cada indicador', level=2)
        for cod, nombre in NOMBRES[dim]:
            pegar_serie(doc, dim, cod, nombre)
            if cod == 'CS.PQRSDF':
                nota_p = ' '.join(r['nota'] for r in anual('CS', 'CS.PQRSDF', 'principal'))
                assert '17-may-2022' in nota_p and '+64 %' in nota_p and 'D86' in nota_p
                for t in ('**Decreto 491:** amplió los plazos para responder solo **hasta el 17-may-2022** (lo derogó la Ley 2207). 2020, 2021 y 2022 T1 se miden con plazos ampliados; 2022 T2 es mixto; 2022 T3–T4 y 2023–2025 con plazos ordinarios (Ley 1755). No comparar los años sin esta salvedad.',
                          '**Salto de 2024:** las recibidas suben de 19.502 a 31.970 (+64 %) por «acciones de mejora en la correcta clasificación de todo lo que ingresa» (informes de 2T y 3T de 2024). Trámite/Servicio pasa del 2,2 % (2023 T1) al 14,5–18,0 %. De 2023 a 2024 no se debe leer como tendencia; 2024 T1 tiene su propia explicación (inicio de año, impuestos, cambio de gobierno).',
                          '**«Sin respuesta» (2020 T4–2021 T4):** no se sabe cuántas estaban dentro del plazo; va al máximo del rango y a la brecha (D86): máximo 90,37 en 2020 y 96,56 en 2021.',
                          '**Sistema:** BPMS hasta 2025 T1; Débora entra a producción en mayo de 2025. El indicador subestima la oportunidad (la auditoría de Control Interno del S1-2025 lo confirma). Brecha en trimestres: 0 de 24.'):
                    escribir(doc.add_paragraph(style='List Bullet'), t, 9)
        if dim == 'PP':
            parrafo(doc, AVISO_PP2020)
        if dim == 'PDM':
            doc.add_heading('Espacio verde: cortes y fuentes', level=2)
            tabla(doc, [1.4, 2.0, 2.2, 5.0, 6.8], EV_ENC, ev_cortes(), tam=8)
            parrafo(doc, EV_NOTA)
            vinetas(doc, EV_URLS)
        if dim == 'CS':
            doc.add_heading('Qué se necesita para seguir midiendo VRC', level=2)
            vinetas(doc, VRC_NECESITA)
        doc.add_heading('Scripts y archivos', level=2)
        d, csv_, scr, nota = [s for s in SCRIPTS if s[0] == dim][0]
        vinetas(doc, [f'**CSV:** {csv_}', f'**Scripts:** {scr}', nota])

    doc.add_heading('Decisiones tomadas (con número y origen)', level=1)
    parrafo(doc, 'Los números D remiten a .claude/memoria/decisiones.md. Origen: Profe, Isabella o propuesta aceptada.')
    tabla(doc, [2.0, 3.0, 8.2, 4.3], ['D', 'Tema', 'Decisión', 'Origen'],
          [[d, f'**{t}**', x, o] for d, t, x, o, _ in DECISIONES], tam=8)

    doc.add_heading('Pendientes', level=1)
    vinetas(doc, [
        '**PQRSDF (CS):** los 24 informes trimestrales 2020–2025 están leídos (23 de forma automática y 1 manual: 2023 T1, cuyo Alcance trae un error de copia). Quedan alertas de comparabilidad: '
        'plazos ampliados por el Decreto 491 hasta el 17-may-2022 (2022 T2 mixto; T3–T4 con plazos ordinarios), columna «Sin respuesta» de 2020 T4–2021 T4 (va al máximo del rango, D86), salto de 2024 y sistema Débora desde mayo de 2025.',
        '**Discapacidad y Víctimas 2021–2023:** el inventario no las tiene; buscar sus rendiciones en la API del portal cuando vuelva, para que el universo de VRC sea el mismo de 2024–2025. Clasificarlas es manual.',
        '**RPO 2021:** rango 0–100 (D87): daría 84,62 si se confirma que el 23449 se publicó a más tardar el 31-dic-2021. El 40433 y el 40434 son encuestas, no respuestas escritas.',
        '**Solicitudes a Planeación y otras dependencias:** ' + ' '.join(f'({i}) {t}' for i, t in enumerate(SOLICITAR, 1)),
        '**Rendición de PP 2025:** se hará en el último trimestre de 2026; entonces se actualiza I4 2025 y el ISPC 2025.',
    ])
    ruta = OUT / 'detalle_revision_indicadores_OCE.docx'
    doc.save(ruta)
    return ruta


# ------------------------------------------------------------------ verificaciones
PROHIBIDAS_PROFE = ('origen', 'claude', 'asistente')
ESPERADOS = {'ICPD-7': ('2019', '62,41'), 'ISPC': ('2025', '61,62'), 'IREPP': ('2025', '95,83'), 'ICRCS': ('2024', '60,59')}


def texto_completo(d):
    partes = [p.text for p in d.paragraphs]
    for t in d.tables:
        partes += [c.text for f in t.rows for c in f.cells]
    for s in d.sections:
        partes += [p.text for p in s.footer.paragraphs] + [p.text for p in s.header.paragraphs]
    return '\n'.join(partes)


# Excepción explícita: la línea de autoría de la portada del Word del profe («Isabella Hoyos Álvarez · fecha · …») SÍ lleva el nombre,
# porque ella es la autora del documento y se lo envía al profe. El filtro de nombres no se aplica a esa línea.
AUTORIA = 'Isabella Hoyos Álvarez'
NOMBRES_PROHIBIDOS_PROFE = ('Isabella', 'Francisco', 'usuaria', 'Analizador')


def verificar(ruta, profe):
    d = Document(ruta)
    sec = d.sections[0]
    util = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
    # 1) anchos de cada tabla menores que el ancho útil
    for i, t in enumerate(d.tables):
        ancho = sum(c.width.cm for c in t.rows[0].cells)
        assert ancho <= util + 0.01, (ruta.name, 'tabla', i, round(ancho, 2), round(util, 2))
    assert not re.findall(r'\bano\b(?!-\d)', re.sub(r'https?://\S+', '', texto_completo(d))), (ruta.name, 'palabra «ano» sin tilde')
    # 2) palabras prohibidas en el Word del profe
    if profe:
        txt = texto_completo(d).lower()
        for p in PROHIBIDAS_PROFE:
            assert p not in txt, (ruta.name, 'contiene', p)
        sin_autoria = '\n'.join(l for l in texto_completo(d).split('\n') if not l.startswith(AUTORIA))
        assert AUTORIA in texto_completo(d), 'falta la línea de autoría de la portada'
        for n in NOMBRES_PROHIBIDOS_PROFE:
            assert n not in sin_autoria, (ruta.name, 'contiene el nombre', n)
        assert not re.search(r'\bD\d{1,3}\b', sin_autoria), (ruta.name, 'número de decisión', re.findall(r'.{20}\bD\d{1,3}\b.{10}', sin_autoria)[:3])
        assert 'Origen' not in [c.text for c in d.tables[0].rows[0].cells]
    # 3) valores de los índices (cuadro bajo el título del índice)
    tablas = {t._tbl: t for t in d.tables}
    titulo, hallados = '', {}
    for el in d.element.body.iterchildren():
        if el.tag == qn('w:p'):
            st = el.find(qn('w:pPr'))
            if st is not None and st.find(qn('w:pStyle')) is not None and 'Heading' in st.find(qn('w:pStyle')).get(qn('w:val')):
                titulo = ''.join(x.text or '' for x in el.iter(qn('w:t')))
        elif el.tag == qn('w:tbl') and titulo:
            t = tablas[el]
            enc = [c.text for c in t.rows[0].cells]
            if enc[:3] == ['Año', 'Valor', 'Semáforo (propuesta)']:
                for clave in {c.split('.')[1].replace('ICPD7', 'ICPD-7') for _n, _d, c, *_r in INDICES_DEF}:
                    if titulo.startswith(clave):
                        hallados[clave] = {f.cells[0].text: f.cells[1].text for f in t.rows[1:]}
    for clave, (anio, valor) in ESPERADOS.items():
        assert hallados[clave][anio] == valor, (ruta.name, clave, anio, hallados[clave].get(anio))
    cs_ = {r['anio']: r for r in anual('CS', 'CS.ICRCS', 'indice')}
    assert fmt(cs_['2024']['valor']) == '60,59' and fmt(cs_['2025']['valor']) == '38,18', ('ICRCS del CSV', cs_['2024']['valor'], cs_['2025']['valor'])
    assert hallados['ICRCS']['2024'] == '60,59' and hallados['ICRCS']['2025'] == '38,18', (ruta.name, 'ICRCS', hallados['ICRCS'])
    # y contra los CSV
    for _, dim, cod, a0, a1, _f in INDICES_DEF:
        clave = cod.split('.')[1].replace('ICPD7', 'ICPD-7')
        for r in anual(dim, cod, 'indice'):
            if a0 <= int(r['anio']) <= a1 and num(r['valor']) is not None:
                assert hallados[clave][r['anio']] == fmt(r['valor']), (ruta.name, clave, r['anio'])
    return d, len(d.paragraphs), len(d.tables), [len(t.rows) for t in d.tables]


def main():
    rutas = (doc_profe(), doc_detalle())
    for ruta, profe in zip(rutas, (True, False)):
        d, np_, nt, filas_ = verificar(ruta, profe)
        titulos = [p.text for p in d.paragraphs if p.style.name.startswith('Heading 1')]
        print(ruta.relative_to(RAIZ), '| párrafos:', np_, '| tablas:', nt)
        print('  secciones:', titulos)
        print('  filas por tabla:', filas_)
    # las dudas del profe deben estar todas marcadas como respondidas
    dudas = (OUT / 'dudas_para_el_profe.md').read_text(encoding='utf-8')
    n_dudas = len(re.findall(r'^\d+b?\. ', dudas, flags=re.M))
    n_resp = len(re.findall(r'Respondida', dudas))
    assert n_dudas == n_resp == 16, (n_dudas, n_resp)
    print('Asserts OK: sin «Origen», «Claude» ni «asistente» en el Word del profe; IREPP 2025 = 95,83, '
          'ICPD-7 2019 = 62,41, ISPC 2025 = 61,62, ICRCS 2024 = 60,59 y 2025 = 38,18 en ambos Word y coinciden con los CSV; anchos de tabla dentro del ancho útil; '
          f'{n_resp} dudas respondidas.')


if __name__ == '__main__':
    main()
