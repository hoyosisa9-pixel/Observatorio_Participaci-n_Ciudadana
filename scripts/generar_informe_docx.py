"""Genera el informe de disponibilidad y brecha en Word (.docx) con el mismo contenido que el PDF de LaTeX.

Uso (desde la raiz del repo):  python3 scripts/generar_informe_docx.py
Requiere openpyxl y python-docx (instalado en el site-packages del usuario: no usar -I).

Fuente unica: el texto sale de scripts/generar_informe_latex.py (funcion construir), que arma el .tex desde el
Excel pulido y los datos de PP. Este script traduce ese mismo .tex a Word con estilos nativos (Titulo, Titulo 1-4,
Normal, Lista con viñetas, Lista numerada), tabla de contenido como campo TOC, tablas nativas con el estilo del PDF
(lineas gruesas arriba y abajo, fina bajo el encabezado, sin verticales), colores de categoria y veredicto tomados
de los \\definecolor del .tex, recuadro "PRELIMINAR, POR VALIDAR" e hipervinculos clicables.
Salida: outputs/disponibilidad/informe_disponibilidad.docx
"""
import re, sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generar_informe_latex as L  # noqa: E402

SALIDA = L.DISP / 'informe_disponibilidad.docx'
FUENTE = 'Cambria'
MONO = 'Consolas'
ANCHO_TEXTO = 21.59 - 2 * 2.5  # cm (carta, márgenes de 2,5 cm)
ENCABEZADO = 'OCE – Disponibilidad y brecha de información'
COLOR_BORDE_AVISO = '990000'  # red!60!black del .tex


# =========================================================================== lectura del .tex
def leer_arg(s, i):
    """Lee un argumento {…} balanceado que empieza en s[i] (saltando espacios). Devuelve (contenido, indice final)."""
    while i < len(s) and s[i] in ' \n':
        i += 1
    if i >= len(s) or s[i] != '{':
        return '', i
    nivel, j = 0, i
    while j < len(s):
        if s[j] == '\\':
            j += 2
            continue
        if s[j] == '{':
            nivel += 1
        elif s[j] == '}':
            nivel -= 1
            if nivel == 0:
                return s[i + 1:j], j + 1
        j += 1
    raise ValueError('llave sin cerrar: ' + s[i:i + 80])


def partir_nivel_superior(s, sep_regex):
    """Divide s por sep_regex solo fuera de llaves (para filas '\\\\' y celdas '&')."""
    partes, nivel, inicio, i = [], 0, 0, 0
    patron = re.compile(sep_regex)
    while i < len(s):
        c = s[i]
        if c == '\\' and nivel == 0:
            m = patron.match(s, i)
            if m:
                partes.append(s[inicio:i])
                i = inicio = m.end()
                continue
            i += 2
            continue
        if c == '\\':
            i += 2
            continue
        if c == '{':
            nivel += 1
        elif c == '}':
            nivel -= 1
        elif nivel == 0:
            m = patron.match(s, i)
            if m:
                partes.append(s[inicio:i])
                i = inicio = m.end()
                continue
        i += 1
    partes.append(s[inicio:])
    return partes


SIMBOLOS = {'ldots': '…', 'textquestiondown': '¿', 'textexclamdown': '¡', 'textasciitilde': '~', 'textbar': '|',
            'textbackslash': '\\', 'textasciicircum': '^', 'textless': '<', 'textgreater': '>',
            'textperiodcentered': '·', 'guillemotleft': '«', 'guillemotright': '»', 'S': '§', 'approx': '≈',
            'times': '×', 'leq': '≤', 'geq': '≥', 'neq': '≠', 'rightarrow': '→', 'leftarrow': '←',
            'leftrightarrow': '↔', 'textbullet': '•', 'quad': '\u2003'}
IGNORAR = {'allowbreak', 'footnotesize', 'scriptsize', 'small', 'large', 'LARGE', 'noindent', 'smallskip', 'par',
           'bfseries', 'centering', 'normalsize', 'raggedright'}
ESTILO = {'textbf': 'b', 'emph': 'i', 'texttt': 'mono', 'textsuperscript': 'sup', 'mathrm': None, 'centerline': None}


def en_linea(s, estado=None, runs=None, celda=None):
    """Traduce texto LaTeX en una lista de runs: dict(t=texto, b, i, mono, sup, link, fill, br)."""
    estado = estado or {}
    runs = [] if runs is None else runs
    buf = []

    def volcar():
        if buf:
            runs.append({**estado, 't': ''.join(buf)})
            buf.clear()

    i = 0
    while i < len(s):
        c = s[i]
        if c == '\\':
            m = re.match(r'\\([a-zA-Z@]+)\*?|\\(.)', s[i:], re.S)
            if m.group(1):
                cmd = m.group(1)
                i += m.end()
                while i < len(s) and s[i] == ' ':  # TeX descarta los espacios tras una orden con nombre
                    i += 1
                if cmd in ESTILO:
                    arg, i = leer_arg(s, i)
                    volcar()
                    clave = ESTILO[cmd]
                    en_linea(arg, {**estado, clave: True} if clave else estado, runs, celda)
                elif cmd == 'href':
                    url, i = leer_arg(s, i)
                    txt, i = leer_arg(s, i)
                    volcar()
                    url = re.sub(r'\\([%#&~])', r'\1', url)
                    en_linea(txt, {**estado, 'link': url}, runs, celda)
                elif cmd == 'colorbox':
                    col, i = leer_arg(s, i)
                    txt, i = leer_arg(s, i)
                    volcar()
                    en_linea(txt, {**estado, 'fill': col}, runs, celda)
                elif cmd == 'cellcolor':
                    col, i = leer_arg(s, i)
                    if celda is not None:
                        celda['fill'] = col
                elif cmd == 'ref':
                    _, i = leer_arg(s, i)
                    buf.append('A')
                elif cmd == 'newline':
                    volcar()
                    runs.append({**estado, 't': '', 'br': True})
                elif cmd in SIMBOLOS:
                    buf.append(SIMBOLOS[cmd])
                    if s.startswith('{}', i):
                        i += 2
                elif cmd in IGNORAR:
                    if s.startswith('{}', i):
                        i += 2
                else:
                    raise ValueError(f'orden LaTeX no prevista: \\{cmd} en «{s[max(0, i - 60):i + 40]}»')
            else:
                buf.append(m.group(2))
                i += 2
        elif c == '{':
            arg, i = leer_arg(s, i)
            volcar()
            en_linea(arg, estado, runs, celda)
        elif c == '$':
            i += 1
        elif c == '~':
            buf.append('\u00a0')
            i += 1
        elif s.startswith('---', i):
            buf.append('—')
            i += 3
        elif s.startswith('--', i):
            buf.append('–')
            i += 2
        elif s.startswith('``', i):
            buf.append('“')
            i += 2
        elif s.startswith("''", i):
            buf.append('”')
            i += 2
        elif c == '\n':
            buf.append(' ')
            i += 1
        else:
            buf.append(c)
            i += 1
    volcar()
    return runs


def texto_plano(runs):
    return ''.join(r['t'] for r in runs).strip()


def columnas_de_spec(spec):
    """Tokens de la especificación de columnas: ('L', cm) | ('Y',) | ('l'|'c'|'r',)."""
    toks, i = [], 0
    while i < len(spec):
        c = spec[i]
        if c in 'Lp':
            arg, i = leer_arg(spec, i + 1)
            toks.append(('L', float(re.match(r'([\d.]+)cm', arg).group(1))))
        elif c in 'YX':
            toks.append(('Y',))
            i += 1
        elif c in 'lcr':
            toks.append((c,))
            i += 1
        else:
            i += 1
    return toks


def leer_filas(cuerpo):
    """Filas de una tabla: lista de (celdas LaTeX, regla_arriba)."""
    filas = []
    for trozo in partir_nivel_superior(cuerpo, r'\\\\(?![a-zA-Z])'):
        regla = bool(re.search(r'\\midrule', trozo))
        t = re.sub(r'\\(toprule|midrule|bottomrule)', '', trozo).strip()
        if not t:
            if regla and filas:
                filas[-1] = (filas[-1][0], filas[-1][1], True)  # la regla queda debajo de la fila anterior
            continue
        filas.append((partir_nivel_superior(t, r'&'), regla, False))
    return filas


def bloques(tex):
    """Convierte el cuerpo del .tex en una lista de bloques."""
    cuerpo = tex[tex.index(r'\begin{document}') + len(r'\begin{document}'):tex.index(r'\end{document}')]
    res, i, apendice = [], 0, False
    tam_tabla = None
    while i < len(cuerpo):
        if cuerpo[i] in ' \n':
            i += 1
            continue
        resto = cuerpo[i:]
        if resto.startswith(r'\hypersetup') or resto.startswith(r'\begingroup') or resto.startswith(r'\endgroup'):
            i = cuerpo.find('\n', i) if '\n' in resto else len(cuerpo)
            continue
        if resto.startswith(r'\begin{titlepage}'):
            fin = cuerpo.index(r'\end{titlepage}', i)
            portada = []
            for m in re.finditer(r'\{\\(LARGE|large)(?:\\bfseries)? (.*?)\\par\}', cuerpo[i:fin]):
                portada.append((m.group(1), texto_plano(en_linea(m.group(2)))))
            res.append(('portada', portada))
            i = fin + len(r'\end{titlepage}')
            continue
        if resto.startswith(r'\tableofcontents'):
            res.append(('indice',))
            i += len(r'\tableofcontents')
            continue
        if resto.startswith(r'\clearpage'):
            res.append(('salto',))
            i += len(r'\clearpage')
            continue
        if resto.startswith(r'\appendix'):
            apendice = True
            i += len(r'\appendix')
            continue
        m = re.match(r'\\(section|subsection|subsubsection|paragraph)(\*?)', resto)
        if m:
            titulo, j = leer_arg(cuerpo, i + m.end())
            res.append(('titulo', m.group(1), bool(m.group(2)), texto_plano(en_linea(titulo)), apendice))
            j2 = re.match(r'\\label\{[^}]*\}', cuerpo[j:])
            i = j + (j2.end() if j2 else 0)
            continue
        m = re.match(r'\{\\(scriptsize|footnotesize)\s*\n?\\begin\{xltabular\}', resto)
        if m:
            fin = cuerpo.index(r'\end{xltabular}', i)
            bloque = cuerpo[i + m.end():fin]
            _, j = leer_arg(bloque, 0)          # {\linewidth}
            spec, j = leer_arg(bloque, j)
            enc = bloque[j:bloque.index(r'\endfirsthead')]
            filas_enc = leer_filas(enc)
            filas = leer_filas(bloque[bloque.index(r'\endlastfoot') + len(r'\endlastfoot'):])
            res.append(('tabla', columnas_de_spec(spec), filas_enc[0][0], filas, 8 if apendice else 9))
            i = fin + len(r'\end{xltabular}')
            m2 = re.match(r'\s*\n\}', cuerpo[i:])
            i += m2.end() if m2 else 0
            continue
        if resto.startswith(r'\begin{center}'):
            fin = cuerpo.index(r'\end{center}', i)
            dentro = cuerpo[i + len(r'\begin{center}'):fin]
            i = fin + len(r'\end{center}')
            if r'\fcolorbox' in dentro:
                m = re.search(r'\\bfseries\s*(.*?)\}\}', dentro, re.S)
                res.append(('aviso', m.group(1).strip()))
                continue
            m = re.search(r'\\begin\{(tabularx|tabular)\}', dentro)
            entorno = m.group(1)
            j = m.end()
            if entorno == 'tabularx':
                _, j = leer_arg(dentro, j)
            spec, j = leer_arg(dentro, j)
            filas = leer_filas(dentro[j:dentro.index(r'\end{' + entorno + '}')])
            res.append(('tabla', columnas_de_spec(spec), filas[0][0], filas[1:], 9))
            continue
        m = re.match(r'\\begin\{(itemize|enumerate)\}', resto)
        if m:
            fin = cuerpo.index(r'\end{' + m.group(1) + '}', i)
            items = [x.strip() for x in partir_nivel_superior(cuerpo[i + m.end():fin], r'\\item\b')[1:]]
            res.append(('lista', m.group(1) == 'enumerate', [en_linea(x) for x in items]))
            i = fin + len(r'\end{' + m.group(1) + '}')
            continue
        # párrafo: hasta línea en blanco o inicio de bloque
        fin = len(cuerpo)
        for patron in (r'\n\s*\n', r'\n\\(?:begin|section|subsection|subsubsection|paragraph|appendix|clearpage)'):
            mm = re.compile(patron).search(cuerpo, i)
            if mm:
                fin = min(fin, mm.start())
        texto = cuerpo[i:fin].strip()
        i = fin
        if texto.startswith(r'\par\smallskip'):
            res.append(('nota', en_linea(texto)))
        elif r'\centerline' in texto:
            antes, despues = texto.split(r'\centerline', 1)
            res.append(('parrafo', en_linea(antes), None))
            arg, _ = leer_arg(despues, 0)
            res.append(('parrafo', en_linea(arg), 'centro'))
        else:
            res.append(('parrafo', en_linea(texto), None))
    return res


# =========================================================================== escritura del .docx
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
    estilo.font.italic = False if estilo.name.startswith('Heading') else estilo.font.italic


def preparar_estilos(doc):
    st = doc.styles
    fijar_fuente(st['Normal'], FUENTE, 11)
    pf = st['Normal'].paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)
    negro = RGBColor(0, 0, 0)
    for nombre, tam in (('Title', 20), ('Heading 1', 15), ('Heading 2', 12.5), ('Heading 3', 11.5),
                        ('Heading 4', 11), ('Heading 5', 11)):
        fijar_fuente(st[nombre], FUENTE, tam, True, negro)
        st[nombre].paragraph_format.space_before = Pt(14 if nombre == 'Heading 1' else 10)
        st[nombre].paragraph_format.space_after = Pt(6)
        st[nombre].paragraph_format.keep_with_next = True
    title_ppr = st['Title'].element.get_or_add_pPr()
    for b in title_ppr.findall(qn('w:pBdr')):
        title_ppr.remove(b)
    for nombre in ('List Bullet', 'List Number', 'TOC Heading'):
        if nombre in [s.name for s in st]:
            fijar_fuente(st[nombre], FUENTE)


def sombrear(elemento_pr, color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    elemento_pr.append(shd)


def agregar_runs(parrafo, runs, colores, tam=None):
    """Escribe runs en el párrafo; los runs consecutivos con el mismo enlace forman un solo hipervínculo."""
    k = 0
    while k < len(runs):
        r = runs[k]
        if r.get('link'):
            grupo = [r]
            while k + 1 < len(runs) and runs[k + 1].get('link') == r['link']:
                k += 1
                grupo.append(runs[k])
            rid = parrafo.part.relate_to(r['link'], RT.HYPERLINK, is_external=True)
            hl = OxmlElement('w:hyperlink')
            hl.set(qn('r:id'), rid)
            for g in grupo:
                run = parrafo.add_run()
                escribir_run(run, g, colores, tam, enlace=True)
                hl.append(run._r)
            parrafo._p.append(hl)
        else:
            escribir_run(parrafo.add_run(), r, colores, tam)
        k += 1


def escribir_run(run, r, colores, tam, enlace=False):
    if r.get('br'):
        run.add_break()
        return
    run.text = r['t']
    run.bold = True if r.get('b') else None
    run.italic = True if r.get('i') else None
    if r.get('mono'):
        run.font.name = MONO
        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:hAnsi'), MONO)
    if r.get('sup'):
        run.font.superscript = True
    if tam:
        run.font.size = Pt(tam - (1 if r.get('mono') else 0))
    elif r.get('mono'):
        run.font.size = Pt(9.5)
    if enlace:
        run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        run.font.underline = True
    if r.get('fill'):
        sombrear(run._r.get_or_add_rPr(), colores[r['fill']])


def borde(padre_pr, etiqueta, lados):
    """lados: {lado: (val, sz)}"""
    b = padre_pr.find(qn(etiqueta))
    if b is None:
        b = OxmlElement(etiqueta)
        padre_pr.append(b)
    for lado, (val, sz) in lados.items():
        e = OxmlElement('w:' + lado)
        e.set(qn('w:val'), val)
        if val != 'nil':
            e.set(qn('w:sz'), str(sz))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), 'auto')
        b.append(e)
    return b


def anchos_columnas(spec, filas):
    fijos = sum(t[1] for t in spec if t[0] == 'L')
    flex = [k for k, t in enumerate(spec) if t[0] != 'L']
    anchos = [t[1] if t[0] == 'L' else 0 for t in spec]
    if flex:
        if all(spec[k][0] == 'Y' for k in flex):
            for k in flex:
                anchos[k] = (ANCHO_TEXTO - fijos) / len(flex)
        else:  # l/c/r: proporcional al texto más largo de cada columna
            largo = [max(len(texto_plano(en_linea(f[k]))) if k < len(f) else 0 for f in filas) + 4 for k in flex]
            for k, n in zip(flex, largo):
                anchos[k] = (ANCHO_TEXTO - fijos) * n / sum(largo)
            minimo = 1.6
            faltante = sum(max(0, minimo - anchos[k]) for k in flex)
            anchos_grandes = [k for k in flex if anchos[k] > minimo]
            total_grandes = sum(anchos[k] for k in anchos_grandes)
            for k in flex:
                if anchos[k] < minimo:
                    anchos[k] = minimo
                elif faltante:
                    anchos[k] -= faltante * anchos[k] / total_grandes
    return anchos


def recortar_espacios(runs):
    runs = [dict(r) for r in runs if r['t'] or r.get('br')]
    while runs and not runs[0].get('br') and not runs[0]['t'].strip():
        runs.pop(0)
    while runs and not runs[-1].get('br') and not runs[-1]['t'].strip():
        runs.pop()
    if runs and not runs[0].get('br'):
        runs[0]['t'] = runs[0]['t'].lstrip()
    if runs and not runs[-1].get('br'):
        runs[-1]['t'] = runs[-1]['t'].rstrip()
    return runs


def escribir_tabla(doc, spec, encabezado, filas, tam, colores):
    todas = [encabezado] + [f[0] for f in filas]
    anchos = anchos_columnas(spec, todas)
    tabla = doc.add_table(rows=len(todas), cols=len(spec))
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla.autofit = False
    tblpr = tabla._tbl.tblPr
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblpr.append(lay)
    borde(tblpr, 'w:tblBorders', {'top': ('single', 12), 'left': ('nil', 0), 'bottom': ('single', 12),
                                  'right': ('nil', 0), 'insideH': ('nil', 0), 'insideV': ('nil', 0)})
    for k, gc in enumerate(tabla._tbl.tblGrid.findall(qn('w:gridCol'))):
        gc.set(qn('w:w'), str(int(anchos[k] / 2.54 * 1440)))
    reglas = [False] + [f[2] for f in filas]
    for r, celdas in enumerate(todas):
        fila = tabla.rows[r]
        trpr = fila._tr.get_or_add_trPr()
        no_partir = OxmlElement('w:cantSplit')
        trpr.append(no_partir)
        if r == 0:
            trpr.append(OxmlElement('w:tblHeader'))
        for k in range(len(spec)):
            celda = fila.cells[k]
            celda.width = Cm(anchos[k])
            info = {}
            runs = en_linea(celdas[k] if k < len(celdas) else '', celda=info)
            if r == 0:
                runs = [{**x, 'b': True} for x in runs]
            runs = recortar_espacios(runs)
            p = celda.paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.line_spacing = 1.0
            if spec[k][0] == 'r':
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif spec[k][0] == 'c':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            agregar_runs(p, [x for x in runs if x['t'] or x.get('br')], colores, tam)
            tcpr = celda._tc.get_or_add_tcPr()
            va = OxmlElement('w:vAlign')
            va.set(qn('w:val'), 'top')
            tcpr.append(va)
            if info.get('fill'):
                sombrear(tcpr, colores[info['fill']])
            if r == 0 or reglas[r]:
                borde(tcpr, 'w:tcBorders', {'bottom': ('single', 4)})
    return tabla


def escribir_aviso(doc, texto):
    tabla = doc.add_table(rows=1, cols=1)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblpr = tabla._tbl.tblPr
    borde(tblpr, 'w:tblBorders', {x: ('single', 8) for x in ('top', 'left', 'bottom', 'right')})
    for e in tblpr.find(qn('w:tblBorders')):
        e.set(qn('w:color'), COLOR_BORDE_AVISO)
    celda = tabla.cell(0, 0)
    celda.width = Cm(ANCHO_TEXTO * 0.92)
    sombrear(celda._tc.get_or_add_tcPr(), 'F8D4D4')
    p = celda.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.add_run(texto).bold = True
    doc.add_paragraph()


def campo(parrafo, instruccion, texto_resultado=''):
    def fld(tipo):
        r = OxmlElement('w:r')
        f = OxmlElement('w:fldChar')
        f.set(qn('w:fldCharType'), tipo)
        r.append(f)
        return r
    parrafo._p.append(fld('begin'))
    r = OxmlElement('w:r')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' {instruccion} '
    r.append(it)
    parrafo._p.append(r)
    parrafo._p.append(fld('separate'))
    if texto_resultado:
        parrafo.add_run(texto_resultado)
    parrafo._p.append(fld('end'))


def numeracion_reiniciada(doc):
    """Nuevo w:num del estilo 'List Number' que empieza en 1 (cada lista numerada del PDF reinicia)."""
    numid_estilo = doc.styles['List Number'].element.pPr.numPr.numId.val
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    abstract = next(n for n in numbering.findall(qn('w:num')) if n.numId == numid_estilo).abstractNumId.val
    nuevo = numbering.add_num(abstract)
    lvl = OxmlElement('w:lvlOverride')
    lvl.set(qn('w:ilvl'), '0')
    start = OxmlElement('w:startOverride')
    start.set(qn('w:val'), '1')
    lvl.append(start)
    nuevo.append(lvl)
    return nuevo.numId


def encabezado_y_paginas(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
    for lado in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
        setattr(sec, lado, Cm(2.5))
    sec.different_first_page_header_footer = True  # portada sin encabezado
    pg = OxmlElement('w:pgNumType')
    pg.set(qn('w:start'), '0')                      # portada = 0, índice = 1 (como en el PDF)
    sec._sectPr.append(pg)
    p = sec.header.paragraphs[0]
    # el estilo Encabezado trae tabulaciones (centro y derecha) que no coinciden con el ancho del texto: se anulan
    for tab in doc.styles['Header'].paragraph_format.tab_stops:
        p.paragraph_format.tab_stops.add_tab_stop(tab.position, WD_TAB_ALIGNMENT.CLEAR)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(ANCHO_TEXTO), WD_TAB_ALIGNMENT.RIGHT)
    run = p.add_run(ENCABEZADO + '\t')
    run.font.size = Pt(9.5)
    campo(p, 'PAGE', '1')
    for r in p.runs:
        r.font.size = Pt(9.5)
    ppr = p._p.get_or_add_pPr()
    borde(ppr, 'w:pBdr', {'bottom': ('single', 4)})


def actualizar_campos_al_abrir(doc):
    ajustes = doc.settings.element
    u = OxmlElement('w:updateFields')
    u.set(qn('w:val'), 'true')
    ajustes.append(u)


def construir_docx():
    tex = L.construir()[0]
    colores = {m.group(1): '%02X%02X%02X' % tuple(int(x) for x in m.group(2).split(','))
               for m in re.finditer(r'\\definecolor\{(\w+)\}\{RGB\}\{([\d,]+)\}', tex)}
    bls = bloques(tex)
    doc = Document()
    preparar_estilos(doc)
    encabezado_y_paginas(doc)
    actualizar_campos_al_abrir(doc)
    cuerpo = doc.element.body
    for p in list(cuerpo.iterchildren(qn('w:p'))):  # párrafo vacío inicial de la plantilla
        cuerpo.remove(p)

    # numeración de títulos como en el PDF (1., 2.1., A.)
    sec, apx, sub, entradas = 0, 0, 0, []
    for b in bls:
        if b[0] == 'titulo' and not b[2] and b[1] in ('section', 'subsection'):
            if b[1] == 'section':
                sub = 0
                if b[4]:                      # apéndice: A., B., …
                    apx += 1
                    actual = 'ABCDEFGHIJ'[apx - 1]
                else:
                    sec += 1
                    actual = str(sec)
                num = actual + '.'
            else:
                sub += 1
                num = f'{actual}.{sub}.'
            entradas.append((b, num))
    numero = {id(b): n for b, n in entradas}

    for b in bls:
        tipo = b[0]
        if tipo == 'portada':
            primera = True
            for tam, texto in b[1]:
                p = doc.add_paragraph(style='Title' if tam == 'LARGE' else None)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if primera:
                    p.paragraph_format.space_before = Pt(90)
                    p.add_run(texto)
                    primera = False
                else:
                    p.paragraph_format.space_before = Pt(150 if texto[:1].isdigit() else 36)
                    r = p.add_run(texto)
                    r.font.size = Pt(14)
            doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)
        elif tipo == 'indice':
            p = doc.add_paragraph('Índice', style='TOC Heading' if 'TOC Heading' in [s.name for s in doc.styles] else 'Heading 1')
            p = doc.add_paragraph()
            lineas = '\n'.join(('' if b2[1] == 'section' else '    ') + n + ' ' + b2[3] for b2, n in entradas)
            campo(p, 'TOC \\o "1-2" \\h \\z \\u', lineas)
        elif tipo == 'salto':
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif tipo == 'titulo':
            _, nivel, estrella, texto, _ = b
            estilo = {'section': 'Heading 1', 'subsection': 'Heading 3' if estrella else 'Heading 2',
                      'subsubsection': 'Heading 4', 'paragraph': 'Heading 5'}[nivel]
            p = doc.add_paragraph(style=estilo)
            p.add_run((numero[id(b)] + ' ' if id(b) in numero else '') + texto)
        elif tipo == 'parrafo':
            runs = recortar_espacios(b[1])
            if not texto_plano(runs):
                continue
            p = doc.add_paragraph()
            if b[2] == 'centro':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            agregar_runs(p, runs, colores)
        elif tipo == 'nota':
            p = doc.add_paragraph()
            agregar_runs(p, [{**r, 'i': True} for r in b[1] if r['t']], colores, 9)
        elif tipo == 'lista':
            numid = numeracion_reiniciada(doc) if b[1] else None
            for item in b[2]:
                p = doc.add_paragraph(style='List Number' if b[1] else 'List Bullet')
                if numid:
                    numpr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numpr.get_or_add_ilvl().val = 0
                    numpr.get_or_add_numId().val = numid
                agregar_runs(p, recortar_espacios(item), colores)
        elif tipo == 'tabla':
            escribir_tabla(doc, b[1], b[2], b[3], b[4], colores)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif tipo == 'aviso':
            escribir_aviso(doc, texto_plano(en_linea(b[1])))
    doc.save(SALIDA)
    return bls


def main():
    bls = construir_docx()
    print(f'Generado {SALIDA.relative_to(L.RAIZ)} ({sum(1 for b in bls if b[0] == "tabla")} tablas)')


if __name__ == '__main__':
    main()
