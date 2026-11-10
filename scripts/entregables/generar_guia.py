"""Genera la guía de estudio de uso interno (PDF, Word y .tex) y la hoja de reunión (PDF de una página).

Uso (desde la raiz del repo):  python3 scripts/entregables/generar_guia.py [--sin-pdf]
Entradas (solo se da formato; no se cambian textos ni cifras):
    outputs/guia/trabajo/guia_contenido.md, hoja_reunion.md, propuesta_variables.csv (control de coincidencia)
Salidas:  outputs/guia/guia_estudio_OCE.{tex,pdf,docx}  y  outputs/guia/hoja_reunion.pdf
Cadena:   Markdown -> .tex (estilo de los informes anteriores) -> PDF con tectonic -> Word con el traductor
          .tex -> .docx de scripts/generar_informe_docx.py (importado, sin modificarlo; las extensiones que hacen
          falta -- secciones apaisadas, bloque monoespaciado, recuadro, sublistas -- viven en este script).
tectonic: se busca en TECTONIC, en el scratchpad de la sesion o en el PATH; si no existe se descarga el binario musl.
"""
import csv, os, re, shutil, subprocess, sys, tarfile, urllib.request
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(RAIZ / 'scripts'))
import generar_informe_latex as L  # noqa: E402
import generar_informe_docx as D  # noqa: E402

TRABAJO = RAIZ / 'outputs' / 'guia' / 'trabajo'
SALIDA = RAIZ / 'outputs' / 'guia'
F_MD = TRABAJO / 'guia_contenido.md'
F_HOJA = TRABAJO / 'hoja_reunion.md'
F_VARS = TRABAJO / 'propuesta_variables.csv'
F_TEX, F_PDF, F_DOCX = (SALIDA / f'guia_estudio_OCE.{e}' for e in ('tex', 'pdf', 'docx'))
F_HOJA_PDF = SALIDA / 'hoja_reunion.pdf'
SCRATCH = Path('/tmp/claude-1000/-workspaces-Observatorio-Participaci-n-Ciudadana/8695b82b-4d33-43a8-8a3c-c302ba25f047/scratchpad')
TMP = SCRATCH / 'guia'
TECTONIC_URL = ('https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.15.0/'
                'tectonic-0.15.0-x86_64-unknown-linux-musl.tar.gz')

TITULO = 'Guía de estudio – uso interno'
SUBTITULO = ('Sistema de indicadores del Observatorio Ciudadano de Envigado: qué se ha hecho, problemas, variables '
             'y preparación de la reunión')
PREPARADA = 'Preparada para: Isabella Hoyos Álvarez'
FECHA = '10 de noviembre de 2026'
TITULO_CORTO = 'OCE -- Guía de estudio (uso interno)'
ENCABEZADO_WORD = 'OCE – Guía de estudio (uso interno)'

ANCHO_V, ANCHO_H = 16.59, 22.9  # cm de texto en vertical / apaisado (carta, margen 2,5 cm)
SEMAFORO = {'Verde': 'cverde', 'Amarillo': 'camarillo', 'Naranja': 'cnaranja', 'Rojo': 'crojo', 'Gris': 'cgris'}
CATEGORIA = {'A': 'ccata', 'B': 'ccatb', 'C': 'ccatc', 'D': 'ccatd'}
APAISADO = {'3.2 Tabla de variables', 'Esquema simple de tablas y relaciones'}
COLORES_TEX = r'''\definecolor{cverde}{RGB}{214,240,214}
\definecolor{camarillo}{RGB}{255,243,191}
\definecolor{cnaranja}{RGB}{255,222,184}
\definecolor{crojo}{RGB}{248,212,212}
\definecolor{cgris}{RGB}{225,225,225}
\definecolor{chead}{RGB}{214,226,240}
\definecolor{cstripe}{RGB}{243,246,250}
\definecolor{cnota}{RGB}{238,243,250}
\definecolor{ccata}{RGB}{214,232,247}
\definecolor{ccatb}{RGB}{212,238,226}
\definecolor{ccatc}{RGB}{255,236,196}
\definecolor{ccatd}{RGB}{246,214,222}'''


# =========================================================================== Markdown -> bloques
def parsear_md(texto):
    """Bloques: ('h', nivel, texto) | ('p', texto) | ('lista', numerada, [(texto, [sub])]) | ('cita', texto)
    | ('codigo', [lineas]) | ('tabla', [filas]) | ('hr',)."""
    L_ = texto.split('\n')
    bl, i = [], 0
    es_li = lambda l: re.match(r'^(\s*)([-*]|\d+\.) +(.*)$', l)
    while i < len(L_):
        l = L_[i]
        if not l.strip():
            i += 1
        elif l.startswith('```'):
            j = i + 1
            while not L_[j].startswith('```'):
                j += 1
            bl.append(('codigo', L_[i + 1:j]))
            i = j + 1
        elif l.strip() == '---':
            bl.append(('hr',))
            i += 1
        elif l.startswith('#'):
            m = re.match(r'^(#+)\s+(.*)$', l)
            bl.append(('h', len(m.group(1)), m.group(2).strip()))
            i += 1
        elif l.startswith('|'):
            j = i
            while j < len(L_) and L_[j].startswith('|'):
                j += 1
            filas = [[c.strip() for c in f.strip().strip('|').split('|')] for f in L_[i:j]]
            bl.append(('tabla', [filas[0]] + filas[2:]))
            i = j
        elif l.startswith('>'):
            j = i
            while j < len(L_) and L_[j].startswith('>'):
                j += 1
            bl.append(('cita', ' '.join(x.lstrip('> ').strip() for x in L_[i:j])))
            i = j
        elif es_li(l) and not es_li(l).group(1):
            numerada = es_li(l).group(2).endswith('.')
            items = []
            j = i
            while j < len(L_):
                m = es_li(L_[j])
                if m and not m.group(1):
                    items.append((m.group(3).strip(), []))
                elif m and m.group(1) and items:
                    items[-1][1].append(m.group(3).strip())
                elif not L_[j].strip():
                    # lineas en blanco entre elementos de la misma lista numerada
                    k = j
                    while k < len(L_) and not L_[k].strip():
                        k += 1
                    n = es_li(L_[k]) if k < len(L_) else None
                    if not (n and not n.group(1) and n.group(2).endswith('.') == numerada):
                        break
                    j = k - 1
                else:
                    break
                j += 1
            bl.append(('lista', numerada, items))
            i = j
        else:
            j = i
            while j < len(L_) and L_[j].strip() and not re.match(r'^(#|\||>|```|---|\s*[-*] |\s*\d+\. )', L_[j]):
                bl.append(('p', L_[j].strip()))
                j += 1
            i = j
    return bl


# =========================================================================== texto en línea
def comillas(t):
    return re.sub(r'"([^"]+)"', '“\\1”', t)


def en_linea(t, tabla=False):
    """Markdown en línea -> LaTeX: `código`, **negrita**, *cursiva*, URLs clicables, comillas tipográficas."""
    t = comillas(t)
    partes = re.split(r'(`[^`]+`|\*\*.+?\*\*|(?<![*\w])\*[^*\s](?:[^*]*[^*\s])?\*(?![*\w])|https?://[^\s)]+)', t)
    out = []
    for p in partes:
        if not p:
            continue
        if p.startswith('`'):
            c = p[1:-1]
            out.append(r'\texttt{' + L.cortable(c) + '}' if tabla else L.ruta(c))
        elif p.startswith('**'):
            out.append(r'\textbf{' + en_linea(p[2:-2], tabla) + '}')
        elif p.startswith('*') and len(p) > 2:
            out.append(r'\emph{' + en_linea(p[1:-1], tabla) + '}')
        elif p.startswith('http'):
            out.append(L.enlace(p))
        else:
            out.append(L.esc(p))
    return ''.join(out)


# =========================================================================== tablas
def anchos(filas, total, tam):
    """Anchos (cm) de columna: las cortas a su tamaño, el resto proporcional al largo medio del texto."""
    n = len(filas[0])
    cw = 0.165 if tam == 'footnotesize' else 0.145
    util = total - n * 0.22
    cuerpo = filas[1:] or filas
    ancho_min, peso = [], []
    for k in range(n):
        col = [re.sub(r'[*`]', '', r[k]) for r in cuerpo]
        enc = re.sub(r'[*`]', '', filas[0][k])
        palabra = max([len(w) for c in col + [enc] for w in c.split()] + [1])
        maxl = max([len(c) for c in col] + [len(enc)])
        medio = sum(len(c) for c in col) / len(col)
        palabra_enc = max([len(w) for w in enc.split()] + [1])
        ancho_min.append(max(min(palabra, 20) * cw * 1.1 + 0.4, palabra_enc * cw * 1.3 + 0.4))
        peso.append(max(medio, 4) ** 0.9)
        if maxl * cw + 0.4 <= 2.4:                      # columna corta: cabe entera en una línea
            ancho_min[-1] = max(ancho_min[-1], maxl * cw * 1.1 + 0.4)
            peso[-1] = 0
    w = list(ancho_min)
    libre = util - sum(w)
    flex = [k for k in range(n) if peso[k] > 0]
    if libre > 0 and flex:
        s = sum(peso[k] for k in flex)
        for k in flex:
            w[k] += libre * peso[k] / s
    elif libre < 0:
        f = util / sum(w)
        w = [x * f for x in w]
    return w


def celda_tex(texto, encabezado_col, fila_par):
    t = texto.strip()
    plano = re.sub(r'\*', '', t)
    relleno = None
    if encabezado_col.startswith('Semáforo') and plano in SEMAFORO:
        relleno = SEMAFORO[plano]
    elif encabezado_col.startswith('Cat') and plano in CATEGORIA:
        relleno = CATEGORIA[plano]
    elif fila_par:
        relleno = 'cstripe'
    return (r'\cellcolor{' + relleno + '}' if relleno else '') + en_linea(t, tabla=True)


def tabla_tex(filas, ancho_total):
    n = len(filas[0])
    tam = 'scriptsize' if n >= 7 else 'footnotesize'
    w = anchos(filas, ancho_total, tam)
    spec = ''.join(r'L{%.2fcm}' % x for x in w)
    cab = [r'\cellcolor{chead}\textbf{' + en_linea(c, True) + '}' for c in filas[0]]
    rayado = len(filas) > 7
    cuerpo = [[celda_tex(c, filas[0][k], rayado and i % 2 == 1) for k, c in enumerate(f)] for i, f in enumerate(filas[1:])]
    lin = [r'{\%s' % tam, r'\begin{xltabular}{\linewidth}{' + spec + '}', r'\toprule', ' & '.join(cab) + r' \\',
           r'\midrule', r'\endfirsthead', r'\toprule', ' & '.join(cab) + r' \\', r'\midrule', r'\endhead',
           r'\midrule \multicolumn{' + str(n) + r'}{r}{\emph{Continúa en la página siguiente}} \\', r'\endfoot',
           r'\bottomrule', r'\endlastfoot']
    lin += [' & '.join(f) + r' \\' for f in cuerpo]
    lin += [r'\end{xltabular}', '}']
    return '\n'.join(lin), (9 if tam == 'footnotesize' else 8)


# =========================================================================== bloques -> LaTeX
def a_latex(bl, word=False, hoja=False):
    """word=True genera la variante que entiende el traductor a Word (marcadores @@..@@ en vez de entornos)."""
    T, tamanos, codigos = [], [], []
    apaisado_hasta = None   # nivel de encabezado que cierra el bloque apaisado
    for b in bl:
        k = b[0]
        if k == 'h':
            _, nivel, txt = b
            if apaisado_hasta is not None and nivel <= apaisado_hasta:
                T.append('@@LAND_OFF@@' if word else r'\clearpage' + '\n' + r'\restoregeometry' + '\n' + r'\setlength{\headwidth}{\textwidth}' + '\n' + r'\renewcommand{\pagsize}{\special{papersize=21.59cm,27.94cm}}')
                apaisado_hasta = None
            if txt in APAISADO:
                T.append('@@LAND_ON@@' if word else r'\clearpage' + '\n' + r'\newgeometry{left=2.5cm,width=22.94cm,top=2.5cm,height=16.5cm,headheight=14pt,headsep=0.8cm}' + '\n' + r'\setlength{\headwidth}{\textwidth}' + '\n' + r'\renewcommand{\pagsize}{\special{papersize=27.94cm,21.59cm}}')
                apaisado_hasta = 2
            cmd = {1: 'section', 2: 'subsection', 3: 'subsubsection'}.get(nivel, 'paragraph')
            cuerpo_t = en_linea(txt)
            if hoja:
                T.append(r'\begin{center}{\Large\bfseries ' + cuerpo_t + r'\par}\end{center}')
                continue
            if nivel == 1 and not word:
                T.append(r'\clearpage')
            T.append('\\' + cmd + '*{' + cuerpo_t + '}')
            if not word and nivel <= 2:
                T.append(r'\phantomsection\addcontentsline{toc}{%s}{%s}' % (cmd, cuerpo_t))
        elif k == 'p':
            txt = b[1]
            if not word and re.fullmatch(r'\*\*[^*]+\*\*', txt):
                T.append(r'\needspace{4\baselineskip}')
            T.append(en_linea(txt))
        elif k == 'cita':
            if word:
                T.append('@@CALLOUT@@ ' + en_linea(b[1]))
            else:
                T.append(r'\begin{center}\fcolorbox{gray!60}{cnota}{\parbox{\dimexpr\linewidth-2\fboxsep-2\fboxrule\relax}'
                         r'{\small ' + en_linea(b[1]) + r'}}\end{center}')
        elif k == 'lista':
            ent = 'enumerate' if b[1] else 'itemize'
            lin = [r'\begin{%s}' % ent]
            for txt, subs in b[2]:
                lin.append(r'\item ' + en_linea(txt))
                if subs and word:
                    lin += [r'\item @@SUB@@' + en_linea(s) for s in subs]
                elif subs:
                    lin += [r'\begin{itemize}'] + [r'\item ' + en_linea(s) for s in subs] + [r'\end{itemize}']
            lin.append(r'\end{%s}' % ent)
            T.append('\n'.join(lin))
        elif k == 'codigo':
            if word:
                T.append(f'@@CODE{len(codigos)}@@')
                codigos.append(b[1])
            else:
                T.append('\\begin{Verbatim}[fontsize=\\footnotesize,frame=single,rulecolor=\\color{gray},'
                         'framesep=2mm]\n' + '\n'.join(b[1]) + '\n\\end{Verbatim}')
        elif k == 'tabla':
            ancho = ANCHO_H if apaisado_hasta else ANCHO_V
            if hoja:
                ancho = ANCHO_HOJA
            if not hoja and b[1][0][0] == 'ID':
                T.append(LEYENDA_TEX)
            t, tam = tabla_tex(b[1], ancho)
            T.append(t)
            tamanos.append(tam)
        elif k == 'hr':
            pass
    if apaisado_hasta is not None:
        T.append('@@LAND_OFF@@' if word else r'\clearpage' + '\n' + r'\restoregeometry' + '\n' + r'\setlength{\headwidth}{\textwidth}' + '\n' + r'\renewcommand{\pagsize}{\special{papersize=21.59cm,27.94cm}}')
    return T, tamanos, codigos


ANCHO_HOJA = 18.5
LEYENDA = [('A', 'Lista para la base de datos'), ('B', 'Lista si los profesores deciden algo'),
           ('C', 'Depende de una solicitud a la Alcaldía'), ('D', 'Propuesta de ajustar, reemplazar o eliminar')]
LEYENDA_TEX = (r'\par\noindent{\footnotesize\textbf{Leyenda de la columna Cat.:} ' + '\\quad '.join(
    r'\colorbox{%s}{~\textbf{%s}~} %s' % (CATEGORIA[c], c, L.esc(t)) for c, t in LEYENDA) + r'\par}\smallskip')


# =========================================================================== documento LaTeX
PREAMBULO = r'''\documentclass[11pt,letterpaper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage[spanish,es-noshorthands,es-tabla]{babel}
\usepackage[margin=MARGEN,headheight=14pt]{geometry}
\usepackage[table]{xcolor}
\usepackage{booktabs}
\usepackage{xltabular}
\usepackage{array}
\usepackage{enumitem}
\usepackage{fancyhdr}
\usepackage{fancyvrb}
\usepackage{needspace}
\usepackage[hidelinks,breaklinks=true]{hyperref}
\usepackage{url}

COLORES
\newcolumntype{L}[1]{>{\raggedright\arraybackslash}p{#1}}
\renewcommand{\arraystretch}{1.18}
\setlength{\tabcolsep}{3pt}
\setlength{\emergencystretch}{3em}
\setlength{\parindent}{0pt}
\newcommand{\pagsize}{\special{papersize=21.59cm,27.94cm}}
\setlength{\parskip}{0.5em}
\setlist{itemsep=2pt,topsep=4pt}
'''


def documento_guia(T):
    cab = PREAMBULO.replace('MARGEN', '2.5cm').replace('COLORES', COLORES_TEX)
    cab += r'''
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\pagsize\small ''' + L.esc(TITULO_CORTO) + r'''}
\fancyhead[R]{\small\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\fancypagestyle{plain}{\fancyhf{}\fancyhead[R]{\small\thepage}}
\hypersetup{pdftitle={''' + L.esc(TITULO) + r'''},pdfauthor={Isabella Hoyos \'Alvarez}}

\begin{document}

\hypersetup{pageanchor=false}
\begin{titlepage}
\centering
\vspace*{3.5cm}
{\LARGE\bfseries ''' + L.esc(TITULO) + r'''\par}
\vspace{1.5cm}
{\large ''' + L.esc(SUBTITULO) + r'''\par}
\vspace{3cm}
{\large ''' + L.esc(PREPARADA) + r'''\par}
\vspace{0.5cm}
{\large ''' + L.esc(FECHA) + r'''\par}
\vfill
{\large Observatorio Ciudadano de Envigado (OCE)\par}
\end{titlepage}
\hypersetup{pageanchor=true}

\tableofcontents
\clearpage

'''
    return cab + '\n\n'.join(T) + '\n\n\\end{document}\n'


def documento_hoja(T):
    cab = PREAMBULO.replace('MARGEN', '1.4cm').replace('COLORES', COLORES_TEX)
    return (cab + '\n\\pagestyle{empty}\n\\renewcommand{\\arraystretch}{1.08}\n\\setlength{\\parskip}{0.3em}\n'
            '\\hypersetup{pdftitle={Hoja de reunión}}\n\\begin{document}\n\\small\n'
            + '\n\n'.join(T) + '\n\n\\end{document}\n')


# =========================================================================== tectonic
def buscar_tectonic():
    cand = [os.environ.get('TECTONIC'), str(SCRATCH / 'tectonic_avance' / 'tectonic'), shutil.which('tectonic')]
    for c in cand:
        if c and Path(c).exists():
            return c
    dest = SCRATCH / 'tectonic_guia'
    dest.mkdir(parents=True, exist_ok=True)
    tgz = dest / 't.tar.gz'
    urllib.request.urlretrieve(TECTONIC_URL, tgz)
    with tarfile.open(tgz) as t:
        t.extractall(dest)
    (dest / 'tectonic').chmod(0o755)
    return str(dest / 'tectonic')


def compilar(tex_path, destino_pdf):
    tec = buscar_tectonic()
    TMP.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([tec, '--keep-logs', '--outdir', str(TMP), str(tex_path)], capture_output=True, text=True)
    (TMP / (tex_path.stem + '.consola.txt')).write_text(r.stdout + r.stderr, encoding='utf-8')
    if r.returncode:
        raise SystemExit('tectonic falló:\n' + (r.stdout + r.stderr)[-3000:])
    shutil.copy(TMP / (tex_path.stem + '.pdf'), destino_pdf)
    return r.stdout + r.stderr


# =========================================================================== Word
def encabezado_seccion(doc, sec, ancho):
    sec.header.is_linked_to_previous = False
    p = sec.header.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for tab in doc.styles['Header'].paragraph_format.tab_stops:
        p.paragraph_format.tab_stops.add_tab_stop(tab.position, WD_TAB_ALIGNMENT.CLEAR)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(ancho), WD_TAB_ALIGNMENT.RIGHT)
    run = p.add_run(D.ENCABEZADO + '\t')
    run.font.size = Pt(9.5)
    D.campo(p, 'PAGE', '1')
    for r in p.runs:
        r.font.size = Pt(9.5)
    D.borde(p._p.get_or_add_pPr(), 'w:pBdr', {'bottom': ('single', 4)})


def nueva_seccion(doc, apaisada):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    ancho, alto = Cm(21.59), Cm(27.94)
    sec.orientation = WD_ORIENT.LANDSCAPE if apaisada else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = (alto, ancho) if apaisada else (ancho, alto)
    sec.different_first_page_header_footer = False
    pg = sec._sectPr.find(qn('w:pgNumType'))
    if pg is not None:
        sec._sectPr.remove(pg)
    encabezado_seccion(doc, sec, ANCHO_H if apaisada else ANCHO_V)
    return sec


def parrafo_sombreado(doc, runs, colores, color='EEF3FA', borde_izq=True):
    p = doc.add_paragraph()
    ppr = p._p.get_or_add_pPr()
    D.borde(ppr, 'w:pBdr', {'top': ('single', 4), 'left': ('single', 4), 'bottom': ('single', 4), 'right': ('single', 4)})
    for e in ppr.find(qn('w:pBdr')):
        e.set(qn('w:color'), '999999')
        e.set(qn('w:space'), '4')
    D.sombrear(ppr, color)
    p.paragraph_format.left_indent = Cm(0.2)
    p.paragraph_format.right_indent = Cm(0.2)
    D.agregar_runs(p, runs, colores, 10)
    return p


def escribir_codigo(doc, lineas):
    for n, l in enumerate(lineas):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        pf.line_spacing = 1.0
        D.sombrear(p._p.get_or_add_pPr(), 'F4F4F4')
        r = p.add_run(l.replace(' ', '\u00a0') if l.strip() else '\u00a0')
        r.font.name = D.MONO
        r._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:hAnsi'), D.MONO)
        r.font.size = Pt(8)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def construir_docx(tex_word, tamanos, codigos):
    colores = {m.group(1): '%02X%02X%02X' % tuple(int(x) for x in m.group(2).split(','))
               for m in re.finditer(r'\\definecolor\{(\w+)\}\{RGB\}\{([\d,]+)\}', tex_word)}
    bls = D.bloques(tex_word)
    doc = Document()
    D.preparar_estilos(doc)
    for nombre in ('List Bullet 2',):
        D.fijar_fuente(doc.styles[nombre], D.FUENTE)
    D.ENCABEZADO = ENCABEZADO_WORD
    D.encabezado_y_paginas(doc)
    D.actualizar_campos_al_abrir(doc)
    cuerpo = doc.element.body
    for p in list(cuerpo.iterchildren(qn('w:p'))):
        cuerpo.remove(p)
    entradas = [b for b in bls if b[0] == 'titulo' and b[1] in ('section', 'subsection')]
    tablas = iter(tamanos)
    for b in bls:
        tipo = b[0]
        if tipo == 'portada':
            textos = [TITULO, SUBTITULO, PREPARADA, FECHA]
            for n, texto in enumerate(textos):
                p = doc.add_paragraph(style='Title' if n == 0 else None)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(110 if n == 0 else 30 if n == 1 else 120 if n == 2 else 8)
                r = p.add_run(texto)
                if n:
                    r.font.size = Pt(14)
            p = doc.add_paragraph('Observatorio Ciudadano de Envigado (OCE)')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(110)
            p.runs[0].font.size = Pt(14)
            p.runs[0].add_break(WD_BREAK.PAGE)
        elif tipo == 'indice':
            doc.add_paragraph('Índice', style='TOC Heading' if 'TOC Heading' in [s.name for s in doc.styles] else 'Heading 1')
            p = doc.add_paragraph()
            lineas = '\n'.join(('' if e[1] == 'section' else '    ') + e[3] for e in entradas)
            D.campo(p, 'TOC \\o "1-2" \\h \\z \\u', lineas)
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif tipo == 'salto':
            continue                              # el salto tras el índice ya se escribió con el campo
        elif tipo == 'titulo':
            _, nivel, _, texto, _ = b
            estilo = {'section': 'Heading 1', 'subsection': 'Heading 2', 'subsubsection': 'Heading 3',
                      'paragraph': 'Heading 4'}[nivel]
            p = doc.add_paragraph(style=estilo)
            if nivel == 'section':
                p.paragraph_format.page_break_before = True
            p.add_run(texto)
        elif tipo == 'parrafo':
            runs = D.recortar_espacios(b[1])
            plano = D.texto_plano(runs)
            if not plano:
                continue
            if plano == '@@LAND_ON@@':
                D.ANCHO_TEXTO = ANCHO_H
                nueva_seccion(doc, True)
            elif plano == '@@LAND_OFF@@':
                D.ANCHO_TEXTO = ANCHO_V
                nueva_seccion(doc, False)
            elif plano.startswith('@@CODE'):
                escribir_codigo(doc, codigos[int(re.match(r'@@CODE(\d+)@@', plano).group(1))])
            elif plano.startswith('@@CALLOUT@@'):
                runs[0]['t'] = runs[0]['t'].replace('@@CALLOUT@@', '').lstrip()
                parrafo_sombreado(doc, [r for r in runs if r['t'] or r.get('br')], colores)
            else:
                p = doc.add_paragraph()
                if b[2] == 'centro':
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                D.agregar_runs(p, runs, colores)
        elif tipo == 'lista':
            numid = D.numeracion_reiniciada(doc) if b[1] else None
            for item in b[2]:
                runs = D.recortar_espacios(item)
                sub = bool(runs) and runs[0]['t'].startswith('@@SUB@@')
                if sub:
                    runs[0]['t'] = runs[0]['t'].replace('@@SUB@@', '')
                p = doc.add_paragraph(style='List Bullet 2' if sub else ('List Number' if b[1] else 'List Bullet'))
                if numid and not sub:
                    numpr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numpr.get_or_add_ilvl().val = 0
                    numpr.get_or_add_numId().val = numid
                p.paragraph_format.space_after = Pt(3)
                D.agregar_runs(p, runs, colores)
        elif tipo == 'tabla':
            D.escribir_tabla(doc, b[1], b[2], b[3], next(tablas), colores)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        else:
            raise ValueError(f'bloque no previsto: {tipo}')
    # la sección apaisada final vacía (si la hubiera) no existe: cada LAND_ON tiene su LAND_OFF
    doc.save(F_DOCX)
    return bls


# =========================================================================== principal
def construir_tex_guia(word):
    bl = parsear_md(F_MD.read_text(encoding='utf-8'))
    # La tabla de variables debe coincidir con el CSV de la propuesta (control de integridad)
    tab = next(b[1] for b in bl if b[0] == 'tabla' and b[1][0][0] == 'ID')
    with open(F_VARS, encoding='utf-8-sig', newline='') as f:
        csvf = list(csv.reader(f, delimiter=';'))
    assert len(csvf) - 1 == len(tab) - 1 == 41, (len(csvf), len(tab))
    assert [r[0] for r in csvf[1:]] == [r[0] for r in tab[1:]], 'IDs de la tabla distintos del CSV'
    assert [r[3] for r in csvf[1:]] == [re.sub(r'\*', '', r[3]) for r in tab[1:]], 'categorías distintas del CSV'
    # los bloques iniciales (título y metadatos) se separan: el título va en la portada
    assert bl[0][0] == 'h' and bl[0][1] == 1
    bl = bl[1:]
    T, tam, cod = a_latex(bl, word=word)
    return T, tam, cod


def main():
    TMP.mkdir(parents=True, exist_ok=True)
    SALIDA.mkdir(parents=True, exist_ok=True)
    T, _, _ = construir_tex_guia(False)
    tex = documento_guia(T)
    raros = sorted({ch for ch in tex if ord(ch) > 127 and ch not in L.LETRAS_OK})
    assert not raros, f'caracteres sin traducir: {raros}'
    F_TEX.write_text(tex, encoding='utf-8')
    # hoja de reunion (una pagina)
    bh = parsear_md(F_HOJA.read_text(encoding='utf-8'))
    Th, _, _ = a_latex(bh, hoja=True)
    f_hoja_tex = TMP / 'hoja_reunion.tex'
    f_hoja_tex.write_text(documento_hoja(Th), encoding='utf-8')
    if '--sin-pdf' not in sys.argv:
        compilar(F_TEX, F_PDF)
        compilar(f_hoja_tex, F_HOJA_PDF)
    # Word: misma estructura con marcadores
    Tw, tam, cod = construir_tex_guia(True)
    cab = documento_guia(Tw)
    construir_docx(cab, tam, cod)
    print('Generados:', *(p.relative_to(RAIZ) for p in (F_TEX, F_PDF, F_DOCX, F_HOJA_PDF)))


if __name__ == '__main__':
    main()
