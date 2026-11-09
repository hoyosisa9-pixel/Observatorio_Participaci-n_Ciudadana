"""Genera el informe LaTeX (Overleaf) de disponibilidad y brecha del sistema de indicadores del OCE.

Uso (desde la raiz del repo):  python3 -I scripts/generar_informe_latex.py
Solo stdlib + openpyxl. Todos los numeros salen de los datos; si cambian, basta con volver a correr el script.

Entradas (outputs/disponibilidad/):
  disponibilidad_brecha_indicadores.xlsx   hojas Variables, Indicadores, Fuentes, Solicitudes
  epp_pp_2025_zonas.csv                    EPP 2025 por zona y municipal (scripts/epp_pp_2025.py)
  documentos_por_descargar.md              limitaciones de acceso (tabla)
  ESTADO.md                                lista "Inconsistencias de la matriz"
  trabajo/revision_fase3.md                seccion 4, decisiones para los profesores
Salidas:
  informe_latex/informe_disponibilidad.tex
  informe_disponibilidad_overleaf.zip      (para Overleaf -> New Project -> Upload Project)

Compatible con pdfLaTeX (por defecto en Overleaf): el script falla si queda algun caracter fuera de Latin-1
sin traducir a una orden LaTeX.
"""
import csv, re, zipfile
from collections import Counter, defaultdict
from pathlib import Path

import sys

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generar_disponibilidad as G  # noqa: E402  (misma limpieza de lenguaje y tildes que el Excel)

RAIZ = Path(__file__).resolve().parent.parent
DISP = RAIZ / 'outputs' / 'disponibilidad'
F_XLSX = DISP / 'disponibilidad_brecha_indicadores.xlsx'
F_ZONAS = DISP / 'epp_pp_2025_zonas.csv'
F_DOCS = DISP / 'interno' / 'documentos_por_descargar.md'
F_ESTADO = DISP / 'ESTADO.md'
F_REVISION = DISP / 'trabajo' / 'revision_fase3.md'
SALIDA = DISP / 'informe_latex'
F_TEX = SALIDA / 'informe_disponibilidad.tex'
F_ZIP = DISP / 'informe_disponibilidad_overleaf.zip'

AUTORA = 'Isabella Hoyos Álvarez'
TITULO = ('Sistema de indicadores del Observatorio Ciudadano de Envigado: '
          'disponibilidad y brecha de información')
TITULO_CORTO = 'OCE -- Disponibilidad y brecha de información'
CATEGORIAS = ['Disponible', 'Construible', 'Parcial', 'Proxy', 'No disponible']
COLOR_CAT = {'Disponible': 'cverde', 'Construible': 'camarillo', 'Parcial': 'camarillo',
             'Proxy': 'cnaranja', 'No disponible': 'crojo'}
VEREDICTOS = ['Viable', 'Viable con ajuste', 'No viable por ahora']
COLOR_VER = {'Viable': 'cverde', 'Viable con ajuste': 'camarillo', 'No viable por ahora': 'crojo'}
SISTEMAS = [('CS', 'Control social'), ('PC', 'Participación ciudadana'),
            ('PDM', 'Cumplimiento del PDM'), ('PP', 'Presupuesto participativo')]
# Versión y fechas de la portada: parámetros fijos (no dependen del día en que se ejecute el script)
VERSION = 'Versión 1.1 – 9 de noviembre de 2026'
FECHA_CORTE = '14 de octubre de 2026'

# --------------------------------------------------------------------------- escape LaTeX
ESPECIALES = {'\\': r'\textbackslash{}', '&': r'\&', '%': r'\%', '$': r'\$', '#': r'\#', '_': r'\_',
              '{': r'\{', '}': r'\}', '~': r'\textasciitilde{}', '^': r'\textasciicircum{}',
              '<': r'\textless{}', '>': r'\textgreater{}', '|': r'\textbar{}'}
UNICODE = {'≈': r'$\approx$', '≤': r'$\leq$', '≥': r'$\geq$', '→': r'$\rightarrow$', '←': r'$\leftarrow$',
           '↔': r'$\leftrightarrow$', '×': r'$\times$', '–': '--', '—': '---', '…': r'\ldots{}',
           '“': "``", '”': "''", '‘': '`', '’': "'", '•': r'\textbullet{}', '−': '-', ' ': '~',
           '✓': 'sí', '≠': r'$\neq$',
           # signos que se escriben con ordenes LaTeX para no depender de la codificacion de la fuente
           '¿': r'\textquestiondown{}', '¡': r'\textexclamdown{}', '²': r'\textsuperscript{2}',
           '³': r'\textsuperscript{3}', 'º': r'\textsuperscript{o}', 'ª': r'\textsuperscript{a}',
           '°': r'\textsuperscript{o}', '§': r'\S{}', '·': r'\textperiodcentered{}', '«': r'\guillemotleft{}',
           '»': r'\guillemotright{}', '€': 'EUR', '´': "'"}
# letras permitidas tal cual (utf8 + T1 en pdfLaTeX)
LETRAS_OK = set('áéíóúÁÉÍÓÚñÑüÜàèìòùÀÈÌÒÙçÇ')
ARCHIVO_LOCAL = 'Archivo local del Observatorio'
NOTA_RECORTE = r'\par\smallskip{\footnotesize\emph{Texto resumido; el detalle completo está en el Excel adjunto (hoja Variables).}}\par'
PAGINAS_PDM = ('Las páginas correctas, en el PDM 2024-2027 aprobado (y entre paréntesis en el documento técnico de '
               'soporte de diciembre de 2025), son: aseguramiento subsidiado p. 602 (640); mortalidad vial p. 706 (738); '
               'espacio público verde p. 717 (748); procesos participativos p. 742 (775); servicios deportivos p. 609 '
               '(646); atención integral a la primera infancia p. 650 (683); Saber 11 p. 652 (685).')


def texto_md(texto):
    """Texto de los documentos de trabajo (Markdown) con la misma limpieza que el Excel y sin referencias a archivos."""
    t = re.sub(r'Las correctas están en `?notas_PDM\.md`?\.?', PAGINAS_PDM, texto)
    t = re.sub(r'\(?`?[\w.-]+\.(?:md|csv|py)`?\)?', 'el Excel adjunto', t)
    t = re.sub(r'\s*Detalle completo en el Excel adjunto §[\d.]+\.?', '', t)
    return G.limpiar(t)


def esc(texto):
    """Escapa caracteres especiales de LaTeX y traduce Unicode fuera de Latin-1."""
    if texto is None:
        return ''
    out = []
    for ch in str(texto):
        if ch in ESPECIALES:
            out.append(ESPECIALES[ch])
        elif ch in UNICODE:
            out.append(UNICODE[ch])
        else:
            out.append(ch)
    return ''.join(out)


def md_inline(texto):
    """Escapa y convierte **negrita** y `codigo` de Markdown."""
    partes = re.split(r'(\*\*.+?\*\*|`.+?`)', texto)
    res = []
    for p in partes:
        if p.startswith('**') and p.endswith('**'):
            res.append(r'\textbf{' + esc(p[2:-2]) + '}')
        elif p.startswith('`') and p.endswith('`'):
            res.append(ruta(p[1:-1]))
        else:
            res.append(esc(p))
    return ''.join(res)


def cortable(texto):
    """Texto escapado con puntos de quiebre tras / . _ - para rutas y URLs largas."""
    s = esc(texto)
    s = re.sub(r'(/|\\_|\.|-|\\&|=|\?)', r'\1\\allowbreak{}', s)
    # nombres de archivo largos sin separadores: un punto de quiebre cada 12 caracteres
    return re.sub(r'[A-Za-z0-9]{16,}', lambda m: r'\allowbreak{}'.join(
        m.group(0)[i:i + 12] for i in range(0, len(m.group(0)), 12)), s)


def esc_q(texto):
    """Escapa y permite cortar tras '/' (p. ej. 'PDF/DOCX')."""
    return esc(texto).replace('/', r'/\allowbreak{}').replace('.', r'.\allowbreak{}')


def ruta(texto):
    return r'\texttt{\footnotesize ' + cortable(texto) + '}'


def enlace(url):
    """URL clicable y divisible en lineas: cada tramo entre '/' es un \\href al mismo destino
    (en XeTeX/dvipdfmx un solo \\href no se puede partir entre lineas)."""
    destino = (url.replace('\\', '/').replace('%', r'\%').replace('#', r'\#').replace('&', r'\&')
               .replace('~', r'\~'))
    tramos = re.findall(r'[^/]*/|[^/]+$', url)
    return (r'\texttt{\footnotesize ' + r'\allowbreak{}'.join(
        r'\href{' + destino + '}{' + cortable(t) + '}' for t in tramos if t) + '}')


def ruta_o_enlaces(texto):
    """Rutas del repo o URLs separadas por ';' o ' | '."""
    piezas = [p.strip() for p in re.split(r'\s*;\s*|\s+\|\s+|\n', texto or '') if p.strip()]
    res = []
    for p in piezas:
        m = re.match(r'(https?://\S+?)([\s(].*)?$', p)
        if m:
            res.append(enlace(m.group(1)) + (' ' + esc(m.group(2).strip()) if m.group(2) else ''))
        elif p == G.SIN_URL or 'data_raw_IIC' in p:
            res.append(esc(ARCHIVO_LOCAL))
        elif p.startswith('(+'):
            res.append(esc(p.replace('en la hoja Enlaces', 'en la hoja Enlaces del Excel adjunto')))
        else:
            # texto con rutas embebidas: cada palabra con '/' se compone como ruta divisible
            res.append(' '.join(ruta(w) if '/' in w and len(w) > 15 else esc_q(w) for w in p.split(' ')))
    return r'; \newline '.join(res)


def corto(texto, n):
    t = ' '.join(str(texto or '').split())
    if len(t) <= n:
        return t
    return t[:n].rsplit(' ', 1)[0].rstrip(',;:.') + '…'


def num(x, dec=0):
    s = f'{float(x):,.{dec}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def pct(x):
    return num(x, 2) + r'~\%'


def celda_color(valor, colores):
    return r'\cellcolor{' + colores.get(valor, 'white') + '}' + esc(valor)


# --------------------------------------------------------------------------- lectura de datos
# encabezados del Excel para los profesores -> claves internas de este script
ALIAS = {'ID de la variable': 'id_variable', 'Nombre': 'nombre', 'Fuente encontrada': 'fuente encontrada',
         'Brecha': 'brecha', 'Disponibilidad': 'disponibilidad', 'Sistema': 'sistema', 'Código': 'codigo',
         'Tipo': 'tipo', 'Veredicto': 'veredicto', 'Ajuste propuesto': 'ajuste_propuesto', 'Emisor': 'emisor',
         'Fuente': 'fuente', 'Formato': 'formato', 'Enlace (URL de origen)': 'url_o_ruta',
         'Dependencia': 'dependencia', 'Dato a solicitar': 'dato_a_solicitar', 'Periodo': 'periodo'}


def hoja(wb, nombre):
    filas = list(wb[nombre].iter_rows(values_only=True))
    cab = [ALIAS.get(str(c), str(c)) for c in filas[0]]
    return [dict(zip(cab, ['' if v is None else str(v) for v in f])) for f in filas[1:] if any(f)]


def lista_md(texto, encabezado, numerada):
    """Items de la lista que sigue a un encabezado de Markdown, hasta el siguiente encabezado."""
    lineas = texto.splitlines()
    i = next(k for k, l in enumerate(lineas) if l.strip().startswith(encabezado))
    items = []
    for l in lineas[i + 1:]:
        if l.startswith('#'):
            break
        m = re.match(r'\s*\d+\.\s+(.*)', l) if numerada else re.match(r'\s*-\s+(.*)', l)
        if m:
            items.append(m.group(1).strip())
    return items


def tabla_md(texto):
    filas = [l for l in texto.splitlines() if l.startswith('|')]
    filas = [[c.strip() for c in f.strip().strip('|').split('|')] for f in filas]
    return [f for f in filas[1:] if not all(re.fullmatch(r':?-{3,}:?', c) for c in f)]


# --------------------------------------------------------------------------- bloques LaTeX
def xltabular(columnas, cabecera, filas, tam=r'\footnotesize'):
    lin = [r'{' + tam, r'\begin{xltabular}{\linewidth}{' + columnas + '}', r'\toprule',
           ' & '.join(r'\textbf{' + c + '}' for c in cabecera) + r' \\', r'\midrule', r'\endfirsthead',
           r'\toprule', ' & '.join(r'\textbf{' + c + '}' for c in cabecera) + r' \\', r'\midrule', r'\endhead',
           r'\midrule \multicolumn{' + str(len(cabecera)) + r'}{r}{\emph{Continúa en la página siguiente}} \\',
           r'\endfoot', r'\bottomrule', r'\endlastfoot']
    lin += [' & '.join(f) + r' \\' for f in filas]
    lin += [r'\end{xltabular}', '}']
    return '\n'.join(lin)


def aviso(texto):
    return (r'\begin{center}\fcolorbox{red!60!black}{crojo}{\parbox{0.92\linewidth}{\centering\bfseries '
            + texto + r'}}\end{center}')


def construir():
    wb = openpyxl.load_workbook(F_XLSX, read_only=True)
    variables = hoja(wb, 'Variables')
    indicadores = hoja(wb, 'Indicadores')
    fuentes = hoja(wb, 'Fuentes')
    solicitudes = hoja(wb, 'Solicitudes')
    with open(F_ZONAS, encoding='utf-8', newline='') as f:
        zonas = list(csv.DictReader(f))
    estado = F_ESTADO.read_text(encoding='utf-8')
    revision = F_REVISION.read_text(encoding='utf-8')
    docs = tabla_md(F_DOCS.read_text(encoding='utf-8'))
    assert len(variables) == 39 and len(indicadores) == 23, (len(variables), len(indicadores))
    for v in variables:
        v['sistema'] = v['id_variable'].split('.')[0]
    for i in indicadores:
        i['sistema'] = i['codigo'].split('.')[0]

    conteo = Counter(v['disponibilidad'] for v in variables)
    conteo_sis = {s: Counter(v['disponibilidad'] for v in variables if v['sistema'] == s) for s, _ in SISTEMAS}
    veredictos = Counter(i['veredicto'] for i in indicadores)
    metas = [v for v in variables if v['disponibilidad'] == 'Disponible' and v['sistema'] == 'PDM'
             and re.search(r'\.M[AS]?$', v['id_variable'])]
    observados = [v for v in variables if v['disponibilidad'] == 'Disponible' and v not in metas]
    indices = [i for i in indicadores if i['tipo'] == 'Índice']
    indices_no = [i for i in indices if i['veredicto'] == 'No viable por ahora']
    muni = next(z for z in zonas if z['nivel'] == 'Municipio')
    filas_z = [z for z in zonas if z['nivel'] != 'Municipio']
    zmin = min(filas_z, key=lambda z: float(z['EPP']))
    zmax = max(filas_z, key=lambda z: float(z['EPP']))
    epp = float(muni['EPP'])

    T = []
    a = T.append
    # ---------------------------------------------------------------- preambulo
    a(r'''\documentclass[11pt,letterpaper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage[spanish,es-noshorthands,es-tabla]{babel}
\usepackage[margin=2.5cm,headheight=14pt]{geometry}
\usepackage[table]{xcolor}
\usepackage{booktabs}
\usepackage{xltabular}
\usepackage{array}
\usepackage{enumitem}
\usepackage{fancyhdr}
\usepackage[hidelinks,breaklinks=true]{hyperref}
\usepackage{url}

\definecolor{cverde}{RGB}{214,240,214}
\definecolor{camarillo}{RGB}{255,243,191}
\definecolor{cnaranja}{RGB}{255,222,184}
\definecolor{crojo}{RGB}{248,212,212}
\newcolumntype{L}[1]{>{\raggedright\arraybackslash}p{#1}}
\newcolumntype{Y}{>{\raggedright\arraybackslash}X}
\renewcommand{\arraystretch}{1.15}
\setlength{\emergencystretch}{3em}
\setlist{itemsep=2pt,topsep=4pt}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small ''' + esc(TITULO_CORTO) + r'''}
\fancyhead[R]{\small\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\fancypagestyle{plain}{\fancyhf{}\fancyhead[R]{\small\thepage}}

\hypersetup{pdftitle={''' + esc(TITULO) + r'''},pdfauthor={''' + esc(AUTORA) + r'''}}

\begin{document}
''')
    # ---------------------------------------------------------------- portada
    a(r'''\hypersetup{pageanchor=false}
\begin{titlepage}
\centering
\vspace*{3cm}
{\LARGE\bfseries ''' + esc(TITULO) + r'''\par}
\vspace{1.5cm}
{\large Documento de trabajo preliminar\par}
\vspace{2cm}
{\large ''' + esc(AUTORA) + r'''\par}
\vspace{0.5cm}
{\large Observatorio Ciudadano de Envigado (OCE)\par}
\vspace{2cm}
{\large Dirigido a: profesores Francisco Correa y Eber Gutiérrez\par}
\vfill
{\large ''' + esc(VERSION) + r'''\par}
\vspace{0.3cm}
{\large Fecha de corte de la información: ''' + esc(FECHA_CORTE) + r'''\par}
\end{titlepage}
\hypersetup{pageanchor=true}

\tableofcontents
\clearpage
''')
    # ---------------------------------------------------------------- 1. resumen ejecutivo
    filas_c = []
    for s, nombre in SISTEMAS:
        c = conteo_sis[s]
        filas_c.append([esc(nombre)] + [str(c.get(k, 0)) for k in CATEGORIAS] + [str(sum(c.values()))])
    filas_c.append([r'\textbf{Total}'] + [r'\textbf{' + str(conteo.get(k, 0)) + '}' for k in CATEGORIAS]
                   + [r'\textbf{' + str(len(variables)) + '}'])
    cab_c = ['Sistema'] + [r'\cellcolor{' + COLOR_CAT[k] + '}' + k for k in CATEGORIAS] + ['Total']
    a(r'\section{Resumen ejecutivo}')
    a(f'Se revisaron {len(variables)} variables y {len(indicadores)} indicadores e índices de los cuatro sistemas '
      r'de la matriz, para 2024, 2025 y lo publicado de 2026. Ningún valor está validado: todas las cifras '
      r'tienen estado \emph{Pendiente}.')
    a(r'\begin{center}\small\begin{tabular}{l' + 'c' * (len(CATEGORIAS) + 1) + r'}\toprule')
    a(' & '.join(r'\textbf{' + c + '}' for c in cab_c) + r' \\ \midrule')
    a('\n'.join(' & '.join(f) + r' \\' for f in filas_c[:-1]) + r' \midrule')
    a(' & '.join(filas_c[-1]) + r' \\ \bottomrule\end{tabular}\end{center}')
    a(r'\noindent\textbf{Veredicto de los indicadores e índices:}'
      + r'\par\smallskip\centerline{' + r'\quad '.join(r'\colorbox{' + COLOR_VER[v] + '}{' + esc(v) + ': ' + str(veredictos.get(v, 0)) + '}'
                  for v in VEREDICTOS) + '}')
    a(r'\paragraph{Hallazgos principales}')
    a(r'\begin{enumerate}')
    a(r'\item \textbf{Hay pocos valores observados publicados.} De los ' + str(conteo['Disponible'])
      + r' datos clasificados como Disponible, ' + str(len(metas)) + r' son metas del PDM a 2027 y solo '
      + str(len(observados)) + ' son valores observados (' + ', '.join(esc(v['id_variable']) for v in observados)
      + r'). La mayor parte de las variables es Parcial (' + str(conteo['Parcial']) + ') o Construible ('
      + str(conteo['Construible']) + ').')
    a(r'\item \textbf{La ejecución presupuestal del presupuesto participativo 2025 es reconstruible:} '
      + pct(epp) + r' (apropiación \$' + num(muni['APP']) + r', obligaciones \$' + num(muni['OPP'])
      + r'), cruzando el plan de acción con el listado de la EFPC. Es una cifra \textbf{preliminar, por validar}; '
      r'entregas (CEP) y metas de atención (CMA) siguen sin soporte.')
    a(r'\item \textbf{Los índices quedan bloqueados.} ' + str(len(indices_no)) + ' de los ' + str(len(indices))
      + ' índices no son calculables por ahora (' + ', '.join(esc(i['codigo']) for i in indices_no)
      + r'), porque alguna de sus entradas no tiene dato y la Guía no redistribuye pesos. En participación '
      r'ciudadana falta un inventario oficial de procesos participativos.')
    a(r'\end{enumerate}')
    # ---------------------------------------------------------------- 2. metodologia
    a(r'''\section{Metodología}
\subsection{Insumos}
\begin{itemize}
\item \textbf{Matriz de los profesores} (\emph{Matriz de información de indicadores del observatorio ciudadano}, hojas Guía, Matriz, Registro, PP\_Proyectos y Resultados): ''' + str(len(variables)) + ' variables y ' + str(len(indicadores)) + r''' indicadores e índices en cuatro sistemas. No se modificó; se trabajó sobre copias.
\item \textbf{Bases recolectadas por el Observatorio}: 2.787 archivos de portales públicos (Alcaldía, Concejo, Contraloría, EFPC, datos.gov.co, DNP), catalogados en un inventario con la URL de origen de cada uno.
\item \textbf{Fuentes externas} nombradas por la matriz o halladas en la búsqueda: MinSalud, Medicina Legal (Forensis), ANSV, ICFES, DANE, TerriData, Área Metropolitana del Valle de Aburrá, entre otras (catálogo completo en el Anexo~\ref{anexo:fuentes}).
\end{itemize}
Periodo objetivo: 2024, 2025 y lo publicado de 2026. Territorio: Municipio de Envigado, con 13 zonas cuando hay desagregación.

\subsection{Escala de disponibilidad}
\begin{center}\small
\begin{tabularx}{\linewidth}{L{3cm}Y}
\toprule \textbf{Categoría} & \textbf{Definición} \\ \midrule
\cellcolor{cverde}Disponible & Publicado, con la unidad y la definición del indicador, para al menos un año entre 2024 y 2026. \\
\cellcolor{camarillo}Construible & No está publicado como cifra, pero se puede contar o calcular con documentos que ya se tienen. \\
\cellcolor{camarillo}Parcial & Existe, pero faltan años, la desagregación o el soporte exigido, o la definición difiere de la matriz. \\
\cellcolor{cnaranja}Proxy & Solo hay un dato parecido, con otra definición, unidad o universo. \\
\cellcolor{crojo}No disponible & No está en el archivo del Observatorio ni en fuentes públicas; hay que pedirlo a la Alcaldía. \\
\bottomrule
\end{tabularx}
\end{center}

\subsection{Reglas}
\begin{itemize}
\item Cada hallazgo lleva evidencia: ruta del archivo o URL, más página, tabla u hoja. Lo que no se pudo abrir se marca ``sin verificar''.
\item Ningún valor se marca como validado: esa decisión corresponde al responsable de la matriz. Los valores encontrados quedan en estado \emph{Pendiente}.
\item Ninguna variable se dejó como Proxy o No disponible sin buscar antes en las fuentes primarias.
\item Las clasificaciones pasaron por una revisión crítica que unificó criterios entre los cuatro sistemas.
\end{itemize}
''')
    # ---------------------------------------------------------------- 3. resultados por sistema
    a(r'\section{Resultados por sistema}')
    for s, nombre in SISTEMAS:
        vs = [v for v in variables if v['sistema'] == s]
        ins = [i for i in indicadores if i['sistema'] == s]
        c = conteo_sis[s]
        a(r'\subsection{' + esc(nombre) + ' (' + s + ')}')
        a(f'{len(vs)} variables: ' + ', '.join(f'{c[k]} {k.lower()}' for k in CATEGORIAS if c.get(k))
          + f'. {len(ins)} indicadores e índices.')
        a(r'\subsubsection*{Variables}')
        a(xltabular(r'L{1.9cm}YL{1.9cm}YY', ['ID', 'Variable', 'Categoría', 'Fuente', 'Brecha'],
                    [[esc(v['id_variable']), esc(corto(v['nombre'], 120)),
                      celda_color(v['disponibilidad'], COLOR_CAT), esc(corto(v['fuente encontrada'], 110)),
                      esc(corto(v['brecha'], 140))] for v in vs], tam=r'\scriptsize'))
        if any(len(' '.join(v[k].split())) > n for v in vs for k, n in (('nombre', 120), ('fuente encontrada', 110),
                                                                      ('brecha', 140))):
            a(NOTA_RECORTE)
        a(r'\subsubsection*{Indicadores e índices}')
        a(xltabular(r'L{1.9cm}L{3.4cm}L{2.1cm}Y', ['Código', 'Indicador', 'Veredicto', 'Ajuste propuesto'],
                    [[esc(i['codigo']), esc(i['nombre']), celda_color(i['veredicto'], COLOR_VER),
                      esc(i['ajuste_propuesto'])] for i in ins], tam=r'\scriptsize'))
    # ---------------------------------------------------------------- 4. EPP
    a(r'\section{Ejecución presupuestal del presupuesto participativo 2025}')
    a(aviso('PRELIMINAR, POR VALIDAR'))
    a(r'''\subsection{Método}
Ninguna fuente pública marca los recursos del presupuesto participativo (PP). Por eso la lista de proyectos se reconstruye:
\begin{enumerate}
\item Del plan de acción del IV trimestre de 2025 (Departamento Administrativo de Planeación, corte 7 de enero de 2026) se toman los proyectos con ``ZONA \emph{n}'' en el nombre, con su apropiación definitiva (APP) y su obligación (OPP).
\item Se excluye el proyecto 0553, \emph{Implementación del presupuesto participativo}, que es el costo del proceso y no un proyecto priorizado.
\item Se contrasta con el listado de proyectos PP 2025 de la Escuela de Formación para la Participación Ciudadana (EFPC): misma cantidad de proyectos por zona y valores que coinciden al peso, salvo los que tuvieron adiciones.
\item $\mathrm{EPP} = 100 \times \mathrm{OPP} / \mathrm{APP}$, por zona y para el consolidado municipal.
\end{enumerate}
''')
    filas_e = [[esc(z['nivel']), z['proyectos_plan'], z['proyectos_efpc'], num(z['APP']),
                num(z['OPP']), pct(z['EPP'])] for z in filas_z]
    a(r'\subsection{Resultados por zona}')
    a(r'\begin{center}\footnotesize\begin{tabular}{lrrrrr}\toprule')
    a(r'\textbf{Nivel} & \textbf{Proy. plan} & \textbf{Proy. EFPC} & \textbf{APP (\$)} & \textbf{OPP (\$)} & \textbf{EPP} \\ \midrule')
    a('\n'.join(' & '.join(f) + r' \\' for f in filas_e) + r' \midrule')
    a(r'\textbf{Municipio} & \textbf{' + muni['proyectos_plan'] + r'} & \textbf{' + muni['proyectos_efpc']
      + r'} & \textbf{' + num(muni['APP']) + r'} & \textbf{' + num(muni['OPP']) + r'} & \textbf{'
      + pct(muni['EPP']) + r'} \\ \bottomrule\end{tabular}\end{center}')
    a(r'Total municipal: EPP 2025 de \textbf{' + pct(epp) + '}. Por zona va de ' + pct(zmin['EPP']) + ' ('
      + esc(zmin['nivel'].lower()) + ') a ' + pct(zmax['EPP']) + ' (' + esc(zmax['nivel'].lower()) + ').')
    a(r'\paragraph{Pendientes antes de usar la cifra}')
    a(r'''\begin{enumerate}
\item Planeación debe confirmar la lista de proyectos PP.
\item Hay que decidir si se incluye el proyecto 0553 (costo del proceso).
\item Las cifras de 2024 y 2026 son aproximadas: 2024 se reconstruye desde el informe de Hacienda con correspondencias parciales, y 2026 no tiene cierre anual (obligaciones al 30 de junio).
\end{enumerate}''')
    # ---------------------------------------------------------------- 5. inconsistencias
    a(r'\section{Inconsistencias encontradas en la matriz}')
    a(r'\begin{itemize}')
    for it in lista_md(estado, '### Inconsistencias de la matriz', numerada=False):
        a(r'\item ' + md_inline(texto_md(it)))
    a(r'\end{itemize}')
    # ---------------------------------------------------------------- 6. decisiones
    a(r'\section{Decisiones que se necesitan de los profesores}')
    a(r'\begin{enumerate}')
    for it in lista_md(revision, '## 4. Preguntas abiertas', numerada=True):
        a(r'\item ' + md_inline(texto_md(it)))
    a(r'\end{enumerate}')
    # ---------------------------------------------------------------- 7. solicitudes
    a(r'\section{Solicitudes de información a la Alcaldía}')
    a(r'\begingroup\sloppy')
    por_dep = defaultdict(list)
    for r in solicitudes:
        por_dep[r['dependencia']].append(r)
    a(f'{len(solicitudes)} solicitudes, agrupadas en {len(por_dep)} dependencias.')
    for dep in sorted(por_dep):
        a(r'\subsection*{' + esc(dep) + '}')
        a(r'\begin{itemize}')
        for r in por_dep[dep]:
            a(r'\item \textbf{' + esc(r['id_variable']) + ':} ' + esc(r['dato_a_solicitar'])
              + (' (' + esc(r['periodo']) + ')' if r['periodo'] else ''))
        a(r'\end{itemize}')
    a(r'\endgroup')
    # ---------------------------------------------------------------- 8. limitaciones
    a(r'\section{Limitaciones y próximos pasos}')
    a(r'\subsection{Documentos que no se pudieron consultar o leer}')
    a(xltabular(r'L{3.6cm}YL{3cm}L{3cm}', ['Documento', 'Enlace', 'Variable que cambiaría', 'Motivo'],
                [[md_inline(texto_md(d[1])),
                  ruta_o_enlaces(' ; '.join(re.findall(r'https?://[^`\s;)]+', d[2])) or texto_md(d[2])),
                  md_inline(texto_md(d[3])), md_inline(texto_md(d[4]))] for d in docs], tam=r'\scriptsize'))
    a(r'''\subsection{Próximos pasos}
\begin{enumerate}
\item Tomar las decisiones de la sección anterior, en especial los universos de conteo.
\item Enviar las solicitudes de información a la Alcaldía.
\item Descargar manualmente los documentos de la tabla anterior y reclasificar las variables afectadas.
\item Validar con Planeación la lista de proyectos del presupuesto participativo.
\item Actualizar la matriz de disponibilidad y este informe con la información nueva.
\end{enumerate}''')
    # ---------------------------------------------------------------- anexo A
    a(r'\appendix')
    a(r'\section{Catálogo de fuentes}\label{anexo:fuentes}')
    a(f'{len(fuentes)} fuentes consultadas.')
    a(xltabular(r'L{2.8cm}YL{1.6cm}Y', ['Emisor', 'Documento', 'Formato', 'Enlace o ruta'],
                [[esc_q(f['emisor']), esc_q(f['fuente']), esc_q(f['formato']), ruta_o_enlaces(f['url_o_ruta'])]
                 for f in fuentes], tam=r'\scriptsize'))
    a(r'\end{document}')
    tex = '\n\n'.join(T) + '\n'
    raros = sorted({ch for ch in tex if ord(ch) > 127 and ch not in LETRAS_OK})
    assert not raros, f'caracteres no ASCII sin traducir a LaTeX: {raros}'
    return tex, conteo, veredictos, epp


def main():
    tex, conteo, veredictos, epp = construir()
    SALIDA.mkdir(parents=True, exist_ok=True)
    F_TEX.write_text(tex, encoding='utf-8')
    with zipfile.ZipFile(F_ZIP, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(SALIDA.iterdir()):
            if f.suffix in ('.tex', '.bib', '.png', '.pdf') and f.name != 'informe_disponibilidad.pdf':
                z.write(f, f.name)
    print('Conteo:', dict(conteo))
    print('Veredictos:', dict(veredictos))
    print(f'EPP 2025: {epp:.2f}')
    print(f'Generado {F_TEX.relative_to(RAIZ)} y {F_ZIP.relative_to(RAIZ)}')


if __name__ == '__main__':
    main()
