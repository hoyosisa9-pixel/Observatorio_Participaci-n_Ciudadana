"""Convierte inventario_IIC_relaciones.md a informe_IIC_overleaf.tex (y copia .txt)."""
import re, shutil, unicodedata
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
MD = RAIZ / 'inventario_IIC_relaciones.md'
TEX = RAIZ / 'informe_IIC_overleaf.tex'

ESC = {'\\': r'\textbackslash{}', '&': r'\&', '%': r'\%', '$': r'\$', '#': r'\#', '_': r'\_',
       '{': r'\{', '}': r'\}', '~': r'\textasciitilde{}', '^': r'\textasciicircum{}',
       '<': r'\textless{}', '>': r'\textgreater{}', '×': r'$\times$', '↔': r'$\leftrightarrow$',
       '…': r'\ldots{}', '"': "''", '├': '|', '└': '`', '─': '-'}

def esc(s):
    return ''.join(ESC.get(c, c) for c in s)

def esc_code(s):
    # texto monoespaciado con puntos de corte tras / _ - . para que ajuste en celdas
    out = esc(s)
    return re.sub(r'(/|\\_|-|\.)', r'\1\\allowbreak{}', out)

def esc_url(u):
    return u.replace('\\', '/').replace('%', r'\%').replace('#', r'\#')

def inline(s):
    """Convierte formato inline de markdown a LaTeX."""
    tokens = []
    def keep(tex):
        tokens.append(tex)
        return f'\x00{len(tokens)-1}\x00'
    def enlace(m):
        texto, destino = m.group(1), m.group(2)
        if not destino.startswith('http'):
            return keep(inline(texto))  # ruta relativa del repositorio: sin hipervínculo en el PDF
        if texto.startswith('http'):
            return keep(r'\url{' + esc_url(destino) + '}')
        return keep(r'\href{' + esc_url(destino) + '}{' + inline(texto) + '}')
    s = re.sub(r'\[([^\]]+)\]\(([^)\s]+)\)', enlace, s)
    s = re.sub(r'`([^`]+)`', lambda m: keep(r'\texttt{' + esc_code(m.group(1)) + '}'), s)
    s = re.sub(r'(?<![{(])(https?://[^\s|)]+)', lambda m: keep(r'\url{' + esc_url(m.group(1)) + '}'), s)
    s = re.sub(r'<br>', lambda m: keep(r'\newline '), s)
    s = esc(s)
    s = re.sub(r'\*\*(.+?)\*\*', r'\\textbf{\1}', s)
    s = re.sub(r'(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?!\w)', r'\\emph{\1}', s)
    s = re.sub(r'\x00(\d+)\x00', lambda m: tokens[int(m.group(1))], s)
    return s

# Anchos de columna por número de columnas (fracciones de \linewidth, sin contar separación)
ANCHOS = {
    3: [0.18, 0.42, 0.40],
    4: [0.26, 0.48, 0.13, 0.13],
    5: [0.19, 0.25, 0.25, 0.10, 0.21],
    6: [0.22, 0.26, 0.07, 0.10, 0.22, 0.13],
}
TAM = {3: r'\footnotesize', 4: r'\small', 5: r'\scriptsize', 6: r'\scriptsize'}

def split_row(line):
    cells = re.split(r'(?<!\\)\|', line.strip().strip('|'))
    return [c.strip().replace('\\|', '|') for c in cells]

def tabla(lineas):
    enc = split_row(lineas[0])
    filas = [split_row(l) for l in lineas[2:]]
    n = len(enc)
    alin = split_row(lineas[1])
    spec = ''.join(
        (r'>{\raggedleft\arraybackslash}' if a.endswith(':') else r'>{\raggedright\arraybackslash}')
        + f'p{{{w:.2f}\\dimexpr\\linewidth-{2*n}\\tabcolsep\\relax}}'
        for w, a in zip(ANCHOS[n], alin))
    head = ' & '.join(r'\textbf{' + inline(c) + '}' for c in enc) + r' \\'
    out = ['{' + TAM[n], r'\setlength{\tabcolsep}{3pt}', r'\begin{longtable}{' + spec + '}', r'\toprule', head, r'\midrule', r'\endfirsthead',
           r'\toprule', head, r'\midrule', r'\endhead', r'\midrule', r'\multicolumn{' + str(n) + r'}{r}{\emph{Continúa en la página siguiente}} \\', r'\endfoot', r'\bottomrule', r'\endlastfoot']
    for f in filas:
        out.append(' & '.join(inline(c) for c in f) + r' \\')
    out += [r'\end{longtable}', '}', '']
    return out

def convertir(md):
    L = md.splitlines()
    out, i = [], 0
    titulo = None
    while i < len(L):
        l = L[i]
        if l.startswith('# '):
            titulo = inline(l[2:]); out.append(r'\section*{' + titulo + '}'); i += 1
        elif l.startswith('### '):
            out.append(r'\subsubsection*{' + inline(l[4:]) + '}'); i += 1
        elif l.startswith('## '):
            out.append(r'\subsection*{' + inline(l[3:]) + '}'); i += 1
        elif l.startswith('```'):
            j = i + 1
            while not L[j].startswith('```'): j += 1
            out += [r'\begin{verbatim}'] + [x.replace('├──', '|--').replace('└──', '`--') for x in L[i+1:j]] + [r'\end{verbatim}']
            i = j + 1
        elif l.startswith('|'):
            j = i
            while j < len(L) and L[j].startswith('|'): j += 1
            out += tabla(L[i:j]); i = j
        elif l.startswith('> '):
            j = i
            bloque = []
            while j < len(L) and L[j].startswith('>'):
                bloque.append(L[j][2:].rstrip()); j += 1
            out.append(r'\begin{quote}')
            for b in bloque:
                if b.startswith('**Fórmula:**'):
                    out += [r'\textbf{Fórmula:}', r'\begin{equation*}',
                            r'IIC = \left( \frac{\text{Propuestas ejecutadas} \times \text{Factor de impacto}}{\text{Total propuestas válidas}} \right) \times 100',
                            r'\end{equation*}',
                            r'donde el \emph{Factor de impacto} es 1 = parcial, 2 = total y 3 = institucionalizada.\\']
                else:
                    out.append(inline(b) + r'\\')
            out.append(r'\end{quote}'); i = j
        elif l.startswith('`IIC ='):
            out += [r'\begin{equation*}',
                    r'IIC = \left( \frac{\text{Propuestas ejecutadas} \times \text{Factor de impacto}}{\text{Total propuestas válidas}} \right) \times 100',
                    r'\end{equation*}']; i += 1
        elif l.startswith('- '):
            out.append(r'\begin{itemize}[leftmargin=*]')
            while i < len(L) and L[i].startswith('- '):
                out.append(r'\item ' + inline(L[i][2:])); i += 1
            out.append(r'\end{itemize}')
        elif re.match(r'\d+\. ', l):
            out.append(r'\begin{enumerate}[leftmargin=*]')
            while i < len(L) and re.match(r'\d+\. ', L[i]):
                out.append(r'\item ' + inline(re.sub(r'^\d+\. ', '', L[i]))); i += 1
            out.append(r'\end{enumerate}')
        elif l.strip() == '---':
            out.append(r'\medskip\noindent\rule{\linewidth}{0.4pt}\medskip'); i += 1
        elif l.strip() == '':
            out.append(''); i += 1
        else:
            out.append(inline(l.rstrip())); i += 1
    return out

PREAMBULO = r'''\documentclass[11pt,a4paper]{article}
\usepackage[spanish]{babel}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[margin=2.5cm]{geometry}
\usepackage{array}
\usepackage{booktabs}
\usepackage{longtable}
\usepackage{tabularx}
\usepackage{amsmath}
\usepackage{amssymb}
\usepackage{enumitem}
\usepackage{xurl}
\usepackage[hidelinks]{hyperref}

% Evita que babel-spanish convierta < > y comillas en tablas
\spanishdeactivate{<>}
\setlength{\LTpre}{6pt}
\setlength{\LTpost}{6pt}
\setlength{\parindent}{0pt}
\setlength{\parskip}{4pt}
\renewcommand{\arraystretch}{1.15}

\title{Inventario y Matriz de Relaciones\\\large Índice de Incidencia Ciudadana (IIC) -- Observatorio Ciudadano de Envigado}
\date{25 de septiembre de 2026}

\begin{document}
'''

cuerpo = convertir(unicodedata.normalize('NFC', MD.read_text(encoding='utf-8')))
TEX.write_text(PREAMBULO + '\n'.join(cuerpo) + '\n\\end{document}\n', encoding='utf-8')
shutil.copyfile(TEX, TEX.with_suffix('.txt'))
print('Generado', TEX, 'y', TEX.with_suffix('.txt'))
