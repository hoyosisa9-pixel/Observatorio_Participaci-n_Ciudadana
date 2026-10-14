"""Convierte un Markdown sencillo (titulos, parrafos, listas, tablas, **negrita**, `codigo`) a .docx.

Uso (desde la raiz del repo):
  python3 scripts/md_a_docx.py outputs/disponibilidad/resumen_reunion.md [salida.docx]
Requiere python-docx (instalado en el site-packages del usuario: no usar -I). Si no se da salida, usa el mismo nombre con extension .docx.
Las tablas se escriben con estilo de cuadricula, encabezado sombreado y fuente de 9 pt.
"""
import re, sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RE_INLINE = re.compile(r'(\*\*.+?\*\*|`.+?`)')


def agregar_texto(parrafo, texto, tam=None, negrita=False):
    for trozo in RE_INLINE.split(texto):
        if not trozo:
            continue
        if trozo.startswith('**') and trozo.endswith('**'):
            run = parrafo.add_run(trozo[2:-2])
            run.bold = True
        elif trozo.startswith('`') and trozo.endswith('`'):
            run = parrafo.add_run(trozo[1:-1])
            run.font.name = 'Consolas'
            run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
        else:
            run = parrafo.add_run(trozo)
            run.bold = negrita or None
        if tam:
            run.font.size = Pt(tam)


def sombrear(celda, color='D9E2F3'):
    tc_pr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tc_pr.append(shd)


def escribir_tabla(doc, filas):
    celdas = [[c.strip() for c in f.strip().strip('|').split('|')] for f in filas]
    celdas = [f for f in celdas if not all(re.fullmatch(r':?-{3,}:?', c) for c in f)]
    ncol = max(len(f) for f in celdas)
    tabla = doc.add_table(rows=len(celdas), cols=ncol)
    tabla.style = 'Table Grid'
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, fila in enumerate(celdas):
        for j in range(ncol):
            celda = tabla.cell(i, j)
            p = celda.paragraphs[0]
            agregar_texto(p, fila[j] if j < len(fila) else '', tam=9, negrita=(i == 0))
            if i == 0:
                sombrear(celda)
    doc.add_paragraph()


def convertir(md, salida):
    doc = Document()
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(2)
        sec.top_margin = sec.bottom_margin = Cm(2)
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10.5)
    lineas = md.read_text(encoding='utf-8').splitlines()
    i = 0
    while i < len(lineas):
        l = lineas[i]
        if l.startswith('|'):
            bloque = []
            while i < len(lineas) and lineas[i].startswith('|'):
                bloque.append(lineas[i])
                i += 1
            escribir_tabla(doc, bloque)
            continue
        m = re.match(r'(#{1,4})\s+(.*)', l)
        if m:
            doc.add_heading(m.group(2).replace('**', ''), level=len(m.group(1)) - 1 if len(m.group(1)) > 1 else 0)
        elif re.match(r'\s*[-*]\s+', l):
            agregar_texto(doc.add_paragraph(style='List Bullet'), re.sub(r'^\s*[-*]\s+', '', l))
        elif re.match(r'\s*\d+\.\s+', l):
            # numero explicito: el estilo 'List Number' continuaria la numeracion entre listas distintas
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.first_line_indent = Cm(-0.6)
            num, resto = re.match(r'\s*(\d+\.)\s+(.*)', l).groups()
            agregar_texto(p, f'{num} {resto}')
        elif l.strip() and not re.fullmatch(r'-{3,}', l.strip()):
            agregar_texto(doc.add_paragraph(), l.strip())
        i += 1
    doc.save(salida)


if __name__ == '__main__':
    entrada = Path(sys.argv[1])
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 else entrada.with_suffix('.docx')
    convertir(entrada, salida)
    print(f'Generado {salida}')
