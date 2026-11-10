"""Genera el informe de avance (Avance 1) en Word con el mismo contenido que el PDF de LaTeX.

Uso (desde la raiz del repo):  python3 scripts/entregables/generar_informe_avance_docx.py
Fuente unica: scripts/entregables/generar_informe_avance_latex.py (construir) arma el .tex; este script lo traduce a
Word con las funciones de scripts/generar_informe_docx.py (estilos nativos Cambria 11, tablas tipo booktabs, encabezado con
numero de pagina, indice como campo TOC, aviso y colores tomados de los \\definecolor del .tex), sin modificarlas.
Salida: outputs/indicadores/entregables/informe_avance_1.docx
"""
import re, sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import generar_informe_avance_latex as AV  # noqa: E402
import generar_informe_docx as D  # noqa: E402

SALIDA = AV.SALIDA / 'informe_avance_1.docx'
D.ENCABEZADO = 'OCE – Avance 1: cálculo provisional de indicadores'


def construir_docx():
    tex = AV.construir()
    colores = {m.group(1): '%02X%02X%02X' % tuple(int(x) for x in m.group(2).split(','))
               for m in re.finditer(r'\\definecolor\{(\w+)\}\{RGB\}\{([\d,]+)\}', tex)}
    bls = D.bloques(tex)
    doc = Document()
    D.preparar_estilos(doc)
    D.encabezado_y_paginas(doc)
    D.actualizar_campos_al_abrir(doc)
    cuerpo = doc.element.body
    for p in list(cuerpo.iterchildren(qn('w:p'))):
        cuerpo.remove(p)
    sec, apx, sub, entradas = 0, 0, 0, []
    for b in bls:
        if b[0] == 'titulo' and not b[2] and b[1] in ('section', 'subsection'):
            if b[1] == 'section':
                sub = 0
                if b[4]:
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
                    p.paragraph_format.space_before = Pt(120 if texto.startswith('Fecha de corte') else
                                                         6 if texto.startswith('Fuentes consultadas') else 36)
                    r = p.add_run(texto)
                    r.font.size = Pt(14)
            doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE)
        elif tipo == 'indice':
            doc.add_paragraph('Índice', style='TOC Heading' if 'TOC Heading' in [s.name for s in doc.styles] else 'Heading 1')
            p = doc.add_paragraph()
            lineas = '\n'.join(('' if b2[1] == 'section' else '    ') + n + ' ' + b2[3] for b2, n in entradas)
            D.campo(p, 'TOC \\o "1-2" \\h \\z \\u', lineas)
        elif tipo == 'salto':
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif tipo == 'titulo':
            _, nivel, estrella, texto, _ = b
            estilo = {'section': 'Heading 1', 'subsection': 'Heading 3' if estrella else 'Heading 2',
                      'subsubsection': 'Heading 4', 'paragraph': 'Heading 5'}[nivel]
            p = doc.add_paragraph(style=estilo)
            p.add_run((numero[id(b)] + ' ' if id(b) in numero else '') + texto)
        elif tipo == 'parrafo':
            runs = D.recortar_espacios(b[1])
            if not D.texto_plano(runs):
                continue
            p = doc.add_paragraph()
            if b[2] == 'centro':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            D.agregar_runs(p, runs, colores)
        elif tipo == 'nota':
            p = doc.add_paragraph()
            D.agregar_runs(p, [{**r, 'i': True} for r in b[1] if r['t']], colores, 9)
        elif tipo == 'lista':
            numid = D.numeracion_reiniciada(doc) if b[1] else None
            for item in b[2]:
                p = doc.add_paragraph(style='List Number' if b[1] else 'List Bullet')
                if numid:
                    numpr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numpr.get_or_add_ilvl().val = 0
                    numpr.get_or_add_numId().val = numid
                D.agregar_runs(p, D.recortar_espacios(item), colores)
        elif tipo == 'tabla':
            D.escribir_tabla(doc, b[1], b[2], b[3], b[4], colores)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif tipo == 'aviso':
            D.escribir_aviso(doc, D.texto_plano(D.en_linea(b[1])))
    doc.save(SALIDA)
    return bls


def main():
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    bls = construir_docx()
    print(f'Generado {SALIDA.relative_to(AV.RAIZ)} ({sum(1 for b in bls if b[0] == "tabla")} tablas)')


if __name__ == '__main__':
    main()
