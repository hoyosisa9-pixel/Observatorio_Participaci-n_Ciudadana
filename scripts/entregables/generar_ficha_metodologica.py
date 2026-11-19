"""Ficha metodológica de los indicadores y los índices del OCE (Word, entrega al profe). Provisional.

Una ficha por indicador principal (los que entran a los índices) y una por índice, en tablas de dos columnas (campo | contenido).
Salida: outputs/entrega_profe/ficha_metodologica_indicadores_OCE.docx
Uso (desde la raíz, sin -I):  python3 scripts/entregables/generar_ficha_metodologica.py
Todo el contenido numérico se lee de outputs/construccion/*/*_indicadores.csv (vía contenido_fichas.py): al actualizar los CSV
basta con volver a correr este script.
"""
import re
from pathlib import Path

from docx import Document

import comun_entregas as ce
import contenido_fichas as cf
import generar_entrega_profe as gep

RAIZ = Path(__file__).resolve().parents[2]
SALIDA = RAIZ / 'outputs' / 'entrega_profe' / 'ficha_metodologica_indicadores_OCE.docx'
ANCHOS = [4.2, 13.3]  # cm; ancho útil vertical = 21,59 - 4,0 = 17,59


def construir(fichas):
    doc = gep.nuevo_documento(horizontal=False)
    doc.add_paragraph(style='Title').add_run('Ficha metodológica de los indicadores y los índices')
    gep.parrafo(doc, f'Observatorio Ciudadano de Envigado · {ce.FECHA} · Resultados **provisionales**')
    gep.parrafo(doc, 'Este documento reúne la metodología de los indicadores principales y de los cuatro índices del observatorio: '
                     'cumplimiento del Plan de Desarrollo (ICPD-7), participación ciudadana (ISPC), presupuesto participativo (IREPP) y control social (ICRCS). '
                     'Cada ficha parte de la definición de la matriz original y la ajusta a lo construido. '
                     'Los resultados por año y la explicación de cada cambio están en el documento «Detalle de la revisión de indicadores».')
    gep.parrafo(doc, 'Cómo leer cada ficha: nombre y código, objetivo y definición; fórmula y cálculo del puntaje de 0 a 100; variables con su unidad y fuente; '
                     'unidad, sentido, periodicidad y cobertura; línea base y referente con su justificación; peso y semáforo (propuesta); '
                     'serie disponible y brecha de información; tratamiento de datos faltantes, rangos y casos especiales; limitaciones; fuentes; fecha y estado. '
                     'Las definiciones comunes (brecha, rango, mismo universo, población, años y semáforo) están en el anexo.')
    doc.add_heading('Contenido', level=1)
    filas = [[f['codigo'], f['nombre'], f['tipo']] for f in fichas]
    gep.tabla(doc, [3.0, 11.0, 3.2], ['Código', 'Nombre', 'Tipo'], filas, tam=9)
    for dim in ce.DIM_ORDEN:
        doc.add_heading(ce.NOMBRE_DIM[dim], level=1)
        for f in [x for x in fichas if ce.dim_de(x['codigo']) == dim]:
            h = doc.add_heading(f"{f['codigo']} · {f['nombre']}", level=2)
            h.paragraph_format.page_break_before = True
            cuerpo = []
            for clave, rotulo in cf.CAMPOS:
                if clave in ('pesos', 'regla', 'anios') and f[clave] == '—':
                    continue
                cuerpo.append([f'**{rotulo}**', f[clave]])
            gep.tabla(doc, ANCHOS, ['Campo', 'Contenido'], cuerpo, tam=9)
    h = doc.add_heading('Anexo. Definiciones comunes', level=1)
    h.paragraph_format.page_break_before = True
    gep.tabla(doc, ANCHOS, ['Término', 'Definición'], [[f'**{t}**', d] for t, d in cf.definiciones_comunes()], tam=9)
    doc.save(SALIDA)
    return SALIDA


def verificar(ruta, fichas):
    d = Document(ruta)
    sec = d.sections[0]
    util = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
    for i, t in enumerate(d.tables):
        ancho = sum(c.width.cm for c in t.rows[0].cells)
        assert ancho <= util + 0.01, ('ancho de tabla', i, round(ancho, 2), round(util, 2))
    texto = gep.texto_completo(d)
    assert not ce.palabras_prohibidas(texto), ce.palabras_prohibidas(texto)
    for extra in ('Isabella', 'Francisco', 'usuaria', 'Analizador'):
        assert extra not in texto, extra
    assert not ce.ano_suelto(texto), ce.ano_suelto(texto)[:3]   # ninguna palabra «ano» suelta (debe ser «año»)
    assert not re.search(r'\bD\d{1,3}\b', texto), re.findall(r'.{20}\bD\d{1,3}\b.{10}', texto)[:3]
    # todo enlace de las fichas existe en los CSV de construcción (o es la raíz de un sitio que ellos citan)
    en_csv = ' '.join(r['enlace'] for dim in ce.DIM_ORDEN for r in ce.DATA[dim])
    for u in set(re.findall(r'https?://[^\s;)\]]+', texto)):
        assert u.rstrip('.,') in en_csv, ('enlace que no está en los CSV', u)
    # una ficha por indicador principal y por índice, contadas contra los CSV
    n_princ = sum(len(ce.principales(dim)) for dim in ce.DIM_ORDEN)
    n_ind = len(ce.NOMBRE_INDICE)
    fichas_word = [p.text.split(' · ')[0] for p in d.paragraphs if p.style.name == 'Heading 2']
    esperados = [c for dim in ce.DIM_ORDEN for c in ce.principales(dim) + [ce.INDICE_DIM[dim]]]
    assert fichas_word == esperados, (fichas_word, esperados)
    assert len(fichas_word) == n_princ + n_ind == len(fichas)
    # índices: los valores de la ficha coinciden con el CSV
    for cod in ce.NOMBRE_INDICE:
        f = next(x for x in fichas if x['codigo'] == cod)
        for a, (v, s) in ce.valores_indice(cod).items():
            if v is not None:
                assert f'{a}: {ce.fmt(v)} ({s})' in f['serie'], (cod, a)
                assert f'{a}: {ce.fmt(v)} ({s})' in texto.replace('\n', ' '), (cod, a)
    esperados_v = {'PP.IREPP': (2025, '95,83'), 'PDM.ICPD7': (2019, '62,41'), 'PC.ISPC': (2025, '61,62'), 'CS.ICRCS': (2024, '60,59')}
    for cod, (a, v) in esperados_v.items():
        assert f'{a}: {v}' in texto, (cod, a, v)
    assert 'CS.ICRCS' and '2025: 38,18' in texto
    # cada ficha trae todos los campos
    for t in d.tables[1:-1]:
        rot = [r.cells[0].text for r in t.rows[1:]]
        for clave, rotulo in cf.CAMPOS:
            if clave in ('pesos', 'regla', 'anios'):
                continue
            assert rotulo in rot, (rotulo, rot[:3])
    return d, len(fichas_word), len(d.tables)


def main():
    fichas = cf.construir_fichas()
    ruta = construir(fichas)
    d, n, nt = verificar(ruta, fichas)
    print(ruta.relative_to(RAIZ), '| fichas:', n, '(indicadores principales:', n - len(ce.NOMBRE_INDICE), '+ índices:', len(ce.NOMBRE_INDICE), ') | tablas:', nt)
    print('  secciones:', [p.text for p in d.paragraphs if p.style.name.startswith('Heading 1')])
    print('Asserts OK: una ficha por indicador principal e índice (contadas contra los CSV); valores de los índices = CSV '
          '(IREPP 2025 = 95,83; ICPD-7 2019 = 62,41; ISPC 2025 = 61,62; ICRCS 2024 = 60,59 y 2025 = 38,18); '
          'sin «Origen», «Claude» ni «asistente»; anchos dentro del margen; todos los campos en cada ficha.')


if __name__ == '__main__':
    main()
