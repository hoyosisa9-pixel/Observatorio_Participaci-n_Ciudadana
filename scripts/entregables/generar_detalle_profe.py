"""Detalle de la revisión de indicadores para el profe (Word vertical). Provisional.

Por cada indicador de la matriz original (incluidos los eliminados y los reemplazados) y por los indicadores nuevos: cómo era en la matriz,
qué pasó (se mantiene / se ajusta / se reemplaza / se elimina / nuevo), por qué cambió, cómo quedó, la serie por año, la brecha y las alertas.
Salida: outputs/entrega_profe/detalle_indicadores_OCE_profe.docx
Uso (desde la raíz, sin -I):  python3 scripts/entregables/generar_detalle_profe.py
Las cifras salen de outputs/construccion/*/*_indicadores.csv; los textos de método, de contenido_fichas.py.
No lleva atribuciones de las decisiones a personas.
"""
import re
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

import comun_entregas as ce
import contenido_fichas as cf
import generar_entrega_profe as gep

RAIZ = Path(__file__).resolve().parents[2]
SALIDA = RAIZ / 'outputs' / 'entrega_profe' / 'detalle_indicadores_OCE_profe.docx'
TITULO_ETQ = {'se mantiene': 'SE MANTIENE', 'se ajusta': 'SE AJUSTA', 'se reemplaza': 'SE REEMPLAZA', 'se elimina': 'SE ELIMINA', 'nuevo': 'NUEVO'}
SIN_BRECHA = ('con dato', 'con dato (rango)')


def items(txt):
    return [t.strip() for t in re.split(r'\n?• ', txt) if t.strip()]


def vinetas_txt(doc, txt, tam=None):
    for it in items(txt):
        gep.escribir(doc.add_paragraph(style='List Bullet'), it, tam)


def nota_fila(r):
    """Nota corta de una fila de la serie: motivo si no hay dato, o aclaración si el dato lleva alguna."""
    b = r['brecha']
    if gep.num(r['valor']) is None:
        t = re.sub(r'^(sin dato|no aplica):\s*', lambda m: 'No aplica: ' if m.group(1) == 'no aplica' else '', b)
        t = gep.quitar_parentesis(t).replace(' ;', ';')
    elif b in SIN_BRECHA:
        return ''
    else:
        t = re.sub(r'^con dato;?\s*', '', b)
        t = t.replace('desconocida', 'sin información sobre veedurías').replace('(con alerta)', 'con alerta')
    t = ce.para_profe(t, decimales=True).strip()
    if len(t) > 170:
        t = t[:170].rsplit(' ', 1)[0].rstrip(',;') + '…'
    return t[0].upper() + t[1:] if t else ''


def filas_serie(dim, cod):
    out = []
    for r in gep.anual(dim, cod, 'principal'):
        mx = gep.maximo(dim, cod, r['anio'])
        vv = gep.valor_fila(r)
        val = gep.fmt(vv) if mx is None or gep.num(mx) == gep.num(r['valor']) else f'{gep.fmt(vv)}–{gep.fmt(mx)}'
        if gep.num(r['valor']) is None:
            val = 'N/D'
        nd = '' if r['numerador'] in ('', 'N/D', 'no aplica') else (f'{gep.fmt(r["numerador"])} / {gep.fmt(r["denominador"])}' if r['denominador'] else gep.fmt(r['numerador']))
        out.append([r['anio'], nd, val, gep.fmt(r['puntaje_0_100']) if gep.num(r['puntaje_0_100']) is not None else '—',
                    r['semaforo'] or '—', nota_fila(r)])
    return out


def serie_indicador(doc, dim, cod, ficha):
    unidad = ficha['unidad']
    gep.parrafo(doc, f'**Serie de {cod}** (valor en {unidad}; para los indicadores con rango, mínimo–máximo):')
    gep.tabla(doc, [1.3, 3.5, 2.3, 1.7, 2.3, 6.4], ['Año', 'Numerador / denominador', 'Valor', 'Puntaje', 'Semáforo (propuesta)', 'Nota'],
              filas_serie(dim, cod), tam=8)


def serie_indice(doc, cod):
    dim, a0 = ce.dim_de(cod), ce.PRIMER_ANIO[cod]
    filas, k, n = gep.datos_indice(dim, cod, a0, gep.ANIO_MAX)
    gep.parrafo(doc, f'**Serie del índice {ce.NOMBRE_INDICE[cod]}** (puntaje de 0 a 100):')
    gep.tabla(doc, [2.0, 6.0, 3.6, 5.8], ['Año', 'Valor', 'Semáforo (propuesta)', 'Componentes que faltan'], filas, tam=8.5)
    return k, n


def bloque_matriz(doc, orig):
    m = cf.datos_matriz(orig)
    f = [['**Objetivo**', m['objetivo']], ['**Fórmula y metodología**', m['formula']], ['**Periodicidad**', m['periodicidad']],
         ['**Semáforo del anexo de la matriz original** (no es el de la propuesta)', m['semaforo']]]
    gep.tabla(doc, [4.0, 13.5], ['Campo', 'Como era en la matriz'], f, tam=8.5)


def bloque_quedo(doc, ficha):
    filas_ = [['**Nombre**', f"{ficha['codigo']} · {ficha['nombre']}"], ['**Fórmula**', ficha['formula']], ['**Cálculo del puntaje**', ficha['puntaje']],
              ['**Unidad y sentido**', f"{ficha['unidad']}; {ficha['sentido']}"], ['**Línea base, referente o meta**', ficha['referente']]]
    if ficha['tipo'] == 'Índice':
        filas_ += [['**Pesos**', ficha['pesos']], ['**Regla**', ficha['regla']]]
    fuentes = '\n'.join(items(ficha['fuentes'])[:3]) if ficha['tipo'] != 'Índice' else ficha['fuentes']
    filas_.append(['**Fuente**', ('\n'.join('• ' + x for x in items(ficha['fuentes'])[:3]) if ficha['tipo'] != 'Índice' else fuentes)])
    gep.tabla(doc, [4.0, 13.5], ['Campo', 'Como quedó'], filas_, tam=8.5)


def bloque(doc, dim, orig, etq, nuevos, porque, fichas):
    cod_titulo = orig or nuevos[0]
    nombre = cf.datos_matriz(orig)['nombre'] if orig else fichas[nuevos[0]]['nombre']
    doc.add_heading(f'{cod_titulo} · {nombre}  [{TITULO_ETQ[etq]}]', level=2)
    if orig:
        gep.parrafo(doc, '**a) Cómo era en la matriz**')
        bloque_matriz(doc, orig)
    else:
        gep.parrafo(doc, '**a) Cómo era en la matriz.** No existía: es un indicador agregado en la revisión.')
    destino = ' y '.join(f"{n}, «{fichas[n]['nombre']}»" for n in nuevos)
    if etq == 'se elimina':
        que = '**Se elimina.** No se calcula ni entra al índice.'
    elif etq == 'nuevo':
        que = f'**Nuevo.** Se agrega {destino}.'
    elif etq == 'se reemplaza':
        que = f'**Se reemplaza** por {destino}.'
    elif etq == 'se ajusta':
        que = f'**Se ajusta:** se conserva el concepto y cambia su medición. Queda como {destino}.'
    else:
        que = f'**Se mantiene** como {destino}; los ajustes de método se explican a continuación.'
    gep.parrafo(doc, '**b) Qué pasó.** ' + que)
    gep.parrafo(doc, ('**c) Por qué se mantiene y qué cambia dentro del indicador.** ' if etq == 'se mantiene' else '**c) Por qué cambió.** ') + porque)
    if not nuevos:
        gep.parrafo(doc, '**d) Cómo quedó.** No se calcula.')
        gep.parrafo(doc, '**e) Serie.** No aplica.')
        gep.parrafo(doc, '**f) Brecha.** No aplica.')
        gep.parrafo(doc, '**g) Alertas.** Ninguna: el indicador ya no hace parte del observatorio.')
        return
    for n in nuevos:
        ficha = fichas[n]
        gep.parrafo(doc, '**d) Cómo quedó**')
        bloque_quedo(doc, ficha)
        gep.parrafo(doc, '**e) Serie por año**')
        if ficha['tipo'] == 'Índice':
            k, tot = serie_indice(doc, n)
            gep.parrafo(doc, f'**f) Brecha.** {ficha["brecha"]}.')
        else:
            serie_indicador(doc, dim, n, ficha)
            if n == 'PDM.CA':
                doc.add_paragraph()
                gep.parrafo(doc, '**Cortes y fuentes del espacio verde**')
                gep.tabla(doc, [1.4, 2.0, 2.2, 5.0, 6.8], gep.EV_ENC, gep.ev_cortes(), tam=8)
                gep.parrafo(doc, ce.para_profe(gep.EV_NOTA))
                for u in gep.EV_URLS[:4]:
                    gep.escribir(doc.add_paragraph(style='List Bullet'), ce.para_profe(u), 8.5)
            gep.parrafo(doc, f'**f) Brecha.** {ficha["brecha"]}.')
        gep.parrafo(doc, '**g) Alertas principales**')
        al = items(ficha['limitaciones'])
        for a in al[:4]:
            gep.escribir(doc.add_paragraph(style='List Bullet'), a, 9.5)
        if n.startswith('PP.') and n in ('PP.EPP',):
            pass


def resumen_indices(doc, fichas):
    cuerpo = []
    for a, vals in ce.resumen_indices():
        fila = [str(a)]
        for cod in ce.NOMBRE_INDICE:
            v, s = vals[cod]
            fila.append(f'**{gep.fmt(v)}** ({s})' if v is not None else 'sin índice')
        cuerpo.append(fila)
    gep.parrafo(doc, 'Valores de los cuatro índices por año (puntaje de 0 a 100; entre paréntesis, el semáforo (propuesta)). '
                     'Un índice se calcula solo en los años en que están todos sus componentes; no se redistribuyen pesos. ' + gep.SEMAFORO)
    gep.tabla(doc, [2.0] + [3.8] * 4, ['Año'] + list(ce.NOMBRE_INDICE.values()), cuerpo, tam=9)
    for cod, nom in ce.NOMBRE_INDICE.items():
        f = fichas[cod]
        gep.escribir(doc.add_paragraph(style='List Bullet'),
                     f'**{nom}.** Años con índice: {f["anios"].split(".")[0].replace("Calculados: ", "")}. Brecha: {f["brecha"]}. Fórmula: {f["formula"].replace(chr(10), " ")}.', 9.5)
    gep.parrafo(doc, 'Lectura: los índices no son comparables entre sí ni forman series continuas; cada uno depende de los años en que coinciden todos sus componentes '
                     '(ver la brecha de información). Con los máximos de los rangos, el ISPC de 2025 sería '
                     f'{ce.vf("PC.ISPC.MAX", 2025)} y el ICRCS sería {ce.vf("CS.ICRCS.MAX", 2024)} en 2024 y {ce.vf("CS.ICRCS.MAX", 2025)} en 2025.')


def construir(fichas_lista):
    fichas = {f['codigo']: f for f in fichas_lista}
    cambios = cf.ORIGINALES
    cuenta = Counter(c[2] for c in cambios)
    doc = gep.nuevo_documento(horizontal=False)
    doc.add_paragraph(style='Title').add_run('Detalle de la revisión de indicadores')
    gep.parrafo(doc, f'Observatorio Ciudadano de Envigado · {ce.FECHA} · Resultados **provisionales**')

    doc.add_heading('Presentación', level=1)
    n_mat = len(ce.matriz())
    gep.parrafo(doc, f'Este documento explica, indicador por indicador, qué se hizo con cada uno de los {n_mat} indicadores e índices de la matriz original y con el indicador que se agregó: '
                     f'{cuenta["se mantiene"]} se mantienen, {cuenta["se ajusta"]} se ajustan, {cuenta["se reemplaza"]} se reemplazan, {cuenta["se elimina"]} se eliminan y {cuenta["nuevo"]} es nuevo. '
                     'Para cada uno se muestra cómo era en la matriz, qué pasó y por qué, cómo quedó (fórmula, referente y fuente), la serie con los resultados por año, '
                     'la brecha de información y las alertas principales. Todos los resultados son provisionales.')
    gep.parrafo(doc, 'Etiquetas: **se mantiene** (el indicador se conserva con su definición; pueden cambiar detalles de medición); '
                     '**se ajusta** (se conserva el concepto, pero cambian componentes, referente o universo); **se reemplaza** (se sustituye por otro indicador que representa lo mismo con datos disponibles); '
                     '**se elimina** (sale del observatorio); **nuevo** (no estaba en la matriz).')

    doc.add_heading('Criterios de la revisión', level=1)
    gep.vinetas(doc, [
        '**Serie en el tiempo.** Cada indicador debe poder verse como una evolución anual. Cada indicador trae todos los años que tiene; el índice solo los años en que coinciden todos sus componentes.',
        '**Datos públicos.** Se trabaja con datos que se puedan encontrar sin pedirlos a la Alcaldía. Lo que no es público queda en «Variables por solicitar».',
        '**Sin autorreporte del 100 %.** No se usan datos en los que la propia Alcaldía se asigna el 100 % (por ejemplo, la promoción del control social o los proyectos «terminados»).',
        '**Mismo universo.** Un dato solo se compara con una meta o línea base de la misma población, el mismo concepto y el mismo periodo; si no existe, se acuerda un referente del mismo universo.',
        '**Brecha de información.** Siempre se reporta: años (o eventos, trimestres o metas) sin dato frente al total del periodo. Un índice no redistribuye pesos: si falta un componente, ese año queda sin índice.',
        '**Población única.** Todas las tasas usan las proyecciones del DANE vigentes al 8 de agosto de 2025 (VP), publicadas en TerriData.',
        '**Semáforo (propuesta).** Verde ≥ 90; Amarillo 75 a < 90; Naranja 50 a < 75; Rojo < 50, igual en las cuatro dimensiones, por validar con el profesor.',
    ])

    doc.add_heading('Resumen de los cuatro índices', level=1)
    resumen_indices(doc, fichas)

    titulos = {'PDM': '1. Cumplimiento del Plan de Desarrollo Municipal (PDM)', 'PC': '2. Participación ciudadana',
               'PP': '3. Presupuesto participativo', 'CS': '4. Control social'}
    for dim in ce.DIM_ORDEN:
        doc.add_heading(titulos[dim], level=1)
        deldim = [c for c in cambios if c[0] == dim]
        gep.parrafo(doc, 'Resumen de lo que pasó con cada indicador de la matriz:')
        gep.tabla(doc, [2.6, 5.7, 2.6, 6.6], ['Código original', 'Indicador de la matriz', 'Qué pasó', 'Queda como'],
                  [[c[1] or '—', cf.datos_matriz(c[1])['nombre'] if c[1] else '(nuevo)', f'**{c[2]}**',
                    ', '.join(f"{n}" for n in c[3]) or '—'] for c in deldim], tam=8.5)
        for _d, orig, etq, nuevos, porque in deldim:
            bloque(doc, dim, orig, etq, nuevos, porque, fichas)
        if dim == 'CS':
            doc.add_heading('Qué se necesita para seguir midiendo VRC', level=2)
            gep.vinetas(doc, gep.VRC_NECESITA)

    doc.add_heading('Decisiones tomadas en la revisión', level=1)
    gep.parrafo(doc, 'Resumen de los acuerdos que definen el método. Se tomaron en las reuniones de revisión (10 y 12 de noviembre de 2026) y en la revisión de cada dimensión.')
    extra = ('Población y veedurías (método)', 'Reglas de cruce de PP')
    sd = ce.pp_anios_sin_dato()
    texto_pp = ('Se calcula 2020–2025 y se reporta la brecha. Un año se calcula solo si se verifica al menos el 50 % de sus proyectos. '
                + ('; '.join(f'{a} queda sin dato' for a, _m in sd) + '.' if sd else 'Todos los años tienen índice.'))
    filas_dec = [[f'**{ce.para_profe(t)}**', ce.para_profe(texto_pp if t == 'Serie de PP' else d).replace('el profe', 'el profesor')]
                 for _c, t, d, _o, v in gep.DECISIONES if v or t in extra]
    gep.tabla(doc, [4.6, 12.9], ['Tema', 'Decisión'], filas_dec, tam=9)

    doc.add_heading('Variables por solicitar', level=1)
    gep.parrafo(doc, 'Información que no es pública o está incompleta y que, de obtenerse, cerraría brechas de información:')
    gep.vinetas(doc, [ce.para_profe(t) for t in gep.SOLICITAR if sd or 'Plan de acción 2021' not in t])
    doc.save(SALIDA)
    return SALIDA


def verificar(ruta, fichas_lista):
    d = Document(ruta)
    sec = d.sections[0]
    util = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
    for i, t in enumerate(d.tables):
        assert sum(c.width.cm for c in t.rows[0].cells) <= util + 0.01, ('ancho de tabla', i)
        assert 'Origen' not in [c.text for c in t.rows[0].cells], ('columna Origen', i)
        for c in t.rows[0].cells:  # todo semáforo propio va rotulado «(propuesta)»
            assert 'emáforo' not in c.text or '(propuesta)' in c.text, c.text
    texto = gep.texto_completo(d)
    assert not ce.palabras_prohibidas(texto), ce.palabras_prohibidas(texto)
    for extra in ('Isabella', 'Francisco', 'usuaria', 'Analizador', 'Ejecutor'):
        assert extra not in texto, extra
    assert not ce.ano_suelto(texto), ce.ano_suelto(texto)[:3]   # ninguna palabra «ano» suelta (debe ser «año»)
    assert not re.search(r'\bD\d{1,3}\b', texto), re.findall(r'.{20}\bD\d{1,3}\b.{10}', texto)[:3]
    for p in d.paragraphs:
        if re.search(r'sem[aá]foro', p.text, re.I) and not p.style.name.startswith('Heading'):
            assert '(propuesta)' in p.text or 'anexo' in p.text.lower() or 'propuesta' in p.text.lower(), p.text[:120]
    en_csv = ' '.join(r['enlace'] for dim in ce.DIM_ORDEN for r in ce.DATA[dim]) + ' ' + ' '.join(gep.EV_URLS)
    for u in set(re.findall(r'https?://[^\s;)\]]+', texto)):
        assert u.rstrip('.,') in en_csv, ('enlace que no está en los CSV', u)
    # 1) todos los indicadores de la matriz original aparecen, con su etiqueta
    heads = [p.text for p in d.paragraphs if p.style.name == 'Heading 2']
    for cod in ce.matriz():
        h = [x for x in heads if x.startswith(cod + ' · ')]
        assert len(h) == 1, (cod, h)
        assert any(e in h[0] for e in cf.TITULO_ETQ_LISTA), h[0]
    assert sum(1 for x in heads if '[NUEVO]' in x) == 1
    # 2) resumen e índices contra los CSV
    tablas = {t._tbl: t for t in d.tables}
    titulo, hallados, resumen_ok = '', {}, False
    for el in d.element.body.iterchildren():
        if el.tag == qn('w:p'):
            txt_p = ''.join(x.text or '' for x in el.iter(qn('w:t')))
            if txt_p.strip():
                titulo = txt_p  # último párrafo no vacío antes de la tabla («Serie del índice X»)
        elif el.tag == qn('w:tbl'):
            t = tablas[el]
            enc = [c.text for c in t.rows[0].cells]
            if enc == ['Año'] + list(ce.NOMBRE_INDICE.values()):
                resumen_ok = True
                for fila in t.rows[1:]:
                    a = int(fila.cells[0].text)
                    for j, cod in enumerate(ce.NOMBRE_INDICE, 1):
                        v, s = ce.valores_indice(cod).get(a, (None, ''))
                        esperado = f'{gep.fmt(v)} ({s})' if v is not None else 'sin índice'
                        assert fila.cells[j].text == esperado, (a, cod, fila.cells[j].text, esperado)
            if enc[:3] == ['Año', 'Valor', 'Semáforo (propuesta)'] and len(enc) == 4:
                for cod, nom in ce.NOMBRE_INDICE.items():
                    if titulo.startswith(f'Serie del índice {nom} '):
                        hallados[cod] = {f.cells[0].text: f.cells[1].text for f in t.rows[1:]}
    assert resumen_ok
    for cod in ce.NOMBRE_INDICE:
        for a, (v, s) in ce.valores_indice(cod).items():
            if v is not None:
                assert hallados[cod][str(a)] == gep.fmt(v), (cod, a, hallados[cod].get(str(a)))
    esperados = {'PP.IREPP': ('2025', '95,83'), 'PDM.ICPD7': ('2019', '62,41'), 'PC.ISPC': ('2025', '61,62'),
                 'CS.ICRCS': ('2024', '60,59')}
    for cod, (a, v) in esperados.items():
        assert hallados[cod][a] == v, (cod, a, hallados[cod])
    assert hallados['CS.ICRCS']['2025'] == '38,18'
    return d, len(d.paragraphs), len(d.tables)


def main():
    fichas = cf.construir_fichas()
    ruta = construir(fichas)
    d, np_, nt = verificar(ruta, fichas)
    print(ruta.relative_to(RAIZ), '| párrafos:', np_, '| tablas:', nt)
    print('  secciones:', [p.text for p in d.paragraphs if p.style.name == 'Heading 1'])
    print('  bloques de indicador (Heading 2):', sum(1 for p in d.paragraphs if p.style.name == 'Heading 2'))
    print('Asserts OK: los', len(ce.matriz()), 'indicadores de la matriz original aparecen con su etiqueta; sin columna «Origen» ni «Claude»/«asistente»; '
          'valores de los índices = CSV (IREPP 2025 = 95,83; ICPD-7 2019 = 62,41; ISPC 2025 = 61,62; ICRCS 2024 = 60,59 y 2025 = 38,18); '
          'semáforo rotulado «(propuesta)»; anchos dentro del margen.')


if __name__ == '__main__':
    main()
