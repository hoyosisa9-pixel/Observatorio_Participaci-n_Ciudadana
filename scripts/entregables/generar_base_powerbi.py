"""Base de datos del OCE en Excel para el tablero de Power BI (Provisional).

Lee los resultados de la construcción (outputs/construccion/<DIM>/<dim>_indicadores.csv, separador ';', BOM)
y escribe outputs/entrega_powerbi/base_indicadores_OCE.xlsx con tablas de Excel listas para importar:

  hechos         una fila por indicador y periodo (año o trimestre): valor, rango mín.-máx., puntaje, semáforo, fuente
  indices        solo los cuatro índices por año (atajo para las tarjetas del tablero)
  brechas        brecha de información por indicador y por índice (periodo completo)
  dim_dimension  las 4 dimensiones y su índice
  dim_indicador  catálogo de indicadores (código único, unidad, sentido, orden y código base)
  dim_anio       calendario de años 2003-2026
  dim_semaforo   rangos del semáforo (propuesta, D60 y D84)
  fuentes        fuentes y enlaces citados
  diccionario    descripción de cada columna
  resumen        los cuatro índices por año en formato ancho, para leer (sin Power BI)
  cambios        qué pasó con cada indicador de la matriz original y por qué
  ficha_metodologica  una fila por indicador principal e índice con su metodología (mismas columnas que la ficha en Word)

Se corre desde la raíz: python3 scripts/entregables/generar_base_powerbi.py
Si cambian los CSV de construcción (p. ej. al completar las rendiciones de Discapacidad y Víctimas con cs_icrcs.py), basta con volver a correrlo.
"""
import csv
import re
import zipfile
from collections import Counter, OrderedDict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import comun_entregas as ce
import contenido_fichas as cf
import excel_formulas

RAIZ = Path(__file__).resolve().parents[2]
SALIDA = RAIZ / 'outputs' / 'entrega_powerbi' / 'base_indicadores_OCE.xlsx'
FECHA = date.today().isoformat()

DIMENSIONES = [
    # codigo, nombre, indice, nombre del índice, fórmula, archivo
    ('PDM', 'Cumplimiento del Plan de Desarrollo Municipal', 'PDM.ICPD7', 'ICPD-7',
     'Promedio simple de 7 componentes: salud, logro educativo, servicios públicos, mortalidad vial, espacio verde, veedurías y homicidio',
     'PDM/pdm_indicadores.csv'),
    ('PC', 'Participación ciudadana', 'PC.ISPC', 'ISPC',
     '0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4', 'PC/pc_indicadores.csv'),
    ('PP', 'Presupuesto participativo', 'PP.IREPP', 'IREPP',
     '0,20·EPP + 0,40·EJ90 + 0,40·FID', 'PP/pp_indicadores.csv'),
    ('CS', 'Control social', 'CS.ICRCS', 'ICRCS',
     '(DSI + PQRSDF + VRC + RPO) / 4', 'CS/cs_indicadores.csv'),
]
# NOMBRE_TABLERO, SEMAFORO, UNIDAD_BASE y MENOR_ES_MEJOR viven en comun_entregas.py (los usan también las fichas metodológicas)

ARIAL = 'Arial'
ENC_FILL = PatternFill('solid', start_color='1F4E78')
ENC_FONT = Font(name=ARIAL, bold=True, color='FFFFFF', size=10)
CUERPO = Font(name=ARIAL, size=10)


# tildes() y TILDES viven en comun_entregas.py (se comparten con los Word de la entrega al profe)
from comun_entregas import (MENOR_ES_MEJOR, NOMBRE_TABLERO, PCT, PUNT, SEMAFORO, TASA, TILDES, UNIDAD_BASE,  # noqa: E402
                            para_profe, tildes)



def num(t):
    """Número en formato punto decimal, o None si no es número."""
    t = (t or '').strip()
    if re.fullmatch(r'-?\d+(\.\d+)?', t):
        return float(t)
    return None


def periodo(t):
    """'2024' -> (2024, None, 'Año'); '2023-T4' -> (2023, 'T4', 'Trimestre'); otro -> None."""
    t = (t or '').strip()
    if re.fullmatch(r'\d{4}', t):
        return int(t), None, 'Año'
    m = re.fullmatch(r'(\d{4})-(T[1-4]|S[12])', t)
    if m:
        return int(m.group(1)), m.group(2), 'Trimestre' if m.group(2).startswith('T') else 'Semestre'
    return None


def tiene_dato(r):
    """Sí si hay número, o si la columna brecha empieza por «con dato» (clasificaciones con dato)."""
    return num(r['valor']) is not None or r['brecha'].lower().startswith('con dato')


def leer():
    hechos, brechas = [], []
    for dim, _nom, _ind, _ni, _f, archivo in DIMENSIONES:
        ruta = RAIZ / 'outputs' / 'construccion' / archivo
        for r in csv.DictReader(ruta.open(encoding='utf-8-sig'), delimiter=';'):
            r = {k: (v or '').strip() for k, v in r.items()}
            for campo in ('indicador', 'nota', 'brecha', 'referente', 'fuente', 'enlace', 'fecha_consulta', 'valor'):
                r[campo] = para_profe(r[campo])  # tildes y sin números de decisión ni atribuciones
            r['dimension'] = dim
            p = periodo(r['anio'])
            if r['tipo'] == 'brecha' or p is None:
                brechas.append(r)
                continue
            r['anio_num'], r['subperiodo'], r['tipo_periodo'] = p
            hechos.append(r)
    return hechos, brechas


# ---------------------------------------------------------------- catálogo: unidad, sentido, orden, código base
ORDEN_DIM = {'PDM': 1, 'PC': 2, 'PP': 3, 'CS': 4}
# orden de los indicadores dentro de cada dimensión (código base): el índice primero
ORDEN_BASE = {
    'PDM': ['PDM.ICPD7', 'PDM.SALUD', 'PDM.ILE', 'PDM.SP', 'PDM.CM', 'PDM.CA', 'PDM.VEED', 'PDM.HOM'],
    'PC': ['PC.ISPC', 'PC.I1', 'PC.I2', 'PC.I3', 'PC.I4'],
    'PP': ['PP.IREPP', 'PP.EPP', 'PP.EJ90', 'PP.FID'],
    'CS': ['CS.ICRCS', 'CS.PQRSDF', 'CS.DSI', 'CS.VRC', 'CS.RPO'],
}
# respaldos cuyo código no empieza por el del indicador al que apoyan
BASE_ESPECIAL = {
    'PP.FID_MON': 'PP.FID', 'PP.COMP': 'PP.EPP', 'PP.EQUIDAD': 'PP.EPP', 'PP.TERMINADOS': 'PP.EJ90',
    'PP.COBERTURA': 'PP.IREPP', 'PP.CONTRACREDITO': 'PP.IREPP', 'PDM.POB': 'PDM.ICPD7', 'PDM.POB05': 'PDM.ILE',
}
UNIDAD_BRECHA = {'CS.PQRSDF.BRECHA': 'trimestres sin informe legible / 24 trimestres (%)', 'CS.DSI.BRECHA': 'metas-año no verificables / metas-año (%)'}
UNIDAD_COD = {
    'PDM.ILE.LA': PUNT, 'PDM.ILE.LN': PUNT, 'PDM.POB': 'habitantes', 'PDM.POB05': 'habitantes',
    'PDM.CM.TD': TASA, 'PDM.HOM.TD': TASA, 'PDM.VEED.MAX': 'conteo (veedurías inscritas)',
    'PC.I3.CONTEO': 'conteo (acciones formativas)', 'PC.I4.VOT': 'conteo (1 = publicado, 0 = no)',
    'PC.I4.RCA': 'conteo (1 = publicado, 0 = no)', 'PC.I4.RPP': 'conteo (1 = publicado, 0 = no)',
    'PC.I1.OFICIAL': '% (IPCM oficial)', 'PC.I1.POB14': '% (votos / población de 14 años y más)',
    'PP.EQUIDAD': PUNT, 'PP.CONTRACREDITO': PCT, 'PP.COBERTURA': PCT, 'PP.TERMINADOS': PCT,
    'CS.DSI.META': 'clasificación (texto)', 'CS.VRC.EVENTO': 'clasificación (texto)', 'PC.ISPC.MAX': PUNT, 'CS.ICRCS.MAX': PUNT,
    'PC.I2.RESP': PCT,
}
SIN_SENTIDO = {'PDM.POB', 'PDM.POB05', 'PC.I3.CONTEO', 'PC.I1.POB14', 'PP.COBERTURA', 'PP.CONTRACREDITO'}
# nombres genéricos del catálogo para las filas de detalle (sin mencionar un elemento concreto)
NOMBRE_DETALLE = {
    'CS.PQRSDF.TRIM': 'PQRSDF por trimestre (detalle: una fila por trimestre)',
    'CS.DSI.META': 'DSI por meta del PDM (detalle: una fila por meta)',
    'CS.VRC.EVENTO': 'VRC por evento de rendición de cuentas (detalle: una fila por evento)',
}
RANGO = {'CS.VRC', 'CS.RPO', 'CS.DSI', 'PC.I4', 'CS.PQRSDF'}  # indicadores en rango: la fila .MAX trae el máximo
ANIO_CORTE = 2025  # el año 2026 está en curso (sin cierre)


def codigo_base(c):
    c = c[:-len('.BRECHA')] if c.endswith('.BRECHA') else c
    if c in BASE_ESPECIAL:
        return BASE_ESPECIAL[c]
    m = re.fullmatch(r'PP\.SENS\d{4}\.(\w+)', c)   # sensibilidades de PP (p. ej. PP.SENS2021.EPP) -> el indicador al que apoyan
    if m:
        return 'PP.' + m.group(1)
    return '.'.join(c.split('.')[:2])


def hoja(wb, nombre, columnas, filas, formatos=None, anchos=None, titulo_tabla=None):
    ws = wb.create_sheet(nombre)
    ws.append(columnas)
    for f in filas:
        ws.append(f)
    for c in ws[1]:
        c.fill, c.font = ENC_FILL, ENC_FONT
        c.alignment = Alignment(vertical='center', wrap_text=True)
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.font = CUERPO
    formatos = formatos or {}
    for i, col in enumerate(columnas, 1):
        letra = get_column_letter(i)
        if col in formatos:
            for c in ws[letra][1:]:
                c.number_format = formatos[col]
        ancho = (anchos or {}).get(col)
        if ancho is None:
            largo = max([len(str(col))] + [len(str(f[i - 1])) for f in filas[:400] if f[i - 1] is not None])
            ancho = min(max(10, largo + 2), 60)
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = 'A2'
    if filas:
        ref = f'A1:{get_column_letter(len(columnas))}{len(filas) + 1}'
        t = Table(displayName=titulo_tabla or nombre, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(t)
    return ws


DICC_DATOS = {
    'id_dato': 'Clave única: codigo|periodo (en filas de detalle se agrega |n: una por meta, evento o trimestre)',
    'dimension': 'PDM, PC, PP o CS', 'nombre_dimension': 'Nombre completo de la dimensión',
    'codigo': 'Código único del indicador o índice (por ejemplo PDM.SALUD, PC.I2, CS.ICRCS)', 'indicador': 'Nombre completo del indicador',
    'nombre_corto': 'Nombre corto para mostrar en el tablero (solo índices e indicadores principales)',
    'tipo': 'indice (el índice de la dimensión), principal (indicador que entra al índice), subcomponente, respaldo, detalle, contexto o brecha. Para el tablero principal use indice y principal',
    'codigo_base': 'Indicador al que pertenece la fila (los respaldos, subcomponentes, detalles y brechas apuntan a su indicador principal)',
    'entra_al_indice': 'Sí para los índices y los indicadores principales', 'codigo_indice': 'Código del índice de la dimensión',
    'orden_dimension': 'Orden de la dimensión: PDM 1, PC 2, PP 3, CS 4', 'orden_indicador': 'Orden dentro de la tabla: el índice primero y después sus componentes, respaldos y detalles',
    'anio': 'Año (número). En PQRSDF trimestral es el año del trimestre. PP = vigencia; PC.I1 = año de la votación. Vacío en las filas de brecha',
    'subperiodo': 'T1–T4 o S1–S2 cuando el dato es trimestral o semestral; vacío si es anual', 'tipo_periodo': 'Año, Trimestre, Semestre o Periodo (filas de brecha)',
    'periodo': 'Periodo tal como viene en la fuente (2024, 2023-T4, 2012-2025 en las brechas)', 'unidad': '%, tasa por 100.000 hab., m²/hab, conteo, puntaje 0–100, habitantes, clasificación (texto) o la unidad de la brecha',
    'sentido': 'mayor es mejor, menor es mejor (mortalidad vial, homicidio y brechas) o no aplica (detalles y conteos de apoyo)',
    'peso_en_indice': 'Peso del indicador en el índice de su dimensión (1/7 = 0,143 en el ICPD-7)',
    'numerador': 'Numerador o primera variable del cálculo (conteos, pesos, habitantes). En las brechas: unidades sin dato',
    'denominador': 'Denominador o segunda variable del cálculo. En las brechas: total del periodo',
    'valor': 'Valor del indicador (número). En los indicadores en rango es el mínimo. Vacío si no hay dato. En las brechas: porcentaje de brecha',
    'valor_min': 'Mínimo del rango (PQRSDF, DSI, VRC, RPO, I4 2025); vacío si no hay rango', 'valor_max': 'Máximo del rango; sale de la fila .MAX del mismo año',
    'valor_texto': 'Texto original cuando no hay número (N/D, no aplica, pendiente…)', 'puntaje_0_100': 'Puntaje de 0 a 100 que entra al índice',
    'semaforo': 'Verde, Amarillo, Naranja o Rojo (propuesta; ver Inicio)', 'color_semaforo': 'Color del semáforo en hexadecimal, para el tablero',
    'tiene_dato': 'Sí si hay número o si la brecha empieza por «con dato»; No si no hay dato (no es un cero)',
    'vigencia_en_curso': 'Sí si el año (2026) está en curso y no se califica', 'brecha': '«con dato» o «sin dato: motivo»; en las filas de brecha, el texto de la brecha del periodo',
    'referente': 'Meta, línea base o referente usado para el puntaje', 'fuente': 'Documento, tabla o página de la fuente', 'enlace': 'Enlace(s) a la fuente (puede haber varios)',
    'enlace_principal': 'Primera dirección web de la columna enlace', 'fecha_consulta': 'Fecha(s) de consulta, tal como se registró',
    'fecha_consulta_fecha': 'Primera fecha de consulta como fecha de Excel', 'estado': 'Provisional', 'nota': 'Método, supuestos y alertas',
}


def main():
    hechos, brechas = leer()
    nombre_dim = {d[0]: d[1] for d in DIMENSIONES}
    indice_de = {d[0]: d[2] for d in DIMENSIONES}
    codigos_indice = {d[2] for d in DIMENSIONES}

    # ---------------- catálogo de indicadores
    cat = OrderedDict()
    for r in hechos + brechas:
        c = r['codigo']
        if c not in cat:
            cat[c] = {'dimension': r['dimension'], 'nombres': Counter(), 'tipos': Counter(), 'pesos': Counter(), 'refs': Counter()}
        cat[c]['nombres'][r['indicador']] += 1
        cat[c]['tipos'][r['tipo']] += 1
        if r['peso_en_indice']:
            cat[c]['pesos'][r['peso_en_indice']] += 1
        if r['referente']:
            cat[c]['refs'][r['referente']] += 1
    orden_tipo = {'indice': 1, 'principal': 2, 'subcomponente': 3, 'respaldo': 4, 'detalle': 5, 'contexto': 6, 'brecha': 7}
    filas_cat = []
    for i, (c, v) in enumerate(cat.items()):
        tipo = v['tipos'].most_common(1)[0][0]
        base = codigo_base(c)
        rank = ORDEN_BASE[v['dimension']].index(base) if base in ORDEN_BASE[v['dimension']] else 99
        filas_cat.append((ORDEN_DIM[v['dimension']], rank, orden_tipo.get(tipo, 9), i, c, v, tipo, base))
    filas_cat.sort(key=lambda x: x[:4])
    dim_ind = []
    for orden, (od, rank, ot, _i, c, v, tipo, base) in enumerate(filas_cat, 1):
        peso = num(v['pesos'].most_common(1)[0][0]) if v['pesos'] else None
        anios = sorted({r['anio_num'] for r in hechos if r['codigo'] == c and tiene_dato(r)})
        nombre = NOMBRE_DETALLE.get(c) or v['nombres'].most_common(1)[0][0]
        unidad = UNIDAD_BRECHA.get(c, 'años sin dato / años del periodo (%)') if tipo == 'brecha' else (UNIDAD_COD.get(c) or UNIDAD_BASE.get(base, PCT))
        if tipo == 'brecha':
            sentido = 'menor es mejor'
        elif tipo == 'detalle' or c in SIN_SENTIDO:
            sentido = 'no aplica'
        else:
            sentido = 'menor es mejor' if c in MENOR_ES_MEJOR or (base in MENOR_ES_MEJOR and not c.endswith('.POB')) else 'mayor es mejor'
        dim_ind.append([c, nombre, NOMBRE_TABLERO.get(c), v['dimension'], nombre_dim[v['dimension']], od, indice_de[v['dimension']],
                        base, tipo, ot, orden, 'Sí' if tipo in ('principal', 'indice') else 'No', peso, unidad, sentido,
                        v['refs'].most_common(1)[0][0] if v['refs'] else None,
                        min(anios) if anios else None, max(anios) if anios else None, len(anios)])
    codigos = {d[0] for d in dim_ind}
    col_cat = ['codigo', 'indicador', 'nombre_tablero', 'dimension', 'nombre_dimension', 'orden_dimension', 'codigo_indice',
               'codigo_base', 'tipo', 'orden_tipo', 'orden_indicador', 'entra_al_indice', 'peso_en_indice', 'unidad', 'sentido',
               'referente', 'primer_anio_con_dato', 'ultimo_anio_con_dato', 'anios_con_dato']

    # ---------------- hechos
    hechos.sort(key=lambda r: (r['dimension'], orden_tipo.get(r['tipo'], 9), r['codigo'], r['anio_num'], r['subperiodo'] or ''))
    maximos = {(r['codigo'][:-4], r['anio']): num(r['valor']) for r in hechos if r['codigo'].endswith('.MAX') and r['codigo'][:-4] in RANGO}
    f_hechos = []
    vistos = Counter()
    for r in hechos:
        v = num(r['valor'])
        vistos[(r['codigo'], r['anio'])] += 1
        n = vistos[(r['codigo'], r['anio'])]
        vmin = vmax = None
        if r['tipo'] == 'principal' and (r['codigo'], r['anio']) in maximos and maximos[(r['codigo'], r['anio'])] is not None:
            vmin, vmax = v, maximos[(r['codigo'], r['anio'])]
        fecha = re.search(r'\d{4}-\d{2}-\d{2}', r['fecha_consulta'])
        url = re.search(r'https?://[^\s;]+', r['enlace'])
        f_hechos.append([
            f"{r['codigo']}|{r['anio']}" + (f"|{n}" if n > 1 or r['tipo'] == 'detalle' else ''), r['dimension'], r['codigo'], r['indicador'], r['tipo'],
            r['anio_num'], r['subperiodo'], r['tipo_periodo'], r['anio'],
            num(r['numerador']), num(r['denominador']), v, vmin, vmax, None if v is not None else (r['valor'] or None),
            num(r['puntaje_0_100']), num(r['peso_en_indice']), r['semaforo'] or None,
            'Sí' if tiene_dato(r) else 'No', r['brecha'] or None, r['referente'] or None,
            r['fuente'] or None, r['enlace'] or None, url.group(0) if url else None, r['fecha_consulta'] or None,
            date.fromisoformat(fecha.group(0)) if fecha else None, r['estado'] or None,
            'Sí' if r['anio_num'] > ANIO_CORTE else 'No', r['nota'] or None])
    col_hechos = ['id_hecho', 'dimension', 'codigo', 'indicador', 'tipo', 'anio', 'subperiodo', 'tipo_periodo', 'periodo',
                  'numerador', 'denominador', 'valor', 'valor_min', 'valor_max', 'valor_texto', 'puntaje_0_100', 'peso_en_indice',
                  'semaforo', 'tiene_dato', 'brecha', 'referente', 'fuente', 'enlace', 'enlace_principal', 'fecha_consulta',
                  'fecha_consulta_fecha', 'estado', 'vigencia_en_curso', 'nota']

    # ---------------- índices (atajo; solo hasta el último año cerrado)
    f_ind = []
    for r in hechos:
        if r['tipo'] == 'indice' and r['codigo'] in codigos_indice and r['anio_num'] <= ANIO_CORTE:
            v = num(r['valor'])
            f_ind.append([r['dimension'], nombre_dim[r['dimension']], r['codigo'], r['indicador'], r['anio_num'], v,
                          r['semaforo'] or None, 'Sí' if tiene_dato(r) else 'No', 'No', r['nota'] or None])
    f_ind.sort(key=lambda x: (x[0], x[4]))

    # ---------------- brechas
    f_br = []
    for r in brechas:
        f_br.append([r['dimension'], r['codigo'], r['indicador'], r['tipo'], r['anio'], num(r['numerador']),
                     num(r['denominador']), num(r['valor']), r['brecha'] or None, r['nota'] or None])
    f_br.sort(key=lambda x: (x[0], x[1]))

    # ---------------- fuentes
    fu = OrderedDict()
    for r in hechos:
        if r['fuente']:
            k = (r['fuente'], r['enlace'])
            fu.setdefault(k, {'dims': set(), 'n': 0, 'fecha': r['fecha_consulta']})
            fu[k]['dims'].add(r['dimension']); fu[k]['n'] += 1
    f_fu = [[i, f, e or None, ', '.join(sorted(v['dims'])), v['n'], v['fecha'] or None]
            for i, ((f, e), v) in enumerate(fu.items(), 1)]

    anios_dim = list(range(2003, 2027))
    sem_nombres = {s[0] for s in SEMAFORO}

    # ---------------- controles
    assert len(codigos) == len(dim_ind), 'códigos de indicador repetidos'
    assert len({d[10] for d in dim_ind}) == len(dim_ind), 'orden_indicador repetido'
    faltan = [d[0] for d in dim_ind if d[11] == 'Sí' and not d[2]]
    assert not faltan, f'falta nombre_tablero: {faltan}'
    assert len({f[0] for f in f_hechos}) == len(f_hechos), 'id_hecho repetido'
    assert len(anios_dim) == len(set(anios_dim))
    assert len({s[0] for s in SEMAFORO}) == len(SEMAFORO)
    for f in f_hechos:
        assert f[2] in codigos, f[2]
        assert f[5] in anios_dim, f[5]
        assert f[17] is None or f[17] in sem_nombres, f[17]
    for x in f_ind:
        assert x[2] in codigos and x[4] in anios_dim and (x[6] is None or x[6] in sem_nombres)
    for x in f_br:
        assert x[1] in codigos, x[1]
    assert all(d[7] in codigos for d in dim_ind), [d[0] for d in dim_ind if d[7] not in codigos]
    for d in dim_ind:
        assert d[13] and d[14] in ('mayor es mejor', 'menor es mejor', 'no aplica'), d[:2]
        if d[8] == 'detalle':
            assert not re.search(r'\d{4}|CS-VRC|meta \(Ap', d[1]), d[1]  # nombre genérico, sin elemento concreto
    # tiene_dato coherente con brecha y con valor
    for f in f_hechos:
        td, br, val = f[18], (f[19] or '').lower(), f[11]
        if val is not None:
            assert td == 'Sí', f[0]
        if td == 'Sí':
            assert not br.startswith('sin dato'), (f[0], br)
        else:
            assert not br.startswith('con dato'), (f[0], br)
    # rangos coherentes
    for f in f_hechos:
        if f[12] is not None:
            assert f[13] >= f[12] - 1e-9, (f[0], f[12], f[13])
    vrc = {(f[2], f[5]): f for f in f_hechos}
    assert vrc[('CS.VRC', 2024)][12] == 66.67 and vrc[('CS.VRC', 2024)][13] == 100.0
    assert vrc[('PC.I4', 2025)][12] == 66.67 and vrc[('PC.I4', 2025)][13] == 100.0
    assert vrc[('CS.RPO', 2021)][13] == 100.0
    assert all(f[23] is None or f[23].startswith('http') for f in f_hechos)
    assert all(f[27] == 'Sí' for f in f_hechos if f[5] == 2026) and all(f[27] == 'No' for f in f_hechos if f[5] <= 2025)
    assert all(x[4] <= ANIO_CORTE for x in f_ind)
    # sin textos sin tilde frecuentes en los campos visibles
    for f in f_hechos:
        for campo in (f[3], f[28]):
            assert not re.search(r'\b(anios|poblacion|Alcaldia|informacion|ejecucion|votacion)\b', campo or ''), (f[0], campo[:80])
    esperados = {('PP', 2025): 95.83, ('PDM', 2019): 62.41, ('PC', 2025): 61.62}
    for (d, a), v in esperados.items():
        got = [x[5] for x in f_ind if x[0] == d and x[4] == a]
        assert got and abs(got[0] - v) < 0.005, (d, a, got)
    cs_csv = {int(r['anio']): (None if r['valor'] in ('N/D', '') else float(r['valor']))
              for r in csv.DictReader((RAIZ / 'outputs' / 'construccion' / 'CS' / 'cs_indicadores.csv').open(encoding='utf-8-sig'), delimiter=';') if r['codigo'] == 'CS.ICRCS'}
    cs_got = {x[4]: x[5] for x in f_ind if x[0] == 'CS'}
    assert cs_got.keys() == cs_csv.keys() and all((cs_got[a] is None and cs_csv[a] is None) or (cs_got[a] is not None and cs_csv[a] is not None and abs(cs_got[a] - cs_csv[a]) < 0.005) for a in cs_got), ('ICRCS distinto del CSV', cs_got, cs_csv)
    assert len([x for x in f_ind if x[0] == 'CS']) == 6 and len([x for x in f_ind if x[0] == 'PDM']) == 7

    # ---------------- libro con fórmulas (ver excel_formulas.py)
    fichas = cf.construir_fichas()
    cat_u = {d[0]: d for d in dim_ind}
    for f in fichas:
        assert (f['unidad'], f['sentido']) == (cat_u[f['codigo']][13], cat_u[f['codigo']][14]), f['codigo']
    info = excel_formulas.construir(hechos, brechas, fichas, dim_ind, codigo_base, SALIDA, FECHA, leame)

    # ---------------- verificación del archivo guardado (valores guardados = CSV)
    import openpyxl
    wd = openpyxl.load_workbook(SALIDA, data_only=True)
    wf = openpyxl.load_workbook(SALIDA)
    assert wd.sheetnames == ['Léame', 'Índices', 'Datos', 'Indicadores', 'Detalle'], wd.sheetnames
    assert sorted(t for w in wf.worksheets for t in w.tables) == ['Datos', 'Detalle', 'Indicadores', 'Indices']
    assert 'fullCalcOnLoad="1"' in zipfile.ZipFile(SALIDA).read('xl/workbook.xml').decode()
    csvv = {(r['codigo'], r['anio']): r for r in hechos}
    tab = lambda w, n: [dict(zip([c.value for c in w[n][1]], row)) for row in w[n].iter_rows(min_row=2, values_only=True)]
    ind = tab(wd, 'Índices')
    assert len(ind) == 25
    for r in ind:
        esp = num(csvv[(r['codigo'], str(r['anio']))]['valor'])
        assert (r['puntaje'] in (None, '') and esp is None) or abs(r['puntaje'] - esp) <= 0.005, r
        assert r['tiene_dato'] == ('Sí' if esp is not None else 'No')
        mx = csvv.get((r['codigo'] + '.MAX', str(r['anio'])))
        if r['puntaje_max'] not in (None, ''):
            assert mx and abs(r['puntaje_max'] - num(mx['valor'])) <= 0.005, r
    got = {(r['codigo'], r['anio']): r for r in ind}
    clave = {('PP.IREPP', 2020): 57.98, ('PP.IREPP', 2021): 90.71, ('PP.IREPP', 2025): 95.83, ('PDM.ICPD7', 2019): 62.41, ('PC.ISPC', 2025): 61.62,
             ('CS.ICRCS', 2024): 60.59, ('CS.ICRCS', 2025): 38.18}
    for k, v in clave.items():
        assert abs(got[k]['puntaje'] - v) <= 0.005, (k, got[k]['puntaje'], v)
    for k, v in {('PC.ISPC', 2025): 69.95, ('CS.ICRCS', 2024): 79.47, ('CS.ICRCS', 2025): 55.71}.items():
        assert abs(got[k]['puntaje_max'] - v) <= 0.005, (k, got[k]['puntaje_max'], v)
    dat = tab(wd, 'Datos')
    assert len(dat) == sum(1 for r in hechos if r['tipo'] == 'principal' and r['tipo_periodo'] == 'Año' and r['anio_num'] <= ANIO_CORTE)
    assert len({r['id'] for r in dat}) == len(dat)
    for r in dat:
        esp = num(csvv[(r['codigo'], str(r['anio']))]['puntaje_0_100'])
        assert (r['puntaje'] in (None, '') and esp is None) or abs(r['puntaje'] - esp) <= 0.005, (r['id'], r['puntaje'], esp)
        assert r['tiene_dato'] in ('Sí', 'No', 'No aplica')
    cat_ = {r['codigo']: r for r in tab(wd, 'Indicadores')}
    assert len(cat_) == 22
    for cod, (s_, t_) in {'PDM.ICPD7': (4, 7), 'PC.ISPC': (5, 6), 'CS.ICRCS': (4, 6), 'PC.I1': (2, 8), 'PC.I2': (1, 6), 'PP.IREPP': (0, 6)}.items():
        assert (cat_[cod]['brecha_sin_dato'], cat_[cod]['brecha_total']) == (s_, t_), (cod, cat_[cod]['brecha_sin_dato'], cat_[cod]['brecha_total'])
    assert (cat_['PDM.VEED']['brecha_sin_dato'], cat_['PDM.VEED']['brecha_total']) == (19, 52)
    assert (cat_['CS.PQRSDF']['brecha_sin_dato'], cat_['CS.PQRSDF']['brecha_total']) == (0, 24)
    assert (cat_['CS.DSI']['brecha_sin_dato'], cat_['CS.DSI']['brecha_total']) == (5, 20)
    assert abs(cat_['PDM.SALUD']['peso'] - 1 / 7) < 1e-12 and abs(cat_['PP.EJ90']['peso'] - 0.4) < 1e-12
    for fila in tab(wd, 'Detalle'):
        assert fila['codigo'] not in ('PDM.CM.TD', 'PDM.HOM.TD', 'PDM.POB') and not str(fila['codigo']).endswith('.MAX') and fila['anio'] <= ANIO_CORTE
    # sin color de semáforo, sin palabras prohibidas, sin números de decisión, sin nombres, sin «ano» suelta
    for w in wf.worksheets:
        for row in w.iter_rows():
            for c in row:
                if c.fill is not None and c.fill.fgColor is not None and str(c.fill.fgColor.rgb)[-6:] in ('C8E6C9', 'FFF3B0', 'FFD8A8', 'F5B7B1'):
                    raise AssertionError(('color de semáforo', w.title, c.coordinate))
                if isinstance(c.value, str):
                    assert not ce.palabras_prohibidas(c.value), (w.title, c.coordinate, c.value[:80])
                    assert not re.search(r'\bD\d{1,3}\b|usuaria|Isabella|Analizador', c.value), (w.title, c.coordinate, c.value[:80])
                    assert not ce.ano_suelto(c.value), (w.title, c.coordinate)
    nform = sum(1 for w in wf.worksheets for row in w.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('='))
    assert all(not re.search(r'XLOOKUP|(?<![A-Z])IFS\(|MINIFS|LET\(|ROUND', c.value) for w in wf.worksheets for row in w.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith('='))
    print(f"{SALIDA.relative_to(RAIZ)} | Índices {info['indices']} filas | Datos {info['datos']} x 20 | Indicadores {info['indicadores']} x 23 | Detalle {info['detalle']} x 17 | fórmulas {nform}")
    print('hojas:', wd.sheetnames)


def leame(ws, fecha):
    """Hoja Léame (texto)."""
    L = [
        ('Base de datos de indicadores del Observatorio Ciudadano de Envigado (OCE)', 'T'),
        (f'Generada el {fecha}. Todos los resultados son PROVISIONALES.', ''),
        ('', ''),
        ('Qué es este archivo', 'S'),
        ('Reúne los cuatro índices del observatorio (cumplimiento del PDM, participación ciudadana, presupuesto participativo y control social), los indicadores que los componen y sus datos de base, con su serie por año.', ''),
        ('Es una base de datos con fórmulas: cambie una entrada y se recalculan el valor, el puntaje, el semáforo (propuesta) y el índice. Cada cifra trae su fuente, su enlace y su fecha de consulta.', ''),
        ('', ''),
        ('Las tablas y cómo se relacionan', 'S'),
        ('Indicadores: catálogo de 22 filas (los 18 indicadores principales y los 4 índices) con definición, unidad, sentido, peso, método del puntaje, línea base, referente, brecha y fuente.', ''),
        ('Datos: una fila por indicador principal y año (hasta 2025): numerador, denominador, valor, puntaje de 0 a 100 y semáforo (propuesta). Se relaciona con Indicadores por codigo.', ''),
        ('Indices: una fila por índice y año: puntaje del índice, máximo (si hay rango) y semáforo (propuesta). Se relaciona con Indicadores por codigo.', ''),
        ('Detalle: entradas de apoyo (subcomponentes, respaldos, eventos, metas y trimestres). Se relaciona con Indicadores por codigo_padre.', ''),
        ('', ''),
        ('La cadena de cálculo', 'S'),
        ('Entradas (Detalle y las columnas numerador, denominador y valor de Datos) → valor (columna valor) → puntaje de 0 a 100 (columna puntaje) → índice (hoja Índices) → brecha (hoja Indicadores).', ''),
        ('Las fórmulas están en las propias celdas y usan solo funciones clásicas (IF, OR, MIN, MAX, COUNTIFS, SUMIFS y AVERAGEIFS). No hay redondeos: el formato muestra 2 decimales, pero el cálculo usa todos.', ''),
        ('Los métodos del puntaje (columna metodo_puntaje de Indicadores): proporcion (el valor ya es el puntaje), tope (máximo 100), inverso (menor es mejor), avance (desde la línea base hacia el referente), veedurias y promedio.', ''),
        ('', ''),
        ('Reglas', 'S'),
        ('No se redistribuyen pesos: un índice solo se calcula si están todos sus componentes del año (componentes_con_dato = componentes). Si falta uno, el puntaje queda vacío.', ''),
        ('«Sin dato» no es cero: una celda vacía significa que no hay dato publicado o verificable. «No aplica» (por ejemplo, el año sin votación) no cuenta en la brecha.', ''),
        ('Rango mínimo–máximo: valor es el mínimo (solo lo documentado) y entra al índice; valor_max suma lo no verificable (sin_verificar); brecha_rango = valor_max − valor. El puntaje_max del índice suma esas brechas con su peso.', ''),
        ('Brecha de información: años (o trimestres, consecutivos o metas) sin dato / total del periodo. Se calcula con fórmula en Indicadores; veedurías, PQRSDF y DSI traen su propia unidad.', ''),
        ('Población: todas las tasas usan las proyecciones del DANE vigentes al 8-ago-2025 (VP). El año 2026 está en curso y no se incluye.', ''),
        ('Semáforo (propuesta, por validar): Verde ≥ 90; Amarillo 75 a < 90; Naranja 50 a < 75; Rojo < 50, sobre el puntaje. En los indicadores con rango se asigna con el mínimo.', ''),
        ('', ''),
        ('Diccionario corto', 'S'),
        ('id: clave única (código|año). dimension: PDM, PC, PP o CS. codigo: código del indicador o índice (ver Indicadores). anio: año (en PC.I1 es el año de la votación; en PP, la vigencia).', ''),
        ('numerador / denominador: variables del cálculo. sin_verificar: parte del numerador que no se pudo verificar (rango). valor: valor del indicador. puntaje: puntaje de 0 a 100.', ''),
        ('tiene_dato: Sí, No o No aplica. motivo_sin_dato: por qué no hay dato. fuente, enlace y otros_enlaces: de dónde sale. fecha_consulta y nota: trazabilidad y método.', ''),
        ('sentido: mayor es mejor o menor es mejor (mortalidad vial y homicidio). unidad: unidad del valor. peso: peso del indicador en el índice de su dimensión.', ''),
        ('', ''),
        ('Nota para Power BI', 'S'),
        ('Importe las tablas Datos, Indices e Indicadores; Detalle es opcional. Las fórmulas traen su valor guardado, así que se pueden leer sin recalcular.', ''),
        ('Relaciones: Indicadores[codigo] 1 → * Datos[codigo] y 1 → * Indices[codigo]; Detalle[codigo_padre] * → 1 Indicadores[codigo].', ''),
        ('Tipos: anio como número entero; enlace y enlace_principal como URL. El color del semáforo no viene en el archivo: defínalo con reglas sobre puntaje (≥ 90, ≥ 75, ≥ 50).', ''),
        ('No sume valor entre indicadores: tienen unidades distintas. Para comparar use puntaje. Los filtros útiles: dimension, codigo y anio.', ''),
        ('', ''),
        ('Cómo se generó', 'S'),
        ('Con scripts/entregables/generar_base_powerbi.py, a partir de los resultados de la construcción (outputs/construccion/*/..._indicadores.csv). Si los resultados se actualizan, se vuelve a correr.', ''),
    ]
    from openpyxl.styles import Alignment, Font
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 150
    for i, (t, e) in enumerate(L, 1):
        c = ws.cell(i, 1, t)
        c.font = Font(name='Arial', size=14 if e == 'T' else 11 if e == 'S' else 10, bold=e in ('T', 'S'), color='1F4E78' if e else '000000')
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if len(t) > 150:
            ws.row_dimensions[i].height = 14 * (len(t) // 150 + 1)


if __name__ == '__main__':
    main()
