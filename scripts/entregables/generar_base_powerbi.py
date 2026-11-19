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
from collections import Counter, OrderedDict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import comun_entregas as ce
import contenido_fichas as cf
import hojas_excel

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

    # ---------------- hoja Datos (tabla larga desnormalizada para Power BI)
    cat = {d[0]: d for d in dim_ind}
    color = {n: c for n, _d, _h, c, _o in SEMAFORO}
    col_datos = ['id_dato', 'dimension', 'nombre_dimension', 'codigo', 'indicador', 'nombre_corto', 'tipo', 'codigo_base', 'entra_al_indice',
                 'codigo_indice', 'orden_dimension', 'orden_indicador', 'anio', 'subperiodo', 'tipo_periodo', 'periodo', 'unidad', 'sentido',
                 'peso_en_indice', 'numerador', 'denominador', 'valor', 'valor_min', 'valor_max', 'valor_texto', 'puntaje_0_100', 'semaforo',
                 'color_semaforo', 'tiene_dato', 'vigencia_en_curso', 'brecha', 'referente', 'fuente', 'enlace', 'enlace_principal',
                 'fecha_consulta', 'fecha_consulta_fecha', 'estado', 'nota']
    f_datos = []

    for f in f_hechos:
        d = cat[f[2]]
        f_datos.append([f[0], f[1], d[4], f[2], f[3], d[2], f[4], d[7], d[11], d[6], d[5], d[10], f[5], f[6], f[7], f[8], d[13], d[14],
                        f[16], f[9], f[10], f[11], f[12], f[13], f[14], f[15], f[17], color.get(f[17]), f[18], f[27], f[19], f[20],
                        f[21], f[22], f[23], f[24], f[25], f[26], f[28]])
    for x in f_br:   # filas de brecha: periodo en texto, anio vacío; valor = brecha en %
        d = cat[x[1]]
        f_datos.append([f'{x[1]}|{x[4]}', x[0], d[4], x[1], x[2], d[2], x[3], d[7], d[11], d[6], d[5], d[10], None, None, 'Periodo', x[4], d[13], d[14],
                        None, x[5], x[6], x[7], None, None, None, None, None, None, 'Sí', 'No', x[8], None, None, None, None, None, None, 'Provisional', x[9]])
    f_datos.sort(key=lambda r: (r[10], r[11], r[3], r[12] if r[12] is not None else 9999, r[13] or '', str(r[15])))
    assert all(len(r) == len(col_datos) for r in f_datos)
    assert len({r[0] for r in f_datos}) == len(f_datos), 'id_dato repetido'
    assert 'origen' not in [c.lower() for c in col_datos]
    assert sum(1 for r in f_datos if r[6] == 'brecha') == len(f_br)
    assert sum(1 for r in f_datos if r[6] != 'brecha') == len(f_hechos)
    # unidad y sentido de las fichas = los del catálogo
    fichas = cf.construir_fichas()
    for f in fichas:
        assert (f['unidad'], f['sentido']) == (cat[f['codigo']][13], cat[f['codigo']][14]), f['codigo']
    n_princ = sum(1 for d in dim_ind if d[8] == 'principal')
    n_idx = sum(1 for d in dim_ind if d[8] == 'indice')
    assert len(fichas) == n_princ + n_idx and {f['codigo'] for f in fichas} == {d[0] for d in dim_ind if d[11] == 'Sí'}

    # series para la hoja Indicadores: {código: {año: (puntaje, semáforo, valor, máximo)}}
    serie = {}
    for f in f_hechos:
        if f[4] in ('principal', 'indice') and f[7] == 'Año' and f[5] <= ANIO_CORTE:
            serie.setdefault(f[2], {})[f[5]] = (f[15], f[17], f[11], f[13])
    maximos_idx = {}
    for f in f_hechos:
        if f[2] in ('PC.ISPC.MAX', 'CS.ICRCS.MAX') and f[11] is not None and f[5] <= ANIO_CORTE:
            maximos_idx.setdefault(f[2][:-4], {})[f[5]] = f[11]
    brechas_txt = {f['codigo']: f['brecha'] for f in fichas}
    nombres = {c: (cat[c][2] or cat[c][1]) for c in brechas_txt}

    # ---------------- libro (5 hojas)
    wb = Workbook()
    dic_datos = [(c, DICC_DATOS[c]) for c in col_datos]
    assert all(c in DICC_DATOS for c in col_datos), [c for c in col_datos if c not in DICC_DATOS]
    hojas_excel.hoja_inicio(wb, FECHA, col_datos, dic_datos)
    hojas_excel.hoja_indices(wb, fichas, maximos_idx)
    hojas_excel.hoja_indicadores(wb, fichas, serie, nombres, brechas_txt)
    _ws_ficha, n_extra = hojas_excel.hoja_ficha(wb, fichas)
    ws = hoja(wb, 'Datos', col_datos, f_datos,
              formatos={'numerador': '#,##0.##', 'denominador': '#,##0.##', 'valor': '0.00', 'valor_min': '0.00', 'valor_max': '0.00',
                        'puntaje_0_100': '0.00', 'peso_en_indice': '0.000', 'anio': '0', 'fecha_consulta_fecha': 'yyyy-mm-dd'},
              anchos={'id_dato': 28, 'indicador': 45, 'nombre_corto': 30, 'nota': 60, 'fuente': 50, 'enlace': 50, 'enlace_principal': 50,
                      'brecha': 40, 'referente': 30, 'nombre_dimension': 30, 'unidad': 24})
    ws.freeze_panes = ws.cell(2, 5)
    hojas_excel.horizontal(ws, '1:1')
    assert wb.sheetnames == ['Inicio', 'Índices', 'Indicadores', 'Ficha metodológica', 'Datos']

    # ningún texto del libro lleva «origen», «Claude» ni «asistente», ni atribuciones, ni números de decisión
    for ws_ in wb.worksheets:
        for fila in ws_.iter_rows():
            for c in fila:
                if isinstance(c.value, str):
                    assert not ce.palabras_prohibidas(c.value), (ws_.title, c.coordinate, c.value[:80])
                    assert not ce.ano_suelto(c.value), (ws_.title, c.coordinate, ce.ano_suelto(c.value)[:2])   # «ano» suelto = falta la eñe
                    assert not re.search(r'\bD\d{1,3}\b|usuaria|Isabella|Analizador', c.value), (ws_.title, c.coordinate, c.value[:80])

    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    wb.save(SALIDA)

    # ---------------- verificación del archivo guardado
    import openpyxl
    wl = openpyxl.load_workbook(SALIDA)
    assert wl.sheetnames == ['Inicio', 'Índices', 'Indicadores', 'Ficha metodológica', 'Datos']
    assert list(wl['Datos'].tables) == ['Datos'] and len(wl['Datos'].tables) == 1 and sum(len(w.tables) for w in wl.worksheets) == 1
    assert wl['Datos'].max_row - 1 == len(f_datos)
    wi = wl['Índices']
    cab = next(r for r in range(1, 30) if wi.cell(r, 1).value == 'Índice')
    anios_h = [int(wi.cell(cab, j).value) for j in range(3, 10)]
    assert anios_h == list(range(2019, 2026))
    for k, cod in enumerate(ce.NOMBRE_INDICE, 1):
        assert wi.cell(cab + k, 1).value == ce.NOMBRE_INDICE[cod]
        vals = ce.valores_indice(cod)
        for j, a in enumerate(anios_h):
            got, esp = wi.cell(cab + k, 3 + j).value, vals.get(a, (None, ''))[0]
            assert (esp is None and got == '—') or (esp is not None and abs(got - esp) < 0.005), (cod, a, got, esp)
    got = {(wi.cell(cab + k, 1).value, a): wi.cell(cab + k, 3 + j).value for k in range(1, 5) for j, a in enumerate(anios_h)}
    assert abs(got[('IREPP', 2025)] - 95.83) < 0.005 and abs(got[('ICPD-7', 2019)] - 62.41) < 0.005 and abs(got[('ISPC', 2025)] - 61.62) < 0.005
    assert abs(got[('ICRCS', 2024)] - 60.59) < 0.005 and abs(got[('ICRCS', 2025)] - 38.18) < 0.005
    # Indicadores: el puntaje de cada celda coincide con Datos
    wd = {(r[3], r[12]): r[25] for r in f_datos if r[6] in ('principal', 'indice') and r[14] == 'Año'}
    wn = wl['Indicadores']
    anios_i = {j: int(wn.cell(4, j).value) for j in range(7, 7 + len(hojas_excel.ANIOS_VISTA))}
    n_celdas = 0
    for r in range(5, wn.max_row + 1):
        nom = wn.cell(r, 1).value
        if wn.cell(r, 2).value == 'Puntaje 0–100':
            cod = next(c for c, n in nombres.items() if n == nom and (wn.cell(r, 3).value == cat[c][13]))
            for j, a in anios_i.items():
                v = wn.cell(r, j).value
                esp = wd.get((cod, a))
                assert (v == '—' and esp is None) or (esp is not None and abs(v - esp) < 1e-9), (cod, a, v, esp)
                n_celdas += 1
    assert n_celdas == (n_princ + n_idx) * len(hojas_excel.ANIOS_VISTA)
    nf = wl['Ficha metodológica']
    assert nf.max_row - 1 == len(fichas) + 2
    texto_ficha = ' '.join(str(c.value) for row in nf.iter_rows() for c in row if c.value)
    assert all(cod in texto_ficha for cod in ce.matriz()), 'faltan indicadores de la matriz original en la ficha'
    assert all(w.page_setup.orientation == 'landscape' for w in wl.worksheets)
    print(f'{SALIDA.relative_to(RAIZ)} | Datos {len(f_datos)} filas x {len(col_datos)} columnas ({len(f_hechos)} de hechos + {len(f_br)} de brecha) | '
          f'Indicadores {n_princ + n_idx} indicadores | Ficha {len(fichas)} + 2 eliminados | Índices 4 x 7 años')
    print('hojas:', wl.sheetnames)


if __name__ == '__main__':
    main()
