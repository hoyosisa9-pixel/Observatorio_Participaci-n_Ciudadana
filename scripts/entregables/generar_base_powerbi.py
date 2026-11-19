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

Se corre desde la raíz: python3 scripts/entregables/generar_base_powerbi.py
Cuando se completen las PQRSDF (cs_descargas.py + cs_icrcs.py), basta con volver a correrlo.
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
NOMBRE_TABLERO = {
    'PDM.ICPD7': 'Índice de cumplimiento del PDM (ICPD-7)', 'PDM.SALUD': 'Cobertura en salud', 'PDM.ILE': 'Logro educativo',
    'PDM.SP': 'Servicios públicos domiciliarios', 'PDM.CM': 'Mortalidad vial', 'PDM.CA': 'Espacio público verde urbano',
    'PDM.VEED': 'Veedurías ciudadanas inscritas', 'PDM.HOM': 'Homicidio',
    'PC.ISPC': 'Índice de participación ciudadana (ISPC)', 'PC.I1': 'Votación del presupuesto participativo',
    'PC.I2': 'Cobertura territorial de la participación', 'PC.I3': 'Cumplimiento de la oferta formativa',
    'PC.I4': 'Publicación de resultados participativos',
    'PP.IREPP': 'Índice de presupuesto participativo (IREPP)', 'PP.EPP': 'Ejecución presupuestal del PP',
    'PP.EJ90': 'Proyectos con ejecución ≥ 90 %', 'PP.FID': 'Fidelidad a lo votado',
    'CS.ICRCS': 'Índice de control social (ICRCS)', 'CS.PQRSDF': 'Oportunidad de respuesta a las PQRSDF',
    'CS.DSI': 'Disponibilidad de información de seguimiento', 'CS.VRC': 'Veedurías en la rendición de cuentas',
    'CS.RPO': 'Respuesta a observaciones ciudadanas',
}
SEMAFORO = [('Verde', 90, 100, '#2E7D32', 1), ('Amarillo', 75, 90, '#F9A825', 2),
            ('Naranja', 50, 75, '#EF6C00', 3), ('Rojo', 0, 50, '#C62828', 4)]

ARIAL = 'Arial'
ENC_FILL = PatternFill('solid', start_color='1F4E78')
ENC_FONT = Font(name=ARIAL, bold=True, color='FFFFFF', size=10)
CUERPO = Font(name=ARIAL, size=10)


# ---------------------------------------------------------------- textos: tildes
# Tabla visible de reemplazos (palabras completas, sin distinguir mayúsculas). Se aplica solo al exportar
# (indicador, nota, brecha, referente y fuente); los CSV de construcción no se tocan ni se aplica dentro de URL.
TILDES = {
    'anios': 'años', 'anio': 'año', 'poblacion': 'población', 'alcaldia': 'Alcaldía', 'regimen': 'régimen',
    'regimenes': 'regímenes', 'educacion': 'educación', 'participacion': 'participación', 'informacion': 'información',
    'publicos': 'públicos', 'publicas': 'públicas', 'publicacion': 'publicación', 'veedurias': 'veedurías',
    'veeduria': 'veeduría', 'rendicion': 'rendición', 'ejecucion': 'ejecución', 'votacion': 'votación',
    'indice': 'índice', 'indices': 'índices', 'dimension': 'dimensión', 'formacion': 'formación', 'atencion': 'atención',
    'ninos': 'niños', 'ninas': 'niñas', 'minimo': 'mínimo', 'minimos': 'mínimos', 'maximo': 'máximo', 'maximos': 'máximos',
    'numero': 'número', 'diagnostico': 'diagnóstico', 'revision': 'revisión', 'gestion': 'gestión',
    'metodologia': 'metodología', 'segun': 'según', 'tambien': 'también', 'despues': 'después', 'ademas': 'además',
    'analisis': 'análisis', 'policia': 'Policía', 'secretaria': 'Secretaría', 'planeacion': 'Planeación',
    'personeria': 'Personería', 'area': 'área', 'terminos': 'términos', 'termino': 'término', 'proyeccion': 'proyección',
    'proyecciones': 'proyecciones', 'definicion': 'definición', 'formula': 'fórmula', 'contracredito': 'contracrédito',
    'apropiacion': 'apropiación', 'obligacion': 'obligación', 'unico': 'único', 'unica': 'única', 'diseno': 'diseño',
    'reunion': 'reunión', 'comite': 'comité', 'caida': 'caída', 'categoria': 'categoría', 'intervencion': 'intervención',
    'intervenciones': 'intervenciones', 'felicitacion': 'felicitación', 'felicitaciones': 'felicitaciones',
    'invitacion': 'invitación', 'remision': 'remisión', 'pagina': 'página', 'paginas': 'páginas',
    'ciudadania': 'ciudadanía', 'politica': 'política', 'politicas': 'políticas', 'ultimo': 'último', 'ultima': 'última',
    'habitat': 'hábitat', 'fisico': 'físico', 'deficit': 'déficit', 'liquidacion': 'liquidación', 'razon': 'razón',
    'verificacion': 'verificación', 'clasificacion': 'clasificación', 'afiliacion': 'afiliación', 'cobertura': 'cobertura',
    'metrica': 'métrica', 'tecnica': 'técnica', 'tecnico': 'técnico', 'economia': 'economía', 'demografica': 'demográfica',
    'accion': 'acción', 'solucion': 'solución', 'funcion': 'función', 'orden': 'orden', 'ocupacion': 'ocupación',
    'cuadros': 'cuadros', 'evaluacion': 'evaluación', 'asignacion': 'asignación', 'estimacion': 'estimación',
    'ubicacion': 'ubicación', 'situacion': 'situación', 'direccion': 'dirección', 'resolucion': 'resolución',
    'comparacion': 'comparación', 'conversion': 'conversión', 'seccion': 'sección', 'sesion': 'sesión', 'sesiones': 'sesiones',
    'audiencia': 'audiencia', 'matricula': 'matrícula', 'relacion': 'relación', 'tenia': 'tenía', 'habia': 'había',
    'mayoria': 'mayoría', 'energia': 'energía', 'mineducacion': 'MinEducación', 'apendice': 'apéndice',
    'descripcion': 'descripción', 'programacion': 'programación', 'priorizacion': 'priorización', 'seria': 'sería',
    'daria': 'daría', 'mediria': 'mediría', 'implicaria': 'implicaría', 'pasaria': 'pasaría', 'podria': 'podría',
    'reduccion': 'reducción', 'convencion': 'convención', 'retroproyeccion': 'retroproyección', 'construccion': 'construcción',
    'medicion': 'medición', 'decision': 'decisión', 'auditoria': 'auditoría', 'subdireccion': 'subdirección',
    'mencion': 'mención', 'identicos': 'idénticos', 'delimitacion': 'delimitación', 'presentacion': 'presentación',
    'formulacion': 'formulación', 'contradiccion': 'contradicción', 'transicion': 'transición', 'inscripcion': 'inscripción',
    'observacion': 'observación', 'version': 'versión', 'excluia': 'excluía', 'electrica': 'eléctrica', 'republica': 'República',
    'asi': 'así', 'capacitacion': 'capacitación', 'actualizacion': 'actualización', 'proporcion': 'proporción', 'grafico': 'gráfico', 'organizaciones': 'organizaciones', 'sera': 'será', 'unicamente': 'únicamente', 'ultimos': 'últimos', 'numeros': 'números', 'biologica': 'biológica', 'ecologica': 'ecológica', 'tematica': 'temática', 'cronologico': 'cronológico',
}
TILDES = {k: v for k, v in TILDES.items() if k != v}
_RE_URL = re.compile(r'(https?://\S+)')
_RE_PAL = re.compile(r'\b(' + '|'.join(sorted(TILDES, key=len, reverse=True)) + r')\b', re.IGNORECASE)
_RE_ANOS = re.compile(r'(\d)(\s*)anos\b')


def _cambia(m):
    original, nuevo = m.group(0), TILDES[m.group(0).lower()]
    if original.isupper() and len(original) > 1:
        return nuevo.upper()
    return nuevo[0].upper() + nuevo[1:] if original[0].isupper() else nuevo


def tildes(texto):
    """Reemplazo controlado de palabras completas sin tilde; no toca las URL."""
    if not texto:
        return texto
    trozos = _RE_URL.split(texto)
    for i in range(0, len(trozos), 2):
        t = _RE_ANOS.sub(r'\1\2años', trozos[i])
        trozos[i] = _RE_PAL.sub(_cambia, t)
    return ''.join(trozos)


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
                r[campo] = tildes(r[campo])
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
PCT, PUNT, TASA = '%', 'puntaje 0–100', 'tasa por 100.000 hab.'
UNIDAD_BASE = {
    'PDM.ICPD7': PUNT, 'PDM.SALUD': PCT, 'PDM.ILE': PUNT, 'PDM.SP': PCT, 'PDM.CM': TASA, 'PDM.CA': 'm²/hab',
    'PDM.VEED': 'conteo (veedurías inscritas)', 'PDM.HOM': TASA, 'PC.ISPC': PUNT, 'PC.I1': '% (IPCM: votos / población)',
    'PC.I2': PCT, 'PC.I3': PCT, 'PC.I4': PCT, 'PP.IREPP': PUNT, 'PP.EPP': PCT, 'PP.EJ90': PCT, 'PP.FID': PCT,
    'CS.ICRCS': PUNT, 'CS.PQRSDF': PCT, 'CS.DSI': PCT, 'CS.VRC': PCT, 'CS.RPO': PCT,
}
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
MENOR_ES_MEJOR = {'PDM.CM', 'PDM.HOM', 'PDM.CM.TD', 'PDM.HOM.TD'}
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
        unidad = 'años sin dato / años del periodo (%)' if tipo == 'brecha' else (UNIDAD_COD.get(c) or UNIDAD_BASE.get(base, PCT))
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
    assert all(x[5] is None for x in f_ind if x[0] == 'CS'), 'ICRCS debería ser N/D mientras falte PQRSDF'
    assert len([x for x in f_ind if x[0] == 'CS']) == 6 and len([x for x in f_ind if x[0] == 'PDM']) == 7

    # ---------------- libro
    wb = Workbook()
    ws = wb.active
    ws.title = 'LEEME'
    texto = [
        ('Base de datos de indicadores del Observatorio Ciudadano de Envigado (OCE)', True),
        (f'Generada el {FECHA}. Todos los resultados son PROVISIONALES.', False),
        ('', False),
        ('Qué contiene', True),
        ('hechos: una fila por indicador y periodo (año o trimestre). Es la tabla principal del tablero.', False),
        ('indices: los cuatro índices por año, hasta 2025 (ICPD-7 desde 2019; ISPC, IREPP e ICRCS desde 2020). Atajo de «hechos» con tipo = indice.', False),
        ('brechas: brecha de información de cada indicador e índice en su periodo completo.', False),
        ('dim_dimension, dim_indicador, dim_anio, dim_semaforo: tablas de dimensiones para relacionar y filtrar.', False),
        ('fuentes: documentos y enlaces citados. diccionario: qué significa cada columna.', False),
        ('', False),
        ('Convención de años', True),
        ('PP: el año es el de la vigencia (el año en que se ejecutan los proyectos).', False),
        ('PC.I1 (votación): el año es el de la votación; los proyectos se ejecutan en t+1.', False),
        ('Los demás indicadores: el año del dato. PC.I4: año de la vigencia. PQRSDF: año del trimestre.', False),
        ('Periodos: el ICPD-7 va desde 2019; ISPC, IREPP e ICRCS, desde 2020. Los indicadores traen todos los años que tengan (p. ej. salud desde 2012).', False),
        ('2026 está en curso: las filas de 2026 llevan vigencia_en_curso = Sí y no entran a «indices».', False),
        ('', False),
        ('Relaciones sugeridas en Power BI (modelo estrella)', True),
        ('hechos[codigo] → dim_indicador[codigo] (muchos a uno)', False),
        ('dim_indicador[dimension] → dim_dimension[dimension] (muchos a uno). Es la única ruta hacia dim_dimension: hechos → dim_indicador → dim_dimension.', False),
        ('No cree una relación directa hechos[dimension] → dim_dimension[dimension]: crearía una ruta ambigua (si la quiere, déjela inactiva y úsela con USERELATIONSHIP).', False),
        ('hechos[anio] → dim_anio[anio] (muchos a uno)', False),
        ('hechos[semaforo] → dim_semaforo[semaforo] (muchos a uno; filas sin semáforo quedan en blanco)', False),
        ('indices[codigo] y brechas[codigo] → dim_indicador[codigo]; indices[anio] → dim_anio[anio]; indices[semaforo] → dim_semaforo[semaforo].', False),
        ('', False),
        ('Cómo leer los datos', True),
        ('tipo: indice (el índice de la dimensión), principal (componente que entra al índice), subcomponente, respaldo, detalle o contexto (no entran al índice).', False),
        ('Para el tablero principal filtre tipo = indice o tipo = principal. Los demás tipos sirven para el detalle y la trazabilidad.', False),
        ('tiene_dato = Sí si hay número o si la columna brecha empieza por «con dato». valor vacío + tiene_dato = No: no hay dato publicado para ese periodo (ver brecha y nota). No es un cero.', False),
        ('valor_texto: texto original cuando no hay número (N/D, no aplica, pendiente, con veeduría, etc.).', False),
        ('valor_min / valor_max: rango de VRC, RPO, DSI, I4 (2025) y PQRSDF. valor es el mínimo; valor_max sale de la fila .MAX del mismo año.', False),
        ('puntaje_0_100: el puntaje que entra al índice. Semáforo (propuesta, para validar con el profe): Verde ≥ 90, Amarillo 75 a < 90, Naranja 50 a < 75, Rojo < 50.', False),
        ('subperiodo: T1–T4 o S1–S2 en Control social (informes trimestrales de PQRSDF). Para series anuales filtre tipo_periodo = Año.', False),
        ('dim_indicador: unidad, sentido (mayor o menor es mejor), orden_dimension, orden_indicador y codigo_base (indicador al que pertenece cada respaldo, subcomponente o detalle).', False),
        ('', False),
        ('Medidas DAX sugeridas', True),
        ('Valor índice = SELECTEDVALUE(indices[valor])', False),
        ('Color semáforo (propuesta) = LOOKUPVALUE(dim_semaforo[color_hex], dim_semaforo[semaforo], SELECTEDVALUE(indices[semaforo]))', False),
        ('% años con dato = DIVIDE(COUNTROWS(FILTER(indices, indices[tiene_dato] = "Sí")), COUNTROWS(indices))', False),
        ('Brecha de información = 1 - [% años con dato]', False),
        ('', False),
        ('Advertencias', True),
        ('Control social (ICRCS) está sin dato en todos los años: faltan los informes de PQRSDF porque el servidor de archivos de la Alcaldía no responde desde el 19-nov-2026. La nota del índice trae el rango posible.', False),
        ('Los cortes de espacio verde (2019, 2022, 2024) son de estudios distintos: no son una tendencia.', False),
        ('Cada fila trae su fuente, enlace y fecha de consulta. Los datos se pueden rastrear hasta el documento original.', False),
        ('Los textos de indicador, nota, brecha, referente y fuente se exportan con tildes (tabla TILDES del script); los CSV de construcción no se modifican.', False),
        ('', False),
        ('Origen técnico', True),
        ('Generado por scripts/entregables/generar_base_powerbi.py a partir de outputs/construccion/*/..._indicadores.csv.', False),
    ]
    for t, negrita in texto:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(name=ARIAL, size=12 if ws.max_row == 1 else 10, bold=negrita)
    ws.column_dimensions['A'].width = 130

    hoja(wb, 'hechos', col_hechos, f_hechos,
         formatos={'numerador': '#,##0.##', 'denominador': '#,##0.##', 'valor': '0.00', 'valor_min': '0.00', 'valor_max': '0.00',
                   'puntaje_0_100': '0.00', 'peso_en_indice': '0.000', 'anio': '0', 'fecha_consulta_fecha': 'yyyy-mm-dd'},
         anchos={'indicador': 45, 'nota': 60, 'fuente': 50, 'enlace': 50, 'enlace_principal': 50, 'brecha': 40, 'referente': 30})
    hoja(wb, 'indices', ['dimension', 'nombre_dimension', 'codigo', 'indice', 'anio', 'valor', 'semaforo', 'tiene_dato', 'vigencia_en_curso', 'nota'],
         f_ind, formatos={'valor': '0.00', 'anio': '0'}, anchos={'nota': 70, 'indice': 40})
    hoja(wb, 'brechas', ['dimension', 'codigo', 'indicador', 'tipo', 'periodo', 'sin_dato', 'total', 'brecha_pct', 'brecha', 'nota'], f_br,
         formatos={'brecha_pct': '0.00'}, anchos={'indicador': 45, 'brecha': 45, 'nota': 60})
    hoja(wb, 'dim_dimension', ['dimension', 'nombre_dimension', 'codigo_indice', 'indice', 'formula_indice', 'orden_dimension'],
         [[d[0], d[1], d[2], d[3], d[4], ORDEN_DIM[d[0]]] for d in DIMENSIONES], anchos={'formula_indice': 80})
    hoja(wb, 'dim_indicador', col_cat, dim_ind, formatos={'peso_en_indice': '0.000'},
         anchos={'indicador': 50, 'nombre_tablero': 45, 'referente': 40, 'unidad': 30})
    hoja(wb, 'dim_anio', ['anio', 'periodo_pdm'],
         [[a, 'PDM 2024-2027' if a >= 2024 else ('PDM 2020-2023' if a >= 2020 else ('PDM 2016-2019' if a >= 2016 else 'Anterior'))]
          for a in anios_dim], formatos={'anio': '0'})
    hoja(wb, 'dim_semaforo', ['semaforo', 'desde', 'hasta', 'color_hex', 'orden', 'estado'],
         [list(s) + ['Propuesta (D60, D84): por validar con el profe'] for s in SEMAFORO], anchos={'estado': 45})
    hoja(wb, 'fuentes', ['id_fuente', 'fuente', 'enlace', 'dimensiones', 'filas_que_la_citan', 'fecha_consulta'], f_fu,
         anchos={'fuente': 80, 'enlace': 70})

    dicc = [
        ('hechos', 'id_hecho', 'Clave única: código|periodo (en filas de detalle se agrega |n, una por elemento: meta, evento o trimestre)'),
        ('hechos', 'dimension', 'PDM, PC, PP o CS (ver dim_dimension)'),
        ('hechos', 'codigo', 'Código único del indicador (ver dim_indicador)'),
        ('hechos', 'indicador', 'Nombre del indicador'),
        ('hechos', 'tipo', 'indice, principal, subcomponente, respaldo, detalle o contexto'),
        ('hechos', 'anio', 'Año (número). En PQRSDF trimestral es el año del trimestre. PP = vigencia; PC.I1 = año de la votación'),
        ('hechos', 'subperiodo', 'T1–T4 o S1–S2 cuando el dato es trimestral o semestral; vacío si es anual'),
        ('hechos', 'tipo_periodo', 'Año, Trimestre o Semestre'),
        ('hechos', 'periodo', 'Periodo tal como viene en la fuente (2024, 2023-T4)'),
        ('hechos', 'numerador / denominador', 'Variables del cálculo (conteos, pesos, habitantes, etc.)'),
        ('hechos', 'valor', 'Valor del indicador (número). En los indicadores en rango es el mínimo. Vacío si no hay dato'),
        ('hechos', 'valor_min / valor_max', 'Rango mínimo y máximo de VRC, RPO, DSI, I4 (2025) y PQRSDF. El máximo sale de la fila .MAX del mismo año; vacío si no hay rango'),
        ('hechos', 'valor_texto', 'Texto original cuando no hay número (N/D, no aplica, pendiente, con veeduría…)'),
        ('hechos', 'puntaje_0_100', 'Puntaje normalizado que entra al índice'),
        ('hechos', 'peso_en_indice', 'Peso del componente en el índice de su dimensión'),
        ('hechos', 'semaforo', 'Verde, Amarillo, Naranja o Rojo (propuesta; ver dim_semaforo)'),
        ('hechos', 'tiene_dato', 'Sí si hay número o si la brecha empieza por «con dato» (clasificaciones con dato); No en N/D, pendiente, desconocida o sin dato'),
        ('hechos', 'brecha', '«con dato» o «sin dato: motivo»'),
        ('hechos', 'referente', 'Meta, línea base o referente usado para el puntaje'),
        ('hechos', 'fuente / enlace / fecha_consulta', 'Trazabilidad del dato (texto con tildes; el enlace puede traer varias URL)'),
        ('hechos', 'enlace_principal', 'Primera URL de la columna enlace; vacía si no hay'),
        ('hechos', 'fecha_consulta_fecha', 'Fecha real (AAAA-MM-DD) tomada de fecha_consulta; la primera si hay varias; vacía si el texto no trae fecha'),
        ('hechos', 'estado', 'Provisional'),
        ('hechos', 'vigencia_en_curso', 'Sí si el año (2026) está en curso y no tiene cierre; No en los demás'),
        ('hechos', 'nota', 'Método, supuestos y alertas'),
        ('indices', '(todas)', 'Igual que hechos, solo para los cuatro índices y hasta 2025; vigencia_en_curso siempre es No'),
        ('brechas', 'sin_dato / total / brecha_pct', 'Años (o trimestres, eventos, metas) sin dato, total del periodo y porcentaje'),
        ('dim_indicador', 'entra_al_indice', 'Sí para índices y componentes principales'),
        ('dim_indicador', 'nombre_tablero', 'Nombre corto para mostrar en el tablero (índices y componentes principales)'),
        ('dim_indicador', 'unidad', 'Unidad del valor: %, tasa por 100.000 hab., m²/hab, conteo, puntaje 0–100, habitantes, clasificación (texto)'),
        ('dim_indicador', 'sentido', 'mayor es mejor, menor es mejor (mortalidad vial, homicidio y brechas) o no aplica (detalles y conteos de apoyo)'),
        ('dim_indicador', 'orden_dimension / orden_indicador', 'Orden para ordenar el tablero: PDM 1, PC 2, PP 3, CS 4; el índice primero y después sus componentes, respaldos y detalles'),
        ('dim_indicador', 'codigo_base', 'Código del indicador al que pertenece la fila (sin sufijos .MAX, .BRECHA, .TRIM, .META, .EVENTO, .TODO, .AMPLIO, etc.)'),
        ('dim_indicador', 'primer/ultimo_anio_con_dato, anios_con_dato', 'Cobertura temporal real del indicador'),
        ('dim_dimension', 'orden_dimension', 'Orden de la dimensión en el tablero'),
        ('dim_semaforo', 'desde / hasta', 'Rango del puntaje (desde inclusive, hasta exclusivo; Verde incluye 100)'),
        ('dim_semaforo', 'estado', 'Propuesta (D60, D84): los rangos se validan con el profe'),
    ]
    hoja(wb, 'diccionario', ['tabla', 'columna', 'descripcion'], [list(d) for d in dicc], anchos={'descripcion': 90})

    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    wb.save(SALIDA)
    print(f'{SALIDA.relative_to(RAIZ)} | hechos {len(f_hechos)} ({len(col_hechos)} columnas) | indices {len(f_ind)} | brechas {len(f_br)} | '
          f'indicadores {len(dim_ind)} ({len(col_cat)} columnas) | fuentes {len(f_fu)}')


if __name__ == '__main__':
    main()
