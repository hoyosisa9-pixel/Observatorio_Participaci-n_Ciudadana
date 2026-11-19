"""Hojas del Excel del OCE (Inicio, Índices, Indicadores, Ficha metodológica). La hoja Datos la arma generar_base_powerbi.py.

Todo se lee de los CSV de construcción (vía comun_entregas y contenido_fichas); no hay cifras escritas a mano.
"""
import math
import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import comun_entregas as ce
import contenido_fichas as cf

ARIAL = 'Arial'
AZUL = '1F4E78'
SOFT = {'Verde': 'C8E6C9', 'Amarillo': 'FFF3B0', 'Naranja': 'FFD8A8', 'Rojo': 'F5B7B1'}
LADO = Side(style='thin', color='D0D0D0')
BORDE = Border(left=LADO, right=LADO, top=LADO, bottom=LADO)
F_CUERPO = Font(name=ARIAL, size=10)
F_NEG = Font(name=ARIAL, size=10, bold=True)
F_ENC = Font(name=ARIAL, size=10, bold=True, color='FFFFFF')
F_GRIS = Font(name=ARIAL, size=9, italic=True, color='7F7F7F')
F_TIT = Font(name=ARIAL, size=14, bold=True, color=AZUL)
F_SUB = Font(name=ARIAL, size=11, bold=True, color=AZUL)
RELLENO_ENC = PatternFill('solid', start_color=AZUL)
RELLENO_GRUPO = PatternFill('solid', start_color='D9E2F3')
RELLENO_SUAVE = PatternFill('solid', start_color='F2F2F2')
ANIOS_VISTA = list(range(2012, 2026))   # años de la hoja Indicadores (los años anteriores están en Datos)
DIM_TXT = {'PDM': 'Cumplimiento del Plan de Desarrollo Municipal (PDM)', 'PC': 'Participación ciudadana',
           'PP': 'Presupuesto participativo', 'CS': 'Control social'}


def horizontal(ws, titulos_filas=None):
    """Impresión horizontal ajustada al ancho."""
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    if titulos_filas:
        ws.print_title_rows = titulos_filas


def altura(ws, fila, anchos, minimo=16, tope=330):
    """Altura estimada de la fila según el texto con ajuste de línea de cada columna."""
    lineas = 1
    for j, a in anchos.items():
        v = ws.cell(fila, j).value
        if isinstance(v, str):
            n = sum(max(1, math.ceil(len(p) / max(a * 1.05, 1))) for p in v.split('\n'))
            lineas = max(lineas, n)
    ws.row_dimensions[fila].height = min(max(minimo, lineas * 12.6 + 4), tope)


def encabezado(ws, fila, textos, desde=1):
    for j, t in enumerate(textos, desde):
        c = ws.cell(fila, j, t)
        c.fill, c.font, c.border = RELLENO_ENC, F_ENC, BORDE
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[fila].height = 32


def corto(texto, n=110):
    t = re.split(r'\s\(|;|\s\+\s', texto or '')[0].strip()
    return t if len(t) <= n else t[:n].rsplit(' ', 1)[0] + '…'


def fmt_num(v, dec=2):
    return ce.fmt(v, dec)


# ------------------------------------------------------------------------------------------------ INICIO
def hoja_inicio(wb, fecha, col_datos, dic_datos):
    ws = wb.active
    ws.title = 'Inicio'
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 118
    fila = [1]

    def linea(a, b='', estilo=None, relleno=None, altura_fija=None):
        r = fila[0]
        ca, cb = ws.cell(r, 1, a), ws.cell(r, 2, b)
        for c in (ca, cb):
            c.font = F_CUERPO
            c.alignment = Alignment(wrap_text=True, vertical='top')
        if estilo == 'titulo':
            ca.font = F_TIT
            ca.alignment = Alignment(wrap_text=False)
        elif estilo == 'seccion':
            for c in (ca, cb):
                c.fill, c.font = RELLENO_GRUPO, F_SUB
        elif estilo == 'etiqueta':
            ca.font = F_NEG
        if relleno:
            ca.fill = PatternFill('solid', start_color=relleno)
            ca.alignment = Alignment(horizontal='center', vertical='center')
            ca.font = F_NEG
        if estilo != 'titulo':
            altura(ws, r, {1: 30, 2: 118}, minimo=16)
        fila[0] += 1

    linea('Base de datos de indicadores del Observatorio Ciudadano de Envigado (OCE)', estilo='titulo')
    linea(f'Generada el {fecha}. Todos los resultados son PROVISIONALES.', estilo=None)
    linea('')
    linea('Qué es este archivo', estilo='seccion')
    linea('', 'Reúne los cuatro índices del observatorio (cumplimiento del PDM, participación ciudadana, presupuesto participativo y control social) y los indicadores que los componen, con su serie por año.')
    linea('', 'Es completo: cada cifra trae su fuente, su enlace y su fecha de consulta, y cada indicador tiene su ficha metodológica y lo que cambió frente a la matriz original.')
    linea('', 'Sirve para leerlo en Excel (hojas 2 a 4) y para alimentar el tablero de Power BI (hoja 5).')
    linea('')
    linea('Cómo leer cada hoja', estilo='seccion')
    linea('1. Inicio', 'Esta hoja: qué es el archivo, el semáforo, el glosario, una nota para Power BI y el diccionario de columnas de «Datos».', estilo='etiqueta')
    linea('2. Índices', 'Los cuatro índices por año (2019–2025), con el color del semáforo (propuesta), la brecha de información y el máximo cuando hay rango. Arriba va la fórmula de cada uno.', estilo='etiqueta')
    linea('3. Indicadores', 'La hoja principal para leer: cada índice y sus componentes, agrupados por dimensión. Muestra el puntaje de 0 a 100 por año con color; debajo, el valor original o el máximo del rango. '
                            'Incluye unidad, sentido, referente, peso, brecha y fuente.', estilo='etiqueta')
    linea('4. Ficha metodológica', 'Una fila por indicador y por índice: definición, fórmula, variables y fuentes, línea base y referente, periodicidad, tratamiento de datos faltantes, limitaciones y qué cambió frente a la matriz original (incluye los indicadores eliminados).', estilo='etiqueta')
    linea('5. Datos', 'La tabla larga para Power BI: una fila por indicador y periodo (también subcomponentes, respaldos, detalle y brechas; use la columna «tipo» para filtrar). Está en un único formato de tabla de Excel.', estilo='etiqueta')
    linea('')
    linea('Semáforo (propuesta, por validar con el profesor)', estilo='seccion')
    descr = {'Verde': 'Puntaje de 90 a 100', 'Amarillo': 'Puntaje de 75 a menos de 90', 'Naranja': 'Puntaje de 50 a menos de 75', 'Rojo': 'Puntaje menor de 50'}
    for nombre, _d, _h, _c, _o in ce.SEMAFORO:
        linea(nombre, descr[nombre], relleno=SOFT[nombre])
    linea('', 'Se aplica al puntaje de 0 a 100, igual en las cuatro dimensiones. En los indicadores con rango se asigna con el mínimo. En «Datos» la columna color_semaforo trae el color en hexadecimal para el tablero.')
    linea('')
    linea('Glosario', estilo='seccion')
    linea('Puntaje de 0 a 100', 'Cada indicador se compara con su referente (línea base y meta) y se convierte a una escala de 0 a 100; el índice combina los puntajes de sus componentes. Un índice solo se calcula si están todos sus componentes: no se redistribuyen pesos.', estilo='etiqueta')
    linea('Brecha de información', 'Parte del periodo (años, trimestres, eventos o metas) que no tiene dato publicado o verificable. Una brecha alta no es mal desempeño: es que no se puede medir. Se reporta siempre.', estilo='etiqueta')
    linea('Rango mínimo–máximo', 'Cuando un dato no se puede verificar del todo, se da un mínimo (solo lo documentado, que es el que entra al índice) y un máximo (suma lo no verificable). La diferencia es la brecha.', estilo='etiqueta')
    linea('«Sin dato» no es cero', 'Una celda vacía o «—» significa que no hay dato (o que no aplica), no que el resultado sea 0. En «Datos», tiene_dato = No.', estilo='etiqueta')
    linea('Convención de años', 'Presupuesto participativo: año de la vigencia. Votación del presupuesto participativo (I1): año de la votación (los proyectos se ejecutan al año siguiente). Los demás: año del dato. '
                               'El año en curso (2026) no se califica. Los cortes de espacio verde son de estudios distintos: no son una tendencia.', estilo='etiqueta')
    linea('')
    linea('Nota para Power BI', estilo='seccion')
    linea('', 'Importe solo la hoja «Datos» (tabla «Datos»); las demás hojas son para leer y no necesitan importarse.')
    linea('', 'Claves: codigo (indicador o índice), anio (año; vacío en las filas de brecha, que usan el texto de «periodo») y codigo_indice (índice de la dimensión). Para el tablero principal filtre tipo = indice o principal; '
              'los demás tipos (subcomponente, respaldo, detalle, contexto y brecha) sirven para el detalle.')
    linea('', 'Use puntaje_0_100 y semaforo / color_semaforo para colorear; valor_min y valor_max para los rangos; tiene_dato para distinguir «sin dato» de cero; orden_dimension y orden_indicador para ordenar.')
    linea('')
    linea('Diccionario de columnas de «Datos»', estilo='seccion')
    for col, desc in dic_datos:
        linea(col, desc, estilo='etiqueta')
    linea('')
    linea('Cómo se generó', estilo='seccion')
    linea('', 'Con scripts/entregables/generar_base_powerbi.py, a partir de los resultados de la construcción (outputs/construccion/*/..._indicadores.csv). Si los resultados se actualizan, se vuelve a correr y todas las hojas se actualizan.')
    horizontal(ws)
    return ws


# ------------------------------------------------------------------------------------------------ ÍNDICES
def hoja_indices(wb, fichas, maximos):
    ws = wb.create_sheet('Índices')
    ws.sheet_view.showGridLines = False
    cods = list(ce.NOMBRE_INDICE)
    anios = list(range(2019, 2026))
    fx = {f['codigo']: f for f in fichas}
    ws['A1'] = 'Los cuatro índices por año (puntaje de 0 a 100)'
    ws['A1'].font = F_TIT
    ws['A2'] = 'Resultados PROVISIONALES. Un índice solo se calcula en los años en que están todos sus componentes; «—» = falta algún componente (ver brecha).'
    ws['A2'].font = F_CUERPO
    r = 4
    ws.cell(r, 1, 'Fórmula de cada índice').font = F_SUB
    for cod in cods:
        r += 1
        c = ws.cell(r, 1, f"{ce.NOMBRE_INDICE[cod]}:  {fx[cod]['formula'].replace(chr(10), ' ')}")
        c.font = F_CUERPO
    r += 2
    ws.cell(r, 1, 'Semáforo (propuesta):').font = F_NEG
    for j, (n, d, h, _c, _o) in enumerate(ce.SEMAFORO):
        c = ws.cell(r, 2 + j, f'{n}  {"≥ 90" if n == "Verde" else f"{d} a < {h}" if n != "Rojo" else "< 50"}')
        c.fill, c.font, c.alignment, c.border = PatternFill('solid', start_color=SOFT[n]), F_CUERPO, Alignment(horizontal='center'), BORDE
    r += 2
    enc = ['Índice', 'Dimensión'] + [str(a) for a in anios] + ['Años con índice', 'Brecha de información', 'Máximo (cuando hay rango)']
    encabezado(ws, r, enc)
    cab = r
    for cod in cods:
        r += 1
        dim = ce.dim_de(cod)
        vals = ce.valores_indice(cod)
        ws.cell(r, 1, ce.NOMBRE_INDICE[cod]).font = F_NEG
        ws.cell(r, 2, DIM_TXT[dim]).font = F_CUERPO
        for j, a in enumerate(anios):
            v, s = vals.get(a, (None, ''))
            c = ws.cell(r, 3 + j)
            if v is None:
                c.value, c.font = '—', F_GRIS
            else:
                c.value, c.number_format, c.font = v, '0.00', F_NEG
                c.fill = PatternFill('solid', start_color=SOFT[s])
            c.alignment = Alignment(horizontal='center', vertical='center')
        k = 3 + len(anios)
        ws.cell(r, k, ', '.join(str(a) for a in ce.anios_con_indice(cod))).font = F_CUERPO
        ws.cell(r, k + 1, ce.brecha_texto(dim, cod)).font = F_CUERPO
        mx = maximos.get(cod, {})
        ws.cell(r, k + 2, '; '.join(f'{a}: {fmt_num(v)}' for a, v in sorted(mx.items())) or '—').font = F_CUERPO
        for j in range(1, k + 3):
            c = ws.cell(r, j)
            c.border = BORDE
            if j in (1, 2, k, k + 1, k + 2):
                c.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[r].height = 34
    anchos = {1: 12, 2: 34, k: 22, k + 1: 36, k + 2: 26}
    for j in range(1, k + 3):
        ws.column_dimensions[get_column_letter(j)].width = anchos.get(j, 9)
    ws.freeze_panes = ws.cell(cab + 1, 3)
    r += 2
    ws.cell(r, 1, 'El detalle de cada componente por año está en la hoja «Indicadores»; la metodología, en «Ficha metodológica».').font = F_GRIS
    horizontal(ws)
    return ws


# ------------------------------------------------------------------------------------------------ INDICADORES
def hoja_indicadores(wb, fichas, serie, nombres, brechas_txt):
    """serie[cod][anio] = (puntaje, semáforo, valor, máximo)."""
    ws = wb.create_sheet('Indicadores')
    ws.sheet_view.showGridLines = False
    fx = {f['codigo']: f for f in fichas}
    ncol_fijas = 6
    cols = ['Indicador', 'Medida', 'Unidad del valor', 'Sentido', 'Referente', 'Peso'] + [str(a) for a in ANIOS_VISTA] + ['Brecha de información', 'Fuente principal', 'Enlace']
    ws['A1'] = 'Indicadores por dimensión: puntaje de 0 a 100 por año'
    ws['A1'].font = F_TIT
    ws['A2'] = ('Cada fila muestra el puntaje (color del semáforo, propuesta). Debajo, si aplica, el valor original en su unidad o el máximo del rango. '
                '«—» = sin dato. Los años anteriores a 2012 (mortalidad vial y homicidio desde 2003) están en la hoja «Datos».')
    ws['A2'].font = F_CUERPO
    encabezado(ws, 4, cols)
    r = 4
    anchos = {1: 36, 2: 15, 3: 20, 4: 14, 5: 40, 6: 8}
    for j in range(len(ANIOS_VISTA)):
        anchos[ncol_fijas + 1 + j] = 7.2
    kb = ncol_fijas + len(ANIOS_VISTA) + 1
    anchos.update({kb: 42, kb + 1: 50, kb + 2: 32})
    for dim in ce.DIM_ORDEN:
        r += 1
        for j in range(1, len(cols) + 1):
            ws.cell(r, j).fill = RELLENO_GRUPO
        ws.cell(r, 1, DIM_TXT[dim]).font = F_SUB
        ws.row_dimensions[r].height = 20
        for cod in [ce.INDICE_DIM[dim]] + ce.principales(dim):
            f = fx[cod]
            es_idx = f['tipo'] == 'Índice'
            s = serie.get(cod, {})
            rango = any(v[3] is not None for v in s.values())
            difiere = any(v[0] is not None and v[2] is not None and abs(v[0] - v[2]) > 0.005 for v in s.values())
            r += 1
            peso = '—'
            if not es_idx:
                p = ce.num([x for x in ce.DATA[dim] if x['codigo'] == cod and x['peso_en_indice']][0]['peso_en_indice'])
                peso = '1/7' if abs(p - 1 / 7) < 1e-4 else f'{p * 100:.0f} %'
            fuente = cf.fuente_serie(cod)
            enlace = ce.enlace_reciente(cod) if not es_idx else ''
            if not es_idx and not enlace:
                enlace = 'ver hoja Datos'
            base = [nombres[cod] if not es_idx else f'{nombres[cod]}', 'Puntaje 0–100', f['unidad'], f['sentido'], f['ref_corto'], peso]
            for j, v in enumerate(base, 1):
                ws.cell(r, j, v)
            for j, a in enumerate(ANIOS_VISTA):
                c = ws.cell(r, ncol_fijas + 1 + j)
                v = s.get(a)
                if v is None or v[0] is None:
                    c.value, c.font = '—', F_GRIS
                else:
                    c.value, c.number_format, c.font = v[0], '0.0', F_NEG
                    if v[1] in SOFT:
                        c.fill = PatternFill('solid', start_color=SOFT[v[1]])
                c.alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(r, kb, brechas_txt[cod])
            ws.cell(r, kb + 1, fuente)
            if enlace:
                c = ws.cell(r, kb + 2, enlace)
                if enlace.startswith('http'):
                    c.hyperlink = enlace
            for j in range(1, len(cols) + 1):
                c = ws.cell(r, j)
                c.border = BORDE
                if j <= ncol_fijas or j >= kb:
                    c.font = F_NEG if j == 1 else (Font(name=ARIAL, size=9, color='0563C1', underline='single') if j == kb + 2 and enlace.startswith('http') else F_CUERPO)
                    c.alignment = Alignment(wrap_text=True, vertical='center')
            altura(ws, r, {1: 36, 5: 40, kb: 42, kb + 1: 50}, minimo=28)
            if rango or difiere:
                r += 1
                etiqueta = 'Máximo (rango)' if rango else 'Valor original'
                ws.cell(r, 2, etiqueta)
                ws.cell(r, 3, f['unidad'] if not rango else '')
                for j, a in enumerate(ANIOS_VISTA):
                    c = ws.cell(r, ncol_fijas + 1 + j)
                    v = s.get(a)
                    x = (v[3] if rango else v[2]) if v else None
                    if x is not None:
                        c.value, c.number_format = x, '0.00'
                    c.alignment = Alignment(horizontal='center')
                for j in range(1, len(cols) + 1):
                    c = ws.cell(r, j)
                    c.border = BORDE
                    c.font = F_GRIS if j <= 2 else Font(name=ARIAL, size=9, color='595959')
                    c.fill = RELLENO_SUAVE
                ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='center')
                ws.row_dimensions[r].height = 18
    for j, a in anchos.items():
        ws.column_dimensions[get_column_letter(j)].width = a
    ws.freeze_panes = ws.cell(5, ncol_fijas + 1)
    r += 2
    ws.cell(r, 1, 'Semáforo (propuesta): Verde ≥ 90 · Amarillo 75 a < 90 · Naranja 50 a < 75 · Rojo < 50. El detalle de cada indicador (subcomponentes, respaldos y trimestres) está en «Datos».').font = F_GRIS
    horizontal(ws, '4:4')
    return ws


# ------------------------------------------------------------------------------------------------ FICHA METODOLÓGICA
COLS_FICHA = [  # rótulo, ancho, función(ficha, cambios) -> texto
    ('Código', 12, lambda f, c: f['codigo']), ('Nombre', 30, lambda f, c: f['nombre']), ('Tipo', 12, lambda f, c: f['tipo']),
    ('Dimensión', 22, lambda f, c: f['dimension']), ('Definición', 55, lambda f, c: f['definicion']),
    ('Fórmula y cálculo del puntaje', 50, lambda f, c: f['formula'] + '\n' + f['puntaje']),
    ('Variables y fuentes', 90, lambda f, c: f['variables'] + '\n\nFuentes:\n' + f['fuentes']),
    ('Línea base y referente', 62, lambda f, c: f['referente']),
    ('Unidad, sentido y peso', 24, lambda f, c: f"{f['unidad']}; {f['sentido']}\nPeso: {f['peso']}"),
    ('Periodicidad y cobertura', 28, lambda f, c: f['periodicidad'] + '\n' + f['cobertura']),
    ('Serie y brecha de información', 38, lambda f, c: f['serie'] + '\n' + f['brecha']),
    ('Tratamiento de datos faltantes, rangos y casos especiales', 66, lambda f, c: f['faltantes']),
    ('Limitaciones y alertas', 66, lambda f, c: f['limitaciones']),
    ('Qué cambió frente a la matriz original', 66, lambda f, c: c.get(f['codigo'], '')),
]


def hoja_ficha(wb, fichas):
    ws = wb.create_sheet('Ficha metodológica')
    ws.sheet_view.showGridLines = False
    cambios, eliminados = cf.cambio_por_codigo()
    encabezado(ws, 1, [c[0] for c in COLS_FICHA])
    anchos = {j: c[1] for j, c in enumerate(COLS_FICHA, 1)}
    r = 1
    for f in fichas:
        r += 1
        for j, (_t, _a, fn) in enumerate(COLS_FICHA, 1):
            c = ws.cell(r, j, fn(f, cambios))
            c.font, c.border = (F_NEG if j <= 2 else F_CUERPO), BORDE
            c.alignment = Alignment(wrap_text=True, vertical='top')
            if f['tipo'] == 'Índice':
                c.fill = RELLENO_GRUPO
        altura(ws, r, anchos, tope=260)
    for dim, orig, nombre, txt in eliminados:
        r += 1
        vals = [orig, nombre, 'Eliminado de la matriz', DIM_TXT[dim]] + [''] * (len(COLS_FICHA) - 5) + [txt]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font, c.border = (F_NEG if j <= 2 else F_CUERPO), BORDE
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.fill = RELLENO_SUAVE
        altura(ws, r, anchos)
    for j, a in anchos.items():
        ws.column_dimensions[get_column_letter(j)].width = a
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS_FICHA))}{r}'
    horizontal(ws, '1:1')
    return ws, r - 1
