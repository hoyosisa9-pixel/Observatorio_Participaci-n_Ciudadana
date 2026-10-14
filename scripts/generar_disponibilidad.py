"""Genera los entregables de la tarea 'Disponibilidad y brecha de informacion del sistema de indicadores'.

Lee los CSV de trabajo de outputs/disponibilidad/trabajo/ (sin modificarlos), aplica las correcciones de la
revision critica de la Fase 3 (revision_fase3.md seccion 1) y escribe:
  - outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx (Lee-me, Variables, Indicadores, Fuentes, Solicitudes)
  - outputs/disponibilidad/disponibilidad_variables.csv (UTF-8, coma)
  - outputs/disponibilidad/disponibilidad_variables_excel.csv (separador ';', UTF-8 con BOM; misma logica de csv_para_excel.py)
Uso: python3 -I scripts/generar_disponibilidad.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
TRABAJO = RAIZ / 'outputs' / 'disponibilidad' / 'trabajo'
SALIDA = RAIZ / 'outputs' / 'disponibilidad'
SISTEMAS = ['CS', 'PC', 'PDM', 'PP']

COLUMNAS_VARIABLES = [
    ('sistema', 'sistema', 14), ('codigo_indicador', 'código indicador', 12), ('id_variable', 'ID variable', 14),
    ('nombre', 'nombre', 40), ('unidad', 'unidad', 12), ('fuente_sugerida', 'fuente sugerida por la matriz', 36),
    ('fuente_encontrada', 'fuente encontrada', 40), ('ruta_o_url', 'ruta en repo o URL', 48),
    ('pagina_tabla', 'página/tabla', 28), ('anios_disponibles', 'años disponibles', 28),
    ('desagregacion', 'desagregación (municipio/zona)', 36), ('formato', 'formato', 14),
    ('disponibilidad', 'disponibilidad', 15), ('valor_encontrado', 'valor encontrado (y año)', 55),
    ('estado_valor', 'estado del valor', 14), ('brecha', 'brecha', 50), ('accion_sugerida', 'acción sugerida', 40),
    ('esfuerzo', 'esfuerzo', 14), ('verificacion', 'verificación', 28), ('notas', 'notas', 50),
]
CATEGORIAS = ['Disponible', 'Construible', 'Parcial', 'Proxy', 'No disponible']
CONTEO_ESPERADO = {'Disponible': 10, 'Construible': 7, 'Parcial': 19, 'Proxy': 2, 'No disponible': 1}


class Anadir(str):
    """Agrega texto al final del valor actual (separado por ' | ')."""


class Prefijar(str):
    """Agrega texto al inicio del valor actual."""


class Sustituir:
    """Reemplaza todas las apariciones de 'viejo' por 'nuevo' en el valor actual."""
    def __init__(self, viejo, nuevo):
        self.viejo, self.nuevo = viejo, nuevo


EPP_PRELIMINAR = '[PRELIMINAR, POR VALIDAR: lista de proyectos PP reconstruida; debe validarla el DAP. EPP 2025 = 98,83 % (APP $20.928.690.218; OPP $20.684.427.720), reproducible con scripts/epp_pp_2025.py]'

# (id_variable, columna, valor_nuevo, motivo). str = reemplazo total; Anadir = agregar; Sustituir = reemplazo parcial.
CORRECCIONES = [
    ('CS.RPO.OT', 'disponibilidad', 'Parcial', 'C1: el protocolo exige matriz publica con todas las observaciones; registro incompleto (3 intervenciones Alcalde 2025), sin cuadros Salud 2024 verificados'),
    ('CS.RPO.OR', 'disponibilidad', 'Parcial', 'C2: OR es subconjunto de OT y hereda el universo incompleto; falta clasificacion "de fondo"'),
    ('CS.RPO.OR', 'brecha', Anadir('2024 (audiencia del Alcalde) construible; 2025 con registro publico incompleto'), 'C2: completar brecha'),
    ('PC.I2.V1', 'disponibilidad', 'Parcial', 'C3: solo cronogramas y reportes agregados; sin el soporte (acta/asistencia por zona) exigido por el protocolo'),
    ('PDM.CS.O', 'disponibilidad', 'Proxy', 'C4: ningun anio tiene la cifra con el metodo del plan; solo definiciones distintas (MinSalud, TerriData)'),
    ('PDM.ILE.S', 'disponibilidad', 'Parcial', 'C5: clasificacion ICFES por plantel no obtenida (URL 404); verificacion "sin verificar"'),
    ('PP.CEP.PEC', 'disponibilidad', 'Parcial', 'C6: autorreporte de ejecucion fisica sin acta de recibo/informe final; regla de coherencia = Parcial'),
    ('PP.CEP.PEC', 'notas', Anadir('Autorreporte; no es entrega verificada.'), 'C6: dejar constancia en notas'),
    ('PP.CEP.PEC', 'valor_encontrado',
     '2025: entre 45 y 55 de 61 proyectos segun la regla de agregacion (49 = todas las filas de diciembre con ejecucion fisica neta = 100 %; 45 = solo filas del 31-dic, 54 proyectos con esa fecha; 55 = ultima fila por indicador). Un proyecto con "Error en la programacion". Autorreporte; regla por fijar. 2024 y 2026: sin cifra',
     'C7: la cifra 49 de 61 no es robusta a la regla de agregacion (verificacion V2)'),
    ('PP.EPP.OPP', 'anios_disponibles',
     '2025 completo (corte 07-01-2026); 2026: APP completa, OPP provisional al 30-jun-2026 (no anual); 2024 aproximado (47 de 50 proyectos, 13 correspondencias aproximadas, 3 sin correspondencia): para 2024 la variable es Parcial',
     'C8: Construible solo vale plenamente para 2025'),
    ('PP.EPP.APP', 'anios_disponibles',
     '2025 completo (corte 07-01-2026); 2026: APP completa, OPP provisional al 30-jun-2026 (no anual); 2024 aproximado (47 de 50 proyectos, 13 correspondencias aproximadas, 3 sin correspondencia): para 2024 la variable es Parcial',
     'C8: Construible solo vale plenamente para 2025'),
    ('PP.EPP.OPP', 'notas',
     Anadir('Proyecto 0553 (costo del proceso) excluido: no figura en el listado EFPC ni es proyecto priorizado; sensibilidad EPP 2025: 98,83 sin 0553 y 98,86 con 0553 (recalculado en Fase 3). CUIPO (-3 %) se usa solo como contraste; la fuente de calculo es el plan de accion/ROCA. Las 3 diferencias con Hacienda 2025 son de lectura del PDF, no del dato fuente.'),
     'C9: decisiones de la seccion 3 y verificacion V1'),
    ('PP.EPP.APP', 'notas',
     Anadir('Proyecto 0553 (costo del proceso) excluido: no figura en el listado EFPC ni es proyecto priorizado; sensibilidad EPP 2025: 98,83 sin 0553 y 98,86 con 0553 (recalculado en Fase 3). CUIPO (-3 %) se usa solo como contraste; la fuente de calculo es el plan de accion/ROCA. Las 3 diferencias con Hacienda 2025 son de lectura del PDF, no del dato fuente.'),
     'C9: decisiones de la seccion 3 y verificacion V1'),
    ('PP.EPP.APP', 'verificacion',
     'verificado (plan de accion IV-trim-2025, recalculado en Fase 3; plan de accion II-trim-2026); la cifra de la Rendicion de Cuentas 2025 p. 61 se toma de la revision del profesor Francisco, no reabierta',
     'C10: el Ejecutor no abrio la rendicion ni el Acta 096'),
    ('PDM.CM.O', 'verificacion',
     'verificado (Forensis 2024 pp. 37 y 440; Policia; TerriData); ANSV 2024 sin verificar (solo prensa, El Colombiano 22-ene-2025)',
     'C11: notas_PDM seccion 7'),
    ('PDM.ILE.A', 'verificacion',
     'parcial: numerador verificado (balance 2024 p. 24; informe 2025 p. 150); denominador DANE 0-6 anos y metodo de la base 16,3 % sin verificar',
     'C12: la brecha de la propia fila lo declara'),
    ('PDM.CP.O', 'verificacion',
     'verificado (busqueda negativa en repo y web); diapositiva p. 4 de la evaluacion T4-2025 leida de imagen',
     'C13: la etiqueta anterior sugeria un valor leido'),
    ('PDM.CP.O', 'estado_valor', 'No disponible', 'C13: estado vacio no esta en la lista de Registro'),
    ('CS.DSI.MS', 'estado_valor', 'Pendiente', 'C14: estado vacio'),
    ('CS.DSI.MS', 'notas',
     Anadir('Variable de definicion metodologica (decision docente), no dato de fuente; insumo: Apendice 9.1.2 del PDM.'),
     'C14: evitar que se lea como dato faltante de la Alcaldia'),
    ('CS.DSI.MD', 'valor_encontrado',
     Anadir('No trasladar al Registro: con universo de productos MD = MS por construccion. Si la seleccion es de metas de resultado, las evaluaciones de la Alcaldia 2024-2025 no publican ningun valor observado; MD solo contaria cifras de fuentes externas oficiales (MinSalud, Forensis, POT).'),
     'C15: evitar que 477/528/539 se lea como DSI = 100'),
    ('PDM.CA.O', 'brecha',
     Anadir('Comparabilidad con la meta 3,88 no verificada: con 5,75 el logro llegaria a 100 por tope sin que la serie sea la de la linea base.'),
     'C16: riesgo de puntaje artificial'),
    ('PP.CEP.PEC', 'desagregacion', Sustituir('Proxy', 'Parcial'), 'C17: consecuencia de C6'),
    ('PDM.CS.O', 'ruta_o_url',
     Sustituir('(zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip; zip ago-2026: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip)',
               '; zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip ; zip ago-2026: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip'),
     'C18: formato (parentesis); valores separados por " ; "'),
    ('PP.EPP.OPP', 'valor_encontrado', Prefijar(EPP_PRELIMINAR), 'C21 (usuaria): marcar la cifra EPP como preliminar, por validar'),
    ('PP.EPP.APP', 'valor_encontrado', Prefijar(EPP_PRELIMINAR), 'C21 (usuaria): marcar la cifra EPP como preliminar, por validar'),
]
# Valores previos esperados en cambios de categoria (si no coinciden, el script se detiene).
PREVIO_DISPONIBILIDAD = {
    'CS.RPO.OT': 'Construible', 'CS.RPO.OR': 'Construible', 'PC.I2.V1': 'Construible',
    'PDM.CS.O': 'Parcial', 'PDM.ILE.S': 'Construible', 'PP.CEP.PEC': 'Proxy',
}

# Normalizacion de dependencias de las solicitudes (clave = nombre original sin acentos ni prefijo).
DEPENDENCIAS = {
    'Alcaldia de Envigado - DAP, Subdireccion de Participacion Ciudadana': 'Departamento Administrativo de Planeación - Subdirección de Participación Ciudadana',
    'Alcaldia de Envigado - DAP': 'Departamento Administrativo de Planeación',
    'Alcaldia de Envigado - DAP / Oficina Asesora de Comunicaciones': 'Departamento Administrativo de Planeación / Oficina Asesora de Comunicaciones',
    'Alcaldia de Envigado - DAP, Direccion de Planeacion': 'Departamento Administrativo de Planeación - Dirección de Planeación',
    'Alcaldia de Envigado - Secretaria de Salud': 'Secretaría de Salud',
    'Alcaldia de Envigado - Bienestar Social': 'Bienestar Social',
    'Departamento Administrativo de Planeación (Subdirección de Participación Ciudadana)': 'Departamento Administrativo de Planeación - Subdirección de Participación Ciudadana',
    'Subdirección de Participación Ciudadana': 'Departamento Administrativo de Planeación - Subdirección de Participación Ciudadana',
    'Personeria de Envigado': 'Personería de Envigado',
    'Secretaría de Salud (Dirección de Aseguramiento)': 'Secretaría de Salud - Dirección de Aseguramiento',
}


def leer_csv(ruta):
    with open(ruta, encoding='utf-8', newline='') as f:
        return list(csv.DictReader(f))


def aplicar_correcciones(filas):
    por_id = {f['id_variable']: f for f in filas}
    aplicadas = 0
    for id_var, col, nuevo, motivo in CORRECCIONES:
        fila = por_id[id_var]
        actual = fila[col]
        if col == 'disponibilidad':
            assert actual == PREVIO_DISPONIBILIDAD[id_var], (id_var, actual)
        if isinstance(nuevo, Prefijar):
            fila[col] = str(nuevo) + ' ' + actual
        elif isinstance(nuevo, Anadir):
            fila[col] = (actual.rstrip() + ' | ' + nuevo) if actual.strip() else str(nuevo)
        elif isinstance(nuevo, Sustituir):
            assert nuevo.viejo in actual, (id_var, col, 'texto a sustituir no encontrado')
            fila[col] = actual.replace(nuevo.viejo, nuevo.nuevo)
        else:
            fila[col] = nuevo
        aplicadas += 1
    return aplicadas


def estados_validos(filas):
    for f in filas:
        if not f['estado_valor'].strip():
            f['estado_valor'] = 'No disponible' if f['disponibilidad'] == 'No disponible' else 'Pendiente'
        assert f['estado_valor'].strip().lower() != 'validado', f['id_variable']


def consolidar_fuentes():
    campos = ['fuente', 'emisor', 'formato', 'periodicidad', 'url_o_ruta', 'descargable_o_manual',
              'variables_que_alimenta', 'estado_acceso', 'notas']
    por_url, orden = {}, []
    for s in SISTEMAS:
        for r in leer_csv(TRABAJO / f'fuentes_{s}.csv'):
            clave = r['url_o_ruta'].strip()
            if clave not in por_url:
                por_url[clave] = dict(r)
                por_url[clave]['_vars'] = []
                orden.append(clave)
            dest = por_url[clave]
            for v in r['variables_que_alimenta'].split(';'):
                v = v.strip()
                if v and v not in dest['_vars']:
                    dest['_vars'].append(v)
            for c in campos:
                if not dest[c].strip() and r[c].strip():
                    dest[c] = r[c]
    filas = []
    for clave in orden:
        d = por_url[clave]
        d['variables_que_alimenta'] = '; '.join(d['_vars'])
        filas.append({c: d[c] for c in campos})
    return campos, filas


def consolidar_solicitudes():
    campos = ['dependencia', 'sistema', 'id_variable', 'dato_a_solicitar', 'periodo', 'desagregacion', 'justificacion']
    filas = []
    for s in SISTEMAS:
        for r in leer_csv(TRABAJO / f'solicitudes_{s}.csv'):
            dep = r['dependencia'].strip()
            r['dependencia'] = DEPENDENCIAS.get(dep, dep)
            r['sistema'] = s
            filas.append({c: r[c] for c in campos})
    filas.sort(key=lambda r: (r['dependencia'].lower(), r['sistema'], r['id_variable']))
    return campos, filas


ENCABEZADO_RELLENO = PatternFill('solid', fgColor='1F4E78')
ENCABEZADO_FUENTE = Font(bold=True, color='FFFFFF')


def escribir_hoja(wb, titulo, encabezados, filas, anchos, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws.append(encabezados)
    for fila in filas:
        ws.append(fila)
    for c in ws[1]:
        c.font, c.fill = ENCABEZADO_FUENTE, ENCABEZADO_RELLENO
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    return ws


def main():
    variables = []
    for s in SISTEMAS:
        filas_s = leer_csv(TRABAJO / f'variables_{s}.csv')
        for f in filas_s:
            f['sistema'] = s  # variables_CS.csv trae 'Control social'; se unifica al codigo del sistema
        variables += filas_s
    assert len(variables) == 39, len(variables)
    n = aplicar_correcciones(variables)
    estados_validos(variables)
    conteo = {c: sum(1 for v in variables if v['disponibilidad'] == c) for c in CATEGORIAS}
    assert conteo == CONTEO_ESPERADO, conteo

    claves = [c[0] for c in COLUMNAS_VARIABLES]
    filas_var = [[v[k] for k in claves] for v in variables]
    indicadores = leer_csv(TRABAJO / 'indicadores_fase3.csv')
    assert len(indicadores) == 23, len(indicadores)
    campos_ind = list(indicadores[0].keys())
    for r in indicadores:
        if r['codigo'] == 'PP.EPP':
            r['razon'] = EPP_PRELIMINAR + ' ' + r['razon']
    campos_fte, fuentes = consolidar_fuentes()
    campos_sol, solicitudes = consolidar_solicitudes()

    wb = Workbook()
    leeme = [
        ['Disponibilidad y brecha de información del sistema de indicadores - Observatorio Ciudadano de Envigado (OCE)'],
        ['Fecha de corte: 14-oct-2026. Periodo objetivo: 2024, 2025 y lo publicado de 2026 (sin cierre anual al corte). Territorio: Municipio; zonas: 13.'],
        ['Escala de disponibilidad: Disponible (publicado con la definición de la Matriz) · Construible (se calcula con fuentes verificadas) · Parcial (solo parte del periodo o sin el soporte exigido) · Proxy (definición, unidad o universo distintos) · No disponible.'],
        [f'Conteo de las 39 variables: ' + ' · '.join(f'{c} {conteo[c]}' for c in CATEGORIAS) + '.'],
        ['Todos los valores encontrados están en estado "Pendiente" (o "No disponible"): ninguno está validado por los docentes ni por la Alcaldía. Las cifras son exploratorias.'],
        ['Esta tarea no calcula el IIC ni ningún índice: solo documenta qué dato existe, dónde y cuál es la brecha.'],
        ['Fuentes: Fases 1 a 3 del repositorio (outputs/disponibilidad/trabajo/: variables, fuentes, solicitudes, notas y revision_fase3.md). Las 20 correcciones de la revisión de la Fase 3 están aplicadas en la hoja Variables, más la marca "preliminar, por validar" de EPP (C21).'],
        ['Hojas: Variables (39) · Indicadores (23, veredicto de viabilidad) · Fuentes (consolidadas por URL o ruta) · Solicitudes (por dependencia).'],
        ['La ejecución presupuestal de PP 2025 (EPP ≈ 98,83 %) es PRELIMINAR, POR VALIDAR: la lista de proyectos PP se reconstruyó cruzando el plan de acción IV-trim-2025 con el listado EFPC y debe confirmarla el DAP. Cálculo: scripts/epp_pp_2025.py (detalle en epp_pp_2025*.csv/.md).'],
        ['Reproducible con: python3 -I scripts/generar_disponibilidad.py'],
    ]
    ws = escribir_hoja(wb, 'Léeme', ['Léeme'], leeme, [150], primera=True)
    ws.auto_filter.ref = None
    escribir_hoja(wb, 'Variables', [c[1] for c in COLUMNAS_VARIABLES], filas_var, [c[2] for c in COLUMNAS_VARIABLES])
    escribir_hoja(wb, 'Indicadores', campos_ind, [[r[c] for c in campos_ind] for r in indicadores],
                  [10, 14, 34, 16, 28, 18, 50, 55, 28])
    escribir_hoja(wb, 'Fuentes', ['fuente', 'emisor', 'formato', 'periodicidad', 'url_o_ruta', 'descargable o extracción manual',
                                  'variables que alimenta', 'estado de acceso', 'notas'],
                  [[r[c] for c in campos_fte] for r in fuentes], [38, 28, 20, 14, 60, 30, 32, 24, 50])
    escribir_hoja(wb, 'Solicitudes', campos_sol, [[r[c] for c in campos_sol] for r in solicitudes],
                  [40, 9, 14, 55, 12, 20, 45])
    destino = SALIDA / 'disponibilidad_brecha_indicadores.xlsx'
    wb.save(destino)

    csv_var = SALIDA / 'disponibilidad_variables.csv'
    with open(csv_var, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        w.writerow([c[1] for c in COLUMNAS_VARIABLES])
        w.writerows(filas_var)
    # Misma logica que scripts/csv_para_excel.py (que solo admite rutas fijas): ';' y UTF-8 con BOM.
    csv_xl = SALIDA / 'disponibilidad_variables_excel.csv'
    with open(csv_var, encoding='utf-8', newline='') as fin, open(csv_xl, 'w', encoding='utf-8-sig', newline='') as fout:
        esc = csv.writer(fout, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        for fila in csv.reader(fin):
            esc.writerow(fila)

    print(f'Correcciones aplicadas: {n} operaciones sobre celdas (C1-C18 y C21; C19 validada; C20 es de documentacion)')
    print('Conteo por disponibilidad:', conteo)
    print(f'Variables {len(filas_var)} · Indicadores {len(indicadores)} · Fuentes {len(fuentes)} · Solicitudes {len(solicitudes)}')
    print(f'Generado {destino}\nGenerado {csv_var}\nGenerado {csv_xl}')


if __name__ == '__main__':
    main()
