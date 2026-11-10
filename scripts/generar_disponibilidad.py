"""Genera la matriz de disponibilidad y brecha del sistema de indicadores para enviar a los profesores.

Lee los CSV de trabajo de outputs/disponibilidad/trabajo/ (sin modificarlos), aplica las correcciones de la
revision critica (revision_fase3.md seccion 1) y una capa de presentacion que no cambia categorias, cifras ni
veredictos: rutas internas -> URL de origen (inventario_detallado_IIC.csv) con hipervinculos, lenguaje sin jerga
interna, tildes, encabezados en espanol, estados de acceso y verificacion normalizados, columna Resumen, celdas de
maximo 600 caracteres (el excedente pasa a Notas), colores por categoria y veredicto. Escribe:
  - outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx (Léeme, Indicadores, Variables, Fuentes,
    Solicitudes, Enlaces)
  - outputs/disponibilidad/disponibilidad_variables.csv (UTF-8, coma)
  - outputs/disponibilidad/disponibilidad_variables_excel.csv (separador ';', UTF-8 con BOM; misma logica de csv_para_excel.py)
Uso: python3 -I scripts/generar_disponibilidad.py
"""
import csv
import re
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
VEREDICTOS = ['Viable', 'Viable con ajuste', 'No viable por ahora']
VEREDICTO_ESPERADO = {'Viable': 1, 'Viable con ajuste': 10, 'No viable por ahora': 12}


class Anadir(str):
    """Agrega texto al final del valor actual (separado por ' | ')."""


class Prefijar(str):
    """Agrega texto al inicio del valor actual."""


class Sustituir:
    """Reemplaza todas las apariciones de 'viejo' por 'nuevo' en el valor actual."""
    def __init__(self, viejo, nuevo):
        self.viejo, self.nuevo = viejo, nuevo


EPP_PRELIMINAR = ('[PRELIMINAR, POR VALIDAR: lista de proyectos del presupuesto participativo reconstruida; debe '
                  'confirmarla Planeación. EPP 2025 = 98,83 % (APP $20.928.690.218; OPP $20.684.427.720)]')

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
    ('PP.CEP.PEC', 'notas',
     Sustituir('Cambia frente a la Fase 1 (Parcial): se reclasifica como Proxy porque',
               'Categoría final: Parcial, igual que en la clasificación preliminar (en la búsqueda se propuso Proxy; '
               'la revisión crítica la dejó en Parcial por tratarse de autorreporte). Motivo de la brecha:'),
     'C22 (QA): la nota decía Proxy y contradecía la categoría final'),
    ('PDM.CS.O', 'ruta_o_url',
     Sustituir('(zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip; zip ago-2026: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip)',
               '; zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip ; zip ago-2026: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip'),
     'C18: formato (parentesis); valores separados por " ; "'),
    ('PP.EPP.OPP', 'valor_encontrado',
     Sustituir('Vigencia 2026: $11.555.041.247 al 30-jun-2026 (Hacienda; 66 proyectos; el XLSX trae un corte anterior de $5.702.528.148)',
               'Vigencia 2026: $5.702.528.148 al 30-jun-2026 (informe de Hacienda y hoja ROCA del plan de acción, coinciden al peso; 66 proyectos)'),
     'C23 (v1.2): la cifra $11.555.041.247 de la versión 1.1 era errónea; OPP 2026 al 30-jun-2026 = $5.702.528.148 (ver config/supuestos.yaml pp_2026_opp_fuente)'),
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


# --------------------------------------------------------------------------- presentación para los profesores
# Solo cambia la forma: enlaces, lenguaje, ortografía, encabezados y formato. No toca categorías, cifras ni veredictos.
F_INVENTARIO = RAIZ / 'data_raw_IIC' / 'inventario_detallado_IIC.csv'
SIN_URL = 'Solo en archivo local del observatorio'
# Documentos del portal de la Alcaldía: el dominio www.envigado.gov.co tiene la cadena de certificados incompleta;
# los mismos archivos responden con certificado válido en el dominio espejo (verificado el 9-nov-2026: 36 de 36).
PORTAL_ARCHIVOS = 'https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/'
ESPEJO_ARCHIVOS = ('https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co'
                   '/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/')
URL_REEMPLAZO = {  # enlaces que cambiaron de dirección (verificado el 9-nov-2026)
    'https://adres.gov.co/entidades-territoriales/bdua': 'https://www.adres.gov.co/entidades-territoriales/bdua/reportes',
}
NOTAS_ENLACE = [  # (prefijo de URL, nota que acompaña al enlace)
    ('https://www.envigado.gov.co', 'puede requerir abrirse desde otro navegador'),
    ('https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf', 'verificar en navegador'),
    ('https://www.adres.gov.co/entidades-territoriales/bdua/reportes',
     'sección de consulta BDUA; enlace actualizado el 9-nov-2026 (la dirección anterior ya no existe)'),
]


def normalizar_url(u):
    u = URL_REEMPLAZO.get(u.rstrip('/'), u)
    if u.startswith(PORTAL_ARCHIVOS):
        u = ESPEJO_ARCHIVOS + u[len(PORTAL_ARCHIVOS):]
    return u


def nota_enlace(u):
    return next((nota for pref, nota in NOTAS_ENLACE if u.startswith(pref)), '')
NOMBRE_SISTEMA = {'CS': 'Control social', 'PC': 'Participación ciudadana', 'PDM': 'Cumplimiento del PDM',
                  'PP': 'Presupuesto participativo'}
MAX_CELDA = 600
MAX_RESUMEN = 200
RE_URL = re.compile(r'https?://[^\s;|,)"\]]+')
RE_RUTA = re.compile(r'data_raw_IIC/[^\s;|,()"\]]+')
# tokens que no se tocan al corregir ortografía: URLs, rutas, nombres de archivo y códigos con guion bajo
RE_PROTEGIDO = re.compile(r'(https?://\S+|data_raw_IIC/\S+|\S+\.(?:pdf|xlsx|xls|csv|docx|zip|json|aspx|md|py)\b'
                          r'|\S*_\S*|\S*/\S*/\S*)')

# Jerga interna -> lenguaje neutro (se aplica en orden, sobre el texto completo de cada celda).
JERGA = [
    (r'\s*\[?\(?(?:ver|listado en|detalle en) notas_\w+\.md\)?\]?', ' (detalle en las notas de trabajo del observatorio)'),
    (r' \(nombre de persona omitido; ruta completa en inventario_detallado_IIC\.csv\)', ' (nombre de persona omitido)'),
    (r'Cambio frente a (?:la )?Fase 1', 'Cambio respecto a la clasificación preliminar'),
    (r'Cambia frente a (?:la )?Fase 1', 'Cambia respecto a la clasificación preliminar'),
    (r'Sin cambio( de categoría)? frente a (?:la )?Fase 1', r'Sin cambio\1 respecto a la clasificación preliminar'),
    (r'Coincide con (?:la )?Fase 1', 'Coincide con la clasificación preliminar'),
    (r'Diferencia con (?:la )?Fase 1', 'Diferencia con la clasificación preliminar'),
    (r'Corrige a (?:la )?Fase 1', 'Corrige la clasificación preliminar'),
    (r'citado en (?:la )?Fase 1', 'citado en la clasificación preliminar'),
    (r'\(referencia Fase 1, sin reabrir\)', '(referencia de la clasificación preliminar, sin reabrir)'),
    (r'referencia de Fase 1', 'referencia de la clasificación preliminar'),
    (r'sin abrir en esta fase', 'sin abrir en esta revisión'),
    (r'en esta fase', 'en esta revisión'),
    (r'^Fase 1:', 'Clasificación preliminar:'),
    (r'(?<=[.|;] )Fase 1:', 'Clasificación preliminar:'),
    (r'\(recalculado en Fase 3\)', '(recalculado en la revisión)'),
    (r', recalculado en Fase 3', ', recalculado en la revisión'),
    (r'No reabierto en Fase 2\.', 'No se reabrió durante la búsqueda.'),
    (r'\bFase \d\b', 'la revisión'),
    (r'Cifras calculadas por Ejecutor con scripts temporales; sin validar',
     'Cifras calculadas por el equipo del observatorio; sin validar'),
    (r'\b(por|del) (?:el )?Ejecutor\b', r'\1 el equipo del observatorio'),
    (r'\bpor el el equipo', 'por el equipo'),
    (r'\bdel el equipo', 'del equipo'),
    (r'\bEjecutor\b', 'equipo del observatorio'),
    (r'\bAnalizador\b', 'equipo del observatorio'),
    (r'con scripts temporales', ''),
    (r'Resumen de WebFetch', 'Resumen automático de la página'),
    (r'Consultado con WebFetch \(resumen automático\)', 'Consultado en línea (resumen automático)'),
    (r'WebFetch falla por', 'La consulta en línea falla por'),
    (r'; con curl -k devuelve solo el esqueleto \([\d.,]+ KB\) de una SPA', '; la descarga directa solo devuelve una página vacía (aplicación JavaScript)'),
    (r'WebSearch no devolvi[oó]', 'La búsqueda web no devolvió'),
    (r' \(WebFetch\)', ''),
    (r' en WebFetch y curl', ' en la consulta en línea'),
    (r'curl error 60 \(cadena de certificados incompleta\)', 'error de certificado (cadena incompleta)'),
    (r'\(WebFetch OK\)', '(consulta en línea correcta)'),
    (r'\bWebFetch\b', 'consulta en línea'),
    (r'\bWebSearch\b', 'búsqueda web'),
    (r'\(HTTP 200 verificado con curl -k, ', '(documento descargado, '),
    (r'\(HTTP 200, Last-Modified ([^;]+); requiere curl -k por cadena de certificado\)',
     r'(documento descargado; publicado el \1; el portal tiene un certificado incompleto)'),
    (r'\(curl, (\d+ s)\)', r'(\1)'),
    (r'API HTTP 200', 'API disponible'),
    (r'Portal responde HTTP 200', 'El portal responde'),
    (r'\(HTTP 200 en la página de listados\)', '(la página de listados responde)'),
    (r'\(HTTP 200\)', '(en línea)'),
    (r'HTTP 200', 'en línea'),
    (r'\bcurl(?: -k)?\b', 'descarga directa'),
    (r' desde el Codespace', ''),
    (r'\bCodespace\b', 'entorno de trabajo'),
    (r', reproducible con scripts/\w+\.py', ''),
    (r'scripts/\w+\.py', 'cálculo del observatorio'),
    (r'repo y web', 'archivo del observatorio y en la web'),
    (r'repo y URL de origen', 'archivo local y URL de origen'),
    (r'\b(?:en el|en) repo\b', 'en el archivo del observatorio'),
    (r'\bdel repo\b', 'del archivo del observatorio'),
    (r'\bel repo\b', 'el archivo del observatorio'),
    (r'\bya en repo\b', 'documento descargado'),
    (r'\brepo\b', 'archivo del observatorio'),
    (r'\bQ&A\b', 'preguntas y respuestas'),
    (r'copias en data_raw_IIC', 'copias en el archivo del observatorio'),
    (r'\bno Validado\b', 'sin validar'),
    (r'  +', ' '),
]
JERGA = [(re.compile(a, re.M), b) for a, b in JERGA]
RE_JERGA_CONTROL = re.compile(r'(?i)(\brepo\b|\bfase \d\b|\bejecutor\b|\banalizador\b|http \d{3}|\bcurl\b|webfetch|'
                              r'websearch|codespace|\w+\.py\b|notas_\w+\.md|revision_fase3|scratchpad|\bIIC\b)')

# Ortografía: palabras sin tilde -> con tilde (minúsculas; se conserva la mayúscula). Se excluyen las ambiguas
# (que, si, donde, cuales, cambio, esta, publica, solo, mas, como, valida, practica), tratadas aparte o sin tocar.
TILDES = {
    'veedurias': 'veedurías', 'veeduria': 'veeduría', 'alcaldia': 'alcaldía', 'comite': 'comité', 'comites': 'comités',
    'apendice': 'apéndice', 'apendices': 'apéndices', 'personeria': 'personería', 'dias': 'días', 'dia': 'día',
    'secretaria': 'secretaría', 'secretarias': 'secretarías', 'codigo': 'código', 'codigos': 'códigos', 'segun': 'según',
    'ultimo': 'último', 'ultima': 'última', 'ultimos': 'últimos', 'ultimas': 'últimas', 'linea': 'línea',
    'lineas': 'líneas', 'metodologica': 'metodológica', 'metodologico': 'metodológico', 'metodologia': 'metodología',
    'estan': 'están', 'ano': 'año', 'anos': 'años', 'anio': 'año', 'anios': 'años', 'calculo': 'cálculo',
    'calculos': 'cálculos', 'contraloria': 'contraloría', 'diagnostico': 'diagnóstico', 'numerico': 'numérico',
    'ningun': 'ningún', 'via': 'vía', 'ademas': 'además', 'tambien': 'también', 'telefonos': 'teléfonos',
    'policia': 'policía', 'busqueda': 'búsqueda', 'busquedas': 'búsquedas', 'leida': 'leída', 'leido': 'leído',
    'metodo': 'método', 'metodos': 'métodos', 'fisica': 'física', 'fisico': 'físico', 'razon': 'razón',
    'imagenes': 'imágenes', 'unica': 'única', 'unico': 'único', 'devolvio': 'devolvió',
    'catalogo': 'catálogo', 'hallo': 'halló', 'pagina': 'página', 'paginas': 'páginas', 'publico': 'público',
    'publicos': 'públicos', 'publicas': 'públicas', 'numero': 'número', 'numeros': 'números', 'tecnico': 'técnico',
    'tecnicos': 'técnicos', 'tecnica': 'técnica', 'tecnicas': 'técnicas', 'analisis': 'análisis', 'indice': 'índice',
    'indices': 'índices', 'minimo': 'mínimo', 'maximo': 'máximo', 'basico': 'básico', 'region': 'región',
    'transito': 'tránsito', 'vehiculo': 'vehículo', 'vehiculos': 'vehículos', 'grafico': 'gráfico',
    'graficos': 'gráficos', 'juridico': 'jurídico', 'juridica': 'jurídica', 'especifico': 'específico',
    'especifica': 'específica', 'electronico': 'electrónico', 'historico': 'histórico', 'historica': 'histórica',
    'estadistica': 'estadística', 'estadisticas': 'estadísticas', 'economico': 'económico', 'politica': 'política',
    'politicas': 'políticas', 'categoria': 'categoría', 'categorias': 'categorías', 'aqui': 'aquí', 'asi': 'así',
    'despues': 'después', 'todavia': 'todavía', 'area': 'área', 'areas': 'áreas', 'medellin': 'Medellín',
    'bogota': 'Bogotá', 'util': 'útil', 'utiles': 'útiles', 'facil': 'fácil', 'dificil': 'difícil',
    'credito': 'crédito', 'deficit': 'déficit', 'periodico': 'periódico', 'academico': 'académico',
    'geografico': 'geográfico', 'automatico': 'automático', 'automatica': 'automática', 'tematica': 'temática',
    'estrategico': 'estratégico', 'estrategica': 'estratégica', 'practicas': 'prácticas', 'informacion': 'información', 'deberia': 'debería', 'tendria': 'tendría',
    'podria': 'podría', 'habria': 'habría', 'seria': 'sería', 'garantia': 'garantía', 'energia': 'energía',
    'ninos': 'niños', 'nino': 'niño', 'ninas': 'niñas', 'nina': 'niña', 'senal': 'señal', 'diseno': 'diseño',
    'pequeno': 'pequeño', 'tamano': 'tamaño', 'espanol': 'español', 'compania': 'compañía', 'campana': 'campaña',
    'enfasis': 'énfasis', 'sintesis': 'síntesis', 'hipotesis': 'hipótesis', 'cronologico': 'cronológico',
    'logistica': 'logística', 'demografico': 'demográfico', 'proposito': 'propósito', 'parametro': 'parámetro', 'parametros': 'parámetros', 'termino': 'término',
    'victimas': 'víctimas', 'victima': 'víctima', 'etnicos': 'étnicos', 'etnico': 'étnico', 'etnicas': 'étnicas',
    'etnica': 'étnica', 'terminos': 'términos', 'genero': 'género', 'publicacion': 'publicación',
}
RE_PALABRA = re.compile(r'[A-Za-zÁÉÍÓÚáéíóúÑñÜü]+')
# 'publica' es adjetivo tras estos sustantivos; en otro caso es verbo ("el PDM no publica")
RE_PUBLICA_ADJ = re.compile(r'\b((?:[Rr]endici[oó]n|[Aa]udiencia|[Mm]atriz|[Pp]ol[ií]tica|[Gg]esti[oó]n|[Cc]onsulta|'
                            r'[Ee]ntidad|[Ii]nformaci[oó]n|[Ii]nversi[oó]n|[Ss]alud|[Ff]unci[oó]n|[Vv][ií]a)) publica\b')


def con_mayusculas(original, nueva):
    if original.isupper() and len(original) > 1:
        return nueva.upper()
    if original[0].isupper():
        return nueva[0].upper() + nueva[1:]
    return nueva


def tildar_segmento(s):
    def palabra(m):
        w = m.group(0)
        low = w.lower()
        if re.fullmatch(r'[a-z]+(cion|sion)', low):
            return con_mayusculas(w, low[:-2] + 'ón')  # -cion/-sion -> -ción/-sión
        if low in TILDES:
            return con_mayusculas(w, TILDES[low])
        return w
    s = RE_PALABRA.sub(palabra, s)
    s = RE_PUBLICA_ADJ.sub(r'\1 pública', s)
    s = re.sub(r'\b([Nn]o) esta\b', r'\1 está', s)
    s = re.sub(r'\bsobre cuales\b', 'sobre cuáles', s)
    s = re.sub(r'\by cuantas\b', 'y cuántas', s)
    s = re.sub(r'\b(oficial|[Ee]l PDM|[Ee]l informe(?: de gestión)? \d{4}) si (es|publica|trae)\b', r'\1 sí \2', s)
    s = re.sub(r'\besta (definid[oa]s?|disponible|publicad[oa]s?|vac[ií][oa])', r'está \1', s)
    return s


def tildar(texto):
    """Corrige tildes y eñes fuera de URLs, rutas y nombres de archivo."""
    partes = RE_PROTEGIDO.split(texto)
    return ''.join(p if i % 2 else tildar_segmento(p) for i, p in enumerate(partes))


def limpiar(texto):
    s = str(texto or '')
    for patron, reemplazo in JERGA:
        s = patron.sub(reemplazo, s)
    s = tildar(s).strip()
    return re.sub(r'\s+([.,;])', r'\1', s)


def cargar_inventario():
    with open(F_INVENTARIO, encoding='utf-8', newline='') as f:
        return {r['ruta_local']: r for r in csv.DictReader(f)}


def url_de_ruta(ruta, inv):
    """URL de origen de un archivo del inventario; para carpetas, la página de origen común de sus archivos."""
    if 'NOMBRE-OMITIDO' in ruta:  # la URL repetiría el nombre de la persona
        return None
    if ruta in inv:
        return inv[ruta]['url_origen'] or None
    if ruta.endswith('/'):
        paginas = {r['pagina_fuente'] for k, r in inv.items() if k.startswith(ruta) and r['pagina_fuente']}
        return paginas.pop() if len(paginas) == 1 else None
    return None


def separar_enlaces(texto, inv):
    """Devuelve (lista de URLs, lista de rutas internas, texto sobrante sin rutas ni URLs)."""
    urls, rutas = [], []
    for r in RE_RUTA.findall(texto or ''):
        r = r.rstrip('.')
        rutas.append(r)
        u = url_de_ruta(r, inv)
        u = normalizar_url(u) if u else u
        if u and u not in urls:
            urls.append(u)
        elif not u and SIN_URL not in urls:
            urls.append(SIN_URL)
    for u in RE_URL.findall(texto or ''):
        u = normalizar_url(u.rstrip('.'))
        if u not in urls:
            urls.append(u)
    resto = RE_URL.sub('', RE_RUTA.sub('', texto or ''))
    resto = re.sub(r'(\s*[;|]\s*)+', '; ', resto).strip(' ;|')
    resto = re.sub(r'^[(\s]*\)|\(\s*\)', '', resto).strip(' ;|')
    return urls, rutas, resto


def estado_acceso(texto, tiene_local):
    t = texto.lower()
    if any(k in t for k in ('no localizado', 'sin hallazgos', 'sin descarga municipal', 'tiempo de espera', 'expiró',
                            'error', 'falla', 'bloqueado', 'sin acceso')):
        return 'Solo en archivo local' if tiene_local else 'No localizado'
    if any(k in t for k in ('(local)', '(repo)', '(en repo)')) and 'url de origen' not in t:
        return 'Solo en archivo local'
    return 'Accesible en línea'


def verificacion(texto):
    t = texto.lower().strip()
    if t.startswith('parcial') or ('sin verificar' in t and 'verificado' in t.replace('sin verificar', '')):
        return 'Parcialmente verificado'
    if t.startswith('sin verificar'):
        return 'Sin verificar'
    return 'Verificado'


def corte_frase(texto, n):
    """Corta en el último fin de frase (o palabra) antes de n caracteres; devuelve (cabeza, resto)."""
    if len(texto) <= n:
        return texto, ''
    trozo = texto[:n]
    k = max(trozo.rfind('. '), trozo.rfind('; '), trozo.rfind(' | '))
    if k < n * 0.4:
        k = trozo.rfind(' ')
    return texto[:k + 1].strip(), texto[k + 1:].strip(' |;')


def resumen(brecha, accion):
    base = brecha or accion
    primera = re.split(r'(?<=\.)\s|\s\|\s', base, maxsplit=1)[0].strip()
    if len(primera) > MAX_RESUMEN:
        # se intenta cerrar en un punto y coma; si no hay, se corta en la última palabra
        k = primera[:MAX_RESUMEN - 1].rfind(';')
        corte = primera[:k] if k > MAX_RESUMEN * 0.4 else ''
        if not corte or corte.count('(') != corte.count(')'):
            corte = primera[:MAX_RESUMEN - 1].rsplit(' ', 1)[0]
            if corte.count('(') > corte.count(')'):
                corte = corte[:corte.rfind('(')].rstrip()
            corte += '…'
        primera = corte
    primera = primera.rstrip(' ,;:')
    return primera if primera.endswith(('.', '…')) else primera + '.'


def acotar(fila, columnas, col_notas, etiquetas):
    """Ninguna celda supera MAX_CELDA: el excedente pasa a Notas (y Notas se reparte en columnas de continuación)."""
    extra = []
    for c in columnas:
        cabeza, resto = corte_frase(fila[c], MAX_CELDA)
        if resto:
            fila[c] = cabeza + ' (continúa en Notas)'
            extra.append(f'{etiquetas[c]} (continuación): {resto}')
    notas = ' | '.join(x for x in [fila[col_notas]] + extra if x)
    partes = []
    while notas:
        cabeza, notas = corte_frase(notas, MAX_CELDA)
        partes.append(cabeza)
    return partes or ['']


RELLENO = {k: PatternFill('solid', fgColor=v) for k, v in {
    'Disponible': 'D6F0D6', 'Construible': 'DCEBFA', 'Parcial': 'FFF3BF', 'Proxy': 'FFDEB8', 'No disponible': 'F8D4D4',
    'Viable': 'D6F0D6', 'Viable con ajuste': 'FFF3BF', 'No viable por ahora': 'F8D4D4'}.items()}
ENCABEZADO_RELLENO = PatternFill('solid', fgColor='1F4E78')
ENCABEZADO_FUENTE = Font(bold=True, color='FFFFFF')
FUENTE_ENLACE = Font(color='0563C1', underline='single')


def escribir_hoja(wb, titulo, encabezados, filas, anchos, primera=False, enlaces=None, colorear=None):
    """enlaces: {indice de columna: lista (por fila) de URL o None}; colorear: indices de columnas con relleno por valor."""
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws.append(encabezados)
    for fila in filas:
        ws.append(fila)
    for c in ws[1]:
        c.font, c.fill = ENCABEZADO_FUENTE, ENCABEZADO_RELLENO
        c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[1].height = 32
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    for r, fila in enumerate(ws.iter_rows(min_row=2), start=0):
        lineas = 1
        for j, c in enumerate(fila):
            c.alignment = Alignment(wrap_text=True, vertical='top')
            texto = '' if c.value is None else str(c.value)
            por_linea = max(anchos[j] * 1.15, 1)
            lineas = max(lineas, sum(max(1, -(-len(t) // int(por_linea))) for t in texto.split('\n')))
            if colorear and j in colorear and texto in RELLENO:
                c.fill = RELLENO[texto]
            if enlaces and j in enlaces and enlaces[j][r]:
                c.hyperlink = enlaces[j][r]
                c.font = FUENTE_ENLACE
        ws.row_dimensions[r + 2].height = min(409, 15 * lineas + 4)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    return ws


def es_url(u):
    return u.startswith('http')


COLUMNAS_VAR_FINAL = [
    ('sistema', 'Sistema', 16), ('codigo_indicador', 'Código del indicador', 12), ('id_variable', 'ID de la variable', 13),
    ('nombre', 'Nombre', 34), ('unidad', 'Unidad', 11), ('fuente_sugerida', 'Fuente sugerida por la matriz', 30),
    ('fuente_encontrada', 'Fuente encontrada', 34), ('enlace', 'Enlace (URL de origen)', 40),
    ('pagina_tabla', 'Página o tabla', 24), ('anios_disponibles', 'Años disponibles', 22),
    ('desagregacion', 'Desagregación (municipio/zona)', 30), ('formato', 'Formato', 12),
    ('disponibilidad', 'Disponibilidad', 14), ('resumen', 'Resumen', 40),
    ('valor_encontrado', 'Valor encontrado (y año)', 48), ('estado_valor', 'Estado del valor', 12),
    ('brecha', 'Brecha', 44), ('accion_sugerida', 'Acción sugerida', 38), ('esfuerzo', 'Esfuerzo', 10),
    ('verificacion', 'Verificación', 14), ('notas', 'Notas', 55),
]
COLUMNAS_IND = [('sistema', 'Sistema', 16), ('codigo', 'Código', 12), ('nombre', 'Nombre', 30), ('tipo', 'Tipo', 14),
                ('variables', 'Variables', 22), ('veredicto', 'Veredicto', 16), ('razon', 'Razón', 50),
                ('ajuste_propuesto', 'Ajuste propuesto', 50), ('anios_calculables', 'Años calculables', 20)]
COLUMNAS_FTE = [('fuente', 'Fuente', 34), ('emisor', 'Emisor', 24), ('formato', 'Formato', 14),
                ('periodicidad', 'Periodicidad', 13), ('enlace', 'Enlace (URL de origen)', 42),
                ('descargable_o_manual', 'Descarga o extracción manual', 28),
                ('variables_que_alimenta', 'Variables que alimenta', 24), ('estado_acceso', 'Estado de acceso', 16),
                ('notas', 'Notas', 50)]
COLUMNAS_SOL = [('dependencia', 'Dependencia', 34), ('sistema', 'Sistema', 16), ('id_variable', 'ID de la variable', 13),
                ('dato_a_solicitar', 'Dato a solicitar', 50), ('periodo', 'Periodo', 12),
                ('desagregacion', 'Desagregación', 18), ('justificacion', 'Justificación', 45)]


def lista_acotada(items, sufijo):
    """Une items por línea sin pasar de MAX_CELDA; si sobran, termina con '(+N <sufijo>)'."""
    texto, n = '', 0
    for k, it in enumerate(items):
        cand = (texto + '\n' + it) if texto else it
        if len(cand) > MAX_CELDA - 45:
            return texto + f'\n(+{len(items) - k} {sufijo})'
        texto = cand
    return texto


def partir_lineas(texto, n=MAX_CELDA):
    """Reparte un texto de varias líneas en trozos de máximo n caracteres, sin cortar líneas."""
    partes, actual = [], ''
    for linea in texto.split('\n'):
        cand = (actual + '\n' + linea) if actual else linea
        if len(cand) > n and actual:
            partes.append(actual)
            actual = linea
        else:
            actual = cand
    return partes + [actual] if actual else partes or ['']


def presentar_enlaces(fila, campo, inv, hoja, ident, tabla_enlaces):
    urls, rutas, resto = separar_enlaces(fila[campo], inv)
    fila['enlace'] = lista_acotada(urls, 'enlaces más en la hoja Enlaces') if urls else SIN_URL
    fila['ruta_interna'] = '\n'.join(rutas)
    notas = []
    for u in urls:
        if es_url(u):
            tabla_enlaces.append([hoja, ident, u, nota_enlace(u)])
            if nota_enlace(u):
                notas.append(f'Enlace {u}: {nota_enlace(u)}.')
    fila['_notas_enlace'] = ' '.join(dict.fromkeys(notas))
    primera = next((u for u in urls if es_url(u)), None)
    return primera, resto


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
    indicadores = leer_csv(TRABAJO / 'indicadores_fase3.csv')
    assert len(indicadores) == 23, len(indicadores)
    for r in indicadores:
        if r['codigo'] == 'PP.EPP':
            r['razon'] = EPP_PRELIMINAR + ' ' + r['razon']
    veredictos = {v: sum(1 for r in indicadores if r['veredicto'] == v) for v in VEREDICTOS}
    assert veredictos == VEREDICTO_ESPERADO, veredictos
    _, fuentes = consolidar_fuentes()
    _, solicitudes = consolidar_solicitudes()
    inv = cargar_inventario()
    tabla_enlaces = []

    # ---- Variables
    filas_var, links_var = [], []
    etiquetas_var = {c[0]: c[1] for c in COLUMNAS_VAR_FINAL}
    for v in variables:
        v = dict(v)
        primera, resto = presentar_enlaces(v, 'ruta_o_url', inv, 'Variables', v['id_variable'], tabla_enlaces)
        detalle_ver = limpiar(v['verificacion'])
        v['verificacion'] = verificacion(v['verificacion'])
        texto = {k: limpiar(v[k]) for k in ('nombre', 'unidad', 'fuente_sugerida', 'fuente_encontrada', 'pagina_tabla',
                                            'anios_disponibles', 'desagregacion', 'formato', 'valor_encontrado',
                                            'brecha', 'accion_sugerida', 'esfuerzo', 'notas')}
        v.update(texto)
        extras = []
        if resto:
            extras.append('Fuente: ' + limpiar(resto))
        if v['_notas_enlace']:
            extras.append(v['_notas_enlace'])
        if detalle_ver.lower() not in ('verificado', 'sin verificar'):
            extras.append('Verificación: ' + detalle_ver)
        v['notas'] = ' | '.join(x for x in [v['notas']] + extras if x)
        v['sistema'] = NOMBRE_SISTEMA[v['sistema']]
        v['resumen'] = resumen(v['brecha'], v['accion_sugerida'])
        notas = acotar(v, ['fuente_encontrada', 'pagina_tabla', 'desagregacion', 'valor_encontrado', 'brecha',
                           'accion_sugerida', 'fuente_sugerida'], 'notas', etiquetas_var)
        v['_notas'] = notas
        filas_var.append(v)
        links_var.append(primera)
    n_notas = max(len(v['_notas']) for v in filas_var)
    for v in filas_var:
        v['_rutas'] = partir_lineas(v['ruta_interna'])
    n_rutas = max(len(v['_rutas']) for v in filas_var)
    cab_var = [c[1] for c in COLUMNAS_VAR_FINAL[:-1]] + ['Notas'] + \
        [f'Notas (continuación {i})' for i in range(1, n_notas)] + ['Ruta interna'] + \
        [f'Ruta interna (continuación {i})' for i in range(1, n_rutas)]
    anch_var = [c[2] for c in COLUMNAS_VAR_FINAL[:-1]] + [55] * n_notas + [60] * n_rutas
    datos_var = [[v[c[0]] for c in COLUMNAS_VAR_FINAL[:-1]] + v['_notas'] + [''] * (n_notas - len(v['_notas']))
                 + v['_rutas'] + [''] * (n_rutas - len(v['_rutas'])) for v in filas_var]

    # ---- Indicadores
    datos_ind = []
    for r in indicadores:
        r = {k: limpiar(x) for k, x in r.items()}
        r['sistema'] = NOMBRE_SISTEMA[r['sistema']]
        for c in ('razon', 'ajuste_propuesto'):
            r[c], resto = corte_frase(r[c], MAX_CELDA)
            assert not resto, (r['codigo'], c)
        datos_ind.append([r[c[0]] for c in COLUMNAS_IND])

    # ---- Fuentes
    filas_fte, links_fte = [], []
    etiquetas_fte = {c[0]: c[1] for c in COLUMNAS_FTE}
    for f in fuentes:
        f = dict(f)
        primera, resto = presentar_enlaces(f, 'url_o_ruta', inv, 'Fuentes', f['fuente'], tabla_enlaces)
        urls_desc = RE_URL.findall(f['descargable_o_manual'])
        for u in map(normalizar_url, urls_desc):
            tabla_enlaces.append(['Fuentes', f['fuente'], u, nota_enlace(u)])
        f['descargable_o_manual'] = RE_URL.sub('(ver enlace en la hoja Enlaces)', f['descargable_o_manual'])
        detalle_acc = limpiar(f['estado_acceso'])
        f['estado_acceso'] = estado_acceso(f['estado_acceso'], bool(f['ruta_interna']))
        f = {k: (limpiar(x) if k not in ('enlace', 'ruta_interna', 'estado_acceso') else x) for k, x in f.items()}
        extras = []
        if resto:
            extras.append('Referencia: ' + limpiar(resto))
        if f['_notas_enlace']:
            extras.append(f['_notas_enlace'])
        if detalle_acc and detalle_acc.lower().rstrip('.') not in ('accesible',):
            extras.append('Acceso: ' + detalle_acc)
        f['notas'] = ' | '.join(x for x in [f['notas']] + extras if x)
        notas = acotar(f, ['fuente', 'descargable_o_manual'], 'notas', etiquetas_fte)
        f['_notas'] = notas
        filas_fte.append(f)
        links_fte.append(primera)
    n_nf = max(len(f['_notas']) for f in filas_fte)
    for f in filas_fte:
        f['_rutas'] = partir_lineas(f['ruta_interna'])
    n_rf = max(len(f['_rutas']) for f in filas_fte)
    cab_fte = [c[1] for c in COLUMNAS_FTE[:-1]] + ['Notas'] + \
        [f'Notas (continuación {i})' for i in range(1, n_nf)] + ['Ruta interna'] + \
        [f'Ruta interna (continuación {i})' for i in range(1, n_rf)]
    anch_fte = [c[2] for c in COLUMNAS_FTE[:-1]] + [50] * n_nf + [60] * n_rf
    datos_fte = [[f[c[0]] for c in COLUMNAS_FTE[:-1]] + f['_notas'] + [''] * (n_nf - len(f['_notas']))
                 + f['_rutas'] + [''] * (n_rf - len(f['_rutas'])) for f in filas_fte]

    # ---- Solicitudes
    datos_sol = []
    for r in solicitudes:
        r = {k: limpiar(x) for k, x in r.items()}
        r['sistema'] = NOMBRE_SISTEMA[r['sistema']]
        datos_sol.append([r[c[0]] for c in COLUMNAS_SOL])

    # ---- Léeme
    leeme = [
        ['Versión',
         'Versión 1.2 – 9 de noviembre de 2026. Fecha de corte de la información: 14 de octubre de 2026.'],
        ['Cambios de la versión 1.2',
         'Cambios de la versión 1.2: se corrige la cifra de obligaciones del presupuesto participativo 2026 al 30 de '
         'junio de 2026 ($5.702.528.148; la cifra $11.555.041.247 de la versión 1.1 era errónea).'],
        ['Qué es este documento',
         'Matriz de disponibilidad y brecha de información del sistema preliminar de indicadores del Observatorio '
         'Ciudadano de Envigado. Para cada una de las 39 variables de la matriz de indicadores dice si el dato existe, '
         'dónde está, en qué formato y qué falta; y para cada uno de los 23 indicadores e índices, si se puede '
         'calcular hoy.'],
        ['Periodo y territorio',
         'Años 2024 y 2025, y lo publicado de 2026 (sin cierre anual a la fecha de corte, 14 de octubre de 2026). '
         'Territorio: Municipio de Envigado; cuando hay desagregación, 13 zonas.'],
        ['Escala de disponibilidad',
         'Disponible: el dato está publicado con la unidad y la definición del indicador para al menos un año entre '
         '2024 y 2026.\nConstruible: no está publicado como cifra, pero se puede contar o calcular con documentos que '
         'ya se tienen.\nParcial: existe, pero faltan años, la desagregación o el soporte exigido, o la definición '
         'difiere de la matriz.\nProxy: solo hay un dato parecido, con otra definición, unidad o universo.\n'
         'No disponible: no está en el archivo del observatorio ni en fuentes públicas; hay que pedirlo a la Alcaldía.'],
        ['Resultados',
         'Variables (39): ' + ', '.join(f'{c} {conteo[c]}' for c in CATEGORIAS) + '.\n'
         'Indicadores e índices (23): ' + ', '.join(f'{v} {veredictos[v]}' for v in VEREDICTOS) + '.'],
        ['Cómo leer las hojas',
         'Recomendamos empezar por la hoja Indicadores: veredicto de cada indicador, razón y ajuste propuesto.\n'
         'Variables: una fila por variable. La columna Resumen dice lo esencial; el detalle está en Brecha, Acción '
         'sugerida y Notas. El enlace lleva al documento de origen.\n'
         'Fuentes: catálogo de documentos y portales consultados, con su estado de acceso.\n'
         'Solicitudes: datos que habría que pedir a la Alcaldía, agrupados por dependencia.\n'
         'Enlaces: todos los enlaces de las hojas Variables y Fuentes, uno por fila, para abrirlos directamente.'],
        ['Colores',
         'Disponibilidad: verde = Disponible, azul = Construible, amarillo = Parcial, naranja = Proxy, '
         'rojo = No disponible. Veredicto: verde = Viable, amarillo = Viable con ajuste, rojo = No viable por ahora.'],
        ['Aviso: nada está validado',
         'Todos los valores encontrados están en estado "Pendiente" (o "No disponible"). Ninguno ha sido validado por '
         'los responsables de la matriz ni por la Alcaldía; las cifras son exploratorias.'],
        ['Aviso: ejecución presupuestal del PP 2025',
         'La ejecución presupuestal del presupuesto participativo 2025 (98,83 %) es PRELIMINAR, POR VALIDAR. La lista '
         'de proyectos se reconstruyó cruzando el plan de acción del IV trimestre de 2025 con el listado de proyectos '
         'de la Escuela de Formación para la Participación Ciudadana. Falta que Planeación confirme la lista, decidir '
         'si se incluye el proyecto 0553 (costo del proceso) y las cifras de 2024 y 2026 son aproximadas.'],
    ]

    wb = Workbook()
    ws = escribir_hoja(wb, 'Léeme', ['Tema', 'Contenido'], leeme, [28, 110], primera=True)
    ws.auto_filter.ref = None
    i_disp = [c[0] for c in COLUMNAS_VAR_FINAL].index('disponibilidad')
    i_enl = [c[0] for c in COLUMNAS_VAR_FINAL].index('enlace')
    escribir_hoja(wb, 'Variables', cab_var, datos_var, anch_var, enlaces={i_enl: links_var}, colorear={i_disp})
    escribir_hoja(wb, 'Indicadores', [c[1] for c in COLUMNAS_IND], datos_ind, [c[2] for c in COLUMNAS_IND],
                  colorear={[c[0] for c in COLUMNAS_IND].index('veredicto')})
    escribir_hoja(wb, 'Fuentes', cab_fte, datos_fte, anch_fte,
                  enlaces={[c[0] for c in COLUMNAS_FTE].index('enlace'): links_fte})
    escribir_hoja(wb, 'Solicitudes', [c[1] for c in COLUMNAS_SOL], datos_sol, [c[2] for c in COLUMNAS_SOL])
    escribir_hoja(wb, 'Enlaces', ['Hoja', 'Variable o fuente', 'Enlace', 'Nota'],
                  [[h, limpiar(i), u, n] for h, i, u, n in tabla_enlaces],
                  [12, 46, 85, 40], enlaces={2: [u for _, _, u, _ in tabla_enlaces]})
    # el orden de hojas pone Indicadores antes de Variables (se recomienda empezar por ahí)
    wb.move_sheet('Indicadores', offset=-1)
    destino = SALIDA / 'disponibilidad_brecha_indicadores.xlsx'
    wb.save(destino)

    csv_var = SALIDA / 'disponibilidad_variables.csv'
    with open(csv_var, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        w.writerow(cab_var)
        w.writerows(datos_var)
    # Misma logica que scripts/csv_para_excel.py (que solo admite rutas fijas): ';' y UTF-8 con BOM.
    csv_xl = SALIDA / 'disponibilidad_variables_excel.csv'
    with open(csv_var, encoding='utf-8', newline='') as fin, open(csv_xl, 'w', encoding='utf-8-sig', newline='') as fout:
        esc = csv.writer(fout, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        for fila in csv.reader(fin):
            esc.writerow(fila)

    print(f'Correcciones aplicadas: {n} operaciones sobre celdas (C1-C18 y C21; C19 validada; C20 es de documentacion)')
    print('Conteo por disponibilidad:', conteo)
    print('Veredictos:', veredictos)
    print(f'Variables {len(datos_var)} · Indicadores {len(datos_ind)} · Fuentes {len(datos_fte)} · '
          f'Solicitudes {len(datos_sol)} · Enlaces {len(tabla_enlaces)}')
    print(f'Generado {destino}\nGenerado {csv_var}\nGenerado {csv_xl}')


if __name__ == '__main__':
    main()
