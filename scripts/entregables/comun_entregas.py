"""Funciones y catálogos comunes de las entregas al profe (Word detallado, ficha metodológica y Excel).

Reúne: el reemplazo de tildes, la limpieza de textos para lectores externos (sin números de decisión ni atribuciones),
los catálogos de unidad y sentido de cada indicador y los lectores de los CSV de construcción.
Los Word de la entrega resumida siguen en generar_entrega_profe.py; este módulo importa sus utilidades y no las duplica.
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

RAIZ = Path(__file__).resolve().parents[2]

# ---------------------------------------------------------------- catálogos compartidos con el Excel (generar_base_powerbi.py)
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

PCT, PUNT, TASA = '%', 'puntaje 0–100', 'tasa por 100.000 hab.'
UNIDAD_BASE = {
    'PDM.ICPD7': PUNT, 'PDM.SALUD': PCT, 'PDM.ILE': PUNT, 'PDM.SP': PCT, 'PDM.CM': TASA, 'PDM.CA': 'm²/hab',
    'PDM.VEED': 'conteo (veedurías inscritas)', 'PDM.HOM': TASA, 'PC.ISPC': PUNT, 'PC.I1': '% (IPCM: votos / población)',
    'PC.I2': PCT, 'PC.I3': PCT, 'PC.I4': PCT, 'PP.IREPP': PUNT, 'PP.EPP': PCT, 'PP.EJ90': PCT, 'PP.FID': PCT,
    'CS.ICRCS': PUNT, 'CS.PQRSDF': PCT, 'CS.DSI': PCT, 'CS.VRC': PCT, 'CS.RPO': PCT,
}
MENOR_ES_MEJOR = {'PDM.CM', 'PDM.HOM', 'PDM.CM.TD', 'PDM.HOM.TD'}

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
    'asi': 'así', 'transito': 'tránsito', 'ampliacion': 'ampliación', 'produccion': 'producción', 'numeracion': 'numeración', 'debora': 'Débora', 'capacitacion': 'capacitación', 'actualizacion': 'actualización', 'proporcion': 'proporción', 'grafico': 'gráfico', 'organizaciones': 'organizaciones', 'sera': 'será', 'unicamente': 'únicamente', 'ultimos': 'últimos', 'numeros': 'números', 'biologica': 'biológica', 'ecologica': 'ecológica', 'tematica': 'temática', 'cronologico': 'cronológico',
    'semaforo': 'semáforo', 'leidas': 'leídas', 'leidos': 'leídos', 'leida': 'leída', 'leido': 'leído', 'metodo': 'método', 'metodos': 'métodos',
    'vacios': 'vacíos', 'vacio': 'vacío', 'minima': 'mínima', 'minimas': 'mínimas', 'homonimo': 'homónimo', 'automatico': 'automático',
    'automaticos': 'automáticos', 'automatica': 'automática', 'ningun': 'ningún', 'algun': 'algún', 'busqueda': 'búsqueda', 'titulos': 'títulos',
    'titulo': 'título', 'ano': 'año', 'encontro': 'encontró', 'divulgacion': 'divulgación', 'planificacion': 'planificación',
    'socializacion': 'socialización', 'posicion': 'posición', 'deliberacion': 'deliberación', 'periodico': 'periódico', 'realizara': 'realizará',
    'derogo': 'derogó', 'explicacion': 'explicación', 'boletin': 'boletín', 'catedra': 'cátedra', 'artisticas': 'artísticas', 'opinion': 'opinión',
    'dialogo': 'diálogo', 'respondio': 'respondió', 'estan': 'están', 'tramites': 'trámites', 'condicion': 'condición', 'priorizo': 'priorizó',
    'excluyo': 'excluyó', 'homogenea': 'homogénea', 'cuantas': 'cuántas', 'aun': 'aún', 'mas': 'más', 'demas': 'demás', 'dejo': 'dejó', 'cuanto': 'cuánto',
    'periodicidad': 'periodicidad', 'unicos': 'únicos', 'estadistica': 'estadística', 'estadisticas': 'estadísticas', 'basico': 'básico',
    'basica': 'básica', 'basicos': 'básicos', 'practica': 'práctica', 'sintesis': 'síntesis', 'indicio': 'indicio', 'extension': 'extensión',
    'extemporanea': 'extemporánea', 'extemporaneas': 'extemporáneas', 'subestima': 'subestima', 'util': 'útil', 'utiles': 'útiles',
    'lectura': 'lectura', 'oportunidad': 'oportunidad', 'electronico': 'electrónico', 'telefono': 'teléfono', 'codigo': 'código',
    'codigos': 'códigos', 'contraccion': 'contracción', 'verificacion': 'verificación', 'rapido': 'rápido', 'valido': 'válido',
    'validos': 'válidos', 'tecnicos': 'técnicos', 'historico': 'histórico', 'historica': 'histórica', 'numerica': 'numérica',
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


# ---------------------------------------------------------------- limpieza de textos para lectores externos
PROHIBIDAS = ('origen', 'claude', 'asistente')  # no pueden aparecer en los archivos para el profe (se busca como fragmento de palabra, sin distinguir mayúsculas)


def palabras_prohibidas(texto):
    return [p for p in PROHIBIDAS if p in texto.lower()]
_SUSTITUCIONES = [
    (r'por D\d{1,3} \((?:decision|decisión) de la usuaria\):', 'por regla,'),
    (r'\bD\d{1,3} \((?:decision|decisión) de la usuaria\):', 'Regla:'),
    (r'por D86 va al MAXIMO', 'va al MAXIMO'),
    (r'Archivo obtenido del ZIP original de Google Drive de la usuaria porque el portal daba HTTP 500; el enlace es el de origen en el portal\.',
     'Archivo obtenido de la copia de respaldo del observatorio porque el portal no respondió (error HTTP 500); el enlace es el del portal.'),
    (r'\s*\((?:decision|decisión) de la usuaria\)', ''),
    (r'Convencion \(QA\)', 'Convención'),
    (r'Convención \(QA\)', 'Convención'),
    (r'\bde la usuaria\b', 'del observatorio'),
    (r'\bla usuaria\b', 'el observatorio'),
    (r'\bdecision de\b', 'decisión de'),
    (r'\basistentes\b', 'participantes'),
    (r'\bAsistentes\b', 'Participantes'),
    (r'\b(audiencia|Audiencia) publica\b', r'\1 pública'),
    (r'\b(?:E|e)spacio publico\b', lambda m: m.group(0)[:-7] + 'público'),
    (r'\bno se voto\b', 'no se votó'),
    (r'\bamplio (los plazos|el plazo|plazos)', r'amplió \1'),
    (r'\bno dejo\b', 'no dejó'),
    (r'\bse publico\b', 'se publicó'),
    (r'ver url_origen de cada ficha en (?:data_raw_IIC|outputs)/[\w/\.\-]+', 'ver el enlace de cada ficha en la documentación de trabajo del observatorio'),
    (r'url_origen', 'enlace de la fuente'),
    (r'\s*\((?:decisiones|decision|decisión) D\d{1,3}(?:\s*[,y/]\s*D\d{1,3})*\)', ''),
    (r'\s*\[(?:S|C|N)\]', ''),   # marcas internas de verificación
    # rutas y archivos de trabajo: el profe no los necesita
    (r'\s*\((?:clasificacion en )?outputs/[\w/\.\-]+\)', ''),
    (r'(?:\(|; )?(?:outputs|scripts|data_raw_IIC)/[\w/\.\-]+\)?', 'la documentación de trabajo del observatorio'),
    (r'\b[\w\-]+\.(?:csv|py)\b(?!\?)', 'la documentación de trabajo del observatorio'),
]
_RE_DREF = re.compile(r'\bD\d{1,3}\b')


def quitar_decisiones(t):
    """Quita las referencias internas a decisiones (D35, D47...) y deja la frase en orden."""
    t = re.sub(r'\s*\(\s*D\d{1,3}(?:\s*[,y]\s*D\d{1,3})*\s*\)', '', t)           # (D35) o (D34, D58)
    t = re.sub(r'\(\s*D\d{1,3}(?:\s*[,y]\s*D\d{1,3})*\s*[:;]\s*', '(', t)        # (D15, D16: texto) -> (texto)
    t = re.sub(r'\s*[;,]\s*D\d{1,3}(?:\s*[,y]\s*D\d{1,3})*(?=\s*[;)\.,]|\s*$)', '', t)   # ...; D35)  /  ..., D58
    t = re.sub(r'\s*\bD\d{1,3}(?:\s*[,y]\s*D\d{1,3})*\s*:\s*', ' ', t)           # «D56: texto»
    t = re.sub(r'\s*\bD\d{1,3}(?:\s*[,y]\s*D\d{1,3})*\b', '', t)                 # restos
    t = re.sub(r'\(\s*\)', '', t)
    t = re.sub(r'[ \t]+([;,.)])', r'\1', t)
    t = re.sub(r'\([ \t]+', '(', t)
    t = re.sub(r'[ \t]{2,}', ' ', t)
    return t.strip()


def coma_decimal(t):
    """Pasa 12.5 a 12,5 en textos de notas (no toca direcciones web)."""
    trozos = re.split(r'(https?://\S+)', t)
    for i in range(0, len(trozos), 2):
        trozos[i] = re.sub(r'(?<![\w.])(\d{1,3}(?:\.\d{3})+)(?![\w])', lambda m: m.group(1), trozos[i])  # miles: se dejan
        trozos[i] = re.sub(r'(?<![\w.])(\d+)\.(\d{1,2})(?![\d\w])', r'\1,\2', trozos[i])
    return ''.join(trozos)


def _limpiar_trozo(t, decimales):
    if not t.strip():
        return t
    t = re.sub(r'\bano-(\d{4}-\d+)', r'ANOSLUG\1', t)   # el nombre de página «ano-2023-855308» del portal no se corrige
    t = _limpiar_trozo2(t, decimales)
    return t.replace('ANOSLUG', 'ano-')


def _limpiar_trozo2(t, decimales):
    ini, fin = re.match(r'\s*', t).group(), re.search(r'\s*$', t).group()
    t = t.strip()
    for pat, rep in _SUSTITUCIONES:
        t = re.sub(pat, rep, t)
    t = quitar_decisiones(t)
    t = tildes(t)
    t = coma_decimal(t) if decimales else t
    return ini + t + fin if t else ini + fin


def para_profe(texto, decimales=False):
    """Texto apto para el profe: con tildes, sin números de decisión ni atribuciones a personas. No toca las direcciones web."""
    if not texto:
        return texto
    trozos = re.split(r'(https?://[^\s;)\]]+)', texto)
    for i in range(0, len(trozos), 2):
        trozos[i] = _limpiar_trozo(trozos[i], decimales)
    return ''.join(trozos)


# ---------------------------------------------------------------- utilidades de Word y lectores de los CSV (de generar_entrega_profe.py)
import csv  # noqa: E402

import generar_entrega_profe as gep  # noqa: E402

DATA = gep.DATA
fmt, num, filas, anual, con_dato, comprimir = gep.fmt, gep.num, gep.filas, gep.anual, gep.con_dato, gep.comprimir
FECHA = gep.FECHA
DIM_ORDEN = ['PDM', 'PC', 'PP', 'CS']
NOMBRE_DIM = {'PDM': 'Cumplimiento del Plan de Desarrollo Municipal (PDM)', 'PC': 'Participación ciudadana',
              'PP': 'Presupuesto participativo', 'CS': 'Control social'}
INDICE_DIM = {'PDM': 'PDM.ICPD7', 'PC': 'PC.ISPC', 'PP': 'PP.IREPP', 'CS': 'CS.ICRCS'}
NOMBRE_INDICE = {'PDM.ICPD7': 'ICPD-7', 'PC.ISPC': 'ISPC', 'PP.IREPP': 'IREPP', 'CS.ICRCS': 'ICRCS'}
SEMAFORO_TXT = ('Verde ≥ 90; Amarillo 75 a < 90; Naranja 50 a < 75; Rojo < 50, aplicado al puntaje de 0 a 100. '
                'Es una propuesta, por validar con el profesor.')
ESTADO = 'Provisional'
PRIMER_ANIO = {'PDM.ICPD7': 2019, 'PC.ISPC': 2020, 'PP.IREPP': 2020, 'CS.ICRCS': 2020}


def dim_de(cod):
    return cod.split('.')[0]


def principales(dim):
    """Códigos de los indicadores principales de la dimensión, en el orden del CSV."""
    out = []
    for r in DATA[dim]:
        if r['tipo'] == 'principal' and r['codigo'] not in out:
            out.append(r['codigo'])
    return out


def val(cod, anio, campo='valor'):
    for r in DATA[dim_de(cod)]:
        if r['codigo'] == cod and r['anio'] == str(anio):
            return r[campo]
    raise KeyError((cod, anio, campo))


def vf(cod, anio, campo='valor', dec=2):
    """Valor del CSV con coma decimal."""
    return fmt(val(cod, anio, campo), dec)


def texto_csv(cod):
    """Referente, nota y fuente de todas las filas del código (sin limpiar), para verificar cifras citadas."""
    return ' '.join(' '.join((r['referente'], r['nota'], r['fuente'], r['brecha'], r['enlace'])) for r in DATA[dim_de(cod)]
                    if r['codigo'] == cod)


def req(cod, *fragmentos):
    """Comprueba que cada cifra o frase citada en un texto escrito a mano exista en el CSV del indicador."""
    t = texto_csv(cod)
    for f in fragmentos:
        assert f in t, (cod, f)


def peso_txt(cod):
    rs = [r for r in DATA[dim_de(cod)] if r['codigo'] == cod and r['peso_en_indice']]
    assert rs, cod
    p = num(rs[0]['peso_en_indice'])
    if abs(p - 1 / 7) < 1e-4:
        return '1/7 (0,1429), igual para los siete componentes del ICPD-7'
    return f'{fmt(p * 100, 0)} % ({fmt(p, 2)})'


def unidad_sentido(cod):
    base = '.'.join(cod.split('.')[:2])
    unidad = UNIDAD_BASE.get(base, PCT)
    sentido = 'menor es mejor' if base in MENOR_ES_MEJOR else 'mayor es mejor'
    return unidad, sentido


def meses_fecha(cod):
    fechas = re.findall(r'\d{4}-\d{2}-\d{2}', ' '.join(r['fecha_consulta'] for r in DATA[dim_de(cod)] if r['codigo'] == cod))
    return max(fechas) if fechas else ''


def fuentes_lista(cod, max_fuentes=6, max_enlaces=2):
    """Ternas (fuente, [enlaces], enlaces adicionales) distintas del indicador, limpias para el profe."""
    grupos = {}
    for r in con_dato(anual(dim_de(cod), cod, 'principal')):
        f = re.split(r'\s+Archivo obtenido', r['fuente'])[0]
        if not f:
            continue
        clave = re.sub(r'\d{4}', '', f)[:80]
        g = grupos.setdefault(clave, [para_profe(f), []])
        urls = re.findall(r'https?://[^\s;]+', r['enlace'])
        urls = [u for u in urls if '?' not in u] or urls[:1]   # se omiten las direcciones con consultas (parámetros) muy largas
        for u in urls:
            if u not in g[1]:
                g[1].append(u)
    salida = [(f, enl[:max_enlaces], max(len(enl) - max_enlaces, 0)) for f, enl in list(grupos.values())[:max_fuentes]]
    if salida and not any(e for _f, e, _m in salida):   # sin enlaces propios: se toman los del detalle por evento (p. ej. VRC)
        det, cod_det = [], ''
        for r in DATA[dim_de(cod)]:
            if r['tipo'] == 'detalle' and r['codigo'].startswith(cod + '.'):
                cod_det = r['codigo']
                det += [u for u in re.findall(r'https?://[^\s;]+', r['enlace']) if u not in det]
        if det:
            salida.append((f'Documentos de cada evento (acta, memoria o cuadro de preguntas): detalle en las filas {cod_det} de la hoja «Detalle» del Excel',
                           det[:max_enlaces], max(len(det) - max_enlaces, 0)))
    return salida


def serie_texto(dim, cod):
    return gep.anios_dato(dim, cod)


def brecha_texto(dim, cod):
    t = para_profe(gep.texto_brecha(dim, cod, corto=False), decimales=True)
    if cod in NOMBRE_INDICE:
        t = t.replace('años sin dato', 'años sin índice')
    t = re.sub(r'\)\s*;\s*(\d{4}(?:, \d{4})*)(\s*;|\s*$)', r'): \1\2', t)      # «(33,3 %); 2020, 2022;» -> «(33,3 %): 2020, 2022»
    t = t.replace('eventos desconocidos', 'eventos sin información sobre veedurías')
    return t.strip().rstrip(';').strip()


def valores_indice(cod):
    """{año: (valor float o None, semáforo)} del índice (hasta 2025)."""
    out = {}
    for r in anual(dim_de(cod), cod, 'indice'):
        v = num(r['valor'])
        out[int(r['anio'])] = (v, r['semaforo'] if v is not None else '')
    return out


def resumen_indices(a0=2019, a1=2025):
    """Filas [año, {cod: (valor, semáforo)}] para la tabla de los cuatro índices."""
    vals = {c: valores_indice(c) for c in NOMBRE_INDICE}
    return [(a, {c: vals[c].get(a, (None, '')) for c in NOMBRE_INDICE}) for a in range(a0, a1 + 1)]


def anios_con_indice(cod):
    return sorted(a for a, (v, _s) in valores_indice(cod).items() if v is not None)


def serie_indice_txt(cod):
    """«2019: 62,41 (Naranja); 2022: ...» a partir del CSV."""
    return '; '.join(f'{a}: {fmt(v)} ({s})' for a, (v, s) in sorted(valores_indice(cod).items()) if v is not None)


def matriz():
    """Indicadores de la matriz original (outputs/disponibilidad/trabajo/matriz.csv), por código único."""
    ruta = RAIZ / 'outputs' / 'disponibilidad' / 'trabajo' / 'matriz.csv'
    with open(ruta, encoding='utf-8-sig', newline='') as f:
        return {r['Código único']: r for r in csv.DictReader(f)}


def una_linea(t):
    return re.sub(r'\s*\n\s*', ' ', t or '').strip()


def pp_anios_sin_dato():
    """[(año, motivo corto)] de los años 2020–2025 sin IREPP, según el CSV de PP."""
    out = []
    for r in anual('PP', 'PP.IREPP', 'indice'):
        if num(r['valor']) is None and 2020 <= int(r['anio']) <= 2025:
            m = re.sub(r'^sin dato:\s*', '', r['brecha'])
            m = gep.quitar_parentesis(m)
            out.append((int(r['anio']), para_profe(m.split(';')[0].strip(), decimales=True)[:200]))
    return out


def vals_i1():
    """{año: votos} de I1 con dato."""
    return {int(r['anio']): r['numerador'] for r in anual('PC', 'PC.I1', 'principal') if num(r['valor']) is not None}


def ano_suelto(texto):
    """Apariciones de la palabra «ano» sin tilde (error de ASCII por «año»), salvo el nombre de página «ano-2023-…» del portal."""
    t = re.sub(r'https?://\S+', '', texto)
    return re.findall(r'.{20}\bano\b(?!-\d).{10}', t, flags=re.IGNORECASE | re.DOTALL)


def enlace_reciente(cod):
    """Primera dirección web de la fuente del año más reciente con dato (sin consultas largas); vacío si no hay."""
    for r in reversed(con_dato(anual(dim_de(cod), cod, 'principal'))):
        urls = [u for u in re.findall(r'https?://[^\s;]+', r['enlace']) if '?' not in u]
        if urls:
            return urls[0]
    return ''
