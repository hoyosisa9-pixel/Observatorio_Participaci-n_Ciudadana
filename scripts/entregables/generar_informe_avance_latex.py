"""Genera el informe de avance (Avance 1) del sistema de indicadores del OCE en LaTeX (Overleaf / tectonic).

Uso (desde la raiz del repo):  python3 scripts/entregables/generar_informe_avance_latex.py
Reutiliza por import las funciones de estilo de scripts/generar_informe_latex.py (escape Unicode, enlace(), xltabular,
aviso, num/pct) sin modificarlas. Todas las cifras salen de los CSV de outputs/indicadores/ (no se recalcula nada);
los textos explicativos y las tablas de decisiones salen de resumen_calculos.md y revision_critica.md.

Entradas: outputs/indicadores/{resultados_indicadores,valores_variables,pp_proyectos,soporte_conteos,
          clasificacion_*,diccionario_datos}.csv, resumen_calculos.md, revision/revision_critica.md,
          config/supuestos.yaml, segunda_lectura/*.csv y la hoja Solicitudes de
          outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx
Salidas:  outputs/indicadores/entregables/informe_avance_1.tex  y  informe_avance_1_overleaf.zip
"""
import csv, re, sys, zipfile
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import openpyxl
import yaml

RAIZ = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(RAIZ / 'scripts'))
import generar_informe_latex as L  # noqa: E402  (estilo, escape y enlaces del informe de disponibilidad)

IND = RAIZ / 'outputs' / 'indicadores'
SALIDA = IND / 'entregables'
F_TEX = SALIDA / 'informe_avance_1.tex'
F_ZIP = SALIDA / 'informe_avance_1_overleaf.zip'
F_XLSX = RAIZ / 'outputs' / 'disponibilidad' / 'disponibilidad_brecha_indicadores.xlsx'

TITULO = ('Sistema de indicadores del Observatorio Ciudadano de Envigado: '
          'cálculo provisional de indicadores viables')
TITULO_CORTO = 'OCE -- Avance 1: cálculo provisional de indicadores'
VERSION = 'Avance 1 – 9 de noviembre de 2026'
PROFESORES = 'Dirigido a: profesores Francisco Correa y Eber Gutiérrez'

esc, num, pct, enlace, aviso, xltabular = L.esc, L.num, L.pct, L.enlace, L.aviso, L.xltabular

COLOR_SEM = {'Verde': 'cverde', 'Amarillo': 'camarillo', 'Naranja': 'cnaranja', 'Rojo': 'crojo', 'Gris': 'cgris'}
MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre',
         'noviembre', 'diciembre']
SISTEMAS = [('CS', 'Control social'), ('PC', 'Participación ciudadana'), ('PDM', 'Cumplimiento del PDM'),
            ('PP', 'Presupuesto participativo')]
# indicadores viables (se calculan); el orden es el del informe
VIABLES = ['CS.PCS', 'CS.DSI', 'CS.VRC', 'CS.RPO', 'CS.ICRCS', 'PC.I2', 'PC.I3', 'PC.I5', 'PDM.CA', 'PDM.CM', 'PP.EPP']
NO_VIABLES = ['PC.I1', 'PC.I4', 'PC.I6', 'PDM.CS', 'PDM.CP', 'PDM.CD', 'PDM.ILE', 'PP.CEP', 'PP.CMA']
INFO = {  # nombre corto, formula de la Matriz, entradas, umbrales del semaforo (hoja Matriz)
    'CS.PCS': ('Promoción del control social (PCS)', 'PCS = 100 × EE / EP, con tope de 100. EE: estrategias de promoción '
               'ejecutadas; EP: estrategias programadas.', 'EE / EP', 'Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50'),
    'CS.DSI': ('Disponibilidad de información de seguimiento (DSI)', 'DSI = 100 × MD / MS. MD: metas seleccionadas con '
               'valor observado publicado para la vigencia; MS: metas seleccionadas.', 'MD / MS',
               'Verde ≥ 90; Amarillo 70 a < 90; Rojo < 70'),
    'CS.VRC': ('Presencia de veedurías en la rendición de cuentas (VRC)', 'VRC = 100 × EV / ET. EV: eventos municipales '
               'de rendición de cuentas con presencia documentada de veedurías; ET: eventos realizados. Si hay eventos '
               'con presencia desconocida, la Matriz indica N/D.', 'EV / ET', 'Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50'),
    'CS.RPO': ('Respuesta pública a observaciones ciudadanas (RPO)', 'RPO = 100 × OR / OT. OR: observaciones con '
               'respuesta de fondo publicada; OT: observaciones ciudadanas registradas públicamente.', 'OR / OT',
               'Verde ≥ 90; Amarillo 70 a < 90; Rojo < 70'),
    'CS.ICRCS': ('Índice de condiciones y respuesta para el control social (ICRCS)', 'ICRCS = (DSI + PCS + VRC + RPO) / 4. '
                 'Exige las cuatro entradas con número; la Matriz no permite redistribuir pesos.',
                 'DSI; PCS; VRC; RPO', 'Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50'),
    'PC.I2': ('I2 Cobertura territorial de la participación', 'I2 = 100 × V1 / V2. V1: zonas con al menos una actividad '
              'participativa realizada; V2: zonas oficiales del municipio.', 'V1 / V2',
              'Verde ≥ 80; Amarillo 50 a < 80; Rojo < 50'),
    'PC.I3': ('I3 Cumplimiento de la oferta formativa', 'I3 = 100 × V1 / V2. V1: acciones formativas programadas que se '
              'realizaron; V2: acciones formativas programadas con finalización prevista en el periodo.', 'V1 / V2',
              'Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50'),
    'PC.I5': ('I5 Cumplimiento de sesiones de las instancias de participación', 'I5 = 100 × V1 / V2. V1: sesiones '
              'realizadas; V2: sesiones previstas de las instancias seleccionadas.', 'V1 / V2',
              'Verde ≥ 75; Amarillo 50 a < 75; Rojo < 50'),
    'PDM.CA': ('Cumplimiento del espacio público verde (CA)', 'CA = 100 × O / M, con tope de 100. O: espacio público verde '
               'observado (m² por habitante); M: meta (m² por habitante).', 'O / M',
               'Verde: 100 sin redondear; Amarillo 75 a < 100; Rojo < 75'),
    'PDM.CM': ('Cumplimiento de la mortalidad vial (CM)', 'CM = 100 × M / O, con tope de 100 (100 si O = 0). O: tasa '
               'observada de mortalidad vial por 100.000 habitantes; M: tasa meta.', 'O / M',
               'Verde: 100 sin redondear; Amarillo 75 a < 100; Rojo < 75'),
    'PP.EPP': ('Ejecución presupuestal del presupuesto participativo (EPP)', 'EPP = 100 × OPP / APP. OPP: obligaciones '
               'presupuestales con cargo al presupuesto participativo; APP: apropiación definitiva.', 'OPP / APP',
               'Verde 90 a 100; Amarillo 75 a < 90; Naranja 50 a < 75; Rojo < 50'),
}
NOMBRE_NV = {'PC.I1': 'PC – I1 (incorporación de aportes)', 'PC.I4': 'PC – I4 (publicación de resultados)',
             'PC.I6': 'PC – I6 (participación virtual)', 'PDM.CS': 'PDM – Cobertura en salud (CS)',
             'PDM.CP': 'PDM – Convivencia y paz (CP)', 'PDM.CD': 'PDM – Deporte (CD)',
             'PDM.ILE': 'PDM – Primera infancia (ILE)', 'PP.CEP': 'PP – Cumplimiento de entregas (CEP)',
             'PP.CMA': 'PP – Metas de atención (CMA)'}
# escenarios (alternativas) mostrados por indicador: (escenario_id exacto, etiqueta)
ESCENARIOS = {
    'CS.DSI': [('ventana_documental_cs=hasta_fecha_consulta', 'Aceptar publicaciones hasta la fecha de consulta'),
               ('dsi_fuente_valor=oficial_incluye_proxy', 'Aceptar también cifras de apoyo (aproximadas)'),
               ('dsi_lista_metas=icpd7_mas3_gobernanza', '10 metas (con tres de gobernanza)'),
               ('dsi_lista_metas=productos_plan_indicativo', 'Productos del plan indicativo (solo control)')],
    'CS.VRC': [('vrc_evento_sin_mencion=cero_siempre', 'Evento sin mención de veedurías cuenta como ausencia'),
               ('control_institucional_545_546', 'Control: solo audiencia del Alcalde y evento de veedurías (consecutivos 545 y 546)')],
    'CS.RPO': [('rpo_criterio_fondo=laxo_texto_respuesta', 'Criterio laxo de respuesta de fondo'),
               ('rpo_filtro_pdm=tematico_pdm', 'Solo observaciones temáticas del plan de desarrollo'),
               ('rpo_filtro_pdm=todas', 'Todas las observaciones, con felicitaciones'),
               ('rpo_universo=alcalde_y_100dias', 'Audiencia del Alcalde y evento de los 100 días'),
               ('rpo_universo=todos_los_cuadros', 'Todos los cuadros de preguntas publicados'),
               ('rpo_registro_incompleto=calcular_con_publicado', 'Calcular con lo publicado aunque el cuadro esté incompleto')],
    'CS.ICRCS': [('vrc_evento_sin_mencion=cero_siempre', 'VRC con evento sin mención contado como ausencia'),
                 ('control_institucional_545_546', 'Control: consecutivos 545 y 546 en los cuatro componentes'),
                 ('envolvente_componentes_min', 'Rango por componente: mínimos'),
                 ('envolvente_componentes_max', 'Rango por componente: máximos')],
    'PC.I2': [('pc_i2_actividades=asambleas_pp', 'Solo asambleas de presupuesto participativo'),
              ('zonas_oficiales=12_con_comite', '12 zonas con comité zonal'),
              ('pc_i2_evidencia=incluye_programacion', 'Incluir asambleas programadas del POT'),
              ('pc_i2_evidencia=acta_por_zona', 'Exigir acta por zona')],
    'PC.I3': [('pc_i3_excluir_servidores=false', 'Incluir la formación para servidores'),
              ('pc_i3_unidad=personas_producto_521', 'Medir en personas')],
    'PC.I5': [('pc_i5_instancias=ampliado_cmpc_ctp_comites', 'Instancias ampliadas')],
    'PDM.CA': [('pdm_ca_fuente=pot_epv_total', 'Total del POT (no comparable con la meta urbana)'),
               ('pdm_ca_fuente=estricto_serie_linea_base', 'Estricto con la serie de la línea base')],
    'PDM.CM': [('pdm_cm_fuente=policia_datosgov;poblacion_dane=dane_postcovid_2020_2035',
                'Datos de la Policía y población DANE posterior a la pandemia')],
    'PP.EPP': [('pp_2024_metodo=solo_exactos', 'Solo cruces exactos'),
               ('pp_incluir_0553=true', 'Incluir el costo del proceso (proyecto 0553)'),
               ('pp_fuente_calculo=cuipo', 'Fuente CUIPO')],
}


# --------------------------------------------------------------------------- lectura
def leer_csv(nombre):
    with open(IND / nombre, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f, delimiter=';'))


def fecha_es(iso, corta=False):
    y, m, d = (int(x) for x in iso.split('-'))
    return f'{d} de {MESES[m - 1]} de {y}'


def fnum(x, dec=2):
    """Numero con coma decimal; enteros sin decimales."""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return 'N/D' if x in ('', None) else str(x)
    if abs(v - round(v)) < 1e-9 and abs(v) >= 1:
        return num(v, 0)
    if abs(v) < 1e-9:
        return '0'
    return num(v, dec)


def lim(t):
    """Limpieza de lenguaje de los textos de los documentos de trabajo (md)."""
    t = t.replace('"Validado"', 'validado').replace('Validado', 'validado')
    t = re.sub(r';? fuera de las opciones configuradas', '', t)
    t = t.replace(' (fuera de las opciones configuradas)', '').replace('(fuera del YAML)', '')
    t = t.replace('en el repositorio', 'en el archivo del Observatorio')
    t = re.sub(r'`([^`]*)`', r'\1', t)
    return t


def md_cel(t):
    return L.md_inline(lim(t))


def seccion_md(texto, prefijo):
    lin = texto.splitlines()
    i = next(k for k, l in enumerate(lin) if l.startswith(prefijo))
    fin = next((k for k in range(i + 1, len(lin)) if lin[k].startswith('## ')), len(lin))
    return '\n'.join(lin[i + 1:fin])


# nombres presentables de fuentes que en el CSV vienen como ruta
def nombre_fuente(fuente, enlace_):
    if '/' not in fuente and 'soporte_conteos' not in fuente:
        return fuente
    claves = [('ejecucion-presupuestal-egresos-al-31122024', 'Secretaría de Hacienda, ejecución presupuestal de egresos al 31-12-2024'),
              ('evaluaciondefinitivaplanaccionivtri', 'Plan de Acción, evaluación definitiva IV trimestre 2025 (Planeación)'),
              ('evaluaciondefinitivaplanaccion2026segundotrimestre', 'Plan de Acción 2026, evaluación definitiva 2T-2026 (Planeación)'),
              ('apendice91planindicativo', 'Apéndice 9.1.2 del Plan Indicativo del PDM 2024-2027')]
    if 'soporte_conteos' in fuente:
        return 'Soporte meta por meta (base de datos, hoja Soporte de conteos)'
    for k, v in claves:
        if k in enlace_ or k in fuente:
            return v
    if 'soporte_conteos' in fuente:
        return 'Soporte meta por meta (base de datos, hoja Soporte de conteos)'
    raise ValueError('fuente sin nombre presentable: ' + fuente)


PAG_REEMPL = [(r'Ver ids_soporte \([^)]*\)\.?', 'detalle por elemento en la hoja Soporte de conteos de la base de datos'),
              (r'Ver ids_soporte en soporte_conteos\.csv', 'detalle por elemento en la hoja Soporte de conteos de la base de datos'),
              (r'\bPI26\b', 'Plan Indicativo de Productos 2T-2026'), (r'\bG25\b', 'Informe de gestión 2025'),
              (r'\bPDM\b consecutivo', 'PDM consecutivo')]


def pagina_limpia(p):
    for a, b in PAG_REEMPL:
        p = re.sub(a, b, p)
    assert not re.search(r'\.csv|\.py|ids_soporte|data_raw', p), p
    return p


def enlace_texto(url, texto):
    """Texto clicable, palabra por palabra (cada palabra es un \\href al mismo destino y puede partirse en lineas)."""
    destino = (url.replace('\\', '/').replace('%', r'\%').replace('#', r'\#').replace('&', r'\&').replace('~', r'\~'))
    return ' '.join(r'\href{' + destino + '}{' + L.cortable(w) + '}' for w in texto.split())


# --------------------------------------------------------------------------- datos
class Datos:
    def __init__(self):
        self.R = leer_csv('resultados_indicadores.csv')
        self.V = leer_csv('valores_variables.csv')
        self.P = leer_csv('pp_proyectos.csv')
        self.S = leer_csv('soporte_conteos.csv')
        self.C24 = leer_csv('clasificacion_pp_2024.csv')
        self.CE = leer_csv('clasificacion_pp_entregas.csv')
        self.D = leer_csv('diccionario_datos.csv')
        self.md = (IND / 'resumen_calculos.md').read_text(encoding='utf-8')
        self.rev = (IND / 'revision' / 'revision_critica.md').read_text(encoding='utf-8')
        self.sup = yaml.safe_load((RAIZ / 'config' / 'supuestos.yaml').read_text(encoding='utf-8'))
        self.vrc_muestra = leer_csv_ruta(IND / 'segunda_lectura' / 'muestra_vrc.csv')
        self.rpo_muestra = leer_csv_ruta(IND / 'segunda_lectura' / 'muestra_rpo.csv')
        self.clas_vrc = leer_csv('clasificacion_vrc.csv')
        self.clas_rpo = leer_csv('clasificacion_rpo.csv')
        self.idx = {}
        for r in self.R:
            if r['territorio'] == 'Municipio':
                clave = (r['codigo'], r['anio'], r['periodo'], 'defecto' if r['es_defecto'] == 'True' else r['escenario_id'])
                self.idx.setdefault(clave, r)
        self.zona = {}
        for r in self.R:
            if r['codigo'] == 'PP.EPP' and r['es_defecto'] == 'True':
                self.zona[(r['territorio'], r['anio'], r['periodo'])] = r

    def fila(self, cod, anio, periodo='Anual', esc_id='defecto'):
        return self.idx.get((cod, str(anio), periodo, esc_id))


def leer_csv_ruta(p):
    with open(p, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f, delimiter=';'))


def celda_resultado(r):
    """(texto de puntaje, texto de semaforo, color) de una fila de resultados."""
    disp, sem = r['disponibilidad'], r['semaforo']
    if disp == 'Calculado':
        pt = fnum(r['puntaje'])
        if sem == 'Gris':
            return pt, 'Sin semáforo', None
        return pt, sem, COLOR_SEM[sem]
    if disp == 'Revisar':
        return 'Revisar', 'Sin color', None
    if disp == 'NE':
        return 'NE', 'Gris', 'cgris'
    return 'N/D', 'Gris', 'cgris'


def sc(color, texto):
    return (r'\cellcolor{' + color + '}' if color else '') + texto


def entradas_txt(r, cod):
    e = [r[f'entrada_{k}'] for k in (1, 2, 3, 4)]
    if cod == 'CS.ICRCS':
        return '; '.join(f'{n} {fnum(v) if v != "" else "N/D"}' for n, v in zip(('DSI', 'PCS', 'VRC', 'RPO'), e))
    if r['disponibilidad'] == 'N/D' and not any(x not in ('', 'N/D') for x in e[:2]):
        return '--'
    a, b = e[0], e[1]
    if cod == 'PP.EPP':
        f = lambda x: r'\$' + num(x, 0) if x not in ('', 'N/D') else 'N/D'
        return f(a) + ' / ' + f(b)
    f = lambda x: fnum(x) if x != '' else 'N/D'
    return f(a) + ' / ' + f(b)


def anio_txt(r):
    if r['periodo'] == 'Semestre 1':
        return r['anio'] + ' (primer semestre)'
    return r['anio']


# --------------------------------------------------------------------------- bloques
def construir():
    d = Datos()
    T = []
    a = T.append
    cortes = sorted({v['fecha_corte_real'] or v['fecha_corte'] for v in d.V
                     if v['es_defecto'] == 'True' and v['estado_informacion'] == 'Pendiente' and v['codigo_indicador'] in VIABLES
                     and (v['fecha_corte_real'] or v['fecha_corte'])})
    corte_min, corte_max = cortes[0], cortes[-1]
    consulta = max(v['fecha_consulta'] for v in d.V if v['fecha_consulta'])

    a(r'''\documentclass[11pt,letterpaper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\usepackage[spanish,es-noshorthands,es-tabla]{babel}
\usepackage[margin=2.5cm,headheight=14pt]{geometry}
\usepackage[table]{xcolor}
\usepackage{booktabs}
\usepackage{xltabular}
\usepackage{array}
\usepackage{enumitem}
\usepackage{fancyhdr}
\usepackage[hidelinks,breaklinks=true]{hyperref}
\usepackage{url}

\definecolor{cverde}{RGB}{214,240,214}
\definecolor{camarillo}{RGB}{255,243,191}
\definecolor{cnaranja}{RGB}{255,222,184}
\definecolor{crojo}{RGB}{248,212,212}
\definecolor{cgris}{RGB}{225,225,225}
\newcolumntype{L}[1]{>{\raggedright\arraybackslash}p{#1}}
\newcolumntype{Y}{>{\raggedright\arraybackslash}X}
\renewcommand{\arraystretch}{1.15}
\setlength{\emergencystretch}{3em}
\setlist{itemsep=2pt,topsep=4pt}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small ''' + esc(TITULO_CORTO) + r'''}
\fancyhead[R]{\small\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\fancypagestyle{plain}{\fancyhf{}\fancyhead[R]{\small\thepage}}

\hypersetup{pdftitle={''' + esc(TITULO) + r'''},pdfauthor={''' + esc(L.AUTORA) + r'''}}

\begin{document}
''')
    a(r'''\hypersetup{pageanchor=false}
\begin{titlepage}
\centering
\vspace*{3cm}
{\LARGE\bfseries ''' + esc(TITULO) + r'''\par}
\vspace{1.5cm}
{\large ''' + esc(VERSION) + r'''\par}
\vspace{2cm}
{\large ''' + esc(L.AUTORA) + r'''\par}
\vspace{0.5cm}
{\large Observatorio Ciudadano de Envigado (OCE)\par}
\vspace{2cm}
{\large ''' + esc(PROFESORES) + r'''\par}
\vfill
{\large Fecha de corte de la información: según el indicador, entre el ''' + esc(fecha_es(corte_min)) + ' y el '
      + esc(fecha_es(corte_max)) + r' (detalle en la sección~2)\par}' + r'''
\vspace{0.3cm}
{\large Fuentes consultadas hasta el ''' + esc(fecha_es(consulta)) + r'''\par}
\end{titlepage}
\hypersetup{pageanchor=true}
''')
    a(aviso('RESULTADOS PROVISIONALES. NADA ESTÁ VALIDADO: todos los valores están en estado Pendiente y los '
            'resultados dependen de decisiones que todavía deben tomar los profesores.'))
    a(r'\tableofcontents' + '\n' + r'\clearpage')

    # ------------------------------------------------------------ 1. resumen ejecutivo
    a(r'\section{Resumen ejecutivo}')
    n_viables = len(VIABLES)
    con_num = [c for c in VIABLES if any(
        (r := d.fila(c, y, p)) and r['disponibilidad'] == 'Calculado' for y, p in
        (('2024', 'Anual'), ('2025', 'Anual'), ('2026', 'Semestre 1')))]
    sin_num = [c for c in VIABLES if c not in con_num]
    a(f'Se calcularon, en el escenario por defecto, los {n_viables} indicadores clasificados como viables o viables con '
      f'ajuste. De ellos, {len(con_num)} tienen resultado numérico en al menos un año y {len(sin_num)} '
      f'({", ".join(esc(x.split(".")[1]) for x in sin_num)}) quedan en N/D. Todo es \\textbf{{provisional}}: las variables '
      'están en estado Pendiente y cada resultado depende de los supuestos por defecto que se describen en la sección~2. '
      'La tabla muestra el puntaje (de 0 a 100) y el color del semáforo definido en la Matriz.')
    filas = []
    cols = [('2024', 'Anual'), ('2025', 'Anual'), ('2026', 'Semestre 1')]
    for c in VIABLES:
        fila = [esc(INFO[c][0])]
        for y, p in cols:
            r = d.fila(c, y, p)
            if r is None:
                fila.append('--')
                continue
            pt, sem, col = celda_resultado(r)
            if r['disponibilidad'] == 'Calculado':
                fila.append(sc(col, f'{pt} ({sem})' if col else f'{pt} ({sem})'))
            else:
                fila.append(sc(col, pt if pt != 'N/D' else 'N/D'))
        filas.append(fila)
    a(r'\begin{center}\footnotesize\begin{tabularx}{\linewidth}{L{6.3cm}YYY}\toprule')
    a(r'\textbf{Indicador} & \textbf{2024} & \textbf{2025} & \textbf{2026 (primer semestre)} \\ \midrule')
    a('\n'.join(' & '.join(f) + r' \\' for f in filas) + r' \bottomrule\end{tabularx}\end{center}')
    a(r'{\footnotesize Escenario por defecto, nivel municipal. \colorbox{cverde}{Verde}, \colorbox{camarillo}{Amarillo}, '
      r'\colorbox{crojo}{Rojo} y \colorbox{cgris}{Gris (N/D)} según los umbrales de la Matriz; el guion (--) indica que no se calcula para '
      'ese año (no hay información). El presupuesto participativo de 2026 corresponde al primer semestre (corte al 30 de '
      'junio) y se deja sin semáforo porque la Matriz solo lo prevé en periodo anual; el de 2024 es aproximado.}')
    # mensajes
    dsi24, dsi24b = d.fila('CS.DSI', 2024), d.fila('CS.DSI', 2024, esc_id='ventana_documental_cs=hasta_fecha_consulta')
    vrc24c, vrc25c = (d.fila('CS.VRC', 2024, esc_id='vrc_evento_sin_mencion=cero_siempre'),
                      d.fila('CS.VRC', 2025, esc_id='vrc_evento_sin_mencion=cero_siempre'))
    ic_a = d.fila('CS.ICRCS', 2024, esc_id='vrc_evento_sin_mencion=cero_siempre')
    ic_b = d.fila('CS.ICRCS', 2024, esc_id='control_institucional_545_546')
    ic_lo, ic_hi = (d.fila('CS.ICRCS', 2024, esc_id='envolvente_componentes_min'),
                    d.fila('CS.ICRCS', 2024, esc_id='envolvente_componentes_max'))
    e24, e25, e26 = d.fila('PP.EPP', 2024), d.fila('PP.EPP', 2025), d.fila('PP.EPP', 2026, 'Semestre 1')
    z10 = d.zona[('Zona 10', '2024', 'Anual')]
    verdes = sum(1 for c in VIABLES for y, p in cols if (r := d.fila(c, y, p)) and r['disponibilidad'] == 'Calculado'
                 and r['semaforo'] == 'Verde')
    rojos = [c for c in VIABLES for y, p in cols if (r := d.fila(c, y, p)) and r['disponibilidad'] == 'Calculado'
             and r['semaforo'] == 'Rojo']
    a(r'\paragraph{Mensajes principales}')
    a(r'\begin{enumerate}')
    a(r'\item \textbf{Resultados con número:} ' + f'{len(con_num)} de {n_viables} indicadores. '
      f'Hay {verdes} celdas en Verde; el único Rojo con número es el DSI. VRC e ICRCS no tienen resultado por defecto.')
    a(r'\item \textbf{DSI en cero por la ventana de publicación:} con el corte del 31 de marzo del año siguiente, el DSI de 2024 y '
      f'de 2025 es {fnum(dsi24["puntaje"])}; si se acepta todo lo publicado hasta la fecha de consulta, el de 2024 sube a '
      f'{fnum(dsi24b["puntaje"])} (sigue en Rojo).')
    a(r'\item \textbf{VRC e ICRCS en N/D por presencia desconocida de veedurías:} del evento del Alcalde solo se publicó un '
      'cuadro de preguntas, sin asistentes. Si la falta de mención de veedurías se cuenta como ausencia, el VRC de 2024 es '
      f'{fnum(vrc24c["puntaje"])} y el de 2025 {fnum(vrc25c["puntaje"])}; el ICRCS de 2024 sería {fnum(ic_a["puntaje"])} '
      f'(o {fnum(ic_b["puntaje"])} en el universo institucional), con un rango por componente de {fnum(ic_lo["puntaje"])} a '
      f'{fnum(ic_hi["puntaje"])}, todo en Amarillo.')
    a(r'\item \textbf{Ejecución del presupuesto participativo alta en 2024 y 2025:} ' + f'{fnum(e24["puntaje"])} (aproximado) y '
      f'{fnum(e25["puntaje"])}, en Verde. En 2026 (primer semestre) es {fnum(e26["puntaje"])}, sin semáforo. La única zona '
      f'que no está en Verde es la zona 10 en 2024 ({fnum(z10["puntaje"])}, {z10["semaforo"]}), por un cruce aproximado de datos.')
    a(r'\item \textbf{Hallazgos de datos para Planeación:} un rubro de déficit fiscal dentro del presupuesto participativo de '
      '2024, cruces débiles entre el listado de la Escuela y Hacienda, 51 y no 50 proyectos con entrega exigible en 2024, y '
      'una cifra de 55 entregas de 2025 que no se reproduce (se obtienen entre 45 y 50). Ver sección~6.')
    a(r'\end{enumerate}')

    # ------------------------------------------------------------ 2. método
    a(r'\section{Método: escenarios y supuestos por defecto}')
    a('Cada indicador se calcula con la fórmula y los umbrales de semáforo de la hoja Resultados de la Matriz de información '
      'de indicadores. Donde la Matriz deja una decisión abierta, se calcularon todas las opciones razonables '
      '(\\emph{escenarios}) y se marcó una como valor por defecto, \\textbf{sin decidir por los profesores}. Todos los supuestos '
      'por defecto son provisionales.')
    a(r'\subsection{Supuestos del escenario por defecto}')
    sup = L.tabla_md(seccion_md(d.md, '## 4.'))
    a(xltabular(r'L{3.4cm}Y', ['Tema', 'Supuesto por defecto'], [[md_cel(r[0]), md_cel(r[1])] for r in sup], tam=r'\footnotesize'))
    a(r'''\subsection{Convenciones de lectura}
\begin{itemize}
\item \textbf{N/D}: no se puede calcular con la información disponible o con la regla de la Matriz. \textbf{NE}: no evaluable (la zona 1 no tiene proyectos de presupuesto participativo). \textbf{Revisar}: el numerador supera al denominador y la Matriz pide revisión, sin color.
\item Los semáforos se asignan con los umbrales de la Matriz, sin redondear antes de clasificar. Los puntajes se muestran con dos decimales.
\item Las variables están en estado Pendiente y los resultados en estado Provisional; ningún dato está validado.
\end{itemize}''')
    a(r'\subsection{Fecha de corte de la información}')
    a('Los datos se registran al cierre de la vigencia (31 de diciembre), pero el documento fuente puede tener un corte real '
      'distinto; la tabla muestra el corte real de las fuentes de cada indicador, por año.')
    filas = []
    for c in VIABLES:
        por_anio = defaultdict(set)
        for v in d.V:
            if (v['es_defecto'] == 'True' and v['codigo_indicador'] == c and v['estado_informacion'] == 'Pendiente'
                    and v['territorio'] == 'Municipio'):
                por_anio[v['anio'] + (' S1' if v['periodo'] == 'Semestre 1' else '')].add(v['fecha_corte_real'] or v['fecha_corte'])
        if c == 'CS.ICRCS':
            texto = 'Combina los cortes de DSI, PCS, VRC y RPO'
        else:
            texto = '; '.join(f'{k}: ' + ' y '.join(fecha_es(x) for x in sorted(vs)) for k, vs in sorted(por_anio.items()))
        filas.append([esc(INFO[c][0]), esc(texto)])
    a(xltabular(r'L{6cm}Y', ['Indicador', 'Corte real de la información'], filas, tam=r'\footnotesize'))
    a(r'\subsection{Verificación}')
    a(r'\begin{itemize}')
    a(r'\item \textbf{Contra la hoja Resultados de la Matriz de los profesores:} se recalculó una copia temporal de la Matriz '
      'con los datos marcados como validados, solo para comparar sus fórmulas, y la copia se borró. De 62 resultados '
      'comparados, 48 coinciden y 14 difieren, todas esperadas: el presupuesto participativo del primer semestre de 2026 '
      '(municipal y 12 zonas), porque la Matriz solo calcula este indicador con periodo anual y devuelve N/D, y la zona 1, '
      'donde el cálculo da NE y la Matriz da N/D.')
    a(r'\item \textbf{Recálculo independiente desde los documentos fuente:} se recalcularon a mano 23 valores (15 de indicadores, '
      '5 observaciones del RPO y 3 proyectos del presupuesto participativo de 2024); los 23 coinciden. La revisión encontró '
      '5 errores de presentación o documentación, que se corrigieron sin cambiar ningún puntaje del escenario por defecto. '
      'Además, 951 de 975 filas verificables reproducen exactamente la regla de la Matriz; las otras 24 son una sensibilidad '
      'declarada de PC.I3.')
    a(r'\item \textbf{Estados y datos personales:} no hay ningún dato en estado validado ni datos personales; las citas de '
      'ciudadanos y funcionarios están anonimizadas.')
    a(r'\end{itemize}')

    # ------------------------------------------------------------ 3. resultados por sistema
    a(r'\section{Resultados por sistema}')
    a('Para cada indicador se presentan la fórmula (de la Matriz), los valores de sus variables con su fuente (documento y '
      'página, con enlace al documento publicado), el resultado por año y cómo cambia según el escenario. '
      'El escenario por defecto aparece primero.')
    vars_por_ind = defaultdict(list)
    for v in d.V:
        if v['es_defecto'] == 'True' and v['territorio'] == 'Municipio' and v['codigo_indicador'] in VIABLES:
            vars_por_ind[v['codigo_indicador']].append(v)
    narr = narrativas(d)
    for s, nombre_s in SISTEMAS:
        inds = [c for c in VIABLES if c.startswith(s + '.')]
        if not inds:
            continue
        a(r'\subsection{' + esc(nombre_s) + ' (' + s + ')}')
        for c in inds:
            nom, formula, ent, umbral = INFO[c]
            a(r'\subsubsection*{' + esc(nom) + '}')
            a(r'\textbf{Fórmula (Matriz):} ' + esc(formula) + r' \textbf{Semáforo:} ' + esc(umbral) + '.')
            vs = vars_por_ind.get(c, [])
            filas = []
            for v in vs:
                if v['estado_informacion'] == 'No disponible' and v['valor'] == '' and not v['fuente']:
                    continue
                val = v['valor']
                if v['unidad'] == 'COP' and val:
                    val = r'\$' + num(val, 0)
                elif val:
                    val = esc(fnum(val))
                else:
                    val = 'N/D'
                if v['id_variable'].endswith('.EV') and not v['valor']:
                    val = 'Desconocida'
                if v['fuente']:
                    fz = nombre_fuente(v['fuente'], v['enlace']).replace('lista definida por parámetro', 'lista definida en los supuestos')
                    pg = pagina_limpia(v['pagina_tabla'])
                    src = (enlace_texto(v['enlace'], fz) if v['enlace'].startswith('http') else esc(fz)) + ', ' + esc(pg)
                else:
                    src = 'Sin fuente (dato no disponible)'
                filas.append([esc(v['id_variable']) + ': ' + esc(re.sub(r'\(lista del parámetro \w+\)', '(lista definida en los supuestos)', v['nombre_variable'])), esc(anio_txt(v)), val, src])
            if filas:
                a(xltabular(r'L{4.4cm}L{1.9cm}L{2.5cm}Y', ['Variable', 'Año', 'Valor', 'Fuente y página'], filas, tam=r'\scriptsize'))
            elif c == 'CS.ICRCS':
                a('El índice no tiene variables propias: se construye con los puntajes de DSI, PCS, VRC y RPO del mismo año.')
            a(r'\paragraph{Resultado y sensibilidad (nivel municipal).} Entradas: ' + esc(ent) + '.')
            filas = []
            def agrega(r, etiqueta, negrita=False):
                pt, sem, col = celda_resultado(r)
                f = [esc(etiqueta), esc(anio_txt(r)), entradas_txt(r, c), esc(pt), sc(col, esc(sem))]
                if negrita:
                    f = [r'\textbf{' + x + '}' if k in (0,) else x for k, x in enumerate(f)]
                filas.append(f)
            for y, p in (('2024', 'Anual'), ('2025', 'Anual'), ('2026', 'Semestre 1')):
                r = d.fila(c, y, p)
                if r is None:
                    continue
                if c == 'PC.I5' and p == 'Anual' and y == '2024' and r['disponibilidad'] != 'Calculado':
                    pass
                agrega(r, 'Por defecto', True)
                dflt = (r['puntaje'], r['semaforo'], r['disponibilidad'], r['entrada_1'], r['entrada_2'])
                for esc_id, etiq in ESCENARIOS.get(c, []):
                    r2 = d.fila(c, y, p, esc_id)
                    if r2 is None:
                        continue
                    if (r2['puntaje'], r2['semaforo'], r2['disponibilidad'], r2['entrada_1'], r2['entrada_2']) == dflt:
                        continue
                    agrega(r2, etiq)
            a(xltabular(r'L{5.0cm}L{1.9cm}L{3.6cm}L{1.8cm}Y', ['Escenario', 'Año', 'Entradas', 'Puntaje', 'Semáforo'],
                        filas, tam=r'\scriptsize'))
            for t in narr.get(c, []):
                a(t)

    # ------------------------------------------------------------ 4. PP por zona
    a(r'\section{Presupuesto participativo por zona}')
    a('Ejecución presupuestal (EPP) en el escenario por defecto: obligaciones sobre apropiación definitiva de los proyectos '
      'de presupuesto participativo de cada zona. El municipal es la suma de obligaciones sobre la suma de apropiaciones, no '
      'un promedio de zonas. La cifra de 2024 es aproximada (cruce del listado de la Escuela de Formación con el informe de '
      'Hacienda) y la de 2026 corresponde al primer semestre, sin semáforo.')
    terr = ['Municipio'] + [f'Zona {k}' for k in range(1, 14)]
    filas = []
    for t in terr:
        fila = [r'\textbf{Municipio}' if t == 'Municipio' else esc(t)]
        for y, p in cols:
            r = (d.fila('PP.EPP', y, p) if t == 'Municipio' else d.zona.get((t, y, p)))
            if r is None:
                fila.append('--')
                continue
            pt, sem, col = celda_resultado(r)
            if r['disponibilidad'] == 'Calculado':
                fila.append(sc(col, f'{pt} ({sem})' if col else pt))
            else:
                fila.append(sc(col, pt))
        filas.append(fila)
    a(r'\begin{center}\footnotesize\begin{tabularx}{\linewidth}{L{3cm}YYY}\toprule')
    a(r'\textbf{Territorio} & \textbf{2024 (aproximado)} & \textbf{2025} & \textbf{2026, primer semestre (sin semáforo)} \\ \midrule')
    a('\n'.join(' & '.join(f) + r' \\' for f in filas) + r' \bottomrule\end{tabularx}\end{center}')
    z1 = d.zona[('Zona 1', '2025', 'Anual')]
    a(f'La zona 1 es {esc(z1["disponibilidad"])} (no evaluable) porque no tiene proyectos de presupuesto participativo; confirmarlo con '
      r'Planeación es una condición de la Guía. En el municipal, 2025 usa obligaciones de \$' + num(e25['entrada_1'], 0)
      + r' sobre apropiación de \$' + num(e25['entrada_2'], 0) + '; 2024, \\$' + num(e24['entrada_1'], 0) + r' sobre \$'
      + num(e24['entrada_2'], 0) + r'; y 2026 (primer semestre), \$' + num(e26['entrada_1'], 0) + r' sobre \$'
      + num(e26['entrada_2'], 0) + '.')

    # ------------------------------------------------------------ 5. variables de no viables
    a(r'\section{Variables construidas de los indicadores no viables}')
    a('Para estos indicadores no se calculó el puntaje ni se asignó semáforo: solo se construyeron las variables que la '
      'información secundaria permite. La tabla siguiente indica por qué no se calcula cada uno.')
    nv = L.tabla_md(seccion_md(d.md, '## 3.'))
    a(xltabular(r'L{3.6cm}YY', ['Indicador', 'Variables construidas', 'Por qué no se calcula'],
                [[md_cel(r[0]), md_cel(r[1]), md_cel(r[2])] for r in nv], tam=r'\scriptsize'))
    filas = []
    for v in d.V:
        if (v['es_defecto'] == 'True' and v['territorio'] == 'Municipio' and v['codigo_indicador'] in NO_VIABLES
                and v['valor'] != ''):
            filas.append([esc(NOMBRE_NV[v['codigo_indicador']]), esc(v['nombre_variable']), esc(anio_txt(v)),
                          esc(fnum(v['valor'])) + (' ' + esc(v['unidad'].replace('Porcentaje (0–100)', '%')) if v['unidad'] not in ('Conteo', 'Número') else '')])
    a(r'\subsection{Valores de las variables}')
    a(xltabular(r'L{4cm}YL{1.6cm}L{2.6cm}', ['Indicador', 'Variable', 'Año', 'Valor'], filas, tam=r'\scriptsize'))
    a('Estado de la información: Pendiente. Las variables de CEP y CMA para cada proyecto están en la hoja PP por proyecto de '
      'la base de datos (Anexo).')

    # ------------------------------------------------------------ 6. hallazgos
    a(r'\section{Hallazgos para Planeación}')
    hallazgos(d, a)

    # ------------------------------------------------------------ 7. decisiones
    a(r'\section{Decisiones pendientes para los profesores}')
    a('Cada fila es una decisión que cambia un resultado. Los rangos son del nivel municipal; el valor por defecto es el '
      'primero de cada rango. «Revisar» y N/D no tienen color de semáforo.')
    dec = L.tabla_md(seccion_md(d.md, '## 5.'))
    # Casos dudosos de la segunda lectura (acuerdo con duda): decisiones adicionales, sin efecto en los escenarios calculados.
    dec = dec + [
        ['Carta de respuesta como prueba de presencia (VRC; de la segunda lectura)',
         '¿Basta una carta de respuesta dirigida a una veeduría para acreditar que la veeduría estuvo presente en el evento, '
         'o se exige acta o lista de asistencia?',
         'Basta la carta / se exige acta o lista de asistencia',
         'VRC, ICRCS',
         'No cambia ningún resultado calculado: el caso (rendición del Comité Municipal de Discapacidad, segundo semestre '
         'de 2024) solo entra en el universo amplio de eventos, que da N/D'],
        ['Respuesta sin la fecha pedida (RPO; de la segunda lectura)',
         '¿Una respuesta que da información sobre el asunto pero no contesta la fecha que pidió la persona («¿para cuándo?») '
         'cuenta como respuesta de fondo?',
         'Sí cuenta / no cuenta',
         'RPO',
         'No cambia ningún resultado calculado: el caso (carta del Comité Municipal de Discapacidad, primer semestre de '
         '2025) solo entra en el universo de todos los cuadros de 2025, que da N/D'],
    ]
    a(xltabular(r'L{2.6cm}L{3.1cm}L{2.5cm}L{2.4cm}Y',
                ['Decisión', 'Pregunta', 'Opciones', 'Indicadores afectados', 'Rango de resultados'],
                [[md_cel(r[0]), md_cel(r[1]), md_cel(r[2]), md_cel(r[3]), md_cel(r[4])] for r in dec], tam=r'\scriptsize'))

    # ------------------------------------------------------------ 8. pendientes
    a(r'\section{Pendientes y próximos pasos}')
    nv_, nr_ = len(d.vrc_muestra), len(d.rpo_muestra)
    forensis = next(v for v in d.V if v['id_variable'] == 'PDM.CM.O' and v['es_defecto'] == 'True')
    a(r'\begin{enumerate}')
    ac_v = sum(1 for x in d.vrc_muestra if x.get('acuerdo_si_no', '').strip().lower() in ('si', 'sí'))
    ac_r = sum(1 for x in d.rpo_muestra if x.get('acuerdo_si_no', '').strip().lower() in ('si', 'sí'))
    assert (ac_v, nv_, ac_r, nr_) == (4, 4, 12, 12), (ac_v, nv_, ac_r, nr_)
    a(r'\item \textbf{Segunda lectura de VRC y RPO (terminada).} Una segunda persona clasificó de forma independiente una muestra '
      r'aleatoria del 20~\% de los eventos (' + str(len(d.clas_vrc)) + ') y de las observaciones (' + str(len(d.clas_rpo)) + ') '
      f'clasificados por el primer lector: {nv_} eventos del VRC y {nr_} observaciones del RPO. La lectura se hizo sobre las citas '
      'y el contexto registrados en la muestra, no releyendo cada documento completo. Acuerdo entre lectores (filas con acuerdo '
      f'sobre filas revisadas): VRC {ac_v}/{nv_} (' + fnum(100 * ac_v / nv_, 0) + r'~\%) y RPO ' + f'{ac_r}/{nr_} ('
      + fnum(100 * ac_r / nr_, 0) + r'~\%). Ningún resultado cambia. Dos casos, en los que hubo acuerdo pero con duda, '
      'se llevan a la sección~7 como decisiones para los profesores.')
    a(r'\item \textbf{Verificación de Forensis 2024.} La mortalidad vial de 2024 (' + fnum(forensis['valor']) + ' por 100.000 '
      'habitantes, ' + esc(forensis['pagina_tabla']) + ') es una cifra citada: el documento no se ha podido descargar '
      '(el sitio de Medicina Legal rechaza descargas automáticas). Hay que bajarlo, confirmar la cifra y el criterio territorial '
      '(lugar del hecho o residencia). Enlace: ' + enlace(forensis['enlace']) + '.')
    a(r'\item \textbf{Tomar las decisiones de la sección~7} y, con ellas, fijar el escenario oficial de cada indicador.')
    a(r'\item \textbf{Confirmar con Planeación} los hallazgos de la sección~6 (lista de proyectos de presupuesto participativo '
      'de cada vigencia, cruces débiles y rubro de déficit fiscal).')
    a(r'\item \textbf{Validar} los valores que los profesores acepten; hoy ninguno lo está.')
    a(r'\end{enumerate}')
    a(r'\subsection{Solicitudes a la Alcaldía relacionadas con los indicadores calculados}')
    wb = openpyxl.load_workbook(F_XLSX, read_only=True)
    sol = L.hoja(wb, 'Solicitudes')
    rel = [r for r in sol if re.match(r'(CS\.(PCS|DSI|VRC|RPO)|PC\.I[235]|PDM\.C[AM]\.|PP\.(EPP|CEP|CMA))', r['id_variable'])]
    a(f'De las {len(sol)} solicitudes de información preparadas, {len(rel)} se relacionan con los indicadores calculados o con '
      'los pendientes de esta sección.')
    a(xltabular(r'L{3.4cm}L{2.2cm}YL{3cm}', ['Dependencia', 'Variable', 'Dato a solicitar', 'Motivo'],
                [[esc(L.corto(r['dependencia'], 80)), esc(r['id_variable']), esc(L.corto(r['dato_a_solicitar'], 200)),
                  esc(L.corto(r['Justificación'] if 'Justificación' in r else '', 140))] for r in sorted(rel, key=lambda r: (r['dependencia'], r['id_variable']))],
                tam=r'\scriptsize'))

    # ------------------------------------------------------------ anexo
    a(r'\appendix')
    a(r'\section{Estructura de la base de datos para el tablero}')
    anexo(d, a)
    a(r'\end{document}')
    tex = '\n\n'.join(T) + '\n'
    for k, v in (('«', r'\guillemotleft{}'), ('»', r'\guillemotright{}'), ('º', r'\textsuperscript{o}')):
        tex = tex.replace(k, v)
    raros = sorted({ch for ch in tex if ord(ch) > 127 and ch not in L.LETRAS_OK})
    assert not raros, f'caracteres no ASCII sin traducir: {raros}'
    prohib = re.findall(r'\b(?:[Ss]cript|[Ff]ase|[Tt]anda|[Rr]epo|[Pp]arcial|[Ee]jecutor|[Aa]nalizador|data_raw|outputs/)\b|.{30}\.csv|.{30}\.py\b|.{30}\.md\b|.{20}Validado',
                        re.sub(r'\\href\{[^}]*\}', '', tex))
    assert not prohib, f'lenguaje interno en el texto: {prohib}'
    return tex


def narrativas(d):
    g = d.fila
    f = lambda r: fnum(r['puntaje'])
    N = defaultdict(list)
    N['CS.PCS'].append('Cómo cambia: no cambia con ningún supuesto probado (fecha de registro, tratamiento de 2026). '
                       'El Departamento Administrativo de Planeación dejó observaciones en la fila del plan indicativo sobre '
                       'la concordancia de la ejecución anual; el resultado coincide con lo reportado, pero conviene atenderlas antes de validar.')
    N['CS.DSI'].append(f'Cómo cambia: el cero depende de la ventana de publicación. Con el corte del 31 de marzo del año siguiente '
                       f'el DSI es {f(g("CS.DSI", 2024))} en 2024 y {f(g("CS.DSI", 2025))} en 2025; con todo lo publicado hasta la fecha '
                       f'de consulta, 2024 es {f(g("CS.DSI", 2024, esc_id="ventana_documental_cs=hasta_fecha_consulta"))}; con cifras de apoyo, 2025 es '
                       f'{f(g("CS.DSI", 2025, esc_id="dsi_fuente_valor=oficial_incluye_proxy"))}. Todos son Rojo. La lista de productos del '
                       'plan indicativo da 100, pero se usa solo como control.')
    N['CS.VRC'].append('Por qué es N/D: la Matriz indica N/D cuando algún evento tiene presencia desconocida. Para la audiencia del '
                       'Alcalde solo se publicó el cuadro de preguntas, sin acta ni asistentes. Si la falta de mención de veedurías se cuenta '
                       f'como ausencia, 2024 da {f(g("CS.VRC", 2024, esc_id="vrc_evento_sin_mencion=cero_siempre"))} (Amarillo) y 2025 '
                       f'{f(g("CS.VRC", 2025, esc_id="vrc_evento_sin_mencion=cero_siempre"))} (Rojo).')
    N['CS.RPO'].append('Cómo cambia: el defecto usa solo la audiencia del Alcalde, excluye felicitaciones y aplica un criterio estricto de '
                       f'respuesta de fondo ({f(g("CS.RPO", 2024))}, Verde). Con todas las observaciones baja a '
                       f'{f(g("CS.RPO", 2024, esc_id="rpo_filtro_pdm=todas"))} (Amarillo) y con todos los cuadros a '
                       f'{f(g("CS.RPO", 2024, esc_id="rpo_universo=todos_los_cuadros"))} (Amarillo). En 2025 el cuadro publicado es '
                       'claramente incompleto (3 intervenciones), por eso queda N/D.')
    N['CS.ICRCS'].append('Cómo cambia: por defecto es N/D porque falta el VRC (y en 2025 también el RPO). Todo ICRCS numérico de 2024 exige apartarse '
                         'de la regla de la Matriz sobre presencia desconocida. Con la falta de mención contada como ausencia, 2024 es '
                         f'{f(g("CS.ICRCS", 2024, esc_id="vrc_evento_sin_mencion=cero_siempre"))}; en el universo institucional, '
                         f'{f(g("CS.ICRCS", 2024, esc_id="control_institucional_545_546"))}; y el rango por componente va de '
                         f'{f(g("CS.ICRCS", 2024, esc_id="envolvente_componentes_min"))} a {f(g("CS.ICRCS", 2024, esc_id="envolvente_componentes_max"))}, '
                         'todo en Amarillo. Es un rango de sensibilidad, no un valor oficial: combina opciones que pueden no ser compatibles entre sí. '
                         'La revisión independiente estimó un mínimo de 43,2 (Rojo) combinando además el universo del Alcalde solo; esa combinación no está calculada.')
    N['PC.I2'].append('Cómo cambia: en 2025 la fuente dice que se hicieron 12 asambleas «cubriendo la totalidad de las zonas» sin listarlas; se '
                      'asignaron a las zonas 2 a 13. Si «totalidad» significa 13 zonas, el resultado sería 13 / 13.')
    N['PC.I3'].append('Cómo cambia: el denominador de 2025 es reconstruido (acciones culminadas más no ejecutadas) porque no hay programación fechada. '
                      'Se excluyó el diplomado en Gestión Pública por ser para servidores, algo que la fuente no acredita; si se cuentan los semilleros como '
                      'acción formativa, el resultado baja a cerca de 80,6.')
    N['PC.I5'].append('Solo hay información de dos instancias (Consejo Municipal de Seguridad Social en Salud y COPACO); 2024 y los semestres de 2025 son N/D '
                      'porque no hay programación ni fechas por sesión.')
    N['PDM.CA'].append('Cómo cambia: el valor de 5,75 corresponde al espacio público verde urbano del POT y no es la misma serie de la línea base '
                       '(3,88); la comparación es aproximada. Si no se aceptara esa diferencia (estricto con la serie base) el resultado sería N/D.')
    N['PDM.CM'].append('Cómo cambia: la cifra observada (7,25; 18 muertes) es una cita de Forensis 2024 pendiente de verificar. '
                       f'Con datos de la Policía y población DANE posterior a la pandemia el resultado es {f(g("PDM.CM", 2024, esc_id="pdm_cm_fuente=policia_datosgov;poblacion_dane=dane_postcovid_2020_2035"))} (Amarillo).')
    N['PP.EPP'].append('Cómo cambia: ver la sección~6 para el rubro de déficit fiscal de 2024 (sin él, 98,64). El municipal es suma de obligaciones sobre suma de '
                       'apropiaciones. El proyecto 0553 (costo del proceso) se excluye por defecto. En 2026 la Matriz no prevé el indicador en periodo '
                       'semestral: el puntaje de medio año se conserva como dato informativo, sin semáforo.')
    return N


def hallazgos(d, a):
    app_r = sum(int(p['app_rubro_2_3_7_06'] or 0) for p in d.C24)
    n_r = sum(1 for p in d.C24 if int(p['app_rubro_2_3_7_06'] or 0) > 0)
    e24 = d.fila('PP.EPP', 2024)
    sin = 100 * (float(e24['entrada_1']) - app_r) / (float(e24['entrada_2']) - app_r)
    cr = {(p['zona'], p['n_en_zona']): p for p in d.C24}
    p1, p2, p3 = cr[('13', '2')], cr[('12', '3')], cr[('10', '2')]
    pe = {y: next(v['valor'] for v in d.V if v['id_variable'] == 'PP.CEP.PE' and v['anio'] == y and v['es_defecto'] == 'True') for y in ('2024', '2025', '2026')}
    ce25 = [r for r in d.CE if r['anio'] == '2025']
    pec = sorted(sum(1 for r in ce25 if r[k] == '1') for k in ('pec_solo_31dic', 'pec_todas_filas_dic_100', 'pec_ultima_fila'))
    z10 = d.zona[('Zona 10', '2024', 'Anual')]
    ex24 = d.fila('PP.EPP', 2024, esc_id='pp_2024_metodo=solo_exactos')
    cnt = Counter(p['tipo_coincidencia'] for p in d.C24)
    a(r'\subsection{Presupuesto participativo}')
    a(r'\begin{itemize}')
    a(r'\item \textbf{Rubro 2.3.7.06 (déficit fiscal) en 2024.} ' + f'{n_r} proyectos incluyen \\${num(app_r, 0)} del rubro 2.3.7.06 '
      '«Disminución de pasivos -- Financiación de déficit fiscal», es decir, pagos de vigencias anteriores ejecutados al 100~\\%. '
      f'No se excluyó por defecto: la decisión es de los profesores. Sin ese rubro el EPP municipal de 2024 es {fnum(sin)} en lugar de {fnum(e24["puntaje"])}. '
      'En 2025 y 2026 no aparece este rubro en los proyectos. Los proyectos afectados quedan marcados como «posible no presupuesto participativo» en la tabla de clasificación de 2024.')
    a(r'\item \textbf{Cruces débiles de 2024.} Se cruzaron ' + f'{cnt["exacta"] + cnt["aproximada"]} de {sum(cnt.values())} proyectos del listado de la Escuela con los rubros de Hacienda '
      f'({cnt["exacta"]} exactos y {cnt["aproximada"]} aproximados; {cnt["sin_correspondencia"]} sin correspondencia). En los aproximados la apropiación de Hacienda puede incluir recursos que no son del presupuesto participativo. '
      f'Planeación debe confirmar, en particular, los cruces 13-2 $\\leftrightarrow$ 0714 (Hacienda {fnum(p1["diferencia_relativa_pct"], 1)}~\\% frente al listado) y '
      f'12-3 $\\leftrightarrow$ 0733 ({fnum(p2["diferencia_relativa_pct"], 1)}~\\%), y el 10-2 $\\leftrightarrow$ 0709 (+{fnum(p3["diferencia_relativa_pct"], 1)}~\\%), que explica la zona 10 en '
      f'{z10["semaforo"]} ({fnum(z10["puntaje"])}). Con solo cruces exactos el EPP de 2024 sería {fnum(ex24["puntaje"])} (Verde).')
    a(r'\item \textbf{51 proyectos con entrega exigible en 2024 (no 50).} El listado de la Escuela de 2024 tiene ' + f'{pe["2024"]} registros; la nota previa de '
      'disponibilidad contaba 50 porque omitía el de la zona 8, cuyo valor viene sin signo «\\$» en el PDF (registro dudoso, corroborado por Hacienda y por el listado de proyectos de inversión). '
      f'Para 2025 y 2026 las entregas exigibles son {pe["2025"]} y {pe["2026"]}.')
    a(r'\item \textbf{Las 55 entregas cumplidas de 2025 no se reproducen.} Con las distintas reglas de conteo se obtienen entre ' + f'{pec[0]} y {pec[-1]} '
      f'({pec[1]} con la regla por defecto), de {pe["2025"]} proyectos con entrega exigible. La entrega cumplida solo se conoce por autorreporte del plan de acción, sin acta de recibo ni informe final.')
    a(r'\item \textbf{Encabezado del plan de acción 2026.} Los encabezados de apropiación y obligación de la hoja dicen «01-04-2026», pero las cifras coinciden '
      'al peso con Hacienda al 30 de junio de 2026 (66 de 66 proyectos). Además, la cifra de \\$11.555.041.247 de obligaciones 2026 que traía la versión 1.1 de la nota de disponibilidad era errónea; '
      'la correcta es \\$5.702.528.148.')
    a(r'\item \textbf{Listas oficiales.} No hay un marcador de presupuesto participativo en la ejecución presupuestal; el universo de proyectos se reconstruyó por nombre y valor. '
      'Planeación debe confirmar la lista de proyectos de cada vigencia y la inexistencia de proyectos en la zona 1.')
    a(r'\end{itemize}')
    a(r'\subsection{Otras inconsistencias de los datos}')
    items = re.findall(r'^(\d+)\. (.*)$', seccion_md(d.md, '## 7.'), re.M)
    elegidos = {1, 7, 9, 10, 11, 12, 13}
    a(r'\begin{itemize}')
    for n, t in items:
        if int(n) in elegidos:
            a(r'\item ' + L.md_inline(lim(t)))
    a(r'\item \textbf{Fechas contradictorias entre documentos.} Para cuatro eventos de rendición de cuentas las fechas difieren entre documentos '
      '(por ejemplo, la rendición del presupuesto participativo de la vigencia 2024 figura el 12 de enero de 2025 en una justificación y el 12 de febrero de 2025 en el informe de gestión). '
      'No afecta el resultado por defecto, pero Planeación debe fijar la fecha oficial de cada evento.')
    a(r'\item \textbf{RPO 2024, intervención n.º 10 del cuadro.} Contiene dos asuntos con su propia respuesta; si se cuentan por separado, '
      f'el RPO de 2024 pasa de {fnum(d.fila("CS.RPO", 2024)["puntaje"])} a 94,44 (sigue en Verde).')
    a(r'\item \textbf{Fuentes y fechas de publicación.} En varios indicadores de presupuesto participativo la fuente figuraba como una ruta de archivo local en lugar de un nombre de publicación, y la fecha de '
      'publicación del documento está vacía en buena parte de las variables. Es un pendiente de presentación, no de cálculo; en este informe se muestra el nombre de la publicación y su enlace.')
    a(r'\end{itemize}')


def anexo(d, a):
    a('La información de este informe se entrega también como una base de datos en Excel para alimentar el tablero. Cada hoja '
      'es una tabla plana (una fila por registro, sin celdas combinadas), con columnas que se unen entre hojas por llaves. '
      'El libro de resultados se llama «base\\_indicadores\\_OCE» y contiene las hojas siguientes.')
    sup = d.sup['parametros']
    n_dec = len(sup)
    n_zona = sum(1 for r in d.R if r['codigo'] == 'PP.EPP' and r['territorio'].startswith('Zona'))
    filas = [
        ['Léeme', 'Descripción del libro, estado provisional, convenciones (N/D, NE, Revisar) y orden de lectura.', '--', '--'],
        ['Resultados', 'Un resultado por indicador, año, periodo, territorio y escenario: entradas, puntaje, semáforo, disponibilidad, '
         'motivo de N/D y fecha de corte.', 'codigo; anio; periodo; territorio; escenario_id (es_defecto marca el escenario por defecto)',
         num(len(d.R), 0)],
        ['Variables', 'Valores de las variables de entrada con unidad, estado, fuente, página o tabla, enlace al documento, '
         'fechas de corte y de consulta.', 'id_variable; anio; periodo; territorio; escenario_id. Se une con Resultados por codigo_indicador y con Soporte de conteos por ids_soporte',
         num(len(d.V), 0)],
        ['PP por zona', 'Ejecución presupuestal (EPP) del presupuesto participativo por zona 1-13 y municipio, por año y escenario: '
         'obligaciones, apropiación, puntaje y semáforo.', 'territorio; anio; periodo; escenario_id', num(n_zona + len([1 for r in d.R if r['codigo'] == 'PP.EPP' and r['territorio'] == 'Municipio']), 0)],
        ['PP por proyecto', 'Proyectos de presupuesto participativo con zona, código, apropiación, compromiso, obligación, tipo de '
         'correspondencia con el listado de la Escuela y si entran al cálculo.', 'anio; periodo; zona; codigo_proyecto', num(len(d.P), 0)],
        ['Soporte de conteos', 'Una fila por elemento contado (zona, evento, observación, proyecto, meta) con documento, página, '
         'enlace, clasificación y justificación.', 'id_soporte (se referencia desde Variables.ids_soporte)', num(len(d.S), 0)],
        ['Supuestos', 'Parámetros, opciones y valor por defecto de cada decisión pendiente (estado provisional).',
         'parametro; opcion', num(n_dec, 0) + ' parámetros'],
        ['Diccionario de datos', 'Descripción de cada columna de cada hoja: tipo, si es obligatoria y significado.', 'tabla; columna', num(len(d.D), 0)],
    ]
    a(xltabular(r'L{2.6cm}YL{4.6cm}L{1.6cm}', ['Hoja', 'Contenido', 'Columnas clave / llaves', 'Filas'],
                [[esc(f[0]), esc(f[1]), esc(f[2]), esc(f[3])] for f in filas], tam=r'\footnotesize'))
    a(r'\paragraph{Cómo se relacionan las hojas.} Resultados y Variables comparten el indicador (codigo / codigo\_indicador), el año, el periodo, el territorio y '
      'el escenario; las variables que sustentan un conteo apuntan a Soporte de conteos mediante ids\\_soporte; PP por zona y PP por proyecto se '
      'enlazan por año, periodo y zona; el Diccionario de datos explica todas las columnas. Todos los datos están en estado '
      'Pendiente y los resultados en estado Provisional.')


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)
    tex = construir()
    F_TEX.write_text(tex, encoding='utf-8')
    with zipfile.ZipFile(F_ZIP, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(F_TEX, F_TEX.name)
    print(f'Generado {F_TEX.relative_to(RAIZ)} y {F_ZIP.relative_to(RAIZ)}')


if __name__ == '__main__':
    main()
