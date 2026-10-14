"""Perfila las llaves de cruce de la matriz B.1 (inventario_IIC_relaciones.md) y mide su coincidencia.

Uso (desde la raiz del repo):  python -I scripts/perfil_llaves_iic.py [--lib RUTA_PYPDF]
Solo usa la biblioteca estandar; pypdf es opcional y solo se necesita para leer el PDF de contratos PPP 2022
(si no esta instalado, indicar la carpeta que lo contiene con --lib; sin el, se omiten los cruces con PPP).

Salidas (sin datos personales: solo conteos, patrones de formato y metricas agregadas) en outputs/:
  perfil_llaves_iic_perfil.csv   perfil por llave
  perfil_llaves_iic_cruces.csv   coincidencia por par de llaves
  perfil_llaves_iic.md           resumen legible con la normalizacion aplicada

Normalizacion (se aplica antes de perfilar y de cruzar):
  - Todas: NFC, quitar espacios y caracteres invisibles, mayusculas. Placeholders -> nulo:
    '', 'NO DEFINIDO', 'NO DEFINIDA', 'N/A', 'NA', 'NULL', 'NONE', 'NINGUNO', 'SIN DATO', '0'.
  - Contratos: se quitan todos los espacios (p. ej. 'ENV-04-09-0549-21' = 'ENV - 04 - 09 - 0549 - 21').
    Adicionalmente se calcula la 'base ENV' (regex ENV-dd-dd-dddd-dd), que ignora sufijos como '-1'.
  - BPIN y CODIGO_PROYECTO: solo digitos, sin ceros a la izquierda ('2020052660141.0' -> '2020052660141').
  - Filtro territorial: se conservan solo filas de Envigado (columna de municipio/ciudad cuando existe).
"""
import argparse, csv, re, sys, unicodedata, zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DATA = RAIZ / 'data_raw_IIC'
SALIDA = RAIZ / 'outputs'
csv.field_size_limit(10**8)

D04 = DATA / '04_datos_abiertos_y_secop' / 'datos_gov_co'
D02 = DATA / '02_ejecucion_y_presupuesto'
F_SECOP1 = D04 / 'datosgov_f789-7hwg_SECOP_I_Procesos_Envigado.csv'
F_SECOP2 = D04 / 'datosgov_jbjy-vk9h_SECOP_II_Contratos_Envigado.csv'
F_INTEG = D04 / 'datosgov_rpmr-utcd_SECOP_Integrado_Envigado.csv'
F_CUIPO = D02 / 'datos_gov_co' / 'datosgov_4f7r-epif_CUIPO_Ejecucion_Gastos_Envigado.csv'
F_PPP = DATA / '01_propuestas_y_acuerdos' / 'envigado_gov_co' / 'contratos-priorizacion-participativa-del-presupuesto' / 'numerocontratoproyectosppp2022.pdf'
F_SEG = [D02 / 'envigado_gov_co' / 'plan-de-accion-unificado-decreto-612-de-2018-ano-2023' / n
         for n in ('seguimientoproyectosdeinversiontercertrimestre2023.xlsx', 'seguimientoproyectosdeinversioncuartotrimestre2023.xlsx')]
F_INV24 = D02 / 'envigado_gov_co' / 'informe-de-gestion-alcaldia-de-envigado' / 'proyectosdeinversion2024.xlsx'

PLACEHOLDERS = {'', 'NO DEFINIDO', 'NO DEFINIDA', 'N/A', 'NA', 'NULL', 'NONE', 'NINGUNO', 'SIN DATO', '0'}
RE_ENV = re.compile(r'ENV-\d{2}-\d{2}-\d{4}-\d{2}')


# ---------- normalizacion ----------
def _base(v):
    v = unicodedata.normalize('NFC', v or '')
    v = ''.join(c for c in v if not unicodedata.category(c).startswith('C'))
    return v.strip().upper()

def norm_contrato(v):
    v = re.sub(r'\s+', '', _base(v))
    return None if v in PLACEHOLDERS else v

def norm_digitos(v):
    v = _base(v)
    v = re.sub(r'\.0+$', '', v)           # '123.0' de Excel
    v = re.sub(r'\D', '', v).lstrip('0')
    return v or None

def base_env(v):
    m = RE_ENV.search(v) if v else None
    return m.group(0) if m else None

def patron(v):
    """Patron de formato: digitos->9, letras->A (sin exponer el valor)."""
    return re.sub(r'[A-Z]', 'A', re.sub(r'\d', '9', v))


# ---------- lectura ----------
def leer_csv(ruta, col_llave, col_filtro=None, valor_filtro=None):
    """Devuelve (valores_crudos_filtrados, filas_totales, filas_tras_filtro)."""
    vals, total = [], 0
    with open(ruta, encoding='utf-8', newline='') as f:
        for fila in csv.DictReader(f):
            total += 1
            if col_filtro and _base(fila.get(col_filtro)) != valor_filtro.upper():
                continue
            vals.append(fila[col_llave])
    return vals, total, len(vals)

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NSR = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def codigos_xlsx(ruta):
    """Lee la columna del codigo de proyecto (primera columna) de cada hoja; solo esa columna."""
    z = zipfile.ZipFile(ruta)
    ss = [''.join(t.text or '' for t in si.iter(NS + 't')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si')] \
        if 'xl/sharedStrings.xml' in z.namelist() else []
    rm = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out = []
    for s in ET.fromstring(z.read('xl/workbook.xml')).iter(NS + 'sheet'):
        t = rm[s.get(NSR + 'id')].lstrip('/')
        t = t if t.startswith('xl/') else 'xl/' + t
        for r in ET.fromstring(z.read(t)).iter(NS + 'row'):
            for c in r.iter(NS + 'c'):
                if re.match(r'A\d+$', c.get('r')):
                    v = c.find(NS + 'v')
                    if v is None:
                        continue
                    val = ss[int(v.text)] if c.get('t') == 's' else v.text
                    if re.fullmatch(r'\s*\d+(\.0+)?\s*', val):   # descarta encabezados y textos
                        out.append(val)
    return out

def contratos_ppp(lib):
    if lib:
        sys.path.insert(0, lib)
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    texto = ' '.join(p.extract_text() or '' for p in PdfReader(str(F_PPP)).pages)
    return RE_ENV.findall(re.sub(r'\s*-\s*', '-', texto))


# ---------- metricas ----------
def perfil(nombre, fuente, columna, crudos, normalizados, filas_tot, filas_filtro, filtro):
    n = len(crudos)
    nulos = sum(1 for v in normalizados if v is None)
    ok = [v for v in normalizados if v is not None]
    c = Counter(ok)
    pats = Counter(patron(v) for v in ok)
    lon = [len(v) for v in ok]
    return {
        'llave': nombre, 'fuente': fuente, 'columna': columna, 'filtro_territorial': filtro,
        'filas_origen': filas_tot, 'filas_tras_filtro': filas_filtro,
        'nulos': nulos, 'pct_nulos': round(100 * nulos / n, 2) if n else 0,
        'valores_unicos': len(c), 'filas_duplicadas': len(ok) - len(c),
        'longitud_min_max': f'{min(lon)}-{max(lon)}' if lon else '',
        'patrones_top3': '; '.join(f'{p} ({k})' for p, k in pats.most_common(3)),
    }

def cruce(par, a_nom, a_vals, b_nom, b_vals):
    A, B = Counter(v for v in a_vals if v), Counter(v for v in b_vals if v)
    inter = set(A) & set(B)
    filas_a = sum(A.values())
    return {
        'par': par, 'llave_A': a_nom, 'llave_B': b_nom,
        'unicos_A': len(A), 'unicos_B': len(B), 'interseccion_unicos': len(inter),
        'pct_unicos_A_en_B': round(100 * len(inter) / len(A), 2) if A else None,
        'pct_unicos_B_en_A': round(100 * len(inter) / len(B), 2) if B else None,
        'filas_A_no_nulas': filas_a,
        'pct_filas_A_con_match': round(100 * sum(k for v, k in A.items() if v in B) / filas_a, 2) if filas_a else None,
    }

def escribir_csv(ruta, filas):
    with open(ruta, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(filas[0]))
        w.writeheader()
        w.writerows(filas)

def md_tabla(filas, cols):
    out = ['| ' + ' | '.join(cols) + ' |', '|' + '---|' * len(cols)]
    for f in filas:
        out.append('| ' + ' | '.join(str(f[c]) for c in cols) + ' |')
    return '\n'.join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--lib', help='carpeta con pypdf (opcional, para leer el PDF PPP 2022)')
    lib = ap.parse_args().lib

    perfiles, claves, avisos = [], {}, []

    def cargar(id_, nombre, fuente, ruta, col, norm, col_f=None, val_f=None):
        crudos, tot, filt = leer_csv(ruta, col, col_f, val_f)
        n = [norm(v) for v in crudos]
        perfiles.append(perfil(nombre, fuente, col, crudos, n, tot, filt,
                               f'{col_f}={val_f}' if col_f else 'ninguno (no hay columna de entidad)'))
        claves[id_] = n

    cargar('s1_contrato', 'numero_de_contrato', 'SECOP I Procesos', F_SECOP1, 'numero_de_contrato', norm_contrato, 'municipio_entidad', 'Envigado')
    cargar('s1_bpin', 'codigo_bpin', 'SECOP I Procesos', F_SECOP1, 'codigo_bpin', norm_digitos, 'municipio_entidad', 'Envigado')
    cargar('s2_ref', 'referencia_del_contrato', 'SECOP II Contratos', F_SECOP2, 'referencia_del_contrato', norm_contrato, 'ciudad', 'Envigado')
    cargar('s2_id', 'id_contrato', 'SECOP II Contratos', F_SECOP2, 'id_contrato', norm_contrato, 'ciudad', 'Envigado')
    cargar('in_num', 'numero_del_contrato', 'SECOP Integrado', F_INTEG, 'numero_del_contrato', norm_contrato, 'municipio_entidad', 'Envigado')
    cargar('cu_bpin', 'bpin', 'CUIPO Ejecucion Gastos', F_CUIPO, 'bpin', norm_digitos, 'nombre_entidad', 'Envigado')

    ppp = contratos_ppp(lib)
    if ppp is None:
        avisos.append('pypdf no disponible: se omitieron los cruces con numerocontratoproyectosppp2022.pdf (use --lib).')
    else:
        n = [norm_contrato(v) for v in ppp]
        perfiles.append(perfil('numero contrato PPP 2022 (regex ENV-dd-dd-dddd-dd)', 'PPP 2022 (PDF)', 'Numero contrato', ppp, n,
                               len(ppp), len(ppp), 'ninguno (documento municipal de Envigado)'))
        claves['ppp'] = n

    seg = []
    for f in F_SEG:
        seg += [norm_digitos(v) for v in codigos_xlsx(f)]
    perfiles.append(perfil('CODIGO_PROYECTO seguimiento 2023 (T3+T4)', 'seguimientoproyectosdeinversion*2023.xlsx', 'CODIGO_PROYECTO',
                           seg, seg, len(seg), len(seg), 'ninguno (documento municipal de Envigado)'))
    inv = [norm_digitos(v) for v in codigos_xlsx(F_INV24)]
    perfiles.append(perfil('CODIGO_PROYECTO proyectosdeinversion2024', 'proyectosdeinversion2024.xlsx', 'CODIGO_PROYECTO',
                           inv, inv, len(inv), len(inv), 'ninguno (documento municipal de Envigado)'))

    # sufijo local del BPIN: AAAA + 05266 (DIVIPOLA Envigado) + codigo local de 4 digitos
    bp_cuipo = [v for v in claves['cu_bpin'] if v]
    sufijos = lambda L: [(v[-4:].lstrip('0') or None) if v and len(v) == 13 and v[4:9] == '05266' else None for v in L]
    pat13 = Counter(1 for v in bp_cuipo if len(v) == 13 and v[4:9] == '05266')
    avisos.append(f'BPIN CUIPO con estructura AAAA05266NNNN (13 digitos): {sum(pat13.values())} de {len(bp_cuipo)} filas no nulas.')

    cruces = []
    def par(*a):
        cruces.append(cruce(*a))
    if 'ppp' in claves:
        ppp_b = [base_env(v) for v in claves['ppp']]
        par('PPP 2022 <-> SECOP I numero_de_contrato (exacto)', 'PPP numero contrato', claves['ppp'], 'numero_de_contrato', claves['s1_contrato'])
        par('PPP 2022 <-> SECOP II referencia_del_contrato (exacto)', 'PPP numero contrato', claves['ppp'], 'referencia_del_contrato', claves['s2_ref'])
        par('PPP 2022 <-> SECOP II referencia_del_contrato (base ENV, sin sufijo)', 'PPP base ENV', ppp_b, 'referencia_del_contrato base ENV', [base_env(v) for v in claves['s2_ref']])
        par('PPP 2022 <-> SECOP II id_contrato (exacto)', 'PPP numero contrato', claves['ppp'], 'id_contrato', claves['s2_id'])
        par('PPP 2022 <-> SECOP Integrado numero_del_contrato (exacto)', 'PPP numero contrato', claves['ppp'], 'numero_del_contrato', claves['in_num'])
        par('PPP 2022 <-> SECOP Integrado numero_del_contrato (base ENV)', 'PPP base ENV', ppp_b, 'numero_del_contrato base ENV', [base_env(v) for v in claves['in_num']])
    par('SECOP I codigo_bpin <-> CUIPO bpin (exacto)', 'codigo_bpin', claves['s1_bpin'], 'bpin', claves['cu_bpin'])
    par('SECOP I numero_de_contrato <-> SECOP II referencia_del_contrato (exacto)', 'numero_de_contrato', claves['s1_contrato'], 'referencia_del_contrato', claves['s2_ref'])
    par('SECOP I numero_de_contrato <-> SECOP Integrado numero_del_contrato (exacto)', 'numero_de_contrato', claves['s1_contrato'], 'numero_del_contrato', claves['in_num'])
    par('SECOP II referencia_del_contrato <-> SECOP Integrado numero_del_contrato (exacto)', 'referencia_del_contrato', claves['s2_ref'], 'numero_del_contrato', claves['in_num'])
    par('SECOP II id_contrato <-> SECOP Integrado numero_del_contrato (exacto)', 'id_contrato', claves['s2_id'], 'numero_del_contrato', claves['in_num'])
    par('CODIGO_PROYECTO seguimiento 2023 <-> BPIN CUIPO (exacto)', 'CODIGO_PROYECTO', seg, 'bpin', claves['cu_bpin'])
    par('CODIGO_PROYECTO seguimiento 2023 <-> sufijo local de BPIN CUIPO (exploratorio)', 'CODIGO_PROYECTO', seg, 'bpin[-4:] sin ceros', sufijos(claves['cu_bpin']))
    par('CODIGO_PROYECTO 2024 <-> sufijo local de BPIN CUIPO (exploratorio)', 'CODIGO_PROYECTO 2024', inv, 'bpin[-4:] sin ceros', sufijos(claves['cu_bpin']))
    par('SECOP I codigo_bpin <-> sufijo local de BPIN CUIPO (control)', 'codigo_bpin[-4:]', sufijos(claves['s1_bpin']), 'bpin[-4:] sin ceros', sufijos(claves['cu_bpin']))
    par('CODIGO_PROYECTO seguimiento 2023 <-> CODIGO_PROYECTO 2024', 'CODIGO_PROYECTO 2023', seg, 'CODIGO_PROYECTO 2024', inv)

    SALIDA.mkdir(exist_ok=True)
    escribir_csv(SALIDA / 'perfil_llaves_iic_perfil.csv', perfiles)
    escribir_csv(SALIDA / 'perfil_llaves_iic_cruces.csv', cruces)
    doc = ['# Perfil de llaves de cruce del IIC', '',
           'Generado por `scripts/perfil_llaves_iic.py`. Solo contiene conteos y patrones de formato (9 = dígito, A = letra); sin datos personales.', '',
           '## Normalización aplicada', '']
    doc += [l.strip() for l in __doc__.split('Normalizacion (se aplica antes de perfilar y de cruzar):')[1].split('"""')[0].strip().splitlines()]
    doc += ['', '## Perfil por llave', '',
            md_tabla(perfiles, ['llave', 'fuente', 'filtro_territorial', 'filas_origen', 'filas_tras_filtro', 'nulos', 'pct_nulos', 'valores_unicos', 'filas_duplicadas', 'longitud_min_max', 'patrones_top3']),
            '', '## Coincidencia entre llaves', '',
            md_tabla(cruces, ['par', 'unicos_A', 'unicos_B', 'interseccion_unicos', 'pct_unicos_A_en_B', 'pct_unicos_B_en_A', 'pct_filas_A_con_match']),
            '', '## Avisos', ''] + [f'- {a}' for a in avisos]
    (SALIDA / 'perfil_llaves_iic.md').write_text('\n'.join(doc) + '\n', encoding='utf-8')
    for f in perfiles:
        print(f"{f['llave'][:45]:45} filas={f['filas_tras_filtro']:>6} nulos={f['nulos']:>6} ({f['pct_nulos']}%) unicos={f['valores_unicos']}")
    for c in cruces:
        print(f"{c['par'][:80]:80} A={c['unicos_A']} B={c['unicos_B']} inter={c['interseccion_unicos']} A%={c['pct_unicos_A_en_B']} B%={c['pct_unicos_B_en_A']} filasA%={c['pct_filas_A_con_match']}")
    for a in avisos:
        print('AVISO:', a)
    print('Escrito en', SALIDA)


if __name__ == '__main__':
    main()
