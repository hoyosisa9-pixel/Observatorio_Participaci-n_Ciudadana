"""Limpia el `bpin` de CUIPO y valida por nombre el puente CODIGO_PROYECTO <-> BPIN.

Uso (desde la raiz del repo):  python -I scripts/limpieza_bpin_cuipo.py
Solo biblioteca estandar. Reutiliza rutas y funciones de scripts/perfil_llaves_iic.py.

Salidas en outputs/ (conteos y codigos de proyecto; sin datos personales, no se lee la columna de responsables):
  limpieza_bpin_cuipo_categorias.csv   categorias de bpin por periodo (anio)
  limpieza_bpin_cuipo_valores.csv      valores malformados distintos con su tratamiento
  puente_codigo_proyecto_bpin.csv      validacion por nombre por cada CODIGO_PROYECTO
  limpieza_bpin_cuipo.md               resumen

Clasificacion del bpin de CUIPO:
  valido        13 digitos con patron AAAA 05266 NNNN
  nulo          '' o '0'
  recuperado    malformado con correccion INEQUIVOCA: la correccion propuesta debe producir un BPIN valido
                que ya exista en CUIPO. Reglas: (a) quitar caracteres no numericos finales/iniciales y '.0';
                (b) 14 digitos con '0' sobrante tras 05266 -> quitar un '0'; (c) 12 digitos sin el '0' tras 5266 -> insertarlo.
  codigo_sectorial  forma NN.NN.NN (codigo de sector, no es BPIN); se marca, no se cruza
  malformado    resto (p. ej. 15 digitos con otro esquema 2025000000...); se marca y se excluye del cruce

Validacion del puente por nombre: para cada CODIGO_PROYECTO se toman como candidatos los BPIN validos cuyo
sufijo local (ultimos 4 digitos sin ceros a la izquierda) coincide. Se compara NOMBRE_PROYECTO con los nombres
programaticos MGA (`nom_programatico_mga`) de ese BPIN en Ejecucion y Programacion de Gastos.
Normalizacion: minusculas, sin tildes, solo letras/numeros, sin stopwords (incluye palabras genericas como
'servicio', 'gestion', 'fortalecimiento'), plural simple (-s/-es) removido.
Puntaje = max( contencion del mejor nombre MGA en el nombre del proyecto,
               SequenceMatcher sobre los tokens ordenados ); si hay < 2 tokens comunes se topa por debajo del umbral de dudoso.
Contencion = |tokens(MGA) y tokens(proyecto)| / |tokens(MGA)|.
Umbrales: confirmado >= 0.50 ; dudoso >= 0.30 y < 0.50 ; rechazado < 0.30.
"""
import csv, difflib, re, sys, unicodedata
from collections import Counter, defaultdict
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / 'scripts'))
import perfil_llaves_iic as P  # noqa: E402

SALIDA = RAIZ / 'outputs'
F_PROG = P.D02 / 'datos_gov_co' / 'datosgov_d9mu-h6ar_CUIPO_Programacion_Gastos_Envigado.csv'
UMBRAL_CONF, UMBRAL_DUD = 0.50, 0.30
RE_VALIDO = re.compile(r'\d{4}05266\d{4}')
STOP = set('de la el los las y en del al a para por con un una o u e se su sus que lo como mas muy sobre entre desde hacia municipio envigado servicio servicios prestacion apoyo gestion calidad fortalecimiento desarrollo accion acciones programa programas proyecto proyectos implementacion mejoramiento integral'.split())


def clasificar(v, validos):
    """Devuelve (categoria, bpin_limpio o None)."""
    raw = (v or '').strip()
    if raw in ('', '0'):
        return 'nulo', None
    if RE_VALIDO.fullmatch(raw):
        return 'valido', raw
    if re.fullmatch(r'\d{2}\.\d{2}\.\d{2}', raw):
        return 'codigo_sectorial', None
    cand = []
    limpio = re.sub(r'\.0+$', '', raw)
    solo = re.sub(r'\D', '', limpio) if re.fullmatch(r'\D*\d+\D*', limpio) else None
    if solo:
        cand.append(solo)
        if len(solo) == 14 and solo[4:9] == '05266' and solo[9] == '0':
            cand.append(solo[:9] + solo[10:])
        if len(solo) == 12 and solo[4:8] == '5266':
            cand.append(solo[:4] + '0' + solo[4:])
    ok = {c for c in cand if RE_VALIDO.fullmatch(c) and c in validos}
    if len(ok) == 1:
        return 'recuperado', ok.pop()
    return 'malformado', None


def norm_nombre(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().lower()
    toks = re.findall(r'[a-z0-9]+', s)
    out = []
    for t in toks:
        if t in STOP or len(t) < 3:
            continue
        if len(t) > 4 and t.endswith('es'):
            t = t[:-2]
        elif len(t) > 3 and t.endswith('s'):
            t = t[:-1]
        out.append(t)
    return out


def puntaje(proyecto, mga):
    tp, tm = set(norm_nombre(proyecto)), set(norm_nombre(mga))
    comun = len(tp & tm)
    cont = comun / len(tm) if tm else 0.0
    seq = difflib.SequenceMatcher(None, ' '.join(sorted(tp)), ' '.join(sorted(tm))).ratio() if tm else 0.0
    # exige >= 2 tokens informativos en comun para que el puntaje pueda superar el umbral de dudoso
    return (max(cont, seq) if comun >= 2 else min(max(cont, seq), UMBRAL_DUD - 0.01)), comun


NS, NSR = P.NS, P.NSR

def codigos_nombres(ruta):
    """(codigo, nombre_proyecto) de las columnas A y B de cada hoja; solo filas con codigo numerico."""
    import zipfile
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(ruta)
    ss = [''.join(t.text or '' for t in si.iter(NS + 't')) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si')]
    rm = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out = set()
    for s in ET.fromstring(z.read('xl/workbook.xml')).iter(NS + 'sheet'):
        t = rm[s.get(NSR + 'id')].lstrip('/')
        t = t if t.startswith('xl/') else 'xl/' + t
        for r in ET.fromstring(z.read(t)).iter(NS + 'row'):
            cel = {}
            for c in r.iter(NS + 'c'):
                col = re.match(r'[A-Z]+', c.get('r')).group()
                if col in ('A', 'B'):
                    v = c.find(NS + 'v')
                    if v is not None:
                        cel[col] = ss[int(v.text)] if c.get('t') == 's' else v.text
            cod = P.norm_digitos(cel.get('A', ''))
            if cod and re.fullmatch(r'\s*\d+(\.0+)?\s*', cel.get('A', '')):
                out.add((cod, re.sub(r'\s+', ' ', cel.get('B', '')).strip()))
    return out


def main():
    # ---- 1. limpieza ----
    filas = []
    for ruta, nombre in ((P.F_CUIPO, 'ejecucion'), (F_PROG, 'programacion')):
        with open(ruta, encoding='utf-8', newline='') as f:
            for r in csv.DictReader(f):
                filas.append((nombre, r['periodo'], r['bpin'], r['nom_programatico_mga']))
    validos = {v.strip() for n, _, v, _ in filas if n == 'ejecucion' and RE_VALIDO.fullmatch(v.strip())}
    cat_periodo, mal_vals = Counter(), Counter()
    bpin_limpio = []   # solo ejecucion (la llave de cruce con SECOP I)
    nombres = defaultdict(set)
    for n, per, v, nom in filas:
        cat, b = clasificar(v, validos)
        cat_periodo[(n, per[:4], cat)] += 1
        if cat not in ('valido', 'nulo'):
            mal_vals[(n, v, cat, b or '')] += 1
        if n == 'ejecucion':
            bpin_limpio.append(b)
        if b and nom:
            nombres[b].add(nom)
    P.escribir_csv(SALIDA / 'limpieza_bpin_cuipo_categorias.csv',
                   [{'dataset': k[0], 'anio_periodo': k[1], 'categoria': k[2], 'filas': c} for k, c in sorted(cat_periodo.items())])
    P.escribir_csv(SALIDA / 'limpieza_bpin_cuipo_valores.csv',
                   [{'dataset': k[0], 'valor_original': k[1], 'categoria': k[2], 'bpin_recuperado': k[3], 'filas': c}
                    for k, c in sorted(mal_vals.items(), key=lambda x: (x[0][2], -x[1]))])
    tot = Counter()
    for (n, _, cat), c in cat_periodo.items():
        if n == 'ejecucion':
            tot[cat] += c

    # SECOP I vs CUIPO limpio
    s1, _, _ = P.leer_csv(P.F_SECOP1, 'codigo_bpin', 'municipio_entidad', 'Envigado')
    s1 = [P.norm_digitos(v) for v in s1]
    cr_antes = P.cruce('SECOP I codigo_bpin <-> CUIPO bpin (crudo normalizado)', 'codigo_bpin', s1, 'bpin', [P.norm_digitos(f[2]) for f in filas if f[0] == 'ejecucion'])
    cr_desp = P.cruce('SECOP I codigo_bpin <-> CUIPO bpin (limpio)', 'codigo_bpin', s1, 'bpin_limpio', bpin_limpio)
    anios_s1 = Counter(v[:4] for v in set(x for x in s1 if x))
    anios_cu = Counter(v[:4] for v in set(x for x in bpin_limpio if x))
    s1_set = {x for x in s1 if x}
    enlazables = {x for x in s1_set if x[:4] >= '2021'}
    cr_2021 = (len(enlazables & set(x for x in bpin_limpio if x)), len(enlazables))

    # ---- 2. puente por nombre ----
    por_sufijo = defaultdict(set)
    for b in nombres:
        por_sufijo[b[-4:].lstrip('0')].add(b)
    pares = []
    for fuente, rutas in (('seguimiento2023', P.F_SEG), ('proyectosinversion2024', [P.F_INV24])):
        cods = defaultdict(set)
        for ruta in rutas:
            for cod, nom in codigos_nombres(ruta):
                cods[cod].add(nom)
        for cod, noms in sorted(cods.items(), key=lambda x: int(x[0])):
            cands = por_sufijo.get(cod, set())
            if not cands:
                continue
            nom_proy = sorted(noms, key=len, reverse=True)[0]
            ev = []
            for b in sorted(cands):
                mejor = max(((puntaje(nom_proy, m)[0], m) for m in nombres[b]), default=(0.0, ''))
                ev.append((mejor[0], b, mejor[1]))
            ev.sort(reverse=True)
            sc, b, m = ev[0]
            n_conf = sum(1 for e in ev if e[0] >= UMBRAL_CONF)
            clase = 'confirmado' if sc >= UMBRAL_CONF else 'dudoso' if sc >= UMBRAL_DUD else 'rechazado'
            pares.append({'fuente': fuente, 'codigo_proyecto': cod, 'sufijo_candidatos_bpin': len(cands),
                          'bpin_mejor': b, 'anio_bpin': b[:4], 'puntaje': round(sc, 3), 'clasificacion': clase,
                          'bpin_confirmados_total': n_conf, 'nombre_proyecto': nom_proy[:90], 'nombre_mga_mejor': m[:90]})
    P.escribir_csv(SALIDA / 'puente_codigo_proyecto_bpin.csv', pares)
    resumen = Counter((p['fuente'], p['clasificacion']) for p in pares)
    ambiguos = Counter(p['fuente'] for p in pares if p['sufijo_candidatos_bpin'] > 1)
    multi_conf = Counter(p['fuente'] for p in pares if p['bpin_confirmados_total'] > 1)

    # ---- resumen ----
    L = ['# Limpieza de `bpin` CUIPO y validación del puente CODIGO_PROYECTO ↔ BPIN', '',
         'Generado por `scripts/limpieza_bpin_cuipo.py`. Conteos y códigos de proyecto; sin datos personales.', '',
         '## Clasificación de `bpin` (CUIPO Ejecución de Gastos, filas)', '',
         '| Categoría | Filas |', '|---|---|'] + [f'| {k} | {v} |' for k, v in tot.most_common()]
    L += ['', 'Por periodo (año) y categoría: `limpieza_bpin_cuipo_categorias.csv`; valores distintos malformados: `limpieza_bpin_cuipo_valores.csv`.', '',
          '## Cobertura temporal de CUIPO', '',
          f"Años de `periodo` en Ejecución de Gastos: {sorted({k[1] for k in cat_periodo if k[0]=='ejecucion'})}. No hay periodos 2018-2022.", '',
          f"Años de formulación (prefijo) de BPIN únicos: SECOP I {dict(sorted(anios_s1.items()))}; CUIPO limpio {dict(sorted(anios_cu.items()))}.", '',
          '## SECOP I `codigo_bpin` ↔ CUIPO `bpin`', '',
          P.md_tabla([cr_antes, cr_desp], ['par', 'unicos_A', 'unicos_B', 'interseccion_unicos', 'pct_unicos_A_en_B', 'pct_filas_A_con_match']), '',
          f'BPIN de SECOP I con prefijo >= 2021 (potencialmente enlazables): {cr_2021[1]}; presentes en CUIPO limpio: {cr_2021[0]}.', '',
          f'## Validación por nombre (umbrales: confirmado >= {UMBRAL_CONF}, dudoso >= {UMBRAL_DUD}, rechazado < {UMBRAL_DUD})', '',
          '| Fuente | Códigos con candidato | Confirmado | Dudoso | Rechazado | Sufijo con varios BPIN | Con varios BPIN confirmados |', '|---|---|---|---|---|---|---|']
    for fu in ('seguimiento2023', 'proyectosinversion2024'):
        c, d, r_ = (resumen[(fu, k)] for k in ('confirmado', 'dudoso', 'rechazado'))
        L.append(f'| {fu} | {c+d+r_} | {c} | {d} | {r_} | {ambiguos[fu]} | {multi_conf[fu]} |')
    L += ['', '### Ejemplos', '']
    for clase in ('confirmado', 'dudoso', 'rechazado'):
        L.append(f'**{clase}**')
        for p in [p for p in pares if p['clasificacion'] == clase][:4]:
            L.append(f"- cód. {p['codigo_proyecto']} / BPIN {p['bpin_mejor']} (puntaje {p['puntaje']}): «{p['nombre_proyecto']}» vs «{p['nombre_mga_mejor']}»")
        L.append('')
    (SALIDA / 'limpieza_bpin_cuipo.md').write_text('\n'.join(L) + '\n', encoding='utf-8')
    print('\n'.join(L))


if __name__ == '__main__':
    main()
