"""Excel del OCE como base de datos con fórmulas (Provisional).

Cinco hojas: Léame, Índices (tabla Indices), Datos (tabla Datos), Indicadores (tabla Indicadores) y Detalle (tabla Detalle).
Cadena de cálculo: Detalle -> Datos -> Indices -> Indicadores. Las fórmulas son clásicas (IF, OR, MIN, MAX, COUNTIFS, SUMIFS, AVERAGEIFS),
sin ROUND, y solo usan celdas de su misma fila o columnas completas con criterio. Para Power BI, después de guardar se escribe en el XML el
valor calculado de cada fórmula (espejo en Python con precisión completa) y se activa fullCalcOnLoad. Todo sale de los CSV de construcción.
"""
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

import comun_entregas as ce

ARIAL = 'Arial'
F_CUERPO = Font(name=ARIAL, size=10)
F_ENC = Font(name=ARIAL, size=10, bold=True, color='FFFFFF')
F_NEG = Font(name=ARIAL, size=10, bold=True)
F_TIT = Font(name=ARIAL, size=14, bold=True, color='1F4E78')
F_SUB = Font(name=ARIAL, size=11, bold=True, color='1F4E78')
ENC_FILL = PatternFill('solid', start_color='1F4E78')
SUB_FILL = PatternFill('solid', start_color='D9E2F3')
ANIO_CORTE = 2025
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Constantes de las fórmulas (una sola fuente): se comprueban contra el texto de los CSV con ce.req
K = {'LB_CA': 4.54, 'REF_CA': 9, 'LB_I1': 3.68, 'REF_I1': 7, 'VEED': 5.4, 'CM': 7.3301, 'HOM': 2.7, 'LA': 19, 'LN': 100.44}
PESOS = {'PP.IREPP': [('PP.EPP', 0.2), ('PP.EJ90', 0.4), ('PP.FID', 0.4)],
         'PC.ISPC': [('PC.I1', 0.3), ('PC.I2', 0.25), ('PC.I3', 0.2), ('PC.I4', 0.25)]}
PROMEDIO = {'PDM.ICPD7': 'PDM', 'CS.ICRCS': 'CS'}
COMPONENTES = {'PDM.ICPD7': 7, 'PC.ISPC': 4, 'PP.IREPP': 3, 'CS.ICRCS': 4}
METODO = {  # código -> (método, descripción de la fórmula del puntaje)
    'PDM.SALUD': 'tope', 'PDM.ILE': 'promedio', 'PDM.SP': 'tope', 'PDM.CM': 'inverso', 'PDM.CA': 'avance', 'PDM.VEED': 'veedurias', 'PDM.HOM': 'inverso',
    'PC.I1': 'avance', 'PC.I2': 'proporcion', 'PC.I3': 'proporcion', 'PC.I4': 'proporcion',
    'PP.EPP': 'proporcion', 'PP.EJ90': 'proporcion', 'PP.FID': 'proporcion',
    'CS.PQRSDF': 'tope', 'CS.DSI': 'proporcion', 'CS.VRC': 'proporcion', 'CS.RPO': 'proporcion'}
FORMULA_TXT = {
    'proporcion': 'puntaje = valor (el valor ya es un porcentaje de cumplimiento)',
    'tope': 'puntaje = MIN(100, valor)',
    'promedio': 'puntaje = valor = promedio de los puntajes LA y LN (hoja Detalle); exige los dos',
    'inverso': 'puntaje = si valor <= 0, 100; si no, MIN(100, 100 * referente / valor)',
    'avance': 'puntaje = MAX(0, MIN(100, 100 * (valor - linea_base) / (referente - linea_base)))',
    'veedurias': 'puntaje = MIN(100, 100 * valor / linea_base)'}
SEMAFORO_F = 'IF({c}="","",IF({c}>=90,"Verde",IF({c}>=75,"Amarillo",IF({c}>=50,"Naranja","Rojo"))))'


def semaforo(v):
    return '' if v == '' else ('Verde' if v >= 90 else 'Amarillo' if v >= 75 else 'Naranja' if v >= 50 else 'Rojo')


def num(t):
    t = (t or '').strip()
    return float(t) if re.fullmatch(r'-?\d+(\.\d+)?', t) else None


def primera_url(enlace):
    u = re.findall(r'https?://[^\s;]+', enlace or '')
    return (u[0] if u else None), (' ; '.join(u[1:]) or None)


class Hoja:
    """Tabla de Excel con fórmulas y su valor espejo."""
    def __init__(self, wb, nombre, tabla, columnas):
        self.ws, self.tabla, self.cols = wb.create_sheet(nombre), tabla, columnas
        self.ws.append(columnas)
        self.cache = {}   # 'L5' -> valor
        self.n = 1

    def fila(self, vals):
        """vals: lista de valores; una tupla (formula, valor_espejo) es una fórmula."""
        self.n += 1
        for j, v in enumerate(vals, 1):
            c = self.ws.cell(self.n, j)
            if isinstance(v, tuple):
                c.value = '=' + v[0].replace('{r}', str(self.n))
                self.cache[c.coordinate] = v[1]
            else:
                c.value = v
        return self.n

    def cierra(self, formatos=None, anchos=None, envolver=()):
        ws = self.ws
        for c in ws[1]:
            c.fill, c.font, c.alignment = ENC_FILL, F_ENC, Alignment(vertical='center', wrap_text=True)
        for fila in ws.iter_rows(min_row=2):
            for c in fila:
                c.font = F_CUERPO
        for j, col in enumerate(self.cols, 1):
            L = get_column_letter(j)
            ws.column_dimensions[L].width = (anchos or {}).get(col, 14)
            if col in (formatos or {}):
                for c in ws[L][1:]:
                    c.number_format = formatos[col]
            if col in envolver:
                for c in ws[L][1:]:
                    c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.freeze_panes = 'D2' if len(self.cols) > 6 else 'B2'
        t = Table(displayName=self.tabla, ref=f'A1:{get_column_letter(len(self.cols))}{self.n}')
        t.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=True)   # sin color de semáforo
        ws.add_table(t)
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = '1:1'


def construir(hechos, brechas, fichas, dim_ind, codigo_base, salida, fecha, leame):
    """Arma el libro y lo guarda en `salida`. Devuelve un dict con tamaños y controles."""
    cat = {d[0]: d for d in dim_ind}
    unidad = {c: d[13] for c, d in cat.items()}
    fx = {f['codigo']: f for f in fichas}
    # constantes de las fórmulas = textos de los CSV
    ce.req('PDM.CA', '4,54', '9 m2/hab'); ce.req('PC.I1', 'LB 3,68', 'referente 7 %'); ce.req('PDM.VEED', 'LB = 5,4')
    ce.req('PDM.CM', '7.3301'); ce.req('PDM.HOM', '2,7 por 100.000', '3,7'); ce.req('PDM.ILE', '19 %', '100,44 %')

    por = {}
    for r in hechos:
        por[(r['codigo'], r['anio'])] = r
    orden_dim = {'PDM': 1, 'PC': 2, 'PP': 3, 'CS': 4}

    wb = Workbook()
    wl = wb.active
    wl.title = 'Léame'
    H_IND = Hoja(wb, 'Índices', 'Indices', ['id', 'dimension', 'codigo', 'indice', 'anio', 'componentes', 'componentes_con_dato', 'puntaje',
                                            'puntaje_max', 'semaforo_propuesta', 'tiene_dato', 'motivo_sin_dato', 'formula'])
    H_DAT = Hoja(wb, 'Datos', 'Datos', ['id', 'dimension', 'codigo', 'indicador', 'anio', 'numerador', 'sin_verificar', 'denominador', 'valor',
                                        'valor_max', 'brecha_rango', 'puntaje', 'semaforo_propuesta', 'tiene_dato', 'motivo_sin_dato', 'fuente',
                                        'enlace', 'otros_enlaces', 'fecha_consulta', 'nota'])
    H_CAT = Hoja(wb, 'Indicadores', 'Indicadores', ['codigo', 'orden', 'dimension', 'nombre_dimension', 'tipo', 'indice', 'indicador', 'definicion',
                                                    'unidad', 'sentido', 'peso', 'metodo_puntaje', 'formula_puntaje', 'linea_base', 'referente',
                                                    'referente_fuente', 'brecha_unidad', 'brecha_sin_dato', 'brecha_total', 'brecha_pct',
                                                    'brecha_texto', 'fuente_principal', 'enlace_principal'])
    H_DET = Hoja(wb, 'Detalle', 'Detalle', ['id', 'dimension', 'codigo_padre', 'codigo', 'tipo', 'descripcion', 'anio', 'subperiodo', 'numerador',
                                            'denominador', 'valor', 'puntaje', 'clasificacion', 'unidad', 'fuente', 'enlace', 'nota'])

    # ------------------------------------------------------------------ DETALLE (entradas)
    EXCL = {'PDM.CM.TD', 'PDM.HOM.TD', 'PDM.POB'}
    det_rows = [r for r in hechos if r['tipo'] in ('subcomponente', 'respaldo', 'detalle', 'contexto') and not r['codigo'].endswith('.MAX')
                and r['codigo'] not in EXCL | {'PDM.POB05', 'PDM.ILE.A'} and r['anio_num'] <= ANIO_CORTE]
    det_rows.sort(key=lambda r: (orden_dim[r['dimension']], codigo_base(r['codigo']), r['codigo'], r['anio_num'], r['subperiodo'] or ''))
    det_val = {}     # (padre, codigo, año) -> valor | puntaje espejo
    det_pt = {}
    vistos = {}
    for r in det_rows:
        cod = r['codigo']
        padre = codigo_base(cod)
        n = vistos.get((cod, r['anio']), 0) + 1
        vistos[(cod, r['anio'])] = n
        url, _otros = primera_url(r['enlace'])
        t = r['valor']
        v = num(t)
        clasif = t if (v is None and t) else None
        numer, deno = num(r['numerador']), num(r['denominador'])
        fila_extra = None
        id_ = f"{cod}|{r['anio']}" + (f'|{n}' if n > 1 or r['tipo'] == 'detalle' else '')
        base = [id_, r['dimension'], padre, cod, r['tipo'], r['indicador'], r['anio_num'], r['subperiodo'] or None]
        cola = [unidad.get(cod), r['fuente'] or None, url, r['nota'] or None]
        if cod == 'PDM.ILE.LA':
            a = por[('PDM.ILE.A', r['anio'])]
            at, pob = num(a['numerador']), num(por[('PDM.POB05', r['anio'])]['valor'])
            val = 100 * at / pob if at is not None and pob else ''
            pt = '' if val == '' else min(100, 100 * val / K['LA'])
            rr = H_DET.n + 1
            H_DET.fila(base + [at, pob, (f'IF(OR(I{{r}}="",J{{r}}=""),"",100*I{{r}}/J{{r}})', val),
                               (f'IF(K{{r}}="","",MIN(100,100*K{{r}}/{K["LA"]}))', pt), None] + cola)
            det_val[(padre, cod, r['anio'])], det_pt[(padre, cod, r['anio'])] = val, pt
        elif cod == 'PDM.ILE.LN':
            N = num(r['numerador'])
            pt = '' if N is None else min(100, 100 * N / K['LN'])
            H_DET.fila(base + [None, None, N, (f'IF(K{{r}}="","",MIN(100,100*K{{r}}/{K["LN"]}))', pt), None if N is not None else (clasif or 'N/D')] + cola)
            det_val[(padre, cod, r['anio'])], det_pt[(padre, cod, r['anio'])] = (N if N is not None else ''), pt
        else:
            H_DET.fila(base + [numer, deno, v, None, clasif] + cola)
            det_val[(padre, cod, r['anio'])] = v if v is not None else ''

    # ------------------------------------------------------------------ DATOS
    P, G, VMAX, BR = {}, {}, {}, {}   # espejo: puntaje, valor, valor_max, brecha_rango por (código, año)
    filas_datos = []
    prin = [r for r in hechos if r['tipo'] == 'principal' and r['tipo_periodo'] == 'Año' and r['anio_num'] <= ANIO_CORTE]
    orden_cod = {c: i for i, c in enumerate(c for d in ce.DIM_ORDEN for c in ce.principales(d))}
    prin.sort(key=lambda r: (orden_cod[r['codigo']], r['anio_num']))
    for r in prin:
        cod, a, A = r['codigo'], r['anio_num'], r['anio']
        met = METODO[cod]
        F, Hh = num(r['numerador']), num(r['denominador'])
        csv_v = num(r['valor'])
        mx = por.get((cod + '.MAX', A))
        sinv = None
        if mx is not None and F is not None and num(mx['numerador']) is not None:
            sinv = num(mx['numerador']) - F
        url, otros = primera_url(r['enlace'])
        entrada_valor = None      # el valor es una entrada (no fórmula)
        calc = None
        if cod in ('PDM.CA', 'PDM.VEED') or (cod == 'PDM.CM' and F is None):
            entrada_valor = csv_v
            F = Hh = None
        tasa = cod == 'PDM.HOM' or (cod == 'PDM.CM' and F is not None)
        # valor (espejo + fórmula)
        if entrada_valor is not None or cod in ('PDM.CA', 'PDM.VEED') or (cod == 'PDM.CM' and F is None):
            vform, v = entrada_valor, ('' if entrada_valor is None else entrada_valor)
            vform = entrada_valor
        elif cod == 'PDM.SP':
            vs = [det_val.get(('PDM.SP', s, A), '') for s in ('PDM.SP.ACU', 'PDM.SP.ALC', 'PDM.SP.ASE', 'PDM.SP.ENE')]
            v = sum(vs) / 4 if all(x != '' for x in vs) else ''
            vform = (f'IF(COUNTIFS(Detalle!$C:$C,$C{{r}},Detalle!$G:$G,$E{{r}},Detalle!$E:$E,"subcomponente",Detalle!$K:$K,">=0")=4,'
                     f'AVERAGEIFS(Detalle!$K:$K,Detalle!$C:$C,$C{{r}},Detalle!$G:$G,$E{{r}},Detalle!$E:$E,"subcomponente"),"")', v)
        elif cod == 'PDM.ILE':
            ps = [det_pt.get(('PDM.ILE', s, A), '') for s in ('PDM.ILE.LA', 'PDM.ILE.LN')]
            v = sum(ps) / 2 if all(x != '' for x in ps) else ''
            vform = (f'IF(COUNTIFS(Detalle!$C:$C,$C{{r}},Detalle!$G:$G,$E{{r}},Detalle!$E:$E,"subcomponente",Detalle!$L:$L,">=0")=2,'
                     f'AVERAGEIFS(Detalle!$L:$L,Detalle!$C:$C,$C{{r}},Detalle!$G:$G,$E{{r}},Detalle!$E:$E,"subcomponente"),"")', v)
        else:
            mult = 100000 if tasa else 100
            v = '' if (F is None or not Hh) else mult * F / Hh
            vform = (f'IF(OR(F{{r}}="",H{{r}}=""),"",{mult}*F{{r}}/H{{r}})', v)
        if isinstance(vform, tuple):
            vcelda = vform
        else:
            vcelda = vform
        # puntaje
        if v == '':
            p = ''
        elif met == 'proporcion' or met == 'promedio':
            p = v
        elif met == 'tope':
            p = min(100, v)
        elif met == 'veedurias':
            p = min(100, 100 * v / K['VEED'])
        elif met == 'inverso':
            ref = K['CM'] if cod == 'PDM.CM' else K['HOM']
            p = 100 if v <= 0 else min(100, 100 * ref / v)
        else:   # avance
            lb, rf = (K['LB_CA'], K['REF_CA']) if cod == 'PDM.CA' else (K['LB_I1'], K['REF_I1'])
            p = max(0, min(100, 100 * (v - lb) / (rf - lb)))
        pf = {'proporcion': 'IF(I{r}="","",I{r})', 'promedio': 'IF(I{r}="","",I{r})', 'tope': 'IF(I{r}="","",MIN(100,I{r}))',
              'veedurias': f'IF(I{{r}}="","",MIN(100,100*I{{r}}/{K["VEED"]}))',
              'inverso': 'IF(I{r}="","",IF(I{r}<=0,100,MIN(100,100*%s/I{r})))' % (K['CM'] if cod == 'PDM.CM' else K['HOM']),
              'avance': 'IF(I{r}="","",MAX(0,MIN(100,100*(I{r}-%s)/(%s-%s))))' % ((K['LB_CA'], K['REF_CA'], K['LB_CA']) if cod == 'PDM.CA' else (K['LB_I1'], K['REF_I1'], K['LB_I1']))}[met]
        # rango
        if sinv is not None:
            vm = '' if (v == '' or not Hh) else min(100, 100 * (F + sinv) / Hh)
            br = '' if (v == '' or vm == '') else vm - v
            fmax = ('IF(OR(I{r}="",G{r}=""),"",MIN(100,100*(F{r}+G{r})/H{r}))', vm)
            fbr = ('IF(OR(I{r}="",J{r}=""),"",J{r}-I{r})', br)
        else:
            vm = br = None
            fmax = fbr = None
        P[(cod, A)], G[(cod, A)] = p, v
        BR[(cod, A)] = br if br not in (None, '') else 0
        # tiene_dato / motivo
        bt = r['brecha']
        if csv_v is not None:
            td, mot = 'Sí', None
        elif bt.lower().startswith('no aplica'):
            td, mot = 'No aplica', re.sub(r'^no aplica:\s*', '', bt, flags=re.I)
        else:
            td, mot = 'No', re.sub(r'^sin dato:\s*', '', bt, flags=re.I)
        if td == 'Sí' and v == '':
            raise AssertionError(('valor del CSV sin reproducir', cod, A))
        fila = H_DAT.fila([f'{cod}|{A}', r['dimension'], cod, r['indicador'], a, F, sinv, Hh, vcelda, fmax, fbr, (pf, p), (SEMAFORO_F.replace('{c}', 'L{r}'), semaforo(p)),
                           td, (mot[:300] if mot else None), r['fuente'] or None, url, otros, r['fecha_consulta'] or None, r['nota'] or None])
        filas_datos.append((cod, a, fila, td, r))
    # control: cada puntaje reproduce el CSV
    dif = []
    for cod, a, _f, _td, r in filas_datos:
        esp = num(r['puntaje_0_100'])
        got = P[(cod, str(a))]
        if (got == '') != (esp is None) or (esp is not None and abs(got - esp) > 0.005):
            dif.append((cod, a, got, esp))
    assert not dif, ('puntajes que no reproducen el CSV', dif[:5])

    # ------------------------------------------------------------------ ÍNDICES
    IX = {}
    for cod in ce.NOMBRE_INDICE:
        dim, nom = ce.dim_de(cod), ce.NOMBRE_INDICE[cod]
        for a in range(ce.PRIMER_ANIO[cod], ANIO_CORTE + 1):
            comp_cod = [c for c in ce.principales(dim)]
            con = sum(1 for c in comp_cod if P.get((c, str(a)), '') != '')
            ncomp = COMPONENTES[cod]
            assert ncomp == len(comp_cod), (cod, ncomp, comp_cod)
            ok = con == ncomp
            sd = lambda c, a=a: f'SUMIFS(Datos!$L:$L,Datos!$C:$C,"{c}",Datos!$E:$E,$E{{r}})'
            if cod in PESOS:
                fp = f'IF(G{{r}}<F{{r}},"",' + '+'.join(f'{w}*{sd(c)}' for c, w in PESOS[cod]) + ')'
                pv = '' if not ok else sum(w * P[(c, str(a))] for c, w in PESOS[cod])
            else:
                fp = f'IF(G{{r}}<F{{r}},"",AVERAGEIFS(Datos!$L:$L,Datos!$B:$B,"{PROMEDIO[cod]}",Datos!$E:$E,$E{{r}}))'
                pv = '' if not ok else sum(P[(c, str(a))] for c in comp_cod) / len(comp_cod)
            brs = lambda c: f'SUMIFS(Datos!$K:$K,Datos!$C:$C,"{c}",Datos!$E:$E,$E{{r}})'
            if cod == 'CS.ICRCS':
                fm = f'IF(H{{r}}="","",H{{r}}+SUMIFS(Datos!$K:$K,Datos!$B:$B,"CS",Datos!$E:$E,$E{{r}})/4)'
                mv = '' if pv == '' else pv + sum(BR[(c, str(a))] for c in comp_cod) / 4
            elif cod == 'PC.ISPC':
                fm = 'IF(H{r}="","",H{r}+' + '+'.join(f'{w}*{brs(c)}' for c, w in PESOS[cod]) + ')'
                mv = '' if pv == '' else pv + sum(w * BR[(c, str(a))] for c, w in PESOS[cod])
            else:
                fm, mv = None, None
            idx_row = por.get((cod, str(a)))
            mot = None
            if not ok:
                mot = re.sub(r'^sin dato:\s*', '', idx_row['brecha'], flags=re.I)[:300] if idx_row else None
            frm = fx[cod]['formula'].replace('\n', ' ')
            H_IND.fila([f'{cod}|{a}', dim, cod, nom, a, ncomp,
                        (f'COUNTIFS(Datos!$B:$B,$B{{r}},Datos!$E:$E,$E{{r}},Datos!$L:$L,">=0")', con), (fp, pv),
                        (fm, mv) if fm else None, (SEMAFORO_F.replace('{c}', 'H{r}'), semaforo(pv)),
                        ('IF(H{r}="","No","Sí")', 'No' if pv == '' else 'Sí'), mot, frm])
            IX[(cod, a)] = (pv, mv, con, ncomp)
    # control contra los CSV
    for (cod, a), (pv, mv, con, nc) in IX.items():
        esp = num(por[(cod, str(a))]['valor'])
        assert (pv == '' and esp is None) or (pv != '' and esp is not None and abs(pv - esp) <= 0.005), (cod, a, pv, esp)
        if mv not in (None, ''):
            mxr = por.get((cod + '.MAX', str(a)))
            assert mxr is not None and abs(mv - num(mxr['valor'])) <= 0.005, ('máximo', cod, a, mv, mxr and mxr['valor'])

    # ------------------------------------------------------------------ INDICADORES
    especial = {}
    for b in brechas:
        if b['codigo'] in ('PDM.VEED.BRECHA', 'CS.PQRSDF.BRECHA', 'CS.DSI.BRECHA'):
            especial[b['codigo'][:-7]] = (num(b['numerador']), num(b['denominador']), b['anio'])
    UNID_BR = {'PDM.VEED': 'consecutivos', 'CS.PQRSDF': 'trimestres', 'CS.DSI': 'metas (año × meta)'}
    LB = {'PDM.CA': (K['LB_CA'], K['REF_CA']), 'PC.I1': (K['LB_I1'], K['REF_I1']), 'PDM.VEED': (K['VEED'], K['VEED']),
          'PDM.CM': (float(ce.val('PDM.CM', 2019)), K['CM']), 'PDM.HOM': (3.7, K['HOM'])}
    cat_rows = {}
    n_orden = 0
    for dim in ce.DIM_ORDEN:
        for cod in [ce.INDICE_DIM[dim]] + ce.principales(dim):
            es_idx = cod == ce.INDICE_DIM[dim]
            f = fx[cod]
            n_orden += 1
            met = None if es_idx else METODO[cod]
            peso = None
            if not es_idx:
                if dim == 'PDM':
                    peso = ('1/7', 1 / 7)
                else:
                    peso = [w for c, w in PESOS.get(ce.INDICE_DIM[dim], []) if c == cod]
                    peso = peso[0] if peso else 0.25
            lb, rf = LB.get(cod, (None, None))
            if cod in ('PDM.SALUD', 'PDM.SP'):
                rf = 100
            elif met == 'proporcion' or cod == 'CS.PQRSDF':
                rf = 100
            ref_txt = '19 y 100.44' if cod == 'PDM.ILE' else None
            if cod in especial:
                sd, tot, per = especial[cod]
                fsd, ftot, fpct = sd, tot, ('IF(S{r}=0,"",100*R{r}/S{r})', 100 * sd / tot)
                bu = UNID_BR[cod]
            else:
                bu = 'índices (años)' if es_idx else 'años'
                if es_idx:
                    sdv = sum(1 for a in range(ce.PRIMER_ANIO[cod], 2026) if IX[(cod, a)][0] == '')
                    tov = 2026 - ce.PRIMER_ANIO[cod]
                    fsd = (f'COUNTIFS(Indices!$C:$C,$A{{r}},Indices!$K:$K,"No")', sdv)
                    ftot = (f'COUNTIFS(Indices!$C:$C,$A{{r}})', tov)
                else:
                    rows = [x for x in filas_datos if x[0] == cod]
                    sdv = sum(1 for x in rows if x[3] == 'No')
                    tov = len(rows) - sum(1 for x in rows if x[3] == 'No aplica')
                    fsd = (f'COUNTIFS(Datos!$C:$C,$A{{r}},Datos!$N:$N,"No")', sdv)
                    ftot = (f'COUNTIFS(Datos!$C:$C,$A{{r}})-COUNTIFS(Datos!$C:$C,$A{{r}},Datos!$N:$N,"No aplica")', tov)
                fpct = ('IF(S{r}=0,"",100*R{r}/S{r})', 100 * sdv / tov)
                sd, tot = sdv, tov
            url = ce.enlace_reciente(cod) if not es_idx else None
            H_CAT.fila([cod, n_orden, dim, ce.NOMBRE_DIM[dim], 'indice' if es_idx else 'principal', ce.INDICE_DIM[dim], f['nombre'], f['definicion'],
                        f['unidad'], f['sentido'], None, met or 'promedio o suma ponderada de componentes',
                        FORMULA_TXT[met] if met else f['formula'].replace('\n', ' '), lb, rf if ref_txt is None else ref_txt, f['ref_corto'], bu, fsd, ftot, fpct,
                        ce.brecha_texto(dim, cod), cf_fuente(cod), url or None])
            cat_rows[cod] = (H_CAT.n, peso, sd, tot)
    # peso: fórmula =1/7 en ICPD-7, número en los demás
    for cod, (fila, peso, sd, tot) in cat_rows.items():
        c = H_CAT.ws.cell(fila, 11)
        if peso is None:
            continue
        if peso[0] == '1/7' if isinstance(peso, tuple) else False:
            c.value = '=1/7'
            H_CAT.cache[c.coordinate] = 1 / 7
        else:
            c.value = peso
    # asserts del catálogo
    assert abs(sum(w for c, w in PESOS['PP.IREPP']) - 1) < 1e-12 and abs(sum(w for c, w in PESOS['PC.ISPC']) - 1) < 1e-12
    pesos_cat = {c: (v[1][1] if isinstance(v[1], tuple) else v[1]) for c, v in cat_rows.items() if v[1] is not None}
    for ix, comps in PESOS.items():
        for c, w in comps:
            assert abs(pesos_cat[c] - w) < 1e-12, (c, w)
    for dim, ix in (('PDM', 'PDM.ICPD7'), ('CS', 'CS.ICRCS')):
        assert abs(sum(pesos_cat[c] for c in ce.principales(dim)) - 1) < 1e-12, ix
    assert abs(sum(pesos_cat[c] for c in ce.principales('PC')) - 1) < 1e-12 and abs(sum(pesos_cat[c] for c in ce.principales('PP')) - 1) < 1e-12
    # las constantes de las fórmulas coinciden con el catálogo
    ws_d = H_DAT.ws
    for cod, (lb, rf) in {'PDM.CA': LB['PDM.CA'], 'PC.I1': LB['PC.I1']}.items():
        fila = [x for x in filas_datos if x[0] == cod][0][2]
        fml = ws_d.cell(fila, 12).value
        assert f'{lb}' in fml and f'{rf}' in fml, (cod, fml)
    for cod, (lb, rf) in LB.items():
        fila = [x for x in cat_rows if x == cod]
        c = H_CAT.ws.cell(cat_rows[cod][0], 15).value
        assert c == rf, (cod, c, rf)
    # brechas esperadas
    esp_br = {'PDM.ICPD7': (4, 7), 'PC.ISPC': (5, 6), 'CS.ICRCS': (4, 6), 'PC.I1': (2, 8), 'PC.I2': (1, 6), 'PP.IREPP': (0, 6), 'PP.EPP': (0, 6)}
    for cod, (s_, t_) in esp_br.items():
        assert (cat_rows[cod][2], cat_rows[cod][3]) == (s_, t_), (cod, cat_rows[cod][2:], (s_, t_))

    # ------------------------------------------------------------------ LÉAME
    leame(wl, fecha)

    # ------------------------------------------------------------------ cierre y guardado
    H_IND.cierra(formatos={'puntaje': '0.00', 'puntaje_max': '0.00', 'anio': '0'}, anchos={'id': 16, 'indice': 10, 'motivo_sin_dato': 50, 'formula': 70, 'dimension': 10},
                 envolver=('motivo_sin_dato', 'formula'))
    H_DAT.cierra(formatos={'numerador': '#,##0.##', 'sin_verificar': '#,##0.##', 'denominador': '#,##0.##', 'valor': '0.00', 'valor_max': '0.00',
                           'brecha_rango': '0.00', 'puntaje': '0.00', 'anio': '0'},
                 anchos={'id': 22, 'indicador': 44, 'motivo_sin_dato': 40, 'fuente': 50, 'enlace': 40, 'otros_enlaces': 30, 'nota': 60})
    H_CAT.cierra(formatos={'peso': '0.0000', 'brecha_pct': '0.0'},
                 anchos={'indicador': 40, 'definicion': 60, 'nombre_dimension': 30, 'formula_puntaje': 55, 'referente_fuente': 40, 'brecha_texto': 40,
                         'fuente_principal': 70, 'enlace_principal': 40, 'metodo_puntaje': 16},
                 envolver=('definicion', 'formula_puntaje', 'referente_fuente', 'brecha_texto', 'fuente_principal', 'indicador'))
    H_DET.cierra(formatos={'numerador': '#,##0.##', 'denominador': '#,##0.##', 'valor': '0.00', 'puntaje': '0.00', 'anio': '0'},
                 anchos={'id': 26, 'descripcion': 50, 'fuente': 40, 'enlace': 30, 'nota': 50, 'clasificacion': 24})
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ce_cache = {'Índices': H_IND.cache, 'Datos': H_DAT.cache, 'Indicadores': H_CAT.cache, 'Detalle': H_DET.cache}
    # catálogo de padres
    pad = {r_[2] for r_ in (H_DET.ws.iter_rows(min_row=2, values_only=True))}
    assert pad <= set(cat_rows), ('codigo_padre fuera del catálogo', pad - set(cat_rows))
    salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(salida)
    escribir_cache(salida, wb, ce_cache)
    return {'indices': H_IND.n - 1, 'datos': H_DAT.n - 1, 'indicadores': H_CAT.n - 1, 'detalle': H_DET.n - 1,
            'formulas': sum(len(v) for v in ce_cache.values())}


def cf_fuente(cod):
    import contenido_fichas as cf
    return cf.fuente_serie(cod)


def escribir_cache(ruta, wb, cache):
    """Escribe en el XML el <v> de cada fórmula (t="str" para texto); sin <v> si el resultado es vacío."""
    nombres = [ws.title for ws in wb.worksheets]
    tmp = Path(tempfile.mkdtemp()) / 'x.xlsx'
    with zipfile.ZipFile(ruta) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            m = re.fullmatch(r'xl/worksheets/sheet(\d+)\.xml', it.filename)
            if m:
                hoja = nombres[int(m.group(1)) - 1]
                if hoja in cache:
                    data = _parchar(data, cache[hoja])
            zo.writestr(it, data)
    shutil.move(str(tmp), str(ruta))


def _parchar(xml, valores):
    root = etree.fromstring(xml)
    q = lambda t: '{%s}%s' % (NS, t)
    for c in root.iter(q('c')):
        f = c.find(q('f'))
        if f is None:
            continue
        for v in c.findall(q('v')):
            c.remove(v)
        ref = c.get('r')
        assert ref in valores, ('fórmula sin espejo', ref)
        val = valores[ref]
        if val == '' or val is None:
            if 't' in c.attrib:
                del c.attrib['t']
            continue           # resultado "" -> sin <v>
        v = etree.SubElement(c, q('v'))
        if isinstance(val, str):
            c.set('t', 'str')
            v.text = val
        else:
            if 't' in c.attrib:
                del c.attrib['t']
            v.text = repr(float(val)) if not isinstance(val, int) else str(val)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
