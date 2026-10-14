"""Reconstruye la ejecucion presupuestal del Presupuesto Participativo (PP.EPP) de la vigencia 2025.

PRELIMINAR, POR VALIDAR: la lista de proyectos PP no esta publicada con un marcador oficial; debe confirmarla
el Departamento Administrativo de Planeacion (DAP).

Uso (desde la raiz del repo):  python3 -I scripts/epp_pp_2025.py
Requiere openpyxl y PyMuPDF (`import pymupdf`).

Metodo:
  1. Plan de accion IV trimestre 2025 (hoja 'PlanAcción27.01.2026', corte 07-01-2026): por cada CODIGO PROYECTO
     se toman APROPIACION DEFINITIVA y EJECUCION OBLIGACION (vienen en una sola fila por proyecto).
  2. Proyectos PP = los que tienen 'ZONA <n>' en el nombre. Se excluye el 0553 'Implementacion del presupuesto
     participativo' (costo del proceso, no un proyecto priorizado) y se informa aparte.
  3. Contraste con el listado EFPC 'Proyectos_PP_2025.pdf' (zona y valor por proyecto): mismo numero de proyectos
     por zona y coincidencia de valores al peso (emparejamiento por zona y valor).
  4. EPP = 100 * OPP / APP (municipal y por zona).

Salidas en outputs/disponibilidad/ (sin datos personales):
  epp_pp_2025_proyectos.csv   proyecto, zona, apropiacion, obligacion, valor EFPC emparejado
  epp_pp_2025_zonas.csv       APP, OPP y EPP por zona y consolidado municipal
  epp_pp_2025.md              resumen
"""
import csv, re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
import pymupdf

RAIZ = Path(__file__).resolve().parent.parent
F_PLAN = (RAIZ / 'data_raw_IIC' / '02_ejecucion_y_presupuesto' / 'envigado_gov_co' / 'seguimiento-a-proyectos-2025'
          / 'evaluaciondefinitivaplanaccionivtri.xlsx')
HOJA_PLAN = 'PlanAcción27.01.2026'
F_EFPC = (RAIZ / 'data_raw_IIC' / '01_propuestas_y_acuerdos' / 'efpc_envigado_gov_co' / 'efpc_2025-07'
          / 'Proyectos_PP_2025.pdf')
SALIDA = RAIZ / 'outputs' / 'disponibilidad'
COSTO_PROCESO = '0553'
RE_ZONA = re.compile(r'ZONA\s*(\d+)')
ETIQUETA = 'PRELIMINAR, POR VALIDAR'


def pesos(x):
    return '$' + f'{x:,.0f}'.replace(',', '.')


def pct(x):
    return f'{x:.2f}'.replace('.', ',') + ' %'


def leer_plan():
    ws = openpyxl.load_workbook(F_PLAN, read_only=True, data_only=True)[HOJA_PLAN]
    proy = {}
    for r in ws.iter_rows(min_row=4, values_only=True):
        cod, nom = r[0], r[1]
        if not cod or not nom:
            continue
        p = proy.setdefault(str(cod), {'nombre': str(nom), 'ap': [], 'ob': []})
        if r[13] not in (None, ''):
            p['ap'].append(float(r[13]))
        if r[14] not in (None, ''):
            p['ob'].append(float(r[14]))
    for cod, p in proy.items():
        # una sola fila con valores financieros por proyecto; si hubiera mas, el script se detiene
        assert len(p['ap']) <= 1 and len(p['ob']) <= 1, f'proyecto {cod} con varias filas financieras'
        p['ap'] = p['ap'][0] if p['ap'] else 0.0
        p['ob'] = p['ob'][0] if p['ob'] else 0.0
    return proy


def leer_efpc():
    """Lista de (zona, valor) del listado EFPC; la zona es una linea con solo el numero antes del proyecto."""
    lineas = [l.strip() for pag in pymupdf.open(F_EFPC) for l in pag.get_text().splitlines()]
    regs, zona = [], None
    for l in lineas:
        if re.fullmatch(r'\d{1,2}', l):
            zona = int(l)
        m = re.fullmatch(r'\$\s*([\d.,]+)', l)
        if m:
            regs.append((zona, int(re.sub(r'[.,]', '', m.group(1)))))
    return regs


def main():
    plan = leer_plan()
    pp = {c: p for c, p in plan.items() if RE_ZONA.search(p['nombre'].upper()) and c != COSTO_PROCESO}
    for p in pp.values():
        p['zona'] = int(RE_ZONA.search(p['nombre'].upper()).group(1))
    efpc = leer_efpc()

    # emparejamiento por zona: primero valores identicos, luego el resto en orden (diferencias por adiciones)
    efpc_z = defaultdict(list)
    for z, v in efpc:
        efpc_z[z].append(v)
    for z in sorted({p['zona'] for p in pp.values()}):
        libres = list(efpc_z[z])
        prz = sorted((c for c in pp if pp[c]['zona'] == z))
        for c in prz:
            v = round(pp[c]['ap'])
            if v in libres:
                libres.remove(v)
                pp[c]['efpc'], pp[c]['coincide'] = v, 'si'
        for c in prz:
            if 'efpc' not in pp[c]:
                pp[c]['efpc'] = libres.pop(0) if libres else None
                pp[c]['coincide'] = 'no'

    SALIDA.mkdir(parents=True, exist_ok=True)
    with open(SALIDA / 'epp_pp_2025_proyectos.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['codigo_proyecto', 'nombre_proyecto', 'zona', 'apropiacion_definitiva', 'obligacion',
                    'valor_listado_efpc', 'apropiacion_igual_a_efpc', 'estado'])
        for c in sorted(pp, key=lambda c: (pp[c]['zona'], c)):
            p = pp[c]
            w.writerow([c, p['nombre'], p['zona'], round(p['ap']), round(p['ob']), p['efpc'], p['coincide'], ETIQUETA])

    zonas = defaultdict(lambda: [0, 0.0, 0.0])
    for p in pp.values():
        zonas[p['zona']][0] += 1
        zonas[p['zona']][1] += p['ap']
        zonas[p['zona']][2] += p['ob']
    app = sum(p['ap'] for p in pp.values())
    opp = sum(p['ob'] for p in pp.values())
    epp = 100 * opp / app
    cp = plan.get(COSTO_PROCESO)
    epp_con = 100 * (opp + cp['ob']) / (app + cp['ap']) if cp else None
    n_efpc = Counter(z for z, _ in efpc)
    with open(SALIDA / 'epp_pp_2025_zonas.csv', 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['nivel', 'proyectos_plan', 'proyectos_efpc', 'APP', 'OPP', 'EPP', 'estado'])
        for z in sorted(zonas):
            n, a, o = zonas[z]
            w.writerow([f'Zona {z}', n, n_efpc.get(z, 0), round(a), round(o), round(100 * o / a, 2), ETIQUETA])
        w.writerow(['Municipio', len(pp), len(efpc), round(app), round(opp), round(epp, 2), ETIQUETA])

    iguales = sum(p['coincide'] == 'si' for p in pp.values())
    md = [
        '# Ejecución presupuestal del Presupuesto Participativo 2025 (PP.EPP)',
        '',
        f'**Estado: {ETIQUETA}.** La lista de proyectos PP se reconstruye; debe validarla el DAP. '
        'Generado por `scripts/epp_pp_2025.py`.',
        '',
        f'- Fuente financiera: `{F_PLAN.relative_to(RAIZ)}`, hoja `{HOJA_PLAN}` (corte 07-01-2026).',
        f'- Contraste: `{F_EFPC.relative_to(RAIZ)}`.',
        f'- Proyectos PP (nombre con "ZONA n", sin el {COSTO_PROCESO}): **{len(pp)}**; en el listado EFPC: **{len(efpc)}**. '
        f'Apropiación igual al valor EFPC al peso: {iguales} de {len(pp)} (el resto difiere por adiciones).',
        f'- Proyectos por zona iguales en ambas fuentes: {"sí" if all(zonas[z][0] == n_efpc.get(z, 0) for z in zonas) else "no"}.',
        f'- **APP = {pesos(app)}** · **OPP = {pesos(opp)}** · **EPP = {pct(epp)}**.',
    ]
    if cp:
        md.append(f'- Con el proyecto {COSTO_PROCESO} (costo del proceso: APP {pesos(cp["ap"])}, OPP {pesos(cp["ob"])}) '
                  f'el EPP sería {pct(epp_con)}. Se excluye por defecto; decisión metodológica pendiente.')
    md += ['', 'Detalle por zona en `epp_pp_2025_zonas.csv` y por proyecto en `epp_pp_2025_proyectos.csv`.', '']
    (SALIDA / 'epp_pp_2025.md').write_text('\n'.join(md), encoding='utf-8')
    print(f'Proyectos {len(pp)} (EFPC {len(efpc)}), iguales {iguales}; APP {app:.0f}; OPP {opp:.0f}; EPP {epp:.2f} %'
          + (f'; con {COSTO_PROCESO}: {epp_con:.2f} %' if cp else ''))


if __name__ == '__main__':
    main()
