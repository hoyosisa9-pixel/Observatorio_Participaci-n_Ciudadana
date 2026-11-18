"""VRC con brecha de información: cota inferior, cota superior y brecha (provisional).

Lee outputs/indicadores/clasificacion_vrc.csv (un evento por fila) y, por año del evento y universo:
  cota inferior = eventos con veeduría documentada / eventos        (desconocida cuenta como no documentada)
  cota superior = (documentada + desconocida) / eventos              (desconocida cuenta como presente)
  brecha        = eventos con presencia desconocida / eventos
Salida: outputs/revision_indicadores/CS/vrc_brecha.csv (; y BOM).
"""
import csv
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
ENTRADA = RAIZ / 'outputs/indicadores/clasificacion_vrc.csv'
SALIDA = RAIZ / 'outputs/revision_indicadores/CS/vrc_brecha.csv'
UNIVERSOS = {
    'u_alcalde_mas_sectoriales_con_acta': 'Principal: audiencia del Alcalde + rendiciones sectoriales con acta o memoria (el usado en el respaldo)',
    'u_amplio_todo_documento': 'Sensibilidad: todo evento de rendición con algún documento publicado',
}


def main():
    filas = list(csv.DictReader(ENTRADA.open(encoding='utf-8-sig'), delimiter=';'))
    salida = []
    for col, desc in UNIVERSOS.items():
        for anio in ('2024', '2025'):
            ev = [f for f in filas if f['anio_evento'] == anio and f[col] == '1']
            si = sum(f['mencion_si'] == '1' for f in ev)
            no = sum(f['mencion_no'] == '1' for f in ev)
            nd = sum(f['mencion_desconocida'] == '1' for f in ev)
            n = len(ev)
            assert si + no + nd == n, (anio, col)
            salida.append({
                'anio': anio, 'universo': col, 'descripcion_universo': desc, 'eventos': n,
                'con_veeduria_documentada': si, 'sin_veeduria_documentada': no, 'presencia_desconocida': nd,
                'cota_inferior_pct': round(100 * si / n, 2), 'cota_superior_pct': round(100 * (si + nd) / n, 2),
                'brecha_informacion_pct': round(100 * nd / n, 2),
                'eventos_desconocidos': ' | '.join(f"{f['id']} {f['evento']}" for f in ev if f['mencion_desconocida'] == '1'),
                'estado': 'Provisional',
            })
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with SALIDA.open('w', encoding='utf-8-sig', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=list(salida[0]), delimiter=';')
        w.writeheader(); w.writerows(salida)
    for s in salida:
        print(s['anio'], s['universo'][:22], 'n', s['eventos'], 'sí', s['con_veeduria_documentada'], 'no', s['sin_veeduria_documentada'],
              'nd', s['presencia_desconocida'], '->', s['cota_inferior_pct'], '–', s['cota_superior_pct'], 'brecha', s['brecha_informacion_pct'])
        print('    ', s['eventos_desconocidos'])


if __name__ == '__main__':
    main()
