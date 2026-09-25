"""Genera una copia del inventario apta para abrir con doble clic en Excel (separador ';', UTF-8 con BOM)."""
import csv
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ORIGEN = RAIZ / 'data_raw_IIC' / 'inventario_detallado_IIC.csv'
DESTINO = RAIZ / 'data_raw_IIC' / 'inventario_detallado_IIC_excel.csv'

with open(ORIGEN, encoding='utf-8', newline='') as fin, open(DESTINO, 'w', encoding='utf-8-sig', newline='') as fout:
    escritor = csv.writer(fout, delimiter=';', quoting=csv.QUOTE_MINIMAL)
    filas = 0
    for fila in csv.reader(fin):
        escritor.writerow(fila)
        filas += 1
print(f'Generado {DESTINO} ({filas - 1} registros)')
