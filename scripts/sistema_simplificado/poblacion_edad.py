#!/usr/bin/env python3
"""Extrae de la proyección DANE PPED 2018-2042 (edad simple) las filas 'Total' de Envigado y guarda una versión
compacta (suma hombres+mujeres por edad simple) en outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv.
Uso (desde la raíz, sin -I): python3 scripts/sistema_simplificado/poblacion_edad.py
"""
from __future__ import annotations
import csv, re, sys
from pathlib import Path
import openpyxl

RAIZ = Path(__file__).resolve().parents[2]
XLSX = RAIZ / "data_raw_IIC/06_descargas_indicadores/dane/PPED-AreaSexoEdadMun-2018-2042_VP.xlsx"
SALIDA = RAIZ / "outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv"


def extraer() -> None:
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    ws = wb["PobMunicipalxÁreaSexoEdad"]
    cab, filas = None, []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 8:
            cab = list(row)
        if row and str(row[2]) == "05266" and row[5] == "Total":
            filas.append(row)
    cols = {}
    for j, c in enumerate(cab):
        m = re.match(r"(Hombres|Mujeres) (\d+)", str(c or ""))
        if m:
            cols.setdefault(int(m.group(2)), []).append(j)
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["anio", "edad", "poblacion"])
        for r in filas:
            for edad in sorted(cols):
                if edad <= 10:
                    w.writerow([r[4], edad, sum(int(r[j]) for j in cols[edad])])
            w.writerow([r[4], "total", int(r[6])])


if __name__ == "__main__":
    extraer()
