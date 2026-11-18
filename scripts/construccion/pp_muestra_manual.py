#!/usr/bin/env python3
"""Registro de la revision manual del cruce de PP: proyectos cuyas filas se leyeron a ojo en la pagina del PDF de Hacienda
(recortes de la pagina renderizada) y se compararon con pp_cruce_proyectos.csv. La lista de proyectos y paginas es la que
se reviso (18-nov-2026); el script solo vuelve a poner al lado los valores del CSV. Todo resultado es PROVISIONAL.
Uso: python3 scripts/construccion/pp_muestra_manual.py   (desde la raiz; despues de pp_cruce.py)
"""
import csv
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
CRUCE = RAIZ / "outputs/construccion/PP/pp_cruce_proyectos.csv"
SALIDA = RAIZ / "outputs/construccion/PP/pp_muestra_manual.csv"
# (anio, sufijo, paginas del PDF leidas)
MUESTRA = [
    (2020, "0216", "14"), (2020, "0475", "21"), (2020, "0480", "24"), (2020, "0252", "29"), (2020, "0447", "31"), (2020, "0495", "20"),
    (2021, "0407", "94, 96"), (2021, "0232", "94, 96"), (2021, "0647", "128"), (2021, "0396", "94, 96"), (2021, "0234", "94, 96"),
    (2022, "0407", "78, 79"), (2022, "0704", "115, 117"), (2022, "0710", "53"), (2022, "0698", "66"), (2022, "0234", "78, 79"),
    (2023, "0407", "64, 65"), (2023, "0729", "91, 92"), (2023, "0735", "64, 65"), (2023, "0400", "64, 65"), (2023, "0386", "64, 65"),
    (2024, "0677", "43"), (2024, "0665", "42"), (2024, "0667", "43"), (2024, "0669", "43"), (2024, "0671", "43"), (2024, "0673", "43"),
]
COLS = ["anio", "id_efpc", "codigo_sufijo", "paginas_pdf_leidas", "valor_votado", "apropiacion_inicial",
        "apropiacion_definitiva", "obligacion", "tipo_cruce", "resultado_revision_manual", "nota_registro"]
filas = []
with open(CRUCE, encoding="utf-8-sig", newline="") as f:
    P = list(csv.DictReader(f, delimiter=";"))
for a, suf, pg in MUESTRA:
    x = next(r for r in P if r["anio"] == str(a) and r["codigo_sufijo"] == suf)
    filas.append({"anio": a, "id_efpc": x["id_efpc"], "codigo_sufijo": suf, "paginas_pdf_leidas": pg, "valor_votado": x["valor_votado"],
                  "apropiacion_inicial": x["apropiacion_inicial"], "apropiacion_definitiva": x["apropiacion_definitiva_sin_2_3_7_06"],
                  "obligacion": x["obligacion_sin_2_3_7_06"], "tipo_cruce": x["tipo_cruce"],
                  "resultado_revision_manual": "coincide: las filas del sufijo en la pagina suman el valor votado y las cifras del CSV",
                  "nota_registro": "'coincide' es el registro de una revision manual hecha por el ejecutor (lectura a ojo de la pagina); este script NO lo verifica"})
with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
    w.writeheader()
    w.writerows(filas)
print(len(filas), "proyectos revisados ->", SALIDA)
