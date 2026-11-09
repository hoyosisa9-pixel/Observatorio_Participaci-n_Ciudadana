"""PDM: las 7 metas del Plan de Desarrollo 2024-2027, leídas del texto de la página del PDF (no copiadas a mano).

Falla si el valor hallado en la página no coincide con el esperado. Contraste con el DTS (22-dic-2025).
CA.M se verifica aquí, pero la fila PDM.CA.M la escribe pdm_ca.py (mismo indicador PDM.CA; evita que un script
borre las filas del otro). Uso: python3 scripts/indicadores/pdm_metas.py   (sin -I, desde la raíz del repo)
"""
import re
import sys
import datetime as dt
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import pymupdf

SCRIPT = "scripts/indicadores/pdm_metas.py"
RUTA_PDM = "data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf"
RUTA_DTS = "data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/documentotecnicosoportepdm22122025.pdf"
ANIOS = (2024, 2025)
NUM = r"\d+(?:[.,]\d+)?"

# id_variable: (indicador, nombre, unidad, etiqueta del indicador en la tabla, unidad en el PDF, pág. PDM, pág. DTS,
#               línea base esperada, meta esperada)
METAS = {
    "PDM.CS.M": ("PDM.CS", "Cobertura meta del aseguramiento en el régimen subsidiado", "Porcentaje (0–100)",
                 r"Cobertura del aseguramiento en salud del Régimen Subsidiado", "Porcentaje", 602, 640, 99.5, 99.5),
    "PDM.CM.M": ("PDM.CM", "Tasa meta de mortalidad vial por 100.000 habitantes", "Muertes / 100.000 hab.",
                 r"Tasa de mortalidad por siniestros viales", r"Tasa por cada 100\.000 habitantes", 706, 738, 10.2, 8.0),
    "PDM.CA.M": ("PDM.CA", "Meta de espacio público verde, en metros cuadrados por habitante", "m² / habitante",
                 r"Espacio público verde", r"Metros cuadrados(?: por habitante)?", 717, 748, 3.88, 3.88),
    "PDM.CP.M": ("PDM.CP", "Porcentaje meta de implementación de esos procesos", "Porcentaje (0–100)",
                 r"Implementación de los procesos de participación ciudadana", "Porcentaje", 742, 775, 75.0, 100.0),
    "PDM.CD.M": ("PDM.CD", "Cobertura porcentual meta de esos servicios deportivos", "Porcentaje (0–100)",
                 r"Cobertura de servicios deportivos de disciplinas tradicionales, nuevas tendencias y adaptadas",
                 "Porcentaje", 609, 646, 5.3, 6.0),
    "PDM.ILE.MA": ("PDM.ILE", "Meta de cobertura de atención integral inicial", "Porcentaje (0–100)",
                   r"Tasa de cobertura en atención integral de niños y niñas de 0 a 6 años", "Porcentaje", 650, 683,
                   16.3, 19.0),
    "PDM.ILE.MS": ("PDM.ILE", "Meta de instituciones oficiales que mejoran su categoría en Saber 11",
                   "Porcentaje (0–100)",
                   r"Instituciones educativas oficiales que mejoran la categoría de clasificación en pruebas Saber 11",
                   "Porcentaje", 652, 685, 0.0, 14.0),
}
ESPERADO = {"PDM.CS.M": 99.5, "PDM.CM.M": 8, "PDM.CA.M": 3.88, "PDM.CP.M": 100, "PDM.CD.M": 6,
            "PDM.ILE.MA": 19, "PDM.ILE.MS": 14}


def _n(s):
    return float(s.replace(",", "."))


def texto_pagina(doc, pag):
    return re.sub(r"\s+", " ", doc[pag - 1].get_text())


def leer_meta(texto, etiqueta, unidad):
    """Devuelve (línea_base, meta) de la fila del indicador en el texto aplanado de la página, o None."""
    etq = etiqueta.replace(" ", r"\s+")
    patron = (etq + r"\s+" + unidad + r"\s+(?:Estable|Creciente|Decreciente|Monitoreo)?\s*(" + NUM +
              r")(?:\s*\(\s*(?:Promedio\s+)?[\d\- ]+\))?\s+(" + NUM + ")")
    m = re.search(patron, texto, flags=re.I)
    return (_n(m.group(1)), _n(m.group(2))) if m else None


def extraer_todas():
    """Lee y verifica las 7 metas en el PDM (obligatorio, falla si no coincide) y en el DTS (contraste)."""
    pdm = pymupdf.open(C.RAIZ / RUTA_PDM)
    dts = pymupdf.open(C.RAIZ / RUTA_DTS)
    res = {}
    for idv, (ind, nom, uni, etq, unid, p_pdm, p_dts, lb_esp, meta_esp) in METAS.items():
        hall = leer_meta(texto_pagina(pdm, p_pdm), etq, unid)
        if hall is None:
            raise RuntimeError(f"{idv}: no se halló la fila del indicador en la p. {p_pdm} del PDM")
        lb, meta = hall
        if abs(meta - ESPERADO[idv]) > 1e-9 or abs(meta - meta_esp) > 1e-9 or abs(lb - lb_esp) > 1e-9:
            raise RuntimeError(f"{idv}: PDM p. {p_pdm} trae línea base {lb} y meta {meta}; esperado {lb_esp} y {meta_esp}")
        h2 = leer_meta(texto_pagina(dts, p_dts), etq, unid)
        if h2 is None:
            contraste = f"DTS p. {p_dts}: no se pudo leer la fila por el formato del texto (sin verificar)"
        elif abs(h2[1] - meta) > 1e-9:
            raise RuntimeError(f"{idv}: DTS p. {p_dts} trae meta {h2[1]} distinta de la del PDM ({meta})")
        else:
            contraste = f"DTS 22-dic-2025 p. {p_dts}: misma meta ({h2[1]:g}) y línea base ({h2[0]:g})"
        res[idv] = dict(meta=meta, linea_base=lb, pag_pdm=p_pdm, pag_dts=p_dts, contraste=contraste)
    return res


def filas_meta(idv, datos, esc):
    """Filas de valores_variables de una meta (2024 y 2025) para un escenario."""
    ind, nom, uni, *_ = METAS[idv]
    horizonte = esc["params"].get("horizonte_meta_pdm", "meta_2027")
    url = C.url_de_ruta(RUTA_PDM)
    filas = []
    for anio in ANIOS:
        real = dt.date(anio, 12, 31)
        obs = (f"Línea base {datos['linea_base']:g}. Meta única del cuatrienio (horizonte 2027). {datos['contraste']}. "
               f"Apéndice 9.1.2 del PDM según notas_PDM.md. Folio impreso = índice del PDF.")
        if idv == "PDM.CA.M":
            obs += " El PDM aprobado rotula la unidad 'Metros cuadrados'; el DTS y el apéndice dicen 'por habitante'."
        if horizonte == "metas_anuales":
            valor, estado = None, "No disponible"
            obs = "No existen metas anuales de resultado en el PDM ni en el DTS (parámetro horizonte_meta_pdm=metas_anuales). " + obs
        else:
            valor, estado = datos["meta"], "Pendiente"
        filas.append(C.fila_valor(
            "PDM", idv, anio, valor, uni, esc, codigo_indicador=ind, nombre_variable=nom, tipo_valor="meta",
            periodo="Anual", territorio="Municipio", estado_informacion=estado, categoria_disponibilidad="Disponible",
            fuente="Plan de Desarrollo 2024-2027 (Acuerdo 017 de 2024), indicadores de resultado programáticos",
            documento_ruta_local=RUTA_PDM, enlace=url, pagina_tabla=f"PDF p. {datos['pag_pdm']} (tabla Indicadores de resultado programáticos)",
            fecha_corte=C.fecha_corte_registro(anio, "Anual", real, esc["params"]), fecha_corte_real=real,
            horizonte_meta=2027, provisional=True, script=SCRIPT, observaciones=obs))
    return filas


def main():
    datos = extraer_todas()
    filas = []
    inds = sorted({v[0] for v in METAS.values()} - {"PDM.CA"})   # CA.M lo escribe pdm_ca.py
    for idv, d in datos.items():
        ind = METAS[idv][0]
        if ind == "PDM.CA":
            continue
        for esc in C.escenarios(ind):
            filas += filas_meta(idv, d, esc)
    ruta = C.escribir_parcial("PDM", "valores_variables", filas, codigos=inds)
    for idv, d in datos.items():
        print(f"{idv}: meta {d['meta']:g} (línea base {d['linea_base']:g}) PDM p. {d['pag_pdm']}; {d['contraste']}")
    print(f"{len(filas)} filas -> {ruta}")


if __name__ == "__main__":
    main()
