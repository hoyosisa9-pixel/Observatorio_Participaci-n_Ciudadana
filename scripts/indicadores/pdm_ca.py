"""PDM.CA (espacio público verde): CA = MIN(100, 100*O/M). O del diagnóstico del POT (2024), M = 3,88 (PDM p. 717).

La serie del POT (EPV 2024, método AMVA) NO es la serie de la línea base del plan (3,88 en 2023): se rotula en la alerta.
Escribe PDM.CA.O, PDM.CA.M (verificada por pdm_metas.extraer_todas) y los resultados de PDM.CA, todos los escenarios.
Uso: python3 scripts/indicadores/pdm_ca.py   (sin -I, desde la raíz del repo)
"""
import re
import sys
import datetime as dt
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import pdm_metas as PM
import pymupdf

SCRIPT = "scripts/indicadores/pdm_ca.py"
RUTA_POT = ("data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/"
            "8-dimension-funcional-_-espacio-publico.pdf")
NOMBRE_O = "Espacio público verde observado, en metros cuadrados por habitante"
ALERTA = "serie distinta de la línea base: EPV del POT 2024 (método AMVA) frente a 3,88 de 2023 del plan"


def leer_pot():
    """Lee y verifica en el PDF del POT los valores de las tablas 38 y 39."""
    doc = pymupdf.open(C.RAIZ / RUTA_POT)
    t86, t87 = PM.texto_pagina(doc, 86), PM.texto_pagina(doc, 87)
    m = re.search(r"Tabla 38\..*?248\.304 9 ([\d.]+) ([\d,]+)", t86)
    if not m or m.group(1) != "17.233.624" or m.group(2) != "69,41":
        raise RuntimeError("POT p. 86: no coincide la tabla 38 (EPV total 69,41)")
    m = re.search(r"240\.655 4\.66 9 ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+)", t87)
    if not m or m.group(1) != "1.382.762" or m.group(2) != "5.75":
        raise RuntimeError("POT p. 87: no coincide la tabla 39 (EPV urbano 5,75)")
    total, urbano = 69.41, 5.75
    # coherencia aritmética de lo publicado (redondeo a 2 decimales)
    if abs(1382762 / 240655 - urbano) > 0.005 or abs(17233624 / 248304 - total) > 0.005:
        raise RuntimeError("POT: el cociente no reproduce el indicador publicado")
    return {"urbano": urbano, "total": total}


def main():
    pot = leer_pot()
    meta = PM.extraer_todas()["PDM.CA.M"]
    url_pot = C.url_de_ruta(RUTA_POT)
    anio = 2024
    real = dt.date(2024, 12, 31)   # el diagnóstico calcula el EPV "al año 2024" con población DANE 2024
    filas_v, filas_r = [], []
    for esc in C.escenarios("PDM.CA"):
        p = esc["params"]
        fuente = p["pdm_ca_fuente"]
        fc = C.fecha_corte_registro(anio, "Anual", real, p)
        if fuente == "pot_epv_urbano":
            O, est = pot["urbano"], "Pendiente"
            pag = "PDF p. 87, tabla 39 (EPV urbano 2024: 1.382.762 m2 / 240.655 hab. = 5,75); línea base AMVA 2022 = 4,66"
            obs = ("EPV urbano 2024 del diagnóstico del POT (método AMVA). Serie distinta de la línea base del plan "
                   "(3,88 en 2023); no se confunda con espacio público efectivo (2,67) ni total (70,54). "
                   "Corte: año calendario 2024, corte real no publicado al día.")
        elif fuente == "pot_epv_total":
            O, est = pot["total"], "Pendiente"
            pag = "PDF p. 86, tabla 38 (EPV total 2024: 17.233.624 m2 / 248.304 hab. = 69,41)"
            obs = ("EPV total 2024 del POT; incluye amplias áreas de reserva natural y el propio diagnóstico advierte "
                   "que la cifra queda distorsionada. Serie distinta de la línea base del plan (3,88). "
                   "Corte: año calendario 2024.")
        else:   # estricto_serie_linea_base
            O, est = None, "No disponible"
            pag = "PDF pp. 86-87 (tablas 38 y 39): serie distinta de la línea base"
            obs = ("Escenario estricto: solo vale la serie de la línea base del plan (3,88 en 2023, fuente "
                   "Secretaría de Medio Ambiente), que no se publica para 2024; el POT no es comparable. N/D.")
        filas_v.append(C.fila_valor(
            "PDM", "PDM.CA.O", anio, O, "m² / habitante", esc, codigo_indicador="PDM.CA", nombre_variable=NOMBRE_O,
            tipo_valor="observado", periodo="Anual", territorio="Municipio", estado_informacion=est,
            categoria_disponibilidad="Parcial",
            fuente="Documento diagnóstico de la revisión del POT, 8. Dimensión funcional - Espacio público (2025)",
            documento_ruta_local=RUTA_POT, enlace=url_pot, pagina_tabla=pag, fecha_corte=fc, fecha_corte_real=real,
            provisional=True, script=SCRIPT, observaciones=obs))
        filas_v += [f for f in PM.filas_meta("PDM.CA.M", meta, esc) if f["anio"] == 2024]
        filas_v += [f for f in PM.filas_meta("PDM.CA.M", meta, esc) if f["anio"] == 2025]
        M = None if p["horizonte_meta_pdm"] == "metas_anuales" else meta["meta"]
        res = C.calcular("PDM.CA", [O, M], {"periodo": "Anual", "territorio": "Municipio"})
        alerta = ALERTA if O is not None else ""
        filas_r.append(C.fila_resultado(
            "PDM.CA", anio, esc, res, [O, M], nombre="Cumplimiento del espacio público verde (CA)",
            tipo="Indicador", fecha_corte_comun=fc, alerta=alerta, script=SCRIPT,
            observaciones=("CA = MIN(100, 100*O/M). " + (f"O={O:g} (POT 2024), M={M:g} (PDM p. 717)."
                           if O is not None and M is not None else res["motivo"]))))
    # PDM.CA.M puede venir repetida entre 2024 y 2025 sin conflicto: la llave incluye el año
    C.escribir_parcial("PDM", "valores_variables", filas_v, codigos=["PDM.CA"])
    C.escribir_parcial("PDM", "resultados_indicadores", filas_r, codigos=["PDM.CA"])
    for f in filas_r:
        print(f"{f['escenario_id'][:70]:70s} O={f.get('entrada_1')} M={f.get('entrada_2')} "
              f"puntaje={f['puntaje']} {f['semaforo']} {f['disponibilidad']}")


if __name__ == "__main__":
    main()
