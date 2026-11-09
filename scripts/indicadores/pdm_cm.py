"""PDM.CM (mortalidad vial 2024): CM = IF(O=0, 100, MIN(100, 100*M/O)); M = 8 (PDM p. 706, leída por pdm_metas).

O según `pdm_cm_fuente` y `poblacion_dane` (todos los escenarios de C.escenarios("PDM.CM")):
  inmlcf_forensis  + serie_publicada_por_fuente : tasa publicada 7,25 (C.forensis_2024(): PDF local o cifra citada, Pendiente)
  inmlcf_forensis  + dane_postcovid / dane_vp   : 18 muertes (Forensis p. 37/440, citada) / población DANE del año
  policia_datosgov + serie_publicada_por_fuente : 7,7 de TerriData (DNP con datos de MinDefensa y DANE; es la tasa publicada
                                                  a partir de homicidios en accidente de tránsito de la Policía)
  policia_datosgov + dane_*                     : suma de `cantidad` 2024 del CSV de datos.gov.co / población DANE del año
Escribe PDM.CM.O, PDM.CM.M, AUX.PDM.CM.MUERTES y AUX.PDM.CM.POBLACION y los resultados de PDM.CM.
Ejecutar DESPUÉS de pdm_metas.py (que también escribe filas PDM.CM.M; aquí se regeneran las mismas, sin conflicto).
Uso: python3 scripts/indicadores/pdm_cm.py   (sin -I, desde la raíz del repo)
"""
import csv
import datetime as dt
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import pdm_metas as PM
import openpyxl

SCRIPT = "scripts/indicadores/pdm_cm.py"
ANIO = 2024
D = C.DIR_DESCARGAS
RUTA_POLICIA = D / "policia_datosgov" / "uav5-b85g_homicidio_accidente_transito_envigado.csv"
URL_POLICIA = "https://www.datos.gov.co/resource/uav5-b85g.csv?cod_muni=05266&$limit=50000"
DANE = {
    "dane_postcovid_2020_2035": ("DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx", "Hoja1",
                                 "https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/"
                                 "DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx"),
    "dane_vp_2018_2042": ("PPED-AreaMun-2018-2042_VP.xlsx", "PobMunicipalxÁrea",
                          "https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/"
                          "PPED-AreaMun-2018-2042_VP.xlsx"),
}
URL_TD = "https://terridata.dnp.gov.co/"
TD_RUTA = "data_raw_IIC/05_fuentes_externas_y_pot/dnp_terridata/TerriData05266f.txt.zip"
NOMBRE_O = "Tasa de mortalidad por siniestros viales por 100.000 habitantes"


def poblacion_dane(clave, anio):
    """Población total de Envigado (05266) del año, leída del xlsx descargado. Devuelve (valor, fila, nombre_archivo)."""
    archivo, hoja, _ = DANE[clave]
    wb = openpyxl.load_workbook(D / "dane" / archivo, read_only=True, data_only=True)
    for i, r in enumerate(wb[hoja].iter_rows(values_only=True), 1):
        if r[2] == "05266" and r[4] == anio and str(r[5]).strip() == "Total":
            wb.close()
            return int(r[6]), i, archivo
    raise RuntimeError(f"{archivo}: no se halló Envigado {anio} Total")


def policia_casos(anio):
    """Suma de `cantidad` del año en el CSV de datos.gov.co (Envigado, homicidio en accidente de tránsito)."""
    tot, n = 0, 0
    with open(RUTA_POLICIA, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if r["cod_muni"] == "05266" and r["fecha_hecho"][:4] == str(anio):
                tot += int(r["cantidad"])
                n += 1
    return tot, n


def terridata_homicidios(anio):
    """Tasa de homicidios en accidentes de tránsito por 100.000 hab. publicada en TerriData."""
    z = zipfile.ZipFile(C.RAIZ / TD_RUTA)
    for l in z.read(z.namelist()[0]).decode("utf-8", "replace").splitlines():
        p = l.split("|")
        if len(p) > 9 and p[6].startswith("Tasa de homicidios en accidentes de tr") and p[9] == str(anio):
            return float(p[7].replace(".", "").replace(",", "."))
    raise RuntimeError("TerriData: no se halló la tasa de homicidios en accidentes de tránsito")


def main():
    meta = PM.extraer_todas()["PDM.CM.M"]
    fo = C.forensis_2024()
    muertes_fo = 18
    cas_pol, n_pol = policia_casos(ANIO)
    td = terridata_homicidios(ANIO)
    pob = {k: poblacion_dane(k, ANIO) for k in DANE}
    # coherencia de lo publicado: 18/248.304 -> 7,25 (Forensis) y 20/260.430 -> 7,7 (TerriData)
    assert abs(100000 * muertes_fo / pob["dane_postcovid_2020_2035"][0] - fo["valor"]) < 0.005
    assert abs(100000 * cas_pol / pob["dane_vp_2018_2042"][0] - td) < 0.05
    real = dt.date(ANIO, 12, 31)
    url_pdm_pop = {k: v[2] for k, v in DANE.items()}
    filas_v, filas_r = [], []
    for esc in C.escenarios("PDM.CM"):
        p = esc["params"]
        fuente, pobp = p["pdm_cm_fuente"], p["poblacion_dane"]
        fc = C.fecha_corte_registro(ANIO, "Anual", real, p)
        M = None if p["horizonte_meta_pdm"] == "metas_anuales" else meta["meta"]
        if fuente == "inmlcf_forensis":
            casos, est_casos = muertes_fo, "Pendiente"
            f_casos = "INMLCF, Forensis 2024: muertes por transporte, Envigado (cifra citada, PDF no verificado)" \
                if not fo["pdf_local"] else "INMLCF, Forensis 2024 (PDF local)"
            pag_casos = "Forensis 2024 p. 37 (tabla 5) y p. 440 (tabla 14)"
            enl_casos, ruta_casos = fo["enlace"], fo["ruta_pdf_local"]
            alerta = "" if fo["pdf_local"] else "cifra citada de Forensis 2024 (7,25; 18 muertes), pendiente de verificar contra el PDF"
            nota_def = ("Definición INMLCF: muertes por transporte (todas las víctimas fatales de siniestros viales).")
        else:
            casos, est_casos = cas_pol, "Pendiente"
            f_casos = "Policía Nacional vía datos.gov.co (uav5-b85g): homicidios en accidente de tránsito, Envigado"
            pag_casos = f"CSV completo, suma de `cantidad` con fecha_hecho en 2024 ({n_pol} registros)"
            enl_casos, ruta_casos = URL_POLICIA, str(RUTA_POLICIA.relative_to(C.RAIZ))
            alerta = "otra definición: homicidio en accidente de tránsito (Policía), no muertes por transporte (INMLCF)"
            nota_def = "Definición Policía: homicidio en accidente de tránsito; no es la misma que la de Forensis (18)."
        # población y tasa
        if pobp == "serie_publicada_por_fuente":
            if fuente == "inmlcf_forensis":
                O, poblacion = fo["valor"], pob["dane_postcovid_2020_2035"][0]
                f_pob = "Población implícita en la tasa publicada por Forensis (248.304 = DANE post-COVID 2024)"
                enl_pob, pag_pob = DANE["dane_postcovid_2020_2035"][2], f"Hoja1, fila {pob['dane_postcovid_2020_2035'][1]} (Total 2024)"
                est_O, fuente_O = fo["estado"], "INMLCF, Forensis 2024, tasa publicada"
                pag_O, enl_O = fo["pagina"], fo["enlace"]
                obs_calc = f"Tasa publicada por la fuente, sin recalcular (18/248.304 x 100.000 = {100000*18/248304:.4f})."
            else:
                O, poblacion = td, pob["dane_vp_2018_2042"][0]
                f_pob = "Población implícita en la tasa de TerriData (260.430 = DANE VP 2024)"
                enl_pob = DANE["dane_vp_2018_2042"][2]
                pag_pob = f"PobMunicipalxÁrea, fila {pob['dane_vp_2018_2042'][1]} (Total 2024)"
                est_O, fuente_O = "Pendiente", "DNP, TerriData: tasa de homicidios en accidentes de tránsito (MinDefensa y DANE)"
                pag_O, enl_O = "TerriData05266f.txt, indicador 'Tasa de homicidios en accidentes de tránsito por cada 100.000 habitantes', 2024", URL_TD
                obs_calc = (f"Tasa publicada por TerriData (7,7) a partir de datos de MinDefensa/Policía; reproduce "
                            f"{cas_pol}/260.430 x 100.000 = {100000*cas_pol/260430:.2f}. Policía no publica tasa propia.")
        else:
            clave = pobp
            poblacion, fila, arch = pob[clave]
            O = 100000 * casos / poblacion
            f_pob = f"DANE, {arch}"
            enl_pob, pag_pob = DANE[clave][2], f"{DANE[clave][1]}, fila {fila} (Total 2024)"
            est_O, fuente_O = est_casos, f_casos + " / " + f_pob
            pag_O, enl_O = f"{pag_casos}; {pag_pob}", enl_casos
            obs_calc = f"Tasa calculada: {casos} / {poblacion} x 100.000 = {O:.4f}."
        contrastes = (f"Contrastes 2024: Forensis 7,25 (18/248.304); DANE VP 18/260.430 = {100000*18/260430:.2f}; "
                      f"Policía {cas_pol} casos: {100000*cas_pol/248304:.2f} (post-COVID) y {100000*cas_pol/260430:.2f} (VP); "
                      "TerriData 7,7. Línea base del plan: 10,2 (2023).")
        comunes = dict(periodo="Anual", territorio="Municipio", fecha_corte=fc, fecha_corte_real=real, provisional=True,
                       script=SCRIPT)
        filas_v.append(C.fila_valor(
            "PDM", "PDM.CM.O", ANIO, O, "Muertes / 100.000 hab.", esc, codigo_indicador="PDM.CM", nombre_variable=NOMBRE_O,
            tipo_valor="observado", estado_informacion=est_O, categoria_disponibilidad="Parcial", fuente=fuente_O,
            documento_ruta_local=ruta_casos if pobp != "serie_publicada_por_fuente" or fuente == "inmlcf_forensis" else TD_RUTA,
            enlace=enl_O, pagina_tabla=pag_O,
            observaciones=f"{nota_def} {obs_calc} {contrastes} Corte: año calendario 2024.", **comunes))
        filas_v.append(C.fila_valor(
            "PDM", "AUX.PDM.CM.MUERTES", ANIO, casos, "Conteo", esc, codigo_indicador="PDM.CM",
            nombre_variable="Muertes viales del año (numerador de la tasa)", tipo_valor="auxiliar",
            estado_informacion=est_casos, categoria_disponibilidad="Parcial", fuente=f_casos,
            documento_ruta_local=ruta_casos, enlace=enl_casos, pagina_tabla=pag_casos,
            observaciones=nota_def + (" Con tasa publicada, el conteo es de apoyo." if pobp == "serie_publicada_por_fuente" and fuente != "inmlcf_forensis" else ""),
            **comunes))
        filas_v.append(C.fila_valor(
            "PDM", "AUX.PDM.CM.POBLACION", ANIO, poblacion, "Habitantes", esc, codigo_indicador="PDM.CM",
            nombre_variable="Población total de Envigado del año (denominador de la tasa)", tipo_valor="auxiliar",
            estado_informacion="Pendiente", categoria_disponibilidad="Disponible", fuente=f_pob, enlace=enl_pob,
            documento_ruta_local="", pagina_tabla=pag_pob,
            observaciones=f"Parámetro poblacion_dane={pobp}.", **comunes))
        filas_v += [f for f in PM.filas_meta("PDM.CM.M", meta, esc)]
        res = C.calcular("PDM.CM", [O, M], {"periodo": "Anual", "territorio": "Municipio"})
        filas_r.append(C.fila_resultado(
            "PDM.CM", ANIO, esc, res, [O, M], nombre="Cumplimiento de la mortalidad vial (CM)", tipo="Indicador",
            fecha_corte_comun=fc, alerta=alerta, script=SCRIPT,
            observaciones=("CM = IF(O=0,100,MIN(100,100*M/O)). " + (f"O={O:.4f} ({fuente}, {pobp}), M={M:g} (PDM p. 706)."
                           if M is not None else res["motivo"]))))
    C.escribir_parcial("PDM", "valores_variables", filas_v, codigos=["PDM.CM"])
    C.escribir_parcial("PDM", "resultados_indicadores", filas_r, codigos=["PDM.CM"])
    for f in filas_r:
        print(f"{f['escenario_id'][:95]:95s} O={f.get('entrada_1')!s:.8} M={f.get('entrada_2')} "
              f"p={f['puntaje']} {f['semaforo']} {f['disponibilidad']}")


if __name__ == "__main__":
    main()
