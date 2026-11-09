"""PDM, indicadores NO viables por ahora (CS, CP, CD, ILE): solo se construyen variables; el indicador queda N/D.

PDM.CS : auxiliares de MinSalud (dic-2025, ago-2026), proxies por pdm_cs_fuente; PDM.CS.O "No disponible" (método del plan).
PDM.CP : meta + AUX.PDM.CP.INDICE_SITUACIONAL (3,23, 2022); O "No disponible".
PDM.CD : AUX.PDM.CD.DISCIPLINAS (38 de 40, 2025); O "No disponible" (la Matriz prohíbe convertir oferta en cobertura).
PDM.ILE: ILE.A construida = 100 * atendidos / población DANE 0-6 (o 0-5) del año, con calibración 2023 contra 16,3;
         AUX.PDM.ILE.S.IE_MEJORAN; ILE.S "No disponible" (sin clasificación ICFES por plantel).
Los resultados de los cuatro indicadores se escriben N/D (Gris), sin puntaje en ningún escenario.
Ejecutar DESPUÉS de pdm_metas.py: se regeneran aquí las filas de metas de estos indicadores (misma lógica) porque
escribir_parcial reemplaza por código de indicador.
Uso: python3 scripts/indicadores/pdm_no_viables.py   (sin -I, desde la raíz del repo)
"""
import datetime as dt
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C
import pdm_metas as PM
import pdm_cm as CM
import openpyxl

SCRIPT = "scripts/indicadores/pdm_no_viables.py"
D = C.DIR_DESCARGAS
RUTA_BALANCE = ("data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/"
                "ppt-balance-pdm-2024_educacion_oficial.pdf")
RUTA_G25 = "data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf"
MS = {
    2025: ("cifras-afiliacion-salud-dic-2025.zip", "cifras-afiliacion-salud-dic-2025/CIFRAS DICIEMBRE-2025.xlsx",
           dt.date(2025, 12, 31), "diciembre de 2025",
           "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip"),
    2026: ("cifras-afiliacion-salud-agosto2026.zip", "CIFRAS AGOSTO-2026.xlsx", dt.date(2026, 8, 31), "agosto de 2026",
           "https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/RBC/cifras-afiliacion-salud-agosto2026.zip"),
}
EDAD_XLSX = "DCD-area-sexo-edad-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx"
URL_EDAD = ("https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/" + EDAD_XLSX)
EDAD_VP = D / "dane" / "PPED-AreaSexoEdadMun-2018-2042_VP.xlsx"   # opcional (3b del plan); no descargado
NOMBRES_O = {"PDM.CS.O": "Cobertura del aseguramiento en el régimen subsidiado (método del plan)",
             "PDM.CP.O": "Porcentaje de implementación de los procesos de participación ciudadana",
             "PDM.CD.O": "Cobertura porcentual de servicios deportivos",
             "PDM.ILE.S": "Instituciones educativas oficiales que mejoran la categoría Saber 11 (porcentaje)"}


def leer_minsalud():
    """Fila ENVIGADO de la hoja CoberturaMunicipio de cada corte: dict[anio] -> valores y fila de hoja."""
    out = {}
    for anio, (zipn, interno, corte, _, _) in MS.items():
        z = zipfile.ZipFile(D / "minsalud" / zipn)
        wb = openpyxl.load_workbook(io.BytesIO(z.read(interno)), read_only=True, data_only=True)
        ws = wb["CoberturaMunicipio"]
        enc = None
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            if i == 2:
                enc = [str(c).strip() if c else "" for c in r]
            if any(isinstance(c, str) and c.strip().upper() == "ENVIGADO" for c in r):
                d = dict(zip(enc, r))
                out[anio] = dict(fila=i, contrib=d["Contributivo"], subs=d["Subsidiado"], exc=d["Excepción & Especiales"],
                                 afil=d["Afiliados"], pob=d["Población"], cob=d["Cobertura"])
                break
        wb.close()
        assert out[anio]["contrib"] + out[anio]["subs"] + out[anio]["exc"] == out[anio]["afil"], "MinSalud: suma de regímenes"
    return out


def leer_edad_dane():
    """Población total de Envigado por edad simple 0-6 para 2023-2025 (post-COVID). dict[anio] -> lista de 7 enteros."""
    wb = openpyxl.load_workbook(D / "dane" / EDAD_XLSX, read_only=True, data_only=True)
    ws = wb.active
    enc, out = None, {}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 9:
            enc = {n: j for j, n in enumerate(r) if n}
        if r[2] == "05266" and str(r[5]).strip() == "Total" and r[4] in (2023, 2024, 2025):
            tot = [r[enc[f"Total_{e}"]] for e in range(7)]
            hm = [(r[enc[f"Hombres_{e}"]] or 0) + (r[enc[f"Mujeres_{e}"]] or 0) for e in range(7)]
            assert tot == hm, f"DANE edad {r[4]}: Total_e distinto de Hombres+Mujeres"
            out[r[4]] = dict(edades=tot, fila=i)
        if i > 2236:
            break
    wb.close()
    assert set(out) == {2023, 2024, 2025}
    return out


def terridata(indicador, anio):
    z = zipfile.ZipFile(C.RAIZ / CM.TD_RUTA)
    for l in z.read(z.namelist()[0]).decode("utf-8", "replace").splitlines():
        p = l.split("|")
        if len(p) > 9 and p[6] == indicador and p[9] == str(anio):
            return float(p[7].replace(".", "").replace(",", "."))
    raise RuntimeError(f"TerriData: sin {indicador} {anio}")


def v(esc, idv, anio, valor, unidad, ind, nombre, tipo, est, cat, fuente, enlace, pagina, real, obs, doc="", periodo="Anual",
      fc=None):
    """Fila de valores_variables."""
    fc = fc or C.fecha_corte_registro(anio, "Anual", real, esc["params"])
    return C.fila_valor("PDM", idv, anio, valor, unidad, esc, codigo_indicador=ind, nombre_variable=nombre, tipo_valor=tipo,
                        periodo=periodo, territorio="Municipio", estado_informacion=est, categoria_disponibilidad=cat,
                        fuente=fuente, documento_ruta_local=doc, enlace=enlace, pagina_tabla=pagina, fecha_corte=fc,
                        fecha_corte_real=real, provisional=True, script=SCRIPT, observaciones=obs)


def resultado(ind, nombre, anio, esc, motivo):
    res = {"puntaje": None, "semaforo": "Gris", "disponibilidad": "N/D", "motivo": motivo}
    return C.fila_resultado(ind, anio, esc, res, None, nombre=nombre, tipo="Indicador",
                            fecha_corte_comun=C.fecha_corte_registro(anio, "Anual", dt.date(anio, 12, 31), esc["params"]),
                            alerta="indicador no viable por ahora: solo variables, sin puntaje", script=SCRIPT,
                            no_calculado_por_etapa=True, observaciones=motivo)


def main():
    metas = PM.extraer_todas()
    ms = leer_minsalud()
    edad = leer_edad_dane()
    filas_v, filas_r = [], []

    # ------------------------------------------------------------------ PDM.CS
    motivo_cs = ("No viable por ahora: la cobertura del régimen subsidiado con el método del plan (BDUA-ADRES SisMaster, "
                 "población objetivo; 99,5 en 2022) no está publicada. MinSalud publica afiliados sobre población total "
                 "(~12 %), que es otra definición; la cobertura total del SGSSS la prohíbe la Matriz.")
    prox = {2025: 29794 / 250767 * 100}
    for esc in C.escenarios("PDM.CS"):
        p = esc["params"]
        for anio, m in MS.items():
            zipn, interno, real, texto, url = m[0], m[1], m[2], m[3], m[4]
            d = ms[anio]
            comun = dict(real=real, fc=real, doc=f"data_raw_IIC/06_descargas_indicadores/minsalud/{zipn}", enlace=url,
                         fuente=f"MinSalud, Cifras de aseguramiento en salud, {texto}",
                         pagina=f"Hoja CoberturaMunicipio, fila {d['fila']} (ENVIGADO) de {interno.split('/')[-1]}")
            nota_pob = ("; el denominador cambió con la actualización de proyecciones DANE (jul-2025): las coberturas de "
                        "diciembre 2025 y agosto 2026 no son comparables" if anio == 2026 else "")
            for idv, nom, val, uni, obs in [
                ("AUX.PDM.CS.AFIL_RS", "Afiliados al régimen subsidiado, Envigado", d["subs"], "Conteo", ""),
                ("AUX.PDM.CS.AFIL_TOTAL", "Afiliados totales al SGSSS, Envigado", d["afil"], "Conteo",
                 "Contributivo + subsidiado + excepción y especiales (suma verificada)."),
                ("AUX.PDM.CS.POB_MINSALUD", "Población de referencia de MinSalud, Envigado", d["pob"], "Habitantes",
                 "Proyección DANE usada por MinSalud en ese corte" + nota_pob + "."),
                ("AUX.PDM.CS.COB_TOTAL_SGSSS", "Cobertura total del SGSSS publicada por MinSalud", d["cob"] * 100,
                 "Porcentaje (0–100)", "La Matriz prohíbe usar la cobertura total como cobertura del régimen subsidiado."),
                ("AUX.PDM.CS.RS_SOBRE_POB_MINSALUD", "Afiliados RS sobre población total de MinSalud (calculada)",
                 d["subs"] / d["pob"] * 100, "Porcentaje (0–100)",
                 "Otra definición: no es la cobertura del plan (99,5), es el peso del régimen subsidiado en la población total."),
            ]:
                filas_v.append(v(esc, idv, anio, val, uni, "PDM.CS", nom, "auxiliar", "Pendiente",
                                 "Parcial", obs=obs + f" Corte {texto} (fin de mes).", **comun))
            # proxy en PDM.CS.O según el parámetro (el indicador sigue N/D)
            src = p["pdm_cs_fuente"]
            if src == "minsalud_rs_sobre_pob_total":
                valor, est, tipo = d["subs"] / d["pob"] * 100, "Pendiente", "proxy"
                obs = ("Proxy de otra definición (RS sobre población total, ~12 %); NO es la cobertura del plan y no se "
                       "compara con la meta 99,5. El indicador no se calcula.")
            elif src == "minsalud_cobertura_total":
                valor, est, tipo = d["cob"] * 100, "Pendiente", "proxy"
                obs = "Cobertura total del SGSSS: la Matriz lo prohíbe como PDM.CS.O. Solo variable; el indicador no se calcula."
            else:
                valor, est, tipo = None, "No disponible", "observado"
                obs = "Método del plan sin publicar. Proxies en las variables auxiliares AUX.PDM.CS.*."
            filas_v.append(v(esc, "PDM.CS.O", anio, valor, "Porcentaje (0–100)", "PDM.CS", NOMBRES_O["PDM.CS.O"], tipo, est,
                             "Proxy" if valor is not None else "No disponible", obs=obs, **comun))
        # TerriData RS / población total 2022-2025
        for anio in (2022, 2023, 2024, 2025):
            tdv = terridata("Cobertura del régimen subsidiado", anio)
            filas_v.append(v(esc, "AUX.PDM.CS.RS_SOBRE_POB_TD", anio, tdv, "Porcentaje (0–100)", "PDM.CS",
                             "Cobertura del régimen subsidiado según TerriData (afiliados RS / población total)", "auxiliar",
                             "Pendiente", "Parcial", "DNP, TerriData Envigado (MinSalud)", CM.URL_TD,
                             "TerriData05266f.txt, indicador 'Cobertura del régimen subsidiado', diciembre",
                             dt.date(anio, 12, 31),
                             "En 2025 coincide con 29.794/250.767 (RS sobre población total): no es el 99,5 del plan.",
                             doc=CM.TD_RUTA))
        for anio, val in ((2020, 99.6), (2021, 99.2), (2022, 99.5)):
            filas_v.append(v(esc, "AUX.PDM.CS.COB_RS_DIAGNOSTICO_PDM", anio, val, "Porcentaje (0–100)", "PDM.CS",
                             "Porcentaje de cobertura del régimen subsidiado del diagnóstico del PDM (método del plan)",
                             "auxiliar", "Pendiente", "Parcial",
                             "Plan de Desarrollo 2024-2027, diagnóstico (BDUA-ADRES SisMaster, año 2022)",
                             C.url_de_ruta(PM.RUTA_PDM), "PDF p. 362", dt.date(anio, 12, 31),
                             "Serie del método del plan; sin continuación publicada para 2024-2026.", doc=PM.RUTA_PDM))
        for f in PM.filas_meta("PDM.CS.M", metas["PDM.CS.M"], esc):
            filas_v.append(f)
        filas_v.append(v(esc, "PDM.CS.O", 2024, None, "Porcentaje (0–100)", "PDM.CS", NOMBRES_O["PDM.CS.O"], "observado",
                         "No disponible", "No disponible", "Sin fuente publicada con el método del plan", "", "",
                         dt.date(2024, 12, 31), motivo_cs + " MinSalud solo trae dic-2024 en Tableau (no consultado)."))
        for anio in PM.ANIOS:
            filas_r.append(resultado("PDM.CS", "Cumplimiento del aseguramiento en el régimen subsidiado (CS)", anio, esc, motivo_cs))

    # ------------------------------------------------------------------ PDM.CP
    motivo_cp = ("No viable por ahora: no hay valor ni método oficial publicado del porcentaje de implementación de los "
                 "procesos de participación ciudadana; el PDM trae además un índice situacional distinto (3,23, 2022).")
    for esc in C.escenarios("PDM.CP"):
        for f in PM.filas_meta("PDM.CP.M", metas["PDM.CP.M"], esc):
            filas_v.append(f)
        for anio in PM.ANIOS:
            filas_v.append(v(esc, "PDM.CP.O", anio, None, "Porcentaje (0–100)", "PDM.CP", NOMBRES_O["PDM.CP.O"], "observado",
                             "No disponible", "No disponible", "Sin fuente publicada (informes de gestión, rendición y "
                             "evaluaciones del PDM 2024-2025 sin cifra)", "", "", dt.date(anio, 12, 31), motivo_cp))
            filas_r.append(resultado("PDM.CP", "Cumplimiento de la implementación de procesos de participación (CP)", anio, esc,
                                     motivo_cp))
        filas_v.append(v(esc, "AUX.PDM.CP.INDICE_SITUACIONAL", 2022, 3.23, "Número", "PDM.CP",
                         "Índice de implementación de los procesos de participación ciudadana (indicador situacional)",
                         "auxiliar", "Pendiente", "Parcial", "Plan de Desarrollo 2024-2027, indicadores situacionales",
                         C.url_de_ruta(PM.RUTA_PDM), "PDF p. 561", dt.date(2022, 12, 31),
                         "Unidad 'Número'; distinto del indicador de resultado (porcentaje, línea base 75, p. 742). Sin ficha "
                         "que relacione ambos.", doc=PM.RUTA_PDM, fc=dt.date(2022, 12, 31)))

    # ------------------------------------------------------------------ PDM.CD
    motivo_cd = ("No viable por ahora: solo hay oferta de disciplinas (38 de 40 en 2025), que la Matriz prohíbe convertir en "
                 "cobertura porcentual de servicios deportivos.")
    for esc in C.escenarios("PDM.CD"):
        for f in PM.filas_meta("PDM.CD.M", metas["PDM.CD.M"], esc):
            filas_v.append(f)
        for anio in PM.ANIOS:
            filas_v.append(v(esc, "PDM.CD.O", anio, None, "Porcentaje (0–100)", "PDM.CD", NOMBRES_O["PDM.CD.O"], "observado",
                             "No disponible", "No disponible", "Sin cobertura porcentual publicada", "", "",
                             dt.date(anio, 12, 31), motivo_cd))
            filas_r.append(resultado("PDM.CD", "Cumplimiento de la cobertura de servicios deportivos (CD)", anio, esc, motivo_cd))
        for idv, nom, val in (("AUX.PDM.CD.DISCIPLINAS", "Disciplinas deportivas implementadas por el Índer", 38),
                              ("AUX.PDM.CD.DISCIPLINAS_META", "Disciplinas deportivas previstas (meta plan)", 40)):
            filas_v.append(v(esc, idv, 2025, val, "Conteo", "PDM.CD", nom, "auxiliar", "Pendiente", "Proxy",
                             "Informe de gestión por dependencias 2025 (Alcaldía de Envigado)", C.url_de_ruta(RUTA_G25),
                             "PDF p. 5", dt.date(2025, 12, 31),
                             "Oferta de disciplinas, no cobertura de población: no sustituye a PDM.CD.O.", doc=RUTA_G25))

    # ------------------------------------------------------------------ PDM.ILE
    motivo_ile = ("No viable por ahora: ILE.S (instituciones oficiales que mejoran la categoría Saber 11) no tiene clasificación "
                  "ICFES por plantel ni total de instituciones; la ILE.A solo está construida (Pendiente), no es dato oficial.")
    nume = {2023: (2059, "TerriData: niños, niñas y mujeres gestantes con educación inicial en el marco de la atención integral",
                   CM.URL_TD, "TerriData05266f.txt, 2023", CM.TD_RUTA, "Incluye mujeres gestantes: universo más amplio."),
            2024: (2161, "Balance del PDM 2024, Secretaría de Educación: niños y niñas con acceso a educación inicial con "
                   "atención integral (logro; meta 1950)", C.url_de_ruta(RUTA_BALANCE), "PDF p. 24", RUTA_BALANCE,
                   "Rango de edad del universo no especificado en la fuente."),
            2025: (2092, "Informe de gestión por dependencias 2025: niños y niñas con acceso a educación inicial con atención "
                   "integral (Formato PA-F-021)", C.url_de_ruta(RUTA_G25), "PDF p. 150", RUTA_G25,
                   "La fuente dice 'niños y niñas <=5 años': coincide con el rango 0-5.")}
    vp_ok = EDAD_VP.exists()
    calib = {}
    for esc in C.escenarios("PDM.ILE"):
        p = esc["params"]
        rango, pobp = p["pdm_ile_a_rango_edad"], p["poblacion_dane"]
        n_e = 7 if rango == "0_6" else 6
        usable = pobp != "dane_vp_2018_2042" or vp_ok
        for anio in (2023, 2024, 2025):
            num, fnum, enl, pag, doc, onum = nume[anio]
            real = dt.date(anio, 12, 31)
            if anio != 2023:
                pass
            pop = sum(edad[anio]["edades"][:n_e])
            if usable:
                a = 100 * num / pop
                est_a = "Pendiente"
                obs = (f"Construida: 100 x {num} / población DANE {rango.replace('_', ' a ')} años {anio} ({pop}) = {a:.4f}. "
                       f"{onum} Población: DANE post-COVID por edad simple"
                       + ("; para 'serie_publicada_por_fuente' no hay serie publicada de ILE.A y se usa la post-COVID." if
                          pobp == "serie_publicada_por_fuente" else "."))
            else:
                a, est_a = None, "No disponible"
                obs = ("Escenario dane_vp_2018_2042: la serie VP por edad simple (PPED-AreaSexoEdadMun-2018-2042_VP.xlsx, "
                       "opcional) no se descargó; N/D.")
            if anio == 2023:   # calibración contra la línea base 16,3
                if a is not None:
                    calib[(rango, pobp)] = a
                filas_v.append(v(esc, "AUX.PDM.ILE.A.CALIBRACION_2023", 2023, a, "Porcentaje (0–100)", "PDM.ILE",
                                 "ILE.A reconstruida 2023 para calibrar contra la línea base 16,3", "auxiliar", est_a,
                                 "Construible", fnum, enl, pag, real,
                                 obs + " Línea base del plan 16,3 (PDM p. 650): diferencia "
                                 + (f"{a - 16.3:+.2f} puntos." if a is not None else "no evaluable."), doc=doc))
                filas_v.append(v(esc, "AUX.PDM.ILE.A.ATENDIDOS", 2023, num, "Conteo", "PDM.ILE",
                                 "Niños y niñas con atención integral (numerador de ILE.A)", "auxiliar", "Pendiente",
                                 "Construible", fnum, enl, pag, real, onum, doc=doc))
                continue
            filas_v.append(v(esc, "PDM.ILE.A", anio, a, "Porcentaje (0–100)", "PDM.ILE",
                             "Cobertura de atención integral inicial construida (niños atendidos / población DANE)",
                             "construido", est_a, "Construible", fnum, enl, pag, real, obs, doc=doc))
            filas_v.append(v(esc, "AUX.PDM.ILE.A.ATENDIDOS", anio, num, "Conteo", "PDM.ILE",
                             "Niños y niñas con atención integral (numerador de ILE.A)", "auxiliar", "Pendiente",
                             "Construible", fnum, enl, pag, real, onum, doc=doc))
            filas_v.append(v(esc, "AUX.PDM.ILE.A.POBLACION", anio, pop if usable else None, "Habitantes", "PDM.ILE",
                             f"Población DANE de {rango.replace('_', ' a ')} años de Envigado (denominador de ILE.A)",
                             "auxiliar", "Pendiente" if usable else "No disponible", "Disponible",
                             "DANE, proyecciones de población municipal por área, sexo y edad simple 2020-2035 (post-COVID)",
                             URL_EDAD, f"Hoja1, fila {edad[anio]['fila']} (Total, edades 0-{n_e - 1}); Total_e = Hombres+Mujeres",
                             real, f"Parámetros pdm_ile_a_rango_edad={rango}, poblacion_dane={pobp}.",
                             doc=f"data_raw_IIC/06_descargas_indicadores/dane/{EDAD_XLSX}"))
        for f in PM.filas_meta("PDM.ILE.MA", metas["PDM.ILE.MA"], esc) + PM.filas_meta("PDM.ILE.MS", metas["PDM.ILE.MS"], esc):
            filas_v.append(f)
        for anio in PM.ANIOS:
            real = dt.date(anio, 12, 31)
            filas_v.append(v(esc, "PDM.ILE.S", anio, None, "Porcentaje (0–100)", "PDM.ILE", NOMBRES_O["PDM.ILE.S"], "observado",
                             "No disponible", "No disponible", "Sin clasificación ICFES por plantel", "", "", real,
                             "Sin total de instituciones ni clasificación por plantel; ver AUX.PDM.ILE.S.IE_MEJORAN."))
            if anio == 2024:
                src = ("Balance del PDM 2024, Secretaría de Educación", C.url_de_ruta(RUTA_BALANCE), "PDF p. 101", RUTA_BALANCE,
                       "Narra el aumento de clasificación Saber 11 de la I.E. San Vicente; el mismo balance (p. 26) reporta el "
                       "incentivo con logro 0 (meta 2): inconsistencia. Cota inferior, sin total de I.E.")
            else:
                src = ("Informe de gestión por dependencias 2025", C.url_de_ruta(RUTA_G25), "PDF p. 161", RUTA_G25,
                       "Incentivo a la I.E. San Vicente Alto de las Flores por mejora en la clasificación de planteles "
                       "(indicador PDM 2.1.2.1.2.12). Cota inferior, sin total de I.E.")
            filas_v.append(v(esc, "AUX.PDM.ILE.S.IE_MEJORAN", anio, 1, "Conteo", "PDM.ILE",
                             "Instituciones educativas oficiales que mejoran la categoría Saber 11 (mínimo reportado)",
                             "auxiliar", "Pendiente", "Construible", src[0], src[1], src[2], real, src[4], doc=src[3]))
            filas_r.append(resultado("PDM.ILE", "Cumplimiento de la atención integral inicial y la calidad educativa (ILE)",
                                     anio, esc, motivo_ile))
    C.escribir_parcial("PDM", "valores_variables", filas_v, codigos=["PDM.CS", "PDM.CP", "PDM.CD", "PDM.ILE"])
    C.escribir_parcial("PDM", "resultados_indicadores", filas_r, codigos=["PDM.CS", "PDM.CP", "PDM.CD", "PDM.ILE"])
    print(f"{len(filas_v)} filas de variables; {len(filas_r)} de resultados")
    for d, m in ms.items():
        print("MinSalud", d, {k: m[k] for k in ("fila", "subs", "afil", "pob")}, f"cob={m['cob'] * 100:.2f}",
              f"RS/pob={m['subs'] / m['pob'] * 100:.3f}")
    for (r, pb), a in sorted(calib.items()):
        print(f"Calibración 2023 rango {r} / {pb}: {a:.3f} (línea base 16,3; dif {a - 16.3:+.2f})")
    for anio in (2024, 2025):
        for n_e, r in ((7, "0_6"), (6, "0_5")):
            pop = sum(edad[anio]["edades"][:n_e])
            print(f"ILE.A {anio} {r}: {nume[anio][0]} / {pop} = {100 * nume[anio][0] / pop:.3f}")


if __name__ == "__main__":
    main()
