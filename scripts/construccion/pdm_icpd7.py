#!/usr/bin/env python3
"""Dimension Cumplimiento del PDM (ICPD-7): 7 componentes con peso 1/7 cada uno y brecha de informacion.

ICPD-7 = promedio simple de SALUD, ILE, SP, CM, CA, VEED y HOM (D48). Si falta un componente el anio queda N/D
(no se redistribuyen pesos). Indice calculado solo para 2019-2025 (anios en que pueden coincidir los 7).
Insumos (todos en data_raw_IIC/07_revision_nov2026/, cada carpeta con manifiesto.csv):
  terridata/TerriData05266.txt.zip                 poblacion (dimension 'Descripcion general'), salud, educacion, servicios, tasas
  dane_vp_retroproyeccion/*.xlsx                   poblacion 0-5 de 2015-2017 (DANE VP)
  men_3y4sdmxy/envigado_3y4s-dmxy.csv              atendidos primera infancia 2015-2023
  pdm_balance_educacion_2024/*.pdf                 atendidos 2024 (balance PDM 2024, Educacion)
  informes_gestion/152890_2025informegestiondependencias.pdf   atendidos 2025 (p. 150)
  policia_transito/envigado_filas_ha6j-pa2r.csv    muertes viales 2010-2025 (conteo propio)
  policia_homicidio/envigado_por_anio_m8fd-ahd9.csv homicidios 2003-2025 (conteo propio)
  personeria_jju5br3w/*.csv                        veedurias (solo campos no personales)
  dane_pped_edad/PPED-AreaSexoEdadMun-2018-2042_VP.xlsx   poblacion por edad 2018-2042 (crudo DANE); se contrasta con
  outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv (derivado)
  espacio_verde/                                   VMR AMVA, nota y grafico Humboldt, POT (fuentes del espacio verde)
Salida  : outputs/construccion/PDM/pdm_indicadores.csv (; y UTF-8 con BOM)
Uso     : python3 scripts/construccion/pdm_icpd7.py   (desde la raiz; necesita openpyxl y pypdf; falla si falta alguno)
Todo resultado es PROVISIONAL. Los archivos descargados se tratan solo como datos.
"""
import csv
import io
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC/07_revision_nov2026"
SALIDA = RAIZ / "outputs/construccion/PDM/pdm_indicadores.csv"
PPED = RAIZ / "outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv"

COLS = ["indicador", "codigo", "tipo", "anio", "numerador", "denominador", "valor", "puntaje_0_100", "peso_en_indice",
        "semaforo", "referente", "fuente", "enlace", "fecha_consulta", "estado", "brecha", "nota"]
PROV = "Provisional"
PESO = 1 / 7
UMBRALES = [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")]     # D60; Rojo por debajo
REF_VIAL = 7.3301          # 9,3459 * 8 / 10,2 (D47)
REF_HOM = 2.7              # D34, D58
LB_VEED = 5.4              # D46
LB_CA, REF_CA = 4.54, 9.0  # D31
META_A, META_N = 19.0, 100.44   # D15 (PDM)
PERIODOS = {"SALUD": range(2012, 2026), "ILE": range(2015, 2026), "SP": range(2019, 2026), "CM": range(2003, 2026),
            "CA": range(2019, 2026), "VEED": range(2012, 2026), "HOM": range(2003, 2026)}
PERIODO_INDICE = range(2019, 2026)
ASSERTS = []


def check(nombre, cond, detalle=""):
    ASSERTS.append((nombre, bool(cond), detalle))
    print(("OK   " if cond else "FALLA"), nombre, detalle)


def man(carpeta, archivo):
    with open(RAW / carpeta / "manifiesto.csv", encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";" if ";" in open(RAW / carpeta / "manifiesto.csv", encoding="utf-8-sig").readline() else ","):
            if r["archivo"] == archivo:
                return r["fecha_consulta"], r["url"]
    raise KeyError((carpeta, archivo))


def semaforo(p):
    if p is None:
        return ""
    for m, c in UMBRALES:
        if p >= m:
            return c
    return "Rojo"


def fmt(v, d=2):
    return "" if v is None else f"{round(v, d):.{d}f}"


def fint(v):
    if v is None or isinstance(v, str):
        return v or ""
    return str(int(round(v))) if abs(v - round(v)) < 1e-9 else f"{v:.4f}".rstrip("0")


def pnum(s):
    """Numero de TerriData: '.' de miles y ',' decimal. Vacio = sin dato."""
    s = (s or "").strip()
    return None if s == "" else float(s.replace(".", "").replace(",", "."))


# ------------------------------------------------------------------ TerriData
TD_ARCH = "TerriData05266.txt.zip"
TD_FECHA, TD_URL = man("terridata", TD_ARCH)
with zipfile.ZipFile(RAW / "terridata" / TD_ARCH) as z:
    TD = list(csv.DictReader(io.StringIO(z.read("TerriData05266.txt").decode("utf-8-sig")), delimiter="|"))


def td(dimension, indicador):
    """serie {anio: valor}; con duplicados por mes se toma el mes mayor; celda vacia = sin dato (no entra)."""
    best = {}
    for r in TD:
        if r["Dimensión"] == dimension and r["Indicador"] == indicador:
            a, m = int(r["Año"]), int(r["Mes"])
            if a not in best or m > best[a][0]:
                best[a] = (m, pnum(r["Dato Numérico"]))
    return {a: v for a, (m, v) in best.items() if v is not None}


POB = td("Descripción general", "Población total")        # D61: DANE VP (8-ago-2025)
check("Poblacion VP 2018 = 228.836 y 2024 = 260.430 (D61)", POB[2018] == 228836 and POB[2024] == 260430)

# poblacion 0-5: retroproyeccion VP 2015-2017 y PPED 2018-2025
import openpyxl  # noqa: E402

DANE_ARCH = "DCD-area-sexo-edad-proypoblacion-Mun-2005-2017_VP.xlsx"
DANE_FECHA, DANE_URL = man("dane_vp_retroproyeccion", DANE_ARCH)
POB05 = {}
wb = openpyxl.load_workbook(RAW / "dane_vp_retroproyeccion" / DANE_ARCH, read_only=True, data_only=True)
ws = wb[wb.sheetnames[0]]
enc, vis = None, False
for row in ws.iter_rows(values_only=True):
    if row[0] == "DP" and row[3] == "MPIO":
        enc = {h: i for i, h in enumerate(row) if h}
        continue
    if enc is None:
        continue
    if str(row[enc["MPIO"]]) == "05266":
        vis = True
        if row[enc["ÁREA GEOGRÁFICA"]] == "Total" and 2015 <= int(row[enc["AÑO"]]) <= 2017:
            POB05[int(row[enc["AÑO"]])] = sum(int(row[enc[f"{s}_{e}"]]) for s in ("Hombres", "Mujeres") for e in range(6))
    elif vis:
        break
wb.close()
check("Poblacion 0-5 VP 2015-2017 = 12.177, 12.220, 12.333", [POB05.get(a) for a in (2015, 2016, 2017)] == [12177, 12220, 12333],
      str([POB05.get(a) for a in (2015, 2016, 2017)]))

import re  # noqa: E402

PPED_ARCH = "PPED-AreaSexoEdadMun-2018-2042_VP.xlsx"
PPED_FECHA, PPED_URL = man("dane_pped_edad", PPED_ARCH)
pped_raw, pped_tot = defaultdict(dict), {}
wb = openpyxl.load_workbook(RAW / "dane_pped_edad" / PPED_ARCH, read_only=True, data_only=True)
ws = wb["PobMunicipalxÁreaSexoEdad"]
cols, vis = None, False
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 8:
        cols = defaultdict(list)
        for j_, c in enumerate(row):
            m = re.match(r"(Hombres|Mujeres) (\d+)", str(c or ""))
            if m:
                cols[int(m.group(2))].append(j_)
        continue
    if cols is None or not row or row[2] is None:
        continue
    if str(row[2]) == "05266":
        vis = True
        if row[5] == "Total":
            pped_tot[int(row[4])] = int(row[6])
            for e in range(11):
                pped_raw[int(row[4])][e] = sum(int(row[j_]) for j_ in cols[e])
wb.close()
for a in range(2018, 2026):
    POB05[a] = sum(pped_raw[a][e] for e in range(6))
check("PPED crudo DANE: total 2018-2025 = TerriData 'Descripcion general' (misma serie VP)", all(pped_tot[a] == POB[a] for a in range(2018, 2026)))
pped = defaultdict(dict)
with open(PPED, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        pped[int(r["anio"])][r["edad"]] = int(r["poblacion"])
check("PPED crudo = poblacion_envigado_edad_pped.csv (edades 0-5 y total, 2018-2025)",
      all(pped[a][str(e)] == pped_raw[a][e] for a in range(2018, 2026) for e in range(6)) and all(pped[a]["total"] == pped_tot[a] for a in range(2018, 2026)))

# ------------------------------------------------------------------ filas
FILAS = []


def fila(ind, cod, tipo, anio, n=None, d=None, v=None, puntaje=None, peso="", ref="", fuente="", enlace="", fecha="",
         brecha=None, nota="", usa_puntaje=True, sem=True, nd_motivo=None, dec=2):
    nd = v is None
    if brecha is None:
        brecha = ("sin dato: " + nd_motivo) if nd else "con dato"
    FILAS.append({
        "indicador": ind, "codigo": cod, "tipo": tipo, "anio": anio,
        "numerador": "N/D" if nd else fint(n), "denominador": "N/D" if nd else fint(d),
        "valor": "N/D" if nd else fmt(v, dec),
        "puntaje_0_100": ("N/D" if puntaje is None else fmt(puntaje)) if (usa_puntaje and (nd or puntaje is not None)) else "",
        "peso_en_indice": f"{peso:.6f}" if isinstance(peso, float) else peso,
        "semaforo": semaforo(puntaje) if (sem and puntaje is not None) else "",
        "referente": ref, "fuente": fuente, "enlace": enlace, "fecha_consulta": fecha, "estado": PROV, "brecha": brecha, "nota": nota})
    if nd and usa_puntaje:
        FILAS[-1]["puntaje_0_100"] = "N/D"
    if nd:
        FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = "N/D" if tipo != "respaldo" else ""


COMP = {k: {} for k in PERIODOS}      # puntaje (o None) por componente y anio
BR = {k: {} for k in PERIODOS}        # motivo si N/D
URL_TD = "https://terridata.dnp.gov.co/ ; " + TD_URL
F_TD = lambda dim: f"TerriData (DNP), archivo masivo TerriData05266, dimension {dim}"

# ================================================================== 1. SALUD
dim = "Salud"
CON, SUB, ESP = (td(dim, "Afiliados al régimen contributivo"), td(dim, "Afiliados al régimen subsidado"), td(dim, "Afiliados a regímenes especiales"))
SGS = td(dim, "Afiliados al SGSSS")
for a in range(2012, 2026):
    tot = CON[a] + SUB[a] + ESP[a]
    cob = 100 * tot / POB[a]
    COMP["SALUD"][a] = min(100.0, cob)
    nota = ("COB = 100*(contributivo + subsidiado + especiales)/poblacion DANE VP; puntaje = min(100, COB). Referente 100 % = cobertura universal (D35). "
            f"Desglose: contributivo {fint(CON[a])}, subsidiado {fint(SUB[a])}, especiales {fint(ESP[a])}. El aseguramiento lo reporta MinSalud.")
    if a == 2017:
        nota += " ATIPICO: los afiliados caen frente a 2016 (contributivo 159.955 -> 154.413) y vuelven a subir en 2018."
    if a <= 2017:
        nota += " 'Afiliados al SGSSS' de TerriData excluye los especiales en 2012-2017, por eso no se usa como numerador (D62)."
    fila("Cobertura en salud", "PDM.SALUD", "principal", a, tot, POB[a], cob, COMP["SALUD"][a], PESO, "100 % (cobertura universal; D35)",
         F_TD(dim) + " (afiliados, fuente original MinSalud) + poblacion DANE VP", URL_TD, TD_FECHA, nota=nota)
    # respaldos
    for cod, nom, s in (("PDM.SALUD.CONTRIB", "Afiliados al regimen contributivo", CON), ("PDM.SALUD.SUBSID", "Afiliados al regimen subsidado", SUB),
                        ("PDM.SALUD.ESPEC", "Afiliados a regimenes especiales", ESP)):
        fila(nom, cod, "respaldo", a, s[a], POB[a], 100 * s[a] / POB[a], None, "", "", F_TD(dim), URL_TD, TD_FECHA, usa_puntaje=False, sem=False,
             nota="Entra a la suma del numerador de PDM.SALUD. Valor = % de la poblacion.")
    ctrl = SGS[a]
    if a >= 2018:
        check(f"SGSSS {a} = contributivo + subsidiado + especiales", ctrl == tot, f"{fint(ctrl)} vs {fint(tot)}")
        nt = "Control (D62): 'Afiliados al SGSSS' publicado = suma de los tres regimenes (assert de igualdad)."
    else:
        check(f"SGSSS {a} = contributivo + subsidiado (sin especiales)", ctrl == CON[a] + SUB[a], f"{fint(ctrl)} vs {fint(CON[a] + SUB[a])}")
        nt = "Control (D62): en 2012-2017 'Afiliados al SGSSS' = contributivo + subsidiado; EXCLUYE los especiales, por eso no se usa como numerador."
    fila("Afiliados al SGSSS (publicado por TerriData)", "PDM.SALUD.SGSSS", "respaldo", a, ctrl, POB[a], 100 * ctrl / POB[a], None, "", "",
         F_TD(dim), URL_TD, TD_FECHA, usa_puntaje=False, sem=False, nota=nt)
# ================================================================== 2. ILE
MEN_ARCH = "envigado_3y4s-dmxy.csv"
MEN_FECHA, MEN_URL = man("men_3y4sdmxy", MEN_ARCH)
MEN = {}
with open(RAW / "men_3y4sdmxy" / MEN_ARCH, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f):
        if r["cod_dane_mun"] == "05266":
            MEN[int(r["fecha"][-4:])] = int(r["indicador_3_total_educaci"])
TD_NUM = td("Educación", "Número de niños, niñas y mujeres gestantes con educación inicial en el marco de la atención integral a la primera infancia")
check("MEN 3y4s-dmxy indicador 3 = numerador TerriData 2015-2023", all(MEN[a] == TD_NUM[a] for a in range(2015, 2024)) and len(MEN) == 9)
BAL_URL = ("https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002717/"
           "135807_ppt-balance-pdm-2024_educacion_oficial.pdf")
BAL_ARCH = "135807_ppt-balance-pdm-2024_educacion_oficial.pdf"
BAL_FECHA, _ = man("pdm_balance_educacion_2024", BAL_ARCH)
IG25_ARCH = "152890_2025informegestiondependencias.pdf"
IG25_FECHA, IG25_URL = man("informes_gestion", IG25_ARCH)
import pypdf  # obligatorio: sin pypdf el script falla (no se saltan asserts)

if True:
    rb = pypdf.PdfReader(str(RAW / "pdm_balance_educacion_2024" / BAL_ARCH))
    t24 = rb.pages[23].extract_text()
    check("Balance PDM 2024 Educacion, p. 24: indicador 2.1.1.1.1 logro 2161 (meta 1950)", "2161" in t24 and "1950" in t24 and "2.1.1.1.1" in t24)
    t56 = rb.pages[55].extract_text()
    check("Balance PDM 2024 Educacion, p. 56: 2.1.1.1.1.12, meta 1950, ejecutada 2161", "2161" in t56 and "1950" in t56)
    ri = pypdf.PdfReader(str(RAW / "informes_gestion" / IG25_ARCH))
    t25 = ri.pages[149].extract_text()
    check("Informe de gestion 2025, p. 150 del PDF: 2.092 ninos y ninas <=5 anos", "2.092" in t25 and "5 años" in t25)
ATEND = dict(MEN)
ATEND[2024], ATEND[2025] = 2161, 2092
NETA = td("Educación", "Cobertura neta en educación - Total")
check("Cobertura neta TerriData 2015-2024 disponible; 2025 sin dato", all(a in NETA for a in range(2015, 2025)) and 2025 not in NETA)
URL_EDU = "https://datos.gov.co/resource/3y4s-dmxy.csv ; " + URL_TD
ALERTA_ILE = {
    2015: "ALERTA de comparabilidad (D57): hasta 2018 el numerador es solo ICBF; desde 2019 es ICBF + preescolar MEN. La retroproyeccion VP de la poblacion 0-5 (2015-2017) viene de otro archivo DANE. No afecta el ICPD.",
    2016: "ALERTA de comparabilidad (D57): hasta 2018 el numerador es solo ICBF; desde 2019, ICBF + MEN. 2016 ATIPICO: atendidos 1.479 frente a 726 (2015) y 754 (2017). No afecta el ICPD.",
    2017: "ALERTA de comparabilidad (D57): hasta 2018 el numerador es solo ICBF; desde 2019, ICBF + MEN. La retroproyeccion VP de la poblacion 0-5 (2015-2017) viene de otro archivo DANE. No afecta el ICPD.",
    2018: "ALERTA de comparabilidad (D57): hasta 2018 el numerador es solo ICBF; desde 2019, ICBF + MEN. No afecta el ICPD.",
    2024: "ALERTA de fuente (D57): desde 2024 el numerador sale de la Alcaldia (balance del PDM 2024 de Educacion: 2.161 en el indicador 'Ninos y ninas con acceso a educacion inicial con atencion integral (CDI-HI-Propia)', pp. 24 y 56), no del MEN; es dato de gestion de la Alcaldia, no verificado con la fuente nacional.",
    2025: "ALERTA de fuente (D57): el numerador (2.092) sale del informe de gestion 2025 de la Alcaldia (p. 150 del PDF, formato PA-F-021), no del MEN; es dato de gestion de la Alcaldia, no verificado con la fuente nacional.",
}
for a in range(2015, 2026):
    at, p5 = ATEND[a], POB05[a]
    A = 100 * at / p5
    LA = min(100.0, 100 * A / META_A)
    N = NETA.get(a)
    LN = None if N is None else min(100.0, 100 * N / META_N)
    ILE = None if LN is None else (LA + LN) / 2
    COMP["ILE"][a] = ILE
    if ILE is None:
        BR["ILE"][a] = "TerriData no publica la cobertura neta en educacion de 2025 (LN sin dato); LA si se calcula"
    if a <= 2023:
        fu_at = f"MEN, datos.gov.co 3y4s-dmxy, indicador 3 'Total educacion inicial ICBF + MEN' (= TerriData); fecha consulta MEN {MEN_FECHA}"
        fe = MEN_FECHA if a >= 2018 else f"{MEN_FECHA}; {DANE_FECHA} (poblacion 0-5)"
        enl_at = URL_EDU
    elif a == 2024:
        fu_at, fe, enl_at = "Alcaldia de Envigado, balance PDM 2024 Secretaria de Educacion (atendidos 2024)", f"{BAL_FECHA} (balance); {TD_FECHA} (TerriData)", BAL_URL + " ; " + URL_TD
    else:
        fu_at, fe, enl_at = "Alcaldia de Envigado, informe de gestion 2025, p. 150 (atendidos 2025)", f"{IG25_FECHA} (informe)", IG25_URL
    fu_p = "DANE VP (retroproyeccion 2005-2017)" if a <= 2017 else "DANE VP (PPED 2018-2042, proyeccion por edad)"
    fu = f"{fu_at}; poblacion 0-5: {fu_p}; cobertura neta: TerriData (MinEducacion)"
    enl = enl_at + (" ; " + DANE_URL if a <= 2017 else "")
    nota_base = (f"A = 100*atendidos/poblacion 0-5 (= {fint(at)}/{fint(p5)}); LA = min(100, 100*A/19); LN = min(100, 100*N/100,44) con N = cobertura neta total tal como se publica; "
                 "ILE = (LA + LN)/2 (D15, D16: 0-5 anos aunque el PDM dice 0-6). ")
    fecha_f = fe if a >= 2018 else f"{MEN_FECHA} (MEN); {DANE_FECHA} (DANE); {TD_FECHA} (TerriData)"
    if a >= 2018 and a <= 2023:
        fecha_f = f"{MEN_FECHA} (MEN); {TD_FECHA} (TerriData)"
    al = ALERTA_ILE.get(a, "")
    if a <= 2014:
        pass
    if a == 2015:
        al += (" Ademas la cobertura neta total de TerriData salta de 70,34 (2014) a 95,88 (2015) y la de transicion de 46,10 a 62,28: salto de serie, causa no documentada.")
    if a <= 2023:
        al += " Universo del numerador MEN/TerriData: 'ninos, ninas y mujeres gestantes' (incluye gestantes)."
    else:
        al += " Universo del numerador de la Alcaldia 2024-2025: solo ninos y ninas (sin gestantes); no es estrictamente comparable con 2015-2023."
    fila("Logro educativo (ILE)", "PDM.ILE", "principal", a, None, None, ILE, ILE, PESO, "PDM: 19 % (primera infancia) y 100,44 % (cobertura neta)",
         fu, enl, fecha_f, nd_motivo=BR["ILE"].get(a), nota=nota_base + al + (" ILE N/D (no se sustituye LN)." if ILE is None else ""))
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
    for cod, nom, n_, d_, v_, pt, ref in (("PDM.ILE.A", "Cobertura de primera infancia A (atendidos / poblacion 0-5)", at, p5, A, None, "meta 19 %"),
                                          ("PDM.ILE.LA", "Logro de primera infancia LA", A, META_A, LA, LA, "meta PDM 19 %"),
                                          ("PDM.ILE.LN", "Logro de cobertura neta LN", N, META_N, LN, LN, "meta PDM 100,44 %")):
        mot = None if v_ is not None else "TerriData no publica la cobertura neta de 2025"
        fila(nom, cod, "subcomponente", a, n_, d_, v_, pt, "", ref, fu, enl, fecha_f, usa_puntaje=pt is not None or v_ is None and cod.endswith("LN"),
             sem=False, nd_motivo=mot, nota=(al if cod == "PDM.ILE.A" else "min(100, 100*" + ("A/19" if cod.endswith("LA") else "N/100,44") + ").") if True else "")
    FILAS[-1]["puntaje_0_100"] = "N/D" if LN is None else FILAS[-1]["puntaje_0_100"]

# ================================================================== 3. SP
dim = "Vivienda y acceso a servicios públicos"
ACU, ALC, ASE = td(dim, "Cobertura de acueducto (REC)"), td(dim, "Cobertura de alcantarillado (REC)"), td(dim, "Cobertura de aseo (REC)")
ENE = td(dim, "Cobertura de energia eléctrica total")
ZON = {z: {s_: td(dim, f"Cobertura de {s_} {z} (REC)") for s_ in ("acueducto", "alcantarillado", "aseo")} for z in ("urbana", "rural")}
check("REC rural = 0 en las 3 coberturas 2016-2020", all(ZON["rural"][s_][a] == 0 for s_ in ZON["rural"] for a in range(2016, 2021)))
check("REC urbano 2019-2020 >= 99,77 (acueducto y alcantarillado = 100)", all(ZON["urbana"][s_][a] >= 99.77 for s_ in ZON["urbana"] for a in (2019, 2020)))
_r21 = [ZON["rural"][s_][a] for s_ in ZON["rural"] for a in range(2021, 2025)]
_u21 = [ZON["urbana"][s_][a] for s_ in ZON["urbana"] for a in range(2021, 2025)]
check("REC 2021-2024: rural entre 8 y 61 %, urbano entre 86 y 93 %", 8 <= min(_r21) and max(_r21) <= 61 and 86 <= min(_u21) and max(_u21) <= 93, f"{min(_r21)}-{max(_r21)}; {min(_u21)}-{max(_u21)}")
ALERTA_SP = ("ALERTA de comparabilidad (D65): en 2016-2020 el REC rural es 0 y el urbano 100 en acueducto y alcantarillado (99,77 en aseo), incluido 2019; desde 2021 el rural va de "
             f"{min(_r21):.0f} a {max(_r21):.0f} % y el urbano de {min(_u21):.0f} a {max(_u21):.0f} %. La serie no es comparable entre 2019-2020 y 2021 en adelante, y la Alcaldia reporta el urbano en 100 %. "
             "Se mantiene el valor (D65). Solicitar a Planeacion el REC rural de 2016 a 2020.")
NOTA_SP = ("Fuente REC: la Alcaldia reporta las coberturas al SUI (Superservicios), TerriData las republica; 2017 y 2020 son artefactos de reporte (rural = 0). "
           "El 100 % de energia es un tope de la UPME (2019-2022 = 100, luego 96,27 y 95,63). Referente 100 % = cobertura universal (D38); "
           "promedio simple de las 4 coberturas (D39). " + ALERTA_SP)
for a in range(2019, 2026):
    vals = [s.get(a) for s in (ACU, ALC, ASE, ENE)]
    if None in vals:
        COMP["SP"][a] = None
        BR["SP"][a] = "TerriData no publica las coberturas de servicios publicos de 2025 (la web muestra 2025 pero no el archivo masivo)"
        sp = None
    else:
        sp = sum(vals) / 4
        COMP["SP"][a] = min(100.0, sp)
    fila("Servicios publicos domiciliarios (promedio de 4 coberturas)", "PDM.SP", "principal", a, None, None, sp, COMP["SP"][a], PESO, "100 % (cobertura universal; D38)",
         F_TD(dim) + " (acueducto, alcantarillado y aseo REC/SUI; energia UPME)", URL_TD, TD_FECHA, nd_motivo=BR["SP"].get(a),
         nota=NOTA_SP + (f" Acueducto {vals[0]}, alcantarillado {vals[1]}, aseo {vals[2]}, energia {vals[3]}." if sp is not None else ""))
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
    for cod, nom, s in (("PDM.SP.ACU", "Cobertura de acueducto (REC)", ACU), ("PDM.SP.ALC", "Cobertura de alcantarillado (REC)", ALC),
                        ("PDM.SP.ASE", "Cobertura de aseo (REC)", ASE), ("PDM.SP.ENE", "Cobertura de energia electrica total", ENE)):
        fila(nom, cod, "subcomponente", a, None, None, s.get(a), None, "", "100 %", F_TD(dim), URL_TD, TD_FECHA, usa_puntaje=False, sem=False,
             nd_motivo="TerriData no publica 2025" if s.get(a) is None else None, nota=NOTA_SP if cod.endswith("ACU") else "")
        FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
for a in [2009] + list(range(2011, 2019)):     # contexto sin puntaje (D50)
    for cod, nom, s in (("PDM.SP.REC.ACU", "Contexto: cobertura de acueducto (REC)", ACU), ("PDM.SP.REC.ALC", "Contexto: cobertura de alcantarillado (REC)", ALC),
                        ("PDM.SP.REC.ASE", "Contexto: cobertura de aseo (REC)", ASE)):
        fila(nom, cod, "contexto", a, None, None, s.get(a), None, "", "100 %", F_TD(dim), URL_TD, TD_FECHA, usa_puntaje=False, sem=False,
             nd_motivo="celda vacia en TerriData" if s.get(a) is None else None,
             nota="Solo contexto, sin puntaje ni peso (D50): la energia total solo existe desde 2019, por eso el indicador completo empieza en 2019. "
                  "REC reportado por la Alcaldia al SUI; 2017 es artefacto de reporte (rural = 0). 2008 y 2010 vienen vacios en TerriData.")
        FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""

# ================================================================== 4. CM (mortalidad vial)
TRA_ARCH = "envigado_filas_ha6j-pa2r.csv"
TRA_FECHA, TRA_URL = man("policia_transito", TRA_ARCH)
MUERTES = defaultdict(int)
with open(RAW / "policia_transito" / TRA_ARCH, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f):
        if r["codigo_dane"] == "05266000":
            MUERTES[int(r["fecha_hecho"][-4:])] += int(r["cantidad"])
TV = td("Convivencia y seguridad ciudadana", "Tasa de homicidios en accidentes de tránsito por cada 100.000 habitantes")
check("Referente vial 7,3301 = 9,3459*8/10,2 y tasa propia 2019 = 9,35",
      abs(9.3459 * 8 / 10.2 - REF_VIAL) < 5e-5 and abs(100000 * MUERTES[2019] / POB[2019] - 9.3459) < 5e-4, f"{100000 * MUERTES[2019] / POB[2019]:.4f}")
bad = []
for a in range(2003, 2026):
    if a >= 2010:
        n, O = MUERTES[a], 100000 * MUERTES[a] / POB[a]
        if round(O, 1) != round(TV[a], 1):
            bad.append((a, round(O, 2), TV[a]))
        fu = "Policia Nacional, datos.gov.co ha6j-pa2r (homicidio culposo en accidente de transito; conteo propio) / poblacion DANE VP"
        enl = "https://www.datos.gov.co/resource/ha6j-pa2r.csv ; " + TRA_URL
        fecha = f"{TRA_FECHA} (Policia); {TD_FECHA} (TerriData: poblacion y control)"
        nota = (f"O = 100.000*muertes/poblacion VP ({n}/{fint(POB[a])}). Referente {REF_VIAL} = reduccion del PDM (8/10,2) aplicada a la LB Policia 2019 (9,3459) (D47); "
                "CM = min(100, 100*7,3301/O). Control: round(tasa propia, 1) = tasa TerriData.")
    else:
        n, O = None, TV[a]
        fu = "TerriData (DNP), tasa publicada de homicidios en accidentes de transito (fuente original Policia Nacional)"
        enl, fecha = URL_TD, TD_FECHA
        nota = ("2003-2009: el conjunto ha6j-pa2r empieza en 2010, asi que se usa la tasa publicada por TerriData (sin conteo ni denominador propio). "
                f"Referente {REF_VIAL} (D47); CM = min(100, 100*7,3301/O).")
    if a == 2023:
        nota += " ATIPICO: 2023 tiene 43 muertes (16,8 por 100.000), el maximo de la serie."
    pt = min(100.0, 100 * REF_VIAL / O)
    COMP["CM"][a] = pt
    fila("Mortalidad vial", "PDM.CM", "principal", a, n, POB[a] if n is not None else None, O, pt, PESO, f"{REF_VIAL} por 100.000 (D47)", fu, enl, fecha, nota=nota)
    if a < 2010:
        FILAS[-1]["denominador"] = ""
    fila("Tasa publicada de homicidios en accidentes de transito (TerriData)", "PDM.CM.TD", "respaldo", a, None, None, TV[a], None, "", "", F_TD("Convivencia y seguridad ciudadana"),
         URL_TD, TD_FECHA, usa_puntaje=False, sem=False, nota="Control: round(tasa propia, 1) = TerriData en 2010-2025." if a >= 2010 else "Dato usado en CM 2003-2009.")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
check("Mortalidad vial: round(tasa propia,1) = TerriData 2010-2025", not bad, str(bad))

# ================================================================== 5. CA (espacio verde)
CA_VAL = {2019: 4.54, 2022: 4.66, 2024: 5.75}
CA_ARCH = {2019: "VMR-EPV-VALLE-DE-ABURRA.pdf", 2022: "humboldt_nota.html", 2024: "146211_8-dimension-funcional-_-espacio-publico.pdf"}
CA_GRAF = "Instituto-Humboldt-plan.jpg"
CA_MAN = {k: man("espacio_verde", v) for k, v in CA_ARCH.items()}
CA_MAN["graf"] = man("espacio_verde", CA_GRAF)
CA_FU = {
    2019: ("AMVA con la Universidad Nacional (estudio de conectividad ecologica, 2020); Reporte Metropolitano Voluntario ODS 11.7, AMVA y ONU-Habitat (2022), tabla 4, p. 24, y tabla 5, p. 25 (tablas en imagen; verificado visualmente el 18-nov: tabla 4, fila Envigado 2019 = 232.903 hab. urbanos, 92,06 ha, 4,54; la tabla 5 repite 4,54 con el encabezado de columna «2009», un error tipográfico de la fuente)",
           CA_MAN[2019][1]),
    2022: ("AMVA con el Instituto Humboldt (actualizacion del Plan Maestro de Espacios Publicos Verdes Urbanos); nota Humboldt (2023), grafico (JPG archivado, columna 2022, fila Envigado = 4,66); repetido en el diagnostico del POT, tabla 39, pp. 86-87",
           CA_MAN[2022][1] + " ; " + CA_MAN["graf"][1]),
    2024: ("Alcaldia de Envigado, diagnostico de la revision del POT, dimension funcional (espacio publico), tabla 39, pp. 86-87: 1.382.762 m2 / 240.655 hab. urbanos (DANE 2024)",
           CA_MAN[2024][1]),
}
CA_FECHA_A = {a: CA_MAN[a][0] for a in CA_ARCH}
_t = pypdf.PdfReader(str(RAW / "espacio_verde" / CA_ARCH[2024]))
_tx = _t.pages[85].extract_text() + _t.pages[86].extract_text()
check("POT tabla 39 (pp. 86-87): 240.655, 1.382.762 y 5.75", all(x in _tx for x in ("240.655", "1.382.762", "5.75")))
check("Humboldt: nota (HTML) y grafico (JPG) archivados", (RAW / "espacio_verde" / CA_GRAF).stat().st_size > 1000 and (RAW / "espacio_verde" / "humboldt_nota.html").stat().st_size > 1000)
CA_FUENTE_ND = "Sin corte publicado con la metodologia del AMVA; fuentes revisadas y archivadas: VMR AMVA (2019), Humboldt (2022), POT (2024)"
avance = lambda v: max(0.0, min(100.0, 100 * (v - LB_CA) / (REF_CA - LB_CA)))
for a in range(2019, 2026):
    if a in CA_VAL:
        v = CA_VAL[a]
        pt = avance(v)
        COMP["CA"][a] = pt
        nota = ("Avance = 100*(valor - 4,54)/(9 - 4,54), acotado entre 0 y 100; LB 2019 = 4,54 (D31), referente 9 m2/hab (tabla 39 del POT). Espacio verde URBANO (D30). "
                "Son cortes de estudios distintos (productores, poblaciones y delimitacion distintos): no leer como tendencia.")
        if a == 2019:
            nota += " Linea base: avance 0 por construccion. El 4,54 no se reproduce con los datos del propio reporte (92,06 ha / 232.903 hab. = 3,95)."
        fila("Espacio publico verde urbano", "PDM.CA", "principal", a, None, None, v, pt, PESO, "LB 2019 = 4,54; referente 9 m2/hab", CA_FU[a][0], CA_FU[a][1], CA_FECHA_A[a],
             nota=nota, dec=2)
        if a == 2024:
            CA24 = FILAS[-1]
    else:
        COMP["CA"][a] = None
        mot = {2020: "no hay medicion con la metodologia del AMVA (el 5,00 de la Secretaria de Medio Ambiente es autorreporte y se excluye)",
               2021: "no hay medicion publicada con la metodologia del AMVA",
               2023: "no hay medicion con la metodologia del AMVA (el 5,00 de la Secretaria de Medio Ambiente es autorreporte y se excluye)",
               2025: "no hay corte 2025 publicado en las fuentes revisadas (POT, PDM, AMVA)"}[a]
        BR["CA"][a] = mot
        fila("Espacio publico verde urbano", "PDM.CA", "principal", a, None, None, None, None, PESO, "LB 2019 = 4,54; referente 9 m2/hab", CA_FUENTE_ND, CA_MAN[2024][1], CA_FECHA_A[2024], nd_motivo=mot,
             nota="No hay serie anual; se solicitara a Planeacion la serie desde 2019 con la metodologia del AMVA (D31).")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = "" if a in CA_VAL else FILAS[-1]["numerador"]

# ================================================================== 6. VEED
VEE_ARCH = "veedurias_jju5-br3w_campos_no_personales.csv"
VEE_FECHA, VEE_URL = man("personeria_jju5br3w", VEE_ARCH)
vee = []
with open(RAW / "personeria_jju5br3w" / VEE_ARCH, encoding="utf-8-sig", newline="") as f:
    hdr = csv.DictReader(f)
    check("Veedurias: el archivo solo trae campos no personales", set(hdr.fieldnames) == {"n", "a_o_registro_inicial", "numero_veeduria", "fecha_inscripcion", "duracion_fecha_final"})
    for r in hdr:
        vee.append((int(r["numero_veeduria"]), int(r["fecha_inscripcion"][:4]), int(r["a_o_registro_inicial"])))
vee.sort()
LIST = defaultdict(int)
LIST_REG = defaultdict(int)
for _, y, yr in vee:
    LIST[y] += 1
    LIST_REG[yr] += 1
FAL = defaultdict(int)          # D66: consecutivos que faltan DENTRO del anio (entre dos listadas del mismo anio)
SIN_ANIO = []                   # huecos entre veedurias de anios distintos: (desde, hasta, n, anio_previo, anio_siguiente)
for (n1, y1, _), (n2, y2, _) in zip(vee, vee[1:]):
    gap = n2 - n1 - 1
    if gap <= 0:
        continue
    if y1 == y2:
        FAL[y1] += gap
    else:
        SIN_ANIO.append((n1 + 1, n2 - 1, gap, y1, y2))
check("Veedurias 2019-2023 = 2, 4, 11, 4, 6 y promedio 5,4", [LIST[a] for a in range(2019, 2024)] == [2, 4, 11, 4, 6] and abs(sum(LIST[a] for a in range(2019, 2024)) / 5 - LB_VEED) < 1e-9,
      str([LIST[a] for a in range(2019, 2024)]))
check("Veedurias: brecha 2022-2024 = 0 (sin consecutivos faltantes en el anio)", all(FAL[a] == 0 for a in (2022, 2023, 2024)))
# D66: periodo 2019-2025 = consecutivos entre la primera veeduria de 2019 y la ultima de 2025
n_ini = min(n for n, y, _ in vee if y == 2019)
n_fin = max(n for n, y, _ in vee if y == 2025)
list_p = sum(1 for n, y, _ in vee if 2019 <= y <= 2025)
fal_p = (n_fin - n_ini + 1) - list_p
sin_p = sum(g for d, h, g, y1, y2 in SIN_ANIO if d > n_ini and h < n_fin)
check("Veedurias 2019-2025: faltantes / (listadas + faltantes) = 19/52 = 36,5 %", fal_p == 19 and list_p + fal_p == 52 and abs(100 * fal_p / (list_p + fal_p) - 36.5) < 0.05, f"{fal_p}/{list_p + fal_p}")
check("Veedurias: faltantes del periodo = dentro del anio + sin anio atribuible", fal_p == sum(FAL[a] for a in range(2019, 2026)) + sin_p, f"{sum(FAL[a] for a in range(2019, 2026))} + {sin_p}")
n_fin_2018 = max(n for n, y, _ in vee if y <= 2018)
list_pre = sum(1 for n, y, _ in vee if y <= 2018)
span_pre = n_fin_2018 - vee[0][0] + 1
check("Veedurias 2012-2018: 11 listadas de 120 consecutivos", (list_pre, span_pre) == (11, 120), f"{list_pre} de {span_pre}")
LB_ALT = (sum(LIST[a] + FAL[a] for a in range(2019, 2022)) + LIST[2022] + LIST[2023]) / 5
check("Veedurias: LB con faltantes seguros de 2019-2021 (1, 8, 3) = 7,8", [FAL[a] for a in (2019, 2020, 2021)] == [1, 8, 3] and abs(LB_ALT - 7.8) < 1e-9, str(LB_ALT))
URL_VEE = "https://www.datos.gov.co/resource/jju5-br3w.csv ; " + VEE_URL
NOTA_LB = f"LB 5,4 (D46); si se contaran los faltantes seguros de 2019-2021 ({FAL[2019]}, {FAL[2020]} y {FAL[2021]}), la LB seria {LB_ALT:.1f} (D68)."
for a in range(2012, 2026):
    V = LIST.get(a, 0)
    pt = min(100.0, 100 * V / LB_VEED)
    COMP["VEED"][a] = pt
    fa = FAL[a]
    if V + fa == 0:
        br = "con dato; brecha no determinable (0 listadas y 0 faltantes dentro del anio)"
    else:
        br = f"con dato; brecha {100 * fa / (V + fa):.1f} % ({fa} consecutivos faltantes dentro del anio / {V + fa})"
    nota = (f"V = veedurias inscritas por anio de fecha_inscripcion (D64), solo las listadas; puntaje = min(100, 100*V/5,4). "
            f"Brecha del anio = consecutivos que faltan dentro del anio / (listadas + esos faltantes) (D66); los consecutivos van de {vee[0][0]} a {vee[-1][0]} y el listado trae solo {len(vee)}. "
            f"Consecutivos sin anio atribuible (huecos entre veedurias de anios distintos) en toda la serie: {sum(x[2] for x in SIN_ANIO)}; no se cuentan en ningun anio. "
            + NOTA_LB)
    if a <= 2018:
        nota += (f" Listado incompleto (D67): en 2012-2018 hay {list_pre} veedurias listadas de {span_pre} consecutivos; el puntaje refleja el listado publicado, no el total real.")
    if V == 0:
        nota += " Sin veedurias listadas en el anio."
    if a == 2025:
        nota += (f" Los consecutivos 291 y 292 tienen anio de registro 2026 (fecha de inscripcion 2025); si se contara por anio de registro, 2025 = {LIST_REG[2025]}.")
    fila("Veedurias ciudadanas inscritas", "PDM.VEED", "principal", a, None, None, float(V), pt, PESO, "LB = 5,4 (promedio 2019-2023; D46)",
         "Personeria de Envigado, datos.gov.co jju5-br3w (campos no personales)", URL_VEE, VEE_FECHA, brecha=br, nota=nota, dec=0)
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
    fila("Maximo posible de veedurias del anio (listadas + faltantes dentro del anio)", "PDM.VEED.MAX", "respaldo", a, None, None, float(V + fa), None, "", "",
         "Personeria de Envigado, datos.gov.co jju5-br3w", URL_VEE, VEE_FECHA, usa_puntaje=False, sem=False, dec=0,
         nota=f"Solo nota: listadas {V} + faltantes dentro del anio {fa} = {V + fa} (un solo valor, D66). No incluye los consecutivos sin anio atribuible. No se usa para el puntaje (D64).")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
fila("Veedurias ciudadanas inscritas", "PDM.VEED", "principal", 2026, None, None, None, None, PESO, "LB = 5,4", "Personeria de Envigado, datos.gov.co jju5-br3w", URL_VEE, VEE_FECHA,
     nd_motivo="2026 esta en curso: no se califica", nota=f"Hay {LIST.get(2026, 0)} veeduria(s) inscrita(s) en 2026 por fecha de inscripcion hasta la fecha de consulta; fuera del periodo y del indice.")
FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""

# ================================================================== 7. HOM
HOM_ARCH = "envigado_por_anio_m8fd-ahd9.csv"
HOM_FECHA, HOM_URL = man("policia_homicidio", HOM_ARCH)
HM = defaultdict(int)
with open(RAW / "policia_homicidio" / HOM_ARCH, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f):
        HM[int(r["y"])] += int(r["n"])
TH = td("Convivencia y seguridad ciudadana", "Tasa de homicidio intencional por cada 100.000 habitantes")
bad = []
for a in range(2003, 2026):
    n = HM[a]
    O = 100000 * n / POB[a]
    if round(O, 1) != round(TH[a], 1):
        bad.append((a, round(O, 2), TH[a]))
    pt = min(100.0, 100 * REF_HOM / O) if O > 0 else 100.0
    COMP["HOM"][a] = pt
    nota = (f"O = 100.000*(todas las caracterizaciones, incluido el feminicidio)/poblacion VP ({n}/{fint(POB[a])}); puntaje = min(100, 100*2,7/O) (D34, D58). "
            "Control: round(tasa propia, 1) = TerriData. La LB 3,7 del PDM sale de la poblacion post-COVID; con la VP, la LB 2023 = 3,52 (D58).")
    fila("Homicidio", "PDM.HOM", "principal", a, n, POB[a], O, pt, PESO, "2,7 por 100.000 (meta PDM; D34, D58)",
         "Policia Nacional, datos.gov.co m8fd-ahd9 (homicidio intencional o con dolo + feminicidio; conteo propio) / poblacion DANE VP", "https://www.datos.gov.co/resource/m8fd-ahd9.csv ; " + HOM_URL,
         f"{HOM_FECHA} (Policia); {TD_FECHA} (TerriData: poblacion y control)", nota=nota)
    fila("Tasa publicada de homicidio intencional (TerriData)", "PDM.HOM.TD", "respaldo", a, None, None, TH[a], None, "", "", F_TD("Convivencia y seguridad ciudadana"), URL_TD, TD_FECHA,
         usa_puntaje=False, sem=False, nota="Control: round(tasa propia, 1) = TerriData en 2003-2025 (assert).")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
check("Homicidio: round(tasa propia,1) = TerriData 2003-2025 (23 anios)", not bad, str(bad))
check("Referente HOM: LB 2023 con VP = 3,52", abs(100000 * HM[2023] / POB[2023] - 3.52) < 0.005, f"{100000 * HM[2023] / POB[2023]:.3f}")

# poblacion total y 0-5 como respaldo
for a in range(2003, 2026):
    fila("Poblacion total (DANE VP)", "PDM.POB", "respaldo", a, None, None, POB[a], None, "", "", "TerriData, dimension 'Descripcion general', indicador 'Poblacion total' (DANE VP 8-ago-2025)",
         URL_TD, TD_FECHA, usa_puntaje=False, sem=False, dec=0, nota="Poblacion unica de todas las tasas (D61). No se usa el indicador homonimo de la dimension 'Censo 2005 y proyecciones DANE' (proyeccion antigua, 2018 = 238.221).")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
for a in range(2015, 2026):
    fu = "DANE, retroproyeccion VP 2005-2017 (hoja NuevaMpal; MPIO 05266, area Total, edades 0-5, hombres + mujeres)" if a <= 2017 else "DANE VP, PPED 2018-2042 (outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv)"
    fila("Poblacion de 0 a 5 anos (DANE VP)", "PDM.POB05", "respaldo", a, None, None, float(POB05[a]), None, "", "", fu,
         DANE_URL if a <= 2017 else "https://www.dane.gov.co/index.php/estadisticas-por-tema/demografia-y-poblacion/proyecciones-de-poblacion",
         DANE_FECHA if a <= 2017 else "2026-11-15 (archivo derivado del PPED; sin manifiesto propio)", usa_puntaje=False, sem=False, dec=0,
         nota="Denominador de ILE.A (D16: 0-5 anos).")
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""

# ================================================================== ICPD-7 e brechas
NOMBRES = {"SALUD": "Salud", "ILE": "ILE", "SP": "Servicios publicos", "CM": "Mortalidad vial", "CA": "Espacio verde", "VEED": "Veedurias", "HOM": "Homicidio"}
ICPD = {}
for a in PERIODO_INDICE:
    vals = {k: COMP[k].get(a) for k in COMP}
    faltan = [k for k, v in vals.items() if v is None]
    if faltan:
        ICPD[a] = None
        mot = "faltan: " + "; ".join(f"{NOMBRES[k]} ({BR[k].get(a, 'N/D')})" for k in faltan)
        fila("ICPD-7, indice de cumplimiento del PDM", "PDM.ICPD7", "indice", a, None, None, None, None, "", "Promedio simple de 7 componentes (D48)", "Calculo propio a partir de los 7 componentes",
             "", "2026-11-18", nd_motivo=mot, nota="No se redistribuyen pesos: el indice solo se calcula con los 7 componentes.")
    else:
        ICPD[a] = sum(vals.values()) / 7
        fila("ICPD-7, indice de cumplimiento del PDM", "PDM.ICPD7", "indice", a, None, None, ICPD[a], ICPD[a], "", "Promedio simple de 7 componentes (D48); semaforo D60",
             "Calculo propio a partir de los 7 componentes", "", "2026-11-18",
             nota="Puntajes: " + "; ".join(f"{NOMBRES[k]} {v:.2f}" for k, v in vals.items()) + ". Cada componente pesa 1/7; valores sin redondear. "
                  "Los componentes usan fuentes y poblaciones distintas (ver cada fila); CA son cortes de estudios distintos, no tendencia."
                  + (" " + ALERTA_SP.replace("ALERTA de comparabilidad (D65)", "ALERTA de comparabilidad de servicios publicos (D65)") if a == 2019 else ""))
    FILAS[-1]["numerador"] = FILAS[-1]["denominador"] = ""
check("ICPD-7 2019 = 62,41", ICPD[2019] is not None and abs(ICPD[2019] - 62.41) <= 0.01, fmt(ICPD[2019]))
check("ICPD-7 2022 = 68,79", ICPD[2022] is not None and abs(ICPD[2022] - 68.79) <= 0.01, fmt(ICPD[2022]))
check("ICPD-7 2024 = 66,85", ICPD[2024] is not None and abs(ICPD[2024] - 66.85) <= 0.01, fmt(ICPD[2024]))
sin = [a for a in PERIODO_INDICE if ICPD[a] is None]
check("Brecha del indice = 4/7 (2020, 2021, 2023, 2025)", sin == [2020, 2021, 2023, 2025], str(sin))
if ICPD[2024] is not None:
    SENS = (sum(COMP[k][2024] for k in COMP if k != "CA") + avance(5.54)) / 7
    check("Sensibilidad ICPD 2024 con espacio verde 5,54 = 66,17", abs(SENS - 66.17) < 0.005, f"{SENS:.4f}")
    CA24["nota"] += (f" Sensibilidad (D59): con la poblacion VP el valor seria 5,54 (avance {avance(5.54):.2f}) y el ICPD 2024 = {SENS:.2f}; se usa el valor publicado 5,75.")

for k, rng in PERIODOS.items():
    if k == "VEED":
        fila("Brecha de informacion de Veedurias (2019-2025)", "PDM.VEED.BRECHA", "brecha", "2019-2025", fal_p, list_p + fal_p, 100 * fal_p / (list_p + fal_p), None, "", "", "Personeria de Envigado, datos.gov.co jju5-br3w",
             URL_VEE, VEE_FECHA, usa_puntaje=False, sem=False, brecha=f"{fal_p} de {list_p + fal_p} consecutivos faltantes (36,5 %)",
             nota=(f"D66: proporcion de consecutivos que faltan entre el {n_ini} y el {n_fin} (periodo 2019-2025) = {fal_p}/{list_p + fal_p}; de los {fal_p}, {fal_p - sin_p} caen dentro de un anio "
                   f"y {sin_p} no tienen anio atribuible (huecos entre anios distintos). Reemplaza a '0 de 14 anios sin dato': todos los anios tienen algun valor, pero el listado esta incompleto. "
                   f"Consecutivos sin anio atribuible en toda la serie 2012-2026: {sum(x[2] for x in SIN_ANIO)}. 2012-2018: {list_pre} listadas de {span_pre} consecutivos (D67)."))
        continue
    con = [a for a in rng if COMP[k].get(a) is not None]
    sn = [a for a in rng if COMP[k].get(a) is None]
    fila(f"Brecha de informacion de {NOMBRES[k]} ({rng[0]}-{rng[-1]})", f"PDM.{k}.BRECHA", "brecha", f"{rng[0]}-{rng[-1]}", len(sn), len(rng), 100 * len(sn) / len(rng), None, "", "", "", "", "2026-11-18",
         usa_puntaje=False, sem=False, brecha=f"{len(sn)} de {len(rng)} anios sin dato",
         nota="Anios con dato: " + (f"{con[0]}-{con[-1]}" if con == list(range(con[0], con[-1] + 1)) else ", ".join(map(str, con))) + "; sin dato: "
              + (", ".join(f"{a} ({BR[k].get(a)})" for a in sn) or "ninguno") + ". Brecha = anios sin dato / anios del periodo.")
fila("Brecha de informacion del ICPD-7 (2019-2025)", "PDM.ICPD7.BRECHA", "brecha", "2019-2025", len(sin), 7, 100 * len(sin) / 7, None, "", "", "", "", "2026-11-18", usa_puntaje=False, sem=False,
     brecha=f"{len(sin)} de 7 anios sin indice",
     nota="Anios con indice: " + ", ".join(str(a) for a in PERIODO_INDICE if ICPD[a] is not None) + "; sin indice: 2020, 2021 y 2023 por espacio verde; 2025 por ILE (sin cobertura neta), servicios publicos (sin coberturas) y espacio verde. "
          "Brecha = anios sin indice / 7.")

# ------------------------------------------------------------------ salida
SALIDA.parent.mkdir(parents=True, exist_ok=True)
with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
    w.writeheader()
    w.writerows(FILAS)
print("\nPuntajes (2019, 2022, 2024):")
for k in COMP:
    print(k, [fmt(COMP[k].get(a)) if COMP[k].get(a) is not None else "N/D" for a in (2019, 2022, 2024)])
print("ICPD-7:", {a: (fmt(v) + " " + semaforo(v) if v is not None else "N/D") for a, v in ICPD.items()})
fallos = [n for n, ok, _ in ASSERTS if not ok]
print("filas:", len(FILAS), "->", SALIDA, "| asserts:", len(ASSERTS), "| fallos:", len(fallos))
if fallos:
    print("ASSERTS FALLIDOS:", fallos)
    sys.exit(1)
