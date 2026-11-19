#!/usr/bin/env python3
"""Dimension Participacion ciudadana (ISPC): I1 votacion de PP, I2 cobertura territorial, I3 oferta formativa, I4 publicacion de resultados.

ISPC = 0,30*I1 + 0,25*I2 + 0,20*I3 + 0,25*I4 (D42). Si falta un componente el anio queda N/D (no se redistribuyen pesos).
Semaforo D60: >= 90 Verde, >= 75 Amarillo, >= 50 Naranja, < 50 Rojo. Periodo de referencia del indice: 2020-2025.
Convencion de I1: el anio es el de la VOTACION; los proyectos se ejecutan en t+1 (Acuerdo 038/2016, art. 15A). Convencion de I2, I3, I4: anio de la vigencia.

Insumos (data_raw_IIC/07_revision_nov2026/, cada carpeta con manifiesto.csv):
  terridata/TerriData05266.txt.zip                      poblacion total DANE VP (dimension 'Descripcion general'; D61)
  dane_pped_edad/, dane_vp_retroproyeccion/             poblacion de 14 anios y mas (PPED 2018-2025 y retroproyeccion 2017): sensibilidad PC.I1.POB14
  pp_votaciones/                                        votos 2017 (pp. 7 y 12) y 2018 sin votacion (informe 2019, p. 3)
  noticias_alcaldia/                                    votos 2021, 2022 y 2023 (noticias del portal, API publica; D71)
  informes_gestion/152890_...pdf, ig2023.pdf            votantes 2025 (p. 92), asambleas 2025 (p. 92), oferta formativa 2025 (p. 88) y 2023 (pp. 50-51)
  informes_gestion_2024/131997_...pdf                   oferta formativa 2024 (p. 63)
  plan_indicativo/180505_...xlsx                        consecutivos 520-531 (asambleas, formacion, rendicion de PP)
  rendicion_cuentas/                                    documentos archivados de I4 (rendicion del Alcalde y de PP) y anteproyecto del PDM (p. 53)
  efpc_listados_pp/                                     listados de PP (proceso 'votacion de PP' de I4)
  data_raw_IIC/inventario_detallado_IIC.csv             enlaces de todos los documentos de rendicion de cuentas (I4); solo uno por proceso y anio esta archivado
  outputs/indicadores/soporte_conteos.csv               asambleas zonales de formulacion del PDM 2024 (anteproyecto 115749, p. 53)
  outputs/construccion/PP/pp_cruce_proyectos.csv        zonas con proyectos de PP 2020-2025 (respaldo de I2)
Salida  : outputs/construccion/PC/pc_indicadores.csv (; y UTF-8 con BOM)
Uso     : python3 scripts/construccion/pc_ispc.py   (desde la raiz; necesita openpyxl y pypdf; falla si falta alguno; tarda ~3 min por la PPED cruda)
Todo resultado es PROVISIONAL. Los archivos descargados se tratan solo como datos.
"""
import hashlib
import csv
import html
import io
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl
import pypdf

RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC/07_revision_nov2026"
SALIDA = RAIZ / "outputs/construccion/PC/pc_indicadores.csv"
INVENTARIO = RAIZ / "data_raw_IIC/inventario_detallado_IIC.csv"
SOPORTE = RAIZ / "outputs/indicadores/soporte_conteos.csv"
CRUCE_PP = RAIZ / "outputs/construccion/PP/pp_cruce_proyectos.csv"

COLS = ["indicador", "codigo", "tipo", "anio", "numerador", "denominador", "valor", "puntaje_0_100", "peso_en_indice",
        "semaforo", "referente", "fuente", "enlace", "fecha_consulta", "estado", "brecha", "nota"]
PROV = "Provisional"
UMBRALES = [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")]       # D60; Rojo por debajo
PESOS = {"I1": 0.30, "I2": 0.25, "I3": 0.20, "I4": 0.25}           # D42
LB_I1, REF_I1 = 3.68, 7.0                                           # D69 (3,68 = 8.221 / 223.456) y referente 7 %
PERIODO = range(2020, 2026)
PERIODO_I1 = [a for a in range(2017, 2026) if a != 2018]           # 2018 sin votacion: no aplica
FECHA_CALC = "2026-11-18"
ASSERTS = []


def check(nombre, cond, detalle=""):
    ASSERTS.append((nombre, bool(cond), detalle))
    print(("OK   " if cond else "FALLA"), nombre, detalle)


def man(carpeta, archivo):
    """(fecha_consulta, url, nota) del manifiesto de la carpeta (separador ; o ,)."""
    p = RAW / carpeta / "manifiesto.csv"
    with open(p, encoding="utf-8-sig", newline="") as f:
        sep = ";" if ";" in f.readline() else ","
    with open(p, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=sep):
            if r["archivo"] == archivo:
                return r["fecha_consulta"], r["url"], r.get("nota", "")
    raise KeyError((carpeta, archivo))


def semaforo(p):
    if p is None:
        return ""
    for m, c in UMBRALES:
        if p >= m:
            return c
    return "Rojo"


def fmt(v, d=2):
    return f"{round(v, d):.{d}f}"


def fint(v):
    return str(int(round(v))) if abs(v - round(v)) < 1e-9 else f"{v:.4f}".rstrip("0")


def norm(s):
    return re.sub(r"\s+", " ", s)


def pdf_pag(ruta, pag):
    """texto normalizado de la pagina `pag` (1 = primera) de un PDF (solo dato)."""
    return norm(pypdf.PdfReader(str(ruta)).pages[pag - 1].extract_text())


def pnum(s):
    s = (s or "").strip()
    return None if s == "" else float(s.replace(".", "").replace(",", "."))


# ------------------------------------------------------------------ filas
FILAS = []


def fila(ind, cod, tipo, anio, n=None, d=None, v=None, puntaje=None, peso="", ref="", fuente="", enlace="", fecha="",
         brecha="", nota="", dec=2, raw=None):
    """raw: dict para forzar textos ('numerador', 'denominador', 'valor', 'puntaje') p. ej. 'N/D', 'no aplica', 'pendiente'."""
    raw = raw or {}
    FILAS.append({
        "indicador": ind, "codigo": cod, "tipo": tipo, "anio": anio,
        "numerador": raw.get("numerador", "" if n is None else fint(n)),
        "denominador": raw.get("denominador", "" if d is None else fint(d)),
        "valor": raw.get("valor", "" if v is None else fmt(v, dec)),
        "puntaje_0_100": raw.get("puntaje", "" if puntaje is None else fmt(puntaje)),
        "peso_en_indice": f"{peso:.2f}" if isinstance(peso, float) else peso,
        "semaforo": semaforo(puntaje) if puntaje is not None else "",
        "referente": ref, "fuente": fuente, "enlace": enlace, "fecha_consulta": fecha, "estado": PROV, "brecha": brecha, "nota": nota})


def fila_nd(ind, cod, anio, peso, ref, motivo, fuente="", enlace="", fecha="", nota="", n_d=True):
    fila(ind, cod, "principal", anio, peso=peso, ref=ref, fuente=fuente, enlace=enlace, fecha=fecha, brecha="sin dato: " + motivo, nota=nota,
         raw={"numerador": "N/D", "denominador": "N/D", "valor": "N/D", "puntaje": "N/D"})


# ================================================================== fuentes archivadas en rendicion_cuentas/ (I4 y anteproyecto del PDM)
RC = {}
with open(RAW / "rendicion_cuentas" / "manifiesto.csv", encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f):
        RC[r["archivo"]] = (r["fecha_consulta"], r["url"], r["nota"])
        h = hashlib.sha256((RAW / "rendicion_cuentas" / r["archivo"]).read_bytes()).hexdigest()
        if h != r["sha256"] or (RAW / "rendicion_cuentas" / r["archivo"]).stat().st_size != int(r["bytes"]):
            raise SystemExit("sha256 o bytes no coinciden con el manifiesto: " + r["archivo"])
RC_FECHA = RC["129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf"][0]


def rc_arch(idd):
    """nombre del archivo archivado cuyo id numerico es idd (falla si no esta)."""
    for n in RC:
        if n.startswith(f"{idd}_"):
            return n
    raise KeyError(idd)


def rc_pdf(idd, pag):
    return pdf_pag(RAW / "rendicion_cuentas" / rc_arch(idd), pag)


check("Archivo rendicion_cuentas: 12 documentos con sha256 y bytes iguales al manifiesto", len(RC) == 12)
check("Doc 6385 (pagina 'Rendicion de Cuentas 2020'): es de la Secretaria de Salud, no del Alcalde", "Secretaría de Salud de Envigado Rendición de cuentas 2020" in rc_pdf(6385, 2))
check("Doc 92036 (rendicion del Alcalde 2022), p. 8: 'Respuestas a preguntas'", "Respuestas a preguntas" in rc_pdf(92036, 8))
check("Anteproyecto PDM (115749), p. 53: asambleas entre el 29 de enero y el 17 de febrero", "29 de enero y el 17 de febrero" in rc_pdf(115749, 53) and "asambleas zonales" in rc_pdf(115749, 53))
check("Doc 23449 (2021), 115926 (2023), 129738 (2024) y 149610 (2025): ejercicio/respuestas de rendicion de cuentas del Alcalde",
      "RENDICIÓN DE CUENTAS DEL SEÑOR ALCALDE" in rc_pdf(23449, 1) and "Rendición de cuentas II semestre 2023" in rc_pdf(115926, 1)
      and "Rendición de Cuentas 2024" in rc_pdf(129738, 1) and "Rendición de Cuentas 2025" in rc_pdf(149610, 1))

# ================================================================== poblacion (TerriData, DANE VP)
TD_ARCH = "TerriData05266.txt.zip"
TD_FECHA, TD_URL, _ = man("terridata", TD_ARCH)
with zipfile.ZipFile(RAW / "terridata" / TD_ARCH) as z:
    TD = list(csv.DictReader(io.StringIO(z.read("TerriData05266.txt").decode("utf-8-sig")), delimiter="|"))
POB = {}
_mes = {}
for r in TD:
    if r["Dimensión"] == "Descripción general" and r["Indicador"] == "Población total":
        a, m, v = int(r["Año"]), int(r["Mes"]), pnum(r["Dato Numérico"])
        if v is not None and (a not in _mes or m > _mes[a]):
            _mes[a], POB[a] = m, v
check("Poblacion VP 2017 / 2021 / 2025 = 223.456 / 247.649 / 264.637 (D61)", (POB[2017], POB[2021], POB[2025]) == (223456, 247649, 264637),
      str((POB[2017], POB[2021], POB[2025])))
URL_TD = "https://terridata.dnp.gov.co/ ; " + TD_URL
F_POB = "poblacion: TerriData (DNP), dimension 'Descripcion general', 'Poblacion total' (DANE VP, 8-ago-2025)"

# ================================================================== I1: votos de PP por anio de votacion
PP_VOT = RAW / "pp_votaciones"
F17, U17, _ = man("pp_votaciones", "14376_informepriorizacionparticipativadelpresupuesto11deoctubrede2017.pdf")
F19, U19, _ = man("pp_votaciones", "5254_informe-ejecutivo-auditoria-priorizacion-participativa-del-presupuesto-2019.pdf")
A17 = PP_VOT / "14376_informepriorizacionparticipativadelpresupuesto11deoctubrede2017.pdf"
A19 = PP_VOT / "5254_informe-ejecutivo-auditoria-priorizacion-participativa-del-presupuesto-2019.pdf"
t7, t12 = pdf_pag(A17, 7), pdf_pag(A17, 12)
check("Votos 2017 = 8.221: aparece en el informe de priorizacion 2017, p. 7 ('8221 votos', IPCM 3,62%) y p. 12 ('8.221 votos')",
      "8221 votos" in t7 and "3,62%" in t7 and "8.221 votos" in t12)
t3 = pdf_pag(A19, 3)
check("2018 sin votacion (vigencia 2019): informe de auditoria 2019, p. 3 'no se sometieron a votacion'", "no se sometieron a votaci" in t3 and "2019" in t3)

NOT = {}
for anio, arch, esperado in ((2021, "api_2021_16828votos.json", "16.828"), (2022, "api_2022_40proyectos.json", "12.541"),
                             (2023, "api_9156-envigadenos-votaron-por-los-proyectos-de-presupuesto.json", "9.156")):
    f_, u_, nota_ = man("noticias_alcaldia", arch)
    d_ = json.load(open(RAW / "noticias_alcaldia" / arch, encoding="utf-8"))
    cuerpo = norm(re.sub(r"<[^>]+>", " ", html.unescape(d_["body"]))) + " " + d_["name"]   # cuerpo y titulo de la noticia
    web = re.search(r"Página web: (\S+?)\.\s", nota_ + " ").group(1)
    check(f"Votos {anio} = {esperado}: aparece en el titulo o el cuerpo de la noticia ({arch})", esperado in cuerpo and d_["creationDate"][:4] == str(anio),
          d_["creationDate"][:10])
    NOT[anio] = (f_, web, d_["creationDate"][:10], d_["name"])
IG25 = "152890_2025informegestiondependencias.pdf"
F25, U25, _ = man("informes_gestion", IG25)
P25 = RAW / "informes_gestion" / IG25
t92 = pdf_pag(P25, 92)
check("Informe de gestion 2025, p. 92: '10,662 votantes directos'", "10,662 votantes directos" in t92)
VOTOS = {2017: 8221, 2021: 16828, 2022: 12541, 2023: 9156, 2025: 10662}
FUENTE_VOTOS = {
    2017: ("Alcaldia de Envigado, Planeacion: Informe de Priorizacion Participativa del Presupuesto (11-oct-2017), pp. 7 y 12", U17, F17),
    2021: ("Portal de la Alcaldia, noticia del %s: «%s»" % (NOT[2021][2], NOT[2021][3]), NOT[2021][1], NOT[2021][0]),
    2022: ("Portal de la Alcaldia, noticia del %s: «%s»" % (NOT[2022][2], NOT[2022][3]), NOT[2022][1], NOT[2022][0]),
    2023: ("Portal de la Alcaldia, noticia del %s: «%s»" % (NOT[2023][2], NOT[2023][3]), NOT[2023][1], NOT[2023][0]),
    2025: ("Alcaldia de Envigado, informe de gestion por dependencias 2025, p. 92 del PDF (votacion del 11 al 24 de agosto)", U25, F25),
}
ip = lambda v, p: 100.0 * v / p
punt1 = lambda x: min(100.0, max(0.0, 100.0 * (x - LB_I1) / (REF_I1 - LB_I1)))
IPCM = {a: ip(VOTOS[a], POB[a]) for a in VOTOS}
P1 = {a: punt1(IPCM[a]) for a in VOTOS}
SIN_ACOTAR = {a: 100.0 * (IPCM[a] - LB_I1) / (REF_I1 - LB_I1) for a in VOTOS}
check("I1 2023 sin acotar = -3,13 y 2017 sin acotar = -0,03", abs(SIN_ACOTAR[2023] + 3.13) < 0.005 and abs(SIN_ACOTAR[2017] + 0.03) < 0.005, f"{SIN_ACOTAR[2023]:.3f} / {SIN_ACOTAR[2017]:.3f}")
check("IPCM 2017 = 3,679 +/- 0,001 y puntaje 0 (LB = 3,68; D69)", abs(IPCM[2017] - 3.679) <= 0.001 and P1[2017] == 0.0, f"{IPCM[2017]:.4f} / {P1[2017]}")
check("LB 3,68 = 8.221 / 223.456 redondeado a 2 decimales", round(IPCM[2017], 2) == LB_I1)
check("I1 2025 = 100*(4,029 - 3,68)/3,32 = 10,5", abs(P1[2025] - 10.51) < 0.01 and abs(IPCM[2025] - 4.029) < 0.001, f"{IPCM[2025]:.4f} / {P1[2025]:.4f}")

# ---- poblacion de 14 anios y mas (PPED crudo 2018-2025 + retroproyeccion 2017)
PPED_ARCH, DANE_ARCH = "PPED-AreaSexoEdadMun-2018-2042_VP.xlsx", "DCD-area-sexo-edad-proypoblacion-Mun-2005-2017_VP.xlsx"
PPED_FECHA, PPED_URL, _ = man("dane_pped_edad", PPED_ARCH)
DANE_FECHA, DANE_URL, _ = man("dane_vp_retroproyeccion", DANE_ARCH)
POB14, POB_TOT_DANE = {}, {}
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
        if row[enc["ÁREA GEOGRÁFICA"]] == "Total" and int(row[enc["AÑO"]]) == 2017:
            cols14 = [i for h, i in enc.items() if re.match(r"(Hombres|Mujeres)_\d+", str(h)) and int(re.match(r"\w+_(\d+)", h).group(1)) >= 14] + \
                     [i for h, i in enc.items() if re.match(r"(Hombres|Mujeres)_\d+ y m", str(h)) and int(re.match(r"\w+_(\d+)", h).group(1)) >= 14]
            cols14 = sorted(set(cols14))
            POB14[2017] = sum(int(row[i]) for i in cols14)
            POB_TOT_DANE[2017] = int(row[enc["Total General"]])
            menos14 = sum(int(row[enc[f"{s}_{e}"]]) for s in ("Hombres", "Mujeres") for e in range(14))
            check("Retroproyeccion 2017: 14+ = total - (0 a 13 anios)", POB14[2017] == POB_TOT_DANE[2017] - menos14, f"{POB14[2017]}")
    elif vis:
        break
wb.close()
wb = openpyxl.load_workbook(RAW / "dane_pped_edad" / PPED_ARCH, read_only=True, data_only=True)
ws = wb["PobMunicipalxÁreaSexoEdad"]
cols = None
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
    if str(row[2]) == "05266" and row[5] == "Total":
        a = int(row[4])
        if a > 2025:
            continue
        tot = int(row[6])
        POB_TOT_DANE[a] = tot
        directo = sum(int(row[j_]) for e, js in cols.items() if e >= 14 for j_ in js)
        menos = sum(int(row[j_]) for e in range(14) for j_ in cols[e])
        check(f"PPED {a}: 14+ (suma de edades) = total - (0 a 13 anios)", directo == tot - menos)
        POB14[a] = directo
wb.close()
check("Poblacion total DANE (PPED 2018-2025 y retroproyeccion 2017) = TerriData 'Descripcion general' (misma serie VP)",
      all(POB_TOT_DANE[a] == POB[a] for a in range(2017, 2026)))

ALERTA_UNIDAD = ("ALERTA de unidad: 2017 y 2021-2023 son 'votos' (las noticias hablan de votos) y 2025 son 'votantes directos' del informe de gestion; "
                 "pueden votar los mayores de 14 anios, pero el IPCM oficial y la LB usan la poblacion total.")
NOTA_BASE = ("IPCM = votos / poblacion DANE VP del mismo anio x 100 (D61); puntaje = min(100, max(0, 100*(IPCM - 3,68)/(7 - 3,68))). LB 3,68 = 8.221/223.456 con la poblacion DANE VP (D69; "
             "la cifra oficial de 2017 es 3,62: ver PC.I1.OFICIAL); referente 7 % = tope de la escala oficial del IPCM (D32). Convencion: el anio es el de la votacion; "
             "los proyectos se ejecutan en t+1 (Acuerdo 038/2016, art. 15A). ")
REF_TXT = "LB 3,68 (2017, D69) -> referente 7 % (tope de la escala oficial del IPCM; D32)"
NOTAS_I1 = {
    2017: " Linea base: puntaje 0 por construccion. Votacion de 2017 para la vigencia 2018 (informe de 2017, p. 7: 'proyectos a ejecutar en el 2018'). Cifra oficial 3,62 % (Alcaldia, poblacion propia no publicada).",
    2021: " La noticia titula 'vigencia 2021-2022' y dice que los proyectos son 'para el ano 2022'. Con la poblacion de 14 anios y mas el IPCM pasaria de 7 % (ver PC.I1.POB14).",
    2022: " Noticia del 13-sep-2022: proyectos de la vigencia 2023.",
    2023: (" Noticia del 11-sep-2023: proyectos de la vigencia 2024. IPCM por debajo de la LB: el puntaje sin acotar seria %.2f y se acota en 0. El 0 de 2017 es por definicion de la LB "
           "(sin acotar seria %.2f, por redondeo de la LB)." % (SIN_ACOTAR[2023], SIN_ACOTAR[2017])),
    2025: " 10.662 votantes directos (informe de gestion 2025, p. 92); el boletin de la Alcaldia dice 'mas de 12.000 participaron' (= inscritos en la plataforma, 13.649 beneficiarios registrados), no se usa. "
          "Solo 2025 tiene los cuatro componentes del ISPC.",
}
for a in range(2017, 2026):
    if a in VOTOS:
        fu, en, fe = FUENTE_VOTOS[a]
        fila("Participacion en la votacion del presupuesto participativo (IPCM)", "PC.I1", "principal", a, VOTOS[a], POB[a], IPCM[a], P1[a], PESOS["I1"], REF_TXT,
             fu + " ; " + F_POB, en + " ; " + URL_TD, fe + " (votos) ; " + TD_FECHA + " (poblacion)", "con dato",
             NOTA_BASE + ALERTA_UNIDAD + NOTAS_I1.get(a, ""), dec=3)
    elif a == 2018:
        fila("Participacion en la votacion del presupuesto participativo (IPCM)", "PC.I1", "principal", a, peso=PESOS["I1"], ref=REF_TXT,
             fuente="Alcaldia de Envigado, informe ejecutivo de auditoria de la priorizacion participativa 2019, p. 3 (los proyectos de la vigencia 2019 no se sometieron a votacion)",
             enlace=U19, fecha=F19, brecha="no aplica: no hubo votacion (vigencia 2019)",
             nota="La vigencia 2019 se priorizo en 2018 y no hubo votacion porque los recursos alcanzaron para todos los proyectos presentados. No cuenta como brecha ni entra al periodo de I1.",
             raw={"numerador": "no aplica", "denominador": "no aplica", "valor": "no aplica", "puntaje": "no aplica"})
    else:
        extra = {
            2019: "Si hubo votacion en 2019 (vigencia 2020), pero no se encontro cifra publicada de votos.",
            2020: ("Votacion de 2020 (vigencia 2021) sin cifra final publicada. Solo se hallaron 5.100 votos parciales (360 Radio), que no se usan. Indicio no usado: la noticia de 2021 dice que la participacion "
                   "de 2021 subio 69 % frente al ano anterior, lo que implicaria cerca de 9.960 votos en 2020 (calculo propio, no es una cifra publicada)."),
            2024: "Votacion de 2024 (vigencia 2025) sin cifra de votos publicada en las fuentes revisadas.",
        }[a]
        fila_nd("Participacion en la votacion del presupuesto participativo (IPCM)", "PC.I1", a, PESOS["I1"], REF_TXT,
                "sin cifra publicada de votos de la votacion de " + str(a), "Portal de la Alcaldia (API publica, busqueda por palabra clave), informes de gestion, listados EFPC: sin cifra",
                "https://www.envigado.gov.co/ ; " + URL_TD, "2026-11-18", nota=NOTA_BASE + extra + " Solicitar a Planeacion los votos y el IPCM por anio y por zona.")

fila("IPCM oficial de la Alcaldia", "PC.I1.OFICIAL", "respaldo", 2017, VOTOS[2017], None, 3.62, None, "", "Escala oficial del IPCM: tramos 1-2,99 %, 3-4,99 % y 5-6,99 % (informe de 2017, p. 6)",
     FUENTE_VOTOS[2017][0], FUENTE_VOTOS[2017][1], FUENTE_VOTOS[2017][2], "con dato",
     "Cifra oficial 3,62 % (informe de 2017, pp. 7 y 12); la Alcaldia usa su propia poblacion de referencia (no publicada; implicaria cerca de 227.000 hab.). Con la poblacion DANE VP el IPCM es 3,679 "
     "(8.221/223.456), valor que se usa como LB (D69). Solo respaldo: no entra al puntaje.", dec=2)
for a in sorted(set(VOTOS) & set(POB14)):
    p14 = ip(VOTOS[a], POB14[a])
    fu, en, fe = FUENTE_VOTOS[a]
    fila("Sensibilidad: votos / poblacion de 14 anios y mas", "PC.I1.POB14", "respaldo", a, VOTOS[a], POB14[a], p14, None, "", "Referente 7 % del IPCM (poblacion total)",
         fu + " ; poblacion de 14+: DANE, " + ("retroproyeccion VP 2005-2017 (MPIO 05266, area Total)" if a == 2017 else "PPED 2018-2042 VP (MPIO 05266, area Total)"),
         en + " ; " + (DANE_URL if a == 2017 else PPED_URL), fe + " (votos) ; " + (DANE_FECHA if a == 2017 else PPED_FECHA) + " (poblacion)", "con dato",
         f"Sensibilidad (no entra al puntaje): pueden votar los mayores de 14 anios. Poblacion total {fint(POB[a])}; de 14+ {fint(POB14[a])} ({100 * POB14[a] / POB[a]:.1f} %). "
         + ("Con esta poblacion 2021 PASARIA DEL 7 % (referente del IPCM); el puntaje de I1 se mantiene con la poblacion total para ser coherente con la LB y el IPCM oficial. " if a == 2021 else
            "El IPCM oficial y la LB usan poblacion total; la comparacion con el referente 7 % se hace con IPCM total.") + (" " + ALERTA_UNIDAD if a == 2025 else ""), dec=3)
check("Sensibilidad 14+: 2021 supera el 7 %", ip(VOTOS[2021], POB14[2021]) > 7.0, f"{ip(VOTOS[2021], POB14[2021]):.3f}")

# ================================================================== I2: cobertura territorial (zonas 2 a 13 / 12; D41)
PI_ARCH = "180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx"
PI_FECHA, PI_URL, _ = man("plan_indicativo", PI_ARCH)
wb = openpyxl.load_workbook(RAW / "plan_indicativo" / PI_ARCH, read_only=True, data_only=True)
PI = {}
for r in wb["Plan-Indicativo-Productos"].iter_rows(values_only=True):
    if r[0] is not None and str(r[0]).strip().isdigit() and 520 <= int(str(r[0]).strip()) <= 531:
        PI[int(str(r[0]).strip())] = {"producto": r[37], "indicador": r[38], "av24": r[64], "av25": r[65], "obs24": r[136] or "", "obs25": r[137] or "", "obs26": r[138] or "", "av26": r[66]}
wb.close()
check("Plan indicativo: consecutivos 520 a 531 presentes", sorted(PI) == list(range(520, 532)))
check("Plan indicativo 528 (rendicion de PP): AVANCE 2026 = 1 y la observacion dice 'ultimo trimestre' (contradiccion de la fuente)", float(PI[528]["av26"]) == 1 and "ltimo trimestre" in norm(PI[528]["obs26"]))
check("Plan indicativo 531 (asambleas zonales): 12 en 2024 y 12 en 2025", float(PI[531]["av24"]) == 12 and float(PI[531]["av25"]) == 12)
t92b = t92
check("Informe de gestion 2025, p. 92: '12 asambleas zonales' y 'totalidad de las zonas'", "12 asambleas zonales" in t92b and "totalidad de las zonas" in t92b)
# asambleas de formulacion del PDM 2024 por zona (soporte del repositorio)
zpdm = set()
with open(SOPORTE, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        if r["codigo_indicador"] == "PC.I2" and r["anio"] == "2024" and "tipo=pdm" in r["banderas"] and "evidencia=realizacion" in r["banderas"]:
            zpdm.add(int(re.search(r"zona=(\d+)", r["banderas"]).group(1)))
check("Asambleas de formulacion del PDM 2024 (anteproyecto 115749, p. 53): zonas 2 a 13 en el soporte (la zona 1 se excluye, D41)", set(range(2, 14)) <= zpdm and len(zpdm) == 13, str(sorted(zpdm)))
ZONAS_PP = defaultdict(set)
with open(CRUCE_PP, encoding="utf-8-sig", newline="") as f:
    for r in csv.DictReader(f, delimiter=";"):
        ZONAS_PP[int(r["anio"])].add(int(r["zona"]))
check("Respaldo I2: en cada anio 2020-2025 las zonas con proyectos de PP = {2..13}", all(ZONAS_PP[a] == set(range(2, 14)) for a in PERIODO), str({a: len(ZONAS_PP[a]) for a in PERIODO}))
FU_I2 = {
    2024: ("Alcaldia de Envigado: asambleas zonales de formulacion del PDM 2024-2027 (anteproyecto 115749, p. 53; zonas 2 a 13 segun outputs/indicadores/soporte_conteos.csv) y asambleas de PP: Plan Indicativo de Productos "
           "(evaluacion definitiva 2T-2026), hoja Plan-Indicativo-Productos, consecutivo 531 = 12 asambleas",
           RC["115749_anteproyecto29022024.pdf"][1] + " ; " + PI_URL,
           RC["115749_anteproyecto29022024.pdf"][0] + " (anteproyecto del PDM, archivado en rendicion_cuentas/) ; " + PI_FECHA + " (plan indicativo)"),
    2025: ("Alcaldia de Envigado, informe de gestion por dependencias 2025, p. 92 del PDF ('12 asambleas zonales ... cubriendo la totalidad de las zonas'); plan indicativo, consecutivo 531 = 12",
           U25 + " ; " + PI_URL, F25 + " (informe) ; " + PI_FECHA + " (plan indicativo)"),
}
NOTA_I2 = ("Zonas con actividad participativa / 12 (zonas 2 a 13; la zona 1, industrial y sin comite zonal, se excluye del numerador y del denominador: D41). ")
for a in PERIODO:
    if a in FU_I2:
        fu, en, fe = FU_I2[a]
        nota = NOTA_I2 + ({
            2024: "En 2024 hay lista por zona de las 13 asambleas de formulacion del PDM (29-ene a 17-feb-2024; zonas 2 a 13 con asamblea) y, como respaldo, el plan indicativo reporta 12 asambleas zonales de PP (consecutivo 531), cuya asignacion a las zonas 2-13 es una inferencia del equipo (agregado anual).",
            2025: ("ALERTA (D73): agregado reportado por la Alcaldia, sin lista por zona; respaldo PP 12/12; se solicita la lista de asambleas por zona. El informe dice 12 asambleas zonales en agosto de 2025 "
                   "que cubren la totalidad de las zonas; plan indicativo, consecutivo 531 = 12; no hay actas por zona publicadas."),
        }[a])
        fila("Cobertura territorial de la participacion", "PC.I2", "principal", a, 12, 12, 100.0, 100.0, PESOS["I2"], "12 zonas (2 a 13; D41)", fu, en, fe, "con dato", nota)
    else:
        fila_nd("Cobertura territorial de la participacion", "PC.I2", a, PESOS["I2"], "12 zonas (2 a 13; D41)",
                "no hay publicacion de actividades participativas por zona en " + str(a), "Informes de gestion y portal de la Alcaldia: sin lista de zonas", "", "2026-11-18",
                nota=NOTA_I2 + "El respaldo PC.I2.RESP (zonas con proyectos de PP) existe, pero no entra al indice. Solicitar a Planeacion las actividades participativas por zona con acta o asistencia.")
for a in PERIODO:
    zs = sorted(ZONAS_PP[a] - {1})
    fila("Respaldo: zonas con proyectos de PP / 12", "PC.I2.RESP", "respaldo", a, len(zs), 12, 100.0 * len(zs) / 12, None, "", "12 zonas (2 a 13; D41)",
         "Listados de proyectos de PP de la EFPC (efpc.envigado.gov.co) cruzados con la ejecucion presupuestal (outputs/construccion/PP/pp_cruce_proyectos.csv)",
         "https://efpc.envigado.gov.co/wp-content/uploads/2025/07/Proyectos_PP_%d.pdf" % a, man("efpc_listados_pp", f"Proyectos_PP_{a}.pdf")[0], "con dato",
         f"Zonas con al menos un proyecto de PP en {a}: {', '.join(map(str, zs))} (la zona 1 no aparece en ningun anio). Respaldo: no entra al indice (los titulos de los listados dicen 'proyectos ejecutados'; "
         "mide proyectos, no actividad participativa) .".replace(" .", "."))
sin2 = [a for a in PERIODO if a not in FU_I2]

# ================================================================== I3: oferta formativa (D49)
IG24 = "131997_24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf"
F24, U24, _ = man("informes_gestion_2024", IG24)
P24 = RAW / "informes_gestion_2024" / IG24
IG23 = "ig2023.pdf"
F23, U23, _ = man("informes_gestion", IG23)
P23 = RAW / "informes_gestion" / IG23
t88 = pdf_pag(P25, 88)
check("Informe de gestion 2025, p. 88: 3 diplomados, 9 seminarios y 9 cursos culminados",
      "3 diplomados (Gesti" in t88 and "9 seminarios y 9 cursos" in t88)
check("Informe de gestion 2025, p. 88: no ejecucion de 2 semilleros, 1 diplomado y 3 seminarios", "no ejecuci" in t88 and "2 semilleros, 1 diplomado y 3 seminarios" in t88)
REAL25 = 3 + 9 + 9                       # D72: el diplomado en Gestion Publica cuenta (la fuente no dice que fuera para servidores)
NOEJ25 = 1 + 3                           # 1 diplomado y 3 seminarios no ejecutados (los 2 semilleros no entran: no son cursos, diplomados ni seminarios)
PROG25 = REAL25 + NOEJ25
check("I3 2025 = 21/25 = 84,00 (D72)", (REAL25, PROG25) == (21, 25) and abs(100 * REAL25 / PROG25 - 84.0) < 0.005)
t50, t51 = pdf_pag(P23, 50), pdf_pag(P23, 51)
check("Informe de gestion 2023, pp. 50-51 (proyecto 0541): 1 diplomado, 11 talleres, 3 conversatorios y 2 talleres de organizaciones/gestion",
      "Proyecto 0541" in t50 and "1 Diplomado en liderazgo social y comunitario" in t50 and "11 talleres" in t50 and "3 conversatorios" in t51
      and "Taller de fortalecimiento para las organizaciones sociales" in t51 and "Taller de gesti" in t51)
CONT23 = 1 + 11 + 3 + 2
t63 = pdf_pag(P24, 63)
ACT24 = ["Diplomado en Planeaci", "Diplomado en Liderazgo", "Curso – Taller: Gesti", "Curso Ley 1757", "Curso Contrataci", "Cátedra de Estudios Afrocolombianos", "Seminario Conciliaci"]
check("Informe de gestion 2024, p. 63: 7 acciones de la Escuela de Formacion Ciudadana (sin los semilleros)", all(x in t63 for x in ACT24) and "Semilleros de liderazgo" in t63)
CONT24 = len(ACT24)
check("Plan indicativo 521, EG 2024: 95 personas certificadas; 525, BM 2024 = 0 funcionarios",
      "95 personas certificadas" in norm(PI[521]["obs24"]) and float(PI[525]["av24"]) == 0 and float(PI[521]["av24"]) == 95)
check("Conteos de formacion: 2023 = 17, 2024 = 7, 2025 = 21 (D72)", (CONT23, CONT24, REAL25) == (17, 7, 21))
I3_NOTA = ("Acciones formativas realizadas / acciones programadas. El denominador es RECONSTRUIDO (no hay programacion publicada; D49): realizadas 21 (3 diplomados, 9 seminarios y 9 cursos culminados, informe de gestion 2025, p. 88) "
           "+ no ejecutadas 4 (1 diplomado y 3 seminarios) = 25 (D72). El diplomado en Gestion Publica SI cuenta: la fuente no dice que fuera para servidores (una version previa lo excluia). "
           "Los semilleros (7 planeados, 5 ejecutados) no se cuentan.")
for a in (2020, 2021, 2022, 2023, 2024):
    fila_nd("Cumplimiento de la oferta formativa", "PC.I3", a, PESOS["I3"], "Acciones programadas (reconstruidas)", "sin programacion publicada", "Informes de gestion de la Alcaldia: sin programacion anual de la oferta formativa",
            "", "2026-11-18", nota="No hay denominador (programacion) publicado. Respaldo PC.I3.CONTEO (numero de acciones, solo ciudadania) en 2023 y 2024. Solicitar la programacion anual a la Subdireccion de Participacion Ciudadana.")
fila("Cumplimiento de la oferta formativa", "PC.I3", "principal", 2025, REAL25, PROG25, 100.0 * REAL25 / PROG25, 100.0 * REAL25 / PROG25, PESOS["I3"], "Acciones programadas (reconstruidas)",
     "Alcaldia de Envigado, informe de gestion por dependencias 2025, p. 88 del PDF (proyecto de formacion para la participacion)", U25, F25, "con dato", I3_NOTA)
CONT = {
    2023: (CONT23, "Informe de gestion por dependencias 2023 (vigencia 2023), pp. 50-51 del PDF, proyecto 0541", U23, F23,
           "1 diplomado en liderazgo social y comunitario + 11 talleres + 3 conversatorios + 2 talleres de organizaciones y de gestion = 17. ALERTA: incluye talleres y conversatorios; con solo cursos, diplomados y seminarios seria 1. "
           "Se excluye la capacitacion a funcionarios de la Subdireccion."),
    2024: (CONT24, "Informe de gestion por dependencias 2024, p. 63 del PDF (Escuela de Formacion Ciudadana); plan indicativo, consecutivo 521 (EG 2024: 95 personas certificadas) y 525 (BM 2024 = 0 funcionarios formados)", U24 + " ; " + PI_URL, F24 + " ; " + PI_FECHA,
           "2 diplomados + 3 cursos + 1 catedra + 1 seminario = 7 formaciones para la ciudadania (95 certificados); los semilleros no se cuentan. Funcionarios formados 2024 = 0 (consecutivo 525), por lo que no hay mezcla con servidores (D49)."),
    2025: (REAL25, "Informe de gestion por dependencias 2025, p. 88 del PDF", U25, F25, "3 diplomados + 9 seminarios + 9 cursos culminados = 21 (D72: incluye el diplomado en Gestion Publica; la fuente no dice que fuera para servidores)."),
}
for a, (n, fu, en, fe, nt) in CONT.items():
    fila("Respaldo: numero de acciones formativas para la ciudadania", "PC.I3.CONTEO", "respaldo", a, None, None, float(n), None, "", "Sin referente (conteo)", fu, en, fe, "con dato",
         nt + " Es un conteo, no una tasa: no entra al indice. Los criterios de conteo no son identicos entre anios.", dec=0)

# ================================================================== I4: publicacion de resultados (D29, D70)
INV = list(csv.DictReader(open(INVENTARIO, encoding="utf-8-sig")))


def inv(idd):
    """(url, pagina, titulo, fecha_publicacion) de un documento del inventario por su id numerico; falla si no esta."""
    for r in INV:
        if re.search(rf"/{idd}_", r["url_origen"]) and r["url_origen"].startswith("https://www.envigado.gov.co"):
            return r["url_origen"], r["pagina_fuente"], r["titulo_publicacion"], r["fecha_publicacion"]
    for r in INV:
        if re.search(rf"/{idd}_", r["url_origen"]):
            return r["url_origen"], r["pagina_fuente"], r["titulo_publicacion"], r["fecha_publicacion"]
    raise KeyError(idd)


FECHA_INV = "2026-09-25 (corte del inventario de 2.787 archivos del repositorio)"
D_RCA = {2021: [23449, 40433], 2022: [92036], 2023: [115926], 2024: [129738], 2025: [149610]}
D_RPP = {2021: list(range(35775, 35782)), 2023: list(range(112889, 112897)), 2024: list(range(161584, 161594))}
ARCH_RCA = {2020: 6385, 2021: 23449, 2022: 92036, 2023: 115926, 2024: 129738, 2025: 149610}     # 2020: documento de Salud (no cuenta)
ARCH_RPP = {2020: 5930, 2021: 35775, 2022: 59050, 2023: 112890, 2024: 161584}
TITULO_OVR = {129738: "Preguntas y respuestas de la Rendición de Cuentas, gestión 2024"}
_faltan = []
for lst in list(D_RCA.values()) + list(D_RPP.values()):
    for i in lst:
        try:
            inv(i)
        except KeyError:
            _faltan.append(i)
check("I4: todos los documentos de rendicion del Alcalde y de PP citados existen en el inventario", not _faltan, str(_faltan))
check("I4: hay un documento archivado por proceso y anio (RCA 2020-2025 y RPP 2020-2024)", all(rc_arch(i) for i in list(ARCH_RCA.values()) + list(ARCH_RPP.values())))
_g23 = [inv(i) for i in D_RPP[2023]]
check("I4 2023: 8 presentaciones de PP (112889-112896) en la pagina ano-2023-855308, publicadas el 23-feb-2024",
      len(_g23) == 8 and all(g[1] == "https://www.envigado.gov.co/ano-2023-855308" and g[3] == "2024-02-23" for g in _g23))
_g24 = [inv(i) for i in D_RPP[2024]]
check("I4 2024: 10 fichas de rendicion de PP (161584-161593) publicadas el 28-feb-2025", len(_g24) == 10 and all(g[3] == "2025-02-28" for g in _g24))
paginas = {r["pagina_fuente"]: r for r in INV}
PAG_RPP = {2020: "https://www.envigado.gov.co/presentaciones-rendicion-proyectos-2020",
           2022: "https://www.envigado.gov.co/fichas-rendicion-de-cuentas-de-priorizacion-participativa-116882"}
for u in PAG_RPP.values():
    check(f"I4: pagina {u.split('/')[-1]} esta en el inventario", u in paginas)
VOT, RCA, RPP = {}, {}, {}
for a in PERIODO:
    VOT[a] = 1
    RCA[a] = 1 if a in D_RCA else 0
    RPP[a] = 1 if a in ARCH_RPP else None      # 2025 pendiente
check("I4: votacion de PP 2020-2025 con listado EFPC archivado", all((RAW / "efpc_listados_pp" / f"Proyectos_PP_{a}.pdf").exists() for a in PERIODO))
I4MIN, I4MAX = {}, {}
for a in PERIODO:
    I4MIN[a] = VOT[a] + RCA[a] + (RPP[a] or 0)
    I4MAX[a] = VOT[a] + RCA[a] + (RPP[a] if RPP[a] is not None else 1)
check("I4 = 2/3, 3/3, 3/3, 3/3, 3/3 y 2025 en rango 2/3 a 3/3", [I4MIN[a] for a in PERIODO] == [2, 3, 3, 3, 3, 2] and I4MAX[2025] == 3 and I4MAX[2020] == 2)
NOTA_I4 = "Procesos con resultados publicados / 3 procesos fijos (votacion de PP, rendicion de cuentas del Alcalde, rendicion de cuentas de PP; D29). Anio = vigencia. "
REF_I4 = "3 procesos fijos (D29)"
FU_I4 = "Portal de la Alcaldia (www.envigado.gov.co), documentos del inventario del repositorio y listados EFPC; detalle en PC.I4.VOT, PC.I4.RCA y PC.I4.RPP"
NOTAS_I4 = {
    2020: "La rendicion de cuentas del Alcalde de 2020 no se encontro: la pagina 'Rendicion de Cuentas 2020' (docs 6384 y 6385) es de la Secretaria de Salud (doc 6385 archivado y verificado: 'Secretaria de Salud de Envigado, Rendicion de cuentas 2020'). No se cuenta (0). Si se localiza, I4 2020 seria 3/3.",
    2022: "Rendicion del Alcalde confirmada: doc 92036, pp. 8-11, 'Respuestas a preguntas'.",
    2025: ("RANGO (D70, D74): al indice va el minimo, 2/3 = 66,67; maximo 3/3 (PC.I4.MAX); brecha 1/3. La busqueda en el portal del 18-nov-2026 no encontro la rendicion de PP de la vigencia 2025 "
           "(las convocatorias halladas son de feb-2024, vigencia 2023). La rendicion de la vigencia 2024 se hizo en feb-2025 (fichas publicadas el 28-feb-2025). El plan indicativo, consecutivo 528, marca AVANCE 2026 = 1, "
           "aunque su observacion dice que la rendicion 'se realizara en el ultimo trimestre del ano' (contradiccion de la fuente). Se actualiza cuando se publique."),
}
for a in PERIODO:
    ob = f"{I4MIN[a]}/3"
    brecha = "con dato" if a != 2025 else "con dato (rango D70): minimo 2/3; maximo 3/3; brecha 1/3 (rendicion de PP 2025 pendiente)"
    fila("Publicacion de resultados de procesos participativos", "PC.I4", "principal", a, I4MIN[a], 3, 100.0 * I4MIN[a] / 3, 100.0 * I4MIN[a] / 3, PESOS["I4"], REF_I4, FU_I4,
         "https://www.envigado.gov.co/ ; ver PC.I4.VOT, PC.I4.RCA, PC.I4.RPP", RC_FECHA + " (documentos archivados en rendicion_cuentas/) ; " + man("efpc_listados_pp", f"Proyectos_PP_{a}.pdf")[0] + " (listado EFPC)", brecha, NOTA_I4 + f"Resultado {ob}. " + NOTAS_I4.get(a, ""))
fila("Publicacion de resultados: maximo posible 2025", "PC.I4.MAX", "respaldo", 2025, 3, 3, 100.0, None, "", REF_I4, "Si se publica la rendicion de cuentas de PP 2025 (consecutivo 528 del plan indicativo)", PI_URL, PI_FECHA, "con dato (rango)",
     "Maximo del rango D70: no entra al indice (al indice va el minimo, 66,67). Brecha 1/3.")


def doc_txt(ids):
    out = []
    for i in ids:
        u, p, t, fp = inv(i)
        out.append(f"{i} ({TITULO_OVR.get(i, t)}; publicado {fp})")
    return "; ".join(out)


def doc_url(ids):
    return " ; ".join(inv(i)[0] for i in ids)


NM = {"VOT": "Votacion de PP: listado de proyectos publicado", "RCA": "Rendicion de cuentas del Alcalde: documento de respuestas publicado", "RPP": "Rendicion de cuentas de PP: informes o fichas publicados"}
for a in PERIODO:
    efpc = man("efpc_listados_pp", f"Proyectos_PP_{a}.pdf")
    fila(NM["VOT"], "PC.I4.VOT", "detalle", a, None, None, None, None, "", REF_I4, f"Listado de proyectos de PP {a} de la Escuela de Formacion para la Participacion Ciudadana (EFPC)", efpc[1], efpc[0], "con dato",
         "1 = publicado. Los titulos del listado dicen 'proyectos ejecutados'; se usa como publicacion del resultado de la votacion." + (" En 2020 el valor del listado es el priorizado (priorizacion de 2019)." if a == 2020 else ""),
         raw={"valor": "1"})
    if a in D_RCA:
        arch = rc_arch(ARCH_RCA[a])
        fila(NM["RCA"], "PC.I4.RCA", "detalle", a, None, None, None, None, "", REF_I4, "Alcaldia de Envigado, rendicion de cuentas del Alcalde, vigencia " + str(a) + ": " + doc_txt(D_RCA[a]) + f"; archivado: {arch}",
             doc_url(D_RCA[a]), RC_FECHA + " (archivado; manifiesto rendicion_cuentas)", "con dato",
             "1 = publicado." + (" " + NOTAS_I4[2022] if a == 2022 else "") + (" Solo el doc 23449 esta archivado; el 40433 (informe de respuestas, jul-2022) se cita del inventario." if a == 2021 else ""), raw={"valor": "1"})
    else:
        arch = rc_arch(ARCH_RCA[a])
        fila(NM["RCA"], "PC.I4.RCA", "detalle", a, None, None, None, None, "", REF_I4, f"Portal de la Alcaldia: no se encontro la rendicion de cuentas del Alcalde de 2020; archivado el doc 6385 para documentar que es de Salud: {arch}",
             "https://www.envigado.gov.co/rendicion-de-cuentas-2020-822054 ; " + RC[arch][1], RC_FECHA + " (archivado; manifiesto rendicion_cuentas)",
             "sin documento encontrado", "0 = no encontrado. " + NOTAS_I4[2020], raw={"valor": "0"})
    if a == 2025:
        fila(NM["RPP"], "PC.I4.RPP", "detalle", a, None, None, None, None, "", REF_I4, "Plan indicativo de productos, consecutivo 528 (eventos de rendicion publica de cuentas de PP); busqueda en el portal del 18-nov-2026: no hay rendicion de PP de la vigencia 2025", PI_URL, PI_FECHA,
             "pendiente", "pendiente: al indice cuenta 0 (minimo) y al maximo 1 (D70, D74). " + NOTAS_I4[2025], raw={"valor": "pendiente"})
    else:
        arch = rc_arch(ARCH_RPP[a])
        if a in D_RPP:
            ids = D_RPP[a]
            pub = inv(ids[0])[3]
            fu = (f"Alcaldia de Envigado, rendicion de cuentas de PP, vigencia {a}: {len(ids)} documentos (ids {ids[0]}-{ids[-1]}; pagina {inv(ids[0])[1]}; publicados {pub}); archivado uno: {arch}")
        else:
            pg = PAG_RPP[a]
            fu = f"Alcaldia de Envigado, rendicion de cuentas de PP, vigencia {a}: pagina del portal con fichas o presentaciones {pg} (publicado {paginas[pg]['fecha_publicacion']}); archivado uno: {arch}"
        en = RC[arch][1] + (" ; " + inv(D_RPP[a][0])[1] if a in D_RPP else " ; " + PAG_RPP[a])
        fila(NM["RPP"], "PC.I4.RPP", "detalle", a, None, None, None, None, "", REF_I4, fu, en, RC_FECHA + " (archivado; manifiesto rendicion_cuentas)", "con dato",
             "1 = publicado." + (" Evidencia solida: 8 presentaciones de PP (112889-112896) en la pagina ano-2023-855308, publicadas el 23-feb-2024." if a == 2023 else "")
             + (" Rendicion de la vigencia 2024 realizada en feb-2025 (fichas publicadas el 28-feb-2025)." if a == 2024 else ""), raw={"valor": "1"})

# ================================================================== ISPC
COMP = {
    "I1": {a: P1.get(a) for a in PERIODO},
    "I2": {a: (100.0 if a in FU_I2 else None) for a in PERIODO},
    "I3": {a: (100.0 * REAL25 / PROG25 if a == 2025 else None) for a in PERIODO},
    "I4": {a: 100.0 * I4MIN[a] / 3 for a in PERIODO},
}
MOT = {"I1": {2020: "sin cifra de votos de la votacion de 2020", 2021: None, 2022: None, 2023: None, 2024: "sin cifra de votos de la votacion de 2024", 2025: None},
       "I2": {a: "sin dato por zona" for a in PERIODO}, "I3": {a: "sin programacion publicada" for a in PERIODO}, "I4": {}}
MOT["I1"][2019] = "sin cifra de votos"
ISPC, FALTAN = {}, {}
for a in PERIODO:
    falt = [k for k in PESOS if COMP[k][a] is None]
    FALTAN[a] = falt
    if falt:
        ISPC[a] = None
        fila("ISPC, indice de participacion ciudadana", "PC.ISPC", "indice", a, peso="", ref="0,30*I1 + 0,25*I2 + 0,20*I3 + 0,25*I4 (D42); semaforo D60", fuente="Calculo propio a partir de los 4 indicadores",
             fecha=FECHA_CALC, brecha="sin dato: faltan " + ", ".join(falt),
             nota=f"N/D: faltan {', '.join(falt)} ({'; '.join(k + ': ' + MOT[k][a] for k in falt)}). No se redistribuyen pesos. Componentes con dato: "
                  + ", ".join(f"{k} {fmt(COMP[k][a])}" for k in PESOS if COMP[k][a] is not None) + ".",
             raw={"numerador": "N/D", "denominador": "N/D", "valor": "N/D", "puntaje": "N/D"})
    else:
        ISPC[a] = sum(PESOS[k] * COMP[k][a] for k in PESOS)
        fila("ISPC, indice de participacion ciudadana", "PC.ISPC", "indice", a, v=ISPC[a], puntaje=ISPC[a], ref="0,30*I1 + 0,25*I2 + 0,20*I3 + 0,25*I4 (D42); semaforo D60",
             fuente="Calculo propio a partir de los 4 indicadores", fecha=FECHA_CALC, brecha="con dato",
             nota="Puntajes: " + "; ".join(f"{k} {COMP[k][a]:.2f}" for k in PESOS) + ". Valores sin redondear. I4 al minimo (2/3, D70): con la rendicion de PP 2025 publicada (I4 = 100) el ISPC seria "
                  f"{ISPC[a] + PESOS['I4'] * (100 - COMP['I4'][a]):.2f}. I1 con LB 3,68 (D69). Convencion (QA): anio = anio calendario de la actividad participativa (I1: votacion para la vigencia t+1; I4: anio de la vigencia). Alertas: I2 2025 es un agregado reportado por la Alcaldia, sin lista por zona (D73); I3 con denominador reconstruido (21/25, D72); I1 con votantes (no votos) y poblacion total.")
        MAXV = ISPC[a] + PESOS["I4"] * (100 - COMP["I4"][a])
        fila("ISPC con I4 al maximo (rango D70)", "PC.ISPC.MAX", "respaldo", a, v=MAXV, ref="Si se publica la rendicion de PP 2025", fuente="Calculo propio", fecha=FECHA_CALC, brecha="con dato (rango)",
             nota="Informativo: no entra a la serie ni al semaforo. Al indice va el minimo.")
check("ISPC 2025 = 61,62 +/- 0,01 (Naranja; D72)", ISPC[2025] is not None and abs(ISPC[2025] - 61.62) <= 0.01 and semaforo(ISPC[2025]) == "Naranja", fmt(ISPC[2025]))
check("ISPC 2025 con I4 = 100 (calculado) = 69,95 +/- 0,01", abs(ISPC[2025] + PESOS["I4"] * (100 - COMP["I4"][2025]) - 69.95) <= 0.01, fmt(ISPC[2025] + PESOS["I4"] * (100 - COMP["I4"][2025])))
check("Pesos suman 1", abs(sum(PESOS.values()) - 1) < 1e-12)
check("ISPC solo calculable en 2025; 2020-2024 N/D", [a for a in PERIODO if ISPC[a] is not None] == [2025])
check("Componentes faltantes: 2020 = I1, I2, I3; 2021-2023 = I2, I3; 2024 = I1, I3",
      FALTAN[2020] == ["I1", "I2", "I3"] and FALTAN[2021] == ["I2", "I3"] and FALTAN[2022] == ["I2", "I3"] and FALTAN[2023] == ["I2", "I3"] and FALTAN[2024] == ["I1", "I3"], str(FALTAN))

# ================================================================== brechas
B = []
con1 = [a for a in PERIODO_I1 if a in VOTOS]
sin1 = [a for a in PERIODO_I1 if a not in VOTOS]
check("Brecha I1 = 3 de 8 anios (2019, 2020 y 2024 sin dato; 2018 no aplica)", sin1 == [2019, 2020, 2024] and len(PERIODO_I1) == 8)
fila("Brecha de informacion de la votacion de PP (2017-2025, sin 2018)", "PC.I1.BRECHA", "brecha", "2017-2025", len(sin1), len(PERIODO_I1), 100.0 * len(sin1) / len(PERIODO_I1), brecha=f"{len(sin1)} de {len(PERIODO_I1)} anios sin dato",
     nota="Con dato: " + ", ".join(map(str, con1)) + "; sin dato: 2019 (sin cifra), 2020 (sin cifra final; 5.100 parciales no se usan) y 2024 (sin cifra). 2018 no cuenta: no hubo votacion (vigencia 2019). Brecha = anios sin dato / anios del periodo. "
          "Fuera del periodo 2020-2025 del indice: 2017 y 2019.", fuente="", enlace="", fecha=FECHA_CALC)
check("Brecha I2 = 4 de 6 (2020-2023)", sin2 == [2020, 2021, 2022, 2023])
fila("Brecha de informacion de la cobertura territorial (2020-2025)", "PC.I2.BRECHA", "brecha", "2020-2025", len(sin2), 6, 100.0 * len(sin2) / 6, brecha=f"{len(sin2)} de 6 anios sin dato",
     nota="Con dato: 2024 y 2025 (2025 como agregado reportado por la Alcaldia sin lista por zona; D73); sin dato: 2020-2023 (solo respaldo PC.I2.RESP). Brecha = anios sin dato / anios del periodo.", fecha=FECHA_CALC)
fila("Brecha de informacion de la oferta formativa (2020-2025)", "PC.I3.BRECHA", "brecha", "2020-2025", 5, 6, 100.0 * 5 / 6, brecha="5 de 6 anios sin dato",
     nota="Con dato: 2025 (denominador reconstruido); sin dato: 2020-2024 (sin programacion publicada; respaldo de conteo en 2023 y 2024). Brecha = anios sin dato / anios del periodo.", fecha=FECHA_CALC)
fila("Brecha de informacion de la publicacion de resultados (2020-2025)", "PC.I4.BRECHA", "brecha", "2020-2025", 0, 6, 0.0, brecha="0 de 6 anios sin dato; 2025 en rango (brecha 1/3); 2 de 18 proceso-anios sin documento",
     nota="Todos los anios tienen dato. Procesos sin resultado publicado: rendicion del Alcalde 2020 (no encontrada) y rendicion de PP 2025 (pendiente, ultimo trimestre de 2026): 2 de 18 proceso-anios (3 procesos x 6 anios). "
          "2025 se muestra como rango 2/3 a 3/3 (D70).", fecha=FECHA_CALC)
sinI = [a for a in PERIODO if ISPC[a] is None]
fila("Brecha de informacion del ISPC (2020-2025)", "PC.ISPC.BRECHA", "brecha", "2020-2025", len(sinI), 6, 100.0 * len(sinI) / 6, brecha=f"{len(sinI)} de 6 anios sin indice",
     nota="Con indice: 2025. Sin indice: " + "; ".join(f"{a} (faltan {', '.join(FALTAN[a])})" for a in sinI) + ". Brecha = anios sin indice / anios del periodo.", fecha=FECHA_CALC)
check("Brecha ISPC = 5 de 6", len(sinI) == 5)

# ------------------------------------------------------------------ salida
SALIDA.parent.mkdir(parents=True, exist_ok=True)
with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
    w.writeheader()
    w.writerows(FILAS)
print("\nI1 (IPCM %, puntaje):", {a: (round(IPCM[a], 3), round(P1[a], 2)) for a in VOTOS})
print("I4:", {a: f"{I4MIN[a]}/3" for a in PERIODO}, "| ISPC:", {a: (fmt(v) + " " + semaforo(v) if v is not None else "N/D") for a, v in ISPC.items()})
fallos = [n for n, ok, _ in ASSERTS if not ok]
print("filas:", len(FILAS), "->", SALIDA, "| asserts:", len(ASSERTS), "| fallos:", len(fallos))
if fallos:
    print("ASSERTS FALLIDOS:", fallos)
    sys.exit(1)
