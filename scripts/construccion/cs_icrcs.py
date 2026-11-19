#!/usr/bin/env python3
"""Dimension Control social (ICRCS): PQRSDF, DSI, VRC y RPO.

ICRCS = (DSI + PQRSDF + VRC + RPO) / 4, peso 0,25 cada uno (D45). VRC, RPO y DSI entran con su MINIMO (D33, D75). No se redistribuyen pesos:
si falta un componente el anio queda N/D. Semaforo D60: >= 90 Verde, >= 75 Amarillo, >= 50 Naranja, < 50 Rojo. Periodo 2020-2025.

Indicadores
  CS.PQRSDF (D45, D78) 100 * ET / R con tope 100; ET = 'Atendida en terminos', R = 'Total' de la fila Total de la tabla 'Atencion en Terminos BPMS' del
                       informe trimestral; el anio es la suma de los 4 trimestres (con menos de 4 informes legibles queda N/D).
                       Rango: maximo = (ET + Pendientes) / R; brecha = Pendientes / R. Respaldo SAC (solo Educacion; no se suma). Lector de PDF generico.
  CS.DSI    (D75)      10 metas del Apendice 9.1.2; una meta cuenta en t si hay valor de la vigencia t con la misma definicion y unidad, publicado por la
                       fuente que nombra el PDM, a mas tardar el 31-mar de t+1 y con fecha de publicacion verificable. Minimo = verificadas / 10;
                       maximo = minimo + no verificables / 10; brecha = no verificables / 10.
  CS.VRC    (D33, D77) eventos de rendicion de cuentas con veeduria documentada (minimo), + desconocidos (maximo); 2021-2023 pendiente de descargas.
  CS.RPO    (D76)      observaciones sin felicitaciones de la audiencia del Alcalde con respuesta de fondo publicada al 31-dic (busqueda hasta 31-mar t+1).

Insumos (data_raw_IIC/07_revision_nov2026/, cada carpeta con manifiesto.csv; el script verifica bytes y sha256):
  pqrs/                         informes trimestrales de PQRSDF (los baja scripts/construccion/cs_descargas.py; hoy solo 2023 T4)
  pdm/116989_...xlsx            Apendice 9.1.2 (LB y meta de las 10 metas del DSI)
  dsi/                          instantanea y metadatos de MEN nudc-7mev (fila 79)
  men_nudc7mev/, terridata/     cobertura neta (2011-2024) y tasa de homicidio (TerriData 2024 y 2025)
  rendicion_cuentas/            cuadros de preguntas y respuestas del Alcalde 2021-2025 (RPO y VRC)
  informes_gestion/152890_...   informe de gestion 2025, p. 102 (fecha de la audiencia del Alcalde: 13-nov-2025)
  outputs/indicadores/clasificacion_vrc.csv, clasificacion_rpo.csv y outputs/revision_indicadores/CS/vrc_brecha.csv
Salida : outputs/construccion/CS/cs_indicadores.csv (; y UTF-8 con BOM)
Uso    : python3 scripts/construccion/cs_icrcs.py   (desde la raiz; necesita openpyxl y pypdf; falla si falta alguno)
Todo resultado es PROVISIONAL. Los archivos descargados se tratan solo como datos.
"""
import csv
import hashlib
import re
import sys
import unicodedata
import zipfile
import io
import json
import warnings
from pathlib import Path

import openpyxl
import pypdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
import cs_descargas as DESC          # lista de URL del inventario (no ejecuta descargas al importarse)

warnings.filterwarnings("ignore")
RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC/07_revision_nov2026"
SALIDA = RAIZ / "outputs/construccion/CS/cs_indicadores.csv"
CLAS_VRC = RAIZ / "outputs/indicadores/clasificacion_vrc.csv"
CLAS_RPO = RAIZ / "outputs/indicadores/clasificacion_rpo.csv"
VRC_BRECHA = RAIZ / "outputs/revision_indicadores/CS/vrc_brecha.csv"

COLS = ["indicador", "codigo", "tipo", "anio", "numerador", "denominador", "valor", "puntaje_0_100", "peso_en_indice",
        "semaforo", "referente", "fuente", "enlace", "fecha_consulta", "estado", "brecha", "nota"]
PROV = "Provisional"
UMBRALES = [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")]       # D60; Rojo por debajo
PESO = 0.25
PERIODO = list(range(2020, 2026))
FECHA_CALC = "2026-11-19"
MOTIVO_PORTAL = "portal sin respuesta el 19-nov; reintentar con cs_descargas.py"
ASSERTS = []


def check(nombre, cond, detalle=""):
    ASSERTS.append((nombre, bool(cond), detalle))
    print(("OK   " if cond else "FALLA"), nombre, detalle)


def norm(s):
    return re.sub(r"\s+", " ", s or "").strip()


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


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


def entero(s):
    return int(s.replace(".", "").replace(",", ""))


# ------------------------------------------------------------------ manifiestos (con verificacion de bytes y sha256)
def _man_rows(carpeta):
    p = RAW / carpeta / "manifiesto.csv"
    with open(p, encoding="utf-8-sig", newline="") as f:
        sep = ";" if ";" in f.readline() else ","
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=sep))


def man(carpeta, archivo):
    """(fecha_consulta, url, nota) del manifiesto; verifica que el archivo exista con los bytes y el sha256 del manifiesto."""
    for r in _man_rows(carpeta):
        if r["archivo"] == archivo:
            p = RAW / carpeta / archivo
            if not p.exists():
                raise SystemExit(f"falta el archivo {p}")
            if p.stat().st_size != int(r["bytes"]) or hashlib.sha256(p.read_bytes()).hexdigest() != r["sha256"]:
                raise SystemExit(f"bytes o sha256 no coinciden con el manifiesto: {carpeta}/{archivo}")
            return r["fecha_consulta"], r["url"], r.get("nota", "")
    raise KeyError((carpeta, archivo))


# ------------------------------------------------------------------ filas
FILAS = []


def fila(ind, cod, tipo, anio, n=None, d=None, v=None, puntaje=None, peso="", ref="", fuente="", enlace="", fecha="", brecha="", nota="", dec=2, raw=None):
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


ND = {"numerador": "N/D", "denominador": "N/D", "valor": "N/D", "puntaje": "N/D"}


def fila_nd(ind, cod, tipo, anio, motivo, peso="", ref="", fuente="", enlace="", fecha="", nota=""):
    fila(ind, cod, tipo, anio, peso=peso, ref=ref, fuente=fuente, enlace=enlace, fecha=fecha, brecha="sin dato: " + motivo, nota=nota, raw=ND)


# ================================================================== PQRSDF: lector generico de informes (trimestrales, semestrales o anuales)
MESES = {m: i + 1 for i, m in enumerate("enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre".split())}
MESES["setiembre"] = 9
ORD = {"primer": 1, "segundo": 2, "tercer": 3, "cuarto": 4}
_MES_RE = "|".join(sorted(MESES, key=len, reverse=True))


def meses_en(texto):
    """Meses (1-12) que nombra un texto, con limites de palabra (mayor no es mayo, genero no es enero) y rangos 'de enero a junio'."""
    t = sin_acentos(texto).lower()
    out = set()
    for a, b in re.findall(rf"\b({_MES_RE})\s+(?:a|hasta|al mes de)\s+({_MES_RE})\b", t):
        if MESES[a] <= MESES[b]:
            out.update(range(MESES[a], MESES[b] + 1))
    out.update(MESES[m] for m in re.findall(rf"\b({_MES_RE})\b", t))
    return sorted(out)


def leer_informe(ruta, anio_hint=None):
    """Lee un informe PQRSDF (solo dato). Devuelve dict con ok, motivo y, si ok: anio, meses, tipo (T/S/A), bpms (atendida, et, pend, total), sac, pagina.
    El anio se contrasta entre la portada, el Alcance y el inventario; si hay dudas devuelve ok=False con el motivo (no adivina)."""
    R = {"ok": False, "motivo": ""}
    try:
        lector = pypdf.PdfReader(str(ruta))
        crudo = [p.extract_text() or "" for p in lector.pages]
    except Exception as e:                                           # PDF danado o no PDF
        R["motivo"] = "PDF ilegible (" + type(e).__name__ + ")"
        return R
    pags = [norm(t) for t in crudo]
    pl = [sin_acentos(p).lower() for p in pags]
    full = " ".join(pags)
    # 1. Alcance: periodo (meses) y sistema
    alc = None
    for m in re.finditer(r"Alcance\s+(.{20,900}?)\s+3\.\s*Responsabilidades", full, re.I):
        if "...." not in m.group(1):
            alc = m.group(1)
            break
    if alc is None:
        R["motivo"] = "plantilla distinta (posible formato 2020-2021): sin seccion 'Alcance' legible"
        return R
    meses = meses_en(alc)
    if not meses:
        R["motivo"] = "'Alcance' sin meses"
        return R
    trims = {(m - 1) // 3 + 1 for m in meses}
    if any(not {3 * q - 2, 3 * q - 1, 3 * q} <= set(meses) for q in trims):
        R["motivo"] = f"los meses del Alcance {meses} no forman trimestres completos; no se adivina el periodo"
        return R
    R["meses"], R["alcance"] = meses, alc
    R["tipo"] = {3: "T", 6: "S", 12: "A"}.get(len(meses))
    if R["tipo"] is None or sorted(trims) != list(range(min(trims), max(trims) + 1)):
        R["motivo"] = f"el periodo del Alcance (meses {meses}) no es trimestre, semestre ni anio"
        return R
    R["sistemas"] = [s for s in ("BPMS", "SAC") if s in alc]
    if "BPMS" not in R["sistemas"]:
        R["motivo"] = "el Alcance no menciona BPMS (informe de otro sistema, p. ej. solo SAC de Educacion)"
        return R
    # anio: portada, Alcance e inventario deben coincidir
    ys_c = set(re.findall(r"\b(20[12]\d)\b", pags[0]))
    ys_a = set(re.findall(r"\b(20[12]\d)\b", alc))
    if ys_a:
        if len(ys_a) != 1 or (ys_c and not ys_a <= ys_c):
            R["motivo"] = "el anio del Alcance no coincide con el de la portada"
            return R
        anio = next(iter(ys_a))
    else:
        if len(ys_c) != 1:
            R["motivo"] = "no se pudo fijar el anio (portada ambigua y Alcance sin anio)"
            return R
        anio = next(iter(ys_c))
    if anio_hint is not None and int(anio) != anio_hint:
        R["motivo"] = f"el anio leido ({anio}) contradice el del inventario ({anio_hint})"
        return R
    R["anio"] = int(anio)
    portada = pl[0]
    mo = re.search(r"(primer|segundo|tercer|cuarto)\s+trimestre", portada)
    ms = re.search(r"(primer|segundo)\s+semestre", portada)
    if R["tipo"] == "T" and mo and ORD[mo.group(1)] != min(trims):
        R["motivo"] = "la portada y los meses del Alcance no coinciden"
        return R
    if R["tipo"] == "S" and ms and ORD[ms.group(1)] != min(trims) // 2 + 1:
        R["motivo"] = "la portada (semestre) y los meses del Alcance no coinciden"
        return R
    # 2. Tabla 'Atencion en Terminos BPMS'
    idx = next((i for i, p in enumerate(pl) if "atencion en terminos bpms" in p), None)
    if idx is None:
        R["motivo"] = "plantilla distinta (posible formato 2020-2021): sin tabla 'Atencion en Terminos BPMS'"
        return R
    bloque = pags[idx][pl[idx].index("atencion en terminos bpms"):]
    if "atendida atendida en terminos pendientes total" not in sin_acentos(bloque).lower():
        R["motivo"] = "encabezado de la tabla BPMS distinto del esperado"
        return R
    m = re.search(r"\bTotal\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)", bloque)
    if not m:
        R["motivo"] = "sin fila Total legible en la tabla BPMS"
        return R
    at, et, pe, tot = (entero(x) for x in m.groups())
    if at + et + pe != tot:
        R["motivo"] = f"atendida + en terminos + pendientes ({at + et + pe}) distinto del total ({tot})"
        return R
    # 3. Total = recibidas por tipo
    rec = None
    for i, p in enumerate(pl):
        if "recibidas en el bpms por cada tipo" in p:
            mm = re.search(r"\bTotal\s+([\d.]+)\s+100", pags[i][p.index("recibidas en el bpms por cada tipo"):])
            if mm:
                rec = entero(mm.group(1))
                break
    if rec is None:
        R["motivo"] = "sin tabla 'recibidas por tipo' legible para contrastar el total"
        return R
    if rec != tot:
        R["motivo"] = f"total de la tabla de atencion ({tot}) distinto de recibidas por tipo ({rec})"
        return R
    suma = [0, 0, 0, 0]                                              # suma de filas (aviso, no bloquea)
    activo = False
    for ln in crudo[idx].splitlines():
        if "TERMINOS BPMS" in sin_acentos(ln).upper():
            activo = True
            continue
        if activo:
            if norm(ln).lower().startswith("total"):
                break
            mm = re.search(r"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$", ln)
            if mm:
                suma = [a + entero(b) for a, b in zip(suma, mm.groups())]
    R["suma_filas_ok"] = suma == [at, et, pe, tot]
    R.update(ok=True, motivo="", bpms=(at, et, pe, tot), recibidas=rec, pagina=idx + 1)
    jdx = next((i for i, p in enumerate(pl) if "atencion en terminos sistema de atencion al ciudadano sac" in p), None)
    R["sac"] = None
    if jdx is not None:
        mm = re.search(r"\bTOTAL\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)", pags[jdx], re.I)
        if mm:
            a, b, c, t = (entero(x) for x in mm.groups())
            if a + b + c == t:
                R["sac"] = (a, b, c, t)
                R["sac_pagina"] = jdx + 1
    return R


def armar_cobertura(conjuntos):
    """D82: busca informes que cubran los 12 meses sin repetirse. conjuntos: lista de tuplas de meses. Devuelve (mejor cobertura o None, numero de coberturas posibles).
    Prefiere la que usa mas informes (mas granular); a igualdad, la de menor orden lexicografico."""
    conjuntos = sorted(set(conjuntos))
    res = []

    def rec(libres, usados):
        if not libres:
            res.append(list(usados))
            return
        m = min(libres)
        for c in conjuntos:
            if min(c) == m and set(c) <= libres:
                rec(libres - set(c), usados + [c])
    rec(set(range(1, 13)), [])
    if not res:
        return None, 0
    res.sort(key=lambda u: (-len(u), u))
    return res[0], len(res)


# pruebas del lector sobre textos y del armado (sin archivos)
check("Meses con limite de palabra: 'mayor' no es mayo, 'genero' no es enero, 'Junio' si", meses_en("el mayor genero, mayoria, Marzo y septiembre, y junio") == [3, 6, 9] and meses_en("mayor genero") == [])
check("Meses: 'octubre, noviembre y diciembre' = 10-12 y 'de enero a junio' = 1-6", meses_en("durante los meses de octubre, noviembre y diciembre") == [10, 11, 12] and meses_en("de enero a junio") == [1, 2, 3, 4, 5, 6])
H1, H2, QQ = (1, 2, 3, 4, 5, 6), (7, 8, 9, 10, 11, 12), [(1, 2, 3), (4, 5, 6), (7, 8, 9), (10, 11, 12)]
check("D82: semestre 1 + T3 + T4 cubre el anio", armar_cobertura([H1, (7, 8, 9), (10, 11, 12)])[0] == [H1, (7, 8, 9), (10, 11, 12)])
check("D82: con S1 y T1+T2 prefiere los trimestres (4 informes) y reporta 2 coberturas; con un hueco no hay cobertura",
      armar_cobertura([H1] + QQ) == ([QQ[0], QQ[1], QQ[2], QQ[3]], 2) and armar_cobertura([H1, (7, 8, 9)])[0] is None and armar_cobertura([(1, 2, 3), (1, 2, 3, 4, 5, 6), (4, 5, 6)] + QQ[2:])[0] == QQ)


def id_url(u):
    m = re.search(r"/(\d+)_", u)
    return int(m.group(1)) if m else 0


# ------------------------------------------------------------------ inventario (URL real de cada informe) e informes disponibles localmente
with open(DESC.INVENTARIO, encoding="utf-8-sig", newline="") as f:
    _inv = list(csv.DictReader(f))
INV_URL = {}                                                          # (anio, periodo) -> [url]
for _c, _n, _u, _nota in DESC.seleccionar(_inv):
    if _c != "pqrs":
        continue
    _m = re.search(r"inventario: (\d{4}|\?) (\w+)", _nota)
    _an = int(_m.group(1)) if _m and _m.group(1) != "?" else None
    _pe = _m.group(2) if _m and _m.group(2) in ("T1", "T2", "T3", "T4", "S1", "S2") else None
    INV_URL.setdefault((_an, _pe), []).append(_u)


def urls_inventario(a, q):
    """URL de los informes del inventario que corresponden al trimestre (a, q), con el semestre o los sin periodo identificado de ese anio."""
    s = INV_URL.get((a, f"T{q}"), []) + INV_URL.get((a, "S1" if q <= 2 else "S2"), [])
    return s or INV_URL.get((a, None), [])


INFORMES = []                                                         # (archivo, url, fecha, nota, resultado)
for r in _man_rows("pqrs"):
    if not r["archivo"].lower().endswith(".pdf"):
        continue
    fecha, url, nota = man("pqrs", r["archivo"])
    h = re.search(r"periodo segun inventario: (20\d\d)", nota) or re.search(r"(20[12]\d)", url.rsplit("/", 1)[-1])
    INFORMES.append((r["archivo"], url, fecha, nota, leer_informe(RAW / "pqrs" / r["archivo"], int(h.group(1)) if h else None)))
print(f"\nInformes PQRSDF locales: {len(INFORMES)}")
for a, u, f, n, R in INFORMES:
    print("  ", a, "->", (f"{R['anio']} meses {R['meses'][0]}-{R['meses'][-1]} BPMS {R['bpms']} p. {R['pagina']}" if R["ok"] else "NO LEGIBLE: " + R["motivo"]))

# aceptados por periodo (anio, meses); duplicados: se usa el mas reciente segun el id del archivo (D82) y se anota
ACEPT, DUP = {}, {}
for inf in INFORMES:
    R = inf[4]
    if R["ok"]:
        ACEPT.setdefault((R["anio"], tuple(R["meses"])), []).append(inf)
SEL = {}
for k, lst in ACEPT.items():
    lst.sort(key=lambda x: id_url(x[1]), reverse=True)
    SEL[k] = lst[0]
    if len(lst) > 1:
        dist = any(x[4]["bpms"] != lst[0][4]["bpms"] for x in lst[1:])
        DUP[k] = ("cifras DISTINTAS" if dist else "cifras iguales") + f": se usa el informe mas reciente (id {id_url(lst[0][1])}) sobre " + ", ".join(str(id_url(x[1])) for x in lst[1:])
NOLEG = [(a, R["motivo"]) for a, u, f, n, R in INFORMES if not R["ok"]]
COBERTO = {}                                                          # (anio, trimestre) -> informe que lo cubre (cualquier periodo aceptado)
for (a, ms), inf in SEL.items():
    for q in {(m - 1) // 3 + 1 for m in ms}:
        COBERTO.setdefault((a, q), inf)

# assert del informe 2023 T4 (BPMS 381 / 3.544 / 224 / 4.149 = 85,42, p. 11; SAC 1.513 / 2 / 80 / 1.595)
r23 = SEL.get((2023, (10, 11, 12)))
check("Informe 2023 T4 legible por el lector generico", r23 is not None)
if r23:
    R = r23[4]
    check("BPMS 2023 T4 = 381 / 3544 / 224 / 4149 (p. 11)", R["bpms"] == (381, 3544, 224, 4149) and R["pagina"] == 11, f"{R['bpms']} p. {R['pagina']}")
    check("2023 T4: 100 * 3544 / 4149 = 85,42", abs(100 * R["bpms"][1] / R["bpms"][3] - 85.42) < 0.005, fmt(100 * R["bpms"][1] / R["bpms"][3]))
    check("SAC 2023 T4 = 1513 / 2 / 80 / 1595", R["sac"] == (1513, 2, 80, 1595), str(R["sac"]))
    check("Total BPMS = recibidas por tipo (4149); Alcance: oct-dic, BPMS y SAC; anio 2023 contrastado", R["recibidas"] == 4149 and R["meses"] == [10, 11, 12] and R["sistemas"] == ["BPMS", "SAC"] and R["anio"] == 2023)
    check("Suma de las filas de la tabla BPMS = fila Total", R["suma_filas_ok"])
check("Lector: un PDF sin la plantilla esperada se marca N/D con motivo, no se adivina",
      (lambda x: x["ok"] is False and "plantilla distinta" in x["motivo"])(leer_informe(RAW / "rendicion_cuentas" / "23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf")))

# ------------------------------------------------------------------ filas CS.PQRSDF
NM_P = "Oportunidad de respuesta a las PQRSDF (atendidas en terminos / recibidas, BPMS)"
REF_P = "100 * atendidas en terminos / recibidas (BPMS), con informes que cubren los 12 meses sin repetirse (D82), tope 100; semaforo D60"
ALERTA_P = ("ALERTA: el informe admite PQRSDF respondidas a tiempo pero cerradas tarde en el sistema, por lo que el indicador SUBESTIMA la oportunidad real; "
            "es un registro administrativo de la Alcaldia (BPMS), no una verificacion externa. Las pendientes en terminos de ley no se cuentan como atendidas en terminos.")
FUENTE_P = "Alcaldia de Envigado, Secretaria General, Direccion de Calidad y SGI: Informe PQRSDF, tabla 'Atencion en Terminos BPMS', fila Total"
ETIQ = lambda ms: (f"T{(ms[0] - 1) // 3 + 1}" if len(ms) == 3 else f"S{(ms[0] - 1) // 6 + 1}" if len(ms) == 6 else "A")
TRIMESTRES = [(a, t) for a in PERIODO for t in (1, 2, 3, 4)]
ANIO_P, COB = {}, {}
for a in PERIODO:
    # filas por informe aceptado (trimestres, semestres, anuales)
    for (aa, ms), inf in sorted(SEL.items()):
        if aa != a:
            continue
        arc, url, fecha, nota, R = inf
        at, et, pe, tot = R["bpms"]
        et_ = ETIQ(ms)
        fila(f"PQRSDF, atencion en terminos BPMS, {a} {et_}", "CS.PQRSDF.TRIM" if et_.startswith("T") else "CS.PQRSDF.SEM", "detalle", f"{a}-{et_}", et, tot, 100 * et / tot,
             ref="Fila Total de la tabla 'Atencion en Terminos BPMS'", fuente=FUENTE_P + f" (p. {R['pagina']})", enlace=url, fecha=fecha, brecha="con dato",
             nota=f"Atendida (fuera de terminos) {at}; atendida en terminos {et}; pendientes {pe} (dentro del plazo legal al corte); total {tot} = recibidas por tipo. "
                  f"Alcance leido: meses {ms[0]}-{ms[-1]}, sistemas {' y '.join(R['sistemas'])}. Archivo local {arc}. Suma de filas de la tabla = fila Total: "
                  f"{'si' if R['suma_filas_ok'] else 'no (aviso)'}." + (f" Duplicado: {DUP[(aa, ms)]}." if (aa, ms) in DUP else ""))
    for q in (1, 2, 3, 4):
        if (a, tuple(range(3 * q - 2, 3 * q + 1))) in SEL:
            continue
        extra = ""
        if (a, q) in COBERTO:
            extra = f" El trimestre esta cubierto por el informe {COBERTO[(a, q)][0]} (periodo {ETIQ(COBERTO[(a, q)][4]['meses'])}); ver fila CS.PQRSDF.SEM."
        if (a, q) == (2022, 1) and (a, q) not in COBERTO:
            extra = " El inventario tiene el informe 34294 'primer semestre 2022': cuando llegue, el lector reconoce si es semestral (cubre T1+T2) o solo T1."
        urls = urls_inventario(a, q)
        fila_nd(f"PQRSDF, atencion en terminos BPMS, {a} T{q}", "CS.PQRSDF.TRIM", "detalle", f"{a}-T{q}",
                (f"sin informe trimestral propio; cubierto por un informe de otro periodo" if (a, q) in COBERTO else MOTIVO_PORTAL), ref="Fila Total de la tabla 'Atencion en Terminos BPMS'",
                fuente=FUENTE_P, enlace="; ".join(urls) if urls else "sin URL en el inventario para este trimestre", fecha="",
                nota="Informe trimestral no disponible localmente." + extra + (f" Hay {len(NOLEG)} informe(s) local(es) no legible(s) (ver CS.PQRSDF.BRECHA); si alguno es este trimestre, el motivo real es ese y no el portal." if NOLEG else ""))
    cand = [ms for (aa, ms) in SEL if aa == a]
    cub, ncub = armar_cobertura(cand)
    COB[a] = (cub, ncub)
    cuartos = {(m - 1) // 3 + 1 for ms in cand for m in ms}
    if cub:
        infs = [SEL[(a, ms)] for ms in cub]
        ET = sum(i[4]["bpms"][1] for i in infs)
        PE = sum(i[4]["bpms"][2] for i in infs)
        RR = sum(i[4]["bpms"][3] for i in infs)
        v = min(100.0, 100.0 * ET / RR)
        ANIO_P[a] = {"v": v, "max": min(100.0, 100.0 * (ET + PE) / RR), "brecha": 100.0 * PE / RR, "ET": ET, "PE": PE, "R": RR, "infs": infs}
        notas_d = [DUP[(a, ms)] for ms in cub if (a, ms) in DUP]
        fila(NM_P, "CS.PQRSDF", "principal", a, ET, RR, v, v, PESO, REF_P, FUENTE_P, "; ".join(i[1] for i in infs), max(i[2] for i in infs), "con dato",
             f"Informes usados (D82): {', '.join(i[0] + ' (' + ETIQ(i[4]['meses']) + ')' for i in infs)}; cubren los 12 meses sin repetirse"
             + (f"; hay {ncub} coberturas posibles y se usa la mas granular" if ncub > 1 else "") + f". Rango: maximo {ANIO_P[a]['max']:.2f} si las {PE} pendientes en terminos se atendieran a tiempo; "
             f"brecha {ANIO_P[a]['brecha']:.2f}. " + (("Duplicados: " + "; ".join(notas_d) + ". ") if notas_d else "") + ALERTA_P)
    else:
        ANIO_P[a] = None
        faltan = [f"T{q}" for q in (1, 2, 3, 4) if q not in cuartos]
        parcial = ""
        if cand:
            parcial = " Con dato parcial (no se suma): " + "; ".join(f"{ETIQ(ms)} = {100 * SEL[(a, ms)][4]['bpms'][1] / SEL[(a, ms)][4]['bpms'][3]:.2f}" for ms in sorted(cand)) + "."
        mot = (f"los informes legibles ({', '.join(ETIQ(ms) for ms in sorted(cand)) or 'ninguno'}) no cubren los 12 meses sin repetirse"
               + (f" (faltan {', '.join(faltan)})" if faltan else "") + f": {MOTIVO_PORTAL}")
        fila_nd(NM_P, "CS.PQRSDF", "principal", a, mot, PESO, REF_P, FUENTE_P,
                "; ".join(sorted({u for q in (1, 2, 3, 4) for u in urls_inventario(a, q)})) or "sin URL en el inventario",
                SEL[sorted(k for k in SEL if k[0] == a)[0]][2] if cand else "", ALERTA_P + parcial)
for a in PERIODO:
    d = ANIO_P[a]
    if d:
        fila(NM_P + ": maximo (si las pendientes en terminos se atienden a tiempo)", "CS.PQRSDF.MAX", "respaldo", a, d["ET"] + d["PE"], d["R"], d["max"], ref="(ET + pendientes) / R",
             fuente=FUENTE_P, fecha=FECHA_CALC, brecha="con dato (rango)", nota=f"Informativo; al indice va el valor de CS.PQRSDF. Brecha = pendientes / R = {d['brecha']:.2f}.")
    else:
        fila_nd(NM_P + ": maximo", "CS.PQRSDF.MAX", "respaldo", a, "faltan informes (portal)", ref="(ET + pendientes) / R", fuente=FUENTE_P, fecha=FECHA_CALC)
# respaldo SAC (solo Educacion; no se suma al principal para no contar dos veces)
FSAC = FUENTE_P.replace("BPMS", "SAC")
for a in PERIODO:
    d = ANIO_P[a]
    if d and all(i[4]["sac"] for i in d["infs"]):
        ETs = sum(i[4]["sac"][0] for i in d["infs"])
        Ts = sum(i[4]["sac"][3] for i in d["infs"])
        fila("PQRSDF, respaldo SAC (Educacion): atendidas en terminos / total", "CS.PQRSDF.SAC", "respaldo", a, ETs, Ts, 100 * ETs / Ts, ref="Tabla 'Atencion en terminos SAC', fila TOTAL",
             fuente=FSAC, fecha=FECHA_CALC, brecha="con dato", nota="Solo Educacion (sistema SAC); NO se suma al principal: las de Educacion que entran por BPMS se trasladan a SAC y se contarian dos veces.")
    else:
        sac23 = SEL.get((2023, (10, 11, 12)))
        extra = ""
        if a == 2023 and sac23 and sac23[4]["sac"]:
            s = sac23[4]["sac"]
            extra = f"2023 T4: atendidas en terminos {s[0]}, extemporaneas {s[1]}, pendientes {s[2]}, total {s[3]} = {100 * s[0] / s[3]:.2f} (p. {sac23[4]['sac_pagina']}); no se suma."
        fila_nd("PQRSDF, respaldo SAC (Educacion): atendidas en terminos / total", "CS.PQRSDF.SAC", "respaldo", a, "faltan informes (portal) o sin tabla SAC legible",
                ref="Tabla 'Atencion en terminos SAC', fila TOTAL", fuente=FSAC, fecha=FECHA_CALC, nota=extra)
# brecha PQRSDF: trimestres sin informe legible que los cubra / 24
nq = sum(1 for k in TRIMESTRES if k not in COBERTO)
fila("Brecha de informacion de PQRSDF (2020-2025)", "CS.PQRSDF.BRECHA", "brecha", "2020-2025", nq, 24, 100.0 * nq / 24, brecha=f"{nq} de 24 trimestres sin informe legible",
     nota="Brecha = trimestres sin informe legible (propio o dentro de un semestre) / 24. El portal no respondio el 19-nov. Anios sin valor: "
          + ", ".join(str(a) for a in PERIODO if ANIO_P[a] is None) + f" ({sum(1 for a in PERIODO if ANIO_P[a] is None)} de 6). Se completa con scripts/construccion/cs_descargas.py y repitiendo este script."
          + (" Informes locales no legibles: " + "; ".join(f"{a} ({m})" for a, m in NOLEG) + "." if NOLEG else ""), fecha=FECHA_CALC)
check("PQRSDF: trimestres sin cubrir = 24 menos los cubiertos por informes legibles (hoy solo 2023 T4)", nq == 24 - len(COBERTO) and (len(SEL) > 1 or nq == 23))
check("PQRSDF anual N/D si los informes no cubren los 12 meses sin repetirse", all((ANIO_P[a] is None) == (COB[a][0] is None) for a in PERIODO))

# ================================================================== DSI (D75)
F_AP, A_AP = "pdm", "116989_apendice91planindicativo27052024.xlsx"
f_ap, u_ap, _ = man(F_AP, A_AP)
wb = openpyxl.load_workbook(RAW / F_AP / A_AP, read_only=True, data_only=True)
rows = list(wb["Apéndice9.1.2."].iter_rows(values_only=True))
FILAS_DSI = [16, 79, 71, 63, 64, 65, 119, 129, 134, 146]
ESPERADO = {16: (99.5, 2022, 99.5), 79: (100.44, 2022, 100.44), 71: (16.3, 2023, 19), 63: (99.73, 2023, 99.9), 64: (99.56, 2023, 99.9), 65: (99.87, 2023, 99.9),
            119: (10.2, 2023, 8), 129: (3.88, 2023, 3.88), 134: (3.7, 2023, 2.7), 146: (75, 2023, 100)}
AP = {}
for n in FILAS_DSI:
    rr = rows[n - 1]
    AP[n] = {"nombre": norm(rr[5]), "unidad": norm(rr[6]), "lb": float(rr[8]), "anio_lb": int(rr[9]), "meta": float(rr[10]), "fuente": norm(rr[13])}
check("Ap. 9.1.2: LB, anio de LB y meta de las 10 metas coinciden con D75 y la ficha",
      all((AP[n]["lb"], AP[n]["anio_lb"], AP[n]["meta"]) == ESPERADO[n] for n in FILAS_DSI), "; ".join(f"{n}: {AP[n]['lb']} ({AP[n]['anio_lb']}) -> {AP[n]['meta']}" for n in FILAS_DSI))
check("DSI: lista de 10 metas (filas 16, 79, 71, 63, 64, 65, 119, 129, 134, 146)", len(FILAS_DSI) == 10 and len(set(FILAS_DSI)) == 10)

# soportes verificados localmente
f_n, u_n, _ = man("dsi/nudc7mev_snapshot", "nudc-7mev_05266.json")
f_m, u_m, _ = man("dsi/nudc7mev_metadatos", "nudc-7mev_views.json")
snap = json.load(open(RAW / "dsi/nudc7mev_snapshot/nudc-7mev_05266.json", encoding="utf-8"))
nudc = {int(x["a_o"]): float(x["cobertura_neta"]) for x in snap}
meta_json = json.load(open(RAW / "dsi/nudc7mev_metadatos/nudc-7mev_views.json", encoding="utf-8"))
import datetime as _dt
rows_upd = _dt.datetime.fromtimestamp(meta_json["rowsUpdatedAt"], _dt.timezone.utc).date().isoformat()
f_ca, u_ca, _ = man("dsi/nudc7mev_created_at", "nudc-7mev_05266_created_at.json")
ca = {int(x["a_o"]): x[":created_at"][:10] for x in json.load(open(RAW / "dsi/nudc7mev_created_at/nudc-7mev_05266_created_at.json", encoding="utf-8"))}
check("MEN nudc-7mev: la fila de 2024 se creo el 2025-11-13 (:created_at), despues del 31-mar-2025 (D79); la de 2023, el 2024-09-06", ca.get(2024) == "2025-11-13" and ca.get(2023) == "2024-09-06" and ca[2024] > "2025-03-31", f"2023 {ca.get(2023)}; 2024 {ca.get(2024)}")
f_ig24, u_ig24, _ = man("informes_gestion_2024", "131997_24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf")
pg59 = norm(pypdf.PdfReader(str(RAW / "informes_gestion_2024/131997_24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf")).pages[58].extract_text())
check("IG 2024, p. 59: contrato ENV-07-09-1191-24 para la ECV y la encuesta de opinion (DAP)", re.search(r"ENV-07-09-\s*1191-24", pg59) is not None and "calidad de vida" in pg59.lower())
f_ig25, u_ig25, _ = man("informes_gestion", "152890_2025informegestiondependencias.pdf")
f_ev, u_ev, _ = man("espacio_verde", "146211_8-dimension-funcional-_-espacio-publico.pdf")
f_bal, u_bal, _ = man("pdm_balance_educacion_2024", "135807_ppt-balance-pdm-2024_educacion_oficial.pdf")
check("MEN nudc-7mev (05266): 2022 = 100,44 (= LB de la fila 79), 2023 = 99,68, 2024 = 98,45; no hay 2025", nudc.get(2022) == 100.44 and nudc.get(2023) == 99.68 and nudc.get(2024) == 98.45 and 2025 not in nudc, str({k: nudc[k] for k in (2022, 2023, 2024)}))
f_td, u_td, _ = man("terridata", "TerriData05266.txt.zip")
with zipfile.ZipFile(RAW / "terridata/TerriData05266.txt.zip") as z:
    td = z.read(z.namelist()[0]).decode("utf-8", "replace").splitlines()
hom = {}
for ln in td:
    x = ln.split("|")
    if len(x) > 11 and x[6] == "Tasa de homicidio intencional por cada 100.000 habitantes":
        hom[int(x[9])] = float(x[7].replace(",", "."))
check("TerriData: tasa de homicidio intencional 2024 = 4,60 y 2025 = 3,00 (DNP con datos de MinDefensa; sin fecha de publicacion)", hom.get(2024) == 4.6 and hom.get(2025) == 3.0, str({k: hom.get(k) for k in (2023, 2024, 2025)}))

# clasificacion por meta y anio (ficha del Analizador, 19-nov): v = verificada en ventana, nv = no verificable, fv = fuera de ventana, sv = sin valor de la definicion y fuente
CL = {
    16: {2024: ("sv", "Sin cifra del aseguramiento del Regimen Subsidiado con el metodo del plan (MinSalud) para la vigencia."), 2025: ("sv", "Sin cifra para la vigencia.")},
    79: {2024: ("fv", f"MEN nudc-7mev: cobertura neta 2024 = {nudc[2024]} (misma definicion, unidad y fuente del PDM; la serie reproduce la LB 100,44 de 2022). La API de datos.gov.co da `:created_at` = {ca[2024]} para la fila de 2024 "
                      f"(la de 2023: {ca[2023]}): publicada despues del 31-mar-2025 (D79). Archivo dsi/nudc7mev_created_at."),
         2025: ("sv", "La serie del MEN llega hasta 2024; no hay valor 2025.")},
    71: {2024: ("sv", "Solo hay conteos de ninos atendidos (Balance PDM 2024, informe de gestion), no la tasa de cobertura 0 a 6 anos del PDM."), 2025: ("sv", "Solo hay conteos de ninos atendidos, no la tasa del PDM.")},
    63: {2024: ("nv", "Fuente del PDM: Encuesta de Calidad de Vida (DAP). El informe de gestion 2024 (p. 59) dice que el contrato ENV-07-09-1191-24 produjo la ECV y la encuesta de opinion, pero no hay informe oficial de la ECV 2024 ni fecha de publicacion (D79)."), 2025: ("sv", "ECV 2025: sin valor (D79).")},
    64: {2024: ("nv", "Fuente del PDM: ECV (DAP). Igual que la fila 63: contrato ENV-07-09-1191-24 (IG 2024, p. 59), sin informe oficial ni fecha (D79)."), 2025: ("sv", "ECV 2025: sin valor (D79).")},
    65: {2024: ("nv", "Fuente del PDM: ECV (DAP). Igual que la fila 63: contrato ENV-07-09-1191-24 (IG 2024, p. 59), sin informe oficial ni fecha (D79)."), 2025: ("sv", "ECV 2025: sin valor (D79).")},
    119: {2024: ("sv", "La tasa de la Secretaria de Movilidad del PDM no esta publicada para 2024; el 7,25 de Medicina Legal (Forensis) es otra fuente: no cuenta."), 2025: ("sv", "Sin valor de la fuente del PDM.")},
    129: {2024: ("fv", "5,75 m2/hab (espacio publico verde urbano 2024, 1.382.762 m2 / 240.655 hab. urbanos) en el documento diagnostico del POT (146211), publicado el 24-oct-2025: despues del 31-mar-2025. ADEMAS la definicion no es la misma: el 3,88 del PDM (Medio Ambiente) tiene metodo sin declarar y ningun documento lo concilia con el 4,66 (AMVA 2022) ni con el 5,75 (decisiones D7, D30)."), 2025: ("sv", "Sin valor 2025.")},
    134: {2024: ("nv", f"TerriData (DNP, con datos de MinDefensa) publica tasa de homicidio intencional 2024 = {hom[2024]:.2f}, sin fecha de publicacion. "
                       f"TerriData 2023 = {hom[2023]:.2f} frente a la LB del PDM 3,7 (Policia Nacional): la diferencia sugiere corte o poblacion distintos."),
          2025: ("nv", f"TerriData: tasa 2025 = {hom[2025]:.2f}, sin fecha de publicacion.")},
    146: {2024: ("sv", "Sin valor observado: las evaluaciones del PDM no reportan este porcentaje."), 2025: ("sv", "Sin valor observado.")},
}
ETQ = {"v": "verificada en ventana", "nv": "no verificable", "fv": "fuera de ventana", "sv": "sin valor de la definicion y fuente"}
DSI = {}
for a in (2024, 2025):
    c = {k: sum(1 for n in FILAS_DSI if CL[n][a][0] == k) for k in ETQ}
    DSI[a] = {"v": c["v"], "nv": c["nv"], "fv": c["fv"], "sv": c["sv"], "min": 100.0 * c["v"] / 10, "max": 100.0 * (c["v"] + c["nv"]) / 10, "brecha": 100.0 * c["nv"] / 10}
check("DSI 2024 (D79): minimo 0, maximo 4/10 = 40 (no verificables 63, 64, 65, 134; fuera de ventana 79 y 129; sin valor 16, 71, 119, 146); brecha 40",
      DSI[2024]["v"] == 0 and DSI[2024]["nv"] == 4 and DSI[2024]["fv"] == 2 and DSI[2024]["sv"] == 4 and [n for n in FILAS_DSI if CL[n][2024][0] == "nv"] == [63, 64, 65, 134]
      and [n for n in FILAS_DSI if CL[n][2024][0] == "fv"] == [79, 129] and DSI[2024]["max"] == 40.0 and DSI[2024]["brecha"] == 40.0)
check("DSI 2025 (D79): minimo 0, maximo 1/10 = 10 (no verificable 134; sin valor 63, 64, 65 y otras 6); brecha 10",
      DSI[2025]["v"] == 0 and DSI[2025]["nv"] == 1 and [n for n in FILAS_DSI if CL[n][2025][0] == "nv"] == [134] and DSI[2025]["sv"] == 9 and DSI[2025]["max"] == 10.0 and DSI[2025]["brecha"] == 10.0)
# Un solo criterio para DSI.TODO (todo lo publicado hasta hoy, sin ventana): la meta cuenta si hay un valor publicado de la fuente que nombra el PDM
# y de la MISMA definicion, y esa definicion se prueba porque la serie publicada reproduce la LB del PDM en el anio de la LB. TerriData solo cuenta si reproduce la LB.
#   79: MEN 2022 = 100,44 = LB -> cuenta (2024 = 98,45).
#  134: TerriData 2023 = 3,50 frente a LB 3,7 (Policia) -> NO reproduce la LB (poblacion distinta) -> no cuenta, igual que la 119 (Forensis 7,25: otra fuente).
#  129: POT 5,75 es espacio verde urbano con metodo AMVA; el 3,88 del PDM tiene metodo sin declarar y ningun documento lo concilia -> no cuenta.
TODO = {2024: [79], 2025: []}
check("DSI.TODO: criterio unico (la serie publicada reproduce la LB del PDM): 79 si; 119, 129 y 134 no", TODO == {2024: [79], 2025: []} and nudc[2022] == AP[79]["lb"] and hom[2023] != AP[134]["lb"])
URL_AP = u_ap
FUENTE_D = "Alcaldia de Envigado, Plan de Desarrollo 2024-2027, Apendice 9.1.2 (lista de metas) y fuentes que nombra el PDM; TerriData (DNP); MEN datos.gov.co nudc-7mev"
REF_D = "10 metas de resultado del PDM (D75); minimo = verificadas en ventana (<= 31-mar t+1, fecha verificable) / 10; semaforo D60"
ALERTA_D = ("ALERTA: DSI mide OPORTUNIDAD de la informacion; las fuentes nacionales con rezago (ECV, MEN, Policia) fallan por diseno aunque el dato exista. "
            "Hoy se publica poco con fecha verificable.")
for a in PERIODO:
    if a in DSI:
        d = DSI[a]
        fila("Disponibilidad de informacion de seguimiento de las metas del PDM (DSI), minimo", "CS.DSI", "principal", a, d["v"], 10, d["min"], d["min"], PESO, REF_D, FUENTE_D, URL_AP, f_ap,
             f"no verificables {d['nv']}/10 = {d['brecha']:.2f}; fuera de ventana {d['fv']}; sin valor de la definicion y fuente {d['sv']}",
             f"Rango: minimo {d['min']:.2f}, maximo {d['max']:.2f} (= minimo + no verificables / 10), brecha {d['brecha']:.2f}. " + ALERTA_D)
        fila("DSI, maximo (minimo + no verificables / 10)", "CS.DSI.MAX", "respaldo", a, d["v"] + d["nv"], 10, d["max"], ref=REF_D, fuente=FUENTE_D, enlace=URL_AP, fecha=f_ap, brecha="con dato (rango)",
             nota="Informativo: no entra al indice ni al semaforo.")
        t = TODO[a]
        fila("DSI, todo lo publicado hasta hoy (sin aplicar la ventana)", "CS.DSI.TODO", "respaldo", a, len(t), 10, 10.0 * len(t), ref="Metas con valor publicado de la misma definicion y fuente del PDM (la serie reproduce la LB), sin ventana",
             fuente=FUENTE_D, enlace=URL_AP, fecha=f_ap, brecha="con dato",
             nota="Metas contadas (filas del Ap. 9.1.2): " + (", ".join(map(str, t)) or "ninguna") + ". Criterio unico: cuenta si hay valor publicado de la fuente del PDM y de la misma definicion, probada porque la serie publicada reproduce la LB del PDM. "
                  "Fila 79 si (MEN 2022 = 100,44 = LB). Fila 119 no (Forensis 7,25 es otra fuente). Fila 134 no (TerriData 2023 = 3,50 frente a LB 3,7 de la Policia: no reproduce la LB). Fila 129 no (POT 5,75 urbano AMVA frente al 3,88 del PDM, metodo sin declarar). Calculo propio; informativo.")
    else:
        fila_nd("Disponibilidad de informacion de seguimiento de las metas del PDM (DSI), minimo", "CS.DSI", "principal", a, "no aplica: la lista de metas es del PDM 2024-2027 (D75: anios 2024 y 2025)", PESO, REF_D,
                FUENTE_D, URL_AP, f_ap, ALERTA_D)
ENL_META = {}
for _a in (2024, 2025):
    ENL_META.update({(16, _a): (URL_AP, f_ap), (119, _a): (URL_AP, f_ap), (146, _a): (URL_AP, f_ap), (79, _a): (u_ca, f_ca), (129, _a): (u_ev, f_ev), (134, _a): (u_td, f_td)})
    for _n in (63, 64, 65):
        ENL_META[(_n, _a)] = (u_ig24, f_ig24) if _a == 2024 else (u_ig25, f_ig25)
ENL_META[(71, 2024)] = (u_bal, f_bal)
ENL_META[(71, 2025)] = (u_ig25, f_ig25)
for n in FILAS_DSI:
    for a in (2024, 2025):
        k, txt = CL[n][a]
        m = AP[n]
        fila(f"DSI meta (Ap. 9.1.2 fila {n}): {m['nombre']} ({m['unidad']})", "CS.DSI.META", "detalle", a, 1 if k == "v" else 0, 1, raw={"valor": ETQ[k]},
             ref=f"LB {m['lb']:g} ({m['anio_lb']}); meta {m['meta']:g}; fuente del PDM: {m['fuente']}", fuente=FUENTE_D, enlace=ENL_META[(n, a)][0],
             fecha=ENL_META[(n, a)][1], brecha=("con dato verificable" if k == "v" else "sin dato verificable"),
             nota=txt + (" Regla: valor de la vigencia con la misma definicion y unidad, de la fuente del PDM, publicado <= 31-mar de t+1 y con fecha verificable." if n == 16 and a == 2024 else ""))
nvt = DSI[2024]["nv"] + DSI[2025]["nv"]
fila("Brecha de informacion del DSI (metas-anio 2024-2025)", "CS.DSI.BRECHA", "brecha", "2024-2025", nvt, 20, 100.0 * nvt / 20, brecha=f"2024: {DSI[2024]['nv']}/10; 2025: {DSI[2025]['nv']}/10 no verificables",
     nota="Brecha = metas no verificables / metas (10 por anio). 2020-2023: no aplica (PDM 2024-2027), 4 de 6 anios del periodo sin DSI. Fuera de ventana: 2 (2024: filas 79 y 129); sin valor de la definicion y fuente: 4 (2024) y 9 (2025), que no son brecha de verificacion sino ausencia del dato.", fecha=f_ap)

# ================================================================== VRC (D33, D77)
def leer_csv(p):
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


CV = leer_csv(CLAS_VRC)
VB = {(r["anio"], r["universo"]): r for r in leer_csv(VRC_BRECHA)}
UP, UA = "u_alcalde_mas_sectoriales_con_acta", "u_amplio_todo_documento"
f_ig, u_ig, _ = man("informes_gestion", "152890_2025informegestiondependencias.pdf")
pg102 = norm(pypdf.PdfReader(str(RAW / "informes_gestion/152890_2025informegestiondependencias.pdf")).pages[101].extract_text())
check("Informe de gestion 2025, p. 102: audiencia publica del Alcalde realizada el 13-nov-2025 (537 ciudadanos)",
      "Audiencia pública realizada el 13 de noviembre de 2025" in pg102 and "537 ciudadanos" in pg102)
FECHA_A25 = "2025-11-13"
# la cifra '10 de diciembre de 2025' del cuadro de preguntas es la fecha de entrega de una respuesta ('ayer'), no de la audiencia
VRC = {}
for a in ("2024", "2025"):
    ev = [r for r in CV if r["anio_evento"] == a and r[UP] == "1"]
    si = sum(r["mencion_si"] == "1" for r in ev)
    nd = sum(r["mencion_desconocida"] == "1" for r in ev)
    no = sum(r["mencion_no"] == "1" for r in ev)
    n = len(ev)
    VRC[a] = {"n": n, "si": si, "nd": nd, "no": no, "min": 100.0 * si / n, "max": 100.0 * (si + nd) / n, "brecha": 100.0 * nd / n}
    b = VB[(a, UP)]
    check(f"VRC {a}: calculo propio = vrc_brecha.csv ({b['cota_inferior_pct']} / {b['cota_superior_pct']} / {b['brecha_informacion_pct']})",
          abs(VRC[a]["min"] - float(b["cota_inferior_pct"])) < 0.01 and abs(VRC[a]["max"] - float(b["cota_superior_pct"])) < 0.01 and abs(VRC[a]["brecha"] - float(b["brecha_informacion_pct"])) < 0.01)
check("VRC 2024 = 66,67 / 100 / 33,33 y 2025 = 20 / 80 / 60", [round(VRC["2024"][k], 2) for k in ("min", "max", "brecha")] == [66.67, 100.0, 33.33] and [round(VRC["2025"][k], 2) for k in ("min", "max", "brecha")] == [20.0, 80.0, 60.0])
AM = {}
for a in ("2024", "2025"):
    ev = [r for r in CV if r["anio_evento"] == a and r[UA] == "1"]
    si = sum(r["mencion_si"] == "1" for r in ev)
    nd = sum(r["mencion_desconocida"] == "1" for r in ev)
    AM[a] = {"n": len(ev), "si": si, "nd": nd, "min": 100.0 * si / len(ev), "max": 100.0 * (si + nd) / len(ev), "brecha": 100.0 * nd / len(ev)}
check("VRC amplio: 2024 50-100 (brecha 50) y 2025 28,57-85,71 (brecha 57,14)", [round(AM["2024"][k], 2) for k in ("min", "max", "brecha")] == [50.0, 100.0, 50.0] and [round(AM["2025"][k], 2) for k in ("min", "max", "brecha")] == [28.57, 85.71, 57.14])
# fechas de consulta: los documentos de eventos 2024-2025 se consultaron para la clasificacion (15-nov-2026, manifiestos de rendicion_cuentas e informes_gestion)
F_VRC = f_ig
REF_V = "eventos con veeduria documentada / eventos (D33); desconocida cuenta como no documentada en el minimo; semaforo D60"
FUENTE_V = "Alcaldia de Envigado: actas, memorias y cuadros de preguntas de los eventos de rendicion de cuentas (clasificacion en outputs/indicadores/clasificacion_vrc.csv)"
ALERTA_V = ("ALERTA: 'con veeduria' solo si el documento lo dice expresamente (una invitacion no cuenta); 'desconocida' cuenta como no documentada en el minimo. "
            "La audiencia del Alcalde nunca publica lista de asistentes. Sin nombres de personas.")
OTRAS = ("Otras rendiciones sectoriales con acta o memoria (Discapacidad, Victimas, etc.)", "pendiente",
         "sin dato: pendiente de busqueda. El inventario no tiene rendiciones de Discapacidad ni de Victimas de 2021-2023 (solo desde 2024); buscarlas en la API del portal cuando vuelva (D81). Clasificarlas es manual", None)
PEND_V = {
    2021: [("Audiencia publica de rendicion de cuentas del Alcalde (29-nov-2021)", "desconocida", "23449 (rendicion_cuentas/, pp. 1-6): cuadro de preguntas y respuestas sin lista de asistentes ni mencion de veedurias (verificado: 0 paginas con 'veedur')", "23449"),
           ("Acta de rendicion de cuentas con evaluacion, feb-2021 (documento 6384)", "pendiente", "sin dato: pendiente de descarga (D81); el 6385, presentacion de Salud ya local, no menciona veedurias pero no decide por ser presentacion", "6384"),
           ("Rendicion de cuentas Secretaria de Salud 2021 (documento 6387)", "pendiente", "sin dato: pendiente de descarga", "6387"),
           ("Acta de rendicion de cuentas 16-dic-2021, 'evaluacion' (documento 24141)", "pendiente", "sin dato: pendiente de descarga; falta confirmar si es evento de Salud o del Alcalde", "24141"), OTRAS],
    2022: [("Audiencia publica de rendicion de cuentas del Alcalde 2022", "desconocida", "92036 (rendicion_cuentas/, 12 pp.): informe del evento de dialogo sin lista de asistentes ni mencion de veedurias (verificado: 0 paginas con 'veedur')", "92036"),
           ("Acta de rendicion de cuentas de Salud 2022-01 (documento 34258)", "pendiente", "sin dato: pendiente de descarga", "34258"),
           ("Acta de rendicion de cuentas de Salud 2022-02, 14-dic-2022 (documento 55874)", "pendiente", "sin dato: pendiente de descarga", "55874"), OTRAS],
    2023: [("Audiencia publica de rendicion de cuentas del Alcalde 2023", "desconocida", "115926 (rendicion_cuentas/, 13 pp.): respuestas a ciudadanos sin lista de asistentes ni mencion de veedurias (verificado: 0 paginas con 'veedur')", "115926"),
           ("Acta de rendicion de cuentas de Salud 2023-01 (documento 111982)", "pendiente", "sin dato: pendiente de descarga", "111982"),
           ("Acta de rendicion de cuentas de Salud 2023-02 (documento 111981)", "pendiente", "sin dato: pendiente de descarga", "111981"), OTRAS],
}
ENL_V = {2021: ("rendicion_cuentas", "23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf"), 2022: ("rendicion_cuentas", "92036_rendicion-de-cuentas-2022.pdf"),
         2023: ("rendicion_cuentas", "115926_preguntasyrespuestasrendiciondecuentas2023.pdf")}
# comprobacion: los tres documentos del Alcalde no mencionan veedurias
for a, (c, f) in ENL_V.items():
    t = [(p.extract_text() or "") for p in pypdf.PdfReader(str(RAW / c / f)).pages]
    man(c, f)
    check(f"Documento del Alcalde {a} ({f.split('_')[0]}) sin mencion de veedurias", not any(re.search(r"(?i)veedur", x) for x in t))
for a in PERIODO:
    k = str(a)
    if k in VRC:
        d = VRC[k]
        fila("Veedurias en la rendicion de cuentas (VRC), minimo", "CS.VRC", "principal", a, d["si"], d["n"], d["min"], d["min"], PESO, REF_V, FUENTE_V, "ver CS.VRC.EVENTO", F_VRC,
             f"desconocida {d['nd']}/{d['n']} = {d['brecha']:.2f}",
             f"Universo principal: audiencia del Alcalde + sectoriales con acta o memoria: {d['n']} eventos; con veeduria {d['si']}, sin veeduria {d['no']}, desconocida {d['nd']}. "
             f"Rango: minimo {d['min']:.2f}, maximo {d['max']:.2f}, brecha {d['brecha']:.2f}. " + ALERTA_V)
        fila("VRC, maximo (documentada + desconocida) / eventos", "CS.VRC.MAX", "respaldo", a, d["si"] + d["nd"], d["n"], d["max"], ref=REF_V, fuente=FUENTE_V, fecha=F_VRC, brecha="con dato (rango)",
             nota="Informativo: no entra al indice ni al semaforo.")
        e = AM[k]
        fila("VRC amplio (todos los eventos con algun documento), minimo", "CS.VRC.AMPLIO", "respaldo", a, e["si"], e["n"], e["min"], ref="Sensibilidad (D33): todo evento de rendicion con documento publicado", fuente=FUENTE_V,
             fecha=F_VRC, brecha=f"desconocida {e['nd']}/{e['n']} = {e['brecha']:.2f}", nota=f"Rango {e['min']:.2f} a {e['max']:.2f}; brecha {e['brecha']:.2f}. Informativo.")
    elif a in PEND_V:
        ev = [x for x in PEND_V[a] if x[3]]
        nd_loc = sum(1 for x in ev if x[1] == "desconocida")
        fila_nd("Veedurias en la rendicion de cuentas (VRC), minimo", "CS.VRC", "principal", a,
                f"pendiente de descarga: {sum(1 for x in ev if x[1] == 'pendiente')} de {len(ev)} eventos sin documento local (D77; cs_descargas.py)", PESO, REF_V, FUENTE_V, "ver CS.VRC.EVENTO", F_VRC,
                f"Eventos previstos con documento identificado: {len(ev)} (audiencia del Alcalde + actas de Salud del inventario; falta confirmar si cada documento es un evento), mas las rendiciones de Discapacidad, Victimas u otras aun no identificadas (D81; el 40433 no es un evento). "
                f"Alcalde desconocida (documento local); {len(ev) - nd_loc} pendientes. Con lo local el minimo posible va de 0 a {100 * (len(ev) - nd_loc) / len(ev):.2f} y el maximo es 100. " + ALERTA_V)
    else:
        fila_nd("Veedurias en la rendicion de cuentas (VRC), minimo", "CS.VRC", "principal", a, "sin eventos clasificados: D77 extiende VRC a 2021-2023; 2020 queda sin dato", PESO, REF_V, FUENTE_V, "", F_VRC,
                "El unico documento local de 2020 (6385) es de la Secretaria de Salud y no menciona veedurias; no se encontro la rendicion del Alcalde de 2020. " + ALERTA_V)
# eventos
for r in sorted(CV, key=lambda r: (r["anio_evento"], r["fecha_evento"] or "9")):
    if r["anio_evento"] not in ("2024", "2025"):
        continue
    cls = "con veeduria" if r["mencion_si"] == "1" else "sin veeduria" if r["mencion_no"] == "1" else "desconocida"
    fecha_ev = r["fecha_evento"]
    nota_f = r["fecha_evento_nota"]
    if r["id"] == "CS-VRC-2025-006":
        fecha_ev = FECHA_A25
        nota_f = "Fecha completada: audiencia del Alcalde el 13-nov-2025 segun el informe de gestion 2025, p. 102 (537 ciudadanos); verificada en el PDF. El cuadro de preguntas menciona el 10-dic-2025 como 'ayer' solo para la entrega de una respuesta. clasificacion_vrc.csv sigue sin fecha."
    fila(f"VRC evento {r['id']}: {r['evento']}", "CS.VRC.EVENTO", "detalle", int(r["anio_evento"]), raw={"numerador": "1" if cls == "con veeduria" else "0", "denominador": "1", "valor": cls},
         ref=f"fecha del evento {fecha_ev or 'sin fecha'}; universo principal: {'si' if r[UP] == '1' else 'no (solo amplio)' if r[UA] == '1' else 'no'}", fuente=f"{r['documento']} ({r['pagina']})", enlace=r["enlace"], fecha=F_VRC,
         brecha="con dato" if cls != "desconocida" else "sin dato: desconocida", nota=(nota_f + " " if nota_f else "") + f"Tipo de documento: {r['tipo_documento']}.")
INV_DOC = {}
for _r in _inv:
    _m = re.search(r"/(\d+)_", _r["url_origen"])
    if _m:
        INV_DOC.setdefault(_m.group(1), _r["url_origen"])
MAN_ALC = {"23449": "23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf", "92036": "92036_rendicion-de-cuentas-2022.pdf", "115926": "115926_preguntasyrespuestasrendiciondecuentas2023.pdf"}
for a, ev in PEND_V.items():
    for nombre, cls, txt, docid in ev:
        if cls == "desconocida":
            _f, _u, _ = man("rendicion_cuentas", MAN_ALC[docid])
        else:
            _f, _u = "", (INV_DOC.get(docid, "sin URL en el inventario (busqueda pendiente en la API del portal)") if docid else "sin URL en el inventario (busqueda pendiente en la API del portal)")
        fila(f"VRC evento {a}: {nombre}", "CS.VRC.EVENTO", "detalle", a, raw={"numerador": "0" if cls == "desconocida" else "N/D", "denominador": "1", "valor": cls if cls == "desconocida" else "N/D"},
             ref="universo principal previsto", fuente="Inventario data_raw_IIC/inventario_detallado_IIC.csv; " + ("documento local" if cls == "desconocida" else "documento no descargado"),
             enlace=_u, fecha=_f, brecha="sin dato: desconocida" if cls == "desconocida" else "sin dato: pendiente de descarga", nota=txt)
sinV = [a for a in PERIODO if str(a) not in VRC]
fila("Brecha de informacion de VRC (2020-2025)", "CS.VRC.BRECHA", "brecha", "2020-2025", len(sinV), 6, 100.0 * len(sinV) / 6, brecha=f"{len(sinV)} de 6 anios sin valor (2020-2023); eventos desconocidos 2024: 1/3, 2025: 3/5",
     nota="Anios sin valor: 2020 (sin eventos) y 2021-2023 (pendiente de descarga, D77). Dentro de 2024 y 2025 la brecha es desconocidos / eventos = 33,33 y 60,00. Se completa con cs_descargas.py y clasificando las actas.", fecha=F_VRC)

# ================================================================== RPO (D76)
REF_R = "observaciones (sin felicitaciones) de la audiencia del Alcalde con respuesta de fondo publicada al 31-dic (busqueda hasta 31-mar t+1) / observaciones (D76); semaforo D60"
ALERTA_R = "ALERTA: una respuesta de fondo no es un acuse ni una remision; los cuadros son los que publica la Alcaldia (no se sabe si registran todas las intervenciones). Sin nombres de ciudadanos."
CR = leer_csv(CLAS_RPO)
RPO = {}
for a in ("2024", "2025"):
    ot = [r for r in CR if r["anio"] == a and r["en_universo_alcalde"] == "1" and r["en_ot_excluir_felicitaciones"] == "1"]
    orr = [r for r in ot if r["fondo_estricto"] == "1" and r["publicada_al_31dic"] == "1"]
    RPO[int(a)] = (len(orr), len(ot))
check("RPO 2024 = 16/17 = 94,12 (129738)", RPO[2024] == (16, 17) and abs(100 * 16 / 17 - 94.12) < 0.005, str(RPO[2024]))
check("RPO 2025 = 1/2 = 50 (149610; cuadro corto)", RPO[2025] == (1, 2), str(RPO[2025]))
# 2021: 23449, pp. 1-6
lec21 = pypdf.PdfReader(str(RAW / "rendicion_cuentas/23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf"))
t21 = " ".join(norm(p.extract_text()) for p in lec21.pages)
nums = sorted({int(x) for x in re.findall(r"(?<!\d)(\d{1,2})\. ¿", t21)})
nfel = t21.count("N/A")
check("RPO 2021: 13 preguntas numeradas + 4 felicitaciones/agradecimientos = 17 intervenciones (23449, pp. 1-6)", nums == list(range(1, 14)) and nfel == 4 and len(lec21.pages) == 6, f"preguntas {nums[-1]}, N/A {nfel}")
RPO[2021] = (11, 13)
RPOMAX = {2021: (13, 13)}
check("RPO 2021 (D80): rango 11/13 = 84,62 a 13/13 = 100", abs(100 * RPO[2021][0] / RPO[2021][1] - 84.62) < 0.005 and RPOMAX[2021] == (13, 13))
RPO[2023] = (0, 6)
t23 = [norm(p.extract_text()) for p in pypdf.PdfReader(str(RAW / "rendicion_cuentas/115926_preguntasyrespuestasrendiciondecuentas2023.pdf")).pages]
c23 = sorted({int(x) for x in re.findall(r"(?<!\d)(\d{1,2})\. Ciudadan[oa]", " ".join(t23))})
check("RPO 2023: 6 observaciones numeradas (115926)", c23 == [1, 2, 3, 4, 5, 6])
check("RPO 2023: 115926, p. 13, dice 'el pasado 23 de enero' (de 2024): escrito despues del 23-ene-2024", len(t23) == 13 and "el pasado 23 de enero" in t23[12])
t22 = pypdf.PdfReader(str(RAW / "rendicion_cuentas/92036_rendicion-de-cuentas-2022.pdf")).pages
check("RPO 2022: 92036 pp. 8-11 son respuestas sin lista de preguntas (p. 8 solo el titulo 'Respuestas a preguntas')", len(t22) == 12 and norm(t22[7].extract_text()) == "Respuestas a preguntas")
RC = {k: man("rendicion_cuentas", f) for k, f in {
    "21": "23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf", "22": "92036_rendicion-de-cuentas-2022.pdf",
    "23": "115926_preguntasyrespuestasrendiciondecuentas2023.pdf", "24": "129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf",
    "25": "149610_preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf"}.items()}
DOC40433 = INV_DOC.get("40433", "")                                   # respuestas escritas de jul-2022 (aun no descargado)
FUENTE_R = "Alcaldia de Envigado: cuadro de preguntas y respuestas de la audiencia publica de rendicion de cuentas del Alcalde"
for a in PERIODO:
    if a == 2020:
        fila_nd("Respuesta a observaciones ciudadanas (RPO)", "CS.RPO", "principal", a, "no se encontro la rendicion del Alcalde de 2020 (6385 es de la Secretaria de Salud)", PESO, REF_R, FUENTE_R, "", "", ALERTA_R)
    elif a == 2022:
        fila_nd("Respuesta a observaciones ciudadanas (RPO)", "CS.RPO", "principal", a, "92036 (pp. 8-11) trae solo respuestas, sin la lista de preguntas (no hay denominador); ademas se publico el 31-ago-2023, fuera de la ventana",
                PESO, REF_R, FUENTE_R, RC["22"][1], RC["22"][0], ALERTA_R)
    else:
        n, d = RPO[a]
        k = str(a)[2:]
        extra = {
            2021: "RANGO (D80): minimo 11/13 = 84,62 (entra al indice) y maximo 13/13 = 100; brecha 2/13 = 15,38. 13 preguntas numeradas (17 intervenciones menos 4 felicitaciones/agradecimientos) en 23449, pp. 1-6. "
                  "El minimo pone en 0 las preguntas 4 y 9: el Alcalde no las respondio en el evento ('si no permiten esta pregunta se la vamos a contestar...'; 'tomaremos atenta nota') y se contestaron por escrito en el documento 40433 (jul-2022). "
                  "ALERTA de publicacion: el 23449 se publico entre el 29-nov-2021 (evento) y el ~27-ene-2022 (cota: los ids 27514-27525 son de una pagina del 2022-01-27); NO esta probado que fuera <= 31-dic-2021, asi que, si la regla de ventana se aplicara al pie de la letra, "
                  "podria ser 0/13. La fecha de pagina del inventario (2021-09-08) es anterior al evento y no sirve. Se verifica con la API del portal cuando vuelva.",
            2023: "6 observaciones, todas con texto de respuesta (la 3 es una respuesta general porque el ciudadano no dejo contacto). Las respuestas se publicaron DESPUES del 31-mar-2024: el propio 115926 (p. 13) dice 'el pasado 23 de enero' de 2024 al contar tramites de la IUE (escrito despues del 23-ene-2024), "
                  "viene despues del 115597 (publicado el 30-abr-2024) y la pagina del inventario es del 2024-04-30; no hay fecha exacta por archivo [S]. Asi, dentro de la ventana el numerador es 0. Ver CS.RPO.TODO (5 o 6 de 6 sin ventana). Se verifica con la API cuando vuelva el portal.",
            2024: "17 observaciones sin felicitaciones; 16 con respuesta de fondo publicada en la ventana (129738, pagina del portal 2024-11-21, documento 2024-12-12).",
            2025: "ALERTA (D76): el cuadro publicado es corto, 3 intervenciones frente a 537 asistentes; 1 felicitacion, 2 observaciones, 1 con respuesta de fondo (la otra es respuesta que remite a otra entidad). "
                  "Valor frágil: el cuadro puede no registrar todas las intervenciones."}[a]
        fila("Respuesta a observaciones ciudadanas (RPO)", "CS.RPO", "principal", a, n, d, 100.0 * n / d, 100.0 * n / d, PESO, REF_R, FUENTE_R, RC[k][1] + ("; " + DOC40433 if a == 2021 else ""), RC[k][0], "con dato (rango: brecha 2/13 = 15,38)" if a == 2021 else "con dato" if a != 2025 else "con dato (con alerta)",
             extra + " " + ALERTA_R)
        mx = RPOMAX.get(a, (n, d))
        fila("RPO, maximo (rango D80)", "CS.RPO.MAX", "respaldo", a, mx[0], mx[1], 100.0 * mx[0] / mx[1], ref=REF_R, fuente=FUENTE_R, enlace=RC[k][1], fecha=RC[k][0], brecha=f"con dato (rango): {100.0 * (mx[0] - n) / mx[1]:.2f}" if a in RPOMAX else "con dato",
             nota="Informativo: no entra al indice. Sin rango distinto del minimo salvo 2021 (D80).")
fila("RPO, sin aplicar la ventana de publicacion (todo lo publicado hasta hoy)", "CS.RPO.TODO", "respaldo", 2023, 5, 6, 100.0 * 5 / 6, ref="Respuestas de fondo publicadas hasta hoy / observaciones (rango 5 a 6 de 6)",
     fuente=FUENTE_R, enlace=RC["23"][1], fecha=RC["23"][0], brecha="con dato (rango)",
     nota="Rango 5/6 = 83,33 (criterio estricto: la observacion 3 recibe respuesta parcial: Enviaseo pide contacto y Seguridad responde solo una parte) a 6/6 = 100,00 (criterio laxo). Informativo: no entra al indice. En 2021, 2024 y 2025 la ventana no cambia el resultado.")
sinR = [a for a in PERIODO if a in (2020, 2022)]
fila("Brecha de informacion de RPO (2020-2025)", "CS.RPO.BRECHA", "brecha", "2020-2025", len(sinR), 6, 100.0 * len(sinR) / 6, brecha=f"{len(sinR)} de 6 anios sin valor (2020, 2022)",
     nota="2020: no se encontro la rendicion del Alcalde. 2022: solo respuestas, sin lista de preguntas, y publicadas fuera de la ventana. 2023 tiene valor 0 dentro de la ventana (no es brecha). 2025 con alerta de cuadro corto.", fecha=RC["24"][0])

# ================================================================== ICRCS
def pt(comp, a):
    if comp == "PQRSDF":
        return ANIO_P[a]["v"] if ANIO_P[a] else None
    if comp == "DSI":
        return DSI[a]["min"] if a in DSI else None
    if comp == "VRC":
        return VRC[str(a)]["min"] if str(a) in VRC else None
    if comp == "RPO":
        return 100.0 * RPO[a][0] / RPO[a][1] if a in RPO else None


MOT = {"PQRSDF": "falta PQRSDF (portal)", "DSI": "falta DSI (no aplica antes de 2024)", "VRC": "falta VRC", "RPO": "falta RPO"}
ICR, FALTAN, RANGO, RANGO_MAX = {}, {}, {}, {}
REF_I = "(DSI + PQRSDF + VRC + RPO) / 4, peso 0,25 cada uno (D45); VRC, RPO y DSI al minimo; sin redistribuir pesos; semaforo D60"
for a in PERIODO:
    falt = [c for c in ("DSI", "PQRSDF", "VRC", "RPO") if pt(c, a) is None]
    FALTAN[a] = falt
    comp = {c: pt(c, a) for c in ("DSI", "PQRSDF", "VRC", "RPO")}
    con = "; ".join(f"{c} {v:.2f}" for c, v in comp.items() if v is not None)
    if falt:
        ICR[a] = None
        rango = ""
        if a in (2024, 2025) and falt == ["PQRSDF"]:
            suma = sum(v for v in comp.values() if v is not None)
            RANGO[a] = (suma / 4, (suma + 100) / 4)
            rango = f" Rango posible segun PQRSDF (P entre 0 y 100): {suma / 4:.2f} (P = 0) a {(suma + 100) / 4:.2f} (P = 100); formula ({comp['DSI']:.2f} + P + {comp['VRC']:.2f} + {comp['RPO']:.2f}) / 4."
        mot = ", ".join(MOT[c] for c in falt)
        fila_nd("ICRCS, indice de control social", "CS.ICRCS", "indice", a, mot, "", REF_I, "Calculo propio a partir de los 4 indicadores", "", FECHA_CALC,
                f"N/D: faltan {', '.join(falt)}. No se redistribuyen pesos. Componentes con valor: {con or 'ninguno'}.{rango}")
    else:
        ICR[a] = sum(comp.values()) / 4
        fila("ICRCS, indice de control social", "CS.ICRCS", "indice", a, v=ICR[a], puntaje=ICR[a], ref=REF_I, fuente="Calculo propio a partir de los 4 indicadores", fecha=FECHA_CALC, brecha="con dato",
             nota="Puntajes: " + con)
# maximos
for a in PERIODO:
    mx = {"DSI": DSI[a]["max"] if a in DSI else None, "VRC": VRC[str(a)]["max"] if str(a) in VRC else None, "RPO": (100.0 * RPOMAX[a][0] / RPOMAX[a][1]) if a in RPOMAX else pt("RPO", a),
          "PQRSDF": ANIO_P[a]["max"] if ANIO_P[a] else None}
    falt = [c for c in ("DSI", "PQRSDF", "VRC", "RPO") if mx[c] is None]
    if falt:
        rango = ""
        if a in (2024, 2025) and falt == ["PQRSDF"]:
            suma = sum(v for v in mx.values() if v is not None)
            RANGO_MAX[a] = (suma / 4, (suma + 100) / 4)
            rango = f" Rango posible segun el maximo de PQRSDF (entre 0 y 100): {suma / 4:.2f} a {(suma + 100) / 4:.2f}; formula ({mx['DSI']:.2f} + P + {mx['VRC']:.2f} + {mx['RPO']:.2f}) / 4."
        fila_nd("ICRCS con los maximos de los componentes", "CS.ICRCS.MAX", "respaldo", a, ", ".join(MOT[c] for c in falt), "", "DSI, VRC, RPO y PQRSDF en su maximo",
                "Calculo propio", "", FECHA_CALC, f"N/D: faltan {', '.join(falt)}.{rango} Informativo: no entra a la serie ni al semaforo.")
    else:
        v = sum(mx.values()) / 4
        fila("ICRCS con los maximos de los componentes", "CS.ICRCS.MAX", "respaldo", a, v=v, ref="DSI, VRC, RPO y PQRSDF en su maximo", fuente="Calculo propio", fecha=FECHA_CALC,
             brecha="con dato (rango)", nota="Informativo.")
conI = [a for a in PERIODO if ICR[a] is not None]
sinI = [a for a in PERIODO if ICR[a] is None]
fila("Brecha de informacion del ICRCS (2020-2025)", "CS.ICRCS.BRECHA", "brecha", "2020-2025", len(sinI), 6, 100.0 * len(sinI) / 6, brecha=f"{len(sinI)} de 6 anios sin indice",
     nota="Sin indice: " + "; ".join(f"{a} (faltan {', '.join(FALTAN[a])})" for a in sinI) + ". Brecha = anios sin indice / anios del periodo. 2024 y 2025 se calculan en cuanto lleguen los informes de PQRSDF.", fecha=FECHA_CALC)
check("ICRCS N/D en 2020-2025 mientras falten informes de PQRSDF (solo 2023 T4 local)", len(SEL) > 1 or conI == [])
check("ICRCS 2024 y 2025 solo faltan por PQRSDF (si no llegaron los informes)", len(SEL) > 1 or (FALTAN[2024] == ["PQRSDF"] and FALTAN[2025] == ["PQRSDF"]), str(FALTAN))
check("ICRCS 2020-2023 sin DSI", all("DSI" in FALTAN[a] for a in (2020, 2021, 2022, 2023)))
check("Pesos: 4 componentes x 0,25 = 1", abs(4 * PESO - 1) < 1e-12)
# el rango del ICRCS se compara con lo calculado de los componentes (no con constantes)
if FALTAN[2024] == ["PQRSDF"] and FALTAN[2025] == ["PQRSDF"]:
    for a_ in (2024, 2025):
        rpo_ = 100.0 * RPO[a_][0] / RPO[a_][1]
        lo = (DSI[a_]["min"] + 0 + VRC[str(a_)]["min"] + rpo_) / 4
        hi = (DSI[a_]["min"] + 100 + VRC[str(a_)]["min"] + rpo_) / 4
        check(f"Rango ICRCS {a_} (P entre 0 y 100) = (DSI + P + VRC + RPO) / 4 calculado de los componentes", abs(RANGO[a_][0] - lo) < 1e-9 and abs(RANGO[a_][1] - hi) < 1e-9, f"{lo:.2f} a {hi:.2f}")
        lo = (DSI[a_]["max"] + 0 + VRC[str(a_)]["max"] + rpo_) / 4
        hi = (DSI[a_]["max"] + 100 + VRC[str(a_)]["max"] + rpo_) / 4
        check(f"Rango de ICRCS.MAX {a_} calculado de los maximos", abs(RANGO_MAX[a_][0] - lo) < 1e-9 and abs(RANGO_MAX[a_][1] - hi) < 1e-9, f"{lo:.2f} a {hi:.2f}")

# ------------------------------------------------------------------ controles finales y salida
check("Ninguna celda contiene 'Ciudadan[oa]:' (sin nombres de ciudadanos)", not any(re.search(r"Ciudadan[oa]:", str(v)) for f in FILAS for v in f.values()))
check("Todas las filas son Provisional y la fecha_consulta sale de manifiestos o del calculo", all(f["estado"] == PROV for f in FILAS))
check("Decimales con punto (sin comas en numerador, denominador, valor ni puntaje)", all("," not in f[c] for f in FILAS for c in ("numerador", "denominador", "valor", "puntaje_0_100")))
print("\nPQRSDF:", {a: (fmt(ANIO_P[a]["v"]) if ANIO_P[a] else "N/D") for a in PERIODO})
print("DSI:", {a: (DSI[a]["min"], DSI[a]["max"]) for a in DSI}, "| VRC:", {a: (fmt(v["min"]), fmt(v["max"]), fmt(v["brecha"])) for a, v in VRC.items()})
print("RPO:", {a: (n, d) for a, (n, d) in sorted(RPO.items())}, "| ICRCS:", {a: (fmt(v) if v is not None else "N/D") for a, v in ICR.items()})
fallos = [n for n, ok, _ in ASSERTS if not ok]
print("filas:", len(FILAS), "->", SALIDA, "| asserts:", len(ASSERTS), "| fallos:", len(fallos))
if fallos:
    print("ASSERTS FALLIDOS:", fallos, "-> la salida NO se escribio (se conserva la anterior)")
    sys.exit(1)
SALIDA.parent.mkdir(parents=True, exist_ok=True)
with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
    w.writeheader()
    w.writerows(FILAS)
print("escrito:", SALIDA)
