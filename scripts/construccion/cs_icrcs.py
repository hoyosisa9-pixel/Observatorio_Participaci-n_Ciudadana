#!/usr/bin/env python3
"""Dimension Control social (ICRCS): PQRSDF, DSI, VRC y RPO.

ICRCS = (DSI + PQRSDF + VRC + RPO) / 4, peso 0,25 cada uno (D45). VRC, RPO y DSI entran con su MINIMO (D33, D75). No se redistribuyen pesos:
si falta un componente el anio queda N/D. Semaforo D60: >= 90 Verde, >= 75 Amarillo, >= 50 Naranja, < 50 Rojo. Periodo 2020-2025.

Indicadores
  CS.PQRSDF (D45, D78, D82, D86) 100 * ET / R con tope 100; ET = 'atendida en terminos' (u 'oportunas'), R = 'Total' de la fila Total de la tabla de oportunidad del sistema general
                       (BPMS hasta 2025 T1 y Debora desde 2025 T2) del informe trimestral. El anio se arma con informes que cubran los 12 meses sin repetirse (trimestres o semestres;
                       si hay duplicados se usa el id mas reciente). Rango: maximo = (ET + pendientes en termino + 'sin respuesta') / R; brecha = (pendientes + sin respuesta) / R
                       (la columna 'sin respuesta' de 2020 T4-2021 T4 va al maximo, D86). Respaldo SAC (solo Educacion; no se suma). Lector de PDF generico con una lectura manual
                       (2023 T1) y alertas de comparabilidad (Decreto 491 hasta el 17-may-2022, cambio de universo de 2024, sistema Debora desde mayo de 2025).
  CS.DSI    (D75, D79) 10 metas del Apendice 9.1.2; una meta cuenta en t si hay valor de la vigencia t con la misma definicion y unidad, publicado por la
                       fuente que nombra el PDM, a mas tardar el 31-mar de t+1 y con fecha de publicacion verificable. Minimo = verificadas / 10;
                       maximo = minimo + no verificables / 10; brecha = no verificables / 10.
  CS.VRC    (D33, D77, D81) eventos de rendicion de cuentas con veeduria documentada (minimo), + desconocidos (maximo); 2021-2023 con la audiencia del Alcalde y las actas de Salud (clasificadas a mano).
  CS.RPO    (D76, D87) observaciones sin felicitaciones de la audiencia del Alcalde con respuesta de fondo publicada al 31-dic (busqueda hasta 31-mar t+1); 2021 como rango 0-100.

Insumos (data_raw_IIC/07_revision_nov2026/, cada carpeta con manifiesto.csv; el script verifica bytes y sha256):
  pqrs/                         24 informes trimestrales de PQRSDF 2020-2025 y cuadros de apoyo (obtenidos del ZIP original de Google Drive de la usuaria, con su autorizacion, porque el portal dio HTTP 500; D85)
  pdm/116989_...xlsx            Apendice 9.1.2 (LB y meta de las 10 metas del DSI)
  dsi/                          instantanea, metadatos y :created_at de MEN nudc-7mev (fila 79)
  men_nudc7mev/, terridata/     cobertura neta (2011-2024) y tasa de homicidio (TerriData 2024 y 2025)
  rendicion_cuentas/            cuadros del Alcalde 2021-2025 (RPO y VRC), actas de Salud 2020-2023 y encuestas 40433 y 40434
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
    """Meses (1-12) que nombra un texto, con limites de palabra (mayor no es mayo, genero no es enero) y rangos 'de enero a junio' o 'entre enero y marzo'."""
    t = sin_acentos(texto).lower()
    out = set()
    for a, b in re.findall(rf"\b({_MES_RE})\s+(?:a|hasta|al mes de|y)\s+({_MES_RE})\b", t):
        if MESES[a] <= MESES[b] and "entre " + a in t or re.search(rf"\b{a}\s+(?:a|hasta|al mes de)\s+{b}\b", t):
            out.update(range(MESES[a], MESES[b] + 1))
    out.update(MESES[m] for m in re.findall(rf"\b({_MES_RE})\b", t))
    return sorted(out)


# columnas de las tablas de oportunidad (el nombre cambia entre plantillas; el papel no)
_TOK = r"\b(atendidas? en terminos?|atendidas? extempora\w*(?:\s+\w*mente)?|inoportunas?|oportunas?|pendientes?|por responder|por atender|vencidas?|sin respuesta|total,? general|total|atendidas?)\b"
_ROL = [(r"atendidas? en termino", "ET"), (r"oportuna", "ET"), (r"inoportuna", "LATE"), (r"extempora", "LATE"), (r"pendiente", "PEND"), (r"por responder", "PEND"),
        (r"por atender", "PEND"), (r"vencida", "VENC"), (r"sin respuesta", "SINR"), (r"total", "TOT"), (r"atendida", "LATE")]


def _rol(tok):
    for pat, rol in _ROL:
        if re.match(pat, tok):
            return rol


def _num(s):
    return int(s.replace(".", "").replace(",", ""))


def tablas_oportunidad(pags):
    """Busca todas las tablas con fila Total cuyo encabezado se reconoce. Devuelve lista de dicts (sistema, roles, valores, pagina)."""
    full, offs = "", []
    for p in pags:
        offs.append(len(full))
        full += p + " "
    out = []
    for m in re.finditer(r"\b(Total general|TOTAL|Total)\s+(\d[\d.]*(?:\s+\d[\d.]*){2,9})(?![\d.,]*\s*%)", full):
        nums_all = [_num(x) for x in re.findall(r"\d[\d.]*", m.group(2))]
        ventana = full[max(0, m.start() - 3500):m.start()]
        toks, hdr = [], ""
        for mk in sorted({mm.start() for x in ("Tipo de", "Tipo ", "Modalidades de", "Modalidad ", "Dependencia") for mm in re.finditer(re.escape(x), ventana)}, reverse=True):
            h = re.split(r"\d", ventana[mk:], maxsplit=1)[0]
            tk = re.findall(_TOK, sin_acentos(h).lower())
            if len(tk) >= 3:
                toks, hdr, k = tk, h, mk
                break
        if not toks:
            continue
        roles = [_rol(t) for t in toks]
        if len(roles) > len(nums_all) or len(set(roles)) != len(roles) or "ET" not in roles:
            continue
        nums = nums_all[:len(roles)]                                 # una fila al final de pagina puede quedar pegada al texto de la pagina siguiente
        vals = dict(zip(roles, nums))
        tot = vals.get("TOT")
        resto = sum(v for r_, v in vals.items() if r_ != "TOT")
        if tot is None:
            tot, sin_tot = resto, True
        else:
            sin_tot = False
            if resto != tot:
                continue
        # sistema: primero la 'Fuente' que sigue a la tabla; si no, el ultimo titulo que la precede
        post = sin_acentos(full[m.end():m.end() + 90]).lower()
        mf = re.search(r"fuente:?\s*sistema (bpms|de atencion|debora)", post)
        if mf:
            sis = "BPMS" if mf.group(1) == "bpms" else "Debora" if mf.group(1) == "debora" else "SAC"
        else:
            previo = sin_acentos(ventana[max(0, k - 900):k + 80]).lower()
            ib = max(previo.rfind("bpms"), previo.rfind("debora"))
            isac = max(previo.rfind("sistema de atencion al ciudadano"), previo.rfind("sac"))
            sis = None if max(ib, isac) < 0 else ("SAC" if isac > ib else "Debora" if previo.rfind("debora") > previo.rfind("bpms") else "BPMS")
        pag = max(i for i, o in enumerate(offs) if o <= m.start()) + 1
        tt = re.findall(r"(?i)(atenci[oó]n en t[eé]rminos\.?(?: bpms| sistema de atenci[oó]n al ciudadano(?: \(?SAC\)?)?)?|oportunidad en la atenci[oó]n y estado(?: bpms| d[eé]bora| sistema de atenci[oó]n al ciudadano(?: \(SAC\))?)?)",
                        ventana[max(0, k - 600):k + 120])
        out.append({"titulo": norm(tt[-1]) if tt else "", "sistema": sis, "roles": roles, "vals": vals, "total": tot, "sin_total_col": sin_tot, "pagina": pag, "encabezado": norm(hdr)[:140]})
    return out


def leer_informe(ruta, anio_hint=None):
    """Lee un informe PQRSDF (solo dato). Devuelve dict con ok, motivo y, si ok: anio, meses, tipo (T/S/A), sistema (BPMS o Debora), bpms = (atendida tarde, en terminos, pendientes, total),
    venc, sinr, sac, pagina. El anio se contrasta entre la portada, el Alcance y el inventario; si hay dudas devuelve ok=False con el motivo real (no adivina)."""
    R = {"ok": False, "motivo": "", "metodo": "automatico"}
    try:
        lector = pypdf.PdfReader(str(ruta))
        crudo = [p.extract_text() or "" for p in lector.pages]
    except Exception as e:                                           # PDF danado o no PDF
        R["motivo"] = "PDF ilegible (" + type(e).__name__ + ")"
        return R
    pags = [norm(t) for t in crudo]
    pl = [sin_acentos(p).lower() for p in pags]
    full = " ".join(pags)
    R["paginas"] = len(pags)
    R["decreto491"] = "decreto 491" in sin_acentos(full).lower()
    R["d491_17may"] = re.search(r"decreto 491.{0,90}17 de mayo de 2022", sin_acentos(full).lower()) is not None
    # 1. Alcance: periodo (meses) y sistema
    alc = None
    for m in re.finditer(r"Alcance\s+(Este informe.{20,900}?)\s+3\.\s*Responsabilidades", full, re.I):
        if "...." not in m.group(1):
            alc = m.group(1)
            break
    if alc is None:
        if "educacion" in pl[0][:400]:
            R["motivo"] = "informe de la Secretaria de Educacion (SAC), sin seccion Alcance: es respaldo y no se suma"
        elif "dependencia denuncia" in pl[0][:200] or len(pags) <= 3:
            R["motivo"] = f"documento de {len(pags)} pagina(s) con un cuadro por dependencia, sin seccion Alcance ni tabla de oportunidad: no es el informe trimestral"
        elif "control interno" in pl[0][:300]:
            R["motivo"] = ("informe de seguimiento de la Oficina de Control Interno (auditoria semestral con muestreo, periodo 1-ene a 30-jun-2025, fuente Debora), distinto del informe "
                           "trimestral de la Direccion de Calidad: sin tabla 'en terminos / total'; no sirve como S1")
        else:
            R["motivo"] = "sin seccion 'Alcance' legible (plantilla distinta)"
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
    al = sin_acentos(alc).lower()
    R["sistemas"] = [s for s, k in (("BPMS", "bpms"), ("Debora", "debora"), ("SAC", "sac")) if re.search(rf"\b{k}\b", al)]
    if "BPMS" not in R["sistemas"] and "Debora" not in R["sistemas"]:
        R["motivo"] = "el Alcance no nombra el sistema general (BPMS o Debora): informe de otro sistema (p. ej. solo SAC de Educacion)"
        return R
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
    # 2. Tablas de oportunidad (varias plantillas): sistema general = BPMS o Debora; SAC aparte
    tabs = tablas_oportunidad(pags)
    gen = [t for t in tabs if t["sistema"] in ("BPMS", "Debora")]
    if not gen:
        R["motivo"] = ("sin tabla de oportunidad del sistema general legible (encabezados leidos: " + "; ".join(sorted({t['encabezado'][:60] for t in tabs})[:3]) + ")") if tabs else \
            "sin tabla de oportunidad con fila Total que cuadre (plantilla no reconocida; revisar a mano)"
        return R
    t = gen[0]
    v = t["vals"]
    tot = t["total"]
    # 3. el total debe coincidir con la tabla de recibidas por tipo o repetirse en el informe
    rec = None
    for i, p in enumerate(pl):
        mr = re.search(r"recibidas en el (?:bpms|debora|sistema bpms) por cada tipo", p)
        if mr:
            mm = re.search(r"\bTotal\s+([\d.]+)\s+100", pags[i][mr.start():])
            if mm:
                rec = _num(mm.group(1))
                break
    reps = len(re.findall(rf"(?<![\d.,]){tot}(?![\d]|[.,]\d)", full)) + len(re.findall(rf"(?<![\d.,]){tot:,}(?![\d]|[.,]\d)".replace(",", "."), full)) if tot >= 1000 else len(re.findall(rf"(?<![\d.,]){tot}(?![\d.,])", full))
    if rec is not None and rec != tot:
        R["motivo"] = f"total de la tabla de oportunidad ({tot}) distinto de recibidas por tipo ({rec})"
        return R
    if rec is None and reps < 2:
        R["motivo"] = f"el total {tot} de la tabla no se repite en el informe (recibidas, nota, mes o canal) para contrastarlo; no se adivina"
        return R
    R["recibidas"], R["total_repeticiones"] = rec, reps
    R.update(ok=True, motivo="", bpms=(v.get("LATE", 0), v["ET"], v.get("PEND", 0), tot), venc=v.get("VENC", 0), sinr=v.get("SINR", 0), sistema=t["sistema"], pagina=t["pagina"],
             roles=t["roles"], encabezado=t["encabezado"], titulo=t["titulo"], suma_filas_ok=True, nota_et=None)
    mn = re.search(r"dentro de los terminos de (?:la )?ley,? ([\d.]+)", sin_acentos(full).lower())
    if mn:
        R["nota_et"] = _num(mn.group(1))
    mp = re.search(r"se encuentran? ([\d.]+) pqrsdf (?:\W{0,3})pendientes", sin_acentos(full).lower())
    R["nota_pend_distinta"] = _num(mp.group(1)) if mp and _num(mp.group(1)) not in (R["bpms"][2], R["bpms"][2] + R["venc"]) else None
    sacs = [x for x in tabs if x["sistema"] == "SAC"]
    R["sac"] = None
    if sacs:
        s = sacs[0]
        R["sac"] = (s["vals"].get("LATE", 0), s["vals"]["ET"], s["vals"].get("PEND", 0), s["total"], s["vals"].get("SINR", 0), s["vals"].get("VENC", 0))
        R["sac_pagina"] = s["pagina"]
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


# Lecturas manuales: informes que el lector automatico rechaza por una incoherencia del propio documento, pero cuyo periodo y cifras se verifican a mano.
# Cada entrada se valida contra la tabla del PDF (pagina y cifras) y lleva cita textual corta; se marca 'lectura manual'.
LECTURAS_MANUALES = {
    "64529_informe-pqrsdf-primer-trimestre.pdf": {
        "anio": 2023, "meses": [1, 2, 3], "sistema": "BPMS", "pagina": 11, "bpms": (485, 4361, 292, 5138),
        "cita": "'Atencion en Terminos. BPMS ... Atendida / Atendida en terminos / Por Atender / Total general'; fila Total general 485 4361 292 5138",
        "motivo_manual": ("la seccion Alcance dice 'entre octubre y diciembre' (error de copia de la plantilla del 4T-2022), pero la portada dice 'Primer trimestre 2023', el objetivo dice 'entre Enero y ...' "
                          "y el cuadro de recibidas por mes (p. 13) trae Enero, Febrero y Marzo con Total general 5138, igual al total de la tabla de oportunidad"),
        "alerta": "Alcance con error de copia (oct-dic) corregido a mano: periodo T1 2023 por portada, objetivo y cuadro mensual (total 5138)."},
}
INFORMES = []                                                         # (archivo, url, fecha, nota, resultado)
for r in _man_rows("pqrs"):
    if not r["archivo"].lower().endswith(".pdf"):
        continue
    fecha, url, nota = man("pqrs", r["archivo"])
    h = re.search(r"periodo segun inventario: (20\d\d)", nota) or re.search(r"(20[12]\d)", url.rsplit("/", 1)[-1])
    R_ = leer_informe(RAW / "pqrs" / r["archivo"], int(h.group(1)) if h else None)
    if not R_["ok"] and r["archivo"] in LECTURAS_MANUALES:
        M = LECTURAS_MANUALES[r["archivo"]]
        pg_ = [norm(x.extract_text() or "") for x in pypdf.PdfReader(str(RAW / "pqrs" / r["archivo"])).pages]
        tb_ = [t for t in tablas_oportunidad(pg_) if t["sistema"] == "BPMS" and t["pagina"] == M["pagina"]]
        mes_ = re.search(r"Tipo de PQRSDF Enero Febrero Marzo Total general.*?Total general (\d+) (\d+) (\d+) (\d+) Fuente", " ".join(pg_))
        if tb_ and (tb_[0]["vals"].get("LATE"), tb_[0]["vals"]["ET"], tb_[0]["vals"].get("PEND"), tb_[0]["total"]) == M["bpms"] and mes_ and int(mes_.group(4)) == M["bpms"][3]:
            R_ = {"ok": True, "motivo": "", "metodo": "manual", "anio": M["anio"], "meses": M["meses"], "tipo": "T", "sistema": M["sistema"], "sistemas": ["BPMS", "SAC"], "pagina": M["pagina"],
                  "bpms": M["bpms"], "venc": 0, "sinr": 0, "recibidas": None, "total_repeticiones": 2, "suma_filas_ok": True, "sac": None, "manual": M,
                  "roles": tb_[0]["roles"], "encabezado": tb_[0]["encabezado"], "titulo": tb_[0].get("titulo", "")}
            _sc = [t for t in tablas_oportunidad(pg_) if t["sistema"] == "SAC"]
            if _sc:
                R_["sac"] = (_sc[0]["vals"].get("LATE", 0), _sc[0]["vals"]["ET"], _sc[0]["vals"].get("PEND", 0), _sc[0]["total"], _sc[0]["vals"].get("SINR", 0), _sc[0]["vals"].get("VENC", 0))
                R_["sac_pagina"] = _sc[0]["pagina"]
        else:
            R_["motivo"] += " | la lectura manual no se pudo validar contra la tabla del PDF"
    INFORMES.append((r["archivo"], url, fecha, nota, R_))
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
    check("SAC 2023 T4 = 1513 / 2 / 80 / 1595", R["sac"][:4] == (2, 1513, 80, 1595), str(R["sac"]))
    check("Total BPMS = recibidas por tipo (4149); Alcance: oct-dic, BPMS y SAC; anio 2023 contrastado", R["recibidas"] == 4149 and R["meses"] == [10, 11, 12] and R["sistemas"] == ["BPMS", "SAC"] and R["anio"] == 2023)
    check("Suma de las filas de la tabla BPMS = fila Total", R["suma_filas_ok"])
check("Lector: un PDF sin la plantilla esperada se marca N/D con motivo, no se adivina",
      (lambda x: x["ok"] is False and "plantilla distinta" in x["motivo"])(leer_informe(RAW / "rendicion_cuentas" / "23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf")))

# ------------------------------------------------------------------ filas CS.PQRSDF
NM_P = "Oportunidad de respuesta a las PQRSDF (atendidas en terminos / recibidas)"
REF_P = "100 * atendidas en terminos / recibidas (BPMS o Debora segun el trimestre), con informes que cubren los 12 meses sin repetirse (D82), tope 100; semaforo D60"
ALERTA_P = ("ALERTA: el informe admite PQRSDF respondidas a tiempo pero cerradas tarde en el sistema, por lo que el indicador SUBESTIMA la oportunidad real; "
            "es un registro administrativo de la Alcaldia (BPMS o Debora), no una verificacion externa. Las pendientes en terminos de ley no se cuentan como atendidas en terminos. "
            "Lo confirma la auditoria independiente de Control Interno (144255, p. 25 del PDF): las solicitudes 'aparecen como extemporaneas por no haber sido cerradas oportunamente en el sistema, a pesar de haber sido atendidas dentro del plazo legal'.")
FUENTE_P = "Alcaldia de Envigado, Secretaria General, Direccion de Calidad y SGI: Informe PQRSDF trimestral, tabla de oportunidad del sistema general ('Atencion en Terminos BPMS', 'Oportunidad en la atencion y estado BPMS' u 'Oportunidad en la atencion y estado Debora'), fila Total"
ETIQ = lambda ms: (f"T{(ms[0] - 1) // 3 + 1}" if len(ms) == 3 else f"S{(ms[0] - 1) // 6 + 1}" if len(ms) == 6 else "A")
TRIMESTRES = [(a, t) for a in PERIODO for t in (1, 2, 3, 4)]
def _per_inv(nota):
    m = re.search(r"periodo segun inventario: (\d{4}|\?) (\w+)", nota)
    return (int(m.group(1)), m.group(2)) if m and m.group(1) != "?" else (None, None)


def motivo_trim(a, q):
    """Motivo real de N/D de un trimestre: lo que dijo el lector sobre el archivo local que el inventario asigna a ese periodo."""
    ms = [(arc, R["motivo"]) for arc, u, f, n, R in INFORMES if not R["ok"] and _per_inv(n) in ((a, f"T{q}"), (a, "S1" if q <= 2 else "S2"))]
    return "; ".join(f"{x} ({m})" for x, m in ms) if ms else "sin archivo local para este trimestre (no se ha podido descargar)"


def regimen_plazos(R):
    """Regimen de plazos del trimestre: el Decreto 491 de 2020 (art. 5) amplio los terminos hasta el 17-may-2022 (lo derogo la Ley 2207); despues rige la Ley 1755."""
    q = (R["meses"][0] - 1) // 3 + 1
    if R["anio"] in (2020, 2021) or (R["anio"] == 2022 and q == 1):
        return "ampliados"
    if R["anio"] == 2022 and q == 2:
        return "mixto"
    if R["anio"] == 2022:
        return "ordinarios"
    return None


def alertas_p(infs):
    """Alertas de comparabilidad de un conjunto de informes."""
    out = []
    Rs = [i[4] for i in infs]
    reg = {regimen_plazos(R) for R in Rs} - {None}
    if "ampliados" in reg:
        out.append("plazos AMPLIADOS por el Decreto 491 de 2020 (art. 5, emergencia): 'en terminos' se mide con mas tiempo y no es comparable con los periodos sin ampliacion")
    if "mixto" in reg:
        out.append("2022 T2 es MIXTO: plazos ampliados hasta el 17-may-2022 (el Decreto 491 lo derogo la Ley 2207) y ordinarios (Ley 1755 de 2015) para lo que ingreso despues (39694, p. 11)")
    if "ordinarios" in reg:
        out.append("2022 T3-T4 se miden con plazos ORDINARIOS (Ley 1755): los informes dicen que lo ingresado despues del 17-may-2022 se tramito segun esa ley (49789, p. 12; 57010, p. 12), aunque su nota repite la cita del Decreto 491")
    sis = [R["sistema"] for R in Rs]
    if "Debora" in sis and "BPMS" in sis:
        out.append("cambio de sistema dentro del anio: BPMS en T1 y Debora (en produccion desde mayo de 2025, 135341, p. 20) en T2-T4")
    elif set(sis) == {"Debora"}:
        out.append("el sistema es Debora (en produccion desde mayo de 2025, 135341, p. 20): serie no homogenea con los anios anteriores (BPMS)")
    if any(R["sinr"] for R in Rs):
        out.append(f"columna 'Sin respuesta' ({sum(R['sinr'] for R in Rs)}): no se sabe cuantas estaban dentro del plazo legal; por D86 va al MAXIMO del rango y a la brecha, junto con las pendientes")
    sp = [f"{R['anio']} T{(R['meses'][0] - 1) // 3 + 1}" for R in Rs if "PEND" not in R["roles"]]
    if sp:
        out.append("sin columna de pendientes en " + ", ".join(sp) + " (solo atendidas en terminos y extemporaneas): no se puede calcular su rango")
    if any(R["venc"] for R in Rs):
        out.append(f"columna 'Vencidas' ({sum(R['venc'] for R in Rs)}): no cuentan como en terminos ni como pendientes en el maximo; segun 118962 (p. 12) esperan aclaraciones de los peticionarios y 'aun se encuentran dentro del termino de Ley', "
                   "asi que el maximo puede estar subestimado")
    if all(R["anio"] == 2024 for R in Rs):
        qs = [(R["meses"][0] - 1) // 3 + 1 for R in Rs]
        if any(q >= 2 for q in qs):
            out.append("CAMBIO DE UNIVERSO en 2024: las recibidas suben de 19.502 (2023) a 31.970 (+64 %) por 'acciones de mejora en el procedimiento de gestion de las PQRSDF, especificamente en la correcta clasificacion de todo lo que ingresa' "
                       "(118962, p. 9; 126322, p. 10); Tramite/Servicio pasa del 2,2 % (2023 T1) al 14,5-18,0 % (2024 T2-T3). No leer 2023 a 2024 como tendencia")
        if 1 in qs:
            out.append("2024 T1 tiene su propia explicacion del aumento: inicio de anio, liquidacion de impuestos y cambio de gobierno (115216, p. 9)")
    if any(R["metodo"] == "manual" for R in Rs):
        out.append("incluye una lectura manual (" + "; ".join(R["manual"]["alerta"] for R in Rs if R["metodo"] == "manual") + ")")
    npd = [R for R in Rs if R.get("nota_pend_distinta")]
    if npd:
        out.append("la nota del informe dice " + ", ".join(str(R["nota_pend_distinta"]) for R in npd) + " pendientes, pero la tabla trae " + ", ".join(str(R["bpms"][2]) + " (+ " + str(R["venc"]) + " vencidas)" for R in npd) + ": la tabla suma al total y es la que se usa")
    ne = [R for R in Rs if R.get("nota_et") and R["nota_et"] != R["bpms"][1]]
    if ne:
        out.append("la nota del informe cita un numero de atendidas en terminos distinto del de la tabla en " + str(len(ne)) + " informe(s)")
    return out


DRIVE = "Archivo obtenido del ZIP original de Google Drive de la usuaria porque el portal daba HTTP 500; el enlace es el de origen en el portal."


def nota_drive(infs):
    return DRIVE if any("ZIP" in i[3] for i in infs) else ""


ANIO_P, COB = {}, {}
for a in PERIODO:
    # filas por informe aceptado (trimestres, semestres, anuales)
    for (aa, ms), inf in sorted(SEL.items()):
        if aa != a:
            continue
        arc, url, fecha, nota, R = inf
        at, et, pe, tot = R["bpms"]
        et_ = ETIQ(ms)
        sis_ = "Debora" if R["sistema"] == "Debora" else "BPMS"
        al_ = alertas_p([inf])
        tit_ = R.get("titulo") or "tabla de oportunidad"
        fila(f"PQRSDF, atencion en terminos {sis_}, {a} {et_}", "CS.PQRSDF.TRIM" if et_.startswith("T") else "CS.PQRSDF.SEM", "detalle", f"{a}-{et_}", et, tot, 100 * et / tot,
             ref=f"Fila Total de la tabla '{tit_}': en terminos / total", fuente=FUENTE_P + f" (p. {R['pagina']}; metodo: {'lectura manual' if R['metodo'] == 'manual' else 'automatico'}; sistema {sis_}). " + nota_drive([inf]), enlace=url, fecha=fecha, brecha="con dato",
             nota=f"Atendida fuera de terminos {at}; atendida en terminos {et}; pendientes en termino {pe}; vencidas {R['venc']}; sin respuesta {R['sinr']}; total {tot}. "
                  f"Alcance leido: meses {ms[0]}-{ms[-1]}, sistemas {' y '.join(R['sistemas'])}. Archivo local {arc}. Total contrastado: "
                  f"{'igual a recibidas por tipo' if R.get('recibidas') else 'repetido ' + str(R.get('total_repeticiones')) + ' veces en el informe'}. Columnas leidas: {R['encabezado'][:90]}."
                  + (f" Duplicado: {DUP[(aa, ms)]}." if (aa, ms) in DUP else "") + (" ALERTA: " + "; ".join(al_) + "." if al_ else "")
                  + (f" LECTURA MANUAL: {R['manual']['motivo_manual']}. Cita: {R['manual']['cita']}." if R["metodo"] == "manual" else ""))
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
                (f"sin informe trimestral propio; cubierto por un informe de otro periodo" if (a, q) in COBERTO else motivo_trim(a, q)), ref="Fila Total de la tabla 'Atencion en Terminos BPMS'",
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
        SR = sum(i[4]["sinr"] for i in infs)
        VE = sum(i[4]["venc"] for i in infs)
        RR = sum(i[4]["bpms"][3] for i in infs)
        v = min(100.0, 100.0 * ET / RR)
        # D86: la columna 'Sin respuesta' (2020 T4-2021 T4) va al maximo y a la brecha, igual que las pendientes en termino
        ANIO_P[a] = {"v": v, "max": min(100.0, 100.0 * (ET + PE + SR) / RR), "brecha": 100.0 * (PE + SR) / RR, "ET": ET, "PE": PE, "SR": SR, "VE": VE, "R": RR, "infs": infs}
        _ss = ["BPMS" if i[4]["sistema"] == "BPMS" else "Debora" for i in infs]
        sist = _ss[0] if len(set(_ss)) == 1 else "mixto (T1 BPMS; T2-T4 Debora)" if a == 2025 else " y ".join(dict.fromkeys(_ss))
        titulos = "; ".join(dict.fromkeys(i[4].get("titulo", "") for i in infs if i[4].get("titulo")))
        notas_d = [DUP[(a, ms)] for ms in cub if (a, ms) in DUP]
        extra144 = ""
        if a == 2025:
            extra144 = (f"Fuente independiente: la auditoria de Control Interno (144255, S1-2025, fuente Debora y SAC) trae un total de 14.965 PQRSD en el semestre, igual a T1 + T2 de los informes trimestrales "
                        f"({SEL[(2025, (1, 2, 3))][4]['bpms'][3]} + {SEL[(2025, (4, 5, 6))][4]['bpms'][3]}), y en la p. 25 del PDF confirma que el indicador subestima la oportunidad. ")
        fila(f"{NM_P}; sistema {sist}", "CS.PQRSDF", "principal", a, ET, RR, v, v, PESO, REF_P, FUENTE_P + (f" Tablas leidas: {titulos}. " if titulos else " ") + nota_drive(infs), "; ".join(i[1] for i in infs), max(i[2] for i in infs), "con dato",
             f"Informes usados (D82): {', '.join(i[0] + ' (' + ETIQ(i[4]['meses']) + ')' for i in infs)}; cubren los 12 meses sin repetirse"
             + (f"; hay {ncub} coberturas posibles y se usa la mas granular" if ncub > 1 else "") + f". Rango: maximo {ANIO_P[a]['max']:.2f} = (en terminos + {PE} pendientes en termino" + (f" + {SR} 'sin respuesta' (D86)" if SR else "") + f") / recibidas; "
             f"brecha {ANIO_P[a]['brecha']:.2f}" + (f" (de la cual 'sin respuesta' solo: {100.0 * SR / RR:.2f}; pendientes: {100.0 * PE / RR:.2f})" if SR else "") + ". " + (f"Vencidas {VE} (no incluidas): si tambien esperaran dentro del plazo el maximo subiria a {min(100.0, 100.0 * (ET + PE + SR + VE) / RR):.2f}. " if VE else "")
             + (("Duplicados: " + "; ".join(notas_d) + ". ") if notas_d else "") + extra144
             + ("ALERTAS DE COMPARABILIDAD: " + "; ".join(alertas_p(infs)) + ". " if alertas_p(infs) else "") + ALERTA_P)
    else:
        ANIO_P[a] = None
        faltan = [f"T{q}" for q in (1, 2, 3, 4) if q not in cuartos]
        parcial = ""
        if cand:
            parcial = " Con dato parcial (no se suma): " + "; ".join(f"{ETIQ(ms)} = {100 * SEL[(a, ms)][4]['bpms'][1] / SEL[(a, ms)][4]['bpms'][3]:.2f}" for ms in sorted(cand)) + "."
        mot = (f"los informes legibles ({', '.join(ETIQ(ms) for ms in sorted(cand)) or 'ninguno'}) no cubren los 12 meses sin repetirse"
               + (f" (faltan {', '.join(faltan)})" if faltan else "") + ": " + "; ".join(motivo_trim(a, int(x[1])) for x in faltan))
        fila_nd(NM_P, "CS.PQRSDF", "principal", a, mot, PESO, REF_P, FUENTE_P,
                "; ".join(sorted({u for q in (1, 2, 3, 4) for u in urls_inventario(a, q)})) or "sin URL en el inventario",
                SEL[sorted(k for k in SEL if k[0] == a)[0]][2] if cand else "", ALERTA_P + parcial)
for a in PERIODO:
    d = ANIO_P[a]
    if d:
        fila(NM_P + ": maximo (si las pendientes en termino y la columna 'sin respuesta' se atienden a tiempo)", "CS.PQRSDF.MAX", "respaldo", a, d["ET"] + d["PE"] + d["SR"], d["R"], d["max"], ref="(ET + pendientes en termino + sin respuesta) / R",
             fuente=FUENTE_P, fecha=FECHA_CALC, brecha="con dato (rango)", nota=f"Informativo; al indice va el valor de CS.PQRSDF. Brecha = (pendientes + sin respuesta) / R = {d['brecha']:.2f}"
             + (f" (pendientes {d['PE']}; sin respuesta {d['SR']}, D86)" if d["SR"] else "") + ".")
    else:
        fila_nd(NM_P + ": maximo", "CS.PQRSDF.MAX", "respaldo", a, "faltan informes legibles que cubran el anio", ref="(ET + pendientes + sin respuesta) / R", fuente=FUENTE_P, fecha=FECHA_CALC)
# respaldo SAC (solo Educacion; no se suma al principal para no contar dos veces)
FSAC = FUENTE_P.replace("BPMS", "SAC")
for a in PERIODO:
    d = ANIO_P[a]
    if d and all(i[4]["sac"] for i in d["infs"]):
        ETs = sum(i[4]["sac"][1] for i in d["infs"])
        Ts = sum(i[4]["sac"][3] for i in d["infs"])
        fila("PQRSDF, respaldo SAC (Educacion): atendidas en terminos / total", "CS.PQRSDF.SAC", "respaldo", a, ETs, Ts, 100 * ETs / Ts, ref="Tabla 'Atencion en terminos SAC', fila TOTAL",
             fuente=FSAC, fecha=FECHA_CALC, brecha="con dato", nota="Solo Educacion (sistema SAC); NO se suma al principal: las de Educacion que entran por BPMS se trasladan a SAC y se contarian dos veces.")
    else:
        sac23 = SEL.get((2023, (10, 11, 12)))
        extra = ""
        if a == 2023 and sac23 and sac23[4]["sac"]:
            s = sac23[4]["sac"]
            extra = f"2023 T4: atendidas en terminos {s[1]}, extemporaneas {s[0]}, pendientes {s[2]}, total {s[3]} = {100 * s[1] / s[3]:.2f} (p. {sac23[4]['sac_pagina']}); no se suma."
        sin_sac = [f"T{(i[4]['meses'][0] - 1) // 3 + 1}" for i in (d["infs"] if d else []) if not i[4]["sac"]]
        fila_nd("PQRSDF, respaldo SAC (Educacion): atendidas en terminos / total", "CS.PQRSDF.SAC", "respaldo", a,
                "LIMITACION DEL LECTOR, no del informe: las tablas SAC existen en los informes, pero la de " + (", ".join(sin_sac) or "algun trimestre") + " no se pudo leer con seguridad (formato distinto o cifras que no cuadran; p. ej. 2021 T3: las columnas suman 1.825 y el total dice 1.826); el SAC es solo respaldo de Educacion",
                ref="Tabla 'Atencion en terminos SAC', fila TOTAL", fuente=FSAC, fecha=FECHA_CALC, nota=extra)
# brecha PQRSDF: trimestres sin informe legible que los cubra / 24
nq = sum(1 for k in TRIMESTRES if k not in COBERTO)
fila("Brecha de informacion de PQRSDF (2020-2025)", "CS.PQRSDF.BRECHA", "brecha", "2020-2025", nq, 24, 100.0 * nq / 24, brecha=f"{nq} de 24 trimestres sin informe legible",
     nota="UNIDAD: trimestres (no anios). Brecha = trimestres sin informe legible (propio o dentro de un semestre) / 24. Los 24 informes se leyeron desde los archivos locales (23 automaticos y 1 lectura manual). Anios sin valor: "
          + (", ".join(str(a) for a in PERIODO if ANIO_P[a] is None) or "ninguno") + f" ({sum(1 for a in PERIODO if ANIO_P[a] is None)} de 6 anios). Informes no usados: 144255 (auditoria de Control Interno), 4355/10914/10915/16495/16496 (cuadros de una pagina por dependencia, de Educacion o resumen); no son duplicados del informe general."
          + (" Informes locales no legibles: " + "; ".join(f"{a} ({m})" for a, m in NOLEG) + "." if NOLEG else ""), fecha=FECHA_CALC)
check("PQRSDF: los 24 trimestres 2020-2025 tienen informe legible (23 automaticos + 1 lectura manual)", nq == 0 and sum(1 for k in SEL.values() if k[4]["metodo"] == "manual") == 1, f"{nq} sin cubrir")
check("PQRSDF anual: cobertura de 12 meses sin repetirse en los 6 anios", all(COB[a][0] is not None and sorted(m for c in COB[a][0] for m in c) == list(range(1, 13)) for a in PERIODO))
_rec = {a: (sum(SEL[(a, ms)][4]["bpms"][1] for ms in COB[a][0]), sum(SEL[(a, ms)][4]["bpms"][3] for ms in COB[a][0])) for a in PERIODO}
check("PQRSDF anual = 100 * suma de en terminos / suma de totales de los trimestres (recalculado)", all(abs(ANIO_P[a]["v"] - min(100.0, 100.0 * _rec[a][0] / _rec[a][1])) < 1e-9 for a in PERIODO), {a: fmt(ANIO_P[a]["v"]) for a in PERIODO})
check("PQRSDF: cada trimestre = en terminos + atendida tarde + pendientes + vencidas + sin respuesta = total", all(R_[4]["ok"] is False or R_[4]["bpms"][0] + R_[4]["bpms"][1] + R_[4]["bpms"][2] + R_[4]["venc"] + R_[4]["sinr"] == R_[4]["bpms"][3] for R_ in INFORMES))
check("PQRSDF 2024 T2: la nota del informe dice 414 pendientes y la tabla 337 (+163 vencidas): se avisa", SEL[(2024, (4, 5, 6))][4]["nota_pend_distinta"] == 414)
check("PQRSDF: 2025 T1 en BPMS y 2025 T2-T4 en Debora (cambio de sistema)", [SEL[(2025, ms)][4]["sistema"] for ms in ((1, 2, 3), (4, 5, 6), (7, 8, 9), (10, 11, 12))] == ["BPMS", "Debora", "Debora", "Debora"])
check("PQRSDF: 2022 T1 es el 34294 (portada 'Primer trimestre', meses 1-3), no un semestre", SEL[(2022, (1, 2, 3))][0].startswith("34294_"))
check("PQRSDF: 144255 (Control Interno, S1 2025) se descarta como otro tipo de informe", any(a_.startswith("144255") and not R_["ok"] and "Control Interno" in R_["motivo"] for a_, u_, f_, n_, R_ in INFORMES))

# evidencia de las alertas de comparabilidad (D85-D87): textos citados
def _pag(arch, n):
    return norm(pypdf.PdfReader(str(RAW / "pqrs" / arch)).pages[n - 1].extract_text() or "")


check("Decreto 491: 39694 p. 11, 49789 p. 12 y 57010 p. 12 dicen 'hasta el 17 de mayo de 2022' y que lo posterior se tramito segun la Ley 1755; 2022 T2 mixto y T3-T4 ordinarios",
      all("hasta el 17 de mayo de 2022" in _pag(f_, n_) and "ley 1755" in _pag(f_, n_).lower() for f_, n_ in (("39694_informepqrsdsgundotrimestre.pdf", 11), ("49789_informepqrsdftercertrimestre.pdf", 12), ("57010_informe-pqrsdf-cuarto-trimestre.pdf", 12)))
      and [regimen_plazos(SEL[(2022, ms)][4]) for ms in ((1, 2, 3), (4, 5, 6), (7, 8, 9), (10, 11, 12))] == ["ampliados", "mixto", "ordinarios", "ordinarios"] and regimen_plazos(SEL[(2023, (1, 2, 3))][4]) is None)
_r23, _r24 = ANIO_P[2023]["R"], ANIO_P[2024]["R"]
check("Cambio de universo 2024: recibidas 19.502 (2023) a 31.970 (2024) = +64 %; 118962 p. 9 y 126322 p. 10 citan la 'correcta clasificacion de todo lo que ingresa' (Tramite 14,52 % y 18,04 %); 115216 p. 9: inicio de anio, impuestos, cambio de gobierno",
      (_r23, _r24) == (19502, 31970) and round(100 * (_r24 / _r23 - 1)) == 64
      and all("correcta clasificación de todo lo que ingresa" in _pag(f_, n_) for f_, n_ in (("118962_informe-pqrsdf-segundo-trimestre_2024.pdf", 9), ("126322_informepqrsdftercertrimestre_2024.pdf", 10)))
      and "14,52%" in _pag("118962_informe-pqrsdf-segundo-trimestre_2024.pdf", 9) and "18,04%" in _pag("126322_informepqrsdftercertrimestre_2024.pdf", 10) and "liquidación de impuestos" in _pag("115216_informe-pqrsd-trimestre-12024.pdf", 9), f"{_r23} -> {_r24}")
check("Tramite/Servicio 2023 T1 = 111 de 5.138 (2,2 %) frente a 14,5-18,0 % en 2024 T2-T3 (cuadro mensual de 64529, p. 13)", "Tramite/Servicio 41 69 1 111" in _pag("64529_informe-pqrsdf-primer-trimestre.pdf", 13) and abs(100 * 111 / 5138 - 2.2) < 0.05)
check("Debora entra a produccion en mayo de 2025 (135341, p. 20) y las tablas de 2025 T2-T4 se llaman 'Oportunidad en la atencion y estado Debora' (p. 12/13)",
      "salida a producción de la herramienta Débora, a partir del mes de mayo" in _pag("135341_informepqrsdfprimertrimestre_2025-1.pdf", 20)
      and all(SEL[(2025, ms)][4]["sistema"] == "Debora" and "d" in SEL[(2025, ms)][4]["titulo"].lower() and "bora" in SEL[(2025, ms)][4]["titulo"].lower() for ms in ((4, 5, 6), (7, 8, 9), (10, 11, 12)))
      and "BPMS" not in SEL[(2025, (7, 8, 9))][4]["titulo"], SEL[(2025, (7, 8, 9))][4]["titulo"])
_p144 = [norm(x.extract_text() or "") for x in pypdf.PdfReader(str(RAW / "pqrs/144255_informesemestralseguimientopqrsd2025_1.pdf")).pages]
check("144255 (Control Interno, S1-2025): total 14.965 = T1 + T2 (7.861 + 7.104) y en la p. 25 dice que los informes no reflejan necesariamente el cumplimiento de terminos",
      SEL[(2025, (1, 2, 3))][4]["bpms"][3] + SEL[(2025, (4, 5, 6))][4]["bpms"][3] == 14965 and any("14.965" in t_ for t_ in _p144) and "no necesariamente refleja la realidad del cumplimiento de los términos legales" in _p144[24])
check("D86: maximo PQRSDF con la columna 'sin respuesta': 2020 = 90,37 y 2021 = 96,56; sin respuesta 484 y 1.720", abs(ANIO_P[2020]["max"] - 90.37) < 0.005 and abs(ANIO_P[2021]["max"] - 96.56) < 0.005 and ANIO_P[2020]["SR"] == 484 and ANIO_P[2021]["SR"] == 1720,
      f"brechas {ANIO_P[2020]['brecha']:.2f} y {ANIO_P[2021]['brecha']:.2f} (solo sin respuesta: {100 * 484 / ANIO_P[2020]['R']:.2f} y {100 * 1720 / ANIO_P[2021]['R']:.2f})")
check("PQRSDF: 2020 T1 y T2 no traen columna de pendientes (se anota)", all("PEND" not in SEL[(2020, ms)][4]["roles"] for ms in ((1, 2, 3), (4, 5, 6))) and "sin columna de pendientes en 2020 T1, 2020 T2" in " ".join(alertas_p([SEL[(2020, (1, 2, 3))], SEL[(2020, (4, 5, 6))]])))
check("PQRSDF 2025: el nombre anual de la fila no atribuye todo el anio a BPMS (T1 BPMS; T2-T4 Debora)", "Debora" in [i[4]["sistema"] for i in ANIO_P[2025]["infs"]] and [i[4]["sistema"] for i in ANIO_P[2025]["infs"]].count("BPMS") == 1)

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
def paginas_doc(carpeta, archivo):
    """Texto normalizado por pagina de un PDF, o [texto completo] de un .docx (solo dato; se lee el XML sin ejecutar nada)."""
    man(carpeta, archivo)
    ruta = RAW / carpeta / archivo
    if archivo.lower().endswith(".docx"):
        with zipfile.ZipFile(ruta) as z:
            x = z.read("word/document.xml").decode("utf-8", "replace")
        return [norm(re.sub(r"<[^>]+>", " ", x.replace("</w:p>", "\n")))]
    return [norm(p_.extract_text() or "") for p_ in pypdf.PdfReader(str(ruta)).pages]


def sq(t):
    return re.sub(r"\s+", "", sin_acentos(t)).lower()


RC_DIR = "rendicion_cuentas"
# Eventos de rendicion de cuentas 2021-2023 (D77, D81): universo = audiencia del Alcalde + actas o memorias de Salud (en el inventario no hay rendiciones de Discapacidad ni de Victimas antes de 2024).
# integral = el documento describe el evento completo (acta). Criterios de clasificacion_vrc.csv: 'con veeduria' solo si el documento lo dice expresamente (lista de asistentes o intervencion);
# 'sin veeduria' solo si un documento integral no las menciona; 'desconocida' si el documento no permite saberlo; una invitacion no cuenta.
EV_V = [
    dict(a=2021, id="23449", f="23449_preguntas-y-respuestas-rendicion-de-cuentas-envigado-2021.pdf", nombre="Audiencia publica de rendicion de cuentas del Alcalde", fecha="2021-11-29", fq="noviembre29", integral=False,
         tipo="cuadro de preguntas y respuestas (no integral)"),
    dict(a=2021, id="6384", f="6384_acta-rendicion-de-cuentas-con-evaluacion-febrero-2021.pdf", nombre="Audiencia publica de rendicion de cuentas 2020-02, Secretaria de Salud", fecha="2021-02-10", fq="febrero10de2021", integral=True,
         tipo="acta con evaluacion (integral)"),
    dict(a=2021, id="6387", f="6387_rendicion-de-cuentas-20211.pdf", nombre="Rendicion de cuentas 2021, Secretaria de Salud (presentacion del Plan de Accion en Salud; sin fecha; evento 2021-01 inferido por la numeracion del acta 2021-02)", fecha="", fq=None, integral=False,
         tipo="presentacion de 54 diapositivas (memoria, no integral)"),
    dict(a=2021, id="24141", f="24141_16122021actarendicioncuentasevaluacion.pdf", nombre="Audiencia publica de rendicion de cuentas 2021-02, Secretaria de Salud", fecha="2021-12-16", fq="diciembre16de2021", integral=True,
         tipo="acta con evaluacion (integral)"),
    dict(a=2022, id="92036", f="92036_rendicion-de-cuentas-2022.pdf", nombre="Audiencia publica de rendicion de cuentas del Alcalde", fecha="", fq=None, integral=False, tipo="informe del evento de dialogo (evaluacion y respuestas; sin lista de asistentes)"),
    dict(a=2022, id="34258", f="34258_2022actarendicioncuentas01.pdf", nombre="Audiencia publica de rendicion de cuentas 2022-01, Secretaria de Salud", fecha="2022-03-17", fq="marzo17de2022", integral=True, tipo="acta (integral)"),
    dict(a=2022, id="55874", f="55874_20221214actarendicioncuentas02.docx", nombre="Audiencia publica de rendicion de cuentas 2022-02, Secretaria de Salud", fecha="2022-12-14", fq="diciembre14de2022", integral=True, tipo="acta (integral, .docx)"),
    dict(a=2023, id="115926", f="115926_preguntasyrespuestasrendiciondecuentas2023.pdf", nombre="Audiencia publica de rendicion de cuentas del Alcalde (II semestre 2023)", fecha="", fq=None, integral=False, tipo="respuestas a ciudadanos (no integral)"),
    dict(a=2023, id="111982", f="111982_2023actarendiciondecuentas01.pdf", nombre="Audiencia publica de rendicion de cuentas 2023-01, Secretaria de Salud", fecha="2023-03-21", fq="marzo21de2023", integral=True, tipo="acta (integral)"),
    dict(a=2023, id="111981", f="111981_2023actarendiciondecuentas02.pdf", nombre="Audiencia publica de rendicion de cuentas 2023-02, Secretaria de Salud", fecha="2023-12-13", fq="diciembre13de2023", integral=True, tipo="acta (integral)"),
]
ESPERADO_V = {"23449": "desconocida", "6384": "con veeduria", "6387": "desconocida", "24141": "con veeduria", "92036": "desconocida", "34258": "con veeduria", "55874": "con veeduria",
              "115926": "desconocida", "111982": "con veeduria", "111981": "con veeduria"}
for ev in EV_V:
    pg = paginas_doc(RC_DIR, ev["f"])
    ment = [i + 1 for i, t in enumerate(pg) if "veedur" in sq(t)]
    asist = sq(pg[0])
    ma = re.search(r"asistentes:(.{0,900})", asist)
    en_asist = bool(ma and "veedur" in ma.group(1)[:700])
    if ev["integral"] and en_asist:
        ev["cls"] = "con veeduria"
    elif ev["integral"] and not ment:
        ev["cls"] = "sin veeduria"
    else:
        ev["cls"] = "desconocida"
    ev["pags"] = ment
    ev["paginas_tot"] = len(pg)
    ev["fecha_ok"] = (ev["fq"] in asist) if ev["fq"] else None
    _f, ev["url"], ev["nm"] = man(RC_DIR, ev["f"])
    ev["fc"] = _f
check("VRC 2021-2023: clasificacion automatica de las 10 piezas = lectura manual (6 actas con veeduria; Alcalde x3 y presentacion 6387 desconocida)", {e["id"]: e["cls"] for e in EV_V} == ESPERADO_V,
      "; ".join(f"{e['id']} {e['cls']} p. {','.join(map(str, e['pags'])) or '-'}" for e in EV_V))
check("VRC 2021-2023: la fecha de cada acta figura en su primera pagina (6384 10-feb-2021; 24141 16-dic-2021; 34258 17-mar-2022; 55874 14-dic-2022; 111982 21-mar-2023; 111981 13-dic-2023) y la del Alcalde 2021 (29-nov-2021)",
      all(e["fecha_ok"] for e in EV_V if e["fq"]))
check("VRC 2021-2023: las 6 actas de Salud son eventos distintos (fechas distintas, numeracion 2020-02, 2021-02, 2022-01, 2022-02, 2023-01, 2023-02)", len({e["fecha"] for e in EV_V if e["id"] in ("6384", "24141", "34258", "55874", "111982", "111981")}) == 6)
check("VRC 2021-2023: los documentos del Alcalde (23449, 92036, 115926) y la presentacion 6387 no mencionan veedurias como asistentes (solo 6387 cita un proyecto de control social sin 'veedur')",
      all(not e["pags"] for e in EV_V if e["id"] in ("23449", "92036", "115926", "6387")))
for ev in EV_V:
    VRC.setdefault(str(ev["a"]), {"n": 0, "si": 0, "no": 0, "nd": 0})
    d_ = VRC[str(ev["a"])]
    d_["n"] += 1
    d_["si" if ev["cls"] == "con veeduria" else "no" if ev["cls"] == "sin veeduria" else "nd"] += 1
for a_ in ("2021", "2022", "2023"):
    d_ = VRC[a_]
    d_["min"], d_["max"], d_["brecha"] = 100.0 * d_["si"] / d_["n"], 100.0 * (d_["si"] + d_["nd"]) / d_["n"], 100.0 * d_["nd"] / d_["n"]
check("VRC 2021 = 2/4 = 50 / 100 / 50; 2022 y 2023 = 2/3 = 66,67 / 100 / 33,33 (calculado)", [round(VRC["2021"][k], 2) for k in ("min", "max", "brecha")] == [50.0, 100.0, 50.0] and
      all([round(VRC[y][k], 2) for k in ("min", "max", "brecha")] == [66.67, 100.0, 33.33] for y in ("2022", "2023")))
ALERTA_V2 = ("ALERTA de universo: 2021-2023 solo tiene la audiencia del Alcalde y las actas o memorias de Salud; el inventario no trae rendiciones de Discapacidad ni de Victimas anteriores a 2024 "
             "(en 2024-2025 el universo incluye 3 a 5 eventos mas), asi que NO es comparable con 2024-2025 hasta buscarlas en la API del portal. El 6387 es una presentacion sin fecha que se cuenta como evento 2021-01 "
             "(supuesto por la numeracion del acta 2021-02): si fuera la presentacion de otro evento, VRC 2021 seria 2/3 = 66,67 con maximo 100. 'Con veeduria' en las actas de Salud se apoya en la lista de asistentes y en la intervencion de integrantes de veedurias.")
OTRAS_ANIO = "Otras rendiciones sectoriales con acta o memoria (Discapacidad, Victimas, etc.)"
INV_DOC = {}
for _r in _inv:
    _m = re.search(r"/(\d+)_", _r["url_origen"])
    if _m:
        INV_DOC.setdefault(_m.group(1), _r["url_origen"])
for a in PERIODO:
    k = str(a)
    if k in VRC:
        d = VRC[k]
        fila("Veedurias en la rendicion de cuentas (VRC), minimo", "CS.VRC", "principal", a, d["si"], d["n"], d["min"], d["min"], PESO, REF_V, FUENTE_V, "ver CS.VRC.EVENTO", F_VRC,
             f"desconocida {d['nd']}/{d['n']} = {d['brecha']:.2f}",
             f"Universo principal: audiencia del Alcalde + sectoriales con acta o memoria: {d['n']} eventos; con veeduria {d['si']}, sin veeduria {d['no']}, desconocida {d['nd']}. "
             f"Rango: minimo {d['min']:.2f}, maximo {d['max']:.2f}, brecha {d['brecha']:.2f}. " + (ALERTA_V2 + " " if a in (2021, 2022, 2023) else "") + ALERTA_V)
        fila("VRC, maximo (documentada + desconocida) / eventos", "CS.VRC.MAX", "respaldo", a, d["si"] + d["nd"], d["n"], d["max"], ref=REF_V, fuente=FUENTE_V, fecha=F_VRC, brecha="con dato (rango)",
             nota="Informativo: no entra al indice ni al semaforo.")
        e = AM.get(k)
        if e: fila("VRC amplio (todos los eventos con algun documento), minimo", "CS.VRC.AMPLIO", "respaldo", a, e["si"], e["n"], e["min"], ref="Sensibilidad (D33): todo evento de rendicion con documento publicado", fuente=FUENTE_V,
             fecha=F_VRC, brecha=f"desconocida {e['nd']}/{e['n']} = {e['brecha']:.2f}", nota=f"Rango {e['min']:.2f} a {e['max']:.2f}; brecha {e['brecha']:.2f}. Informativo.")
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
for ev in EV_V:
    pgtxt = (f"paginas con mencion de veedurias: {', '.join(map(str, ev['pags']))}" if ev["pags"] else "sin mencion de veedurias en ninguna pagina")
    fila(f"VRC evento {ev['a']}: {ev['nombre']} (documento {ev['id']})", "CS.VRC.EVENTO", "detalle", ev["a"],
         raw={"numerador": "1" if ev["cls"] == "con veeduria" else "0", "denominador": "1", "valor": ev["cls"]},
         ref=f"fecha del evento {ev['fecha'] or 'sin fecha en el documento'}; universo principal: si", fuente=f"{ev['tipo']}; {pgtxt} ({ev['paginas_tot']} paginas); clasificacion manual con los criterios de clasificacion_vrc.csv. " + (DRIVE if "ZIP" in ev["nm"] else ""),
         enlace=ev["url"], fecha=ev["fc"], brecha="con dato" if ev["cls"] != "desconocida" else "sin dato: desconocida",
         nota=("Con veeduria: la lista de asistentes del acta incluye integrantes de veedurias" + (" y un integrante de una veeduria interviene o presenta informe" if ev["id"] in ("6384", "24141", "55874", "111982", "111981") else "; el acta solo las nombra como asistentes y en el agradecimiento inicial") + ". "
               if ev["cls"] == "con veeduria" else "Desconocida: el documento no es integral (cuadro de preguntas o presentacion) y no permite saber si hubo veedurias. ") + "Sin nombres de personas.")
for a in (2021, 2022, 2023):
    fila_nd(f"VRC evento {a}: {OTRAS_ANIO}", "CS.VRC.EVENTO", "detalle", a, "pendiente de busqueda: el inventario no tiene rendiciones de Discapacidad ni de Victimas de 2021-2023 (solo desde 2024); buscarlas en la API del portal cuando vuelva (D81)",
            ref="universo principal (D33) si existen", fuente="Inventario data_raw_IIC/inventario_detallado_IIC.csv", enlace="sin URL en el inventario", fecha="", nota="Clasificarlas es manual, con los criterios de clasificacion_vrc.csv.")
sinV = [a for a in PERIODO if str(a) not in VRC]
fila("Brecha de informacion de VRC (2020-2025)", "CS.VRC.BRECHA", "brecha", "2020-2025", len(sinV), 6, 100.0 * len(sinV) / 6, brecha=f"{len(sinV)} de 6 anios sin valor (2020); eventos desconocidos: 2021 2/4, 2022 1/3, 2023 1/3, 2024 1/3, 2025 3/5",
     nota="Anio sin valor: 2020 (sin eventos clasificados; el 6385 es presentacion de Salud). Dentro de cada anio la brecha es desconocidos / eventos: 50,00 (2021), 33,33 (2022), 33,33 (2023), 33,33 (2024), 60,00 (2025). Falta buscar rendiciones de Discapacidad y Victimas de 2021-2023. Se completa con cs_descargas.py y clasificando las actas.", fecha=F_VRC)

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
RPO[2021] = (0, 13)                                               # D87: minimo = lo verificado en la ventana
RPO21_SI_VENTANA = (11, 13)                                       # si se confirma la publicacion <= 31-dic-2021
RPOMAX = {2021: (13, 13)}
_t21p = [norm(x.extract_text()) for x in lec21.pages]
check("RPO 2021 (D87): rango 0/13 = 0 a 13/13 = 100; con la publicacion confirmada daria 11/13 = 84,62 (preguntas 4 y 9 en 0 por el criterio de fondo, 23449 pp. 2-3)",
      RPO[2021] == (0, 13) and RPOMAX[2021] == (13, 13) and abs(100 * RPO21_SI_VENTANA[0] / RPO21_SI_VENTANA[1] - 84.62) < 0.005
      and "se la vamos a contestar" in _t21p[1] and "4. ¿La carrera 41" in _t21p[1] and "tomaremos atenta nota" in _t21p[2] and "9. ¿Qué vamos a hacer por el adulto mayor artista" in _t21p[2])
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
DOC40433 = man("rendicion_cuentas", "40433_preguntasrendicioncuentas2021.pdf")[1] + "; " + man("rendicion_cuentas", "40434_rendicioncuentas.pdf")[1]   # documentos revisados (jul-2022), sin respuestas escritas
_t40 = " ".join(" ".join(paginas_doc(RC_DIR, f_)) for f_ in ("40433_preguntasrendicioncuentas2021.pdf", "40434_rendicioncuentas.pdf"))
check("RPO 2021 (D80): los documentos 40433 (encuesta de PP) y 40434 (evaluacion de la audiencia) NO traen respuestas escritas a las preguntas 4 y 9 (sin 'carrera 41', 'artista' ni 'pavimentacion'): el 40433 y el 40434 son encuestas",
      not re.search(r"(?i)carrera 41|artista|pavimenta|Direcci[oó]n de Envejecimiento", _t40) and "Encuesta Rendición de Cuentas. 2021" in _t40 and "audiencia pública de rendición de cuentas cumplió sus expectativas" in _t40)
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
            2021: "RANGO 0 a 100 (D87): el minimo es 0/13 porque la publicacion dentro de la ventana NO esta verificada: el 23449 se publico entre el 29-nov-2021 (evento) y el ~27-ene-2022 (cota: los ids 27514-27525 son de una pagina del 2022-01-27) y no esta probado que fuera "
                  "<= 31-dic-2021. SI se confirma esa fecha, el valor seria 11/13 = 84,62 (ver CS.RPO.TODO) y el maximo 13/13 = 100. Las preguntas 4 y 9 van en 0 por el criterio de fondo: el Alcalde no las respondio de fondo en el cuadro (23449, pp. 2-3: 'si no permiten esta pregunta se la vamos a contestar...'; "
                  "'tomaremos atenta nota'). 13 preguntas numeradas (17 intervenciones menos 4 felicitaciones/agradecimientos), pp. 1-6. El 40433 y el 40434 (mismo grupo del portal, jul-2022) son ENCUESTAS (PP 2021 y evaluacion de la audiencia), no respuestas escritas: "
                  "ninguno menciona la carrera 41 ni el adulto mayor artista; las respuestas escritas, si existen, estan en otro documento no localizado. La fecha de pagina del inventario (2021-09-08) es anterior al evento y no sirve. Se verifica con la API del portal cuando vuelva.",
            2023: "6 observaciones, todas con texto de respuesta (la 3 es una respuesta general porque el ciudadano no dejo contacto). Las respuestas se publicaron DESPUES del 31-mar-2024: el propio 115926 (p. 13) dice 'el pasado 23 de enero' de 2024 al contar tramites de la IUE (escrito despues del 23-ene-2024), "
                  "viene despues del 115597 (publicado el 30-abr-2024) y la pagina del inventario es del 2024-04-30; no hay fecha exacta por archivo [S]. Asi, dentro de la ventana el numerador es 0. Ver CS.RPO.TODO (5 o 6 de 6 sin ventana). Se verifica con la API cuando vuelva el portal.",
            2024: "17 observaciones sin felicitaciones; 16 con respuesta de fondo publicada en la ventana (129738, pagina del portal 2024-11-21, documento 2024-12-12).",
            2025: "ALERTA (D76): el cuadro publicado es corto, 3 intervenciones frente a 537 asistentes; 1 felicitacion, 2 observaciones, 1 con respuesta de fondo (la otra es respuesta que remite a otra entidad). "
                  "Valor frágil: el cuadro puede no registrar todas las intervenciones."}[a]
        fila("Respuesta a observaciones ciudadanas (RPO)", "CS.RPO", "principal", a, n, d, 100.0 * n / d, 100.0 * n / d, PESO, REF_R, FUENTE_R, RC[k][1] + ("; " + DOC40433 if a == 2021 else ""), RC[k][0], "con dato (rango 0 a 100; brecha 100: publicacion en ventana no verificada)" if a == 2021 else "con dato" if a != 2025 else "con dato (con alerta)",
             extra + " " + ALERTA_R)
        mx = RPOMAX.get(a, (n, d))
        fila("RPO, maximo (rango D87)", "CS.RPO.MAX", "respaldo", a, mx[0], mx[1], 100.0 * mx[0] / mx[1], ref=REF_R, fuente=FUENTE_R, enlace=RC[k][1], fecha=RC[k][0], brecha=f"con dato (rango): {100.0 * (mx[0] - n) / mx[1]:.2f}" if a in RPOMAX else "con dato",
             nota="Informativo: no entra al indice. Sin rango distinto del minimo salvo 2021 (D87).")
fila("RPO, sin aplicar la ventana de publicacion (todo lo publicado hasta hoy)", "CS.RPO.TODO", "respaldo", 2023, 5, 6, 100.0 * 5 / 6, ref="Respuestas de fondo publicadas hasta hoy / observaciones (rango 5 a 6 de 6)",
     fuente=FUENTE_R, enlace=RC["23"][1], fecha=RC["23"][0], brecha="con dato (rango)",
     nota="Rango 5/6 = 83,33 (criterio estricto: la observacion 3 recibe respuesta parcial: Enviaseo pide contacto y Seguridad responde solo una parte) a 6/6 = 100,00 (criterio laxo). Informativo: no entra al indice. En 2021, 2024 y 2025 la ventana no cambia el resultado.")
fila("RPO 2021 si se confirma la publicacion <= 31-dic-2021 (preguntas 4 y 9 en 0)", "CS.RPO.TODO", "respaldo", 2021, RPO21_SI_VENTANA[0], RPO21_SI_VENTANA[1], 100.0 * RPO21_SI_VENTANA[0] / RPO21_SI_VENTANA[1],
     ref="11 de 13 con respuesta de fondo; condicion: 23449 publicado a mas tardar el 31-dic-2021", fuente=FUENTE_R, enlace=RC["21"][1], fecha=RC["21"][0], brecha="con dato (condicionado a la fecha)",
     nota="Informativo: no entra al indice. Pasaria a ser el valor de CS.RPO 2021 si la API del portal confirma la fecha de publicacion del 23449.")
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


MOT = {"PQRSDF": "falta PQRSDF", "DSI": "falta DSI (no aplica antes de 2024)", "VRC": "falta VRC", "RPO": "falta RPO"}
ICR, FALTAN, RANGO, RANGO_MAX, ICRMAX = {}, {}, {}, {}, {}
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
             nota="Puntajes: " + con + ". Rango con los maximos: ver CS.ICRCS.MAX. Alertas: PQRSDF subestima la oportunidad y no es homogenea entre anios (plazos ampliados por el Decreto 491 de 2020, sistema Debora desde el 2T-2025); DSI 2024-2025 en 0 por regla de ventana (mide oportunidad); VRC y RPO 2025 con alerta (ver sus filas).")
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
        ICRMAX[a] = v
        fila("ICRCS con los maximos de los componentes", "CS.ICRCS.MAX", "respaldo", a, v=v, ref="DSI, VRC, RPO y PQRSDF en su maximo", fuente="Calculo propio", fecha=FECHA_CALC,
             brecha="con dato (rango)", nota="Informativo.")
conI = [a for a in PERIODO if ICR[a] is not None]
sinI = [a for a in PERIODO if ICR[a] is None]
fila("Brecha de informacion del ICRCS (2020-2025)", "CS.ICRCS.BRECHA", "brecha", "2020-2025", len(sinI), 6, 100.0 * len(sinI) / 6, brecha=f"{len(sinI)} de 6 anios sin indice",
     nota="Sin indice: " + "; ".join(f"{a} (faltan {', '.join(FALTAN[a])})" for a in sinI) + ". Brecha = anios sin indice / anios del periodo. 2020-2023 no tienen DSI (la lista de metas es del PDM 2024-2027).", fecha=FECHA_CALC)
check("ICRCS calculado = (DSI + PQRSDF + VRC + RPO) / 4 con los componentes (recalculado) en los anios con 4 componentes",
      all(abs(ICR[a_] - (DSI[a_]["min"] + ANIO_P[a_]["v"] + VRC[str(a_)]["min"] + 100.0 * RPO[a_][0] / RPO[a_][1]) / 4) < 1e-9 for a_ in conI) and conI == [a_ for a_ in PERIODO if not FALTAN[a_]], f"anios {conI}")
check("ICRCS 2024 y 2025: ya no falta PQRSDF; 2020-2023 sin indice por falta de DSI", all("PQRSDF" not in FALTAN[a_] for a_ in PERIODO) and FALTAN[2024] == [] and FALTAN[2025] == [], str(FALTAN))
check("ICRCS 2020-2023 sin DSI", all("DSI" in FALTAN[a] for a in (2020, 2021, 2022, 2023)))
check("Pesos: 4 componentes x 0,25 = 1", abs(4 * PESO - 1) < 1e-12)
# el rango del ICRCS se compara con lo calculado de los componentes (no con constantes)
for a_ in conI:
    rpo_ = 100.0 * RPO[a_][0] / RPO[a_][1]
    hi = (DSI[a_]["max"] + ANIO_P[a_]["max"] + VRC[str(a_)]["max"] + (100.0 * RPOMAX[a_][0] / RPOMAX[a_][1] if a_ in RPOMAX else rpo_)) / 4
    check(f"ICRCS.MAX {a_} = (DSI max + PQRSDF max + VRC max + RPO max) / 4 recalculado y >= ICRCS", abs(ICRMAX[a_] - hi) < 1e-9 and ICRMAX[a_] >= ICR[a_], f"{ICR[a_]:.2f} a {ICRMAX[a_]:.2f}")
check("Semaforo D60 del ICRCS: 2024 Naranja (>= 50) y 2025 Rojo (< 50)", semaforo(ICR[2024]) == "Naranja" and semaforo(ICR[2025]) == "Rojo", f"{ICR[2024]:.2f} / {ICR[2025]:.2f}")

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
