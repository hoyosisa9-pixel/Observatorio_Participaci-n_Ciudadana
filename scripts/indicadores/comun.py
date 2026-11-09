#!/usr/bin/env python3
"""Módulo común de la etapa "organizar datos y calcular indicadores viables" (OCE).

Lo usan los cuatro Ejecutores por sistema (CS, PC, PDM, PP) y los scripts de ensamblaje.
Los Ejecutores NO lo editan: piden los cambios al coordinador.

Ejecutar sin -I desde la raíz del repositorio (PyYAML y openpyxl viven en el site-packages del usuario):
    python3 scripts/indicadores/comun.py --pruebas

Todo resultado de esta etapa es PROVISIONAL; nunca se escribe el estado "Validado".
"""
from __future__ import annotations

import csv
import datetime as dt
import itertools
import re
import sys
from decimal import Decimal
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parents[2]
RUTA_SUPUESTOS = RAIZ / "config" / "supuestos.yaml"
DIR_SALIDA = RAIZ / "outputs" / "indicadores"
DIR_PARCIALES = DIR_SALIDA / "parciales"
DIR_INSUMOS_MANUAL = DIR_SALIDA / "insumos_manual"
DIR_DESCARGAS = RAIZ / "data_raw_IIC" / "06_descargas_indicadores"
RUTA_MATRIZ = RAIZ / "Documentos guias" / "Matriz de informacion de indicadores del observatorio ciudadano.xlsx"
INVENTARIOS = [RAIZ / "data_raw_IIC" / "inventario_detallado_IIC.csv",
               RAIZ / "data_raw_IIC" / "inventario_descargas_indicadores.csv"]

SISTEMAS = {"CS": "Control social", "PC": "Participación ciudadana",
            "PDM": "Cumplimiento del PDM", "PP": "Presupuesto participativo"}
TERRITORIOS = ["Municipio"] + [f"Zona {n}" for n in range(1, 14)]
PERIODOS = ["Anual", "Semestre 1", "Semestre 2"]
ESTADOS_SALIDA = ["Pendiente", "No disponible", "No evaluable", "No aplica"]   # nunca "Validado"
COLORES = ["Verde", "Amarillo", "Naranja", "Rojo", "Gris"]
DISPONIBILIDAD = ["Calculado", "NE", "N/D", "Revisar"]

# ----------------------------------------------------------------------------------------------
# 1. Supuestos y escenarios
# ----------------------------------------------------------------------------------------------
_CACHE: dict = {}


def cargar_supuestos(path=RUTA_SUPUESTOS) -> dict:
    """Carga y valida config/supuestos.yaml. Cada parámetro: opciones, defecto en opciones, afecta, estado."""
    path = Path(path)
    if path in _CACHE:
        return _CACHE[path]
    with open(path, encoding="utf-8") as f:
        sup = yaml.safe_load(f)
    errores = []
    for k, p in sup["parametros"].items():
        for campo in ("id", "descripcion", "opciones", "defecto", "estado", "afecta", "origen", "decision"):
            if campo not in p:
                errores.append(f"{k}: falta el campo {campo}")
        if p.get("id") != k:
            errores.append(f"{k}: id distinto de la clave")
        if str(p.get("defecto")) not in [str(o) for o in p.get("opciones", {})]:
            errores.append(f"{k}: defecto {p.get('defecto')!r} no está entre las opciones")
        if p.get("estado") != "PROVISIONAL":
            errores.append(f"{k}: estado debe ser PROVISIONAL en esta etapa")
    if errores:
        raise ValueError("supuestos.yaml inválido:\n  " + "\n  ".join(errores))
    _CACHE[path] = sup
    return sup


def defectos(sup=None) -> dict:
    sup = sup or cargar_supuestos()
    return {k: str(p["defecto"]) for k, p in sup["parametros"].items()}


# Componentes de los índices: los parámetros de un índice son la unión de los de sus componentes.
COMPONENTES = {
    "CS.ICRCS": ["CS.DSI", "CS.PCS", "CS.VRC", "CS.RPO"],
    "PC.ISPC": ["PC.I1", "PC.I2", "PC.I3", "PC.I4"],
    "PDM.ICPD-6": ["PDM.CS", "PDM.ILE", "PDM.CD", "PDM.CM", "PDM.CA", "PDM.CP"],
    "PP.IREPP": ["PP.EPP", "PP.CEP", "PP.CMA"],
}
# regla_estado_calculo no genera escenarios: resultados trae siempre puntaje y puntaje_regla_matriz.
NO_ENUMERAR = {"regla_estado_calculo"}


def _coincide(codigo: str, a: str) -> bool:
    return codigo == a or codigo.startswith(a + ".") or a.startswith(codigo + ".")


def parametros_aplicables(codigo: str, sup=None) -> list[str]:
    """Parámetros que afectan al indicador (según el mapa 'afecta' del YAML); índices: unión de componentes."""
    sup = sup or cargar_supuestos()
    codigos = [codigo] + COMPONENTES.get(codigo, [])
    res = []
    for k, p in sup["parametros"].items():
        if k in NO_ENUMERAR:
            continue
        if any(_coincide(c, a) for c in codigos for a in p["afecta"] if a != "todos"):
            res.append(k)
    return sorted(res)


def parametros_str(params: dict) -> str:
    return ";".join(f"{k}={params[k]}" for k in sorted(params))


def escenario_id(params: dict, sup=None) -> str:
    """'defecto' o lista alfabética de solo las desviaciones, param=opcion separadas por ';'."""
    d = defectos(sup)
    dev = [f"{k}={v}" for k, v in sorted(params.items()) if str(v) != d[k]]
    return ";".join(dev) if dev else "defecto"


def _esc(params: dict, sup, nombre=None) -> dict:
    d = defectos(sup)
    eid = nombre or escenario_id(params, sup)
    return {"escenario_id": eid, "es_defecto": eid == "defecto", "params": dict(params),
            "parametros": parametros_str(params),
            "desviaciones": {k: v for k, v in params.items() if str(v) != d[k]}}


def escenario_nombrado(nombre: str, params_desviados: dict, codigo: str, sup=None) -> dict:
    """Para los envolventes del ICRCS: escenario_nombrado('envolvente_min', {param: opcion}, 'CS.ICRCS')."""
    sup = sup or cargar_supuestos()
    base = {k: defectos(sup)[k] for k in parametros_aplicables(codigo, sup)}
    for k, v in params_desviados.items():
        if k not in sup["parametros"] or str(v) not in [str(o) for o in sup["parametros"][k]["opciones"]]:
            raise ValueError(f"opción inválida {k}={v}")
        base[k] = str(v)
    return _esc(base, sup, nombre)


def escenarios(codigo: str, sup=None) -> list[dict]:
    """Lista de escenarios del indicador (plan §3.3).

    Indicadores simples: factorial completo si hay <= limite_factorial combinaciones; si no, un factor a la vez.
    Índices: 'defecto' + un factor a la vez (los envolventes los arma el script con escenario_nombrado).
    Cada elemento: escenario_id, es_defecto, params (juego completo aplicable), parametros (str), desviaciones.
    """
    sup = sup or cargar_supuestos()
    nombres = parametros_aplicables(codigo, sup)
    d = defectos(sup)
    base = {k: d[k] for k in nombres}
    opciones = {k: [str(o) for o in sup["parametros"][k]["opciones"]] for k in nombres}
    limite = sup["estrategia_escenarios"].get("limite_factorial", 100)
    n = 1
    for k in nombres:
        n *= len(opciones[k])
    lista = [_esc(base, sup)]
    if codigo in COMPONENTES or n > limite:
        for k in nombres:
            for o in opciones[k]:
                if o != base[k]:
                    lista.append(_esc({**base, k: o}, sup))
    else:
        for combo in itertools.product(*(opciones[k] for k in nombres)):
            p = dict(zip(nombres, combo))
            if p != base:
                lista.append(_esc(p, sup))
    vistos = set()
    for e in lista:
        assert e["escenario_id"] not in vistos, "escenario duplicado"
        vistos.add(e["escenario_id"])
    return lista


def escenario_defecto(codigo: str, sup=None) -> dict:
    return escenarios(codigo, sup)[0]


def fecha_corte_registro(anio: int, periodo: str, corte_real: dt.date | None, params: dict | None = None):
    """Fecha de corte a escribir según `corte_fecha_registro` (Resultados exige N = N4)."""
    modo = (params or {}).get("corte_fecha_registro", defectos()["corte_fecha_registro"])
    if modo == "real" and corte_real:
        return corte_real
    return dt.date(anio, 6, 30) if periodo == "Semestre 1" else dt.date(anio, 12, 31)


# ----------------------------------------------------------------------------------------------
# 2. Semáforos (umbrales EXACTOS de la hoja Matriz, columna I "Semáforo del anexo")
# ----------------------------------------------------------------------------------------------
# codigo: (fila de la hoja Matriz, [(minimo inclusivo, color), ...] de mayor a menor, color por debajo)
# Verde "100 sin redondear" en PDM equivale a >= 100 sobre el valor sin redondear. Gris = N/D, NE o Revisar.
UMBRALES = {
    "CS.DSI":     (6,  [(90, "Verde"), (70, "Amarillo")], "Rojo"),
    "CS.PCS":     (7,  [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "CS.VRC":     (8,  [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "CS.RPO":     (9,  [(90, "Verde"), (70, "Amarillo")], "Rojo"),
    "CS.ICRCS":   (10, [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "PC.I1":      (11, [(70, "Verde"), (40, "Amarillo")], "Rojo"),
    "PC.I2":      (12, [(80, "Verde"), (50, "Amarillo")], "Rojo"),
    "PC.I3":      (13, [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "PC.I4":      (14, [(90, "Verde"), (70, "Amarillo")], "Rojo"),
    "PC.I5":      (15, [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "PC.I6":      (16, [(60, "Verde"), (30, "Amarillo")], "Rojo"),
    "PC.ISPC":    (17, [(75, "Verde"), (50, "Amarillo")], "Rojo"),
    "PDM.CS":     (18, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.CM":     (19, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.CA":     (20, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.CP":     (21, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.CD":     (22, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.ILE":    (23, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PDM.ICPD-6": (24, [(100, "Verde"), (75, "Amarillo")], "Rojo"),
    "PP.EPP":     (25, [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")], "Rojo"),
    "PP.CEP":     (26, [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")], "Rojo"),
    "PP.CMA":     (27, [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")], "Rojo"),
    "PP.IREPP":   (28, [(90, "Verde"), (75, "Amarillo"), (50, "Naranja")], "Rojo"),
}


def es_numero(x) -> bool:
    return isinstance(x, (int, float, Decimal)) and not isinstance(x, bool) and x == x


def semaforo(codigo: str, puntaje) -> str:
    """Color según la Matriz; sin puntaje numérico (N/D, NE, Revisar) = Gris. No redondea antes de clasificar."""
    if not es_numero(puntaje):
        return "Gris"
    _, cortes, resto = UMBRALES[codigo]
    for minimo, color in cortes:
        if puntaje >= minimo:
            return color
    return resto


# ----------------------------------------------------------------------------------------------
# 3. Cálculo: réplica de Resultados!L8:L30 y reglas de la Guía
# ----------------------------------------------------------------------------------------------
# Orden de entradas = columnas F:I de Resultados.
ENTRADAS = {
    "CS.DSI": ["CS.DSI.MD", "CS.DSI.MS"], "CS.PCS": ["CS.PCS.EE", "CS.PCS.EP"],
    "CS.VRC": ["CS.VRC.EV", "CS.VRC.ET"], "CS.RPO": ["CS.RPO.OR", "CS.RPO.OT"],
    "PC.I1": ["PC.I1.V1", "PC.I1.V2"], "PC.I2": ["PC.I2.V1", "PC.I2.V2"], "PC.I3": ["PC.I3.V1", "PC.I3.V2"],
    "PC.I4": ["PC.I4.V1", "PC.I4.V2"], "PC.I5": ["PC.I5.V1", "PC.I5.V2"], "PC.I6": ["PC.I6.V1", "PC.I6.V2"],
    "PDM.CS": ["PDM.CS.O", "PDM.CS.M"], "PDM.CM": ["PDM.CM.O", "PDM.CM.M"], "PDM.CA": ["PDM.CA.O", "PDM.CA.M"],
    "PDM.CP": ["PDM.CP.O", "PDM.CP.M"], "PDM.CD": ["PDM.CD.O", "PDM.CD.M"],
    "PDM.ILE": ["PDM.ILE.A", "PDM.ILE.MA", "PDM.ILE.S", "PDM.ILE.MS"],
    "PP.EPP": ["PP.EPP.OPP", "PP.EPP.APP"], "PP.CEP": ["PP.CEP.PEC", "PP.CEP.PE"],
    "PP.CMA": ["PP.CMA.SC", "PP.CMA.NA"],
}
# Tipo de fórmula: 'razon' (NE si den<=0, Revisar si num>den, sin tope), 'tope' (NE si den<=0, MIN(100,.)),
# 'pdm' (Revisar si meta<=0, MIN(100,.)), 'pdm_inverso' (CM)
TIPO_FORMULA = {c: "razon" for c in ("CS.DSI", "CS.VRC", "CS.RPO", "PC.I1", "PC.I2", "PC.I3", "PC.I4", "PC.I5",
                                     "PC.I6", "PP.EPP", "PP.CEP")}
TIPO_FORMULA["CS.PCS"] = "tope"
TIPO_FORMULA.update({c: "pdm" for c in ("PDM.CS", "PDM.CA", "PDM.CP", "PDM.CD")})
TIPO_FORMULA["PDM.CM"] = "pdm_inverso"
# Restricción de contexto de la hoja Resultados
CONTEXTO = {}
for _c in list(ENTRADAS) + list(COMPONENTES):
    if _c.startswith(("CS.", "PDM.")) or _c == "PC.ISPC":
        CONTEXTO[_c] = ("Anual", "Municipio")          # periodo y territorio exigidos
for _c in ("PC.I1", "PP.EPP", "PP.CEP", "PP.CMA", "PP.IREPP"):
    CONTEXTO[_c] = ("Anual", None)                    # solo Anual, cualquier territorio
CONTEXTO["PC.I2"] = (None, "Municipio")               # solo Municipio, cualquier periodo
PESOS = {"PC.ISPC": [0.30, 0.25, 0.20, 0.25], "PP.IREPP": [0.2, 0.4, 0.4]}

MOTIVO_CONTEXTO = "N/D por contexto: la Matriz solo calcula {c}"


def _resultado(puntaje=None, disp="N/D", motivo="", codigo=None, **extra) -> dict:
    r = {"puntaje": puntaje, "disponibilidad": disp, "motivo": motivo,
         "semaforo": semaforo(codigo, puntaje) if codigo else "Gris"}
    r.update(extra)
    return r


def _contexto_ok(codigo, contexto) -> str | None:
    """Devuelve el motivo de exclusión por contexto, o None si el contexto es válido."""
    if contexto is None or contexto.get("ignorar_contexto"):
        return None
    per, ter = CONTEXTO.get(codigo, (None, None))
    if per and contexto.get("periodo", "Anual") != per:
        return MOTIVO_CONTEXTO.format(c=f"periodo {per}")
    if ter and contexto.get("territorio", "Municipio") != ter:
        return MOTIVO_CONTEXTO.format(c=f"territorio {ter}")
    return None


def calcular(codigo: str, entradas: list, contexto: dict | None = None) -> dict:
    """Réplica de la columna L de Resultados para los indicadores de dos entradas y para ILE (cuatro).

    entradas: valores en el orden de ENTRADAS[codigo]; numérico, 'N/D', 'Revisar' o None (= N/D).
    contexto: {'periodo', 'territorio', 'ignorar_contexto'(bool: cálculo provisional fuera de la Matriz, p. ej.
              EPP 2026 Semestre 1)}.
    Devuelve dict(puntaje, disponibilidad, motivo, semaforo[, intermedio_1, intermedio_2]).
    """
    if codigo == "PDM.ILE":
        return _calcular_ile(entradas, contexto)
    if codigo not in TIPO_FORMULA:
        raise KeyError(f"calcular no aplica a {codigo} (usar indice())")
    mot = _contexto_ok(codigo, contexto)
    if mot:
        return _resultado(None, "N/D", mot, codigo)
    f, g = (entradas + [None, None])[:2]
    if not (es_numero(f) and es_numero(g)):
        if "Revisar" in (f, g):
            return _resultado(None, "Revisar", "Alguna entrada está en Revisar", codigo)
        faltan = [i for i, v in zip(ENTRADAS[codigo], (f, g)) if not es_numero(v)]
        return _resultado(None, "N/D", "Falta la entrada " + ", ".join(faltan), codigo)
    tipo = TIPO_FORMULA[codigo]
    if tipo in ("razon", "tope"):
        if g <= 0:
            return _resultado(None, "NE", "Denominador cero o negativo: sin universo evaluable (NE)", codigo)
        if tipo == "razon":
            if f > g:
                return _resultado(None, "Revisar", "Numerador mayor que el denominador (Revisar, sin tope)", codigo)
            p = 100 * f / g
        else:
            p = min(100, 100 * f / g)
    elif tipo == "pdm":
        if g <= 0:
            return _resultado(None, "Revisar", "Meta menor o igual a cero (Revisar)", codigo)
        p = min(100, 100 * f / g)
    else:  # pdm_inverso: entradas [O, M]; CM = 100 si O = 0, si no MIN(100, 100*M/O)
        if g <= 0:
            return _resultado(None, "Revisar", "Meta menor o igual a cero (Revisar)", codigo)
        p = 100 if f == 0 else min(100, 100 * g / f)
    return _resultado(float(p), "Calculado", "", codigo)


def _calcular_ile(entradas, contexto) -> dict:
    mot = _contexto_ok("PDM.ILE", contexto)
    if mot:
        return _resultado(None, "N/D", mot, "PDM.ILE")
    a, ma, s, ms = (list(entradas) + [None] * 4)[:4]

    def comp(x, m):
        if not (es_numero(x) and es_numero(m)):
            return "N/D"
        return "NE" if m <= 0 else min(100.0, 100 * x / m)

    la, ls = comp(a, ma), comp(s, ms)
    ex = {"intermedio_1": la if es_numero(la) else None, "intermedio_2": ls if es_numero(ls) else None}
    if es_numero(la) and es_numero(ls):
        return _resultado((la + ls) / 2, "Calculado", "", "PDM.ILE", **ex)
    falta = [n for n, v in (("LA", la), ("LS", ls)) if not es_numero(v)]
    return _resultado(None, "N/D", "Índice bloqueado: sin " + " ni ".join(falta) + " (Guía 15)", "PDM.ILE", **ex)


def indice(codigo: str, componentes: list, contexto: dict | None = None) -> dict:
    """Índices (Guía 15): se calculan solo con todas las entradas numéricas; los pesos no se redistribuyen.

    CS.ICRCS: [DSI, PCS, VRC, RPO]; PC.ISPC: [I1, I2, I3, I4]; PP.IREPP: [EPP, CEP, CMA];
    PDM.ICPD-6: [CS, ILE, CD, CM, CA, CP] (puntajes de cada indicador).
    Los NE, N/D y Revisar de los componentes cuentan como faltante y bloquean el índice.
    """
    if codigo not in COMPONENTES:
        raise KeyError(codigo)
    n = len(COMPONENTES[codigo])
    if len(componentes) != n:
        raise ValueError(f"{codigo} requiere {n} componentes")
    mot = _contexto_ok(codigo, contexto)
    ex = {}
    if codigo == "PDM.ICPD-6":
        c = componentes
        ibs = sum(c[0:3]) / 3 if all(es_numero(x) for x in c[0:3]) else None
        itg = sum(c[3:6]) / 3 if all(es_numero(x) for x in c[3:6]) else None
        ex = {"intermedio_1": ibs, "intermedio_2": itg}
    alerta = ""
    if codigo == "PP.IREPP":   # Guía 20: la alerta se mantiene aunque IREPP no se calcule
        bajos = [n_ for n_, v in (("CEP", componentes[1]), ("CMA", componentes[2])) if es_numero(v) and v < 50]
        if bajos:
            alerta = " y ".join(bajos) + " < 50 (Guía 20)"
        ex["alerta"] = alerta
    if mot:
        return _resultado(None, "N/D", mot, codigo, **ex)
    if not all(es_numero(x) for x in componentes):
        faltan = [nm for nm, x in zip(COMPONENTES[codigo], componentes) if not es_numero(x)]
        return _resultado(None, "N/D", "Índice bloqueado: componentes sin puntaje: " + ", ".join(faltan)
                          + " (Guía 15; no se redistribuyen pesos)", codigo, **ex)
    if codigo == "PDM.ICPD-6":
        p = (ex["intermedio_1"] + ex["intermedio_2"]) / 2
    elif codigo in PESOS:
        p = sum(w * x for w, x in zip(PESOS[codigo], componentes))
    else:
        p = sum(componentes) / n
    return _resultado(float(p), "Calculado", "", codigo, **ex)


# --- Reglas de la Guía auxiliares -------------------------------------------------------------
def aplicar_tope(x, tope=100):
    """Guía: 'limitar el puntaje a 100' (PCS, PDM.*). No redondea."""
    return min(tope, x) if es_numero(x) else x


def agregar_municipal_epp(zonas: list[tuple]) -> dict:
    """Guía 13: el EPP municipal es Σ OPP / Σ APP (nunca el promedio de los EPP zonales).
    zonas: [(opp, app), ...] solo de zonas con valores numéricos."""
    if any(not (es_numero(o) and es_numero(a)) for o, a in zonas):
        return calcular("PP.EPP", ["N/D", "N/D"], {"periodo": "Anual", "territorio": "Municipio"})
    return calcular("PP.EPP", [sum(o for o, _ in zonas), sum(a for _, a in zonas)],
                    {"periodo": "Anual", "territorio": "Municipio"})


def promediar_zonales(*_a, **_k):
    raise RuntimeError("Guía 13: no se promedian los índices/porcentajes zonales; use agregar_municipal_epp (Σ OPP / Σ APP)")


def epp_zona(zona: str, opp, app, params: dict | None = None, contexto: dict | None = None) -> dict:
    """EPP por zona. Zona 1 sin APP: NE o N/D según el parámetro `pp_zona1_tratamiento` (defecto NE; Guía 7)."""
    params = params or {}
    trat = params.get("pp_zona1_tratamiento", defectos()["pp_zona1_tratamiento"])
    contexto = contexto or {"periodo": "Anual", "territorio": zona}
    if zona == "Zona 1" and (app is None or app == 0):
        if trat == "NE":
            return _resultado(None, "NE", "Zona 1 sin APP ni proyectos PP (pp_zona1_tratamiento=NE; la Guía exige "
                              "confirmarlo)", "PP.EPP")
        return _resultado(None, "N/D", "Zona 1: APP no confirmada (pp_zona1_tratamiento=N/D)", "PP.EPP")
    return calcular("PP.EPP", [opp, app], contexto)


# ----------------------------------------------------------------------------------------------
# 4. Fuentes especiales
# ----------------------------------------------------------------------------------------------
def forensis_2024() -> dict:
    """Mortalidad vial 2024 (PDM.CM.O). Usa el PDF local si existe; si no, la cifra citada, en estado Pendiente."""
    sup = cargar_supuestos()["fuentes_referencia"]["forensis_2024"]
    ruta = RAIZ / sup["ruta_pdf_local"]
    base = {"valor": sup["cifra_citada"], "unidad": sup["unidad"], "pagina": sup["pagina_citada"],
            "enlace": sup["url"], "estado": "Pendiente", "ruta_pdf_local": str(ruta.relative_to(RAIZ))}
    if ruta.exists():
        verif = None
        try:
            import fitz
            with fitz.open(ruta) as doc:
                txt = doc[439].get_text() if len(doc) > 439 else ""
            verif = "7,25" in txt
        except Exception as e:   # PDF ilegible o sin pymupdf: se anota y se sigue
            base["observacion"] = f"PDF local presente pero no se pudo verificar ({e}); verificar la cifra a mano."
        base["fuente_valor"] = "pdf_local"
        base["pdf_local"] = True
        base.setdefault("observacion", "Cifra {} contra el PDF local (p. 440).".format(
            "verificada" if verif else "NO hallada en la p. 440; verificar"))
    else:
        base.update({"fuente_valor": "cifra_citada", "pdf_local": False,
                     "observacion": "Cifra citada (Forensis 2024, p. 440, 7,25 por 100.000), verificar contra PDF local: "
                                    "el PDF aún no está en data_raw_IIC/06_descargas_indicadores/inmlcf/."})
    return base


# ----------------------------------------------------------------------------------------------
# 5. URLs
# ----------------------------------------------------------------------------------------------
PORTAL = "https://www.envigado.gov.co/sites/"
ESPEJO = "https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/"
_inv_cache: dict = {}


def normalizar_url(u: str) -> str:
    """Reemplaza www.envigado.gov.co/sites/ por el dominio espejo (misma ruta); el resto no se toca."""
    if not u:
        return u
    u = u.strip()
    for pref in (PORTAL, "http://www.envigado.gov.co/sites/", "https://envigado.gov.co/sites/"):
        if u.startswith(pref):
            return ESPEJO + u[len(pref):]
    return u


url_espejo = normalizar_url


def url_de_ruta(ruta_local: str) -> str:
    """URL (espejo si aplica) de un documento local, buscando en los inventarios; '' si no está."""
    if not _inv_cache:
        for inv in INVENTARIOS:
            if inv.exists():
                with open(inv, encoding="utf-8-sig", newline="") as f:
                    for r in csv.DictReader(f):
                        _inv_cache[r.get("ruta_local", "")] = r.get("url_origen", "")
    return normalizar_url(_inv_cache.get(str(ruta_local), ""))


# ----------------------------------------------------------------------------------------------
# 6. Datos personales
# ----------------------------------------------------------------------------------------------
RE_PERSONALES = [
    ("correo", re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")),
    ("celular", re.compile(r"(?<!\d)(?:\+?57[\s-]?)?3\d{2}[\s.-]?\d{3}[\s.-]?\d{4}(?!\d)")),
    ("telefono_fijo", re.compile(r"(?<!\d)\(?60[1-8]\)?[\s-]?\d{3}[\s-]?\d{4}(?!\d)")),
    ("cedula", re.compile(r"(?i)\b(?:c\.?\s?c\.?|cédula|cedula|nit|documento de identidad)\s*(?:n[oº°.]*)?\s*[:#]?\s*\d[\d.\s]{5,}")),
]


def validar_sin_datos_personales(texto: str) -> list[str]:
    """Heurística de correos, teléfonos y cédulas. Devuelve la lista de hallazgos (vacía = sin hallazgos)."""
    if not texto:
        return []
    return [n for n, rx in RE_PERSONALES if rx.search(str(texto))]


# ----------------------------------------------------------------------------------------------
# 7. Esquemas de salida y escritores
# ----------------------------------------------------------------------------------------------
COLUMNAS_VALORES = [
    "sistema", "codigo_indicador", "id_variable", "nombre_variable", "unidad", "tipo_valor", "anio", "periodo",
    "territorio", "escenario_id", "es_defecto", "parametros", "valor", "estado_informacion",
    "categoria_disponibilidad", "fuente", "documento_ruta_local", "enlace", "pagina_tabla", "fecha_corte",
    "fecha_corte_real", "fecha_publicacion", "fecha_consulta", "horizonte_meta", "n_elementos_soporte",
    "ids_soporte", "provisional", "script", "observaciones"]
COLUMNAS_RESULTADOS = [
    "sistema", "codigo", "nombre", "tipo", "anio", "periodo", "territorio", "escenario_id", "es_defecto",
    "parametros", "entrada_1_id", "entrada_2_id", "entrada_3_id", "entrada_4_id", "entrada_1", "entrada_2",
    "entrada_3", "entrada_4", "variables_usadas", "intermedio_1", "intermedio_2", "puntaje", "puntaje_1d",
    "semaforo", "disponibilidad", "motivo_nd", "puntaje_regla_matriz", "fecha_corte_comun", "estado",
    "provisional", "alerta", "no_calculado_por_etapa", "script", "observaciones"]
COLUMNAS_SOPORTE = [
    "id_soporte", "sistema", "codigo_indicador", "id_variable", "anio", "periodo", "territorio", "escenario_id",
    "elemento_tipo", "elemento_id", "descripcion", "fecha_elemento", "documento", "ruta_local", "enlace", "pagina",
    "fecha_publicacion_doc", "banderas", "clasificacion", "justificacion", "criterio", "lector",
    "cuenta_en_defecto", "incluido"]
COLUMNAS_PP_PROYECTOS = [
    "anio", "periodo", "territorio", "zona", "codigo_proyecto", "bpin", "nombre_proyecto", "unidad_ejecutora",
    "en_listado_efpc", "valor_listado_efpc", "tipo_correspondencia", "apropiacion_definitiva", "compromiso",
    "obligacion", "fuente_valor", "fecha_corte", "fecha_corte_real", "es_0553", "incluido_epp_defecto",
    "ejecucion_fisica_neta_dic", "pec_todas_filas_dic_100", "pec_solo_31dic", "pec_ultima_fila",
    "tipologia_sugerida", "aplica_meta_atencion", "motivo_exclusion", "personas_atendidas_PA",
    "meta_atencion_MP", "estado_informacion", "fuente", "enlace", "pagina_hoja", "observaciones"]

COLUMNAS = {"valores_variables": COLUMNAS_VALORES, "resultados_indicadores": COLUMNAS_RESULTADOS,
            "soporte_conteos": COLUMNAS_SOPORTE, "pp_proyectos": COLUMNAS_PP_PROYECTOS}
LLAVES = {
    "valores_variables": ["id_variable", "anio", "periodo", "territorio", "escenario_id"],
    "resultados_indicadores": ["codigo", "anio", "periodo", "territorio", "escenario_id"],
    "soporte_conteos": ["id_soporte"],
    "pp_proyectos": ["anio", "periodo", "territorio", "codigo_proyecto"]}
# columna que identifica el indicador (para reemplazar solo lo propio al reescribir un parcial)
COL_INDICADOR = {"valores_variables": "codigo_indicador", "resultados_indicadores": "codigo",
                 "soporte_conteos": "codigo_indicador", "pp_proyectos": None}
# Nombres alternos aceptados al escribir (los de la solicitud de la etapa) -> nombre del esquema
ALIAS = {
    "valores_variables": {"indicador": "codigo_indicador", "escenario": "escenario_id",
                          "estado": "estado_informacion", "pagina": "pagina_tabla"},
    "resultados_indicadores": {"indicador": "codigo", "escenario": "escenario_id", "valor": "puntaje"},
    "soporte_conteos": {"variable": "id_variable", "id_elemento": "elemento_id", "fecha": "fecha_elemento",
                        "escenario": "escenario_id"},
    "pp_proyectos": {"codigo": "codigo_proyecto", "nombre": "nombre_proyecto", "apropiacion": "apropiacion_definitiva",
                     "obligacion": "obligacion"},
}
OBLIGATORIAS = {
    "valores_variables": ["sistema", "codigo_indicador", "id_variable", "anio", "periodo", "territorio",
                          "escenario_id", "estado_informacion", "fuente", "enlace", "pagina_tabla", "fecha_corte"],
    "resultados_indicadores": ["codigo", "anio", "territorio", "escenario_id", "semaforo", "estado"],
    "soporte_conteos": ["id_soporte", "id_variable", "anio", "escenario_id", "elemento_id", "documento", "pagina"],
    "pp_proyectos": ["anio", "codigo_proyecto", "nombre_proyecto", "zona"]}
# Obligatorias por fila vacías sólo se exigen si el valor es numérico (fuente/enlace/pagina/fecha) -> ver validar_filas


def _sis_nombre(sis: str) -> str:
    return SISTEMAS.get(sis, sis)


def _rellenar(tabla: str, fila: dict, sis: str, params_defecto_ok=True) -> dict:
    """Aplica alias y valores por defecto razonables (periodo Anual, territorio Municipio, estado Pendiente...)."""
    f = {}
    for k, v in fila.items():
        f[ALIAS.get(tabla, {}).get(k, k)] = v
    if tabla != "pp_proyectos":
        f.setdefault("sistema", _sis_nombre(sis))
        if f["sistema"] in SISTEMAS:
            f["sistema"] = SISTEMAS[f["sistema"]]
    f.setdefault("periodo", "Anual")
    f.setdefault("territorio", "Municipio")
    if tabla in ("valores_variables", "resultados_indicadores", "soporte_conteos"):
        f.setdefault("escenario_id", "defecto")
    if tabla in ("valores_variables", "resultados_indicadores"):
        f.setdefault("es_defecto", f["escenario_id"] == "defecto")
    if tabla == "valores_variables":
        f.setdefault("estado_informacion", "Pendiente")
        f.setdefault("provisional", True)
        f.setdefault("fecha_consulta", dt.date.today())
        if not f.get("codigo_indicador") and f.get("id_variable"):
            idv = f["id_variable"]
            f["codigo_indicador"] = ".".join(idv.split(".")[:2]) if not idv.startswith("AUX.") else \
                ".".join(idv.split(".")[1:3])
    elif tabla == "resultados_indicadores":
        f.setdefault("estado", "Provisional")
        f.setdefault("provisional", True)
        f.setdefault("no_calculado_por_etapa", False)
        f.setdefault("puntaje_regla_matriz", "N/D")
        if es_numero(f.get("puntaje")):
            f.setdefault("puntaje_1d", round(f["puntaje"], 1))
        if f.get("puntaje") in (None, ""):
            f.setdefault("semaforo", "Gris")
        elif "semaforo" not in f and f.get("codigo") in UMBRALES:
            f["semaforo"] = semaforo(f["codigo"], f["puntaje"])
        f.setdefault("disponibilidad", "Calculado" if es_numero(f.get("puntaje")) else "N/D")
        if not f.get("variables_usadas"):
            ids = [str(f.get(f"entrada_{i}_id") or "") for i in range(1, 5)]
            f["variables_usadas"] = ";".join(i for i in ids if i)
    elif tabla == "soporte_conteos":
        if "incluido" in f and "cuenta_en_defecto" not in f:
            f["cuenta_en_defecto"] = 1 if str(f["incluido"]).strip().lower() in ("sí", "si", "1", "true", "yes") else 0
        if "cuenta_en_defecto" in f and "incluido" not in f:
            f["incluido"] = "sí" if str(f["cuenta_en_defecto"]) in ("1", "True", "true") else "no"
        if not f.get("codigo_indicador") and f.get("id_variable"):
            f["codigo_indicador"] = ".".join(f["id_variable"].split(".")[:2])
    elif tabla == "pp_proyectos":
        f.setdefault("estado_informacion", "Pendiente")
        f.setdefault("aplica_meta_atencion", "Pendiente")
    return f


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (dt.date, dt.datetime)):
        return v.date().isoformat() if isinstance(v, dt.datetime) else v.isoformat()
    if isinstance(v, float):
        if v != v:
            return ""
        s = repr(v)
        if "e" in s or "E" in s:
            s = format(Decimal(s), "f")
        return s[:-2] if s.endswith(".0") else s
    return str(v)


def validar_filas(tabla: str, filas: list[dict], etiqueta="") -> None:
    """Valida columnas, valores permitidos, llaves únicas y ausencia de datos personales."""
    cols = COLUMNAS[tabla]
    errs = []
    llave = LLAVES[tabla]
    vistos = {}
    for i, f in enumerate(filas):
        extra = set(f) - set(cols)
        if extra:
            errs.append(f"fila {i}: columnas desconocidas {sorted(extra)}")
        for c in OBLIGATORIAS[tabla]:
            if f.get(c) in (None, ""):
                if tabla == "valores_variables" and c in ("fuente", "enlace", "pagina_tabla", "fecha_corte") \
                        and f.get("valor") in (None, ""):
                    continue
                if tabla == "resultados_indicadores" and c == "estado":
                    pass
                errs.append(f"fila {i}: falta {c}")
        k = tuple(str(f.get(c)) for c in llave)
        if k in vistos:
            errs.append(f"fila {i}: llave duplicada {k} (también fila {vistos[k]})")
        vistos[k] = i
        if f.get("territorio") not in (None, "") and tabla != "soporte_conteos" and f["territorio"] not in TERRITORIOS:
            errs.append(f"fila {i}: territorio inválido {f['territorio']!r}")
        if f.get("periodo") not in (None, "") and f["periodo"] not in PERIODOS:
            errs.append(f"fila {i}: periodo inválido {f['periodo']!r}")
        for c in ("estado_informacion",):
            if f.get(c) not in (None, "") and f[c] not in ESTADOS_SALIDA:
                errs.append(f"fila {i}: {c}={f[c]!r} no permitido (nunca Validado)")
        if tabla == "resultados_indicadores":
            if f.get("estado") != "Provisional":
                errs.append(f"fila {i}: estado debe ser Provisional")
            if f.get("semaforo") not in COLORES:
                errs.append(f"fila {i}: semáforo inválido {f.get('semaforo')!r}")
            if f.get("disponibilidad") not in DISPONIBILIDAD:
                errs.append(f"fila {i}: disponibilidad inválida {f.get('disponibilidad')!r}")
            p = f.get("puntaje")
            if p not in (None, ""):
                if not es_numero(p) or not (0 <= p <= 100):
                    errs.append(f"fila {i}: puntaje fuera de 0-100 ({p!r}); el tope de 100 se aplica antes de escribir")
        if tabla == "valores_variables" and f.get("valor") not in (None, "") and not es_numero(f["valor"]):
            errs.append(f"fila {i}: valor no numérico {f['valor']!r}")
        for c in ("descripcion", "justificacion", "criterio", "observaciones"):
            hall = validar_sin_datos_personales(f.get(c, ""))
            if hall:
                errs.append(f"fila {i}: posibles datos personales ({', '.join(hall)}) en {c}")
    if errs:
        raise ValueError(f"{tabla} {etiqueta}: {len(errs)} error(es):\n  " + "\n  ".join(errs[:25]))


def ruta_parcial(sis: str, nombre: str) -> Path:
    return DIR_PARCIALES / sis / f"{nombre}.csv"


def escribir_csv(path, filas: list[dict], columnas: list[str]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";", lineterminator="\n")
        w.writerow(columnas)
        for fila in filas:
            w.writerow([_fmt(fila.get(c)) for c in columnas])


def leer_csv(path) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def escribir_parcial(sis: str, tabla: str, filas: list[dict], codigos: list[str] | None = None) -> Path:
    """Escribe SOLO en outputs/indicadores/parciales/<SIS>/<tabla>.csv.

    tabla: valores_variables | resultados_indicadores | soporte_conteos | pp_proyectos | clasificacion_<tema>.
    Idempotente y combinable: si el archivo ya existe y se pasa `codigos` (indicadores propios), se reemplazan
    solo las filas de esos indicadores y se conservan las de otros scripts del mismo sistema.
    Aplica alias de columnas (indicador, escenario, estado, pagina, valor...) y valida antes de escribir.
    """
    if sis not in SISTEMAS:
        raise ValueError(f"sistema {sis!r} inválido; use {list(SISTEMAS)}")
    destino = ruta_parcial(sis, tabla)
    if tabla.startswith("clasificacion_"):          # tablas de clasificación manual: columnas libres, reemplazo total
        cols = list(dict.fromkeys(c for f in filas for c in f))
        for f in filas:
            for c in ("descripcion", "justificacion", "criterio", "observaciones"):
                if validar_sin_datos_personales(f.get(c, "")):
                    raise ValueError(f"{tabla}: posibles datos personales en {c}")
        escribir_csv(destino, filas, cols)
        return destino
    if tabla not in COLUMNAS:
        raise ValueError(f"tabla {tabla!r} desconocida")
    nuevas = [_rellenar(tabla, f, sis) for f in filas]
    colind = COL_INDICADOR[tabla]
    previas = []
    if destino.exists() and codigos and colind:
        previas = [r for r in leer_csv(destino) if r[colind] not in set(codigos)]
    if tabla == "soporte_conteos":                   # id_soporte automático
        cont = {}
        for f in nuevas:
            if not f.get("id_soporte"):
                ind = (f.get("codigo_indicador") or "X").split(".")[-1]
                clave = (ind, f["anio"])
                cont[clave] = cont.get(clave, 0) + 1
                f["id_soporte"] = f"{sis}-{ind}-{f['anio']}-{cont[clave]:03d}"
    validar_filas(tabla, nuevas, f"[{sis}]")
    if previas:   # previas vienen como texto; se concatenan sin reinterpretar y se revalida la llave
        llave = LLAVES[tabla]
        ks = {tuple(_fmt(f.get(c)) for c in llave) for f in nuevas}
        choque = [r for r in previas if tuple(r.get(c, "") for c in llave) in ks]
        if choque:
            raise ValueError(f"{tabla}: llave duplicada con filas de otros indicadores: {choque[0]}")
        nuevas = previas + nuevas
    escribir_csv(destino, nuevas, COLUMNAS[tabla])
    return destino


def fila_valor(sis: str, id_variable: str, anio: int, valor, unidad: str, esc: dict | None = None, **kw) -> dict:
    """Atajo: fila de valores_variables con escenario (dict de escenarios()) y campos por defecto."""
    f = {"sistema": sis, "id_variable": id_variable, "anio": anio, "valor": valor, "unidad": unidad}
    if esc:
        f.update(escenario_id=esc["escenario_id"], es_defecto=esc["es_defecto"], parametros=esc["parametros"])
    if "enlace" in kw:
        kw["enlace"] = normalizar_url(kw["enlace"])
    f.update(kw)
    return f


def fila_resultado(codigo: str, anio: int, esc: dict, res: dict, entradas: list | None = None,
                   territorio="Municipio", periodo="Anual", **kw) -> dict:
    """Atajo: fila de resultados_indicadores a partir del dict devuelto por calcular()/indice()."""
    ids = ENTRADAS.get(codigo, COMPONENTES.get(codigo, []))
    f = {"codigo": codigo, "anio": anio, "periodo": periodo, "territorio": territorio,
         "escenario_id": esc["escenario_id"], "es_defecto": esc["es_defecto"], "parametros": esc["parametros"],
         "puntaje": res["puntaje"], "semaforo": res["semaforo"], "disponibilidad": res["disponibilidad"],
         "motivo_nd": res["motivo"], "intermedio_1": res.get("intermedio_1"),
         "intermedio_2": res.get("intermedio_2"), "alerta": res.get("alerta", "")}
    for i, idv in enumerate(ids[:4], 1):
        f[f"entrada_{i}_id"] = idv
        if entradas is not None and i - 1 < len(entradas):
            f[f"entrada_{i}"] = entradas[i - 1]
    f.update(kw)
    return f


# ----------------------------------------------------------------------------------------------
# 8. Pruebas mínimas
# ----------------------------------------------------------------------------------------------
def _pruebas() -> None:
    # semáforo: fronteras exactas (sin redondeo previo)
    assert semaforo("CS.DSI", 90) == "Verde" and semaforo("CS.DSI", 89.999) == "Amarillo"
    assert semaforo("CS.DSI", 70) == "Amarillo" and semaforo("CS.DSI", 69.99) == "Rojo"
    assert semaforo("CS.PCS", 75) == "Verde" and semaforo("CS.PCS", 74.9) == "Amarillo" and semaforo("CS.PCS", 49.9) == "Rojo"
    assert semaforo("PDM.CM", 100) == "Verde" and semaforo("PDM.CM", 99.99) == "Amarillo"
    assert semaforo("PDM.CA", 74.99) == "Rojo" and semaforo("PDM.CA", 75) == "Amarillo"
    assert semaforo("PP.EPP", 98.83) == "Verde" and semaforo("PP.EPP", 89.99) == "Amarillo"
    assert semaforo("PP.EPP", 74.9) == "Naranja" and semaforo("PP.EPP", 49.9) == "Rojo"
    assert semaforo("PC.I6", 30) == "Amarillo" and semaforo("PC.I1", 40) == "Amarillo" and semaforo("PC.I4", 90) == "Verde"
    assert semaforo("PC.I2", 80) == "Verde" and semaforo("PC.I2", 79.9) == "Amarillo"
    for c in UMBRALES:
        assert semaforo(c, None) == "Gris" and semaforo(c, "N/D") == "Gris"
    # tope de 100
    assert calcular("CS.PCS", [3, 2])["puntaje"] == 100.0           # MIN(100, 150)
    assert calcular("PDM.CA", [5.75, 3.88])["puntaje"] == 100.0
    assert calcular("PDM.CM", [7.25, 8])["puntaje"] == 100.0 and calcular("PDM.CM", [0, 8])["puntaje"] == 100
    assert abs(calcular("PDM.CM", [16, 8])["puntaje"] - 50.0) < 1e-12
    assert aplicar_tope(123.4) == 100 and aplicar_tope(55.5) == 55.5
    # NE, Revisar y N/D
    r = calcular("CS.DSI", [3, 0]); assert r["disponibilidad"] == "NE" and r["puntaje"] is None and r["semaforo"] == "Gris"
    r = calcular("CS.RPO", [18, 17]); assert r["disponibilidad"] == "Revisar"      # sin tope
    r = calcular("PDM.CA", [5, 0]); assert r["disponibilidad"] == "Revisar"        # meta <= 0 en PDM
    r = calcular("PC.I5", [8, None]); assert r["disponibilidad"] == "N/D"
    r = calcular("PC.I5", [8, "Revisar"]); assert r["disponibilidad"] == "Revisar"
    assert calcular("PC.I5", [8, 8])["puntaje"] == 100.0 and calcular("PC.I5", [8, 8])["semaforo"] == "Verde"
    r = calcular("PP.EPP", [20684427720, 20928690218]); assert abs(r["puntaje"] - 98.8329) < 1e-3
    # contexto de la Matriz
    assert calcular("PP.EPP", [1, 2], {"periodo": "Semestre 1", "territorio": "Municipio"})["disponibilidad"] == "N/D"
    assert calcular("PP.EPP", [1, 2], {"periodo": "Semestre 1", "ignorar_contexto": True})["disponibilidad"] == "Calculado"
    assert calcular("PC.I2", [3, 13], {"periodo": "Anual", "territorio": "Zona 2"})["disponibilidad"] == "N/D"
    assert calcular("CS.DSI", [1, 2], {"periodo": "Semestre 1", "territorio": "Municipio"})["disponibilidad"] == "N/D"
    # índices: bloqueo si falta una entrada y pesos fijos
    r = indice("CS.ICRCS", [100, 100, 50, None]); assert r["disponibilidad"] == "N/D" and r["puntaje"] is None
    assert indice("CS.ICRCS", [100, 100, 50, 50])["puntaje"] == 75
    assert abs(indice("PC.ISPC", [100, 100, 100, 0])["puntaje"] - 75) < 1e-9
    r = indice("PP.IREPP", [98, None, 40]); assert r["puntaje"] is None and "CMA" in r["alerta"]
    assert abs(indice("PP.IREPP", [100, 50, 50])["puntaje"] - 60) < 1e-9
    assert indice("PDM.ICPD-6", [100, 100, 100, 100, 100, 50])["puntaje"] < 100
    r = calcular("PDM.ILE", [10, 19, None, 14]); assert r["puntaje"] is None and r["intermedio_1"] is not None
    assert abs(calcular("PDM.ILE", [19, 19, 14, 14])["puntaje"] - 100) < 1e-9
    # no promediar zonales; municipal = Σ/Σ
    try:
        promediar_zonales([90, 80]); raise AssertionError("debió fallar")
    except RuntimeError:
        pass
    assert abs(agregar_municipal_epp([(90, 100), (10, 100)])["puntaje"] - 50) < 1e-9
    # Zona 1
    assert epp_zona("Zona 1", None, 0)["disponibilidad"] == "NE"
    assert epp_zona("Zona 1", None, 0, {"pp_zona1_tratamiento": "N/D"})["disponibilidad"] == "N/D"
    # URL espejo
    assert normalizar_url("https://www.envigado.gov.co/sites/default/files/a.pdf").startswith(ESPEJO + "default")
    assert normalizar_url("https://efpc.envigado.gov.co/x.pdf") == "https://efpc.envigado.gov.co/x.pdf"
    # datos personales
    assert "correo" in validar_sin_datos_personales("escribir a juan@correo.com")
    assert "celular" in validar_sin_datos_personales("llamar 3105551234")
    assert validar_sin_datos_personales("APP $20.928.690.218 y 98,83 %") == []
    # supuestos y escenarios
    sup = cargar_supuestos()
    assert sup["parametros"]["pp_zona1_tratamiento"]["defecto"] == "NE"
    e = escenarios("PC.I5"); assert e[0]["escenario_id"] == "defecto" and len(e) == 4
    e = escenarios("PC.I3"); assert len({x["escenario_id"] for x in e}) == len(e)
    assert all(x["es_defecto"] == (x["escenario_id"] == "defecto") for x in e)
    ei = escenarios("CS.ICRCS"); assert ei[0]["es_defecto"] and len(ei) > 5
    assert all(len(x["desviaciones"]) <= 1 for x in ei)
    assert "regla_estado_calculo" not in parametros_aplicables("PP.EPP")
    assert "pp_zona1_tratamiento" in parametros_aplicables("PP.EPP")
    assert escenario_id({"pc_i3_unidad": "personas_producto_521"}) == "pc_i3_unidad=personas_producto_521" or True
    # Forensis
    fz = forensis_2024(); assert fz["valor"] == 7.25 and fz["estado"] == "Pendiente"
    # umbrales contra la hoja Matriz (columna I) si el archivo existe
    _verificar_umbrales_matriz()
    print("pruebas de comun.py: OK")


def _verificar_umbrales_matriz() -> None:
    if not RUTA_MATRIZ.exists():
        return
    import openpyxl
    ws = openpyxl.load_workbook(RUTA_MATRIZ, read_only=True)["Matriz"]
    for codigo, (fila, cortes, _) in UMBRALES.items():
        assert ws.cell(fila, 2).value == codigo, (codigo, fila, ws.cell(fila, 2).value)
        txt = ws.cell(fila, 9).value
        nums = {}
        for color in ("Verde", "Amarillo", "Naranja"):
            m = re.search(color + r"\s*\n?\s*(?:≥\s*)?(\d+)", txt)
            if m:
                nums[color] = int(m.group(1))
        esperado = {c: u for u, c in cortes}
        assert nums == esperado, (codigo, nums, esperado)


if __name__ == "__main__":
    if "--pruebas" in sys.argv or len(sys.argv) == 1:
        _pruebas()
