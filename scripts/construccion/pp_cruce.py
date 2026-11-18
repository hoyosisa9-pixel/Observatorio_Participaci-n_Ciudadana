#!/usr/bin/env python3
"""Cruce por proyecto del Presupuesto Participativo (PP), 2020-2025.

Llave de cruce (D52/D53): valor votado en el listado EFPC = apropiacion INICIAL de Hacienda, de un rubro o de la suma
de los rubros de inversion (prefijo 2.3) con el mismo sufijo de proyecto (4 ultimos digitos del rubro). El nombre y la zona
solo confirman (y, en 2020, desempatan: la descripcion del rubro trae el nombre y la zona).
  exacto             : la apropiacion inicial coincide al peso (diferencia < 1 peso).
  diferencia minima  : hasta $1.000; donde hay nombre de proyecto (2020, 2024, 2025) hasta max($1.000; 0,02 % del valor
                       votado) si ademas la zona y el nombre confirman. Siempre documentada en la fila.
  no verificado      : ninguna coincidencia, varias coincidencias sin desempate, o el sufijo lo reclaman dos proyectos.
Se excluye el rubro 2.3.7.06 (deficit fiscal / pagos de vigencias anteriores; D51) de apropiacion y obligacion; el monto
excluido queda en columnas aparte. Se suman todas las fuentes de financiacion del mismo sufijo.
Fuentes: 2020-2024 PDF de ejecucion de egresos al 31-dic (Hacienda); 2025 hoja ROCA del Plan de accion IV trimestre
(misma estructura de rubros; no hay PDF de Hacienda 2025 en el repositorio de este Codespace).

Salida: outputs/construccion/PP/pp_cruce_proyectos.csv (; y BOM) y ajustes_pp.csv.
Uso   : python3 scripts/construccion/pp_cruce.py   (desde la raiz; sin -I)
Todo resultado es PROVISIONAL.
"""
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pp_lectores as L  # noqa: E402

RAIZ = L.RAIZ
SALIDA = RAIZ / "outputs/construccion/PP/pp_cruce_proyectos.csv"
SALIDA_AJ = RAIZ / "outputs/construccion/PP/pp_ajustes.csv"
ANIOS = range(2020, 2026)
DIF_MIN = 1000          # pesos: tope minimo de la "diferencia minima" entre valor votado y apropiacion inicial
DIF_REL = 0.0002        # ... o 0,02 % del valor votado, lo que sea mayor (la diferencia minima exige confirmar con nombre y zona)
DEFICIT = "23706"       # rubro 2.3.7.06 (financiacion de deficit fiscal)
FECHA = "2026-11-18"


def _fechas():
    """fecha_consulta por archivo, tomada del manifiesto de cada carpeta."""
    out = {}
    for carpeta in ("efpc_listados_pp", "hacienda_ejecucion", "plan_accion_2025"):
        with open(L.BASE / carpeta / "manifiesto.csv", encoding="utf-8", newline="") as f:
            for r in csv.DictReader(f):
                out[f"{carpeta}/{r['archivo']}"] = r["fecha_consulta"]
    return out


FECHAS = _fechas()
URL_EFPC = "https://efpc.envigado.gov.co/wp-content/uploads/2025/07/Proyectos_PP_{a}.pdf"
FUENTES = {
    2020: ("Hacienda: ejecucion de egresos al 31-dic-2020 (PDF)", "https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/000324/16192_dic312020ejecuciondegastos.pdf", "hacienda_ejecucion/eje2020q4.pdf"),
    2021: ("Hacienda: ejecucion de egresos IV trimestre 2021 (PDF)", "https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/000589/29402_2021ejecucionpresupuestalegresostrimestre04.pdf", "hacienda_ejecucion/eje2021q4.pdf"),
    2022: ("Hacienda: ejecucion de egresos IV trimestre 2022 (PDF)", "https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/001182/59091_2022ejecucion-presupuestal-egresoscuarto-trimestre.pdf", "hacienda_ejecucion/eje2022q4.pdf"),
    2023: ("Hacienda: ejecucion de egresos IV trimestre 2023 (PDF)", "https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002256/112776_ejecucion-presupuestal-egresos-2023_4.pdf", "hacienda_ejecucion/eje2023q4.pdf"),
    2024: ("Hacienda: ejecucion de egresos al 31-dic-2024 (PDF)", "https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002660/132983_ejecucion-presupuestal-egresos-al-31122024.pdf", "hacienda_ejecucion/eje2024q4.pdf"),
    2025: ("Plan de accion IV trimestre 2025, hoja ROCA07.01.26 (plan_accion_roca)", "https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003057/152845_evaluaciondefinitivaplanaccionivtri.xlsx", "plan_accion_2025/evaluaciondefinitivaplanaccionivtri.xlsx"),
}
COLS = ["anio", "id_efpc", "zona", "nombre_efpc", "valor_votado", "codigo_sufijo", "apropiacion_inicial",
        "apropiacion_definitiva_sin_2_3_7_06", "compromiso_sin_2_3_7_06", "obligacion_sin_2_3_7_06", "excluido_2_3_7_06_apropiacion",
        "excluido_2_3_7_06_obligacion", "diferencia_votado_menos_inicial", "tipo_cruce", "pagina_pdf",
        "fuente_valor", "fuente", "enlace_ejecucion", "enlace_listado_efpc", "fecha_consulta", "fecha_consulta_listado", "observacion"]


def agrupar(filas):
    """sufijo -> estadisticas de los rubros de inversion (prefijo 23), separando el 2.3.7.06."""
    g = defaultdict(lambda: {"ini": 0.0, "dfn": 0.0, "obl": 0.0, "comp": 0.0, "x_dfn": 0.0, "x_obl": 0.0, "rows": [],
                             "pag": [], "desc": []})
    for r in filas:
        if not r["rubro"].startswith("23") or not r["ok"]:
            continue
        s = g[r["sufijo"]]
        if r["rubro"].startswith(DEFICIT):
            s["x_dfn"] += r["ap_def"]
            s["x_obl"] += r["obl"]
            s["pag"].append(r["pagina"])
            continue
        s["ini"] += r["ap_ini"]
        s["dfn"] += r["ap_def"]
        s["obl"] += r["obl"]
        s["comp"] += r["comp"]
        s["rows"].append(r["ap_ini"])
        s["pag"].append(r["pagina"])
        s["desc"].append(r["descripcion"])
    return g


STOP = {"zona", "zonas", "municipio", "envigado", "proyecto", "para", "como", "desde", "entre", "mediante", "del", "las", "los", "con"}


def _norm(t):
    import unicodedata
    t = unicodedata.normalize("NFD", t.lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def _zona_en(txt):
    m = re.findall(r"(?i)zona\s*(\d{1,2})|\b(\d{1,2})\s+del\s+municipio", txt)
    return int([x for x in m[-1] if x][0]) if m else None


def confirma(nom_efpc, zona, nom_h):
    """Zona y nombre del proyecto en Hacienda/plan de accion coinciden con los del listado EFPC (solo confirma)."""
    h = _norm(nom_h)
    m = re.findall(r"zona\s*(\d{1,2})", h)
    if m and int(m[-1]) != zona:
        return False
    a = {w[:5] for w in re.findall(r"[a-z]{4,}", _norm(nom_efpc)) if w not in STOP}
    b = {w[:5] for w in re.findall(r"[a-z]{4,}", h) if w not in STOP}
    return bool(a and b) and len(a & b) / len(a | b) >= 0.30 and (bool(m) or len(a & b) / len(a | b) >= 0.5)


def candidatos(v, g, con_nombres=True):
    """Sufijos cuya apropiacion inicial (suma o un rubro) coincide con v: exactos y de diferencia minima (por separado)."""
    ex, mn = {}, {}
    tol = max(DIF_MIN, DIF_REL * v) if con_nombres else DIF_MIN   # sin nombre que confirme, solo hasta $1.000
    for suf, s in g.items():
        if s["ini"] <= 0:
            continue
        d_sum = abs(s["ini"] - v)
        d_row = min((abs(x - v) for x in s["rows"] if x > 0), default=1e18)
        d = min(d_sum, d_row)
        via = "suma del sufijo" if d_sum <= d_row else "un rubro del sufijo"
        if d < 1:
            ex[suf] = via
        elif d <= tol:
            mn[suf] = (via, d)
    return ex, mn


def cruzar(anio, efpc, g, nombres=None):
    """nombres: {sufijo: nombre del proyecto} cuando la fuente lo trae (solo confirma y desempata)."""
    res = []
    cand = []
    nombres = nombres or {}
    for e in efpc:
        ex, mn = candidatos(e["valor"], g, bool(nombres))
        nota = ""
        if nombres:     # diferencia > $1.000: exige nombre y zona confirmados; hasta $1.000 solo que la zona no contradiga
            mn = {s_: v_[0] for s_, v_ in mn.items()
                  if (s_ in nombres and confirma(e["nombre"], e["zona"], nombres[s_]))
                  or (v_[1] <= DIF_MIN and not (s_ in nombres and _zona_en(nombres[s_]) not in (None, e["zona"])))}
        else:
            mn = {s_: v_[0] for s_, v_ in mn.items()}
        if ex:
            c, tipo = ex, "exacto"
        else:
            c, tipo = mn, "diferencia minima"
        if len(c) > 1 and nombres:   # desempate por nombre y zona (solo 2024 y 2025)
            z = {s_: v_ for s_, v_ in c.items() if s_ in nombres and confirma(e["nombre"], e["zona"], nombres[s_])}
            if len(z) == 1:
                c, nota = z, f"desempate por nombre y zona ({', '.join(sorted(ex or mn))} coinciden en valor); "
        cand.append([c, tipo, nota])
    # desempate por zona solo en 2020 (la descripcion del rubro trae el nombre del proyecto y, casi siempre, la zona)
    if anio == 2020:
        for e, k in zip(efpc, cand):
            if len(k[0]) > 1:
                zs = {s_: _zona_en(" ".join(g[s_]["desc"])) for s_ in k[0]}
                exp = [s_ for s_, z in zs.items() if z == e["zona"]]
                neu = [s_ for s_, z in zs.items() if z is None]
                if len(exp) == 1:
                    k[0], k[2] = {exp[0]: k[0][exp[0]]}, f"desempate por la zona {e['zona']} de la descripcion del rubro; "
                elif not exp and len(neu) == 1:
                    k[0], k[2] = {neu[0]: k[0][neu[0]]}, ("desempate por descarte: la descripcion de los demas sufijos con el mismo valor "
                                                         f"nombra otra zona (la de este no trae numero de zona); ")
    for i, e in enumerate(efpc):
        if (anio, e["id"]) in ASIGNACION_SUFIJO:
            suf, zz = ASIGNACION_SUFIJO[(anio, e["id"])]
            if suf in g and abs(g[suf]["ini"] - e["valor"]) < 1:
                cand[i] = [{suf: "suma del sufijo"}, "exacto", f"D55: sufijo {suf} asignado a {zz} (cada sufijo conserva su zona en todos los anios; el valor coincide al peso); "]
    reclamos = defaultdict(list)
    for i, k in enumerate(cand):
        if (anio, efpc[i]["id"]) in DUPLICADOS:
            continue
        if len(k[0]) == 1:
            reclamos[next(iter(k[0]))].append(i)
    for i, (e, k) in enumerate(zip(efpc, cand)):
        c, tipo, nota = k
        base = {"anio": anio, "zona": e["zona"], "nombre_efpc": e["nombre"], "valor_votado": e["valor"], "id": e["id"]}
        if (anio, e["id"]) in DUPLICADOS:
            base.update(codigo_sufijo="", ini=None, dfn=None, obl=None, x_dfn=None, x_obl=None, tipo="duplicado del listado", pag="",
                        dif=None, comp=None, obs=f"D55: mismo proyecto que {DUPLICADOS[(anio, e['id'])]} (misma zona, nombre y valor); se cuenta una vez y queda fuera del denominador")
            res.append(base)
            continue
        if len(c) == 1 and len(reclamos[next(iter(c))]) == 1:
            suf = next(iter(c))
            s = g[suf]
            d = e["valor"] - s["ini"]
            obs = nota
            if c[suf] == "un rubro del sufijo":
                obs += (f"el valor votado coincide con un solo rubro; la apropiacion inicial total del sufijo es {round(s['ini']):,} "
                        "(otras fuentes o rubros del mismo sufijo); ")
            if tipo == "diferencia minima":
                obs += (f"diferencia de {round(e['valor'] - (s['ini'] if c[suf] == 'suma del sufijo' else min(s['rows'], key=lambda x: abs(x - e['valor']))))} pesos "
                        f"({100 * abs(d) / e['valor']:.4f} % del valor votado)"
                        + ("; nombre y zona confirmados; " if nombres and suf in nombres and confirma(e["nombre"], e["zona"], nombres[suf])
                           else "; el nombre no se pudo confirmar (diferencia <= $1.000 y la zona no contradice); " if nombres else "; "))
            elif nombres and suf in nombres and not confirma(e["nombre"], e["zona"], nombres[suf]):
                zz = _zona_en(nombres[suf])
                obs += (f"ALERTA: la zona del nombre en la fuente ({zz}) difiere de la del listado; " if zz not in (None, e["zona"])
                        else "nombre no confirmado automaticamente (texto de la fuente parcial o redactado distinto); la zona no contradice; ")
            elif nombres and suf in nombres:
                obs += "nombre y zona confirmados; "
            pg_ = sorted({p for p in s["pag"]}, key=lambda p: (isinstance(p, str), p if isinstance(p, int) else 0, str(p)))
            pag = ", ".join(str(p) for p in pg_) if len(pg_) <= 6 else f"{pg_[0]} ... {pg_[-1]} ({len(pg_)} paginas)"
            base.update(codigo_sufijo=suf, ini=s["ini"], dfn=s["dfn"], obl=s["obl"], x_dfn=s["x_dfn"], x_obl=s["x_obl"],
                        tipo=tipo, pag=pag, dif=d, obs=obs.strip("; "), comp=s["comp"])
        else:
            if len(c) > 1:
                m = f"varias coincidencias de apropiacion inicial ({', '.join(sorted(c))}) sin desempate"
            elif len(c) == 1:
                m = f"el sufijo {next(iter(c))} lo reclaman {len(reclamos[next(iter(c))])} proyectos con el mismo valor"
                if any(j != i and efpc[j]["zona"] == e["zona"] and abs(efpc[j]["valor"] - e["valor"]) < 1 for j in reclamos[next(iter(c))]):
                    m += " (el listado EFPC trae dos filas con la misma zona y el mismo valor: posible duplicado del listado)"
            else:
                m = "ninguna apropiacion inicial de Hacienda coincide con el valor votado (ni por diferencia minima)"
            base.update(codigo_sufijo="", ini=None, dfn=None, obl=None, x_dfn=None, x_obl=None, tipo="no verificado",
                        pag="", dif=None, obs=(nota + m).strip("; "), comp=None)
        res.append(base)
    return res


# Decisiones de la usuaria que prevalecen sobre la regla de cruce (D52). Cada una queda documentada en la fila y en ajustes.
FORZADOS_NO_VERIFICADO = {}   # D54 (18-nov): se retiro EFPC2024-Z9-4; el cruce con el sufijo 0755 se acepta

# D55 (18-nov): 2022. Z6-5 es duplicado de Z6-2 (mismo proyecto 0347); cada sufijo conserva su zona en todos los anios.
DUPLICADOS = {(2022, "EFPC2022-Z6-5"): "EFPC2022-Z6-2"}
ASIGNACION_SUFIJO = {(2022, "EFPC2022-Z3-4"): ("0346", "zona 3"), (2022, "EFPC2022-Z9-4"): ("0348", "zona 9"),
                     (2022, "EFPC2022-Z6-2"): ("0347", "zona 6")}


def nombres_de(anio, g):
    """{sufijo: nombre del proyecto} segun la fuente de cada anio; solo para confirmar/desempatar."""
    if anio == 2020:
        return {s_: " ".join(dict.fromkeys(v["desc"])) for s_, v in g.items() if v["desc"]}
    if anio == 2024:
        with open(RAIZ / "outputs/indicadores/pp_proyectos.csv", encoding="utf-8-sig", newline="") as f:
            return {r["codigo_proyecto"]: r["nombre_proyecto"] for r in csv.DictReader(f, delimiter=";")
                    if r["anio"] == "2024" and r["en_listado_efpc"] == "True" and re.fullmatch(r"\d{4}", r["codigo_proyecto"])}
    if anio == 2025:
        import openpyxl
        ws = openpyxl.load_workbook(L.BASE / FUENTES[2025][2], read_only=True, data_only=True)["PlanAcción27.01.2026"]
        out = {}
        for r in ws.iter_rows(min_row=4, values_only=True):
            if r[0] and r[1]:
                out.setdefault(str(r[0]).zfill(4), str(r[1]))
        return out
    return {}


def cargar():
    todas = []
    for a in ANIOS:
        efpc = L.leer_efpc(a)
        cont = defaultdict(int)
        for e in efpc:
            cont[e["zona"]] += 1
            e["id"] = f"EFPC{a}-Z{e['zona']}-{cont[e['zona']]}"
            e["pag_efpc"] = e["pagina"]
        filas = L.leer_roca_2025() if a == 2025 else L.leer_hacienda(a)
        g = agrupar(filas)
        r = cruzar(a, efpc, g, nombres_de(a, g))
        for x, e in zip(r, efpc):
            x["pag_efpc"] = e["pag_efpc"]
            if (a, x["id"]) in FORZADOS_NO_VERIFICADO:
                x["obs"] = (f"[{x['tipo']} por regla -> sufijo {x['codigo_sufijo']}] " if x["codigo_sufijo"] else "") + FORZADOS_NO_VERIFICADO[(a, x["id"])]
                x["regla_sufijo"], x["regla_tipo"] = x["codigo_sufijo"], x["tipo"]
                x.update(codigo_sufijo="", ini=None, dfn=None, obl=None, x_dfn=None, x_obl=None, tipo="no verificado", pag="", dif=None, comp=None)
        todas += r
    return todas


def escribir(filas):
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    f_ = lambda v: "" if v is None else (f"{v:.0f}" if abs(v - round(v)) < 1e-9 else f"{v:.2f}")  # no se redondea el valor votado con decimales
    with open(SALIDA, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter=";")
        w.writeheader()
        for x in filas:
            a = x["anio"]
            fv = FUENTES[a]
            pag = x["pag"] if x["pag"] else ""
            w.writerow({"anio": a, "id_efpc": x["id"], "zona": x["zona"], "nombre_efpc": x["nombre_efpc"], "valor_votado": f_(x["valor_votado"]),
                        "codigo_sufijo": x["codigo_sufijo"], "apropiacion_inicial": f_(x["ini"]),
                        "apropiacion_definitiva_sin_2_3_7_06": f_(x["dfn"]), "compromiso_sin_2_3_7_06": f_(x["comp"]), "obligacion_sin_2_3_7_06": f_(x["obl"]),
                        "excluido_2_3_7_06_apropiacion": f_(x["x_dfn"]), "excluido_2_3_7_06_obligacion": f_(x["x_obl"]),
                        "diferencia_votado_menos_inicial": f_(x["dif"]), "tipo_cruce": x["tipo"],
                        "pagina_pdf": (f"{pag}" if a != 2025 else pag) + (f" | listado EFPC p. {x['pag_efpc']}" if pag else f"listado EFPC p. {x['pag_efpc']}"),
                        "fuente_valor": "plan_accion_roca" if a == 2025 else "hacienda",
                        "fuente": fv[0], "enlace_ejecucion": fv[1], "enlace_listado_efpc": URL_EFPC.format(a=a),
                        "fecha_consulta": FECHAS[fv[2]], "fecha_consulta_listado": FECHAS[f"efpc_listados_pp/Proyectos_PP_{a}.pdf"], "observacion": x["obs"]})


def ajustes(T):
    """Tabla de ajustes explicitos respecto de outputs/indicadores/pp_proyectos.csv (que NO se edita)."""
    AJ = []
    old = {}
    with open(RAIZ / "outputs/indicadores/pp_proyectos.csv", encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r["en_listado_efpc"] == "True" and re.fullmatch(r"\d{4}", r["codigo_proyecto"]):
                old[(int(r["anio"]), r["codigo_proyecto"])] = r
    # 1. valores votados emparejados distinto (el emparejamiento "en orden" de scripts/indicadores/pp_epp.py:188-191)
    for x in T:
        if x["anio"] != 2025 or not x["codigo_sufijo"]:
            continue
        o = old.get((2025, x["codigo_sufijo"]))
        if o and abs(float(o["valor_listado_efpc"]) - x["valor_votado"]) >= 1:
            AJ.append({"anio": 2025, "proyecto": x["codigo_sufijo"], "campo": "valor votado (listado EFPC)",
                       "valor_anterior": o["valor_listado_efpc"], "valor_corregido": x["valor_votado"],
                       "motivo": "emparejamiento 'en orden' de pp_epp.py intercambio los valores de los proyectos de la zona 12 que no coincidian al peso",
                       "evidencia": (f"listado EFPC 2025 p. {x['pag_efpc']}, zona {x['zona']}: '{x['nombre_efpc'][:70]}' = "
                                     + f"{round(x['valor_votado']):,}".replace(",", ".") + f"; {x['pag']}: apropiacion inicial {x['ini']:.0f} "
                                     f"({x['tipo']}; " + ("nombre y zona confirmados" if "nombre y zona confirmados" in x["obs"] else "nombre no confirmado, la zona no contradice") + ")"),
                       "efecto": f"cruce {x['tipo']}; apropiacion definitiva vs votado = {100 * (x['dfn'] / x['valor_votado'] - 1):+.2f} %"})
    # 2. D51: rubro 2.3.7.06
    for x in T:
        if x["codigo_sufijo"] and x["x_dfn"]:
            AJ.append({"anio": x["anio"], "proyecto": x["codigo_sufijo"], "campo": "apropiacion y obligacion (sin rubro 2.3.7.06)",
                       "valor_anterior": f"{x['dfn'] + x['x_dfn']:.0f} / {x['obl'] + x['x_obl']:.0f}",
                       "valor_corregido": f"{x['dfn']:.0f} / {x['obl']:.0f}",
                       "motivo": "D51: se excluye el rubro 2.3.7.06 (deficit fiscal, pagos de vigencias anteriores)",
                       "evidencia": f"Hacienda 2024 p. {x['pag']}: rubros de prefijo 23706 con sufijo {x['codigo_sufijo']}",
                       "efecto": f"excluido: apropiacion {x['x_dfn']:.0f}, obligacion {x['x_obl']:.0f}"})
    # 3. D52: forzados a no verificado
    for (a, i), m in FORZADOS_NO_VERIFICADO.items():
        AJ.append({"anio": a, "proyecto": i, "campo": "tipo de cruce", "valor_anterior": "(regla de cruce) exacto", "valor_corregido": "no verificado",
                   "motivo": m, "evidencia": "decisiones.md D52", "efecto": "cuenta en el denominador de EJ90 y FID; fuera del numerador y de EPP"})
    # D54 y D55
    z94 = next(x for x in T if x["id"] == "EFPC2024-Z9-4")
    AJ.append({"anio": 2024, "proyecto": "EFPC2024-Z9-4", "campo": "tipo de cruce", "valor_anterior": "no verificado (D52)", "valor_corregido": f"{z94['tipo']} con sufijo {z94['codigo_sufijo']}",
               "motivo": "D54 (usuaria): se acepta el cruce con 0755; 8 rubros con apropiacion inicial que suman 500.398.021 = valor votado; definitiva y obligacion dan lo mismo",
               "evidencia": "eje2024q4.pdf pp. 80-82 (revision de QA); sector 3206; fuente 15317001002", "efecto": "entra al numerador de EJ90 y FID y a EPP"})
    for (a, i), ref in DUPLICADOS.items():
        AJ.append({"anio": a, "proyecto": i, "campo": "universo", "valor_anterior": "proyecto votado", "valor_corregido": "duplicado del listado (fuera del denominador)",
                   "motivo": f"D55: mismo proyecto que {ref} (0347, $275.041.582)", "evidencia": f"listado EFPC {a} p. 1: dos filas con misma zona, nombre y valor", "efecto": "universo 2022 = 39"})
    for (a, i), (suf, zz) in ASIGNACION_SUFIJO.items():
        if i in DUPLICADOS.values():
            continue
        AJ.append({"anio": a, "proyecto": i, "campo": "sufijo asignado", "valor_anterior": "varias coincidencias de valor", "valor_corregido": suf,
                   "motivo": f"D55: el sufijo {suf} conserva la {zz} en todos los anios (2020 descripciones de Hacienda, 2024 plan de accion)",
                   "evidencia": "valor votado = apropiacion inicial del sufijo al peso", "efecto": "cruce exacto"})
    # 4. proyectos que el cruce anterior trataba distinto (2024 y 2025)
    nuevo = {(x["anio"], x["codigo_sufijo"]): x for x in T if x["codigo_sufijo"]}
    for (a, suf), o in sorted(old.items()):
        if a in (2024, 2025) and (a, suf) not in nuevo:
            AJ.append({"anio": a, "proyecto": suf, "campo": "tipo de cruce", "valor_anterior": o["tipo_correspondencia"], "valor_corregido": "no verificado",
                       "motivo": "la apropiacion inicial de Hacienda no coincide con el valor votado (D52 / regla general)",
                       "evidencia": f"valor votado {o['valor_listado_efpc']}; apropiacion definitiva anterior {o['apropiacion_definitiva']}", "efecto": "fuera de EPP"})
    return AJ


if __name__ == "__main__":
    T = cargar()
    escribir(T)
    AJ = ajustes(T)
    with open(SALIDA_AJ, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["anio", "proyecto", "campo", "valor_anterior", "valor_corregido", "motivo", "evidencia", "efecto"], delimiter=";")
        w.writeheader()
        w.writerows(AJ)
    for a in ANIOS:
        t = [x for x in T if x["anio"] == a]
        c = defaultdict(int)
        for x in t:
            c[x["tipo"]] += 1
        print(a, len(t), dict(c))
    print("ajustes:", len(AJ), "->", SALIDA_AJ)
