#!/usr/bin/env python3
"""PP, tanda 2: variables de los indicadores NO viables (PP.CEP, PP.CMA). NO calcula CEP, CMA ni IREPP.

Variables que construye (todas PROVISIONAL, estado "Pendiente" o "No disponible"; nada "Validado"):
  PP.CEP.PE   proyectos financiados con entrega exigible en la vigencia (2024, 2025, 2026), municipal y zonas 2-13.
  PP.CEP.PEC  PROXY de autorreporte: proyectos con todos sus productos al 100 % en el plan de acción IV-trim-2025
              (sin acta de recibo ni informe final), con las 3 reglas de `pp_cep_regla_agregacion`. Solo 2025.
  PP.CMA.NA   proyectos con meta de atención aplicable, según `pp_cma_aplicabilidad`.
  PP.CMA.PA / PP.CMA.MP  personas atendidas / meta de atención: N/D razonado (no hay datos por proyecto).

Escribe SOLO en outputs/indicadores/parciales/PP/ con comun.escribir_parcial:
  valores_variables (codigos=["PP.CEP","PP.CMA"]), soporte_conteos (codigos=["PP.CEP"]; una fila por proyecto) y
  clasificacion_pp_entregas (una fila por proyecto x año; tabla propia). NO toca pp_proyectos de pp_epp.py.

Uso (desde la raíz, SIN -I):  python3 scripts/indicadores/pp_no_viables.py
"""
from __future__ import annotations

import collections
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
import pymupdf

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
sys.path.insert(0, str(RAIZ / "scripts"))
import comun as C       # noqa: E402
import pp_epp as P      # noqa: E402  (lectura del plan de acción y universos; se importa, no se modifica)

SCRIPT = "scripts/indicadores/pp_no_viables.py"
F_EFPC24 = RAIZ / "data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf"
ZONAS = list(range(2, 14))
REGLAS = ["todas_filas_dic_100", "solo_31dic", "ultima_fila_por_indicador"]
COL_REGLA = {"todas_filas_dic_100": "pec_todas_filas_dic_100", "solo_31dic": "pec_solo_31dic",
             "ultima_fila_por_indicador": "pec_ultima_fila"}
ESPERADO_PE = {2025: 61, 2026: 66}
RANGO_PEC_2025 = (45, 55)   # rango citado por el plan (la regla 3 da un valor propio; ver observaciones)


def rel(p: Path) -> str:
    return str(p.relative_to(RAIZ))


# ----------------------------------------------------------------------------------------------
# Lectura de fuentes
# ----------------------------------------------------------------------------------------------
def leer_efpc24() -> list[dict]:
    """Listado EFPC 2024: una entrada por proyecto (zona, texto inicial, valor). El valor puede venir sin '$'."""
    regs, cur = [], None
    for pi, pg in enumerate(pymupdf.open(F_EFPC24)):
        for l in pg.get_text().splitlines():
            l = l.strip()
            if re.fullmatch(r"\d{1,2}", l):
                cur = {"zona": int(l), "t": [], "pag": pi + 1, "valor": None, "valor_sin_signo": False}
                regs.append(cur)
            elif cur is None or not l:
                continue
            elif re.fullmatch(r"\$\s*[\d.,]+", l):
                cur["valor"] = int(re.sub(r"\D", "", l))
            elif re.fullmatch(r"\d{1,3}(\.\d{3})+", l) and cur["valor"] is None:
                cur["valor"], cur["valor_sin_signo"] = int(l.replace(".", "")), True
            else:
                cur["t"].append(l)
    for r in regs:
        assert r["valor"] is not None, r
    cont = collections.Counter()
    for r in regs:
        cont[r["zona"]] += 1
        r["n"] = cont[r["zona"]]
        r["nombre"] = " ".join(r["t"])[:110]
    return regs


def filas_fisicas_2025(codigos: set) -> dict:
    """Por proyecto: filas (fecha, ejecución física neta) del plan de acción IV-trim-2025, agrupadas por
    indicador de producto (o por actividad si la fila no trae indicador)."""
    ws = openpyxl.load_workbook(P.F_PLAN25, read_only=True, data_only=True)[P.HOJA_PLAN25]
    out = collections.defaultdict(lambda: collections.defaultdict(list))
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] in codigos and r[8] is not None and isinstance(r[8], dt.datetime):
            neta = r[12] if isinstance(r[12], (int, float)) else None
            out[r[0]][(r[4] or r[5] or "")].append((r[8].date(), neta, isinstance(r[12], str)))
    return out


def evaluar_pec(grupos: dict) -> dict:
    """Aplica las 3 reglas a un proyecto. Valor no numérico (p. ej. 'Error en la programación') = no cumple."""
    todas = [f for fs in grupos.values() for f in fs]
    dic = [f for f in todas if f[0].month == 12]
    dic31 = [f for f in dic if f[0].day == 31]
    ok = lambda fs: bool(fs) and all(f[1] is not None and f[1] >= 1 - 1e-9 for f in fs)   # noqa: E731
    ultimas = []
    for fs in grupos.values():
        m = max(f[0] for f in fs)
        ultimas += [f for f in fs if f[0] == m]
    return {"pec_todas_filas_dic_100": int(ok(dic)), "pec_solo_31dic": int(ok(dic31)),
            "pec_ultima_fila": int(ok(ultimas)), "n_filas_dic": len(dic),
            "n_filas_dic_no_numericas": sum(1 for f in dic if f[2]),
            "n_productos": len(grupos)}


def tipologia(nombre: str) -> tuple[str, str]:
    """Tipología sugerida por palabras del nombre (heurística, a verificar por el equipo)."""
    n = nombre.upper()
    if "ACUEDUCTO" in n or "INFRAESTRUCTURA VERDE" in n or "RECUPERACIÓN DE ESPACIOS URBANOS" in n \
            or "RECUPERACION DE ESPACIOS URBANOS" in n:
        return "infraestructura", "Obra o infraestructura: sin personas atendidas como unidad de meta"
    if "MEDIO DE COMUNICACI" in n or "MEDIO ALTERNATIVO" in n or "MEDIO DE ALTERNATIVO" in n:
        return "medios", "Medio de comunicación comunitaria: sin personas atendidas como unidad de meta"
    return "servicio a personas", ""


# ----------------------------------------------------------------------------------------------
def main():
    sup = C.cargar_supuestos()
    esc_cep = C.escenarios("PP.CEP", sup)
    esc_cma = C.escenarios("PP.CMA", sup)
    assert esc_cep[0]["es_defecto"] and esc_cma[0]["es_defecto"]
    pdef = esc_cep[0]["params"]
    regla_def = pdef["pp_cep_regla_agregacion"]

    # --- universo de proyectos por año --------------------------------------------------------
    proy = []     # dicts: anio, zona, codigo, nombre, es_0553, doc, pagina...
    efpc24 = leer_efpc24()
    for r in efpc24:
        proy.append({"anio": 2024, "zona": r["zona"], "codigo": f"EFPC2024-Z{r['zona']}-{r['n']}", "nombre": r["nombre"],
                     "es_0553": False, "doc": rel(F_EFPC24), "pagina": f"p. {r['pag']}; valor {r['valor']:,}".replace(",", "."),
                     "obs": ("Valor en el PDF sin signo '$' (403.014.352); la nota de disponibilidad contaba 50 proyectos sin este registro. "
                             "Verificar con Planeación" if r["valor_sin_signo"] else ""),
                     "dudoso": r["valor_sin_signo"]})
    u = {2025: P.universo(2025), 2026: P.universo(2026)}
    for anio, uu in u.items():
        for c, p in sorted(uu["pp"].items(), key=lambda kv: (kv[1]["zona"] or 99, kv[0])):
            proy.append({"anio": anio, "zona": p["zona"], "codigo": c, "nombre": p["nombre"], "es_0553": p["es_0553"],
                         "doc": rel(uu["f_plan"]), "pagina": f"Hoja {uu['hoja']}", "obs": "", "dudoso": False})
    fis = filas_fisicas_2025({p["codigo"] for p in proy if p["anio"] == 2025 and not p["es_0553"]})

    for p in proy:
        p["pe"] = 0 if p["es_0553"] else 1
        p["tipo"], p["motivo"] = tipologia(p["nombre"])
        p["aplica_tip"] = "No" if p["tipo"] != "servicio a personas" else "Sí"
        if p["es_0553"]:
            p["tipo"], p["motivo"], p["aplica_tip"] = "otro", "Costo del proceso (asambleas), no un proyecto priorizado", "No"
        p["pec"] = {}
        if p["anio"] == 2025 and not p["es_0553"]:
            p["pec"] = evaluar_pec(fis[p["codigo"]])

    # --- regresiones --------------------------------------------------------------------------
    pe = {a: sum(p["pe"] for p in proy if p["anio"] == a) for a in (2024, 2025, 2026)}
    assert pe[2025] == 61 and pe[2026] == 66, pe
    pe24_sin_dudoso = sum(1 for p in proy if p["anio"] == 2024 and not p["dudoso"])
    assert pe24_sin_dudoso == 50, pe24_sin_dudoso
    pec = {r: sum(p["pec"][COL_REGLA[r]] for p in proy if p["anio"] == 2025 and p["pec"]) for r in REGLAS}
    assert pec["todas_filas_dic_100"] == 49 and pec["solo_31dic"] == 45, pec

    # --- tabla propia: clasificacion_pp_entregas ---------------------------------------------
    cl = []
    for p in proy:
        d = {"anio": p["anio"], "zona": p["zona"] if p["zona"] is not None else "N/A", "codigo_proyecto": p["codigo"],
             "nombre_proyecto": p["nombre"], "es_0553": p["es_0553"], "pe_cuenta": p["pe"],
             "cuenta_en_defecto_pe": 1 if p["pe"] else 0}
        for r in REGLAS:
            d[COL_REGLA[r]] = p["pec"].get(COL_REGLA[r], "")
        d["pec_regla_defecto"] = regla_def
        d["pec_defecto"] = p["pec"].get(COL_REGLA[regla_def], "")
        d["n_productos_o_actividades"] = p["pec"].get("n_productos", "")
        d["n_filas_dic"] = p["pec"].get("n_filas_dic", "")
        d["n_filas_dic_no_numericas"] = p["pec"].get("n_filas_dic_no_numericas", "")
        d["tipologia_sugerida"] = p["tipo"]
        d["aplica_meta_atencion_defecto"] = "Pendiente"
        d["aplica_meta_atencion_tipologia"] = p["aplica_tip"]
        d["motivo_exclusion"] = p["motivo"]
        d["personas_atendidas_PA"] = ""
        d["meta_atencion_MP"] = ""
        d["estado_informacion"] = "Pendiente"
        d["fuente"] = p["doc"]
        d["pagina_hoja"] = p["pagina"]
        obs = []
        if p["anio"] == 2025 and p["pec"]:
            obs.append("PEC = autorreporte del plan de acción, no entrega verificada (sin acta de recibo ni informe final)")
        if p["anio"] == 2026:
            obs.append("2026: corte 30-jun-2026; no hay cierre de vigencia ni avance físico anual")
        if p["obs"]:
            obs.append(p["obs"])
        if p["anio"] == 2024:
            obs.append("Nombre truncado del listado EFPC (texto inicial); sin avance físico 2024 en el repositorio")
        d["observaciones"] = ". ".join(obs)
        cl.append(d)
    # sin datos personales (lo valida escribir_parcial para observaciones; se revisa aquí el nombre)
    for d in cl:
        assert not C.validar_sin_datos_personales(d["nombre_proyecto"]), d["codigo_proyecto"]

    # --- valores_variables --------------------------------------------------------------------
    def terr_cuenta(anio, terr, f):
        if terr == "Municipio":
            return sum(f(p) for p in proy if p["anio"] == anio)
        z = int(terr.split()[1])
        return sum(f(p) for p in proy if p["anio"] == anio and p["zona"] == z)

    def meta_fuente(anio):
        if anio == 2024:
            return rel(F_EFPC24), P.enlace(F_EFPC24, "https://efpc.envigado.gov.co/wp-content/uploads/2025/07/Proyectos_PP_2024.pdf"), \
                "Listado EFPC 2024: un registro por proyecto (zona, nombre, valor)", dt.date(2024, 12, 31)
        uu = u[anio]
        real = dt.date(2026, 1, 7) if anio == 2025 else dt.date(2026, 6, 30)
        return rel(uu["f_plan"]), P.enlace(uu["f_plan"]), \
            f"Hoja {uu['hoja']}: proyectos con 'ZONA n' en el nombre (sin 0553)", real

    filas_v = []
    for e in esc_cep:
        prm = e["params"]
        z1 = prm["pp_zona1_tratamiento"]
        for anio in (2024, 2025, 2026):
            doc, enl, pag, real = meta_fuente(anio)
            fc = C.fecha_corte_registro(anio, "Anual", real, prm)
            terrs = ["Municipio"] + [f"Zona {z}" for z in range(1, 14)]
            for terr in terrs:
                zero1 = terr == "Zona 1"
                obs_base = f"Escenario: {'; '.join(f'{k}={v}' for k, v in e['desviaciones'].items()) or 'defecto'}"
                # ---- PE
                if zero1:
                    vpe, cat, est = None, ("NE" if z1 == "NE" else "N/D"), ("No evaluable" if z1 == "NE" else "No disponible")
                    obs = "Zona 1 sin proyectos PP (PDM p. 558: 13 zonas, 12 comités); la Guía exige confirmarlo con Planeación. " + obs_base
                else:
                    vpe = terr_cuenta(anio, terr, lambda p: p["pe"])
                    cat, est = "Construible", "Pendiente"
                    obs = ("Proyectos con entrega exigible en la vigencia = proyectos financiados con apropiación en la vigencia (excluye el 0553, "
                           "costo del proceso). " + obs_base)
                    if anio == 2024:
                        obs += (". Listado EFPC 2024 con 51 registros; la nota de disponibilidad contaba 50 (omitía el registro de la zona 8 "
                                "con valor sin signo '$'). Aquí se cuentan 51; con 50 el municipal sería 50 y la zona 8 tendría 4")
                    if anio == 2026:
                        obs += ". Vigencia 2026 en curso: universo del plan de acción II trimestre (corte 30-jun-2026)"
                filas_v.append(C.fila_valor(
                    "PP", "PP.CEP.PE", anio, vpe, "Conteo", e, nombre_variable="Proyectos con entrega exigible en la vigencia (PE)",
                    tipo_valor="construido", periodo="Anual", territorio=terr, fuente=doc, enlace=enl, pagina_tabla=pag,
                    fecha_corte=fc, fecha_corte_real=real, estado_informacion=est, categoria_disponibilidad=cat,
                    provisional=True, script=SCRIPT, observaciones=obs))
                # ---- PEC
                regla = prm["pp_cep_regla_agregacion"]
                if zero1:
                    vpec, cat, est = None, ("NE" if z1 == "NE" else "N/D"), ("No evaluable" if z1 == "NE" else "No disponible")
                    obs = "Zona 1 sin proyectos PP. " + obs_base
                elif anio == 2025:
                    vpec = terr_cuenta(anio, terr, lambda p: p["pec"].get(COL_REGLA[regla], 0))
                    cat, est = "Proxy", "Pendiente"
                    obs = (f"PROXY de autorreporte (regla {regla}): proyectos con todos sus productos/actividades al 100 % de ejecución física "
                           "neta en el plan de acción IV-trim-2025; NO es entrega verificada (sin acta de recibo ni informe final). "
                           "Valores no numéricos del plan ('Error en la programación') cuentan como no cumplidos. " + obs_base)
                else:
                    vpec, cat, est = None, "No disponible", "No disponible"
                    obs = ("Sin avance físico por proyecto equivalente al de 2025 en el repositorio" if anio == 2024
                           else "Avance físico a 30-jun-2026 es parcial; el cierre de vigencia no existe aún") + ". " + obs_base
                filas_v.append(C.fila_valor(
                    "PP", "PP.CEP.PEC", anio, vpec, "Conteo", e, nombre_variable="Proyectos con entrega cumplida (PEC), proxy de autorreporte",
                    tipo_valor="proxy" if anio == 2025 else "construido", periodo="Anual", territorio=terr, fuente=doc, enlace=enl,
                    pagina_tabla=pag + ("; col. EJECUCIÓN FÍSICA NETA" if anio == 2025 else ""),
                    fecha_corte=fc, fecha_corte_real=real, estado_informacion=est, categoria_disponibilidad=cat,
                    provisional=True, script=SCRIPT, observaciones=obs))

    # CMA: NA (municipal, con la aplicabilidad), PA y MP (N/D razonado)
    for e in esc_cma:
        prm = e["params"]
        modo = prm["pp_cma_aplicabilidad"]
        for anio in (2024, 2025, 2026):
            doc, enl, pag, real = meta_fuente(anio)
            fc = C.fecha_corte_registro(anio, "Anual", real, prm)
            obs_base = f"Escenario: {'; '.join(f'{k}={v}' for k, v in e['desviaciones'].items()) or 'defecto'}"
            if modo == "pendiente":
                vna, cat, est = None, "N/D", "No disponible"
                obs = ("Aplicabilidad de la meta de atención pendiente de decisión del equipo (pp_cma_aplicabilidad=pendiente); "
                       "la Matriz exige el universo completo verificado. Tipología sugerida en clasificacion_pp_entregas. " + obs_base)
            else:
                vna = sum(1 for p in proy if p["anio"] == anio and p["aplica_tip"] == "Sí")
                cat, est = "Parcial", "Pendiente"
                obs = ("Conteo de proyectos de 'servicio a personas' según tipología sugerida por palabras del nombre (heurística, "
                       "no verificada); infraestructura y medios = No con motivo. " + obs_base)
            zterr = ["Municipio"] + [f"Zona {z}" for z in range(1, 14)]
            for terr in zterr:
                if terr == "Zona 1":
                    z1 = prm["pp_zona1_tratamiento"]
                    v, c_, s_ = None, ("NE" if z1 == "NE" else "N/D"), ("No evaluable" if z1 == "NE" else "No disponible")
                    o = "Zona 1 sin proyectos PP. " + obs_base
                elif terr == "Municipio":
                    v, c_, s_, o = vna, cat, est, obs
                else:
                    z = int(terr.split()[1])
                    v = None if vna is None else sum(1 for p in proy if p["anio"] == anio and p["zona"] == z and p["aplica_tip"] == "Sí")
                    c_, s_, o = cat, est, obs
                filas_v.append(C.fila_valor(
                    "PP", "PP.CMA.NA", anio, v, "Conteo", e, nombre_variable="Proyectos con meta de atención aplicable (NA)",
                    tipo_valor="construido", periodo="Anual", territorio=terr, fuente=doc, enlace=enl, pagina_tabla=pag,
                    fecha_corte=fc, fecha_corte_real=real, estado_informacion=s_, categoria_disponibilidad=c_,
                    provisional=True, script=SCRIPT, observaciones=o))
            # PA y MP: solo municipal y solo con la configuración por defecto (no dependen del escenario)
            if e["es_defecto"]:
                for idv, nom, motivo in (
                        ("PP.CMA.PA", "Personas atendidas (PA), por proyecto",
                         "No hay personas únicas atendidas por proyecto. " + (
                             "En 2024 solo hay atenciones por componente sin deduplicar (INDER, rendición PP 2024 pp. 5-9) y 'inscritos' o "
                             "'asistentes' de algunos operadores; no distinguen proyecto." if anio == 2024
                             else "El plan de acción mide actividades, porcentajes y paquetes, no personas.")),
                        ("PP.CMA.MP", "Meta de atención (MP), por proyecto",
                         "No hay meta anual de personas verificable por proyecto; solo menciones aisladas de población objeto de algunos operadores "
                         "(p. ej. Fundarte zona 9, ecozona zona 10 en 2024). El plan de acción no la reporta.")):
                    filas_v.append(C.fila_valor(
                        "PP", idv, anio, None, "Conteo", e, nombre_variable=nom, tipo_valor="construido", periodo="Anual",
                        territorio="Municipio", fuente=doc, enlace=enl, pagina_tabla=pag, fecha_corte=fc, fecha_corte_real=real,
                        estado_informacion="No disponible", categoria_disponibilidad="N/D", provisional=True, script=SCRIPT,
                        observaciones=motivo + " " + obs_base))

    # --- soporte_conteos: una fila por proyecto (escenario por defecto) -----------------------
    filas_s = []
    doc_enl = {a: meta_fuente(a) for a in (2024, 2025, 2026)}
    for p in proy:
        if p["es_0553"]:
            continue
        a = p["anio"]
        doc, enl, pag, _ = doc_enl[a]
        terr = f"Zona {p['zona']}"
        base = {"codigo_indicador": "PP.CEP", "anio": a, "periodo": "Anual", "territorio": terr, "escenario_id": "defecto",
                "elemento_tipo": "proyecto", "elemento_id": p["codigo"], "descripcion": p["nombre"][:200], "documento": Path(doc).name,
                "ruta_local": doc, "enlace": enl, "pagina": p["pagina"], "lector": "ejecutor_PP"}
        filas_s.append({**base, "id_variable": "PP.CEP.PE", "banderas": f"zona={p['zona']};vigencia={a}"
                        + (";registro_dudoso=1" if p["dudoso"] else ""),
                        "clasificacion": "proyecto con entrega exigible",
                        "justificacion": "Proyecto financiado de la vigencia" + ("; valor sin signo '$' en el listado" if p["dudoso"] else ""),
                        "cuenta_en_defecto": 1})
        if a == 2025:
            pc = p["pec"]
            filas_s.append({**base, "id_variable": "PP.CEP.PEC",
                            "banderas": (f"todas_filas_dic_100={pc['pec_todas_filas_dic_100']};solo_31dic={pc['pec_solo_31dic']};"
                                         f"ultima_fila={pc['pec_ultima_fila']};n_filas_dic={pc['n_filas_dic']};"
                                         f"n_no_numericas={pc['n_filas_dic_no_numericas']}"),
                            "clasificacion": "autorreporte 100 %" if pc[COL_REGLA[regla_def]] else "autorreporte < 100 %",
                            "justificacion": f"Regla por defecto {regla_def}; autorreporte del plan de acción, sin acta de recibo",
                            "cuenta_en_defecto": pc[COL_REGLA[regla_def]]})

    ruta_v = C.escribir_parcial("PP", "valores_variables", filas_v, codigos=["PP.CEP", "PP.CMA"])
    ruta_s = C.escribir_parcial("PP", "soporte_conteos", filas_s, codigos=["PP.CEP"])
    ruta_c = C.escribir_parcial("PP", "clasificacion_pp_entregas", cl)

    # --- resumen ------------------------------------------------------------------------------
    print("PE:", pe, "(2024 sin el registro dudoso: 50)")
    for a in (2024, 2025, 2026):
        zs = [terr_cuenta(a, f"Zona {z}", lambda p: p["pe"]) for z in ZONAS]
        print(f"  PE {a}: municipal={pe[a]} zonas 2-13: {min(zs)}-{max(zs)}")
    print("PEC 2025 por regla:", pec)
    for r in REGLAS:
        zs = [sum(p["pec"][COL_REGLA[r]] for p in proy if p["anio"] == 2025 and p["pec"] and p["zona"] == z) for z in ZONAS]
        print(f"  {r}: municipal={pec[r]} zonas {min(zs)}-{max(zs)}: {zs}")
    for r in ("todas_filas_dic_100",):
        zt = [(z, sum(p["pec"][COL_REGLA[r]] for p in proy if p["anio"] == 2025 and p["pec"] and p["zona"] == z),
               sum(1 for p in proy if p["anio"] == 2025 and p["zona"] == z)) for z in ZONAS]
        print("  PEC defecto por zona (cumple/total):", zt)
    tip = collections.Counter((p["anio"], p["aplica_tip"]) for p in proy if not p["es_0553"])
    print("Aplica meta (tipología sugerida):", dict(tip))
    print(f"Filas: valores={len(filas_v)}, soporte={len(filas_s)}, clasificacion={len(cl)}")
    print("Escrito:", rel(ruta_v), rel(ruta_s), rel(ruta_c), sep="\n  ")


if __name__ == "__main__":
    main()
