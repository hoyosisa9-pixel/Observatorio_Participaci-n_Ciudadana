"""PC.I3 - Cumplimiento de la oferta formativa (principal). PROVISIONAL.

I3 = 100 * V1 / V2; V1 = acciones formativas programadas para la ciudadania que se realizaron; V2 = acciones programadas con
finalizacion prevista en el periodo; V1 > V2 -> Revisar (o tope 100 segun pc_i3_v1_mayor_v2).
Unidad por defecto 'acciones' (informes de gestion): 2025 = 21 culminadas de 25 programadas (inferido); 2024 solo hay culminadas
(7), sin programadas -> N/D. Alternativa 'personas_producto_521' (plan indicativo, consecutivo 521): programadas 10 / 15 y
ejecutadas 95 / 186 (2024 / 2025) -> V1 > V2 -> Revisar.
Soporte: una fila por accion formativa (2024 y 2025) y una fila por anio del producto 521 (personas).
Ejecutar desde la raiz: python3 scripts/indicadores/pc_i3.py
"""
import sys
from pathlib import Path
RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "indicadores"))
import comun as C
import openpyxl
import warnings
warnings.filterwarnings("ignore")

SIS, IND, SCRIPT = "PC", "PC.I3", "scripts/indicadores/pc_i3.py"
NOMBRES = {"PC.I3.V1": "Número de acciones formativas programadas para la ciudadanía que se realizaron",
           "PC.I3.V2": "Número total de acciones formativas programadas para la ciudadanía con finalización prevista en el periodo"}
_B = "data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/"
DOCS = {"G25": ("Informe de gestión por dependencias 2025", _B + "anon-2026/2025informegestiondependencias.pdf"),
        "G24": ("Informe de gestión por dependencias 2024",
                _B + "ano-2025-298942/24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf"),
        "PI26": ("Plan Indicativo de Productos, evaluación definitiva 2T-2026", _B +
                 "informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx")}
FILA_521 = 524


def leer_521():
    ws = openpyxl.load_workbook(RAIZ / DOCS["PI26"][1], data_only=True)["Plan-Indicativo-Productos"]
    r = next(f for f in ws.iter_rows(min_row=FILA_521, max_row=FILA_521, values_only=True))
    assert r[0] == 521 and r[39] == "Número" and "Personas formadas" in r[38], r[:3]
    return dict(prog24=int(r[44]), prog25=int(r[45]), ej24=int(r[64]), ej25=int(r[65]))


def enl(k):
    return C.url_de_ruta(DOCS[k][1])


def construir_soporte(d521):
    filas = []
    n = {2024: 0, 2025: 0}

    def add(anio, tipo, nombre, culm, prog, serv, clave, pag, just, vars_="PC.I3.V1;PC.I3.V2", et="accion_formativa"):
        n[anio] += 1
        doc, ruta = DOCS[clave]
        incl = culm == 1 and serv == "no"
        filas.append(dict(
            id_soporte=f"PC-I3-{anio}-{n[anio]:03d}", codigo_indicador=IND, id_variable=vars_, anio=anio,
            escenario_id="defecto", elemento_tipo=et, elemento_id=f"{anio}-{tipo.upper()}-{n[anio]:02d}",
            descripcion=nombre, fecha_elemento=None, documento=doc, ruta_local=ruta, enlace=enl(clave), pagina=pag,
            banderas=f"tipo={tipo};programada={prog};culminada={culm};servidores={serv};unidad=acciones",
            clasificacion=("programada y culminada" if culm == 1 and prog == "1" else
                           "programada no ejecutada" if culm == 0 else "culminada, programación no publicada"),
            justificacion=just, criterio="Cuenta si es curso, diplomado o seminario completo para la ciudadanía, culminado en "
            "el periodo; no cuentan clases, semilleros ni capacitación por demanda; formación a servidores se excluye "
            "(pc_i3_excluir_servidores)", lector="ejecutor_PC", incluido="sí" if incluida(culm, serv) else "no",
            cuenta_en_defecto=1 if incluida(culm, serv) else 0))

    def incluida(culm, serv):
        return culm == 1 and serv == "no"

    p25 = "p. 88 (Oferta Académica y Observaciones); p. 87 (producto de funcionarios formados)"
    j25 = "El informe reporta culminados 3 diplomados, 9 seminarios y 9 cursos; el detalle por seminario y curso no se publica"
    add(2025, "diplomado", "Diplomado en Gestión Pública", 1, "1", "probable", "G25", p25,
        j25 + ". Posible formación para servidores: el informe reporta 47 funcionarios certificados y nombra este diplomado; "
        "no se acredita el público. Se trata como servidores=probable (excluida con pc_i3_excluir_servidores=true)")
    add(2025, "diplomado", "Diplomado en Derechos Humanos de las Mujeres", 1, "1", "no", "G25", p25, j25)
    add(2025, "diplomado", "Diplomado en Gestión Pública para el Desarrollo Local", 1, "1", "no", "G25", p25, j25)
    for k in range(1, 10):
        add(2025, "seminario", f"Seminario culminado {k} de 9 (sin nombre publicado)", 1, "1", "no", "G25", p25, j25)
    for k in range(1, 10):
        add(2025, "curso", f"Curso culminado {k} de 9 (sin nombre publicado)", 1, "1", "no", "G25", p25, j25)
    jn = ("Observaciones del informe: no ejecución de 1 diplomado y 3 seminarios inicialmente planeados (presupuesto "
          "reorientado al Observatorio). No se reportan cursos no ejecutados; programación 2025 inferida = culminadas + no ejecutadas")
    add(2025, "diplomado", "Diplomado planeado y no ejecutado", 0, "1", "no", "G25", "p. 88 (Observaciones)", jn)
    for k in range(1, 4):
        add(2025, "seminario", f"Seminario planeado y no ejecutado {k} de 3", 0, "1", "no", "G25", "p. 88 (Observaciones)", jn)
    p24 = "pp. 62-63 (proyecto 0541, actividades de la Escuela) y PI26 consecutivo 521, col. EG"
    j24 = ("Listado de formaciones realizadas en 2024 (informe de gestión y observaciones del plan indicativo, con personas "
           "certificadas); no hay programación publicada de acciones, por eso V2 2024 = N/D")
    for tipo, nom in [("diplomado", "Diplomado en Planeación para el Desarrollo"),
                      ("diplomado", "Diplomado en Liderazgo Social y Comunitario"),
                      ("curso", "Curso-taller Gestión Documental"), ("curso", "Curso Ley 1757 de 2015 y Control Social"),
                      ("curso", "Curso Contratación Estatal"), ("seminario", "Seminario Conciliación Comunal"),
                      ("catedra", "Cátedra de Estudios Afrocolombianos")]:
        add(2024, tipo, nom, 1, "desconocida", "no", "G24", p24, j24 + ("; la cátedra no es curso, diplomado ni seminario "
            "en sentido estricto: se cuenta con duda" if tipo == "catedra" else ""))
    # personas (producto 521)
    for anio, prog, ej in ((2024, d521["prog24"], d521["ej24"]), (2025, d521["prog25"], d521["ej25"])):
        n[anio] += 1
        filas.append(dict(
            id_soporte=f"PC-I3-{anio}-{n[anio]:03d}", codigo_indicador=IND, id_variable="PC.I3.V1;PC.I3.V2", anio=anio,
            escenario_id="defecto", elemento_tipo="meta", elemento_id=f"{anio}-PRODUCTO-521-PERSONAS",
            descripcion=f"Producto 521 'Personas formadas para la participación incidente': programadas {prog}, ejecutadas {ej}",
            documento=DOCS["PI26"][0], ruta_local=DOCS["PI26"][1], enlace=enl("PI26"),
            pagina=f"Hoja Plan-Indicativo-Productos, fila {FILA_521} (consecutivo 521): col. {'AS' if anio == 2024 else 'AT'} "
                   f"programada, col. {'BM' if anio == 2024 else 'BN'} ejecutada",
            banderas=f"tipo=producto521;programadas={prog};ejecutadas={ej};unidad=personas_producto_521",
            clasificacion="meta anual de personas", justificacion="Unidad personas (no acciones); la ejecución supera la meta "
            "porque la meta está subestimada (el informe 2025 recomienda ajustarla). Producto separado del de funcionarios",
            criterio="Solo para el escenario pc_i3_unidad=personas_producto_521", lector="ejecutor_PC",
            cuenta_en_defecto=0, incluido="no"))
    return filas


def main():
    d521 = leer_521()
    assert (d521["prog24"], d521["prog25"], d521["ej24"], d521["ej25"]) == (10, 15, 95, 186), d521
    soporte = construir_soporte(d521)
    C.escribir_parcial(SIS, "soporte_conteos", soporte, codigos=[IND])
    valores, resultados = [], []
    base = dict(codigo_indicador=IND, categoria_disponibilidad="Parcial", script=SCRIPT, tipo_valor="observado")
    for esc in C.escenarios(IND):
        p = esc["params"]
        for anio in (2024, 2025):
            fc = C.fecha_corte_registro(anio, "Anual", None, p)
            ac = [f for f in soporte if f["anio"] == anio and "unidad=acciones" in f["banderas"]]
            if p["pc_i3_unidad"] == "acciones":
                def ok(f, solo_culm):
                    b = dict(kv.split("=", 1) for kv in f["banderas"].split(";"))
                    if p["pc_i3_excluir_servidores"] == "true" and b["servidores"] in ("probable", "si"):
                        return False
                    return (b["culminada"] == "1") if solo_culm else b["programada"] == "1"
                ids1 = [f["id_soporte"] for f in ac if ok(f, True)]
                ids2 = [f["id_soporte"] for f in ac if ok(f, False)]
                v1 = len(ids1)
                v2 = len(ids2) if ids2 else None
                clave = "G25" if anio == 2025 else "G24"
                pag1 = "p. 88" if anio == 2025 else "pp. 62-63; PI26 consecutivo 521 col. EG"
                obs2 = ("Inferido: culminadas + no ejecutadas (1 diplomado y 3 seminarios); no se reportan cursos no ejecutados"
                        if anio == 2025 else "")
                unid = "Acciones"
            else:
                v1 = d521["ej24"] if anio == 2024 else d521["ej25"]
                v2 = d521["prog24"] if anio == 2024 else d521["prog25"]
                ids1 = ids2 = [f["id_soporte"] for f in soporte if f["anio"] == anio and "producto521" in f["banderas"]]
                clave, pag1 = "PI26", f"Fila {FILA_521} (consecutivo 521), col. " + ("BM" if anio == 2024 else "BN")
                obs2 = ""
                unid = "Personas"
            serv_obs = (" Se excluyó el diplomado en Gestión Pública (probable formación a servidores)."
                        if p["pc_i3_unidad"] == "acciones" and p["pc_i3_excluir_servidores"] == "true" and anio == 2025 else "")
            valores.append(C.fila_valor(SIS, "PC.I3.V1", anio, v1, "Conteo", esc, nombre_variable=NOMBRES["PC.I3.V1"],
                estado_informacion="Pendiente", **base, fuente=DOCS[clave][0], documento_ruta_local=DOCS[clave][1],
                enlace=enl(clave), pagina_tabla=pag1, fecha_corte=fc, n_elementos_soporte=len(ids1),
                ids_soporte=",".join(ids1), observaciones=f"Unidad: {unid.lower()}." + serv_obs +
                (" En 2024 solo hay acciones culminadas (sin programadas)." if unid == "Acciones" and anio == 2024 else "")))
            if v2 is None:
                valores.append(C.fila_valor(SIS, "PC.I3.V2", anio, None, "Conteo", esc, nombre_variable=NOMBRES["PC.I3.V2"],
                    estado_informacion="No disponible", **base, fecha_corte=fc,
                    observaciones="2024: no se publica programación de acciones formativas; solo la meta en personas (consecutivo 521)"))
            else:
                valores.append(C.fila_valor(SIS, "PC.I3.V2", anio, v2, "Conteo", esc, nombre_variable=NOMBRES["PC.I3.V2"],
                    estado_informacion="Pendiente", **base, fuente=DOCS[clave][0], documento_ruta_local=DOCS[clave][1],
                    enlace=enl(clave), pagina_tabla=("p. 88 (Observaciones)" if unid == "Acciones" else
                                                     f"Fila {FILA_521}, col. " + ("AS" if anio == 2024 else "AT")),
                    fecha_corte=fc, n_elementos_soporte=len(ids2), ids_soporte=",".join(ids2),
                    observaciones=f"Unidad: {unid.lower()}. {obs2}".strip() + serv_obs))
            res = C.calcular(IND, [v1, v2])
            alerta = ""
            if res["disponibilidad"] == "Revisar" and p["pc_i3_v1_mayor_v2"] == "tope_100" and v2:
                res = C.calcular(IND, [v2, v2])
                res["motivo"] = ""
                alerta = f"V1 ({v1}) mayor que V2 ({v2}): puntaje limitado a 100 por el escenario tope_100"
            elif res["disponibilidad"] == "Revisar":
                alerta = f"V1 ({v1}) > V2 ({v2}): la meta está subestimada o V1 no pertenece a lo programado; revisar"
            elif res["disponibilidad"] != "Calculado" and v2 is None:
                res["motivo"] = "2024 sin programación de acciones formativas publicada (V2 N/D)"
            resultados.append(C.fila_resultado(IND, anio, esc, res, [v1, v2], sistema=C._sis_nombre(SIS),
                nombre="Cumplimiento de la oferta formativa", tipo="Principal", fecha_corte_comun=fc, script=SCRIPT,
                alerta=alerta, observaciones=f"Unidad: {unid.lower()}; solo Municipio"))
    C.escribir_parcial(SIS, "valores_variables", valores, codigos=[IND])
    C.escribir_parcial(SIS, "resultados_indicadores", resultados, codigos=[IND])
    print(f"soporte={len(soporte)} valores={len(valores)} resultados={len(resultados)}")
    from collections import Counter
    for r in resultados:
        if r["es_defecto"] or (r["escenario_id"].count("pc_i3_unidad=personas") and r["escenario_id"].count(";") <= 1):
            print(r["escenario_id"][:70], r["anio"], r["entrada_1"], r["entrada_2"], r["puntaje"], r["semaforo"], r["disponibilidad"])
    print(Counter((r["anio"], r["disponibilidad"], r["semaforo"]) for r in resultados))


if __name__ == "__main__":
    main()
