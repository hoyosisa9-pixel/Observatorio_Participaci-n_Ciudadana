#!/usr/bin/env python3
"""Series anuales de Envigado (05266) para la dimensión PDM (revisión de nov-2026). Provisional.

Descarga (cada fuente en su carpeta nueva de data_raw_IIC/07_revision_nov2026/, con manifiesto.csv):
  terridata/            archivo masivo de TerriData de Envigado (DNP)
  men_nudc7mev/         cobertura neta del MEN (datos.gov.co, nudc-7mev), filtrada por municipio
  policia_homicidio/    m8fd-ahd9 (conteo anual agregado por SoQL)
  policia_hurto_personas/ 4rxi-8m8d (ídem)
  policia_lesiones/     jr6v-i33g (ídem)
  policia_transito/     ha6j-pa2r (filas de Envigado, código 05266000)
Población: UNA sola proyección DANE (PPED 2018-2042, total municipal) ya extraída en
outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv (misma que usa calcular.py para salud).

Salida: outputs/revision_indicadores/PDM/series_pdm.csv y series_candidatos_sexto.csv (; y UTF-8 con BOM).
Uso (desde la raíz, sin -I): python3 scripts/revision_indicadores/pdm_series.py
Los archivos descargados se tratan solo como datos.
"""
from __future__ import annotations
import csv, datetime as dt, hashlib, io, ssl, sys, urllib.parse, urllib.request, zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
RAW = RAIZ / "data_raw_IIC" / "07_revision_nov2026"
OUT = RAIZ / "outputs" / "revision_indicadores" / "PDM"
POB = RAIZ / "outputs/sistema_simplificado/trabajo/poblacion_envigado_edad_pped.csv"
SIMPL = RAIZ / "outputs/sistema_simplificado/indicadores_simplificados.csv"
UA = "Mozilla/5.0 (observatorio-indicadores; investigacion academica)"
SOC = "https://www.datos.gov.co/resource/"
URL_TD_ZIP = "https://terridata.dnp.gov.co/assets/docs/txt/entidades/TerriData05266.txt.zip"
URL_TD = "https://terridata.dnp.gov.co/"
ENV = "(cod_muni in('05266','5266') or upper(municipio) like '%ENVIGADO%')"
CAMPOS_M = ["archivo", "url", "fecha_consulta", "bytes", "sha256"]
HOY = dt.date.today().isoformat()


def bajar(carpeta: str, nombre: str, url: str) -> bytes:
    """Descarga y registra en el manifiesto. No sobrescribe cortes anteriores: si ya existe un archivo con el mismo
    sha256 no lo vuelve a guardar; si el contenido cambió lo guarda en una subcarpeta con la fecha."""
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=180, context=ssl.create_default_context()) as r:
        datos = r.read(250 * 1024 * 1024)
    sha = hashlib.sha256(datos).hexdigest()
    d = RAW / carpeta
    d.mkdir(parents=True, exist_ok=True)
    man = d / "manifiesto.csv"
    filas = []
    if man.exists():
        with open(man, encoding="utf-8-sig", newline="") as f:
            filas = list(csv.DictReader(f, delimiter=";"))
    if any(r["url"] == url and r["sha256"] == sha for r in filas):
        return datos  # mismo corte ya guardado
    rel = nombre if not (d / nombre).exists() and not any(r["archivo"] == nombre for r in filas) else f"{HOY}/{nombre}"
    if (d / rel).exists():
        rel = f"{HOY}/{sha[:8]}_{nombre}"
    (d / rel).parent.mkdir(parents=True, exist_ok=True)
    (d / rel).write_bytes(datos)
    filas.append({"archivo": rel, "url": url, "fecha_consulta": HOY, "bytes": len(datos), "sha256": sha})
    with open(man, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_M, delimiter=";")
        w.writeheader(); w.writerows(filas)
    return datos


def soql(rec: str, **kw) -> str:
    return SOC + rec + ".csv?" + urllib.parse.urlencode({"$" + k: v for k, v in kw.items()}, quote_via=urllib.parse.quote)


def csv_dicts(b: bytes) -> list[dict]:
    return list(csv.DictReader(io.StringIO(b.decode("utf-8-sig"))))


# ----------------------------------------------------------------------------- fuentes
def terridata() -> dict:
    """Archivo COMPLETO de TerriData. Para cada (indicador, año) usa el corte de diciembre (mes 12); si ese año no
    tiene mes 12, el mayor mes disponible. Devuelve {indicador: {año: (valor, fuente, mes)}}."""
    z = bajar("terridata", "TerriData05266.txt.zip", URL_TD_ZIP)
    t = zipfile.ZipFile(io.BytesIO(z)).read("TerriData05266.txt").decode("utf-8")
    d = {}
    for l in t.splitlines()[1:]:
        p = l.split("|")
        if p[2] == "05266" and p[7]:
            a, m = int(p[9]), int(p[10] or 0)
            previo = d.setdefault(p[6], {}).get(a)
            if previo is None or m > previo[2]:
                d[p[6]][a] = (float(p[7].replace(".", "").replace(",", ".")), p[11].strip(), m)
    return d


def poblacion() -> tuple[dict, dict]:
    tot, edad = {}, {}
    with open(POB, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            a = int(r["anio"])
            if r["edad"] == "total":
                tot[a] = int(r["poblacion"])
            else:
                edad.setdefault(a, {})[int(r["edad"])] = int(r["poblacion"])
    return tot, edad


def policia_anual(carpeta: str, rec: str, extra_group: str = "") -> dict:
    g = "y" + (", " + extra_group if extra_group else "")
    url = soql(rec, select="date_extract_y(fecha_hecho) as y" + (", " + extra_group if extra_group else "") + ", sum(cantidad::number) as n",
               where=ENV, group=g, order="y", limit="1000")
    filas = csv_dicts(bajar(carpeta, f"envigado_por_anio_{rec}.csv", url))
    res = {}
    for r in filas:
        res[int(r["y"])] = res.get(int(r["y"]), 0) + int(float(r["n"]))
    return res, filas


def transito() -> dict:
    url = soql("ha6j-pa2r", where="codigo_dane='05266000'", limit="5000")
    filas = csv_dicts(bajar("policia_transito", "envigado_filas_ha6j-pa2r.csv", url))
    res = {}
    for r in filas:
        a = int(r["fecha_hecho"][-4:])
        res[a] = res.get(a, 0) + int(float(r["cantidad"]))
    return res, filas


def men() -> dict:
    url = soql("nudc-7mev", select="a_o,c_digo_municipio,municipio,poblaci_n_5_16,cobertura_neta", where="upper(municipio)='ENVIGADO' and c_digo_departamento in('5','05') and c_digo_municipio in('5266','05266')",
               order="a_o", limit="100")
    return {int(r["a_o"]): (float(r["cobertura_neta"]), r["c_digo_municipio"]) for r in csv_dicts(bajar("men_nudc7mev", "envigado_nudc-7mev.csv", url))}


# ----------------------------------------------------------------------------- salida
COLS = ["indicador", "variable", "anio", "valor", "unidad", "numerador", "denominador", "fuente", "enlace", "estado", "nota"]
FILAS: list[dict] = []


def add(ind, var, anio, valor, unidad, fuente, enlace, nota="", numerador="", denominador=""):
    def f(x):
        return "" if x is None or x == "" else (round(x, 4) if isinstance(x, float) else x)
    FILAS.append(dict(zip(COLS, [ind, var, anio, f(valor), unidad, f(numerador), f(denominador), fuente, enlace, "Provisional", nota])))


def main() -> int:
    td = terridata()
    pob, pedad = poblacion()
    PP = "DANE, proyección de población PPED 2018-2042 (total municipal; una sola proyección para todas las tasas)"

    # ============================ 1. SALUD
    I = "Cobertura de salud (aseguramiento total)"
    tot = {a: v[0] for a, v in td["Afiliados al SGSSS"].items() if 2018 <= a <= 2025}
    exist = {}
    with open(SIMPL, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r["indicador"] == "Cobertura de salud" and r["tipo"] == "principal" and r["valor"]:
                exist[int(r["anio"])] = float(r["valor"])
    cs = {}
    for a in sorted(tot):
        cs[a] = 100 * tot[a] / pob[a]
        var = "" if a - 1 not in cs else f"Variación anual: {cs[a] - cs[a-1]:+.2f} puntos porcentuales. "
        cmp_ = f"Serie existente en indicadores_simplificados.csv: {exist[a]:.2f}; diferencia {cs[a]-exist[a]:+.2f} pp. " if a in exist else "Año nuevo respecto de la serie existente (archivo completo de TerriData). "
        add(I, "Cobertura total del aseguramiento (afiliados SGSSS / población)", a, cs[a], "Porcentaje", "TerriData (MinSalud, corte diciembre); " + PP, URL_TD,
            var + cmp_, tot[a], pob[a])
        if a - 1 in cs:
            add(I, "Variación anual de la cobertura (puntos porcentuales)", a, cs[a] - cs[a - 1], "Puntos porcentuales", "Cálculo propio", "", "")
    regs = ("Afiliados al régimen contributivo", "Afiliados al régimen subsidado", "Afiliados a regímenes especiales")
    for nom in regs:
        for a, (v, fu, m) in sorted(td[nom].items()):
            if 2018 <= a <= 2025:
                add(I, nom, a, v, "Personas", "TerriData (MinSalud)", URL_TD, "")
    for a in sorted(tot):
        if all(a in td[n] for n in regs):
            suma = sum(td[n][a][0] for n in regs)
            add(I, "Control: contributivo + subsidiado + especiales menos total SGSSS", a, suma - tot[a], "Personas", "Cálculo propio", "",
                "Debe ser 0 si el total es la suma de los tres regímenes.")

    # ============================ 2. ILE
    I = "ILE (cobertura neta + primera infancia)"
    cn_td = {a: v[0] for a, v in td["Cobertura neta en educación - Total"].items()}
    cn_men = men()
    for a in sorted(cn_td):
        if a >= 2015:
            add(I, "Cobertura neta total en educación (TerriData)", a, cn_td[a], "Porcentaje", "TerriData (MinEducación)", URL_TD,
                ("2015-2017: TerriData (95,9; 96,1; 94,8) y el MEN (68,0; 66,5; 64,2) difieren mucho y no se sabe con qué población se calcularon; el ILE usa solo 2018 en adelante. El MEN en datos.gov.co se muestra como comparativo." if a <= 2017 else ""))
    for a, (v, cod) in sorted(cn_men.items()):
        if a >= 2015:
            dif = f"; diferencia con TerriData {v - cn_td[a]:+.2f}" if a in cn_td else ""
            add(I, "Cobertura neta total (MEN, datos.gov.co nudc-7mev), comparativo", a, v, "Porcentaje", "MinEducación (datos.gov.co, nudc-7mev)",
                "https://www.datos.gov.co/resource/nudc-7mev.json",
                f"Código de municipio en el conjunto: {cod}{dif}." + (" Antes de 2018 la serie del MEN es discontinua (64,24 en 2017 frente a 96,96 en 2018), probablemente por el cambio de proyección." if a == 2017 else ""))
    atend = {a: v[0] for a, v in td["Número de niños, niñas y mujeres gestantes con educación inicial en el marco de la atención integral a la primera infancia"].items()}
    # 2024-2025: informe de gestión (ya en el sistema simplificado)
    ig = {}
    with open(SIMPL, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r["indicador"] == "ILE primera infancia - cobertura de atención integral" and int(r["anio"]) >= 2024:
                ig[int(r["anio"])] = (int(r["numerador"]), r["enlace"])
    nota_pi = "El numerador incluye mujeres gestantes (el denominador solo niños). "

    def nota_fuente(a):
        if a < 2024:
            return ""
        return ("Cambio de fuente en 2024: hasta 2023 el numerador es de TerriData (MinEducación) y desde 2024 es del informe de gestión de la Alcaldía; "
                "no se verificó que cuente igual (niños + gestantes). No se ve un salto: +1,02 pp en 2023 y +0,97 pp en 2024.")
    pi5, pi6 = {}, {}
    for a in range(2018, 2026):
        n_ = atend.get(a) if a <= 2023 else ig[a][0]
        fu = "TerriData (MinEducación)" if a <= 2023 else "Informe de gestión de la Alcaldía (ya recogido en indicadores_simplificados.csv); la fuente cambia en 2024"
        en = URL_TD if a <= 2023 else ig[a][1]
        p5 = sum(pedad[a][e] for e in range(0, 6)); p6 = sum(pedad[a][e] for e in range(0, 7))
        pi5[a], pi6[a] = 100 * n_ / p5, 100 * n_ / p6
        add(I, "Atendidos con educación inicial en atención integral (niños, niñas y gestantes)", a, float(n_), "Personas", fu, en, nota_pi + nota_fuente(a))
        add(I, "Cobertura de atención integral a primera infancia, 0-5 años", a, pi5[a], "Porcentaje", fu + "; " + PP, en,
            nota_pi + ("Año base del PDM: 0-5 da 16,56 y 0-6 da 14,02; LB del PDM 16,3 (dice 0-6 años): 0-5 reproduce mejor la LB. " if a == 2023 else "") + nota_fuente(a), float(n_), p5)
        add(I, "Cobertura de atención integral a primera infancia, 0-6 años (comparativo)", a, pi6[a], "Porcentaje", fu + "; " + PP, en, "", float(n_), p6)
    ile = {}
    for a in range(2018, 2025):
        LA = min(100.0, 100 * pi5[a] / 19)
        LN = min(100.0, 100 * cn_td[a] / 100.44)
        ile[a] = (LA + LN) / 2
        add(I, "LA = min(100, 100·A/19), A = primera infancia 0-5", a, LA, "Puntos (0-100)", "Cálculo propio (matriz PDM.ILE; meta MA = 19, Ap. 9.1.2)", "", "", pi5[a], 19)
        add(I, "LN = min(100, 100·N/100,44), N = cobertura neta TerriData", a, LN, "Puntos (0-100)", "Cálculo propio (MN = 100,44, meta del PDM, Ap. 9.1.2)", "", "", cn_td[a], 100.44)
        add(I, "ILE = (LA + LN)/2", a, ile[a], "Puntos (0-100)", "Cálculo propio", "", "Provisional; sin ILE 2025: la cobertura neta 2025 no está publicada.")
    a = 2025
    LA = min(100.0, 100 * pi5[a] / 19)
    add(I, "LA = min(100, 100·A/19), A = primera infancia 0-5", a, LA, "Puntos (0-100)", "Cálculo propio", "", "LN y ILE 2025: N/D (cobertura neta 2025 sin publicar).", pi5[a], 19)

    # ============================ 3. SERVICIOS PÚBLICOS (archivo completo de TerriData)
    I = "Servicios públicos domiciliarios (cobertura total)"
    S = "Superintendencia de Servicios Públicos (REC), vía TerriData"
    U = "UPME / IPSE, vía TerriData"
    dup = "Duplicados de mes 9 y mes 12 en el mismo año: se usa el corte de diciembre. "
    anom_rec = ("Anomalías del REC (registro de suscriptores y viviendas): caídas a cerca de 68 en 2017 (solo corte de septiembre) y en 2020; "
                "cobertura rural en 0 de 2016 a 2020 con urbana en 100, lo que no cuadra con un total de 68 en 2020; valores de 2016 y 2018 idénticos; "
                "no es la ECV de la meta del PDM (C7). ")
    anom_en = ("Energía: la caída de 2023 viene de las dos zonas (urbana 100 → 98,50 → 97,39 y rural 100 → 64,43 → 69,77), no solo de la rural. ")
    SERV = (("acueducto", "Cobertura de acueducto (REC)", S, anom_rec + dup, True),
            ("alcantarillado", "Cobertura de alcantarillado (REC)", S, anom_rec + dup, True),
            ("aseo", "Cobertura de aseo (REC)", S, anom_rec + dup + "Sin meta en el PDM (no tiene indicador de resultado). 2022 repite el valor de 2021 (87,67). ", False),
            ("energía", "Cobertura de energia eléctrica total", U, anom_en + "UPME/IPSE publica solo desde 2019 en TerriData. ", True))
    tot_sp = {}
    for serv, clave, fuente, nota, meta in SERV:
        tot_sp[serv] = {a: v[0] for a, v in td[clave].items() if 2018 <= a <= 2024}
        for a in sorted(tot_sp[serv]):
            add(I, f"Cobertura de {serv} (total)", a, tot_sp[serv][a], "Porcentaje", fuente, URL_TD, nota + f"Indicador de TerriData: «{clave}».")
            if meta:
                add(I, f"Logro de {serv} = min(100, 100·O/99,9)", a, min(100.0, 100 * tot_sp[serv][a] / 99.9), "Puntos (0-100)",
                    "Cálculo propio (meta PDM 99,9, Ap. 9.1.2)", "", "Meta del PDM de otra fuente (ECV): ver C7.", tot_sp[serv][a], 99.9)
    for a in range(2018, 2025):
        if all(a in tot_sp[s_] for s_ in tot_sp):
            add(I, "Promedio simple de las 4 coberturas (acueducto, alcantarillado, energía, aseo)", a, sum(tot_sp[s_][a] for s_ in tot_sp) / 4, "Porcentaje",
                "Cálculo propio", "", "Con las cuatro series totales; 2018 sin energía (UPME desde 2019) queda fuera.")
    # desagregación urbana / rural (contexto de las anomalías)
    for serv, base in (("acueducto", "Cobertura de acueducto"), ("alcantarillado", "Cobertura de alcantarillado"), ("aseo", "Cobertura de aseo")):
        for z in ("urbana", "rural"):
            clave = f"{base} {z} (REC)"
            for a, (v, _, m) in sorted(td[clave].items()):
                if 2016 <= a <= 2024:
                    add(I, f"{clave} (desagregación, no es el total)", a, v, "Porcentaje", S, URL_TD, "Solo zona " + z + ". " + dup)
    for z in ("urbana", "eléctrica rural"):
        clave = "Cobertura de energia eléctrica urbana" if z == "urbana" else "Cobertura de energía eléctrica rural"
        for a, (v, _, m) in sorted(td[clave].items()):
            if a >= 2018:
                add(I, clave + " (desagregación, no es el total)", a, v, "Porcentaje", U, URL_TD, "Solo zona " + ("urbana" if z == "urbana" else "rural") + ".")
    for serv, clave in (("acueducto", "Cobertura de acueducto (Censo)"), ("alcantarillado", "Cobertura de alcantarillado (Censo)"), ("aseo", "Cobertura de aseo (Censo)")):
        for a, (v, _, m) in td[clave].items():
            add(I, f"Cobertura de {serv} (total): censo", a, v, "Porcentaje", "DANE, Censo 2018, vía TerriData", URL_TD, "Referencia censal 2018; no comparable con el REC.")
    for a, (v, _, m) in sorted(td["Acceso a servicios públicos"].items()):
        if 2016 <= a <= 2024:
            add(I, "Respaldo: puntaje MDM «Acceso a servicios públicos»", a, v, "Puntos (0-100)", "DNP, Medición de Desempeño Municipal (SPT), vía TerriData", URL_TD,
                "Serie de respaldo, no se mezcla con las coberturas." + (" Solo hay corte de octubre (mes 10)." if m == 10 else ""))

    # ============================ 4. SEXTO INDICADOR (candidatos)
    hom, _ = policia_anual("policia_homicidio", "m8fd-ahd9", "spoa_caracterizacion")
    hur, _ = policia_anual("policia_hurto_personas", "4rxi-8m8d")
    les, _ = policia_anual("policia_lesiones", "jr6v-i33g")
    ref = {
        "homicidio": ("Homicidio (intencional o con dolo, incluye feminicidio)", hom, "m8fd-ahd9", 3.7, 2.7, "Tasa de homicidio", "Decreciente."),
        "hurto a personas": ("Hurto a personas", hur, "4rxi-8m8d", 632.5, 631.5, "Tasa de hurto a personas", "Decreciente."),
        "lesiones personales": ("Lesiones personales", les, "jr6v-i33g", 113.7, 112.7, "Tasa de lesiones personales", "Decreciente."),
    }
    tdtasa = {"homicidio": td["Tasa de homicidio intencional por cada 100.000 habitantes"], "hurto a personas": td["Tasa de hurto a personas por cada 100.000 habitantes"]}
    rows_c = []
    for k, (nom, ser, rec, lb, meta, _, _) in ref.items():
        I = f"Sexto indicador, candidato: {k}"
        url = f"https://www.datos.gov.co/resource/{rec}.json"
        tasa = {a: 100000 * ser.get(a, 0) / pob[a] for a in range(2019, 2026)}
        for a in range(2019, 2026):
            add(I, f"Conteo anual de {k}", a, float(ser.get(a, 0)), "Casos", f"Policía Nacional / MinDefensa (datos.gov.co {rec})", url,
                "Hechos ocurridos en Envigado (cod_muni 05266).")
            nota = ""
            if a == 2023:
                nota = f"LB del PDM (2023) = {lb}; con esta población da {tasa[a]:.2f}. Meta 2027 = {meta}."
            add(I, f"Tasa de {k} por 100.000 hab.", a, tasa[a], "Por 100.000 hab.", f"Policía Nacional (datos.gov.co {rec}); " + PP, url, nota, float(ser.get(a, 0)), pob[a])
            if k == "homicidio" and a >= 2021:
                pm = sum(tasa[b] for b in (a - 2, a - 1, a)) / 3
                add(I, "Promedio móvil de 3 años de la tasa de homicidio", a, pm, "Por 100.000 hab.", "Cálculo propio", "", "Promedio de la tasa de los años a-2, a-1 y a.")
            rows_c.append({"candidato": k, "anio": a, "conteo": ser.get(a, 0), "poblacion": pob[a], "tasa_100mil": round(tasa[a], 2),
                           "tasa_TerriData": (round(tdtasa[k][a][0], 2) if k in tdtasa and a in tdtasa[k] else ""), "linea_base_2023_PDM": lb, "meta_2027_PDM": meta,
                           "prom_movil_3a": (round(sum(100000 * ser.get(b, 0) / pob[b] for b in (a - 2, a - 1, a)) / 3, 2) if k == "homicidio" and a >= 2021 else ""),
                           "fuente": f"datos.gov.co {rec}"})

    # ============================ 5. MORTALIDAD VIAL
    I = "Mortalidad vial (Policía, homicidio culposo en accidente de tránsito)"
    tr, _ = transito()
    td_vial = td["Tasa de homicidios en accidentes de tránsito por cada 100.000 habitantes"]
    for a in range(2019, 2026):
        t = 100000 * tr.get(a, 0) / pob[a]
        nota = f"TerriData (misma fuente Policía) publica {td_vial[a][0]:.1f}. " if a in td_vial else ""
        if a == 2023:
            nota += ("Salto: 43 muertes (29 en 2022 y 20 en 2024). El conjunto registra hechos ocurridos en Envigado, no residentes; las 43 filas están repartidas durante todo el año "
                     "(no es un solo evento: máximo 3 víctimas en un día (8-jun y 5-jul)). LB del PDM 10,2 (2023, fuente Secretaría de Movilidad, otro universo): no comparable (C7).")
        if a == 2025:
            nota += "Meta del PDM 8 (Movilidad, otro universo). Datos de la Policía hasta jul-2026."
        add(I, "Muertes en tránsito (hechos en Envigado)", a, float(tr.get(a, 0)), "Muertes", "Policía Nacional (datos.gov.co ha6j-pa2r)", "https://www.datos.gov.co/resource/ha6j-pa2r.json", "", "", "")
        add(I, "Tasa de mortalidad vial por 100.000 hab.", a, t, "Por 100.000 hab.", "Policía Nacional (ha6j-pa2r); " + PP, "https://www.datos.gov.co/resource/ha6j-pa2r.json", nota, float(tr.get(a, 0)), pob[a])

    # ============================ escribir
    OUT.mkdir(parents=True, exist_ok=True)
    with open(OUT / "series_pdm.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter=";"); w.writeheader(); w.writerows(FILAS)
    with open(OUT / "series_candidatos_sexto.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows_c[0]), delimiter=";"); w.writeheader(); w.writerows(rows_c)
    print(len(FILAS), "filas en series_pdm.csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
