"""Lectores de PDF para PP: listado EFPC (valor votado) y ejecucion presupuestal de egresos de Hacienda.

Los PDF se leen como texto (pymupdf); no se ejecuta nada de ellos. Todo resultado es PROVISIONAL.
"""
import re
from collections import defaultdict
from pathlib import Path

import pymupdf

RAIZ = Path(__file__).resolve().parents[2]
BASE = RAIZ / "data_raw_IIC/07_revision_nov2026"
MONEY_W = re.compile(r"^\$?\d{1,3}(?:[.,]\d{3}){2,}(?:[.,]\d{1,2})?$")


def valor_colombiano(s):
    """'$ 178,140,439' -> 178140439 ; '159.630.940,00' -> 159630940 ; '203.335.424,86' -> 203335424.86"""
    s = s.replace("$", "").strip()
    m = re.search(r"[.,](\d{2})$", s)
    if m:
        ent = re.sub(r"[.,]", "", s[: m.start()])
        return float(f"{ent}.{m.group(1)}") if m.group(1) != "00" else float(ent)
    return float(re.sub(r"[.,]", "", s))


def leer_efpc(anio):
    """Lista de dicts {zona, nombre, valor, pagina} del listado EFPC (valor total del proyecto).
    Filas ancladas por la columna de valor (importe). La zona es el digito de la columna izquierda a la misma altura;
    si la celda de zona esta combinada (sin digito en esa fila) hereda la zona de la fila anterior. El nombre sale de la
    columna del proyecto y solo sirve para confirmar."""
    doc = pymupdf.open(BASE / f"efpc_listados_pp/Proyectos_PP_{anio}.pdf")
    regs, zona_prev = [], None
    for pn, pg in enumerate(doc, 1):
        words = [w for w in pg.get_text("words") if w[1] > 62]
        cand = [w for w in words if w[0] < 97 and re.fullmatch(r"\d{1,2}", w[4])]
        x_min = min((w[0] for w in cand), default=0)
        zonas = [w for w in cand if w[0] <= x_min + 6]
        anclas = sorted(((w[1], w[4]) for w in words if MONEY_W.match(w[4])))
        for i, (y, t) in enumerate(anclas):
            ylo = (anclas[i - 1][0] + y) / 2 if i else 62
            yhi = (y + anclas[i + 1][0]) / 2 if i + 1 < len(anclas) else 9999
            z = [w for w in zonas if ylo <= w[1] < yhi]
            if len(z) > 1:
                raise ValueError(f"EFPC {anio} p.{pn}: dos zonas en una fila")
            zona = int(z[0][4]) if z else zona_prev
            zona_prev = zona
            nm = [w for w in words if 97 <= w[0] < 243 and ylo <= w[1] < yhi and not (pn == 1 and w[1] < 82)]
            nm.sort(key=lambda w: (round(w[1] / 3), w[0]))
            regs.append({"zona": zona, "nombre": " ".join(w[4] for w in nm), "valor": valor_colombiano(t),
                         "pagina": pn, "zona_heredada": not z})
    return regs


RE_COD = re.compile(r"\d{18,}")
RE_NUM = re.compile(r"^-$|^\$?-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}$|^\$?\s*0[.,]00$")


def leer_hacienda(anio_o_ruta):
    """Filas hoja de rubro de un PDF de ejecucion de egresos, leidas por posicion (palabras con coordenadas).
    Cada rubro (codigo solo de digitos, >= 18) agrupa los importes y la descripcion que quedan en su franja vertical.
    Devuelve lista de dict {rubro, sufijo, pagina, descripcion, ap_ini, ap_def, comp, obl, ok}; 'ok' = 11 importes.
    Importes en orden: inicial, adiciones, disminuciones, traslado cred., contracred., definitiva, certificado,
    registro (compromiso), ejecucion (obligacion), pagos, disponible. El '-' vale 0."""
    ruta = anio_o_ruta if isinstance(anio_o_ruta, Path) else BASE / f"hacienda_ejecucion/eje{anio_o_ruta}q4.pdf"
    out = []
    for pn, pg in enumerate(pymupdf.open(ruta), 1):
        W = pg.get_text("words")
        # 2024 trae los importes con "$" en palabra aparte y en columna; se trata aparte
        cods = sorted((w for w in W if RE_COD.fullmatch(w[4]) and w[0] < 100), key=lambda w: w[1])
        if not cods:
            continue
        ys = [w[1] for w in cods]
        for i, c in enumerate(cods):
            ylo = (ys[i - 1] + ys[i]) / 2 if i else 0
            yhi = (ys[i] + ys[i + 1]) / 2 if i + 1 < len(ys) else 9999
            blk = [w for w in W if ylo <= w[1] < yhi and w[0] > 100]
            nums = sorted((w for w in blk if w[0] > 260 and abs(w[1] - c[1]) <= 3 and RE_NUM.match(w[4])), key=lambda w: w[0])
            vals = [0.0 if w[4] == "-" else valor_colombiano(w[4]) for w in nums[-11:]]
            desc = " ".join(w[4] for w in sorted((w for w in blk if 140 < w[0] < 330 and not RE_NUM.match(w[4])),
                                                 key=lambda w: (round(w[1] / 3), w[0])))
            row = {"rubro": c[4], "sufijo": c[4][-4:], "pagina": pn, "descripcion": desc, "ok": len(vals) == 11}
            if len(vals) == 11:
                row.update(ap_ini=vals[0], ap_def=vals[5], comp=vals[7], obl=vals[8])
            out.append(row)
    return out


def leer_roca_2025():
    """ROCA 2025 (hoja ROCA07.01.26 del Plan de accion IV trim. 2025): misma estructura que la ejecucion de egresos de Hacienda
    (rubro de 47 digitos, columna PROYECTO, apropiacion inicial/definitiva, compromiso y obligacion acumulados)."""
    import openpyxl
    f = BASE / "plan_accion_2025/evaluaciondefinitivaplanaccionivtri.xlsx"
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)["ROCA07.01.26"]
    out, ix, hdr = [], {}, None
    for n, r in enumerate(ws.iter_rows(values_only=True), 1):
        if r[0] == "SECCIÓN":
            hdr = [str(h or "").replace("\n", " ").strip() for h in r]
            ix = {k: hdr.index(v) for k, v in {"proy": "PROYECTO", "rubro": "RUBRO PRESUPUESTAL", "desc": "DESCRIPCIÓN DEL RUBRO PRESUPUESTAL",
                  "ini": "APROPIACIÓN INICIAL", "dfn": "APROPIACIÓN DEFINITIVA", "comp": "REGISTRO COMPROMISO ACUMULADO",
                  "obl": "EJECUCIÓN OBLIGACIÓN ACUMULADO", "bpin": "BANCO PROYECTO DE INVERSIÓN-BPIN"}.items()}
            continue
        if not hdr or r[ix["rubro"]] in (None, ""):
            continue
        cod = str(r[ix["rubro"]])
        if not RE_COD.fullmatch(cod):
            continue
        f_ = lambda k: float(str(r[ix[k]] or 0).replace(",", "")) if r[ix[k]] not in (None, "") else 0.0
        out.append({"rubro": cod, "sufijo": str(r[ix["proy"]]).zfill(4), "pagina": f"ROCA07.01.26 fila {n}",
                    "descripcion": str(r[ix["desc"]] or ""), "ok": True, "ap_ini": f_("ini"), "ap_def": f_("dfn"),
                    "comp": f_("comp"), "obl": f_("obl")})
    return out
