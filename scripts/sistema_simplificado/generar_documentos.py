#!/usr/bin/env python3
"""Genera los documentos del sistema simplificado (PROVISIONAL), con las cifras leídas de
outputs/sistema_simplificado/indicadores_simplificados.csv y variables_por_dimension.csv:

  outputs/sistema_simplificado/sistema_indicadores_simplificado.{pdf,docx,xlsx}   (PDF y Word: máximo 6 páginas)
  outputs/sistema_simplificado/variables_para_solicitar.{docx,xlsx}               (máximo 2 páginas)

Uso (desde la raíz, sin -I):  python3 scripts/sistema_simplificado/generar_documentos.py [--sin-pdf]
Cadena: Markdown -> .tex (estilo de los informes anteriores) -> PDF con tectonic -> Word con el traductor .tex -> .docx
de scripts/generar_informe_docx.py. Se importan (sin modificarlos) generar_informe_latex, generar_informe_docx y
scripts/entregables/generar_guia; sus funciones se reutilizan y se ajustan aquí en memoria (anchos de tabla, enlaces
con texto corto, colores de prioridad).
"""
from __future__ import annotations

import csv
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts"))
sys.path.insert(0, str(RAIZ / "scripts" / "entregables"))
import generar_informe_latex as L  # noqa: E402
import generar_informe_docx as D  # noqa: E402
import generar_guia as G  # noqa: E402

OUT = RAIZ / "outputs" / "sistema_simplificado"
F_CSV = OUT / "indicadores_simplificados.csv"
F_VARS = OUT / "variables_por_dimension.csv"
TMP = G.SCRATCH / "simplificado2"
G.TMP = TMP
FECHA = "10 de noviembre de 2026"
MARGEN = 1.6                      # cm
ANCHO = 27.94 - 2 * MARGEN        # cm de texto en horizontal (carta)
G.ANCHO_V = G.ANCHO_H = ANCHO
ESTADO_COLOR = {"Disponible": "cverde", "Construible": "ccata", "No disponible": "crojo"}
PRIORIDAD = {"Alta": "crojo", "Media": "camarillo", "Baja": "cverde"}

# ------------------------------------------------------------------------------------------------ datos
with open(F_CSV, encoding="utf-8-sig", newline="") as f:
    FILAS = list(csv.DictReader(f, delimiter=";"))
for r in FILAS:
    r["anio"] = int(r["anio"])
    r["v"] = float(r["valor"]) if r["valor"] != "" else None


with open(F_VARS, encoding="utf-8-sig", newline="") as f:
    VARS = list(csv.DictReader(f, delimiter=";"))
for r in VARS:
    r["anio"] = int(r["anio"]) if r["anio"] != "" else ""
    r["v"] = float(r["valor"]) if r["valor"] != "" else None
    r["detalle"] = r["detalle_valor"]


def fm(x, dec=2):
    """Número con coma decimal y punto de miles; enteros sin decimales."""
    if x is None:
        return "sin valor"
    if abs(x - round(x)) < 1e-9:
        return f"{int(round(x)):,}".replace(",", ".")
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fila(ind, anio, tipo=None):
    c = [r for r in FILAS if r["indicador"] == ind and r["anio"] == anio and (tipo is None or r["tipo"] == tipo)]
    assert len(c) == 1, (ind, anio, tipo, len(c))
    return c[0]


def v(ind, anio, tipo=None):
    return fila(ind, anio, tipo)["v"]


def razon(ind, anio, tipo=None):
    r = fila(ind, anio, tipo)
    return f"{fm(float(r['numerador']))}/{fm(float(r['denominador']))}"


def enlace(ind, anio):
    return fila(ind, anio)["enlace"]


def con_valor(nombres):
    """True si alguno de los indicadores (nombres exactos) tiene al menos un año con valor."""
    return any(r["v"] is not None for r in FILAS if r["indicador"] in nombres)


def lk(texto, url):
    return f"[{texto}]({url})" if url else texto


def pares(*items):
    """'a (2024) y b (2025)' a partir de [(valor_texto, anio)], agrupando valores iguales."""
    grupos = []
    for t, a in items:
        if grupos and grupos[-1][0] == t:
            grupos[-1][1].append(str(a))
        else:
            grupos.append([t, [str(a)]])
    return "; ".join(f"{t} ({' y '.join(an)})" for t, an in grupos)


def sem_ultimo(ind, anios):
    for a in sorted(anios, reverse=True):
        r = fila(ind, a)
        if r["v"] is not None and r["semaforo"] in ("Verde", "Amarillo", "Naranja", "Rojo"):
            return r["semaforo"]
    return "Gris"


# ------------------------------------------------------------------------------------------------ ajustes en memoria de generar_guia
_en_linea_orig = G.en_linea
_RE_LINK = re.compile(r"(\[[^\]]+\]\(https?://[^)\s]+\))")


def _destino(url):
    return url.replace("\\", "/").replace("%", r"\%").replace("#", r"\#").replace("&", r"\&").replace("~", r"\~")


def _en_linea(t, tabla=False):
    out = []
    for p in _RE_LINK.split(t):
        m = re.fullmatch(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", p)
        out.append(r"\href{" + _destino(m.group(2)) + "}{" + L.esc(m.group(1)) + "}" if m else _en_linea_orig(p, tabla))
    return "".join(out)


G.en_linea = _en_linea

FRACCIONES = {"vars": [0.19, 0.09, 0.26, 0.16, 0.21, 0.09], "inds": [0.20, 0.085, 0.215, 0.235, 0.07, 0.195],
              "res": [0.19, 0.09, 0.09, 0.09, 0.09, 0.11, 0.12, 0.12, 0.10], "solic": [0.30, 0.22, 0.23, 0.15, 0.10]}


def _anchos(filas, total, tam):
    n = len(filas[0])
    util = total - n * 0.22
    clave = {9: "res", 5: "solic"}.get(n) or ("vars" if filas[0][0] == "Variable" else "inds")
    return [util * f for f in FRACCIONES[clave]]


G.anchos = _anchos
_celda_orig = G.celda_tex


def _celda(texto, enc, par):
    plano = re.sub(r"\*", "", texto.strip())
    if enc.startswith("Semáforo") and plano == "Sin semáforo":
        return r"\cellcolor{cgris}" + G.en_linea(texto.strip(), tabla=True)
    if enc == "Estado" and plano in ESTADO_COLOR:
        return r"\cellcolor{" + ESTADO_COLOR[plano] + "}" + G.en_linea(texto.strip(), tabla=True)
    if enc.startswith("Prioridad") and plano in PRIORIDAD:
        return r"\cellcolor{" + PRIORIDAD[plano] + "}" + G.en_linea(texto.strip(), tabla=True)
    return _celda_orig(texto, enc, par)


G.celda_tex = _celda
_tabla_orig = G.tabla_tex


def _tabla(filas, ancho):
    """Igual que la original pero siempre en letra pequeña (7 pt) para que las tablas por dimensión quepan."""
    n = len(filas[0])
    tam = "footnotesize" if n == 5 else "scriptsize"
    w = G.anchos(filas, ancho, tam)
    spec = "".join(r"L{%.2fcm}" % x for x in w)
    cab = [r"\cellcolor{chead}\textbf{" + G.en_linea(c, True) + "}" for c in filas[0]]
    rayado = len(filas) > 7
    cuerpo = [[G.celda_tex(c, filas[0][k], rayado and i % 2 == 1) for k, c in enumerate(f)] for i, f in enumerate(filas[1:])]
    fin = r" \\"
    lin = ["{" + (chr(92) + "fontsize{6.5}{7.6}" + chr(92) + "selectfont" if n != 5 else chr(92) + tam), r"\begin{xltabular}{\linewidth}{" + spec + "}", r"\toprule", " & ".join(cab) + fin,
           r"\midrule", r"\endfirsthead", r"\toprule", " & ".join(cab) + fin, r"\midrule", r"\endhead",
           r"\midrule \multicolumn{" + str(n) + r"}{r}{\emph{Continúa en la página siguiente}} \\", r"\endfoot",
           r"\bottomrule", r"\endlastfoot"]
    lin += [" & ".join(f) + fin for f in cuerpo]
    lin += [r"\end{xltabular}", "}"]
    return "\n".join(lin), (8 if n == 5 else 6.5)


G.tabla_tex = _tabla


# ------------------------------------------------------------------------------------------------ LaTeX / Word
def preambulo(titulo_corto, titulo_pdf):
    p = G.PREAMBULO.replace("\\documentclass[11pt,letterpaper]{article}", "\\documentclass[10pt,letterpaper]{article}")
    p = p.replace("\\usepackage[margin=MARGEN,headheight=14pt]{geometry}",
                  f"\\usepackage[landscape,left={MARGEN}cm,right={MARGEN}cm,top=1.6cm,bottom=1.2cm,headheight=12pt,headsep=0.3cm]{{geometry}}")
    p = p.replace("[hidelinks,breaklinks=true]", "[colorlinks=true,linkcolor=black,urlcolor=blue!55!black,breaklinks=true]")
    p = p.replace("COLORES", G.COLORES_TEX)
    p = p.replace("\\newcommand{\\pagsize}{\\special{papersize=21.59cm,27.94cm}}",
                  "\\newcommand{\\pagsize}{\\special{papersize=27.94cm,21.59cm}}")
    p = p.replace("\\setlength{\\parskip}{0.5em}", "\\setlength{\\parskip}{0.35em}")
    p = p.replace("\\renewcommand{\\arraystretch}{1.18}", "\\renewcommand{\\arraystretch}{1.04}")
    p += (r"\pagestyle{fancy}" "\n" r"\fancyhf{}" "\n" r"\fancyhead[L]{\pagsize\small " + L.esc(titulo_corto) + "}\n"
          r"\fancyhead[R]{\small\thepage}" "\n" r"\renewcommand{\headrulewidth}{0.4pt}" "\n"
          r"\fancypagestyle{plain}{\fancyhf{}\fancyhead[R]{\small\thepage}}" "\n"
          r"\hypersetup{pdftitle={" + L.esc(titulo_pdf) + "}}\n")
    return p


def documento_tex(titulo_corto, titulo, cuerpo_tex, word):
    if word:
        cab = r"\section*{" + L.esc(titulo) + "}\n" + L.esc(FECHA + ". Observatorio Ciudadano de Envigado (OCE).") + "\n\n"
    else:
        cab = (r"\begin{center}{\Large\bfseries " + L.esc(titulo) + r"\par}\vspace{3pt}{\small " +
               L.esc(FECHA + ". Observatorio Ciudadano de Envigado (OCE).") + r"\par}\end{center}" "\n\n")
    aviso = L.aviso(L.esc("PROVISIONAL: propuesta de trabajo. Ninguna cifra está validada ni confirmada por la Alcaldía.")) + "\n\n"
    return (preambulo(titulo_corto, titulo) + "\n\\begin{document}\n" + ("" if word else "\\small\n") + "\\hypersetup{pageanchor=true}\n\n" + cab + aviso +
            cuerpo_tex + "\n\n\\end{document}\n")


def a_tex(md, word):
    bl = G.parsear_md(md)
    T, tam, _ = G.a_latex(bl, word=word)
    tex = "\n\n".join(T).replace(r"\needspace{4\baselineskip}", r"\needspace{6\baselineskip}")
    if word:   # el traductor a Word no lee \fontsize: el tamaño real (tam) se aplica al escribir la tabla
        tex = tex.replace(r"{\fontsize{6.5}{7.6}\selectfont", r"{\scriptsize")
    return tex, tam


def construir_word(tex_word, tamanos, destino, encabezado):
    bls = D.bloques(tex_word)
    colores = {m.group(1): "%02X%02X%02X" % tuple(int(x) for x in m.group(2).split(","))
               for m in re.finditer(r"\\definecolor\{(\w+)\}\{RGB\}\{([\d,]+)\}", tex_word)}
    doc = Document()
    D.preparar_estilos(doc)
    D.fijar_fuente(doc.styles["Normal"], D.FUENTE, 9)
    doc.styles["Normal"].paragraph_format.space_after = Pt(4)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.05
    D.ENCABEZADO = encabezado
    D.ANCHO_TEXTO = ANCHO
    D.encabezado_y_paginas(doc)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(27.94), Cm(21.59)
    for lado in ("left_margin", "right_margin"):
        setattr(sec, lado, Cm(MARGEN))
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.2)
    sec.different_first_page_header_footer = False
    pg = sec._sectPr.find(qn("w:pgNumType"))
    if pg is not None:
        sec._sectPr.remove(pg)
    G.encabezado_seccion(doc, sec, ANCHO)
    D.actualizar_campos_al_abrir(doc)
    cuerpo = doc.element.body
    for p in list(cuerpo.iterchildren(qn("w:p"))):
        cuerpo.remove(p)
    tablas = iter(tamanos)
    primero = True
    for b in bls:
        t = b[0]
        if t == "titulo":
            _, nivel, _, texto, _ = b
            if primero:
                p = doc.add_paragraph(style="Title")
                p.add_run(texto).font.size = Pt(16)
                primero = False
            else:
                estilo = {"section": "Heading 1", "subsection": "Heading 2", "subsubsection": "Heading 3", "paragraph": "Heading 4"}[nivel]
                p = doc.add_paragraph(style=estilo)
                p.add_run(texto)
                if estilo == "Heading 2":
                    p.runs[0].font.size = Pt(12)
                    p.paragraph_format.space_before = Pt(8)
                    p.paragraph_format.space_after = Pt(3)
        elif t == "aviso":
            D.escribir_aviso(doc, b[1])
        elif t == "parrafo":
            runs = D.recortar_espacios(b[1])
            if not D.texto_plano(runs):
                continue
            p = doc.add_paragraph()
            D.agregar_runs(p, runs, colores)
        elif t == "lista":
            numid = D.numeracion_reiniciada(doc) if b[1] else None
            for item in b[2]:
                runs = D.recortar_espacios(item)
                p = doc.add_paragraph(style="List Number" if b[1] else "List Bullet")
                if numid:
                    numpr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numpr.get_or_add_ilvl().val = 0
                    numpr.get_or_add_numId().val = numid
                p.paragraph_format.space_after = Pt(2)
                D.agregar_runs(p, runs, colores)
        elif t == "tabla":
            D.escribir_tabla(doc, b[1], b[2], b[3], next(tablas), colores)
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
        elif t in ("salto",):
            continue
        else:
            raise ValueError(f"bloque no previsto: {t}")
    doc.save(destino)


def compilar_pdf(tex_text, nombre, destino_pdf):
    TMP.mkdir(parents=True, exist_ok=True)
    raros = sorted({ch for ch in tex_text if ord(ch) > 127 and ch not in L.LETRAS_OK})
    assert not raros, f"caracteres sin traducir: {raros}"
    f_tex = TMP / f"{nombre}.tex"
    f_tex.write_text(tex_text, encoding="utf-8")
    G.compilar(f_tex, destino_pdf)
    return f_tex


def paginas_pdf(path):
    import pymupdf
    return len(pymupdf.open(path))


def paginas_word(docx, carpeta):
    carpeta.mkdir(parents=True, exist_ok=True)
    env = {"HOME": str(TMP), "PATH": "/usr/bin:/bin"}
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(carpeta), str(docx)],
                   check=True, capture_output=True, env=env, timeout=240)
    return paginas_pdf(carpeta / (Path(docx).stem + ".pdf")), carpeta / (Path(docx).stem + ".pdf")


# ================================================================================================ DOCUMENTO 1
DIMS = ["Control social", "Participación ciudadana", "Cumplimiento del PDM", "Presupuesto participativo"]
DIM_TIT = {"Control social": "Control social (CS)", "Participación ciudadana": "Participación ciudadana (PC)",
           "Cumplimiento del PDM": "Cumplimiento del Plan de Desarrollo Municipal (PDM)", "Presupuesto participativo": "Presupuesto participativo (PP)"}
SIN_SEM = "Sin semáforo"
SEMS = ("Verde", "Amarillo", "Naranja", "Rojo")
TIPO_TXT = {"principal": "Principal", "respaldo": "Respaldo", "complemento": "Complemento"}


def tabla_md(cab, filas):
    lin = ["| " + " | ".join(cab) + " |", "|" + "---|" * len(cab)]
    for f in filas:
        assert len(f) == len(cab), f
        lin.append("| " + " | ".join(x.replace("|", "/") for x in f) + " |")
    return "\n".join(lin)


def mm(x):
    """Pesos en millones, con coma decimal."""
    return "$" + fm(x / 1e6, 1) + " millones"


def fv(x, unidad):
    if unidad == "pesos":
        return mm(x)
    return fm(x)


def hoy(vs):
    """Texto 'qué hay hoy' de una variable (lista de filas del CSV de variables) con valor y años, más la fuente."""
    v0 = vs[0]
    con = [r for r in vs if r["valor"] != ""]
    if not con:
        txt = "Sin dato"
    else:
        con.sort(key=lambda r: r["anio"])
        det = {r["detalle"] for r in con}
        un = v0["unidad"]
        if len(con) > 3:
            txt = (f"Serie {con[0]['anio']}-{con[-1]['anio']}: {fv(con[0]['v'], un)} ({con[0]['anio']}) a {fv(con[-1]['v'], un)} ({con[-1]['anio']})")
        else:
            items = []
            for r in con:
                t = fv(r["v"], un)
                if r["anio"] == 2027:
                    items.append(f"meta 2027: {t}")
                else:
                    items.append(f"{t} ({r['anio']})")
            txt = "; ".join(items)
        if len(det) == 1 and next(iter(det)) and not (len(con) == 1 and con[0]["anio"] == 2027):
            txt += f" [{next(iter(det))}]"
        elif len(con) == 1 and con[0]["anio"] != 2027 and con[0]["detalle"]:
            txt += f" [{con[0]['detalle']}]"
    f = v0["fuente"]
    return txt + ". " + (lk(f, v0["enlace"]) if v0["enlace"] else f)


def vars_dim(dim):
    orden, d = [], {}
    for r in VARS:
        if r["dimension"] == dim:
            if r["id_variable"] not in d:
                orden.append(r["id_variable"])
            d.setdefault(r["id_variable"], []).append(r)
    return [d[k] for k in orden]


def tabla_variables(dim):
    cab = ["Variable", "Para qué indicador", "Qué hay hoy (valor, años y fuente)", "Qué falta", "Limitación", "Estado"]
    filas = []
    for vs in vars_dim(dim):
        v0 = vs[0]
        filas.append([v0["variable"], v0["para_indicador"], hoy(vs), v0["que_falta"], v0["limitacion"], v0["estado"]])
    return tabla_md(cab, filas)


def ser(ind, anios, tipo=None, con_razon=False):
    items = []
    for a in anios:
        r = fila(ind, a, tipo)
        if r["v"] is None:
            t = "N/D"
        else:
            t = fm(r["v"])
            if con_razon and r["numerador"] != "" and r["denominador"] != "":
                t += f" ({razon(ind, a, tipo)})"
        items.append((t, a))
    return pares(*items)


def sem_de(ind, anios, tipo=None):
    return sem_ultimo(ind, anios) if tipo is None else next(
        (fila(ind, a, tipo)["semaforo"] for a in sorted(anios, reverse=True)
         if fila(ind, a, tipo)["v"] is not None and fila(ind, a, tipo)["semaforo"] in SEMS), SIN_SEM)


def _sem(x):
    return SIN_SEM if x == "Gris" or x not in SEMS else x


CUENTA: list[tuple] = []     # (dimensión, tipo, [nombres de indicadores del CSV])


def IND(dim, nombre, tipo, formula, valores, sem, limit, nombres):
    CUENTA.append((dim, tipo, nombres))
    return [nombre, TIPO_TXT[tipo], formula, valores, _sem(sem), limit]


def tabla_indicadores(filas):
    return tabla_md(["Indicador", "Tipo", "Fórmula", "Valor por año", "Semáforo", "Limitación"], filas)


def ind_cs():
    D, A = "Control social", (2024, 2025)
    pcs, pcs_r = "PCS - Promoción del control social", "PCS - Respaldo: estrategias ejecutadas (número)"
    dsi, dsi_r = "DSI - Disponibilidad de información de seguimiento", "DSI - Respaldo: con todo lo publicado hasta la fecha de consulta"
    vrc, vrc_r = "VRC - Veedurías en la rendición de cuentas", "VRC - Respaldo: evento sin mención de veeduría = ausencia"
    rpo, rpo_r = "RPO - Respuesta a observaciones ciudadanas", "RPO - Respaldo: con el cuadro de preguntas y respuestas publicado"
    icr, icr_r = "ICRCS - Índice de Condiciones y Respuesta para el Control Social", "ICRCS - Respaldo: con los respaldos de VRC y RPO"
    return [
        IND(D, "**PCS** Promoción del control social", "principal", "100 × estrategias ejecutadas / programadas", ser(pcs, A), sem_ultimo(pcs, A),
            "Dato autodeclarado por la dependencia", [pcs]),
        IND(D, "PCS: estrategias ejecutadas", "respaldo", "Conteo de estrategias ejecutadas", ser(pcs_r, A), "Gris", "Conteo sin meta: no tiene semáforo", [pcs_r]),
        IND(D, "**DSI** Disponibilidad de información de seguimiento", "principal", "100 × metas con valor publicado a tiempo / metas seleccionadas",
            ser(dsi, A, con_razon=True), sem_ultimo(dsi, A), "Depende de la lista de metas y de que se publiquen valores de resultado", [dsi]),
        IND(D, "DSI: con todo lo publicado hasta la consulta", "respaldo", "Igual, sin exigir que se publique a tiempo", ser(dsi_r, (2024,), con_razon=True), sem_ultimo(dsi_r, (2024,)),
            "Cuenta lo publicado después del cierre de la vigencia", [dsi_r]),
        IND(D, "**VRC** Veedurías en la rendición de cuentas", "principal", "100 × eventos con veeduría / eventos", ser(vrc, A), sem_ultimo(vrc, A),
            "Faltan listas de asistencia y el universo de eventos", [vrc]),
        IND(D, "VRC: evento sin mención = ausencia", "respaldo", "Igual, contando como ausencia el evento que no menciona veedurías", ser(vrc_r, A, con_razon=True), sem_ultimo(vrc_r, A),
            "El supuesto puede subestimar la presencia real", [vrc_r]),
        IND(D, "**RPO** Respuesta a observaciones ciudadanas", "principal", "100 × observaciones con respuesta de fondo / observaciones", ser(rpo, A, con_razon=True), sem_ultimo(rpo, A),
            "2025 sin cuadro completo de preguntas y respuestas", [rpo]),
        IND(D, "RPO: con el cuadro publicado", "respaldo", "Igual, con el único cuadro publicado de 2025", ser(rpo_r, (2025,), con_razon=True), sem_ultimo(rpo_r, (2025,)),
            "Solo 2 observaciones: resultado frágil", [rpo_r]),
        IND(D, "**ICRCS** Índice de condiciones y respuesta para el control social", "principal", "(DSI + PCS + VRC + RPO) / 4", ser(icr, A), sem_ultimo(icr, A),
            "No se calcula si falta algún componente", [icr]),
        IND(D, "ICRCS: con los respaldos de VRC y RPO", "respaldo", "Mismo promedio, con los respaldos de VRC y RPO", ser(icr_r, A), sem_ultimo(icr_r, A),
            "Mezcla principales y respaldos; provisional", [icr_r]),
    ]


def ind_pc():
    D, A = "Participación ciudadana", (2024, 2025)
    i1, i2, i3, i4 = "I1 - Incorporación de aportes", "I2 - Cobertura territorial", "I3 - Oferta formativa", "I4 - Publicación de resultados"
    r1 = "I1 - Respaldo: proyectos votados con presupuesto asignado"
    r2 = "I2 - Respaldo: zonas con asamblea de presupuesto participativo"
    r3 = "I3 - Respaldo: acciones formativas culminadas (número)"
    r4 = "I4 - Respaldo: procesos con documento de resultados (universo exploratorio)"
    io, ir = "ISPC - Índice (fórmula original: 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4)", "ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)"
    return [
        IND(D, "**I1** Incorporación de aportes", "principal", "100 × procesos con aporte incorporado / procesos cerrados", ser(i1, A), sem_ultimo(i1, A),
            "Sin inventario de procesos cerrados", [i1]),
        IND(D, "I1: proyectos votados con presupuesto asignado", "respaldo", "100 × proyectos votados con presupuesto / proyectos votados", ser(r1, A, con_razon=True), sem_ultimo(r1, A),
            "Usa los datos de la fidelidad de PP: no sumar en dos dimensiones", [r1]),
        IND(D, "**I2** Cobertura territorial", "principal", "100 × zonas con actividad / 13 zonas", ser(i2, A, con_razon=True), sem_ultimo(i2, A),
            "Faltan actas por zona", [i2]),
        IND(D, "I2: zonas con asamblea de PP", "respaldo", "100 × zonas con asamblea de presupuesto participativo / 13", ser(r2, A, con_razon=True), sem_ultimo(r2, A),
            "La zona 1 no tiene comité zonal", [r2]),
        IND(D, "**I3** Oferta formativa", "principal", "100 × acciones realizadas / programadas", ser(i3, A, con_razon=True), sem_ultimo(i3, A),
            "2024 sin programación publicada; 2025 reconstruida", [i3]),
        IND(D, "I3: acciones formativas culminadas", "respaldo", "Conteo de acciones culminadas", ser(r3, A), "Gris",
            "Incluye acciones para servidores; sin meta, sin semáforo", [r3]),
        IND(D, "**I4** Publicación de resultados", "principal", "100 × procesos con resultados publicados / procesos cerrados", ser(i4, A), sem_ultimo(i4, A),
            "Sin inventario de procesos cerrados", [i4]),
        IND(D, "I4: procesos del universo exploratorio", "respaldo", "Conteo de procesos del universo exploratorio", ser(r4, A), "Gris",
            "Universo no oficial; sin semáforo", [r4]),
        IND(D, "**ISPC** Índice de seguimiento de la participación ciudadana", "principal", "0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4", ser(io, A), sem_ultimo(io, A),
            "Faltan I1 e I4", [io]),
        IND(D, "ISPC: opción redistribuida", "respaldo", "0,556·I2 + 0,444·I3 (pesos 0,25/0,45 y 0,20/0,45)", ser(ir, A), sem_ultimo(ir, A),
            "No es el índice original; 2024 sin I3", [ir]),
    ]


def ind_pdm():
    D = "Cumplimiento del PDM"
    salud = "Cobertura de salud"
    cm, cmr = "CM - Mortalidad vial", "CM - Respaldo: cifra de la Policía"
    ca, car = "CA - Espacio público verde", "CA - Respaldo: espacio público verde total del POT"
    cp, cpr = "CP - Procesos participativos", "CP - Respaldo: índice situacional de implementación"
    cd, cdr = "CD - Servicios deportivos", "CD - Respaldo: disciplinas deportivas atendidas"
    ile, ile_av = "ILE primera infancia - cobertura de atención integral", "ILE primera infancia - avance hacia la meta 2027"
    clasif = "ILE Saber 11 - clasificación de planteles (meta 14 %)"
    sab = "ILE Saber 11 - puntaje global promedio, instituciones oficiales urbanas"
    icpd, icpd_r = "ICPD-6 - Índice", "ICPD-6 - Promedio parcial con los indicadores disponibles"
    dis = "Inclusión: personas con discapacidad en el registro nacional (acumulado)"
    ms_d, ms_a = "Cobertura de salud, cifra publicada por MinSalud (diciembre)", "Cobertura de salud, cifra publicada por MinSalud (agosto)"
    pol = "Logros de políticas públicas (nueva dimensión)"
    r = lambda ind, a: f"{fm(v(ind, a))}"
    return [
        IND(D, "**Cobertura de salud** (reemplaza el aseguramiento subsidiado)", "principal", "100 × afiliados (contributivos + subsidiados + especiales) / población DANE",
            ser(salud, (2024, 2025)) + f"; serie 2020-2025 calculada", "Gris",
            f"Sin meta oficial. MinSalud publicó {r(ms_d, 2025)} (dic. 2025) y {r(ms_a, 2026)} (ago. 2026) con otras poblaciones", [salud]),
        IND(D, "Cobertura de salud: discapacidad registrada", "respaldo", "Personas con discapacidad en el registro nacional (acumulado)",
            f"{fm(v(dis, 2020))} personas (acumulado a 2020)", "Gris", "Registro sin actualizar desde 2020; no es una tasa", [dis]),
        IND(D, "**CM** Mortalidad vial", "principal", "Logro frente a la meta (7,25 frente a 8 por 100.000 hab.)", ser(cm, (2024,)), sem_ultimo(cm, (2024,)),
            "Solo 2024; meta final 2027, sin metas anuales", [cm]),
        IND(D, "CM: cifra de la Policía", "respaldo", "Igual, con homicidios en accidente de tránsito de la Policía", ser(cmr, (2024,)), sem_ultimo(cmr, (2024,)),
            "Otra definición que la del INMLCF", [cmr]),
        IND(D, "**CA** Espacio público verde", "principal", "Logro frente a la meta (5,75 frente a 3,88 m² por hab.)", ser(ca, (2024,)), sem_ultimo(ca, (2024,)),
            "No concilia con la línea base del plan", [ca]),
        IND(D, "CA: espacio público verde total del POT", "respaldo", "Igual, con el total del POT (69,41 m² por hab.)", ser(car, (2024,)), sem_ultimo(car, (2024,)),
            "Otra serie: no comparable con la línea base", [car]),
        IND(D, "**CP** Procesos participativos", "principal", "Logro frente a la meta (100 %)", ser(cp, (2024, 2025)), sem_ultimo(cp, (2024, 2025)),
            "No hay método ni cifra oficial", [cp]),
        IND(D, "CP: índice situacional", "respaldo", "Índice del diagnóstico del PDM", f"{fm(v(cpr, 2022))} (2022)", "Gris", "Sin meta comparable; sin semáforo", [cpr]),
        IND(D, "**CD** Servicios deportivos", "principal", "Logro frente a la meta (cobertura de 6 %)", ser(cd, (2025,)), sem_ultimo(cd, (2025,)),
            "Falta la cifra de cobertura del Índer", [cd]),
        IND(D, "CD: disciplinas atendidas", "respaldo", "100 × disciplinas implementadas / previstas", ser(cdr, (2025,), con_razon=True), sem_ultimo(cdr, (2025,)),
            "Mide oferta, no población atendida", [cdr]),
        IND(D, "**ILE** Logro educativo (evolución)", "principal", "Serie anual, variación anual y avance = 100 × (valor − 16,3) / (19 − 16,3), entre 0 y 100",
            f"Primera infancia (0 a 5 años): {fm(v(ile, 2024))} (2024) y {fm(v(ile, 2025))} (2025); avance {fm(v(ile_av, 2024))} y {fm(v(ile_av, 2025))}. Serie 2019-2025 calculada. Planteles que mejoran en Saber 11: N/D",
            "Gris", f"2023 da {fm(v(ile, 2023))}, no 16,3 de la línea base; sin clasificación de planteles", [ile, clasif]),
        IND(D, "ILE: puntaje de Saber 11", "respaldo", "Puntaje global promedio de oficiales urbanas",
            pares((fm(v(sab, 2019)), 2019), (fm(v(sab, 2020)), 2020), (fm(v(sab, 2022)), 2022)), "Gris", "Datos públicos solo hasta 2022; sin semáforo", [sab]),
        IND(D, "**ICPD-6** Índice de cumplimiento del PDM", "principal", "IBS = (salud + ILE + CD) / 3; ITG = (CM + CA + CP) / 3; (IBS + ITG) / 2", "N/D (2024 y 2025)", "Gris",
            "Faltan salud, CP, CD y planteles", [icpd]),
        IND(D, "ICPD-6: promedio con las entradas disponibles", "respaldo", "Promedio simple de las entradas disponibles",
            f"{fm(v(icpd_r, 2024))} (2024, {fm(float(fila(icpd_r, 2024)['denominador']))} de 6 entradas); {fm(v(icpd_r, 2025))} (2025, {fm(float(fila(icpd_r, 2025)['denominador']))} de 6)",
            "Gris", "Referencial: no cumple la regla de la Guía de no calcular índices con entradas faltantes", [icpd_r]),
        IND(D, "**Logros de políticas públicas** (nueva)", "principal", "100 × políticas con informe de seguimiento publicado / políticas vigentes", "N/D (2025)", "Gris",
            "Sin inventario oficial de políticas vigentes", [pol]),
        IND(D, "Políticas: con plan de acción y metas cumplidas", "respaldo", "Políticas con plan de acción / vigentes; metas cumplidas / metas", "N/D", "Gris",
            "Para más adelante; sin fuente identificada", ["__politicas_respaldo__"]),
    ]


def ind_pp():
    D, A = "Presupuesto participativo", (2024, 2025, 2026)
    fid, mon = "Fidelidad a lo votado", "Fidelidad a lo votado, versión monetaria"
    ter, ejt = "Proyectos votados que se terminaron", "Ejecución presupuestal de los proyectos terminados"
    epp = "EPP, ejecución presupuestal del presupuesto participativo"
    epp26 = epp + " (1.er semestre)"
    tc, p90 = "Tasa de compromiso", "Proyectos con ejecución de al menos 90 %"
    eq = "Equidad entre zonas (brecha de ejecución)"
    return [
        IND(D, "**Fidelidad a lo votado**", "principal", "100 × proyectos con apropiación igual al valor votado (más o menos 1 %) / proyectos votados", ser(fid, A, con_razon=True), sem_ultimo(fid, A),
            "2024 aproximada; 2026 sin el valor votado de cada proyecto", [fid]),
        IND(D, "**Proyectos votados que se terminaron**", "principal", "100 × proyectos votados con avance físico de 100 % / proyectos votados", ser(ter, A, con_razon=True), sem_ultimo(ter, A),
            "Autorreporte de las dependencias, sin acta de recibo", [ter]),
        IND(D, "**EPP** Ejecución presupuestal", "principal", "100 × obligaciones (OPP) / apropiación definitiva (APP)",
            f"{fm(v(epp, 2024))} (2024); {fm(v(epp, 2025))} (2025); {fm(v(epp26, 2026))} (1.er semestre de 2026)", sem_ultimo(epp, (2024, 2025)),
            "2024 aproximada (50 de 51 proyectos cruzados); 2026 sin semáforo", [epp, epp26]),
        IND(D, "Ejecución de los proyectos terminados", "respaldo", "100 × obligaciones / apropiación, solo de los proyectos con avance físico de 100 %", ser(ejt, A), sem_ultimo(ejt, A),
            "Autorreporte de las dependencias, sin acta de recibo", [ejt]),
        IND(D, "Tasa de compromiso", "respaldo", "100 × compromisos / apropiación definitiva", ser(tc, (2024, 2025)), sem_ultimo(tc, (2024, 2025)),
            "2024: Hacienda no separa compromisos", [tc]),
        IND(D, "Proyectos con ejecución de al menos 90 %", "respaldo", "100 × proyectos con obligaciones ≥ 90 % de su apropiación / proyectos", ser(p90, (2024, 2025), con_razon=True),
            sem_ultimo(p90, (2024, 2025)), "2024 sobre 50 proyectos cruzados", [p90]),
        IND(D, "Equidad entre zonas", "respaldo", "Zona con mayor EPP − zona con menor EPP (puntos)", ser(eq, (2024, 2025)), "Gris",
            "Sin umbral: no tiene semáforo", [eq]),
        IND(D, "Fidelidad a lo votado, versión monetaria", "complemento", "100 × apropiación total / valor votado total", ser(mon, (2024, 2025)), "Gris",
            "Complemento de la fidelidad; puede superar 100", [mon]),
    ]


def _cuenta_dim(dim, tipo):
    gr = [n for d, t, n in CUENTA if d == dim and t == tipo]
    return len(gr), sum(1 for g in gr if con_valor(g))


def documento_1():
    CUENTA.clear()
    ind = {"Control social": ind_cs(), "Participación ciudadana": ind_pc(), "Cumplimiento del PDM": ind_pdm(), "Presupuesto participativo": ind_pp()}
    fid24 = fila("Fidelidad a lo votado", 2024)
    ter25 = fila("Proyectos votados que se terminaron", 2025)
    ejt25 = fila("Ejecución presupuestal de los proyectos terminados", 2025)
    partes = []
    cab_res = ["Dimensión", "Principales", "Con valor hoy", "Respaldos", "Con valor hoy", "Variables disponibles", "Variables construibles", "Variables no disponibles", "Total de variables"]
    res = []
    tot = [0] * 8
    for d in DIMS:
        p, pv = _cuenta_dim(d, "principal")
        rp, rv = _cuenta_dim(d, "respaldo")
        est = {e: len({x["id_variable"] for x in VARS if x["dimension"] == d and x["estado"] == e}) for e in ("Disponible", "Construible", "No disponible")}
        fila_r = [DIM_TIT[d], p, pv, rp, rv, est["Disponible"], est["Construible"], est["No disponible"]]
        fila_r.append(sum(est.values()))
        res.append(fila_r)
        for k in range(8):
            tot[k] += fila_r[k + 1]
    res.append(["**Total**"] + [f"**{x}**" for x in tot])
    tab_res = tabla_md(cab_res, [[str(x) for x in f] for f in res])
    for d in DIMS:
        partes.append(f"## {DIM_TIT[d]}\n\n**Variables**\n\n{tabla_variables(d)}\n\n**Indicadores**\n\n{tabla_indicadores(ind[d])}\n")
    md = f"""
## Introducción

Este documento reemplaza al anterior tras la reunión con el profesor Francisco Correa. Qué cambió: **(1)** documentos más cortos, organizados por dimensión; en cada una se ve variable por variable qué hay, qué falta y con qué limitaciones, y los indicadores se construyen desde esas variables; **(2)** regla del respaldo: cada indicador lleva un respaldo (o una variable proxy) que da una lectura provisional cuando el principal no se puede calcular, sin reemplazarlo ni sumarse a él; **(3)** en participación ciudadana se eliminan I5 e I6; **(4)** en el cumplimiento del PDM la cobertura de salud reemplaza al aseguramiento subsidiado, el logro educativo (ILE) pasa a medirse como evolución de una serie y se añade la dimensión de políticas públicas; **(5)** en presupuesto participativo los principales son la fidelidad a lo votado, los proyectos votados que se terminaron y la ejecución presupuestal (EPP).

**Nada está validado** ni se validó con Planeación: las cifras son provisionales y salen de documentos públicos consultados hasta el {FECHA}. **Estado de las variables.** *Disponible*: hay valor publicado para 2024 y 2025 (en las metas, para el horizonte del plan). *Construible*: no está publicado tal cual o faltan años, pero se puede calcular o armar con datos existentes o con una solicitud concreta. *No disponible*: no hay dato ni forma de armarlo con lo reunido. **Semáforo:** del año más reciente con valor; en presupuesto participativo usa los umbrales de PP (verde desde 90, amarillo desde 75, naranja desde 50), una propuesta. N/D: sin dato.

{chr(10).join(partes)}
## Resumen

Indicadores por dimensión y cuántos tienen valor hoy (al menos un año con cifra), y variables por estado. La versión monetaria de la fidelidad es un complemento y no se cuenta como respaldo.

{tab_res}

## Siglas

**APP** apropiación definitiva del presupuesto participativo. **CA** espacio público verde. **CD** servicios deportivos. **CM** mortalidad vial. **CP** procesos participativos. **DAP** Departamento Administrativo de Planeación. **DANE** Departamento Administrativo Nacional de Estadística. **DSI** disponibilidad de información de seguimiento. **EFPC** Escuela de Formación para la Participación Ciudadana. **EPP** ejecución del presupuesto participativo. **EPV** espacio público verde. **I1 a I4** incorporación de aportes, cobertura territorial, oferta formativa y publicación de resultados. **ICFES** Instituto Colombiano para la Evaluación de la Educación. **ICPD-6** índice de cumplimiento del PDM (IBS: bienestar social; ITG: gestión territorial). **ICRCS** índice de condiciones y respuesta para el control social. **ILE** logro educativo. **Índer** Instituto de Deporte y Recreación de Envigado. **INMLCF** Instituto Nacional de Medicina Legal y Ciencias Forenses. **ISPC** índice de seguimiento de la participación ciudadana. **MinSalud** Ministerio de Salud y Protección Social. **N/D** sin dato. **OPP** obligaciones del presupuesto participativo. **PCS** promoción del control social. **PDM** Plan de Desarrollo Municipal 2024-2027. **POT** Plan de Ordenamiento Territorial. **RPO** respuesta a observaciones ciudadanas. **SGSSS** Sistema General de Seguridad Social en Salud. **TerriData** portal de datos territoriales del Departamento Nacional de Planeación. **VRC** veedurías en la rendición de cuentas. Las letras entre paréntesis en las variables (MS, MD, EE, EP y similares) son las de la Matriz de los profesores.
"""
    return md.strip() + "\n", dict(res=res, tot=tot, fid24=fid24, ter25=ter25, ejt25=ejt25, ind=ind)


# ================================================================================================ EXCEL DEL SISTEMA
def plano(t):
    """Texto sin los enlaces en formato Markdown."""
    return re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", t)


def excel_sistema(destino, info):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    azul = PatternFill("solid", fgColor="1F4E78")
    fino = Side(style="thin", color="BFBFBF")
    col_estado = {"Disponible": "D6F0D6", "Construible": "DCEBFA", "No disponible": "F8D4D4"}
    col_sem = {"Verde": "D6F0D6", "Amarillo": "FFF3BF", "Naranja": "FFDEB8", "Rojo": "F8D4D4", SIN_SEM: "E1E1E1"}

    def cabecera(ws, cols, anchos):
        ws.append(cols)
        for c in ws[1]:
            c.fill, c.font, c.alignment = azul, Font(bold=True, color="FFFFFF"), Alignment(wrap_text=True, vertical="center")
        for i, w in enumerate(anchos):
            ws.column_dimensions[chr(65 + i) if i < 26 else "A" + chr(65 + i - 26)].width = w
        ws.freeze_panes = "A2"

    def enlace_celda(c, url):
        if url:
            c.value, c.hyperlink = url, url
            c.font = Font(color="0563C1", underline="single")

    # --- Variables por dimensión (una fila por variable; columnas de años con el valor)
    wv = wb.active
    wv.title = "Variables por dimensión"
    anios = list(range(2019, 2028))
    cols = ["Dimensión", "ID de la variable", "Variable", "Origen", "Para qué indicador", "Estado", "Qué hay hoy (valor, años y fuente)", "Qué falta", "Limitación",
            "Unidad"] + [str(a) if a != 2027 else "2027 (meta)" for a in anios] + ["Fuente", "Enlace"]
    cabecera(wv, cols, [20, 17, 40, 9, 22, 14, 48, 36, 44, 14] + [11] * len(anios) + [38, 40])
    for d in DIMS:
        for vs in vars_dim(d):
            v0 = vs[0]
            porano = {r["anio"]: r["v"] for r in vs if r["v"] is not None}
            fila_ = [DIM_TIT[d], v0["id_variable"], v0["variable"], v0["origen"], v0["para_indicador"], v0["estado"], plano(hoy(vs)), v0["que_falta"], v0["limitacion"], v0["unidad"]]
            fila_ += [porano.get(a) for a in anios] + [v0["fuente"], None]
            wv.append(fila_)
            r_ = wv.max_row
            wv.cell(r_, 6).fill = PatternFill("solid", fgColor=col_estado[v0["estado"]])
            for k in range(11, 11 + len(anios)):
                wv.cell(r_, k).number_format = "#,##0.00" if v0["unidad"] not in ("pesos", "personas") else "#,##0"
            enlace_celda(wv.cell(r_, 10 + len(anios) + 2), v0["enlace"])
    for row in wv.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=fino)
    wv.auto_filter.ref = wv.dimensions

    # --- Indicadores por dimensión
    wi = wb.create_sheet("Indicadores por dimensión")
    cabecera(wi, ["Dimensión", "Indicador", "Tipo", "Fórmula", "Valor por año", "Semáforo (año más reciente con valor)", "Limitación", "Con valor hoy", "Valor 2024", "Valor 2025", "Valor 2026",
                  "Fuente", "Enlace"], [20, 40, 12, 46, 46, 16, 44, 10, 11, 11, 11, 34, 40])
    for d in DIMS:
        nombres = [n for dd, t, n in CUENTA if dd == d]
        for row, nom in zip(info["ind"][d], nombres):
            filas_i = [r for r in FILAS if r["indicador"] in nom]
            def valor(a):
                c = [r for r in filas_i if r["indicador"] == nom[0] and r["anio"] == a]
                return c[0]["v"] if c else None
            con = [r for r in filas_i if r["enlace"]]
            ult = sorted(con, key=lambda r: r["anio"])[-1] if con else None
            wi.append([DIM_TIT[d], plano(row[0]).replace("**", ""), row[1], row[2], plano(row[3]), row[4], row[5], "Sí" if con_valor(nom) else "No",
                       valor(2024), valor(2025), valor(2026), ult["fuente"] if ult else "", None])
            r_ = wi.max_row
            wi.cell(r_, 6).fill = PatternFill("solid", fgColor=col_sem[row[4]])
            for k in (9, 10, 11):
                wi.cell(r_, k).number_format = "#,##0.00"
            enlace_celda(wi.cell(r_, 13), ult["enlace"] if ult else "")
    for row in wi.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=fino)
    wi.auto_filter.ref = wi.dimensions

    # --- Léeme
    lm = wb.create_sheet("Léeme", 0)
    cabecera(lm, ["Tema", "Contenido"], [30, 120])
    res, tot = info["res"], info["tot"]
    filas_lm = [
        ("Versión", f"{FECHA}. PROVISIONAL: propuesta de trabajo del Observatorio Ciudadano de Envigado (OCE)."),
        ("Qué es este archivo", "Sistema de indicadores simplificado, organizado por dimensión (control social, participación ciudadana, cumplimiento del Plan de Desarrollo Municipal y presupuesto participativo). "
                                "Para cada dimensión dice, variable por variable, qué hay, qué falta y con qué limitaciones, y los indicadores que se construyen desde esas variables."),
        ("Regla del respaldo", "Cada indicador principal lleva un respaldo (o una variable proxy) que da una lectura provisional cuando el principal no se puede calcular; el respaldo no lo reemplaza ni se suma a él."),
        ("Estado de las variables", "Disponible: hay valor publicado para 2024 y 2025 (en las metas, para el horizonte del plan).\n"
                                    "Construible: no está publicado tal cual o faltan años, pero se puede calcular o armar con datos existentes o con una solicitud concreta.\n"
                                    "No disponible: no hay dato ni forma de armarlo con lo reunido."),
        ("Semáforo", "Corresponde al año más reciente con valor. En presupuesto participativo usa los umbrales de PP (verde desde 90, amarillo desde 75, naranja desde 50), que son una propuesta. "
                     "Los indicadores sin meta o umbral aparecen como Sin semáforo."),
        ("Resultados", "\n".join(f"{r[0]}: {r[1]} indicadores principales ({r[2]} con valor hoy) y {r[3]} de respaldo ({r[4]} con valor hoy); variables: {r[5]} disponibles, {r[6]} construibles, {r[7]} no disponibles."
                                  for r in res[:-1]) +
                       f"\nTotal: {tot[0]} principales ({tot[1]} con valor hoy), {tot[2]} respaldos ({tot[3]} con valor hoy); {tot[7]} variables ({tot[4]} disponibles, {tot[5]} construibles, {tot[6]} no disponibles)."),
        ("Cómo leer las hojas", "Variables por dimensión: una fila por variable, con el valor de cada año en las columnas de años y el enlace a la fuente. "
                                "Indicadores por dimensión: una fila por indicador, con tipo (principal, respaldo o complemento), fórmula, valores y semáforo coloreado. Ambas hojas tienen filtros."),
        ("Colores", "Estado: verde = Disponible, azul = Construible, rojo = No disponible. Semáforo: verde, amarillo, naranja y rojo; gris = sin semáforo."),
        ("Aviso: nada está validado", "Ninguna cifra ha sido validada ni confirmada por la Alcaldía ni por Planeación. Las cifras de presupuesto participativo se basan en una lista de proyectos reconstruida; "
                                      "los proyectos terminados son un autorreporte de las dependencias, sin acta de recibo."),
    ]
    for t, c in filas_lm:
        lm.append([t, c])
    for row in lm.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        row[0].font = Font(bold=True)
    wb.active = 0
    wb.save(destino)



# ================================================================================================ DOCUMENTO 2
VARIABLES = [
    # (variable, para qué indicador, a quién, formato, prioridad)
    ("Inventario de procesos participativos cerrados (fecha, zona, si el aporte entró a una decisión y si se publicó el resultado)",
     "I1 y I4 de participación ciudadana (y su respaldo)", "Subdirección de Participación Ciudadana (DAP)", "Hoja de cálculo, una fila por proceso", "Alta"),
    ("Valor votado de cada proyecto de presupuesto participativo 2026 (listado oficial)",
     "Fidelidad a lo votado 2026; EPP por proyecto", "EFPC; Subdirección de Participación Ciudadana", "Hoja de cálculo o PDF del listado", "Alta"),
    ("Lista oficial de códigos de proyecto financiados con presupuesto participativo (2024 a 2026)",
     "EPP y fidelidad a lo votado", "DAP; Secretaría de Hacienda", "Hoja de cálculo", "Alta"),
    ("Ejecución de egresos por código de proyecto (apropiación, compromisos y obligaciones, con corte de cierre)",
     "EPP, tasa de compromiso, proyectos con 90 % o más, equidad entre zonas", "Secretaría de Hacienda", "Hoja de cálculo o CSV", "Alta"),
    ("Actas de recibo, informes finales y de supervisión de los proyectos de presupuesto participativo",
     "Proyectos votados que se terminaron (verificar el autorreporte)", "Subdirección de Participación Ciudadana; unidades ejecutoras", "Carpeta de PDF con un índice", "Alta"),
    ("Avance físico por proyecto de 2024 y al cierre de 2026 (plan de acción con corte de diciembre)",
     "Proyectos votados que se terminaron y ejecución de los terminados", "DAP; unidades ejecutoras", "Hoja de cálculo (plan de acción)", "Alta"),
    ("Actas de rendición de cuentas con listas de asistencia (sin datos personales) y calendario de eventos",
     "VRC de control social", "Secretarías (Salud, Bienestar Social y otras); DAP", "PDF y tabla de eventos", "Alta"),
    ("Cuadros completos de preguntas y respuestas de las rendiciones 2024 a 2026",
     "RPO de control social", "DAP; Oficina Asesora de Comunicaciones; Secretaría de Salud", "Hoja de cálculo", "Alta"),
    ("Corte actualizado del registro de veedurías vigentes", "VRC de control social", "Personería de Envigado", "Hoja de cálculo", "Media"),
    ("Valores observados 2024 y 2025 de las metas del plan seleccionadas", "DSI de control social", "DAP, Dirección de Planeación", "Hoja de cálculo (evaluación del plan indicativo)", "Media"),
    ("Soportes de las estrategias de promoción del control social ejecutadas", "PCS de control social", "DAP, Subdirección de Participación", "PDF o enlaces", "Baja"),
    ("Plan de formación aprobado y relación de acciones formativas programadas y culminadas (sin datos personales)",
     "I3 de participación ciudadana", "EFPC", "Hoja de cálculo", "Alta"),
    ("Actas o listados de asistencia de las asambleas por zona", "I2 de participación ciudadana", "Subdirección de Participación Ciudadana", "PDF o tabla por zona", "Media"),
    ("Ficha técnica del indicador de procesos participativos (valor, fórmula y fuente)", "CP del PDM", "DAP", "Ficha en PDF", "Alta"),
    ("Cobertura deportiva: numerador, denominador y valor", "CD del PDM", "Índer", "Hoja de cálculo", "Alta"),
    ("Clasificación de planteles de Saber 11 por institución, 2023 a 2025", "ILE (Saber 11) del PDM", "Secretaría de Educación", "Hoja de cálculo por institución", "Alta"),
    ("Resultados de Saber 11 por establecimiento de 2021 y de 2023 a 2025", "ILE (respaldo: puntaje promedio) del PDM", "Secretaría de Educación; ICFES", "Hoja de cálculo por institución", "Media"),
    ("Niños y niñas atendidos de forma integral 2023 a 2025 y la población usada para la línea base", "ILE (primera infancia) del PDM", "Secretaría de Educación", "Hoja de cálculo", "Alta"),
    ("Afiliados al sistema de salud por régimen y población de referencia, cortes de diciembre y el último de 2026", "Cobertura de salud del PDM",
     "Secretaría de Salud, Dirección de Aseguramiento", "Hoja de cálculo", "Media"),
    ("Registro actualizado de personas con discapacidad del municipio", "Respaldo de inclusión (cobertura de salud)",
     "Secretaría de Salud o Secretaría de Inclusión Social", "Hoja de cálculo sin datos personales", "Media"),
    ("Inventario de políticas públicas municipales vigentes, con su acto, informes de seguimiento y plan de acción", "Logros de políticas públicas del PDM",
     "Concejo Municipal; DAP", "Hoja de cálculo con enlaces", "Alta"),
    ("Serie anual de muertes por siniestros viales 2023 a 2026", "CM del PDM", "Secretaría de Movilidad", "Hoja de cálculo", "Baja"),
    ("Ficha del indicador de espacio público verde (metros cuadrados por habitante)", "CA del PDM", "Secretaría de Medio Ambiente y Desarrollo Agropecuario", "Ficha en PDF", "Baja"),
]


def documento_2_md():
    cab = ["Variable", "Para qué indicador sirve", "A quién pedirla", "Formato ideal", "Prioridad"]
    lin = ["| " + " | ".join(cab) + " |", "|" + "---|" * 5]
    lin += ["| " + " | ".join(x.replace("|", "/") for x in f) + " |" for f in VARIABLES]
    return ("Listado de lo que falta para calcular los indicadores del sistema simplificado y los respaldos propuestos. Es solo una lista de trabajo, no una solicitud formal: "
            "no incluye datos personales y cada ítem debe ratificarse con la dependencia antes de pedirse.\n\n" + "\n".join(lin) + "\n\n" +
            "**Siglas.** **DAP** Departamento Administrativo de Planeación. **EFPC** Escuela de Formación para la Participación Ciudadana. **EPP** ejecución del presupuesto participativo. "
            "**Índer** Instituto de Deporte y Recreación de Envigado. **PDM** Plan de Desarrollo Municipal. **I1 a I4** indicadores de incorporación de aportes, cobertura territorial, oferta formativa y publicación de resultados. "
            "**PCS, DSI, VRC y RPO** indicadores de control social (promoción, disponibilidad de información, veedurías y respuesta a observaciones). **CM, CA, CP y CD** indicadores del PDM (mortalidad vial, espacio público verde, procesos participativos y servicios deportivos). "
            "**ILE** logro educativo. **Saber 11** prueba de Estado para estudiantes de último grado.\n")


def excel_variables(destino):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "Variables"
    cab = ["Variable", "Para qué indicador sirve", "A quién pedirla", "Formato ideal", "Prioridad"]
    ws.append(cab)
    for f in VARIABLES:
        ws.append(list(f))
    azul = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.fill = azul
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    color = {"Alta": "F8D4D4", "Media": "FFF3BF", "Baja": "D6F0D6"}
    fino = Side(style="thin", color="BFBFBF")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=fino)
        row[4].fill = PatternFill("solid", fgColor=color[row[4].value])
    for col, w in zip("ABCDE", (62, 40, 40, 30, 11)):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    lm = wb.create_sheet("Léeme", 0)
    for i, t in enumerate(["Variables para solicitar - sistema de indicadores simplificado (propuesta, provisional)", FECHA,
                           "Listado de trabajo, no una solicitud formal. No contiene datos personales.",
                           "Prioridad: Alta = sin esta variable un indicador principal no se puede calcular; Media = mejora o confirma un resultado; Baja = complementa."], 1):
        lm.cell(i, 1, t).alignment = Alignment(wrap_text=True, vertical="top")
    lm["A1"].font = Font(bold=True, size=13)
    lm.column_dimensions["A"].width = 110
    wb.active = 1
    wb.save(destino)


# ================================================================================================ principal
def main():
    sin_pdf = "--sin-pdf" in sys.argv
    TMP.mkdir(parents=True, exist_ok=True)
    # ---------- documento 1
    md1, info1 = documento_1()
    t_pdf, tam = a_tex(md1, False)
    t_word, tam_w = a_tex(md1, True)
    titulo1 = "Sistema de indicadores del Observatorio Ciudadano de Envigado: versión simplificada (propuesta)"
    corto1 = "OCE -- Sistema de indicadores simplificado (propuesta, provisional)"
    tex_pdf = documento_tex(corto1, titulo1, t_pdf, False)
    tex_word = documento_tex(corto1, titulo1, t_word, True)
    pdf1, docx1 = OUT / "sistema_indicadores_simplificado.pdf", OUT / "sistema_indicadores_simplificado.docx"
    if not sin_pdf:
        compilar_pdf(tex_pdf, "sistema_indicadores_simplificado", pdf1)
    construir_word(tex_word, tam_w, docx1, "OCE – Sistema de indicadores simplificado (propuesta, provisional)")
    excel_sistema(OUT / "sistema_indicadores_simplificado.xlsx", info1)
    # ---------- documento 2
    md2 = documento_2_md()
    t2_pdf, _ = a_tex(md2, False)
    t2_word, tam2_w = a_tex(md2, True)
    titulo2 = "Variables que faltan para los indicadores del Observatorio Ciudadano de Envigado"
    corto2 = "OCE -- Variables para solicitar (propuesta, provisional)"
    tex2_pdf = documento_tex(corto2, titulo2, t2_pdf, False)
    tex2_word = documento_tex(corto2, titulo2, t2_word, True)
    docx2, xlsx2 = OUT / "variables_para_solicitar.docx", OUT / "variables_para_solicitar.xlsx"
    pdf2_tmp = TMP / "salida" / "variables_para_solicitar.pdf"
    pdf2_tmp.parent.mkdir(parents=True, exist_ok=True)
    if not sin_pdf:
        compilar_pdf(tex2_pdf, "variables_para_solicitar", pdf2_tmp)
    construir_word(tex2_word, tam2_w, docx2, "OCE – Variables para solicitar (propuesta, provisional)")
    excel_variables(xlsx2)
    # ---------- páginas
    if not sin_pdf:
        print("PDF documento 1:", paginas_pdf(pdf1), "páginas")
        print("PDF documento 2 (LaTeX):", paginas_pdf(pdf2_tmp), "páginas")
    n1, _p = paginas_word(docx1, TMP / "word_pdf")
    n2, _p = paginas_word(docx2, TMP / "word_pdf")
    print("Word documento 1:", n1, "páginas;  Word documento 2:", n2, "páginas")


if __name__ == "__main__":
    main()
