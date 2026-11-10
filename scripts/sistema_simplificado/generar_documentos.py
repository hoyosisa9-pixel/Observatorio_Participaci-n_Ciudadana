#!/usr/bin/env python3
"""Genera los dos documentos cortos del sistema simplificado (PROVISIONAL), con las cifras leídas de
outputs/sistema_simplificado/indicadores_simplificados.csv:

  outputs/sistema_simplificado/sistema_indicadores_simplificado.{pdf,docx}   (máximo 5 páginas)
  outputs/sistema_simplificado/variables_para_solicitar.{docx,xlsx}           (máximo 2 páginas)

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
TMP = G.SCRATCH / "simplificado"
G.TMP = TMP
FECHA = "10 de noviembre de 2026"
MARGEN = 1.6                      # cm
ANCHO = 27.94 - 2 * MARGEN        # cm de texto en horizontal (carta)
G.ANCHO_V = G.ANCHO_H = ANCHO
PRIORIDAD = {"Alta": "crojo", "Media": "camarillo", "Baja": "cverde"}

# ------------------------------------------------------------------------------------------------ datos
with open(F_CSV, encoding="utf-8-sig", newline="") as f:
    FILAS = list(csv.DictReader(f, delimiter=";"))
for r in FILAS:
    r["anio"] = int(r["anio"])
    r["v"] = float(r["valor"]) if r["valor"] != "" else None


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

FRACCIONES = {7: [0.11, 0.14, 0.15, 0.18, 0.08, 0.20, 0.14], 5: [0.30, 0.22, 0.23, 0.15, 0.10]}


def _anchos(filas, total, tam):
    n = len(filas[0])
    util = total - n * 0.22
    return [util * f for f in FRACCIONES[n]]


G.anchos = _anchos
_celda_orig = G.celda_tex


def _celda(texto, enc, par):
    plano = re.sub(r"\*", "", texto.strip())
    if enc.startswith("Prioridad") and plano in PRIORIDAD:
        return r"\cellcolor{" + PRIORIDAD[plano] + "}" + G.en_linea(texto.strip(), tabla=True)
    return _celda_orig(texto, enc, par)


G.celda_tex = _celda
_tabla_orig = G.tabla_tex


def _tabla(filas, ancho):
    t, tam = _tabla_orig(filas, ancho)
    return t, tam


# ------------------------------------------------------------------------------------------------ LaTeX / Word
def preambulo(titulo_corto, titulo_pdf):
    p = G.PREAMBULO.replace("\\documentclass[11pt,letterpaper]{article}", "\\documentclass[10pt,letterpaper]{article}")
    p = p.replace("\\usepackage[margin=MARGEN,headheight=14pt]{geometry}",
                  f"\\usepackage[landscape,margin={MARGEN}cm,headheight=14pt,headsep=0.5cm]{{geometry}}")
    p = p.replace("[hidelinks,breaklinks=true]", "[colorlinks=true,linkcolor=black,urlcolor=blue!55!black,breaklinks=true]")
    p = p.replace("COLORES", G.COLORES_TEX)
    p = p.replace("\\newcommand{\\pagsize}{\\special{papersize=21.59cm,27.94cm}}",
                  "\\newcommand{\\pagsize}{\\special{papersize=27.94cm,21.59cm}}")
    p = p.replace("\\setlength{\\parskip}{0.5em}", "\\setlength{\\parskip}{0.35em}")
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
    return (preambulo(titulo_corto, titulo) + "\n\\begin{document}\n\\hypersetup{pageanchor=true}\n\n" + cab + aviso +
            cuerpo_tex + "\n\n\\end{document}\n")


def a_tex(md, word):
    bl = G.parsear_md(md)
    T, tam, _ = G.a_latex(bl, word=word)
    return "\n\n".join(T), tam


def construir_word(tex_word, tamanos, destino, encabezado):
    bls = D.bloques(tex_word)
    colores = {m.group(1): "%02X%02X%02X" % tuple(int(x) for x in m.group(2).split(","))
               for m in re.finditer(r"\\definecolor\{(\w+)\}\{RGB\}\{([\d,]+)\}", tex_word)}
    doc = Document()
    D.preparar_estilos(doc)
    D.fijar_fuente(doc.styles["Normal"], D.FUENTE, 10)
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
    sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.5)
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
def tabla_md(filas):
    cab = ["Indicador", "Qué mide", "Fórmula", "Estado hoy (valor y año)", "Semáforo", "Respaldo o proxy", "Fuente"]
    lin = ["| " + " | ".join(cab) + " |", "|" + "---|" * len(cab)]
    for f in filas:
        assert len(f) == 7, f
        lin.append("| " + " | ".join(x.replace("|", "/") for x in f) + " |")
    return "\n".join(lin)


def ind_cs():
    A = (2024, 2025)
    I = "Control social"
    pcs = "PCS - Promoción del control social"
    dsi = "DSI - Disponibilidad de información de seguimiento"
    vrc, rpo = "VRC - Veedurías en la rendición de cuentas", "RPO - Respuesta a observaciones ciudadanas"
    icr = "ICRCS - Índice de Condiciones y Respuesta para el Control Social"
    filas = [
        ["**PCS** Promoción del control social", "Estrategias de promoción ejecutadas frente a las programadas", "100 × ejecutadas / programadas",
         pares(*[(fm(v(pcs, a)), a) for a in A]), sem_ultimo(pcs, A),
         f"Número de estrategias ejecutadas: {fm(v('PCS - Respaldo: estrategias ejecutadas (número)', 2024))} por año",
         lk("Plan Indicativo de Productos (Planeación)", enlace(pcs, 2024))],
        ["**DSI** Disponibilidad de información de seguimiento", "Metas del plan con valor publicado a tiempo", "100 × metas con valor / metas seleccionadas",
         pares(*[(fm(v(dsi, a)), a) for a in A]), sem_ultimo(dsi, A),
         f"Con todo lo publicado hasta hoy: {fm(v('DSI - Respaldo: con todo lo publicado hasta la fecha de consulta', 2024))} (2024), "
         f"{fila('DSI - Respaldo: con todo lo publicado hasta la fecha de consulta', 2024)['semaforo']}",
         lk("Plan Indicativo, informes de gestión", enlace(dsi, 2024))],
        ["**VRC** Veedurías en la rendición de cuentas", "Eventos de rendición con presencia de veedurías", "100 × eventos con veeduría / eventos",
         "No disponible (2024 y 2025)", "Gris",
         f"Evento sin mención = ausencia: {fm(v('VRC - Respaldo: evento sin mención de veeduría = ausencia', 2024))} (2024, "
         f"{fila('VRC - Respaldo: evento sin mención de veeduría = ausencia', 2024)['semaforo']}) y "
         f"{fm(v('VRC - Respaldo: evento sin mención de veeduría = ausencia', 2025))} (2025, "
         f"{fila('VRC - Respaldo: evento sin mención de veeduría = ausencia', 2025)['semaforo']}); provisional",
         lk("Actas y cuadros de rendición", enlace(vrc, 2024))],
        ["**RPO** Respuesta a observaciones ciudadanas", "Observaciones con respuesta de fondo publicada", "100 × respondidas / observaciones",
         f"{fm(v(rpo, 2024))} (2024); no disponible (2025)", sem_ultimo(rpo, A),
         f"2025 con el cuadro publicado: {fm(v('RPO - Respaldo: con el cuadro de preguntas y respuestas publicado', 2025))}, "
         f"{fila('RPO - Respaldo: con el cuadro de preguntas y respuestas publicado', 2025)['semaforo']} (solo 2 observaciones)",
         lk("Cuadros de preguntas y respuestas", enlace(rpo, 2024))],
        ["**ICRCS** Índice de Condiciones y Respuesta para el Control Social", "Resumen del componente", "(DSI + PCS + VRC + RPO) / 4",
         "No disponible (2024 y 2025)", "Gris",
         f"Con los respaldos de VRC y RPO: {fm(v('ICRCS - Respaldo: con los respaldos de VRC y RPO', 2024))} (2024, "
         f"{fila('ICRCS - Respaldo: con los respaldos de VRC y RPO', 2024)['semaforo']}) y "
         f"{fm(v('ICRCS - Respaldo: con los respaldos de VRC y RPO', 2025))} (2025, "
         f"{fila('ICRCS - Respaldo: con los respaldos de VRC y RPO', 2025)['semaforo']}); provisional",
         "Los cuatro indicadores anteriores"],
    ]
    return filas, dict(
        principales=[[pcs], [dsi], [vrc], [rpo], [icr]],
        respaldos=[["PCS - Respaldo: estrategias ejecutadas (número)"], ["DSI - Respaldo: con todo lo publicado hasta la fecha de consulta"],
                   ["VRC - Respaldo: evento sin mención de veeduría = ausencia"],
                   ["RPO - Respaldo: con el cuadro de preguntas y respuestas publicado"], ["ICRCS - Respaldo: con los respaldos de VRC y RPO"]])


def ind_pc():
    A = (2024, 2025)
    i1, i2, i3, i4 = ("I1 - Incorporación de aportes", "I2 - Cobertura territorial", "I3 - Oferta formativa", "I4 - Publicación de resultados")
    r1 = "I1 - Respaldo: proyectos votados con presupuesto asignado"
    r2 = "I2 - Respaldo: zonas con asamblea de presupuesto participativo"
    r3 = "I3 - Respaldo: acciones formativas culminadas (número)"
    r4 = "I4 - Respaldo: procesos con documento de resultados (universo exploratorio)"
    ispc_o = "ISPC - Índice (fórmula original: 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4)"
    ispc_r = "ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)"
    filas = [
        ["**I1** Incorporación de aportes", "Procesos en que un aporte ciudadano entró a una decisión formal", "100 × procesos con aporte incorporado / procesos cerrados",
         "No disponible (sin inventario de procesos)", "Gris",
         f"Proyectos votados con presupuesto asignado / votados: {razon(r1, 2024)} (2024), {razon(r1, 2025)} (2025); provisional. "
         "Usa los mismos datos de la fidelidad de presupuesto participativo; no sumar en dos componentes",
         lk("Listado EFPC, Hacienda y plan de acción", enlace(r1, 2024))],
        ["**I2** Cobertura territorial", "Zonas con al menos una actividad participativa", "100 × zonas con actividad / 13",
         pares((fm(v(i2, 2024)), 2024), (fm(v(i2, 2025)), 2025)), sem_ultimo(i2, A),
         f"Asambleas de presupuesto participativo: {razon(r2, 2024)} (2024 y 2025)", lk("Informe de gestión 2025, PDM", enlace(i2, 2025))],
        ["**I3** Oferta formativa", "Acciones formativas ejecutadas frente a las programadas", "100 × ejecutadas / programadas",
         f"{fm(v(i3, 2025))} (2025); no disponible (2024)", sem_ultimo(i3, A),
         f"Acciones culminadas: {fm(v(r3, 2024))} (2024), {fm(v(r3, 2025))} (2025)", lk("Informe de gestión 2025", enlace(i3, 2025))],
        ["**I4** Publicación de resultados", "Procesos con resultados publicados", "100 × con resultados publicados / procesos cerrados",
         "No disponible (sin inventario de procesos)", "Gris",
         f"Procesos con documento de resultados (universo exploratorio): {fm(v(r4, 2024))} (2024), {fm(v(r4, 2025))} (2025); construible",
         "Portal de la Alcaldía; Subdirección de Participación"],
        ["**ISPC** Índice de Seguimiento de la Participación Ciudadana", "Resumen del componente",
         "Original: 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4. Opción: 0,556·I2 + 0,444·I3",
         "No disponible (original)", "Gris",
         f"Opción redistribuida: {fm(v(ispc_r, 2025))} (2025), {fila(ispc_r, 2025)['semaforo']}; 2024 no disponible. Provisional",
         "I2 e I3"],
    ]
    return filas, dict(principales=[[i1], [i2], [i3], [i4], [ispc_o]], respaldos=[[r1], [r2], [r3], [r4], [ispc_r]])


def ind_pdm():
    salud = "Cobertura de salud"
    cm, cmr = "CM - Mortalidad vial", "CM - Respaldo: cifra de la Policía"
    ca, car = "CA - Espacio público verde", "CA - Respaldo: espacio público verde total del POT"
    cp, cpr = "CP - Procesos participativos", "CP - Respaldo: índice situacional de implementación"
    cd, cdr = "CD - Servicios deportivos", "CD - Respaldo: disciplinas deportivas atendidas"
    ile = "ILE primera infancia - cobertura de atención integral"
    ile_av = "ILE primera infancia - avance hacia la meta 2027"
    sab = "ILE Saber 11 - puntaje global promedio, instituciones oficiales urbanas"
    icpd = "ICPD-6 - Promedio parcial con los indicadores disponibles"
    dis = "Inclusión: personas con discapacidad en el registro nacional (acumulado)"
    ms_d, ms_a = "Cobertura de salud, cifra publicada por MinSalud (diciembre)", "Cobertura de salud, cifra publicada por MinSalud (agosto)"
    ser_s = [(a, v(salud, a)) for a in range(2020, 2026)]
    sab_s = [(a, v(sab, a)) for a in (2019, 2020, 2022)]
    filas = [
        ["**Cobertura de salud** (nuevo; reemplaza el aseguramiento subsidiado)", "Población afiliada al sistema de salud",
         "100 × afiliados (contributivo + subsidiado + especiales) / población DANE",
         f"Serie 2020-2025 con una sola proyección de población: {fm(ser_s[0][1])} (2020) a {fm(ser_s[-1][1])} (2025). "
         f"MinSalud publicó {fm(v(ms_d, 2025))} (dic-2025) y {fm(v(ms_a, 2026))} (ago-2026) con otras poblaciones", "Gris",
         f"Discapacidad o inclusión: {fm(v(dis, 2020))} personas en el registro nacional (acumulado a 2020); por verificar",
         lk("MinSalud, TerriData", "https://terridata.dnp.gov.co/")],
        ["**CM** Mortalidad vial", "Muertes en vías por 100.000 habitantes frente a la meta", "Logro frente a la meta (Matriz)",
         f"{fm(v(cm, 2024))} (2024; {fm(float(fila(cm, 2024)['numerador']))} frente a {fm(float(fila(cm, 2024)['denominador']))})",
         sem_ultimo(cm, [2024]),
         f"Policía: {fm(float(fila(cmr, 2024)['numerador']))} → {fm(v(cmr, 2024))}, {fila(cmr, 2024)['semaforo']}",
         lk("Forensis (INMLCF), Policía", enlace(cm, 2024))],
        ["**CA** Espacio público verde", "Metros cuadrados por habitante frente a la meta", "Logro frente a la meta",
         f"{fm(v(ca, 2024))} (2024; {fm(float(fila(ca, 2024)['numerador']))} frente a {fm(float(fila(ca, 2024)['denominador']))})", sem_ultimo(ca, [2024]),
         f"EPV total del POT {fm(float(fila(car, 2024)['numerador']))} (otra serie)", lk("POT, PDM", enlace(ca, 2024))],
        ["**CP** Procesos participativos", "Implementación oficial de los procesos de participación", "Logro frente a la meta",
         "No disponible (sin ficha oficial)", "Gris", f"Índice situacional {fm(v(cpr, 2022))} (2022)", lk("PDM; ficha al DAP", enlace(cp, 2024))],
        ["**CD** Servicios deportivos", "Cobertura deportiva", "Logro frente a la meta", "No disponible", "Gris",
         f"Disciplinas atendidas {razon(cdr, 2025)} = {fm(v(cdr, 2025))} (2025), {fila(cdr, 2025)['semaforo']}; provisional",
         lk("Informe de gestión (Índer)", enlace(cdr, 2025))],
        ["**ILE** Logro educativo (evolución)", "Avance de la serie hacia la meta 2027",
         "Serie anual; variación anual; avance = 100 × (valor − línea base) / (meta − línea base), entre 0 y 100",
         f"Primera infancia (0 a 5 años): {fm(v(ile, 2024))} (2024) y {fm(v(ile, 2025))} (2025); avance {fm(v(ile_av, 2024))} y {fm(v(ile_av, 2025))}. "
         f"Serie 2019-2025 calculada", "Gris",
         f"Saber 11, puntaje promedio de oficiales urbanas: {fm(sab_s[0][1])} (2019), {fm(sab_s[1][1])} (2020), {fm(sab_s[2][1])} (2022); "
         "sin clasificación de planteles",
         lk("TerriData, ICFES", "https://www.datos.gov.co/d/kgxf-xxbe") + "; Secretaría de Educación"],
        ["**ICPD-6** Índice de cumplimiento del PDM", "Resumen del componente", "IBS = (salud + ILE + CD)/3; ITG = (CM + CA + CP)/3; (IBS + ITG)/2",
         "No disponible", "Gris",
         f"Promedio con las entradas disponibles: {fm(v(icpd, 2024))} (2024) y {fm(v(icpd, 2025))} (2025), con {fm(float(fila(icpd, 2024)['denominador']))} y "
         f"{fm(float(fila(icpd, 2025)['denominador']))} de 6 entradas; referencial: no cumple la regla de la Guía de no calcular índices con entradas faltantes",
         "Los indicadores anteriores"],
        ["**Logros de políticas públicas** (nueva dimensión)", "Seguimiento de las políticas públicas municipales",
         "a) 100 × políticas con informe de seguimiento publicado / políticas vigentes; b) con plan de acción vigente / vigentes; c) metas cumplidas / metas",
         "No disponible (sin inventario oficial); a) construible armando la lista con los acuerdos del Concejo", "Gris",
         "a) hoy; b) y c) en el futuro", "Concejo (acuerdos); Planeación"],
    ]
    return filas, dict(
        principales=[[salud], [cm], [ca], [cp], [cd], [ile, ile_av, "ILE Saber 11 - clasificación de planteles (meta 14 %)"], [icpd.replace("Promedio parcial", "Índice")],
                     ["Logros de políticas públicas (nueva dimensión)"]],
        respaldos=[[dis], [cmr], [car], [cpr], [cdr], [sab], [icpd], ["__politicas_respaldo__"]])


def ind_pp():
    A = (2024, 2025)
    fid = "Fidelidad a lo votado"
    mon = "Fidelidad a lo votado, versión monetaria"
    epp = "EPP, ejecución presupuestal del presupuesto participativo"
    epp26 = epp + " (1.er semestre)"
    tc, p90 = "Tasa de compromiso", "Proyectos con ejecución de al menos 90 %"
    eq, af = "Equidad entre zonas (brecha de ejecución)", "Avance físico (proyectos reportados terminados)"
    filas = [
        ["**Fidelidad a lo votado** (principal)", "Proyectos cuya apropiación coincide con lo votado",
         "100 × proyectos con apropiación igual al valor votado (más o menos 1 %, parámetro) / proyectos votados",
         f"{fm(v(fid, 2024))} (2024, aprox.; {razon(fid, 2024)}), {fila(fid, 2024)['semaforo']}; {fm(v(fid, 2025))} (2025; {razon(fid, 2025)}), {fila(fid, 2025)['semaforo']}; 2026 construible (falta el valor votado)",
         sem_ultimo(fid, A),
         f"Complemento monetario: apropiación total / valor votado total = {fm(v(mon, 2024))} (2024) y {fm(v(mon, 2025))} (2025). "
         "Semáforo con los umbrales de presupuesto participativo (propuesta)",
         lk("Listado EFPC", enlace(fid, 2024)) + "; Hacienda; plan de acción"],
        ["**EPP** Ejecución presupuestal (principal)", "Dinero del presupuesto participativo convertido en obligaciones", "100 × OPP / APP",
         f"{fm(v(epp, 2024))} (2024, aprox.); {fm(v(epp, 2025))} (2025); {fm(v(epp26, 2026))} (1.er semestre de 2026, sin semáforo)",
         sem_ultimo(epp, A), "Sin respaldo propio: lo respaldan los cuatro siguientes", lk("Hacienda, plan de acción", enlace(epp, 2025))],
        ["Tasa de compromiso (respaldo)", "Dinero comprometido", "100 × compromisos / apropiación",
         f"{fm(v(tc, 2025))} (2025); 2024 no disponible (Hacienda no separa compromisos)", sem_ultimo(tc, A), "—", "Hacienda"],
        ["Proyectos bien ejecutados (respaldo)", "Proyectos con ejecución de al menos 90 %", "proyectos con 90 % o más / proyectos",
         f"{razon(p90, 2024)} (2024); {razon(p90, 2025)} (2025)", sem_ultimo(p90, A), "—", "Hacienda"],
        ["Equidad entre zonas (respaldo)", "Brecha de ejecución entre zonas", "zona más alta − zona más baja (puntos)",
         f"{fm(v(eq, 2024))} (2024: {fm(float(fila(eq, 2024)['denominador']))} a {fm(float(fila(eq, 2024)['numerador']))}); "
         f"{fm(v(eq, 2025))} (2025: {fm(float(fila(eq, 2025)['denominador']))} a {fm(float(fila(eq, 2025)['numerador']))})", "Gris", "—", "Hacienda"],
        ["Avance físico (respaldo de entrega)", "Proyectos reportados como terminados", "proyectos con avance físico de 100 % / proyectos",
         f"{razon(af, 2025)} = {fm(v(af, 2025))} (2025, autorreporte)", sem_ultimo(af, A), "—", lk("Plan de acción 2025", enlace(af, 2025))],
    ]
    return filas, dict(principales=[[fid], [epp, epp26]], respaldos=[[tc], [p90], [eq], [af]])


def _cuenta(grupos):
    """(total, con valor) contando cada grupo como un indicador; el grupo con '__politicas_respaldo__' no tiene valor."""
    return len(grupos), sum(1 for g in grupos if con_valor(g))


def documento_1():
    A = {}
    cs, cs_c = ind_cs()
    pc, pc_c = ind_pc()
    pdm, pdm_c = ind_pdm()
    pp, pp_c = ind_pp()
    fid24 = fila("Fidelidad a lo votado", 2024)
    ile23 = [r for r in FILAS if r["indicador"] == "ILE primera infancia - cobertura de atención integral" and r["anio"] == 2023][0]
    m = re.search(r"2023 da ([\d,]+) con 0 a 5 años y ([\d,]+) con 0 a 6 años", ile23["nota"])
    v05, v06 = m.group(1), m.group(2)
    sal25, ms_dic, ms_ago = v("Cobertura de salud", 2025), v("Cobertura de salud, cifra publicada por MinSalud (diciembre)", 2025), v("Cobertura de salud, cifra publicada por MinSalud (agosto)", 2026)
    isp = fila("ISPC - Opción redistribuida (0,556·I2 + 0,444·I3)", 2025)
    cs_t = {k: _cuenta(cs_c[k]) for k in cs_c}
    pc_t = {k: _cuenta(pc_c[k]) for k in pc_c}
    pdm_t = {k: _cuenta(pdm_c[k]) for k in pdm_c}
    pp_t = {k: _cuenta(pp_c[k]) for k in pp_c}
    resumen = [("Control social", cs_t), ("Participación ciudadana", pc_t), ("Cumplimiento del PDM", pdm_t), ("Presupuesto participativo", pp_t)]
    tot_p = sum(t["principales"][0] for _, t in resumen), sum(t["principales"][1] for _, t in resumen)
    tot_r = sum(t["respaldos"][0] for _, t in resumen), sum(t["respaldos"][1] for _, t in resumen)
    tab_res = ["| Componente | Principales | Principales con valor hoy | Respaldos | Respaldos con valor hoy |", "|---|---|---|---|---|"]
    for n, t in resumen:
        tab_res.append(f"| {n} | {t['principales'][0]} | {t['principales'][1]} | {t['respaldos'][0]} | {t['respaldos'][1]} |")
    tab_res.append(f"| **Total** | **{tot_p[0]}** | **{tot_p[1]}** | **{tot_r[0]}** | **{tot_r[1]}** |")
    md = f"""
## Introducción

Tras la reunión con el profesor Francisco Correa se aprobó simplificar el sistema de indicadores. En control social se mantienen los indicadores; en participación ciudadana se eliminan I5 e I6; en el cumplimiento del Plan de Desarrollo (PDM) la cobertura de salud reemplaza al aseguramiento subsidiado y se añade una dimensión de políticas públicas; en presupuesto participativo se mantienen la ejecución presupuestal (EPP) y la fidelidad a lo votado, y salen por ahora la entrega cumplida, la meta de atención y el índice del presupuesto participativo.

**Regla del respaldo.** Cada indicador lleva un indicador de respaldo o una variable proxy: si el principal no se puede calcular hoy, el respaldo ofrece una lectura provisional con los datos que sí existen. El respaldo no reemplaza al principal ni se suma a él. Las cifras de este documento son provisionales: nada está validado y no hubo tiempo de validar con Planeación. Las series se calcularon con documentos públicos consultados hasta el {FECHA}. En las tablas, el semáforo de cada fila corresponde al año más reciente con valor y el de los respaldos se indica en su texto.

## Control social

{tabla_md(cs)}

## Participación ciudadana

{tabla_md(pc)}

*Notas.* El ISPC original ya usa solo I1 a I4; eliminar I5 e I6 no cambia sus pesos. La opción redistribuida (0,556·I2 + 0,444·I3, es decir 0,25/0,45 y 0,20/0,45) se ofrece a los profesores mientras I1 e I4 no sean viables. El respaldo de I1 usa los mismos datos de la fidelidad de presupuesto participativo, por eso no debe sumarse en dos componentes.

## Cumplimiento del Plan de Desarrollo (PDM)

{tabla_md(pdm)}

*Notas.* **Rango de edad (ILE).** El PDM fija la línea base 16,3 en 2023; con la población DANE por edad simple de una misma proyección, 2023 da {v05} con 0 a 5 años y {v06} con 0 a 6 años: ninguno la reproduce con tolerancia de 0,1, y se usa 0 a 5 años (el más cercano). La clasificación de planteles de Saber 11 de Envigado no está en datos abiertos; sin ella el avance hacia la meta de 14 % no se puede calcular y el respaldo es el puntaje promedio (el conjunto público termina en 2022). **Cobertura de salud.** La serie divide los afiliados de cada diciembre entre una sola proyección de población DANE (la de 2018-2042), por eso es comparable entre años: {fm(sal25)} en 2025. La cifra de MinSalud de diciembre de 2025 ({fm(ms_dic)}) y la de agosto de 2026 ({fm(ms_ago)}) difieren porque MinSalud actualizó la población con otra proyección, no por una pérdida de afiliados. **ICPD-6.** Solo se muestra como referencial. **Políticas públicas.** Hay documentos de cuatro políticas, pero no un inventario oficial de las vigentes.

## Presupuesto participativo

{tabla_md(pp)}

*Notas.* **Fidelidad 2024:** se cuentan los {fid24['denominador']} proyectos votados del listado de la EFPC; si solo se cuentan los 50 que se pudieron cruzar con Hacienda, el resultado sería 39/50 = 78,00. La versión monetaria es un complemento de la fidelidad, no un respaldo aparte; por eso en este componente los principales son la fidelidad y la EPP, y los respaldos son los cuatro últimos.

## Resumen

Cuántos indicadores principales y respaldos tiene cada componente, y cuántos tienen valor hoy (al menos un año con cifra, aunque sea referencial).

{chr(10).join(tab_res)}

## Siglas

**CA** espacio público verde. **CD** servicios deportivos. **CM** mortalidad vial. **CP** procesos participativos. **DAP** Departamento Administrativo de Planeación. **DANE** Departamento Administrativo Nacional de Estadística. **DSI** disponibilidad de información de seguimiento. **EFPC** Escuela de Formación para la Participación Ciudadana. **EPP** ejecución del presupuesto participativo. **EPV** espacio público verde. **I1 a I4** indicadores de incorporación de aportes, cobertura territorial, oferta formativa y publicación de resultados. **ICFES** Instituto Colombiano para la Evaluación de la Educación. **ICPD-6** índice de cumplimiento del PDM. **ICRCS** índice de condiciones y respuesta para el control social. **ILE** logro educativo. **Índer** Instituto de Deporte y Recreación de Envigado. **INMLCF** Instituto Nacional de Medicina Legal y Ciencias Forenses. **ISPC** índice de seguimiento de la participación ciudadana. **MinSalud** Ministerio de Salud y Protección Social. **OPP, APP** obligaciones y apropiación definitiva del presupuesto participativo. **PCS** promoción del control social. **PDM** Plan de Desarrollo Municipal 2024-2027. **POT** Plan de Ordenamiento Territorial. **RPO** respuesta a observaciones ciudadanas. **VRC** veedurías en la rendición de cuentas.
"""
    return md.strip() + "\n", dict(resumen=resumen)


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
     "Avance físico (respaldo de entrega)", "Subdirección de Participación Ciudadana; unidades ejecutoras", "Carpeta de PDF con un índice", "Media"),
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
    md1, _ = documento_1()
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
