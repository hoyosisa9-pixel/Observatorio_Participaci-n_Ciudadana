"""Excel de presentación del sistema de indicadores simplificado (una hoja por dimensión).

Lee (sin recalcular):
  - outputs/sistema_simplificado/sistema_indicadores_simplificado.xlsx, hoja "Indicadores por dimensión"
    (nombres, fórmulas, limitaciones, valores 2024-2026, fuente y enlace ya revisados);
  - outputs/sistema_simplificado/indicadores_simplificados.csv (semáforo por año y series 2019-2025).
Escribe outputs/sistema_simplificado/indicadores_OCE_presentacion.xlsx. Todo es PROVISIONAL.
Ejecutar desde la raíz: python3 scripts/sistema_simplificado/generar_excel_presentacion.py
"""
import csv
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parents[2]
D = RAIZ / "outputs/sistema_simplificado"
F_XLSX = D / "sistema_indicadores_simplificado.xlsx"
F_CSV = D / "indicadores_simplificados.csv"
SAL = D / "indicadores_OCE_presentacion.xlsx"

AZUL, AZUL_CLARO, GRIS_TXT = "1F4E78", "DDEBF7", "7F7F7F"
SEM = {"Verde": "C6EFCE", "Amarillo": "FFEB9C", "Naranja": "F8CBAD", "Rojo": "FFC7CE"}
SEM_TXT = {"Verde": "006100", "Amarillo": "7F6000", "Naranja": "833C0B", "Rojo": "9C0006"}
DIMS = [  # (nombre en la hoja fuente, nombre de pestaña, componente en el CSV, descripción)
    ("Control social (CS)", "Control social", "Control social",
     "Mide si la ciudadanía tiene información y respuestas para vigilar la gestión pública: promoción del control social, "
     "información de seguimiento publicada, veedurías en la rendición de cuentas y respuesta a observaciones."),
    ("Participación ciudadana (PC)", "Participación ciudadana", "Participación ciudadana",
     "Mide si los espacios de participación llegan a todo el territorio, si la oferta formativa se cumple y si los aportes "
     "ciudadanos se incorporan y sus resultados se publican. Se eliminaron I5 (sesiones de instancias) e I6 (canal virtual)."),
    ("Cumplimiento del Plan de Desarrollo Municipal", "Plan de Desarrollo", "Cumplimiento del PDM",
     "Mide el avance de resultados seleccionados del Plan de Desarrollo 2024-2027 (PDM): salud, educación, mortalidad vial, "
     "espacio público verde, procesos participativos, deporte y la nueva dimensión de logros de políticas públicas."),
    ("Presupuesto participativo (PP)", "Presupuesto participativo", "Presupuesto participativo",
     "Mide si los proyectos que la comunidad votó recibieron el presupuesto votado, si se ejecutaron y si se terminaron."),
]
FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)


def n(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def cabecera(ws, fila, textos, col0=1):
    for j, t in enumerate(textos, col0):
        c = ws.cell(fila, j, t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDE
    ws.row_dimensions[fila].height = 30


def banda(ws, fila, texto, ncols, tam=16, fondo=AZUL, color="FFFFFF", alto=30):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, texto)
    c.font = Font(bold=True, size=tam, color=color)
    c.fill = PatternFill("solid", fgColor=fondo)
    c.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[fila].height = alto


def texto(ws, fila, t, ncols, alto=32, negrita=False, color="000000", fondo=None):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, t)
    c.font = Font(bold=negrita, color=color)
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    if fondo:
        c.fill = PatternFill("solid", fgColor=fondo)
    ws.row_dimensions[fila].height = alto


def valor(c, v, sem):
    """Escribe un valor con formato y color de semáforo."""
    f = n(v)
    if v in (None, ""):
        c.value = "N/D"
        c.font = Font(color=GRIS_TXT, italic=True)
    elif f is None:
        c.value = v
    else:
        c.value = f
        c.number_format = "#,##0" if f == int(f) else "#,##0.00"
        if sem in SEM:
            c.fill = PatternFill("solid", fgColor=SEM[sem])
            c.font = Font(bold=True, color=SEM_TXT[sem])
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDE


def leyenda(ws, fila):
    ws.cell(fila, 1, "Semáforo:").font = Font(bold=True)
    for j, (s, d) in enumerate([("Verde", "Cumple o va bien"), ("Amarillo", "Alerta"), ("Rojo", "Bajo"),
                                ("N/D", "Sin dato")], 2):
        c = ws.cell(fila, j, f"{s}: {d}")
        if s in SEM:
            c.fill = PatternFill("solid", fgColor=SEM[s])
            c.font = Font(color=SEM_TXT[s], bold=True)
        else:
            c.font = Font(color=GRIS_TXT, italic=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.cell(fila + 1, 1, "Sin color: el indicador no tiene semáforo definido (conteos, series o corte parcial).").font = \
        Font(italic=True, color=GRIS_TXT, size=9)


def main():
    src = openpyxl.load_workbook(F_XLSX)["Indicadores por dimensión"]
    hdr = [c.value for c in src[1]]
    ind = [dict(zip(hdr, r)) for r in src.iter_rows(min_row=2, values_only=True) if r[0]]
    for r in ind:  # nombres oficiales de la Matriz
        r["Indicador"] = r["Indicador"].replace("Índice de seguimiento de la participación ciudadana",
                                                "Índice Simplificado de Participación Ciudadana")
    csvf = list(csv.DictReader(open(F_CSV, encoding="utf-8-sig"), delimiter=";"))
    assert not any(r.get("estado") == "Validado" for r in csvf)

    def sem_anio(comp, anio, v):
        """Semáforo del año tomado del CSV: misma dimensión, año y valor."""
        f = n(v)
        if f is None:
            return ""
        for r in csvf:
            if r["componente"] == comp and r["anio"] == str(anio) and n(r["valor"]) is not None \
                    and abs(n(r["valor"]) - f) < 0.006:
                if "sin semáforo" in r["nota"].lower():
                    return ""
                return r["semaforo"] if r["semaforo"] in SEM else ""
        return ""

    wb = openpyxl.Workbook()
    # ------------------------------------------------------------------ Resumen
    ws = wb.active
    ws.title = "Resumen"
    NC = 7
    banda(ws, 1, "Sistema de indicadores del Observatorio Ciudadano de Envigado (OCE)", NC, 18, alto=36)
    banda(ws, 2, "Versión simplificada – resultados provisionales – 10 de noviembre de 2026", NC, 11, AZUL_CLARO, AZUL, 22)
    texto(ws, 3, "PROVISIONAL: ninguna cifra está validada ni confirmada por la Alcaldía. Los valores salen de documentos "
                 "públicos consultados hasta el 10 de noviembre de 2026.", NC, 30, True, "C00000", "FCE4E4")
    texto(ws, 4, "Cómo leer este libro: esta hoja muestra los indicadores principales de las cuatro dimensiones. Cada "
                 "dimensión tiene su hoja con la fórmula, la lectura de cada resultado y sus indicadores de respaldo (un "
                 "respaldo da una lectura cuando el principal no se puede calcular). La hoja «Series en el tiempo» muestra "
                 "la evolución 2019-2025 de salud y educación.", NC, 58)
    leyenda(ws, 6)
    f0 = 9
    cabecera(ws, f0, ["Dimensión", "Indicador principal", "2024", "2025", "2026", "Lectura", "Hoja"])
    fila = f0 + 1
    cuenta = {}
    for dsrc, pest, comp, _ in DIMS:
        prin = [r for r in ind if r["Dimensión"].startswith(dsrc) and r["Tipo"] == "Principal"]
        cuenta[pest] = (len(prin), sum(1 for r in prin if r["Con valor hoy"] == "Sí"),
                        sum(1 for r in ind if r["Dimensión"].startswith(dsrc) and r["Tipo"] != "Principal"))
        ini = fila
        for r in prin:
            ws.cell(fila, 2, r["Indicador"])
            for j, a in enumerate((2024, 2025, 2026), 3):
                valor(ws.cell(fila, j), r[f"Valor {a}"], sem_anio(comp, a, r[f"Valor {a}"]))
            ws.cell(fila, 6, r["Limitación"])
            c = ws.cell(fila, 7, f"Ver «{pest}»")
            c.hyperlink = f"#'{pest}'!A1"
            c.font = Font(color="0563C1", underline="single")
            for j in (2, 6, 7):
                ws.cell(fila, j).alignment = Alignment(wrap_text=True, vertical="center")
            for j in range(1, NC + 1):
                ws.cell(fila, j).border = BORDE
            fila += 1
        ws.merge_cells(start_row=ini, start_column=1, end_row=fila - 1, end_column=1)
        c = ws.cell(ini, 1, pest)
        c.font = Font(bold=True, color=AZUL)
        c.fill = PatternFill("solid", fgColor=AZUL_CLARO)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    fila += 1
    banda(ws, fila, "Cuántos indicadores hay", NC, 12, AZUL_CLARO, AZUL, 22)
    fila += 1
    cabecera(ws, fila, ["Dimensión", "Indicadores principales", "Con valor hoy", "Respaldos y complementos"])
    for pest, (p, cv, r) in cuenta.items():
        fila += 1
        for j, v in enumerate((pest, p, cv, r), 1):
            c = ws.cell(fila, j, v)
            c.border = BORDE
            c.alignment = Alignment(horizontal="center" if j > 1 else "left", vertical="center")
    fila += 1
    tot = [sum(x[i] for x in cuenta.values()) for i in range(3)]
    for j, v in enumerate(("Total", *tot), 1):
        c = ws.cell(fila, j, v)
        c.font = Font(bold=True)
        c.border = BORDE
        c.alignment = Alignment(horizontal="center" if j > 1 else "left")
    for j, w in enumerate([22, 46, 11, 11, 11, 52, 18], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f"A{f0 + 1}"

    # ------------------------------------------------------------------ Hojas por dimensión
    for dsrc, pest, comp, desc in DIMS:
        ws = wb.create_sheet(pest)
        NC = 7
        banda(ws, 1, pest, NC, 18, alto=34)
        texto(ws, 2, desc, NC, 44)
        texto(ws, 3, "Provisional: nada está validado. N/D = sin dato.", NC, 18, False, "C00000")
        fila = 5
        for titulo, filtro in (("Indicadores principales", lambda t: t == "Principal"),
                               ("Respaldos y complementos (se usan cuando el principal no se puede calcular)",
                                lambda t: t != "Principal")):
            filas = [r for r in ind if r["Dimensión"].startswith(dsrc) and filtro(r["Tipo"])]
            banda(ws, fila, titulo, NC, 12, AZUL_CLARO, AZUL, 22)
            fila += 1
            cabecera(ws, fila, ["Indicador", "Cómo se calcula", "2024", "2025", "2026", "Lectura y limitación", "Fuente"])
            fila += 1
            for r in filas:
                es_ind = any(k in r["Indicador"] for k in ("ICRCS", "ISPC", "ICPD"))
                nombre = r["Indicador"] + ("  (complemento)" if r["Tipo"] == "Complemento" else "")
                ws.cell(fila, 1, nombre).font = Font(bold=True, color=AZUL if es_ind else "000000")
                ws.cell(fila, 2, r["Fórmula"])
                for j, a in enumerate((2024, 2025, 2026), 3):
                    valor(ws.cell(fila, j), r[f"Valor {a}"], sem_anio(comp, a, r[f"Valor {a}"]))
                    # años sin fila en la fuente: celda vacía en vez de N/D si el indicador no aplica ese año
                    if r[f"Valor {a}"] in (None, "") and str(a) not in str(r["Valor por año"]):
                        ws.cell(fila, j).value = "—"
                ws.cell(fila, 6, r["Limitación"])
                c = ws.cell(fila, 7, r["Fuente"])
                if r.get("Enlace"):
                    c.hyperlink = r["Enlace"]
                    c.font = Font(color="0563C1", underline="single")
                for j in (1, 2, 6, 7):
                    ws.cell(fila, j).alignment = Alignment(wrap_text=True, vertical="center")
                for j in range(1, NC + 1):
                    ws.cell(fila, j).border = BORDE
                if es_ind:
                    for j in (1, 2, 6, 7):
                        ws.cell(fila, j).fill = PatternFill("solid", fgColor="F2F2F2")
                ws.row_dimensions[fila].height = 45
                fila += 1
            fila += 1
        leyenda(ws, fila)
        for j, w in enumerate([38, 44, 11, 11, 11, 46, 30], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ------------------------------------------------------------------ Series en el tiempo
    ws = wb.create_sheet("Series en el tiempo")
    anios = [str(a) for a in range(2019, 2026)]
    NC = 2 + len(anios)
    banda(ws, 1, "Series en el tiempo (evolución 2019-2025)", NC, 18, alto=34)
    texto(ws, 2, "Indicadores que el profesor pidió mirar como evolución en años. N/D = el dato público no existe para ese año.",
          NC, 30)
    series = [
        ("Plan de Desarrollo", "Cobertura de salud (%)", "Cobertura de salud"),
        ("Plan de Desarrollo", "Primera infancia: niños de 0 a 5 años con atención integral (%)",
         "ILE primera infancia - cobertura de atención integral"),
        ("Plan de Desarrollo", "Primera infancia: avance hacia la meta 2027 (0 a 100)",
         "ILE primera infancia - avance hacia la meta 2027"),
        ("Plan de Desarrollo", "Primera infancia: variación anual (puntos)",
         "ILE primera infancia - variación anual (puntos porcentuales)"),
        ("Plan de Desarrollo", "Saber 11: puntaje global promedio, colegios oficiales urbanos", "ILE Saber 11 - puntaje global promedio, instituciones o"),
        ("Plan de Desarrollo", "Saber 11: puntaje global promedio, todos los colegios", "ILE Saber 11 - puntaje global promedio, todos los plant"),
    ]
    cabecera(ws, 4, ["Dimensión", "Serie"] + anios)
    fila = 5
    for dim, rot, clave in series:
        regs = {r["anio"]: r for r in csvf if (r["indicador"] == clave if clave == "Cobertura de salud"
                                               else r["indicador"].startswith(clave))}
        ws.cell(fila, 1, dim)
        ws.cell(fila, 2, rot).font = Font(bold=True)
        for j, a in enumerate(anios, 3):
            c = ws.cell(fila, j)
            if a in regs:
                valor(c, regs[a]["valor"], "")
            else:
                c.value = "—"
                c.alignment = Alignment(horizontal="center")
                c.border = BORDE
        for j in (1, 2):
            ws.cell(fila, j).alignment = Alignment(wrap_text=True, vertical="center")
            ws.cell(fila, j).border = BORDE
        ws.row_dimensions[fila].height = 32
        fila += 1
    fila += 1
    for t in ("Cobertura de salud = afiliados al sistema de salud (contributivo + subsidiado + especiales) / población DANE, "
              "con una sola proyección de población para todos los años. Fuente: TerriData (MinSalud).",
              "Primera infancia: atendidos según TerriData (2019-2023) e informes de gestión (2024-2025); población DANE de "
              "0 a 5 años. Avance = 100 × (valor − 16,3) / (19 − 16,3), entre 0 y 100.",
              "Saber 11: datos abiertos del ICFES; el conjunto público solo llega a 2022 (falta 2021)."):
        texto(ws, fila, t, NC, 32, False, GRIS_TXT)
        fila += 1
    for j, w in enumerate([20, 52] + [10] * len(anios), 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.print_options.horizontalCentered = True
    wb.save(SAL)
    print(f"Escrito {SAL}")


if __name__ == "__main__":
    main()
