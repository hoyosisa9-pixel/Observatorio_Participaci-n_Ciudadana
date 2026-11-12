"""Excel con una hoja por dimensión: indicadores e índices calculados del sistema simplificado.

Lee outputs/sistema_simplificado/indicadores_simplificados.csv (no recalcula nada) y escribe
outputs/sistema_simplificado/indicadores_por_dimension.xlsx. Todo es PROVISIONAL.
Ejecutar desde la raíz: python3 scripts/sistema_simplificado/generar_excel_dimensiones.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

RAIZ = Path(__file__).resolve().parents[2]
ENT = RAIZ / "outputs/sistema_simplificado/indicadores_simplificados.csv"
SAL = RAIZ / "outputs/sistema_simplificado/indicadores_por_dimension.xlsx"

DIMS = [("Control social", "Control social"), ("Participación ciudadana", "Participación ciudadana"),
        ("Cumplimiento del PDM", "Plan de Desarrollo"), ("Presupuesto participativo", "Presupuesto participativo")]
COLOR = {"Verde": "C6EFCE", "Amarillo": "FFEB9C", "Naranja": "F8CBAD", "Rojo": "FFC7CE", "Gris": "D9D9D9"}
TIPO = {"principal": "Principal", "respaldo": "Respaldo", "complemento": "Complemento", "referencial": "Referencial"}
CAB = PatternFill("solid", fgColor="1F4E78")
BORDE = Border(bottom=Side(style="thin", color="BFBFBF"))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def fmt(v):
    f = num(v)
    if f is None:
        return v or ""
    return f"{f:,.0f}".replace(",", ".") if f == int(f) and abs(f) >= 1000 else (
        f"{f:.2f}".rstrip("0").rstrip(".").replace(".", ","))


def es_indice(nombre):
    n = nombre.upper()
    return any(k in n for k in ("ICRCS", "ISPC", "ICPD"))


def main():
    filas = list(csv.DictReader(open(ENT, encoding="utf-8-sig"), delimiter=";"))
    for f in filas:  # una sola fila por indicador: el corte semestral de 2026 se indica en la nota
        if f["indicador"].endswith(" (1.er semestre)"):
            f["indicador"] = f["indicador"][: -len(" (1.er semestre)")]
            f["nota"] = "2026: primer semestre (corte 30 de junio), sin semáforo. " + f["nota"]
    assert not any(f.get("estado") == "Validado" for f in filas)
    anios = sorted({f["anio"] for f in filas if f["anio"]})
    wb = Workbook()
    le = wb.active
    le.title = "Léeme"
    textos = [
        ("Indicadores e índices por dimensión – Observatorio Ciudadano de Envigado (OCE)", True),
        ("PROVISIONAL: ninguna cifra está validada ni confirmada por la Alcaldía.", True),
        ("Una hoja por dimensión: control social, participación ciudadana, Plan de Desarrollo y presupuesto participativo.", False),
        ("Cada fila es un indicador o índice; las columnas de años muestran su valor, coloreado según el semáforo "
         "(verde, amarillo, naranja, rojo; gris = sin dato o N/D; sin color = sin semáforo definido).", False),
        ("Tipo: Principal = indicador del sistema; Respaldo = indicador o variable proxy que da una lectura cuando el "
         "principal no se puede calcular; Complemento = otra forma de leer el principal; Referencial = solo de referencia.", False),
        ("Los índices (ICRCS, ISPC, ICPD-6) van al final de cada hoja, resaltados.", False),
        ("Debajo de cada tabla está el detalle por año: numerador, denominador, fuente y nota.", False),
        ("Corte de la información: documentos públicos consultados hasta el 10 de noviembre de 2026.", False),
    ]
    for t, b in textos:
        le.append([t])
        le.cell(le.max_row, 1).font = Font(bold=b, color="C00000" if "PROVISIONAL" in t else "000000")
        le.cell(le.max_row, 1).alignment = Alignment(wrap_text=True)
    le.column_dimensions["A"].width = 120

    for comp, hoja in DIMS:
        ws = wb.create_sheet(hoja)
        sub = [f for f in filas if f["componente"] == comp]
        orden, vistos = [], set()
        for f in sub:
            k = (f["indicador"], f["tipo"])
            if k not in vistos:
                vistos.add(k)
                orden.append(k)
        orden.sort(key=lambda k: (es_indice(k[0]), ["principal", "complemento", "respaldo", "referencial"].index(k[1])
                                  if k[1] in TIPO else 9))
        ws.append([f"{hoja}: indicadores e índices (provisional)"])
        ws.cell(1, 1).font = Font(bold=True, size=13)
        cab = ["Indicador o índice", "Tipo"] + anios + ["Fuente", "Nota"]
        ws.append(cab)
        for c in ws[2]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = CAB
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        for k in orden:
            reg = {f["anio"]: f for f in sub if (f["indicador"], f["tipo"]) == k}
            ultimo = reg[max(a for a in reg if not (a == "2026" and "primer semestre" in reg[a]["nota"]))] if len(reg) > 1 else reg[max(reg)]
            fila = [k[0], TIPO.get(k[1], k[1])] + [fmt(reg[a]["valor"]) if a in reg and reg[a]["valor"] else
                                                     ("N/D" if a in reg else "") for a in anios]
            fila += [ultimo["fuente"], ultimo["nota"]]
            ws.append(fila)
            r = ws.max_row
            for j, a in enumerate(anios, 3):
                if a in reg:
                    s = reg[a]["semaforo"] or ("Gris" if not reg[a]["valor"] else "")
                    if reg[a]["valor"] and "sin semáforo" in reg[a]["nota"].lower():
                        s = ""  # p. ej. EPP 2026 primer semestre: puntaje informativo, sin color
                    if s in COLOR:
                        ws.cell(r, j).fill = PatternFill("solid", fgColor=COLOR[s])
                ws.cell(r, j).alignment = Alignment(horizontal="center", vertical="top")
            if es_indice(k[0]):
                ws.cell(r, 1).font = Font(bold=True, color="1F4E78")
            if ultimo.get("enlace"):
                ws.cell(r, len(cab) - 1).hyperlink = ultimo["enlace"]
                ws.cell(r, len(cab) - 1).font = Font(color="0563C1", underline="single")
            for c in ws[r]:
                c.border = BORDE
                if c.column in (1, 2, len(cab) - 1, len(cab)):
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        fin_tabla = ws.max_row
        ws.auto_filter.ref = f"A2:{ws.cell(2, len(cab)).column_letter}{fin_tabla}"
        ws.freeze_panes = "C3"
        # Detalle por año
        ws.append([])
        ws.append(["Detalle por año"])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
        dcab = ["Indicador o índice", "Tipo", "Año", "Numerador", "Denominador", "Valor", "Semáforo", "Estado", "Fuente", "Nota"]
        ws.append(dcab)
        for c in ws[ws.max_row]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = CAB
        for k in orden:
            for f in sorted((f for f in sub if (f["indicador"], f["tipo"]) == k), key=lambda f: f["anio"]):
                ws.append([k[0], TIPO.get(k[1], k[1]), f["anio"], fmt(f["numerador"]), fmt(f["denominador"]),
                           fmt(f["valor"]) or "N/D", f["semaforo"] or "Sin semáforo", f["estado"], f["fuente"], f["nota"]])
                r = ws.max_row
                sin_sem = f["valor"] and "sin semáforo" in f["nota"].lower()
                if sin_sem:
                    ws.cell(r, 7).value = "Sin semáforo"
                elif f["semaforo"] in COLOR:
                    ws.cell(r, 7).fill = PatternFill("solid", fgColor=COLOR[f["semaforo"]])
                if f.get("enlace"):
                    ws.cell(r, 9).hyperlink = f["enlace"]
                    ws.cell(r, 9).font = Font(color="0563C1", underline="single")
                for c in ws[r]:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        anchos = [42, 13] + [10] * len(anios) + [34, 70]
        for j, w in enumerate(anchos, 1):
            ws.column_dimensions[ws.cell(2, j).column_letter].width = w
    wb.save(SAL)
    print(f"Escrito {SAL} ({len(filas)} filas de origen; años {anios[0]}-{anios[-1]})")


if __name__ == "__main__":
    main()
