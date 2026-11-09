#!/usr/bin/env python3
"""Ensambla los parciales por sistema en las tablas finales de outputs/indicadores/.

Entradas : outputs/indicadores/parciales/<SIS>/{valores_variables,resultados_indicadores,soporte_conteos,pp_proyectos}.csv
Salidas  : outputs/indicadores/{valores_variables,resultados_indicadores,pp_proyectos,soporte_conteos,diccionario_datos}.csv
           (separador ';', UTF-8 con BOM) y outputs/indicadores/base_indicadores.xlsx (una hoja por tabla,
           más diccionario_datos y Supuestos).
Los sistemas sin parciales se omiten con aviso. Ejecutar desde la raíz: python3 scripts/indicadores/ensamblar.py
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import comun as C  # noqa: E402

TABLAS = ["valores_variables", "resultados_indicadores", "pp_proyectos", "soporte_conteos"]

# (tipo, descripción) por columna. tipo: texto | entero | decimal | bool | fecha
D = {
 "valores_variables": {
  "sistema": ("texto", "Sistema: Control social / Participación ciudadana / Cumplimiento del PDM / Presupuesto participativo"),
  "codigo_indicador": ("texto", "Indicador al que pertenece la variable (CS.DSI, PDM.CM...)"),
  "id_variable": ("texto", "ID de la Matriz (CS.VRC.ET) o auxiliar AUX.<SIS>.<IND>.<NOMBRE>; siempre el ID completo"),
  "nombre_variable": ("texto", "Nombre completo de la variable (Registro col. D)"),
  "unidad": ("texto", "Conteo, Porcentaje (0–100), COP, Muertes / 100.000 hab., m² / habitante u otra para auxiliares"),
  "tipo_valor": ("texto", "observado, meta, construido, proxy o auxiliar"),
  "anio": ("entero", "Año del dato"), "periodo": ("texto", "Anual, Semestre 1 o Semestre 2"),
  "territorio": ("texto", "Municipio o Zona 1 ... Zona 13"),
  "escenario_id": ("texto", "'defecto' o lista alfabética de desviaciones param=opcion separadas por ';'"),
  "es_defecto": ("bool", "True solo en el escenario por defecto (PROVISIONAL)"),
  "parametros": ("texto", "Juego completo de parámetros aplicables, clave=valor;..."),
  "valor": ("decimal", "Valor sin redondear; vacío si no hay dato"),
  "estado_informacion": ("texto", "Pendiente, No disponible, No evaluable o No aplica. Nunca Validado"),
  "categoria_disponibilidad": ("texto", "Disponible, Construible, Parcial, Proxy o No disponible"),
  "fuente": ("texto", "Publicación o fuente (Registro col. K)"),
  "documento_ruta_local": ("texto", "Ruta bajo data_raw_IIC/ (columna interna)"),
  "enlace": ("texto", "URL de la evidencia (dominio espejo para documentos de la Alcaldía)"),
  "pagina_tabla": ("texto", "Página, tabla, hoja o celda (Registro col. M)"),
  "fecha_corte": ("fecha", "Fecha de corte escrita en el Registro, según corte_fecha_registro"),
  "fecha_corte_real": ("fecha", "Corte real del dato"), "fecha_publicacion": ("fecha", "Fecha interna de publicación del documento"),
  "fecha_consulta": ("fecha", "Fecha de ejecución del script"), "horizonte_meta": ("entero", "Horizonte de la meta (2027)"),
  "n_elementos_soporte": ("entero", "Filas de soporte que suman el valor"), "ids_soporte": ("texto", "IDs de soporte_conteos separados por coma"),
  "provisional": ("bool", "Siempre True mientras existan parámetros PROVISIONAL"),
  "script": ("texto", "Script que generó la fila"), "observaciones": ("texto", "Brechas, discrepancias y alternativas"),
 },
 "resultados_indicadores": {
  "sistema": ("texto", "Sistema del indicador"), "codigo": ("texto", "Código único del indicador o índice"),
  "nombre": ("texto", "Nombre del indicador"), "tipo": ("texto", "Indicador, Complementario o Índice"),
  "anio": ("entero", "Año observado"), "periodo": ("texto", "Anual, Semestre 1 o Semestre 2"),
  "territorio": ("texto", "Municipio o Zona n"), "escenario_id": ("texto", "Ver valores_variables.escenario_id"),
  "es_defecto": ("bool", "Escenario por defecto"), "parametros": ("texto", "Parámetros aplicables clave=valor;..."),
  "entrada_1_id": ("texto", "ID de la entrada 1 (orden de Resultados F:I)"), "entrada_2_id": ("texto", "ID de la entrada 2"),
  "entrada_3_id": ("texto", "ID de la entrada 3"), "entrada_4_id": ("texto", "ID de la entrada 4"),
  "entrada_1": ("decimal", "Valor de la entrada 1 o N/D, Revisar"), "entrada_2": ("decimal", "Valor de la entrada 2"),
  "entrada_3": ("decimal", "Valor de la entrada 3"), "entrada_4": ("decimal", "Valor de la entrada 4"),
  "variables_usadas": ("texto", "IDs de las variables usadas, separados por ';'"),
  "intermedio_1": ("decimal", "LA (ILE) o IBS (ICPD-6)"), "intermedio_2": ("decimal", "LS (ILE) o ITG (ICPD-6)"),
  "puntaje": ("decimal", "Puntaje 0-100 sin redondear (regla preliminar_pendiente); vacío si no se calcula"),
  "puntaje_1d": ("decimal", "Puntaje redondeado a 1 decimal, solo para mostrar"),
  "semaforo": ("texto", "Verde, Amarillo, Naranja, Rojo o Gris (umbrales de la Matriz)"),
  "disponibilidad": ("texto", "Calculado, NE, N/D o Revisar"), "motivo_nd": ("texto", "Por qué no hay puntaje"),
  "puntaje_regla_matriz": ("texto", "Lo que mostraría la hoja Resultados con estados Pendiente (N/D)"),
  "fecha_corte_comun": ("fecha", "Corte común de las entradas"), "estado": ("texto", "Siempre Provisional"),
  "provisional": ("bool", "True"), "alerta": ("texto", "Alertas, p. ej. CEP o CMA < 50 (Guía 20)"),
  "no_calculado_por_etapa": ("bool", "True para indicadores No viables: solo variables"),
  "script": ("texto", "Script que generó la fila"), "observaciones": ("texto", "Notas"),
 },
 "pp_proyectos": {
  "anio": ("entero", "Año"), "periodo": ("texto", "Anual o Semestre 1"), "territorio": ("texto", "Municipio"),
  "zona": ("entero", "Zona del proyecto (2-13)"), "codigo_proyecto": ("texto", "Código único del proyecto (4 dígitos)"),
  "bpin": ("texto", "Código BPIN"), "nombre_proyecto": ("texto", "Nombre del proyecto"),
  "unidad_ejecutora": ("texto", "Dependencia ejecutora"), "en_listado_efpc": ("bool", "Aparece en el listado de la EFPC"),
  "valor_listado_efpc": ("decimal", "Valor del listado EFPC (COP)"),
  "tipo_correspondencia": ("texto", "exacta, aproximada o sin_correspondencia"),
  "apropiacion_definitiva": ("decimal", "APP del proyecto (COP)"), "compromiso": ("decimal", "Compromiso (COP)"),
  "obligacion": ("decimal", "OPP del proyecto (COP)"), "fuente_valor": ("texto", "plan_accion_roca, hacienda o cuipo"),
  "fecha_corte": ("fecha", "Corte registrado"), "fecha_corte_real": ("fecha", "Corte real"),
  "es_0553": ("bool", "Proyecto 0553 (costo del proceso)"), "incluido_epp_defecto": ("bool", "Entra en el EPP del escenario por defecto"),
  "ejecucion_fisica_neta_dic": ("decimal", "Ejecución física neta a diciembre"),
  "pec_todas_filas_dic_100": ("entero", "Proxy PEC regla todas_filas_dic_100 (0/1)"),
  "pec_solo_31dic": ("entero", "Proxy PEC regla solo_31dic (0/1)"), "pec_ultima_fila": ("entero", "Proxy PEC regla ultima_fila_por_indicador (0/1)"),
  "tipologia_sugerida": ("texto", "servicio a personas, infraestructura, medios u otro"),
  "aplica_meta_atencion": ("texto", "Sí, No o Pendiente"), "motivo_exclusion": ("texto", "Motivo de exclusión"),
  "personas_atendidas_PA": ("entero", "Personas atendidas (vacío; sin datos personales)"),
  "meta_atencion_MP": ("entero", "Meta de atención (vacío)"), "estado_informacion": ("texto", "Pendiente u otro; nunca Validado"),
  "fuente": ("texto", "Fuente del registro"), "enlace": ("texto", "Enlace a la evidencia"),
  "pagina_hoja": ("texto", "Página u hoja"), "observaciones": ("texto", "Notas"),
 },
 "soporte_conteos": {
  "id_soporte": ("texto", "<SIS>-<IND>-<aaaa>-<nnn>"), "sistema": ("texto", "Sistema"),
  "codigo_indicador": ("texto", "Indicador"), "id_variable": ("texto", "Variable que alimenta el elemento"),
  "anio": ("entero", "Año"), "periodo": ("texto", "Periodo"), "territorio": ("texto", "Territorio"),
  "escenario_id": ("texto", "Escenario al que se refiere la fila"),
  "elemento_tipo": ("texto", "evento, observacion, meta, accion_formativa, sesion, zona_actividad, proceso, proyecto, actividad_virtual"),
  "elemento_id": ("texto", "Identificador estable del elemento"),
  "descripcion": ("texto", "Asunto resumido, sin datos personales"), "fecha_elemento": ("fecha", "Fecha del elemento"),
  "documento": ("texto", "Documento fuente"), "ruta_local": ("texto", "Ruta local bajo data_raw_IIC/"),
  "enlace": ("texto", "URL (espejo)"), "pagina": ("texto", "Página o tabla"),
  "fecha_publicacion_doc": ("fecha", "Fecha de publicación verificada del documento"),
  "banderas": ("texto", "clave=valor;... atributos que usan los parámetros"),
  "clasificacion": ("texto", "Etiqueta legible de la clasificación"), "justificacion": ("texto", "Por qué se clasificó así"),
  "criterio": ("texto", "Criterio aplicado para incluir o excluir"), "lector": ("texto", "Quién clasificó"),
  "cuenta_en_defecto": ("entero", "1 si entra en el conteo del escenario por defecto"),
  "incluido": ("texto", "sí / no, equivalente legible de cuenta_en_defecto"),
 },
}
EXTRA_DICC = [
 ("matriz_diligenciada.Registro", "A:Q", "texto", "no", "Sistema, indicador, ID, nombre, unidad (A:E); año, periodo, territorio (F:H); valor (I); estado (J); fuente (K); enlace (L); página (M); corte (N); consulta (O); horizonte (P); observaciones (Q). R:S son fórmulas y no se tocan."),
 ("matriz_diligenciada.PP_Proyectos", "A:M", "texto", "no", "Año, periodo, territorio, código, nombre, ¿aplica meta?, PA, MP, fuente, enlace, corte, estado y observaciones/soporte. N:P son fórmulas y no se tocan."),
]
HEAD_DICC = ["tabla", "columna", "tipo", "obligatoria", "descripcion"]


def construir_diccionario() -> list[dict]:
    filas = []
    for t in TABLAS:
        oblig = set(C.OBLIGATORIAS[t])
        for col in C.COLUMNAS[t]:
            tipo, desc = D[t][col]
            filas.append({"tabla": t, "columna": col, "tipo": tipo, "obligatoria": "sí" if col in oblig else "no",
                          "descripcion": desc})
    for t, col, tipo, ob, desc in EXTRA_DICC:
        filas.append({"tabla": t, "columna": col, "tipo": tipo, "obligatoria": ob, "descripcion": desc})
    return filas


def _tipado(tabla: str, col: str, v: str):
    """Convierte el texto del CSV al tipo nativo para el xlsx."""
    if v == "":
        return None
    tipo = D[tabla][col][0]
    try:
        if tipo == "entero":
            return int(float(v))
        if tipo == "decimal":
            return float(v)
        if tipo == "bool":
            return v == "True"
        if tipo == "fecha":
            return dt.date.fromisoformat(v)
    except ValueError:
        return v        # N/D, Revisar, etc.
    return v


def leer_parciales(tabla: str) -> tuple[list[dict], list[str]]:
    filas, usados = [], []
    for sis in C.SISTEMAS:
        p = C.ruta_parcial(sis, tabla)
        if p.exists():
            r = C.leer_csv(p)
            filas += r
            usados.append(f"{sis}({len(r)})")
    return filas, usados


def main() -> int:
    C.DIR_SALIDA.mkdir(parents=True, exist_ok=True)
    existentes = [s for s in C.SISTEMAS if (C.DIR_PARCIALES / s).is_dir()]
    for s in C.SISTEMAS:
        if s not in existentes:
            print(f"AVISO: no hay parciales de {s} ({C.DIR_PARCIALES / s}); se omite")
    tablas = {}
    for t in TABLAS:
        filas, usados = leer_parciales(t)
        # llaves únicas entre sistemas
        llave = C.LLAVES[t]
        vistos = {}
        for i, f in enumerate(filas):
            k = tuple(f.get(c, "") for c in llave)
            if k in vistos:
                raise SystemExit(f"ERROR {t}: llave duplicada {k}")
            vistos[k] = i
        for f in filas:
            for c in ("estado_informacion", "estado"):
                if f.get(c) == "Validado":
                    raise SystemExit(f"ERROR {t}: estado Validado no permitido")
        tablas[t] = filas
        C.escribir_csv(C.DIR_SALIDA / f"{t}.csv", filas, C.COLUMNAS[t])
        print(f"{t}: {len(filas)} filas ({', '.join(usados) or 'sin parciales'})")
    dicc = construir_diccionario()
    C.escribir_csv(C.DIR_SALIDA / "diccionario_datos.csv", dicc, HEAD_DICC)
    print(f"diccionario_datos: {len(dicc)} filas")

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    wb.remove(wb.active)
    cab = PatternFill("solid", fgColor="1F4E78")

    def hoja(nombre, cols, filas, tipado=None):
        ws = wb.create_sheet(nombre[:31])
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = cab
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for f in filas:
            fila = [tipado(c, f.get(c, "")) if tipado else f.get(c) for c in cols]
            ws.append(fila)
        for j, c in enumerate(cols, 1):
            ancho = min(60, max(len(c), *(len(str(f.get(c, ""))) for f in filas[:200])) + 2) if filas else len(c) + 2
            ws.column_dimensions[get_column_letter(j)].width = ancho
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                for cell in row:
                    if isinstance(cell.value, dt.date):
                        cell.number_format = "dd/mm/yyyy"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    leeme = wb.create_sheet("Léeme")
    for linea in ["Base de indicadores del Observatorio Ciudadano de Envigado (OCE) - ESTADO: PROVISIONAL.",
                  "Todos los valores están en estado Pendiente y los resultados son Provisionales; nada está Validado.",
                  "Cada fila lleva el escenario de supuestos con que se calculó (hoja Supuestos y config/supuestos.yaml).",
                  "Fechas dd/mm/aaaa; puntajes en escala 0-100; montos en COP.",
                  f"Generado: {dt.date.today().isoformat()} por scripts/indicadores/ensamblar.py."]:
        leeme.append([linea])
    leeme.column_dimensions["A"].width = 110
    for t in TABLAS:
        hoja(t, C.COLUMNAS[t], tablas[t], lambda c, v, t=t: _tipado(t, c, v) if isinstance(v, str) else v)
    hoja("diccionario_datos", HEAD_DICC, dicc)
    sup = C.cargar_supuestos()
    filas_sup = []
    for k, p in sup["parametros"].items():
        filas_sup.append({"id": k, "descripcion": p["descripcion"], "defecto": str(p["defecto"]),
                          "opciones": " | ".join(f"{o}: {d}" for o, d in p["opciones"].items()),
                          "estado": p["estado"], "afecta": ", ".join(p["afecta"]), "decision": p["decision"]})
    hoja("Supuestos", ["id", "descripcion", "defecto", "opciones", "estado", "afecta", "decision"], filas_sup)
    destino = C.DIR_SALIDA / "base_indicadores.xlsx"
    wb.save(destino)
    print(f"escrito {destino}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
