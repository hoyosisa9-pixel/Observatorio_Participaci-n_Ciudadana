# Mapa de búsqueda – Disponibilidad y brecha de información del sistema de indicadores (Fase 1)

**Observatorio Ciudadano de Envigado (OCE)** · Elaborado por el Agente Analizador · Fecha: 14-oct-2026
**Propósito:** guiar a los ejecutores en la verificación, variable por variable, de si el dato existe, dónde, en qué formato y cuál es la brecha. **No se calcula ningún indicador ni índice.** El trabajo previo sobre el IIC solo se usa como insumo (inventario de fuentes, llaves, hallazgos CUIPO/SECOP).

**Periodo objetivo:** vigencias 2024, 2025 y lo publicado de 2026 (corte de descarga del inventario: 25-sep-2026). **Territorio:** Municipio de Envigado (se anota cuando hay desagregación por zona; Envigado se organiza en zonas, no en comunas).

**Insumos leídos**
- `Documentos guias/Matriz de informacion de indicadores del observatorio ciudadano.xlsx` (solo lectura; copia byte a byte en `outputs/disponibilidad/trabajo/matriz_copia.xlsx`, verificada con `cmp`).
- Volcados para ejecutores: `outputs/disponibilidad/trabajo/registro.csv` (156 filas de referencia, columnas A:Q), `outputs/disponibilidad/trabajo/matriz.csv` (23 filas + encabezado), `outputs/disponibilidad/trabajo/guia.txt`.
- `.memory/knowledge_graph.md`, `inventario_IIC_relaciones.md`, `outputs/perfil_llaves_iic.md`, `outputs/limpieza_bpin_cuipo.md`, `outputs/puente_codigo_proyecto_bpin.csv`.
- `data_raw_IIC/inventario_detallado_IIC.csv` (2.787 filas). Inspección breve (solo encabezados o páginas puntuales) de ~15 archivos candidatos clave, señalados abajo como **[inspeccionado]**.

---

## 0. Escala de clasificación (usar literalmente)

| Categoría | Criterio operativo para esta tarea |
|---|---|
| **Disponible** | El valor de la variable, con la definición de la matriz, está publicado por una fuente oficial para el año, y puede citarse directamente (documento + página/tabla/celda). |
| **Construible** | No está publicado como cifra, pero puede obtenerse **con la misma definición** contando, filtrando o sumando registros de fuentes oficiales disponibles (repo o web), sin supuestos externos fuertes. |
| **Parcial** | Existe dato con la definición correcta, pero solo para parte del periodo (p. ej. 2024 sí, 2025 no), del universo (algunas instancias, algunos proyectos, algunas zonas) o sin el soporte exigido por el protocolo. |
| **Proxy** | Solo existe un dato de definición, universo, unidad o fuente distinta que aproxima la variable (debe declararse la diferencia; no sustituye el dato). |
| **No disponible** | No se identificó fuente pública ni registro primario que permita obtener o construir el dato en el periodo. |

Esta escala **no** es la columna "Estado de información" de la hoja Registro (cuya lista es Pendiente / Validado / No disponible / No evaluable / No aplica). Ver sección 2.

---

## 1. Verificación de conteos de la matriz

| Control | Esperado | Encontrado | Resultado |
|---|---|---|---|
| Indicadores en hoja Matriz (Tipo = Indicador o Complementario) | 19 | 19 (CS 4, PC 6 [I5 e I6 complementarios], PDM 6, PP 3) | Cuadra |
| Índices en hoja Matriz (Tipo = Índice) | 4 | 4 (ICRCS, ISPC, ICPD-6, IREPP) | Cuadra |
| Total filas de catálogo | 23 | 23 | Cuadra |
| Variables únicas (ID) en hoja Registro | 39 | 39 (CS 8, PC 12, PDM 14, PP 5) | Cuadra |
| Filas de referencia en Registro | 156 = 39 × 4 años (2024–2027), Anual, Municipio | 156 | Cuadra |
| Metas precargadas | 7 valores | 7 IDs con valor (28 filas: 7 × 4 años), estado Pendiente, horizonte 2027 | Cuadra |
| Filas capturadas en PP_Proyectos | 0 (sin datos precargados) | 0 (filas 6–305 solo con fórmulas) | Cuadra |

**Observación:** las fórmulas de la Matriz usan **42 símbolos de base**, no 39. La diferencia es de diseño: en PP.CMA, **PA** y **MP** se capturan por proyecto en la hoja `PP_Proyectos` (no en Registro) y **SC** se calcula automáticamente. Solo `PP.CMA.NA` va en Registro. Los ejecutores deben mapear también PA y MP (ver 6.4.6).

---

## 2. Reglas de la Guía relevantes para el registro (resumen operativo)

Fuente: `guia.txt` (hoja Guía, filas 5–28). Numeración de la Guía entre paréntesis.

1. **Estructura (2, 5):** una fila por variable / año / periodo / territorio; clave calculada `ID|Año|Periodo|Territorio` (col. R). Filas libres hasta la 1005; no escribir en R:S. Duplicados → "Revisar".
2. **Evidencia mínima (3):** para que un valor se use: Estado = **Validado**; valor numérico ≥ 0; **nombre de la fuente en K**; **fecha de corte en N**. Añadir enlace (L), página/tabla/anexo (M) y fecha de consulta (O) para verificabilidad. Conteos enteros; porcentajes entre 0 y 100.
3. **Estados (lista de validación col. J):** Pendiente, Validado, No disponible, No evaluable, No aplica. "Pendiente no se utiliza en los cálculos" (fila 2 de Registro). **Validado es una decisión del responsable de la matriz** (19): los ejecutores de esta fase **no deben marcar Validado**; deben dejar Pendiente y documentar la evidencia.
4. **Corte (4):** Anual = 31-dic del año evaluado. Resultados exige que N (fecha de corte) sea **exactamente igual** al corte común N4; de lo contrario la entrada queda N/D.
5. **Sin dato (7):** N/D = faltante, no validado o fuera de corte; NE = base cero sin universo; Revisar = duplicados o numerador fuera del universo. Un cero observado y validado con base positiva da 0 puntos.
6. **Metas PDM (8, 9):** solo se precargan las metas de formulación de los anexos; **verificar modificaciones oficiales antes de validarlas**; registrar en P el horizonte. Con horizonte 2027, el resultado es logro al corte frente a la meta final, no cumplimiento anual.
7. **Educación y ambiente (10):** ILE no se reemplaza por promedio de puntajes Saber; ambiente = espacio público **verde** m²/hab. No inferir resultados desde número de actividades.
8. **PP (11–13):** en PP_Proyectos una fila por proyecto/año/periodo/territorio; PA = personas atendidas únicas en el proyecto; MP = meta anual positiva; exclusiones documentadas en M con "No". PP.CMA.NA = proyectos aplicables del universo completo. Universo municipal deduplicado; no promediar porcentajes zonales; EPP con obligaciones y apropiación definitiva **excluyendo cofinanciación**.
9. **Ámbitos (14):** Control social y PDM: Municipio y Anual. PC: I2 e ISPC municipales; I1 e ISPC anuales; I3–I6 admiten semestre. PP: anual, municipal o por zona.
10. **Control social (17):** universos definidos antes de medir; convocatoria ≠ presencia de veedurías; acuse ≠ respuesta de fondo; cierre sustantivo 31-dic; búsqueda documental hasta 31-mar siguiente; no contar respuestas posteriores al cierre.
11. **Formatos (18):** 99,5 % se digita 99,5; fechas dd/mm/aaaa; presupuestos en COP.
12. **Observaciones (19):** registrar discrepancias, modificaciones, definiciones de universo y versiones de programación en la col. Q.

**Cómo encaja la escala de disponibilidad con el Registro:** la clasificación (Disponible/Construible/Parcial/Proxy/No disponible) debe ir en el entregable de disponibilidad y, si se toca el Registro de una copia de trabajo, en **Observaciones (Q)**, nunca en Estado (J). Un Proxy **no** debe digitarse como valor obtenido (I).

---

## 3. Catálogo de los 23 indicadores e índices

| Sistema | Código | Indicador / índice | Tipo | Fórmula (Matriz) | Variables (ID Registro) | Periodicidad / ámbito | Semáforo (verde / amarillo / rojo) |
|---|---|---|---|---|---|---|---|
| CS | CS.DSI | Disponibilidad de información de seguimiento | Indicador | DSI = 100·MD/MS | CS.DSI.MD, CS.DSI.MS | Anual, Municipio | ≥90 / 70–<90 / <70 |
| CS | CS.PCS | Cumplimiento de la promoción del control social | Indicador | PCS = 100·EE/EP (tope 100) | CS.PCS.EE, CS.PCS.EP | Anual, Municipio | ≥75 / 50–<75 / <50 |
| CS | CS.VRC | Presencia de veedurías en la rendición de cuentas | Indicador | VRC = 100·EV/ET | CS.VRC.EV, CS.VRC.ET | Anual, Municipio | ≥75 / 50–<75 / <50 |
| CS | CS.RPO | Respuesta pública a observaciones ciudadanas | Indicador | RPO = 100·OR/OT | CS.RPO.OR, CS.RPO.OT | Anual, Municipio | ≥90 / 70–<90 / <70 |
| CS | CS.ICRCS | Índice de Condiciones y Respuesta para el Control Social | Índice | (DSI+PCS+VRC+RPO)/4 | (4 puntajes) | Anual, Municipio | ≥75 / 50–<75 / <50 |
| PC | PC.I1 | Incorporación de aportes ciudadanos | Indicador (principal) | I1 = 100·V1/V2 | PC.I1.V1, PC.I1.V2 | Anual | ≥70 / 40–<70 / <40 |
| PC | PC.I2 | Cobertura territorial de la participación | Indicador (principal) | I2 = 100·V1/V2 | PC.I2.V1, PC.I2.V2 | Semestral y anual; municipal | ≥80 / 50–<80 / <50 |
| PC | PC.I3 | Cumplimiento de la oferta formativa | Indicador (principal) | I3 = 100·V1/V2 | PC.I3.V1, PC.I3.V2 | Semestral y anual | ≥75 / 50–<75 / <50 |
| PC | PC.I4 | Publicación de resultados participativos | Indicador (principal) | I4 = 100·V1/V2 | PC.I4.V1, PC.I4.V2 | Semestral y anual | ≥90 / 70–<90 / <70 |
| PC | PC.I5 | Cumplimiento de sesiones de participación | Complementario | I5 = 100·V1/V2 | PC.I5.V1, PC.I5.V2 | Semestral y anual | ≥75 / 50–<75 / <50 |
| PC | PC.I6 | Disponibilidad de participación virtual | Complementario | I6 = 100·V1/V2 | PC.I6.V1, PC.I6.V2 | Semestral y anual | ≥60 / 30–<60 / <30 |
| PC | PC.ISPC | Índice Simplificado de Participación Ciudadana de Envigado | Índice | 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4 (I5, I6 excluidos) | (4 puntajes) | Anual, Municipio | ≥75 / 50–<75 / <50 |
| PDM | PDM.CS | Logro de la cobertura de aseguramiento subsidiado | Indicador | CS = 100·O/M (máx. 100) | PDM.CS.O, PDM.CS.M | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.CM | Logro de la reducción de mortalidad vial | Indicador | CM = 100·M/O (máx. 100; si O=0 → 100) | PDM.CM.O, PDM.CM.M | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.CA | Logro del mantenimiento de espacio público verde | Indicador | CA = 100·O/M (máx. 100) | PDM.CA.O, PDM.CA.M | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.CP | Logro de implementación de procesos participativos | Indicador | CP = 100·O/M (máx. 100) | PDM.CP.O, PDM.CP.M | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.CD | Logro de cobertura de servicios deportivos | Indicador | CD = 100·O/M (máx. 100) | PDM.CD.O, PDM.CD.M | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.ILE | Índice de Logro Educativo | Indicador (compuesto) | LA = 100·A/MA; LS = 100·S/MS (cada uno tope 100); ILE = (LA+LS)/2 | PDM.ILE.A, PDM.ILE.MA, PDM.ILE.S, PDM.ILE.MS | Anual, Municipio | 100 / 75–<100 / <75 |
| PDM | PDM.ICPD-6 | Índice de cumplimiento de resultados seleccionados del PDM | Índice | IBS = (CS+ILE+CD)/3; ITG = (CM+CA+CP)/3; ICPD-6 = (IBS+ITG)/2 | (6 puntajes) | Anual, Municipio | 100 / 75–<100 / <75 |
| PP | PP.EPP | Ejecución presupuestal del PP | Indicador | EPP = 100·OPP/APP | PP.EPP.OPP, PP.EPP.APP | Trimestral si hay reporte; cierre anual; municipal o zona | 90–100 / 75–<90 / naranja 50–<75 / <50 |
| PP | PP.CEP | Cumplimiento de entrega de proyectos | Indicador | CEP = 100·PEC/PE | PP.CEP.PEC, PP.CEP.PE | Ídem | Ídem |
| PP | PP.CMA | Cumplimiento de metas de atención | Indicador | por proyecto 100·mín(PA/MP;1); CMA = SC/NA | PP.CMA.NA (+ PA, MP por proyecto en PP_Proyectos; SC calculado) | Ídem | Ídem |
| PP | PP.IREPP | Índice de Resultados de Ejecución del PP | Índice | 0,20·EPP + 0,40·CEP + 0,40·CMA | (3 puntajes) | Anual | Ídem; alerta si CEP o CMA < 50 |

Reglas de índices (Guía 15): faltantes no cambian pesos y **bloquean** el índice (N/D). Por tanto, una sola variable No disponible deja el índice de su sistema en N/D.

---

## 4. Metas precargadas de la Guía y contraste preliminar con el PDM 2024–2027

La Guía (8) precarga: salud 99,5; movilidad 8; ambiente 3,88; participación 100; deporte 6; educación 19 y 14. En Registro están en las 4 vigencias (2024–2027), estado Pendiente, horizonte 2027, con la observación "Precargada desde el anexo; verificar meta vigente y horizonte antes de validar".

Contraste preliminar realizado por el Analizador (lectura puntual de páginas; **el Ejecutor debe confirmarlo y citarlo formalmente**):

| Meta | Variable | Indicador de resultado del PDM | Línea base (PDM) | Apéndice 9.1.2 (27-may-2024) | PDM aprobado (Alcaldía, 807 p.) | DTS modificado 22-dic-2025 (841 p.) | Matriz cita |
|---|---|---|---|---|---|---|---|
| 99,5 | PDM.CS.M | Cobertura del aseguramiento en salud del Régimen Subsidiado (%; estable) | 99,5 (2022) | fila 16: meta 2027 = 99,5 | p. 602: 99,5 | p. 640: 99,5 | PDM 1.1.4, p. 518 |
| 8 | PDM.CM.M | Tasa de mortalidad por siniestros viales (por 100.000 hab.; decreciente) | 10,2 (año: 2023 en apéndice; sin año en PDM y DTS) | fila 119: meta 2027 = 8 | p. 706: 8 | p. 738: 8 | programa 3.2.2, p. 599 |
| 3,88 | PDM.CA.M | Espacio público verde (m²/hab.; estable) | 3,88 (2023) | fila 129: 3,88 | p. 717: 3,88 | p. 748: 3,88 | programa 3.3.2, p. 607 |
| 100 | PDM.CP.M | Implementación de los procesos de participación ciudadana (%; creciente) | 75 (2023 en apéndice; sin año en PDM y DTS) | fila 146: 100 | p. 742: 100 | p. 775: 100 | programa 5.1.1, p. 628 |
| 6 | PDM.CD.M | Cobertura de servicios deportivos de disciplinas tradicionales, nuevas tendencias y adaptadas (%) | 5,3 (2023) | fila 24: 6 | p. 609: 6 | p. 646: 6 | programa 1.2.1, pp. 522–523 |
| 19 | PDM.ILE.MA | Tasa de cobertura en atención integral de niños y niñas de 0 a 6 años (%) | 16,3 (2023) | fila 71: 19 | p. 650: 19 | p. 683: 19 | programas 2.1.1, pp. 553–554 |
| 14 | PDM.ILE.MS | IE oficiales que mejoran la categoría de clasificación en pruebas Saber 11 (%) | 0 (2023) | fila 74: 14 | p. 652: 14 | p. 685: 14 | programa 2.1.2, pp. 553–554 |

Archivos para el contraste (ruta_local exacta):
- `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` **[inspeccionado]** – hoja `Apéndice9.1.2.` (indicadores de resultado, línea base, META 2027, fuente, responsable); hoja `Apéndice9.1.1` (productos con metas 2024–2027).
- `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` **[inspeccionado]** – PDM aprobado (Acuerdo 017 de 2024).
- `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/documentotecnicosoportepdm22122025.pdf` **[inspeccionado]** – DTS ajustado tras la modificación de dic-2025.
- `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/expedientecompletoproyectoacuerdo0592025.pdf` **[inspeccionado, solo p. 1]** – Proyecto de Acuerdo (dic-2025) que **modifica parcialmente el Acuerdo 017 de 2024**. Verificar si fue aprobado y si toca algún indicador de resultado (en la muestra revisada solo elimina/ajusta productos).
- `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/alternativas-de-solucion-planes-zonales/116988_acuerdo017de2024.pdf` – Acuerdo 017 de 2024.
- `data_raw_IIC/05_fuentes_externas_y_pot/concejoenvigado_gov_co/media_/PLAN-DE-DESARROLLO-MUNICIPAL-2024-2027.pdf` – copia del Concejo (314 p.; no coincide con la paginación citada por la Matriz).

**Conclusión preliminar:** las 7 metas coinciden con el PDM aprobado y con el DTS modificado de dic-2025. Ninguna es meta anual: todas son **meta 2027** (horizonte final). El PDM no publica metas anuales de resultado, por lo que la lectura "evaluación anual" (Guía 9) no tiene metas comparables oficiales.

---

## 5. Hallazgos transversales útiles para todos los sistemas

- **Plan indicativo de productos con programado y ejecutado por año:** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` **[inspeccionado]**, hoja `Plan-Indicativo-Productos` (593 filas × 151 col.): columnas `META PROGRAMADA 2024..2027`, `AVANCE EJECUTADO 2024..2027`, eficacia, ejecución financiera, columnas `ZONA 1`…`ZONA 13` (planes zonales) y códigos de políticas públicas (p. ej. `06-PP-PAR-AC-2022`). Hoja `ROCA-2026-07-01` con rubros, fuente, proyecto y BPIN. Es la fuente más directa para productos del PDM (2024, 2025 y 2026 al 2.º trimestre).
- **Plan de acción (seguimiento a proyectos), con apropiación y obligación por proyecto:** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx` **[inspeccionado]** (hojas `PlanAcción27.01.2026` con CÓDIGO PROYECTO, actividad, valor programado/ejecutado, APROPIACIÓN DEFINITIVA, EJECUCIÓN OBLIGACIÓN; y `ROCA07.01.26` con RECURSO, FUENTE, PROYECTO, BPIN, APROPIACIÓN DEFINITIVA, OBLIGACIÓN ACUMULADA). Equivalentes: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/e_def_planaccion3tri2025.xlsx`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2026/evaluaciondefinitivaplanaccion2026segundotrimestre20260730.xlsx`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2026/evaluaciondefinitivaplanaccion2026primertrimestre20260430.xlsx`, y para 2024 `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/indicadoresdegestionplandeaccion2024.xlsx`.
- **Evaluaciones trimestrales del PDM (presentaciones PDF, poco texto extraíble; incluyen sección "Indicadores de resultado"):** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202431032025.pdf` (cierre 2024), `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202525022026.pdf` (cierre 2025) **[inspeccionado: 19 p., texto mínimo, requiere lectura visual]**, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionprovisionalplandesarrollomunicipalcuartotrimestre202527012026.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalcuartotrimestre202430012025-1.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalprimertrimestre202513052025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalsegundotrimestre202522072025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/eplandllomun3tri2025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollotercertrimestre2024.pdf`.
- **Informes de gestión por dependencias (Control Interno):** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` (vigencia 2025, 304 p., texto extraíble) y `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf` (vigencia 2024).
- **Informe de rendición de cuentas 2024 del alcalde:** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/2024--2027-raul-cardona-gonzalez/rendiciondecuentas2024digital.pdf`.
- **Atención:** 5 archivos de evaluación del PDM 2025–2026 tienen `ventana_temporal = 2022` en el inventario (error de catalogación). Buscar por nombre, no solo por ventana (ver 8.2).

---

## 6. Mapa por variable

Formato de cada ficha: **ID · nombre · unidad · indicador · fuente sugerida por la Matriz · candidatos en el repo (ruta_local) · fuentes externas · palabras clave · clasificación preliminar y motivo · territorio/periodo.** Todas las fuentes externas figuran como "URL por buscar" porque la Matriz **no trae ninguna URL** (0 hipervínculos en las 5 hojas).

### 6.1 Control social (CS) – 8 variables

#### CS.DSI.MD – Número de metas seleccionadas con valor observado publicado para la vigencia [Conteo] · Indicador CS.DSI
- **Fuente sugerida (Matriz):** PDM para identificar metas; informe público de seguimiento o evaluación para valores; repositorio del PDM y actas del Concejo como rutas. Exigir código o nombre, valor observado, vigencia y unidad; la línea base no es ejecución.
- **Candidatos repo:**
  - `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` (AVANCE EJECUTADO 2024/2025 por producto)
  - `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/planindicativoproductos20261trimestre.xlsx`
  - Evaluaciones trimestrales del PDM (sección 5), en particular `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202431032025.pdf` y `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202525022026.pdf`
  - `data_raw_IIC/05_fuentes_externas_y_pot/dnp_desempeno_territorial/Copia_de_Resultados_Eficacia_2023.xlsx` (formato DNP, solo 2023: referencia)
- **Fuentes externas:** Portal de transparencia / seguimiento del PDM de la Alcaldía (URL por buscar); KPT / Sistema de Seguimiento a Planes de Desarrollo del DNP (URL por buscar).
- **Palabras clave:** repo: `planindicativo`, `evaluacion plan desarrollo`, `avance ejecutado`, `eficacia`; web: "Envigado evaluación plan de desarrollo 2025 indicadores de resultado", "seguimiento plan indicativo Envigado".
- **Clasificación preliminar: Construible.** Las cifras por producto existen (xlsx); pero MD depende de una selección previa de metas (MS) que aún no está definida. Una vez fijada la lista, el conteo es directo para 2024 y 2025; 2026 solo al 2.º trimestre (no hay cierre anual).
- **Territorio/periodo:** Municipio, anual. Hay columnas por zona en el plan indicativo (no requeridas).

#### CS.DSI.MS – Número total de metas incluidas previamente en la selección de seguimiento [Conteo] · CS.DSI
- **Fuente sugerida:** decisión metodológica (Matriz: "Iniciar con 10 metas como propuesta manejable; mantener la lista").
- **Candidatos repo:** `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` (universo de metas para seleccionar: 9.1.1 productos, 9.1.2 resultados).
- **Fuentes externas:** ninguna; es una definición del observatorio.
- **Palabras clave:** —
- **Clasificación preliminar: Construible** (por definición; requiere que los docentes aprueben la lista de 10 metas antes de medir, Guía 17).
- **Riesgo:** sin lista acordada, DSI = N/D para todos los años.

#### CS.PCS.EE – Número de estrategias de promoción del control social ejecutadas y reportadas en el periodo [Conteo] · CS.PCS
- **Fuente sugerida:** reporte público de ejecución de Planeación; PDM proyecto 5.1.2.1. Contar estrategias del reporte, no cada curso o sesión. La oferta de la EFPC no prueba ejecución.
- **Candidatos repo:**
  - `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` **[inspeccionado]**: consecutivo 544, producto "Servicio de formación para el control social", indicador "Estrategias de promoción del control social" (5.1.2.1.2, unidad 07), AVANCE EJECUTADO 2024 = 1 y 2025 = 1; con observaciones de la evaluación.
  - `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx` (actividades del proyecto)
  - `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/indicadoresdegestionplandeaccion2024.xlsx` (2024)
  - Soportes de ejecución: `data_raw_IIC/03_actas_y_comites_tecnicos/efpc_envigado_gov_co/efpc_2025-03/CONTROL-SOCIAL-A-LA-GESTION-PUBLICA.pdf`, `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/veedurias-en-salud/capacitacioncontrolsocial.pdf`, `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/veedurias-en-salud/fortalecimientoveeduriasciudadanassalud.pdf` (material; no prueban por sí solos la estrategia)
- **Fuentes externas:** reporte de ejecución del DAP – Subdirección de Participación Ciudadana (URL por buscar).
- **Palabras clave:** repo: `control social`, `5.1.2.1`, `estrategias de promocion`; web: "Envigado estrategia promoción control social 2025".
- **Clasificación preliminar: Disponible** (2024 y 2025, cifra oficial de ejecución de producto). 2026: **Parcial** (solo avance a 2.º trimestre).
- **Brecha:** el reporte es autodeclarado; verificar soporte de ejecución (Guía 3) y que "estrategia" sea la unidad del proyecto 5.1.2.1.

#### CS.PCS.EP – Número de estrategias de promoción del control social programadas oficialmente para el mismo periodo [Conteo] · CS.PCS
- **Fuente sugerida:** programación oficial; meta de formulación 4 estrategias en el cuatrienio ("no asumir una por año").
- **Candidatos repo:** el mismo xlsx del plan indicativo (META PROGRAMADA 2024 = 1, 2025 = 1, 2026 = 1, 2027 = 1); `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` (Apéndice 9.1.1, fila 553: meta plan 4; 1/1/1/1); PDM p. 748 y DTS 2025 p. 781.
- **Fuentes externas:** —
- **Palabras clave:** `meta programada`, `estrategias de promocion del control social`.
- **Clasificación preliminar: Disponible** (programación anual oficial publicada). La advertencia de la Matriz de "no asumir una por año" queda resuelta por la programación del plan indicativo; documentarlo en Observaciones.

#### CS.VRC.EV – Número de eventos municipales de rendición de cuentas realizados con presencia documentada de veedurías [Conteo] · CS.VRC
- **Fuente sugerida:** balance de rendición de cuentas; actas, memorias o informes de eventos. Acreditar presencia; convocatoria no prueba asistencia.
- **Candidatos repo:**
  - Plan indicativo 2026-2T (consecutivo 545: "Eventos de rendición pública de cuentas con veedurías ciudadanas realizados", programado 1/año, ejecutado 2024 = 1, 2025 = 1) – dato institucional, no acredita presencia.
  - Actas/evaluaciones de eventos: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/acta-audiencia-publica-de-rendicion-de-cuentas-2024/2024actarendicioncuentassecretariasalud02docx.pdf`, `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-202501-secretaria/2025actarendicioncuentassecretariasalud01final.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/primera-rendicion-de-cuentas-secretaria-de-salud-de/134500_2025actarendicioncuentassecretariasalud01final.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026actarendicioncuentassecretariasalud01.pdf`, `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/evaluacionjornadarendiciondecuentas2024.xlsx`
  - Veeduría en PP: `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/rendicion-de-cuentas-presupuesto-participativo-2024/veeduria-118--informe-gestion-2024-1.pdf`
  - Registro de veedurías (para identificar veedurías activas): `data_raw_IIC/03_actas_y_comites_tecnicos/datos_gov_co/datosgov_jju5-br3w_Veedurias_Ciudadanas_Personeria_Envigado.csv` **[inspeccionado: 45 registros, años de registro 2012–2025; contiene datos personales]**
- **Fuentes externas:** Personería de Envigado – registro de veedurías (URL por buscar); Red Institucional de Apoyo a Veedurías (URL por buscar).
- **Palabras clave:** repo: `veedur`, `acta audiencia`, `rendicion de cuentas`, `asistencia`; web: "Envigado rendición de cuentas 2025 acta veeduría".
- **Clasificación preliminar: Parcial.** Existe el conteo institucional (1 por año) y actas de algunos eventos; falta acreditar presencia de veedurías evento por evento para el universo ET.

#### CS.VRC.ET – Número total de eventos municipales de rendición de cuentas realizados en el periodo [Conteo] · CS.VRC
- **Fuente sugerida:** balance completo de rendición de cuentas de la Alcaldía. Los debates del Concejo no integran ET.
- **Candidatos repo (eventos identificables 2024–2026):**
  - Alcalde: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf` (100 días 2024)
  - Sectoriales: Salud (2024 ×2, 2025, 2026: rutas en CS.VRC.EV y `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/primera-rendicion-de-cuentas-de-la-secretaria-de-salud/20240321primerarendicincuentassalud-1.pdf`); Víctimas (`data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-de-la-politica-publica-de-victimas/infrendcuentasvca2025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/presentacion-rendicion-de-cuentas-vca-2024/programa-victimas.pdf`); Grupos étnicos (`data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-de-la-poblacion-afro-indigena-y/rendiciondecuentasgruposetnicos2025.pdf`); Comité de Discapacidad (2024 y 2025: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/justificacionnoplandemejoramiento20242.pdf`, `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/rendicion-de-cuentas-del-comite-municipal-de-discapacidad-961114/justificacionnoplandemejoramiento20251.pdf`); PP 2024 (`data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/rendicion-de-cuentas-presupuesto-participativo-2024/justificacionnoplandemejoramiento.pdf`).
  - Programación: plan indicativo (consecutivo de "Eventos de rendición pública de cuentas realizados", 5.1.2.1.4: 1/año).
  - Excluir: `data_raw_IIC/02_ejecucion_y_presupuesto/concejoenvigado_gov_co/media_/Rendicion-de-Cuentas-2024_Concejo-de-Envigado-30112024.pdf` (Concejo).
- **Fuentes externas:** Estrategia de rendición de cuentas / micrositio de rendición de cuentas de la Alcaldía (URL por buscar); FURAG – política de rendición de cuentas (`data_raw_IIC/05_fuentes_externas_y_pot/datos_gov_co/datosgov_daed-z4fw_FURAG_MIPG_Envigado.csv`, solo 2022–2024, proxy de contexto).
- **Palabras clave:** repo: `rendicion de cuentas`, `audiencia publica`, `justificacion no plan de mejoramiento` (cada justificación marca un evento cerrado); web: "Envigado estrategia de rendición de cuentas 2025 cronograma".
- **Clasificación preliminar: Construible.** Hay evidencia de múltiples eventos; requiere fijar el universo ("eventos municipales": ¿solo audiencia del alcalde o también sectoriales y de instancias?) antes de contar (Guía 17). Sin esa decisión, ET puede variar entre 1 y ~8 por año.

#### CS.RPO.OR – Número de observaciones ciudadanas del universo evaluado con respuesta de fondo publicada al corte [Conteo] · CS.RPO
- **Fuente sugerida:** matriz pública de preguntas y respuestas, acta integral o evaluación de rendición de cuentas con todas las observaciones.
- **Candidatos repo:** `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` (duplicado en `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-alcalde/preguntasyrecomendacionesrendiciondecuentas2024.pdf`) **[inspeccionado: 24 p., cuadro numerado con pregunta y respuesta]**; `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf` **[inspeccionado: 5 p., mismo formato]**; `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf`; Salud: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-rendicion-de-cuentas-secretaria/preguntasyrespuestasrendicioncuentassaludi2025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026preguntasrespuestasrendicioncuentas01.pdf`; consultas previas: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/envigado-invita-a-la-comunidad-a-proponer-temas-para/formato_consulta_ciudadana_rendicion_cuentas-2025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/envigado-abre-consulta-ciudadana-para-construir-la-rendicion/formato-consulta-ciudadana-rendicion-cuentas-2026.docx` (formatos, no respuestas).
- **Fuentes externas:** micrositio de rendición de cuentas (URL por buscar).
- **Palabras clave:** `preguntas y respuestas`, `preguntas y recomendaciones`, `respuesta`; web: "Envigado preguntas y respuestas rendición de cuentas 2025".
- **Clasificación preliminar: Construible** (2024, 2025). Requiere clasificar cada respuesta como "de fondo" vs acuse y filtrar observaciones "sobre el PDM". **Contiene nombres de ciudadanos: no publicar datos personales.**

#### CS.RPO.OT – Número total de observaciones ciudadanas sobre el PDM registradas públicamente en ejercicios cerrados de rendición de cuentas del periodo [Conteo] · CS.RPO
- **Fuente sugerida / candidatos:** los mismos de CS.RPO.OR (los cuadros numeran todas las intervenciones, incluidas sugerencias y felicitaciones).
- **Clasificación preliminar: Construible.** Riesgo: definir qué intervenciones son "observaciones sobre el PDM" (excluir felicitaciones) y si el universo son solo eventos del alcalde o también sectoriales; mantener coherencia con CS.VRC.ET. 2026: solo evento de Salud publicado al corte.

### 6.2 Participación ciudadana (PC) – 12 variables

Fuentes comunes del sistema:
- Política Pública de Participación (Acuerdo 033 de 2022): `data_raw_IIC/03_actas_y_comites_tecnicos/efpc_envigado_gov_co/efpc_2024-12/ACUERDO-033-DE-2022-Politica-Publica-de-Partipacion-Ciudadana-Envigado.pdf`; ejes: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/estrategia-de-participacion-ciudadana--2024/ejes-estrategicos_politicapublica_participacionciudadana.pdf`.
- Calendario de la Subdirección de Participación: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/calendario-anual-subdireccion-participacion-ciudadana/calendarioparticipacionciudadana2024.pdf` (solo 2024).
- Plan indicativo y plan de acción (sección 5), productos 5.1.1.x (formación, PP, instancias, observatorio).
- Política de Participación Social en Salud (PPSS, Res. 2063/2017), programación y seguimiento con fechas, población y enlaces de evidencia: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion/pps190spps20241231mu000000005266.xlsx` (2024, 0,0 MB: verificar contenido) y `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` **[inspeccionado: hojas PROGRAMACIÓN y SEGUIMIENTO]**. Es sectorial (salud), no la política general.

#### PC.I1.V1 – Número de procesos participativos cerrados que incorporaron al menos un aporte ciudadano en una decisión formal [Conteo] · PC.I1
- **Fuente sugerida:** informes de cierre, actas de instancias y decisiones de Participación y dependencias; actas del Concejo cuando corresponda. Vincular aporte ↔ decisión.
- **Candidatos repo:** PP (iniciativas → proyectos priorizados): `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/88-iniciativas-ciudadanas-fueron-postuladas-para-el-515924/resumeniniciativascomunitariaspp2024.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/90-iniciativas-ciudadanas-fueron-postuladas-para-el-694496/90-iniciativas-ciudadanas-fueron-postuladas-para-el-proceso-de-presupuesto-participativo-2025--2026.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2025.pdf`; planeación participativa del PDM: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/plan-de-desarrollo-20242027/conceptoctp13042024.pdf`, `data_raw_IIC/05_fuentes_externas_y_pot/concejoenvigado_gov_co/media_24898/` (metodología de participación en debates del PDM; ver inventario), `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice961articulacionplaneszonalesdesarrollo.xlsx`; POT: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/programacion-de-las-asambleas-zonales-del-plan-de-ordenamiento/cronograma-asambleas-pot.pdf`; consulta de decretos: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/consulta-publica-para-el-proyecto-decreto-de-modificacion/` (ver inventario).
- **Fuentes externas:** informes de cierre de procesos de la Subdirección de Participación (URL por buscar); Observatorio de Participación (producto 5.1.1.4.5 del PDM: "Informes de gestión del conocimiento") (URL por buscar).
- **Palabras clave:** repo: `iniciativas`, `priorizacion`, `asambleas zonales`, `concepto CTP`, `consulta publica`; web: "Envigado informe de cierre proceso participativo 2025", "Envigado presupuesto participativo proyectos priorizados 2025".
- **Clasificación preliminar: Parcial.** Solo es trazable para algunos procesos (PP: iniciativa → proyecto; PDM: CTP → ajustes). No existe inventario oficial de procesos participativos cerrados.

#### PC.I1.V2 – Número total de procesos participativos cerrados en los que se recibieron aportes ciudadanos [Conteo] · PC.I1
- **Candidatos repo:** los de PC.I1.V1 + calendario 2024.
- **Clasificación preliminar: Parcial.** Requiere construir el inventario de procesos (decisión de universo, Guía 17). Sin inventario oficial; riesgo alto de subconteo.

#### PC.I2.V1 – Número de zonas con al menos una actividad participativa realizada para su comunidad [Conteo] · PC.I2
- **Fuente sugerida:** consolidado de Participación, comités zonales y organizadores con fecha, zona y acta o asistencia. Los planes zonales no prueban actividad.
- **Candidatos repo:** plan de acción 2025 IV trim. (`data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx`, proyecto 0553 "Implementación del presupuesto participativo": "Eventos de asambleas zonales realizados", programado 12 / ejecutado 12) **[inspeccionado]**; `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/programacion-de-las-asambleas-zonales-y-puntos-descentralizados/asambleas-zonales-y-puntos-descentralizados-pp2027.pdf` (programación 2026); `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/estructuracion-tecnica-de-proyectos-de-presupuesto-participativo/espaciosdeliberatoriospresupuestosparticipativos.pdf` (2024); calendario 2024; POT asambleas 2025.
- **Fuentes externas:** consolidado de asistencia por zona (Subdirección de Participación) (URL por buscar).
- **Palabras clave:** `asambleas zonales`, `espacios deliberatorios`, `zona`, `asistencia`.
- **Clasificación preliminar: Construible** (al menos las asambleas zonales de PP cubren las zonas; requiere acta o asistencia por zona para acreditar). Disponible por zona (cobertura sí/no).

#### PC.I2.V2 – Número total de zonas oficiales del municipio [Conteo] · PC.I2
- **Fuente sugerida:** catálogo oficial de zonas; validar con Planeación.
- **Candidatos repo:** plan indicativo 2026-2T (columnas `ZONA 1` a `ZONA 13`); planes zonales: `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-06/PLAN-ZONA-2.pdf` … `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-06/PLAN-ZONA-13.pdf` (12 archivos, **no hay PLAN-ZONA-1**); `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/` (POT); `data_raw_IIC/05_fuentes_externas_y_pot/efpc_envigado_gov_co/efpc_2024-12/mapa1.pdf`, `data_raw_IIC/05_fuentes_externas_y_pot/efpc_envigado_gov_co/efpc_2024-12/mapa2.pdf`.
- **Fuentes externas:** Acuerdo del POT vigente / decreto de división territorial (URL por buscar).
- **Palabras clave:** `zona 1`, `division territorial`, `zonas urbanas rurales`.
- **Clasificación preliminar: Disponible**, con verificación pendiente: el plan indicativo usa 13 zonas pero el PP y los planes zonales solo listan zonas 2–13. Confirmar si la zona 1 existe/participa (¿zona sin población o rural?) antes de fijar V2 = 13 o 12.

#### PC.I3.V1 – Número de acciones formativas programadas para la ciudadanía que se realizaron [Conteo] · PC.I3
- **Fuente sugerida:** programación fechada y reporte de ejecución de la Escuela (EFPC) con soporte de finalización; excluir formación para servidores.
- **Candidatos repo:** plan indicativo 2026-2T **[inspeccionado]**: consecutivo 521 "Personas formadas para la participación incidente" (programado 10/15/15/10; ejecutado 2024 = 95, 2025 = 186), consecutivo 524 semillero (40/año; ejecutado 45 y 80), consecutivo 525 funcionarios (excluir); "Personas que participan en actividades de capacitación para la participación" (Apéndice 9.1.1 fila 529). Material EFPC: `data_raw_IIC/03_actas_y_comites_tecnicos/efpc_envigado_gov_co/efpc_2024-12/Lineamientos-EFPC-2022.pdf`, `data_raw_IIC/05_fuentes_externas_y_pot/efpc_envigado_gov_co/efpc_2024-12/CARTILLA.pdf`.
- **Fuentes externas:** portal de la Escuela de Formación para la Participación Ciudadana (oferta de cursos/diplomados) (URL por buscar).
- **Palabras clave:** repo: `efpc`, `diplomado`, `curso`, `formacion participacion incidente`; web: "Escuela de Formación para la Participación Ciudadana Envigado diplomado 2025".
- **Clasificación preliminar: Proxy.** Lo publicado cuenta **personas**, no acciones formativas; no se halló programación fechada de cursos.

#### PC.I3.V2 – Número total de acciones formativas programadas para la ciudadanía con finalización prevista en el periodo [Conteo] · PC.I3
- **Candidatos repo:** los de PC.I3.V1 (metas programadas en personas).
- **Clasificación preliminar: Proxy** (la programación oficial está en personas). Si el portal de la EFPC publica la oferta fechada, podría subir a Construible.

#### PC.I4.V1 – Número de procesos participativos cerrados con resultados publicados y consultables [Conteo] · PC.I4
- **Fuente sugerida:** inventario de procesos cerrados y enlaces públicos; convocatorias y fotos no son resultados.
- **Candidatos repo:** los resultados publicados de PP (Proyectos_PP_2024/2025, resumen de iniciativas 2024, rendición PP 2024), concepto CTP del PDM, planes zonales 2025, síntesis del diagnóstico POT (`data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/publicacion-de-la-sintesis-del-diagnostico-del-pot/`), `data_raw_IIC/03_actas_y_comites_tecnicos/efpc_envigado_gov_co/efpc_2025-04/INFORME-FINAL-EXPERIENCIAS-DE-PARTICIPACION-NO-FORMAL-EN-ENVIGADO-2023.pdf`. Los `url_origen` del inventario sirven como "enlaces públicos" de evidencia.
- **Palabras clave:** `resultados`, `informe final`, `sintesis`, `publicacion`.
- **Clasificación preliminar: Parcial** (mismo problema de universo que I1; la publicación sí es verificable proceso por proceso con los url_origen).

#### PC.I4.V2 – Número total de procesos participativos cerrados [Conteo] · PC.I4
- **Clasificación preliminar: Parcial.** Debe compartir el inventario de procesos con PC.I1.V2 (I1.V2 es subconjunto: procesos con aportes).

#### PC.I5.V1 – Número de sesiones previstas de las instancias seleccionadas que se realizaron [Conteo] · PC.I5
- **Fuente sugerida:** calendarios y actas de las instancias; Consejo Municipal de Participación como referencia.
- **Candidatos repo:** COPACO: `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/comite-de-participacion-comunitaria-copaco/20251125_instancia-copaco-1.pdf`; Comité de Discapacidad (rendiciones 2024–2026); PPSS (pps190 xlsx, hojas PROGRAMACIÓN/SEGUIMIENTO); producto "Asistencias técnicas a instancias de participación" y "Eventos de integración de instancias" (Apéndice 9.1.1 filas 543–544); acuerdos de instancias en `data_raw_IIC/03_actas_y_comites_tecnicos/efpc_envigado_gov_co/efpc_2025-09/` (reglamentos, sin calendario).
- **Fuentes externas:** actas del Consejo Municipal de Participación Ciudadana, CTP, JAL/JAC (URL por buscar).
- **Palabras clave:** `consejo municipal de participacion`, `CTP`, `sesion`, `quorum`, `acta`, `instancia`.
- **Clasificación preliminar: Parcial.** Hay soportes dispersos de algunas instancias; no hay calendarios de sesiones previstas para la mayoría.

#### PC.I5.V2 – Número total de sesiones previstas de esas mismas instancias para el periodo [Conteo] · PC.I5
- **Clasificación preliminar: Parcial** (programación disponible solo para PPSS-salud y quizá COPACO). Requiere seleccionar instancias antes de medir.

#### PC.I6.V1 – Número de actividades participativas realizadas que habilitaron un canal virtual para intervenir [Conteo] · PC.I6
- **Fuente sugerida:** relación de actividades y evidencia del canal habilitado; una transmisión para observar no basta.
- **Candidatos repo:** solo indicios: rendiciones transmitidas por YouTube (p. ej. cuadro 2025 cita el enlace de transmisión), formatos de consulta previa 2025/2026, "Plataforma web para presupuesto participativo en operación" (Apéndice 9.1.1 fila 538).
- **Palabras clave:** `virtual`, `hibrido`, `formulario`, `plataforma`, `YouTube`, `en vivo`.
- **Clasificación preliminar: No disponible.** No hay relación de actividades con evidencia de canal para intervenir; las transmisiones no cuentan.

#### PC.I6.V2 – Número total de actividades participativas realizadas [Conteo] · PC.I6
- **Candidatos repo:** calendario 2024, plan de acción (actividades del proyecto 0553 y de la Subdirección), eventos de rendición, asambleas zonales.
- **Clasificación preliminar: Parcial** (2024 con calendario; 2025–2026 solo actividades de PP y rendición).

### 6.3 Cumplimiento del PDM (PDM) – 14 variables

#### PDM.CS.O – Cobertura porcentual observada del aseguramiento en el régimen subsidiado [Porcentaje 0–100] · PDM.CS
- **Fuente sugerida:** reporte público de aseguramiento de Salud; fuente del plan: Ministerio de Salud. No usar cobertura total del SGSSS ni ejecución presupuestal.
- **Candidatos repo:** evaluaciones del PDM (sección "Indicadores de resultado", lectura visual); informes de rendición de Salud: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/informe-rendicion-de-cuentas-2026-pub-web.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/socializacion-plan-de-accion-secretaria-de-salud-2025/presentacion-rendicion-de-cuentas.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/socializacion-plan-de-accion-2024-secretaria-de-salud/2024presentacionrendicioncuentassalud02.pdf`; `data_raw_IIC/05_fuentes_externas_y_pot/dnp_terridata/TerriData05266f.txt.zip` **[inspeccionado]**.
- **Fuentes externas:** MinSalud / ADRES – BDUA, cifras de aseguramiento por municipio (URL por buscar); SISPRO (URL por buscar).
- **Palabras clave:** repo: `regimen subsidiado`, `aseguramiento`, `cobertura`; web: "cobertura régimen subsidiado Envigado 2025 MinSalud", "BDUA afiliados Envigado".
- **Clasificación preliminar: Proxy.** TerriData publica "Cobertura del régimen subsidiado" (2024 = 12,20; 2025 = 11,88, "valor multiplicado por 100"), que es **otra definición** (afiliados RS sobre población) y no comparable con la meta 99,5. Sube a Disponible si la evaluación del PDM o Salud publica el indicador de resultado con el método del plan.

#### PDM.CS.M – Cobertura meta del aseguramiento en el régimen subsidiado [%] · PDM.CS
- **Clasificación preliminar: Disponible.** 99,5 (2027), coincide en Apéndice 9.1.2, PDM p. 602 y DTS dic-2025 p. 640 (sección 4).

#### PDM.CM.O – Tasa anual observada de mortalidad vial por 100.000 habitantes [Muertes/100.000 hab.] · PDM.CM
- **Fuente sugerida:** informe de seguridad vial; responsable Secretaría de Movilidad. Comparar tasas con igual población; verificar lugar del siniestro o residencia.
- **Candidatos repo:** TerriData (txt.zip) **[inspeccionado]**: "Tasa de homicidios en accidentes de tránsito por cada 100.000 hab." 2023 = 16,8; 2025 = 11,3; **sin 2024** (DNP con datos MinDefensa/DANE); `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/2025informedefinitivoseguimientoplanesmovilidad.pdf`; `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/planes-de-mejoramiento-auditorias-internas-2025/2025_planmejora_seguimientomovilidad.pdf`; evaluaciones del PDM; `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/secretaria-de-movilidad-217778/informeconcejomovilidad2023.pdf` (solo 2023).
- **Fuentes externas:** Observatorio Nacional de Seguridad Vial – ANSV (URL por buscar); INML – Forensis / datos de lesiones fatales de tránsito (URL por buscar); Secretaría de Movilidad de Envigado (URL por buscar); proyecciones de población DANE (URL por buscar).
- **Palabras clave:** `siniestros viales`, `mortalidad vial`, `fallecidos transito`, `seguridad vial`; web: "Envigado fallecidos siniestros viales 2024 2025 tasa".
- **Clasificación preliminar: Proxy.** El único valor numérico hallado proviene de otra fuente (Policía/MinDefensa) y difiere de la línea base del plan (10,2 en 2023 vs 16,8 en TerriData). Falta 2024.

#### PDM.CM.M – Tasa meta de mortalidad vial [Muertes/100.000 hab.] · **Disponible** (8; PDM p. 706, DTS p. 738).

#### PDM.CA.O – Espacio público verde observado, m²/habitante · PDM.CA
- **Fuente sugerida:** informe ambiental o anexo de seguimiento; Secretaría de Medio Ambiente y Desarrollo Agropecuario. No sustituir por espacio público efectivo, árboles o hectáreas.
- **Candidatos repo:** `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/2025informedefinitivoseguimientoplanesmedioam.pdf`; `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/` (14 archivos, 2025; buscar "espacio público verde"); evaluaciones del PDM.
- **Fuentes externas:** inventario de espacio público verde de la Secretaría de Medio Ambiente (URL por buscar); AMVA – indicadores de espacio público verde (URL por buscar).
- **Palabras clave:** `espacio publico verde`, `zonas verdes`, `m2 por habitante`, `EPV`.
- **Clasificación preliminar: No disponible.** No se identificó cifra 2024–2025; el POT diagnóstico podría aportar espacio público **efectivo** (Proxy, no sustituible según la Matriz).

#### PDM.CA.M – **Disponible** (3,88; PDM p. 717, DTS p. 748).

#### PDM.CP.O – Porcentaje observado de implementación de procesos de participación ciudadana, según el método oficial [%] · PDM.CP
- **Fuente sugerida:** informe de seguimiento del PDM o balance de participación con resultado agregado; responsable Planeación. No inferirlo de reuniones, asistentes o cursos.
- **Candidatos repo:** evaluaciones del PDM (indicadores de resultado, lectura visual); `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/2025informedefinitivoseguimientoplanesdap.pdf`; informes de gestión por dependencias (DAP).
- **Fuentes externas:** hoja de vida / ficha técnica del indicador de resultado "Implementación de los procesos de participación ciudadana" (DAP) (URL por buscar).
- **Palabras clave:** `implementacion de los procesos de participacion`, `indicador de resultado`, `ficha tecnica`.
- **Clasificación preliminar: No disponible.** No se halló la cifra ni el método oficial de cálculo (línea base 75 sin ficha). Riesgo: si la evaluación del PDM lo reporta, pasa a Disponible, pero sin método publicado no es verificable.

#### PDM.CP.M – **Disponible** (100; PDM p. 742, DTS p. 775).

#### PDM.CD.O – Cobertura porcentual observada de los servicios deportivos incluidos en el resultado [%] · PDM.CD
- **Fuente sugerida:** informe de gestión o seguimiento del Índer con la cobertura calculada.
- **Candidatos repo:** `data_raw_IIC/02_ejecucion_y_presupuesto/contraloria_envigado_gov_co/informe-definitivo-de-la-auditoria-financiera-de-gestion/informedefinitivoafgr122025inder.pdf` (auditoría Contraloría al Índer, 2025); `data_raw_IIC/02_ejecucion_y_presupuesto/contraloria_envigado_gov_co/informe-definitivo-de-auditoria-abreviada-para-fenecimiento-662265/20260729_informe_definitivo_aafn09.pdf`; `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/rendicion-de-cuentas-presupuesto-participativo-2024/inder--rendicion-de-cuentas-1.pdf`; evaluaciones del PDM (unidad 4.3 Deporte en el resumen ejecutivo del plan indicativo).
- **Fuentes externas:** Índer Envigado – informes de gestión / rendición de cuentas (URL por buscar).
- **Palabras clave:** `cobertura servicios deportivos`, `inder`, `disciplinas`, `nuevas tendencias`, `adaptadas`.
- **Clasificación preliminar: Parcial.** Es probable que el Índer reporte el indicador en su informe de gestión o en la evaluación del PDM, pero no se confirmó para 2024 y 2025; el Índer es ente descentralizado (sus informes no están en el inventario salvo auditorías).

#### PDM.CD.M – **Disponible** (6; PDM p. 609, DTS p. 646).

#### PDM.ILE.A – Cobertura observada de atención integral de niños y niñas de 0 a 6 años [%] · PDM.ILE
- **Fuente sugerida:** porcentajes agregados publicados por Educación.
- **Candidatos repo:** `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` (balance PDM 2024 de Educación); `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-298942/2025informedefinitivoseguimientoplaneseducacion.pdf`; TerriData (número de niños con educación inicial, solo hasta 2023: Proxy); plan indicativo (producto "Niños y niñas con acceso a la educación inicial con atención integral", número: Proxy).
- **Fuentes externas:** MEN – SIMAT / Cuéntame (atención integral a la primera infancia) (URL por buscar); proyecciones DANE 0–6 años (URL por buscar).
- **Palabras clave:** `atencion integral`, `primera infancia`, `educacion inicial`, `0 a 6`.
- **Clasificación preliminar: Parcial** (probable en el balance 2024 de Educación; 2025 no confirmado; los conteos absolutos son Proxy).

#### PDM.ILE.MA – **Disponible** (19; PDM p. 650, DTS p. 683).

#### PDM.ILE.S – Porcentaje observado de instituciones oficiales que mejoran su categoría en Saber 11 [%] · PDM.ILE
- **Fuente sugerida:** porcentajes publicados por Educación; contraste con clasificación de planteles del ICFES. No reemplazar por promedio de puntajes.
- **Candidatos repo:** balance PDM 2024 Educación (arriba); TerriData solo trae puntajes promedio (no usar).
- **Fuentes externas:** ICFES – resultados de clasificación de planteles Saber 11 por año (DataIcfes / FTP) (URL por buscar); directorio de IE oficiales de Envigado (URL por buscar).
- **Palabras clave:** `clasificacion de planteles`, `categoria A+ A B C D`, `Saber 11`, `IE oficiales`.
- **Clasificación preliminar: Construible.** La clasificación por plantel es pública por año; el porcentaje de IE oficiales que mejoran se obtiene comparando t vs t−1 (2024 vs 2023; 2025 vs 2024). Aún no descargada en el repo.

#### PDM.ILE.MS – **Disponible** (14; PDM p. 652, DTS p. 685).

### 6.4 Presupuesto participativo (PP) – 5 variables de Registro (+ PA y MP por proyecto)

Fuentes comunes del sistema:
- Catálogo de proyectos PP por año (zona, nombre, objetivo, valor): `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf` **[inspeccionado: 4 p., texto extraíble, zona + proyecto + objetivo + valor COP; sin código]**, `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2025.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-06/Proyectos_PP_EFC.pdf`.
- Proyectos de inversión 2024 con CODIGO_PROYECTO, META y PRESUPUESTO DEFINITIVO por dependencia: `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/proyectosdeinversion2024.xlsx` **[inspeccionado]** (incluye proyectos zonales de PP, p. ej. inglés por zona).
- Avance de ejecución PP 2024 por zona y operador (35 PDF): carpeta `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/avance-de-ejecucion-de-presupuestos-participativos-2024/` (p. ej. `ejecucionppzona2inder.pdf`, `ejecucionppzona12oopp.pdf`).
- Rendición de cuentas PP 2024 (12 PDF por sector): carpeta `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/rendicion-de-cuentas-presupuesto-participativo-2024/`.
- Proyectos PP 2025 presentados al Concejo: `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-BIENESTAR-ANIMAL.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-DISCAPACIDAD.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-SALUD-MENTAL.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-SALUD-MENTAL-1.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-SEC-MOVILIDAD.pdf`, `data_raw_IIC/01_propuestas_y_acuerdos/concejoenvigado_gov_co/media_/PP-TURISMO.pdf`.
- Plan de acción con apropiación/obligación por proyecto y BPIN (sección 5); ejecución presupuestal de egresos de Hacienda (PDF trimestrales): `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2024-325557/ejecucion-presupuestal-egresos-al-31122024.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2025-576917/ejecucion-presupuestal-egresos-4-al-31122025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2026-358126/ejecucion-presupuestal-egresos-2-al-30062026.pdf`.
- CUIPO: `data_raw_IIC/02_ejecucion_y_presupuesto/datos_gov_co/datosgov_4f7r-epif_CUIPO_Ejecucion_Gastos_Envigado.csv` y `data_raw_IIC/02_ejecucion_y_presupuesto/datos_gov_co/datosgov_d9mu-h6ar_CUIPO_Programacion_Gastos_Envigado.csv` **[inspeccionado]**: periodos 2023-03 a 2026-06; columnas `bpin`, `cod/nom_fuentes_financiacion`, `nom_programatico_mga`, `compromisos`, `obligaciones`, `pagos`. **No hay marcador de fuente "presupuesto participativo"** (búsqueda de "particip" en todas las columnas solo devuelve programáticos MGA genéricos y participación en salud/víctimas).
- Hallazgos previos (insumo IIC): `outputs/limpieza_bpin_cuipo.md` – el puente CODIGO_PROYECTO ↔ BPIN por sufijo **no es válido** (0 confirmados); BPIN de CUIPO limpio 2023–2026. Por tanto, **no** cruzar el catálogo PP con CUIPO por BPIN inferido; usar el plan de acción/ROCA (misma fuente municipal, con CÓDIGO PROYECTO y BPIN en la misma fila).
- Programación de asambleas PP 2027 (publicada 2026): `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/programacion-de-las-asambleas-zonales-y-puntos-descentralizados/asambleas-zonales-y-puntos-descentralizados-pp2027.pdf`.
- **Desfase de vigencia:** las iniciativas postuladas en un año ("PP 2025–2026", 90 iniciativas) se ejecutan en la vigencia siguiente. El Ejecutor debe fijar qué catálogo corresponde a la **vigencia de ejecución** 2024, 2025 y 2026.

#### PP.EPP.OPP – Valor de las obligaciones presupuestales registradas con cargo a recursos de PP [COP] · PP.EPP
- **Fuente sugerida:** informe público de ejecución presupuestal y modificaciones; distinguir obligaciones, compromisos y pagos; separar cofinanciación.
- **Candidatos repo:** plan de acción/ROCA 2024, 2025 (IV trim.) y 2026 (II trim.) (columnas OBLIGACIÓN ACUMULADA por proyecto y fuente); avance de ejecución PP 2024 por zona; ejecución de egresos de Hacienda al 31-dic.
- **Fuentes externas:** CUIPO/CHIP (en repo, sin marcador PP); informe de ejecución de PP de la Subdirección de Participación (URL por buscar).
- **Palabras clave:** `obligacion`, `presupuesto participativo`, `PP`, `zona`, `fuente`, `cofinanciacion`.
- **Clasificación preliminar: Construible** (2024, 2025; 2026 parcial). Se obtiene sumando obligaciones de los proyectos del catálogo PP en el plan de acción/ROCA. **Brecha:** no hay fuente de financiación PP identificada; separar cofinanciación puede no ser posible (documentar).

#### PP.EPP.APP – Valor de la apropiación definitiva de recursos de PP [COP] · PP.EPP
- **Candidatos repo:** los de OPP (APROPIACIÓN DEFINITIVA); valores asignados en Proyectos_PP_AAAA.pdf y PRESUPUESTO DEFINITIVO en proyectosdeinversion2024.xlsx (contraste).
- **Clasificación preliminar: Construible** (mismas condiciones que OPP). Riesgo: proyectos PP que comparten código con recursos ordinarios.

#### PP.CEP.PEC – Número de proyectos con todas las entregas exigibles completas y verificadas al corte [Conteo] · PP.CEP
- **Fuente sugerida:** informes finales, actas de recibo o reportes de supervisión publicados; mantener en el denominador los no iniciados o atrasados.
- **Candidatos repo:** avance de ejecución PP 2024 (35 PDF, incluye `evidenciasejecucionzona10dlloeconomico.pdf`); rendición PP 2024 (12 PDF); plan de acción (ejecución física por actividad); SECOP II contratos (`data_raw_IIC/04_datos_abiertos_y_secop/datos_gov_co/datosgov_jbjy-vk9h_SECOP_II_Contratos_Envigado.csv`, estado del contrato/liquidación; sin BPIN).
- **Palabras clave:** `acta de recibo`, `informe final`, `supervision`, `liquidacion`, `ejecucion fisica`.
- **Clasificación preliminar: Parcial.** 2024 tiene soportes por zona; 2025 y 2026 solo avance físico agregado del plan de acción (no "entrega completa verificada").

#### PP.CEP.PE – Número total de proyectos financiados con entrega exigible en la vigencia [Conteo] · PP.CEP
- **Candidatos repo:** Proyectos_PP_2024.pdf, Proyectos_PP_2025.pdf, proyectosdeinversion2024.xlsx, PP-*.pdf del Concejo (2025).
- **Clasificación preliminar: Construible.** Universo enumerable por año y zona; requiere deduplicar proyectos compartidos entre zonas (p. ej. "Zonas en movimiento" cubre zonas 2–13) y definir "entrega exigible en la vigencia" (Guía 13).

#### PP.CMA.NA – Número de proyectos con meta de atención aplicable en el universo completo y verificado [Conteo] · PP.CMA
- **Fuente sugerida:** fichas públicas con meta principal de atención.
- **Candidatos repo:** proyectosdeinversion2024.xlsx (columna META), plan de acción (indicadores de producto con unidad "personas"), fichas en Proyectos_PP_AAAA.pdf (objetivo).
- **Clasificación preliminar: Parcial.** Se puede marcar aplicabilidad para 2024; para 2025–2026 faltan fichas con meta de atención.

#### (PP_Proyectos) PA – Personas efectivamente atendidas por proyecto · y MP – Meta anual de personas por atender
- **Candidatos repo:** avance de ejecución PP 2024 y rendición PP 2024 (reportan beneficiarios); plan de acción (valor programado/ejecutado de indicadores en personas); proyectosdeinversion2024.xlsx (META).
- **Clasificación preliminar:** PA **Parcial** (2024; sin garantía de deduplicación interna); MP **Parcial** (meta anual en personas solo para algunos proyectos).
- **Territorio:** por zona (se debe consolidar universo municipal deduplicado, Guía 13).

---

## 7. Conteo preliminar esperado por categoría

| Sistema | Disponible | Construible | Parcial | Proxy | No disponible | Total |
|---|---|---|---|---|---|---|
| CS | 2 (PCS.EE, PCS.EP) | 5 (DSI.MD, DSI.MS, VRC.ET, RPO.OR, RPO.OT) | 1 (VRC.EV) | 0 | 0 | 8 |
| PC | 1 (I2.V2) | 1 (I2.V1) | 7 (I1.V1, I1.V2, I4.V1, I4.V2, I5.V1, I5.V2, I6.V2) | 2 (I3.V1, I3.V2) | 1 (I6.V1) | 12 |
| PDM | 7 (las 7 metas) | 1 (ILE.S) | 2 (CD.O, ILE.A) | 2 (CS.O, CM.O) | 2 (CA.O, CP.O) | 14 |
| PP | 0 | 3 (EPP.OPP, EPP.APP, CEP.PE) | 2 (CEP.PEC, CMA.NA) | 0 | 0 | 5 |
| **Total** | **10** | **10** | **12** | **4** | **3** | **39** |

Lectura: de los 10 "Disponible", **7 son metas precargadas**; solo **3 valores observados** están publicados directamente (PCS.EE, PCS.EP, I2.V2). Ningún índice podría calcularse hoy sin construcción adicional; ICPD-6 queda bloqueado por CA.O y CP.O; ISPC por I3 (Proxy) y el universo de procesos; IREPP por CEP y CMA parciales. 2026 no tiene cierre anual (corte de datos: 2.º trimestre).

---

## 8. Inconsistencias y riesgos detectados

### 8.1 En la matriz (Documentos guias/…xlsx)
1. **Conteos:** cuadran (39 variables; 19 + 4 = 23). Pero las fórmulas usan 42 símbolos: PA y MP (PP_Proyectos) y SC (calculado) no están en Registro. Correcto por diseño; documentarlo para no reportar "faltantes".
2. **Colisión de códigos:** el prefijo de sistema `CS` (Control social) coincide con el indicador `PDM.CS` (aseguramiento subsidiado), y en la fórmula de ICPD-6 "CS" significa PDM.CS. Igual con símbolos de variable: `MS` = metas seleccionadas (CS.DSI.MS) y meta Saber (PDM.ILE.MS); `M` y `O` se reutilizan. Los IDs completos son únicos; exigir siempre el ID completo en los entregables.
3. **Paginación de las citas del PDM no verificable:** la Matriz cita "p. 518, 522–523, 553–554, 599, 607, 628, 633 de la copia del Concejo". La copia del Concejo en el repo tiene 314 páginas; en el PDM de la Alcaldía (807 p.) los indicadores están en pp. 602, 609, 650, 652, 706, 717, 742 y 748, y en el DTS de dic-2025 en pp. 640, 646, 683, 685, 738, 748, 775, 781. Reemplazar las citas por documento + página verificados.
4. **Programa de movilidad:** la Matriz asigna mortalidad vial al "programa 3.2.2"; en el Apéndice 9.1.2 está en Línea 3, Nodo 2, Programa 3, y el proyecto de seguridad vial es el 3.2.3.2 → probablemente programa **3.2.3**. Verificar.
5. **Línea base sin año:** la Matriz dice que CM (10,2) y CP (75) no tienen año en la fila del plan; es cierto en el PDM y el DTS, pero el Apéndice 9.1.2 les asigna **2023**. Registrar la discrepancia en Observaciones.
6. **PCS – programación anual:** la Matriz advierte "4 estrategias en el cuatrienio; no asumir una por año", pero el plan indicativo programa oficialmente 1 por año (2024–2027). La advertencia es conservadora pero ya existe programación anual oficial.
7. **Metas sin horizonte anual:** las 7 metas son meta 2027; se repiten en 2024–2026. Para metas "creciente" (CD 6, MA 19, MS 14, CP 100) la comparación de años intermedios contra la meta final castiga los primeros años. El PDM no publica metas anuales de resultado, así que la "evaluación anual" de la Guía (9) no tiene insumo oficial.
8. **Periodicidad sin opción en las listas:** PP declara "Trimestral si hay reporte", pero las listas de periodo de Registro y PP_Proyectos solo admiten Anual / Semestre 1 / Semestre 2, y Resultados exige Anual para PP. Lo trimestral no se puede registrar.
9. **Territorio sin catálogo:** la columna H de Registro, la C de PP_Proyectos y el selector H4 de Resultados no tienen lista de validación; cualquier variación ("Zona 2" vs "zona 2") rompe la clave. Falta el catálogo oficial de zonas (y hay ambigüedad 12 vs 13 zonas, ver PC.I2.V2).
10. **Corte exacto:** Resultados solo usa un valor si su fecha de corte (N) es **igual** al corte común (31-dic). Datos con corte de mes (TerriData), fecha de examen (ICFES) o cierre de evaluación (25-feb-2026) quedarán N/D si no se normalizan; la Guía no explica cómo tratar el "último año completo disponible" (PDM.CS, CA, CP, CD).
11. **Estados vs escala de disponibilidad:** la lista de estados incluye "No disponible", "No evaluable", "No aplica", que la Guía no define (solo Pendiente/Validado y N/D/NE/Revisar). La escala de esta tarea no tiene columna propia; usar Observaciones.
12. **Nombre de PP.CMA.NA distinto:** Matriz "Número de proyectos con meta de atención aplicable"; Registro "…aplicable en el universo completo y verificado".
13. **Etiquetas de entradas:** IREPP dice "3 variables" (son 3 puntajes); ICRCS e ISPC dicen "4 entradas"; ICPD-6 no indica número.
14. **EPP "excluyendo cofinanciación":** ninguna fuente pública identifica la fuente PP (CUIPO no tiene marcador; ROCA usa códigos de fuente sin diccionario). La regla puede no ser verificable.
15. **Fuentes metodológicas ausentes:** la Guía cita cuatro anexos Word ("Matriz de indicadores de control social.docx", "…participacion ciudadana Envigado.docx", "…cumplimiento PDM Envigado.docx", "…presupuesto participativo.docx") que **no están en el repositorio**. Tampoco el DOCX OPCEN (pendiente del knowledge graph).
16. **Registro con 2027:** las 39 × 4 filas incluyen 2027, que no es observable; no es error, pero no forma parte del periodo objetivo.
17. **PDM.CS y TerriData:** la Matriz prohíbe usar cobertura total del SGSSS, pero no advierte que el indicador homónimo de TerriData ("Cobertura del régimen subsidiado", ~12 %) usa otra definición. Riesgo alto de uso indebido.

### 8.2 En el inventario y los archivos
- **Ventana temporal mal catalogada:** 11 archivos tienen un año en el nombre que no coincide con `ventana_temporal`. Los 5 críticos son evaluaciones del PDM 2024–2026 catalogadas como "2022": `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalcuartotrimestre202430012025-1.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalprimertrimestre202513052025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionplandesarrollomunicipalsegundotrimestre202522072025.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluacionprovisionalplandesarrollomunicipalcuartotrimestre202527012026.pdf`, `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/planindicativoproductos20261trimestre.xlsx`.
- **PDF de evaluación del PDM con poco texto (presentaciones con gráficos):** requieren lectura visual u OCR; las cifras de "Indicadores de resultado" probablemente están en imágenes.
- **Datos personales:** los cuadros de preguntas y respuestas y el CSV de veedurías incluyen nombres y correos. Solo contar; no transcribir.
- **Lector de PDF:** no hay `pdftotext`; `pypdf` existe solo en un scratchpad de otra sesión (`/tmp/claude-1000/-workspaces-Observatorio-Participaci-n-Ciudadana/5c988072-7b80-4261-aef7-cf04a02bb425/scratchpad/pylib`), ruta efímera. Los ejecutores deben confirmar que está disponible o reinstalarla en su scratchpad.
- **Cobertura temporal 2026:** corte de descarga 25-sep-2026; los insumos de 2026 llegan al 2.º trimestre (plan indicativo y plan de acción al 30-jun/30-jul-2026; ejecución de egresos al 30-jun-2026). Ningún valor anual 2026 puede tener corte 31-dic-2026.

---

## 9. Orden de trabajo sugerido para la Fase 2 (Ejecutor)

1. **PDM (alto rendimiento):** confirmar formalmente las 7 metas (sección 4) y revisar visualmente las evaluaciones IV trim. 2024 y 2025 para los indicadores de resultado (CS.O, CM.O, CA.O, CP.O, CD.O, ILE.A, ILE.S). Si no aparecen, dejar la clasificación de la sección 6.3.
2. **CS:** extraer del plan indicativo 2026-2T los consecutivos 544 y 545 (y el de "Eventos de rendición pública de cuentas realizados"); enumerar eventos de rendición 2024–2026 y contar intervenciones/respuestas en los dos cuadros de preguntas (sin datos personales).
3. **PP:** construir el catálogo PP por vigencia de ejecución (Proyectos_PP_AAAA + proyectosdeinversion2024) y verificar si sus códigos aparecen en el plan de acción/ROCA 2024–2026 (APROPIACIÓN DEFINITIVA, OBLIGACIÓN ACUMULADA). No usar el puente por sufijo BPIN.
4. **PC:** decidir con los docentes los universos (procesos cerrados, instancias, eventos municipales) antes de contar; mientras tanto, reportar Parcial con la lista de evidencias.
5. **Web (solo lo marcado "URL por buscar"):** ICFES clasificación de planteles; BDUA/MinSalud; ANSV/INML; portal EFPC; micrositio de rendición de cuentas; Índer.
6. No modificar `Documentos guias/` ni el knowledge graph en esta fase; trabajar sobre `outputs/disponibilidad/trabajo/matriz_copia.xlsx` si se necesita escribir.
