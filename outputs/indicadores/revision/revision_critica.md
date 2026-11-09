# Revisión crítica independiente: cálculo de indicadores viables (tanda 4)

**Autor:** Agente Analizador. **Fecha:** 9-nov-2026. **Alcance:** tandas 1-3 y CS.ICRCS, en el estado del repositorio a esta fecha (`resultados_indicadores.csv` con 999 filas).
**Método:** recalculé cada valor leyendo yo mismo el documento fuente con pymupdf u openpyxl, sin usar los scripts. Comparé las fórmulas con `Resultados!L8:N30` y con la hoja `Matriz` del archivo original, y apliqué las reglas 3-20 de la hoja `Guía`. Revisé la coherencia con `config/supuestos.yaml`. **No modifiqué scripts ni salidas.**
**Clasificación de hallazgos:** **ERROR** = corregir antes de entregar · **ADVERTENCIA** = documentar · **OK**.

---

## 0. Resumen ejecutivo

- **Recálculo a mano:** 23 comprobaciones de 13 indicadores o escenarios: 15 valores de indicador, 5 observaciones de RPO y 3 proyectos de PP 2024. **Las 23 coinciden** con el script.
  - PDM.CM coincide en la aritmética. La cifra observada (7,25) sigue siendo una cita: Forensis 2024 no está en el repositorio. La reconstruí con 18 muertes y la población DANE 2024 (248.304), lo que da 7,249.
- **Fórmulas y semáforos:** recalculé 975 filas (todas salvo las 24 de PDM.ILE, que son N/D). De ellas, 951 reproducen exactamente la regla de la Matriz, incluidos los topes, Revisar y NE. Las otras 24 son las filas `pc_i3_v1_mayor_v2=tope_100`, una sensibilidad declarada en el YAML. ICRCS no redistribuye pesos. El EPP municipal es Σ OPP / Σ APP, no un promedio de zonas.
- **Escenario por defecto:** coincide con el YAML en el 100 % de las filas, y todas las opciones usadas existen en el YAML.
- **"Validado" y datos personales:** no hay ningún "Validado" en estados. Todo dice Provisional, Pendiente, No disponible o No evaluable. No hay datos personales: los 6 "celulares" detectados son montos en pesos.
- **5 ERRORES**, todos de presentación o documentación; ninguno cambia un puntaje del escenario por defecto:
  - E1. Semáforo asignado a PP.EPP 2026 S1 fuera del contexto de la Matriz.
  - E2. Nombre equivocado de la instancia CMSSS en PC.I5.
  - E3. 515 filas de PP.EPP sin `nombre` ni `tipo`.
  - E4. Las "envolventes" de CS.ICRCS no son envolventes, y el texto de 2025 es falso.
  - E5. Descripciones desactualizadas en `supuestos.yaml`.
- **Advertencias de mayor peso:**
  - A1. Todo ICRCS numérico de 2024 exige apartarse de una regla literal de la Matriz (VRC = N/D si hay presencia desconocida). El rango alcanzable va de 43,2 (Rojo) a por lo menos 73,8 (Amarillo).
  - A2. El APP de PP 2024 incluye recursos que probablemente no son de PP: 15 cruces aproximados, hasta +60,8 %, y $627,2 millones del rubro 2.3.7.06 "Disminución de pasivos – Financiación de déficit fiscal".
  - A3. PC.I2 2025 asigna a las zonas 2-13 un dato agregado. La fuente dice "la totalidad de las zonas".

---

## 1. Recálculo a mano de una muestra

| # | Indicador / valor | Valor del script | Valor recalculado | Documento y página / celda | ¿Coincide? |
|---|---|---|---|---|---|
| 1 | **CS.PCS 2024** | EE=1, EP=1 → 100 Verde | EE=1, EP=1 → MIN(100, 100·1/1) = 100 Verde | PI26 `evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx`, hoja `Plan-Indicativo-Productos`, fila 547 (CONSECUTIVO 544, indicador AM "Estrategias de promoción del control social"): AS (META PROGRAMADA 2024) = 1; BM (AVANCE EJECUTADO 2024) = 1 | Sí |
| 2 | **CS.DSI 2024** | MD=0, MS=7 → 0 Rojo | MS = 7 (filas 16, 24, 71, 74, 119, 129, 146 del Apéndice 9.1.2). MD = 0: el POT, la única cifra de la misma unidad con fecha verificable, se publicó el 24-oct-2025, después del límite del 31-mar-2025 (Guía 17); Forensis no tiene fecha verificable; las demás metas no tienen cifra → 0/7 = 0 Rojo | POT diagnóstico, dimensión funcional, PDF pp. 86-87 (EPV 2024); `soporte_conteos` CS-DSI-2024-001…007 | Sí (con reserva: si Forensis 2024 hubiera salido antes del 31-mar-2025, DSI = 14,3) |
| 3 | **CS.RPO 2024** | OR=16, OT=17 → 94,12 Verde | El cuadro tiene 22 intervenciones; 5 son felicitaciones o agradecimientos con respuesta "N/A" (n.º 3, 4, 5, 7, 8) → OT = 17. La n.º 9 solo remite ("Se da respuesta en el evento…") → OR = 16 → 94,12 Verde. Publicado el 12-dic-2024 (metadatos del PDF), antes del cierre | QA24 `129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf`, pp. 1-24 | Sí |
| 3a | RPO obs. n.º 1 (CS-RPO-2024-001) | pregunta; fondo estricto = 1 | "¿tendrá Envigado una casa del cuidado?" → "No vamos a contar con una casa de cuidado… sí se cuenta con una estrategia…": negativa motivada = fondo | QA24 p. 1 | Sí |
| 3b | RPO obs. n.º 3 (-003) | felicitación; fuera de OT | "Alcalde Dios lo bendiga… ha cumplido con todo…", respuesta N/A | QA24 p. 2 | Sí |
| 3c | RPO obs. n.º 9 (-009) | sugerencia; estricto 0, laxo 1 | "Sugiero… ayudar con mejoras de hogar…" → "Se da respuesta en el evento en la intervención de la Secretaria de Bienestar Social": es una remisión sin contenido | QA24 pp. 4-5 | Sí |
| 3d | RPO obs. n.º 21 (-021) | pregunta; fondo 1 | Plusvalía → "…se llevarán a cabo los procesos necesarios para implementar este tributo…; se incluirá en el cambio del POT": compromiso = fondo | QA24 p. 23 | Sí |
| 3e | RPO obs. n.º 22 (-022) | pregunta; tema_pdm 0; fondo 1; dudoso | "¿dónde se puede encontrar la calificación de desempeño…?" → "reposan en las respectivas historias laborales…": responde el "dónde" | QA24 pp. 23-24 | Sí (bien marcada como dudosa) |
| 4 | **CS.VRC 2024** (componentes) | ET=3; EV desconocido → N/D | Universo `alcalde_mas_sectoriales_con_acta`: (a) Salud 21-mar-2024: acta-informe p. 1 "Asistentes: … integrantes de las veedurías en salud" → sí; (b) Salud 26-nov-2024: acta p. 12 "Continúan sus presentaciones los integrantes de las veedurías…" → sí; (c) Alcalde 2024: solo cuadro de preguntas, sin asistentes ni mención → desconocida. ET = 3, EV = 2 + 1 desconocido → N/D (regla de la Matriz) | ACS24a p. 1 y p. 17; ACS24b pp. 1, 8-9, 12; QA24 p. 1 | Sí |
| 5 | **PC.I2 2025** | 12/13 → 92,31 Verde | "Se cumplió con la realización de 12 asambleas zonales durante el mes de agosto de 2025" → V1 = 12; V2 = 13 zonas (PDM p. 558: "12 Comités Zonales… de las 13 zonas establecidas") → 92,31 Verde (≥ 80) | G25 `2025informegestiondependencias.pdf`, PDF p. 92; PDM p. 558 | Sí (ver A3) |
| 6 | **PC.I3 2025** | 20/24 → 83,33 Verde | Culminados: 3 diplomados + 9 seminarios + 9 cursos = 21; no ejecutados: 1 diplomado + 3 seminarios (los semilleros se excluyen) → programados 25; sin el diplomado "Gestión Pública" (para servidores) → 20/24 = 83,33 Verde | G25 PDF pp. 87-88 | Sí (ver A8) |
| 7 | **PC.I5 2025** | 8/8 → 100 Verde | PROGRAMACIÓN fila 58 (E5Ld-M01-A01, reuniones del CMSSS) = 4 y fila 59 (E5Ld-M02-A01, COPACO) = 4; SEGUIMIENTO filas 58 y 59 = 4 y 4. La fila 60 (A02) es adicional → 8/8 = 100 | PPSS25 `pps190ppps20251231mu000000005266.xlsx`, hojas PROGRAMACIÓN col. I y SEGUIMIENTO col. G | Sí |
| 8 | **PDM.CA 2024** | O=5,75, M=3,88 → 100 Verde | Tabla 39: EPV urbano 1.382.762 m² / 240.655 hab. = 5,746 → 5,75; meta "Espacio público verde… 3,88 (2023) 3,88" → MIN(100, 100·5,75/3,88 = 148,2) = 100 Verde | POTEP PDF pp. 86-87; PDM PDF p. 717 | Sí |
| 9 | **PDM.CM 2024** | O=7,25, M=8 → 100 Verde | M = 8 (PDM p. 706: "Tasa de mortalidad por siniestros viales… 10,2 / 8"). O: Forensis no está en el repositorio. 18 / 248.304 × 100.000 = 7,249 (DANE post-COVID 2024, Total = 248.304, fila 2232) → MIN(100, 100·8/7,25 = 110,3) = 100 Verde | PDM p. 706; `06_descargas_indicadores/dane/DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx` | Sí en la aritmética; **O no verificable** contra Forensis |
| 10 | **PP.EPP 2025 municipal** | 20.684.427.720 / 20.928.690.218 → 98,833 Verde | Filas con "ZONA n" en NOMBRE PROYECTO: 61 proyectos sin el 0553. Σ col. N (APROPIACIÓN DEFINITIVA) = 20.928.690.218; Σ col. O (EJECUCIÓN OBLIGACIÓN) = 20.684.427.720 → 98,833 | PA25 `evaluaciondefinitivaplanaccionivtri.xlsx`, hoja `PlanAcción27.01.2026` | Sí |
| 10a | PP.EPP 2025 zona 2 | 1.436.299.907 / 1.475.002.700 → 97,376 | Igual | ídem | Sí |
| 10b | PP.EPP 2025 zona 9 | 2.734.256.520 / 2.809.659.342 → 97,316 | Igual | ídem | Sí |
| 11 | **PP.EPP 2024 municipal** | 18.357.265.731 / 18.601.236.386 → 98,688 Verde | Rubros de 47 dígitos con sufijo igual al código de proyecto, para los 50 proyectos con `incluido_completa=True`: Σ Apropiación definitiva = 18.601.236.386; Σ Obligación = 18.357.265.731 → 98,688 (0 diferencias proyecto por proyecto) | HAC24 `ejecucion-presupuestal-egresos-al-31122024.pdf` | Sí |
| 11a | Proyecto 0742 (zona 2, exacta) | APP 346.571.149; OPP 346.571.149 | APP 346.571.149; OPP 346.571.149; EFPC 346.571.149 | HAC24; EFPC24 p. 1 | Sí |
| 11b | Proyecto 0346 (zona 3, aprox.) | APP 424.965.037; OPP 424.813.971 | Igual (5 rubros; EFPC = 369.555.287, +15 %) | HAC24; EFPC24 p. 1 | Sí (ver A2) |
| 11c | Proyecto 0714 (zona 13, aprox.) | APP 685.162.170; OPP 655.187.630 | Igual (23 rubros; EFPC = 425.998.943, +60,8 %) | HAC24; EFPC24 p. 4 | Sí (cruce débil, ver A2) |
| 12 | **PP.EPP 2026 S1 municipal** | 5.702.528.148 / 24.519.738.737 → 23,257 | 66 proyectos "ZONA n" de la hoja `Plan-Acción`: Σ APP = 24.519.738.737, Σ OPP = 5.702.528.148. **Además**, con los rubros de Hacienda al 30-jun-2026 por sufijo, las mismas cifras al peso (66/66 proyectos, 0 diferencias) | PA26 `evaluaciondefinitivaplanaccion2026segundotrimestre20260730.xlsx`; HAC26 `ejecucion-presupuestal-egresos-2-al-30062026.pdf` | Sí (ver E1 y A10) |
| 13 | **CS.ICRCS 2024, escenario `vrc_evento_sin_mencion=cero_siempre`** | 65,196 Amarillo | VRC = 2/3 = 66,667 (el Alcalde pasa a ausencia). ICRCS = (0 + 100 + 66,667 + 94,118)/4 = 65,196 → Amarillo (50 a < 75) | Componentes 1-4 | Sí |

**Total: 23 comprobaciones, 23 coincidencias.** La única salvedad es que la cifra observada de PDM.CM no es verificable contra su fuente primaria.

---

## 2. Fórmulas, semáforos y reglas de la Guía

Recalculé por programa las 999 filas de `resultados_indicadores.csv` con las fórmulas transcritas de `Resultados!L8:L30` y `M8:M30`.

| Verificación | Resultado | Clase |
|---|---|---|
| Puntaje = fórmula de la Matriz para CS.DSI, PCS, VRC, RPO, PC.I1-I6, PDM.CS/CA/CP/CD/CM y PP.EPP/CEP | De las 975 filas verificables (se excluyen las 24 de PDM.ILE, que son N/D), 951 son idénticas, incluido CS.ICRCS. Las 24 filas `pc_i3_v1_mayor_v2=tope_100` dan 100 donde la Matriz da "Revisar"; es una sensibilidad declarada en el YAML | OK |
| Umbrales de semáforo: CS.DSI/RPO 90/70; PCS/VRC/ICRCS 75/50; PC.I2 80/50; I3/I5 75/50; PDM verde solo en 100; PP de 4 colores 90/75/50 | Todos coinciden con la Matriz y la Guía 16. Se clasifica sin redondear (`puntaje`; `puntaje_1d` solo para mostrar) | OK |
| Tope MIN(100): solo CS.PCS y PDM | Correcto. DSI, VRC, RPO, PC y PP llevan "Revisar" si el numerador supera al denominador | OK |
| PDM.CM inverso: `IF(O=0,100,MIN(100,100·M/O))`, con O = entrada 1 y M = entrada 2 | Correcto (F21 = O, G21 = M) | OK |
| N/D frente a NE (Guía 7): zona 1 de PP = NE (APP = 0); faltantes = N/D | Correcto; ambos van en Gris y sin puntaje | OK |
| Índices sin redistribuir pesos (Guía 15): ICRCS solo con 4 entradas numéricas | Correcto: el defecto 2024 da N/D aunque haya 3 componentes | OK |
| PP municipal = Σ OPP / Σ APP (Guía 13), sin promediar zonas | Verificado en 2024, 2025 y 2026 S1: el municipal coincide con la suma de los proyectos y no con un promedio de zonas | OK |
| Contexto: CS y PDM solo Anual + Municipio; PC.I2 solo Municipio | 0 filas con puntaje fuera de contexto | OK |
| **PP solo Anual (Guía 14; `Resultados!L27` exige E4 = "Anual")** | 135 filas de PP.EPP 2026 "Semestre 1" tienen puntaje y **semáforo de desempeño** (Rojo en todas, municipal y zonas). La columna `alerta` dice "fuera del contexto de la Matriz", pero el color Rojo compara medio año contra umbrales anuales | **ERROR E1** |
| `puntaje_regla_matriz` (regla estricta, solo Validado) | Es "N/D" en todas las filas, lo cual es correcto porque todo está Pendiente | OK |
| Guía 17: convocatoria ≠ presencia; acuse ≠ respuesta | Aplicado: n.º 9 de QA24 no cuenta; la invitación de ACS24a p. 15 no se usa como evidencia | OK, pero ver A5 |
| Matriz VRC: "Si hay casos con presencia desconocida, VRC = N/D" | El defecto lo cumple. El único ICRCS 2024 numérico (`cero_siempre`) la contradice abiertamente | ADVERTENCIA A1 |
| Matriz CA: "No sustituir por espacio público efectivo total" | Por defecto se usa el EPV urbano, lo cual cumple. El escenario `pot_epv_total` (69,41, incluye sistemas orográfico e hídrico) no es comparable con la meta urbana 3,88 | ADVERTENCIA A9 |
| Matriz EPP: "Misma vigencia, zona, corte y fuente de financiación… separar cofinanciación" | 2025 y 2026 cumplen (ICLD; mismo corte). En 2024 hay rubros que probablemente no corresponden a PP | ADVERTENCIA A2 |

---

## 3. Trazabilidad de valores por defecto y soporte de conteos

Revisé las 213 filas de `valores_variables.csv` del escenario por defecto que tienen valor.

| Hallazgo | Detalle | Clase |
|---|---|---|
| Todas tienen `fuente`, `pagina_tabla`, `fecha_corte` y `fecha_consulta`, y enlace o ruta | 0 huecos en esos campos | OK |
| `fuente` es una ruta de archivo y no un nombre de publicación | 84 filas de PP.EPP y 52 de PP.CEP. Ejemplo: `data_raw_IIC/02_ejecucion…/ejecucion-presupuestal-egresos-al-31122024.pdf`. Esa ruta pasa tal cual a `Registro!K` de `matriz_diligenciada.xlsx`; la Guía 3 pide el "nombre de la fuente en K" | ADVERTENCIA A11 |
| `fecha_publicacion` vacía | En casi todas las filas (PP.EPP 84, PDM 40+, CS 8, PC 13). Es crítica en CS, porque la ventana documental depende de ella. Hoy solo vive en `soporte_conteos` o en las tablas de clasificación | ADVERTENCIA A12 |
| Conteos sin filas en `soporte_conteos` | CS.PCS.EE y EP (2024, 2025 y 2026 S1): ninguna fila por estrategia. PC.I2.V2 (13 zonas): `ids_soporte` vacío. Variables auxiliares AUX.PDM.CM.MUERTES, AUX.PDM.CS.AFIL_RS, AUX.PDM.ILE.A.*, AUX.PDM.CD.DISCIPLINAS y PP.CEP.PE (52 filas): sin `ids_soporte`. El soporte de PP.CEP está en `clasificacion_pp_entregas.csv`, sin vínculo por id | ADVERTENCIA A12 |
| PDM.CM.O y AUX.PDM.CM.MUERTES | `documento_ruta_local` apunta a `…/inmlcf/Forensis_2024.pdf`, que no existe | ADVERTENCIA (pendiente conocido) |
| `pp_proyectos.csv`: columnas `pec_todas_filas_dic_100`, `pec_solo_31dic` y `pec_ultima_fila` | Vacías en las 181 filas; los valores (49/45/50) están en `clasificacion_pp_entregas.csv`. Son columnas muertas | ADVERTENCIA |
| CS-VRC-2024-008 | `ruta_local` termina en "[archivo con nombre de persona omitido]" y el enlace apunta solo a la carpeta 002627, así que un tercero no puede ubicar el archivo. Sugerencia: usar el ID numérico del archivo en `content/files` sin el nombre | ADVERTENCIA |
| `clasificacion_rpo` (56), `clasificacion_vrc` (17), `clasificacion_pp_2024` (51), `pp_proyectos` (181) | 0 vacíos en documento, página, enlace, ruta y clasificación | OK |

---

## 4. Coherencia entre escenarios

| Verificación | Resultado | Clase |
|---|---|---|
| ¿`parametros` de toda fila `es_defecto=True` = `defecto` de `supuestos.yaml`? | Sí, 36 de 36 parámetros, en todas las filas | OK |
| ¿Todos los `param=opcion` usados existen en el YAML? | Sí, 0 opciones desconocidas. El `escenario_id` concuerda con `parametros` y no contiene desviaciones iguales al defecto | OK |
| Parámetros nunca variados | `regla_estado_calculo` (está cubierto por la columna `puntaje_regla_matriz`), `pp_cep_regla_agregacion` y `pp_cma_aplicabilidad` (solo afectan variables; las alternativas están en `clasificacion_pp_entregas`) | OK |
| **Descripciones del YAML desactualizadas** | `pp_2024_metodo`: líneas 532-533 dicen "47 de 50 proyectos" y "34 proyectos con correspondencia exacta"; los resultados dan **50 de 51** y **35 exactos**. `pp_cep_regla_agregacion`: línea 561 dice "Ultima fila por indicador (2025: 55)"; `clasificacion_pp_entregas.csv` da **50** | **ERROR E5** |
| Escenarios creados fuera del YAML | `control_institucional_545_546` (CS.VRC y CS.ICRCS) y `envolvente_min/max` (ICRCS). El de control está bien rotulado: su `parametros` lista las 4 desviaciones (`institucional_pi_545_546`, `cero_siempre`, `hasta_fecha_consulta` y la observación) y `es_defecto=False`. Las envolventes no son lo que dice su nombre (ver E4) | OK / **ERROR E4** |
| `pp_fuente_calculo` en 2024 | La fila por defecto de 2024 dice `pp_fuente_calculo=plan_accion_roca`, pero se calculó con Hacienda (el YAML anota "2024 solo admite hacienda"). El escenario `pp_fuente_calculo=hacienda` duplica el defecto en 2024. La observación sí lo dice | ADVERTENCIA |
| `pp_2026_opp_fuente` | Las dos opciones dan la misma cifra al peso, así que el escenario es redundante (y está bien que lo sea) | OK |
| `plan/plan_calculo.md` | Valores de control desactualizados: §1.2 fila PP.EPP ("50 proyectos; 47 con correspondencia, 34 exactas") y línea 139 ("17.199.119.162 / 17.332.821.550 ≈ 99,2") | ADVERTENCIA (documento histórico; anotar la fe de erratas) |
| `ESTADO.md` y `README.md` | ESTADO cita `resultados_indicadores.csv` (939) cuando hoy son 999, y deja la tanda 4 como "Pendiente". El "orden probado" del README omite `cs_vrc`, `cs_rpo`, `cs_segunda_lectura`, `pp_epp_2024` y `cs_icrcs` | ADVERTENCIA |

---

## 5. "Validado", estados y datos personales

| Verificación | Resultado | Clase |
|---|---|---|
| Estados en CSV | `resultados.estado` = Provisional (999/999; `provisional=True` en todas); `valores_variables` = Pendiente 2754, No disponible 776, No evaluable 48; `pp_proyectos` y `clasificacion_pp_2024` = Pendiente | OK |
| `matriz_diligenciada.xlsx` | `Registro!J`: Pendiente 259, No disponible 91, No evaluable 9; `PP_Proyectos!L`: Pendiente 178. "Validado" solo aparece en las fórmulas originales (`Registro!S`, `PP_Proyectos!O`) y en 3 celdas de texto de la Guía | OK |
| Correos, celulares y cédulas en `outputs/indicadores/**` (CSV, MD, XLSX) | 0 correos y 0 cédulas. Los 6 "celulares" de 10 dígitos que empiezan por 3 son montos de APP/OPP por zona (p. ej., 3.162.025.221) = falsos positivos | OK |
| Nombres de personas | Busqué los nombres que aparecen en QA24, ACS24b, PI26 (`EU`, responsable con correo) y en los metadatos de los PDF: 0 apariciones. Las citas usan "[ciudadano/a]" o "[funcionario/a]". Las fuentes sí contienen nombres, celulares y un enlace de OneDrive personal (PPSS25, SEGUIMIENTO col. L), y nada de eso se copió | OK |

---

## 6. ERRORES (corregir antes de entregar)

**E1. Semáforo de desempeño en PP.EPP 2026 "Semestre 1" (135 filas).**
- **Problema:**
  - `Resultados!L27` solo calcula PP con periodo Anual, y la Guía 14 dice "PP es anual".
  - El script fuerza el cálculo con `{"ignorar_contexto": True}` (`scripts/indicadores/pp_epp.py:283`) y además pinta el semáforo: Rojo en el municipal (23,26) y en las 12 zonas.
  - El propio `avance_1.md` admite que "no es comparable".
- **Corrección:**
  - En `pp_epp.py`, después de construir cada fila con `periodo != "Anual"` (llamada `C.fila_resultado` en las líneas 368-370), conservar `puntaje` y fijar `semaforo = "Gris"`.
  - Usar `disponibilidad = "Calculado (fuera de contexto)"` o mantener "Calculado" con `alerta` explícita "sin semáforo: periodo no previsto por la Matriz".
  - Aplicarlo también a `tratamiento_2026=anual_provisional`, porque es un corte de 30-jun rotulado como Anual.

**E2. Nombre incorrecto de la instancia CMSSS.**
- **Dónde:** `scripts/indicadores/pc_i5.py:28` la llama "Comité Municipal de Participación Social en Salud (CMSSS)". El error se propaga a `soporte_conteos.csv` (PC-I5-2025-001 a 004) y a `base_indicadores.xlsx`.
- **Valor correcto:** "Consejo Municipal de Seguridad Social en Salud (CMSSS)".
- **Fuentes:**
  - Acta de la primera rendición de Salud 2024, ACS24a p. 1: "Consejo Municipal de Seguridad Social en Salud –CMSSS-".
  - PPSS25, hoja PROGRAMACIÓN fila 58: meta "CTSSS constituido…", actividad "reuniones ordinarias de la instancia CMSSS".

**E3. PP.EPP sin `nombre` ni `tipo` en 515 filas de `resultados_indicadores.csv`** (también en `base_indicadores.xlsx`).
- **Corrección:** en `pp_epp.py:368`, pasar a `C.fila_resultado(...)` los argumentos `nombre="Ejecución presupuestal del presupuesto participativo (EPP)"` y `tipo="Indicador"`. Fuente: hoja `Matriz`, fila PP.EPP, columnas D y L.

**E4. Las "envolventes" de CS.ICRCS no son envolventes, y el texto de 2025 es falso.**
- **2024:**
  - `envolvente_min` = 65,20 y `envolvente_max` = 68,17 son el mínimo y el máximo de solo 2 escenarios: un factor a la vez más el de control.
  - Combinando opciones del YAML compatibles entre sí (misma ventana para DSI, VRC y RPO), el rango alcanzable es mucho mayor:
    - **Mínimo 43,18 (Rojo):** DSI 0 + PCS 100 + VRC 0 + RPO 72,73. Usa `vrc_universo=alcalde` con `cero_siempre` (0/1) y `rpo_filtro_pdm=todas`.
    - **Máximo de al menos 73,81 (Amarillo):** DSI 28,57 + PCS 100 + VRC 66,67 + RPO 100. Usa `hasta_fecha_consulta`, `cero_siempre` y `rpo_criterio_fondo=laxo_texto_respuesta`. Excluyo `dsi_lista_metas=productos_plan_indicativo`, que es solo de control.
  - Reportar 65,2-68,2 como "envolvente" subestima la incertidumbre y oculta que el semáforo puede ser Rojo.
- **2025:**
  - Las filas `envolvente_*` dicen "DSI=N/D, PCS=N/D, VRC=N/D, RPO=N/D", pero DSI 2025 = 0 y PCS 2025 = 100. El texto sale de `cs_icrcs.py:98-101`, que pasa `[None]*4`.
  - Tampoco es cierto que no haya escenarios calculables. Con dos factores (`cero_siempre` + `rpo_registro_incompleto=calcular_con_publicado`): ICRCS 2025 = (0 + 100 + 20 + 50)/4 = **42,5 (Rojo)**.
- **Corrección:** o se calcula la envolvente combinando los extremos por componente con restricción de ventana común, o se renombra como `min_evaluado` / `max_evaluado` con la leyenda "extremos de los escenarios de un factor; no es el rango posible". En los dos casos, el texto de 2025 debe listar los puntajes reales de DSI y PCS.

**E5. Descripciones desactualizadas en `config/supuestos.yaml`.**
- Línea 532: "47 de 50 proyectos" → "50 de 51 proyectos (rotulado aproximado; 35 exactos + 15 aproximados)".
- Línea 533: "34 proyectos" → "35 proyectos con correspondencia exacta". Fuente: `clasificacion_pp_2024.csv`, columna `tipo_coincidencia`.
- Línea 561: "(2025: 55)" → "(2025: 50)". Fuente: `clasificacion_pp_entregas.csv`, `pec_ultima_fila` = 50.
- El YAML lo mantiene el coordinador, así que la corrección le corresponde a él.

---

## 7. ADVERTENCIAS (documentar)

- **A1. ICRCS 2024 depende de apartarse de la Matriz.**
  - El único ICRCS numérico del escenario de un factor (65,2) exige `vrc_evento_sin_mencion=cero_siempre`, contrario a la regla literal de la Matriz.
  - Con el defecto, ICRCS = N/D, y es lo correcto.
  - Rango alcanzable: 43,2 (Rojo) a por lo menos 73,8 (Amarillo); ver E4.
  - Decisión de los profesores.
- **A2. El APP de PP 2024 incluye recursos que probablemente no son de PP.**
  - En 15 cruces "aproximados", la apropiación de Hacienda supera el valor EFPC: +60,8 % (0714), +32,5 % (0709), +19,9 % (0349), +17,3 % (0746), entre 10 % y 15 % (0346, 0347, 0348, 0758) y +9,7 % (0272).
  - Son proyectos con recursos que no vienen de PP, y la Matriz pide la "misma fuente de financiación".
  - Además, 2024 incluye $627.204.058 de rubros 2.3.7.06 "Disminución de pasivos – Financiación de déficit fiscal" (pagos de vigencias anteriores, ejecutados al 100 %). Sin ellos, EPP 2024 = 17.730.061.673 / 17.974.032.328 = **98,64** (frente a 98,69).
  - En 2025 y 2026 no aparece el rubro 2.3.7 en los proyectos de PP.
  - El Amarillo de la zona 10 (82,58) se debe sobre todo al 0709.
  - Con `solo_exactos`: 99,28.
- **A3. PC.I2 2025: la asignación a zonas es inferida.**
  - G25 p. 92 dice "12 asambleas zonales… cubriendo la totalidad de las zonas del municipio".
  - El script asigna una asamblea a cada zona de la 2 a la 13 ("realizada (inferida del agregado)").
  - Si "totalidad" significa 13 zonas, el resultado sería 13/13. Con `12_con_comite` también da 100.
- **A4. CS.RPO: la intervención n.º 10 tiene dos asuntos.**
  - En QA24 p. 5, la fila n.º 10 contiene dos asuntos con su propia respuesta: el subsidio a cuidadores y "Incentivos para los líderes comunales RTA: …".
  - Si se cuentan por separado: OT = 18, OR = 17 → 94,4 (Verde). No cambia el semáforo.
- **A5. CS.VRC: la presencia en ACS24a descansa en un párrafo de plantilla.**
  - En ACS24a, la presencia de veedurías se acredita con el párrafo "Asistentes: … integrantes de las veedurías en salud…", que se repite casi textual en ACS24b y ACS25 y en la p. 17 es un saludo.
  - En ACS24b la evidencia es fuerte: p. 12, las veedurías presentan hallazgos.
  - Con `lista_asistencia` todo queda en N/D.
  - Hay fechas contradictorias entre documentos: por ejemplo, la rendición de PP de la vigencia 2024 figura el 12-ene-2025 en la justificación y el 12-feb-2025 en G25 p. 92. No afecta el defecto.
- **A6. PDM.CM: cifra observada sin verificar.**
  - 7,25 y 18 muertes son cifras citadas, porque Forensis 2024 no está en el repositorio.
  - Falta verificar el criterio territorial: lugar del hecho o residencia (protocolo de la Matriz).
  - Con Policía y DANE post-COVID da 99,3 (Amarillo).
- **A7. CS.PCS: hay reparos del DAP.**
  - La fila 547 de PI26 trae observaciones del DAP (col. EL/EM): "No concuerda la información de ejecución anual con G+…".
  - El resultado coincide con lo reportado, pero hay que anotarlo antes de cualquier "Validado".
  - La Matriz advierte: "Meta de formulación: 4 estrategias en el cuatrienio; no asumir una por año". La programación oficial (AS:AV = 1/1/1/1) sí es de una por año, así que el uso es correcto.
- **A8. PC.I3 2025: el denominador es reconstruido.**
  - V2 sale de "culminadas + no ejecutadas", porque no hay programación fechada publicada.
  - Se excluyen los semilleros (7 planeados, 5 ejecutados, G25 p. 88). Si se cuentan como acciones formativas: 25/31 = 80,6 (Verde).
  - La exclusión del diplomado "Gestión Pública" como formación para servidores no está acreditada en la fuente.
- **A9. PDM.CA: comparabilidad de series.**
  - El EPV urbano del POT (5,75; población de cabecera 240.655) no es la serie de la línea base (3,88; el AMVA daba 4,66 en 2022).
  - El PDM rotula la unidad como "Metros cuadrados" (p. 717).
  - En `soporte_conteos` (CS-DSI-2024-006) se clasifica como `valor_misma_definicion` aunque se anota "serie distinta"; es una incoherencia de rótulo.
- **A10. PP 2026: encabezado engañoso en la fuente.**
  - En la hoja `Plan-Acción` de PA26, los encabezados de APROPIACIÓN y OBLIGACIÓN dicen "01-04-2026", pero las cifras coinciden al peso con Hacienda al 30-jun-2026 (comprobado en los 66 proyectos).
  - Hay que documentarlo en Observaciones de Registro.
- **A11. `fuente` es una ruta, no un nombre.** En PP, `fuente` (y por tanto `Registro!K`) es una ruta local. Sugerencia: "Secretaría de Hacienda – Ejecución presupuestal de egresos al 31-12-2024 (PDF)", con la ruta en otra columna.
- **A12. Huecos de trazabilidad.** `fecha_publicacion` vacía y conteos sin fila de soporte (ver §3).
- **A13. Segunda lectura sin hacer.** No está diligenciada (`segunda_lectura/muestra_vrc.csv`, 4 filas; `muestra_rpo.csv`, 12). RPO y VRC siguen siendo de un solo lector.
- **A14. Documentación de estado.** `ESTADO.md` (939 filas; tanda 4 pendiente) y el orden del `README.md` están desactualizados. `scripts/indicadores/cs_icrcs.py` no está versionado (git status `??`).

---

## 8. Decisiones metodológicas abiertas para los profesores

Valores del escenario por defecto → alternativas, en el nivel Municipio. "Revisar" y "N/D" no tienen color.

| Parámetro (YAML) / decisión | Indicadores afectados | Defecto → alternativas (resultado y semáforo) |
|---|---|---|
| `vrc_evento_sin_mencion` | CS.VRC, CS.ICRCS | VRC 2024: N/D → 66,7 Amarillo (`cero_siempre`); VRC 2025: N/D → 20 Rojo; ICRCS 2024: N/D → 65,2 Amarillo |
| `vrc_universo` (combinado con el anterior) | CS.VRC, CS.ICRCS | VRC 2024: 0 (Alcalde) · 50 (institucional 545/546, control) · 66,7 (defecto + cero) · amplio 3/7 = 42,9; ICRCS 2024: **43,2 a ≥ 73,8 (Rojo a Amarillo)** |
| `vrc_acreditacion_presencia`, `vrc_asignacion_anio` | CS.VRC | Todas las opciones dan N/D mientras persista la presencia desconocida |
| `ventana_documental_cs` | CS.DSI, CS.VRC, CS.ICRCS | DSI 2024: 0 → 28,6 (Rojo en ambos) |
| `dsi_lista_metas` | CS.DSI, CS.ICRCS | 7 metas: 0 · 10 metas: 0 · productos del plan indicativo: 100 (solo control). La Matriz pide "iniciar con 10 metas" |
| `dsi_fuente_valor` | CS.DSI | DSI 2025: 0 → 28,6 (con proxies) |
| `rpo_filtro_pdm` | CS.RPO, CS.ICRCS | 94,1 Verde → 93,8 Verde (temático) / **72,7 Amarillo** (todas, con felicitaciones) |
| `rpo_universo` | CS.RPO | 94,1 → 93,1 Verde (con 100 días) / **89,5 Amarillo** (todos los cuadros) |
| `rpo_criterio_fondo` | CS.RPO | 94,1 → 100 |
| `rpo_registro_incompleto` | CS.RPO 2025, ICRCS 2025 | N/D → **50 Rojo**; con `cero_siempre`, ICRCS 2025 = 42,5 Rojo |
| Separar la intervención n.º 10 de QA24 (fuera del YAML) | CS.RPO 2024 | 94,1 → 94,4 |
| `pc_i2_actividades` | PC.I2 | 2024: 100 → 92,3 (solo PP) |
| `pc_i2_evidencia` | PC.I2 | 2025: 92,3 → 100 (con programación POT); `acta_por_zona` → N/D en ambos años |
| `zonas_oficiales` | PC.I2, PP.EPP | PC.I2 2025: 92,3 → 100 (12 zonas). PP no cambia |
| `pc_i3_unidad` | PC.I3 | 2025: 83,3 → Revisar (personas 186/15); 2024: N/D → Revisar (95/10) |
| `pc_i3_excluir_servidores` | PC.I3 | 83,3 → 84,0 |
| Contar semilleros como acción formativa (fuera del YAML) | PC.I3 | 83,3 → 80,6 |
| `pc_i5_instancias` | PC.I5 | 100 → N/D (ampliado) |
| `horizonte_meta_pdm` | PDM.CA, PDM.CM (y todo PDM) | 100 → N/D (no hay metas anuales) |
| `pdm_ca_fuente` | PDM.CA, CS.DSI | 100 → 100 (EPV total, no comparable) / N/D (estricto) |
| `pdm_cm_fuente` × `poblacion_dane` | PDM.CM | 100 Verde → **99,3 Amarillo** (Policía + DANE post-COVID); el resto, 100 |
| `pp_2024_metodo` | PP.EPP 2024 | 98,69 → 99,28 (solo exactos) / N/D |
| Excluir rubros 2.3.7 y la parte no-PP de los cruces aproximados (fuera del YAML) | PP.EPP 2024 | 98,69 → 98,64 (sin 2.3.7); zona 10 Amarillo sujeta al cruce del 0709 |
| `pp_incluir_0553` | PP.EPP | 2025: 98,83 → 98,86; 2026 S1: 23,26 → 23,92 |
| `pp_fuente_calculo` | PP.EPP | 2025: 98,83 (Hacienda, igual) / N/D (CUIPO) |
| `pp_zona1_tratamiento` | PP.EPP, PP.CEP, PP.CMA (zona 1) | NE ↔ N/D |
| `tratamiento_2026` | PP.EPP, CS.PCS, PC.I2, PC.I3 | S1 provisional (23,3) / excluir / anual provisional (contradice la Guía 4) |
| `corte_fecha_registro` | Todos en la matriz diligenciada | En Python no cambia. En la hoja Resultados, `real` hace que todo quede en N/D (N ≠ N4) |
| `regla_estado_calculo` | Todos | Con la regla estricta de la Matriz, todo es N/D mientras no haya "Validado" |
| `pdm_ile_a_rango_edad`, `pp_cep_regla_agregacion`, `pp_cma_aplicabilidad`, `pc_universo_procesos` | Solo variables de indicadores no viables | ILE.A 2024: 15,58 (0-6) / 18,43 (0-5); PEC 2025: 49 / 45 / 50 de 61 |

---

## 9. Archivos revisados

- Salidas:
  - `outputs/indicadores/resultados_indicadores.csv`, `valores_variables.csv`, `soporte_conteos.csv`, `pp_proyectos.csv`
  - `clasificacion_rpo.csv`, `clasificacion_vrc.csv`, `clasificacion_pp_2024.csv`, `clasificacion_pp_entregas.csv`
  - `matriz_diligenciada.xlsx`, `base_indicadores.xlsx`
- Configuración y documentación:
  - `config/supuestos.yaml`, `outputs/indicadores/plan/plan_calculo.md`
  - `ESTADO.md`, `avance_1.md`, `avance_2.md`, `avance_3.md`
  - `scripts/indicadores/README.md`
- Matriz: `Documentos guias/Matriz de informacion de indicadores del observatorio ciudadano.xlsx`, hojas Guía, Matriz y Resultados; solo lectura.
- Fuentes leídas directamente (en `data_raw_IIC/`):
  - Instrumentos de planeación y seguimiento: PI26, PA25, PA26, PPSS25.
  - Ejecución presupuestal de Hacienda: HAC24, HAC25 (contraste del rubro 2.3.7) y HAC26.
  - Informes y planes: G25, PDM, POTEP.
  - Rendición de cuentas: QA24, ACS24a, ACS24b.
  - Población: DANE post-COVID.
- Script auxiliar de lectura de Hacienda, fuera del repositorio: `/tmp/claude-1000/-workspaces-Observatorio-Participaci-n-Ciudadana/8695b82b-4d33-43a8-8a3c-c302ba25f047/scratchpad/hac.py`.
