# Plan de cálculo de indicadores viables y construcción de variables (Fase 1 – Agente Analizador)

**Observatorio Ciudadano de Envigado (OCE)** · Etapa: organizar datos y calcular indicadores viables · Fecha del plan: 9-nov-2026 · Base: `outputs/disponibilidad/` versión 1.1 (corte de información 14-oct-2026).

**Encargo (profesor Francisco):**
- (a) organizar los datos de los indicadores calculables y construir las variables obtenibles de información secundaria;
- (b) calcular los indicadores viables mientras los profesores Francisco Correa y Eber Gutiérrez y la usuaria deciden qué indicadores se ajustan.

**Regla central:** el equipo técnico no toma ninguna decisión metodológica pendiente.
- Cada decisión es un parámetro de `config/supuestos.yaml` con opciones explícitas y un valor **por defecto PROVISIONAL**.
- Se calcula con todas las opciones (estrategia en §3.3).
- El borrador del archivo está en `outputs/indicadores/plan/supuestos_borrador.yaml`.

**Fuentes leídas para este plan:**
- `outputs/disponibilidad/ESTADO.md`;
- Excel `disponibilidad_brecha_indicadores.xlsx` (hojas Indicadores y Variables);
- `trabajo/revision_fase3.md` §2-§6 y `trabajo/notas_{CS,PC,PDM,PP}.md`;
- `trabajo/guia.txt` y `trabajo/matriz.csv`;
- la matriz original, leída con openpyxl: hojas Registro, PP_Proyectos y Resultados, con sus fórmulas y validaciones;
- `scripts/epp_pp_2025.py` y sus salidas;
- `data_raw_IIC/inventario_detallado_IIC.csv`;
- Apéndice 9.1.2 del PDM y encabezados del Plan Indicativo 2T-2026.

---

## 0. Confirmación del alcance

La hoja Indicadores del Excel confirma **11 indicadores calculables** en esta etapa: 1 Viable y 10 Viables con ajuste.

| Código | Veredicto | Años calculables (hoja Indicadores) | Años que calcula esta etapa |
|---|---|---|---|
| CS.DSI | Viable con ajuste | 2024; 2025 (2026 sin cierre) | 2024, 2025 |
| CS.PCS | **Viable** | 2024; 2025 (2026 al 2T: 0, sin cierre) | 2024, 2025 (2026 solo como variable, sin puntaje) |
| CS.VRC | Viable con ajuste | 2024; 2025 (condicionados) | 2024, 2025 |
| CS.RPO | Viable con ajuste | 2024 (2025 solo con acta integral) | 2024 (2025 solo con el parámetro `rpo_registro_incompleto=calcular_con_publicado`) |
| CS.ICRCS | Viable con ajuste (índice) | 2024 | 2024 (2025 da N/D en el escenario por defecto) |
| PC.I2 | Viable con ajuste | 2024; 2025 | 2024, 2025 |
| PC.I3 | Viable con ajuste | 2024; 2025 (opción A); 2025 aprox. (opción B) | 2024 y 2025 según la unidad |
| PC.I5 | Viable con ajuste (complementario) | 2025 (acotado a dos instancias) | 2025 (2024 si el PPSS 2024 trae la programación, a verificar) |
| PDM.CM | Viable con ajuste | 2024 | 2024 |
| PDM.CA | Viable con ajuste | 2024 (con salvedad) | 2024 |
| PP.EPP | Viable con ajuste | 2025; 2026 provisional; 2024 aproximado | 2024 (aprox.), 2025, 2026 (Semestre 1, provisional); Municipio + Zonas 1-13 |

**12 No viables por ahora: solo se construyen variables, sin calcular el indicador.**
- PC: I1, I4, I6 e ISPC.
- PDM: CS, CP, CD, ILE e ICPD-6.
- PP: CEP, CMA e IREPP.

---

## 1. Indicadores que se calculan

### 1.1 Reglas de cálculo comunes (réplica exacta de la hoja Resultados)

Las fórmulas se transcribieron de `Resultados!L8:L30` y `M8:M30`. Se usan los nombres de Resultados: F = entrada 1, G = entrada 2. Ningún valor se redondea antes de clasificarlo.

| Regla | Indicadores |
|---|---|
| Si falta una entrada (no numérica): N/D; si alguna entrada es "Revisar": Revisar | Todos |
| Denominador ≤ 0: **NE** | CS.DSI, CS.PCS, CS.VRC, CS.RPO, PC.I1-I6, PP.EPP, PP.CEP |
| Denominador (meta) ≤ 0: **Revisar** | PDM.* |
| Numerador > denominador: **Revisar** (sin tope) | CS.DSI, CS.VRC, CS.RPO, PC.I1-I6, PP.EPP, PP.CEP |
| Tope `MIN(100, ·)` | CS.PCS, PDM.CS, PDM.CA, PDM.CP, PDM.CD, PDM.CM (inverso), PDM.ILE (LA y LS por separado) |
| PDM.CM inverso | `IF(O=0, 100, MIN(100, 100*M/O))` |
| Solo contexto Anual + Municipio | CS.* (incluido ICRCS), PC.ISPC, PDM.* |
| Solo Anual (cualquier territorio) | PC.I1, PP.EPP, PP.CEP, PP.CMA, PP.IREPP |
| Solo Municipio (cualquier periodo) | PC.I2 |
| Sin restricción de contexto | PC.I3, PC.I4, PC.I5, PC.I6 |
| Índices: se calculan solo con todas las entradas numéricas; no se redistribuyen pesos | ICRCS = promedio de 4; ISPC = 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4; ICPD-6 = promedio(IBS, ITG); IREPP = 0,2·EPP + 0,4·CEP + 0,4·CMA (alerta si CEP o CMA < 50) |

Consecuencias del contexto:
- PP.EPP en "Semestre 1" (2026) da N/D en la matriz porque la fórmula exige Anual. El cálculo en Python lo reporta como **provisional, fuera del contexto de la Matriz** (ver §3, `tratamiento_2026`).
- PC.I2 por zona da N/D en la matriz. La Matriz pide "por zona: cobertura sí o no", que va solo en `soporte_conteos`.

### 1.2 Tabla de indicadores a calcular

Abreviaturas de documentos (ruta local bajo `data_raw_IIC/`). Para los archivos de `www.envigado.gov.co/sites/...` el enlace que se registra es el **espejo** `https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/...`, con la misma ruta.

| Sigla | Documento | Ruta local (`data_raw_IIC/…`) | Carpeta y archivo en `content/files/` |
|---|---|---|---|
| PI26 | Plan Indicativo de Productos, evaluación definitiva 2T-2026 (xlsx), hoja `Plan-Indicativo-Productos`, llave = col. A CONSECUTIVO | `02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | 003611/180505_… |
| AP91 | Apéndice 9.1 Plan Indicativo (hojas `Apéndice9.1.1`, `Apéndice9.1.2.`) | `05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` | 002340/116989_… |
| PDM | Plan de Desarrollo 2024-2027 (Acuerdo 017/2024) | `05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | 002562/128075_… |
| DTS | Documento técnico soporte PDM 22-dic-2025 | `…/plan-de-desarrollo-20242027/documentotecnicosoportepdm22122025.pdf` | 003109/155431_… |
| EV24 / EV25 | Evaluación PDM IV trim. 2024 / 2025 (PDF) | `02_…/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202431032025.pdf` / `…202525022026.pdf` | 002703/135123_… / 003149/157439_… |
| QA24 | Preguntas y recomendaciones, audiencia del Alcalde 2024 | `03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | 002595/129738_… |
| QA25 | Preguntas y respuestas, rendición 2025 (Alcalde) | `02_…/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf` | 002993/149610_… |
| QA100 | Preguntas y respuestas, 100 días (abr-2024) | `02_…/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf` | 002312/115597_… |
| QAS25 / QAS26 | Preguntas y respuestas, Salud 2025 / 2026 | `02_…/preguntas-y-respuestas-rendicion-de-cuentas-secretaria/preguntasyrespuestasrendicioncuentassaludi2025.pdf` / `02_…/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026preguntasrespuestasrendicioncuentas01.pdf` | 002690/134454_… / 003187/159305_… |
| ACS24a / ACS24b / ACS25 / ACS26 | Actas de rendición de Salud (21-mar-2024, 26-nov-2024, 19-mar-2025, 18-mar-2026) | `02_…/primera-rendicion-de-cuentas-de-la-secretaria-de-salud/20240321primerarendicincuentassalud-1.pdf`; `03_…/acta-audiencia-publica-de-rendicion-de-cuentas-2024/2024actarendicioncuentassecretariasalud02docx.pdf`; `03_…/audiencia-publica-de-rendicion-de-cuentas-202501-secretaria/2025actarendicioncuentassecretariasalud01final.pdf` (duplicado byte a byte en `02_…/primera-rendicion-de-cuentas-secretaria-de-salud-de/134500_…`: contar una vez); `02_…/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026actarendicioncuentassecretariasalud01.pdf` | 002317 / 002602 / 002690 / 003233 |
| CMD | Rendición del Comité Municipal de Discapacidad (cartas, memorias, preguntas y respuestas; 2024-2026) | `03_…/rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/`, `…-571162/`, `03_…/el-comite-municipal-de-discapacidad-presentara-su-rendicion/` | 002994, 003525… |
| VCA | Rendición de la política de víctimas 2024-2025 | `02_…/presentacion-rendicion-de-cuentas-vca-2024/`, `02_…/rendicion-de-cuentas-de-poblacion-victima-del-conflicto/`, `02_…/rendicion-de-cuentas-de-la-politica-publica-de-victimas/` | 002755, 002594, 002934 |
| ETN | Rendición de grupos étnicos, vigencia 2025 (evento del 3-sep-2026) | `02_…/rendicion-de-cuentas-de-la-poblacion-afro-indigena-y/` | 003671 |
| JUS | Justificaciones de no plan de mejoramiento (una por evento) | `rg 'justificacion' inventario_detallado_IIC.csv` (14 archivos) | varias |
| G24 / G25 | Informe de gestión por dependencias 2024 / 2025 | `02_…/ano-2025-298942/24informegestionpordependenciasalcaldiamunicipalenvigado2024.pdf` / `02_…/anon-2026/2025informegestiondependencias.pdf` | 002640/131997_… / 003058/152890_… |
| PPSS24 / PPSS25 | Seguimiento del plan de acción de la política de participación (xlsx), hojas PROGRAMACIÓN y SEGUIMIENTO | `02_…/seguimiento-al-plan-de-accion-de-la-politica-de-participacion/pps190spps20241231mu000000005266.xlsx` / `01_…/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | 002652/132552_… / 003452/172562_… |
| ANTE | Anteproyecto del PDM 29-feb-2024 (p. 53) | `05_…/plan-de-desarrollo-20242027/anteproyecto29022024.pdf` | 002315/115749_… |
| POTC | Cronograma de asambleas del POT 2025 | `03_…/programacion-de-las-asambleas-zonales-del-plan-de-ordenamiento/cronograma-asambleas-pot.pdf` | 002670/133472_… |
| POTEP | Diagnóstico POT, dimensión funcional, espacio público (pp. 86-87, tablas 38 y 39) | `05_…/documento-diagnostico-pot/8-dimension-funcional-_-espacio-publico.pdf` | 002925/146211_… |
| PA25 / PA26 | Plan de acción IV trim. 2025 (hojas `PlanAcción27.01.2026`, `ROCA07.01.26`) / 2T-2026 | `02_…/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx` / `02_…/seguimiento-a-proyectos-2026/evaluaciondefinitivaplanaccion2026segundotrimestre20260730.xlsx` | 003057/152845_… / 003611/180504_… |
| HAC24 / HAC26 | Ejecución presupuestal de egresos al 31-dic-2024 / al 30-jun-2026 (PDF de Hacienda) | `02_…/ano-2024-325557/ejecucion-presupuestal-egresos-al-31122024.pdf` / `02_…/ano-2026-358126/ejecucion-presupuestal-egresos-2-al-30062026.pdf` | 002660/132983_… / 003478/173890_… |
| EFPC24 / EFPC25 | Listados de proyectos PP de la EFPC | `01_…/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf` / `…/Proyectos_PP_2025.pdf` | `https://efpc.envigado.gov.co/wp-content/uploads/2025/07/…` |
| PI24X | Proyectos de inversión 2024 (xlsx) | `02_…/informe-de-gestion-alcaldia-de-envigado/proyectosdeinversion2024.xlsx` | 002363/118141_… |
| CUIPO | CUIPO, ejecución de gastos (contraste) | `02_…/datos_gov_co/datosgov_4f7r-epif_CUIPO_Ejecucion_Gastos_Envigado.csv` | datos.gov.co 4f7r-epif |
| FOR24 | Forensis 2024 (INMLCF), p. 37 tabla 5 y p. 440 tabla 14 | **por descargar a mano** (ver §6.3) | `https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf` |
| DANE | Proyecciones municipales de población | **por descargar** (§6.3) | dane.gov.co |
| TD | TerriData Envigado (txt) | `05_fuentes_externas_y_pot/dnp_terridata/TerriData05266f.txt.zip` | terridata.dnp.gov.co |

**Indicadores calculables.** Fórmula y semáforo exactos de la Matriz. Escenarios: identificadores de §3.

| Código | Fórmula exacta (Matriz) | Semáforo exacto | Variables y fuente (página, hoja o celda) | Años | Territorio | Escenarios que lo afectan |
|---|---|---|---|---|---|---|
| **CS.DSI** | DSI = (100 × MD) / MS; sin tope (MD > MS → Revisar) | Verde ≥ 90 · Amarillo 70 a < 90 · Rojo < 70 · Gris N/D | **MS:** lista de metas (AP91 hoja `Apéndice9.1.2.`, filas 16, 24, 71, 74, 119, 129, 146 en el escenario por defecto). **MD:** por meta y vigencia, ¿hay valor observado publicado con la misma definición y unidad? Fuentes: EV24/EV25 (solo productos; pp. 12-18), G24/G25, POTEP pp. 86-87, FOR24, MinSalud, TD. Soporte: una fila por meta × año, con documento, página y fecha de publicación | 2024, 2025 | Municipio | `dsi_lista_metas`, `dsi_fuente_valor`, `ventana_documental_cs`, `regla_estado_calculo` |
| **CS.PCS** | PCS = (100 × EE) / EP; "limitar el puntaje a 100" | Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50 · Gris N/D | **EE:** PI26, CONSECUTIVO 544 (producto 5.1.2.1.2), col. BM/BN/BO "AVANCE EJECUTADO 2024/2025/2026". **EP:** misma fila, col. AS/AT/AU "META PROGRAMADA 2024/2025/2026". Observaciones de la UE y del DAP en las columnas de observaciones y evidencias | 2024, 2025 (2026 solo variable: EE = 0 al 2T) | Municipio | `tratamiento_2026`, `regla_estado_calculo`, `corte_fecha_registro` |
| **CS.VRC** | VRC = (100 × EV) / ET; EV ⊆ ET; MD > MS → Revisar; "si hay casos con presencia desconocida, VRC = N/D" | Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50 · Gris N/D | **ET:** eventos de rendición de cuentas del universo elegido. Fuentes: PI26 consecutivos 545 y 546 (obs. UE), QA24, QA100, QA25, ACS24a, ACS24b, ACS25, ACS26, CMD, VCA, ETN, JUS. **EV:** eventos con presencia de veedurías según el criterio elegido (mención en el acta: ACS24a p. 1 y p. 17, ACS24b pp. 8-9 y 12, ACS25 p. 8, ACS26 p. 9, carta CMD p. 1; evento de veedurías del 12-dic-2024 en PI26 consecutivo 545) | 2024, 2025 | Municipio | `vrc_universo`, `vrc_acreditacion_presencia`, `vrc_evento_sin_mencion`, `vrc_asignacion_anio`, `ventana_documental_cs` |
| **CS.RPO** | RPO = (100 × OR) / OT; OR ⊆ OT; sin tope | Verde ≥ 90 · Amarillo 70 a < 90 · Rojo < 70 · Gris N/D | **OT:** intervenciones ciudadanas registradas en los cuadros de preguntas y respuestas del universo (QA24, cuadro 1-22, 24 pp.; QA100, 14 pares; QA25, 3 intervenciones; QAS25, 6 ítems; QAS26, 3; cartas CMD). **OR:** las que tienen respuesta de fondo publicada al corte, según la rúbrica de §3 | 2024 (2025 condicionado) | Municipio | `rpo_universo`, `rpo_filtro_pdm`, `rpo_criterio_fondo`, `rpo_registro_incompleto`, `ventana_documental_cs` |
| **CS.ICRCS** | ICRCS = (DSI + PCS + VRC + RPO) / 4; requiere las 4 entradas numéricas con el mismo corte; no redistribuye pesos | Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50 · Gris N/D | Puntajes de los cuatro indicadores anteriores (lee `parciales/CS.*_resultados.csv`) | 2024 (2025 N/D por defecto) | Municipio | Todos los parámetros CS; se aplica la estrategia por factor de §3.3 |
| **PC.I2** | I2 = (100 × V1) / V2; V1 > V2 → Revisar | Verde ≥ 80 · Amarillo 50 a < 80 · Rojo < 50 · Gris N/D | **V2:** PDM p. 558 (13 zonas; 12 comités; zona 1 pendiente), PI26 columnas C:O "ZONA 1…ZONA 13". **V1:** zonas con al menos una actividad realizada. 2024: ANTE p. 53 (reuniones del 29-ene al 17-feb-2024 en cada una de las 13 zonas), PI26 consecutivo 531 (12 asambleas PP). 2025: G25 p. 92 (12 asambleas zonales PP, ago-2025, zonas 2-13), POTC p. 1 (cronograma POT con zonas 1 a 13: programación, no ejecución). Soporte: una fila por zona × año × actividad | 2024, 2025 | Municipio (en soporte, zona sí/no) | `pc_i2_actividades`, `pc_i2_evidencia`, `zonas_oficiales` |
| **PC.I3** | I3 = (100 × V1) / V2; V1 debe pertenecer a lo programado; V1 > V2 → Revisar | Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50 · Gris N/D | **Unidad "acciones" (por defecto):** 2025: G25 pp. 87-88 (proyecto 0541: 3 diplomados, 9 seminarios y 9 cursos culminados; no ejecutados: 1 diplomado, 3 seminarios y 2 semilleros); 2024: G24 pp. 62-63 y PI26 consecutivo 521 (obs. UE 2024), sin programación, así que V2 2024 = N/D. **Unidad "personas":** PI26 consecutivo 521 (META PROGRAMADA AS/AT vs AVANCE EJECUTADO BM/BN); antes de usarlo, verificar en el archivo la unidad y los valores (las notas citan programación 10/15/15/10 frente a ejecución 95 y 186, lo que daría V1 > V2) | 2025 (acciones); 2024-2025 (personas) | Municipio | `pc_i3_unidad`, `pc_i3_excluir_servidores`, `pc_i3_v1_mayor_v2` |
| **PC.I5** | I5 = (100 × V1) / V2; adicionales aparte; V1 > V2 → Revisar | Verde ≥ 75 · Amarillo 50 a < 75 · Rojo < 50 · Gris N/D | PPSS25, hojas PROGRAMACIÓN (col. "EXPRESIÓN NUMÉRICA DE LA ACTIVIDAD") y SEGUIMIENTO, líneas E5Ld-M01-A01 (CMSSS, trimestral) y E5Ld-M02-A01 (COPACO, trimestral): 4 + 4 previstas. Revisar si PPSS24 trae las mismas líneas para 2024. La línea E5Ld-M02-A02 está rotulada "VIGENCIA 2024" dentro de 2025: anotarlo | 2025 (Anual; Semestre 1 y 2 si las fechas del seguimiento lo permiten) | Municipio | `pc_i5_instancias` |
| **PDM.CM** | CM = (100 × M) / O; máximo 100; si O = 0 con tasa válida, CM = 100 | Verde = 100 sin redondear · Amarillo 75 a < 100 · Rojo 0 a < 75 · Gris N/D | **M = 8** (AP91 `Apéndice9.1.2.` fila 119; PDM p. 706; DTS p. 738; horizonte 2027). **O:** FOR24 p. 440 tabla 14 (18 muertes por transporte; tasa publicada 7,25 por 100.000) y p. 37 tabla 5. Alternativa Policía: datos.gov.co `uav5-b85g` (cod_muni 05266) más población DANE. TD como contraste (7,7 en 2024) | 2024 | Municipio | `pdm_cm_fuente`, `poblacion_dane`, `horizonte_meta_pdm` |
| **PDM.CA** | CA = (100 × O) / M; máximo 100 | Verde = 100 sin redondear · Amarillo 75 a < 100 · Rojo 0 a < 75 · Gris N/D | **M = 3,88 m²/hab** (AP91 fila 129; PDM p. 717, rotulado "Metros cuadrados"; DTS p. 748). **O:** POTEP p. 87, tabla 39, EPV urbano 2024 = 5,75 (1.382.762 m² / 240.655 hab.); p. 86, tabla 38, EPV total = 69,41; línea base AMVA 2022 = 4,66 (contexto) | 2024 | Municipio | `pdm_ca_fuente`, `horizonte_meta_pdm` |
| **PP.EPP** | EPP = 100 × OPP / APP; misma vigencia, zona, corte y fuente de financiación; OPP > APP → Revisar | Verde 90 a 100 · Amarillo 75 a < 90 · Naranja 50 a < 75 · Rojo 0 a < 50 · N/D o NE sin puntaje | **2025:** PA25 hoja `PlanAcción27.01.2026`, col. N (APROPIACIÓN DEFINITIVA, índice 13) y col. O (EJECUCIÓN OBLIGACIÓN, índice 14), filas con "ZONA n" en el nombre (61, sin el 0553); BPIN de la hoja `ROCA07.01.26`; contraste EFPC25 (pp. 1-5), HAC (sufijo del rubro = código de proyecto) y CUIPO. **2026:** APP de PA26 (66 proyectos, $24.519.738.737); OPP de HAC26 al 30-jun-2026 ($5.702.528.148). **2024:** HAC24 (rubros con sufijo de proyecto) × EFPC24 (50 proyectos; 47 con correspondencia, 34 exactas) y PI24X. Lógica 2025 reutilizable de `scripts/epp_pp_2025.py` | 2024 (aprox.), 2025, 2026 S1 (provisional) | Municipio + Zona 2…13; **Zona 1 = NE** (APP = 0, sin proyectos) | `pp_identificacion`, `pp_incluir_0553`, `pp_fuente_calculo`, `pp_2024_metodo`, `pp_2026_opp_fuente`, `tratamiento_2026`, `corte_fecha_registro` |

**Valores de control.** Las cifras de las notas de la Fase 2 **no son insumo**: salieron de scripts temporales no versionados. Cada script debe rederivar sus cifras y compararlas con este rango; si no cuadra, debe anotar la diferencia.

| Indicador | Control esperado (escenario por defecto) |
|---|---|
| CS.PCS | 2024 = 100 (1/1); 2025 = 100 (1/1) |
| CS.RPO 2024 (Alcalde, sin felicitaciones) | OT ≈ 17; OR ≤ 17 (una respuesta remite a "lo respondido en el evento"; con la rúbrica estricta probablemente no cuenta) |
| CS.VRC | ET: 1 a 8 según el universo; con el universo institucional (545 + 546): EV = 1, ET = 2 |
| PC.I2 | 2024 = 13/13 (asambleas PDM); 2025 = 12/13 (solo PP) o 13/13 si se acredita la realización de las asambleas POT |
| PC.I3 2025 (acciones) | 21/25 = 84 antes de excluir la formación para servidores |
| PC.I5 2025 | 8/8 = 100 |
| PDM.CM 2024 | O = 7,25, M = 8, CM = 100 (verde) |
| PDM.CA 2024 | O = 5,75, M = 3,88, CA = 100 (verde, con salvedad de comparabilidad) |
| PP.EPP 2025 | APP $20.928.690.218, OPP $20.684.427.720, EPP 98,83 (98,86 con el 0553) |
| PP.EPP 2026 S1 | 5.702.528.148 / 24.519.738.737 ≈ 23,3 (provisional). Corregido el 9-nov-2026: la cifra 11.555.041.247 citada en octubre no se reproduce con el informe de Hacienda al 30-jun-2026 ni con ROCA |
| PP.EPP 2024 | aprox. 17.199.119.162 / 17.332.821.550 ≈ 99,2 (47 de 50 proyectos) |

---

## 2. Variables que se construyen para los No viables (sin calcular el indicador)

En todos los casos, el script del sistema escribe la variable en `valores_variables` y su soporte, y deja el indicador con `disponibilidad = "N/D"` y `motivo_nd` explícito. **No se escribe puntaje**, ni siquiera en escenarios alternativos.

| Indicador (motivo por el que no se calcula) | Variables que se construyen | Fuente y procedimiento | Estado en Registro |
|---|---|---|---|
| **PC.I1** (no hay inventario oficial de procesos cerrados ni registro aporte-decisión) | PC.I1.V2 y PC.I1.V1 sobre el universo exploratorio `pc_universo_procesos` (2024: 3 procesos; 2025: 5), con una fila por proceso: nombre, fecha de cierre, documento de decisión y si incorporó al menos un aporte | G24 pp. 62-64, G25 pp. 90-92, ANTE p. 53, POTC p. 1, PI26 consecutivos 530 y 531, comunicados de las 88/90 iniciativas | Valor con estado "Pendiente" y Observaciones "universo exploratorio no oficial; indicador no calculado" |
| **PC.I4** (mismo universo inexistente; sin auditoría documento por documento) | PC.I4.V2 (= universo de I1) y PC.I4.V1 con lista de chequeo: identifica el proceso, sintetiza aportes, comunica decisión; cada criterio 0/1 en el soporte | QA24, QA25, EFPC25, comunicados, PLAN-ZONA-n | Ídem |
| **PC.I6** (no hay registro único de actividades ni evidencia del canal) | Lista parcial de actividades con canal virtual para intervenir (2024 ≥ 1: QA24 pp. 6-8; 2025 ≥ 1: G25 p. 92, votación web PP). V2 no se totaliza | QA24, G25, PI26 consecutivo 529 | V1: "Pendiente" con nota "cota inferior"; V2: "No disponible" |
| **PC.ISPC** (I1 e I4 N/D; Guía 15) | Ninguna adicional | — | — |
| **PDM.CS** (la cobertura RS con el método del plan no está publicada) | Auxiliares: AUX.PDM.CS.AFIL_RS (29.794 dic-2025; 31.997 ago-2026), AUX.PDM.CS.AFIL_TOTAL, AUX.PDM.CS.POB_MINSALUD (250.767; 268.990), AUX.PDM.CS.COB_TOTAL_SGSSS (87,65; 82,59), AUX.PDM.CS.RS_SOBRE_POB_TD (TD 2022-2025). Meta PDM.CS.M = 99,5 | MinSalud xlsx, hoja `CoberturaMunicipio`, fila ENVIGADO (descarga, §6.3); TD; PDM p. 362 (diagnóstico 99,6/99,2/99,5) y p. 602 | PDM.CS.O "No disponible", con los proxies en Observaciones; meta "Pendiente" |
| **PDM.CP** (no hay valor ni método oficial; el PDM tiene además un índice 3,23) | Solo la meta PDM.CP.M = 100 y AUX.PDM.CP.INDICE_SITUACIONAL = 3,23 (2022, PDM p. 561) | PDM p. 742, DTS p. 775, AP91 fila 146 | O: "No disponible" |
| **PDM.CD** (solo hay oferta de disciplinas, que la Matriz prohíbe convertir en cobertura) | AUX.PDM.CD.DISCIPLINAS (38 de 40, 2025). Meta 6 | G25 p. 5; PDM p. 609; AP91 fila 24 | O: "No disponible" |
| **PDM.ILE** (LS sin clasificación ICFES por plantel) | **PDM.ILE.A construida** = 100 × niños con atención integral / población DANE de 0 a 6 años (2024: 2.161, balance de Educación p. 24; 2025: 2.092, G25 p. 150). **Calibración obligatoria:** 2.059 (2023, TD) / población 0-6 de 2023 debe dar ≈ 16,3 (línea base); si no, probar 0-5 (`pdm_ile_a_rango_edad`) y anotar. ILE.S: AUX.PDM.ILE.S.IE_MEJORAN (≥ 1 en 2024 y ≥ 1 en 2025; balance de Educación p. 101; G25 p. 161), sin total de I.E. Si la usuaria descarga la clasificación ICFES, se construye S | Balance de Educación 2024 (`05_…/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf`), G25, TD, DANE por edad simple | A: "Pendiente" (construida); S: "No disponible" |
| **PDM.ICPD-6** (CS, ILE, CD y CP N/D) | Ninguna adicional | — | — |
| **PP.CEP** (PEC sin actas de recibo ni informes finales; autorreporte) | **PP.CEP.PE construida**: 50 / 61 / 66 (2024 / 2025 / 2026) por Municipio y Zona 2-13. **PEC proxy** (solo variable, rotulada "autorreporte"): con las 3 reglas de `pp_cep_regla_agregacion` (rango 2025: 45-55 de 61) | EFPC24, EFPC25, PA26; PA25 hoja `PlanAcción27.01.2026`, indicadores de producto, col. EJECUCIÓN FÍSICA NETA | PE: "Pendiente"; PEC: "No disponible" con el proxy en Observaciones |
| **PP.CMA** (no hay meta anual de personas ni personas únicas por proyecto) | Universo de proyectos para `pp_proyectos` y la hoja PP_Proyectos (177 filas municipales), con "¿Aplica meta de atención?" según `pp_cma_aplicabilidad`. PP.CMA.NA queda "No disponible". PA/MP parciales de 2024 (INDER pp. 5-9, Fundarte zona 9, ecozona 10) solo en Observaciones, sin datos personales | EFPC24/25, PA25/PA26, rendición PP 2024 (`01_…/rendicion-de-cuentas-presupuesto-participativo-2024/`) | NA: "No disponible" |
| **PP.IREPP** (CEP y CMA N/D; Guías 15 y 20) | Ninguna adicional. Se publica EPP suelto, con la alerta de que CEP y CMA no tienen soporte | — | — |

---

## 3. Catálogo de decisiones pendientes y escenarios

### 3.1 Parámetros

Formato listo para YAML; el borrador completo está en `supuestos_borrador.yaml`. Todos los valores por defecto son **PROVISIONALES**: no son una decisión de los docentes y deben rotularse así en cada salida.

| ID | Pregunta abierta | Opciones | Por defecto (PROVISIONAL) | Indicadores afectados |
|---|---|---|---|---|
| `corte_fecha_registro` | ¿Qué fecha de corte se escribe para datos de cierre de vigencia con corte real distinto del 31-dic (p. ej., plan de acción con corte 7-ene-2026)? | `real` · `normalizada_cierre_vigencia` (31/12 del año en la columna N; corte real en Observaciones y en `fecha_corte_real`) | `normalizada_cierre_vigencia` | Todos (Resultados exige N = N4) |
| `tratamiento_2026` | ¿Cómo tratar 2026 sin cierre (corte 30-jun)? | `excluir` · `semestre1_provisional` (Registro "Semestre 1", corte 30/06/2026, rotulado provisional) · `anual_provisional` (viola la Guía 4; solo sensibilidad) | `semestre1_provisional` | PP.EPP, CS.PCS (variable), PC.I2, PC.I3 |
| `regla_estado_calculo` | ¿Se calcula con valores en estado "Pendiente"? | `estricta_matriz` (solo Validado: todo N/D) · `preliminar_pendiente` | `preliminar_pendiente` en la columna `puntaje`; `estricta_matriz` siempre en `puntaje_regla_matriz` | Todos |
| `zonas_oficiales` | ¿Universo de zonas? | `13` (PDM p. 558) · `12_con_comite` | `13` | PC.I2, PP por zona |
| `poblacion_dane` | ¿Qué población usar en las tasas? | `serie_publicada_por_fuente` (la tasa tal como la publica la fuente) · `dane_postcovid_2020_2035` · `dane_vp_2018_2042` | `serie_publicada_por_fuente` | PDM.CM, PDM.ILE.A, AUX de PDM.CS |
| `horizonte_meta_pdm` | ¿Metas 2027 en años intermedios? | `meta_2027` (Guía 4 y 9: logro al corte frente a la meta final) · `metas_anuales` (no existen, da N/D) | `meta_2027` | PDM.* |
| `ventana_documental_cs` | ¿Hasta cuándo se aceptan publicaciones para la vigencia? | `guia_31mar_siguiente` (Guía 17; respuestas publicadas al 31-dic para RPO) · `hasta_fecha_consulta` | `guia_31mar_siguiente` | CS.DSI, CS.VRC, CS.RPO |
| `dsi_lista_metas` | ¿Cuáles metas forman MS? | `icpd7` (AP91 filas 16, 24, 71, 74, 119, 129, 146; MS = 7) · `icpd7_mas3_gobernanza` (más filas 147, 148 y 151: percepción de claridad de la información, IDI y Gobierno Digital; MS = 10; **propuesta del Analizador**) · `productos_plan_indicativo` (todos los productos con meta > 0; da 100 por construcción; solo control) | `icpd7` | CS.DSI, ICRCS |
| `dsi_fuente_valor` | ¿Qué cuenta como "valor observado publicado"? | `solo_alcaldia` · `oficial_misma_definicion` (Alcaldía o entidad oficial: INMLCF, MinSalud, DANE, ICFES, siempre que la definición y la unidad coincidan con la meta) · `oficial_incluye_proxy` | `oficial_misma_definicion` | CS.DSI, ICRCS |
| `vrc_universo` | ¿Qué es un "evento municipal de rendición de cuentas"? | `alcalde` (audiencia del Alcalde, consecutivo 546) · `institucional_pi_545_546` (audiencia + evento de veedurías) · `alcalde_mas_sectoriales_con_acta` · `amplio_todo_documento` (1 a 8 por año) | `alcalde_mas_sectoriales_con_acta` (recomendación de la Fase 3) | CS.VRC, ICRCS |
| `vrc_acreditacion_presencia` | ¿La mención de veedurías en el acta acredita presencia? | `mencion_en_documento` · `lista_asistencia` (listas no publicadas: EV = N/D) | `mencion_en_documento` | CS.VRC, ICRCS |
| `vrc_evento_sin_mencion` | Evento del universo cuyo documento no menciona veedurías | `nd_siempre` (lectura literal: presencia desconocida, VRC N/D) · `cero_si_documento_integral` (acta o memoria completa con asistentes: ausencia; cuadro de preguntas o justificación: desconocida, N/D) · `cero_siempre` | `cero_si_documento_integral` | CS.VRC, ICRCS |
| `vrc_asignacion_anio` | Evento que rinde la vigencia t realizado en t+1 (p. ej., PP 2024 en ene/feb-2025) | `fecha_evento` (Matriz: "eventos realizados en la vigencia") · `vigencia_rendida` | `fecha_evento` | CS.VRC |
| `rpo_universo` | ¿Qué ejercicios integran OT? | `alcalde` · `alcalde_y_100dias` · `todos_los_cuadros` (Alcalde, 100 días, Salud, CMD) | `alcalde` (recomendación de la Fase 3) | CS.RPO, ICRCS |
| `rpo_filtro_pdm` | ¿Qué es una observación "sobre el PDM"? | `excluir_felicitaciones` · `tematico_pdm` (solo las que remiten a un programa, meta o sector del PDM; excluye trámites individuales) · `todas` | `excluir_felicitaciones` | CS.RPO |
| `rpo_criterio_fondo` | ¿Qué es "respuesta de fondo"? | `estricto_rubrica` (atiende el asunto con dato, decisión, compromiso fechado o negativa motivada; no cuentan acuse, agradecimiento, "se tendrá en cuenta" ni remisión sin contenido) · `laxo_texto_respuesta` (cualquier texto de respuesta salvo acuse explícito) | `estricto_rubrica` | CS.RPO, ICRCS |
| `rpo_registro_incompleto` | Cuadro publicado claramente incompleto (Alcalde 2025: 3 intervenciones) | `nd` (Matriz: un listado parcial no permite conocer OT) · `calcular_con_publicado` | `nd` | CS.RPO 2025 |
| `pc_i2_actividades` | ¿Qué actividades cuentan para cubrir una zona? | `asambleas_pp` · `asambleas_pp_pdm_pot` | `asambleas_pp_pdm_pot` | PC.I2 |
| `pc_i2_evidencia` | ¿Qué acredita la actividad en la zona? | `realizacion_reportada` (informe de gestión, anteproyecto o ejecución del plan indicativo) · `incluye_programacion` (cronogramas) · `acta_por_zona` (no publicadas: N/D) | `realizacion_reportada` | PC.I2 |
| `pc_i3_unidad` | ¿Acciones formativas o personas? | `acciones` · `personas_producto_521` | `acciones` | PC.I3 |
| `pc_i3_excluir_servidores` | ¿Excluir la formación exclusiva para servidores? (la Matriz lo exige) | `true` · `false` | `true` | PC.I3 |
| `pc_i3_v1_mayor_v2` | Ejecución mayor que la programación | `revisar` (Resultados) · `tope_100` | `revisar` | PC.I3 |
| `pc_i5_instancias` | ¿Qué instancias se seleccionan? | `salud_copaco_cmsss` · `ampliado_cmpc_ctp_comites` (sin calendario: N/D) | `salud_copaco_cmsss` | PC.I5 |
| `pc_universo_procesos` | Universo de procesos cerrados (solo variables) | `exploratorio_pp_pdm_pot_rc` · `inventario_oficial` (N/D) | `exploratorio_pp_pdm_pot_rc` | PC.I1, PC.I4 (variables) |
| `pdm_cm_fuente` | Fuente de mortalidad vial | `inmlcf_forensis` · `policia_datosgov` (homicidio en accidente de tránsito; otra definición) | `inmlcf_forensis` | PDM.CM, CS.DSI |
| `pdm_ca_fuente` | ¿El POT (5,75) frente a la línea base 3,88? | `pot_epv_urbano` · `pot_epv_total` · `estricto_serie_linea_base` (N/D) | `pot_epv_urbano` (rotulado "serie distinta de la línea base") | PDM.CA, CS.DSI |
| `pdm_cs_fuente` | ¿MinSalud con denominador propio? (solo variables) | `metodo_plan` (N/D) · `minsalud_rs_sobre_pob_total` · `minsalud_cobertura_total` (la Matriz lo prohíbe) | `metodo_plan` | PDM.CS (variables), CS.DSI |
| `pdm_ile_a_rango_edad` | Denominador de ILE.A | `0_6` · `0_5` | `0_6` (el que mejor calibre con 16,3 en 2023; documentar) | PDM.ILE.A |
| `pp_identificacion` | ¿Regla para identificar proyectos PP? | `efpc_x_plan_accion` · `solo_plan_accion_zona` | `efpc_x_plan_accion` | PP.* |
| `pp_incluir_0553` | ¿Incluir el costo del proceso (0553)? | `false` · `true` | `false` | PP.EPP |
| `pp_fuente_calculo` | ¿Fuente de APP y OPP? | `plan_accion_roca` · `hacienda` · `cuipo` (solo contraste) | `plan_accion_roca` (2025-2026); 2024 solo `hacienda` | PP.EPP |
| `pp_2024_metodo` | ¿Cómo tratar la correspondencia aproximada de 2024? | `hacienda_correspondencia_completa` (47/50, rotulado aproximado) · `solo_exactos` (34) · `nd` | `hacienda_correspondencia_completa` | PP.EPP 2024 |
| `pp_2026_opp_fuente` | ¿Fuente de OPP 2026 al corte? | `hacienda_30jun` · `plan_accion_xlsx` (corte anterior) | `hacienda_30jun` | PP.EPP 2026 |
| `pp_cep_regla_agregacion` | Regla del proxy de PEC (solo variable) | `todas_filas_dic_100` (49) · `solo_31dic` (45) · `ultima_fila_por_indicador` (55) | `todas_filas_dic_100` | PP.CEP (variable) |
| `pp_cma_aplicabilidad` | "¿Aplica meta de atención?" en PP_Proyectos | `pendiente` · `tipologia_sugerida` (infraestructura y medios = "No", con motivo) | `pendiente` (la tipología se anota en Observaciones) | PP.CMA (variables) |

Decisiones agregadas a la lista inicial:
- `corte_fecha_registro`, `regla_estado_calculo` y `ventana_documental_cs`: esta última aplica la Guía 17; la fecha de publicación de cada documento pasa a ser dato del soporte.
- `vrc_evento_sin_mencion` y `vrc_asignacion_anio`.
- `rpo_registro_incompleto` y `pc_i2_evidencia`.
- `pc_i3_excluir_servidores` y `pc_i3_v1_mayor_v2`.
- `pdm_ile_a_rango_edad`, `pp_2026_opp_fuente`, `pp_cma_aplicabilidad` y `zonas_oficiales`.

### 3.2 Convención de escenarios

- `escenario_id = "defecto"` cuando todos los parámetros que afectan al indicador están en su valor por defecto.
- Si no, `escenario_id` = lista ordenada alfabéticamente de **solo las desviaciones**, `param=opcion` separadas por `;`. Ejemplo: `rpo_criterio_fondo=laxo_texto_respuesta;rpo_universo=todos_los_cuadros`.
- La columna `parametros` guarda siempre el juego completo de los parámetros aplicables, en formato `clave=valor;…`, para que la fila sea autosuficiente.
- `es_defecto` vale `True` solo en la fila "defecto". Toda fila lleva `provisional = True` mientras exista algún parámetro con `estado: PROVISIONAL`, es decir, siempre en esta etapa.

### 3.3 Estrategia de cálculo

- **Indicadores simples:** factorial completo de los parámetros que los afectan (máximo 48 combinaciones; RPO es el mayor, con 3 × 3 × 2 × 2 × 2 = 72 incluida la ventana). Si una combinación supera 100, usar un factor a la vez.
- **ICRCS:**
  - escenario "defecto";
  - **un factor a la vez** (cada opción no por defecto de cada parámetro CS, manteniendo el resto en el valor por defecto);
  - dos envolventes, `envolvente_min` y `envolvente_max`: la combinación de opciones que da el puntaje mínimo y la que da el máximo, elegidas a partir de los puntajes de los cuatro componentes ya calculados.
- `comun.escenarios(codigo)` genera la lista a partir de `supuestos.yaml`, con un mapa `afecta:` por parámetro, para que los cuatro Ejecutores usen la misma enumeración.

---

## 4. Esquema exacto de las tablas de salida

**Convenciones generales:**
- CSV con separador `;`, codificación UTF-8 con BOM (`utf-8-sig`), punto decimal y sin separador de miles.
- Fechas en ISO `aaaa-mm-dd` en el CSV. En el Excel, celdas de fecha con formato `dd/mm/aaaa` (Guía 18).
- Porcentajes y puntajes en escala 0-100 (99,5, no 0,995). Montos en COP enteros.
- Cada CSV tiene una hoja homónima en `outputs/indicadores/indicadores_calculados.xlsx`, con tipos nativos, más las hojas `Léeme` y `Supuestos` (volcado del YAML).
- **Territorio:** `Municipio`, `Zona 1` … `Zona 13`, sin ceros a la izquierda.
- **Periodo:** `Anual`, `Semestre 1`, `Semestre 2` (listas de la Matriz).
- **IDs de variable:** los de la Matriz (`CS.VRC.ET`, `PDM.CM.O`…). Las variables auxiliares usan `AUX.<SISTEMA>.<IND>.<NOMBRE>` (p. ej., `AUX.PDM.CM.MUERTES`, `AUX.PDM.CM.POBLACION`, `AUX.PP.EPP.APP_EFPC`). Siempre se usa el ID completo; nunca "CS" solo, por la colisión CS / PDM.CS.

### 4.1 `valores_variables` (una fila por variable × año × periodo × territorio × escenario)

| Columna | Tipo | Descripción |
|---|---|---|
| sistema | texto | Control social / Participación ciudadana / Cumplimiento del PDM / Presupuesto participativo |
| codigo_indicador | texto | CS.DSI, …, o el indicador de un AUX |
| id_variable | texto | ID de la Matriz o AUX |
| nombre_variable | texto | Nombre completo (Registro col. D) |
| unidad | texto | Lista de la Matriz: Conteo, Porcentaje (0–100), COP, Muertes / 100.000 hab., m² / habitante; los AUX pueden usar otra unidad |
| tipo_valor | texto | observado · meta · construido · proxy · auxiliar |
| anio | entero | 2024-2027 |
| periodo | texto | Anual / Semestre 1 / Semestre 2 |
| territorio | texto | Municipio / Zona n |
| escenario_id | texto | §3.2 |
| es_defecto | bool | |
| parametros | texto | `clave=valor;…` |
| valor | decimal (vacío si no hay) | Sin redondear |
| estado_informacion | texto | Pendiente · No disponible · No evaluable · No aplica (**nunca Validado**) |
| categoria_disponibilidad | texto | Disponible / Construible / Parcial / Proxy / No disponible (hoja Variables) |
| fuente | texto | Nombre de la publicación (va a Registro col. K) |
| documento_ruta_local | texto | Ruta bajo `data_raw_IIC/` (columna interna) |
| enlace | texto | URL espejo o URL de origen (va a Registro col. L) |
| pagina_tabla | texto | Página PDF, tabla, hoja y celda o consecutivo (va a col. M) |
| fecha_corte | fecha | La que va a la col. N, según `corte_fecha_registro` |
| fecha_corte_real | fecha | Corte real del dato |
| fecha_publicacion | fecha o vacío | Fecha interna del documento o carpeta `content/files/NNNNNN`; **no usar** la `fecha_publicacion` del inventario sin verificarla (ver §7) |
| fecha_consulta | fecha | Fecha de ejecución del script (col. O) |
| horizonte_meta | entero o vacío | 2027 en las metas (col. P) |
| n_elementos_soporte | entero o vacío | Filas de soporte que suman el valor |
| ids_soporte | texto | IDs de `soporte_conteos` separados por `,` |
| provisional | bool | |
| script | texto | `scripts/indicadores/<archivo>.py` |
| observaciones | texto | Brecha, discrepancias y alternativas (va a col. Q) |

Llave única: (id_variable, anio, periodo, territorio, escenario_id).

### 4.2 `resultados_indicadores` (una fila por indicador × año × periodo × territorio × escenario)

| Columna | Tipo | Descripción |
|---|---|---|
| sistema, codigo, nombre, tipo | texto | tipo = Indicador / Complementario / Índice |
| anio, periodo, territorio | | |
| escenario_id, es_defecto, parametros | | |
| entrada_1_id … entrada_4_id | texto | IDs de las entradas, en el orden de Resultados (F:I) |
| entrada_1 … entrada_4 | decimal o texto | Valor, o "N/D" / "Revisar" |
| intermedio_1, intermedio_2 | decimal o vacío | LA y LS (ILE); IBS e ITG (ICPD-6) |
| puntaje | decimal o vacío | Sin redondear, con la regla `preliminar_pendiente` |
| puntaje_1d | decimal | Redondeado a 1 decimal, solo para mostrar |
| semaforo | texto | Verde / Amarillo / Naranja / Rojo / Gris |
| disponibilidad | texto | Calculado / NE / N/D / Revisar |
| motivo_nd | texto | Por qué no hay puntaje |
| puntaje_regla_matriz | texto | Lo que mostraría la matriz con estados Pendiente: "N/D" en esta etapa |
| fecha_corte_comun | fecha | |
| provisional | bool | |
| alerta | texto | Por ejemplo, "CEP o CMA < 50" (Guía 20) o "serie distinta de la línea base" (CA) |
| no_calculado_por_etapa | bool | True para los No viables |
| script, observaciones | texto | |

Llave única: (codigo, anio, periodo, territorio, escenario_id).

### 4.3 `pp_proyectos` (una fila por proyecto × año × periodo × territorio de cálculo)

Columnas: anio · periodo · territorio (`Municipio`) · zona (2-13) · codigo_proyecto (texto de 4 dígitos; en 2024, sin correspondencia: `EFPC2024-Z<zona>-<n>`) · bpin · nombre_proyecto · unidad_ejecutora · en_listado_efpc (bool) · valor_listado_efpc (COP) · tipo_correspondencia (exacta / aproximada / sin_correspondencia) · apropiacion_definitiva · compromiso · obligacion · fuente_valor (plan_accion_roca / hacienda / cuipo) · fecha_corte · fecha_corte_real · es_0553 (bool) · incluido_epp_defecto (bool) · ejecucion_fisica_neta_dic (decimal) · pec_todas_filas_dic_100 / pec_solo_31dic / pec_ultima_fila (0/1) · tipologia_sugerida (servicio a personas / infraestructura / medios / otro) · aplica_meta_atencion (Sí / No / Pendiente) · motivo_exclusion · personas_atendidas_PA (vacío) · meta_atencion_MP (vacío) · estado_informacion · fuente · enlace · pagina_hoja · observaciones.

Llave: (anio, periodo, territorio, codigo_proyecto). Volumen: 50 + 61 + 66 = 177 filas, más el 0553 de 2025 y 2026 marcado `incluido_epp_defecto = False`.

### 4.4 `soporte_conteos` (un registro por elemento contado)

| Columna | Tipo | Descripción |
|---|---|---|
| id_soporte | texto | `<SIS>-<IND>-<aaaa>-<nnn>`, p. ej. `CS-VRC-2024-003` |
| sistema, codigo_indicador, id_variable | texto | |
| anio, periodo, territorio | | |
| elemento_tipo | texto | evento · observacion · meta · accion_formativa · sesion · zona_actividad · proceso · proyecto · actividad_virtual |
| elemento_id | texto | Identificador estable (p. ej., número de fila del cuadro QA24) |
| descripcion | texto | **Sin datos personales:** tema o asunto resumido, nunca nombres, teléfonos ni correos |
| fecha_elemento | fecha o vacío | |
| documento | texto | Nombre del documento |
| ruta_local, enlace, pagina | texto | |
| fecha_publicacion_doc | fecha o vacío | Ver §7 |
| banderas | texto | `clave=valor;…`, con los atributos que usan los parámetros. Ejemplos: `es_alcalde=1;tiene_acta=1;doc_integral=1;mencion_veedurias=0;vigencia_rendida=2024`; `tipo=pregunta;sobre_pdm=1;respuesta_texto=1;fondo_estricto=0` |
| clasificacion | texto | Etiqueta legible (fondo / acuse / felicitación / sin respuesta…) |
| justificacion | texto | Por qué se clasificó así (cita breve parafraseada) |
| lector | texto | `ejecutor_CS` u otro; habrá segunda lectura por muestreo |
| cuenta_en_defecto | 0/1 | Si entra en el conteo del escenario por defecto |

Regla: **los conteos se derivan de `banderas` + parámetros**. Ningún script escribe un conteo que no salga de filtrar esta tabla. Las clasificaciones manuales se guardan en `outputs/indicadores/insumos_manual/<SIS>_<tema>.csv`, versionadas y con las mismas columnas. Los scripts las leen y validan: que no tengan datos personales y que no haya IDs duplicados.

### 4.5 `diccionario_datos`

Columnas: tabla · columna · tipo · obligatoria (sí/no) · descripcion · valores_permitidos · ejemplo · origen (script o columna de la Matriz). Una fila por columna de las tablas 4.1-4.4, más las hojas de la matriz diligenciada (Registro A:Q y PP_Proyectos A:M).

### 4.6 Rutas de salida

| Ruta | Contenido |
|---|---|
| `outputs/indicadores/parciales/<CODIGO>_{valores,resultados,soporte}.csv` | Escritos por cada script de indicador; nadie más los toca |
| `outputs/indicadores/{valores_variables,resultados_indicadores,pp_proyectos,soporte_conteos,diccionario_datos}.csv` | Escritos solo por el ensamblador |
| `outputs/indicadores/indicadores_calculados.xlsx` | Las 5 tablas + Léeme + Supuestos |
| `outputs/indicadores/matriz_diligenciada_escenario_defecto.xlsx` | Copia de la matriz con Registro y PP_Proyectos diligenciadas |
| `outputs/indicadores/resumen_calculos.md` | Resumen |
| `config/supuestos.yaml` | Parámetros, desde el borrador |

---

## 5. Reglas de la Guía aplicables y diligenciamiento de la matriz

### 5.1 Reglas

- **Guía 3. Evidencia mínima:** un valor solo es "utilizable" con Estado = Validado, valor numérico ≥ 0, fuente en K y fecha de corte en N (fórmula `Registro!S`). Además, los conteos deben ser enteros y los porcentajes ≤ 100.
  - En esta etapa **todo queda "Pendiente"**, así que la hoja Resultados de la copia mostrará N/D en todos los indicadores. Es el comportamiento correcto de la Matriz (Guía 19: "Validado" lo decide el responsable).
  - Los puntajes preliminares viven en `resultados_indicadores.puntaje`; `puntaje_regla_matriz` = "N/D".
- **Guía 4. Corte:** Resultados busca filas cuyo N sea igual a N4, y N4 vale 31/12 del año (Anual) o 30/06 (Semestre 1).
  - Un dato con corte 07/01/2026 escrito tal cual nunca entra en el cálculo; de ahí `corte_fecha_registro`.
  - **Las metas también exigen N = N4 y P = K4 (2027)**: en las filas de meta de cada año se escribe N = 31/12/año y P = 2027.
- **Guía 5:** una sola fila por variable/año/periodo/territorio (la llave R detecta duplicados: "Revisar"). Hay fórmulas hasta la fila 1005. **No escribir en R:S.**
- **Guía 7: N/D frente a NE.**
  - N/D = faltante, no validado o fuera del corte.
  - NE = base cero, sin universo (p. ej., EPP de la Zona 1, sin APP).
  - Revisar = duplicados o numerador fuera del universo.
  - Un cero observado con base positiva da 0 puntos.
- **Guía 8-9:** las metas del PDM tienen horizonte 2027; el resultado es "logro al corte frente a la meta final", no cumplimiento anual. Verificar las modificaciones antes de validar; el acuerdo 059/2025 toca productos, no resultados (sin verificar en su totalidad).
- **Guías 10, 13, 17 y 20:**
  - Guía 10: no inferir CA ni ILE desde actividades.
  - Guía 13: EPP sin cofinanciación (en 2025-2026 todo es ICLD, pero hay que verificarlo por fila); no promediar porcentajes zonales: el municipal es Σ OPP / Σ APP.
  - Guía 17: universos ex ante; una convocatoria no acredita presencia y un acuse no acredita respuesta; cierre al 31-dic y búsqueda hasta el 31-mar.
  - Guía 20: la alerta de PP se mantiene aunque IREPP no se calcule.
- **Guía 14:** CS y PDM, solo Municipio y Anual; I1 e ISPC anuales; I3-I6 admiten semestre; PP anual, municipal o por zona.
- **Guía 15:** un faltante bloquea el índice; los pesos no cambian.
- **Guía 16:** clasificar sin redondear. PDM: verde solo en 100. PP: cuatro colores.
- **Guía 18:** 99,5 y no 0,995; fechas dd/mm/aaaa; COP.

### 5.2 Registro (copia en `outputs/indicadores/matriz_diligenciada_escenario_defecto.xlsx`)

- Copiar el original con `shutil.copy2` y verificar con `cmp`. **El original de `Documentos guias/` no se modifica.**
- Abrir la copia con `openpyxl.load_workbook(...)` (sin `data_only`) y escribir solo **A:Q**. Las filas 6-161 ya tienen A:E, F, G y H prellenadas (39 variables × 2024-2027, Municipio, Anual): se actualizan in situ, buscando por (C, F, G, H).
- Las filas nuevas se agregan desde la 162: zonas de PP y 2026 "Semestre 1". Máximo 1.000 filas, hasta la 1005; volumen estimado ≈ 156 + 12 zonas × 3 años × 4 variables PP + ~15 de 2026 ≈ 320.

| Col. | Contenido |
|---|---|
| A-E | Sistema, indicador, ID, nombre y unidad (copiar de la fila base de la variable) |
| F | Año |
| G | Periodo |
| H | Territorio |
| I | Valor del **escenario por defecto** (numérico; vacío si no hay) |
| J | "Pendiente" si hay valor; "No disponible" si no lo hay. Nunca "Validado" |
| K | Fuente o publicación |
| L | Enlace (espejo) |
| M | Página, tabla o anexo |
| N | Fecha de corte (`datetime.date`, formato dd/mm/aaaa) |
| O | Fecha de consulta |
| P | Horizonte de la meta (2027; solo metas) |
| Q | `"ESCENARIO POR DEFECTO PROVISIONAL: <param=valor;…>. Corte real: <fecha>. <brecha>. Alternativas: <escenario>=<valor>; …"` (≤ 600 caracteres; el resto queda en el CSV) |

- Las filas de 2027 quedan con J "Pendiente" y Q "Vigencia futura", salvo las metas, que ya vienen precargadas.

### 5.3 PP_Proyectos

- Escribir **A:M**, filas 6-305 (capacidad 300). Solo filas `Municipio`: el universo municipal deduplicado (Guía 13) suma 177 o 179 filas. Las filas por zona superarían la capacidad, así que la zona se anota en M.

| Col. | Contenido |
|---|---|
| A | Año |
| B | "Anual" (2026: "Semestre 1") |
| C | "Municipio" |
| D | Código del proyecto (texto) |
| E | Nombre |
| F | "Pendiente" (`pp_cma_aplicabilidad=pendiente`) |
| G, H | Vacías |
| I | Fuente |
| J | Enlace |
| K | Fecha de corte |
| L | "Pendiente" |
| M | `"Zona n; APP …; OPP …; tipología sugerida: …; correspondencia EFPC: …"` |

- **No escribir en N:P.**

### 5.4 Resultados

- No se modifica: selectores B4 = 2024, E4 = Anual, H4 = Municipio, K4 = 2027.
- Marcar `wb.calculation.fullCalcOnLoad = True`.
- Verificar después de guardar:
  - que las fórmulas de R:S, N:P y Resultados sean idénticas al original (comparar el texto de las celdas);
  - que se conserven las 5 validaciones de Registro, las 4 de PP_Proyectos y las 2 de Resultados;
  - que se conserve el formato condicional de Resultados (1).
- Si openpyxl pierde algo, documentarlo y no entregar la copia.

---

## 6. Asignación a los 4 Ejecutores, módulo común y orden de trabajo

### 6.1 Módulo común (lo escribe primero un único responsable: la sesión principal o el "Ejecutor 0"; los demás **no lo editan** y piden cambios al coordinador)

`scripts/indicadores/comun.py`:
- `cargar_supuestos(path='config/supuestos.yaml')`:
  - valida la estructura: cada parámetro tiene opciones, un valor por defecto que pertenece a las opciones y una lista `afecta`;
  - **requiere PyYAML**, que hoy solo está en el site-packages del usuario, invisible con `-I`. Hay que instalarlo en el intérprete activo (`python3 -m pip install pyyaml`) o ejecutar los scripts sin `-I` desde la raíz del repo, sin módulos homónimos.
- `escenarios(codigo) -> list[dict]`: aplica §3.3 (factorial, un factor a la vez o envolventes).
- `escenario_id(params, codigo)`.
- `UMBRALES` y `semaforo(codigo, puntaje)`: tabla de §1.2.
- `calcular(codigo, entradas, contexto)`: réplica de Resultados (§1.1); devuelve puntaje, disponibilidad y motivo.
- `indice(codigo, componentes)`.
- `COLUMNAS_{VALORES,RESULTADOS,SOPORTE,PP_PROYECTOS}` y `escribir_parcial(codigo, tabla, filas)`; esta última valida columnas, tipos y llaves únicas.
- `url_espejo(url)`: reemplaza `https://www.envigado.gov.co/sites/` por `https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/`.
- `url_de_ruta(ruta_local)`: busca en `inventario_detallado_IIC.csv` y en `data_raw_IIC/inventario_descargas_indicadores.csv`.
- `TERRITORIOS` y `PERIODOS`.
- `validar_sin_datos_personales(texto)`: heurística de correos, teléfonos y cédulas.
- `pruebas_comun.py`: casos de prueba de `calcular` (NE, Revisar, tope y CM con O = 0).

**Coordinación:**
- Cada Ejecutor escribe **solo** en sus scripts, en sus `parciales/<CODIGO>_*` y en sus `insumos_manual/<SIS>_*`.
- Solo el ensamblador escribe las tablas finales y la matriz diligenciada.
- `config/supuestos.yaml` lo crea el responsable de `comun.py` a partir del borrador; los Ejecutores no lo editan.

### 6.2 Asignación por sistema

| Ejecutor | Scripts (`scripts/indicadores/`) | Insumos manuales | Esfuerzo estimado |
|---|---|---|---|
| **CS** | `cs_pcs.py`, `cs_dsi.py`, `cs_vrc.py`, `cs_rpo.py`, `cs_icrcs.py` (se ejecuta al final y lee los 4 parciales) | `CS_eventos_rendicion.csv`: una fila por evento 2024-2026 con banderas `es_alcalde`, `es_545`, `es_546`, `tiene_acta`, `doc_integral`, `mencion_veedurias`, `pagina_mencion`, `vigencia_rendida`, `fecha_evento` y `fecha_publicacion_doc`. `CS_observaciones_rpo.csv`: una fila por intervención de QA24, QA100, QA25, QAS25, QAS26 y CMD, con `tipo`, `sobre_pdm`, `respuesta_texto`, `fondo_estricto` y `fecha_publicacion_respuesta`. `CS_metas_dsi.csv`: una fila por meta × año con `valor_publicado`, `fuente`, `misma_definicion`, `es_alcaldia` y `fecha_publicacion` | 1,5 a 2 días (lectura de ~20 documentos y ~60 intervenciones) |
| **PC** | `pc_i2.py`, `pc_i3.py`, `pc_i5.py`, `pc_variables_no_viables.py` (I1, I4, I6) | `PC_zonas_actividades.csv` (zona × año × actividad × evidencia), `PC_acciones_formativas.csv` (acción, tipo, público ciudadano o servidor, programada o realizada, página), `PC_procesos.csv`, `PC_actividades_virtuales.csv` | 1 día |
| **PDM** | `pdm_metas.py` (las 7 metas × 2024-2027), `pdm_cm.py`, `pdm_ca.py`, `pdm_variables_no_viables.py` (CS.O auxiliares, CP, CD, ILE.A con calibración, ILE.S) | `PDM_lecturas_pdf.csv` (cifras leídas de PDF o imagen, con página y la nota "leído de imagen" cuando aplique) | 0,5 a 1 día, más descargas |
| **PP** | `pp_epp.py` (2024, 2025 y 2026; Municipio y zonas; contrastes con Hacienda y CUIPO), `pp_variables_no_viables.py` (PE, PEC proxy y universo de `pp_proyectos`) | `PP_correspondencia_2024.csv` (proyecto EFPC24 ↔ rubro de Hacienda; exacta, aproximada o sin correspondencia) | 1 a 1,5 días; 2024 es lo más costoso |
| **Ensamblador** (el Ejecutor PP al terminar, o la sesión principal) | `ensamblar.py`: concatena los parciales, valida llaves y esquema, escribe los CSV, el xlsx y la matriz diligenciada (§5), redacta `resumen_calculos.md` y ejecuta las verificaciones | — | 0,5 a 1 día |

`scripts/epp_pp_2025.py` y sus salidas en `outputs/disponibilidad/` **no se modifican**. `pp_epp.py` reimplementa o importa su lógica, y debe reproducir 98,83 al peso como prueba de regresión.

### 6.3 Descargas nuevas

Destino: `data_raw_IIC/06_descargas_indicadores/<emisor>/`, ignorado por git; nunca `/tmp`. Cada descarga va en su propio directorio vacío y se registra en `data_raw_IIC/inventario_descargas_indicadores.csv`, con las mismas 11 columnas del inventario. El inventario original no se modifica.

| # | Descarga | URL | Tamaño (HEAD del 9-nov-2026) | Para | Responsable |
|---|---|---|---|---|---|
| 1 | DANE, población municipal por área 2020-2035 (post-COVID) | `https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/DCD-area-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx` | 1,9 MB | CM (escenario), CS.O auxiliar | PDM |
| 2 | DANE, población municipal por área 2018-2042 (VP) | `…/Municipal/PPED-AreaMun-2018-2042_VP.xlsx` | 3,9 MB | CM (escenario) | PDM |
| 3 | DANE, área, sexo y edad simple 2020-2035 (post-COVID) | `…/Municipal/DCD-area-sexo-edad-proypoblacion-Mun-2020-2035-ActPostCOVID-19.xlsx` | 69,5 MB | ILE.A (0-6) | PDM |
| 3b | (opcional) DANE, área, sexo y edad 2018-2042 (VP) | `…/Municipal/PPED-AreaSexoEdadMun-2018-2042_VP.xlsx` | 131,6 MB | ILE.A con la serie VP | PDM |
| 4 | MinSalud, cifras de afiliación dic-2025 y ago-2026 | `https://minsalud.gov.co/sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip`; `…/VP/RBC/cifras-afiliacion-salud-agosto2026.zip` | 1,3 MB + 1,3 MB | PDM.CS auxiliares | PDM |
| 5 | Policía, homicidios en accidentes de tránsito (Envigado) | `https://www.datos.gov.co/resource/uav5-b85g.csv?cod_muni=05266` (filtrado; < 1 MB) | < 1 MB | CM (escenario `policia_datosgov`) | PDM |
| 6 | **Forensis 2024 (INMLCF)** | `https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf` | Responde **403** a clientes automáticos: **la usuaria debe descargarlo a mano** desde el navegador y dejarlo en `data_raw_IIC/06_descargas_indicadores/inmlcf/`. Tamaño no verificable (estimado decenas de MB) | PDM.CM 2024 (escenario por defecto) | Usuaria y luego PDM |
| 7 | (opcional) Clasificación de planteles ICFES 2023-2025 | `https://www.icfes.gov.co/analisis-de-datos/` (la ruta directa dio 404 en oct-2026) | Desconocido | PDM.ILE.S (variable) | PDM, solo si aparece |
| 8 | (opcional) Rendición de Cuentas 2025, p. 61; Acuerdo 007 de 2025; Acta 096 de 2026 | Espejo `…/content/files/002947/147324_rc2025.pdf`; `https://www.concejoenvigado.gov.co/wp-content/uploads/2025/02/Acuerdo-007-de-2025.pdf`; `…/2026/09/Acta-096-de-2026.pdf` | Desconocido (PDF; estimado < 30 MB) | Contraste de APP 2026; acto de las 13 zonas | PP / PC |

**Total obligatorio:** ≈ 78 MB (≈ 210 MB con los opcionales 3b y 8). Espacio libre: 9,0 GB. Las copias de octubre en `/tmp` ya no existen; hay que volver a descargar.

Mientras Forensis no esté, PDM.CM en el escenario por defecto se registra con la cifra **publicada y citada** (7,25; Forensis p. 440). La celda queda con la nota "verificar contra PDF local" y el valor "Pendiente". No se inventa ninguna página que no se haya visto.

### 6.4 Orden de trabajo (de lo más rápido a lo más costoso)

0. Responsable común: `config/supuestos.yaml`, `comun.py` con sus pruebas e instalación de PyYAML (2-3 h). **Bloquea a los demás.**
1. En paralelo, la primera ola de tareas rápidas, de lectura de xlsx:
   - CS: `cs_pcs.py`.
   - PDM: `pdm_metas.py` y `pdm_ca.py`.
   - PC: `pc_i5.py`.
   - PP: `pp_epp.py` 2025 y 2026, con la regresión 98,83.
2. Segunda ola:
   - PDM: descargas 1-5, `pdm_cm.py` y la calibración de ILE.A.
   - PC: `pc_i2.py` y `pc_i3.py`.
   - CS: `cs_dsi.py`.
   - PP: `pp_variables_no_viables.py` (PE, PEC proxy, `pp_proyectos`).
3. Tercera ola, la más costosa:
   - CS: `cs_vrc.py` y `cs_rpo.py` (clasificación manual).
   - PP: EPP 2024 (correspondencia con Hacienda).
   - PC: variables I1, I4 e I6.
   - PDM: variables de los No viables.
4. `cs_icrcs.py` (cuando estén los 4 parciales de CS).
5. Segunda lectura por muestreo (Analizador): 20 % de las filas de `CS_observaciones_rpo.csv` y de `CS_eventos_rendicion.csv`, y 5 proyectos de `PP_correspondencia_2024.csv`. Las discrepancias se anotan.
6. Ensamblador: tablas finales, xlsx, matriz diligenciada, `resumen_calculos.md` y verificaciones (llaves únicas; ningún "Validado"; R:S y N:P intactos; conteos de soporte = valores; los controles de §1.2 se reproducen).
7. Actualizar `.memory/knowledge_graph.md` y crear `outputs/indicadores/ESTADO.md`.

**Esfuerzo total:** ≈ 4-6 días-agente en paralelo, ≈ 2-3 días de calendario.

---

## 7. Riesgos e inconsistencias

1. **Resultados de la copia mostrará N/D en todo.**
   - Los estados son "Pendiente", así que la columna S da "N/D".
   - Hay que explicarlo en el Léeme y en el resumen: los puntajes preliminares están en `resultados_indicadores`.
   - Para verlos en la matriz habría que validar, y eso lo deciden los docentes.
2. **Fecha de corte.**
   - La matriz solo usa filas con N igual a N4, también en las metas.
   - Si se registran cortes reales (7-ene-2026), EPP 2025 queda N/D en la matriz; de ahí el parámetro `corte_fecha_registro`.
   - La normalización es una decisión provisional.
3. **2026 y semestres.**
   - EPP, CS y PDM exigen "Anual" en Resultados, así que el EPP de 2026 S1 solo existe en el CSV.
   - PC.I3 a I6 sí admiten semestre.
4. **La `fecha_publicacion` del inventario no es confiable.**
   - Ejemplos: Plan Indicativo 2T-2026 = "2022-04-28"; Hacienda dic-2024 = "2024-04-22"; DTS 22-dic-2025 = "2024-05-03"; Q&A 2025 = "2025-09-16", aunque el documento cita el 10-dic-2025.
   - La ventana del 31-mar (Guía 17) y la regla de respuestas al 31-dic deben usar la fecha interna del documento o el orden de la carpeta `content/files/NNNNNN`, anotando el método.
5. **Datos personales** en los cuadros de preguntas y respuestas, la encuesta, el CSV de veedurías, el calendario de 2024 y los PDF de PP 2024. El soporte describe el asunto, nunca a la persona; `validar_sin_datos_personales` se aplica a `descripcion` y `justificacion`.
6. **Subjetividad de las clasificaciones de RPO y VRC** (un solo lector). Se mitiga con rúbrica escrita en el YAML, banderas binarias y segunda lectura por muestreo. Hay que reportar la sensibilidad de los escenarios.
7. **PC.I3 en personas** probablemente da V1 > V2 ("Revisar"). Además, las personas capacitadas en 2025 difieren entre fuentes: 65 en el plan indicativo frente a 170 en el informe de gestión.
8. **Comparabilidad del PDM.**
   - CA = 100 por tope con una serie (POT 5,75) que no se reconcilia con la línea base 3,88; hay que rotularlo.
   - CM usa Forensis, pero la línea base 10,2 no declara fuente.
   - Las poblaciones difieren: 248.304 (Forensis, post-COVID), 260.430 (VP) y 268.990 (MinSalud, ago-2026).
   - Las páginas de la Matriz (518, 599, 607…) no corresponden al PDM de la Alcaldía; usar las de `notas_PDM.md`.
   - La Matriz ubica la mortalidad vial en el programa 3.2.2; según el Apéndice es 3.2.3 (fila 119: línea 3, nodo 2, programa 3).
9. **PP.**
   - No hay marcador PP oficial: la lista reconstruida debe confirmarla Planeación.
   - Hay 3 % de diferencia con CUIPO.
   - En 2024, 13 correspondencias son aproximadas y 3 no tienen correspondencia.
   - 5 proyectos de 2025 difieren del valor EFPC (adiciones).
   - El OPP 2026 tiene dos cortes distintos (5,70 frente a 11,56 mil millones; corregido el 9-nov-2026: la cifra 11.555.041.247 citada en octubre no se reproduce con el informe de Hacienda al 30-jun-2026 ni con ROCA; el OPP correcto es 5.702.528.148 en ambas fuentes).
   - La cofinanciación no es verificable.
   - Zona 1 = NE.
   - PP_Proyectos admite 300 filas, así que solo caben las filas municipales.
10. **openpyxl** puede alterar el libro al guardar. Se mitiga con la verificación de §5.4. La matriz no usa extensiones de validación (sin advertencias), pero el Apéndice 9.1 sí: solo se lee.
11. **PyYAML no está disponible con `python3 -I`**: hay que instalarlo antes de empezar (§6.1).
12. **Forensis 403**: depende de una descarga manual. Sin ella, CM se apoya en la cifra citada, "Pendiente".
13. **Concurrencia**: cuatro Ejecutores podrían chocar. Se mitiga con propiedad exclusiva de archivos (§6.1) y un ensamblador único.
14. **Cifras previas no versionadas** (notas de la Fase 2): son solo control; todo valor final sale de un script.
15. **Inconsistencias heredadas que deben ir en Observaciones:**
    - fecha de la audiencia del Alcalde 2024 (20 frente a 26-nov);
    - fecha de la rendición de PP 2024 (12-ene frente a 12-feb-2025);
    - la línea E5Ld-M02-A02 rotulada "VIGENCIA 2024" en el PPSS 2025;
    - el producto 531 dice "12 zonas", frente a 13 oficiales;
    - la rendición de grupos étnicos de la vigencia 2025, del 3-sep-2026, queda fuera de la ventana del 31-mar;
    - el acta de Salud 2025 está duplicada en el repositorio.
16. **Inconsistencias de la Matriz:**
    - colisión de códigos CS / PDM.CS y MS / ILE.MS;
    - PP declara periodicidad trimestral, pero las listas no admiten trimestres;
    - la columna Territorio no tiene catálogo;
    - los estados "No evaluable" y "No aplica" no están definidos en la Guía;
    - el nombre de PP.CMA.NA difiere entre Matriz y Registro;
    - la fórmula de EPP no tiene tope (OPP > APP da Revisar).
17. **`ESTADO.md` de disponibilidad, §4:** la tabla de conteo por categoría de la versión 1.1 muestra **todos los valores en 0**, error del generador. Los conteos correctos (D 10 · C 7 · P 19 · Proxy 2 · ND 1) están en `revision_fase3.md` §6. Hay que corregirlo en una tarea aparte; está fuera del alcance de esta carpeta.
