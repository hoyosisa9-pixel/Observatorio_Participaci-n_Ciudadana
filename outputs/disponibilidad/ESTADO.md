# ESTADO – Disponibilidad y brecha de información del sistema de indicadores (OCE Envigado)

**Punto de control:** 14-oct-2026 · Repositorio `/workspaces/Observatorio_Participaci-n_Ciudadana` · Rama `main`

Este archivo es autosuficiente: explica la tarea, lo hecho, lo pendiente y cómo retomar.

---

## 1. La tarea (resumen del encargo de la usuaria)

Los profesores **Francisco Correa y Eber** definieron un sistema preliminar de indicadores para entregar a la Alcaldía de Envigado (15-oct). Se reúnen para decidir qué indicadores ajustar según la información que **realmente exista**. El encargo es decir, **variable por variable**, si el dato está disponible, dónde, en qué formato y cuál es la brecha.

**Esta tarea NO calcula el IIC ni ningún índice.** El trabajo previo del repo sobre el IIC (inventario, scripts, hallazgos CUIPO/SECOP/BPIN en `outputs/`) es solo insumo.

- **Matriz de los profesores (NO modificar):** `Documentos guias/Matriz de informacion de indicadores del observatorio ciudadano.xlsx` (hojas Guía, Matriz, Registro, PP_Proyectos, Resultados). Copia de trabajo y volcados en `outputs/disponibilidad/trabajo/`.
- **Alcance:** 4 sistemas (CS control social, PC participación ciudadana, PDM cumplimiento del PDM, PP presupuesto participativo); **39 variables** (Registro) y **23 indicadores e índices** (19 + 4). Periodo objetivo 2024, 2025 y lo publicado de 2026; territorio Municipio (anotar si hay por zona). **Zonas: 13** (regla de la usuaria; confirmada por el PDM p. 558: 13 zonas, 12 comités zonales, zona 1 pendiente de conformar).
- **Escala (etiquetas exactas):** Disponible · Construible · Parcial · Proxy · No disponible (definiciones en `trabajo/mapa_busqueda.md` §0).
- **Reglas:** no inventar valores, URLs ni páginas (lo no verificado se marca "sin verificar"); cada hallazgo con evidencia (ruta o URL + página/tabla/hoja); ningún valor "Validado" (todos "Pendiente"); portal caído = brecha; no reproducir datos personales; cifras leídas de imágenes se marcan "leído de imagen" con página. Ninguna variable puede quedar "No disponible" o "Proxy" sin búsqueda web en fuentes primarias.
- **Método con agentes** (`.claude/agents/analizador.md` y `ejecutor.md`; ambos con todas las herramientas, incluida la web; los subagentes no lanzan subagentes, coordina la sesión principal).

## 2. Fases

| Fase | Responsable | Estado |
|---|---|---|
| 1. Plan y mapa de búsqueda | Analizador | **Terminada** → `trabajo/mapa_busqueda.md` |
| 2. Búsqueda y evidencia por sistema (4 Ejecutores en paralelo) | Ejecutor ×4 | **A medias: CS, PC y PDM terminados; PP NO lanzado** |
| 3. Revisión crítica y viabilidad de los 23 indicadores | Analizador | Pendiente |
| 4. Entregables finales + verificación + memoria | Ejecutor | Pendiente |

**Qué falta exactamente en la Fase 2 (sistema PP):**
- Lanzar el Ejecutor de PP con el mismo formato de prompt que los otros tres (ver §8) y estas instrucciones específicas:
  - reutilizar `outputs/limpieza_bpin_cuipo.md`, `outputs/perfil_llaves_iic.md` y `outputs/puente_codigo_proyecto_bpin.csv`;
  - **no** usar el puente CODIGO_PROYECTO↔BPIN por sufijo (invalidado);
  - usar el plan de acción/ROCA (código de proyecto y BPIN en la misma fila);
  - fijar la vigencia de ejecución (las iniciativas de un año se ejecutan al siguiente);
  - evaluación por zona (2 a 13) **más** consolidado municipal deduplicado (Guía 13).
- **Bloqueo:** la usuaria pidió pasar al Ejecutor de PP "información adicional sobre PP" (4 fuentes verificadas, lo que falta, evaluación por zona + consolidado municipal). **Esa información no llegó a la conversación ni está en el repo.** Hay que pedírsela a la usuaria (o que confirme lanzar sin ella).
- Variables de PP: PP.EPP.OPP, PP.EPP.APP, PP.CEP.PEC, PP.CEP.PE, PP.CMA.NA (+ PA y MP por proyecto en la hoja PP_Proyectos; SC es calculado).

## 3. Archivos generados

| Ruta | Contenido |
|---|---|
| `outputs/disponibilidad/ESTADO.md` | Este punto de control |
| `outputs/disponibilidad/trabajo/mapa_busqueda.md` | Fase 1: escala, verificación de conteos, reglas de la Guía, catálogo de 23 indicadores con fórmulas, metas y páginas, fichas de búsqueda de las 39 variables, inconsistencias (§8) y orden de trabajo |
| `outputs/disponibilidad/trabajo/matriz_copia.xlsx` | Copia byte a byte de la matriz original (verificada con `cmp`) |
| `outputs/disponibilidad/trabajo/guia.txt` | Volcado de la hoja Guía |
| `outputs/disponibilidad/trabajo/matriz.csv` | Volcado de la hoja Matriz (23 indicadores/índices) |
| `outputs/disponibilidad/trabajo/registro.csv` | Volcado de la hoja Registro (39 variables × 4 años = 156 filas) |
| `outputs/disponibilidad/trabajo/variables_{CS,PC,PDM}.csv` | Fase 2: una fila por variable (8, 12 y 14) con 20 columnas: sistema, codigo_indicador, id_variable, nombre, unidad, fuente_sugerida, fuente_encontrada, ruta_o_url, pagina_tabla, anios_disponibles, desagregacion, formato, disponibilidad, valor_encontrado, estado_valor, brecha, accion_sugerida, esfuerzo, verificacion, notas |
| `outputs/disponibilidad/trabajo/fuentes_{CS,PC,PDM}.csv` | Catálogo de fuentes usadas (emisor, formato, periodicidad, URL/ruta, descargable o manual, estado de acceso) |
| `outputs/disponibilidad/trabajo/solicitudes_{CS,PC,PDM}.csv` | Datos a pedir a la Alcaldía, por dependencia |
| `outputs/disponibilidad/trabajo/notas_{CS,PC,PDM}.md` | Hallazgos detallados, cambios frente a la Fase 1, URLs que fallaron, páginas exactas de las metas (PDM) |

Aún **no existen** los entregables finales (Fase 4): `disponibilidad_brecha_indicadores.xlsx` (hojas Variables, Indicadores, Fuentes, Solicitudes), `resumen_reunion.md` y el CSV de la hoja Variables.

## 4. Conteo de variables por categoría

CS, PC y PDM son resultados de la Fase 2 (con búsqueda web). **PP es la clasificación preliminar de la Fase 1** (sin Ejecutor todavía).

| Sistema | Disponible | Construible | Parcial | Proxy | No disponible | Total |
|---|---|---|---|---|---|---|
| CS | 2 | 5 | 1 | 0 | 0 | 8 |
| PC | 1 | 1 | 10 | 0 | 0 | 12 |
| PDM | 7 | 2 | 3 | 1 | 1 | 14 |
| PP (preliminar) | 0 | 3 | 2 | 0 | 0 | 5 |
| **Total** | **10** | **11** | **16** | **1** | **1** | **39** |

De los 10 "Disponible", 7 son las metas del PDM. Solo 3 son valores observados publicados: CS.PCS.EE, CS.PCS.EP y PC.I2.V2.

## 5. Tabla de las 39 variables

| Sistema | ID variable | Categoría | Origen | Fuente principal (ruta o URL) | Brecha |
|---|---|---|---|---|---|
| CS | CS.DSI.MD | Construible | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003611/180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-de-desarrollo-20242027/evaluacionplandesarrollomunicipalcuartotrimestre202431032025.pdf | ...cuartotrimestre202525022026.pdf` | La lista de metas seleccionadas (MS) no esta definida por los docentes. Las evaluaciones PDM publicadas (IV trim. 2024 y 2025) solo cubren indicadores de producto; no… |
| CS | CS.DSI.MS | Construible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` | Falta la decision de los docentes sobre cuales son las metas y cuantas (la Matriz sugiere 10). No es un dato publicable por la Alcaldia. |
| CS | CS.PCS.EE | Disponible | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003611/180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | Dato autodeclarado por la dependencia. El DAP anoto "Sí, pero con observaciones" sobre evidencias 2024 y 2025 y reparos de coherencia con G+. 2026 sin cierre. No se… |
| CS | CS.PCS.EP | Disponible | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | https://www.envigado.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/003611/180505_evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` | La Matriz advierte "no asumir una por año"; la programacion oficial si es 1 por anio, por lo que la advertencia queda resuelta. No se buscaron modificaciones… |
| CS | CS.VRC.EV | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/acta-audiencia-publica-de-rendicion-de-cuentas-2024/2024actarendicioncuentassecretariasalud02docx.pdf | data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-202501-secretaria/2025actarendicioncuentassecretariasalud01final.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026actarendicioncuentassecretariasalud01.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/primera-rendicion-de-cuentas-de-la-secretaria-de-salud/20240321primerarendicincuentassalud-1.pdf | data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/rendicion-de-cuentas-comite-municipal-de-discapacidad-905995/respuesta_rendicion_cuentas_ricardo_gaviria.pdf` | Las actas de Salud dicen que "se anexan listados de asistencia" pero los listados no estan en el repo; sin ellos la presencia es mencion textual y no asistencia… |
| CS | CS.VRC.ET | Construible | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx | data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf | justificaciones de no plan de mejoramiento por evento (ver notas_CS.md) | https://www.envigado.gov.co/rendicion-de-cuentas (sin verificar: portal con render JavaScript)` | No hay un balance publicado que fije el universo de "eventos municipales"; la cifra depende de la definicion (1 a 8 por anio). Fechas inconsistentes en algunos… |
| CS | CS.RPO.OR | Construible | Fase 2 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-rendicion-de-cuentas-secretaria/preguntasyrespuestasrendicioncuentassaludi2025.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026preguntasrespuestasrendicioncuentas01.pdf` | Falta la clasificacion de cada respuesta como "de fondo" vs acuse (lectura manual) y del universo evaluado (¿solo Alcalde o tambien sectoriales?). Una respuesta puede… |
| CS | CS.RPO.OT | Construible | Fase 2 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/rendicion-de-cuentas-2025-envigado-vamos-adelante/preguntas-y-respuestas_rendicion-de-cuentas_2025.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-de-la-rendicion-de-cuentas-100/preguntasyrespuestasrendicioncuentasabril2024.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/preguntas-y-respuestas-rendicion-de-cuentas-secretaria/preguntasyrespuestasrendicioncuentassaludi2025.pdf | data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/plan-territoral-de-salud-2024-2027-en-envigado-vamos/2026preguntasrespuestasrendicioncuentas01.pdf` | Los cuadros no distinguen "sobre el PDM"; hay que excluir felicitaciones y temas ajenos. El cuadro del Alcalde 2025 trae solo 3 intervenciones para un evento amplio,… |
| PC | PC.I1.V1 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay inventario oficial de procesos cerrados ni registro que vincule el aporte con la decisión formal; la matriz exige esa vinculación proceso por proceso. |
| PC | PC.I1.V2 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Falta definición oficial de "proceso participativo cerrado"; riesgo alto de subconteo (consultas sectoriales, comités, instancias no están inventariados). |
| PC | PC.I2.V1 | Construible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/anteproyecto29022024.pdf` | La matriz exige acreditar fecha, zona y acta o asistencia; el repo solo tiene cronogramas y reportes agregados, no actas ni listados por zona. Los planes zonales no… |
| PC | PC.I2.V2 | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | No se obtuvo el texto del Acuerdo 007 de 2025 ni del acto que crea las zonas (sin verificar). El Acuerdo 033 de 2022 no arrojó coincidencia textual de "13 zonas" en… |
| PC | PC.I3.V1 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El reporte público oficial (plan indicativo) cuenta personas formadas (95 en 2024; 186 en 2025), no acciones. Las 21 acciones de 2025 pueden incluir formación para… |
| PC | PC.I3.V2 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay programación oficial de acciones formativas (la meta oficial está en personas); el 25 es una reconstrucción que supone que los cursos programados fueron todos… |
| PC | PC.I4.V1 | Parcial | Fase 2 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Universo de procesos sin inventario oficial; no se auditó si cada publicación comunica decisiones o respuestas; enlaces públicos no revisados al corte por falla del… |
| PC | PC.I4.V2 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Igual que PC.I1.V2; I1.V2 debe ser subconjunto de I4.V2, pero ambos comparten un universo no definido. |
| PC | PC.I5.V1 | Parcial | Fase 2 | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | La matriz exige seleccionar las instancias antes de medir; solo hay cifras de dos instancias sectoriales. Sin calendario de sesiones del CMPC, CTP, comités zonales,… |
| PC | PC.I5.V2 | Parcial | Fase 2 | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | Programación solo de dos instancias del sector salud y solo 2025; no existe programación consolidada de las instancias de participación general. |
| PC | PC.I6.V1 | Parcial | Fase 2 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | No hay relación de actividades con canal virtual para intervenir ni evidencia de funcionamiento por actividad. El portal envigado.gov.co no respondió (error de… |
| PC | PC.I6.V2 | Parcial | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay registro único de actividades participativas realizadas; el calendario 2024 es programación, no ejecución; para 2025 y 2026 no hay calendario publicado en el… |
| PDM | PDM.CS.O | Parcial | Fase 2 | `https://minsalud.gov.co/Proteccion-Social/Paginas/cifras-aseguramiento-salud.aspx (zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip` | La cifra con el método del plan (99,5 en 2022; BDUA-ADRES/SisMaster, probablemente afiliados RS sobre población objetivo Sisbén) no está publicada para 2024-2026: el… |
| PDM | PDM.CS.M | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial.  |
| PDM | PDM.CM.O | Parcial | Fase 2 | `https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf` | Falta 2025 y 2026 de INMLCF/ANSV (Forensis 2025 aún no publicado; el sitio ANSV expiró por tiempo desde el Codespace). Las series difieren en definición: INMLCF… |
| PDM | PDM.CM.M | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CA.O | Parcial | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/8-dimension-funcional-_-espacio-publico.pdf` | No se reconcilia la base del plan (3,88 en 2023) con la AMVA 2022 (4,66) ni con el POT 2024 (5,75): definición de área o versión de inventario distinta. No hay cifra… |
| PDM | PDM.CA.M | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial.  |
| PDM | PDM.CP.O | No disponible | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay cifra ni método oficial. El PDM usa dos magnitudes distintas: un índice numérico (3,23 en 2022, p. 561) y un porcentaje (75, base 2023, meta 100, p. 742); sin… |
| PDM | PDM.CP.M | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CD.O | Proxy | Fase 2 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El indicador de resultado es una cobertura porcentual de servicios deportivos (base 5,3 %, meta 6 %); lo hallado es la oferta de disciplinas. No hay cobertura 2024,… |
| PDM | PDM.CD.M | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial.  |
| PDM | PDM.ILE.A | Construible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | El % de cobertura no está publicado. Hay que dividir el numerador por la población DANE de 0 a 6 años (no descargada) y confirmar que ese es el método de la base 16,3… |
| PDM | PDM.ILE.MA | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial.  |
| PDM | PDM.ILE.S | Construible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | Falta la clasificación por plantel (A+, A, B, C, D) 2023-2025 de las I.E. oficiales de Envigado y el total de I.E. oficiales. Los "Cuadros de salida" Saber 11 del… |
| PDM | PDM.ILE.MS | Disponible | Fase 2 | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial.  |
| PP | PP.EPP.OPP | Construible | Fase 1 (preliminar) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/ano-2024-325557/ejecucion-presupuestal-egresos-al-31122024.pdf (+ plan de acción/ROCA)` | No hay fuente de financiación "PP" identificada (CUIPO sin marcador); sumar obligaciones de los proyectos del catálogo PP; cofinanciación quizá no separable. |
| PP | PP.EPP.APP | Construible | Fase 1 (preliminar) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/proyectosdeinversion2024.xlsx` | Igual que OPP; riesgo de proyectos PP que comparten código con recursos ordinarios. |
| PP | PP.CEP.PEC | Parcial | Fase 1 (preliminar) | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/avance-de-ejecucion-de-presupuestos-participativos-2024/` | 2024 con soportes por zona; 2025-2026 solo avance físico agregado, no "entrega completa verificada". |
| PP | PP.CEP.PE | Construible | Fase 1 (preliminar) | `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf` | Universo enumerable por año y zona; deduplicar proyectos multi-zona y definir "entrega exigible en la vigencia". |
| PP | PP.CMA.NA | Parcial | Fase 1 (preliminar) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/proyectosdeinversion2024.xlsx` | Aplicabilidad marcable para 2024; para 2025-2026 faltan fichas con meta de atención. |

## 6. Indicadores e índices (23)

**Veredicto: pendiente de Fase 3** para los 23 (CS.DSI, CS.PCS, CS.VRC, CS.RPO, CS.ICRCS; PC.I1–I6, PC.ISPC; PDM.CS, PDM.CM, PDM.CA, PDM.CP, PDM.CD, PDM.ILE, PDM.ICPD-6; PP.EPP, PP.CEP, PP.CMA, PP.IREPP). Fórmulas y variables de cada uno en `trabajo/mapa_busqueda.md` §3.

**Regla clave (Guía 15):** en los índices, una entrada faltante no redistribuye pesos y bloquea el índice (N/D). Lectura preliminar sin veredicto:
- **CS.ICRCS**: depende de que los docentes definan universos (eventos municipales, metas seleccionadas, "respuesta de fondo").
- **PC.ISPC**: I1 e I4 dependen de un inventario de procesos que no existe oficialmente.
- **PDM.ICPD-6**: CP.O no está disponible y CD.O es Proxy.
- **PP.IREPP**: pendiente del Ejecutor de PP.

## 7. Hallazgos, inconsistencias y preguntas abiertas

### Hallazgos más importantes
1. **Metas del PDM correctas.** Las 7 metas de la Guía (99,5; 8; 3,88; 100; 6; 19; 14) coinciden en el PDM (Acuerdo 017/2024), el DTS modificado del 22-dic-2025 y el Apéndice 9.1.2. Todas son **meta 2027**; no hay metas anuales de resultado. El proyecto de acuerdo 059/2025 modifica productos, no resultados (en los artículos revisados); su aprobación queda sin verificar.
2. **La búsqueda en fuentes primarias mejoró varias clasificaciones del PDM.**
   - PDM.CS.O: MinSalud publica afiliados al régimen subsidiado (29.794 a dic-2025), pero no la cobertura con el método del plan. La cifra de TerriData (≈12 %) es otra definición y no se compara con 99,5.
   - PDM.CM.O: Forensis 2024 da 18 muertes y 7,25 por 100.000. Para 2025 solo hay dato de la Policía.
   - PDM.CA.O: el POT da 5,75 m²/hab urbano en 2024, pero no cuadra con la línea base del plan (3,88).
   - PDM.CP.O sigue **No disponible**: solo se publican productos y no hay método oficial.
3. **PC carece de un inventario oficial de procesos participativos.** Eso vuelve Parciales a I1, I4 e I6. PC.I3 sí tiene conteo de acciones formativas en 2025 (21 culminadas de 25) y PC.I6.V1 tiene evidencia parcial de canal virtual.
4. **CS es mayormente Construible** desde el Plan Indicativo y los cuadros de preguntas de rendición de cuentas. PCS se cumple 1/1 en 2024 y 2025.
5. **Brecha de acceso transversal.** `www.envigado.gov.co` falla por certificado TLS (no se eludió); solo bajan archivos directos `/content/files/`. También fallaron `mi.envigado.gov.co` (timeout), la ANSV, el Tableau de MinSalud (histórico), la descarga ICFES por plantel y la sección Observatorio de la EFPC (404).
6. **Inconsistencias de cifras entre fuentes.**
   - Personas capacitadas 2025: 65 en el plan indicativo frente a 170 en el informe de gestión.
   - Fecha de la audiencia del Alcalde 2024: 20-nov frente a 26-nov.
   - Rendición de PP 2024: 12-ene frente a 12-feb-2025.
   - Espacio verde: 3,88 / 4,66 / 5,75.

### Inconsistencias de la matriz
Detalle completo en `trabajo/mapa_busqueda.md` §8.1.
- **Páginas del PDM:** las páginas citadas "de la copia del Concejo" (518, 599, 607…) no corresponden a los documentos. Las correctas están en `notas_PDM.md`.
- **Programa de mortalidad vial:** la Matriz lo ubica en el programa 3.2.2; probablemente es el 3.2.3.
- **Códigos repetidos:** "CS" es a la vez el sistema de control social y PDM.CS. "MS" son las metas seleccionadas y también la meta Saber.
- **Periodicidad de PP:** se declara trimestral, pero las listas de periodo no admiten trimestres.
- **Territorio:** no hay catálogo de zonas en Registro, PP_Proyectos ni Resultados.
- **Fecha de corte:** Resultados exige corte exacto al 31-dic.
- **Estados:** la lista de estados ("No evaluable", "No aplica") no está definida en la Guía.
- **Nombre de PP.CMA.NA:** difiere entre la Matriz y Registro.
- **Cofinanciación en EPP:** la regla de excluirla es inverificable, porque ninguna fuente marca recursos PP.
- **Documentos ausentes:** faltan en el repo los 4 anexos Word metodológicos y el DOCX OPCEN.
- **Advertencia de PCS:** dice "no asumir una por año", pero el plan indicativo sí programa 1 por año.

### Preguntas abiertas para los profesores
1. **Universos de conteo:**
   - "Evento municipal de rendición de cuentas": según el criterio, salen de 1 a 8 eventos por año.
   - "Proceso participativo cerrado": el universo exploratorio da 3 procesos en 2024 y 5 en 2025.
   - Instancias seleccionadas para PC.I5.
   - Lista de metas de CS.DSI (la Matriz sugiere 10).
   - Criterio de "respuesta de fondo".
2. **Metas 2027 en años intermedios:** ¿evaluar 2024-2026 contra la meta 2027 o pedir metas anuales al DAP?
3. **Indicadores con dato de otra definición:** ¿aceptar Forensis para CM, el POT para CA y MinSalud con denominador propio para CS?
4. **Indicadores sin dato:** ¿redefinir PDM.CP (sin método oficial) y PDM.CD (solo hay conteo de disciplinas)?
5. **Fechas de corte:** ¿cómo tratar datos con corte distinto al 31-dic y 2026 sin cierre (datos al 2.º trimestre)?

## 8. Cómo retomar (sesión nueva sin memoria)

1. Lee `CLAUDE.md`, `.memory/knowledge_graph.md` (§12) y este archivo. No leas el resto de la conversación: no existe.
2. Instala dependencias si faltan: `pip install openpyxl pdfplumber pypdf pymupdf`. No hay `pdftotext` ni `pdftoppm`. Para cifras en imágenes, renderiza la página con PyMuPDF a PNG y léela con la herramienta Read.
3. **Fase 2 – PP.** Pregunta a la usuaria por la "información adicional sobre PP" (§2) o si lanzar sin ella. Lanza un subagente `ejecutor` con un prompt equivalente al de los otros sistemas:
   - **Insumos:** `trabajo/mapa_busqueda.md` §6.4, `guia.txt`, `registro.csv`, `matriz.csv`, el inventario `data_raw_IIC/inventario_detallado_IIC.csv` y los hallazgos de `outputs/` (CUIPO, SECOP, BPIN).
   - **Reglas:** las de §1.
   - **Salidas:** `trabajo/variables_PP.csv` (mismas 20 columnas), `fuentes_PP.csv`, `solicitudes_PP.csv` y `notas_PP.md`.
   - **Validación:** que los CSV se relean con Python y que haya 5 filas.
4. **Fase 3.** Lanza un subagente `analizador` para:
   - revisar las 39 filas consolidadas (`trabajo/variables_*.csv`): que cada clasificación esté respaldada por la evidencia, que nada sin verificar se presente como verificado y que la escala se aplique igual en los 4 sistemas;
   - clasificar los 23 indicadores e índices (Viable / Viable con ajuste / No viable por ahora) con razón y propuesta de ajuste, diciendo para los índices si se calculan con todas sus entradas;
   - devolver una lista concreta de correcciones.
5. **Fase 4.** Lanza un subagente `ejecutor` para:
   - aplicar las correcciones y generar en `outputs/disponibilidad/`:
     - `disponibilidad_brecha_indicadores.xlsx`, con las hojas Variables (39 filas, columnas del encargo), Indicadores (23), Fuentes (consolidar `fuentes_*.csv`) y Solicitudes (consolidar `solicitudes_*.csv` por dependencia);
     - `resumen_reunion.md`, de 1-2 páginas: semáforo por sistema, qué es viable ya, ajustes propuestos y preguntas abiertas;
     - el CSV de Variables, con `scripts/csv_para_excel.py` si sirve;
   - reabrir el xlsx con openpyxl para verificarlo;
   - actualizar `.memory/knowledge_graph.md` y este ESTADO.md.
6. **Cierre para la usuaria:** resumen de 5 líneas con cuántas variables hay por categoría y los 3 hallazgos principales.
7. **Git:** la usuaria autorizó commit y push de `outputs/disponibilidad/` y `.memory/` para este punto de control. Para commits posteriores, confirmar con ella.
