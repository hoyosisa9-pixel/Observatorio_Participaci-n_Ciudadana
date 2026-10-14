# ESTADO – Disponibilidad y brecha de información del sistema de indicadores (OCE Envigado)

**Punto de control:** 14-oct-2026 (actualizado tras la Fase 2 completa) · Repositorio `/workspaces/Observatorio_Participaci-n_Ciudadana` · Rama `main`

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
| 2. Búsqueda y evidencia por sistema (4 Ejecutores) | Ejecutor ×4 | **Terminada: 39/39 variables (CS 8, PC 12, PDM 14, PP 5)** |
| 3. Revisión crítica y viabilidad de los 23 indicadores | Analizador | **Pendiente** (un primer intento se detuvo sin dejar archivos) |
| 4. Entregables finales + verificación + memoria | Ejecutor | Pendiente |

**PP (Fase 2) se hizo con la información adicional del profesor Francisco.**
- **Fuentes verificadas por su revisión previa** (registradas en `fuentes_PP.csv`, sin volver a buscarlas):
  - Listado PP 2025 de la EFPC.
  - Acuerdo 007 de 2025.
  - Rendición de Cuentas 2025, p. 61: 66 proyectos priorizados por $24.519 M para 2026.
  - Acta 096 de 2026, p. 3.
- **Conclusión previa del profesor:** EPP, CEP y CMA en N/D.
- **Lo que encontró el Ejecutor:**
  - Los 61 proyectos "ZONA n" del plan de acción del IV trimestre de 2025 cruzan 1 a 1 con el listado EFPC 2025.
  - Con ese cruce se reconstruyen APP y OPP 2025 (≈ $20.929 M / $20.684 M, corte 7-ene-2026).
  - APP 2026 suma $24.519,7 M, igual a la rendición.
  - 2024 es aproximado: 47 de 50 proyectos, desde los rubros del PDF de Hacienda.
- **Puntos sin resolver:**
  - diferencia de 3 % con CUIPO;
  - 3 proyectos no concilian con Hacienda;
  - falta decidir si el proyecto 0553 (costo del proceso, $776,96 M) entra en el cálculo;
  - PEC solo tiene avance físico autorreportado (Proxy).
- **Cobertura territorial:** zona 1 sin proyectos; zonas 2 a 13 con la misma categoría que el consolidado municipal.
- Detalle en `trabajo/notas_PP.md`. **La Fase 3 debe evaluar críticamente esta reconstrucción frente a la conclusión del profesor.**

**Qué falta:** Fase 3 (revisión de las 39 filas, veredicto de los 23 indicadores, correcciones, lista de documentos no descargables) y Fase 4 (entregables). Ver §8.

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
| `outputs/disponibilidad/trabajo/fuentes_{CS,PC,PDM,PP}.csv` | Catálogo de fuentes usadas (emisor, formato, periodicidad, URL/ruta, descargable o manual, estado de acceso) |
| `outputs/disponibilidad/trabajo/solicitudes_{CS,PC,PDM,PP}.csv` | Datos a pedir a la Alcaldía, por dependencia |
| `outputs/disponibilidad/trabajo/notas_{CS,PC,PDM,PP}.md` | Hallazgos detallados, cambios frente a la Fase 1, URLs que fallaron, páginas exactas de las metas (PDM), PA/MP y disponibilidad por zona (PP) |

Aún **no existen** los entregables finales (Fase 4): `disponibilidad_brecha_indicadores.xlsx` (hojas Variables, Indicadores, Fuentes, Solicitudes), `resumen_reunion.md` y el CSV de la hoja Variables.

## 4. Conteo de variables por categoría

Resultado de la Fase 2 (los 4 sistemas con búsqueda en repo y web). Sujeto a la revisión crítica de la Fase 3.

| Sistema | Disponible | Construible | Parcial | Proxy | No disponible | Total |
|---|---|---|---|---|---|---|
| CS | 2 | 5 | 1 | 0 | 0 | 8 |
| PC | 1 | 1 | 10 | 0 | 0 | 12 |
| PDM | 7 | 2 | 3 | 1 | 1 | 14 |
| PP | 0 | 3 | 1 | 1 | 0 | 5 |
| **Total** | **10** | **11** | **15** | **2** | **1** | **39** |

De los 10 "Disponible", 7 son las metas del PDM; solo 3 son valores observados publicados (CS.PCS.EE, CS.PCS.EP, PC.I2.V2).

## 5. Tabla de las 39 variables

Detalle completo (valor encontrado, años, desagregación, formato, acción sugerida, esfuerzo, verificación) en `trabajo/variables_<SIS>.csv`.

| Sistema | ID variable | Categoría | Fuente principal (ruta o URL) | Brecha |
|---|---|---|---|---|
| CS | CS.DSI.MD | Construible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | La lista de metas seleccionadas (MS) no esta definida por los docentes. Las evaluaciones PDM publicadas (IV trim. 2024 y 2025) solo cubren indicadores de producto; no… |
| CS | CS.DSI.MS | Construible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` | Falta la decision de los docentes sobre cuales son las metas y cuantas (la Matriz sugiere 10). No es un dato publicable por la Alcaldia. |
| CS | CS.PCS.EE | Disponible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | Dato autodeclarado por la dependencia. El DAP anoto "Sí, pero con observaciones" sobre evidencias 2024 y 2025 y reparos de coherencia con G+. 2026 sin cierre. No se… |
| CS | CS.PCS.EP | Disponible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | La Matriz advierte "no asumir una por año"; la programacion oficial si es 1 por anio, por lo que la advertencia queda resuelta. No se buscaron modificaciones… |
| CS | CS.VRC.EV | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | Las actas de Salud dicen que "se anexan listados de asistencia" pero los listados no estan en el repo; sin ellos la presencia es mencion textual y no asistencia… |
| CS | CS.VRC.ET | Construible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | No hay un balance publicado que fije el universo de "eventos municipales"; la cifra depende de la definicion (1 a 8 por anio). Fechas inconsistentes en algunos… |
| CS | CS.RPO.OR | Construible | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Falta la clasificacion de cada respuesta como "de fondo" vs acuse (lectura manual) y del universo evaluado (¿solo Alcalde o tambien sectoriales?). Una respuesta puede… |
| CS | CS.RPO.OT | Construible | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Los cuadros no distinguen "sobre el PDM"; hay que excluir felicitaciones y temas ajenos. El cuadro del Alcalde 2025 trae solo 3 intervenciones para un evento amplio,… |
| PC | PC.I1.V1 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay inventario oficial de procesos cerrados ni registro que vincule el aporte con la decisión formal; la matriz exige esa vinculación proceso por proceso. |
| PC | PC.I1.V2 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Falta definición oficial de "proceso participativo cerrado"; riesgo alto de subconteo (consultas sectoriales, comités, instancias no están inventariados). |
| PC | PC.I2.V1 | Construible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/anteproyecto29022024.pdf` | La matriz exige acreditar fecha, zona y acta o asistencia; el repo solo tiene cronogramas y reportes agregados, no actas ni listados por zona. Los planes zonales no… |
| PC | PC.I2.V2 | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | No se obtuvo el texto del Acuerdo 007 de 2025 ni del acto que crea las zonas (sin verificar). El Acuerdo 033 de 2022 no arrojó coincidencia textual de "13 zonas" en… |
| PC | PC.I3.V1 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El reporte público oficial (plan indicativo) cuenta personas formadas (95 en 2024; 186 en 2025), no acciones. Las 21 acciones de 2025 pueden incluir formación para… |
| PC | PC.I3.V2 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay programación oficial de acciones formativas (la meta oficial está en personas); el 25 es una reconstrucción que supone que los cursos programados fueron todos… |
| PC | PC.I4.V1 | Parcial | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Universo de procesos sin inventario oficial; no se auditó si cada publicación comunica decisiones o respuestas; enlaces públicos no revisados al corte por falla del… |
| PC | PC.I4.V2 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Igual que PC.I1.V2; I1.V2 debe ser subconjunto de I4.V2, pero ambos comparten un universo no definido. |
| PC | PC.I5.V1 | Parcial | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | La matriz exige seleccionar las instancias antes de medir; solo hay cifras de dos instancias sectoriales. Sin calendario de sesiones del CMPC, CTP, comités zonales,… |
| PC | PC.I5.V2 | Parcial | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | Programación solo de dos instancias del sector salud y solo 2025; no existe programación consolidada de las instancias de participación general. |
| PC | PC.I6.V1 | Parcial | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | No hay relación de actividades con canal virtual para intervenir ni evidencia de funcionamiento por actividad. El portal envigado.gov.co no respondió (error de… |
| PC | PC.I6.V2 | Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay registro único de actividades participativas realizadas; el calendario 2024 es programación, no ejecución; para 2025 y 2026 no hay calendario publicado en el… |
| PDM | PDM.CS.O | Parcial | `https://minsalud.gov.co/Proteccion-Social/Paginas/cifras-aseguramiento-salud.aspx (zip dic-2025: /sites/rid/Lists/BibliotecaDigital/RIDE/VP/DOA/cifras-afiliacion-salud-dic-2025.zip` | La cifra con el método del plan (99,5 en 2022; BDUA-ADRES/SisMaster, probablemente afiliados RS sobre población objetivo Sisbén) no está publicada para 2024-2026: el… |
| PDM | PDM.CS.M | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.CM.O | Parcial | `https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf` | Falta 2025 y 2026 de INMLCF/ANSV (Forensis 2025 aún no publicado; el sitio ANSV expiró por tiempo desde el Codespace). Las series difieren en definición: INMLCF… |
| PDM | PDM.CM.M | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CA.O | Parcial | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/8-dimension-funcional-_-espacio-publico.pdf` | No se reconcilia la base del plan (3,88 en 2023) con la AMVA 2022 (4,66) ni con el POT 2024 (5,75): definición de área o versión de inventario distinta. No hay cifra… |
| PDM | PDM.CA.M | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.CP.O | No disponible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay cifra ni método oficial. El PDM usa dos magnitudes distintas: un índice numérico (3,23 en 2022, p. 561) y un porcentaje (75, base 2023, meta 100, p. 742); sin… |
| PDM | PDM.CP.M | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CD.O | Proxy | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El indicador de resultado es una cobertura porcentual de servicios deportivos (base 5,3 %, meta 6 %); lo hallado es la oferta de disciplinas. No hay cobertura 2024,… |
| PDM | PDM.CD.M | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.ILE.A | Construible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | El % de cobertura no está publicado. Hay que dividir el numerador por la población DANE de 0 a 6 años (no descargada) y confirmar que ese es el método de la base 16,3… |
| PDM | PDM.ILE.MA | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.ILE.S | Construible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | Falta la clasificación por plantel (A+, A, B, C, D) 2023-2025 de las I.E. oficiales de Envigado y el total de I.E. oficiales. Los "Cuadros de salida" Saber 11 del… |
| PDM | PDM.ILE.MS | Disponible | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PP | PP.EPP.OPP | Construible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx (hojas PlanAcción27.01.2026 y ROCA07.01.26)` | No existe marcador de PP en las fuentes: el universo se reconstruye cruzando el catálogo EFPC con los códigos de proyecto del plan de acción. Para 2024 no hay plan de… |
| PP | PP.EPP.APP | Construible | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx` | Mismas que OPP: sin marcador PP; universo reconstruido; 2024 sin plan de acción por proyecto; dos fuentes de valor (catálogo = valor del proyecto; apropiación… |
| PP | PP.CEP.PEC | Proxy | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx (hoja PlanAcción27.01.2026, filas con fecha 2025-12-31, col. EJECUCIÓN FÍSICA NETA)` | El indicador exige entrega exigible completa y verificada con acta de recibo, informe final o supervisión. No se encontró ninguno publicado por proyecto en el repo ni… |
| PP | PP.CEP.PE | Construible | `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf` | Falta definir "entrega exigible en la vigencia" (los 61 de 2025 son "en ejecución"; la entrega exigible puede ser parcial o cruzar vigencias). No hay listado oficial… |
| PP | PP.CMA.NA | Parcial | `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2025.pdf` | Falta la ficha del proyecto con meta anual de personas por proyecto. El plan de acción mide actividades, porcentajes y paquetes, no personas únicas. Las cifras de… |

## 6. Indicadores e índices (23)

**Veredicto: pendiente de Fase 3** para los 23 (CS.DSI, CS.PCS, CS.VRC, CS.RPO, CS.ICRCS; PC.I1–I6, PC.ISPC; PDM.CS, PDM.CM, PDM.CA, PDM.CP, PDM.CD, PDM.ILE, PDM.ICPD-6; PP.EPP, PP.CEP, PP.CMA, PP.IREPP). Fórmulas y variables de cada uno en `trabajo/mapa_busqueda.md` §3.

**Regla clave (Guía 15):** en los índices, una entrada faltante no redistribuye pesos y bloquea el índice (N/D). Lectura preliminar sin veredicto:
- **CS.ICRCS**: depende de que los docentes definan universos (eventos municipales, metas seleccionadas, "respuesta de fondo").
- **PC.ISPC**: I1 e I4 dependen de un inventario de procesos que no existe oficialmente.
- **PDM.ICPD-6**: CP.O no está disponible y CD.O es Proxy.
- **PP.IREPP**: EPP construible (2025 completo, 2026 parcial, 2024 aproximado); CEP solo con proxy; CMA parcial.

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
3. **Fase 2:** terminada (no repetir).
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
