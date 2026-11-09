# ESTADO – Disponibilidad y brecha de información del sistema de indicadores (OCE Envigado)

**Punto de control:** 14-oct-2026 (actualizado el 9-nov-2026: versión 1.1 tras el control de calidad) · Repositorio `/workspaces/Observatorio_Participaci-n_Ciudadana` · Rama `main`

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
| 2. Búsqueda y evidencia por sistema (4 Ejecutores) | Ejecutor ×4 | **Terminada: 39/39 variables** |
| 3. Revisión crítica y viabilidad de los 23 indicadores | Analizador | **Terminada** → `trabajo/revision_fase3.md`, `trabajo/indicadores_fase3.csv` (dos intentos previos se cortaron por límite de uso) |
| 4. Entregables finales + verificación + memoria | Ejecutor + sesión principal | **Terminada** → Excel, CSV, `resumen_reunion.md`, `interno/documentos_por_descargar.md`, `scripts/generar_disponibilidad.py` |

**La tarea está completa.** Lo que queda abierto depende de los profesores (preguntas de §7) o de que la usuaria descargue manualmente los documentos de `interno/documentos_por_descargar.md`.

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


## 2b. Versión 1.1 (9-nov-2026): correcciones del control de calidad

Sin cambios en clasificaciones, cifras ni veredictos (10/7/19/2/1; 1/10/12; EPP 98,83 %).
- **Enlaces:** los documentos del portal de la Alcaldía (`www.envigado.gov.co/sites/.../content/files/...`) apuntan al dominio espejo `alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co`, con la misma ruta (36 de 36 responden con certificado válido; 38 enlaces espejo en total). Las 15 páginas de sección del portal, sin espejo, quedan con la nota "puede requerir abrirse desde otro navegador" (Excel: Notas y hoja Enlaces, columna Nota).
- **ADRES:** `https://www.adres.gov.co/entidades-territoriales/bdua/reportes` reemplaza la dirección anterior (404). **Forensis 2024:** se mantiene, con la nota "verificar en navegador" (responde 403 a clientes automáticos).
- **PP.CEP.PEC:** la nota ya no dice "se reclasifica como Proxy"; ahora indica que la categoría final es Parcial.
- **Portada:** parámetros fijos en `generar_informe_latex.py` (`VERSION`, `FECHA_CORTE`): "Versión 1.1 – 9 de noviembre de 2026" y "Fecha de corte de la información: 14 de octubre de 2026".
- **Documentos internos:** `resumen_reunion.md`, `resumen_reunion.docx` y `documentos_por_descargar.md` se movieron a `outputs/disponibilidad/interno/` (con `LEEME.md`); no se entregan. El informe formal sigue leyendo `interno/documentos_por_descargar.md` para su sección 8.1.
- **Espacio:** los 5 ZIP originales están respaldados en Google Drive (carpeta Bases_Datos_IIC_Envigado); se eliminaron del Codespace el 9-nov-2026 por espacio.
- **Entregables vigentes:** `disponibilidad_brecha_indicadores.xlsx`, `informe_disponibilidad.docx`, `informe_latex/informe_disponibilidad.pdf` e `informe_disponibilidad_overleaf.zip`.

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

**Entregables finales (Fase 4):**

| Ruta | Contenido |
|---|---|
| `outputs/disponibilidad/disponibilidad_brecha_indicadores.xlsx` | Versión para los profesores: hojas Léeme, Indicadores, Variables (39), Fuentes (66), Solicitudes (32) y Enlaces (218 hipervínculos). Rutas internas reemplazadas por la URL de origen del inventario (ruta interna en columna aparte al final), lenguaje sin jerga, tildes, encabezados en español, sistema con nombre completo, estado de acceso y verificación normalizados, columna Resumen (≤ 200 caracteres), celdas ≤ 600 caracteres (excedente en Notas), colores por categoría y veredicto; EPP marcado "preliminar, por validar", Indicadores (23), Fuentes (66, deduplicadas por URL o ruta) y Solicitudes (32, por dependencia). Verificado con openpyxl |
| `outputs/disponibilidad/disponibilidad_variables.csv` | Hoja Variables en CSV (UTF-8, coma) |
| `outputs/disponibilidad/disponibilidad_variables_excel.csv` | Ídem para Excel (`;`, UTF-8 con BOM) |
| `outputs/disponibilidad/interno/resumen_reunion.md` (interno, no se entrega) | Resumen de 1-2 páginas para los profesores: semáforo, viables, ajustes, EPP 2025, preguntas, documentos por descargar |
| `outputs/disponibilidad/interno/documentos_por_descargar.md` (interno, no se entrega) | 8 documentos a descargar o leer manualmente, con su link y la variable que cambiarían |
| `outputs/disponibilidad/interno/resumen_reunion.docx` (interno, no se entrega) | Versión Word del resumen (generada con `python3 scripts/md_a_docx.py outputs/disponibilidad/resumen_reunion.md`; sin `-I` porque python-docx está en el site-packages del usuario) |
| `scripts/epp_pp_2025.py` → `outputs/disponibilidad/epp_pp_2025.md`, `epp_pp_2025_zonas.csv`, `epp_pp_2025_proyectos.csv` | Cálculo reproducible de EPP 2025, **preliminar, por validar**: 61 proyectos "ZONA n" del plan de acción IV-trim-2025 (sin el 0553) contrastados con el listado EFPC (61, misma cantidad por zona, 56 iguales al peso). APP $20.928.690.218; OPP $20.684.427.720; EPP 98,83 % (98,86 % con el 0553). Confirmó la cifra de la Fase 2 |
| `outputs/disponibilidad/informe_latex/informe_disponibilidad.tex` (+ `.pdf` de control) y `informe_disponibilidad_overleaf.zip` | Informe formal en LaTeX para Overleaf (New Project → Upload Project). Lo genera `python3 -I scripts/generar_informe_latex.py` leyendo el Excel, `epp_pp_2025_zonas.csv`, `interno/documentos_por_descargar.md`, este ESTADO.md (inconsistencias) y `trabajo/revision_fase3.md` §4 (decisiones). Compila sin errores con tectonic (XeTeX); compatible con pdfLaTeX |
| `outputs/disponibilidad/informe_disponibilidad.docx` | Informe en Word para el informe a Envigado, mismo contenido que el PDF de LaTeX. Lo genera `python3 scripts/generar_informe_docx.py` (sin `-I`: python-docx está en ~/.local) traduciendo el mismo .tex que arma `generar_informe_latex.py`: estilos nativos, campo TOC, tablas nativas con colores del .tex, hipervínculos |
| `scripts/generar_disponibilidad.py` | Regenera el Excel y los CSV desde `trabajo/` aplicando las 20 correcciones. Ejecutar con `python3 -I scripts/generar_disponibilidad.py`; se detiene si el conteo no cuadra |
| `outputs/disponibilidad/trabajo/revision_fase3.md` | Fase 3: muestreo, 20 correcciones, veredictos, ajustes, preguntas, documentos y conteo final |
| `outputs/disponibilidad/trabajo/indicadores_fase3.csv` | Veredicto de los 23 indicadores e índices |

## 4. Conteo de variables por categoría (final, tras Fase 3)

Fuente: hoja Variables de `disponibilidad_brecha_indicadores.xlsx` (Fase 2 + 20 correcciones de `trabajo/revision_fase3.md` §1).

| Sistema | Disponible | Construible | Parcial | Proxy | No disponible | Total |
|---|---|---|---|---|---|---|
| CS | 2 | 3 | 3 | 0 | 0 | 8 |
| PC | 1 | 0 | 11 | 0 | 0 | 12 |
| PDM | 7 | 1 | 3 | 2 | 1 | 14 |
| PP | 0 | 3 | 2 | 0 | 0 | 5 |
| **Total** | **10** | **7** | **19** | **2** | **1** | **39** |

Cambios frente a la Fase 2 (D 10 · C 11 · P 15 · Proxy 2 · ND 1):
- CS.RPO.OR y CS.RPO.OT: Construible → Parcial.
- PC.I2.V1: Construible → Parcial.
- PDM.ILE.S: Construible → Parcial.
- PDM.CS.O: Parcial → Proxy.
- PP.CEP.PEC: Proxy → Parcial.

De los 10 Disponibles, 7 son metas del PDM; solo 3 son valores observados (CS.PCS.EE, CS.PCS.EP, PC.I2.V2).

**Reglas de coherencia de la escala (Fase 3):**
- El autorreporte cuenta como Disponible si la variable pide solo "reportado"; si el protocolo exige acta o informe, es Parcial.
- Proxy queda solo para datos con otra definición, unidad o universo.
- Las metas del PDM son Disponibles para el horizonte 2027.

## 5. Tabla de las 39 variables

Detalle completo en la hoja Variables del Excel y en `disponibilidad_variables.csv`.

| Sistema | ID variable | Categoría | Fuente principal (ruta o URL) | Brecha |
|---|---|---|---|---|
| Control social | CS.DSI.MD | 2024; 2025; 2026 (corte 2T, sin cierre anual) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | La lista de metas seleccionadas (MS) no esta definida por los docentes. Las evaluaciones PDM publicadas (IV trim. 2024 y 2025) solo cubren indicadores de producto; no… |
| Control social | CS.DSI.MS | Definicion fija para el periodo (lista sin cambios) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/apendice91planindicativo27052024.xlsx` | Falta la decision de los docentes sobre cuales son las metas y cuantas (la Matriz sugiere 10). No es un dato publicable por la Alcaldia. |
| Control social | CS.PCS.EE | 2024; 2025; 2026 (avance 0 al 2T; en planeacion en el T1) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | Dato autodeclarado por la dependencia. El DAP anoto "Sí, pero con observaciones" sobre evidencias 2024 y 2025 y reparos de coherencia con G+. 2026 sin cierre. No se… |
| Control social | CS.PCS.EP | 2024; 2025; 2026; 2027 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | La Matriz advierte "no asumir una por año"; la programacion oficial si es 1 por anio, por lo que la advertencia queda resuelta. No se buscaron modificaciones… |
| Control social | CS.VRC.EV | 2024; 2025 (solo cifra institucional); 2026 (Salud, 18-mar) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | Las actas de Salud dicen que "se anexan listados de asistencia" pero los listados no estan en el repo; sin ellos la presencia es mencion textual y no asistencia… |
| Control social | CS.VRC.ET | 2024; 2025; 2026 (parcial) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/informe-de-gestion-alcaldia-de-envigado/evaluaciondefinitivaplanindicativoproductos2026segundotrimestre20260730.xlsx` | No hay un balance publicado que fije el universo de "eventos municipales"; la cifra depende de la definicion (1 a 8 por anio). Fechas inconsistentes en algunos… |
| Control social | CS.RPO.OR | 2024; 2025; 2026 (solo Salud y Discapacidad) | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Falta la clasificacion de cada respuesta como "de fondo" vs acuse (lectura manual) y del universo evaluado (¿solo Alcalde o tambien sectoriales?). Una respuesta puede… |
| Control social | CS.RPO.OT | 2024; 2025; 2026 (parcial) | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Los cuadros no distinguen "sobre el PDM"; hay que excluir felicitaciones y temas ajenos. El cuadro del Alcalde 2025 trae solo 3 intervenciones para un evento amplio,… |
| PC | PC.I1.V1 | 2024; 2025 (2026 en curso, sin cierre) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay inventario oficial de procesos cerrados ni registro que vincule el aporte con la decisión formal; la matriz exige esa vinculación proceso por proceso. |
| PC | PC.I1.V2 | 2024; 2025 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Falta definición oficial de "proceso participativo cerrado"; riesgo alto de subconteo (consultas sectoriales, comités, instancias no están inventariados). |
| PC | PC.I2.V1 | 2024; 2025; 2026 solo programado | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/anteproyecto29022024.pdf` | La matriz exige acreditar fecha, zona y acta o asistencia; el repo solo tiene cronogramas y reportes agregados, no actas ni listados por zona. Los planes zonales no… |
| PC | PC.I2.V2 | 2024; 2025; 2026 (vigente) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | No se obtuvo el texto del Acuerdo 007 de 2025 ni del acto que crea las zonas (sin verificar). El Acuerdo 033 de 2022 no arrojó coincidencia textual de "13 zonas" en… |
| PC | PC.I3.V1 | 2024 (construible); 2025 (publicado); 2026 solo oferta sin fechas | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El reporte público oficial (plan indicativo) cuenta personas formadas (95 en 2024; 186 en 2025), no acciones. Las 21 acciones de 2025 pueden incluir formación para… |
| PC | PC.I3.V2 | 2025 (inferido) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay programación oficial de acciones formativas (la meta oficial está en personas); el 25 es una reconstrucción que supone que los cursos programados fueron todos… |
| PC | PC.I4.V1 | 2024; 2025 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | Universo de procesos sin inventario oficial; no se auditó si cada publicación comunica decisiones o respuestas; enlaces públicos no revisados al corte por falla del… |
| PC | PC.I4.V2 | 2024; 2025 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | Igual que PC.I1.V2; I1.V2 debe ser subconjunto de I4.V2, pero ambos comparten un universo no definido. |
| PC | PC.I5.V1 | 2025 (COPACO y CMSSS); 2024 solo enlaces a carpetas privadas | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | La matriz exige seleccionar las instancias antes de medir; solo hay cifras de dos instancias sectoriales. Sin calendario de sesiones del CMPC, CTP, comités zonales,… |
| PC | PC.I5.V2 | 2025 | `data_raw_IIC/01_propuestas_y_acuerdos/envigado_gov_co/seguimiento-al-plan-de-accion-de-la-politica-de-participacion-596131/pps190ppps20251231mu000000005266.xlsx` | Programación solo de dos instancias del sector salud y solo 2025; no existe programación consolidada de las instancias de participación general. |
| PC | PC.I6.V1 | 2024; 2025 | `data_raw_IIC/03_actas_y_comites_tecnicos/envigado_gov_co/audiencia-publica-de-rendicion-de-cuentas-2024-alcalde/129738_preguntasyrecomendacionesrendiciondecuentas2024.pdf` | No hay relación de actividades con canal virtual para intervenir ni evidencia de funcionamiento por actividad. El portal envigado.gov.co no respondió (error de… |
| PC | PC.I6.V2 | 2024 (calendario programado y reporte); 2025 (reporte); 2026 parcial al 2.º trimestre | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay registro único de actividades participativas realizadas; el calendario 2024 es programación, no ejecución; para 2025 y 2026 no hay calendario publicado en el… |
| PDM | PDM.CS.O | Dic-2025 y ago-2026 (MinSalud, descarga verificada); 2024 y 2025 solo vía TerriData (RS/población); 2024 mensual de MinSalud solo en tablero Tableau (no accesible) | `https://minsalud.gov.co/Proteccion-Social/Paginas/cifras-aseguramiento-salud.aspx` | La cifra con el método del plan (99,5 en 2022; BDUA-ADRES/SisMaster, probablemente afiliados RS sobre población objetivo Sisbén) no está publicada para 2024-2026: el… |
| PDM | PDM.CS.M | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.CM.O | 2024 (INMLCF; ANSV según prensa 22-ene-2025); 2025 y 2026 parcial solo Policía; 2025 INMLCF/ANSV no localizado | `https://www.medicinalegal.gov.co/documents/20143/1269218/Forensis_2024.pdf` | Falta 2025 y 2026 de INMLCF/ANSV (Forensis 2025 aún no publicado; el sitio ANSV expiró por tiempo desde el Codespace). Las series difieren en definición: INMLCF… |
| PDM | PDM.CM.M | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CA.O | 2024 (cálculo del diagnóstico POT, publicado 2025); línea base AMVA 2022; 2025 y 2026 no hallados | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/documento-diagnostico-pot/8-dimension-funcional-_-espacio-publico.pdf` | No se reconcilia la base del plan (3,88 en 2023) con la AMVA 2022 (4,66) ni con el POT 2024 (5,75): definición de área o versión de inventario distinta. No hay cifra… |
| PDM | PDM.CA.M | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.CP.O | Ninguno 2024-2026 | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | No hay cifra ni método oficial. El PDM usa dos magnitudes distintas: un índice numérico (3,23 en 2022, p. 561) y un porcentaje (75, base 2023, meta 100, p. 742); sin… |
| PDM | PDM.CP.M | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. Línea base sin… |
| PDM | PDM.CD.O | 2025 (producto, no resultado) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/anon-2026/2025informegestiondependencias.pdf` | El indicador de resultado es una cobertura porcentual de servicios deportivos (base 5,3 %, meta 6 %); lo hallado es la oferta de disciplinas. No hay cobertura 2024,… |
| PDM | PDM.CD.M | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.ILE.A | Numerador 2024 y 2025; el porcentaje de cobertura no está publicado | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | El % de cobertura no está publicado. Hay que dividir el numerador por la población DANE de 0 a 6 años (no descargada) y confirmar que ese es el método de la base 16,3… |
| PDM | PDM.ILE.MA | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PDM | PDM.ILE.S | 2024 y 2025 (menciones de casos); el % no está publicado | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/secretaria-de-educacion-426550/ppt-balance-pdm-2024_educacion_oficial.pdf` | Falta la clasificación por plantel (A+, A, B, C, D) 2023-2025 de las I.E. oficiales de Envigado y el total de I.E. oficiales. Los "Cuadros de salida" Saber 11 del… |
| PDM | PDM.ILE.MS | Meta 2027 (única; no hay metas anuales de resultado) | `data_raw_IIC/05_fuentes_externas_y_pot/envigado_gov_co/plan-de-desarrollo-20242027/plandesarrollo20242027.pdf` | Solo meta final 2027; el PDM no publica metas anuales de resultado, así que la lectura anual de la Guía (punto 9) carece de meta intermedia oficial. |
| PP | PP.EPP.OPP | 2025 completo (corte 07-01-2026); 2026: APP completa, OPP provisional al 30-jun-2026 (no anual); 2024 aproximado (47 de 50 proyectos, 13 correspondencias aproximadas, 3 sin correspondencia): para 2024 la variable es Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx (hojas PlanAcción27.01.2026 y ROCA07.01.26)` | No existe marcador de PP en las fuentes: el universo se reconstruye cruzando el catálogo EFPC con los códigos de proyecto del plan de acción. Para 2024 no hay plan de… |
| PP | PP.EPP.APP | 2025 completo (corte 07-01-2026); 2026: APP completa, OPP provisional al 30-jun-2026 (no anual); 2024 aproximado (47 de 50 proyectos, 13 correspondencias aproximadas, 3 sin correspondencia): para 2024 la variable es Parcial | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx` | Mismas que OPP: sin marcador PP; universo reconstruido; 2024 sin plan de acción por proyecto; dos fuentes de valor (catálogo = valor del proyecto; apropiación… |
| PP | PP.CEP.PEC | 2024 (parcial, documentos por zona-operador), 2025 (indicativo), 2026 (solo avance a mitad de año) | `data_raw_IIC/02_ejecucion_y_presupuesto/envigado_gov_co/seguimiento-a-proyectos-2025/evaluaciondefinitivaplanaccionivtri.xlsx (hoja PlanAcción27.01.2026, filas con fecha 2025-12-31, col. EJECUCIÓN FÍSICA NETA)` | El indicador exige entrega exigible completa y verificada con acta de recibo, informe final o supervisión. No se encontró ninguno publicado por proyecto en el repo ni… |
| PP | PP.CEP.PE | 2024, 2025, 2026 | `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2024.pdf` | Falta definir "entrega exigible en la vigencia" (los 61 de 2025 son "en ejecución"; la entrega exigible puede ser parcial o cruzar vigencias). No hay listado oficial… |
| PP | PP.CMA.NA | 2024 (parcial por operador); 2025 y 2026 sin meta de personas | `data_raw_IIC/01_propuestas_y_acuerdos/efpc_envigado_gov_co/efpc_2025-07/Proyectos_PP_2025.pdf` | Falta la ficha del proyecto con meta anual de personas por proyecto. El plan de acción mide actividades, porcentajes y paquetes, no personas únicas. Las cifras de… |

## 6. Indicadores e índices (23) – veredicto de la Fase 3

**Resumen:** 1 Viable, 10 Viables con ajuste y 12 No viables por ahora.

**Índices:** solo CS.ICRCS se calcula con todas sus entradas, y únicamente para 2024 tras definir los universos. PC.ISPC, PDM.ICPD-6 y PP.IREPP quedan bloqueados por la Guía 15.

El ajuste propuesto de cada indicador está en la hoja Indicadores y en `trabajo/revision_fase3.md` §2-§3.

| Sistema | Código | Veredicto | Razón | Años calculables |
|---|---|---|---|---|
| CS | CS.DSI | **Viable con ajuste** | Registros existen (plan indicativo, evaluaciones PDM) pero MS no está definido; con universo de productos DSI = 100 por construcción. | 2024; 2025 (2026 sin cierre) |
| CS | CS.PCS | **Viable** | Programación (1/año) y ejecución reportada (1 en 2024 y 2025) publicadas en el plan indicativo consec. 544; PCS = 100. | 2024; 2025 (2026 al 2T: 0, sin cierre) |
| CS | CS.VRC | **Viable con ajuste** | ET va de 1 a 8 eventos según la definición y la presencia solo consta como mención textual en actas; con presencia desconocida la Matriz manda N/D. | 2024; 2025 (condicionados al universo y al criterio de presencia) |
| CS | CS.RPO | **Viable con ajuste** | El cuadro de preguntas y respuestas del Alcalde 2024 permite contar (17 con respuesta de 22); el de 2025 está incompleto (3) y falta el criterio "de fondo". | 2024 (2025 solo con acta integral) |
| CS | CS.ICRCS | **Viable con ajuste** | Las cuatro entradas son calculables en 2024 una vez fijadas las definiciones; en 2025 RPO y VRC quedan bloqueadas por registros incompletos. | 2024 (con todas las entradas, tras definiciones); 2025 N/D por ahora |
| PC | PC.I1 | **No viable por ahora** | No existe inventario oficial de procesos participativos cerrados ni registro aporte-decisión; el universo 3/5 es una propuesta del Ejecutor. | Ninguno (exploratorio 2024 y 2025) |
| PC | PC.I2 | **Viable con ajuste** | V2 = 13 publicado (PDM p. 558); V1 sale de reportes agregados (12 asambleas PP; 13 POT 2025) sin acta por zona. | 2024; 2025 |
| PC | PC.I3 | **Viable con ajuste** | La programación oficial está en personas, no en acciones; el 25 programadas de 2025 es inferido y mezcla formación a servidores. | 2024; 2025 (opción A); 2025 aproximado (opción B) |
| PC | PC.I4 | **No viable por ahora** | Comparte el universo inexistente de I1 y no se auditó documento por documento; portal envigado.gov.co inaccesible. | Ninguno (exploratorio 2024 y 2025) |
| PC | PC.I5 | **Viable con ajuste** | Solo hay programación y ejecución de dos instancias de salud (COPACO y CMSSS, 8/8 en 2025). | 2025 (acotado a dos instancias) |
| PC | PC.I6 | **No viable por ahora** | No hay registro único de actividades realizadas ni evidencia del canal por actividad; los conteos son listas incompletas. | Ninguno |
| PC | PC.ISPC | **No viable por ahora** | I1 e I4 no se pueden calcular; por la Guía 15 el índice queda bloqueado (no se redistribuyen pesos). | Ninguno (no se calcula con todas sus entradas) |
| PDM | PDM.CS | **No viable por ahora** | Meta 99,5 disponible, pero la cobertura con el método del plan no se publica 2024-2026; MinSalud y TerriData usan otro denominador (Proxy). | Ninguno |
| PDM | PDM.CM | **Viable con ajuste** | Forensis 2024 da 7,25 por 100.000 (logro 100 con meta 8), pero la línea base 10,2 no tiene fuente declarada y 2025 solo existe en Policía (otra definición). | 2024 |
| PDM | PDM.CA | **Viable con ajuste** | Solo hay EPV urbano 2024 del POT (5,75) no reconciliable con la base 3,88; daría 100 por tope. | 2024 (con salvedad de comparabilidad) |
| PDM | PDM.CP | **No viable por ahora** | No hay valor ni método oficial del porcentaje de implementación (el PDM tiene además un índice 3,23). | Ninguno |
| PDM | PDM.CD | **No viable por ahora** | Solo hay oferta de disciplinas (38 de 40), que la Matriz prohíbe convertir en cobertura. | Ninguno |
| PDM | PDM.ILE | **No viable por ahora** | LA es construible (numerador 2024-2025 + población DANE), pero LS carece de la clasificación ICFES por plantel; sin LS el ILE es N/D. | Ninguno por ahora (LA: 2024 y 2025) |
| PDM | PDM.ICPD-6 | **No viable por ahora** | CS, ILE, CD y CP son N/D; por la Guía 15 el índice queda bloqueado. | Ninguno (no se calcula con todas sus entradas) |
| PP | PP.EPP | **Viable con ajuste** | El cruce 1:1 de los 61 proyectos "ZONA n" del plan de acción con el listado EFPC permite reconstruir APP y OPP (EPP 2025 = 98,83, recalculado); no hay marcador PP oficial. | 2025; 2026 solo provisional (OPP al 30-jun); 2024 aproximado (47 de 50) |
| PP | PP.CEP | **No viable por ahora** | PE es construible (50/61/66), pero PEC no tiene actas de recibo ni informes finales; el autorreporte da 74-90 % según la regla. | Ninguno (2025 solo provisional con autorreporte) |
| PP | PP.CMA | **No viable por ahora** | No hay meta anual de personas ni personas únicas atendidas por proyecto; NA no se puede fijar. | Ninguno |
| PP | PP.IREPP | **No viable por ahora** | CEP y CMA son N/D; por la Guía 15 y la Guía 20 no se calcula ni se compensa con EPP. | Ninguno (no se calcula con todas sus entradas) |

## 7. Hallazgos, inconsistencias y preguntas abiertas

### Hallazgos más importantes
1. **Metas del PDM correctas.** Las 7 metas de la Guía (99,5; 8; 3,88; 100; 6; 19; 14) coinciden en el PDM (Acuerdo 017/2024), el DTS modificado del 22-dic-2025 y el Apéndice 9.1.2. Todas son **meta 2027**; no hay metas anuales de resultado. El proyecto de acuerdo 059/2025 modifica productos, no resultados (en los artículos revisados); su aprobación queda sin verificar.
2. **La búsqueda en fuentes primarias mejoró varias clasificaciones del PDM.**
   - PDM.CS.O: MinSalud publica afiliados al régimen subsidiado (29.794 a dic-2025), pero no la cobertura con el método del plan. La cifra de TerriData (≈12 %) es otra definición y no se compara con 99,5.
   - PDM.CM.O: Forensis 2024 da 18 muertes y 7,25 por 100.000. Para 2025 solo hay dato de la Policía.
   - PDM.CA.O: el POT da 5,75 m²/hab urbano en 2024, pero no cuadra con la línea base del plan (3,88).
   - PDM.CP.O sigue **No disponible**: solo se publican productos y no hay método oficial.
3. **PC carece de un inventario oficial de procesos participativos.** Eso vuelve Parciales a I1, I4 e I6. PC.I3 sí tiene conteo de acciones formativas en 2025 (21 culminadas de 25) y PC.I6.V1 tiene evidencia parcial de canal virtual.
4. **CS: 3 Construibles y 3 Parciales.** Las variables salen del Plan Indicativo y de los cuadros de preguntas de rendición de cuentas. PCS se cumple 1/1 en 2024 y 2025 y es el único indicador Viable.
7. **EPP 2025 es reconstruible (complemento a la revisión del profesor Francisco).**
   - Con sus 4 fuentes, EPP queda N/D.
   - Con el plan de acción IV-trim-2025 se calcula APP $20.928.690.218 y OPP $20.684.427.720, EPP ≈ 98,8 (exploratorio, Pendiente).
   - Planeación debe validar la lista de proyectos.
   - CEP y CMA siguen N/D y por eso IREPP no se calcula.
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
3. Las fases 1 a 4 están terminadas. Para cambiar una clasificación:
   - edita la lista de correcciones en `scripts/generar_disponibilidad.py` o los `trabajo/variables_<SIS>.csv`;
   - regenera con `python3 -I scripts/generar_disponibilidad.py`;
   - actualiza §4-§6 de este archivo.
4. **Siguiente paso natural:**
   - incorporar las decisiones de los profesores (universos, aceptación de proxies, metas anuales) y los documentos descargados manualmente (`interno/documentos_por_descargar.md`);
   - reclasificar y regenerar.
5. **PP:** EPP 2025 ya es reproducible (`python3 -I scripts/epp_pp_2025.py`) y está marcado "preliminar, por validar" en el Excel (C21 en `generar_disponibilidad.py`) y en el resumen. Falta: validación del DAP de la lista de proyectos, decisión sobre el 0553, script para 2024 y 2026, y fijar la regla de agregación del proxy de PEC (da entre 45 y 55 de 61).
6. **Git:** la usuaria autorizó commit y push de `outputs/disponibilidad/` y `.memory/` para este punto de control. Para commits posteriores, confirmar con ella.
