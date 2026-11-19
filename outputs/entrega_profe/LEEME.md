# Entrega al profe: indicadores, cambios y resultados

Estado al 19-nov-2026 (todo Provisional). Los resultados vigentes están en los dos Word, en `outputs/construccion/` y en la base para Power BI.

| Archivo | Contenido | Estado |
|---|---|---|
| `indicadores_cambios_y_datos_OCE.docx` | Word para el profe: cambios, resultados de los índices, decisiones y variables por solicitar | **Vigente** (lo genera `scripts/entregables/generar_entrega_profe.py`) |
| `detalle_revision_indicadores_OCE.docx` | Word para Isabella: series completas, fuentes, scripts, decisiones con número y origen, pendientes | **Vigente** (mismo generador) |
| `detalle_indicadores_OCE_profe.docx` | Word detallado para el profe: por cada indicador de la matriz original, qué pasó, por qué, cómo quedó, serie por año, brecha y alertas (sin origen de las decisiones) | **Vigente** (`scripts/entregables/generar_detalle_profe.py`) |
| `ficha_metodologica_indicadores_OCE.docx` | Ficha metodológica: una por indicador principal (18) y por índice (4), más anexo de definiciones | **Vigente** (`scripts/entregables/generar_ficha_metodologica.py`) |
| `registro_cambios.md` | Lista corta de todos los cambios y sus orígenes | **Vigente** |
| `dudas_para_el_profe.md` | Las 16 dudas, todas respondidas, con la decisión posterior de Isabella aparte | **Vigente** (cerrada) |
| `1_PDM.md`, `2_PC.md`, `3_PP.md`, `4_CS.md` | Cuadros de la revisión del 15-nov | **Histórico** (valores y decisiones anteriores) |
| `../entrega_powerbi/base_indicadores_OCE.xlsx` | Excel completo: sirve para el profe (hojas Inicio, Índices, Indicadores y Ficha metodológica) y para Power BI (hoja Datos) (`scripts/entregables/generar_base_powerbi.py`) | **Vigente** |

Soporte técnico: `outputs/construccion/` (CSV por dimensión), `scripts/construccion/` (scripts) y `outputs/revision_indicadores/` (series y QA).
Para regenerar todo después de cambiar los datos (los informes de PQRSDF y las actas ya están en `data_raw_IIC/07_revision_nov2026/`, obtenidos del Google Drive de Isabella porque el portal dio HTTP 500; si el portal vuelve, `cs_descargas.py` solo baja lo que falte): `python3 scripts/construccion/cs_descargas.py && python3 scripts/construccion/cs_icrcs.py && python3 scripts/entregables/generar_entrega_profe.py && python3 scripts/entregables/generar_base_powerbi.py`
