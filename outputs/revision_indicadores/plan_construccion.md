# Plan de construcción de indicadores e índices (OCE)

**Estado:** Provisional · 18-nov-2026 · Basado en las respuestas del profe a las dudas y en las decisiones de Isabella (D34–D49 en `.claude/memoria/decisiones.md`). Ya no quedan decisiones abiertas para construir.
**Flujo por dimensión:** Analizador (ficha final de cada indicador) → aprobación de Isabella → Ejecutor (script + CSV) → QA → actualizar memoria y entregas.

## Reglas transversales (aplican a todo)
1. **Brecha de información siempre.** Cada indicador reporta, por año: valor, «sin dato» y por qué. Cada índice reporta los años calculados y los años sin índice por falta de algún componente. Brecha = años (o eventos) sin dato / años (o eventos) del periodo.
2. **Trazabilidad.** Cada dato lleva documento, tabla o página, enlace, fecha de consulta y archivo descargado en `data_raw_IIC/07_revision_nov2026/<fuente>/` con `manifiesto.csv`.
3. **Mismo universo.** Solo se usan metas del mismo universo que el dato. Si no hay, se usa el referente acordado.
4. **Sin autorreportes del 100 %** como dato principal.
5. **No redistribuir pesos.** Si falta un componente, el índice de ese año queda N/D y se anota en la brecha.
6. **Cada indicador con todos sus años.** El indicador se construye con toda su serie disponible; el índice solo con los años en que coinciden todos sus componentes. Las tasas usan **una sola serie de población**: DANE vigente (VP, 8-ago-2025) = TerriData, dimensión «Descripción general», «Población total» (2018 = 228.836; 2024 = 260.430). El 238.221 de 2018 sale de otro indicador con el mismo nombre (dimensión «Censo 2005 y proyecciones DANE»): es la proyección antigua y **no se usa**. Filtrar siempre por dimensión (Analizador, 18-nov).
7. **Salidas:** `outputs/construccion/<dimensión>/` con un CSV por dimensión (`;` y BOM): indicador, año, variables, valor, puntaje 0–100, referente, fuente, enlace, estado y brecha. Scripts en `scripts/construccion/`.

## Orden
PP (el más listo) → PDM → Participación ciudadana → Control social (de último, como pediste).

## 1. Presupuesto participativo · IREPP = 0,20·EPP + 0,40·EJ90 + 0,40·FID (D44) — **CONSTRUIDO** (QA aprobado con observaciones, 18-nov)
Scripts: `scripts/construccion/pp_lectores.py` → `pp_cruce.py` → `pp_muestra_manual.py` → `pp_irepp.py`. Salidas: `outputs/construccion/PP/` (indicadores, cruce de 349 proyectos, ajustes y muestra manual). Decisiones D43, D44, D51–D56.

| Año | EPP | EJ90 | FID | IREPP |
|---|---|---|---|---|
| 2020* | 82,42 | 22/69 = 31,88 | 38/69 = 55,07 | 51,27 Naranja |
| 2021 | N/D (39/89 verificados < 50 %, D56) | | | |
| 2022 | 94,22 | 32/39 = 82,05 | 33/39 = 84,62 | 85,51 Amarillo |
| 2023 | 97,03 | 34/39 = 87,18 | 34/39 = 87,18 | 89,15 Amarillo |
| 2024 | 99,15 | 47/51 = 92,16 | 47/51 = 92,16 | 93,56 Verde |
| 2025 | 98,83 | 57/61 = 93,44 | 59/61 = 96,72 | 95,83 Verde |

Brecha: 1 de 6 años (2021). *2020: valor priorizado (en 2019 no hubo votación); 20 de 69 proyectos con contracrédito total ($3.627.742.979); obligación / valor votado = 46,34 %.

## 2. Cumplimiento del PDM · ICPD-7 = promedio simple de 7 (D34, D48) — **CONSTRUIDO** (QA aprobado con observaciones, 18-nov)
Script: `scripts/construccion/pdm_icpd7.py` (44 asserts; tarda unos 3 minutos porque lee la PPED cruda del DANE). Salida: `outputs/construccion/PDM/pdm_indicadores.csv`. Decisiones D57–D68. Población única = DANE VP (D61).

| Componente | Serie | 2019 | 2022 | 2024 | Brecha |
|---|---|---|---|---|---|
| Salud (ref. 100 %) | 2012–2025 | 82,37 | 85,03 | 83,76 | 0/14 |
| Logro educativo | 2015–2024 (2015–2018 con alerta) | 83,48 | 90,91 | 95,15 | 1/11 (2025) |
| Servicios públicos (ref. 100 %) | 2019–2024 (2019 con alerta, D65) | 91,97 | 89,56 | 89,31 | 1/7 (2025) |
| Mortalidad vial (ref. 7,33) | 2003–2025 | 78,43 | 63,69 | 95,45 | 0/23 |
| Espacio verde (avance 4,54 → 9) | 2019, 2022, 2024 | 0,00 | 2,69 | 27,13 | 4/7 |
| Veedurías (LB 5,4) | 2012–2025 | 37,04 | 74,07 | 18,52 | 19/52 consecutivos faltantes en 2019–2025 (D66) |
| Homicidio (meta 2,7) | 2003–2025 | 63,56 | 75,59 | 58,60 | 0/23 |
| **ICPD-7** | | **62,41 Naranja** | **68,79 Naranja** | **66,85 Naranja** | 4/7 años (2020, 2021, 2023 y 2025) |

Fuentes archivadas con manifiesto en `data_raw_IIC/07_revision_nov2026/` (espacio_verde, dane_pped_edad, dane_vp_retroproyeccion, men_3y4sdmxy, personeria_jju5br3w, pdm_balance_educacion_2024). El 4,54 de 2019 se verificó visualmente en la tabla 4 del VMR del AMVA.

## 3. Participación ciudadana · ISPC = 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4 (D42) — **CONSTRUIDO** (QA aprobado con observaciones, 18-nov)
Script: `scripts/construccion/pc_ispc.py` (56 asserts; tarda unos 3 minutos). Salida: `outputs/construccion/PC/pc_indicadores.csv`. Decisiones D69–D74. Convención: el año es el año calendario de la actividad participativa (I1 = votación para la vigencia t+1; I4 = año de la vigencia).

| Año | I1 votos → IPCM → puntaje (LB 3,68; ref. 7) | I2 (12 zonas) | I3 | I4 | ISPC |
|---|---|---|---|---|---|
| 2017 | 8.221 → 3,68 % → 0 (LB) | | | | |
| 2018 | no hubo votación | | | | |
| 2020 | sin dato | sin dato | sin dato | 2/3 | N/D |
| 2021 | 16.828 → 6,80 % → 93,83 | sin dato | sin dato | 3/3 | N/D |
| 2022 | 12.541 → 4,98 % → 39,08 | sin dato | sin dato | 3/3 | N/D |
| 2023 | 9.156 → 3,58 % → 0 | sin dato | sin dato | 3/3 | N/D |
| 2024 | sin dato | 12/12 | sin dato | 3/3 | N/D |
| 2025 | 10.662 → 4,03 % → 10,51 | 12/12 (alerta D73) | 21/25 = 84,00 | 2/3 (rango hasta 3/3, D74) | **61,62 Naranja** (máx. 69,95) |

Brechas: I1 3/8 años, I2 4/6, I3 5/6, I4 0/6 (2 de 18 proceso-años sin documento), ISPC 5/6. Fuentes de I4 archivadas en `data_raw_IIC/07_revision_nov2026/rendicion_cuentas/`; los votos, en `noticias_alcaldia/`.

## 4. Control social · ICRCS = (DSI + PQRSDF + VRC + RPO) / 4 — **CONSTRUIDO PARCIAL** (QA aprobado con observaciones, 19-nov; falta PQRSDF por la caída del portal)
Scripts: `scripts/construccion/cs_descargas.py` (descarga idempotente) → `scripts/construccion/cs_icrcs.py` (48 asserts; el CSV solo se escribe si pasan todos). Salida: `outputs/construccion/CS/cs_indicadores.csv`. Decisiones D75–D82. VRC, RPO y DSI entran con su mínimo.

| Año | PQRSDF | DSI (mín.–máx.) | VRC (mín.–máx.) | RPO | ICRCS |
|---|---|---|---|---|---|
| 2020 | N/D | no aplica | N/D | N/D | N/D |
| 2021 | N/D | no aplica | pendiente de descarga | 84,62–100 | N/D |
| 2022 | N/D | no aplica | pendiente | N/D | N/D |
| 2023 | solo T4 = 85,42 | no aplica | pendiente | 0 (fuera de ventana) | N/D |
| 2024 | N/D | 0–40 | 66,67–100 | 94,12 | N/D: entre 40,20 y 65,20 según PQRSDF |
| 2025 | N/D | 0–10 | 20–80 | 50 (alerta) | N/D: entre 17,50 y 42,50 según PQRSDF |

**Pendiente (portal www.envigado.gov.co caído desde el 19-nov):** 23 de los 24 informes de PQRSDF y las actas de Salud 2021–2023. Cuando responda: `python3 scripts/construccion/cs_descargas.py` y luego `python3 scripts/construccion/cs_icrcs.py`. Las actas se clasifican a mano. Hay que buscar también las rendiciones de Discapacidad y Víctimas 2021–2023 en la API.

## Cierre
- Actualizar los dos Word (cuadros con valores, puntajes y brecha), `registro_cambios.md` y la memoria.
- QA final de todo el sistema: cifras contra la fuente, universos, fórmulas y brechas.
