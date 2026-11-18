# Plan de construcción de indicadores e índices (OCE)

**Estado:** Provisional · 18-nov-2026 · Basado en las respuestas del profe a las dudas y en las decisiones de Isabella (D34–D49 en `.claude/memoria/decisiones.md`). Ya no quedan decisiones abiertas para construir.
**Flujo por dimensión:** Analizador (ficha final de cada indicador) → aprobación de Isabella → Ejecutor (script + CSV) → QA → actualizar memoria y entregas.

## Reglas transversales (aplican a todo)
1. **Brecha de información siempre.** Cada indicador reporta, por año: valor, «sin dato» y por qué. Cada índice reporta los años calculados y los años sin índice por falta de algún componente. Brecha = años (o eventos) sin dato / años (o eventos) del periodo.
2. **Trazabilidad.** Cada dato lleva documento, tabla o página, enlace, fecha de consulta y archivo descargado en `data_raw_IIC/07_revision_nov2026/<fuente>/` con `manifiesto.csv`.
3. **Mismo universo.** Solo se usan metas del mismo universo que el dato. Si no hay, se usa el referente acordado.
4. **Sin autorreportes del 100 %** como dato principal.
5. **No redistribuir pesos.** Si falta un componente, el índice de ese año queda N/D y se anota en la brecha.
6. **Cada indicador con todos sus años.** El indicador se construye con toda su serie disponible; el índice solo con los años en que coinciden todos sus componentes. Las tasas usan **una sola serie de población** en todo el periodo (la DANE de TerriData 2003–2026 o la PPED 2018–2042; no mezclar: en 2018 difieren, 238.221 frente a 228.836).
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

## 2. Cumplimiento del PDM · ICPD-7 (D34)
| Componente | Puntaje | Referente | Años | Brecha |
|---|---|---|---|---|
| Salud | cobertura total (afiliados / población DANE) | 100 % universal (D35) | **2012–2025** (afiliados en TerriData desde 2012) | — |
| Logro educativo (ILE) | (min(100, 100·A/19) + min(100, 100·N/100,44)) / 2 | metas del PDM (D15) | **2015–2024** (cobertura neta desde 2005; primera infancia desde 2015) | 2025 sin cobertura neta; verificar la población 0–5 antes de 2018 |
| Servicios públicos | promedio simple de 4 coberturas (D39) | **100 %** (D38) | 2019–2024 con las 4; acueducto, alcantarillado y aseo **2008–2024** | energía total solo desde 2019; 2025 |
| Mortalidad vial | min(100, 100·7,33/O) (fórmula de la matriz) | **7,33** = reducción del PDM (−21,6 %) sobre la LB Policía 2019 (D47) | **2003–2025** (TerriData, con datos de la Policía) | 2023 atípico |
| Espacio verde | (valor − 4,54) / (9 − 4,54) | LB 2019, referente 9 (D31) | 2019, 2022, 2024 | 2020, 2021, 2023, 2025 |
| Veedurías | min(100, 100·valor/5,4) | LB = promedio 2019–2023 = **5,4** (D46) | 2012–2026 | listado incompleto (consecutivos 118–293) |
| **Homicidio (nuevo)** | 100·2,7 / O, tope 100 | meta PDM 2,7 (Policía, mismo universo) | **2003–2025** (Policía y TerriData) | — |
| **ICPD-7** | **promedio simple de los 7** (D48) | | **2019, 2022, 2024** (por espacio verde) | 2020, 2021, 2023, 2025 |

## 3. Participación ciudadana · ISPC = 0,30·I1 + 0,25·I2 + 0,20·I3 + 0,25·I4 (D42)
| Indicador | Puntaje | Años | Brecha |
|---|---|---|---|
| I1 IPCM | (IPCM − 3,62) / (7 − 3,62) (D32) | 2017, 2025 | 2018 y 2020–2024 por solicitar; 2019 sin votación |
| I2 cobertura territorial | zonas con actividad / **12** (D41) | 2024, 2025; respaldo 2020–2025 | 2020–2023 solo respaldo |
| I3 oferta formativa | realizadas / programadas | 2025 (denominador reconstruido) | sin programación publicada en ningún año; respaldo: conteo 17, 7, 20, solo ciudadanía (D49) |
| I4 publicación de resultados | procesos con resultados / 3 procesos fijos (D29) | 2021–2025 | 2022 por confirmar |
| **ISPC** | | **2025** | demás años N/D (I1 e I3) |

## 4. Control social · ICRCS = (DSI + PQRSDF + VRC + RPO) / 4 (D45)
| Indicador | Puntaje | Años | Brecha |
|---|---|---|---|
| PQRSDF (antes PCS) | atendidas a tiempo / recibidas, suma de trimestres | 2020–2025 | trimestres faltantes |
| DSI | metas con información de seguimiento / metas de la lista **actualizada** | 2024, 2025 | metas sin dato publicado |
| VRC | mínimo (al índice), máximo y brecha (D33) | 2024, 2025 | 33,3 % y 60 % |
| RPO | con respuesta de fondo / observaciones | 2024; 2025 incompleto | 2025 |
| **ICRCS** | | **2024** (2025 si se completa RPO) | |

## Cierre
- Actualizar los dos Word (cuadros con valores, puntajes y brecha), `registro_cambios.md` y la memoria.
- QA final de todo el sistema: cifras contra la fuente, universos, fórmulas y brechas.
