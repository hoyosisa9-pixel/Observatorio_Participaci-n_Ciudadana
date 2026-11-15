# Reporte del Ejecutor: series PDM de Envigado (provisional, 15-nov-2026)

Script: `scripts/revision_indicadores/pdm_series.py`. Salidas: `series_pdm.csv` (196 filas) y `series_candidatos_sexto.csv`. Descargas con manifiesto en `data_raw_IIC/07_revision_nov2026/` (terridata, men_nudc7mev, policia_homicidio, policia_hurto_personas, policia_lesiones, policia_transito).
**Población única** para todas las tasas: DANE PPED 2018-2042, total municipal (2023 = 256.042; 2024 = 260.430; 2025 = 264.637), la misma que usa calcular.py para salud. Se corrige así la 248.304 usada antes en la Policía 2024.
**TerriData sí se automatiza**: archivo masivo COMPLETO por entidad `https://terridata.dnp.gov.co/assets/docs/txt/entidades/TerriData05266.txt.zip` (separador |, ago-2026). La versión resumida «f» (que usé al principio y que también lee calcular.py) omite años; no usarla. Duplicados de mes 9 y 12: se usa diciembre. Las descargas no se sobrescriben: cada corrida con contenido distinto va a una subcarpeta con la fecha, con manifiesto.

## Salud
Afiliados SGSSS (TerriData/MinSalud, dic.) / población, 2018-2025: 82,71; 82,37; 82,29; 85,21; 85,03; 84,23; 83,76; 83,06. Variación (pp): -0,34 (2019); -0,08; +2,92; -0,18; -0,80; -0,46; -0,71 (2025). 2020-2025 idénticos a la serie existente («Cobertura de salud»); 2018-2019 son nuevos (afiliados 189.271 y 193.893). Control: contributivo+subsidiado+especiales = total en todos los años (diferencia 0). Sin meta del PDM (otro universo). Nota: calcular.py lee el archivo «f» y por eso no tiene 2018-2019.

## ILE
- Cobertura neta total TerriData (MEN) 2015-2024: 95,88; 96,05; 94,81; 96,96; 97,06; 97,95; 98,65; 100,44; 99,68; 98,45. El MEN en datos.gov.co (nudc-7mev; ahora filtrado por departamento 5/05 y municipio 5266/05266, 2011-2024 guardado) coincide en 2018-2019 y 2022-2024 pero difiere en 2020-2021 (96,85 y 98,04) y 2015-2017 TerriData (95,9; 96,1; 94,8) y MEN (68,0; 66,5; 64,2) no coinciden y no se sabe con qué población: el ILE usa desde 2018. Usar TerriData. 2025 no publicada.
- Primera infancia 0-5: TerriData hasta 2023; 2024 y 2025 del informe de gestión (la fuente cambia; el cambio queda anotado en la nota de cada fila; ya estaba en indicadores_simplificados.csv). El numerador incluye gestantes. 2023: 0-5 = 16,56 y 0-6 = 14,02 frente a LB 16,3: 0-5 reproduce mejor la LB.
- ILE = (LA+LN)/2, con LA tope 100 (meta 19) y LN tope 100 (MN 100,44): 2018 83,39; 2019 83,48; 2020 86,24; 2021 89,18; 2022 90,91; 2023 93,21; 2024 95,15; 2025 N/D (sin cobertura neta; LA 2025 = 90,08).
- Anomalía: LN 2022 alcanza 100 por construcción (la meta es la LB de 2022).

## Servicios públicos (cobertura total, archivo completo de TerriData)
Hay serie anual 2018-2024 de los cuatro servicios (REC, Superservicios, para acueducto, alcantarillado y aseo; UPME/IPSE para energía, desde 2019). Logro = min(100, 100·O/99,9); aseo sin meta.
| Año | Acueducto (logro) | Alcantarillado (logro) | Aseo | Energía (logro) | Promedio de las 4 |
|---|---|---|---|---|---|
| 2018 | 92,98 (93,07) | 93,05 (93,14) | 93,88 | N/D | N/D |
| 2019 | 88,75 (88,84) | 88,86 (88,95) | 90,27 | 100 (100) | 91,97 |
| 2020 | 68,31 (68,38) | 68,54 (68,61) | 72,02 | 100 (100) | 77,22 |
| 2021 | 85,60 (85,69) | 85,67 (85,76) | 87,67 | 100 (100) | 89,74 |
| 2022 | 84,16 (84,24) | 86,42 (86,51) | 87,67 | 100 (100) | 89,56 |
| 2023 | 87,05 (87,14) | 87,35 (87,44) | 88,38 | 96,27 (96,37) | 89,76 |
| 2024 | 86,85 (86,94) | 88,14 (88,23) | 86,64 | 95,63 (95,73) | 89,31 |
Respaldo: puntaje MDM «Acceso a servicios públicos» 2016-2024 (73,35; 73,97 [corte de octubre]; 74,30; 73,21; 66,87; 72,95; 64,89; 65,20; 62,04). Censo 2018 (referencia): acueducto 98,96, alcantarillado 97,34, aseo 99,67.
Anomalías del REC: caídas a cerca de 68 en 2017 (solo corte de septiembre) y en 2020; cobertura rural en 0 de 2016 a 2020 con urbana en 100 (no cuadra con el total); valores de 2016 y 2018 idénticos; aseo 2022 repite 2021. Energía 2023: bajan las dos zonas (urbana 100, 98,50, 97,39; rural 100, 64,43, 69,77). El REC es registro de suscriptores y la meta del PDM viene de la ECV (hogares): no son comparables (C7), y la serie es volátil: leerla con cautela. Superservicios: el informe nacional 2024 no trae tabla municipal; no se encontró un conjunto municipal en datos.gov.co.

## Sexto indicador (candidatos), tasa por 100.000 (Policía, hechos en Envigado)
| Año | Homicidio (n) tasa; prom. móvil 3 años | Hurto a personas (n) tasa | Lesiones personales (n) tasa |
|---|---|---|---|
| 2019 | 10 / 4,25 | 1.601 / 680,1 | 376 / 159,7 |
| 2020 | 7 / 2,89 | 1.080 / 445,6 | 239 / 98,6 |
| 2021 | 3 / 1,21; 2,78 | 1.560 / 629,9 | 318 / 128,4 |
| 2022 | 9 / 3,57; 2,56 | 1.554 / 616,8 | 316 / 125,4 |
| 2023 | 9 / 3,52; 2,77 | 1.643 / 641,7 | 294 / 114,8 |
| 2024 | 12 / 4,61; 3,90 | 1.329 / 510,3 | 245 / 94,1 |
| 2025 | 8 / 3,02; 3,72 | 1.327 / 501,4 | 265 / 100,1 |
LB 2023 y meta 2027 (Ap. 9.1.2): homicidio 3,7 a 2,7; hurto 632,5 a 631,5; lesiones 113,7 a 112,7. Con la población única: homicidio 3,52 (LB 3,7), hurto 641,7 (632,5), lesiones 114,8 (113,7): la diferencia se debe a otra población. Homicidio incluye feminicidio (2020: 1; 2024: 1); coincide con TerriData (4,2; 2,9; 1,2; 3,6; 3,5; 4,6; 3,0). Hurto coincide con TerriData salvo 2025 (501,4 frente a 497,7). Lesiones no está en TerriData (jr6v-i33g, MinDefensa). 2026 es parcial (a sep-2026) y no se usa.

## Mortalidad vial (Policía ha6j-pa2r, código 05266000)
Muertes 2019-2025: 22, 12, 21, 29, 43, 20, 30; tasa: 9,35; 4,95; 8,48; 11,51; 16,79; 7,68; 11,34 (TerriData, misma fuente, publica 9,3; 5,0; 8,5; 11,5; 16,8; 7,7; 11,3). Salto 2023: son hechos ocurridos en Envigado (el conjunto no distingue residencia); las 43 filas están repartidas todo el año, sin un evento masivo (máximo 3 víctimas en un día: 8-jun y 5-jul-2023; 40 de 43 con medio «no reportado»). No hay manera de explicarlo con estos datos; la LB 10,2 del PDM es de Movilidad y otro universo. Forensis no se mezcló.

## Espacio verde (búsqueda acotada): no encontrado
Revisados con búsqueda de texto el informe de gestión por dependencias 2024 (198 págs.), el de 2025 (304 págs.) y la evaluación del plan de acción 2025 (xlsx): ninguno trae m²/habitante de espacio público verde. Solo mencionan mantenimiento de jardines (562.908 m² en 2025) y acciones de adquisición de predios. La búsqueda web tampoco halló cifras de Envigado (el AMVA/Humboldt solo dice que perdió espacio verde por habitante entre 2006 y 2022). Los PDF se bajaron al área temporal, no a data_raw. Pendiente: pedirlo a Planeación o al AMVA.

## Recomendación para el sexto indicador
Tasa de homicidio como indicador principal: es la única de las tres que aún no alcanza su meta (3,52 en 2023, 4,61 en 2024 y 3,02 en 2025 frente a 2,7). Hurto a personas y lesiones personales están por debajo de la meta desde 2024, así que su logro daría 100 fijo. Como lectura de apoyo, el promedio móvil de 3 años de la tasa de homicidio (2,78; 2,56; 2,77; 3,90; 3,72 en 2021-2025), que suaviza los conteos pequeños (un homicidio mueve la tasa unos 0,4 puntos). Lesiones personales queda como respaldo.
