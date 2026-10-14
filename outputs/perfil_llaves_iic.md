# Perfil de llaves de cruce del IIC

Generado por `scripts/perfil_llaves_iic.py`. Solo contiene conteos y patrones de formato (9 = dígito, A = letra); sin datos personales.

## Normalización aplicada

- Todas: NFC, quitar espacios y caracteres invisibles, mayusculas. Placeholders -> nulo:
'', 'NO DEFINIDO', 'NO DEFINIDA', 'N/A', 'NA', 'NULL', 'NONE', 'NINGUNO', 'SIN DATO', '0'.
- Contratos: se quitan todos los espacios (p. ej. 'ENV-04-09-0549-21' = 'ENV - 04 - 09 - 0549 - 21').
Adicionalmente se calcula la 'base ENV' (regex ENV-dd-dd-dddd-dd), que ignora sufijos como '-1'.
- BPIN y CODIGO_PROYECTO: solo digitos, sin ceros a la izquierda ('2020052660141.0' -> '2020052660141').
- Filtro territorial: se conservan solo filas de Envigado (columna de municipio/ciudad cuando existe).

## Perfil por llave

| llave | fuente | filtro_territorial | filas_origen | filas_tras_filtro | nulos | pct_nulos | valores_unicos | filas_duplicadas | longitud_min_max | patrones_top3 |
|---|---|---|---|---|---|---|---|---|---|---|
| numero_de_contrato | SECOP I Procesos | municipio_entidad=Envigado | 2531 | 2531 | 1 | 0.04 | 2283 | 247 | 8-32 | AAA-99-99-9999-99 (1135); AAA-99-99-999-99 (435); 99-99-99-99-999-99 (346) |
| codigo_bpin | SECOP I Procesos | municipio_entidad=Envigado | 2531 | 2531 | 2307 | 91.15 | 63 | 161 | 13-13 | 9999999999999 (224) |
| referencia_del_contrato | SECOP II Contratos | ciudad=Envigado | 25854 | 25854 | 0 | 0.0 | 25549 | 305 | 1-167 | AAA-99-99-9999-99 (9025); AAAAAA-999-99-999-99 (3582); 99-9-99999-99 (2722) |
| id_contrato | SECOP II Contratos | ciudad=Envigado | 25854 | 25854 | 0 | 0.0 | 25854 | 0 | 17-18 | AA9.AAAAAA.9999999 (23322); AA9.AAAAAA.999999 (2532) |
| numero_del_contrato | SECOP Integrado | municipio_entidad=Envigado | 58464 | 57536 | 0 | 0.0 | 47819 | 9717 | 9-18 | AA9.AAAAAA.9999999 (30848); 99-99-9999999 (13895); 99-9-9999999 (3034) |
| bpin | CUIPO Ejecucion Gastos | nombre_entidad=Envigado | 62413 | 62413 | 19700 | 31.56 | 450 | 42263 | 6-15 | 9999999999999 (42346); 999999999999999 (165); 99999999999999 (148) |
| numero contrato PPP 2022 (regex ENV-dd-dd-dddd-dd) | PPP 2022 (PDF) | ninguno (documento municipal de Envigado) | 52 | 52 | 0 | 0.0 | 14 | 38 | 17-17 | AAA-99-99-9999-99 (52) |
| CODIGO_PROYECTO seguimiento 2023 (T3+T4) | seguimientoproyectosdeinversion*2023.xlsx | ninguno (documento municipal de Envigado) | 419 | 419 | 0 | 0.0 | 217 | 202 | 2-3 | 999 (360); 99 (59) |
| CODIGO_PROYECTO proyectosdeinversion2024 | proyectosdeinversion2024.xlsx | ninguno (documento municipal de Envigado) | 207 | 207 | 0 | 0.0 | 207 | 0 | 2-3 | 999 (181); 99 (26) |

## Coincidencia entre llaves

| par | unicos_A | unicos_B | interseccion_unicos | pct_unicos_A_en_B | pct_unicos_B_en_A | pct_filas_A_con_match |
|---|---|---|---|---|---|---|
| PPP 2022 <-> SECOP I numero_de_contrato (exacto) | 14 | 2283 | 1 | 7.14 | 0.04 | 23.08 |
| PPP 2022 <-> SECOP II referencia_del_contrato (exacto) | 14 | 25549 | 11 | 78.57 | 0.04 | 71.15 |
| PPP 2022 <-> SECOP II referencia_del_contrato (base ENV, sin sufijo) | 14 | 9105 | 12 | 85.71 | 0.13 | 73.08 |
| PPP 2022 <-> SECOP II id_contrato (exacto) | 14 | 25854 | 0 | 0.0 | 0.0 | 0.0 |
| PPP 2022 <-> SECOP Integrado numero_del_contrato (exacto) | 14 | 47819 | 0 | 0.0 | 0.0 | 0.0 |
| PPP 2022 <-> SECOP Integrado numero_del_contrato (base ENV) | 14 | 0 | 0 | 0.0 | None | 0.0 |
| SECOP I codigo_bpin <-> CUIPO bpin (exacto) | 63 | 450 | 5 | 7.94 | 1.11 | 2.68 |
| SECOP I numero_de_contrato <-> SECOP II referencia_del_contrato (exacto) | 2283 | 25549 | 3 | 0.13 | 0.01 | 0.12 |
| SECOP I numero_de_contrato <-> SECOP Integrado numero_del_contrato (exacto) | 2283 | 47819 | 0 | 0.0 | 0.0 | 0.0 |
| SECOP II referencia_del_contrato <-> SECOP Integrado numero_del_contrato (exacto) | 25549 | 47819 | 819 | 3.21 | 1.71 | 3.17 |
| SECOP II id_contrato <-> SECOP Integrado numero_del_contrato (exacto) | 25854 | 47819 | 23478 | 90.81 | 49.1 | 90.81 |
| CODIGO_PROYECTO seguimiento 2023 <-> BPIN CUIPO (exacto) | 217 | 450 | 0 | 0.0 | 0.0 | 0.0 |
| CODIGO_PROYECTO seguimiento 2023 <-> sufijo local de BPIN CUIPO (exploratorio) | 217 | 225 | 55 | 25.35 | 24.44 | 26.01 |
| CODIGO_PROYECTO 2024 <-> sufijo local de BPIN CUIPO (exploratorio) | 207 | 225 | 51 | 24.64 | 22.67 | 24.64 |
| SECOP I codigo_bpin <-> sufijo local de BPIN CUIPO (control) | 59 | 225 | 46 | 77.97 | 20.44 | 62.5 |
| CODIGO_PROYECTO seguimiento 2023 <-> CODIGO_PROYECTO 2024 | 217 | 207 | 194 | 89.4 | 93.72 | 88.78 |

## Avisos

- BPIN CUIPO con estructura AAAA05266NNNN (13 digitos): 42346 de 42713 filas no nulas.
