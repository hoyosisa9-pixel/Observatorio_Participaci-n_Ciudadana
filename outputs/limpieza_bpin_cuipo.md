# Limpieza de `bpin` CUIPO y validación del puente CODIGO_PROYECTO ↔ BPIN

Generado por `scripts/limpieza_bpin_cuipo.py`. Conteos y códigos de proyecto; sin datos personales.

## Clasificación de `bpin` (CUIPO Ejecución de Gastos, filas)

| Categoría | Filas |
|---|---|
| valido | 42319 |
| nulo | 19700 |
| recuperado | 191 |
| malformado | 165 |
| codigo_sectorial | 38 |

Por periodo (año) y categoría: `limpieza_bpin_cuipo_categorias.csv`; valores distintos malformados: `limpieza_bpin_cuipo_valores.csv`.

## Cobertura temporal de CUIPO

Años de `periodo` en Ejecución de Gastos: ['2023', '2024', '2025', '2026']. No hay periodos 2018-2022.

Años de formulación (prefijo) de BPIN únicos: SECOP I {'2018': 21, '2019': 6, '2020': 32, '2021': 4}; CUIPO limpio {'2020': 30, '2021': 144, '2022': 10, '2023': 38, '2024': 192}.

## SECOP I `codigo_bpin` ↔ CUIPO `bpin`

| par | unicos_A | unicos_B | interseccion_unicos | pct_unicos_A_en_B | pct_filas_A_con_match |
|---|---|---|---|---|---|
| SECOP I codigo_bpin <-> CUIPO bpin (crudo normalizado) | 63 | 450 | 5 | 7.94 | 2.68 |
| SECOP I codigo_bpin <-> CUIPO bpin (limpio) | 63 | 414 | 5 | 7.94 | 2.68 |

BPIN de SECOP I con prefijo >= 2021 (potencialmente enlazables): 4; presentes en CUIPO limpio: 4.

## Validación por nombre (umbrales: confirmado >= 0.5, dudoso >= 0.3, rechazado < 0.3)

| Fuente | Códigos con candidato | Confirmado | Dudoso | Rechazado | Sufijo con varios BPIN | Con varios BPIN confirmados |
|---|---|---|---|---|---|---|
| seguimiento2023 | 55 | 0 | 2 | 53 | 47 | 0 |
| proyectosinversion2024 | 51 | 0 | 2 | 49 | 43 | 0 |

### Ejemplos

**confirmado**

**dudoso**
- cód. 11 / BPIN 2021052660011 (puntaje 0.456): «DESARROLLAR ESTRATEGIAS PARA EL ACCESO Y PERMANENCIA DE LA EDUCACIÓN MEDIA, TÉCNICA, TECNO» vs «calidad, cobertura y fortalecimiento de la educacion inicial, prescolar, basica y media.»
- cód. 24 / BPIN 2023052660024 (puntaje 0.456): «FORTALECIMIENTO DE LAS ESTRATEGIAS DE GESTION Y ANALISIS DEL ECOSISTEMA DEL ENTORNO EDUCAT» vs «Servicio de desarrollo de contenidos educativos para la educacion inicial, preescolar, bas»
- cód. 11 / BPIN 2021052660011 (puntaje 0.456): «DESARROLLAR ESTRATEGIAS PARA EL ACCESO Y PERMANENCIA DE LA EDUCACIÓN MEDIA, TÉCNICA, TECNO» vs «calidad, cobertura y fortalecimiento de la educacion inicial, prescolar, basica y media.»
- cód. 24 / BPIN 2023052660024 (puntaje 0.456): «FORTALECIMIENTO DE LAS ESTRATEGIAS DE GESTION Y ANALISIS DEL ECOSISTEMA DEL ENTORNO EDUCAT» vs «Servicio de desarrollo de contenidos educativos para la educacion inicial, preescolar, bas»

**rechazado**
- cód. 12 / BPIN 2024052660012 (puntaje 0.29): «INCLUSIÓN Y DIVERSIDAD EN LA CALIDAD EDUCATIVA EN EL MUNICIPIO DE ENVIGADO» vs «promocion y acceso efectivo a procesos culturales y artisticos»
- cód. 13 / BPIN 2024052660013 (puntaje 0.29): «I AM EDUCATION. ECOSISTEMA EDUCATIVO CON ACCESO AL BILINGÜISMO EN EL MUNICIPIO DE ENVIGADO» vs «promocion y acceso efectivo a procesos culturales y artisticos»
- cód. 14 / BPIN 2024052660014 (puntaje 0.29): «ECOSISTEMA EDUCATIVO CON ARTICULACIÓN ENTRE SUS NIVELES Y CICLOS EN EL MUNICIPIO DE ENVIGA» vs «promocion y acceso efectivo a procesos culturales y artisticos»
- cód. 15 / BPIN 2024052660015 (puntaje 0.29): «MODELOS PEDAGÓGICOS FLEXIBLES CON ENFOQUE DIFERENCIAL, GÉNERO E INCLUSIÓN PARA UN ECOSISTE» vs «Servicio de gestion territorial para atencion en salud -pandemias- a poblacion afectada po»

