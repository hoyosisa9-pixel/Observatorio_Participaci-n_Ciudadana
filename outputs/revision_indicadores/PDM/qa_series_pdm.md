# QA de las series PDM del Ejecutor (15-nov-2026)

Revisé `scripts/revision_indicadores/pdm_series.py`, `series_pdm.csv`, `series_candidatos_sexto.csv`, `reporte_ejecutor.md` y las descargas de `data_raw_IIC/07_revision_nov2026/` (con sus manifiestos). Los contrasté con D15–D25 (`.claude/memoria/decisiones.md`), `plan_revision.md` y la fila PDM.ILE de `outputs/disponibilidad/trabajo/matriz.csv`. No corregí nada. Lo único que bajé fue a la carpeta temporal de QA, no a `data_raw`.

## 1. Reproducibilidad

- Corrí el script **en una copia** (scratchpad, con la misma estructura de carpetas y los dos insumos locales), sin `-I`. Así evité sobrescribir `data_raw` y `outputs`, porque el script siempre vuelve a descargar. Necesita red y funcionó: tardó unos 10 s e imprimió «196 filas».
- `series_pdm.csv` y `series_candidatos_sexto.csv` salieron **idénticos byte a byte**. Los 6 archivos descargados también son idénticos a los del repositorio (mismo contenido; el manifiesto solo cambia la fecha).
- Seguridad: el ZIP se lee en memoria y por el nombre exacto del miembro, sin extraerlo. Los CSV se leen con `csv`. No hay `eval`, `exec` ni `pickle`. Los parámetros SoQL van codificados y la lectura tiene un tope de 250 MB. La carpeta del script no contiene archivos descargados. **OK.**
- Datos personales: ninguna salida trae nombres, correos ni documentos. El conjunto de tránsito trae fecha, género, grupo de edad y medio por víctima, pero las salidas solo publican conteos anuales. El conjunto de veedurías (jju5-br3w, que tiene nombres y correos) **no forma parte** de este trabajo: no aparece ni en el script ni en las salidas. **OK.**

## 2. Muestra recalculada a mano

| Cifra | Ejecutor | QA (cómo) | ¿Coincide? |
|---|---|---|---|
| Homicidio 2024 | 12 → 4,61 | API m8fd-ahd9, `cod_muni='05266'`: 11 intencionales + 1 feminicidio = 12; 12/260.430 = 4,608 | Sí |
| Homicidio 2025 | 8 → 3,02 | 8 (sin feminicidio); 8/264.637 = 3,023 | Sí |
| Lesiones 2023 | 294 → 114,82 | API jr6v-i33g: 204 filas, suma de `cantidad` = 294 (F 114 + M 179 + 1 sin sexo); 294/256.042 = 114,82 | Sí |
| Lesiones 2025 | 265 → 100,14 | 198 filas, suma = 265; 265/264.637 = 100,14 | Sí |
| Hurto a personas 2025 | 1.327 → 501,44 | 349 filas, suma = 1.327; 1.327/264.637 = 501,44 | Sí |
| Muertes en tránsito 2023 | 43 → 16,79 | Archivo descargado: 43 filas con `cantidad = 1`, todas «Art. 109 homicidio culposo en accidente de tránsito»; 43/256.042 = 16,79 | Sí |
| Cobertura neta 2022 / 2024 | 100,44 / 98,45 | Archivo TerriData y API nudc-7mev: las dos fuentes dan lo mismo | Sí |
| Primera infancia 0–5, 2023 | 16,56 | 2.059 (TerriData, niños + gestantes) / 12.431 (DANE PPED, suma de 0 a 5 años) = 16,563. Con 0–6: 2.059/14.685 = 14,02 | Sí |
| ILE 2024 | 95,15 | A = 2.161/12.324 = 17,535 → LA = min(100; 100·17,535/19) = 92,289. LN = min(100; 100·98,45/100,44) = 98,019. ILE = 95,154 | Sí |
| Salud 2025 | 83,06 | 219.795/264.637 = 83,055 (TerriData, mes 12) | Sí |
| Acueducto REC 2021 / energía 2024 | 85,60 / 95,63 | Archivo TerriData | Sí |

**Lesiones (jr6v-i33g).** Los metadatos del conjunto dicen «LESIONES PERSONALES… a excepción de aquellas ocurridas en accidentes de tránsito. Se mide en víctimas». No tiene una columna de tipo de delito porque todo el conjunto es lesiones personales. Con el filtro solo aparece `05266 / ENVIGADO / ANTIOQUIA`, así que no se cuela otro municipio. Las filas vienen agregadas (por ejemplo, 204 filas suman 294 víctimas en 2023, con hasta 5 víctimas en una fila). Sumar `cantidad` es lo correcto: contar filas daría menos. No hay filas repetidas por fecha y sexo en 2023, así que **no hay doble conteo**. Lo mismo vale para el hurto (no hay filas repetidas por fecha en 2025).

## 3. Hallazgos

| Severidad | Indicador | Hallazgo | Evidencia | Sugerencia |
|---|---|---|---|---|
| **Bloqueante** | PDM.SP (servicios públicos) | **TerriData sí tiene series anuales de acueducto, alcantarillado y aseo (REC, total, urbana y rural) de 2008 a 2024.** El Ejecutor usó la versión resumida de la ficha (`TerriData05266f.txt.zip`, 58 KB, 6.416 filas), que solo trae REC 2018 y 2021. En la misma ruta está el archivo completo `TerriData05266.txt.zip` (112 KB, 13.238 filas, también actualizado el 07-ago-2026, sha256 `e137c93e…bcc1`). Es un superconjunto del resumido: los 6.416 valores del resumido coinciden en el completo. Así, la afirmación «no hay serie anual completa» y la de que el promedio de las 4 coberturas «solo» se puede calcular para 2021 son **falsas** | `curl -I https://terridata.dnp.gov.co/assets/docs/txt/entidades/TerriData05266.txt.zip` → 200, 112.050 bytes. Valores en la tabla 4 | Cambiar a la URL sin «f» y rehacer PDM.SP 2019–2024. Revisar la calidad de la serie (ver la fila siguiente) antes de proponerla al profe (D17) |
| Importante | PDM.SP | La serie REC completa es **volátil y tiene inconsistencias internas**. Acueducto: 2017 = 67,78 y 2020 = 68,31, entre años de 85 a 93. La zona rural vale 0,00 de 2016 a 2020 mientras la urbana vale 100, lo que no cuadra con un total de 68,31 en 2020. Los valores de 2016 y 2018 son idénticos (92,98 / 93,05 / 93,88) y están duplicados, con mes 9 y mes 12. El aseo de 2022 repite el de 2021 (87,67). El salto de 2018 a 2021 que el Ejecutor llama «probable cambio de método» es parte de esta volatilidad: no son dos puntos aislados | Archivo completo, Vivienda y acceso a servicios públicos | Documentar el REC como autorreporte municipal de suscriptores y viviendas (de estratificación) y anotar la C7 frente a la ECV. Al leer el archivo completo, el script debe tener una regla de desempate, porque hoy `d[ind][año]` se queda con la última fila que lee |
| Importante | PDM.SALUD | El reporte dice «Falta 2019 (TerriData no publica el total SGSSS) y los especiales 2019-2020». El archivo completo sí los trae. Afiliados SGSSS: 2018 = 189.271 y 2019 = 193.893 (mes 12). Especiales: 2019 = 3.057 y 2020 = 3.139. El control da 0 (por ejemplo, 2019: 176.793 + 14.043 + 3.057 = 193.893). Con la población única quedan 2018 = 82,71 y 2019 = 82,37 | Archivo completo, Salud/Aseguramiento | Extender la serie a 2018–2025. El mismo error viene de `calcular.py:516-524`, que también lee el archivo «f» |
| Importante | ILE (comparativo MEN) | El filtro del MEN (`c_digo_departamento='05'`) solo recupera **2023 y 2024**. Hasta 2022, el conjunto usa el código de departamento `5` y el de municipio `5266`, como ya advertía el plan, §3 paso 5. El CSV trae 2 filas del comparativo. Lo que dice el reporte de 2020–2021 (96,85 y 98,04) y de 2017 (64,24) es cierto según la API, pero **no sale del archivo guardado ni del script**: no es trazable. La nota de 2017 del código nunca se imprime | `data_raw_IIC/07_revision_nov2026/men_nudc7mev/envigado_nudc-7mev.csv` (2 filas). API con `c_digo_municipio in('5266','05266')`: 2011–2024 | Filtrar por los dos códigos y volver a guardar |
| Importante | Todas (trazabilidad) | Cada vez que se corre, el script vuelve a descargar y **sobrescribe** `data_raw` y el manifiesto: no tiene modo sin conexión ni guarda copias por fecha. Las cifras de la Policía se actualizan cada mes (el hurto 2025 ya difiere de TerriData: 1.327 frente a unas 1.317 implícitas en 497,7), así que una nueva corrida cambia el CSV sin dejar rastro del corte anterior | `pdm_series.py:40-56` | Agregar una opción para leer lo ya descargado o guardar las descargas con la fecha |
| Menor | ILE | El cambio de fuente de primera infancia en 2024 (de TerriData al informe de gestión) aparece en la columna `fuente`, pero no en la `nota`, y no se verificó que el informe de gestión cuente igual (niños + gestantes). No se ve un salto: la variación es de +1,02 pp en 2023 y de +0,97 pp en 2024 | `series_pdm.csv`, filas de primera infancia 2024 y 2025 | Anotar el cambio en la nota y confirmar la definición con Educación |
| Menor | ILE | La fila de 2015 dice «Serie continua 2015-2024 (TerriData)», pero para 2015–2017 TerriData da 95,9 / 96,1 / 94,8 y el MEN da 68,0 / 66,5 / 64,2: no se sabe con qué población se calcularon esos años. No afecta al ILE, que empieza en 2018 | API nudc-7mev | Publicar solo desde 2018 o anotar la duda |
| Menor | PDM.SP (energía) | La nota dice que la caída de 2023 se debe al área rural (64,43), pero la cobertura urbana también baja (100 → 98,50 → 97,39). Con un peso rural implícito de cerca del 6,5 %, la caída urbana aporta unos 1,4 pp y la rural unos 2,3 pp | Archivo TerriData (urbana, rural y total) | Corregir la nota. Sí es cobertura total municipal (es coherente con el promedio ponderado urbano-rural) |
| Menor | Sexto indicador | D19 pide una **recomendación** entre homicidio, hurto y lesiones, y el reporte da los datos pero no recomienda | `reporte_ejecutor.md`, §Sexto indicador | Que la recomiende quien corresponda (el plan sugería homicidio con promedio móvil) |
| Menor | Código | Hay código muerto: `num()` no se usa, la nota de los regímenes especiales de 2019 nunca se imprime y hay un `if False` en el conteo de la Policía. La desagregación de la energía rural cita «UPME» y la total cita «UPME/IPSE»: está bien, pero conviene anotarlo | `pdm_series.py:62, 152, 239` | Limpieza |

## 4. Servicios públicos en el archivo completo de TerriData (cobertura total)

| Año | Acueducto REC | Alcantarillado REC | Aseo REC | Energía UPME/IPSE | Promedio simple de las 4 |
|---|---|---|---|---|---|
| 2019 | 88,75 | 88,86 | 90,27 | 100,00 | 91,97 |
| 2020 | 68,31 | 68,54 | 72,02 | 100,00 | 77,22 |
| 2021 | 85,60 | 85,67 | 87,67 | 100,00 | 89,73 |
| 2022 | 84,16 | 86,42 | 87,67 | 100,00 | 89,56 |
| 2023 | 87,05 | 87,35 | 88,38 | 96,27 | 89,76 |
| 2024 | 86,85 | 88,14 | 86,64 | 95,63 | 89,31 |

Más series del archivo completo:

- REC desde 2008: hay datos en 2009 y de 2011 a 2024.
- Urbana REC 2021–2024: acueducto 87,86 / 86,68 / 90,21 / 90,71; alcantarillado 88,14 / 89,30 / 90,88 / 92,35; aseo 88,64 / 89,16 / 89,96 / 88,92.
- Censo 2005 y 2018 (incluye energía, gas y aseo).
- «Acceso a servicios públicos» de la MDM, 2016–2024 (73,35 … 62,04), que en el archivo resumido solo aparece para 2021.

No hay series del SUI con otro nombre ni del ICEE aparte del indicador UPME/IPSE. En datos.gov.co no encontré un conjunto municipal de coberturas para Envigado (la búsqueda solo devuelve conjuntos de otras ciudades). Eso respalda lo que dice el Ejecutor, aunque no lo pude comprobar de forma exhaustiva.

## 5. Coherencia (puntos pedidos)

- **Población única:** todas las tasas usan `poblacion_envigado_edad_pped.csv` (DANE PPED 2018–2042 VP): las de salud, la Policía, el tránsito y la primera infancia por edad. **OK.**
- **ILE con 0–5 (D16):** se aplica, la fila de 2023 dice que el PDM habla de 0–6 y se publica el comparativo 0–6. La fórmula sigue la matriz con LS reemplazado por LN (D15): MA = 19, MN = 100,44, tope de 100 y N/D en 2025 porque falta un componente (regla de la matriz). **OK.**
- **Homicidio con feminicidio:** confirmado. Hay feminicidios en 2017 (1), 2020 (1) y 2024 (1), y las tasas de TerriData solo cuadran si se suman: en 2020, 7 da 2,89 y 6 daría 2,48, frente a 2,9 de TerriData. **OK.**
- **Tránsito 2023:** confirmado. Son 43 víctimas repartidas en el año (junio tiene 8; hay como máximo 3 en un día: 08-jun y 05-jul), 40 tienen el medio «no reportado» y TerriData da 16,8. El conjunto ubica los hechos por el lugar donde ocurrieron (`codigo_dane` del municipio del hecho). La explicación del Ejecutor se sostiene. No hay manera de descartar duplicados porque el conjunto no trae identificador: hay 7 fechas de 2023 con filas idénticas.
- **No se mezclan fuentes dentro de una serie:** se cumple, salvo el cambio de primera infancia en 2024, que está anotado en la fuente. El tránsito usa solo la Policía y la cobertura neta solo TerriData.

## Veredicto

**Rechazado**, solo por la parte de servicios públicos (y la de salud 2018–2019). Las cifras calculadas son correctas y reproducibles, y no hay problemas de seguridad ni de datos personales. Pero la conclusión central de PDM.SP («TerriData no tiene una serie anual») es falsa: el Ejecutor leyó el archivo resumido en lugar del completo. Hay que rehacer servicios públicos y salud con `TerriData05266.txt.zip` y arreglar el filtro del MEN antes de que esto llegue al Word del profe (D25). Lo demás (ILE, sexto indicador, tránsito) queda **aprobado con observaciones**.
