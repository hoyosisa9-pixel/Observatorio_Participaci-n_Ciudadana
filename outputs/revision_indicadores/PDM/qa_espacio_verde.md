# QA: espacio público verde urbano por habitante (PDM.CA), Envigado

**Estado:** Provisional · QA · 15-nov-2026
**Objeto:** `outputs/revision_indicadores/PDM/espacio_verde_serie.md` y cortes elegidos por la usuaria: 2019 = 4,54 · 2022 = 4,66 · 2024 = 5,75 (verde urbano); contexto: 2024 total = 69,41 y 2013 = 4,16.
**Método:** QA descargó de nuevo los documentos (consulta del 15-nov-2026), extrajo el texto con pypdf y leyó las tablas que vienen como imagen. Los cocientes se recalcularon a mano.

## 1. Cifras verificadas contra la fuente

| Cifra | Dónde aparece (verificado) | Concepto en la fuente | Recálculo | Resultado |
|---|---|---|---|---|
| 5,75 (2024) | Diagnóstico POT, dimensión funcional, espacio público ([PDF](https://alcaldia-municipal-de-envigado-en-antioquia.micolombiadigital.gov.co/sites/alcaldia-municipal-de-envigado-en-antioquia/content/files/002925/146211_8-dimension-funcional-_-espacio-publico.pdf)), tabla 39, p. 86 (encabezado) y p. 87 (fila) | «Indicador EPV Urbano por habitante» | 1.382.762 / 240.655 = 5,7458 → **5,75** | Confirmada |
| 1.382.762 m² | Mismo PDF, tabla 31, p. 81: «Espacio Público Verde en suelo urbano» | EPV en suelo urbano (excluye 31.056 m² de expansión) | n. a. | Confirmada |
| 240.655 hab. | Mismo PDF, tabla 32, p. 84: población urbana de 2024, «a partir de proyecciones DANE 2018» | Urbana | n. a. | Confirmada |
| 69,41 (2024, contexto) | Mismo PDF, tabla 38, p. 86 | «Indicador EPV por habitante» (total) | 17.233.624 / 248.304 = 69,4053 → **69,41** | Confirmada |
| 4,66 (2022) | Mismo PDF, tabla 39, p. 86-87, columna «Línea base AMVA, 2022»; y gráfico del Humboldt (columna 2022, fila Envigado = 4.66, 2006 = «NA») | EPV urbano per cápita (actualización PMEPVU) | No se puede recalcular (no hay área ni población) | Confirmada |
| 4,22 (2022) | Mismo PDF, p. 11: «El proceso de actualización del PMEPVU aporta un indicador para el municipio de Envigado de 4,22 m2/habitante de espacio público verde al año 2022» | Igual que el 4,66 | n. a. | Confirmada: hay una contradicción interna en el POT |
| 4,54 (2019) | [Reporte ODS 11.7 AMVA/ONU-Hábitat](https://urbanpolicyplatform.org/wp-content/uploads/2023/06/VMR-EPV-VALLE-DE-ABURRA.pdf), tabla 4, p. 24 del PDF (folio impreso «23»), fila ENVIGADO: población urbana 232.903; áreas verdes urbanas públicas 92,06 ha; **4,54** | Columna «Espacios Públicos Verdes Efectivos (m²) por habitante» | 920.600 / 232.903 = **3,95 ≠ 4,54** | El valor publicado está confirmado y es de Envigado; no se puede reproducir (ver H2) |
| 4,54 (2019) | Mismo reporte, tabla 5, p. 25 | Columna «Índice de Espacio Público Verde Urbano (m²/hab.)», con el año rotulado «2009» (errata por 2019) | n. a. | Confirmada |
| 4,16 (2013, contexto) | [ICAU 2013 Minambiente](https://www.minambiente.gov.co/wp-content/uploads/2021/06/Informe_Nacional_de_Calidad_Ambiental_Urbana_poblacion_100.000_y_500.000_habitantes._Ano_2013.pdf), tabla 16 «Resultados ICAU – Envigado», p. 41 | «Superficie de área verde por habitante» | n. a. | Confirmada |

La ficha cita bien las páginas y las tablas (`espacio_verde_serie.md`, líneas 21, 25, 26, 29 y 43).

## 2. Hallazgos

| Severidad | Indicador | Hallazgo | Evidencia | Sugerencia |
|---|---|---|---|---|
| Importante | PDM.CA (2019) | **H1. Nombre del concepto.** La tabla 4 llama al 4,54 «Espacio Público Verde **Efectivo**» y la tabla 5 lo llama «Índice de Espacio Público Verde Urbano». El texto aclara que es el «**llamado** espacio público verde efectivo» y que mide áreas verdes urbanas **públicas** sobre la población urbana. No es el EPE del Decreto 1077 (el EPE de Envigado es 2,67 a 3,43). Por eso el concepto equivale al verde urbano público, pero la etiqueta puede confundir. | VMR, pp. 23-25 (títulos de las tablas 4 y 5) | En la ficha, anotar que la fuente lo rotula «efectivo», pero que corresponde al índice de EPV urbano público y no al EPE. |
| Importante | PDM.CA (2019) | **H2. El 4,54 no se puede reproducir.** 92,06 ha / 232.903 hab. = 3,95. La diferencia aparece en **todos** los municipios de la tabla 4 (Medellín: 1.123,89 ha / 2.508.452 = 4,48 frente a 4,86; Itagüí: 4,04 frente a 4,67; Sabaneta: 6,71 frente a 5,09). No es una errata de Envigado: el índice viene directamente del estudio AMVA-UNAL (2020) con un método que el reporte no explica. | VMR, tabla 4, p. 24 | Mantener el 4,54 como valor publicado, con la nota de que no cuadra con el área y la población de la misma tabla. La ficha ya lo advierte (línea 21). |
| Importante | PDM.CA (serie 2019-2022-2024) | **H3. Los tres cortes no usan el mismo método.** (a) Productores e inventarios distintos: AMVA-UNAL 2020, con cartografía de conectividad ecológica; actualización del PMEPVU de AMVA-Humboldt 2022; inventario propio del POT 2024. (b) Poblaciones distintas: «Aburrá en Datos» (232.903) en 2019; proyección DANE 2018 de Planeación municipal (240.655) en 2024; no publicada en 2022. (c) Delimitación: el estudio de 2019 incluyó verdes en **suelo de expansión urbana** (VMR, p. 25, explicación de Itagüí), mientras que el POT 2024 solo cuenta **suelo urbano** (tabla 31; sumar los 31.056 m² de expansión daría +0,13 m²/hab). (d) La actualización de 2022 re-estimó los valores municipales de 2006 con cifras distintas a las del VMR: Medellín 6,17 frente a 6,52; Bello 4,10 frente a 4,02; Itagüí 4,20 frente a 3,86. Eso indica un método distinto al del reporte de 2019. (e) El propio VMR advierte que, «a nivel municipal es difícil hacer una comparación entre los indicadores de los dos estudios». | VMR, pp. 23 y 25; gráfico del Humboldt; POT, tablas 31, 32 y 39 | Presentar 4,54 → 4,66 → 5,75 como **cortes de fuentes distintas**, no como tendencia. No atribuir la diferencia de 1,09 (tabla 39) a una gestión municipal. La línea 82 de la ficha («del mismo concepto [...] el método sigue el criterio del AMVA») es exacta en el concepto, pero exagera la equivalencia del método. |
| Importante | PDM.CA (2022) | **H4. La cifra de 2022 se cita de dos formas.** 4,22 en la p. 11 y 4,66 en la tabla 39 del mismo diagnóstico. El 4,66 coincide con el gráfico original del Humboldt; el 4,22 no tiene respaldo externo. | POT, p. 11 y tabla 39 | Usar 4,66 (respaldado por el Humboldt) y dejar constancia del 4,22. La ficha ya lo hace (línea 26). |
| Menor | Contexto 2024 | **H5. El total de EPV no es uniforme en el POT.** Tablas 30 y 38: 17.233.624 m². Tabla 31: 17.233.674 (que es la suma correcta de sus componentes). Tabla 29: 17.405.793. Con el total de la tabla 31 el indicador sigue en 69,41; con el de la tabla 29 sería 70,10. | POT, pp. 80-81 y 86 | Citar la tabla 38 como está. Opcionalmente, anotar la diferencia. |
| Menor | PDM.CA (2019) | **H6. Paginación y erratas.** La tabla 4 está en la p. 24 del PDF, pero el folio impreso dice «23». En la tabla 5, la columna del 4,54 está rotulada «2009» (errata por 2019; el resto de la tabla dice 2006-2019). | VMR, pp. 24-25 | Citar «p. 24 del PDF (folio 23)». |
| Menor | PDM.CA (2024) | **H7. El referente de la tabla 39 no coincide con el texto.** El texto dice que la meta es la del PMEPVU (AMVA-Humboldt), pero la columna del déficit (3,25) usa el estándar de 9 m²/hab, atribuido a la OMS. | POT, pp. 86-87 | Sin efecto sobre 5,75. Coincide con la decisión D7 (referente 9). |

Los cuatro criterios de las reuniones (`decisiones.md`) se cumplen en lo que se revisó. Ninguno de los cortes elegidos es un autorreporte del 100 % ni una meta copiada como logro (los 5,00 quedan excluidos). Los datos son públicos y se pueden encontrar. El universo urbano es el mismo en el numerador y el denominador de cada corte, y los totales (69,41 y 4,16) quedan solo como contexto, así que no hay doble conteo.

## 3. Qué no se pudo verificar
- El estudio primario AMVA-UNAL (2020): de ahí saldría la fórmula exacta del 4,54 y la explicación de la diferencia con 3,95.
- El documento técnico de la actualización del PMEPVU 2022 (AMVA-Humboldt): faltan el área, la población y si aplica la escala de beneficio ambiental. Solo se vio el gráfico de la nota de prensa.
- Si el inventario de EPV urbano del POT 2024 contiene solo áreas públicas. El inventario se titula «elementos del espacio público», pero la p. 11 define el EPV como áreas «públicas o privadas».
- La fuente exacta de la población urbana de 232.903 (2019).

## Veredicto
**Aprobado con observaciones.** Las cifras 4,54, 4,66, 5,75, 69,41, 4,22 y 4,16 están en las páginas y tablas citadas, con el año y el concepto indicados, y los cocientes de 2024 cuadran. Para la ficha hay que dejar constancia de H1 a H3: el 4,54 está rotulado «efectivo», no se puede reproducir con los datos de su propia tabla, y los tres cortes vienen de métodos, poblaciones y delimitaciones distintos. Por eso deben presentarse como cortes, no como una serie comparable.
