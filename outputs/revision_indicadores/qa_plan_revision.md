# QA del plan de revisión indicador por indicador

**Objeto:** `outputs/revision_indicadores/plan_revision.md` (Analizador, 15-nov-2026, PROVISIONAL).
**Revisor:** agente QA · 15-nov-2026 · no se corrigió ningún archivo.
**Contrastado con:** `.claude/memoria/decisiones.md`, `.claude/memoria/grafo_conocimiento.md`, informe de la reunión del 12-nov-2026 (texto sacado de `word/document.xml`), `outputs/sistema_simplificado/trabajo/propuesta_aprobada.md`, `indicadores_simplificados.csv`, `variables_por_dimension.csv`, `trabajo/poblacion_envigado_edad_pped.csv`, `outputs/indicadores/valores_variables.csv`, `scripts/sistema_simplificado/calcular.py` y consultas a la web hechas hoy (API de datos.gov.co, Apéndice 9.1 del PDM en .xlsx, diagnóstico del POT en PDF y encabezados HTTP de los PDF citados).
Las descargas quedaron en `/tmp/claude-1000/-workspaces-Observatorio-Participaci-n-Ciudadana/aec5e3c3-2ba7-48db-85f6-d014a15fc2a6/scratchpad/descargas_qa/`.

---

## 1. Tabla de hallazgos

| # | Severidad | Indicador | Hallazgo | Evidencia | Sugerencia |
|---|---|---|---|---|---|
| 1 | Importante | PC.VEED | La hipótesis de que el listado «solo guarda las vigentes» no cuadra con el propio conjunto de datos. (a) Hay al menos 3 veedurías con fecha final anterior a la actualización del 07-oct-2026: n.º 278 (vence el 23-jun-2025), 218 (21-may-2026) y 283 (06-jul-2026). La 266 trae una fecha de 1931, que es un error de digitación. (b) La columna `n` solo tiene números impares (1, 3, 5 … 89). Eso sugiere que el listado original tenía unas 89 filas y que la publicación perdió una de cada dos, por ejemplo por celdas combinadas. Esta explicación compite con la de «solo vigentes». Aun así, el paso 1 del §3 dice «los datos ya están verificados. *Fácil*». Lo verificado es solo que el conjunto existe y cuántas filas trae, no que la serie sirva. | API `jju5-br3w` (45 filas): campos `duracion_fecha_final` y `n`. Plan, líneas 39, 72 y 114 | Bajar el paso 1 a «fuente encontrada; la serie no es confiable todavía». Pedir a la Personería, o buscar en su esquema de publicación (`rmdh-csmg`), el listado completo de consecutivos. Agregar la hipótesis de las filas perdidas a P7. Ojo: 2024 tiene 1 sola inscripción, así que una línea base en 2024 sería muy inestable |
| 2 | Importante | PDM.CM | El paso 4 propone «agregar 2025 con la Policía», pero el valor principal de 2024 es de Forensis (7,25). La serie quedaría con Forensis en 2024 y Policía en 2025, dos fuentes con definiciones distintas. El propio repositorio dice que «la Policía y el INMLCF miden distinto». Además, el §2.5 compara la tasa de la Policía de 2025 (11,34) con la meta 8, que es de la Secretaría de Movilidad. El mismo plan advierte que es otro universo (C7). | Plan, líneas 26, 99 y 117. `variables_por_dimension.csv`, PDM.CM.O, columna «limitación» | Escoger una sola fuente para toda la serie (la Policía 2019–2025 es la única completa) y dejar Forensis como contraste. Leer el resultado frente a la meta solo después de resolver el salto de 2023 (43 muertes contra una LB de 10,2) |
| 3 | Importante | PDM.ILE (primera infancia) | El plan da por buena la serie de primera infancia 2019–2025 [V-rep] y no menciona dos limitaciones que el repositorio ya documenta. (a) El numerador incluye **mujeres gestantes**, pero el denominador es la población de 0 a 5 años: no es el mismo universo (C7). (b) La fuente cambia en 2024: TerriData hasta 2023 y el informe de gestión en 2024–2025, lo que rompe la comparabilidad de la serie (C5). Esto afecta P1 y P2. | `calcular.py:409-412` (variable «Número de niños, niñas y mujeres gestantes…»; 2024–2025 desde AUX.PDM.ILE.A.ATENDIDOS). `variables_por_dimension.csv`, PDM.ILE.ATEND: «Fuente distinta antes y después de 2023; incluye gestantes» | Agregar las dos limitaciones a la fila de ILE y a P2. Averiguar si TerriData o la Secretaría de Educación separan a las gestantes |
| 4 | Importante | PDM.ICPD | El «año común probable 2020–2024 [S]» no es compatible con lo que el mismo plan dice de los componentes. PDM.CA tiene una serie «débil» (4,66 en 2022 y una sola medición de 5,75, que el CSV fecha en 2024). PDM.SP no tiene serie confirmada. La regla de la Guía es no calcular el índice con entradas faltantes. Con esa regla, el ICPD sería calculable en uno o dos años como mucho, lo que choca con C5. | Plan, líneas 25, 28 y 30 | Hacer la cuenta de años disponibles por componente antes de proponer la ventana. Incluir en P5 qué hacer si CA no tiene serie anual |
| 5 | Menor | PDM.ILE | El plan dice «con 0–6 da 14,56». Esa cifra usa la proyección DANE post-COVID 2020–2035 (población de 14.140), mientras que 16,56 (0–5) usa la proyección 2018–2042. Son dos proyecciones distintas. Con la misma proyección de la hoja, 0–6 da **14,02**, que es lo que ya dice la nota del CSV. La conclusión (0–5 se acerca más a 16,3) no cambia. | `valores_variables.csv:1008` (2059/14140 = 14,56, post-COVID). `indicadores_simplificados.csv:80` (14,02). Recalculado: 2059/14.685 = 14,02 y 2059/12.431 = 16,56 | Usar 14,02 para cumplir la regla de una sola población DANE |
| 6 | Menor | PDM.CA | La fórmula que se propone, 100·(valor − 4,66)/(9 − 4,66), usaría 5,75 como valor observado. Sin embargo, el CSV advierte que 5,75 «no concilia con la LB del AMVA (4,66): otra definición de área». El plan pide explicar la diferencia, pero mientras tanto presenta la fórmula como veredicto. | `variables_por_dimension.csv`, PDM.CA.O. Plan, línea 25 | Dejar la fórmula condicionada a P13 y a que 4,66 y el valor observado se midan igual |
| 7 | Menor | PC.I1 | El veredicto «Se reemplaza (por PC.VEED)» va más allá de D13, que pide **diferenciar** I1 de PP.FID, no reemplazarlo. El grafo lo deja como «Pendiente». El plan sí remite a P8. | `decisiones.md:31`. Grafo, línea 55. Plan, línea 40 | Rotularlo «Pendiente (propuesta: reemplazar)» |
| 8 | Menor | PDM.SEXTO | En la reunión se nombraron como candidatos la «violencia interpersonal» y las lesiones. El Ap. 9.1.2 tiene «Tasa de lesiones personales» (línea 4, R2; LB 113,7 en 2023; meta 112,7; fuente Policía Nacional) y «Tasa de violencia en el contexto familiar» (LB 783, meta 745). El plan no los evalúa. Tampoco revisa «juventudes». | Informe, §5 (45:15–58:59). Ap. 9.1.2, filas A=4 B=1 R2 y A=4 B=2 R1 | Agregarlos al §2.3, aunque sea en la lista de descartados |
| 9 | Menor | PDM.SP | El plan dice «coberturas ≈ 99–100 %: el indicador casi no varía». En la reunión se dijo que el **aseo** tenía una serie con más variación y que el acueducto urbano mostraba saltos. | Informe, §5 (33:00–44:20). Plan, línea 28 | Matizarlo: es cierto para energía y no necesariamente para el aseo |
| 10 | Menor | Fuentes | Tres URL marcadas [S] «solo en el buscador» sí existen: responden HTTP 200 con `application/pdf`. Son las de Superservicios AAA 2023 y UPME ICEE 2024. El plan tampoco anota que en `nudc-7mev` el código de Envigado cambia de `5266` (2011–2022) a `05266` (2023–2024). Si el Ejecutor filtra por código, pierde años. Además, la población de 2021 viene como texto, «27.863». | `curl -I` de hoy. API `nudc-7mev` | Pasar las URL a «existe» y avisar al Ejecutor que filtre por nombre de municipio o por los dos códigos |
| 11 | Menor | PP.resp | El plan trata «proyectos terminados» como respaldo, igual que `propuesta_aprobada.md` («Avance físico (respaldo de entrega)»). Pero en `indicadores_simplificados.csv` figura como `principal`. El plan no señala esa diferencia del repositorio. | `propuesta_aprobada.md:57`. CSV: «Proyectos votados que se terminaron;principal» | Anotar que hay que corregir la columna `tipo` del CSV |
| 12 | Menor | Consistencia | El §2.3 dice «los cuatro candidatos tienen un indicador de resultado», pero la tabla trae 3 y los otros 2 aparecen como descartados. Los cuatro de D10 cuentan seguridad como uno solo. Con homicidio y hurto separados, la frase se presta a confusión. | Plan, líneas 81-91 | Redactarlo de nuevo |
| 13 | Menor | Decisiones | El plan no menciona D14 (encuesta de percepción). No es un indicador, pero es una decisión registrada. P7 sugiere que PC.VEED podría ir en CS, lo que reabriría D12 (CS definido) y contradiría D13, que lo pone en PC. | `decisiones.md:31-32`. Plan, línea 139 | Agregar D14 como «sin acción en este plan». En P7, advertir que mover VEED a CS reabre D12 |

No encontré hallazgos **bloqueantes**.

---

## 2. Verificado sin problema (comprobado por mí hoy)

**Cifras [V] contra la fuente web (muestra de 9 grupos de cifras):**
1. **Veedurías `jju5-br3w`:** el conjunto existe con el título «Veedurías Ciudadanas Personería Envigado», lo publica la Personería de Envigado y se actualizó el 07-oct-2026. Tiene **45 filas** y las 10 columnas que describe el plan. Conteos por año de inscripción: 2012: 1, 2014: 2, 2015: 2, 2016: 2, 2017: 1, 2018: 3, 2019: 2, 2020: 4, **2021: 11**, 2022: 4, 2023: 6, 2024: 1, 2025: 5, 2026: 1. Todos coinciden. El consecutivo va de 118 a 293. Las dos filas con año de registro 2026 y fecha de 2025 son las 291 y 292. El rango de 2021 va de 262 a 275. El conjunto trae nombres y correos. `rmdh-csmg` existe y es el esquema de publicación de la Personería. (La afirmación «solo vigentes» queda cuestionada: ver hallazgo 1).
2. **Cobertura neta `nudc-7mev`, Envigado:** 96,96 (2018), 97,06, 96,85, 98,04, **100,44 (2022)**, 99,68 y 98,45 (2024). Coinciden. La serie va de 2011 a 2024. El salto de 64,24 (2017) a 96,96 (2018) coincide con el cambio de la población de 5 a 16 años, de 41.605 a 28.116, lo que respalda la explicación del cambio de proyección. La cobertura neta de transición (64,78 en 2018 y 63,47 en 2024) también coincide.
3. **Muertes en accidente de tránsito `ha6j-pa2r`, Envigado (05266000):** 22, 12, 21, 29, **43, 20, 30** (2019–2025). Coinciden. La descripción dice «datos hasta 31-jul-2026» y la actualización es del 01-sep-2026. Tasas indicativas: 30/264.637 = 11,34 y 43/256.042 = 16,79. Coinciden.
4. **Homicidios `m8fd-ahd9`, Envigado:** 10, 7, 3, 9, 9, 12, 8 (2019–2025). Coinciden. La serie va de 2003 al 30-sep-2026. 9/256.042 = 3,52. Coincide.
5. **Hurto a personas `4rxi-8m8d`, Envigado (suma de `cantidad`):** 1.601, 1.080, 1.560, 1.554, 1.643, 1.329, 1.327. Coinciden. 1.643/256.042 = 641,7. Coincide.
6. **Apéndice 9.1.2 del PDM** (descargué el .xlsx; la hoja se llama «Apéndice9.1.2.»). Coinciden todas las LB, años, metas y fuentes citadas:
   - homicidio: 3,7 (2023) → 2,7, Policía;
   - hurto a personas: 632,5 → 631,5;
   - mortalidad vial: 10,2 (2023) → 8, Secretaría de Movilidad;
   - espacio público verde: 3,88 (2023) → 3,88, «Estable»; espacio público efectivo: 3,33 → 4;
   - cobertura neta: 100,44 (2022) → 100,44, «Monitoreo»;
   - atención integral de **0 a 6 años**: 16,3 (2023) → 19;
   - acueducto, alcantarillado y energía: 99,73, 99,56 y 99,87 (2023) → 99,9, fuente ECV; **el aseo no tiene indicador de resultado**;
   - régimen subsidiado: 99,5 (2022) → 99,5;
   - procesos de participación: 75 (2023) → 100, fuente DAP;
   - cobertura deportiva: 5,3 → 6; actividad física: 8 (promedio 2020–2023) → 9; población inscrita en programas deportivos: 12,2 (2021, ECV) → 13;
   - déficit cuantitativo de vivienda: 5,07 → 4,07; déficit cualitativo: 11,23;
   - pobreza por ingreso: 13,69 (2023, ECV).

   El «hallazgo transversal» es correcto.
7. **MDM `nkjx-rsq7`:** para Envigado solo hay datos de 2016 a 2020. Coincide. **`tt8i-w5qr`** es «Déficit habitacional, según municipio 2018». Coincide.
8. **POT, diagnóstico de espacio público:** el PDF existe. La «Tabla 39. Indicador de Espacio Público Verde EPV-urbano» tiene las columnas «Línea base AMVA, 2022», «Estándar mínimo EPV (OMS)» y «Medición EPV Urbano». El estándar de 9 m²/hab aparece en el texto. La Tabla 38 da el EPV total, 69,41, con 248.304 habitantes (DANE 2024).
9. **URL:** el listado de la EFPC de 2024 y los informes AAA y de la UPME responden HTTP 200.

**Cifras [V-rep] contra los CSV del repositorio (todas coinciden):**
- PP: FID 76,47 (39/51) y 95,08 (58/61); versión monetaria 103,55 y 100,74; EPP 98,69, 98,83 y 23,26 (numerador 2026 = 5.702.528.148); compromiso 99,56; ≥90 %: 48/50 y 57/61; brecha entre zonas 17,42 y 2,51; terminados 49/61.
- PC: I1 50/51 y 61/61; I2 100 y 92,31; I3 83,33; acciones culminadas 7 y 21; I4 3 y 5; ISPC 88,32.
- CS: PCS 100/100; DSI 0/0 y respaldo 28,57; VRC 66,67 y 20; RPO 94,12 y 50; ICRCS 65,20 y 42,50.
- PDM:
  - Salud: 82,29 … 83,76 y 83,06; MinSalud 87,65 y 82,59.
  - Mortalidad vial: 7,25 → 100; Policía 8,05 → 99,32. La población de 248.304 está confirmada: 20/248.304 = 8,0546. La de salud de 2024 es 260.430.
  - Espacio verde: 5,75 contra 3,88; EPV total 69,41.
  - Primera infancia: 13,36 (2019), 16,56 (2023), 17,53 y 17,12.
  - Saber 11: 267,16, 263,6 y 269,69.
  - CD 38/40; CP 3,23.
- La contradicción que señala el plan entre `ESTADO.md` («ILE.A usa 0-6 por defecto») y el CSV simplificado (0–5) existe (`outputs/indicadores/ESTADO.md:48`).

**Fidelidad a las reuniones (coinciden):**
- D1 (I5 e I6 eliminados), D2 y D8 (ILE: cobertura neta + primera infancia, conservar el nombre, fórmula por definir, conciliar 0–5 y 0–6).
- D3 y D5 (cobertura total; la meta del subsidiado no sirve como meta). Revisar la tabla del subsidiado 2022 quedó resuelto con el Ap. 9.1.2.
- D6, D7 (4,66, 5,75 y 9; separar verde, efectivo y total; buscar 2024–2025).
- D9 (4 servicios, cobertura total, sin internet, ponderación como posibilidad).
- D10 (vínculo con el PDM y serie de 4–5 años: los dos candidatos principales lo cumplen).
- D11 (CP sale; CD sin tasa del Índer). D12 (PP y CS definidos: el plan no los cambia, solo pregunta en P10 y P11). D4 queda como pregunta. D13 (veedurías con LB; mantener I2).
- Compromisos: el orden de los pasos sigue el §6 del informe, y la base de datos va al final.

**Consistencia interna:** la columna «Paso» de las tablas coincide con el orden del §3 (1 = VEED, 2 = SEXTO, … 13 = POLPUB). Las preguntas P1–P13 cubren todos los «Qué falta» que piden una decisión. Las marcas [V], [V-rep], [V-reu] y [S] se usan con cuidado: lo que no se pudo ver quedó como [S]. El aviso sobre datos personales en el conjunto de veedurías es correcto.

---

## 3. No verificable (y por qué)

- **Forensis 2024 (7,25, p. 440) y Forensis 2025:** no abrí el PDF del INMLCF. El plan dice que el sitio responde 403. Queda como [V-rep], tomado del CSV.
- **Valores numéricos de la Tabla 39 del POT (4,66 y 5,75):** las celdas no aparecen en los flujos de texto del PDF (probablemente están como imagen) y aquí no hay herramientas de PDF instaladas. Encontré la tabla y sus encabezados. Los valores se apoyan en el informe de la reunión [V-reu].
- **TerriData** (servicios públicos, salud, atendidos de primera infancia): la aplicación se carga con JavaScript. No la consulté.
- **Contenido municipal de los informes AAA de Superservicios y del ICEE de la UPME:** solo comprobé que las URL existen.
- **Explicación de por qué hay 43 muertes en 2023 (Policía) frente a la LB 10,2 de Movilidad**, y si se cuenta el lugar del hecho o la residencia: no se puede saber solo con el conjunto de datos.
- **Notas manuscritas de las reuniones (R1 y hoja 3 de R2):** no están en el repositorio. Contrasté contra su transcripción en `decisiones.md`. El informe del 12-nov no menciona las veedurías, así que D13 depende solo de las notas.
- **Ap. 9.1.2, columna ODS** (p. ej., 16.1.1 para el homicidio): no la revisé.

---

## 4. Veredicto

**Aprobado con observaciones.**

El plan es fiel a las decisiones D1–D14 y a los criterios C1–C8. Las cifras [V] de la muestra coinciden con la fuente, y las [V-rep] con los CSV. No hay bloqueantes. Antes de que el Ejecutor empiece conviene atender los cuatro hallazgos importantes:
1. la serie de veedurías no se puede usar todavía como flujo anual;
2. hay que escoger una sola fuente para la mortalidad vial;
3. hay que documentar que el numerador de primera infancia incluye gestantes y que la fuente cambia en 2024;
4. hay que revisar la ventana de años del ICPD.
